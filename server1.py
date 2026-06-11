"""
Word / ODT / RTF → PDF conversion backend
==========================================
Engine   : LibreOffice headless (primary) — works on Render free Docker tier
Fallback : docx2pdf / Microsoft Word      — Windows & macOS only

Render free-tier notes
----------------------
• Uses Docker runtime (see Dockerfile.word-to-pdf) so apt-get installs LO once.
• workers=1  → stays within 512 MB RAM (LO alone needs ~250 MB resident).
• timeout=120 → first-call LO warm-up can take 20-40 s; default 30 s kills it.
• Each conversion gets an isolated LO user-profile dir so concurrent requests
  (or a retry on the same process) never collide.
• Temp dirs are removed in a daemon thread 30 s after the response is sent,
  which is more reliable than response.call_on_close under Gunicorn.
"""

from __future__ import annotations

import json
import logging
import os
import platform
import shutil
import subprocess
import tempfile
import threading
import time
import uuid
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

# Optional: Microsoft Word conversion (Windows/macOS only)
try:
    from docx2pdf import convert as _word_convert
except Exception:
    _word_convert = None

# ──────────────────────────────────────────────────────────────────────────────
# Logging
# ──────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

# ──────────────────────────────────────────────────────────────────────────────
# Flask app  (static_folder="." lets the root route serve wordtopdf.html
#             if it exists alongside this file — matches existing server1 behaviour)
# ──────────────────────────────────────────────────────────────────────────────
app = Flask(__name__, static_folder=".", static_url_path="")
CORS(app, resources={r"/*": {"origins": "*"}})
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024   # 100 MB

ALLOWED_EXTENSIONS: set[str] = {".docx", ".doc", ".odt", ".rtf"}

# ──────────────────────────────────────────────────────────────────────────────
# PDF quality presets — fed to LibreOffice writer_pdf_Export filter
# ──────────────────────────────────────────────────────────────────────────────
QUALITY_PRESETS: dict[str, dict] = {
    "standard": {
        "UseLosslessCompression": False,
        "Quality": 85,
        "ReduceImageResolution": False,
        "MaxImageResolution": 300,
    },
    "high": {
        "UseLosslessCompression": False,
        "Quality": 95,
        "ReduceImageResolution": False,
        "MaxImageResolution": 300,
    },
    "compressed": {
        "UseLosslessCompression": False,
        "Quality": 65,
        "ReduceImageResolution": True,
        "MaxImageResolution": 150,
    },
    "print": {
        "UseLosslessCompression": True,
        "Quality": 100,
        "ReduceImageResolution": False,
        "MaxImageResolution": 300,
    },
}

# ──────────────────────────────────────────────────────────────────────────────
# LibreOffice serialisation lock
# LO cannot run two conversions against the same user-profile simultaneously.
# The semaphore limits concurrent LO calls to 1 within a single process.
# (With workers=1 on Render free this is belt-and-suspenders, but harmless.)
# ──────────────────────────────────────────────────────────────────────────────
_LO_LOCK = threading.Semaphore(1)
_LO_TIMEOUT: int = int(os.environ.get("LO_TIMEOUT", "120"))


# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────

def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def _jbool(v: bool) -> str:
    return "true" if v else "false"


def build_pdf_filter_options(quality: str) -> str:
    """
    Serialise a LibreOffice writer_pdf_Export filter-options JSON string.
    SelectPdfVersion 17 = PDF 1.7 (maximally compatible).
    EmbedStandardFonts keeps text rendering accurate on any viewer.
    """
    preset = QUALITY_PRESETS.get(quality, QUALITY_PRESETS["high"])
    options = {
        "UseLosslessCompression":  {"type": "boolean", "value": _jbool(preset["UseLosslessCompression"])},
        "Quality":                 {"type": "long",    "value": str(int(preset["Quality"]))},
        "ReduceImageResolution":   {"type": "boolean", "value": _jbool(preset["ReduceImageResolution"])},
        "MaxImageResolution":      {"type": "long",    "value": str(int(preset["MaxImageResolution"]))},
        "SelectPdfVersion":        {"type": "long",    "value": "17"},
        "ExportBookmarks":         {"type": "boolean", "value": "true"},
        "EmbedStandardFonts":      {"type": "boolean", "value": "true"},
    }
    return json.dumps(options, separators=(",", ":"))


def locate_soffice() -> str | None:
    """Find the LibreOffice binary on PATH or well-known install paths."""
    candidates = [
        "soffice",
        "libreoffice",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/usr/lib/libreoffice/program/soffice",
        "/opt/libreoffice/program/soffice",
    ]
    for c in candidates:
        found = shutil.which(c) or (c if Path(c).is_file() else None)
        if found:
            logger.info("Located LibreOffice at: %s", found)
            return found
    return None


def _cleanup_later(path: Path, delay: float = 30.0) -> None:
    """Delete *path* in a daemon thread after *delay* seconds."""
    def _run():
        time.sleep(delay)
        shutil.rmtree(path, ignore_errors=True)
        logger.debug("Cleaned up temp dir: %s", path)
    threading.Thread(target=_run, daemon=True).start()


# ──────────────────────────────────────────────────────────────────────────────
# Conversion engines
# ──────────────────────────────────────────────────────────────────────────────

def convert_with_libreoffice(
    input_path: Path,
    output_dir: Path,
    quality: str = "high",
) -> Path:
    """
    Convert *input_path* to PDF using LibreOffice headless.

    Critical flags
    --------------
    --nolockcheck         : lets multiple processes co-exist (each gets its own profile)
    SAL_USE_VCLPLUGIN=svp : headless VCL plugin — no display required
    DISPLAY=""            : prevent any accidental X11 connection attempts
    Isolated profile dir  : separate UUID dir per request so parallel calls never clash
    Semaphore             : prevents two LO subprocesses racing inside one Python process
    """
    soffice = locate_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice not found. "
            "Ensure the Docker image installs libreoffice-writer."
        )

    filter_opts = build_pdf_filter_options(quality)
    convert_to_arg = f"pdf:writer_pdf_Export:{filter_opts}"

    # Per-request isolated profile — essential for concurrency safety
    profile_dir = output_dir / f"lo-profile-{uuid.uuid4().hex}"
    profile_dir.mkdir(parents=True, exist_ok=True)
    profile_uri = profile_dir.resolve().as_uri()

    env = os.environ.copy()
    env.setdefault("LANG", "en_US.UTF-8")
    env.setdefault("LC_ALL", "en_US.UTF-8")
    env["DISPLAY"] = ""                   # no display
    env["SAL_USE_VCLPLUGIN"] = "svp"      # headless VCL
    env["HOME"] = str(output_dir)         # prevent ~/.config collisions

    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        "--nolockcheck",
        "--norestore",
        "--nocrashreport",
        f"-env:UserInstallation={profile_uri}",
        "--convert-to", convert_to_arg,
        "--outdir", str(output_dir),
        str(input_path),
    ]

    logger.info("LO cmd: %s", " ".join(cmd))

    acquired = _LO_LOCK.acquire(timeout=_LO_TIMEOUT)
    if not acquired:
        raise RuntimeError("LibreOffice is busy — please retry in a moment.")

    try:
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            env=env,
            timeout=_LO_TIMEOUT,
            check=False,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError(f"LibreOffice timed out after {_LO_TIMEOUT}s.")
    finally:
        _LO_LOCK.release()

    output_pdf = output_dir / f"{input_path.stem}.pdf"

    if result.returncode != 0 or not output_pdf.exists():
        detail = result.stderr.strip() or result.stdout.strip() or "LibreOffice conversion failed."
        logger.error("LO error (rc=%d): %s", result.returncode, detail)
        raise RuntimeError(detail)

    logger.info(
        "LO produced %s (%d bytes)", output_pdf.name, output_pdf.stat().st_size
    )
    return output_pdf


def convert_with_word(input_path: Path, output_dir: Path) -> Path:
    """Convert using Microsoft Word via docx2pdf (Windows / macOS only)."""
    if _word_convert is None:
        raise RuntimeError("docx2pdf is not installed.")
    if platform.system().lower() not in {"windows", "darwin"}:
        raise RuntimeError("Microsoft Word conversion only works on Windows / macOS.")

    _word_convert(str(input_path), str(output_dir))
    output_pdf = output_dir / f"{input_path.stem}.pdf"
    if not output_pdf.exists():
        raise RuntimeError("Word conversion produced no output.")
    return output_pdf


# ──────────────────────────────────────────────────────────────────────────────
# Routes
# ──────────────────────────────────────────────────────────────────────────────

@app.get("/")
def home():
    """
    Serve wordtopdf.html if it exists next to this file,
    otherwise return a JSON status response (matches existing behaviour).
    """
    html = Path(__file__).parent / "wordtopdf.html"
    if html.exists():
        return app.send_static_file("wordtopdf.html")
    return jsonify(
        status="running",
        tool="word-to-pdf",
        engine="libreoffice",
        accepts=sorted(ALLOWED_EXTENSIONS),
        quality_presets=sorted(QUALITY_PRESETS.keys()),
    )


@app.get("/health")
def health():
    soffice = locate_soffice()
    return jsonify(
        status="ok" if soffice else "degraded",
        server="word-to-pdf",
        version="2.0",
        libreoffice=bool(soffice),
        libreoffice_bin=soffice,
        platform=platform.system(),
        word_available=_word_convert is not None,
        accepts=sorted(ALLOWED_EXTENSIONS),
        quality_presets=sorted(QUALITY_PRESETS.keys()),
    )


@app.route("/convert", methods=["POST", "OPTIONS"])
def convert_word():
    """
    POST /convert
    ─────────────
    Form fields:
      file     (required) — .docx / .doc / .odt / .rtf
      quality  (optional) — standard | high | compressed | print   (default: high)
      engine   (optional) — auto | libreoffice | word               (default: auto)

    Response headers (on success):
      X-Engine  — engine actually used
      X-Quality — quality preset applied
    """
    if request.method == "OPTIONS":
        return ("", 204)

    upload = request.files.get("file")
    if not upload:
        return jsonify(error="No file uploaded."), 400
    if not upload.filename:
        return jsonify(error="Empty filename."), 400
    if not allowed_file(upload.filename):
        return jsonify(
            error=f"Unsupported file type. Allowed: {', '.join(sorted(ALLOWED_EXTENSIONS))}"
        ), 400

    quality = request.form.get("quality", "high").strip().lower()
    if quality not in QUALITY_PRESETS:
        quality = "high"

    engine = request.form.get("engine", "auto").strip().lower()

    request_id = uuid.uuid4().hex
    work_dir = Path(tempfile.gettempdir()) / f"wordpdf_{request_id}"
    work_dir.mkdir(parents=True, exist_ok=True)

    filename = secure_filename(upload.filename)
    input_path = work_dir / filename

    try:
        upload.save(str(input_path))
        logger.info(
            "Request %s: file=%s size=%d quality=%s engine=%s",
            request_id, filename, input_path.stat().st_size, quality, engine,
        )

        # ── Pick engine ──────────────────────────────────────────────────────
        engine_used: str

        if engine == "word":
            output_pdf = convert_with_word(input_path, work_dir)
            engine_used = "word"

        elif engine == "libreoffice":
            output_pdf = convert_with_libreoffice(input_path, work_dir, quality)
            engine_used = "libreoffice"

        else:
            # auto: try LibreOffice first (works on Linux / Docker / Render),
            #       fall back to Word on Windows/macOS if LO is absent.
            try:
                output_pdf = convert_with_libreoffice(input_path, work_dir, quality)
                engine_used = "libreoffice"
            except RuntimeError as lo_err:
                logger.warning("LO failed (%s), falling back to Word…", lo_err)
                try:
                    output_pdf = convert_with_word(input_path, work_dir)
                    engine_used = "word"
                except RuntimeError as word_err:
                    raise RuntimeError(
                        f"Both engines failed — "
                        f"LibreOffice: {lo_err} | Word: {word_err}"
                    ) from None

        # ── Stream response ──────────────────────────────────────────────────
        response = send_file(
            output_pdf,
            as_attachment=True,
            download_name=f"{Path(filename).stem}.pdf",
            mimetype="application/pdf",
            conditional=False,
            max_age=0,
        )
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
        response.headers["X-Engine"] = engine_used
        response.headers["X-Quality"] = quality

        _cleanup_later(work_dir, delay=30.0)
        return response

    except RuntimeError as exc:
        _cleanup_later(work_dir, delay=5.0)
        logger.error("Conversion error [%s]: %s", request_id, exc)
        return jsonify(error=str(exc)), 500

    except Exception as exc:
        _cleanup_later(work_dir, delay=5.0)
        logger.exception("Unexpected error [%s]", request_id)
        return jsonify(error=f"Unexpected error: {exc}"), 500


@app.errorhandler(413)
def file_too_large(_):
    return jsonify(error="File too large (max 100 MB)."), 413


# ──────────────────────────────────────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    logger.info("Starting word-to-pdf server on port %d (debug=%s)", port, debug)
    app.run(host="0.0.0.0", port=port, debug=debug)
