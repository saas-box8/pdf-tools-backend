import json
import os
import shutil
import subprocess
import tempfile
import uuid
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename


app = Flask(__name__, static_folder=".", static_url_path="")
CORS(app)

app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024

ALLOWED_EXTENSIONS = {".docx", ".doc", ".odt", ".rtf"}

QUALITY_PRESETS = {
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


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def json_bool(value: bool) -> str:
    return "true" if value else "false"


def build_pdf_filter_options(quality: str) -> str:
    preset = QUALITY_PRESETS.get(
        quality,
        QUALITY_PRESETS["high"],
    )

    options = {
        "UseLosslessCompression": {
            "type": "boolean",
            "value": json_bool(preset["UseLosslessCompression"]),
        },
        "Quality": {
            "type": "long",
            "value": str(int(preset["Quality"])),
        },
        "ReduceImageResolution": {
            "type": "boolean",
            "value": json_bool(preset["ReduceImageResolution"]),
        },
        "MaxImageResolution": {
            "type": "long",
            "value": str(int(preset["MaxImageResolution"])),
        },
        "SelectPdfVersion": {
            "type": "long",
            "value": "17",
        },
        "ExportBookmarks": {
            "type": "boolean",
            "value": "true",
        },
    }

    return json.dumps(options, separators=(",", ":"))


def locate_soffice():
    for candidate in ("soffice", "libreoffice"):
        path = shutil.which(candidate)
        if path:
            return path
    return None


def convert_with_libreoffice(
    input_path: Path,
    output_dir: Path,
    quality: str,
) -> Path:
    soffice = locate_soffice()

    if not soffice:
        raise RuntimeError(
            "LibreOffice was not found. "
            "Please contact support."
        )

    filter_options = build_pdf_filter_options(quality)
    convert_to_arg = f"pdf:writer_pdf_Export:{filter_options}"

    profile_dir = output_dir / "lo-profile"
    profile_dir.mkdir(parents=True, exist_ok=True)
    profile_uri = profile_dir.resolve().as_uri()

    env = os.environ.copy()
    env.setdefault("LANG", "en_US.UTF-8")

    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        "--nolockcheck",
        "--norestore",
        f"-env:UserInstallation={profile_uri}",
        "--convert-to",
        convert_to_arg,
        "--outdir",
        str(output_dir),
        str(input_path),
    ]

    completed = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        env=env,
        check=False,
    )

    output_pdf = output_dir / f"{input_path.stem}.pdf"

    if completed.returncode != 0 or not output_pdf.exists():
        raise RuntimeError(
            completed.stderr.strip()
            or completed.stdout.strip()
            or "LibreOffice conversion failed."
        )

    return output_pdf


@app.get("/")
def home():
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
        libreoffice=bool(soffice),
        libreoffice_bin=soffice,
    )


@app.route("/convert", methods=["POST"])
def convert_word():
    upload = request.files.get("file")

    if not upload:
        return jsonify(error="No file uploaded."), 400

    if not upload.filename:
        return jsonify(error="No selected file."), 400

    if not allowed_file(upload.filename):
        return jsonify(
            error=f"Unsupported file type. Allowed: {', '.join(sorted(ALLOWED_EXTENSIONS))}"
        ), 400

    quality = request.form.get("quality", "high").strip().lower()
    if quality not in QUALITY_PRESETS:
        quality = "high"

    request_id = uuid.uuid4().hex

    work_dir = (
        Path(tempfile.gettempdir()) / f"wordpdf_{request_id}"
    )
    work_dir.mkdir(parents=True, exist_ok=True)

    filename = secure_filename(upload.filename)
    input_path = work_dir / filename

    try:
        upload.save(str(input_path))

        output_pdf = convert_with_libreoffice(
            input_path,
            work_dir,
            quality,
        )

        response = send_file(
            output_pdf,
            as_attachment=True,
            download_name=f"{input_path.stem}.pdf",
            mimetype="application/pdf",
            max_age=0,
        )

        response.headers["Cache-Control"] = "no-store"
        response.headers["X-Engine"] = "libreoffice"
        response.headers["X-Quality"] = quality

        @response.call_on_close
        def cleanup():
            try:
                shutil.rmtree(work_dir, ignore_errors=True)
            except Exception:
                pass

        return response

    except RuntimeError as exc:
        shutil.rmtree(work_dir, ignore_errors=True)
        return jsonify(error=str(exc)), 500

    except Exception:
        shutil.rmtree(work_dir, ignore_errors=True)
        return jsonify(error="Conversion failed. Please try again."), 500


@app.errorhandler(413)
def file_too_large(_):
    return jsonify(error="File too large (max 100 MB)."), 413


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.getenv("PORT", "5000")),
        debug=False,
    )
