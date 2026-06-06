"""
HTML / URL → PDF converter — pure Python, no Playwright/Chromium required.
Uses WeasyPrint for HTML→PDF rendering. Works on Render free plan.
"""

from __future__ import annotations

import io
import logging
import os
import shutil
import tempfile
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from flask import Flask, after_this_request, jsonify, request, send_file
from flask_cors import CORS
from pypdf import PdfReader, PdfWriter

# WeasyPrint — pure Python HTML/CSS → PDF
try:
    import weasyprint
    WEASYPRINT_OK = True
except Exception as _wp_err:
    WEASYPRINT_OK = False
    print(f"[web.py] WeasyPrint unavailable: {_wp_err}")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
log = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024

DEFAULT_TIMEOUT = 30

PAGE_SIZES = {
    "a4": "A4",
    "a3": "A3",
    "a5": "A5",
    "letter": "Letter",
    "legal": "Legal",
    "tabloid": "Tabloid",
}


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def fetch_url_html(url: str) -> tuple[str, str]:
    url = (url or "").strip()
    if not url:
        raise ValueError("URL is empty.")
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    resp = requests.get(
        url,
        timeout=DEFAULT_TIMEOUT,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            )
        },
    )
    resp.raise_for_status()
    return resp.text, resp.url


def ensure_full_html(html: str, title: str = "Document", base_url: str | None = None) -> str:
    html = html or ""
    lower = html.lower()
    safe_title = title.replace("<", "&lt;").replace(">", "&gt;")
    base_tag = f'<base href="{base_url}">' if base_url else ""

    if "<html" in lower and "<body" in lower:
        # Inject base tag if missing
        if base_url and "<base" not in lower:
            html = html.replace("<head>", f"<head>{base_tag}", 1)
            html = html.replace("<HEAD>", f"<HEAD>{base_tag}", 1)
        return html

    body = html if "<body" in lower else f"<body>{html}</body>"
    return (
        f"<!doctype html><html><head>"
        f"<meta charset='utf-8'>{base_tag}"
        f"<title>{safe_title}</title>"
        f"<style>body{{font-family:Arial,sans-serif;font-size:11pt;line-height:1.4;margin:0}}"
        f"img{{max-width:100%;height:auto}}"
        f"table{{border-collapse:collapse;width:100%}}"
        f"td,th{{border:1px solid #ccc;padding:4px 8px}}</style>"
        f"</head>{body}</html>"
    )


def build_page_css(page_size: str, orientation: str, margin_mm: str,
                   backgrounds: bool, custom_width: str, custom_height: str) -> str:
    if page_size == "custom":
        size_val = f"{custom_width}mm {custom_height}mm"
    else:
        size_val = PAGE_SIZES.get(page_size, "A4")
        if orientation == "landscape":
            size_val += " landscape"

    bg = "print-color-adjust: exact;" if backgrounds else ""

    try:
        margin = float(margin_mm)
    except (ValueError, TypeError):
        margin = 10.0

    return f"""
@page {{
    size: {size_val};
    margin: {margin}mm;
    {bg}
}}
"""


def html_to_pdf_bytes(html: str, *, page_size: str = "a4", orientation: str = "portrait",
                      margin_mm: str = "10", backgrounds: bool = False,
                      custom_width: str = "210", custom_height: str = "297",
                      base_url: str | None = None) -> bytes:
    if not WEASYPRINT_OK:
        raise RuntimeError(
            "WeasyPrint is not available. Make sure it is installed."
        )

    page_css = build_page_css(page_size, orientation, margin_mm,
                               backgrounds, custom_width, custom_height)

    # Inject @page CSS into <head>
    if "<head>" in html:
        html = html.replace("<head>", f"<head><style>{page_css}</style>", 1)
    elif "<HEAD>" in html:
        html = html.replace("<HEAD>", f"<HEAD><style>{page_css}</style>", 1)
    else:
        html = f"<style>{page_css}</style>" + html

    doc = weasyprint.HTML(string=html, base_url=base_url)
    return doc.write_pdf()


def compress_pdf(pdf_bytes: bytes) -> bytes:
    """Re-write PDF through pypdf to strip metadata and compress streams."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    for page in reader.pages:
        page.compress_content_streams()
        writer.add_page(page)
    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out.read()


# ─────────────────────────────────────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────────────────────────────────────

@app.get("/health")
def health():
    return jsonify({
        "status": "ok",
        "engine": "weasyprint" if WEASYPRINT_OK else "unavailable",
    })


@app.get("/")
def home():
    return jsonify({"status": "running", "tool": "html-to-pdf"})


@app.post("/preview-url")
def preview_url():
    url = (request.json or {}).get("url", "").strip()
    if not url:
        return jsonify({"error": "URL is required."}), 400
    try:
        html, final_url = fetch_url_html(url)
        soup = BeautifulSoup(html, "html.parser")
        title = soup.title.string.strip() if soup.title and soup.title.string else final_url
        return jsonify({"title": title, "url": final_url, "ok": True})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.post("/convert")
def convert():
    try:
        if not WEASYPRINT_OK:
            return jsonify({"error": "WeasyPrint rendering engine is not available."}), 500

        files = request.files.getlist("files")
        html_code = (request.form.get("html_code") or "").strip()
        url = (request.form.get("url") or "").strip()

        page_size = (request.form.get("page_size", "a4")).lower().strip()
        orientation = (request.form.get("orientation", "portrait")).lower().strip()
        margin_mm = request.form.get("margin_mm", "10")
        backgrounds = request.form.get("backgrounds", "0") == "1"
        compress = request.form.get("compress", "0") == "1"
        custom_width = request.form.get("custom_width", "210")
        custom_height = request.form.get("custom_height", "297")

        if page_size not in set(PAGE_SIZES.keys()) | {"custom"}:
            page_size = "a4"
        if orientation not in {"portrait", "landscape"}:
            orientation = "portrait"

        if not files and not html_code and not url:
            return jsonify({"error": "Provide HTML files, pasted code, or a URL."}), 400

        title = "Document"
        final_html = ""
        base_url_hint: str | None = None

        if files and any(f.filename for f in files):
            # Find first HTML file
            html_file = next((f for f in files if f.filename and
                               Path(f.filename).suffix.lower() in {".html", ".htm"}), None)
            if not html_file:
                return jsonify({"error": "No HTML file found in upload."}), 400
            title = Path(html_file.filename).stem
            final_html = ensure_full_html(
                html_file.read().decode("utf-8", errors="replace"), title=title
            )

        elif html_code:
            title = "Pasted HTML"
            final_html = ensure_full_html(html_code, title=title)

        else:
            fetched_html, final_url_str = fetch_url_html(url)
            title = final_url_str
            base_url_hint = final_url_str
            final_html = ensure_full_html(fetched_html, title=title, base_url=base_url_hint)

        render_kwargs = dict(
            page_size=page_size,
            orientation=orientation,
            margin_mm=margin_mm,
            backgrounds=backgrounds,
            custom_width=custom_width,
            custom_height=custom_height,
            base_url=base_url_hint,
        )

        pdf_bytes = html_to_pdf_bytes(final_html, **render_kwargs)

        if compress:
            try:
                pdf_bytes = compress_pdf(pdf_bytes)
            except Exception:
                pass  # Return uncompressed on failure

        safe_title = "".join(c for c in title if c.isalnum() or c in "-_ ")[:60].strip() or "document"
        download_name = f"{safe_title}.pdf"

        return send_file(
            io.BytesIO(pdf_bytes),
            as_attachment=True,
            download_name=download_name,
            mimetype="application/pdf",
            max_age=0,
        )

    except requests.RequestException as exc:
        return jsonify({"error": f"Could not fetch URL: {exc}"}), 400
    except Exception as exc:
        log.exception("Unexpected error during conversion")
        return jsonify({"error": str(exc)}), 500


if __name__ == "__main__":
    from waitress import serve as waitress_serve
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", 5000))
    log.info("Starting HTML→PDF server on http://%s:%d", host, port)
    waitress_serve(app, host=host, port=port, threads=4)
