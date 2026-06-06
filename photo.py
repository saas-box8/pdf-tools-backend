"""
PDF → JPG converter — pure Python, no LibreOffice required.
Supports PDF input only (office file conversion removed for free-plan compatibility).
Uses PyMuPDF for rendering. Works on Render free plan.
"""

from __future__ import annotations

import io
import shutil
import tempfile
import uuid
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import fitz  # PyMuPDF
from flask import Flask, jsonify, render_template, request, send_file
from flask_cors import CORS
from PIL import Image
from werkzeug.utils import secure_filename

app = Flask(__name__, template_folder="templates", static_folder="static")
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024

CORS(app, resources={r"/*": {"origins": "*"}})

BASE_DIR = Path(tempfile.gettempdir()) / "pdf_to_jpg_tool"
BASE_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_EXTS = {".pdf"}


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTS


def render_pdf_to_jpgs(pdf_path: Path, outdir: Path, scale: float):
    outdir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf_path))
    matrix = fitz.Matrix(scale, scale)
    jpg_paths = []
    try:
        for i in range(len(doc)):
            page = doc[i]
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            out = outdir / f"page_{i+1:03d}.jpg"
            img.save(out, "JPEG", quality=90, optimize=True)
            jpg_paths.append(out)
    finally:
        doc.close()
    return jpg_paths


def make_zip(files):
    buffer = io.BytesIO()
    with ZipFile(buffer, "w", ZIP_DEFLATED) as z:
        for f in files:
            z.write(f, arcname=f.name)
    buffer.seek(0)
    return buffer


@app.get("/")
def index():
    return jsonify({"status": "running", "tool": "pdf-to-jpg"})


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


@app.post("/convert")
def convert():
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        return jsonify({"error": "No file uploaded"}), 400
    if not allowed_file(uploaded.filename):
        return jsonify({"error": "Only PDF files are supported"}), 400

    quality = (request.form.get("quality") or "high").lower()
    size_mode = (request.form.get("size") or "original").lower()
    scale_map = {"small": 1.5, "original": 2.0, "large": 3.0}
    scale = scale_map.get(size_mode, 2.0)

    job_id = uuid.uuid4().hex
    workdir = BASE_DIR / job_id
    input_dir = workdir / "input"
    output_dir = workdir / "output"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    filename = secure_filename(uploaded.filename)
    input_path = input_dir / filename
    uploaded.save(input_path)

    try:
        jpgs = render_pdf_to_jpgs(input_path, output_dir, scale)

        q = {"compressed": 75, "standard": 85, "high": 95}.get(quality, 90)
        for img_path in jpgs:
            img = Image.open(img_path).convert("RGB")
            img.save(img_path, "JPEG", quality=q, optimize=True)

        zip_buffer = make_zip(jpgs)
        return send_file(
            zip_buffer,
            mimetype="application/zip",
            as_attachment=True,
            download_name=f"{input_path.stem}_jpgs.zip",
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


if __name__ == "__main__":
    from waitress import serve
    serve(app, host="0.0.0.0", port=5000)
