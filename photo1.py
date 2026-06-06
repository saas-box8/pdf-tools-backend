"""
JPG/PNG → PDF converter — pure Python, no LibreOffice required.
Supports image input (JPG/JPEG/PNG) and PDF passthrough only.
Office file conversion removed for free-plan compatibility.
Uses PyMuPDF + Pillow. Works on Render free plan.
"""

from __future__ import annotations

import io
import logging
import os
import shutil
import tempfile
import uuid
from pathlib import Path

import fitz
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from PIL import Image, UnidentifiedImageError
from werkzeug.utils import secure_filename

IMAGE_EXTS = {".jpg", ".jpeg", ".png"}
PDF_EXTS = {".pdf"}
ALLOWED_EXTS = IMAGE_EXTS | PDF_EXTS

BASE_DIR = Path(tempfile.gettempdir()) / "jpg_to_pdf_tool"
BASE_DIR.mkdir(parents=True, exist_ok=True)

MAX_UPLOAD_MB = 100

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
CORS(app, resources={r"/*": {"origins": "*"}})


@app.after_request
def add_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Headers"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    return response


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTS


def cleanup_dir(path: Path) -> None:
    shutil.rmtree(path, ignore_errors=True)


def jpeg_quality_from_mode(mode: str) -> int:
    return {"a4": 95, "letter": 85, "fit-image": 75}.get((mode or "a4").lower().strip(), 95)


def margin_to_points(margin_mode: str) -> int:
    return {"none": 0, "small": 8, "medium": 18, "large": 28, "xlarge": 40}.get(
        (margin_mode or "small").lower().strip(), 8
    )


def build_pdf_from_images(image_paths: list[Path], output_pdf: Path,
                           quality_mode: str, margin_mode: str) -> None:
    jpeg_quality = jpeg_quality_from_mode(quality_mode)
    margin_pt = margin_to_points(margin_mode)
    doc = fitz.open()
    try:
        for image_path in image_paths:
            try:
                with Image.open(image_path) as img:
                    if img.mode != "RGB":
                        img = img.convert("RGB")
                    img_w, img_h = img.size
                    if img_h == 0:
                        raise RuntimeError(f"Invalid image dimensions: {image_path.name}")
                    ratio = img_w / img_h
                    content_w = 842 if ratio >= 1 else 595
                    content_h = content_w / ratio
                    page_w = content_w + (margin_pt * 2)
                    page_h = content_h + (margin_pt * 2)
                    page = doc.new_page(width=page_w, height=page_h)
                    rect = fitz.Rect(margin_pt, margin_pt,
                                     margin_pt + content_w, margin_pt + content_h)
                    buffer = io.BytesIO()
                    img.save(buffer, format="JPEG", quality=jpeg_quality, optimize=True)
            except UnidentifiedImageError:
                raise RuntimeError(f"Invalid image file: {image_path.name}")
            except Exception as exc:
                raise RuntimeError(f"Could not process image: {image_path.name}") from exc
            page.insert_image(rect, stream=buffer.getvalue(), keep_proportion=False)
        doc.save(str(output_pdf))
    finally:
        doc.close()


@app.get("/")
def home():
    return jsonify({"status": "running", "service": "JPG to PDF API"})


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/convert", methods=["GET", "POST", "OPTIONS"])
def convert():
    if request.method == "GET":
        return jsonify({"status": "ready", "message": "Send POST request with files"}), 200
    if request.method == "OPTIONS":
        return jsonify({"status": "ok"}), 200

    files = request.files.getlist("files")
    if not files:
        single = request.files.get("file")
        if single:
            files = [single]
    if not files:
        return jsonify({"error": "No files uploaded."}), 400

    job_id = uuid.uuid4().hex
    workdir = BASE_DIR / job_id
    input_dir = workdir / "input"
    output_dir = workdir / "output"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        saved_paths: list[Path] = []
        for idx, uploaded in enumerate(files, start=1):
            if not uploaded.filename:
                continue
            if not allowed_file(uploaded.filename):
                cleanup_dir(workdir)
                return jsonify({"error": f"Unsupported file type: {uploaded.filename}. Only JPG/PNG/PDF supported."}), 400
            safe_name = secure_filename(uploaded.filename) or f"file_{idx}"
            unique_name = f"{uuid.uuid4().hex}_{safe_name}"
            input_path = input_dir / unique_name
            uploaded.save(input_path)
            saved_paths.append(input_path)

        if not saved_paths:
            cleanup_dir(workdir)
            return jsonify({"error": "No valid files uploaded."}), 400

        quality = (request.form.get("quality") or "a4").lower().strip()
        margin_mode = (request.form.get("margin") or "small").lower().strip()
        ext_set = {p.suffix.lower() for p in saved_paths}

        # CASE 1: Images → PDF
        if ext_set.issubset(IMAGE_EXTS):
            output_name = f"{saved_paths[0].stem}.pdf" if len(saved_paths) == 1 else "images_to_pdf.pdf"
            output_pdf = output_dir / output_name
            build_pdf_from_images(saved_paths, output_pdf, quality, margin_mode)
            response = send_file(str(output_pdf), mimetype="application/pdf",
                                  as_attachment=True, download_name=output_name)
            response.call_on_close(lambda: cleanup_dir(workdir))
            return response

        # CASE 2: PDF passthrough
        if len(saved_paths) == 1 and saved_paths[0].suffix.lower() in PDF_EXTS:
            response = send_file(str(saved_paths[0]), mimetype="application/pdf",
                                  as_attachment=True, download_name=f"{saved_paths[0].stem}.pdf")
            response.call_on_close(lambda: cleanup_dir(workdir))
            return response

        cleanup_dir(workdir)
        return jsonify({"error": "Please upload JPG/PNG images only, or a single PDF file."}), 400

    except Exception as exc:
        logger.exception("Conversion failed")
        cleanup_dir(workdir)
        return jsonify({"error": str(exc)}), 500


if __name__ == "__main__":
    from waitress import serve
    serve(app, host="0.0.0.0", port=5000, threads=8)
