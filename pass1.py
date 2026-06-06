from __future__ import annotations

import io
import os
from io import BytesIO
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from pypdf import PdfReader, PdfWriter
from pypdf.errors import WrongPasswordError
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)

app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_UPLOAD_MB", "50")) * 1024 * 1024
app.config["JSON_SORT_KEYS"] = False


def is_pdf(filename: str) -> bool:
    return Path(filename.lower()).suffix == ".pdf"


def clean_metadata(metadata):
    if not metadata:
        return None

    cleaned = {}
    for key, value in metadata.items():
        if key is None or value is None:
            continue
        cleaned[str(key)] = str(value)

    return cleaned or None


@app.get("/")
def home():
    return jsonify({"status": "running", "tool": "pdf unlock"}), 200


@app.get("/health")
def health():
    return jsonify({"ok": True}), 200


@app.post("/unlock")
def unlock_pdf():
    uploaded_file = request.files.get("file")
    password = (request.form.get("password") or "").strip()

    if not uploaded_file or not uploaded_file.filename:
        return jsonify({"error": "No PDF file uploaded."}), 400

    if not is_pdf(uploaded_file.filename):
        return jsonify({"error": "Only PDF files are allowed."}), 400

    if not password:
        return jsonify({"error": "Please enter the PDF password."}), 400

    try:
        pdf_bytes = uploaded_file.read()
        if not pdf_bytes:
            return jsonify({"error": "Uploaded file is empty."}), 400

        reader = PdfReader(BytesIO(pdf_bytes))

        if reader.is_encrypted:
            try:
                result = reader.decrypt(password)
            except WrongPasswordError:
                return jsonify({"error": "Wrong password or unsupported encryption."}), 400
            except Exception as exc:
                return jsonify({"error": f"Could not decrypt PDF: {str(exc)}"}), 400

            if result == 0:
                return jsonify({"error": "Wrong password or unsupported encryption."}), 400

        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        try:
            metadata = clean_metadata(reader.metadata)
            if metadata:
                writer.add_metadata(metadata)
        except Exception:
            pass

        output = io.BytesIO()
        writer.write(output)
        output.seek(0)

        original_name = secure_filename(uploaded_file.filename)
        stem = Path(original_name).stem
        output_name = f"{stem}_unlocked.pdf"

        response = send_file(
            output,
            as_attachment=True,
            download_name=output_name,
            mimetype="application/pdf",
            conditional=False,
            max_age=0,
        )
        response.headers["Cache-Control"] = "no-store"
        return response

    except Exception as exc:
        return jsonify({"error": f"Unlock failed: {str(exc)}"}), 500


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)