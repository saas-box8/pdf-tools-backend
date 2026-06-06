from __future__ import annotations

import io
import secrets
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from pypdf import PdfReader, PdfWriter
from pypdf.constants import UserAccessPermissions
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

app = Flask(__name__)
CORS(app)

app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB


def build_permissions(
    disable_print: bool,
    disable_edit: bool,
    disable_copy: bool,
) -> UserAccessPermissions:
    permissions = (
        UserAccessPermissions.PRINT
        | UserAccessPermissions.PRINT_TO_REPRESENTATION
        | UserAccessPermissions.MODIFY
        | UserAccessPermissions.ADD_OR_MODIFY
        | UserAccessPermissions.ASSEMBLE_DOC
        | UserAccessPermissions.FILL_FORM_FIELDS
        | UserAccessPermissions.EXTRACT
        | UserAccessPermissions.EXTRACT_TEXT_AND_GRAPHICS
    )

    if disable_print:
        permissions &= ~UserAccessPermissions.PRINT
        permissions &= ~UserAccessPermissions.PRINT_TO_REPRESENTATION

    if disable_edit:
        permissions &= ~UserAccessPermissions.MODIFY
        permissions &= ~UserAccessPermissions.ADD_OR_MODIFY
        permissions &= ~UserAccessPermissions.ASSEMBLE_DOC
        permissions &= ~UserAccessPermissions.FILL_FORM_FIELDS

    if disable_copy:
        permissions &= ~UserAccessPermissions.EXTRACT
        permissions &= ~UserAccessPermissions.EXTRACT_TEXT_AND_GRAPHICS

    return permissions


def create_footer_overlay(page_width: float, page_height: float, footer_text: str) -> io.BytesIO:
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))

    text = (footer_text or "Protected PDF").strip()
    font_name = "Helvetica"
    font_size = 8

    text_width = stringWidth(text, font_name, font_size)
    x = max((page_width - text_width) / 2, 10)
    y = 12

    c.setFont(font_name, font_size)
    c.setFillColorRGB(0.45, 0.45, 0.45)
    c.drawString(x, y, text)
    c.showPage()
    c.save()

    packet.seek(0)
    return packet


def add_footer_to_writer(reader: PdfReader, footer_text: str) -> PdfWriter:
    writer = PdfWriter()

    for page in reader.pages:
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)

        overlay_stream = create_footer_overlay(page_width, page_height, footer_text)
        overlay_reader = PdfReader(overlay_stream)
        overlay_page = overlay_reader.pages[0]

        page.merge_page(overlay_page)
        writer.add_page(page)

    return writer


@app.get("/")
def home():
    return jsonify({"message": "Protect PDF backend is running"})


@app.get("/health")
def health():
    return jsonify({"ok": True})


@app.post("/protect")
def protect_pdf():
    pdf_file = request.files.get("file")
    password = (request.form.get("password") or "").strip()
    confirm_password = (request.form.get("confirm_password") or "").strip()
    owner_note = (request.form.get("owner_note") or "Protected by PDF Tool").strip()

    disable_print = request.form.get("disable_print") == "true"
    disable_edit = request.form.get("disable_edit") == "true"
    disable_copy = request.form.get("disable_copy") == "true"

    if not pdf_file or not pdf_file.filename:
        return jsonify({"error": "No PDF file uploaded."}), 400

    if not pdf_file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF files are allowed."}), 400

    if not password:
        return jsonify({"error": "Password is required."}), 400

    if password != confirm_password:
        return jsonify({"error": "Password and confirm password do not match."}), 400

    try:
        reader = PdfReader(pdf_file.stream)

        if reader.is_encrypted:
            return jsonify({"error": "This PDF is already encrypted."}), 400

        writer = add_footer_to_writer(reader, owner_note)
        permissions = build_permissions(disable_print, disable_edit, disable_copy)

        owner_password = secrets.token_urlsafe(16)

        writer.encrypt(
            user_password=password,
            owner_password=owner_password,
            permissions_flag=permissions,
            algorithm="AES-256",
        )

        output = io.BytesIO()
        writer.write(output)
        output.seek(0)

        safe_name = f"{Path(pdf_file.filename).stem}_protected.pdf"

        response = send_file(
            output,
            as_attachment=True,
            download_name=safe_name,
            mimetype="application/pdf",
            conditional=False,
            max_age=0,
        )
        response.headers["Cache-Control"] = "no-store"
        return response

    except Exception as exc:
        return jsonify({"error": f"Failed to protect PDF: {str(exc)}"}), 500


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)