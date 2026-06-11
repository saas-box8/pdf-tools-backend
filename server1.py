"""
Word (DOCX) → PDF converter — pure Python, no LibreOffice required.
Uses python-docx to read the DOCX and reportlab to write the PDF.
Works on Render free plan.
"""

from __future__ import annotations

import io
import os
import re
import shutil
import tempfile
import uuid
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

# python-docx
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn  # FIX: needed for correct XML element detection

# reportlab
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm  # FIX: removed invalid 'pt' import (not in reportlab.lib.units)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors as rl_colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table,
    TableStyle, PageBreak, HRFlowable,
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

app = Flask(__name__, static_folder=".", static_url_path="")
CORS(app)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024

ALLOWED_EXTENSIONS = {".docx"}


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


# ── Alignment map ─────────────────────────────────────────────────────────────
_ALIGN_MAP = {
    WD_ALIGN_PARAGRAPH.LEFT: TA_LEFT,
    WD_ALIGN_PARAGRAPH.CENTER: TA_CENTER,
    WD_ALIGN_PARAGRAPH.RIGHT: TA_RIGHT,
    WD_ALIGN_PARAGRAPH.JUSTIFY: TA_JUSTIFY,
    None: TA_LEFT,
}


def _hex_color(rgb: RGBColor | None):
    if rgb is None:
        return None
    return "#{:02x}{:02x}{:02x}".format(rgb.red, rgb.green, rgb.blue)


def _pt(val) -> float:
    """Return point value from a docx Pt/Length, or default."""
    if val is None:
        return 11.0
    try:
        return val.pt
    except Exception:
        return 11.0


def _has_page_break(para) -> bool:
    """
    FIX: detect only true page breaks (<w:br w:type='page'/>), not soft line
    breaks (<w:br/> or <w:br w:type='textWrapping'/>).  The old string-based
    search was also matching 'w:lastRenderedPageBreak' which is a hint added
    by Word's renderer, not an actual break instruction.
    """
    for br in para._element.findall(".//" + qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return True
    return False


def _build_para_text(para, base_size: float = 11.0) -> str:
    """Build reportlab-compatible XML markup from a docx paragraph."""
    parts: list[str] = []
    for run in para.runs:
        text = run.text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        if not text:
            continue
        tags_open: list[str] = []
        tags_close: list[str] = []
        size = _pt(run.font.size) if run.font.size else base_size
        color = _hex_color(run.font.color.rgb if run.font.color and run.font.color.type else None)
        if run.bold:
            tags_open.append("<b>"); tags_close.insert(0, "</b>")
        if run.italic:
            tags_open.append("<i>"); tags_close.insert(0, "</i>")
        if run.underline:
            tags_open.append("<u>"); tags_close.insert(0, "</u>")
        font_tag = f'<font size="{size:.1f}"'
        if color:
            font_tag += f' color="{color}"'
        font_tag += ">"
        tags_open.insert(0, font_tag)
        tags_close.append("</font>")
        parts.append("".join(tags_open) + text + "".join(tags_close))
    return "".join(parts)


def docx_to_pdf_bytes(docx_bytes: bytes) -> bytes:
    doc = Document(io.BytesIO(docx_bytes))

    buf = io.BytesIO()
    pdf = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=2.5 * cm,
        rightMargin=2.5 * cm,
        topMargin=2.5 * cm,
        bottomMargin=2.5 * cm,
    )

    story = []
    base_style = ParagraphStyle("base", fontSize=11, leading=14, spaceAfter=4)

    for para in doc.paragraphs:
        # FIX: only insert a PageBreak for real page-break instructions,
        # not for every paragraph that happens to contain a soft line break.
        if _has_page_break(para):
            story.append(PageBreak())
            # Don't 'continue' — the paragraph may also have text before the break
            # so fall through and render its text content normally.

        raw = _build_para_text(para)
        if not raw.strip():
            story.append(Spacer(1, 6))
            continue

        align = _ALIGN_MAP.get(para.alignment, TA_LEFT)

        # Detect heading by style name
        style_name = (para.style.name or "").lower()
        if "heading 1" in style_name:
            ps = ParagraphStyle("h1", fontSize=18, leading=22, bold=True,
                                spaceAfter=8, spaceBefore=10, alignment=align)
        elif "heading 2" in style_name:
            ps = ParagraphStyle("h2", fontSize=15, leading=19, bold=True,
                                spaceAfter=6, spaceBefore=8, alignment=align)
        elif "heading 3" in style_name:
            ps = ParagraphStyle("h3", fontSize=13, leading=17, bold=True,
                                spaceAfter=4, spaceBefore=6, alignment=align)
        else:
            base_sz = _pt(para.style.font.size) if para.style.font.size else 11.0
            ps = ParagraphStyle("body", fontSize=base_sz, leading=base_sz * 1.3,
                                spaceAfter=4, alignment=align)

        try:
            story.append(Paragraph(raw, ps))
        except Exception:
            # Strip markup on failure
            plain = re.sub(r"<[^>]+>", "", raw)
            story.append(Paragraph(plain.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">"), ps))

    # Tables
    for table in doc.tables:
        data = []
        for row in table.rows:
            data.append([
                Paragraph(
                    cell.text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
                    base_style
                )
                for cell in row.cells
            ])
        if data:
            col_count = max(len(r) for r in data)
            avail = A4[0] - 5 * cm
            col_w = avail / col_count
            t = Table(data, colWidths=[col_w] * col_count)
            t.setStyle(TableStyle([
                ("GRID", (0, 0), (-1, -1), 0.5, rl_colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor("#DBEAFE")),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]))
            story.append(Spacer(1, 8))
            story.append(t)
            story.append(Spacer(1, 8))

    if not story:
        story.append(Paragraph("(empty document)", base_style))

    pdf.build(story)
    buf.seek(0)
    return buf.read()


@app.get("/")
def home():
    return jsonify({"status": "running", "tool": "word-to-pdf"})


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/convert", methods=["POST"])
def convert_word():
    upload = request.files.get("file")
    if not upload:
        return jsonify(error="No file uploaded"), 400
    if not upload.filename:
        return jsonify(error="No selected file"), 400
    if not allowed_file(upload.filename):
        return jsonify(error="Only DOCX files allowed."), 400

    filename = secure_filename(upload.filename)
    try:
        docx_bytes = upload.read()
        pdf_bytes = docx_to_pdf_bytes(docx_bytes)
        buf = io.BytesIO(pdf_bytes)
        response = send_file(
            buf,
            as_attachment=True,
            download_name=f"{Path(filename).stem}.pdf",
            mimetype="application/pdf",
            max_age=0,
        )
        response.headers["Cache-Control"] = "no-store"
        return response
    except Exception as exc:
        return jsonify(error=str(exc)), 500


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
