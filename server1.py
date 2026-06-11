"""
Word (DOCX) -> PDF converter — pure Python, no LibreOffice required.
Uses python-docx to read the DOCX and reportlab to write the PDF.
Works on Render free plan (512 MB RAM).
"""

from __future__ import annotations

import io
import os
import re
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors as rl_colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Table, TableStyle, PageBreak,
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY

app = Flask(__name__, static_folder=".", static_url_path="")
CORS(app)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024

_ALIGN_MAP = {
    WD_ALIGN_PARAGRAPH.LEFT:    TA_LEFT,
    WD_ALIGN_PARAGRAPH.CENTER:  TA_CENTER,
    WD_ALIGN_PARAGRAPH.RIGHT:   TA_RIGHT,
    WD_ALIGN_PARAGRAPH.JUSTIFY: TA_JUSTIFY,
    None:                        TA_LEFT,
}


def _hex(rgb: RGBColor | None) -> str | None:
    if rgb is None:
        return None
    return "#{:02x}{:02x}{:02x}".format(rgb.red, rgb.green, rgb.blue)


def _pt(val, default: float = 11.0) -> float:
    if val is None:
        return default
    try:
        return float(val.pt)
    except Exception:
        return default


def _has_page_break(para) -> bool:
    """True only for explicit page-break runs (<w:br w:type='page'/>)."""
    for br in para._element.findall(".//" + qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return True
    return False


def _para_markup(para, base_size: float = 11.0) -> str:
    """Convert a docx paragraph's runs to ReportLab XML markup."""
    parts: list[str] = []
    for run in para.runs:
        text = (run.text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        if not text:
            continue
        size  = _pt(run.font.size, base_size)
        color = _hex(run.font.color.rgb if (run.font.color and run.font.color.type) else None)

        open_tags:  list[str] = [f'<font size="{size:.1f}"' + (f' color="{color}"' if color else "") + ">"]
        close_tags: list[str] = ["</font>"]

        if run.bold:
            open_tags.append("<b>");  close_tags.insert(0, "</b>")
        if run.italic:
            open_tags.append("<i>");  close_tags.insert(0, "</i>")
        if run.underline:
            open_tags.append("<u>");  close_tags.insert(0, "</u>")

        parts.append("".join(open_tags) + text + "".join(close_tags))
    return "".join(parts)


def _safe_para(markup: str, style: ParagraphStyle) -> Paragraph:
    """Return a Paragraph; fall back to plain text if markup is broken."""
    try:
        return Paragraph(markup, style)
    except Exception:
        plain = re.sub(r"<[^>]+>", "", markup)
        plain = plain.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
        return Paragraph(plain, style)


def docx_to_pdf_bytes(docx_bytes: bytes) -> bytes:
    doc = Document(io.BytesIO(docx_bytes))

    buf = io.BytesIO()
    pdf = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2.5 * cm, rightMargin=2.5 * cm,
        topMargin=2.5 * cm,  bottomMargin=2.5 * cm,
    )

    base_style = ParagraphStyle("base", fontSize=11, leading=14, spaceAfter=4)
    story: list = []

    # ── Paragraphs ────────────────────────────────────────────────────────────
    for para in doc.paragraphs:
        if _has_page_break(para):
            story.append(PageBreak())

        markup = _para_markup(para)
        if not markup.strip():
            story.append(Spacer(1, 6))
            continue

        align      = _ALIGN_MAP.get(para.alignment, TA_LEFT)
        style_name = (para.style.name or "").lower()

        if "heading 1" in style_name:
            ps = ParagraphStyle("h1", fontSize=18, leading=22, bold=True,
                                spaceBefore=10, spaceAfter=8,  alignment=align)
        elif "heading 2" in style_name:
            ps = ParagraphStyle("h2", fontSize=15, leading=19, bold=True,
                                spaceBefore=8,  spaceAfter=6,  alignment=align)
        elif "heading 3" in style_name:
            ps = ParagraphStyle("h3", fontSize=13, leading=17, bold=True,
                                spaceBefore=6,  spaceAfter=4,  alignment=align)
        else:
            sz = _pt(para.style.font.size if para.style else None)
            ps = ParagraphStyle("body", fontSize=sz, leading=sz * 1.3,
                                spaceAfter=4, alignment=align)

        story.append(_safe_para(markup, ps))

    # ── Tables ────────────────────────────────────────────────────────────────
    for table in doc.tables:
        rows = []
        for row in table.rows:
            rows.append([
                _safe_para(
                    cell.text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"),
                    base_style,
                )
                for cell in row.cells
            ])
        if not rows:
            continue

        col_n  = max(len(r) for r in rows)
        col_w  = (A4[0] - 5 * cm) / col_n
        tbl    = Table(rows, colWidths=[col_w] * col_n)
        tbl.setStyle(TableStyle([
            ("GRID",          (0, 0), (-1, -1), 0.5, rl_colors.grey),
            ("BACKGROUND",    (0, 0), (-1,  0), rl_colors.HexColor("#DBEAFE")),
            ("FONTSIZE",      (0, 0), (-1, -1), 9),
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.extend([Spacer(1, 8), tbl, Spacer(1, 8)])

    if not story:
        story.append(Paragraph("(empty document)", base_style))

    pdf.build(story)
    buf.seek(0)
    return buf.read()


# ── Routes ────────────────────────────────────────────────────────────────────

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
    if Path(upload.filename).suffix.lower() != ".docx":
        return jsonify(error="Only .docx files are accepted."), 400

    stem = Path(secure_filename(upload.filename)).stem
    try:
        pdf_bytes = docx_to_pdf_bytes(upload.read())
        buf = io.BytesIO(pdf_bytes)
        resp = send_file(buf, as_attachment=True,
                         download_name=f"{stem}.pdf",
                         mimetype="application/pdf",
                         max_age=0)
        resp.headers["Cache-Control"] = "no-store"
        return resp
    except Exception as exc:
        return jsonify(error=str(exc)), 500


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
