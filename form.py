"""
PDF Forms Filler — form.py

Install:
    pip install flask flask-cors pymupdf waitress

Run (production):
    waitress-serve --host=0.0.0.0 --port=5000 form:app

Run (development):
    python form.py
"""

import base64
import io
import json
import re
import threading
import time
import uuid

import fitz
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS

app = Flask(__name__)

# CORS for local / same-origin frontend
CORS(
    app,
    resources={r"/*": {"origins": "*"}},
    methods=["GET", "POST", "OPTIONS"],
    allow_headers=["Content-Type", "Authorization"],
)

RENDER_SCALE = 2.0
TTL_SECONDS = 30 * 60

_store: dict = {}
_store_lock = threading.Lock()


# ── CORS after-request hook ──────────────────────────────────────────────────
@app.after_request
def add_cors_headers(resp):
    origin = request.headers.get("Origin")
    if origin:
        resp.headers["Access-Control-Allow-Origin"] = origin
        resp.headers["Vary"] = "Origin"
    else:
        resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    resp.headers["Access-Control-Max-Age"] = "86400"
    return resp


# ── PyMuPDF widget-type constants ─────────────────────────────────────────────
def _raw_const(name, default):
    v = getattr(fitz, name, None)
    return v if v is not None else default


_WT_TEXT     = _raw_const("PDF_WIDGET_TYPE_TEXT",     4)
_WT_CHECKBOX = _raw_const("PDF_WIDGET_TYPE_CHECKBOX", 2)
_WT_COMBOBOX = _raw_const("PDF_WIDGET_TYPE_COMBOBOX", 6)
_WT_LISTBOX  = _raw_const("PDF_WIDGET_TYPE_LISTBOX",  5)

_FTYPE_TO_PDF_FT = {
    "text":      "Tx",
    "multiline": "Tx",
    "checkbox":  "Btn",
    "radio":     "Ch",
    "combobox":  "Ch",
    "listbox":   "Ch",
    "signature": "Tx",
    "date":      "Tx",
}

_FTYPE_TO_WIDGET_TYPE = {
    "text":      _WT_TEXT,
    "multiline": _WT_TEXT,
    "checkbox":  _WT_CHECKBOX,
    "radio":     _WT_COMBOBOX,
    "combobox":  _WT_COMBOBOX,
    "listbox":   _WT_LISTBOX,
    "signature": _WT_TEXT,
    "date":      _WT_TEXT,
}

_FTYPE_TO_FF = {
    "text":      0,
    "date":      0,
    "signature": 0,
    "checkbox":  0,
    "multiline": 4096,
    "combobox":  131072,
    "listbox":   0,
    "radio":     32768,
}

print(f"[form.py] PyMuPDF version: {getattr(fitz, 'version', '?')}", flush=True)
print(
    f"[form.py] WT TEXT={_WT_TEXT} CHECKBOX={_WT_CHECKBOX} "
    f"COMBOBOX={_WT_COMBOBOX} LISTBOX={_WT_LISTBOX}",
    flush=True,
)


# ── TTL cleanup thread ────────────────────────────────────────────────────────
def _cleanup_loop():
    while True:
        time.sleep(60)
        cutoff = time.time() - TTL_SECONDS
        with _store_lock:
            expired = [t for t, v in _store.items() if v["last_access"] < cutoff]
            for t in expired:
                _store.pop(t, None)


threading.Thread(target=_cleanup_loop, daemon=True).start()


# ── In-memory store helpers ──────────────────────────────────────────────────
def _store_get(token: str) -> bytes | None:
    with _store_lock:
        entry = _store.get(token)
        if entry is None:
            return None
        entry["last_access"] = time.time()
        return bytes(entry["doc_bytes"])


def _store_put(token: str, doc_bytes: bytes) -> None:
    with _store_lock:
        _store[token] = {
            "doc_bytes": bytearray(doc_bytes),
            "last_access": time.time(),
        }


# ── Page rendering ────────────────────────────────────────────────────────────
def _render_page(doc: fitz.Document, page_num: int) -> dict:
    page = doc[page_num]
    mat  = fitz.Matrix(RENDER_SCALE, RENDER_SCALE)
    pix  = page.get_pixmap(matrix=mat, alpha=False)
    return {
        "img": base64.b64encode(pix.tobytes("png")).decode(),
        "iw":  pix.width,
        "ih":  pix.height,
        "pw":  page.rect.width,
        "ph":  page.rect.height,
    }


# ── Widget-type resolution ────────────────────────────────────────────────────
_TYPE_MAP: dict = {}
for _ta, _tl in [
    ("PDF_WIDGET_TYPE_TEXT",        "text"),
    ("PDF_WIDGET_TYPE_CHECKBOX",    "checkbox"),
    ("PDF_WIDGET_TYPE_RADIOBUTTON", "radio"),
    ("PDF_WIDGET_TYPE_COMBOBOX",    "combobox"),
    ("PDF_WIDGET_TYPE_LISTBOX",     "listbox"),
    ("PDF_WIDGET_TYPE_SIGNATURE",   "signature"),
    ("PDF_WIDGET_TYPE_BUTTON",      "button"),
]:
    _tv = getattr(fitz, _ta, None)
    if _tv is not None:
        _TYPE_MAP[_tv] = _tl


def _wtype(w) -> str:
    wt = getattr(w, "widget_type", None)
    if wt is not None:
        return _TYPE_MAP.get(wt, "text")
    ft = getattr(w, "field_type", None)
    if isinstance(ft, str):
        return {"Tx": "text", "Btn": "checkbox", "Ch": "combobox", "Sig": "signature"}.get(ft, "text")
    return "text"


# ── Field parsing ─────────────────────────────────────────────────────────────
def _parse_fields(doc: fitz.Document) -> list:
    fields: list = []
    seen: dict   = {}

    for pn in range(len(doc)):
        page    = doc[pn]
        widgets = page.widgets() or []

        for w in widgets:
            wt = _wtype(w)
            if wt == "button":
                continue

            raw = (w.field_name or "").strip()

            if raw in seen:
                seen[raw] += 1
                uname = f"{raw}_{seen[raw]}"
            else:
                seen[raw] = 0
                uname = raw

            val = w.field_value
            if val is None:
                val = ""
            elif isinstance(val, bytes):
                val = val.decode("utf-8", errors="replace")
            elif isinstance(val, bool):
                pass  # keep bool
            else:
                val = str(val)

            opts: list = []
            if getattr(w, "choice_values", None):
                for o in w.choice_values:
                    opts.append(o.decode("utf-8", errors="replace") if isinstance(o, bytes) else str(o))

            flags    = w.field_flags or 0
            fl_multi = getattr(fitz, "PDF_FIELD_IS_MULTILINE", 0)
            fl_req   = getattr(fitz, "PDF_FIELD_IS_REQUIRED",  0)
            multiline = bool(flags & fl_multi) if fl_multi else False
            required  = bool(flags & fl_req)   if fl_req  else False

            on_state = ""
            if wt == "checkbox":
                try:
                    on_state = str(w.on_state() or "Yes")
                except Exception:
                    on_state = "Yes"

            r = w.rect
            fields.append({
                "page":      pn,
                "name":      uname,
                "type":      wt,
                "rect":      [round(r.x0, 2), round(r.y0, 2), round(r.x1, 2), round(r.y1, 2)],
                "value":     val,
                "options":   opts,
                "multiline": multiline,
                "required":  required,
                "on_state":  on_state,
                "manual":    False,
            })

    return fields


# ── Widget creation helpers ──────────────────────────────────────────────────
def _apply_style(w, ftype: str, options: list) -> None:
    if ftype == "multiline":
        try:
            w.field_flags = getattr(fitz, "PDF_FIELD_IS_MULTILINE", 4096)
        except Exception:
            pass

    if ftype in ("radio", "combobox", "listbox"):
        try:
            w.choice_values = options if options else (["Yes", "No"] if ftype == "radio" else ["Option 1", "Option 2"])
        except Exception:
            pass

    for attr, val in [
        ("fill_color",   (0.93, 0.93, 1.0) if ftype == "signature" else (1.0, 1.0, 1.0)),
        ("border_color", (0.41, 0.35, 1.0)),
        ("text_color",   (0.0,  0.0,  0.0)),
        ("border_width", 1),
        ("text_fontsize", 10),
    ]:
        try:
            setattr(w, attr, val)
        except Exception:
            pass


def _pdf_escape(s: str) -> str:
    return s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def _add_field_lowlevel(doc, page, ftype: str, name: str, rect, options: list) -> None:
    """Last-resort: write the widget annotation directly into the PDF object stream."""
    ft  = _FTYPE_TO_PDF_FT.get(ftype, "Tx")
    ff  = _FTYPE_TO_FF.get(ftype, 0)
    ph  = page.rect.height
    px0, py0 = rect.x0, ph - rect.y1
    px1, py1 = rect.x1, ph - rect.y0

    rect_str  = "[%.4f %.4f %.4f %.4f]" % (px0, py0, px1, py1)
    safe_name = _pdf_escape(name)

    parts = [
        "/Type /Annot", "/Subtype /Widget",
        f"/FT /{ft}", f"/T ({safe_name})",
        f"/Rect {rect_str}", "/DA (/Helv 10 Tf 0 g)",
        f"/Ff {ff}", f"/P {page.xref} 0 R",
        "/MK << /BG [1 1 1] /BC [0.41 0.35 1] >>",
    ]
    if ftype == "checkbox":
        parts += ["/AS /Off", "/V /Off"]
    elif ftype != "signature":
        parts.append("/V ()")
    if options and ftype in ("combobox", "listbox", "radio"):
        opt_parts = [f"({_pdf_escape(o)})" for o in options]
        parts.append("/Opt [" + " ".join(opt_parts) + "]")

    annot_dict = "<< " + " ".join(parts) + " >>"
    annot_xref = doc.get_new_xref()
    doc.update_object(annot_xref, annot_dict)
    annot_ref  = f"{annot_xref} 0 R"

    page_obj = doc.xref_object(page.xref, compressed=False)
    if "/Annots" in page_obj:
        page_obj = re.sub(r"/Annots\s*\[", "/Annots [" + annot_ref + " ", page_obj, count=1)
    else:
        stripped = page_obj.rstrip()
        sep = "\n/Annots [" + annot_ref + "]\n>>"
        page_obj = (stripped[:-2].rstrip() + sep) if stripped.endswith(">>") else (page_obj + "\n/Annots [" + annot_ref + "]")
    doc.update_object(page.xref, page_obj)

    catalog_xref = doc.pdf_catalog()
    catalog_obj  = doc.xref_object(catalog_xref, compressed=False)
    m_ind        = re.search(r"/AcroForm\s+(\d+)\s+0\s+R", catalog_obj)

    if m_ind:
        af_xref = int(m_ind.group(1))
        af_obj  = doc.xref_object(af_xref, compressed=False)
        if "/Fields" in af_obj:
            af_obj = re.sub(r"/Fields\s*\[", "/Fields [" + annot_ref + " ", af_obj, count=1)
        else:
            stripped = af_obj.rstrip()
            af_obj = (stripped[:-2].rstrip() + "\n/Fields [" + annot_ref + "]\n>>") if stripped.endswith(">>") else (af_obj + "\n/Fields [" + annot_ref + "]")
        doc.update_object(af_xref, af_obj)

    elif "/AcroForm" in catalog_obj:
        if "/Fields" in catalog_obj:
            catalog_obj = re.sub(r"/Fields\s*\[", "/Fields [" + annot_ref + " ", catalog_obj, count=1)
            doc.update_object(catalog_xref, catalog_obj)
        else:
            m_inline = re.search(r"/AcroForm\s*(<<.*?>>)", catalog_obj, re.DOTALL)
            if m_inline:
                inner = m_inline.group(1).rstrip()
                inner = (inner[:-2].rstrip() + "\n/Fields [" + annot_ref + "]\n>>") if inner.endswith(">>") else (inner + "\n/Fields [" + annot_ref + "]")
                af_xref = doc.get_new_xref()
                doc.update_object(af_xref, inner)
                catalog_obj = catalog_obj.replace(m_inline.group(0), "/AcroForm " + str(af_xref) + " 0 R")
                doc.update_object(catalog_xref, catalog_obj)
    else:
        af_xref = doc.get_new_xref()
        doc.update_object(af_xref, "<< /Fields [" + annot_ref + "] /DA (/Helv 10 Tf 0 g) >>")
        stripped = catalog_obj.rstrip()
        catalog_obj = (stripped[:-2].rstrip() + "\n/AcroForm " + str(af_xref) + " 0 R\n>>") if stripped.endswith(">>") else (catalog_obj + "\n/AcroForm " + str(af_xref) + " 0 R")
        doc.update_object(catalog_xref, catalog_obj)


def _create_widget(doc, page, ftype: str, name: str, rect, options: list) -> None:
    wt_const = _FTYPE_TO_WIDGET_TYPE.get(ftype, _WT_TEXT)
    ft_str   = _FTYPE_TO_PDF_FT.get(ftype, "Tx")
    errors: list = []

    for attempt, setup in enumerate([
        lambda w: setattr(w, "field_type", ft_str),
        lambda w: setattr(w, "widget_type", wt_const),
        None,  # constructor variant
    ]):
        try:
            w = fitz.Widget(wt_const) if attempt == 2 else fitz.Widget()
            if setup:
                setup(w)
            w.field_name = name
            w.rect       = rect
            _apply_style(w, ftype, options)
            page.add_widget(w)
            print(f"[widget] S{attempt+1} OK", flush=True)
            return
        except Exception as e:
            errors.append(f"S{attempt+1}:{e}")

    # Try raw int types
    for iv in [4, 2, 6, 5]:
        try:
            w = fitz.Widget()
            w.widget_type = iv
            w.field_name  = name
            w.rect        = rect
            page.add_widget(w)
            print(f"[widget] S4 OK int={iv}", flush=True)
            return
        except Exception as e:
            errors.append(f"S4(int={iv}):{e}")

    # Low-level fallback
    try:
        _add_field_lowlevel(doc, page, ftype, name, rect, options)
        print("[widget] S5 OK low-level PDF", flush=True)
        return
    except Exception as e:
        errors.append(f"S5:{e}")

    raise RuntimeError("All strategies failed: " + " | ".join(errors))


def _delete_widget(page, widget) -> bool:
    for method in [
        lambda: page.delete_widget(widget) if hasattr(page, "delete_widget") else (_ for _ in ()).throw(AttributeError()),
        lambda: page.delete_annot(widget),
    ]:
        try:
            method()
            return True
        except Exception:
            pass
    return False


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/debug", methods=["GET", "OPTIONS"])
def debug_info():
    if request.method == "OPTIONS":
        return ("", 204)
    info = {
        "pymupdf_version": str(getattr(fitz, "version", "unknown")),
        "resolved_types":  {k: str(v) for k, v in {"TEXT": _WT_TEXT, "CHECKBOX": _WT_CHECKBOX, "COMBOBOX": _WT_COMBOBOX, "LISTBOX": _WT_LISTBOX}.items()},
        "all_constants":   {attr: str(getattr(fitz, attr, None)) for attr in sorted(dir(fitz)) if "WIDGET" in attr or "FIELD_IS" in attr},
    }
    return jsonify(info)


@app.route("/upload", methods=["POST", "OPTIONS"])
def upload():
    if request.method == "OPTIONS":
        return ("", 204)
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file part"}), 400
        raw = request.files["file"].read()
        if not raw:
            return jsonify({"error": "Empty file"}), 400

        doc = fitz.open(stream=raw, filetype="pdf")
        buf = io.BytesIO()
        doc.save(buf, garbage=4, deflate=True)
        doc.close()
        doc_bytes = buf.getvalue()

        token = str(uuid.uuid4())
        _store_put(token, doc_bytes)

        doc2   = fitz.open(stream=doc_bytes, filetype="pdf")
        fields = _parse_fields(doc2)
        page0  = _render_page(doc2, 0)
        pages  = len(doc2)
        doc2.close()

        return jsonify({
            "token":        token,
            "pages":        pages,
            "fields":       fields,
            "field_count":  len(fields),
            "render_scale": RENDER_SCALE,
            "page":         page0,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/render/<token>/<int:page_num>", methods=["GET", "OPTIONS"])
def render(token, page_num):
    if request.method == "OPTIONS":
        return ("", 204)
    try:
        doc_bytes = _store_get(token)
        if doc_bytes is None:
            return jsonify({"error": "Token expired"}), 404
        doc = fitz.open(stream=doc_bytes, filetype="pdf")
        if not (0 <= page_num < len(doc)):
            doc.close()
            return jsonify({"error": "Page out of range"}), 400
        result = _render_page(doc, page_num)
        doc.close()
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/add-field/<token>", methods=["POST", "OPTIONS"])
def add_field(token):
    if request.method == "OPTIONS":
        return ("", 204)
    try:
        doc_bytes = _store_get(token)
        if doc_bytes is None:
            return jsonify({"error": "Token expired"}), 404

        body     = request.get_json(force=True)
        page_num = int(body.get("page", 0))
        name     = (body.get("name") or "").strip()
        ftype    = (body.get("type", "text") or "text").strip().lower()
        rect_raw = body.get("rect", [50, 700, 250, 720])
        options  = [str(o) for o in body.get("options", [])]

        doc = fitz.open(stream=doc_bytes, filetype="pdf")
        if not (0 <= page_num < len(doc)):
            doc.close()
            return jsonify({"error": "Page out of range"}), 400

        if not name:
            existing = {w.field_name for pg in doc for w in (pg.widgets() or []) if w.field_name}
            idx, candidate = 1, f"{ftype}_1"
            while candidate in existing:
                idx += 1
                candidate = f"{ftype}_{idx}"
            name = candidate

        rect = fitz.Rect(float(rect_raw[0]), float(rect_raw[1]), float(rect_raw[2]), float(rect_raw[3]))
        page = doc[page_num]

        try:
            _create_widget(doc, page, ftype, name, rect, options)
        except RuntimeError as add_err:
            doc.close()
            return jsonify({"error": str(add_err)}), 500

        buf = io.BytesIO()
        doc.save(buf, garbage=4, deflate=True)
        new_bytes = buf.getvalue()
        doc.close()
        _store_put(token, new_bytes)

        doc2   = fitz.open(stream=new_bytes, filetype="pdf")
        fields = _parse_fields(doc2)
        page_d = _render_page(doc2, page_num)
        doc2.close()

        return jsonify({"ok": True, "name": name, "fields": fields, "field_count": len(fields), "page": page_d})
    except Exception as e:
        print(f"[add-field] unhandled: {e}", flush=True)
        return jsonify({"error": str(e)}), 500


@app.route("/update-field/<token>", methods=["POST", "OPTIONS"])
def update_field(token):
    if request.method == "OPTIONS":
        return ("", 204)
    try:
        doc_bytes = _store_get(token)
        if doc_bytes is None:
            return jsonify({"error": "Token expired"}), 404

        body     = request.get_json(force=True)
        page_num = int(body.get("page", 0))
        name     = (body.get("name") or "").strip()
        rect_raw = body.get("rect")

        if not name or not rect_raw:
            return jsonify({"error": "Missing name or rect"}), 400

        doc = fitz.open(stream=doc_bytes, filetype="pdf")
        if not (0 <= page_num < len(doc)):
            doc.close()
            return jsonify({"error": "Page out of range"}), 400

        page   = doc[page_num]
        target = next((w for w in (page.widgets() or []) if (w.field_name or "") == name), None)

        if target is None:
            doc.close()
            return jsonify({"error": "Field not found"}), 404

        new_rect = fitz.Rect(float(rect_raw[0]), float(rect_raw[1]), float(rect_raw[2]), float(rect_raw[3]))
        try:
            target.rect = new_rect
            target.update()
        except Exception as e:
            doc.close()
            return jsonify({"error": f"Failed to update widget rect: {e}"}), 500

        buf = io.BytesIO()
        doc.save(buf, garbage=4, deflate=True)
        new_bytes = buf.getvalue()
        doc.close()
        _store_put(token, new_bytes)

        doc2 = fitz.open(stream=new_bytes, filetype="pdf")
        doc2.reload_page(doc2[page_num])
        fields = _parse_fields(doc2)
        page_d = _render_page(doc2, page_num)
        doc2.close()

        return jsonify({"ok": True, "fields": fields, "page": page_d})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/remove-field/<token>", methods=["POST", "OPTIONS"])
def remove_field(token):
    if request.method == "OPTIONS":
        return ("", 204)
    try:
        doc_bytes = _store_get(token)
        if doc_bytes is None:
            return jsonify({"error": "Token expired"}), 404

        body     = request.get_json(force=True)
        page_num = int(body.get("page", 0))
        name     = body.get("name", "")

        doc  = fitz.open(stream=doc_bytes, filetype="pdf")
        page = doc[page_num]

        for w in list(page.widgets() or []):
            if (w.field_name or "") == name:
                _delete_widget(page, w)
                break

        buf = io.BytesIO()
        doc.save(buf, garbage=4, deflate=True)
        new_bytes = buf.getvalue()
        doc.close()
        _store_put(token, new_bytes)

        doc2   = fitz.open(stream=new_bytes, filetype="pdf")
        fields = _parse_fields(doc2)
        page_d = _render_page(doc2, page_num)
        doc2.close()

        return jsonify({"ok": True, "fields": fields, "page": page_d})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/fill/<token>", methods=["POST", "OPTIONS"])
def fill(token):
    if request.method == "OPTIONS":
        return ("", 204)
    try:
        doc_bytes = _store_get(token)
        if doc_bytes is None:
            return jsonify({"error": "Token expired"}), 404

        if request.is_json:
            body       = request.get_json(force=True)
            fields_raw = body.get("fields", {})
            flatten    = bool(body.get("flatten", False))
        else:
            fields_raw = json.loads(request.form.get("fields", "{}"))
            flatten    = request.form.get("flatten", "false").lower() in ("1", "true", "yes")

        doc = fitz.open(stream=doc_bytes, filetype="pdf")

        for page in doc:
            for widget in page.widgets() or []:
                fname = widget.field_name or ""
                if fname not in fields_raw:
                    continue
                val   = fields_raw[fname]
                wtype = _wtype(widget)
                try:
                    if wtype == "checkbox":
                        if isinstance(val, bool):
                            widget.field_value = val
                        elif isinstance(val, str):
                            widget.field_value = val.lower() in ("true", "yes", "1", "on")
                        else:
                            widget.field_value = bool(val)
                    else:
                        widget.field_value = str(val)
                    widget.update()
                except Exception:
                    pass

        if flatten:
            try:
                if hasattr(doc, "bake"):
                    doc.bake()
            except Exception:
                pass
            for pg in doc:
                for w in list(pg.widgets() or []):
                    try:
                        pg.delete_widget(w)
                    except Exception:
                        try:
                            pg.delete_annot(w)
                        except Exception:
                            pass
                try:
                    pg.clean_contents()
                except Exception:
                    pass

        out    = io.BytesIO()
        doc.save(out, garbage=4, deflate=True, incremental=False)
        filled = out.getvalue()
        doc.close()
        _store_put(token, filled)

        return send_file(
            io.BytesIO(filled),
            mimetype="application/pdf",
            as_attachment=True,
            download_name="filled.pdf",
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── Entrypoint ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        from waitress import serve
        print("[form.py] Starting production server via waitress on 0.0.0.0:5000", flush=True)
        serve(app, host="0.0.0.0", port=5000, threads=8)
    except ImportError:
        print("[form.py] waitress not found — falling back to Flask dev server", flush=True)
        app.run(host="0.0.0.0", port=5000, debug=False, use_reloader=False)