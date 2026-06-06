# Render Deployment Guide — Free Plan

All 13 backends run on the **Render free plan** with zero system packages.
Every dependency is a pure Python package installed via pip.

## What changed from the original code

| File | Change |
|------|--------|
| `server1.py` | Removed LibreOffice → now uses python-docx + reportlab |
| `app1.py` | Removed LibreOffice → now uses python-pptx + reportlab |
| `photo.py` | Removed LibreOffice office-file branch → PDF-only input |
| `photo1.py` | Removed LibreOffice office-file branch → images + PDF only |
| `backend.py` | Removed LibreOffice XLS export → now uses xlwt (pure Python) |
| `server.py` | Removed Tesseract + Poppler → now uses RapidOCR + PyMuPDF |
| `web.py` | Removed Playwright/Chromium → now uses WeasyPrint (pure Python) |

---

## How to deploy (5 steps)

1. **Create a GitHub repo** — push this entire folder to it (all files at root level)
2. Go to **[dashboard.render.com](https://dashboard.render.com)**
3. Click **New → Blueprint**
4. Select your GitHub repo
5. Click **Apply** — all 13 services deploy automatically

Each service gets its own URL, e.g. `https://pdf-to-jpg.onrender.com`

---

## Services & URLs

| Service name | File | What it does |
|-------------|------|--------------|
| `pdf-to-jpg` | photo.py | PDF → JPG images (ZIP) |
| `jpg-to-pdf` | photo1.py | JPG/PNG images → PDF |
| `pdf-to-excel` | backend.py | PDF → Excel (.xlsx / .xls) |
| `excel-to-pdf` | backend1.py | Excel → PDF |
| `pdf-form-filler` | form.py | Fill PDF form fields |
| `pdf-to-word` | server.py | PDF → Word (.docx) |
| `word-to-pdf` | server1.py | Word (.docx) → PDF |
| `pdf-protect` | pass.py | Add password & permissions to PDF |
| `pdf-unlock` | pass1.py | Remove password from PDF |
| `ppt-editor` | app.py | PPT/PPTX editor & PDF export |
| `pptx-to-pdf` | app1.py | PPT/PPTX → PDF |
| `html-to-pdf` | web.py | HTML file, pasted code, or URL → PDF |
| `pdf-to-html` | web1.py | PDF → HTML |

---

## Health check

Every service has `GET /health` → `{"status": "ok"}`.
Set this as the health check path in the Render dashboard.

---

## Notes

- **Free plan spin-down**: Services sleep after 15 min of inactivity; first request takes ~30s to wake up.
- **photo.py / photo1.py**: Office files (.docx, .pptx) are no longer accepted as input — PDF and image inputs still work perfectly.
- **web.py**: WeasyPrint renders HTML very accurately but differently from a browser; JavaScript is not executed.
