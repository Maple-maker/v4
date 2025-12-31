import io
import os
import re
import tempfile
from dataclasses import dataclass
from typing import List, Optional, Tuple

from flask import Flask, request, send_file, render_template_string

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import simpleSplit

app = Flask(__name__)

# -----------------------------
# Parsing helpers
# -----------------------------

@dataclass
class ItemRow:
    description: str
    nsn: str
    qty: int

def _clean_nsn(token: str) -> str:
    # Keep only digits; preserve leading zeros
    d = re.sub(r"[^0-9]", "", token or "")
    return d

def _as_int(token: str) -> Optional[int]:
    token = (token or "").strip()
    if not token:
        return None
    # If OCR introduces commas or stray chars, strip.
    token = re.sub(r"[^0-9\-]", "", token)
    if not token:
        return None
    try:
        return int(token)
    except Exception:
        return None

def parse_bom_pdf(bom_bytes: bytes, start_page: int = 0) -> List[ItemRow]:
    """
    Extract rows from a BOM PDF that looks like:
      LV | Description | ... | OH Qty
    We ONLY use:
      - item nomenclature from Description column
      - material/NSN from the left-side material number inside the description block (often appears as digits)
      - OH QTY as qty
    Strategy:
      1) Use pdfplumber to pull tables where possible.
      2) Fallback to text-line parsing and a regex that matches a description block + OH qty.
    """
    rows: List[ItemRow] = []
    with pdfplumber.open(io.BytesIO(bom_bytes)) as pdf:
        for pno in range(start_page, len(pdf.pages)):
            page = pdf.pages[pno]
            # Try table extraction first (works for many BOMs)
            try:
                tables = page.extract_tables()
            except Exception:
                tables = []
            for table in tables or []:
                if not table or len(table) < 2:
                    continue
                # Find header indexes
                header = [str(c or "").strip().upper() for c in table[0]]
                def idx_of(name_opts):
                    for opt in name_opts:
                        for i, h in enumerate(header):
                            if opt in h:
                                return i
                    return None
                idx_desc = idx_of(["DESCRIPTION"])
                idx_oh = idx_of(["OH QTY", "OHQTY", "ON HAND", "ON-HAND", "OH  QTY"])
                idx_mat = idx_of(["MATERIAL", "NSN", "MAT"])
                # Some BOMs embed material in description; allow idx_mat None.
                if idx_desc is None or idx_oh is None:
                    continue
                for r in table[1:]:
                    if not r:
                        continue
                    desc = str(r[idx_desc] or "").strip()
                    if not desc:
                        continue
                    oh = _as_int(str(r[idx_oh] or ""))
                    if oh is None or oh == 0:
                        continue
                    mat = _clean_nsn(str(r[idx_mat] or "")) if idx_mat is not None else ""
                    # If material empty, attempt to find a 9-10 digit NSN-like number in desc
                    if not mat:
                        m = re.search(r"\b(\d{9,10})\b", desc.replace(" ", ""))
                        if m:
                            mat = m.group(1)
                    # Clean description (drop excessive whitespace/newlines)
                    desc = re.sub(r"\s+", " ", desc).strip()
                    if mat and desc:
                        rows.append(ItemRow(description=desc, nsn=mat, qty=int(oh)))

            # Fallback: parse text lines using a state machine.
            txt = page.extract_text() or ""
            if not txt.strip():
                continue
            lines = [re.sub(r"\s+", " ", l).strip() for l in txt.splitlines() if l.strip()]
            # Detect if this looks like a BOM content page (has "LV" and "OH Qty")
            if not any("OH" in l.upper() and "QTY" in l.upper() for l in lines):
                continue

            # Heuristic: in many PDFs, description appears on a line by itself and qty is a trailing token.
            # We'll look for lines that end with an integer and have letters in front.
            for l in lines:
                # Skip obvious headers/footers
                UL = l.upper()
                if UL.startswith("LV ") or "DESCRIPTION" in UL or "WTY" in UL or "ARC" in UL or "CIIC" in UL:
                    continue
                # Try match "... <qty>" at end
                m = re.match(r"^(.*\D)\s+(\d{1,6})$", l)
                if not m:
                    continue
                left, qty_s = m.group(1).strip(), m.group(2)
                qty = int(qty_s)
                if qty == 0:
                    continue
                # left may include an LV at start like "B BASE ASSEMBLY, OUTRIGGER"
                left = re.sub(r"^[A-Z]\s+", "", left)
                # Extract likely NSN/material from left (first 9-10 digit run)
                nsn = ""
                m2 = re.search(r"\b(\d{9,10})\b", left.replace(" ", ""))
                if m2:
                    nsn = m2.group(1)
                    # Remove that numeric run from description
                    left = re.sub(m2.group(1), "", left)
                desc = re.sub(r"\s+", " ", left).strip(" -:")
                if desc and nsn:
                    rows.append(ItemRow(description=desc, nsn=nsn, qty=qty))

    # Deduplicate: keep max qty for same (desc, nsn)
    dedup = {}
    for r in rows:
        key = (r.description.upper(), r.nsn)
        if key not in dedup or r.qty > dedup[key].qty:
            dedup[key] = r
    return list(dedup.values())

# -----------------------------
# Rendering helpers
# -----------------------------

def _wrap_desc(desc: str, max_width: float, font_name: str, font_size: int, max_lines: int = 2) -> List[str]:
    lines = simpleSplit(desc, font_name, font_size, max_width)
    if len(lines) <= max_lines:
        return lines
    # truncate last line with ellipsis
    kept = lines[:max_lines]
    if len(kept[-1]) > 3:
        kept[-1] = kept[-1][: max(0, len(kept[-1]) - 3)] + "..."
    return kept

def build_overlay(items: List[ItemRow], page_size=letter) -> bytes:
    """
    Draw ONLY the list area:
      Box no, Description+NSN, UOI=EA, Initial=qty, Spares=0, Total=qty
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=page_size)

    # Layout (tuned to "blank_flat" 1750 template; adjust as needed)
    # X centers for columns:
    X_BOX = 85
    X_DESC_L = 125
    DESC_W = 285
    X_UOI = 430
    X_INIT = 485
    X_SPARES = 535
    X_TOTAL = 585

    # Vertical layout
    Y_TOP = 525   # start higher to remove top gap
    Y_BOTTOM = 95  # stop above signature block
    ROW_H = 24     # row height so we can fill page without huge gaps

    FONT = "Helvetica"
    c.setFont(FONT, 9)

    idx = 0
    while idx < len(items):
        y = Y_TOP
        # per-page loop
        while idx < len(items) and (y - ROW_H) >= Y_BOTTOM:
            row_num = idx + 1
            it = items[idx]

            # Box number centered
            c.setFont(FONT, 9)
            c.drawCentredString(X_BOX, y - 15, str(row_num))

            # Description wrapped
            desc_lines = _wrap_desc(it.description.upper(), DESC_W, FONT, 10, max_lines=2)
            c.setFont(FONT, 10)
            c.drawString(X_DESC_L, y - 10, desc_lines[0])
            if len(desc_lines) > 1:
                c.drawString(X_DESC_L, y - 21, desc_lines[1])

            # NSN line under description (always numeric)
            c.setFont(FONT, 9)
            c.drawString(X_DESC_L, y - 33, f"NSN: {it.nsn}")

            # UOI, quantities
            c.setFont(FONT, 10)
            c.drawCentredString(X_UOI, y - 15, "EA")
            c.drawCentredString(X_INIT, y - 15, str(it.qty))
            c.drawCentredString(X_SPARES, y - 15, "0")
            c.drawCentredString(X_TOTAL, y - 15, str(it.qty))

            y -= ROW_H
            idx += 1

        c.showPage()

    c.save()
    return buf.getvalue()

def merge_overlay_on_template(template_bytes: bytes, overlay_pdf_bytes: bytes) -> bytes:
    tpl = PdfReader(io.BytesIO(template_bytes))
    ov = PdfReader(io.BytesIO(overlay_pdf_bytes))
    out = PdfWriter()

    # Repeat template first page for each overlay page
    base = tpl.pages[0]
    for p in ov.pages:
        newp = base
        # Need copy to avoid mutating base repeatedly
        newp = newp.copy()
        newp.merge_page(p)
        out.add_page(newp)

    out_buf = io.BytesIO()
    out.write(out_buf)
    return out_buf.getvalue()

# -----------------------------
# Minimal UI
# -----------------------------

HTML = """
<!doctype html>
<title>BOM → DD1750</title>
<style>
body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial; margin: 24px; }
.row { display:flex; gap:24px; }
.col { flex:1; }
label { font-weight:600; display:block; margin-top:12px; }
input[type=file] { width: 100%; }
input[type=number] { width: 120px; }
button { margin-top: 16px; padding: 10px 14px; font-weight:600; }
small { color:#555; }
</style>
<h2>BOM → DD1750 (Flask v4 patched)</h2>
<form method="post" enctype="multipart/form-data">
<div class="row">
  <div class="col">
    <label>BOM (PDF)</label>
    <input name="bom" type="file" accept=".pdf" required>
    <label>Start parsing at page (0-based)</label>
    <input name="start_page" type="number" value="0" min="0">
    <small>Use 0 for normal BOMs.</small>
  </div>
  <div class="col">
    <label>DD1750 template (flat PDF)</label>
    <input name="template" type="file" accept=".pdf" required>
    <small>Upload your blank_flat.pdf template.</small>
  </div>
</div>
<button type="submit">Generate DD1750</button>
</form>
{% if error %}
<p style="color:#b00020; font-weight:600;">{{ error }}</p>
{% endif %}
"""

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template_string(HTML, error=None)

    try:
        bom = request.files.get("bom")
        template = request.files.get("template")
        if not bom or not template:
            return render_template_string(HTML, error="Please upload both BOM and template.")
        start_page = int(request.form.get("start_page", "0") or "0")

        bom_bytes = bom.read()
        tpl_bytes = template.read()

        items = parse_bom_pdf(bom_bytes, start_page=start_page)
        if not items:
            return render_template_string(HTML, error="No items parsed from BOM. If this BOM is scanned, convert to a text PDF or provide an Excel export.")

        overlay = build_overlay(items)
        merged = merge_overlay_on_template(tpl_bytes, overlay)

        return send_file(
            io.BytesIO(merged),
            mimetype="application/pdf",
            as_attachment=True,
            download_name="DD1750.pdf",
        )
    except Exception as e:
        return render_template_string(HTML, error=f"error: {e}")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8080")))
