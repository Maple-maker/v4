import os, tempfile, re
from flask import Flask, request, send_file, render_template_string

from pypdf import PdfReader, PdfWriter
from pypdf._page import PageObject
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch

VERSION = "2025-12-31.flask.v4"

HTML = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>BOM → DD1750</title>
  <style>
    body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;max-width:980px;margin:40px auto;padding:0 16px}
    .box{border:1px solid #ddd;border-radius:10px;padding:18px;margin:16px 0}
    label{display:block;margin:10px 0 4px}
    input,select{padding:8px;width:100%}
    button{padding:10px 14px;border:0;border-radius:8px;background:#111;color:#fff;cursor:pointer}
    small{color:#555}
    .row{display:flex;gap:18px}
    .col{flex:1}
  </style>
</head>
<body>
  <h2>BOM → DD1750</h2>
  <p><small>Version: {{version}}</small></p>

  <div class="box">
    <form method="post" action="/generate" enctype="multipart/form-data">
      <div class="row">
        <div class="col">
          <label>BOM (PDF)</label>
          <input type="file" name="bom" accept=".pdf" required>
        </div>
        <div class="col">
          <label>Blank DD1750 template (PDF)</label>
          <input type="file" name="template" accept=".pdf" required>
        </div>
      </div>

      <div class="row">
        <div class="col">
          <label>Label under description</label>
          <select name="label">
            <option value="NSN" selected>NSN</option>
            <option value="SN">SN</option>
          </select>
        </div>
        <div class="col">
          <label>Start parsing BOM at page (0-based)</label>
          <input type="number" name="start_page" value="0" min="0">
        </div>
      </div>

      <p><small>Uses PDF text extraction. If your BOM is scanned, OCR it (Adobe “Recognize Text”) before upload.</small></p>

      <button type="submit">Generate DD1750</button>
    </form>
  </div>
</body>
</html>
"""

app = Flask(__name__)

# ---------- BOM parsing (BOM tables like your B49 style) ----------

LV_RE = re.compile(r"^\s*([A-Z])\s*$")
MAT_RE = re.compile(r"^\s*(\d{7,12})\s*$")  # material/NSN numbers are digits
QTY_RE = re.compile(r"(\d+)\s*$")

def looks_like_qty_line(s: str) -> bool:
    u = s.upper()
    # Lines that commonly contain the OH QTY at end include UI/SCMC fields like "X U EA 9G 1"
    return any(tok in u for tok in [" X ", " U ", " EA ", " AY ", "9G", "9K", "SCMC", "CIIC"])

def is_header_noise(s: str) -> bool:
    u = s.upper()
    return any(h in u for h in [
        "LV", "DESCRIPTION", "WTY", "ARC", "CIIC", "UI", "SCMC", "AUTH", "OH QTY",
        "COMPONENT OF END ITEM", "PAGE", "COEI"
    ])

def normalize_desc(s: str) -> str:
    s = re.sub(r"\s{2,}", " ", s).strip()
    # keep only the real description text, strip trailing column junk if present
    return s[:90]

def extract_items_bom_style(pdf_path: str, start_page: int = 0):
    reader = PdfReader(pdf_path)
    items = []
    cur = {"desc": None, "mat": None, "qty": None}

    def flush():
        nonlocal cur
        if cur["desc"] and cur["mat"] and cur["qty"] is not None and cur["qty"] > 0:
            items.append({"desc": cur["desc"], "mat": cur["mat"], "qty": cur["qty"]})
        cur = {"desc": None, "mat": None, "qty": None}

    for pi in range(start_page, len(reader.pages)):
        txt = reader.pages[pi].extract_text() or ""
        if not txt.strip():
            continue
        lines = [l.rstrip() for l in txt.splitlines() if l.strip()]
        for raw in lines:
            s = raw.strip()
            if not s or is_header_noise(s):
                continue

            # Start of a new LV block
            if LV_RE.match(s):
                flush()
                continue

            # Material number (digits only line)
            mm = MAT_RE.match(s)
            if mm:
                cur["mat"] = mm.group(1)
                continue

            # Qty: often at end of "X U EA ... <qty>" line
            if looks_like_qty_line(" " + s + " "):
                qm = QTY_RE.search(s)
                if qm:
                    qty = int(qm.group(1))
                    # hard guardrail: ignore insane OCR-style quantities
                    if 0 <= qty <= 999:
                        cur["qty"] = qty
                continue

            # Description line: first meaningful non-header, non-mat, non-qty line
            if cur["desc"] is None:
                # avoid grabbing lone codes
                if len(s) >= 3 and not MAT_RE.match(s):
                    cur["desc"] = normalize_desc(s)
                continue

            # Some descriptions wrap; if second line looks like continuation and we have space, append
            if cur["desc"] and cur["mat"] is None and len(cur["desc"]) < 80 and len(s) < 40 and not looks_like_qty_line(" "+s+" "):
                cur["desc"] = normalize_desc(cur["desc"] + " " + s)

    flush()
    # Remove duplicates that happen from repeated headers/page breaks
    dedup = []
    seen = set()
    for it in items:
        key = (it["desc"], it["mat"], it["qty"])
        if key in seen:
            continue
        seen.add(key)
        dedup.append(it)
    return dedup

def paginate(items, per_page=18):
    return [items[i:i+per_page] for i in range(0, len(items), per_page)] or [[]]

# ---------- DD1750 overlay (align to template columns) ----------

def make_overlay(pages, label="NSN"):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    c = canvas.Canvas(tmp.name, pagesize=letter)

    # Column centers (tuned to "blank_flat.pdf" style)
    x_box_center = 1.10*inch
    x_contents_left = 1.62*inch

    # Move UNIT OF ISSUE column LEFT vs earlier builds (user feedback)
    x_uoi_center = 6.18*inch

    # Quantity columns
    x_init_center = 7.05*inch
    x_spares_center = 7.83*inch
    x_total_center = 8.58*inch

    top = 6.55*inch     # start at first writable row (fix "starts halfway down")
    bottom = 1.06*inch
    per_page = 18
    row_h = (top-bottom)/per_page

    for p, rows in enumerate(pages, start=1):
        y = top - row_h/2.0
        for idx, it in enumerate(rows, start=1):
            line_no = (p-1)*per_page + idx

            # Box number centered
            c.setFont("Helvetica", 9)
            s = str(line_no)
            c.drawString(x_box_center - c.stringWidth(s,"Helvetica",9)/2, y-3, s)

            # Contents: description + NSN/SN on next line
            c.setFont("Helvetica", 9)
            c.drawString(x_contents_left, y+2, it["desc"])
            c.setFont("Helvetica", 8)
            c.drawString(x_contents_left, y-10, f"{label}: {it['mat']}")

            # Unit of issue fixed EA, centered in its column
            c.setFont("Helvetica", 9)
            unit = "EA"
            c.drawString(x_uoi_center - c.stringWidth(unit,"Helvetica",9)/2, y-3, unit)

            # Initial operation (d) = qty; Running spares always 0; Total = qty
            q = str(it["qty"])
            c.drawString(x_init_center - c.stringWidth(q,"Helvetica",9)/2, y-3, q)

            z = "0"
            c.drawString(x_spares_center - c.stringWidth(z,"Helvetica",9)/2, y-3, z)

            c.drawString(x_total_center - c.stringWidth(q,"Helvetica",9)/2, y-3, q)

            y -= row_h

        c.showPage()

    c.save()
    return tmp.name

def merge_with_template(template_pdf: str, overlay_pdf: str, out_pdf: str):
    tpl = PdfReader(template_pdf)
    ov = PdfReader(overlay_pdf)
    writer = PdfWriter()
    base = tpl.pages[0]

    for ovp in ov.pages:
        merged = PageObject.create_blank_page(width=base.mediabox.width, height=base.mediabox.height)
        merged.merge_page(base)
        merged.merge_page(ovp)
        writer.add_page(merged)

    with open(out_pdf, "wb") as f:
        writer.write(f)

@app.get("/")
def home():
    return render_template_string(HTML, version=VERSION)

@app.post("/generate")
def generate():
    if "bom" not in request.files or "template" not in request.files:
        return "Missing files", 400

    bom = request.files["bom"]
    template = request.files["template"]
    label = request.form.get("label", "NSN")
    start_page = int(request.form.get("start_page", "0") or "0")

    with tempfile.TemporaryDirectory() as td:
        bom_path = os.path.join(td, "bom.pdf")
        tpl_path = os.path.join(td, "template.pdf")
        bom.save(bom_path)
        template.save(tpl_path)

        items = extract_items_bom_style(bom_path, start_page=start_page)

        # drop qty=0 (per your rule)
        items = [it for it in items if int(it["qty"]) > 0]

        pages = paginate(items, per_page=18)
        overlay = make_overlay(pages, label=label)

        out_pdf = os.path.join(td, "DD1750_OUTPUT.pdf")
        merge_with_template(tpl_path, overlay, out_pdf)

        return send_file(out_pdf, as_attachment=True, download_name="DD1750_OUTPUT.pdf", mimetype="application/pdf")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port)
