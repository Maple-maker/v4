from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
import tempfile
import os

# =============================
# DD1750 LAYOUT CONSTANTS
# =============================
ROW_HEIGHT = 11.5
TOP_Y = 640
BOTTOM_Y = 95

X_BOX_NO  = 45
X_DESC    = 90
X_UNIT    = 360
X_INITIAL = 415
X_SPARES  = 470
X_TOTAL   = 525


# =============================
# BASIC PDF TEXT EXTRACTION
# (v4-style, no OCR)
# =============================
def extract_pdf_text_rows(pdf_path):
    reader = PdfReader(pdf_path)
    rows = []

    for page in reader.pages:
        text = page.extract_text()
        if not text:
            continue

        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue

            rows.append(line)

    return rows


# =============================
# VERY CONSERVATIVE BOM PARSER
# (matches prior working behavior)
# =============================
def parse_items_from_rows(rows):
    items = []
    current = {}

    for line in rows:
        upper = line.upper()

        if upper.startswith("NSN"):
            nsn = line.split(":")[-1].strip()
            current["nsn"] = nsn
            continue

        # Quantity heuristic (last numeric token)
        tokens = line.split()
        qty = None
        for t in reversed(tokens):
            if t.isdigit():
                qty = int(t)
                break

        if qty is not None and "description" in current:
            current["qty"] = qty
            items.append(current)
            current = {}
            continue

        # Otherwise treat as description
        if len(line) > 3:
            current = {
                "description": line,
                "nsn": "",
                "qty": 0
            }

    return items


# =============================
# CORE DD1750 GENERATOR
# =============================
def generate_dd1750_from_pdf(bom_pdf, template_pdf, output_pdf):
    rows = extract_pdf_text_rows(bom_pdf)
    items = parse_items_from_rows(rows)

    packet = canvas.Canvas(output_pdf, pagesize=letter)

    def new_page():
        packet.showPage()

    y = TOP_Y
    box_no = 1

    for item in items:
        qty = int(item.get("qty", 0))
        if qty <= 0:
            continue

        desc = item.get("description", "").strip()
        nsn = item.get("nsn", "").strip()

        if y < BOTTOM_Y:
            new_page()
            y = TOP_Y

        packet.drawString(X_BOX_NO, y, str(box_no))
        packet.drawString(X_DESC, y, desc)

        if nsn:
            packet.drawString(X_DESC, y - 5, f"NSN: {nsn}")

        packet.drawString(X_UNIT, y, "EA")
        packet.drawString(X_INITIAL, y, str(qty))
        packet.drawString(X_SPARES, y, "0")
        packet.drawString(X_TOTAL, y, str(qty))

        y -= ROW_HEIGHT
        box_no += 1

    packet.save()

    merge_template(template_pdf, output_pdf)


# =============================
# TEMPLATE MERGE (SAFE)
# =============================
def merge_template(template_pdf, overlay_pdf):
    base_reader = PdfReader(template_pdf)
    overlay_reader = PdfReader(overlay_pdf)

    writer = PdfWriter()

    for i, overlay_page in enumerate(overlay_reader.pages):
        base_page = base_reader.pages[0]
        base_page.merge_page(overlay_page)
        writer.add_page(base_page)

    with open(overlay_pdf, "wb") as f:
        writer.write(f)
