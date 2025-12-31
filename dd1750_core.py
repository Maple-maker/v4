from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
import textwrap

# =============================
# DD1750 LAYOUT CONSTANTS
# Adjust ONLY if your template differs
# =============================
ROW_HEIGHT = 18          # more readable spacing, avoids weird bottom gaps
TOP_Y = 610              # where line 1 starts
BOTTOM_Y = 105           # stop before signature block

# Left edges / centers of columns (tuned for your blank_flat template look)
X_BOX_CENTER  = 70

X_DESC_LEFT   = 110
DESC_MAX_WIDTH_CHARS = 52   # wrap threshold for your template scale

# Centers for quantity columns (so text is centered in the boxes)
X_UNIT_CENTER    = 368
X_INITIAL_CENTER = 418
X_SPARES_CENTER  = 474
X_TOTAL_CENTER   = 530

FONT_MAIN = "Helvetica"
FONT_SIZE_DESC = 9
FONT_SIZE_NSN  = 8
FONT_SIZE_QTY  = 10


# =============================
# BASIC PDF TEXT EXTRACTION
# =============================
def extract_pdf_text_rows(pdf_path: str):
    reader = PdfReader(pdf_path)
    rows = []
    for page in reader.pages:
        txt = page.extract_text() or ""
        for line in txt.splitlines():
            line = line.strip()
            if line:
                rows.append(line)
    return rows


# =============================
# PARSER: pull Description + NSN + OH QTY
# (conservative: avoids garbage values)
# =============================
def parse_items_from_rows(rows):
    items = []
    current_desc = None
    current_nsn = None

    def flush_if_ready(qty):
        nonlocal current_desc, current_nsn
        if not current_desc:
            return
        # drop zero qty
        if qty is None or qty <= 0:
            current_desc, current_nsn = None, None
            return
        items.append({
            "description": current_desc.strip(),
            "nsn": (current_nsn or "").strip(),
            "qty": int(qty),
        })
        current_desc, current_nsn = None, None

    for line in rows:
        u = line.upper()

        # Skip obvious headers / noise
        if any(k in u for k in [
            "COMPONENT LISTING", "HAND RECEIPT", "LV", "DESCRIPTION", "WTY",
            "ARC", "CIIC", "UI", "SCMC", "AUTH QTY", "OH QTY"
        ]):
            continue

        # Capture NSN line (your BOMs use NSN or MATERIAL-like tokens sometimes)
        if u.startswith("NSN"):
            # e.g. "NSN: 012821882"
            parts = line.split(":")
            maybe = parts[-1].strip() if len(parts) > 1 else line.replace("NSN", "").strip()
            # keep only digits
            digits = "".join(ch for ch in maybe if ch.isdigit())
            current_nsn = digits if digits else current_nsn
            continue

        # If the line ends with a reasonable qty, treat that as OH QTY
        # (prevents insane values like 11677003 etc. by bounding)
        tokens = line.split()
        qty = None
        if tokens:
            last = tokens[-1]
            if last.isdigit():
                q = int(last)
                if 0 <= q <= 500:   # sanity bound; adjust if your kits legitimately exceed this
                    qty = q

        # If we found qty and we already have a description, flush
        if qty is not None and current_desc:
            flush_if_ready(qty)
            continue

        # Otherwise treat as (part of) description
        # Keep the first “real” line as the description.
        if len(line) >= 4:
            # Some BOMs show a code first like "BII-015341887 Basic Issue Items"
            # Strip leading codes if present
            cleaned = line
            if cleaned.upper().startswith("BII-") or cleaned.upper().startswith("COEI-"):
                # keep words after the code if they exist
                cleaned = " ".join(cleaned.split()[1:]) or line

            # If description already exists, ignore extra continuations to reduce run-ons.
            if not current_desc:
                current_desc = cleaned

    return items


# =============================
# DRAW HELPERS
# =============================
def draw_centered(c: canvas.Canvas, x: float, y: float, text: str, font=FONT_MAIN, size=10):
    c.setFont(font, size)
    c.drawCentredString(x, y, text)

def draw_wrapped_description(c: canvas.Canvas, x_left: float, y: float, desc: str, nsn: str):
    """
    Draw description wrapped to 1–2 lines max, then NSN line under it.
    Prevents text running off the page.
    """
    c.setFont(FONT_MAIN, FONT_SIZE_DESC)
    lines = textwrap.wrap(desc, width=DESC_MAX_WIDTH_CHARS)
    if not lines:
        lines = [""]

    # Force max 2 lines, truncate last line with ellipsis if needed
    if len(lines) > 2:
        lines = lines[:2]
        if len(lines[1]) > 3:
            lines[1] = lines[1][:-3] + "..."

    c.drawString(x_left, y, lines[0])
    if len(lines) > 1:
        c.drawString(x_left, y - 9, lines[1])

    # NSN on the next line down
    if nsn:
        c.setFont(FONT_MAIN, FONT_SIZE_NSN)
        c.drawString(x_left, y - 18, f"NSN: {nsn}")


# =============================
# MERGE OVERLAY ONTO TEMPLATE
# =============================
def merge_template(template_pdf: str, overlay_pdf: str, out_pdf: str):
    base_reader = PdfReader(template_pdf)
    overlay_reader = PdfReader(overlay_pdf)

    writer = PdfWriter()
    base_page = base_reader.pages[0]

    for overlay_page in overlay_reader.pages:
        page = base_page
        page.merge_page(overlay_page)
        writer.add_page(page)

    with open(out_pdf, "wb") as f:
        writer.write(f)


# =============================
# MAIN: BOM PDF -> DD1750 PDF
# =============================
def generate_dd1750_from_pdf(bom_pdf: str, template_pdf: str, output_pdf: str):
    rows = extract_pdf_text_rows(bom_pdf)
    items = parse_items_from_rows(rows)

    # build overlay
    overlay_path = output_pdf.replace(".pdf", "_overlay.pdf")
    c = canvas.Canvas(overlay_path, pagesize=letter)

    y = TOP_Y
    box_no = 1

    for item in items:
        qty = int(item["qty"])
        if qty <= 0:
            continue

        # If we need more vertical space per item because of wrapped desc + nsn,
        # we consume ~ ROW_HEIGHT each item, but description draws down to y-18.
        # So we ensure we have at least ~24px left.
        if y < (BOTTOM_Y + 24):
            c.showPage()
            y = TOP_Y

        # BOX NO centered
        draw_centered(c, X_BOX_CENTER, y, str(box_no), size=10)

        # Description + NSN wrapped
        draw_wrapped_description(c, X_DESC_LEFT, y, item["description"], item["nsn"])

        # Quantities (centered in correct columns)
        draw_centered(c, X_UNIT_CENTER,    y, "EA", size=FONT_SIZE_QTY)
        draw_centered(c, X_INITIAL_CENTER, y, str(qty), size=FONT_SIZE_QTY)
        draw_centered(c, X_SPARES_CENTER,  y, "0", size=FONT_SIZE_QTY)
        draw_centered(c, X_TOTAL_CENTER,   y, str(qty), size=FONT_SIZE_QTY)

        y -= ROW_HEIGHT
        box_no += 1

    c.save()

    # merge onto template (multi-page)
    merge_template(template_pdf, overlay_path, output_pdf)
