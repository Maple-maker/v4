"""
dd1750_core.py (PATCH for v4)

Fixes requested:
- Move the rendered table UP (reduce top/bottom dead-space)
- Force proper column alignment:
    Unit of issue = EA
    Initial operation = OH QTY
    Running spares = 0
    Total = Initial
- Ensure description WRAPS (handles comma-joined tokens like "CABLE,ASSEMBLY")
- Prevent text running off the page / into other columns
"""

from __future__ import annotations

import textwrap
from typing import List, Dict, Any

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter


# =============================
# LAYOUT CONSTANTS (TUNE HERE)
# =============================

# Vertical placement
ROW_HEIGHT = 19
TOP_Y = 655          # moved UP compared to earlier builds
BOTTOM_Y = 110       # keep above signature block

# Column X positions
X_BOX_CENTER = 70
X_DESC_LEFT = 110

# IMPORTANT: v4 output shows qty columns shifted one column RIGHT.
# These values shift them LEFT to match the DD1750 template.
X_UNIT_CENTER = 348
X_INITIAL_CENTER = 405
X_SPARES_CENTER = 468
X_TOTAL_CENTER = 526

# Wrapping behavior
DESC_MAX_WIDTH_CHARS = 56
MAX_DESC_LINES = 2  # keep tight to preserve row readability

# Fonts
FONT_MAIN = "Helvetica"
FONT_SIZE_DESC = 9
FONT_SIZE_NSN = 8
FONT_SIZE_QTY = 10


# =============================
# TEXT EXTRACTION (non-OCR path)
# =============================
def extract_pdf_text_rows(pdf_path: str) -> List[str]:
    reader = PdfReader(pdf_path)
    rows: List[str] = []
    for page in reader.pages:
        txt = page.extract_text() or ""
        for line in txt.splitlines():
            line = line.strip()
            if line:
                rows.append(line)
    return rows


# =============================
# PARSER: Description + NSN + OH QTY
# =============================
def _digits_only(s: str) -> str:
    return "".join(ch for ch in s if ch.isdigit())


def parse_items_from_rows(rows: List[str]) -> List[Dict[str, Any]]:
    """
    Conservative parser:
    - Uses the first reasonable line as description
    - Uses "NSN:" line for stock number (digits-only)
    - Uses the last token on a line as OH QTY if it's a sane integer
    - Drops OH QTY == 0
    """
    items: List[Dict[str, Any]] = []
    current_desc: str | None = None
    current_nsn: str | None = None

    def flush(qty: int | None):
        nonlocal current_desc, current_nsn
        if not current_desc:
            return
        if qty is None or qty <= 0:
            current_desc, current_nsn = None, None
            return
        items.append({"description": current_desc.strip(), "nsn": (current_nsn or "").strip(), "qty": int(qty)})
        current_desc, current_nsn = None, None

    for line in rows:
        u = line.upper()

        # Skip obvious headers/noise
        if any(k in u for k in [
            "COMPONENT LISTING", "HAND RECEIPT",
            "LV", "DESCRIPTION", "WTY", "ARC", "CIIC", "UI", "SCMC",
            "AUTH QTY", "OH QTY"
        ]):
            continue

        # NSN line
        if u.startswith("NSN"):
            parts = line.split(":")
            maybe = parts[-1].strip() if len(parts) > 1 else line.replace("NSN", "").strip()
            d = _digits_only(maybe)
            if d:
                current_nsn = d
            continue

        # Try detect qty as last token
        qty: int | None = None
        tokens = line.split()
        if tokens and tokens[-1].isdigit():
            q = int(tokens[-1])
            # Sanity bound prevents OCR/parse garbage like 85090307
            if 0 <= q <= 500:
                qty = q

        # If we already have a desc, and this line carries a qty -> flush item
        if qty is not None and current_desc:
            flush(qty)
            continue

        # Otherwise, capture description if we don't have one yet
        if current_desc is None and len(line) >= 4:
            cleaned = line.strip()
            # remove obvious prefixes
            if cleaned.upper().startswith(("BII-", "COEI-")):
                cleaned = " ".join(cleaned.split()[1:]) or line
            current_desc = cleaned

    return items


# =============================
# DRAW HELPERS
# =============================
def draw_centered(c: canvas.Canvas, x: float, y: float, text: str, size: int = 10) -> None:
    c.setFont(FONT_MAIN, size)
    c.drawCentredString(x, y, text)


def normalize_for_wrap(s: str) -> str:
    """
    Ensure comma-joined tokens wrap by inserting a space after commas:
      "CABLE,ASSEMBLY" -> "CABLE, ASSEMBLY"
    """
    # normalize weird spacing
    s = s.replace(", ", ",")
    out = []
    for i, ch in enumerate(s):
        out.append(ch)
        if ch == ",":
            nxt = s[i + 1] if i + 1 < len(s) else ""
            if nxt and nxt != " ":
                out.append(" ")
    return "".join(out).strip()


def draw_wrapped_description(c: canvas.Canvas, x_left: float, y: float, desc: str, nsn: str) -> None:
    desc = normalize_for_wrap(desc)

    c.setFont(FONT_MAIN, FONT_SIZE_DESC)
    lines = textwrap.wrap(
        desc,
        width=DESC_MAX_WIDTH_CHARS,
        break_long_words=True,
        break_on_hyphens=False,
    )
    if not lines:
        lines = [""]

    if len(lines) > MAX_DESC_LINES:
        lines = lines[:MAX_DESC_LINES]
        # truncate the last line slightly
        if len(lines[-1]) > 3:
            lines[-1] = lines[-1][:-3] + "..."

    c.drawString(x_left, y, lines[0])
    if len(lines) > 1:
        c.drawString(x_left, y - 9, lines[1])

    if nsn:
        c.setFont(FONT_MAIN, FONT_SIZE_NSN)
        c.drawString(x_left, y - 18, f"NSN: {nsn}")


# =============================
# MERGE OVERLAY ON TEMPLATE
# =============================
def merge_template(template_pdf: str, overlay_pdf: str, out_pdf: str) -> None:
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
# PUBLIC API USED BY app.py
# =============================
def generate_dd1750_from_pdf(
    bom_pdf: str,
    template_pdf: str,
    output_pdf: str,
) -> List[Dict[str, Any]]:
    """
    Returns the parsed items and writes output_pdf.
    """
    rows = extract_pdf_text_rows(bom_pdf)
    items = parse_items_from_rows(rows)

    overlay_path = output_pdf.replace(".pdf", "_overlay.pdf")
    c = canvas.Canvas(overlay_path, pagesize=letter)

    y = TOP_Y
    box_no = 1

    for item in items:
        qty = int(item["qty"])
        if qty <= 0:
            continue

        # Need headroom for description+nsn drawn down to y-18
        if y < (BOTTOM_Y + 28):
            c.showPage()
            y = TOP_Y

        # Box number
        draw_centered(c, X_BOX_CENTER, y, str(box_no), size=10)

        # Contents + NSN (wrapped)
        draw_wrapped_description(c, X_DESC_LEFT, y, item["description"], item["nsn"])

        # Columns forced as requested
        draw_centered(c, X_UNIT_CENTER, y, "EA", size=FONT_SIZE_QTY)
        draw_centered(c, X_INITIAL_CENTER, y, str(qty), size=FONT_SIZE_QTY)
        draw_centered(c, X_SPARES_CENTER, y, "0", size=FONT_SIZE_QTY)
        draw_centered(c, X_TOTAL_CENTER, y, str(qty), size=FONT_SIZE_QTY)

        y -= ROW_HEIGHT
        box_no += 1

    c.save()
    merge_template(template_pdf, overlay_path, output_pdf)
    return items
