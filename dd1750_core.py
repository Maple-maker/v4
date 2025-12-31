from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.lib.units import inch
from io import BytesIO
import fitz

def extract_bom_items(bom_path, start_page=0):
    doc = fitz.open(bom_path)
    items = []
    for page in doc.pages(start=start_page):
        blocks = page.get_text("blocks")
        for b in blocks:
            txt = b[4].strip()
            if txt and "NSN" in txt:
                lines = txt.split("\n")
                if len(lines) >= 2:
                    name = lines[0].strip()
                    nsn = lines[1].replace("NSN:", "").strip()
                    items.append((name, nsn))
    return items

def parse_bom_and_fill_dd1750(bom_path, template_path, output_path, start_page=0):
    items = extract_bom_items(bom_path, start_page)
    if not items:
        raise ValueError("No items parsed from BOM. If this BOM is scanned, convert to a text PDF or provide an Excel export.")

    reader = PdfReader(template_path)
    writer = PdfWriter()
    items_per_page = 29
    total_pages = (len(items) + items_per_page - 1) // items_per_page

    for page_num in range(total_pages):
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFont("Helvetica", 6.8)

        y_start = 580
        y_offset = 13
        x_box = 60
        x_nomen = 110
        x_uoi = 450
        x_init = 500
        x_spare = 530
        x_total = 560

        for i in range(items_per_page):
            idx = page_num * items_per_page + i
            if idx >= len(items):
                break
            name, nsn = items[idx]
            y = y_start - i * y_offset

            box_number = str(idx + 1)
            description = f"{name}\nNSN: {nsn}"
            qty = "1"

            can.drawString(x_box, y, box_number)
            text_object = can.beginText(x_nomen, y)
            for line in description.split("\n"):
                text_object.textLine(line)
            can.drawText(text_object)

            can.drawCentredString(x_uoi + 5, y, "EA")
            can.drawCentredString(x_init + 5, y, qty)
            can.drawCentredString(x_spare + 5, y, "0")
            can.drawCentredString(x_total + 5, y, qty)

        can.save()
        packet.seek(0)
        overlay_pdf = PdfReader(packet)
        base_page = reader.pages[0]
        base_page.merge_page(overlay_pdf.pages[0])
        writer.add_page(base_page)

    with open(output_path, "wb") as f_out:
        writer.write(f_out)