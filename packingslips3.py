import os
from io import BytesIO
from uuid import uuid4

from docx import Document
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


app = FastAPI(title="Ark Swimwear Thank You Note Service")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "generated_output")
TEMPLATE_DOCX = os.path.join(BASE_DIR, "mailmerge.docx")
HANDWRITING_FONT = os.path.join(BASE_DIR, "Reneeshandwriting-Regular (2).otf")

os.makedirs(OUTPUT_DIR, exist_ok=True)


class ThankYouRequest(BaseModel):
    name: str


def get_first_name(full_name: str):
    clean_name = " ".join(full_name.split())
    if not clean_name:
        raise HTTPException(status_code=400, detail="Name is required")
    return clean_name.split()[0].capitalize()


def require_file(path: str, description: str):
    if not os.path.exists(path):
        raise HTTPException(
            status_code=500,
            detail=f"{description} not found at {path}",
        )


def render_template_text(first_name: str):
    doc = Document(TEMPLATE_DOCX)
    lines = []

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text:
            lines.append(text.replace("{{ first_name }}", first_name))

    return lines


def wrap_text(text: str, font_name: str, font_size: float, max_width: float):
    words = text.split()
    lines = []
    current_line = ""

    for word in words:
        candidate = f"{current_line} {word}".strip()
        if pdfmetrics.stringWidth(candidate, font_name, font_size) <= max_width:
            current_line = candidate
        else:
            if current_line:
                lines.append(current_line)
            current_line = word

    if current_line:
        lines.append(current_line)

    return lines


def draw_ark_footer(c, page_width: float):
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(page_width / 2, 25.2, "ARK SWIMWEAR")
    c.setFillColorRGB(0.45, 0.45, 0.45)
    c.setFont("Helvetica-Bold", 4.5)
    c.drawCentredString(page_width / 2, 17.3, "SYDNEY AUSTRALIA")
    c.setFillColorRGB(0, 0, 0)


def create_thank_you_pdf(first_name: str, output_pdf_path: str):
    page_width = 538.67999
    page_height = 368.51999
    font_name = "ReneeHandwriting"

    if font_name not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(font_name, HANDWRITING_FONT))

    template_lines = render_template_text(first_name)
    title = template_lines[0] if template_lines else f"Thank you {first_name}!"
    body = template_lines[1:] or [
        "We hope you create some beautiful memories with your new swimwear. Note, your pieces will stretch when worn wet the first few times, so ensure they fit tight at first."
    ]

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=(page_width, page_height))
    c.setTitle(f"Thank you {first_name}")

    left_margin = 28.32
    max_width = 490
    font_size = 48.024

    c.setFont(font_name, font_size)
    c.drawString(left_margin, 300.22, title)

    y = 250.27
    leading = 49.82
    for paragraph in body:
        for line in wrap_text(paragraph, font_name, font_size, max_width):
            c.drawString(left_margin, y, line)
            y -= leading

    signature_size = 51.984
    c.setFont(font_name, signature_size)
    c.drawString(140.66, 49.2, "With love, Renée and the team")
    c.drawString(472.9, 26.4, "xx")
    draw_ark_footer(c, page_width)

    c.save()
    buffer.seek(0)

    with open(output_pdf_path, "wb") as f:
        f.write(buffer.read())


@app.post("/thank-you")
def create_thank_you_note(request: ThankYouRequest):
    first_name = get_first_name(request.name)
    require_file(TEMPLATE_DOCX, "Template DOCX")
    require_file(HANDWRITING_FONT, "Handwriting font")

    output_filename = f"thank_you_{uuid4().hex}.pdf"
    output_path = os.path.join(OUTPUT_DIR, output_filename)
    create_thank_you_pdf(first_name, output_path)

    return FileResponse(
        output_path,
        media_type="application/pdf",
        filename=f"thank_you_{first_name}.pdf",
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
