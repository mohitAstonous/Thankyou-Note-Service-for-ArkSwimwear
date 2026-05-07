import os
from io import BytesIO
from uuid import uuid4

from docx import Document
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fontTools.pens.basePen import BasePen
from fontTools.ttLib import TTFont as OpenTypeFont
from pydantic import BaseModel
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
import uharfbuzz as hb


app = FastAPI(title="Ark Swimwear Thank You Note Service")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "generated_output")
TEMPLATE_DOCX = os.path.join(BASE_DIR, "mailmerge.docx")
HANDWRITING_FONT = os.path.join(BASE_DIR, "Reneeshandwriting-Regular (2).otf")
ARK_LOGO = os.path.join(BASE_DIR, "Ark Swimwear Sydney Australia Logo@3x.png")

os.makedirs(OUTPUT_DIR, exist_ok=True)

WORD_HANDWRITING_FEATURES = {
    # Keep the handwriting unconnected by disabling joining substitutions.
    "liga": False,
    "clig": False,
    "calt": False,
    "kern": False,
}

XX_SIGNATURE_FEATURES = {
    "liga": True,
    "clig": True,
    "calt": True,
    "kern": False,
}


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


class OpenTypeTextRenderer:
    def __init__(self, font_path: str):
        with open(font_path, "rb") as font_file:
            self.font_data = font_file.read()

        self.face = hb.Face(self.font_data)
        self.font = hb.Font(self.face)
        hb.ot_font_set_funcs(self.font)
        self.tt_font = OpenTypeFont(font_path)
        self.glyph_set = self.tt_font.getGlyphSet()
        self.units_per_em = self.tt_font["head"].unitsPerEm
        self.font.scale = (self.units_per_em, self.units_per_em)

    def shape(self, text: str, features=None):
        buffer = hb.Buffer()
        buffer.add_str(text)
        buffer.direction = "ltr"
        buffer.script = "Latn"
        buffer.language = "en"
        hb.shape(self.font, buffer, features or WORD_HANDWRITING_FEATURES)
        return zip(buffer.glyph_infos, buffer.glyph_positions)

    def text_width(self, text: str, font_size: float, features=None):
        scale = font_size / self.units_per_em
        return sum(
            position.x_advance * scale
            for _, position in self.shape(text, features=features)
        )

    def draw(self, c, text: str, x: float, y: float, font_size: float, features=None):
        scale = font_size / self.units_per_em
        cursor_x = 0

        c.saveState()
        c.setFillColorRGB(0, 0, 0)
        c.translate(x, y)

        for info, position in self.shape(text, features=features):
            glyph_name = self.tt_font.getGlyphName(info.codepoint)
            glyph = self.glyph_set[glyph_name]
            glyph_x = cursor_x + position.x_offset * scale
            glyph_y = position.y_offset * scale

            c.saveState()
            c.translate(glyph_x, glyph_y)
            c.scale(scale, scale)
            path = c.beginPath()
            glyph.draw(ReportLabPathPen(self.glyph_set, path))
            c.drawPath(path, fill=1, stroke=0)
            c.restoreState()

            cursor_x += position.x_advance * scale

        c.restoreState()


class ReportLabPathPen(BasePen):
    def __init__(self, glyph_set, path):
        super().__init__(glyph_set)
        self.path = path

    def _moveTo(self, point):
        self.path.moveTo(*point)

    def _lineTo(self, point):
        self.path.lineTo(*point)

    def _curveToOne(self, point1, point2, point3):
        self.path.curveTo(*point1, *point2, *point3)

    def _qCurveToOne(self, point1, point2):
        point0 = self._getCurrentPoint()
        curve1 = (
            point0[0] + (2.0 / 3.0) * (point1[0] - point0[0]),
            point0[1] + (2.0 / 3.0) * (point1[1] - point0[1]),
        )
        curve2 = (
            point2[0] + (2.0 / 3.0) * (point1[0] - point2[0]),
            point2[1] + (2.0 / 3.0) * (point1[1] - point2[1]),
        )
        self.path.curveTo(*curve1, *curve2, *point2)

    def _closePath(self):
        self.path.close()


def wrap_text(text: str, font_size: float, max_width: float, text_width):
    words = text.split()
    lines = []
    current_line = ""

    for word in words:
        candidate = f"{current_line} {word}".strip()
        if text_width(candidate, font_size) <= max_width:
            current_line = candidate
        else:
            if current_line:
                lines.append(current_line)
            current_line = word

    if current_line:
        lines.append(current_line)

    return lines


def draw_ark_footer(c, page_width: float):
    logo = ImageReader(ARK_LOGO)
    image_width, image_height = logo.getSize()
    display_width = 110
    display_height = display_width * image_height / image_width
    x = (page_width - display_width) / 2
    c.drawImage(
        logo,
        x,
        15,
        width=display_width,
        height=display_height,
        mask="auto",
        preserveAspectRatio=True,
    )


def create_thank_you_pdf(first_name: str, output_pdf_path: str):
    page_width = 538.67999
    page_height = 368.51999
    font_name = "ReneeHandwriting"

    if font_name not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(font_name, HANDWRITING_FONT))

    handwriting = OpenTypeTextRenderer(HANDWRITING_FONT)

    template_lines = render_template_text(first_name)
    title = template_lines[0] if template_lines else f"Thank you {first_name}!"
    body = template_lines[1:] or [
        "We hope you create some beautiful memories with your new swimwear. Note, your pieces will stretch when worn wet the first few times, so ensure they fit tight at first."
    ]

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=(page_width, page_height))
    c.setTitle(f"Thank you {first_name}")

    left_margin = 10.32
    max_width = 520.36
    font_size = 52

    handwriting.draw(c, title, left_margin, 300.22, font_size)

    y = 250.27
    leading = 49.82
    for paragraph in body:
        for line in wrap_text(
            paragraph,
            font_size,
            max_width,
            handwriting.text_width,
        ):
            handwriting.draw(c, line, left_margin, y, font_size)
            y -= leading

    signature_size = 51.984
    handwriting.draw(c, "With love, Renée and the team", 140.66, 49.2, signature_size)
    handwriting.draw(c, "xx", 472.9, 26.4, signature_size, features=XX_SIGNATURE_FEATURES)
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
    require_file(ARK_LOGO, "Ark footer logo")

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
