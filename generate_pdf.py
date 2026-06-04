#!/usr/bin/env python3
"""
Generate PDF from the presentation data using ReportLab.
Creates a 30-page PDF matching the PPTX content.
"""

import io, os, sys, requests
from pathlib import Path
from PIL import Image as PILImage

from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.units import inch, mm
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, KeepTogether
from reportlab.platypus.flowables import HRFlowable
from reportlab.lib.colors import HexColor, white, black
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

PAGE_W, PAGE_H = 13.33*inch, 7.5*inch   # 16:9

# Colors
TIMBER_DARK  = HexColor('#6B481A')
TIMBER_MED   = HexColor('#9B7335')
TIMBER_LIGHT = HexColor('#F5EDDC')
STEEL_DARK   = HexColor('#1E3552')
STEEL_MED    = HexColor('#3B5C84')
STEEL_LIGHT  = HexColor('#DCE8F5')
MASON_DARK   = HexColor('#7A2810')
MASON_MED    = HexColor('#B04522')
MASON_LIGHT  = HexColor('#F5E5DD')
NEAR_BLACK   = HexColor('#1A1A1A')
MID_GRAY     = HexColor('#909090')
LIGHT_GRAY   = HexColor('#F4F4F4')

SECTION_COLORS = {
    'TIMBER':  (TIMBER_DARK, TIMBER_MED, TIMBER_LIGHT),
    'STEEL':   (STEEL_DARK,  STEEL_MED,  STEEL_LIGHT),
    'MASONRY': (MASON_DARK,  MASON_MED,  MASON_LIGHT),
}

# Import slide data from the main script
sys.path.insert(0, '/home/user/studyo_rehberi')

from generate_presentation import SLIDES, IMAGES, REFERENCES

HEADERS = {"User-Agent": "Mozilla/5.0 (academic PDF generator)"}

def fetch_image(url, fallback_color=(180, 180, 200)):
    try:
        r = requests.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
        if r.status_code == 200 and len(r.content) > 5000:
            img = PILImage.open(io.BytesIO(r.content))
            img.verify()
            return io.BytesIO(r.content)
    except Exception:
        pass
    # Fallback placeholder
    img = PILImage.new("RGB", (600, 450), fallback_color)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    buf.seek(0)
    return buf

def draw_slide(c, slide_data, img_cache):
    sec = slide_data['section']
    dark, med, light = SECTION_COLORS[sec]
    num = slide_data['num']

    W, H = PAGE_W, PAGE_H

    # Background
    c.setFillColor(white)
    c.rect(0, 0, W, H, fill=1, stroke=0)

    # Left accent bar
    c.setFillColor(dark)
    c.rect(0, 0, 0.12*inch, H, fill=1, stroke=0)

    # Header strip
    c.setFillColor(dark)
    c.rect(0.12*inch, H - 0.42*inch, W - 0.12*inch, 0.42*inch, fill=1, stroke=0)

    # Section label in header
    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(0.3*inch, H - 0.30*inch, sec)

    # Slide number in header
    c.setFont("Helvetica", 9)
    c.setFillColor(HexColor('#CCCCCC'))
    num_txt = f"{num} / 30"
    c.drawRightString(W - 0.25*inch, H - 0.30*inch, num_txt)

    # Title
    title = slide_data['title']
    c.setFillColor(dark)
    c.setFont("Helvetica-Bold", 22)
    # Wrap title if needed
    _draw_wrapped_text(c, title, 0.25*inch, H - 1.25*inch, 8.3*inch, 22, dark, bold=True)

    # Subtitle
    subtitle = slide_data['subtitle']
    c.setFillColor(HexColor('#505050'))
    c.setFont("Helvetica-Oblique", 11)
    c.drawString(0.25*inch, H - 1.68*inch, subtitle[:80])

    # Divider line
    c.setStrokeColor(med)
    c.setLineWidth(1.5)
    c.line(0.25*inch, H - 1.82*inch, 8.5*inch, H - 1.82*inch)

    # Content bullets
    cy = H - 1.97*inch
    for (heading, body) in slide_data['bullets']:
        if cy < 1.2*inch:
            break
        # Bullet marker
        c.setFillColor(med)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(0.25*inch, cy, "▪")

        # Heading
        c.setFillColor(dark)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(0.48*inch, cy, heading[:70])
        cy -= 0.22*inch

        # Body text (wrapped)
        c.setFillColor(NEAR_BLACK)
        c.setFont("Helvetica", 9)
        lines = _wrap_text(body, 9, 7.8*inch)
        for line in lines[:3]:
            if cy < 1.2*inch:
                break
            c.drawString(0.48*inch, cy, line)
            cy -= 0.175*inch
        cy -= 0.07*inch

    # Highlight box
    hbox_y = 0.85*inch
    c.setFillColor(light)
    c.rect(0.25*inch, hbox_y, 7.95*inch, 0.72*inch, fill=1, stroke=0)
    c.setFillColor(med)
    c.rect(0.25*inch, hbox_y, 0.07*inch, 0.72*inch, fill=1, stroke=0)

    c.setFillColor(dark)
    c.setFont("Helvetica-Bold", 8.5)
    hl_text = "◆  " + slide_data['highlight']
    lines = _wrap_text(hl_text, 8.5, 7.5*inch)
    hl_y = hbox_y + 0.53*inch
    for ln in lines[:2]:
        c.drawString(0.42*inch, hl_y, ln)
        hl_y -= 0.18*inch

    # Citation
    c.setFillColor(MID_GRAY)
    c.setFont("Helvetica-Oblique", 8)
    c.drawString(0.25*inch, 0.6*inch, slide_data['citation'][:100])

    # Image (right column)
    key = slide_data['img_key']
    buf = img_cache.get(key)
    if buf:
        buf.seek(0)
        try:
            img_x = 8.5*inch
            img_y = 1.5*inch
            img_w = 4.5*inch
            img_h = 5.5*inch

            pil_img = PILImage.open(buf)
            orig_w, orig_h = pil_img.size
            aspect = orig_w / orig_h
            if aspect > img_w/img_h:
                draw_w = img_w
                draw_h = img_w / aspect
            else:
                draw_h = img_h
                draw_w = img_h * aspect

            # Center in image area
            off_x = (img_w - draw_w) / 2
            off_y = (img_h - draw_h) / 2

            buf.seek(0)
            c.drawImage(buf,
                        img_x + off_x,
                        img_y + off_y,
                        draw_w, draw_h)
        except Exception as e:
            pass

    # Image caption
    url, caption, credit, lic = IMAGES[key]
    c.setFillColor(MID_GRAY)
    c.setFont("Helvetica-Oblique", 7)
    cap = f"{caption[:55]}  {credit}"
    c.drawString(8.5*inch, 1.35*inch, cap[:75])
    c.drawString(8.5*inch, 1.2*inch, lic)


def draw_title_slide(c):
    W, H = PAGE_W, PAGE_H
    c.setFillColor(HexColor('#121212'))
    c.rect(0, 0, W, H, fill=1, stroke=0)

    # Left accent bars
    c.setFillColor(TIMBER_DARK)
    c.rect(0, 0, 0.18*inch, H, fill=1, stroke=0)
    c.setFillColor(STEEL_DARK)
    c.rect(0.18*inch, 0, 0.18*inch, H, fill=1, stroke=0)
    c.setFillColor(MASON_DARK)
    c.rect(0.36*inch, 0, 0.18*inch, H, fill=1, stroke=0)

    c.setFillColor(HexColor('#AAAAAA'))
    c.setFont("Helvetica-Oblique", 13)
    c.drawString(0.75*inch, H - 2.2*inch, "ARCHITECTURE & MATERIALS")

    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 40)
    c.drawString(0.75*inch, H - 3.1*inch, "Timber / Steel / Masonry")

    c.setFillColor(HexColor('#CCCCCC'))
    c.setFont("Helvetica-Oblique", 15)
    c.drawString(0.75*inch, H - 3.7*inch,
                 "Structural Logic, Material Culture, and Environmental Performance")

    c.setStrokeColor(HexColor('#555555'))
    c.setLineWidth(1)
    c.line(0.75*inch, H - 4.15*inch, W - 0.75*inch, H - 4.15*inch)

    c.setFillColor(HexColor('#888888'))
    c.setFont("Helvetica", 11)
    c.drawString(0.75*inch, H - 4.45*inch,
                 "30-Slide Academic Presentation  ·  Architecture & Structural Engineering")

    c.setFillColor(TIMBER_MED)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(0.75*inch, H - 5.2*inch, "I. TIMBER  (Slides 1–10)")
    c.setFillColor(STEEL_MED)
    c.drawString(4.8*inch, H - 5.2*inch, "II. STEEL  (Slides 11–20)")
    c.setFillColor(MASON_MED)
    c.drawString(9.0*inch, H - 5.2*inch, "III. MASONRY  (Slides 21–30)")

    c.setFillColor(HexColor('#666666'))
    c.setFont("Helvetica", 9)
    c.drawString(0.75*inch, 0.55*inch, "Prepared: June 2026")


def draw_references_slide(c):
    W, H = PAGE_W, PAGE_H
    c.setFillColor(white)
    c.rect(0, 0, W, H, fill=1, stroke=0)
    c.setFillColor(NEAR_BLACK)
    c.rect(0, 0, 0.12*inch, H, fill=1, stroke=0)
    c.rect(0.12*inch, H - 0.42*inch, W - 0.12*inch, 0.42*inch, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(0.3*inch, H - 0.30*inch, "REFERENCES")

    c.setFillColor(NEAR_BLACK)
    c.setFont("Helvetica-Bold", 17)
    c.drawString(0.25*inch, H - 0.95*inch, "Selected Academic and Authoritative Sources")

    mid = len(REFERENCES) // 2
    c.setFont("Helvetica", 7.5)
    c.setFillColor(NEAR_BLACK)
    for i, ref in enumerate(REFERENCES[:mid]):
        y = H - 1.35*inch - i*0.38*inch
        c.drawString(0.25*inch, y, f"[{i+1}] {ref[:80]}")
    for i, ref in enumerate(REFERENCES[mid:]):
        y = H - 1.35*inch - i*0.38*inch
        c.drawString(6.8*inch, y, f"[{mid+i+1}] {ref[:80]}")


def _wrap_text(text, fontsize, max_width):
    """Naive word-wrap returning list of lines."""
    words = text.split()
    lines = []
    current = []
    # Approximate char width
    char_w = fontsize * 0.55
    max_chars = int(max_width / char_w)
    for word in words:
        test = " ".join(current + [word])
        if len(test) <= max_chars:
            current.append(word)
        else:
            if current:
                lines.append(" ".join(current))
            current = [word]
    if current:
        lines.append(" ".join(current))
    return lines

def _draw_wrapped_text(c, text, x, y, max_w, size, color, bold=False):
    c.setFillColor(color)
    fn = "Helvetica-Bold" if bold else "Helvetica"
    c.setFont(fn, size)
    lines = _wrap_text(text, size, max_w)
    for i, line in enumerate(lines[:2]):
        c.drawString(x, y - i*(size*1.2), line)


def main():
    out_path = Path("/home/user/studyo_rehberi/presentation_output/Timber_Steel_Masonry_Academic_Presentation.pdf")

    print("Loading images for PDF...")
    img_cache = {}
    for key, (url, caption, credit, lic) in IMAGES.items():
        buf = fetch_image(url)
        img_cache[key] = buf
        print(f"  {key} ok")

    print("\nBuilding PDF...")
    c = canvas.Canvas(str(out_path), pagesize=(PAGE_W, PAGE_H))
    c.setTitle("Timber / Steel / Masonry — Academic Presentation")
    c.setAuthor("Architecture & Structural Engineering")
    c.setSubject("Structural Materials in Architecture")

    # Title slide
    draw_title_slide(c)
    c.showPage()

    # 30 content slides
    for sd in SLIDES:
        draw_slide(c, sd, img_cache)
        c.showPage()
        print(f"  slide {sd['num']:02d}/30  {sd['title'][:50]}")

    # References slide
    draw_references_slide(c)
    c.showPage()

    c.save()
    size = out_path.stat().st_size
    print(f"\n  ✔  PDF saved: {out_path}  ({size:,} bytes)")

if __name__ == "__main__":
    main()
