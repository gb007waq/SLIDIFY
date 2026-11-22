"""Slidify deck builder.

Builds a .pptx in the Slidify signature style (see ../STYLE.md) from a JSON
spec, and refuses any spec that breaks the style rules.

Usage:
    python build_deck.py spec.json [output.pptx]

Spec format:
{
  "title": "Deck title",
  "date": "NOV, 2026",                       # optional, shown under the title
  "agenda": true,                            # optional "Sequence" slide
  "slides": [
    {"heading": "Heading", "lines": ["One-liner", "Item | Paired item"]},
    {"heading": "Heading", "left": ["..."], "right": ["..."]}
  ],
  "closing": "THANK YOU"                     # optional; default THANK YOU, "" to omit
}

Optional logos: put PNG files in the skill's assets/ folder:
    logo_title.png   centred on the title slide
    logo_top.png     top-right of every content slide
    logo_corner.png  bottom-right of every content slide
"""

import json
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ---- Signature style -------------------------------------------------------
FONT = "Nirmala Text"
TITLE_PT = 36               # title slide and THANK YOU
HEADING_SIZES = (24, 22)    # slide headings: 24, or 22 if 24 won't fit on one line
BODY_SIZES = (20, 18)       # body lines: 20, or 18 for the whole slide if needed
HEADING_PT = HEADING_SIZES[-1]
BODY_PT = BODY_SIZES[-1]

WHITE = RGBColor(0xFF, 0xFF, 0xFF)       # background
GREY_HEADING = RGBColor(0xA6, 0xA6, 0xA6)  # slide headings, date
GREY_CLOSING = RGBColor(0x80, 0x80, 0x80)  # title and THANK YOU
BLACK = RGBColor(0x00, 0x00, 0x00)       # body lines, column divider

MAX_SLIDES = 12
MAX_LINES = 7            # per slide (per column on two-column slides)
MAX_LINE_WORDS = 12
MAX_HEADING_WORDS = 5

SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)
LEFT = Inches(0.7)
TEXT_W = SLIDE_W - LEFT - Inches(1.6)    # keep clear of the right-hand logos
INSET = Inches(0.05)
COL_LEFT_X = Inches(1.0)
COL_W = Inches(5.3)
DIVIDER_X = Inches(6.55)
COL_RIGHT_X = Inches(6.95)
LINE_GAP_PT = 32                         # space between one-liners

SKILL_DIR = Path(__file__).resolve().parent.parent
ASSETS = SKILL_DIR / "assets"

FONT_FILE = Path(r"C:\Windows\Fonts\Nirmala.ttc")
FONT_INDEX = 3  # "Nirmala Text Regular" inside Nirmala.ttc


# ---- Line-fit measurement --------------------------------------------------
def _make_measurer():
    try:
        from PIL import ImageFont
    except ImportError:
        return None
    if not FONT_FILE.exists():
        return None
    cache = {}

    def measure(text, pt):
        if pt not in cache:
            cache[pt] = ImageFont.truetype(str(FONT_FILE), pt * 10, index=FONT_INDEX)
        return cache[pt].getlength(text) / 10

    return measure


MEASURE = _make_measurer()


def fits_one_line(text, pt, box_width_emu):
    avail = Emu(box_width_emu - 2 * INSET).pt * 0.95
    if MEASURE:
        return MEASURE(text, pt) <= avail
    return len(text) * pt * 0.55 <= avail


# ---- Validation ------------------------------------------------------------
def clean(text):
    text = re.sub(r"\s+", " ", str(text)).strip()
    text = re.sub(r"\s*\|\s*", " | ", text)   # tidy pipe pairs
    return text.rstrip(".")                   # one-liners carry no full stop


def slide_count(spec):
    return (1 + (1 if spec.get("agenda") else 0) + len(spec.get("slides", []))
            + (1 if spec.get("closing", "THANK YOU") else 0))


def validate(spec):
    errors = []
    slides = spec.get("slides", [])
    total = slide_count(spec)
    if total > MAX_SLIDES:
        errors.append(f"Deck has {total} slides; the maximum is {MAX_SLIDES}.")
    if not spec.get("title"):
        errors.append("Missing deck title.")
    if spec.get("agenda") and len(slides) > MAX_LINES:
        errors.append(f"Agenda would list {len(slides)} headings (max {MAX_LINES}); drop the agenda or merge slides.")

    def check(label, text, limit, pt, width):
        if "\n" in str(text):
            errors.append(f"{label}: contains a line break: {text!r}")
        t = clean(text)
        if not t:
            errors.append(f"{label}: empty.")
        if len(t.split()) > limit:
            errors.append(f"{label}: {len(t.split())} words (max {limit}): {t!r}")
        if not fits_one_line(t, pt, width):
            errors.append(f"{label}: too long for one line: {t!r}")

    check("Title", spec.get("title", ""), 8, TITLE_PT, SLIDE_W - 2 * LEFT)

    for i, s in enumerate(slides, start=2 + (1 if spec.get("agenda") else 0)):
        check(f"Slide {i} heading", s.get("heading", ""), MAX_HEADING_WORDS, HEADING_PT, TEXT_W)
        if "lines" in s:
            cols = [("", s["lines"], TEXT_W)]
        else:
            cols = [("left ", s.get("left", []), COL_W), ("right ", s.get("right", []), COL_W)]
        for name, lines, width in cols:
            if not lines:
                errors.append(f"Slide {i}: {name}column has no lines.")
            if len(lines) > MAX_LINES:
                errors.append(f"Slide {i}: {name}{len(lines)} lines (max {MAX_LINES}).")
            for j, line in enumerate(lines, start=1):
                check(f"Slide {i} {name}line {j}", line, MAX_LINE_WORDS, BODY_PT, width)
    return errors


# ---- Rendering -------------------------------------------------------------
def white(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = WHITE


def textbox(slide, left, top, width, height, anchor=MSO_ANCHOR.TOP):
    tf = slide.shapes.add_textbox(left, top, width, height).text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = INSET
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    return tf


def write(paragraph, text, pt, color, align=None):
    if align is not None:
        paragraph.alignment = align
    r = paragraph.add_run()
    r.text = text
    f = r.font
    f.name, f.size, f.bold = FONT, Pt(pt), False
    f.color.rgb = color
    rPr = r._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", FONT)


def picture(slide, name, left, top, height):
    path = ASSETS / name
    if path.exists():
        slide.shapes.add_picture(str(path), left, top, height=height)


def blank(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    white(s)
    return s


def pick_size(sizes, texts, width):
    """Largest allowed size at which every text fits on one line."""
    for pt in sizes:
        if all(fits_one_line(clean(t), pt, width) for t in texts):
            return pt
    return sizes[-1]


def heading(slide, text):
    tf = textbox(slide, LEFT, Inches(0.45), TEXT_W, Inches(0.7), MSO_ANCHOR.MIDDLE)
    pt = pick_size(HEADING_SIZES, [text], TEXT_W)
    write(tf.paragraphs[0], clean(text), pt, GREY_HEADING)
    picture(slide, "logo_top.png", SLIDE_W - Inches(0.95), 0, Inches(1.0))
    picture(slide, "logo_corner.png", SLIDE_W - Inches(1.65), SLIDE_H - Inches(0.85), Inches(0.55))


def lines_box(slide, left, top, width, lines, pt):
    tf = textbox(slide, left, top, width, SLIDE_H - top - Inches(0.6))
    for k, line in enumerate(lines):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.space_after = Pt(LINE_GAP_PT)
        write(p, clean(line), pt, BLACK)
    return tf


def title_slide(prs, title, date=None):
    s = blank(prs)
    logo = ASSETS / "logo_title.png"
    if logo.exists():
        pic = s.shapes.add_picture(str(logo), 0, Inches(2.2), height=Inches(1.6))
        pic.left = int((SLIDE_W - pic.width) / 2)
        top = Inches(4.1)
    else:
        top = Inches(3.3)
    tf = textbox(s, LEFT, top, SLIDE_W - 2 * LEFT, Inches(0.9), MSO_ANCHOR.MIDDLE)
    write(tf.paragraphs[0], clean(title), TITLE_PT, GREY_CLOSING, PP_ALIGN.CENTER)
    if date:
        tf = textbox(s, LEFT, Inches(6.3), SLIDE_W - 2 * LEFT, Inches(0.5), MSO_ANCHOR.MIDDLE)
        write(tf.paragraphs[0], clean(date).upper(), BODY_PT, GREY_HEADING, PP_ALIGN.CENTER)


def content_slide(prs, spec):
    s = blank(prs)
    heading(s, spec["heading"])
    top = Inches(1.65)
    if "lines" in spec:
        pt = pick_size(BODY_SIZES, spec["lines"], TEXT_W)
        lines_box(s, LEFT, top, TEXT_W, spec["lines"], pt)
    else:
        pt = pick_size(BODY_SIZES, spec["left"] + spec["right"], COL_W)
        lines_box(s, COL_LEFT_X, top, COL_W, spec["left"], pt)
        lines_box(s, COL_RIGHT_X, top, COL_W, spec["right"], pt)
        n = max(len(spec["left"]), len(spec["right"]))
        pitch = Pt(pt * 1.2 + LINE_GAP_PT)
        bottom = top + int(pitch * n)
        d = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, DIVIDER_X, top - Inches(0.1),
                                   DIVIDER_X, bottom - Inches(0.2))
        d.line.color.rgb = BLACK
        d.line.width = Pt(1.25)


def closing_slide(prs, text):
    s = blank(prs)
    tf = textbox(s, LEFT, Inches(3.3), SLIDE_W - 2 * LEFT, Inches(0.9), MSO_ANCHOR.MIDDLE)
    write(tf.paragraphs[0], clean(text).upper(), TITLE_PT, GREY_CLOSING, PP_ALIGN.CENTER)


def build(spec, out_path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    title_slide(prs, spec["title"], spec.get("date"))
    if spec.get("agenda"):
        content_slide(prs, {"heading": "Sequence",
                            "lines": [s["heading"] for s in spec["slides"]]})
    for s in spec.get("slides", []):
        content_slide(prs, s)
    closing = spec.get("closing", "THANK YOU")
    if closing:
        closing_slide(prs, closing)
    prs.save(out_path)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    spec_path = Path(sys.argv[1])
    spec = json.loads(spec_path.read_text(encoding="utf-8-sig"))
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else spec_path.with_suffix(".pptx")

    errors = validate(spec)
    if errors:
        print("STYLE CHECK FAILED - fix these and run again:")
        for e in errors:
            print("  - " + e)
        sys.exit(1)

    build(spec, out)
    print(f"OK: {slide_count(spec)} slides -> {out.resolve()}")
    if MEASURE is None:
        print("Note: Pillow or the Nirmala font file was not found; line fit was estimated.")


if __name__ == "__main__":
    main()
