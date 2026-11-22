"""Slidify deck builder.

Builds a .pptx in the Slidify signature style from a JSON spec, and refuses
any spec that breaks the style rules.

Usage:
    python build_deck.py spec.json [output.pptx]

Spec format:
{
  "title": "Deck title",                      # title slide heading
  "subtitle": "Optional one-liner",           # optional
  "slides": [
    {"heading": "Slide heading", "bullets": ["One line", "Another line"]}
  ],
  "closing": "Thank you"                      # optional closing slide line
}
"""

import json
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ---- Signature style -------------------------------------------------------
FONT = "Nirmala Text"
HEADING_PT = 22
BODY_PT = 18

GREY_DARK = RGBColor(0x33, 0x33, 0x33)   # headings, title-slide background
GREY_MID = RGBColor(0x59, 0x59, 0x59)    # body text
GREY_LIGHT = RGBColor(0xA6, 0xA6, 0xA6)  # bullet markers, rules, subtitles
GREY_PALE = RGBColor(0xF2, 0xF2, 0xF2)   # content-slide background

MAX_SLIDES = 12          # whole deck, title and closing slides included
MAX_BULLETS = 5          # per slide
MAX_BULLET_WORDS = 8
MAX_HEADING_WORDS = 6

SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)
MARGIN = Inches(0.9)
TEXT_W = SLIDE_W - 2 * MARGIN
BULLET_INDENT = Inches(0.35)
INSET = Inches(0.1)

FONT_FILE = Path(r"C:\Windows\Fonts\Nirmala.ttc")
FONT_INDEX = 3  # "Nirmala Text Regular" inside Nirmala.ttc


# ---- Line-fit measurement --------------------------------------------------
def _make_measurer():
    """Return f(text, pt) -> width in points, or None if Pillow/font missing."""
    try:
        from PIL import ImageFont
    except ImportError:
        return None
    if not FONT_FILE.exists():
        return None
    cache = {}

    def measure(text, pt):
        if pt not in cache:
            # Size in "pixels" == points when we treat 1px as 1pt.
            cache[pt] = ImageFont.truetype(str(FONT_FILE), pt * 10, index=FONT_INDEX)
        return cache[pt].getlength(text) / 10

    return measure


MEASURE = _make_measurer()


def emu_to_pt(v):
    return Emu(v).pt


def fits_one_line(text, pt, box_width_emu):
    avail = emu_to_pt(box_width_emu) * 0.95  # 5% safety margin
    if MEASURE:
        return MEASURE(text, pt) <= avail
    # Fallback: average glyph is ~0.55 em wide.
    return len(text) * pt * 0.55 <= avail


# ---- Validation ------------------------------------------------------------
def words(text):
    return len(text.split())


def clean(text):
    text = re.sub(r"\s+", " ", str(text)).strip()
    return text.rstrip(".")  # one-liners carry no full stop


def validate(spec):
    errors = []
    slides = spec.get("slides", [])
    total = 1 + len(slides) + (1 if spec.get("closing") else 0)
    if total > MAX_SLIDES:
        errors.append(f"Deck has {total} slides; the maximum is {MAX_SLIDES}.")
    if not spec.get("title"):
        errors.append("Missing deck title.")

    def check_line(label, text, limit, pt, width):
        if "\n" in str(text):
            errors.append(f"{label}: contains a line break: {text!r}")
        t = clean(text)
        if words(t) > limit:
            errors.append(f"{label}: {words(t)} words (max {limit}): {t!r}")
        if not fits_one_line(t, pt, width):
            errors.append(f"{label}: too long to fit on one line: {t!r}")

    check_line("Title", spec.get("title", ""), MAX_HEADING_WORDS, HEADING_PT, TEXT_W - 2 * INSET)
    if spec.get("subtitle"):
        check_line("Subtitle", spec["subtitle"], MAX_BULLET_WORDS, BODY_PT, TEXT_W - 2 * INSET)
    if spec.get("closing"):
        check_line("Closing", spec["closing"], MAX_HEADING_WORDS, HEADING_PT, TEXT_W - 2 * INSET)

    for i, s in enumerate(slides, start=2):
        check_line(f"Slide {i} heading", s.get("heading", ""), MAX_HEADING_WORDS, HEADING_PT, TEXT_W - 2 * INSET)
        bullets = s.get("bullets", [])
        if not bullets:
            errors.append(f"Slide {i}: no bullets.")
        if len(bullets) > MAX_BULLETS:
            errors.append(f"Slide {i}: {len(bullets)} bullets (max {MAX_BULLETS}).")
        for j, b in enumerate(bullets, start=1):
            check_line(f"Slide {i} bullet {j}", b, MAX_BULLET_WORDS, BODY_PT,
                       TEXT_W - 2 * INSET - BULLET_INDENT)
    return errors


# ---- Rendering -------------------------------------------------------------
def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def textbox(slide, left, top, width, height, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = INSET
    tf.margin_top = tf.margin_bottom = Inches(0.05)
    return tf


def style_run(run, pt, color, bold=False):
    f = run.font
    f.name = FONT
    f.size = Pt(pt)
    f.bold = bold
    f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", FONT)


def add_bullet(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(int(BULLET_INDENT)))
    pPr.set("indent", str(-int(BULLET_INDENT)))
    for child in list(pPr):
        if child.tag in (qn("a:buNone"), qn("a:buChar"), qn("a:buClr"), qn("a:buFont")):
            pPr.remove(child)
    buClr = pPr.makeelement(qn("a:buClr"), {})
    srgb = buClr.makeelement(qn("a:srgbClr"), {"val": str(GREY_LIGHT)})
    buClr.append(srgb)
    buFont = pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"})
    buChar = pPr.makeelement(qn("a:buChar"), {"char": "•"})
    pPr.append(buClr)
    pPr.append(buFont)
    pPr.append(buChar)


def rule(slide, top, color=GREY_LIGHT, width=None):
    line = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, MARGIN + INSET, top, MARGIN + INSET + (width or Inches(1.2)), top)
    line.line.color.rgb = color
    line.line.width = Pt(1)


def title_slide(prs, title, subtitle=None):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, GREY_DARK)
    tf = textbox(s, MARGIN, Inches(2.9), TEXT_W, Inches(0.7), MSO_ANCHOR.BOTTOM)
    style_run(tf.paragraphs[0].add_run(), HEADING_PT, GREY_PALE, bold=True)
    tf.paragraphs[0].runs[0].text = clean(title)
    rule(s, Inches(3.75))
    if subtitle:
        tf2 = textbox(s, MARGIN, Inches(3.95), TEXT_W, Inches(0.6))
        r = tf2.paragraphs[0].add_run()
        r.text = clean(subtitle)
        style_run(r, BODY_PT, GREY_LIGHT)
    return s


def bullet_slide(prs, heading, bullets):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, GREY_PALE)
    tf = textbox(s, MARGIN, Inches(0.6), TEXT_W, Inches(0.7), MSO_ANCHOR.BOTTOM)
    r = tf.paragraphs[0].add_run()
    r.text = clean(heading)
    style_run(r, HEADING_PT, GREY_DARK, bold=True)
    rule(s, Inches(1.45))

    body = textbox(s, MARGIN, Inches(1.85), TEXT_W, Inches(5.0))
    for k, b in enumerate(bullets):
        p = body.paragraphs[0] if k == 0 else body.add_paragraph()
        add_bullet(p)
        p.space_after = Pt(BODY_PT * 0.9)
        r = p.add_run()
        r.text = clean(b)
        style_run(r, BODY_PT, GREY_MID)
    return s


def build(spec, out_path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    title_slide(prs, spec["title"], spec.get("subtitle"))
    for s in spec.get("slides", []):
        bullet_slide(prs, s["heading"], s["bullets"])
    if spec.get("closing"):
        title_slide(prs, spec["closing"])
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
    n = 1 + len(spec.get("slides", [])) + (1 if spec.get("closing") else 0)
    print(f"OK: {n} slides -> {out.resolve()}")
    if MEASURE is None:
        print("Note: Pillow or the Nirmala font file was not found; line fit was estimated.")


if __name__ == "__main__":
    main()
