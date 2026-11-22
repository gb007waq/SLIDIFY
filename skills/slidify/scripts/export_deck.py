"""Slidify exports: slide pictures (PNG), PDF and a Word version of a deck.

Uses the same layout and sizing rules as build_deck.py, so the pictures and
PDF match the PowerPoint file.

Usage:
    python export_deck.py spec.json out_folder [--formats png,pdf,docx,zip]
"""

import json
import sys
import zipfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_deck as bd  # noqa: E402

PX_PER_IN = 144                      # 1920 x 1080 slides
W, H = int(13.333 * PX_PER_IN), int(7.5 * PX_PER_IN)


def _px(emu):
    return int(emu / 914400 * PX_PER_IN)


def _pt(pt):
    return int(pt / 72 * PX_PER_IN)


_fonts = {}


def _font(pt):
    if pt not in _fonts:
        try:
            _fonts[pt] = ImageFont.truetype(str(bd.FONT_FILE), _pt(pt), index=bd.FONT_INDEX)
        except OSError:
            _fonts[pt] = ImageFont.load_default(_pt(pt))
    return _fonts[pt]


def _hex(rgb):
    return "#" + str(rgb)


def _text(d, x, y, text, pt, color):
    f = _font(pt)
    # Align the glyph box top with y, like a PowerPoint text line.
    d.text((x, y + _pt(pt) * 0.1), text, font=f, fill=_hex(color))


def _centered(d, box_top_in, box_h_in, text, pt, color):
    f = _font(pt)
    tw = d.textlength(text, font=f)
    line_h = _pt(pt) * 1.2
    y = int(box_top_in * PX_PER_IN + (box_h_in * PX_PER_IN - line_h) / 2)
    _text(d, (W - tw) / 2, y, text, pt, color)


def _logo(img, name, x, y, h):
    path = bd.ASSETS / name
    if not path.exists():
        return 0
    logo = Image.open(path).convert("RGBA")
    w = int(logo.width * h / logo.height)
    logo = logo.resize((w, h))
    if x is None:
        x = (W - w) // 2
    img.paste(logo, (int(x), int(y)), logo)
    return w


def _lines(d, x_emu, top_emu, lines, pt):
    x = _px(x_emu + bd.INSET)
    y = _px(top_emu) + _pt(2)
    step = _pt(pt * 1.2 + bd.LINE_GAP_PT)
    for line in lines:
        _text(d, x, y, bd.clean(line), pt, bd.BLACK)
        y += step


def render_slides(spec):
    """Return a list of PIL images, one per slide."""
    slides = []

    # Title
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    has_logo = _logo(img, "logo_title.png", None, int(2.2 * PX_PER_IN), int(1.6 * PX_PER_IN))
    _centered(d, 4.1 if has_logo else 3.3, 0.9, bd.clean(spec["title"]), bd.TITLE_PT, bd.GREY_CLOSING)
    if spec.get("date"):
        _centered(d, 6.3, 0.5, bd.clean(spec["date"]).upper(), bd.BODY_SIZES[-1], bd.GREY_HEADING)
    slides.append(img)

    content = list(spec.get("slides", []))
    if spec.get("agenda"):
        content.insert(0, bd.agenda_slide(spec))

    for s in content:
        img = Image.new("RGB", (W, H), "white")
        d = ImageDraw.Draw(img)
        hpt = bd.pick_size(bd.HEADING_SIZES, [s["heading"]], bd.TEXT_W)
        hy = int(0.45 * PX_PER_IN + (0.7 * PX_PER_IN - _pt(hpt) * 1.2) / 2)
        _text(d, _px(bd.LEFT + bd.INSET), hy, bd.clean(s["heading"]), hpt, bd.GREY_HEADING)
        _logo(img, "logo_top.png", W - int(0.95 * PX_PER_IN), 0, int(1.0 * PX_PER_IN))
        _logo(img, "logo_corner.png", W - int(1.65 * PX_PER_IN), H - int(0.85 * PX_PER_IN), int(0.55 * PX_PER_IN))
        top = int(1.65 * 914400)
        if "lines" in s:
            pt = bd.pick_size(bd.BODY_SIZES, s["lines"], bd.TEXT_W)
            _lines(d, bd.LEFT, top, s["lines"], pt)
        else:
            pt = bd.pick_size(bd.BODY_SIZES, s["left"] + s["right"], bd.COL_W)
            _lines(d, bd.COL_LEFT_X, top, s["left"], pt)
            _lines(d, bd.COL_RIGHT_X, top, s["right"], pt)
            n = max(len(s["left"]), len(s["right"]))
            x = _px(bd.DIVIDER_X)
            y0 = _px(top) - int(0.1 * PX_PER_IN)
            y1 = _px(top) + n * _pt(pt * 1.2 + bd.LINE_GAP_PT) - int(0.2 * PX_PER_IN)
            d.line([(x, y0), (x, y1)], fill="black", width=max(2, _pt(1.25)))
        slides.append(img)

    closing = spec.get("closing", "THANK YOU")
    if closing:
        img = Image.new("RGB", (W, H), "white")
        _centered(ImageDraw.Draw(img), 3.3, 0.9, bd.clean(closing).upper(), bd.TITLE_PT, bd.GREY_CLOSING)
        slides.append(img)
    return slides


def write_pngs(images, folder, stem="slide"):
    folder.mkdir(parents=True, exist_ok=True)
    paths = []
    for i, im in enumerate(images, 1):
        p = folder / f"{stem}-{i:02d}.png"
        im.save(p)
        paths.append(p)
    return paths


def write_pdf(images, path):
    images[0].save(path, save_all=True, append_images=images[1:], resolution=PX_PER_IN)


def write_zip(paths, path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            z.write(p, p.name)


def write_docx(spec, path):
    """A Word version: one page per slide, same fonts, sizes and greys."""
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt, RGBColor

    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Inches(11.69), Inches(8.27)
    for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, m, Inches(0.8))

    def para(text, pt, rgb, align=None, before=0, after=12, page_break=False):
        p = doc.add_paragraph()
        if align is not None:
            p.alignment = align
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        r = p.add_run(text)
        r.font.name = bd.FONT
        r._element.rPr.rFonts.set(qn("w:eastAsia"), bd.FONT)
        r._element.rPr.rFonts.set(qn("w:cs"), bd.FONT)
        r.font.size = Pt(pt)
        r.font.color.rgb = RGBColor.from_string(str(rgb))
        if page_break:
            r.add_break(WD_BREAK.PAGE)
        return p

    center = WD_ALIGN_PARAGRAPH.CENTER
    para(bd.clean(spec["title"]), bd.TITLE_PT, bd.GREY_CLOSING, center, before=170, after=24)
    if spec.get("date"):
        para(bd.clean(spec["date"]).upper(), bd.BODY_SIZES[-1], bd.GREY_HEADING, center)
    doc.paragraphs[-1].runs[-1].add_break(WD_BREAK.PAGE)

    content = list(spec.get("slides", []))
    if spec.get("agenda"):
        content.insert(0, {"heading": "Sequence", "lines": [s["heading"] for s in spec["slides"]]})
    for s in content:
        para(bd.clean(s["heading"]), bd.HEADING_SIZES[0], bd.GREY_HEADING, after=24)
        lines = s.get("lines") or (s.get("left", []) + s.get("right", []))
        for line in lines:
            para(bd.clean(line), bd.BODY_SIZES[0], bd.BLACK, after=16)
        doc.paragraphs[-1].runs[-1].add_break(WD_BREAK.PAGE)

    closing = spec.get("closing", "THANK YOU")
    if closing:
        para(bd.clean(closing).upper(), bd.TITLE_PT, bd.GREY_CLOSING, center, before=200)
    else:
        last = doc.paragraphs[-1].runs[-1]._element
        for br in last.findall(qn("w:br")):
            last.remove(br)
    doc.save(path)


def export_all(spec, folder, stem="deck"):
    """Write deck.pptx, deck.pdf, deck.docx, slide PNGs and slides.zip into folder."""
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    out = {}
    out["pptx"] = folder / f"{stem}.pptx"
    bd.build(spec, out["pptx"])
    images = render_slides(spec)
    out["png"] = write_pngs(images, folder / "slides")
    out["pdf"] = folder / f"{stem}.pdf"
    write_pdf(images, out["pdf"])
    out["zip"] = folder / f"{stem} - slides.zip"
    write_zip(out["png"], out["zip"])
    out["docx"] = folder / f"{stem}.docx"
    write_docx(spec, out["docx"])
    return out


def main():
    args, (lo, hi) = bd.split_args(sys.argv[1:])
    if len(args) < 2:
        print(__doc__)
        sys.exit(2)
    spec = json.loads(Path(args[0]).read_text(encoding="utf-8-sig"))
    errors = bd.validate(spec, lo, hi)
    if errors:
        print("STYLE CHECK FAILED - fix these and run again:")
        for e in errors:
            print("  - " + e)
        sys.exit(1)
    out = export_all(spec, args[1], stem=bd.clean(spec["title"]))
    for k, v in out.items():
        print(k, ":", f"{len(v)} files" if isinstance(v, list) else v)


if __name__ == "__main__":
    main()
