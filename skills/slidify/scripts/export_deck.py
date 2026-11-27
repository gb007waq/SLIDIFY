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
    key = (str(bd.FONT_FILE), bd.FONT_INDEX, pt)
    if key not in _fonts:
        try:
            _fonts[key] = ImageFont.truetype(str(bd.FONT_FILE), _pt(pt), index=bd.FONT_INDEX)
        except (OSError, TypeError):
            _fonts[key] = ImageFont.load_default(_pt(pt))
    return _fonts[key]


def _hex(rgb):
    return "#" + str(rgb)


def _text(d, x, y, text, pt, color, stroke=0):
    f = _font(pt)
    # Align the glyph box top with y, like a PowerPoint text line.
    d.text((x, y + _pt(pt) * 0.1), text, font=f, fill=_hex(color),
           stroke_width=stroke, stroke_fill=_hex(color) if stroke else None)


def _semibold(img, x, y, text, pt, color):
    """'A little thick': draw at 3x with a 1px outline, then shrink, so the outline is a third of a pixel."""
    k = 3
    f = ImageFont.truetype(str(bd.FONT_FILE), _pt(pt) * k, index=bd.FONT_INDEX)
    w = int(f.getlength(text)) + 8 * k
    h = int(_pt(pt) * 1.6) * k
    layer = Image.new("L", (w, h), 0)
    ImageDraw.Draw(layer).text((2 * k, _pt(pt) * 0.1 * k), text, font=f, fill=255, stroke_width=1, stroke_fill=255)
    mask = layer.resize((w // k, h // k), Image.LANCZOS)
    img.paste(Image.new("RGB", mask.size, _hex(color)), (int(x) - 2, int(y)), mask)


def _footer(d, s):
    """Sources line: 9 pt light grey at the foot of the slide, wrapping to two lines if needed."""
    text = bd.sources_text(s)
    if not text:
        return
    f = _font(bd.SOURCE_PT)
    x, y = _px(bd.LEFT + bd.INSET), _px(bd.SOURCE_TOP) + _pt(2)
    avail = _px(bd.TEXT_W - 2 * bd.INSET)
    line, lines = "", []
    for word in text.split(" "):
        test = (line + " " + word).strip()
        if d.textlength(test, font=f) > avail and line:
            lines.append(line)
            line = word
        else:
            line = test
    lines.append(line)
    for ln in lines[:2]:
        _text(d, x, y, ln, bd.SOURCE_PT, bd.SOURCE_COLOR)
        y += _pt(bd.SOURCE_PT * 1.2)


def _data(img, d, s):
    """Draw KPIs and the grey bar chart the same way data_slide lays them out."""
    lay = bd.data_layout(s)
    for n, k in enumerate(s.get("kpis") or []):
        x = _px(bd.LEFT + n * lay["cell_w"] + bd.INSET)
        top = _px(lay["kpi_top"])
        vy = top + int(0.8 * PX_PER_IN) - int(_pt(bd.KPI_PT) * 1.25)
        _semibold(img, x, vy, bd.clean(k["value"]), bd.KPI_PT, bd.TIER_COLORS[0])
        _text(d, x, top + int(0.8 * PX_PER_IN) + _pt(3), bd.clean(k["label"]), bd.KPI_LABEL_PT, bd.TIER_COLORS[2])
    chart = s.get("chart")
    if not chart:
        return
    bars = chart["bars"]
    x0, y0, cw, chh = (_px(v) for v in lay["chart_box"])
    label_h, value_h = _pt(bd.CHART_PT) * 2, _pt(bd.CHART_PT) * 1.8
    base_y = y0 + chh - label_h
    top_y = y0 + value_h
    slot = cw / len(bars)
    bar_w = slot / 1.7                                    # matches PowerPoint gap width 70%
    vmax = max(float(b["value"]) for b in bars) or 1
    unit = bd.chart_unit(chart)
    d.line([(x0, base_y), (x0 + cw, base_y)], fill=_hex(bd.CHART_AXIS), width=2)
    fv, fl = _font(bd.CHART_PT), _font(bd.CHART_PT)
    for n, b in enumerate(bars):
        cx = x0 + slot * n + slot / 2
        h = (base_y - top_y) * float(b["value"]) / vmax
        d.rectangle([cx - bar_w / 2, base_y - h, cx + bar_w / 2, base_y], fill=_hex(bd.TIER_COLORS[2]))
        val = bd.fmt_value(b["value"], unit)
        d.text((cx - d.textlength(val, font=fv) / 2, base_y - h - _pt(bd.CHART_PT) * 1.35), val,
               font=fv, fill=_hex(bd.TIER_COLORS[0]))
        lab = bd.clean(b["label"])
        d.text((cx - d.textlength(lab, font=fl) / 2, base_y + _pt(6)), lab, font=fl, fill=_hex(bd.TIER_COLORS[1]))


def _centered(d, box_top_in, box_h_in, text, pt, color):
    f = _font(pt)
    tw = d.textlength(text, font=f)
    line_h = _pt(pt) * 1.2
    y = int(box_top_in * PX_PER_IN + (box_h_in * PX_PER_IN - line_h) / 2)
    _text(d, (W - tw) / 2, y, text, pt, color)


def _chosen_logo(img, spec, left, top):
    """The presentation's logo, fitted in a 1 cm square at (left, top) EMU."""
    path = bd.logo_path(spec)
    if not path:
        return
    x, y, w, h = bd.logo_box(path, left, top)
    logo = Image.open(path).convert("RGBA").resize((max(1, _px(w)), max(1, _px(h))), Image.LANCZOS)
    img.paste(logo, (_px(x), _px(y)), logo)


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


def _lines(d, x_emu, top_emu, lines, pt, n=None):
    x = _px(x_emu + bd.INSET)
    y = _px(top_emu) + _pt(2)
    step = _pt(pt * 1.2 + bd.line_gap(n or len(lines), pt))
    for line in lines:
        _text(d, x, y, bd.clean(line), pt, bd.BLACK)
        y += step


def render_slides(spec):
    """Return a list of PIL images, one per slide."""
    slides = []

    # Title
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    _chosen_logo(img, spec, (bd.SLIDE_W - bd.LOGO_SIZE) // 2, bd.title_logo_top(spec))
    parts, tpt, (_, ty, _, th) = bd.title_layout(spec)        # dark line, lighter line, tight
    line_h = _pt(tpt * 1.08)
    y = _px(ty) + (_px(th) - line_h * len(parts)) / 2
    for n, part in enumerate(parts):
        f = _font(tpt)
        _text(d, (W - d.textlength(part, font=f)) / 2, y, part, tpt, bd.TITLE_COLORS[min(n, 1)])
        y += line_h
    if spec.get("date"):
        _centered(d, 6.3, 0.5, bd.clean(spec["date"]).upper(), bd.BODY_SIZES[-1], bd.GREY_HEADING)
    slides.append(img)

    content = list(spec.get("slides", []))
    if spec.get("agenda"):
        content.insert(0, bd.agenda_slide(spec))

    for s in content:
        img = Image.new("RGB", (W, H), "white")
        d = ImageDraw.Draw(img)
        if s.get("kind") == "section":                    # Pyramid "Appendix" divider
            _centered(d, 3.3, 0.9, bd.clean(s["heading"]), bd.TITLE_PT, bd.GREY_HEADING)
            slides.append(img)
            continue
        if s.get("kind") == "data":                       # Pyramid base: KPIs and / or a bar chart
            hpt = bd.pick_size(bd.HEADING_SIZES, [s["heading"]], bd.TEXT_W)
            hy = int(0.45 * PX_PER_IN + (0.7 * PX_PER_IN - _pt(hpt) * 1.2) / 2)
            _text(d, _px(bd.LEFT + bd.INSET), hy, bd.clean(s["heading"]), hpt, bd.GREY_HEADING)
            _data(img, d, s)
            _footer(d, s)
            _chosen_logo(img, spec, *bd.corner_logo_pos())
            slides.append(img)
            continue
        hpt = bd.pick_size(bd.HEADING_SIZES, [s["heading"]], bd.TEXT_W)
        hy = int(0.45 * PX_PER_IN + (0.7 * PX_PER_IN - _pt(hpt) * 1.2) / 2)
        _text(d, _px(bd.LEFT + bd.INSET), hy, bd.clean(s["heading"]), hpt, bd.GREY_HEADING)
        top = int(bd.BODY_TOP)
        if bd.is_tiered(s):                               # Pyramid: fact / proof / evidence
            pt = bd.pick_size(bd.BODY_SIZES, [t for t, _ in bd.tier_lines(s)], bd.TEXT_W)
            rows, gaps = bd.tier_layout(s, pt)
            x, y = _px(bd.LEFT + bd.INSET), _px(top) + _pt(2)
            for (text, tier), gap in zip(rows, gaps):
                if tier == 0:
                    _semibold(img, x, y, bd.clean(text), pt, bd.TIER_COLORS[tier])
                else:
                    _text(d, x, y, bd.clean(text), pt, bd.TIER_COLORS[tier])
                y += _pt(pt * 1.2 + gap)
        elif "lines" in s:
            pt = bd.pick_size(bd.BODY_SIZES, s["lines"], bd.TEXT_W)
            _lines(d, bd.LEFT, top, s["lines"], pt)
        else:
            pt = bd.pick_size(bd.BODY_SIZES, s["left"] + s["right"], bd.COL_W)
            n = max(len(s["left"]), len(s["right"]))
            _lines(d, bd.COL_LEFT_X, top, s["left"], pt, n)
            _lines(d, bd.COL_RIGHT_X, top, s["right"], pt, n)
            x = _px(bd.DIVIDER_X)
            y0 = _px(top) - int(0.1 * PX_PER_IN)
            y1 = _px(top) + n * _pt(pt * 1.2 + bd.line_gap(n, pt)) - int(0.2 * PX_PER_IN)
            d.line([(x, y0), (x, y1)], fill="black", width=max(2, _pt(1.25)))
        _footer(d, s)
        _chosen_logo(img, spec, *bd.corner_logo_pos())
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
    parts = bd.title_parts(spec)
    for n, part in enumerate(parts):
        para(part, bd.TITLE_SLIDE_SIZES[0], bd.TITLE_COLORS[min(n, 1)], center,
             before=170 if n == 0 else 0, after=24 if n == len(parts) - 1 else 0)
    if spec.get("date"):
        para(bd.clean(spec["date"]).upper(), bd.BODY_SIZES[-1], bd.GREY_HEADING, center)
    doc.paragraphs[-1].runs[-1].add_break(WD_BREAK.PAGE)

    content = list(spec.get("slides", []))
    if spec.get("agenda"):
        content.insert(0, {"heading": "Sequence", "lines": [s["heading"] for s in spec["slides"]]})
    for s in content:
        if s.get("kind") == "section":
            para(bd.clean(s["heading"]), bd.TITLE_PT, bd.GREY_HEADING, center, before=170)
            doc.paragraphs[-1].runs[-1].add_break(WD_BREAK.PAGE)
            continue
        para(bd.clean(s["heading"]), bd.HEADING_SIZES[0], bd.GREY_HEADING, after=24)
        if s.get("kind") == "data":                       # KPIs and chart values as text
            for k in s.get("kpis") or []:
                p = para(f"{bd.clean(k['value'])}  {bd.clean(k['label'])}", bd.BODY_SIZES[0], bd.TIER_COLORS[0], after=10)
                p.runs[0].font.bold = True
            chart = s.get("chart")
            if chart:
                unit = bd.chart_unit(chart)
                for b in chart["bars"]:
                    para(f"{bd.clean(b['label'])}: {bd.fmt_value(b['value'], unit)}", bd.BODY_SIZES[0],
                         bd.TIER_COLORS[2], after=6)
        elif bd.is_tiered(s):                               # fact / proof / evidence, each lighter
            rows, _ = bd.tier_layout(s, bd.BODY_SIZES[0])
            for text, tier in rows:
                p = para(bd.clean(text), bd.BODY_SIZES[0], bd.TIER_COLORS[tier], after=16)
                if tier == 0:
                    p.runs[0].font.bold = True   # Word has no hairline outline; bold stands in
        else:
            lines = s.get("lines") or (s.get("left", []) + s.get("right", []))
            for line in lines:
                para(bd.clean(line), bd.BODY_SIZES[0], bd.BLACK, after=16)
        if bd.sources_text(s):
            para(bd.sources_text(s), bd.SOURCE_PT, bd.SOURCE_COLOR, before=18, after=0)
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
    with bd.use_style(spec.get("style")):     # same font & sizes in every format
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
    args, (lo, hi), style = bd.split_args(sys.argv[1:])
    if len(args) < 2:
        print(__doc__)
        sys.exit(2)
    spec = json.loads(Path(args[0]).read_text(encoding="utf-8-sig"))
    if style:
        spec["style"] = style
    errors = bd.validate(spec, lo, hi)
    if errors:
        print("STYLE CHECK FAILED - fix these and run again:")
        for e in errors:
            print("  - " + e)
        sys.exit(1)
    mode = "Pyramid" if spec.get("structure") == "pyramid" else "Standard"
    out = export_all(spec, args[1], stem=f"{bd.full_title(spec)} - {mode}")
    for k, v in out.items():
        print(k, ":", f"{len(v)} files" if isinstance(v, list) else v)


if __name__ == "__main__":
    main()
