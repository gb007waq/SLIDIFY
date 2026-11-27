"""Slidify deck builder.

Builds a .pptx in the Slidify signature style (see ../STYLE.md) from a JSON
spec, and refuses any spec that breaks the style rules.

Usage:
    python build_deck.py spec.json [output.pptx] [--slides N | --slides A-B] [--style minimus1]

Font style: --style minimus1 (or "style" in the spec). Styles live in ../fonts.json.

Structure: "structure": "pyramid" in the spec gives answer-first decks. Each slide
then has a "kind": summary (first, the main answer; its lines repeat every pillar
heading word-for-word plus at most 2 ask lines), pillar (2-4), evidence (after its
pillar), next (detail on the ask, after the pillars), section (an "Appendix"
divider, no lines) and appendix (supplementary detail only).
Headings and lines are complete sentences (up to 14 words, still one line).

Slide count (title, agenda and THANK YOU included): at most 12 by default.
--slides 8 means exactly 8; --slides 15-20 means between 15 and 20.

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
import threading
from contextlib import contextmanager
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ---- Signature style -------------------------------------------------------
FONT = "Nirmala Text"
TITLE_PT = 36               # THANK YOU and Appendix divider
TITLE_SLIDE_SIZES = (44, 40)  # title slide: 44, or 40 if a line would not fit
HEADING_SIZES = (24, 22)    # slide headings: 24, or 22 if 24 won't fit on one line
BODY_SIZES = (20, 18)       # body lines: 20, or 18 for the whole slide if needed
HEADING_PT = HEADING_SIZES[-1]
BODY_PT = BODY_SIZES[-1]

WHITE = RGBColor(0xFF, 0xFF, 0xFF)       # background
GREY_HEADING = RGBColor(0xA6, 0xA6, 0xA6)  # slide headings, date
GREY_CLOSING = RGBColor(0x80, 0x80, 0x80)  # THANK YOU
# Title slide: first part dark grey, second part (under it) lighter grey.
TITLE_COLORS = (RGBColor(0x2C, 0x31, 0x37), RGBColor(0x8A, 0x92, 0x9C))
TITLE_SEPARATORS = (": ", " | ", " - ", " – ", " — ", ", ")
BLACK = RGBColor(0x00, 0x00, 0x00)       # body lines, column divider

MAX_SLIDES = 12
MIN_LINES = 6            # every content slide: 6-8 lines, more when the topic needs detail
MAX_LINES = 8            # per slide (per column on two-column slides)
MAX_LINE_WORDS = 12
MAX_HEADING_WORDS = 5

# Pyramid structure: answer first, 2-4 pillars with evidence, appendix last.
# Headings and lines are complete sentences, still one line each.
PYRAMID_KINDS = ("summary", "pillar", "evidence", "data", "next", "section", "appendix")
PYRAMID_MAX_ASKS = 2      # summary lines that are not pillars (the ask / next step)

# Pyramid pillar / evidence / next slides use three tiers, each a step lighter
# ("Urban Smoke and Steel" greys): the fact (2, slightly thick), proving it (4), evidence (2).
TIERS = (("fact", 2), ("proof", 4), ("evidence", 2))
TIER_KINDS = ("pillar", "evidence", "next")
TIER_COLORS = (RGBColor(0x2C, 0x31, 0x37), RGBColor(0x5C, 0x66, 0x70), RGBColor(0x8A, 0x92, 0x9C))
SEMIBOLD_PT = 0.3             # hairline outline that makes the fact lines a little thick


# Pyramid data slides (the base): big KPI numbers and / or one grey bar chart.
# Sources line at the foot of every content slide: small, light grey, no encyclopedias.
SOURCE_PT = 9
SOURCE_COLOR = RGBColor(0x8A, 0x92, 0x9C)
SOURCE_COUNT = (1, 3)
SOURCE_TOP = Inches(6.92)
BANNED_SOURCES = re.compile(r"wiki|britannica|encyclop|encarta|world ?book|fandom|investopedia", re.I)

KPI_COUNT = (2, 4)
KPI_PT, KPI_LABEL_PT, CHART_PT = 36, 18, 18
CHART_BARS = (2, 8)
CHART_AXIS = RGBColor(0xD0, 0xD3, 0xD8)


def is_tiered(s):
    return any(k in s for k, _ in TIERS)
PYRAMID_PILLARS = (2, 4)
PYRAMID_HEADING_WORDS = 14
PYRAMID_LINE_WORDS = 14
PYRAMID_SHORT_MIN = 3     # executive summary and appendix slides may be shorter

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


# ---- Font styles (Minimus 1, Minimus 2, Minimus 3...) --------------------------------
STYLES_FILE = SKILL_DIR / "fonts.json"
_style_lock = threading.RLock()


def load_styles():
    data = json.loads(STYLES_FILE.read_text(encoding="utf-8"))
    return data.get("default", "minimus1"), data["styles"]


def style_ready(st):
    return bool(st.get("font"))


def style_file(st):
    """First font file that exists (file may be one path or a list), else the fallback.
    Returns (path or None, is_fallback)."""
    import os
    files = st.get("file") or []
    if isinstance(files, str):
        files = [files]
    for f in files:
        p = Path(os.path.expandvars(f))
        if not p.is_absolute():
            p = SKILL_DIR / p
        if p.exists():
            return p, False
    fb = st.get("fallback_file")
    if fb and Path(os.path.expandvars(fb)).exists():
        return Path(os.path.expandvars(fb)), True
    return None, True


@contextmanager
def use_style(name=None):
    """Apply a font style (font, file and sizes) while building or checking a deck."""
    global FONT, FONT_FILE, FONT_INDEX, TITLE_PT, HEADING_SIZES, BODY_SIZES, HEADING_PT, BODY_PT, TITLE_SLIDE_SIZES
    default, styles = load_styles()
    name = name or default
    if name not in styles:
        raise ValueError(f"Unknown font style {name!r}. Choose one of: {', '.join(styles)}")
    st = styles[name]
    if not style_ready(st):
        raise ValueError(f"{st.get('label', name)} has no font yet.")
    with _style_lock:
        saved = (FONT, FONT_FILE, FONT_INDEX, TITLE_PT, HEADING_SIZES, BODY_SIZES, TITLE_SLIDE_SIZES)
        f, fallback = style_file(st)
        FONT, FONT_FILE = st["font"], f
        FONT_INDEX = 0 if fallback else int(st.get("index", 0))
        TITLE_PT = st.get("title_pt", TITLE_PT)
        TITLE_SLIDE_SIZES = tuple(st.get("title_slide_pt", TITLE_SLIDE_SIZES))
        HEADING_SIZES = tuple(st.get("heading_pt", HEADING_SIZES))
        BODY_SIZES = tuple(st.get("body_pt", BODY_SIZES))
        HEADING_PT, BODY_PT = HEADING_SIZES[-1], BODY_SIZES[-1]
        try:
            yield st
        finally:
            FONT, FONT_FILE, FONT_INDEX, TITLE_PT, HEADING_SIZES, BODY_SIZES, TITLE_SLIDE_SIZES = saved
            HEADING_PT, BODY_PT = HEADING_SIZES[-1], BODY_SIZES[-1]


# ---- Line-fit measurement --------------------------------------------------
_measure_cache = {}


def MEASURE(text, pt):
    """Width in points of text in the current style's font, or None if it can't be measured."""
    try:
        from PIL import ImageFont
    except ImportError:
        return None
    if FONT_FILE is None or not Path(FONT_FILE).exists():
        return None
    key = (str(FONT_FILE), FONT_INDEX, pt)
    if key not in _measure_cache:
        _measure_cache[key] = ImageFont.truetype(str(FONT_FILE), pt * 10, index=FONT_INDEX)
    return _measure_cache[key].getlength(text) / 10


def fits_one_line(text, pt, box_width_emu):
    avail = Emu(box_width_emu - 2 * INSET).pt * 0.95
    width = MEASURE(text, pt)
    if width is not None:
        return width <= avail
    return len(text) * pt * 0.55 <= avail


# ---- Validation ------------------------------------------------------------
def clean(text):
    text = re.sub(r"\s+", " ", str(text)).strip()
    text = re.sub(r"\s*\|\s*", " | ", text)   # tidy pipe pairs
    return text.rstrip(".")                   # one-liners carry no full stop


def slide_count(spec):
    return (1 + (1 if spec.get("agenda") else 0) + len(spec.get("slides", []))
            + (1 if spec.get("closing", "THANK YOU") else 0))


MAX_AGENDA = 2 * MAX_LINES   # agenda goes to two columns after 7 headings


def parse_slides(arg):
    """'12' -> (12, 12); '15-20' -> (15, 20)."""
    a, _, b = str(arg).partition("-")
    lo, hi = int(a), int(b or a)
    if lo < 3 or hi < lo:
        raise ValueError(f"Bad slide count: {arg!r}")
    return lo, hi


def agenda_slide(spec):
    """The Sequence slide: one column up to 7 headings, then two columns."""
    heads = [s["heading"] for s in spec.get("slides", [])]
    if len(heads) <= MAX_LINES:
        return {"heading": "Sequence", "lines": heads}
    half = (len(heads) + 1) // 2
    return {"heading": "Sequence", "left": heads[:half], "right": heads[half:]}


def validate(spec, min_slides=1, max_slides=MAX_SLIDES):
    """Style check in the deck's own font style (spec['style'], default Minimus 1)."""
    try:
        with use_style(spec.get("style")):
            return _validate(spec, min_slides, max_slides)
    except ValueError as e:
        return [str(e)]


def _check_line(errors, label, text, limit, pt, width):
    if "\n" in str(text):
        errors.append(f"{label}: contains a line break: {text!r}")
    t = clean(text)
    if not t:
        errors.append(f"{label}: empty.")
    if len(t.split()) > limit:
        errors.append(f"{label}: {len(t.split())} words (max {limit}): {t!r}")
    if not fits_one_line(t, pt, width):
        errors.append(f"{label}: too long for one line: {t!r}")


def sources_text(s):
    srcs = [clean(x) for x in (s.get("sources") or []) if str(x).strip()]
    return ("Sources: " + "; ".join(srcs)) if srcs else ""


def source_errors(s, i):
    """Every content slide cites 1-3 real, non-encyclopedic sources that fit in two small lines."""
    srcs = s.get("sources")
    if not isinstance(srcs, list) or not SOURCE_COUNT[0] <= len([x for x in srcs if str(x).strip()]) <= SOURCE_COUNT[1]:
        return [f"Slide {i}: add 'sources' - {SOURCE_COUNT[0]}-{SOURCE_COUNT[1]} real sources "
                "(publisher, document, year), no encyclopedias."]
    errors = []
    for src in srcs:
        if BANNED_SOURCES.search(str(src)):
            errors.append(f"Slide {i}: source {clean(src)!r} is an encyclopedia; cite the primary source instead "
                          "(government report, official statistics, law, research paper or the owner's file).")
    width = MEASURE(sources_text(s), SOURCE_PT)
    avail = Emu(TEXT_W - 2 * INSET).pt * 0.95 * 2
    if width is not None and width > avail:
        errors.append(f"Slide {i}: sources are too long for two small lines; shorten them.")
    return errors


def slide_errors(s, i, structure="standard", heading_width=None):
    """Rules for one content slide (number i in the deck)."""
    errors = []
    pyramid = structure == "pyramid"
    kind = s.get("kind")
    heading_width = heading_width or TEXT_W
    if pyramid and kind == "section":                      # e.g. "Appendix" divider
        _check_line(errors, f"Slide {i} heading", s.get("heading", ""), 4, TITLE_PT, SLIDE_W - 2 * LEFT)
        return errors
    errors += source_errors(s, i)
    _check_line(errors, f"Slide {i} heading", s.get("heading", ""),
                PYRAMID_HEADING_WORDS if pyramid else MAX_HEADING_WORDS, HEADING_PT, heading_width)
    if pyramid and kind == "data":                           # KPIs and / or a grey bar chart
        return errors + data_errors(s, i)
    if pyramid and kind in TIER_KINDS:                       # 2 fact / 4 proof / 2 evidence
        for name, want in TIERS:
            got = s.get(name)
            if not isinstance(got, list) or len(got) != want:
                n = len(got) if isinstance(got, list) else 0
                errors.append(f"Slide {i}: '{name}' needs exactly {want} sentences (has {n}). "
                              "Tiered slides are 2 fact, 4 proof, 2 evidence.")
                continue
            for j, line in enumerate(got, start=1):
                _check_line(errors, f"Slide {i} {name} {j}", line, PYRAMID_LINE_WORDS, BODY_PT, TEXT_W)
        return errors
    if "lines" in s:
        cols = [("", s["lines"], TEXT_W)]
    else:
        cols = [("left ", s.get("left", []), COL_W), ("right ", s.get("right", []), COL_W)]
    low = PYRAMID_SHORT_MIN if pyramid and kind in ("summary", "appendix") else MIN_LINES
    total_lines = sum(len(c[1]) for c in cols)
    if total_lines < low:
        errors.append(f"Slide {i}: {total_lines} lines (needs {low}-{MAX_LINES}); add detail.")
    for name, lines, width in cols:
        if not lines:
            errors.append(f"Slide {i}: {name}column has no lines.")
        if len(lines) > MAX_LINES:
            errors.append(f"Slide {i}: {name}{len(lines)} lines (max {MAX_LINES}).")
        for j, line in enumerate(lines, start=1):
            _check_line(errors, f"Slide {i} {name}line {j}", line,
                        PYRAMID_LINE_WORDS if pyramid else MAX_LINE_WORDS, BODY_PT, width)
    return errors


def _norm(text):
    return re.sub(r"[^\w%]+", " ", clean(text).lower()).strip()


def pyramid_errors(slides):
    """Answer first, 2-4 pillars each followed by its evidence, next steps, then an optional appendix."""
    errors = _pyramid_order_errors(slides)
    if errors:
        return errors
    # The summary must state every pillar word-for-word, in order, plus at most 2 ask lines.
    summary = slides[0]
    lines = summary.get("lines") or (summary.get("left", []) + summary.get("right", []))
    said = [_norm(x) for x in lines]
    pillars = [s.get("heading", "") for s in slides if s.get("kind") == "pillar"]
    missing = [p for p in pillars if _norm(p) not in said]
    for p in missing:
        errors.append(f"Pyramid: pillar heading {clean(p)!r} must appear word-for-word as a line on the "
                      "executive summary.")
    if not missing:
        order = [said.index(_norm(p)) for p in pillars]
        if order != sorted(order):
            errors.append("Pyramid: list the pillars on the executive summary in the same order as the pillar slides.")
        asks = len(said) - len(pillars)
        if asks > PYRAMID_MAX_ASKS:
            errors.append(f"Pyramid: the executive summary has {asks} lines that are not pillars; keep at most "
                          f"{PYRAMID_MAX_ASKS} (the ask / next step) - every other line must be a pillar.")
    return errors


def _pyramid_order_errors(slides):
    errors, kinds = [], [s.get("kind") for s in slides]
    bad = sorted({str(k) for k in kinds if k not in PYRAMID_KINDS})
    if bad:
        errors.append(f"Pyramid: every slide needs a kind ({', '.join(PYRAMID_KINDS)}); got {', '.join(bad)}.")
        return errors
    if not kinds or kinds[0] != "summary":
        errors.append("Pyramid: the first content slide must be the executive summary (kind 'summary') "
                      "with the main answer as its heading.")
    if kinds.count("summary") != 1:
        errors.append("Pyramid: exactly one executive summary slide.")
    pillars = kinds.count("pillar")
    if not PYRAMID_PILLARS[0] <= pillars <= PYRAMID_PILLARS[1]:
        errors.append(f"Pyramid: {pillars} pillar slides; use {PYRAMID_PILLARS[0]}-{PYRAMID_PILLARS[1]} "
                      "(3 is usually best).")
    seen_pillar = in_next = in_appendix = False
    for n, k in enumerate(kinds[1:], start=2):
        if k == "summary":
            continue
        if k == "section":
            if in_appendix:
                errors.append("Pyramid: only one Appendix divider.")
            in_appendix = True
        elif k in ("pillar", "evidence", "data", "next") and in_appendix:
            errors.append(f"Pyramid: content slide {n} comes after the Appendix divider; use kind 'appendix'.")
        elif k in ("pillar", "evidence", "data") and in_next:
            errors.append(f"Pyramid: slide {n} ({k}) comes after the next-steps slides; put every pillar and "
                          "its evidence before the next steps.")
        elif k in ("evidence", "data") and not seen_pillar:
            errors.append(f"Pyramid: {k} slide {n} must follow the pillar it supports.")
        elif k == "next" and not seen_pillar:
            errors.append(f"Pyramid: next-steps slide {n} must come after the pillars.")
        elif k == "appendix" and not in_appendix:
            errors.append(f"Pyramid: appendix slide {n} needs an Appendix divider (kind 'section') before it.")
        if k == "pillar":
            seen_pillar = True
        if k == "next":
            in_next = True
    if in_appendix and kinds[-1] == "section":
        errors.append("Pyramid: the Appendix divider needs at least one appendix slide after it.")
    return errors


def _validate(spec, min_slides, max_slides):
    errors = []
    slides = spec.get("slides", [])
    structure = spec.get("structure", "standard")
    total = slide_count(spec)
    if total > max_slides or total < min_slides:
        want = f"exactly {min_slides}" if min_slides == max_slides else (
            f"between {min_slides} and {max_slides}" if min_slides > 1 else f"at most {max_slides}")
        errors.append(f"Deck has {total} slides (title, agenda & THANK YOU included); it must have {want}.")
    if not spec.get("title"):
        errors.append("Missing deck title.")
    if spec.get("agenda") and len(slides) > MAX_AGENDA:
        errors.append(f"Agenda would list {len(slides)} headings (max {MAX_AGENDA}); drop the agenda.")
    if structure == "pyramid":
        if spec.get("agenda"):
            errors.append("Pyramid: no Sequence slide - the executive summary is the agenda.")
        errors += pyramid_errors(slides)

    parts = title_parts(spec)
    for n, part in enumerate(parts, start=1):
        _check_line(errors, f"Title line {n}", part, 7, TITLE_SLIDE_SIZES[-1], SLIDE_W - 2 * LEFT)
    width = COL_W if spec.get("agenda") and len(slides) > MAX_LINES else TEXT_W
    for i, s in enumerate(slides, start=2 + (1 if spec.get("agenda") else 0)):
        errors += slide_errors(s, i, structure, width)
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


def write(paragraph, text, pt, color, align=None, semibold=False):
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
    if semibold:   # "a little thick": hairline outline in the text colour (lighter than bold)
        ln = rPr.makeelement(qn("a:ln"), {"w": str(int(Pt(SEMIBOLD_PT)))})
        fill = ln.makeelement(qn("a:solidFill"), {})
        fill.append(fill.makeelement(qn("a:srgbClr"), {"val": str(color)}))
        ln.append(fill)
        rPr.insert(0, ln)


def tier_lines(spec):
    """[(text, tier index)] for a tiered slide: fact, proof, evidence."""
    return [(t, n) for n, (name, _) in enumerate(TIERS) for t in spec.get(name, [])]


def tier_layout(spec, pt):
    """Space after each line (pt): one even spacing, the same as every other slide -
    the tiers are told apart by their grey, not by extra gaps."""
    rows = tier_lines(spec)
    gap = line_gap(len(rows), pt)
    return rows, [gap] * (len(rows) - 1) + [0]


def tiered_box(slide, spec, pt):
    rows, gaps = tier_layout(spec, pt)
    tf = textbox(slide, LEFT, BODY_TOP, TEXT_W, SLIDE_H - BODY_TOP - Inches(0.6))
    for k, ((text, tier), gap) in enumerate(zip(rows, gaps)):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        write(p, clean(text), pt, TIER_COLORS[tier], semibold=(tier == 0))


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


BODY_TOP = Inches(1.65)
BODY_BOTTOM = SLIDE_H - Inches(0.9)      # leaves clear room for the sources line


def line_gap(n, pt):
    """Space after each line: the usual 32 pt, tightened so 7-8 lines still fit the slide."""
    if n <= 1:
        return LINE_GAP_PT
    room = Emu(BODY_BOTTOM - BODY_TOP).pt - n * pt * 1.2
    return max(8, min(LINE_GAP_PT, room / (n - 1)))


def lines_box(slide, left, top, width, lines, pt, n=None):
    tf = textbox(slide, left, top, width, SLIDE_H - top - Inches(0.6))
    gap = line_gap(n or len(lines), pt)
    for k, line in enumerate(lines):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        write(p, clean(line), pt, BLACK)
    return tf


def title_parts(spec):
    """['Kill Tourism', 'Build Tourservation']: from "title" + "title2", or split once at
    a colon, bar, dash or comma in the title. One part when there is nothing to split."""
    title = clean(spec.get("title", "")).rstrip(",:;")
    second = clean(spec.get("title2", "")) if spec.get("title2") else ""
    if second:
        return [title, second]
    for sep in TITLE_SEPARATORS:
        if sep in title:
            a, b = title.split(sep, 1)
            if a.strip() and b.strip():
                return [a.strip(), b.strip()]
    return [title]


def full_title(spec):
    return " ".join(title_parts(spec))


def title_layout(spec):
    """Title-slide lines, their size and the box they sit in (EMU), tight together."""
    parts = title_parts(spec)
    width = SLIDE_W - 2 * LEFT
    pt = next((p for p in TITLE_SLIDE_SIZES if all(fits_one_line(x, p, width) for x in parts)),
              TITLE_SLIDE_SIZES[-1])
    line_h = Pt(pt * 1.08)
    box_h = int(line_h * len(parts)) + Inches(0.1)
    middle = Inches(4.55) if (ASSETS / "logo_title.png").exists() else Inches(3.75)
    return parts, pt, (LEFT, middle - box_h // 2, width, box_h)


def title_slide(prs, spec):
    s = blank(prs)
    logo = ASSETS / "logo_title.png"
    if logo.exists():
        pic = s.shapes.add_picture(str(logo), 0, Inches(2.2), height=Inches(1.6))
        pic.left = int((SLIDE_W - pic.width) / 2)
    parts, pt, (x, y, w, h) = title_layout(spec)
    tf = textbox(s, x, y, w, h, MSO_ANCHOR.MIDDLE)
    for n, part in enumerate(parts):
        p = tf.paragraphs[0] if n == 0 else tf.add_paragraph()
        p.line_spacing = 0.9          # minimum space between the two lines
        p.space_before = p.space_after = Pt(0)
        write(p, part, pt, TITLE_COLORS[min(n, 1)], PP_ALIGN.CENTER)
    date = spec.get("date")
    if date:
        tf = textbox(s, LEFT, Inches(6.3), SLIDE_W - 2 * LEFT, Inches(0.5), MSO_ANCHOR.MIDDLE)
        write(tf.paragraphs[0], clean(date).upper(), BODY_PT, GREY_HEADING, PP_ALIGN.CENTER)


def sources_footer(slide, spec):
    text = sources_text(spec)
    if text:
        tf = textbox(slide, LEFT, SOURCE_TOP, TEXT_W, Inches(0.45))
        write(tf.paragraphs[0], text, SOURCE_PT, SOURCE_COLOR)


def content_slide(prs, spec):
    s = blank(prs)
    heading(s, spec["heading"])
    top = BODY_TOP
    if is_tiered(spec):
        pt = pick_size(BODY_SIZES, [t for t, _ in tier_lines(spec)], TEXT_W)
        tiered_box(s, spec, pt)
    elif "lines" in spec:
        pt = pick_size(BODY_SIZES, spec["lines"], TEXT_W)
        lines_box(s, LEFT, top, TEXT_W, spec["lines"], pt)
    else:
        pt = pick_size(BODY_SIZES, spec["left"] + spec["right"], COL_W)
        n = max(len(spec["left"]), len(spec["right"]))
        lines_box(s, COL_LEFT_X, top, COL_W, spec["left"], pt, n)
        lines_box(s, COL_RIGHT_X, top, COL_W, spec["right"], pt, n)
        pitch = Pt(pt * 1.2 + line_gap(n, pt))
        bottom = top + int(pitch * n)
        d = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, DIVIDER_X, top - Inches(0.1),
                                   DIVIDER_X, bottom - Inches(0.2))
        d.line.color.rgb = BLACK
        d.line.width = Pt(1.25)
    sources_footer(s, spec)


def closing_slide(prs, text):
    s = blank(prs)
    tf = textbox(s, LEFT, Inches(3.3), SLIDE_W - 2 * LEFT, Inches(0.9), MSO_ANCHOR.MIDDLE)
    write(tf.paragraphs[0], clean(text).upper(), TITLE_PT, GREY_CLOSING, PP_ALIGN.CENTER)


def data_layout(s):
    """Where the KPI row and the chart sit on a Pyramid data slide (EMU)."""
    kpis, chart = s.get("kpis") or [], s.get("chart")
    body_h = BODY_BOTTOM - BODY_TOP
    if kpis and chart:
        kpi_top, chart_top = BODY_TOP, BODY_TOP + Inches(1.55)
    elif kpis:
        kpi_top, chart_top = BODY_TOP + int((body_h - Inches(1.3)) / 2), None
    else:
        kpi_top, chart_top = None, BODY_TOP
    cell_w = int(TEXT_W / max(1, len(kpis))) if kpis else 0
    return {"kpi_top": kpi_top, "cell_w": cell_w, "chart_top": chart_top,
            "chart_box": (LEFT, chart_top, TEXT_W, BODY_BOTTOM - chart_top) if chart else None}


def chart_unit(chart):
    u = str(chart.get("unit") or "").strip()
    return (" " + u) if u and u[0].isalnum() else u


def fmt_value(v, unit=""):
    v = float(v)
    txt = f"{v:,.0f}" if v.is_integer() else f"{v:,.1f}"
    return txt + unit


def data_errors(s, i):
    errors, kpis, chart = [], s.get("kpis") or [], s.get("chart")
    if not kpis and not chart:
        return [f"Slide {i}: a data slide needs 'kpis' (2-4 big numbers) and/or a 'chart'."]
    if kpis:
        if not KPI_COUNT[0] <= len(kpis) <= KPI_COUNT[1]:
            errors.append(f"Slide {i}: {len(kpis)} KPIs; use {KPI_COUNT[0]}-{KPI_COUNT[1]}.")
        cell = int(TEXT_W / max(1, len(kpis)))
        for j, k in enumerate(kpis, start=1):
            if not isinstance(k, dict) or not str(k.get("value", "")).strip() or not str(k.get("label", "")).strip():
                errors.append(f"Slide {i} KPI {j}: needs a 'value' (e.g. \"70%\") and a short 'label'.")
                continue
            _check_line(errors, f"Slide {i} KPI {j} value", k["value"], 3, KPI_PT, cell)
            _check_line(errors, f"Slide {i} KPI {j} label", k["label"], 6, KPI_LABEL_PT, cell)
    if chart:
        bars = chart.get("bars") if isinstance(chart, dict) else None
        if not isinstance(bars, list) or not CHART_BARS[0] <= len(bars) <= CHART_BARS[1]:
            errors.append(f"Slide {i}: the chart needs {CHART_BARS[0]}-{CHART_BARS[1]} bars.")
        else:
            slot = int(TEXT_W / len(bars))
            for j, b in enumerate(bars, start=1):
                v = b.get("value") if isinstance(b, dict) else None
                if isinstance(v, bool) or not isinstance(v, (int, float)) or v < 0:
                    errors.append(f"Slide {i} bar {j}: 'value' must be a number of 0 or more.")
                    continue
                _check_line(errors, f"Slide {i} bar {j} label", b.get("label", ""), 3, CHART_PT, slot)
    return errors


def data_slide(prs, spec):
    """Pyramid base: 2-4 big KPI numbers and / or one grey bar chart (a real PowerPoint chart)."""
    s = blank(prs)
    heading(s, spec["heading"])
    lay = data_layout(spec)
    for n, k in enumerate(spec.get("kpis") or []):
        x = LEFT + n * lay["cell_w"]
        tf = textbox(s, x, lay["kpi_top"], lay["cell_w"], Inches(0.8), MSO_ANCHOR.BOTTOM)
        write(tf.paragraphs[0], clean(k["value"]), KPI_PT, TIER_COLORS[0], semibold=True)
        tf = textbox(s, x, lay["kpi_top"] + Inches(0.8), lay["cell_w"], Inches(0.5))
        write(tf.paragraphs[0], clean(k["label"]), KPI_LABEL_PT, TIER_COLORS[2])
    chart = spec.get("chart")
    if chart:
        bars = chart["bars"]
        cd = CategoryChartData()
        cd.categories = [clean(b["label"]) for b in bars]
        cd.add_series(clean(spec["heading"]), [float(b["value"]) for b in bars])
        x, y, cx, cy = lay["chart_box"]
        ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, cd).chart
        ch.has_legend = False
        ch.has_title = False
        ch.font.name, ch.font.size = FONT, Pt(CHART_PT)
        ch.font.color.rgb = TIER_COLORS[1]
        plot = ch.plots[0]
        plot.gap_width = 70
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.number_format = f'#,##0.##"{chart_unit(chart)}"'.replace(".##", "") if all(
            float(b["value"]).is_integer() for b in bars) else f'#,##0.0"{chart_unit(chart)}"'
        dl.number_format_is_linked = False
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
        dl.font.size, dl.font.name = Pt(CHART_PT), FONT
        dl.font.color.rgb = TIER_COLORS[0]
        fill = plot.series[0].format.fill
        fill.solid()
        fill.fore_color.rgb = TIER_COLORS[2]
        va = ch.value_axis
        va.visible = False
        va.has_major_gridlines = False
        va.minimum_scale = 0
        ca = ch.category_axis
        ca.has_major_gridlines = False
        ca.format.line.color.rgb = CHART_AXIS
        ca.tick_labels.font.size, ca.tick_labels.font.name = Pt(CHART_PT), FONT
        ca.tick_labels.font.color.rgb = TIER_COLORS[1]
    sources_footer(s, spec)


def section_slide(prs, text):
    """Pyramid divider (e.g. Appendix): one grey word, centred, like the title slide."""
    s = blank(prs)
    tf = textbox(s, LEFT, Inches(3.3), SLIDE_W - 2 * LEFT, Inches(0.9), MSO_ANCHOR.MIDDLE)
    write(tf.paragraphs[0], clean(text), TITLE_PT, GREY_HEADING, PP_ALIGN.CENTER)


def build(spec, out_path):
    with use_style(spec.get("style")):
        _build(spec, out_path)


def _build(spec, out_path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    title_slide(prs, spec)
    if spec.get("agenda"):
        content_slide(prs, agenda_slide(spec))
    for s in spec.get("slides", []):
        if s.get("kind") == "section":
            section_slide(prs, s["heading"])
        elif s.get("kind") == "data":
            data_slide(prs, s)
        else:
            content_slide(prs, s)
    closing = spec.get("closing", "THANK YOU")
    if closing:
        closing_slide(prs, closing)
    prs.save(out_path)


def split_args(argv):
    """Pull '--slides N' or '--slides A-B' out of argv. Default: at most 12."""
    args, limits, style = [], (1, MAX_SLIDES), None
    it = iter(argv)
    for a in it:
        if a == "--slides":
            limits = parse_slides(next(it))
        elif a.startswith("--slides="):
            limits = parse_slides(a.split("=", 1)[1])
        elif a == "--style":
            style = next(it)
        elif a.startswith("--style="):
            style = a.split("=", 1)[1]
        else:
            args.append(a)
    return args, limits, style


def main():
    args, (lo, hi), style = split_args(sys.argv[1:])
    if not args:
        print(__doc__)
        sys.exit(2)
    spec_path = Path(args[0])
    spec = json.loads(spec_path.read_text(encoding="utf-8-sig"))
    if style:
        spec["style"] = style
    out = Path(args[1]) if len(args) > 1 else spec_path.with_suffix(".pptx")

    errors = validate(spec, lo, hi)
    if errors:
        print("STYLE CHECK FAILED - fix these and run again:")
        for e in errors:
            print("  - " + e)
        sys.exit(1)

    build(spec, out)
    print(f"OK: {slide_count(spec)} slides -> {out.resolve()}")


if __name__ == "__main__":
    main()
