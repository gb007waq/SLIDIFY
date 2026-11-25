---
name: slidify
description: Make a PowerPoint presentation in the owner's signature style (white slides, grey headings, plain black one-liners without bullet markers, Nirmala Text at 36/24/20 pt, 12 slides max), modelled on their reference deck. Use whenever the user asks for a presentation, deck, slides or PPT, or says "slidify" something.
---

# Slidify

Turn any topic, notes or document into a `.pptx` in the owner's signature style.
The builder script enforces the look; your job is the content and its wording.

## Before writing anything

Read [STYLE.md](STYLE.md). It is the full style guide, taken from the owner's reference deck in `reference/`.
If you can read PDFs, look at `reference/CMPO Gov.pdf` too: match its tone, line length and slide types.

## The rules in short

- White background, Nirmala Text, nothing else.
- Title slide 36 pt. Grey headings 22-24 pt. Black body lines 18-20 pt. The builder picks the size within each range.
- **No bullet markers.** Each point is a plain one-liner with space around it.
- One line per point, 12 words at most (aim for 4-8). Never two lines.
- Pair related items with ` | `, and use `&` for "and". No full stops.
- **6-8 lines on every content slide**: 6 when the topic is simple, 7-8 when it needs more detail. Use two columns for lists of short items (up to 8 per column).
- First slide is the title (with date), then optionally a "Sequence" agenda. Last slide is THANK YOU.
- 12 slides by default (title, Sequence and THANK YOU included). If the user asks for fewer, 15-20, or an exact number, follow that and pass it to the builder with `--slides`.

## Workflow

1. **Understand the ask:** topic, audience and any source material. Ask one short question only if the topic is too vague to outline.
2. **Outline:** one point per slide, headings of 1-4 words.
3. **Write in the owner's voice:** direct statements, compact figures and dates, no filler. Cut every word that can go.
4. **Write the spec** as UTF-8 JSON:

   ```json
   {
     "title": "Deck title",
     "date": "NOV, 2026",
     "agenda": true,
     "slides": [
       {"heading": "Functions", "lines": ["Act as Secretariat to CM Policy Initiatives", "Work as Think-Tank for CM KP"]},
       {"heading": "Work Ethic", "left": ["High Value", "Low Cost | Low HR"], "right": ["Driven by Creativity", "Do More with Less"]}
     ]
   }
   ```

   `closing` defaults to THANK YOU. Use today's month and year for `date` unless told otherwise.

5. **Build** with the script in this skill's base directory (shown when the skill loads):

   ```
   python "<skill base directory>/scripts/build_deck.py" spec.json "Deck Title.pptx" --slides 12
   ```

   Pyramid mode: if the user asks for Pyramid, answer-first or an executive/recommendation deck, follow "Pyramid mode" in STYLE.md: set `"structure": "pyramid"`, give every slide a `"kind"` (summary, pillar, evidence, data, next, section, appendix), give pillar/evidence/next slides the three tiers `"fact"` (2), `"proof"` (4) and `"evidence"` (2) instead of `"lines"`, add `data` slides with `"kpis"` and/or `"chart"` where real figures exist, write complete-sentence headings and lines, and leave out the Sequence slide. The builder checks the pyramid order.

   Font style: Minimus 1 (Nirmala Text) is the default. If the user names another style (Minimus 2, Minimus 3), add `--style minimus2` etc.; styles are set in `fonts.json` and the builder says if one has no font yet.

   `--slides 8` means exactly 8 slides; `--slides 15-20` means 15 to 20. Leave it out for "at most 12".

   It needs `python-pptx`; Pillow gives exact line-fit checks. If they are missing, run `python -m pip install --user python-pptx pillow`.

6. **If the style check fails,** shorten the flagged lines (never split them) and build again until it prints `OK`.
7. **Other formats:** if the user wants PDF, slide pictures or Word as well, run
   `python "<skill base directory>/scripts/export_deck.py" spec.json <output folder>`.
   It writes `.pptx`, `.pdf`, `.docx`, one PNG per slide and a zip of the PNGs.
8. **Report** the file paths and list the slide headings.

## Logos

PNGs placed in `assets/` are added automatically: `logo_title.png` on the title slide, and `logo_top.png` and `logo_corner.png` at the top right and bottom right of content slides. Add them only when the owner asks for logos on the deck.

## Where to save

Save to the folder the user names. Otherwise use `~/Documents/Slidify`. Name the file after the deck title plus the presentation mode, e.g. `Solar for KP Schools - Pyramid.pptx` or `Employment Generation in KP - Standard.pptx`.
