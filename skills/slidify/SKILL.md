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
- 3-7 lines per slide. Use two columns for lists of short items.
- First slide is the title (with date), then optionally a "Sequence" agenda. Last slide is THANK YOU.
- 12 slides at most in total.

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
   python "<skill base directory>/scripts/build_deck.py" spec.json "Deck Title.pptx"
   ```

   It needs `python-pptx`; Pillow gives exact line-fit checks. If they are missing, run `python -m pip install --user python-pptx pillow`.

6. **If the style check fails,** shorten the flagged lines (never split them) and build again until it prints `OK`.
7. **Report** the `.pptx` path and list the slide headings.

## Logos

PNGs placed in `assets/` are added automatically: `logo_title.png` on the title slide, and `logo_top.png` and `logo_corner.png` at the top right and bottom right of content slides. Add them only when the owner asks for logos on the deck.

## Where to save

Save to the folder the user names. Otherwise use `~/Documents/Slidify`. Name the file after the deck title.
