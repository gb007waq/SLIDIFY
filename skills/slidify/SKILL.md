---
name: slidify
description: Make a PowerPoint presentation in the owner's signature minimalist style (grey palette, Nirmala Text, one-line bullets, 12 slides max). Use whenever the user asks for a presentation, deck, slides or PPT, or says "slidify" something.
---

# Slidify

Turn any topic, notes, document or brief into a `.pptx` in one fixed signature style.
The builder script enforces the style; your job is the content.

## The style (non-negotiable)

- **Palette:** greys only - `#333333`, `#595959`, `#A6A6A6`, `#F2F2F2`. No other colours, no images, no icons.
- **Font:** Nirmala Text everywhere. Headings 22 pt, body 18 pt. No other sizes.
- **Bullets:** every point is one line, a short one-liner of **8 words or fewer**. Never two lines.
- **Headings:** 6 words or fewer.
- **Per slide:** 3-5 bullets.
- **Length:** 12 slides maximum in total (title and closing slides count). Aim for 10-12 for a full talk; use fewer for small topics.
- **Wording:** minimum words. Cut articles, filler and hedging. No full stops at line ends. No sub-bullets.

## Workflow

1. **Understand the ask.** Topic, audience and any source material. If the user gave none of these and the topic is too vague to outline, ask one short question; otherwise just proceed.
2. **Outline.** Title slide, then up to 10 content slides, then an optional closing slide. Each content slide makes one point.
3. **Write tight.** Draft each bullet, then shorten it. Prefer nouns and verbs: "Revenue up 20% in Q3", not "We saw that revenue increased by about 20% during Q3."
4. **Write the spec** as JSON (UTF-8) next to where the deck should go:

   ```json
   {
     "title": "Deck title",
     "subtitle": "Optional one-liner",
     "slides": [
       {"heading": "Slide heading", "bullets": ["One line", "Another line"]}
     ],
     "closing": "Thank you"
   }
   ```

5. **Build** with the script in this skill's folder:

   ```
   python "<this skill's folder>/scripts/build_deck.py" spec.json output.pptx
   ```

   The skill folder is normally `~/.claude/skills/slidify`. Needs `python-pptx` (`pip install python-pptx`); Pillow gives exact line-fit checks.

6. **If the style check fails**, rewrite the flagged lines shorter (never split them into two lines) and run again until it prints `OK`.
7. **Report** the path to the `.pptx` and list the slide headings in one short list.

## Where to save

Save to the folder the user names. If none is named, use their current project folder, or `~/Documents/Slidify` if there is none. Name the file after the deck title.
