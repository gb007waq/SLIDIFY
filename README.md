# SLIDIFY

A Claude Code agent that makes PowerPoint presentations in one signature style, taken from the owner's reference deck.

## The style

- White slides, grey headings, plain black one-liners with no bullet markers
- Nirmala Text: title slide 36 pt, headings 22-24 pt, body 18-20 pt
- One line per point, minimum words, `&` for "and", ` | ` to pair items
- Title slide with date, optional Sequence agenda, THANK YOU at the end
- 12 slides by default; fewer, 15-20 or any number on request

The full guide is in [skills/slidify/STYLE.md](skills/slidify/STYLE.md). The builder refuses any deck that breaks the rules, so the agent rewrites until it fits.

## Install (any computer)

1. Install Python, then the libraries:

   ```powershell
   python -m pip install --user python-pptx pillow
   ```

2. In Claude Code, add this repo and install the plugin:

   ```
   /plugin marketplace add gb007waq/SLIDIFY
   /plugin install slidify@slidify
   ```

3. Restart Claude Code.

## Use

- Ask: *"Slidify a presentation on employment generation"*, or paste notes or a document.
- Or call the agent directly: *"Use the slidify agent to make a deck on …"*.

Decks are saved to `Documents\Slidify` unless you name another folder.

## The SLIDIFY app

A web page on your computer for making decks without typing in Claude Code.

- **Put in:** a topic or notes (typed or spoken), plus Word (.docx), PDF, pictures or text files
- **Speak:** press Speak, say what the presentation is about in English or Urdu, press Stop. It makes the slides (needs Chrome or Edge)
- **Slides:** 12 (standard), fewer than 12, 15 to 20, or any custom number from 3 to 40
- **Get out:** PowerPoint, PDF, slide pictures (.zip) and Word, with a preview of every slide
- **Change something:** type what to change ("merge slides 4 & 5") and it rebuilds the deck

Start it (installs what it needs, then opens your browser at http://127.0.0.1:8765):

```powershell
powershell -ExecutionPolicy Bypass -File .\app\start.ps1
```

The app asks your installed Claude Code to write the slides, so you need Claude Code installed and signed in. There's no API key and no separate bill. Each deck takes about 1-2 minutes.

## What's inside

| Path | What it is |
|---|---|
| `agents/slidify.md` | The SLIDIFY agent |
| `skills/slidify/SKILL.md` | Step-by-step instructions the agent follows |
| `skills/slidify/STYLE.md` | The signature style guide |
| `skills/slidify/reference/` | The reference deck the style comes from |
| `skills/slidify/scripts/build_deck.py` | Builds the `.pptx` and enforces the style |
| `skills/slidify/scripts/export_deck.py` | Exports a deck as PDF, slide pictures and Word |
| `app/` | The SLIDIFY web app (`start.ps1` runs it) |
| `skills/slidify/assets/` | Optional logos: `logo_title.png`, `logo_top.png`, `logo_corner.png` |
| `examples/` | A sample spec and the deck it builds |

Build by hand from a spec:

```powershell
python skills\slidify\scripts\build_deck.py examples\sample.json examples\sample.pptx
```

`install.ps1` is an alternative to the plugin: it copies the skill straight into `~/.claude/skills`.
