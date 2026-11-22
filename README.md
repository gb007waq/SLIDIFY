# Slidify

A Claude Code skill that makes presentations in one signature minimalist style.

## The style

- Greys only: `#333333`, `#595959`, `#A6A6A6`, `#F2F2F2`
- Nirmala Text: headings 22 pt, body 18 pt
- One-line bullets, 8 words max, 3-5 per slide
- 12 slides max, title and closing included

The builder refuses any deck that breaks these rules, so Claude rewrites until it fits.

## Install

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

This copies `skills/slidify` to `~/.claude/skills/slidify` and installs `python-pptx` and `pillow`.

## Use

In Claude Code: *"Slidify a presentation on onboarding new hires"*, or give it notes or a document.

Build by hand from a JSON spec:

```powershell
python skills\slidify\scripts\build_deck.py examples\sample.json examples\sample.pptx
```

See [examples/sample.json](examples/sample.json) for the spec format.
