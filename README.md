# SLIDIFY

A Claude Code agent that makes PowerPoint presentations in one signature style, taken from the owner's reference deck.

## The style

- White slides, grey headings, plain black one-liners with no bullet markers
- Nirmala Text: title slide 36 pt, headings 22-24 pt, body 18-20 pt
- One line per point, minimum words, `&` for "and", ` | ` to pair items
- Title slide with date, optional Sequence agenda, THANK YOU at the end
- 12 slides at most

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

## What's inside

| Path | What it is |
|---|---|
| `agents/slidify.md` | The SLIDIFY agent |
| `skills/slidify/SKILL.md` | Step-by-step instructions the agent follows |
| `skills/slidify/STYLE.md` | The signature style guide |
| `skills/slidify/reference/` | The reference deck the style comes from |
| `skills/slidify/scripts/build_deck.py` | Builds the `.pptx` and enforces the style |
| `skills/slidify/assets/` | Optional logos: `logo_title.png`, `logo_top.png`, `logo_corner.png` |
| `examples/` | A sample spec and the deck it builds |

Build by hand from a spec:

```powershell
python skills\slidify\scripts\build_deck.py examples\sample.json examples\sample.pptx
```

`install.ps1` is an alternative to the plugin: it copies the skill straight into `~/.claude/skills`.
