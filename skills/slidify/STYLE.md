# Signature style

Taken from the owner's reference deck, `reference/CMPO Gov.pdf`. Open it whenever in doubt: it is the source of truth.

## Look

- **Background:** plain white. No colour fills, shapes, icons, stock images or decoration.
- **Greys only:** headings `#A6A6A6`, title and THANK YOU `#808080`, body text black.
- **Font:** Nirmala Text, regular weight (never bold).
- **Sizes:** title slide 36 pt (THANK YOU too). Slide headings 22-24 pt. Body lines 18-20 pt.
  The builder uses 24 and 20, and drops to 22 or 18 only when a line would not otherwise fit on one line. Body lines on a slide always share one size.
- **Heading:** top left, grey, short.
- **Body:** plain lines aligned left, with **no bullet markers**, generous space between lines.
- **Logos (optional):** small, top right and bottom right of content slides.

## Slide types

| Type | When | Spec |
|---|---|---|
| Title | Always first | `"title"`, optional `"date"` such as `"OCT, 2026"` |
| Sequence | Agenda listing the content headings | `"agenda": true` |
| Lines | Statements, one per line | `{"heading", "lines": [...]}` |
| Two columns | A list of short items, split down a black divider | `{"heading", "left": [...], "right": [...]}` |
| THANK YOU | Always last, grey capitals, centred | default; `"closing": ""` to omit |

## Writing

- Every point is **one line**. Never wrap and never split a thought across two lines.
- Minimum words: 12 per line at most; aim for 4-8.
- Headings are 1-4 word noun phrases: "Functions", "Priority Sectors", "Progress so far …".
- Start each line with a capital letter. No full stop at the end.
- Use `&` for "and".
- Pair related items on one line with ` | `: "Low Cost | Low HR", "Roads | Infrastructure".
- Plain, direct statements in the active voice: "CMPO will not replicate any function already assigned in the Government".
- Keep dates and figures compact: "13th Aug, 24", "200 M PKR", "~50 M PKR".
- 3-7 lines per slide. A list of more than 7 short items goes in two columns.
- 12 slides at most in total; 10-12 for a full talk.

## Reference examples

```
Functions
  Act as Secretariat to CM Policy Initiatives
  Work as Think-Tank for CM KP
  Repository of all Government Data for CM KP

CMPO Work Ethic            (two columns)
  High Value               | Notgovernment
  High Quality             | No hierarchy - Communal Workspace
  Low Cost | Low HR        | Do More with Less
```
