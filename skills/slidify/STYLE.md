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
- **6-8 lines on every content slide.** Use 6 when the topic is simple and 7-8 when it needs more detail; never fewer than 6. A long list of short items goes in two columns (up to 8 per column). Spacing tightens slightly at 8 lines so they fit.
- **12 slides by default**, counting the title, Sequence and THANK YOU slides. When the owner asks for another number (fewer, 15-20, or any exact count), use that instead.
- The Sequence slide lists up to 8 headings in one column and up to 16 in two columns. With more than 16 content slides, leave it out.

## Pyramid mode (answer first)

Used when the owner picks **Pyramid**, or asks for an answer-first, executive, recommendation or "pyramid" presentation. Same look as above; only the structure and wording change. Set `"structure": "pyramid"` and give every slide a `"kind"`:

| Level | Slide `kind` | What it holds |
|---|---|---|
| **Top: the answer** | `summary` (first content slide, exactly one) | Heading = the core recommendation or conclusion, as one complete sentence. 3-6 lines: each pillar as a complete sentence, plus the ask or next step |
| **Middle: the pillars** | `pillar` (2-4, usually 3) | Heading = one distinct, non-overlapping reason the answer holds, as a complete-sentence claim. 6-8 lines of the strongest evidence |
| **Base: the evidence** | `evidence` (after the pillar it supports) | Heading = a complete-sentence takeaway. 6-8 lines of facts, figures, KPIs, results and comparisons |
| **Appendix** | `section` divider (heading "Appendix", no lines), then `appendix` slides (3-8 lines) | Extra detail that would slow the story. Only when there is genuine extra detail |

- Headings and lines are **complete sentences** (subject + verb), up to 14 words, still one line, no full stops.
- Headings state the "so what", never a topic label: "Solar cuts school power bills by 70 percent", not "Cost Savings".
- Lead with numbers where you have them.
- No Sequence slide: the executive summary is the agenda.

```
KP should solarise all public schools by 2028                      (summary)
  Solar cuts school power bills by about 70 percent
  Reliable power keeps classrooms running during load shedding
  Costs are recovered within four years from savings

Solar cuts school power bills by about 70 percent                  (pillar)
  Pilot schools in Mardan saw bills fall from 40k to 12k rupees
  ...
```

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
