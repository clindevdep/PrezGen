---
name: zentiva-prez-gen
version: v020
description: "Generate Zentiva-branded PowerPoint presentations from markdown or Python specs (Brand_v002.pptx)."
---

# Zentiva Presentation Generator (v020)

Generates Zentiva-branded PPTX decks with `python-pptx` from the official template **Brand_v002.pptx**
(16 template slides, 24 layouts). Template slides are reused for the title, statement and split slides
so gradients, photos and the logo stay intact; all other slides are added from the template layouts.

## Project location

```
/home/clindevdep/AI/PrezGen/
├── Templates/Brand_v002.pptx              # current template (Brand_v001 kept for old decks)
├── Templates/presentation_template_v020.md
└── zentiva-prez-gen/
    ├── scripts/generate_pptx.py           # generator (VERSION = v020)
    ├── assets/Brand_v002.pptx             # default template (copy)
    ├── references/branding.md             # colours, fonts, layouts
    └── verification/                      # rendered checks
```
Run with uv: `cd /home/clindevdep/AI/PrezGen/zentiva-prez-gen && uv run python scripts/generate_pptx.py …`

## Quick start

```bash
cp /home/clindevdep/AI/PrezGen/Templates/presentation_template_v020.md my_deck.md   # edit the ```yaml blocks
cd /home/clindevdep/AI/PrezGen/zentiva-prez-gen
uv run python scripts/generate_pptx.py --from-markdown my_deck.md my_deck_v001.pptx
uv run python scripts/generate_pptx.py --test verification/test_output             # all slide types
```

Python API:
```python
import sys; sys.path.insert(0, "/home/clindevdep/AI/PrezGen/zentiva-prez-gen")
from scripts.generate_pptx import generate_presentation
generate_presentation("out.pptx", None, [            # None = assets/Brand_v002.pptx
    {"type": "title", "title": "Tadalafil 20 mg FCT", "subtitle": "Feasibility"},
    {"type": "content", "title": "Key findings", "subtitle": "Regulatory",
     "content": ["EMA PSG: **fasting + fed**", ["Rev. 3, [PSG](https://www.ema.europa.eu/…)", 1]],
     "sources": ["[EMA PSG Rev. 3](https://…)", "[EPAR Mylan](https://…)"]},
    {"type": "table", "title": "CVw", "header": ["PK", "CVw", "N"], "rows": [["Cmax", "14.2%", "4"]]},
])
```
The deck follows the order of the spec list. Unused template slides are removed. Slide numbers
("n / total") are added to every visible slide except the title.

## Slide types

| type | built from | fields |
|---|---|---|
| `title` | template slide 1 (photo + gradient), date added | title, subtitle, image, date |
| `quote` | template slide 2 (gradient statement); further quotes → content slide | title |
| `split` | template slide 3 (panel + photo) | title, subtitle, image |
| `content` | layout `2_Title and Content` | title, subtitle, content |
| `two_column` | layout `17_Title and Content` (right gradient circle) | title, content, content2 |
| `highlight` | layout `25_Title and Content` (style of template slides 15–16) | title, content with `<<…>>` |
| `text_image` | layout `12_Title and Content` (photo right) | title, subtitle, content, image |
| `table` | layout `2_Title and Content` + table in the style of template slide 11 | title, subtitle, header, rows, col_widths, max_rows (default 12; longer tables continue on "(cont.)" slides) |
| `conclusion` | layout `2_Title and Content` | title, subtitle, takeaways |

Every slide also takes `sources` (footnote line above the footer: `[label](url)`, text, or `[label, url]`)
and `hidden: true`.

Bullets: `"text"` = level 0 (blue, bold-ish 20 pt), `["text", 1]` = level 1 (green 16 pt), `["text", 2]` = 14 pt.
Inline formatting in all text and table cells: `[label](https://…)` hyperlink, `**bold**`, `<<highlight>>` (teal).

## Brand (from the Brand_v002 theme and master)

| role | value |
|---|---|
| Titles, body text, table header | `#004876` (theme accent1) |
| Subtitles, emphasis | `#00BD8E` (theme accent2) |
| Master bullet colour | `#3EC89F` |
| Footnotes | `#323232` (dk1) |
| Fonts | Arial in the master text styles (theme minor/major: Calibri); tables in the template use Roboto |
| Other theme accents | `#1C6DC3`, `#B9207B`, `#FC4C02`, `#FFCC00` (charts) |
| Table style | Medium Style 2 – Accent 2 `{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}`, header row accent1 |

v019 and earlier used `#0C4160` / `#00A98F`, which are not in the template.

## Verification

```bash
cd verification && libreoffice --headless --convert-to pdf deck.pptx && pdftoppm -png -r 60 deck.pdf deck
```
Look at the PNGs: no "Click to add text", nothing outside the slide, date inside the title photo.

## Changes in v020 (2026-10-14)
- Default template Brand_v002.pptx. v019 removed only the first 14 template slides, so the two RWE example
  slides of Brand_v002 leaked into every deck; template slides are now removed by identity.
- Deck order = spec order (v019 put reused template slides — e.g. a quote — always at positions 1–3).
- All slide types available in `generate_presentation` (text_image and conclusion were test-only); new `table`,
  `sources`, `hidden`; links / bold / highlight in all text.
- Theme colours; title date moved inside the photo; images now inserted into picture placeholders
  (v019 checked a class name that never matched); `--from-markdown` implemented (it was documented, not built).

## Troubleshooting
- "Click to add text" visible → a placeholder was left: `hide_unused_placeholders(slide, keep_indices=[…])`.
- Table too long → it continues on "(cont.)" slides (`max_rows`); keep cells short.
- Old decks built on Brand_v001: pass the template explicitly (`--template …/Brand_v001.pptx`).
