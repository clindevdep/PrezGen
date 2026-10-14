# Zentiva Brand Guidelines for Presentations

This reference defines the visual standards for creating PowerPoint presentations branded for Zentiva.

## Color Palette (Brand_v002 theme, v020)

| Element | Hex Code | Source | Usage |
| :--- | :--- | :--- | :--- |
| **Zentiva Blue** | `#004876` | theme accent1 (dk2 `#004875`) | Titles, body text, table header, dark panels |
| **Zentiva Green** | `#00BD8E` | theme accent2 | Subtitles, emphasis, gradients |
| **Bullet Green** | `#3EC89F` | master body style | Bullet characters |
| **Light Teal** | `#E8F5F2` | — | Light rows / panels |
| **Text Grey** | `#323232` | theme dk1 | Footnotes |
| Accents | `#1C6DC3`, `#B9207B`, `#FC4C02`, `#FFCC00` | theme accent3–6 | Charts |

Fonts: Arial in the master title/body styles (title 28 pt bold accent1; subtitle placeholder 20 pt accent2;
body 28/24/20 pt); theme fonts Calibri Light / Calibri; template tables use Roboto 11 pt.
v019 and earlier documented `#0C4160` / `#00A98F`, which do not occur in the template.

## Template Layouts

The template `Brand_v002.pptx` (16 template slides) contains 24 slide layouts, the same as Brand_v001. Key layouts used by the generator:

| Layout Name | Purpose | Placeholders |
| :--- | :--- | :--- |
| `1_Title Slide` | Full-image title or quote slide | CENTER_TITLE, PICTURE |
| `2_Title and Content` | Standard content with bullets | TITLE, BODY (idx 14) |
| `3_Title Slide` | Split layout with left gradient panel | CENTER_TITLE, SUBTITLE, PICTURE |
| `17_Title and Content` | Two-column content layout | TITLE, BODY (idx 14, 15) |

## Slide Types

### Title Slide (Template Slide 0)
- Full-width background image with blue gradient overlay
- White centered title text
- Optional subtitle
- Zentiva logo in footer

### Quote/Statement Slide (Template Slide 1)
- Full gradient background (dark blue to teal)
- White centered statement text
- Zentiva logo in footer
- Page number in teal (bottom-right)

### Split Layout (Template Slide 2)
- Left panel: gradient background with title/subtitle
- Right panel: full-height image
- White text on gradient panel

### Content Slide (Added via layout)
- White background
- Dark blue title at top
- Teal bullet points with hierarchical indentation
- Zentiva logo in footer

### Two-Column Layout (Added via layout)
- Circular gradient accent shape on right
- Left column: teal bullets on white
- Right column: white text inside gradient circle
- Teal gradient accent at bottom

## Placeholder Reference

| Type | Index | Description |
| :--- | :--- | :--- |
| CENTER_TITLE | 0 | Main title (type 3) |
| SUBTITLE | - | Secondary text (type 4) |
| BODY | 14 | Primary content area |
| BODY | 15 | Secondary content (two-column) |
| PICTURE | 13 | Image placeholder (type 18) |
| FOOTER | 11 | Footer text |
| SLIDE_NUMBER | 12 | Page number |

## Design Principles

- **Logo**: Always visible in bottom-left footer (inherited from master)
- **Page Numbers**: Teal color, bottom-right corner
- **Gradients**: Dark blue (#0C4160) to teal (#00A98F) diagonal
- **Text on Dark**: White (#FFFFFF)
- **Text on Light**: Dark blue titles, teal body text
- **Simplicity**: Clean, minimalist layouts with ample white space
