# Zentiva Presentation Generator

**Version:** v020 (template Brand_v002.pptx; see zentiva-prez-gen/SKILL.md for the changes)

Generate professional Zentiva-branded PowerPoint presentations programmatically using Python.

## Overview

This project provides a skill for generating PPTX presentations that adhere to Zentiva corporate branding guidelines. It preserves visual elements (gradients, logos, colors) by modifying the official Brand_v002.pptx template directly.

## Features

- **Template-based generation** - Preserves gradients, logos, and brand elements
- **Multiple slide types** - Title, Content, Two-Column, Text+Image, Quote, Conclusion
- **Hierarchical bullets** - Level-based styling with proper indentation and colors
- **Automatic elements** - Date on title slide, slide numbers in "current / total" format
- **PNG verification** - Export to PNG for visual comparison

## Project Structure

```
/home/clindevdep/AI/PrezGen/
├── README.md                           # This file
├── Templates/
│   └── presentation_template_v018.md   # Markdown input template
│
└── zentiva-prez-gen/
    ├── SKILL.md                        # Skill documentation
    ├── assets/
    │   └── Brand_v001.pptx             # Official Zentiva template
    ├── references/
    │   └── branding.md                 # Brand guidelines
    ├── scripts/
    │   └── generate_pptx.py            # Main generator script
    └── verification/
        └── test_output_v018.pptx       # Test output
```

## Quick Start

### Prerequisites

```bash
pip install python-pptx lxml
```

### Generate a Test Presentation

```bash
cd /home/clindevdep/AI/PrezGen/zentiva-prez-gen
python scripts/generate_pptx.py --test verification/test_output
```

Output: `verification/test_output_v018.pptx`

### Using the Python API

```python
import sys
sys.path.insert(0, '/home/clindevdep/AI/PrezGen/zentiva-prez-gen')
from scripts.generate_pptx import (
    generate_presentation,
    add_content_slide,
    add_conclusion_slide,
    ZENTIVA_DARK_BLUE,
    ZENTIVA_TEAL
)

# Define slides
slides = [
    {
        'type': 'title',
        'title': 'Q1 2026 Results',
        'subtitle': 'Quarterly Business Review'
    },
    {
        'type': 'content',
        'title': 'Key Highlights',
        'subtitle': 'Performance Overview',
        'content': [
            'Revenue up 15% YoY',
            ('Market expansion in CEE', 1),
            ('New product launches', 1),
            'Operational efficiency gains'
        ]
    },
    {
        'type': 'two_column',
        'title': 'Financial Overview',
        'content': ['Revenue: $2.4B', 'EBITDA: $680M'],
        'content2': ['Net Income: $340M', 'EPS: $1.42']
    }
]

# Generate presentation
generate_presentation(
    output_path='my_presentation.pptx',
    template_path='/home/clindevdep/AI/PrezGen/zentiva-prez-gen/assets/Brand_v001.pptx',
    slides_spec=slides
)
```

### Using the Markdown Template

1. Copy the template:
   ```bash
   cp /home/clindevdep/AI/PrezGen/Templates/presentation_template_v018.md my_presentation.md
   ```

2. Edit with your content (YAML format)

3. Generate (requires implementing markdown parser - see Advanced Usage)

## Slide Types

### Title Slide
```python
{'type': 'title', 'title': 'Main Title', 'subtitle': 'Optional Subtitle'}
```
- Full gradient background with centered white title
- Automatic date display (YYYY-MMM-DD)
- No slide number

### Content Slide
```python
{'type': 'content',
 'title': 'Slide Title',
 'subtitle': 'Green subtitle',
 'content': [
    'Level 0 item (blue)',
    ('Level 1 item (green)', 1),
    ('Level 2 item (green)', 2)
]}
```
- Hierarchical bullet points with color-coded levels
- Optional green subtitle

### Two-Column Slide
```python
{'type': 'two_column',
 'title': 'Comparison',
 'content': ['Left 1', 'Left 2'],
 'content2': ['Right 1', 'Right 2']}
```
- Side-by-side content areas
- Decorative gradient circle on right

### Text+Image Slide
```python
{'type': 'text_image',
 'title': 'Feature Title',
 'subtitle': 'Feature subtitle',
 'content': ['Point 1', ('Detail', 1)],
 'image': '/path/to/image.png'}
```
- Text content on left
- Image placeholder on right

### Quote Slide
```python
{'type': 'quote', 'title': 'Your impactful statement'}
```
- Full gradient background
- Centered white text

### Conclusion Slide
```python
{'type': 'conclusion',
 'title': 'Conclusions',
 'subtitle': 'Summary',
 'takeaways': ['Point 1', 'Point 2', 'Point 3']}
```
- Key takeaways with large bullet points

## Brand Colors

| Color | Hex | RGB | Usage |
|-------|-----|-----|-------|
| Dark Blue | #0C4160 | (12, 65, 96) | Titles, level 0 text, slide numbers |
| Teal | #00A98F | (0, 169, 143) | Subtitles, level 1+ text, accents |
| White | #FFFFFF | (255, 255, 255) | Title slide text, backgrounds |

## Bullet Hierarchy

| Level | Font Size | Text Color | Bullet Color |
|-------|-----------|------------|--------------|
| 0 | 20pt | Dark Blue | Dark Blue |
| 1 | 16pt | Teal | Teal |
| 2 | 14pt | Teal | Teal |

## Verification Workflow

Generate PNG screenshots for visual verification:

```bash
cd /home/clindevdep/AI/PrezGen/zentiva-prez-gen/verification

# Convert to PDF
libreoffice --headless --convert-to pdf test_output_v018.pptx

# Convert to PNG (150 DPI)
pdftoppm -png -r 150 test_output_v018.pdf test_output_v018

# View slides
ls test_output_v018-*.png
```

## API Reference

### Main Functions

| Function | Description |
|----------|-------------|
| `generate_presentation(output, template, slides)` | Generate full presentation |
| `generate_test_presentation(output, template)` | Generate test with all layouts |
| `get_versioned_filename(base, version)` | Create versioned filename |

### Slide Functions

| Function | Description |
|----------|-------------|
| `add_content_slide(prs, title, bullets, subtitle)` | Add content slide |
| `add_two_column_slide(prs, title, left, right)` | Add two-column slide |
| `add_text_image_slide(prs, title, subtitle, content, image)` | Add text+image slide |
| `add_conclusion_slide(prs, title, subtitle, takeaways)` | Add conclusion slide |

### Utility Functions

| Function | Description |
|----------|-------------|
| `add_slide_number(slide, current, total)` | Add "current / total" number |
| `move_slide_to_end(prs, index)` | Reorder slides |
| `hide_slide(prs, index)` | Hide slide from presentation |
| `hide_unused_placeholders(slide, keep)` | Remove placeholder ghosts |
| `set_bullet_format(para, char, color)` | Set explicit bullet styling |

## AI Skill Integration

This project is integrated as an ai-rulez skill:

**Skill Location:** `/home/clindevdep/AI/System/.ai-rulez/skills/zentiva-prez-gen/SKILL.md`

**Trigger Phrases:**
- "create Zentiva presentation"
- "generate Zentiva slides"
- "make Zentiva pptx"

## Troubleshooting

### "Click to add Text" appearing
Use `hide_unused_placeholders()` to remove unused OBJECT placeholders.

### Bullets not visible
Template layouts have `buNone` set. The generator uses explicit `set_bullet_format()`.

### Date not visible
Date positioned at (9.0", 6.2") in white 20pt bold. Requires dark gradient background.

### Slide numbers not at far right
Numbers positioned at 11.8" on 13.33" wide slide with right alignment.

### Gradients not preserved
Must modify existing template slides (0, 1, 2) rather than creating new ones from layouts.

## Version History

| Version | Date | Changes |
|---------|------|---------|
| v018 | 2026-02-15 | Date visibility fix, slide numbers at far right |
| v017 | 2026-02-14 | Subtitle spacing, bullet-text spacing, slide number format |
| v016 | 2026-02-14 | Green level 1+ bullets, date on title, slide numbers, conclusions |
| v015 | 2026-02-14 | Text+image slide, move/hide slides |
| v014 | 2026-02-14 | Subtitle integration into body placeholder |
| v013 | 2026-02-14 | Hierarchical font styling |
| v012 | 2026-02-14 | Hide unused placeholders |
| v011 | 2026-02-14 | Explicit bullet formatting |
| v010 | 2026-02-14 | Initial versioned release |

## License

Internal use - Zentiva corporate branding

## Author

Generated with Claude Code assistance

---

*Zentiva Presentation Generator v018*
