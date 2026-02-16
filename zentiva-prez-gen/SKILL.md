---
name: zentiva-prez-gen
version: v019
description: |
  Generate Zentiva-branded PowerPoint presentations from markdown templates.
  Uses the official Brand_v001.pptx template to preserve gradients, logos, and visual elements.

  TRIGGERS - Use this skill when user says:
  - "create Zentiva presentation" / "generate Zentiva slides" / "make Zentiva pptx"
  - "create slides from markdown" / "convert md to pptx" / "presentation from template"
  - Any request for Zentiva-branded presentations or slide decks
---

# Zentiva Presentation Generator (v019)

This skill enables you to create branded Zentiva presentations programmatically using Python and the `python-pptx` library.

## Quick Start

```python
from scripts.generate_pptx import generate_presentation

slides = [
    {'type': 'title', 'title': 'Q1 2026 Results'},
    {'type': 'quote', 'title': 'Driving sustainable growth through innovation'},
    {'type': 'content', 'title': 'Key Highlights', 'content': [
        'Revenue up 15% YoY',
        ('Market expansion in CEE', 1),
        ('New product launches', 1),
        'Operational efficiency gains'
    ]},
    {'type': 'two_column', 'title': 'Financial Overview',
     'content': ['Revenue: $2.4B', 'EBITDA: $680M'],
     'content2': ['Net Income: $340M', 'EPS: $1.42']}
]

generate_presentation('output.pptx', 'assets/Brand_v001.pptx', slides)
```

## Supported Slide Types

### 1. Title Slide (`type: 'title'`)
Full-width background image with gradient overlay and centered white title. Includes automatic date display (YYYY-MMM-DD format) in bottom right.

```python
{'type': 'title', 'title': 'Presentation Title', 'subtitle': 'Optional Subtitle'}
```

### 2. Quote/Statement Slide (`type: 'quote'`)
Full gradient background (dark blue to teal) with centered white statement text.

```python
{'type': 'quote', 'title': 'Your impactful statement here'}
```

### 3. Content Slide (`type: 'content'`)
Standard slide with title and hierarchical bullet points.

```python
{'type': 'content', 'title': 'Slide Title', 'content': [
    'First level bullet',
    ('Second level indented', 1),
    'Back to first level'
]}
```

### 4. Two-Column Slide (`type: 'two_column'`)
Side-by-side content areas with decorative gradient circle on right.

```python
{'type': 'two_column', 'title': 'Comparison',
 'content': ['Left item 1', 'Left item 2'],
 'content2': ['Right item 1', 'Right item 2']}
```

### 5. Split Layout (`type: 'split'`)
Left gradient panel with text, right side image.

```python
{'type': 'split', 'title': 'Section Title', 'subtitle': 'Description'}
```

### 6. Highlight Content Slide (`type: 'highlight'`) - NEW in v019
Content slide with inline text highlighting. Key phrases are emphasized in teal color.
Use `<<text>>` syntax to mark portions that should be highlighted.

```python
{'type': 'highlight', 'title': 'Key Benefits', 'content': [
    '<<Primary advantage>> with supporting explanation',
    ('Detail with <<emphasis>> in the middle', 1),
    'Another point with <<multiple>> inline <<highlights>>',
    ('Sub-point containing <<key data>> and context', 1)
]}
```

**Highlight syntax:**
- Wrap text in `<<` and `>>` to highlight in teal
- Works at any bullet level
- Multiple highlights per line supported
- Level 0: Bold dark blue with teal highlights
- Level 1+: Regular dark blue with teal highlights

## Bullet Point Hierarchy

Content supports nested bullets using tuples:

```python
content = [
    'Level 0 bullet',           # First level (teal)
    ('Level 1 bullet', 1),      # Indented with bullet
    ('Level 2 bullet', 2),      # Further indented
    'Back to level 0'
]
```

## API Reference

### `generate_presentation(output_path, template_path, slides_spec)`

Generate a complete presentation from a list of slide specifications.

**Parameters:**
- `output_path` (str): Path to save the generated .pptx file
- `template_path` (str): Path to Brand_v001.pptx template
- `slides_spec` (list): List of slide dictionaries

**Slide Dictionary Keys:**
| Key | Type | Description |
|-----|------|-------------|
| `type` | str | 'title', 'quote', 'content', 'two_column', 'split', 'highlight' |
| `title` | str | Main title or statement text |
| `subtitle` | str | Secondary text (title/split slides) |
| `content` | list | Bullet points for left/main content |
| `content2` | list | Bullet points for right column (two_column only) |
| `image` | str | Path to image file (title/split slides) |

### `generate_test_presentation(output_path, template_path)`

Generate a test presentation with all layout types for verification.

```bash
python scripts/generate_pptx.py --test output
# Creates: output_v010.pptx
```

### `get_versioned_filename(base_name, version=None)`

Generate a versioned filename following the naming convention.

```python
from scripts.generate_pptx import get_versioned_filename, VERSION
print(get_versioned_filename("quarterly_report"))  # quarterly_report_v010.pptx
```

## Design Principles

The generator preserves Zentiva brand identity:

- **Colors**: Dark Blue (#0C4160), Teal (#00A98F)
- **Logo**: Automatically appears in footer (from template master)
- **Gradients**: Visual elements preserved from template slides
- **Typography**: Template fonts maintained
- **Layout**: Clean, minimalist with ample white space
- **Slide Numbers**: "current / total" format in dark blue, positioned at far right
- **Date**: YYYY-MMM-DD format on title slide in white

## Resources

- **Template**: `assets/Brand_v001.pptx`
- **Brand Guide**: `references/branding.md`
- **Generator Script**: `scripts/generate_pptx.py`
- **Verification**: `verification/` (PNG comparison infrastructure)

## Technical Notes

### Why Template Modification?

The generator modifies existing template slides rather than creating new ones from layouts. This approach ensures:

1. **Gradient preservation**: Background shapes and gradients are slide-specific, not layout properties
2. **Visual fidelity**: All decorative elements (circles, accent shapes) are preserved
3. **Brand consistency**: Master slide elements (logo, footer) automatically included

### Template Slide Mapping

| Index | Type | Visual Elements |
|-------|------|-----------------|
| 0 | Title | Background image, gradient overlay |
| 1 | Quote | Full gradient background |
| 2 | Split | Left gradient panel, right image area |
| 3+ | Various | Additional layouts (removed after use) |

New content and two-column slides are added from layouts, while title/quote/split slides modify existing template slides to preserve their visual elements.
