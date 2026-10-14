# Presentation Template v020

This template defines the structure for generating Zentiva-branded PowerPoint presentations.
Edit the sections below and use with the zentiva-prez-gen skill.

---

## Metadata

```yaml
title: "Your Presentation Title"
subtitle: "Optional Subtitle"
author: "Author Name"
date: auto  # Use 'auto' for current date or specify 'YYYY-MMM-DD'
output: "output_filename"  # Will be saved as output_filename_v020.pptx
```

---

## Slides

### Slide 1: Title

```yaml
type: title
title: "Main Presentation Title"
subtitle: "Subtitle or tagline"
```

### Slide 2: Content with Bullets

```yaml
type: content
title: "Key Highlights"
subtitle: "Overview of main points"
content:
  - "First main point (level 0 - blue)"
  - ["Supporting detail", 1]  # Level 1 - green
  - ["Another detail", 1]
  - "Second main point"
  - ["Supporting detail", 1]
  - "Third main point"
```

### Slide 3: Two Column Comparison

```yaml
type: two_column
title: "Comparison Title"
content:
  - "Left column point 1"
  - "Left column point 2"
  - "Left column point 3"
content2:
  - "Right column point 1"
  - "Right column point 2"
  - "Right column point 3"
```

### Slide 4: Text and Image

```yaml
type: text_image
title: "Feature Overview"
subtitle: "Key capabilities and benefits"
content:
  - "Main feature one"
  - ["Detail about feature one", 1]
  - "Main feature two"
  - ["Detail about feature two", 1]
image: "/path/to/image.png"  # Optional - leave empty for placeholder
```

### Slide 5: Quote/Statement

```yaml
type: quote
title: "Impactful statement or quote that summarizes a key message"
```

### Slide 6: Conclusions

```yaml
type: conclusion
title: "Conclusions"
subtitle: "Key takeaways from this presentation"
takeaways:
  - "First key takeaway point"
  - "Second key takeaway point"
  - "Third key takeaway point"
```

### Slide 7: Table (v020)

```yaml
type: table
title: "Within-subject variability"
subtitle: "Pooled CVw from EPARs and FDA reviews"
header: ["Parameter", "CVw", "Studies", "Source"]
rows:
  - ["Cmax", "14.2%", "4", "[EPAR Tadalafil Mylan](https://www.ema.europa.eu/en/medicines/human/EPAR/tadalafil-mylan)"]
  - ["AUC", "12.3%", "4", "EPAR + FDA review"]
col_widths: [2, 1, 1, 4]   # optional, relative
sources:
  - "[Feasibility workspace](https://library.clindevdep.com/)"
```

### Slide 8: Highlight (v019)

```yaml
type: highlight
title: "Why it matters"
content:
  - "<<Key phrase>> in teal, the rest in blue"
  - ["Detail with a [link](https://www.zentiva.com)", 1]
```

### Inline formatting (all text, v020)

- `[label](https://…)` → hyperlink (underlined)
- `**bold**` → bold
- `<<highlight>>` → teal emphasis
- `sources:` on any slide → small footnote line with the links above the footer
- `hidden: true` → slide kept but hidden

---

## Usage Instructions

### Quick Start

1. Copy this template:
   ```bash
   cp /home/clindevdep/AI/PrezGen/Templates/presentation_template_v020.md my_presentation.md
   ```

2. Edit the YAML blocks with your content

3. Generate the presentation:
   ```bash
   cd /home/clindevdep/AI/PrezGen/zentiva-prez-gen
   uv run python scripts/generate_pptx.py --from-markdown ../Templates/my_presentation.md output.pptx
   ```

### Content Guidelines

#### Bullet Point Levels
- **Level 0** (no bracket): Blue text, blue bullet, larger font (20pt)
- **Level 1** `["text", 1]`: Green text, green bullet, medium font (16pt)
- **Level 2** `["text", 2]`: Green text, green bullet, smaller font (14pt)

#### Slide Types Reference

| Type | Purpose | Required Fields |
|------|---------|-----------------|
| `title` | Opening slide with date | title, subtitle (optional) |
| `content` | Bulleted content | title, content, subtitle (optional) |
| `two_column` | Side-by-side comparison | title, content, content2 |
| `text_image` | Text left, image right | title, subtitle, content, image (optional) |
| `quote` | Full-screen statement | title |
| `conclusion` | Summary/takeaways | title, takeaways, subtitle (optional) |

#### Best Practices

1. **Keep titles concise** - Max 6-8 words
2. **Limit bullets** - 3-5 main points per slide
3. **Use hierarchy** - Level 1 for supporting details only
4. **Be consistent** - Similar slides should have similar structure
5. **Professional tone** - Clear, evidence-based language

---

## Brand Elements (Automatic)

The generator automatically applies:
- **Zentiva Dark Blue** (#0C4160) - Titles, level 0 text, slide numbers
- **Zentiva Teal** (#00A98F) - Subtitles, level 1+ text
- **Date** - YYYY-MMM-DD format on title slide
- **Slide Numbers** - "current / total" format, bottom right
- **Logo** - From template master slide

---

## Example: Complete Presentation

```yaml
# Metadata
title: "Q1 2026 Business Review"
subtitle: "Quarterly Performance Analysis"
author: "Business Team"
date: auto
output: "q1_2026_review"

# Slides
slides:
  - type: title
    title: "Q1 2026 Business Review"
    subtitle: "Driving Growth Through Innovation"

  - type: content
    title: "Financial Highlights"
    subtitle: "Strong performance across all metrics"
    content:
      - "Revenue: $2.4B (+15% YoY)"
      - ["Organic growth: 12%", 1]
      - ["Acquisitions: 3%", 1]
      - "EBITDA margin: 28.3%"
      - "Net income: $340M"

  - type: two_column
    title: "Regional Performance"
    content:
      - "Europe: $1.2B"
      - "Growth: +18%"
      - "Market share: 24%"
    content2:
      - "Asia-Pacific: $800M"
      - "Growth: +22%"
      - "New markets: 3"

  - type: text_image
    title: "Product Pipeline"
    subtitle: "Innovation driving future growth"
    content:
      - "5 new products launched"
      - ["3 in cardiology", 1]
      - ["2 in neurology", 1]
      - "Pipeline value: $1.8B"

  - type: conclusion
    title: "Conclusions"
    subtitle: "Key messages for stakeholders"
    takeaways:
      - "Strong financial performance with 15% revenue growth"
      - "Successful expansion in Asia-Pacific markets"
      - "Robust product pipeline supporting future growth"
```

---

*Template Version: v018 - Compatible with zentiva-prez-gen v018*
