# PPT Skill -- Complete User Guide

> A JSON-driven native PPTX auto-generation system.

---

## 1. Installation and Environment

### Dependencies

- Python 3.10+
- Pre-installed: python-pptx, Pillow, PyMuPDF
- Optional: LibreOffice (high-fidelity rendering)

### Verify Installation

```bash
cd ppt_skill
python scripts/build_pptx.py examples/demo_deck.json /tmp/test.pptx
```

---

## 2. Quick Start

```json
{
  "title": "AI Agent Overview",
  "slides": [
    { "layout": "cover", "title": "Cover" },
    { "layout": "title_content", "title": "Content", "content": ["A", "B"] },
    { "layout": "closing", "title": "Thank You" }
  ]
}
```

```bash
python scripts/build_pptx.py deck.json output.pptx
python scripts/build_pptx.py deck.json output_dark.pptx --theme dark
```

---

## 3. JSON Specification

### Top-Level Structure

| Field | Type | Required | Description |
|------|------|------|------|
| title | string | No | Deck title |
| subtitle | string | No | Subtitle |
| slides | array | Yes | Array of slides |

### Common Fields

| Field | Type | Default | Description |
|------|------|--------|------|
| layout | string | inferred | Layout type |
| accent_line | bool | true | Title underline |
| shapes | array | -- | Free-form shapes |
| hyperlink | string | -- | Hyperlink |

### Content Formats

- Plain string: `"content": "text"`
- Bullets: `"content": ["A", "B"]`
- Rich text: `"content": [{"text": "bold text", "bold": true}]`

---

## 4. Layout Reference

| Layout | Purpose | Key Fields |
|------|------|---------|
| cover | Cover slide | title, subtitle, date |
| toc | Table of contents | title, items |
| title_content | Title + body | title, content |
| two_column | Two-column | title, left_content, right_content |
| chart | Chart | title, description, chart_data |
| table | Table | title, description, table_data |
| timeline | Timeline | title, timeline_data |
| image_content | Image + text | title, image_path, content |
| quote | Quote | quote, author |
| team | Team | title, team_data |
| data_highlight | Big number | title, big_number, label |
| process | Process | title, process_data |
| tree | Tree diagram | title, tree_data |
| waterfall | Waterfall chart | title, description, chart_data |
| funnel | Funnel chart | title, description, chart_data |
| gantt | Gantt chart | title, description, chart_data |
| wordcloud | Word cloud | title, description, words |
| closing | Closing | title, subtitle |

---

## 5. Charts

Native chart types: column, bar, line, pie, doughnut, area, scatter, radar, bubble

Advanced charts: waterfall, funnel, gantt, wordcloud

---

## 6. Themes

Built-in themes: default, dark, warm, forest, cyber, minimal, rose, ocean

```bash
python scripts/build_pptx.py deck.json out.pptx --theme dark
```

Custom themes: create a JSON under `references/themes/` containing primary, secondary, accent, background, text, light_bg, font_title, font_body

---

## 7. Advanced Features

### Free-Form Shapes

Coordinate formats: `60%`, `2in`, `5cm`, `100px`, or EMU  
Shape types: rectangle, rounded_rectangle, oval, triangle, diamond, pentagon, hexagon, star, arrow_*, chevron, parallelogram, trapezoid, donut

### Image Masks

mask: circle, rounded_rectangle, oval, triangle, diamond, hexagon, star, pentagon, heart, cloud, sun, moon

### Animations

```bash
python scripts/build_pptx.py deck.json out.pptx --animate fade
```

Effects: fade, fly_in, wipe, zoom, bounce, appear, swivel, float

---

## 8. Export and Preview

```bash
# Render preview
python scripts/render_slides.py deck.pptx preview_dir --dpi 150

# PDF
python scripts/export.py deck.pptx --format pdf --output deck.pdf

# Long image
python scripts/export.py deck.pptx --format long_image --output deck.png

# HTML
python scripts/export.py deck.pptx --format html --output deck.html

# Snapshot
python scripts/snapshot.py deck.pptx --page 3 --output page3.png
```

---

## 9. Quality Checks

```bash
# QA check
python scripts/qa_check.py output.pptx

# Color contrast audit
python scripts/color_guard.py output.pptx --theme-json references/themes/dark.json --format markdown
```

color_guard checks: text-background contrast (WCAG), shape fill-text contrast, background glare  
Output levels: CRITICAL (< 2.0) / WARN (< 4.5) / INFO

---

## 10. CLI Reference

### build_pptx.py

| Option | Description |
|------|------|
| --template PATH | Custom template |
| --theme NAME | Theme or JSON path |
| --animate EFFECT | Animation |
| --no-stagger | Disable stagger |
| --auto-route | Auto-infer layout |
| --export FORMAT | pdf / long_image / html |
| --audit | Auto color audit |

### render_slides.py

| Option | Description |
|------|------|
| --dpi N | Render DPI (default 150) |
| --soft | Pure Python soft rendering |
| --pages LIST | Render only specified pages |

---

## 11. FAQ

**Q: Fonts display abnormally on Linux?**  
A: Install fonts-noto-cjk or copy the bundled NotoSansCJKsc-Regular.otf from the project.

**Q: Animations not visible in preview?**  
A: Animations must be viewed in PowerPoint / WPS to see the effect.

**Q: Chart colors too light or too dark?**  
A: Use `--theme` to specify a theme, or run color_guard to check contrast.

**Q: How to customize a template?**  
A: Use `scripts/template_analyzer.py` to analyze an existing PPTX, generate contract.json, then convert it to a schema.

---

## 12. Markdown → Deck Conversion

Convert a Markdown file directly into a JSON deck:

```bash
python scripts/md_to_deck.py input.md --output deck.json --theme dark
```

### Supported Markdown Syntax

| Syntax | Output |
|------|------|
| `# Title` | cover layout |
| `## Title` | title_content layout |
| `- Item` | Bulleted list |
| `1. Item` | Numbered list |
| `> Quote -- Author` | quote layout |
| `\|Table\|` | table layout |
| `\`\`\`python` | Code block (see below) |
| `\`\`\`chart` | chart/tree/highlight/wordcloud/team |

### Nested Lists

Use indentation to control hierarchy (2 spaces = 1 level):

```markdown
- Level 1 item
  - Level 2 item
    - Level 3 item
```

Renders as indented hierarchical bullets.

### Code Blocks

Markdown code blocks are automatically rendered as labeled code boxes:

```markdown
    ```python
    def hello():
        return "world"
    ```
```

Effect:
- A gray top tab bar showing the language name (e.g., `python`)
- Monospaced font (Courier New)
- Light gray background (#F3F4F6)
- Supported languages: `python`, `js`, `java`, `cpp`, `json`, `yaml`, `xml`, `sql`, `bash`, `sh`

### URL Auto-Hyperlinks

URLs in Markdown are automatically converted to blue underlined hyperlinks:

```markdown
- Official site: https://example.com
```

---

## 14. Constraint Layout System

### 14.1 Overview

Constraint layout is a core feature of ppt_skill. It uses a **constraint solver** to automatically calculate the maximum readable font size per page, and **attach chains** to automatically arrange element positions.

**Core features:**
- Binary search for maximum font size
- font_scale ratio control (large titles, small body text)
- attach relative positioning (avoids hard-coded coordinates)
- Automatic expand when content is sparse
- Automatic shrink to compress line spacing when content is dense
- Truncation and warning on overflow

### 14.2 Available Constraint Templates

| Template | Purpose | Key Fields |
|------|------|---------|
| `cover_constraint` | Cover | `title`, `subtitle`, `meta` |
| `section_constraint` | Section page | `title`, `subtitle`, `chapter_num` |
| `title_content_constraint` | Title + body | `title`, `content` |
| `two_column_constraint` | Two-column comparison | `title`, `left_content`, `right_content` |
| `quote_constraint` | Quote | `quote`, `author` |
| `data_highlight_constraint` | Big number | `title`, `big_number`, `description` |
| `three_column_constraint` | Three-column | `title`, `col_1`, `col_2`, `col_3` |
| `chart_constraint` | Chart | `title`, `description`, `chart_data` |
| `table_constraint` | Table | `title`, `description`, `table_data` |
| `timeline_constraint` | Timeline | `title`, `timeline_data` |
| `team_constraint` | Team | `title`, `team_data` |
| `process_constraint` | Process | `title`, `process_data` |
| `tree_constraint` | Tree diagram | `title`, `tree_data` |
| `image_text_split_constraint` | Left image, right text | `title`, `image`, `content` |

### 14.3 Data Format

All constraint templates use the same data format as fixed layouts, but with an additional set of templates suffixed by `_constraint`:

```json
{
  "layout": "title_content_constraint",
  "title": "Constraint Layout Demo",
  "content": ["Largest possible font", "Auto-arranged margins", "Auto-truncate on overflow"]
}
```

### 14.4 Special Data Formats

**timeline_data:**
```json
{
  "timeline_data": {
    "milestones": [
      {"date": "2024 Q1", "title": "Kickoff", "description": "Product kickoff"},
      {"date": "2024 Q2", "title": "MVP", "description": "Minimum viable product"}
    ]
  }
}
```

**team_data:**
```json
{
  "team_data": {
    "members": [
      {"name": "Zhang San", "role": "CEO", "desc": "Strategy"},
      {"name": "Member A", "role": "CTO", "desc": "Technology"}
    ]
  }
}
```

**process_data:**
```json
{
  "process_data": {
    "steps": [
      {"step": "01", "title": "Requirements", "desc": "Collect requirements"},
      {"step": "02", "title": "Design", "desc": "Solution design"}
    ]
  }
}
```

**tree_data:**
```json
{
  "tree_data": {
    "root": "Product Architecture",
    "children": [
      {"name": "Frontend", "children": [{"name": "Web"}, {"name": "App"}]},
      {"name": "Backend", "children": [{"name": "API"}]}
    ]
  }
}
```

### 14.5 Building Constraint Template Decks

```bash
python scripts/build_pptx.py deck.json output.pptx --theme minimal
```

Constraint templates support all themes.

---

## 15. Changelog

### 2026-07-01

- **Code block rendering**: Supports `[lang]...[/code]` tags with language tab bar and monospaced font
- **Nested lists**: Supports 4-level indentation controlled via `indent:N|` prefix
- **URL hyperlinks**: Auto-detected and converted to blue underlined links
- **Radar chart fix**: Falls back to native XL_CHART_TYPE.RADAR_MARKERS rendering
- **Code block background**: Tab bar + code area double-rectangle seamless join

---

*Document version: 2026-07-01*
