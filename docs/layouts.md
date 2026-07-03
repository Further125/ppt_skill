# Layout Reference

Every slide in a presentation (deck spec) must include a `layout` field. This document describes all supported layouts, their corresponding fields, and visual effects.

---

## `cover`

Full-screen title slide with a large title, subtitle, and optional date.

```json
{
  "layout": "cover",
  "title": "Presentation Title",
  "subtitle": "Subtitle",
  "date": "2026-06-28"
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `title` | string | Yes | Main title text |
| `subtitle` | string | No | Subtitle text |
| `date` | string | No | Date string |

---

## `toc`

Table of contents page with numbered items.

```json
{
  "layout": "toc",
  "title": "Table of Contents",
  "items": ["Introduction", "Methods", "Results", "Conclusion"]
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `title` | string | No | Section title (default: "Table of Contents") |
| `items` | string[] | Yes | List of section names |

---

## `title_content`

The most commonly used layout: title at the top and a content area below.

```json
{
  "layout": "title_content",
  "title": "Key Findings",
  "content": "Plain text paragraph or bulleted list",
  "accent_line": true
}
```

`content` can be:
- A plain string (rendered as a paragraph)
- A list of strings (rendered as bullets)
- An array of rich-text run dictionaries (see [advanced.md](advanced.md#rich-text))

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `title` | string | Yes | Slide title |
| `content` | string / string[] / object[] | No | Body text |
| `accent_line` | bool | No | Whether to show a blue underline (default: true) |

---

## `two_column`

Two-column text layout placed side by side.

```json
{
  "layout": "two_column",
  "title": "Comparison",
  "left_content": ["Point A1", "Point A2"],
  "right_content": ["Point B1", "Point B2"]
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `title` | string | Yes | Slide title |
| `left_content` | string / string[] / object[] | No | Left column text |
| `right_content` | string / string[] / object[] | No | Right column text |

---

## `chart`

Native PowerPoint charts (bar, column, line, pie, doughnut, area, scatter, radar, bubble, etc.).

```json
{
  "layout": "chart",
  "title": "Sales Growth",
  "description": "Year-over-year comparison",
  "chart_data": {
    "type": "bar",
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "values": [120, 150, 180, 210],
    "series_name": "Revenue"
  }
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `title` | string | Yes | Slide title |
| `description` | string | No | Subtitle text |
| `chart_data` | object | Yes | Chart configuration |

**`chart_data` fields:**
- `type`: `"bar"`, `"column"`, `"line"`, `"pie"`, `"doughnut"`, `"area"`, `"scatter"`, `"radar"`, `"bubble"`
- `categories`: string[] — X-axis labels
- `values`: number[] — Y-axis data
- `series_name`: string — Legend label

---

## `table`

Data table with styled header and alternating row colors.

```json
{
  "layout": "table",
  "title": "Framework Comparison",
  "description": "Key Metrics",
  "table_data": {
    "headers": ["Name", "Stars", "Language"],
    "rows": [
      ["LlamaFactory", "72K", "Python"],
      ["Unsloth", "67K", "Python"]
    ]
  }
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `title` | string | Yes | Slide title |
| `description` | string | No | Subtitle text |
| `table_data` | object | Yes | Table data |

**`table_data` fields:**
- `headers`: string[] — Column headers
- `rows`: string[][] — Data rows

---

## `timeline`

Timeline with milestones alternating above and below a horizontal line.

```json
{
  "layout": "timeline",
  "title": "Project Roadmap",
  "timeline_data": {
    "milestones": [
      {"date": "Q1 2025", "title": "Research", "desc": "Literature Review"},
      {"date": "Q2 2025", "title": "Design", "desc": "Architecture Design"},
      {"date": "Q3 2025", "title": "Build", "desc": "Feature Implementation"},
      {"date": "Q4 2025", "title": "Release", "desc": "Production Launch"}
    ]
  }
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `title` | string | Yes | Slide title |
| `timeline_data` | object | Yes | Milestone data |

**`timeline_data` fields:**
- `milestones`: array of `{date, title, desc}`

---

## `image_content`

Image on the left, text on the right.

```json
{
  "layout": "image_content",
  "title": "Architecture Diagram",
  "image_path": "diagram.png",
  "content": ["Key Point 1", "Key Point 2"]
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `title` | string | Yes | Slide title |
| `image_path` | string | Yes | Image file path |
| `content` | string / string[] / object[] | No | Body text |

---

## `quote`

Centered quote and attribution.

```json
{
  "layout": "quote",
  "quote": "Simplicity is the ultimate sophistication.",
  "author": "Leonardo da Vinci"
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `quote` | string | Yes | Quote text |
| `author` | string | No | Attribution |

---

## `team`

Team member cards with avatar, name, role, and description.

```json
{
  "layout": "team",
  "title": "Our Team",
  "team_data": {
    "members": [
      {"name": "Alice", "role": "Product Manager", "desc": "Product Strategy"},
      {"name": "Bob", "role": "Engineer", "desc": "Backend Systems"}
    ]
  }
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `title` | string | Yes | Slide title |
| `team_data` | object | Yes | Member data |

---

## `data_highlight`

Displays a single large number and label to emphasize key data.

```json
{
  "layout": "data_highlight",
  "title": "Market Size",
  "big_number": "$12B",
  "label": "Total Addressable Market"
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `title` | string | Yes | Slide title |
| `big_number` | string | Yes | Large number text |
| `label` | string | Yes | Label below the number |

---

## `process`

Horizontal step cards with numbering and titles.

```json
{
  "layout": "process",
  "title": "Workflow",
  "process_data": {
    "steps": [
      {"title": "Input"},
      {"title": "Process"},
      {"title": "Output"}
    ]
  }
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `title` | string | Yes | Slide title |
| `process_data` | object | Yes | Step data |

**`process_data` fields:**
- `steps`: array of `{title}` — each step becomes a numbered colored card

---

## `waterfall`

Waterfall chart showing cumulative changes.

```json
{
  "layout": "waterfall",
  "title": "Profit Analysis",
  "description": "From Revenue to Net Profit",
  "chart_data": {
    "categories": ["Revenue", "COGS", "Operating Expenses", "Net Profit"],
    "values": [100, -40, -30, 30]
  }
}
```

Requires `scripts/advanced_charts.py`.

---

## `funnel`

Funnel chart showing step-by-step decrease.

```json
{
  "layout": "funnel",
  "title": "Conversion Funnel",
  "chart_data": {
    "categories": ["Visitors", "Leads", "Customers"],
    "values": [1000, 300, 80]
  }
}
```

---

## `gantt`

Gantt chart for project scheduling.

```json
{
  "layout": "gantt",
  "title": "Project Schedule",
  "chart_data": {
    "tasks": [
      {"name": "Design", "start": 1, "duration": 3},
      {"name": "Build", "start": 4, "duration": 5}
    ]
  }
}
```

---

## `wordcloud`

Word cloud image generated from word frequency data.

```json
{
  "layout": "wordcloud",
  "title": "Keyword Cloud",
  "words": [
    {"text": "AI", "weight": 100},
    {"text": "ML", "weight": 80},
    {"text": "NLP", "weight": 60}
  ]
}
```

---

## `closing`

Thank you / closing slide.

```json
{
  "layout": "closing",
  "title": "Thank You for Listening",
  "subtitle": "Any Questions?"
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `title` | string | No | Closing text (default: "Thank You") |
| `subtitle` | string | No | Subtitle text |

---

## Constraint Layouts

Constraint layouts use a solver to automatically calculate font sizes and positions. They use the `_constraint` suffix, but the data format is the same as regular layouts.

### `cover_constraint`

```json
{"layout": "cover_constraint", "title": "Presentation Title", "subtitle": "Subtitle", "meta": "2026-06-28"}
```

### `section_constraint`

```json
{"layout": "section_constraint", "title": "Section Title", "chapter_num": "01", "subtitle": "Subtitle"}
```

### `title_content_constraint`

```json
{"layout": "title_content_constraint", "title": "Title", "content": ["Point A", "Point B"]}
```

### `two_column_constraint`

```json
{"layout": "two_column_constraint", "title": "Title", "left_content": ["Left content"], "right_content": ["Right content"]}
```

### `quote_constraint`

```json
{"layout": "quote_constraint", "quote": "Quote text", "author": "Author"}
```

### `data_highlight_constraint`

```json
{"layout": "data_highlight_constraint", "title": "Title", "big_number": "99", "description": ["Description 1", "Description 2"]}
```

### `three_column_constraint`

```json
{"layout": "three_column_constraint", "title": "Title", "col_1": ["Column 1"], "col_2": ["Column 2"], "col_3": ["Column 3"]}
```

### `chart_constraint`

```json
{"layout": "chart_constraint", "title": "Title", "description": ["Description"], "chart_data": {...}}
```

### `table_constraint`

```json
{"layout": "table_constraint", "title": "Title", "description": ["Description"], "table_data": {...}}
```

### `timeline_constraint`

```json
{"layout": "timeline_constraint", "title": "Title",
 "timeline_data": {"milestones": [{"date": "Q1", "title": "Milestone", "description": "Description"}]}}
```

### `team_constraint`

```json
{"layout": "team_constraint", "title": "Title",
 "team_data": {"members": [{"name": "Name", "role": "Role", "desc": "Description"}]}}
```

### `process_constraint`

```json
{"layout": "process_constraint", "title": "Title",
 "process_data": {"steps": [{"step": "01", "title": "Step", "desc": "Description"}]}}
```

### `tree_constraint`

```json
{"layout": "tree_constraint", "title": "Title",
 "tree_data": {"root": "Root Node", "children": [{"name": "Child Node"}]}}
```

### `image_text_split_constraint`

```json
{"layout": "image_text_split_constraint", "title": "Title", "image": "path.png", "content": ["Description"]}
```

---

## Common Fields (All Layouts)

Every slide can also contain the following fields:

| Field | Type | Description |
|-------|------|-------------|
| `accent_line` | bool | Whether to show a blue title underline (default: true) |
| `template_slide_name` | string | Template slide to clone |
| `shapes` | object[] | Free-form shapes array (see [advanced.md](advanced.md)) |
