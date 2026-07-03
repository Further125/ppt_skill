# PPT Skill — JSON-Driven PowerPoint Generator

A powerful, template-based PowerPoint generation system. Write a JSON spec, get a polished `.pptx`.

[Chinese Docs](README_CN.md)

## Features

- **Constraint Layout System**: Binary-search-based font sizing with attach-chain auto-positioning. Text expands to fill available space while staying readable; overflows trigger graceful shrink or truncation warnings.
- **Declarative Layout Schemas**: JSON-defined layouts with percentage coordinates and attach-based relative positioning.
- **Intent Router**: LLM can omit `layout` field — system auto-infers from content fields (`--auto-route`).
- **Relative Coordinates**: Use `%`, `px`, `cm`, `mm`, `in` or predefined regions instead of raw EMU.
- **Template Analysis**: Extract visual contract (colors, fonts, layout structure) from any PPTX template.
- **Schema Auto-Generation**: Convert analyzed template into declarative layout schema.
- **Quick Preview**: Single-slide JSON to PNG rendering for LLM visual feedback loops.
- **20+ slide layouts**: cover, title_content, two_column, chart, table, timeline, process, team, quote, data_highlight, funnel, waterfall, gantt, wordcloud, image_content, closing, and more.
- **Theme engine**: switch color palettes via `--theme`.
- **Auto-fit text**: font sizes automatically adjusted to fit shapes.
- **Rich text mixing**: bold, color, size per run inside any text box.
- **Free-form shapes**: place text, images, shapes, charts anywhere on a slide.
- **Image masking & cropping**: circle masks, rounded rectangles, rectangular crops, z-order control.
- **Native charts & tables**: bar, line, pie, doughnut, radar, bubble, etc.
- **Export tools**: PDF, long image, HTML slideshow.
- **Snapshot tool**: capture any page or clipped region from a PPTX.

## Documentation

- [docs/QUICK_START.md](docs/QUICK_START.md) — Quick start
- [docs/USAGE.md](docs/USAGE.md) — Usage guide
- [docs/user-guide.md](docs/user-guide.md) — Detailed user guide
- [docs/layouts.md](docs/layouts.md) — Layout reference
- [docs/THEMES.md](docs/THEMES.md) — Theme system
- [docs/TEMPLATES.md](docs/TEMPLATES.md) — Template system
- [docs/advanced.md](docs/advanced.md) — Advanced features
- [docs/architecture.md](docs/architecture.md) — Architecture design
- [docs/comparison.md](docs/comparison.md) — Comparison with other tools
- [docs/EXAMPLES.md](docs/EXAMPLES.md) — Example gallery
- [docs/FAQ.md](docs/FAQ.md) — Frequently asked questions
- [docs/TESTING.md](docs/TESTING.md) — Testing guide
- [README_CN.md](README_CN.md) — Chinese version

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Generate a PPT from JSON
python scripts/build_pptx.py examples/demo_deck.json output/my_deck.pptx

# Auto-route layouts (omit layout field in JSON)
python scripts/build_pptx.py examples/auto_route_test.json output/auto.pptx --auto-route

# Quick single-slide preview for LLM visual feedback
python scripts/quick_preview.py examples/preview_test.json --output preview.png --auto-route

# Analyze a template and extract visual contract
python scripts/template_analyzer.py templates/base_template.pptx --output contract.json

# Convert contract to declarative schema
python scripts/contract_to_schema.py contract.json --layout cover --output schema.json

# With a theme
python scripts/build_pptx.py examples/demo_deck.json output/my_deck.pptx --theme ocean_tech

# Export to PDF
python scripts/export.py output/my_deck.pptx --format pdf --output my_deck.pdf

# Snapshot page 3
python scripts/snapshot.py output/my_deck.pptx --page 3 --output page3.png
```

## Project Structure

```
ppt_skill/
├── scripts/                  # Core scripts
│   ├── build_pptx.py         # Core builder (JSON → PPTX)
│   ├── layout_schema.py      # Declarative schema engine
│   ├── intent_router.py      # Auto-infer layout from content
│   ├── template_analyzer.py  # Extract visual contract from PPTX
│   ├── contract_to_schema.py # Convert contract to schema
│   ├── quick_preview.py      # Single-slide JSON → PNG preview
│   ├── export.py             # PDF / long_image / HTML export
│   ├── snapshot.py           # Page/region capture
│   ├── theme_engine.py       # Color palette engine
│   ├── text_fitter.py        # Auto-fit font size calculator
│   ├── animator.py           # Animation helper
│   └── advanced_charts.py    # Waterfall, funnel, gantt, wordcloud
├── templates/                # PPTX templates
│   └── base_template.pptx
├── references/               # Layout schemas and themes
│   ├── layout_schemas/
│   └── themes/
├── examples/                 # Example JSON specs
├── tests/                    # Test fixtures
├── docs/                     # English documentation
│   ├── QUICK_START.md
│   ├── USAGE.md
│   ├── LAYOUTS.md
│   ├── user-guide.md
│   ├── architecture.md
│   └── advanced.md
├── docs/zh/                  # Chinese documentation
└── output/                   # Output directory
```

## JSON Deck Spec

A deck spec is a JSON object with `title`, optional `subtitle`, and a `slides` array.

```json
{
  "title": "My Presentation",
  "subtitle": "Generated with ppt_skill",
  "slides": [
    {
      "layout": "cover",
      "title": "Hello World",
      "subtitle": "A demo deck"
    },
    {
      "layout": "title_content",
      "title": "Key Points",
      "content": [
        "First bullet point",
        "Second bullet point",
        "Third bullet point"
      ]
    }
  ]
}
```

## Layouts

| Layout | Description | Key Fields |
|---|---|---|
| `cover` | Title slide | `title`, `subtitle`, `date` |
| `toc` | Table of contents | `title`, `items` |
| `title_content` | Title + body text | `title`, `content` |
| `two_column` | Two-column text | `title`, `left_content`, `right_content` |
| `chart` | Chart slide | `title`, `description`, `chart_data` |
| `table` | Data table | `title`, `description`, `table_data` |
| `timeline` | Horizontal timeline | `title`, `timeline_data` |
| `image_content` | Image + text | `title`, `image_path`, `content` |
| `quote` | Quote block | `quote`, `author` |
| `team` | Team member cards | `title`, `team_data` |
| `data_highlight` | Big number highlight | `title`, `big_number`, `label` |
| `process` | Step process | `title`, `process_data` |
| `waterfall` | Waterfall chart | `title`, `description`, `chart_data` |
| `funnel` | Funnel chart | `title`, `description`, `chart_data` |
| `gantt` | Gantt chart | `title`, `description`, `chart_data` |
| `wordcloud` | Word cloud | `title`, `description`, `words` |
| `closing` | Thank you slide | `title`, `subtitle` |

See [docs/layouts.md](docs/layouts.md) for full details.

## Constraint Layout System

A constraint solver that automatically computes the largest readable font size for every slide, then positions all elements via attach-chains.

1. **Binary search** finds the maximum base font size that fits all textboxes.
2. **Font scaling** applies per-shape ratios.
3. **Attach chains** position shapes relative to each other.
4. **Expand** grows short text to fill whitespace; **shrink** reduces line-spacing when content is dense.
5. **Overflow guard** truncates overflowing text and prints warnings.

### Constraint templates

| Template | Description |
|---|---|
| `cover_constraint` | Title + subtitle + meta, auto-sized |
| `section_constraint` | Chapter divider with number |
| `title_content_constraint` | Title + bullet list |
| `two_column_constraint` | Side-by-side comparison |
| `quote_constraint` | Large quote + attribution |
| `data_highlight_constraint` | Big number + description |
| `three_column_constraint` | Three equal columns |
| `chart_constraint` | Title + native chart |
| `table_constraint` | Title + data table |
| `timeline_constraint` | Horizontal timeline |
| `team_constraint` | Member cards |
| `process_constraint` | Step process cards |
| `tree_constraint` | Hierarchical tree diagram |
| `image_text_split_constraint` | Left image + right text |

## Declarative Layout Schemas

Layouts can be defined as JSON schemas with percentage coordinates:

```json
{
  "name": "title_content",
  "shapes": [
    {
      "role": "title",
      "type": "textbox",
      "region": {"left": "5%", "top": "5%", "width": "90%", "height": "12%"},
      "style": "heading-1",
      "content_source": "title",
      "adaptive": {"strategy": "shrink", "min_size": 28}
    }
  ]
}
```

## Auto-Routing

Omit the `layout` field — the system infers it from content:

```json
{"quote": "Best code is no code", "author": "Dev"}        → quote
{"big_number": "99.9%", "label": "Uptime"}                → data_highlight
{"left_content": ["A"], "right_content": ["B"]}           → two_column
{"title": "Hello", "subtitle": "World"}                    → cover
```

```bash
python scripts/build_pptx.py deck.json output.pptx --auto-route
```

## Advanced Features

### Rich Text

Content can be a string, a bullet list, or an array of styled runs:

```json
{
  "type": "text",
  "content": [
    {"text": "Bold red text", "bold": true, "color": "#FF0000", "size": 24},
    {"text": " normal text"}
  ]
}
```

### Free-Form Shapes

```json
{
  "layout": "title_content",
  "title": "Demo",
  "shapes": [
    {
      "type": "shape",
      "shape_type": "rounded_rectangle",
      "left": "60%",
      "top": "20%",
      "width": "30%",
      "height": "15%",
      "fill": "#3B82F6",
      "content": "Hello"
    }
  ]
}
```

See [docs/advanced.md](docs/advanced.md) for more.

## Themes

Built-in themes: `default`, `dark`, `warm`, `forest`, `minimal`.

Custom themes can be added to `references/themes/`.

## Changelog

### 2026-07-02
- **Constraint Layout System**: 14 constraint templates with binary-search font sizing, attach-chain positioning, expand/shrink strategies, and overflow guard
- **Theme Engine Coordination**: Smart text color selection based on underlying shape fill brightness
- **Auto Shrink**: Graceful degradation for dense content
- **Image-Text Joint Constraint**: Aspect-ratio-aware image sizing with text sidebars

## License

MIT
