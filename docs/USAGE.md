# Usage Guide

## Command Line Interface

`scripts/build_pptx.py` is the main entry point.

```bash
python scripts/build_pptx.py <input.json> <output.pptx> [options]
```

Common options:

| Option | Description |
|---|---|
| `--theme <name>` | Apply a built-in or custom theme |
| `--auto-route` | Infer layout from slide content |
| `--animate <effect>` | Add entrance animations |
| `--template <path>` | Use a custom PPTX template |
| `--verbose` | Print detailed debug info |

## Layout Auto-Routing

Omit `layout` in your JSON and use `--auto-route`:

```json
{
  "title": "Auto Route Demo",
  "slides": [
    {"title": "Cover", "subtitle": "Auto inferred"},
    {"quote": "Less is more", "author": "Someone"},
    {"big_number": "99.9%", "label": "Uptime"}
  ]
}
```

```bash
python scripts/build_pptx.py deck.json output.pptx --auto-route
```

## Themes

Built-in themes: `default`, `ocean`, `sunset`, `forest`, `berry`, `monochrome`, `minimal`, `dark`.

```bash
python scripts/build_pptx.py deck.json output.pptx --theme ocean
```

Custom theme file:

```bash
python scripts/build_pptx.py deck.json output.pptx --theme references/themes/my_theme.json
```

## QA Check

```bash
python scripts/qa_check.py output.pptx
```

Checks for empty slides, text overflow, and missing placeholders.

## Other Tools

| Script | Purpose |
|---|---|
| `render_slides.py` | PPTX to PNG preview |
| `export.py` | PDF / long image / HTML export |
| `snapshot.py` | Capture a specific page or region |
| `template_analyzer.py` | Extract visual contract from a PPTX template |
| `contract_to_schema.py` | Convert template contract to declarative schema |
| `quick_preview.py` | Single-slide JSON to PNG for LLM feedback |
