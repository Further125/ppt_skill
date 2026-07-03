# Examples Gallery

This directory contains example JSON decks demonstrating various features of PPT Skill.

## Getting Started Examples

| File | Description |
|---|---|
| `demo_deck.json` | Minimal example with cover, content, and closing slides |
| `full_demo.json` | Comprehensive showcase of multiple layouts and features |
| `preview_test.json` | Single-slide example for `quick_preview.py` |

## Layout Examples

| File | Description |
|---|---|
| `accent_line_demo.json` | Demonstrates accent line toggle |
| `title_content` layouts | See `demo_deck.json`, `full_demo.json` |
| `two_column` / `three_column` | See `full_demo.json` |
| `image_content` | See `image_mask_demo.json` |

## Chart and Table Examples

| File | Description |
|---|---|
| `charts_anim_demo.json` | Native charts with animations |
| `table_test.json` | Table layout demonstration |
| `advanced_demo.json` | Advanced charts: waterfall, funnel, gantt |

## Styling Examples

| File | Description |
|---|---|
| `style_dark_cyber.json` | Dark cyber theme example |
| `style_dark_cyber_built.pptx` | Pre-built output of the dark cyber example |
| `style_minimal_corporate.json` | Minimal corporate theme |
| `style_infographic_rich.json` | Rich infographic style |

## Advanced Features

| File | Description |
|---|---|
| `rich_text_demo.json` | Mixed formatting runs (bold, color, size) |
| `free_shape_demo.json` | Free-form shapes and positioning |
| `image_mask_demo.json` | Image masks and cropping |
| `schema_test.json` | Declarative layout schema example |
| `schema_comprehensive.json` | Comprehensive schema layout examples |
| `compare_test.json` | Side-by-side comparison layout |

## Intent Routing

| File | Description |
|---|---|
| `auto_route_test.json` | Example for `--auto-route` automatic layout inference |

## Running Examples

```bash
python scripts/build_pptx.py examples/demo_deck.json output/demo.pptx
python scripts/build_pptx.py examples/full_demo.json output/full.pptx --theme ocean_tech
python scripts/build_pptx.py examples/auto_route_test.json output/auto.pptx --auto-route
```

## Creating Your Own

Start from `examples/demo_deck.json` and modify the `slides` array. Each slide needs a `layout` field (unless using `--auto-route`) and layout-specific content fields.
