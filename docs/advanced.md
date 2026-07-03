# Advanced Features

## Rich Text

Any text field that accepts a string can also accept an array of **run dicts** for mixed formatting.

```json
{
  "content": [
    {"text": "Rust", "bold": true, "color": "#3B82F6", "size": 24},
    {"text": " is blazingly fast."}
  ]
}
```

**Run dict fields:**

| Field | Type | Description |
|-------|------|-------------|
| `text` | string | Text content (required) |
| `bold` | bool | Bold text |
| `italic` | bool | Italic text |
| `underline` | bool | Underlined text |
| `color` | string | Hex color (`#RRGGBB`) |
| `size` | number | Font size in points |
| `font` | string | Font family name |

> **Tip:** When `\n` appears in a run's text, it automatically creates a new paragraph. This is useful for vertical centering in shapes.

---

## Free-Form Shapes

Any slide can include a `shapes` array for precise positioning of custom elements.

### Shape Types

#### Text Box

```json
{
  "type": "text",
  "left": 1000000,
  "top": 2000000,
  "width": 3000000,
  "height": 1000000,
  "content": "Hello World",
  "font_size": 18,
  "color": "#374151",
  "alignment": "center",
  "vertical_center": true
}
```

#### Image

```json
{
  "type": "image",
  "path": "photo.jpg",
  "left": 5000000,
  "top": 2000000,
  "width": 2000000,
  "height": 1500000,
  "mode": "fit"
}
```

`mode`: `"fit"` (preserve aspect ratio, letterbox) or `"stretch"` (fill exactly).

#### Shape (rectangle, oval, etc.)

```json
{
  "type": "shape",
  "shape_type": "rounded_rectangle",
  "left": 1000000,
  "top": 4000000,
  "width": 2000000,
  "height": 1000000,
  "fill": "#3B82F6",
  "border_color": "#1E3A5F",
  "border_width": 2,
  "content": "Click me",
  "alignment": "center",
  "vertical_center": true
}
```

**Available shape_types:**

- `rectangle`
- `rounded_rectangle`
- `oval`
- `triangle`
- `diamond`
- `pentagon`
- `hexagon`
- `star`
- `arrow_right`, `arrow_left`, `arrow_up`, `arrow_down`
- `chevron`
- `parallelogram`
- `trapezoid`
- `donut`

#### Chart

```json
{
  "type": "chart",
  "left": 1000000,
  "top": 4000000,
  "width": 8000000,
  "height": 4000000,
  "chart_data": {
    "type": "bar",
    "categories": ["A", "B", "C"],
    "values": [10, 20, 30]
  }
}
```

### Common Shape Fields

All shapes support:

| Field | Type | Description |
|-------|------|-------------|
| `left` | int | X position in EMU |
| `top` | int | Y position in EMU |
| `width` | int | Width in EMU |
| `height` | int | Height in EMU |
| `z_order` | int | Stacking order (higher = more front) |

---

## Image Masking

Mask an image into a shape (circle, rounded rectangle, etc.):

```json
{
  "type": "image",
  "path": "avatar.jpg",
  "mask": "circle",
  "left": 1000000,
  "top": 1000000,
  "width": 1500000,
  "height": 1500000
}
```

**Available masks:**

- `circle` / `oval`
- `rounded_rectangle`
- `rectangle`
- `triangle`
- `diamond`
- `hexagon`
- `star`
- `pentagon`
- `heart`
- `cloud`
- `sun`
- `moon`

---

## Image Cropping

Crop an image from any edge using ratios (0.0–1.0):

```json
{
  "type": "image",
  "path": "photo.jpg",
  "crop": {
    "left": 0.1,
    "right": 0.1,
    "top": 0,
    "bottom": 0
  }
}
```

Each value is the fraction of the image dimension to crop from that edge.

---

## Z-Order (Layer Control)

Control which shapes appear on top of others:

```json
{
  "type": "shape",
  "shape_type": "rectangle",
  "left": 0,
  "top": 0,
  "width": 10000000,
  "height": 5000000,
  "fill": "#F3F4F6",
  "z_order": 0
},
{
  "type": "image",
  "path": "overlay.png",
  "z_order": 5
}
```

- `z_order: 0` = back (bottom layer)
- Higher numbers = front (top layer)

---

## Accent Line Toggle

The blue underline decoration below titles can be disabled per slide:

```json
{
  "layout": "title_content",
  "title": "Clean Title",
  "accent_line": false,
  "content": "No underline on this slide."
}
```

Default is `true` (show accent line).

---

## Auto-Fit System

The builder uses Pillow to measure text and binary search to find the largest font size that fits a shape. This is applied automatically to:

- Titles and subtitles
- Content text boxes
- Table cells
- Timeline/process/team text boxes
- Big numbers in data_highlight

You can disable auto-fit by not passing `auto_fit=True` to `replace_placeholder_text` (for custom scripts).

---

## Themes

Switch color palettes via CLI:

```bash
python scripts/build_pptx.py deck.json output.pptx --theme ocean
```

Built-in themes:

| Theme | Description |
|-------|-------------|
| `default` | Navy blue + white |
| `ocean` | Teal + light blue |
| `sunset` | Orange + coral |
| `forest` | Green + sage |
| `berry` | Purple + magenta |
| `monochrome` | Gray scale |

Custom themes can be added to `references/themes/` as JSON files with `primary`, `secondary`, `accent`, `dark`, `light`, `background`, `text` color fields.

---

## Export Formats

### PDF

```bash
python scripts/export.py deck.pptx --format pdf --output deck.pdf
```

Requires LibreOffice (`soffice`). Falls back to text report if unavailable.

### Long Image

Stitches all slides vertically into one PNG:

```bash
python scripts/export.py deck.pptx --format long_image --output deck.png
```

### HTML Slideshow

Simple reveal.js-style HTML:

```bash
python scripts/export.py deck.pptx --format html --output deck.html
```

---

## Snapshot Tool

Capture a specific page or region:

```bash
# Full page
python scripts/snapshot.py deck.pptx --page 3 --output page3.png

# Pixel crop (x, y, w, h)
python scripts/snapshot.py deck.pptx --page 3 --crop "200,300,800,600" --output clip.png

# EMU crop (same unit as shapes)
python scripts/snapshot.py deck.pptx --page 3 --crop-emu "500000,1000000,2000000,1500000" --output clip.png

# List page dimensions
python scripts/snapshot.py deck.pptx --list
```

---

## EMU Coordinate System

All positions in `shapes` use **EMU** (English Metric Units):

- `1 inch = 914,400 EMU`
- `1 cm = 360,000 EMU`
- Standard 16:9 slide: `12,192,000 × 6,858,000 EMU`

Quick reference for a 16:9 slide:

| Position | EMU |
|----------|-----|
| Left edge | 0 |
| Right edge | 12,192,000 |
| Top edge | 0 |
| Bottom edge | 6,858,000 |
| Center X | 6,096,000 |
| Center Y | 3,429,000 |
