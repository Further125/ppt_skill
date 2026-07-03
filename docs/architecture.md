# PPT Skill -- Architecture Design Document

> Internal architecture, module design, and extension guide for developers.

---

## 1. Project Structure

```
ppt_skill/
├── scripts/
│   ├── build_pptx.py         # Core build engine
│   ├── theme_engine.py       # Theme engine
│   ├── color_guard.py        # WCAG color audit
│   ├── layout_schema.py      # Declarative layout schema
│   ├── intent_router.py      # Layout inference
│   ├── template_analyzer.py  # Template visual contract extraction
│   ├── contract_to_schema.py # Contract to schema
│   ├── text_fitter.py        # Font auto-fit
│   ├── animator.py           # Animation engine
│   ├── advanced_charts.py    # Advanced charts
│   ├── render_slides.py      # PPTX to PNG
│   ├── export.py             # Export tools
│   ├── snapshot.py           # Snapshot capture
│   └── qa_check.py           # QA check
├── templates/
│   └── base_template.pptx    # Master template
├── references/
│   ├── themes/               # Theme JSON
│   └── layout_schemas/       # Layout schemas
├── examples/                 # Example decks
└── docs/                     # Documentation
```

---

## 2. Core Build Flow

`build_pptx.py` uses a three-stage build:

### Stage 1: Parsing and Routing
1. Load the deck JSON.
2. When `--auto-route` is used, call `intent_router` to infer the layout.
3. Load the template PPTX and extract the available layouts.

### Stage 2: Slide-by-Slide Construction
For each `slide_spec`:
1. Prefer schema rendering (`layout_schema.render_slide_from_schema`).
2. If the schema does not exist, fall back to template cloning (`clone_slide`).
3. Call the corresponding rendering function based on the layout type.
4. Synchronize the title accent line (`sync_accent_line`).
5. Process `hyperlink` / `video_path` / `shapes`.
6. Detect overflow and adjust (`_adjust_slide_for_overflow`).

### Stage 3: Post-Processing
1. Apply the theme (`theme_engine.apply_theme`).
2. Apply animations (`animator`).
3. Run the color audit (`color_guard`).
4. Export (`export`).

---

## 3. Layout System

### 3.1 Traditional Layout (Template Cloning)
- Find the corresponding template slide in `base_template.pptx`.
- Clone the XML element into a new slide using `copy.deepcopy`.
- Placeholder text matching (e.g., `__TITLE__`, `__CONTENT__`).
- Pros: high compatibility, inherits master styles.
- Cons: fixed layouts, difficult to control precisely.

### 3.2 Declarative Schema Layout (New)
JSON Schema defines the layout and supports:
- Percentage coordinates (`5%`, `left-half`, `center`).
- `attach` relative positioning (`attach to title bottom + 1%`).
- `content_source` field mapping.
- `adaptive` overflow strategies (`shrink` / `expand` / `truncate`).

---

## 4. Theme Engine

`theme_engine.py` design principles:

### 4.1 Color Mapping
- Define legacy fixed color constants (`OLD_ACCENT_BLUE`, `OLD_DARK_TEXT`, etc.).
- Iterate over all shapes and map old colors to the new theme colors.
- Background shapes (>80% area) are set to `theme.background`.
- Chart backgrounds are forced to light colors and text to dark colors (to ensure readability).
- Tables keep light backgrounds and dark text.

### 4.2 Smart Text Color Selection
Calculate WCAG contrast based on the shape fill color and automatically choose white or black text:
- Compute the contrast of white and dark text against the fill color.
- Select the color with the higher contrast.
- Large dark backgrounds use dimmed white (`#E2E8F0`) to avoid glare.

---

## 5. Color Contrast Audit

`color_guard.py` implementation:

### Detection Dimensions
1. `shape_fill_text`: text contrast on shape fills.
2. `textbox`: text contrast for text boxes against the background.
3. `bg_bg_contrast`: chart/table background contrast against the slide background (INFO).

### Contrast Calculation
WCAG 2.1 standard:
- Relative luminance: `L = 0.2126*R + 0.7152*G + 0.0722*B`
- Contrast: `(L1 + 0.05) / (L2 + 0.05)`
- `CRITICAL`: < 2.0 | `WARN`: < 4.5 | `INFO`: >= 4.5

### Background Sampling
- If the center point of the text shape falls inside another shape, use that shape's fill as the background.
- Otherwise, use the slide-level background color.

---

## 6. Text Auto-Fitting

`text_fitter.py` is implemented with Pillow:
1. Load fonts (Noto Sans CJK / DejaVu).
2. Binary search for the maximum usable font size:
   - `low = 1`, `high = initial font size`
   - `mid = (low + high) // 2`
   - Pillow measures the text width and height at size `mid`.
   - If it exceeds the bounds -> `high = mid - 1`
   - Otherwise -> `low = mid + 1`
3. Return the largest non-overflowing font size.
4. Actual rendering uses Microsoft YaHei, with a width fine-tuning factor of `1.04`.

---

## 7. Declarative Layout Schema

`layout_schema.py` core concepts:

### Region System
- Percentages: `5%`, `left-half`, `top-third`, `center`
- Absolute units: `2in`, `5cm`, `100px`
- Raw EMU: `5000000`

### Attach System
Relative positioning to avoid hard-coding:
```json
{ "attach": { "to": "title", "edge": "bottom", "offset": "1%" } }
```

### Adaptive Strategies
- `shrink`: reduce font size until it fits.
- `expand`: extend the shape height.
- `truncate`: truncate the text.

---

## 8. Intent Routing

`intent_router.py` infers layouts from content fields:

Matching rule examples:
- `{quote, author}` -> `quote`
- `{big_number, label}` -> `data_highlight`
- `{left_content, right_content}` -> `two_column`
- `{title, subtitle}` -> `cover`
- `{chart_data}` -> `chart`
- `{table_data}` -> `table`

Matched in priority order; if `slide_spec` contains all the keys of a rule, the corresponding layout is assigned.

---

## 9. Template Analysis

`template_analyzer.py` extracts the visual contract:
1. Iterate over all slides in the template.
2. Extract shape type, position, size, fill color, font, and font size.
3. Recognize placeholder text patterns (such as `__TITLE__`).
4. Output `contract.json` (colors, fonts, layout structure, placeholder mapping).

`contract_to_schema.py` converts it into a declarative Schema:
- Compute relative coordinates and percentages.
- Generate `attach` relationships.
- Output JSON Schema.

---

## 10. Rendering and Export

### 10.1 Rendering

`render_slides.py` provides two modes:

LibreOffice mode (default):
- `soffice --headless --convert-to pdf`
- `pdftoppm` or PyMuPDF `PDF -> PNG`
- Pros: pixel-perfect, consistent font rendering.
- Cons: requires LibreOffice, slow first-time load.

Soft rendering (`--soft`):
- `python-pptx` reads shape text and colors.
- Pillow draws to an image.
- Pros: pure Python, fast.
- Cons: approximate rendering; complex charts may be distorted.

### 10.2 Export
- PDF: converted via `soffice`.
- Long image: vertically concatenate all slide PNGs.
- HTML: reveal.js-style slides.

---

## 12. Constraint Solving System

### 12.1 Design Goal

"Maximize font size while auto-arranging bounding boxes" -- the core goal of the constraint solving system.

### 12.2 Architecture

```
constraint_layout.py
├── _pptx_wrap_text()          # Pillow text wrapping simulation
├── _measure_required_height() # Compute required text height
├── _truncate_text_to_height() # Overflow truncation
├── _expand_font_size()        # Fill available space
├── solve_constraints()        # Main solver
│   ├── 1. Parse schema shapes
│   ├── 2. Topological sort by attach edges
│   ├── 3. Collect textboxes + images
│   ├── 4. Binary search best_base
│   ├── 5. Compute actual height + expand/shrink
│   ├── 6. Apply region + attach chain
│   └── 7. Return (region, style) list
└── _compute_image_height()    # Maintain image aspect ratio
```

### 12.3 Binary Search Algorithm

```python
lo, hi = 8, 96
best_base = 8
while lo <= hi:
    mid = (lo + hi) // 2
    if _fits(mid):      # All textboxes fit?
        best_base = mid
        lo = mid + 1    # Try larger
    else:
        hi = mid - 1    # Too large, shrink
```

`_fits(base_size)` checks every textbox:
- Actual font size = `base_size * font_scale`
- Compute the wrapped text height
- If height <= available height -> fits
- In single-line mode, width is also checked

### 12.4 Attach Chain Calculation

```python
# Topological sort: attach targets come first
sorted_shapes = topological_sort(shapes, attach_edges)

# Compute region one by one
for shape in sorted_shapes:
    if shape.attach:
        parent = shape.attach.to
        parent_region = computed[parent]
        shape.region.top = parent_region.top + parent_region.height + offset
```

Supports 8 directions: `top`, `bottom`, `left`, `right`, `top-left`, `top-right`, `bottom-left`, `bottom-right`.

### 12.5 Expand Strategy

When text height is less than 50% of the region height, try to increase the font size until it fills 60%:

```python
if h < max_height * 0.50:
    expanded = _expand_font_size(text, width, max_height, current_size)
    actual_size = expanded
    h = measure(text, expanded, width)
```

### 12.6 Shrink Strategy

When `best_base < 14pt`, automatically degrade:

1. **Line-spacing compression**: `1.40 -> 1.20` (+14% vertical space)
2. **Font-scale compression**: all `font_scale * 0.85`
3. **Final warning**: if still < 10pt, prompt the user to reduce content

```python
if best_base < 14:
    best_base_tight = _run_solver(line_spacing=1.20)
    if best_base_tight > best_base:
        best_base = best_base_tight
    if best_base < 10:
        # Compress font_scale and retry
        shrunk_textboxes = [tb.font_scale *= 0.85 for tb in textboxes]
        best_base_shrunk = _run_solver(shrunk_textboxes, line_spacing=1.20)
```

### 12.7 Theme Coordination

When applying a theme, `theme_engine.py` detects the fill color of the `auto_shape` beneath a textbox:

```python
def _detect_underlying_fill(textbox, slide):
    # Check all auto_shapes overlapping the textbox
    # Return the fill color of the largest one
    # Choose white or black text based on fill brightness
```

This resolves the issue where white text on colored shapes (such as process cards) is globally changed to a dark color by the theme.

---

## 13. Extension Guide

### 11.1 Adding a New Layout

1. Create a template slide in `base_template.pptx`.
2. Add an `elif layout_type == xxx` branch in `build_pptx.py`.
3. Implement the rendering function (refer to existing layouts).
4. Add documentation in `layouts.md`.
5. (Optional) Create a declarative Schema at `layout_schemas/xxx.json`.

### 11.2 Adding a New Theme

1. Create a JSON file under `references/themes/`.
2. Include fields: `name`, `primary`, `secondary`, `accent[]`, `background`, `text`, `light_bg`, `font_title`, `font_body`.
3. Test with `--theme /path/to/theme.json`.

### 11.3 Adding a New Chart Type

1. Implement the drawing function in `advanced_charts.py`.
2. Add a type check in the `chart` layout branch of `build_pptx.py`.
3. Embed using `matplotlib` + `python-pptx` image insertion.

---

*Document version: 2026-07-01*
