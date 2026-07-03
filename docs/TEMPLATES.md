# Templates

PPT Skill uses a base PPTX template with predefined slide layouts and placeholder shapes.

## Base Template

`templates/base_template.pptx` is the default template. It contains placeholder shapes for all supported layouts.

## How Templates Work

- The builder loads the template.
- Each slide uses a layout from the template.
- Placeholder shapes named `__TITLE__`, `__CONTENT__`, `__IMAGE__`, etc. are replaced with generated content.
- The constraint layout system can bypass template placeholders and compute positions dynamically.

## Using a Custom Template

```bash
python scripts/build_pptx.py deck.json output.pptx --template templates/my_template.pptx
```

## Creating a Template

1. Create a PPTX file in PowerPoint or LibreOffice.
2. Add slides for each layout you need.
3. Add placeholder shapes with recognized names:
   - `__TITLE__` — slide title
   - `__SUBTITLE__` — subtitle
   - `__CONTENT__` — body content
   - `__IMAGE__` — image placeholder
   - `__CHART__` — chart placeholder
   - `__TABLE__` — table placeholder
4. Save the template and reference it with `--template`.

## Template Analysis

Extract the visual contract (colors, fonts, layout structure) from any template:

```bash
python scripts/template_analyzer.py templates/base_template.pptx --output contract.json
```

Convert the contract to a declarative schema:

```bash
python scripts/contract_to_schema.py contract.json --layout cover --output schema.json
```

## Template Market

`templates/market/` contains community or alternative template packages. Each package includes:

- `template.pptx` — the template file
- `config.json` — metadata and supported layouts
- `analysis.json` — extracted visual contract
