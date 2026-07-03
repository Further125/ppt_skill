# Template Market

Install and manage custom PPTX templates for ppt-skill.

## Install a Template

```bash
python scripts/install_template.py install my_template.pptx --name consulting --category business --description "McKinsey-style consulting template"
```

## List Installed Templates

```bash
python scripts/install_template.py list
```

## Use a Market Template

```bash
python scripts/build_pptx.py deck.json out.pptx --template templates/market/consulting/template.pptx
```

## Template Structure

Each template is a directory containing:

```
templates/market/{name}/
├── template.pptx    # The actual template file
├── config.json      # Placeholder mappings and metadata
└── analysis.json    # Full structure analysis
```

## Creating Templates

Any PPTX file with named placeholders can be used as a template.

Placeholder naming conventions:
- `Title` / `__TITLE__` -> slide title
- `Content` / `__CONTENT__` -> body text
- `Subtitle` / `__SUBTITLE__` -> subtitle
- `Date` / `__DATE__` -> date field

Run `python scripts/template_analyzer.py your.pptx --output contract.json` for detailed analysis.

## Built-in Templates

| Name | Category | Description |
|------|----------|-------------|
| default | system | Base template with 13+ layouts |
