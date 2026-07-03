# Quick Start

## 1. Install

```bash
pip install -r requirements.txt
```

Recommended Python: 3.10+

## 2. Write a JSON Deck

```json
{
  "title": "My First Deck",
  "slides": [
    {
      "layout": "cover",
      "title": "Hello PPT Skill",
      "subtitle": "JSON-driven PowerPoint"
    },
    {
      "layout": "title_content",
      "title": "Key Points",
      "content": [
        "Write JSON once",
        "Generate PPTX automatically",
        "Apply themes and layouts"
      ]
    }
  ]
}
```

Save it as `my_deck.json`.

## 3. Build

```bash
python scripts/build_pptx.py my_deck.json output/my_deck.pptx
```

## 4. Preview

```bash
# Soft preview (no LibreOffice required)
python scripts/render_slides.py output/my_deck.pptx output/preview --soft
```

## 5. Export

```bash
python scripts/export.py output/my_deck.pptx --format pdf --output my_deck.pdf
```
