# PPT Skill — AI-Powered PowerPoint Generator

A Python + python-pptx based system for automatically generating PowerPoint presentations from JSON specifications.

Supports 20+ slide layouts, native charts and tables, timelines, process diagrams, team cards, animations, theme switching, and multimedia extensions.

---

## Quick Start

**Requirements**: Python 3.10+, dependencies `python-pptx`, `Pillow`, `PyMuPDF`. Optional: LibreOffice for high-fidelity rendering, Noto Sans CJK for Chinese display.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

```bash
# Build a PPT
python scripts/build_pptx.py deck.json output.pptx

# Render preview images
python scripts/render_slides.py output.pptx preview_dir

# QA check
python scripts/qa_check.py output.pptx
```

## Project Structure

```
ppt_skill/
├── README.md                    # English documentation
├── README.zh-CN.md              # Chinese documentation
├── scripts/                     # Core scripts
│   ├── build_pptx.py            # Core builder
│   ├── render_slides.py         # PPTX to PNG preview
│   ├── qa_check.py              # Quality checks
│   └── export.py                # Export tools
├── templates/                   # PPTX templates
├── references/                  # Layout schemas and themes
├── examples/                    # Example JSON decks
├── tests/                       # Test fixtures
└── docs/                        # Documentation
```

## Supported Layouts

- `cover`, `toc`, `section`
- `title_content`, `two_column`, `three_column`
- `chart`, `table`
- `timeline`, `process`, `tree`
- `team`, `quote`, `data_highlight`
- `image_content`, `image_text_split`
- `wordcloud`, `closing`

Constraint variants (`*_constraint`) are available for automatic font sizing and positioning.

## Chart Types

`column`, `line`, `bar`, `pie`, `doughnut`, `area`, `radar`, `bubble`, `scatter`, `combo`.

## Themes

```bash
python scripts/build_pptx.py deck.json output.pptx --theme dark
```

## JSON Format

```json
{
  "title": "Presentation Title",
  "slides": [
    {"layout": "cover", "title": "Hello", "subtitle": "World"},
    {"layout": "title_content", "title": "Content", "content": ["A", "B"]}
  ]
}
```

## Notes

- Default font is "Microsoft YaHei"; install `fonts-noto-cjk` on Linux for CJK support.
- Soft preview is approximate; install LibreOffice for pixel-accurate rendering.
- Animations are visible only when opened in PowerPoint / WPS.

## License

MIT
