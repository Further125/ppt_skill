# Frequently Asked Questions

## General

### What is PPT Skill?

PPT Skill is a Python tool that generates PowerPoint files from JSON specifications. It is designed for automation, LLM integration, and batch presentation generation.

### Do I need Microsoft PowerPoint installed?

No. PPT Skill uses `python-pptx` to create `.pptx` files. You only need PowerPoint or a compatible viewer to open the output.

### Do I need LibreOffice?

No. LibreOffice is optional. It is used only for high-fidelity PDF/PNG rendering. The soft preview fallback works without it.

## Usage

### Why is my text overflowing?

- Use constraint layouts (`*_constraint`) for automatic font sizing.
- Reduce content length or split across multiple slides.
- Use `qa_check.py` to detect overflow.

### Why do Chinese characters show as squares?

The default font is "Microsoft YaHei". On Linux, install a CJK font:

```bash
# Debian/Ubuntu
sudo apt-get install fonts-noto-cjk
```

Or register the bundled font:

```bash
sudo mkdir -p /usr/share/fonts/opentype/noto
sudo cp fonts/NotoSansCJKsc-Regular.otf /usr/share/fonts/opentype/noto/
sudo fc-cache -fv
```

### How do I add my own layout?

1. Define a declarative schema in `references/layout_schemas/`.
2. Or add a layout function in `scripts/build_pptx.py`.
3. Register the layout in `references/layouts.json`.

### Can I use Markdown instead of JSON?

Yes. Use `scripts/md_to_deck.py`:

```bash
python scripts/md_to_deck.py input.md output.json
```

## Troubleshooting

### `ModuleNotFoundError: No module named 'pptx'`

Install dependencies:

```bash
pip install -r requirements.txt
```

### Generated PDF is a text report instead of a real PDF

LibreOffice is not installed. Install it for PDF export, or use the long image/HTML export options which do not require LibreOffice.

### Animations do not appear in previews

Animations are embedded in the PPTX but only visible when opened in PowerPoint or WPS. The PNG preview does not render animations.
