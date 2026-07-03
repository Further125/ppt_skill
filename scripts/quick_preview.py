#!/usr/bin/env python3
"""Quick preview: render a single slide from JSON spec to PNG.

Usage:
  python quick_preview.py slide_spec.json --output preview.png
  python quick_preview.py slide_spec.json --output preview.png --auto-route

Designed for LLM visual feedback loops.
"""
import sys, os, json, argparse, tempfile


sys.path.insert(0, SCRIPT_DIR)


def main():
    parser = argparse.ArgumentParser(description="Quick single-slide preview")
    parser.add_argument("slide_json", help="JSON file with single slide spec")
    parser.add_argument("--output", "-o", required=True, help="Output PNG path")
    parser.add_argument("--template", default=os.path.join(os.path.dirname(SCRIPT_DIR), "templates", "base_template.pptx"))
    parser.add_argument("--auto-route", action="store_true", help="Auto-infer layout")
    parser.add_argument("--theme", default="default", help="Theme name")
    args = parser.parse_args()

    with open(args.slide_json, "r", encoding="utf-8") as f:
        slide_spec = json.load(f)

    # Auto-route if requested
    if args.auto_route:
        try:
            import intent_router
            deck = {"slides": [slide_spec]}
            routed = intent_router.auto_route(deck)
            slide_spec = routed["slides"][0]
            if slide_spec.get("_inferred"):
                print(f"Inferred layout: {slide_spec['layout']} (conf={slide_spec['_confidence']})")
        except Exception as e:
            print(f"Auto-route failed: {e}")

    # Build a single-slide deck
    deck = {"title": "Preview", "slides": [slide_spec]}

    # Build PPTX
    from build_pptx import build_deck
    tmp_pptx = tempfile.mktemp(suffix=".pptx")
    build_deck(deck, args.template, tmp_pptx)

    # Apply theme
    try:
        from theme_engine import apply_theme
        from pptx import Presentation
        prs = Presentation(tmp_pptx)
        apply_theme(prs, args.theme)
        prs.save(tmp_pptx)
    except Exception as e:
        print(f"Theme apply failed: {e}")

    # Render to PNG
    try:
        from render_slides import render_with_libreoffice
        tmp_dir = tempfile.mkdtemp()
        render_with_libreoffice(tmp_pptx, tmp_dir)
        # Find the rendered PNG
        pngs = sorted([f for f in os.listdir(tmp_dir) if f.endswith(".png")])
        if pngs:
            import shutil
            src = os.path.join(tmp_dir, pngs[0])
            shutil.copy2(src, args.output)
            print(f"Preview saved: {args.output}")
        else:
            print("Rendering failed: no PNG output")
    except Exception as e:
        print(f"Render failed: {e}")
    finally:
        if os.path.exists(tmp_pptx):
            os.remove(tmp_pptx)


if __name__ == "__main__":
    main()
