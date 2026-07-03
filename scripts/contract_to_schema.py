#!/usr/bin/env python3
"""Convert a template visual contract into a layout schema.

Usage:
  python contract_to_schema.py contract.json --layout cover --output schema.json
"""
import sys, os, json, argparse


def contract_to_schema(contract, layout_name):
    """Convert a template contract into a layout schema."""
    layouts = contract.get("layouts", [])
    target = None
    for l in layouts:
        if l["name"] == layout_name:
            target = l
            break
    if not target:
        print(f"Layout '{layout_name}' not found in contract")
        return None

    theme = contract.get("theme", {})
    primary = theme.get("primary", "#3B82F6")
    bg = theme.get("background", "#FFFFFF")
    text_color = theme.get("text", "#1F2937")
    inverse_text = "#FFFFFF" if bg and bg != "#FFFFFF" else "#1F2937"

    shapes = []
    for s in target["shapes"]:
        shape_def = {
            "role": s.get("role", "shape"),
            "type": "textbox" if s["type"] == "TEXT_BOX" else "shape",
            "region": {
                "left": f"{s['rel_left']*100:.1f}%",
                "top": f"{s['rel_top']*100:.1f}%",
                "width": f"{s['rel_width']*100:.1f}%",
                "height": f"{s['rel_height']*100:.1f}%",
            },
        }

        # Infer style
        style = {}
        if s.get("font_size"):
            style["font_size"] = s["font_size"]
        if s.get("bold"):
            style["bold"] = True
        if s.get("font_name"):
            style["font_name"] = s["font_name"]

        # Detect if dark background
        is_dark = bg and bg.upper() not in ("#FFFFFF", "#FFF", "#F3F4F6")
        if is_dark and s.get("rel_top", 0) > 0.1:  # not background
            style["color"] = inverse_text
        elif s.get("color"):
            style["color"] = s["color"]
        else:
            style["color"] = text_color

        # Background shapes
        if s["type"] == "AUTO_SHAPE" and s.get("rel_width", 0) > 0.8 and s.get("rel_height", 0) > 0.8:
            shape_def["shape_type"] = "rectangle"
            shape_def["type"] = "shape"
            if s.get("fill"):
                shape_def["style"] = {"fill": s["fill"]}
            else:
                shape_def["style"] = {"fill": bg}
            shapes.append(shape_def)
            continue

        # Accent line detection: thin horizontal/vertical line
        aspect = s.get("rel_width", 1) / max(s.get("rel_height", 0.001), 0.001)
        if aspect > 20 or aspect < 0.05:
            shape_def["shape_type"] = "rectangle"
            shape_def["type"] = "shape"
            shape_def["style"] = {"fill": primary}
            shape_def["role"] = "accent_line"
            shapes.append(shape_def)
            continue

        # Fix null color
        if "color" in style and style["color"] is None:
            del style["color"]

        shape_def["style"] = style

        # Fix type: if it has text, it's a textbox
        if s["type"] == "TEXT_BOX":
            shape_def["type"] = "textbox"

        # Detect content role from text preview
        text_preview = s.get("text_preview", "")
        if text_preview == "__TITLE__":
            shape_def["content_source"] = "title"
            shape_def["adaptive"] = {"strategy": "shrink", "min_size": 20}
        elif text_preview == "__SUBTITLE__":
            shape_def["content_source"] = "subtitle"
        elif text_preview == "__DATE__":
            shape_def["content_source"] = "date"
        elif text_preview in ("__CONTENT__", "__ITEMS__"):
            shape_def["content_source"] = "content"
            shape_def["adaptive"] = {"strategy": "shrink", "min_size": 12}
        elif text_preview == "__LEFT__":
            shape_def["content_source"] = "left_content"
        elif text_preview == "__RIGHT__":
            shape_def["content_source"] = "right_content"
        elif "__QUOTE__" in text_preview:
            shape_def["content_source"] = "quote"
        elif "__AUTHOR__" in text_preview:
            shape_def["content_source"] = "author"
        elif "__BIG_NUMBER__" in text_preview:
            shape_def["content_source"] = "big_number"
        elif "__LABEL__" in text_preview:
            shape_def["content_source"] = "label"

        shapes.append(shape_def)

    return {
        "name": layout_name,
        "description": f"Auto-generated schema from template analysis",
        "shapes": shapes,
    }


def main():
    parser = argparse.ArgumentParser(description="Convert contract to schema")
    parser.add_argument("contract", help="Contract JSON from template_analyzer")
    parser.add_argument("--layout", required=True, help="Layout name to convert")
    parser.add_argument("--output", "-o", required=True, help="Output schema JSON")
    args = parser.parse_args()

    with open(args.contract, "r", encoding="utf-8") as f:
        contract = json.load(f)

    schema = contract_to_schema(contract, args.layout)
    if schema:
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(schema, f, ensure_ascii=False, indent=2)
        print(f"Schema saved: {args.output}")


if __name__ == "__main__":
    main()
