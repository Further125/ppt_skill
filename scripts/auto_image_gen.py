#!/usr/bin/env python3
"""
auto_image_gen.py — Auto-generate diagram images for image_content slides.

Scans a deck JSON for image_content slides lacking an image_path,
generates a themed diagram via matplotlib, writes the image next to
the JSON, updates the JSON in-place, and optionally rebuilds the PPTX.

Usage:
    python auto_image_gen.py examples/style_dark_cyber.json
    python auto_image_gen.py deck.json --build --theme references/themes/cyber_neon.json
"""
import sys, os, json, argparse, textwrap, math

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Circle, Wedge
import numpy as np


def _hex_to_rgb(hex_str):
    h = hex_str.lstrip("#")
    return tuple(int(h[i:i+2], 16) / 255.0 for i in (0, 2, 4))


def _lightness(hex_color):
    r, g, b = _hex_to_rgb(hex_color)
    return 0.299 * r + 0.587 * g + 0.114 * b


def _choose_theme_colors(theme_dict):
    """Pick diagram colors that harmonise with the deck theme."""
    bg = theme_dict.get("bg", "#0F172A") if isinstance(theme_dict, dict) else "#0F172A"
    primary = theme_dict.get("primary", "#0EA5E9") if isinstance(theme_dict, dict) else "#0EA5E9"
    accent = theme_dict.get("accent", "#A855F7") if isinstance(theme_dict, dict) else "#A855F7"
    text = theme_dict.get("text", "#E2E8F0") if isinstance(theme_dict, dict) else "#E2E8F0"

    is_dark = _lightness(bg) < 0.5
    if is_dark:
        return {
            "bg": bg,
            "node_bg": theme_dict.get("light_bg", "#1E293B") if isinstance(theme_dict, dict) else "#1E293B",
            "node_border": "#334155",
            "center_border": primary,
            "line": "#334155",
            "glow": primary,
            "text": text,
            "accent_text": primary,
        }
    else:
        return {
            "bg": bg,
            "node_bg": "#F8FAFC",
            "node_border": "#CBD5E1",
            "center_border": primary,
            "line": "#CBD5E1",
            "glow": primary,
            "text": text,
            "accent_text": primary,
        }


def _extract_nodes(title, content_lines):
    """Try to extract a central concept + satellite nodes from slide text."""
    center = title.strip() if title else "System"
    nodes = []
    for line in content_lines:
        line = line.strip()
        if ":" in line:
            label = line.split(":", 1)[0].strip()
            nodes.append(label)
        elif "→" in line:
            label = line.split("→", 1)[0].strip()
            nodes.append(label)
        elif line:
            # Take first 2-3 words as label
            words = line.split()
            nodes.append(" ".join(words[:2]) if len(words) > 1 else words[0])
    if not nodes:
        nodes = ["Component A", "Component B", "Component C"]
    return center, nodes[:6]  # cap at 6 satellites


def generate_hub_diagram(title, content, out_path, theme_dict=None):
    """Draw a hub-and-spoke diagram based on slide text."""
    colors = _choose_theme_colors(theme_dict or {})
    center_label, nodes = _extract_nodes(title, content)

    fig, ax = plt.subplots(1, 1, figsize=(10.24, 5.76), facecolor=colors["bg"])
    ax.set_facecolor(colors["bg"])
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")

    cx, cy = 50, 50
    n = len(nodes)
    radius = 32 if n <= 4 else 28

    # Draw connections first (behind nodes)
    for i in range(n):
        angle = 2 * math.pi * i / n - math.pi / 2
        nx = cx + radius * math.cos(angle)
        ny = cy + radius * math.sin(angle)
        ax.plot([nx, cx], [ny, cy], color=colors["line"], linewidth=2, zorder=1)
        # Glowing dot on line
        mid_x, mid_y = (nx + cx) / 2, (ny + cy) / 2
        glow_dot = Circle((mid_x, mid_y), 1.2, facecolor=colors["glow"], alpha=0.5, zorder=2)
        ax.add_patch(glow_dot)

    # Center hub
    cw, ch = 24, 14
    center = FancyBboxPatch(
        (cx - cw/2, cy - ch/2), cw, ch,
        boxstyle="round,pad=1.2", facecolor=colors["node_bg"],
        edgecolor=colors["center_border"], linewidth=2.5, zorder=4
    )
    ax.add_patch(center)
    ax.text(cx, cy, center_label, fontsize=13, color=colors["accent_text"],
            ha="center", va="center", fontweight="bold", zorder=5)

    # Satellite nodes
    for i, label in enumerate(nodes):
        angle = 2 * math.pi * i / n - math.pi / 2
        nx = cx + radius * math.cos(angle)
        ny = cy + radius * math.sin(angle)

        # Glow
        glow = Circle((nx, ny), 10.5, facecolor="none", edgecolor=colors["glow"],
                      linewidth=8, alpha=0.12, zorder=1)
        ax.add_patch(glow)

        # Node
        node = Circle((nx, ny), 10, facecolor=colors["node_bg"],
                      edgecolor=colors["node_border"], linewidth=2, zorder=3)
        ax.add_patch(node)

        wrapped = "\n".join(textwrap.wrap(label, width=14))
        ax.text(nx, ny, wrapped, fontsize=10, color=colors["text"],
                ha="center", va="center", fontweight="bold", zorder=5)

    # Decorative particles
    rng = np.random.default_rng(42)
    for _ in range(25):
        px, py = rng.uniform(5, 95), rng.uniform(5, 95)
        if not (40 < px < 60 and 40 < py < 60):
            ax.plot(px, py, "o", color=colors["glow"],
                    markersize=rng.uniform(1, 3), alpha=rng.uniform(0.08, 0.35), zorder=0)

    plt.tight_layout(pad=0)
    plt.savefig(out_path, dpi=150, facecolor=colors["bg"],
                edgecolor="none", bbox_inches="tight", pad_inches=0.1)
    plt.close()
    print(f"  Diagram saved: {out_path}")


def process_deck(json_path, build=False, theme_path=None, template=None):
    with open(json_path, "r", encoding="utf-8") as f:
        deck = json.load(f)

    deck_dir = os.path.dirname(os.path.abspath(json_path))
    theme = None
    if theme_path and os.path.exists(theme_path):
        with open(theme_path, "r", encoding="utf-8") as f:
            theme = json.load(f)
    elif "theme" in deck and isinstance(deck["theme"], dict):
        theme = deck["theme"]

    modified = False
    for idx, slide in enumerate(deck.get("slides", [])):
        if slide.get("layout") != "image_content":
            continue
        if slide.get("image_path") and os.path.exists(slide["image_path"]):
            continue  # already has a real image

        title = slide.get("title", "")
        content = slide.get("content", [])
        if isinstance(content, str):
            content = [content]
        if not content:
            continue

        out_name = f"auto_img_slide_{idx+1:03d}.png"
        out_path = os.path.join(deck_dir, out_name)
        print(f"Generating diagram for slide {idx+1}: {title}")
        generate_hub_diagram(title, content, out_path, theme_dict=theme)
        slide["image_path"] = out_path
        modified = True

    if modified:
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(deck, f, ensure_ascii=False, indent=2)
        print(f"Updated JSON: {json_path}")
    else:
        print("No missing images found — nothing to generate.")

    if build and modified:
        build_script = os.path.join(SCRIPT_DIR, "build_pptx.py")
        if os.path.exists(build_script):
            out_pptx = os.path.splitext(json_path)[0] + "_built.pptx"
            cmd = [sys.executable, build_script, json_path, out_pptx]
            if theme_path:
                cmd += ["--theme", theme_path]
            import subprocess
            subprocess.run(cmd)
        else:
            print("Warning: build_pptx.py not found, skipping PPTX build.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Auto-generate images for image_content slides")
    parser.add_argument("deck_json", help="Path to deck JSON")
    parser.add_argument("--build", action="store_true", help="Rebuild PPTX after generating images")
    parser.add_argument("--theme", help="Theme JSON path for color harmony")
    parser.add_argument("--template", help="Template PPTX path (passed to build_pptx.py)")
    args = parser.parse_args()
    process_deck(args.deck_json, build=args.build, theme_path=args.theme, template=args.template)
