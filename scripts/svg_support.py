#!/usr/bin/env python3
"""
SVG support for ppt-skill.

Converts SVG files to images that can be embedded in PPTX.

Usage:
    from svg_support import svg_to_image
    svg_to_image("icon.svg", "icon.png", width=200, height=200)
"""

import os
import subprocess
import sys
import tempfile


def _has_cairosvg():
    try:
        import cairosvg
        return True
    except ImportError:
        return False


def _has_svglib():
    try:
        from svglib.svglib import svg2rlg
        return True
    except ImportError:
        return False


def _has_inkscape():
    try:
        subprocess.run(["inkscape", "--version"], capture_output=True, check=True)
        return True
    except Exception:
        return False


def svg_to_image(svg_path, output_path=None, width=None, height=None, dpi=150):
    """Convert an SVG file to PNG.

    Returns the output path. Tries multiple backends in order:
    1. CairoSVG (best quality, pure Python)
    2. Inkscape (if available)
    3. svglib + reportlab (fallback)
    """
    if output_path is None:
        output_path = os.path.splitext(svg_path)[0] + ".png"

    # Try CairoSVG first
    if _has_cairosvg():
        import cairosvg
        kwargs = {"output_width": width, "output_height": height} if width or height else {}
        cairosvg.svg2png(url=svg_path, write_to=output_path, dpi=dpi, **kwargs)
        return output_path

    # Try Inkscape
    if _has_inkscape():
        cmd = ["inkscape", svg_path, "--export-type=png", f"--export-filename={output_path}"]
        if dpi:
            cmd.append(f"--export-dpi={dpi}")
        if width:
            cmd.append(f"--export-width={width}")
        if height:
            cmd.append(f"--export-height={height}")
        subprocess.run(cmd, capture_output=True, check=True)
        return output_path

    # Fallback to svglib + reportlab
    if _has_svglib():
        from svglib.svglib import svg2rlg
        from reportlab.graphics import renderPM
        drawing = svg2rlg(svg_path)
        if drawing:
            renderPM.drawToFile(drawing, output_path, fmt="PNG")
            return output_path

    raise RuntimeError(
        "No SVG converter available. Install one of:\n"
        "  pip install cairosvg      (recommended)\n"
        "  apt install inkscape      (system package)\n"
        "  pip install svglib reportlab"
    )


def add_svg_to_slide(slide, svg_path, left, top, width, height, **kwargs):
    """Add an SVG image to a slide, converting on-the-fly.

    Args:
        slide: python-pptx Slide object
        svg_path: Path to SVG file
        left, top, width, height: Position and size in EMU
        **kwargs: Passed to svg_to_image (dpi, etc.)

    Returns:
        The picture shape
    """
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        png_path = tmp.name

    try:
        svg_to_image(svg_path, png_path, **kwargs)
        shape = slide.shapes.add_picture(png_path, left, top, width, height)
        return shape
    finally:
        if os.path.exists(png_path):
            os.remove(png_path)


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Convert SVG to PNG")
    parser.add_argument("svg", help="Input SVG file")
    parser.add_argument("--output", "-o", help="Output PNG file")
    parser.add_argument("--width", type=int, help="Output width in pixels")
    parser.add_argument("--height", type=int, help="Output height in pixels")
    parser.add_argument("--dpi", type=int, default=150, help="DPI")
    args = parser.parse_args()

    out = svg_to_image(args.svg, args.output, args.width, args.height, args.dpi)
    print(f"Converted: {out}")
