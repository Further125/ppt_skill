#!/usr/bin/env python3
"""Snapshot: export a specific page (or a clipped region) from a PPTX/PDF.

Usage:
  # Full page 3 as PNG
  python snapshot.py deck.pptx --page 3 --output page3.png

  # Clip a region (x, y, width, height in pixels at 144 DPI)
  python snapshot.py deck.pptx --page 3 --crop "100,200,400,300" --output clip.png

  # Clip using EMU coordinates (same unit as free shapes)
  python snapshot.py deck.pptx --page 3 --crop-emu "1000000,2000000,3000000,4000000" --output clip.png

  # From an existing PDF
  python snapshot.py deck.pdf --page 3 --output page3.png
"""
import sys, os, argparse, subprocess, shutil, hashlib


try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

from pptx import Presentation

DEFAULT_DPI = 144
EMU_PER_INCH = 914400


def _libreoffice_env():
    env = dict(os.environ)
    env["HOME"] = "/tmp"
    return env


def _pptx_to_pdf(pptx_path):
    """Convert PPTX to PDF (cached in /tmp). Returns PDF path."""
    base = os.path.basename(pptx_path)
    cache_key = hashlib.md5((os.path.abspath(pptx_path) + str(os.path.getmtime(pptx_path))).encode()).hexdigest()
    cached_pdf = os.path.join("/tmp", f"ppt_skill_snapshot_{cache_key}.pdf")

    if os.path.exists(cached_pdf):
        return cached_pdf

    if not shutil.which("soffice"):
        raise RuntimeError("LibreOffice (soffice) not found. Cannot convert PPTX to PDF.")

    out_dir = "/tmp"
    cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, pptx_path]
    profile_dir = os.path.join("/tmp", f"lo_profile_{os.getpid()}")
    os.makedirs(profile_dir, exist_ok=True)
    cmd.append(f"-env:UserInstallation=file://{profile_dir}")

    result = subprocess.run(cmd, capture_output=True, text=True, env=_libreoffice_env(), timeout=120)
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

    generated = os.path.join(out_dir, base.replace(".pptx", ".pdf").replace(".PPTX", ".pdf"))
    if not os.path.exists(generated):
        raise RuntimeError(f"LibreOffice did not produce expected PDF: {generated}")

    os.rename(generated, cached_pdf)
    return cached_pdf


def _emu_to_px(emu, dpi):
    return int(emu / EMU_PER_INCH * dpi)


def _px_to_emu(px, dpi):
    return int(px / dpi * EMU_PER_INCH)


def _parse_crop(s, dpi):
    """Parse 'x,y,w,h' string and return (x, y, w, h) in pixels."""
    parts = [float(p.strip()) for p in s.split(",")]
    if len(parts) != 4:
        raise ValueError("crop must be 'x,y,w,h'")
    return tuple(int(p) for p in parts)


def _parse_crop_emu(s, dpi):
    """Parse 'x,y,w,h' EMU string and convert to pixels."""
    parts = [float(p.strip()) for p in s.split(",")]
    if len(parts) != 4:
        raise ValueError("crop-emu must be 'x,y,w,h'")
    return tuple(_emu_to_px(int(p), dpi) for p in parts)


def snapshot(pdf_path, page_num, crop_rect, output_path, dpi=DEFAULT_DPI):
    if not HAS_PYMUPDF:
        raise RuntimeError("PyMuPDF (fitz) is required. Install: pip install pymupdf")

    doc = fitz.open(pdf_path)
    if page_num < 1 or page_num > len(doc):
        raise ValueError(f"Page {page_num} out of range (1-{len(doc)})")

    page = doc[page_num - 1]

    # Render page at specified DPI
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat)

    if crop_rect:
        x, y, w, h = crop_rect
        scale = dpi / 72.0
        # Convert pixel coords to page points for fitz clip
        clip_rect = fitz.Rect(
            x / scale, y / scale,
            (x + w) / scale, (y + h) / scale,
        )
        # Clamp to page rect
        page_rect = page.rect
        clip_rect &= page_rect
        if clip_rect.is_empty:
            raise ValueError("Crop region is empty or outside page bounds")

        pix = page.get_pixmap(matrix=mat, clip=clip_rect)

    pix.save(output_path)
    doc.close()
    print(f"Snapshot saved: {output_path} ({pix.width}x{pix.height} px, {dpi} DPI)")


def list_pages(pdf_path):
    """List all pages with dimensions."""
    if not HAS_PYMUPDF:
        raise RuntimeError("PyMuPDF (fitz) is required.")
    doc = fitz.open(pdf_path)
    for i in range(len(doc)):
        page = doc[i]
        rect = page.rect
        print(f"Page {i+1}: {rect.width:.1f} x {rect.height:.1f} pt")
    doc.close()


def main():
    parser = argparse.ArgumentParser(description="Snapshot a PPTX/PDF page or region")
    parser.add_argument("input", help="Input PPTX or PDF file")
    parser.add_argument("--page", "-p", type=int, required=True, help="Page number (1-based)")
    parser.add_argument("--output", "-o", required=True, help="Output PNG file path")
    parser.add_argument("--crop", help="Crop region: 'x,y,w,h' in pixels at render DPI")
    parser.add_argument("--crop-emu", help="Crop region: 'x,y,w,h' in EMU (same as PPTX coords)")
    parser.add_argument("--dpi", type=int, default=DEFAULT_DPI, help=f"Render DPI (default {DEFAULT_DPI})")
    parser.add_argument("--list", action="store_true", help="List page dimensions and exit")
    args = parser.parse_args()

    input_path = args.input
    if input_path.lower().endswith(".pptx"):
        pdf_path = _pptx_to_pdf(input_path)
    else:
        pdf_path = input_path

    if args.list:
        list_pages(pdf_path)
        return

    crop = None
    if args.crop_emu:
        crop = _parse_crop_emu(args.crop_emu, args.dpi)
    elif args.crop:
        crop = _parse_crop(args.crop, args.dpi)

    snapshot(pdf_path, args.page, crop, args.output, args.dpi)


if __name__ == "__main__":
    main()
