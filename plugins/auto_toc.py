"""
Plugin: Auto Table of Contents

If deck_spec has "generate_toc": true, automatically inserts
a TOC slide after the cover slide.
"""


def pre_build(deck_spec, context):
    """Insert TOC slide if requested."""
    if not deck_spec.get("generate_toc"):
        return deck_spec

    slides = deck_spec.get("slides", [])
    if not slides:
        return deck_spec

    # Collect section titles from slides
    toc_items = []
    for i, slide in enumerate(slides):
        layout = slide.get("layout", "")
        title = slide.get("title", "")
        if title and layout not in ("cover", "closing"):
            toc_items.append(title)

    if not toc_items:
        return deck_spec

    # Find cover index
    cover_idx = -1
    for i, slide in enumerate(slides):
        if slide.get("layout") == "cover":
            cover_idx = i
            break

    toc_slide = {
        "layout": "toc",
        "title": deck_spec.get("toc_title", "目录"),
        "items": toc_items[:10]  # limit to 10 items
    }

    insert_idx = cover_idx + 1 if cover_idx >= 0 else 0
    slides.insert(insert_idx, toc_slide)
    deck_spec["slides"] = slides

    return deck_spec
