#!/usr/bin/env python3
"""Intent router: infer the best layout from slide content fields.

This lets LLMs write simpler slide specs without memorizing layout names.
"""

# Ordered by specificity (most specific first)
INTENT_RULES = [
    # Quote / testimonial
    {"fields": ["quote", "author"], "layout": "quote", "confidence": 1.0},
    {"fields": ["quote"], "layout": "quote", "confidence": 0.9},

    # Data highlight
    {"fields": ["big_number", "label"], "layout": "data_highlight", "confidence": 1.0},
    {"fields": ["big_number"], "layout": "data_highlight", "confidence": 0.8},

    # Two-column comparison
    {"fields": ["left_content", "right_content"], "layout": "two_column", "confidence": 1.0},

    # Image + content
    {"fields": ["image_path", "content"], "layout": "image_content", "confidence": 1.0},
    {"fields": ["image_path"], "layout": "image_content", "confidence": 0.7},

    # Chart
    {"fields": ["chart_data"], "layout": "chart", "confidence": 1.0},

    # Table
    {"fields": ["table_data"], "layout": "table", "confidence": 1.0},

    # Timeline
    {"fields": ["timeline_data"], "layout": "timeline", "confidence": 1.0},

    # Team
    {"fields": ["team_data"], "layout": "team", "confidence": 1.0},

    # Process
    {"fields": ["process_data"], "layout": "process", "confidence": 1.0},

    # TOC
    {"fields": ["items"], "layout": "toc", "confidence": 0.8},

    # Cover
    {"fields": ["title", "subtitle", "date"], "layout": "cover", "confidence": 0.9},
    {"fields": ["title", "subtitle"], "layout": "cover", "confidence": 0.8},

    # Closing
    {"fields": ["title"], "layout": "title_content", "confidence": 0.5},
]


def infer_layout(slide_spec):
    """Infer the best layout for a slide spec.

    Returns (layout_name, confidence) or (None, 0) if no match.
    """
    # If layout is explicitly specified, trust it
    explicit = slide_spec.get("layout")
    if explicit:
        return explicit, 1.0

    # Otherwise, infer from content fields
    fields = set(slide_spec.keys())
    best_match = None
    best_confidence = 0.0

    for rule in INTENT_RULES:
        required = set(rule["fields"])
        if required.issubset(fields):
            if rule["confidence"] > best_confidence:
                best_match = rule["layout"]
                best_confidence = rule["confidence"]

    return best_match, best_confidence


def auto_route(deck_spec):
    """Auto-route all slides in a deck spec.

    Adds inferred layout to slides that don't have one.
    Returns a new deck spec with layouts filled in.
    """
    new_deck = dict(deck_spec)
    new_slides = []
    for slide in deck_spec.get("slides", []):
        new_slide = dict(slide)
        layout, confidence = infer_layout(new_slide)
        if layout and not new_slide.get("layout"):
            new_slide["layout"] = layout
            new_slide["_inferred"] = True
            new_slide["_confidence"] = confidence
        new_slides.append(new_slide)
    new_deck["slides"] = new_slides
    return new_deck


if __name__ == "__main__":
    import json
    test_specs = [
        {"title": "Hello", "subtitle": "World", "date": "2026"},
        {"quote": "Best code is no code", "author": "Dev"},
        {"big_number": "99.9%", "label": "Uptime"},
        {"left_content": ["A"], "right_content": ["B"]},
        {"title": "Chart", "chart_data": {"type": "column"}},
        {"title": "Table", "table_data": {"headers": ["A"], "rows": [[1]]}},
        {"title": "Normal", "content": ["A", "B"]},
    ]
    for spec in test_specs:
        layout, conf = infer_layout(spec)
        print(f"{spec.keys()} -> {layout} (conf={conf})")
