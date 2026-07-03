#!/usr/bin/env python3
"""
Validate a deck JSON/YAML file against the schema.

Usage:
    python scripts/validate_deck.py deck.json
    python scripts/validate_deck.py deck.yaml
"""

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Simple schema validation without external dependencies
VALID_LAYOUTS = {
    "cover", "toc", "title_content", "two_column", "chart", "table",
    "timeline", "image_content", "quote", "team", "data_highlight",
    "process", "tree", "waterfall", "funnel", "gantt", "wordcloud", "closing"
}

CHART_TYPES = {
    "column", "bar", "line", "pie", "doughnut", "area", "scatter", "radar", "bubble"
}


def validate_deck(deck_spec):
    """Validate deck spec and return list of errors."""
    errors = []
    warnings = []

    if not isinstance(deck_spec, dict):
        errors.append("Deck spec must be a JSON object")
        return errors, warnings

    slides = deck_spec.get("slides", [])
    if not isinstance(slides, list):
        errors.append("'slides' must be an array")
        return errors, warnings

    if not slides:
        warnings.append("Deck has no slides")

    for i, slide in enumerate(slides):
        prefix = f"Slide {i + 1}"
        if not isinstance(slide, dict):
            errors.append(f"{prefix}: must be an object")
            continue

        layout = slide.get("layout")
        if layout and layout not in VALID_LAYOUTS:
            errors.append(f"{prefix}: unknown layout '{layout}'. Valid: {', '.join(sorted(VALID_LAYOUTS))}")

        # Layout-specific validation
        if layout == "chart":
            chart_data = slide.get("chart_data")
            if not chart_data:
                errors.append(f"{prefix}: 'chart_data' is required for chart layout")
            elif not isinstance(chart_data, dict):
                errors.append(f"{prefix}: 'chart_data' must be an object")
            else:
                chart_type = chart_data.get("type")
                if chart_type and chart_type not in CHART_TYPES:
                    errors.append(f"{prefix}: unknown chart type '{chart_type}'. Valid: {', '.join(sorted(CHART_TYPES))}")
                if "categories" not in chart_data:
                    warnings.append(f"{prefix}: 'chart_data.categories' is missing")
                if "values" not in chart_data:
                    warnings.append(f"{prefix}: 'chart_data.values' is missing")

        elif layout == "table":
            table_data = slide.get("table_data")
            if not table_data:
                errors.append(f"{prefix}: 'table_data' is required for table layout")
            elif not isinstance(table_data, dict):
                errors.append(f"{prefix}: 'table_data' must be an object")
            else:
                if "headers" not in table_data:
                    warnings.append(f"{prefix}: 'table_data.headers' is missing")
                if "rows" not in table_data:
                    warnings.append(f"{prefix}: 'table_data.rows' is missing")

        elif layout == "timeline":
            timeline_data = slide.get("timeline_data")
            if not timeline_data:
                errors.append(f"{prefix}: 'timeline_data' is required for timeline layout")
            elif not isinstance(timeline_data, dict):
                errors.append(f"{prefix}: 'timeline_data' must be an object")
            elif "milestones" not in timeline_data:
                warnings.append(f"{prefix}: 'timeline_data.milestones' is missing")

        elif layout == "tree":
            tree_data = slide.get("tree_data")
            if not tree_data:
                errors.append(f"{prefix}: 'tree_data' is required for tree layout")
            elif not isinstance(tree_data, dict):
                errors.append(f"{prefix}: 'tree_data' must be an object")
            elif "root" not in tree_data:
                errors.append(f"{prefix}: 'tree_data.root' is required")

        elif layout == "team":
            team_data = slide.get("team_data")
            if not team_data:
                errors.append(f"{prefix}: 'team_data' is required for team layout")
            elif not isinstance(team_data, dict):
                errors.append(f"{prefix}: 'team_data' must be an object")
            elif "members" not in team_data:
                warnings.append(f"{prefix}: 'team_data.members' is missing")

        elif layout == "image_content":
            if "image_path" not in slide:
                warnings.append(f"{prefix}: 'image_path' is missing for image_content layout")

        # Check for common issues
        if slide.get("accent_line") is not None and not isinstance(slide.get("accent_line"), bool):
            errors.append(f"{prefix}: 'accent_line' must be a boolean")

        if slide.get("shapes") is not None and not isinstance(slide.get("shapes"), list):
            errors.append(f"{prefix}: 'shapes' must be an array")

    return errors, warnings


def main():
    parser = argparse.ArgumentParser(description="Validate deck JSON/YAML")
    parser.add_argument("deck", help="Deck file (JSON or YAML)")
    parser.add_argument("--strict", action="store_true", help="Treat warnings as errors")
    args = parser.parse_args()

    with open(args.deck, "r", encoding="utf-8") as f:
        if args.deck.lower().endswith((".yaml", ".yml")):
            try:
                import yaml
                deck_spec = yaml.safe_load(f)
            except ImportError:
                print("Error: PyYAML is required for YAML files")
                sys.exit(1)
        else:
            deck_spec = json.load(f)

    errors, warnings = validate_deck(deck_spec)

    print(f"Validation results for {args.deck}:")
    print(f"  Slides: {len(deck_spec.get('slides', []))}")
    print(f"  Errors: {len(errors)}")
    print(f"  Warnings: {len(warnings)}")

    if errors:
        print("\nErrors:")
        for e in errors:
            print(f"  [ERROR] {e}")

    if warnings:
        print("\nWarnings:")
        for w in warnings:
            print(f"  [WARN] {w}")

    if errors or (args.strict and warnings):
        print("\nValidation FAILED")
        sys.exit(1)
    else:
        print("\nValidation PASSED")
        sys.exit(0)


if __name__ == "__main__":
    main()
