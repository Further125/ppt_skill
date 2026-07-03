#!/usr/bin/env python3
"""
Conditional logic and loops for deck specs.

Supports:
- Variables: deck_spec["vars"] = {"version": "1.0", "show_charts": true}
- Conditions: slide{"if": "show_charts"} or slide{"if": "slides.length > 5"}
- Loops: slide{"for": "items", "as": "item", "template": {...}}

Usage:
    from deck_logic import evaluate_deck
    expanded_deck = evaluate_deck(deck_spec)
"""

import copy
import re


def _resolve_value(expr, context):
    """Resolve an expression in the given context.

    Supports:
    - Direct variable lookup: "show_charts"
    - Comparison: "slides.length > 5"
    - Equality: "theme == 'dark'"
    - Boolean: "has_charts and has_tables"
    """
    # Direct boolean
    if expr in context:
        val = context[expr]
        if isinstance(val, bool):
            return val
        return bool(val)

    # Comparison patterns
    # slides.length > 5
    match = re.match(r'^(\w+(?:\.\w+)*)\s*([<>!=]+)\s*(.+)$', expr.strip())
    if match:
        left_expr, op, right_expr = match.groups()
        left_val = _get_nested(context, left_expr)
        right_val = _parse_literal(right_expr.strip())

        if op == '>':
            return left_val > right_val
        elif op == '>=':
            return left_val >= right_val
        elif op == '<':
            return left_val < right_val
        elif op == '<=':
            return left_val <= right_val
        elif op == '==':
            return left_val == right_val
        elif op == '!=':
            return left_val != right_val

    # Boolean expressions: "a and b", "a or b"
    if ' and ' in expr.lower():
        parts = re.split(r'\s+and\s+', expr, flags=re.IGNORECASE)
        return all(_resolve_value(p.strip(), context) for p in parts)
    if ' or ' in expr.lower():
        parts = re.split(r'\s+or\s+', expr, flags=re.IGNORECASE)
        return any(_resolve_value(p.strip(), context) for p in parts)

    # Default: truthy check
    return bool(expr)


def _get_nested(obj, path):
    """Get a nested value by dot path."""
    parts = path.split('.')
    current = obj
    for part in parts:
        if isinstance(current, dict):
            current = current.get(part)
        elif isinstance(current, list):
            if part == 'length':
                return len(current)
            try:
                current = current[int(part)]
            except (ValueError, IndexError):
                return None
        else:
            return None
        if current is None:
            return None
    return current


def _parse_literal(s):
    """Parse a literal value from string."""
    s = s.strip()
    if s.lower() == 'true':
        return True
    if s.lower() == 'false':
        return False
    if s.lower() == 'none' or s.lower() == 'null':
        return None
    if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
        return s[1:-1]
    try:
        if '.' in s:
            return float(s)
        return int(s)
    except ValueError:
        return s


def _apply_filter(val, filter_name):
    """Apply a simple filter to a value."""
    if filter_name == "length":
        if isinstance(val, (list, dict, str)):
            return len(val)
        return 0
    if filter_name == "upper":
        return str(val).upper()
    if filter_name == "lower":
        return str(val).lower()
    if filter_name == "trim":
        return str(val).strip()
    return val


def _interpolate_string(s, context):
    """Replace {{var}} and {{var | filter}} placeholders in strings."""
    if not isinstance(s, str):
        return s

    def replacer(match):
        expr = match.group(1).strip()
        # Check for filter syntax: var | filter
        if "|" in expr:
            parts = expr.split("|")
            var_expr = parts[0].strip()
            filter_name = parts[1].strip()
            val = _get_nested(context, var_expr)
            if val is None:
                return match.group(0)
            val = _apply_filter(val, filter_name)
            return str(val)
        else:
            val = _get_nested(context, expr)
            if val is None:
                return match.group(0)
            return str(val)

    return re.sub(r'\{\{\s*(.+?)\s*\}\}', replacer, s)


def _interpolate_dict(d, context):
    """Recursively interpolate all strings in a dict."""
    if isinstance(d, dict):
        return {k: _interpolate_dict(v, context) for k, v in d.items()}
    elif isinstance(d, list):
        return [_interpolate_dict(item, context) for item in d]
    elif isinstance(d, str):
        return _interpolate_string(d, context)
    return d


def evaluate_deck(deck_spec):
    """Evaluate all conditions and loops in a deck spec.

    Returns a new deck spec with all conditions resolved and loops expanded.
    """
    deck = copy.deepcopy(deck_spec)
    vars_dict = deck.get("vars", {})

    # Build evaluation context
    context = {
        **vars_dict,
        "slides": deck.get("slides", []),
        "title": deck.get("title", ""),
    }

    # Process slides
    processed_slides = []
    for slide in deck.get("slides", []):
        # Check condition
        if "if" in slide:
            condition = slide.pop("if")
            if not _resolve_value(condition, context):
                continue

        # Process loops
        if "for" in slide:
            loop_result = _process_loop(slide, context)
            processed_slides.extend(loop_result)
            continue

        # Interpolate variables in the slide
        slide = _interpolate_dict(slide, context)
        processed_slides.append(slide)

    deck["slides"] = processed_slides
    return deck


def _process_loop(slide, context):
    """Process a for-loop slide."""
    loop_spec = slide.pop("for")
    items_expr = loop_spec if isinstance(loop_spec, str) else loop_spec.get("items")
    var_name = loop_spec.get("as", "item") if isinstance(loop_spec, dict) else "item"
    template = slide

    items = _get_nested(context, items_expr)
    if items is None:
        return []
    if not isinstance(items, list):
        items = [items]

    results = []
    for i, item in enumerate(items):
        item_context = dict(context)
        item_context[var_name] = item
        item_context[f"{var_name}_index"] = i
        item_context[f"{var_name}_first"] = (i == 0)
        item_context[f"{var_name}_last"] = (i == len(items) - 1)

        expanded = copy.deepcopy(template)
        expanded = _interpolate_dict(expanded, item_context)
        results.append(expanded)

    return results


if __name__ == "__main__":
    import json

    test_deck = {
        "title": "Report",
        "vars": {
            "show_charts": True,
            "version": "2.0",
            "regions": ["APAC", "EMEA", "AMER"]
        },
        "slides": [
            {"layout": "cover", "title": "Report v{{ version }}"},
            {"layout": "chart", "if": "show_charts", "title": "Overview", "chart_data": {"type": "column", "categories": ["Q1"], "values": [100]}},
            {"layout": "title_content", "for": {"items": "regions", "as": "region"}, "title": "{{ region }} Region", "content": ["Analysis for {{ region }}"]},
            {"layout": "closing", "if": "slides.length > 3", "title": "Thank You"}
        ]
    }

    result = evaluate_deck(test_deck)
    print(json.dumps(result, ensure_ascii=False, indent=2))
