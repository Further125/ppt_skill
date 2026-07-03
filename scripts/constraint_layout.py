#!/usr/bin/env python3
"""Constraint-based layout solver for ppt-skill.

Given a set of shapes with text content, font scales, and width constraints,
finds the maximum base font size that lets all text fit, then auto-computes
each shape's height and position based on attach chains.
"""

import os, sys, copy
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Import text measurement from text_fitter
try:
    sys.path.insert(0, SCRIPT_DIR)
    from text_fitter import _pptx_wrap_text, _pptx_text_height, _pptx_char_width, _emu_to_px
except ImportError:
    _pptx_wrap_text = _pptx_text_height = _pptx_char_width = _emu_to_px = None


def _text_content_to_string(content):
    """Flatten content (list/str/dict) to a single string for measurement.
    
    Adds '• ' prefix to match render_text_shape behavior for list items.
    """
    if not content:
        return ""
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for item in content:
            if isinstance(item, str):
                parts.append("• " + item)
            elif isinstance(item, dict):
                parts.append(item.get("text", ""))
        return "\n".join(parts)
    return str(content)


def _compute_image_height(image_path, width_emu, max_height_emu, aspect_ratio=None):
    """Compute image height preserving aspect ratio."""
    if aspect_ratio is not None and aspect_ratio > 0:
        h_emu = int(width_emu * aspect_ratio)
        return min(h_emu, max_height_emu)
    try:
        from PIL import Image
        if not os.path.exists(image_path):
            return max_height_emu  # fallback: use max height
        with Image.open(image_path) as img:
            img_w, img_h = img.size
            if img_w > 0:
                aspect = img_h / img_w
                h_emu = int(width_emu * aspect)
                return min(h_emu, max_height_emu)
    except Exception:
        pass
    return max_height_emu


def _measure_required_height(text, font_size_pt, max_width_emu, line_spacing=1.40):
    """Return required height in EMU for text at given font size and width.
    
    Uses configurable line_spacing with NO para_gap (matches render_text_shape
    which sets space_before=space_after=0).
    """
    if not _pptx_wrap_text or not text.strip():
        return int(font_size_pt * line_spacing * 12700)  # fallback: one line
    max_w_pt = _emu_to_px(max_width_emu) * 0.75
    lines = _pptx_wrap_text(text, font_size_pt, max_w_pt)
    if not lines:
        return 0
    # Line height only (no para_gap since we set space_before/after=0)
    line_h = font_size_pt * line_spacing
    h_pt = line_h * len(lines)
    # pt -> EMU (1 pt = 12700 EMU)
    return int(h_pt * 12700)


def _truncate_text_to_height(text, font_size_pt, max_width_emu, max_height_emu, line_spacing=1.40):
    """Truncate text to fit within max_height, return (truncated_text, was_truncated, kept_lines, total_lines)."""
    if not _pptx_wrap_text or not text.strip():
        return text, False, 0, 0
    max_w_pt = _emu_to_px(max_width_emu) * 0.75
    lines = _pptx_wrap_text(text, font_size_pt, max_w_pt)
    if not lines:
        return text, False, 0, 0
    line_h = font_size_pt * line_spacing
    max_h_pt = max_height_emu / 12700

    total_h = 0
    keep_lines = []
    for line in lines:
        if total_h + line_h <= max_h_pt:
            keep_lines.append(line)
            total_h += line_h
        else:
            break

    was_truncated = len(keep_lines) < len(lines)
    truncated = "\n".join(keep_lines) if was_truncated else text
    return truncated, was_truncated, len(keep_lines), len(lines)


def _topological_sort(shapes_defs):
    """Sort shapes by attach dependency (parents before children)."""
    # Build adjacency list
    role_to_idx = {s.get("role", f"__{i}"): i for i, s in enumerate(shapes_defs)}
    in_degree = [0] * len(shapes_defs)
    children = [[] for _ in range(len(shapes_defs))]

    for i, s in enumerate(shapes_defs):
        attach = s.get("attach")
        if attach and attach.get("to") in role_to_idx:
            parent_idx = role_to_idx[attach["to"]]
            children[parent_idx].append(i)
            in_degree[i] += 1

    # Kahn's algorithm
    queue = [i for i, d in enumerate(in_degree) if d == 0]
    sorted_idx = []
    while queue:
        idx = queue.pop(0)
        sorted_idx.append(idx)
        for child in children[idx]:
            in_degree[child] -= 1
            if in_degree[child] == 0:
                queue.append(child)

    # Append any remaining (cycles - shouldn't happen)
    for i, d in enumerate(in_degree):
        if d > 0 and i not in sorted_idx:
            sorted_idx.append(i)

    return [shapes_defs[i] for i in sorted_idx]


def solve_constraints(shapes_defs, slide_spec, slide_width, slide_height,
                      resolve_region_fn, resolve_style_fn, theme=None):
    """Solve layout constraints and return computed (region, style) for each shape.

    Returns: list of (shape_def, computed_region, computed_style)
    """
    if not _pptx_wrap_text:
        # Fallback: return original regions
        return [(s, resolve_region_fn(s.get("region", "full"), slide_width, slide_height),
                 resolve_style_fn(s.get("style", "body"), theme))
                for s in shapes_defs]

    # 1. Topological sort
    sorted_shapes = _topological_sort(shapes_defs)

    # 2. Build role -> shape mapping for original order
    role_to_shape = {s.get("role", f"__{i}"): s for i, s in enumerate(shapes_defs)}

    # 3. Collect textbox constraints + image aspect ratios
    textboxes = []  # list of textbox constraint dicts
    image_shapes = []  # list of image shape dicts

    for s in sorted_shapes:
        shape_type = s.get("type", "textbox")
        region_spec = s.get("region", "full")
        region = resolve_region_fn(region_spec, slide_width, slide_height)

        if shape_type == "textbox":
            content_source = s.get("content_source")
            content = slide_spec.get(content_source, "") if content_source else s.get("content", "")
            text = _text_content_to_string(content)
            if text.strip():
                width_emu = region["width"]
                max_height_emu = region["height"]
                style = resolve_style_fn(s.get("style", "body"), theme)
                font_scale = style.get("font_scale", 1.0)
                min_height_emu = region.get("min_height", 0)
                textboxes.append({
                    "role": s.get("role", ""),
                    "text": text,
                    "width_emu": width_emu,
                    "max_height_emu": max_height_emu,
                    "min_height_emu": min_height_emu,
                    "font_scale": font_scale,
                    "max_font_size": style.get("max_font_size"),
                    "no_wrap": s.get("no_wrap", False),
                    "uniform_group": s.get("uniform_group"),
                    "margin_left": 91440,
                    "margin_right": 91440,
                    "margin_top": 45720,
                    "margin_bottom": 45720,
                })

        elif shape_type == "image":
            content_source = s.get("content_source")
            image_path = slide_spec.get(content_source, "") if content_source else s.get("content", "")
            aspect_ratio = s.get("aspect_ratio")
            if image_path or aspect_ratio:
                image_shapes.append({
                    "role": s.get("role", ""),
                    "path": str(image_path) if image_path else "",
                    "width_emu": region["width"],
                    "max_height_emu": region["height"],
                    "min_height_emu": region.get("min_height", 0),
                    "aspect_ratio": aspect_ratio,
                })

    if not textboxes:
        # No textboxes, return original
        return [(s, resolve_region_fn(s.get("region", "full"), slide_width, slide_height),
                 resolve_style_fn(s.get("style", "body"), theme))
                for s in shapes_defs]

    # 4. Binary search for max base font size
    # Respect per-textbox max_font_size limits
    def _run_solver(textboxes_list, line_spacing=1.40):
        global_max_base = 96
        for tb in textboxes_list:
            if tb.get("max_font_size"):
                max_base_for_tb = int(tb["max_font_size"] / tb["font_scale"])
                global_max_base = min(global_max_base, max_base_for_tb)
        lo, hi = 8, global_max_base
        best_base = 8

        def _fits(base_size):
            for tb in textboxes_list:
                actual_size = base_size * tb["font_scale"]
                actual_size = max(actual_size, 8)
                # Enforce max_font_size per textbox
                if tb.get("max_font_size") and actual_size > tb["max_font_size"]:
                        return False
                # Deduct text frame margins for accurate fit
                content_w = max(tb["width_emu"] - tb["margin_left"] - tb["margin_right"], 100000)
                content_h = max(tb["max_height_emu"] - tb["margin_top"] - tb["margin_bottom"], 100000)
                if tb.get("no_wrap"):
                    # No-wrap: text must fit in a single line
                    total_w = sum(_pptx_char_width(ch, actual_size) for ch in tb["text"])
                    max_w_pt = _emu_to_px(content_w) * 0.75
                    if total_w > max_w_pt:
                        return False
                    # Single line height must fit
                    line_h = actual_size * line_spacing * 12700
                    if line_h > content_h:
                        return False
                else:
                    h = _measure_required_height(tb["text"], actual_size, content_w, line_spacing)
                    if h > content_h:
                        return False
                if tb["min_height_emu"] > 0:
                    min_h = _measure_required_height(tb["text"], actual_size, content_w, line_spacing)
                    if min_h < tb["min_height_emu"] and min_h < content_h:
                        pass
            return True

        while lo <= hi:
            mid = (lo + hi) // 2
            if _fits(mid):
                best_base = mid
                lo = mid + 1
            else:
                hi = mid - 1
        return best_base

    active_line_spacing = 1.40
    best_base = _run_solver(textboxes, line_spacing=active_line_spacing)

    # Shrink strategy: if best_base is too small, try tighter line spacing
    if best_base < 14:
        best_base_tight = _run_solver(textboxes, line_spacing=1.20)
        if best_base_tight > best_base:
            best_base = best_base_tight
            active_line_spacing = 1.20
            print(f"  [shrink] line spacing reduced 1.40→1.20, best_base={best_base}pt")
        # If still too small, reduce all font_scales and retry
        if best_base < 10:
            shrunk_textboxes = []
            for tb in textboxes:
                stb = dict(tb)
                stb["font_scale"] = tb["font_scale"] * 0.85
                shrunk_textboxes.append(stb)
            best_base_shrunk = _run_solver(shrunk_textboxes, line_spacing=1.20)
            if best_base_shrunk >= 10:
                textboxes = shrunk_textboxes
                best_base = best_base_shrunk
                active_line_spacing = 1.20
                print(f"  [shrink] font_scale ×0.85 + line spacing 1.20, best_base={best_base}pt")
            else:
                print(f"  [shrink] content too dense even after shrink (best_base={best_base}pt). Consider reducing content.")


    # 5. Compute actual heights for textboxes + images + overflow detection
    def _expand_font_size(text, width_emu, max_height_emu, start_size, max_font_size=None):
        """Find larger font size that fills ~60% of max_height without overflow."""
        if not text.strip():
            return start_size
        upper = min(96, max_font_size) if max_font_size else 96
        lo, hi = int(start_size), upper
        best = start_size
        target_h = max_height_emu * 0.60
        while lo <= hi:
            mid = (lo + hi) // 2
            h = _measure_required_height(text, mid, width_emu)
            if h <= max_height_emu and h >= target_h:
                best = mid
                lo = mid + 1
            elif h < target_h:
                best = mid
                lo = mid + 1
            else:
                hi = mid - 1
        return best

    def _expand_font_size_nowrap(text, width_emu, start_size, max_font_size=None):
        """Find larger font size that fills ~90% of width without wrapping."""
        if not text.strip():
            return start_size
        upper = min(96, max_font_size) if max_font_size else 96
        lo, hi = int(start_size), upper
        best = start_size
        max_w_pt = _emu_to_px(width_emu) * 0.75
        target_w = max_w_pt * 0.90
        while lo <= hi:
            mid = (lo + hi) // 2
            total_w = sum(_pptx_char_width(ch, mid) for ch in text)
            if total_w <= max_w_pt and total_w >= target_w:
                best = mid
                lo = mid + 1
            elif total_w < target_w:
                best = mid
                lo = mid + 1
            else:
                hi = mid - 1
        return best

    role_to_height = {}
    role_to_truncated_text = {}  # role -> truncated text string
    overflow_warnings = []

    for tb in textboxes:
        actual_size = best_base * tb["font_scale"]
        actual_size = max(actual_size, 8)
        content_w = max(tb["width_emu"] - tb["margin_left"] - tb["margin_right"], 100000)
        content_h = max(tb["max_height_emu"] - tb["margin_top"] - tb["margin_bottom"], 100000)
        h = _measure_required_height(tb["text"], actual_size, content_w, active_line_spacing)

        # Check overflow at minimum size (8pt)
        min_size = 8 * tb["font_scale"]
        min_size = max(min_size, 8)
        min_h = _measure_required_height(tb["text"], min_size, content_w, active_line_spacing)

        if min_h > content_h:
            # Overflow even at minimum size - truncate
            truncated, was_trunc, kept, total = _truncate_text_to_height(
                tb["text"], min_size, content_w, content_h, active_line_spacing
            )
            role_to_truncated_text[tb["role"]] = truncated
            overflow_warnings.append(
                f"  [overflow] '{tb['role']}' truncated: {kept}/{total} lines kept at {min_size:.1f}pt"
            )
            # Recalculate height with truncated text
            h = _measure_required_height(truncated, actual_size, content_w, active_line_spacing)
        elif h > content_h:
            # Overflow at current best_base but not at minimum
            h = content_h

        # Expand: try larger font to fill available space
        content_w = max(tb["width_emu"] - tb["margin_left"] - tb["margin_right"], 100000)
        content_h = max(tb["max_height_emu"] - tb["margin_top"] - tb["margin_bottom"], 100000)
        if tb.get("no_wrap"):
            # No-wrap expand: fill width up to 90%
            expanded = _expand_font_size_nowrap(tb["text"], content_w, actual_size, tb.get("max_font_size"))
            if expanded > actual_size:
                actual_size = expanded
                h = int(actual_size * active_line_spacing * 12700)
        elif h < tb["max_height_emu"] * 0.50:
            expanded = _expand_font_size(tb["text"], content_w, content_h, actual_size, tb.get("max_font_size"))
            if expanded > actual_size:
                actual_size = expanded
                h = _measure_required_height(tb["text"], actual_size, content_w, active_line_spacing)

        # Enforce min_height
        if tb["min_height_emu"] > 0:
            h = max(h, tb["min_height_emu"])
        role_to_height[tb["role"]] = min(h, tb["max_height_emu"])
        tb["computed_size"] = actual_size

    # 5b. Uniform font size groups: all members share the minimum computed size
    uniform_font_groups = {}
    for tb in textboxes:
        group = tb.get("uniform_group")
        if group:
            uniform_font_groups.setdefault(group, []).append(tb)
    for group_name, members in uniform_font_groups.items():
        sizes = [tb["computed_size"] for tb in members]
        if sizes:
            unified_size = min(sizes)
            for tb in members:
                tb["computed_size"] = unified_size
                # Recalculate height with unified size
                if tb.get("no_wrap"):
                    role_to_height[tb["role"]] = int(unified_size * active_line_spacing * 12700)
                else:
                    role_to_height[tb["role"]] = _measure_required_height(tb["text"], unified_size, tb["width_emu"], active_line_spacing)

    # 5c. Compute image heights preserving aspect ratio
    for img in image_shapes:
        h = _compute_image_height(img["path"], img["width_emu"], img["max_height_emu"], img.get("aspect_ratio"))
        if img["min_height_emu"] > 0:
            h = max(h, img["min_height_emu"])
        role_to_height[img["role"]] = min(h, img["max_height_emu"])

    if overflow_warnings:
        print("\n".join(overflow_warnings))

    # 5b. Uniform height groups (bottom alignment)
    # Collect groups from original shapes_defs
    uniform_groups = {}
    for s in shapes_defs:
        group = s.get("uniform_group")
        role = s.get("role")
        if group and role and role in role_to_height:
            uniform_groups.setdefault(group, []).append(role)
    for group_name, roles in uniform_groups.items():
        max_h = max(role_to_height.get(r, 0) for r in roles)
        if max_h > 0:
            # Add 5% safety margin to prevent clipping
            max_h = int(max_h * 1.05)
            for r in roles:
                role_to_height[r] = max_h

    # 5c. Dynamic width adjustment for flex groups
    # Collect flex groups from original shapes_defs
    flex_groups = {}
    for s in shapes_defs:
        flex = s.get("flex_width")
        role = s.get("role")
        if flex and role:
            flex_groups.setdefault(flex, []).append(role)

    # Compute content lengths for width redistribution
    role_to_content_len = {}
    for tb in textboxes:
        role_to_content_len[tb["role"]] = len(tb["text"])

    # Apply width adjustments
    for group_name, roles in flex_groups.items():
        # Get original regions
        group_regions = []
        for r in roles:
            s = role_to_shape.get(r)
            if s:
                rs = s.get("region", "full")
                region = resolve_region_fn(rs, slide_width, slide_height)
                group_regions.append((r, region, role_to_content_len.get(r, 1)))

        if len(group_regions) < 2:
            continue

        # Calculate total content and total original width
        total_content = sum(c for _, _, c in group_regions)
        total_width = sum(r["width"] for _, r, _ in group_regions)
        avg_content = total_content / len(group_regions)

        # Redistribute widths proportionally (±30% cap to avoid extreme distortion)
        new_widths = {}
        for role, region, content in group_regions:
            ratio = content / avg_content if avg_content > 0 else 1.0
            ratio = max(0.7, min(1.3, ratio))  # cap between 70% and 130%
            new_widths[role] = int(region["width"] * ratio)

        # Normalize to keep total width constant
        scale = total_width / sum(new_widths.values()) if sum(new_widths.values()) > 0 else 1.0
        for role in new_widths:
            new_widths[role] = int(new_widths[role] * scale)

        # Update role_to_width for later region computation
        for role, _, _ in group_regions:
            if role in new_widths:
                # Store in a temporary dict that will be used during region computation
                pass  # We'll handle this in the region computation loop

    # 6. Compute regions following attach chains
    computed = {}  # role -> (region, style)
    shape_registry = {}  # role -> computed region dict

    # Pre-compute flex width adjustments
    role_to_width = {}
    for group_name, roles in flex_groups.items():
        group_regions = []
        for r in roles:
            s_def = role_to_shape.get(r)
            if s_def:
                rs = s_def.get("region", "full")
                region = resolve_region_fn(rs, slide_width, slide_height)
                group_regions.append((r, region, role_to_content_len.get(r, 1)))
        if len(group_regions) < 2:
            continue
        total_content = sum(c for _, _, c in group_regions)
        total_width = sum(r["width"] for _, r, _ in group_regions)
        avg_content = total_content / len(group_regions)
        new_widths = {}
        for role, region, content in group_regions:
            ratio = content / avg_content if avg_content > 0 else 1.0
            ratio = max(0.7, min(1.3, ratio))
            new_widths[role] = int(region["width"] * ratio)
        scale = total_width / sum(new_widths.values()) if sum(new_widths.values()) > 0 else 1.0
        for role in new_widths:
            role_to_width[role] = int(new_widths[role] * scale)

    for s in sorted_shapes:
        role = s.get("role", "")
        region_spec = s.get("region", "full")
        base_region = resolve_region_fn(region_spec, slide_width, slide_height)
        attach = s.get("attach")

        if attach and attach.get("to") in shape_registry:
            parent_region = shape_registry[attach["to"]]
            edge = attach.get("edge", "bottom")
            offset = 0
            # Parse offset from attach or default
            if "offset" in attach:
                offset_val = attach["offset"]
                if isinstance(offset_val, str) and offset_val.endswith("%"):
                    offset = int(slide_height * float(offset_val[:-1]) / 100.0)
                else:
                    offset = int(offset_val)

            region = dict(base_region)
            if edge == "bottom":
                region["top"] = parent_region["top"] + parent_region["height"] + offset
            elif edge == "top":
                region["top"] = parent_region["top"] - region["height"] - offset
            elif edge == "right":
                region["left"] = parent_region["left"] + parent_region["width"] + offset
            elif edge == "left":
                region["left"] = parent_region["left"] - region["width"] - offset
        else:
            region = dict(base_region)

        # Apply flex width adjustment
        if role in role_to_width:
            region["width"] = role_to_width[role]

        # Override height for textboxes with computed height
        if role in role_to_height:
            region["height"] = role_to_height[role]

        style = resolve_style_fn(s.get("style", "body"), theme)
        # For textboxes in constraint mode, set computed font size
        if s.get("type", "textbox") == "textbox" and role in role_to_height:
            style = copy.deepcopy(style)
            # Use pre-computed expanded size if available
            computed_size = None
            for tb in textboxes:
                if tb["role"] == role and "computed_size" in tb:
                    computed_size = tb["computed_size"]
                    break
            if computed_size is None:
                computed_size = best_base * style.get("font_scale", 1.0)
            style["font_size"] = max(int(computed_size), 8)
            style.pop("font_scale", None)

        computed[role] = (region, style)
        shape_registry[role] = region
    # Return in original order + truncated content map
    result = []
    for s in shapes_defs:
        role = s.get("role", "")
        region, style = computed.get(role, (resolve_region_fn(s.get("region", "full"), slide_width, slide_height),
                     resolve_style_fn(s.get("style", "body"), theme)))
        result.append((s, region, style))
    return result, role_to_truncated_text
