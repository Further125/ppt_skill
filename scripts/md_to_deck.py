#!/usr/bin/env python3
"""
Convert Markdown to JSON deck spec.

Usage:
    python scripts/md_to_deck.py input.md --output deck.json
    python scripts/md_to_deck.py input.md --output deck.json --theme dark
"""

import argparse
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def parse_markdown(md_text):
    """Parse markdown text into slide structures."""
    lines = md_text.splitlines()
    slides = []
    pending = None  # dict: {title, layout, content[], table, chart, tree, team, highlight, wordcloud, quote}
    in_code_block = False
    code_buffer = []
    code_lang = None

    def flush_pending():
        nonlocal pending
        if pending is None:
            return
        slide = {"layout": pending.get("layout", "title_content"), "title": pending.get("title", "")}
        if pending.get("content"):
            slide["content"] = pending["content"]
        if pending.get("table_data"):
            slide["table_data"] = pending["table_data"]
        if pending.get("chart_data"):
            slide["chart_data"] = pending["chart_data"]
        if pending.get("tree_data"):
            slide["tree_data"] = pending["tree_data"]
        if pending.get("team_data"):
            slide["team_data"] = pending["team_data"]
        if pending.get("big_number"):
            slide["big_number"] = pending["big_number"]
        if pending.get("label"):
            slide["label"] = pending["label"]
        if pending.get("words"):
            slide["words"] = pending["words"]
        if pending.get("quote"):
            slide["quote"] = pending["quote"]
        if pending.get("author"):
            slide["author"] = pending["author"]

        # Only add non-empty slides
        has_content = any(k in slide for k in ["content", "table_data", "chart_data", "tree_data", "team_data", "big_number", "words", "quote"])
        if slide.get("title") or has_content:
            slides.append(slide)
        pending = None

    def init_pending(title, layout="title_content"):
        nonlocal pending
        flush_pending()
        pending = {"title": title, "layout": layout, "content": []}

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Code blocks
        if stripped.startswith("```"):
            if not in_code_block:
                in_code_block = True
                code_lang = stripped[3:].strip().lower()
                code_buffer = []
            else:
                in_code_block = False
                code_text = "\n".join(code_buffer)

                def reuse_pending(layout, title_fallback=""):
                    """Reuse pending slide if it has no content yet, otherwise flush and create new."""
                    nonlocal pending
                    if pending and not pending.get("content") and not pending.get("table_data") and not pending.get("chart_data") and not pending.get("tree_data") and not pending.get("team_data") and not pending.get("big_number") and not pending.get("words"):
                        pending["layout"] = layout
                        return pending
                    init_pending(pending["title"] if pending else title_fallback, layout)
                    return pending

                if code_lang == "chart":
                    try:
                        chart_data = json.loads(code_text)
                        ctype = chart_data.get("type", "")
                        if ctype == "tree":
                            reuse_pending("tree", "Tree")["tree_data"] = chart_data
                        elif ctype == "highlight":
                            p = reuse_pending("data_highlight")
                            p["big_number"] = chart_data.get("big_number", "")
                            p["label"] = chart_data.get("label", "")
                        elif ctype == "wordcloud":
                            p = reuse_pending("wordcloud")
                            p["words"] = chart_data.get("words", [])
                        elif ctype == "team":
                            reuse_pending("team")["team_data"] = chart_data
                        else:
                            reuse_pending("chart")["chart_data"] = chart_data
                    except json.JSONDecodeError:
                        pass

                elif code_lang == "tree":
                    try:
                        tree_data = json.loads(code_text)
                        reuse_pending("tree")["tree_data"] = tree_data
                    except json.JSONDecodeError:
                        pass

                elif code_lang == "team":
                    try:
                        team_data = json.loads(code_text)
                        reuse_pending("team")["team_data"] = team_data
                    except json.JSONDecodeError:
                        pass

                elif code_lang == "highlight":
                    try:
                        h = json.loads(code_text)
                        p = reuse_pending("data_highlight")
                        p["big_number"] = h.get("big_number", "")
                        p["label"] = h.get("label", "")
                    except json.JSONDecodeError:
                        pass

                elif code_lang == "wordcloud":
                    try:
                        wc = json.loads(code_text)
                        p = reuse_pending("wordcloud")
                        p["words"] = wc.get("words", [])
                    except json.JSONDecodeError:
                        pass

                else:
                    # Unrecognized code block - preserve as formatted text
                    if pending is None:
                        init_pending("")
                    # Add language label, spacer row, and code content
                    pending["content"].append(f"[{code_lang}]")
                    pending["content"].append("")  # spacer row for label bar
                    for code_line in code_text.splitlines():
                        pending["content"].append(code_line)
                    pending["content"].append("[/code]")

                code_lang = None
            i += 1
            continue

        if in_code_block:
            code_buffer.append(line)
            i += 1
            continue

        # Horizontal rule
        if re.match(r'^---+\s*$', stripped):
            i += 1
            continue

        # H1 -> Cover
        if stripped.startswith("# ") and not stripped.startswith("## "):
            init_pending(stripped[2:].strip(), "cover")
            i += 1
            continue

        # H2 -> Title slide
        if stripped.startswith("## "):
            init_pending(stripped[3:].strip(), "title_content")
            i += 1
            continue

        # H3 -> Subheading, add as bold content
        if stripped.startswith("### "):
            sub = stripped[4:].strip()
            if pending is None:
                init_pending("")
            pending["content"].append(f"**{sub}**")
            i += 1
            continue

        def pending_is_empty():
            """Check if pending slide is just a heading with no real content."""
            if not pending:
                return False
            return not pending.get("content") and not pending.get("table_data") and not pending.get("chart_data") and not pending.get("tree_data") and not pending.get("team_data") and not pending.get("big_number") and not pending.get("words") and not pending.get("quote")

        # Blockquote -> Quote slide
        if stripped.startswith("> "):
            quote_text = stripped[2:].strip()
            match = re.match(r'^(.*?)\s*[-–—]{2}\s*(.+)$', quote_text)
            # If pending is just a heading, transform it into a quote slide
            if pending_is_empty():
                pending["layout"] = "quote"
                pending["quote"] = match.group(1).strip() if match else quote_text
                if match:
                    pending["author"] = match.group(2).strip()
            else:
                quote_title = pending["title"] if pending else ""
                flush_pending()
                quote_slide = {"layout": "quote", "quote": match.group(1).strip() if match else quote_text}
                if match:
                    quote_slide["author"] = match.group(2).strip()
                if quote_title:
                    quote_slide["title"] = quote_title
                slides.append(quote_slide)
            i += 1
            continue

        # Table
        if stripped.startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1

            rows = []
            for tl in table_lines:
                if re.match(r'^\|[-:\s|]+\|$', tl):
                    continue
                cells = [c.strip() for c in tl.split("|")]
                cells = [c for c in cells if c]
                if cells:
                    rows.append(cells)

            if rows:
                # If pending is just a heading, transform it into a table slide
                if pending_is_empty():
                    pending["layout"] = "table"
                    pending["table_data"] = {"headers": rows[0], "rows": rows[1:]}
                else:
                    table_title = pending["title"] if pending else ""
                    flush_pending()
                    slides.append({
                        "layout": "table",
                        "title": table_title or rows[0][0],
                        "table_data": {"headers": rows[0], "rows": rows[1:]}
                    })
            continue

        # Bullet list — preserve indentation for nested lists
        bullet_match = re.match(r'^(\s*)([*\-+])\s+(.*)$', line)
        if bullet_match and not stripped.startswith("|---"):
            indent_spaces = len(bullet_match.group(1))
            bullet = bullet_match.group(3).strip()
            if bullet:
                if pending is None:
                    init_pending("")
                # 2 spaces = 1 indent level
                level = indent_spaces // 2
                if level > 0:
                    pending["content"].append(f"indent:{level}|{bullet}")
                else:
                    pending["content"].append(bullet)
            i += 1
            continue

        # Numbered list
        match = re.match(r'^\d+\.\s+(.*)$', stripped)
        if match:
            item = match.group(1).strip()
            if pending is None:
                init_pending("")
            pending["content"].append(item)
            i += 1
            continue

        # Plain text
        if stripped:
            if pending is None:
                init_pending("")
            pending["content"].append(stripped)

        i += 1

    # Flush any remaining pending content
    flush_pending()

    # Remove truly empty slides
    slides = [s for s in slides if s.get("title") or any(k in s for k in ["content", "table_data", "chart_data", "tree_data", "team_data", "big_number", "words", "quote"])]

    # Add closing slide if not present
    if slides and slides[-1].get("layout") not in ("closing", "cover"):
        slides.append({"layout": "closing", "title": "谢谢", "subtitle": "Q & A"})

    return slides


def build_deck(slides, title="", theme=None):
    deck = {"title": title, "slides": slides}
    if theme:
        deck["theme"] = theme
    return deck


def main():
    parser = argparse.ArgumentParser(description="Convert Markdown to JSON deck")
    parser.add_argument("md", help="Input Markdown file")
    parser.add_argument("--output", "-o", default="-", help="Output JSON file")
    parser.add_argument("--title", help="Deck title (default: filename)")
    parser.add_argument("--theme", help="Theme name")
    args = parser.parse_args()

    with open(args.md, "r", encoding="utf-8") as f:
        md_text = f.read()

    slides = parse_markdown(md_text)
    title = args.title or os.path.splitext(os.path.basename(args.md))[0]
    deck = build_deck(slides, title=title, theme=args.theme)

    json_str = json.dumps(deck, ensure_ascii=False, indent=2)

    if args.output == "-":
        print(json_str)
    else:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(json_str)
        print(f"Converted {len(slides)} slides from {args.md}")
        print(f"Output written to {args.output}")


if __name__ == "__main__":
    main()
