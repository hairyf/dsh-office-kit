#!/usr/bin/env python3
"""Replace text in a Word document without touching the original.

Only changed spans are rewritten, preserving the original text nodes and run
formatting even when a match straddles runs. Changed text uses the formatting
at the start of its span. Body, tables, and every section header and footer
are covered.

Usage:
  replace_text.py --input IN.docx [--out OUT.docx] --find OLD --replace NEW [--find ...]
  replace_text.py --input IN.docx [--out OUT.docx] --spec edits.json

The spec file is {"replace": [{"find": "18%", "replace": "21%"}]} or a bare list
of those objects.

Exit codes: 0 replaced, 1 nothing replaced (with --require-all) or invalid
input, 2 argparse usage error.
"""
from __future__ import annotations

import argparse
import json
import sys
from difflib import SequenceMatcher
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import docx_kit as kit  # noqa: E402

# Windows consoles default to a legacy code page; every payload is UTF-8 JSON.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass


def paragraph_text_nodes(paragraph) -> list:
    """Every w:t below a paragraph, including runs inside hyperlinks."""
    return paragraph._p.xpath(".//w:t")


def replace_in_paragraph(paragraph, find: str, replacement: str) -> int:
    if not find:
        return 0
    ranges = []
    joined = ""
    for node in paragraph_text_nodes(paragraph):
        text = node.text or ""
        if text:
            ranges.append((node, len(joined), len(joined) + len(text)))
            joined += text
    matches = []
    offset = joined.find(find)
    while offset != -1:
        matches.append(offset)
        offset = joined.find(find, offset + len(find))
    if not matches:
        return 0

    changes = [opcode for opcode in SequenceMatcher(
        None, find, replacement, autojunk=False
    ).get_opcodes() if opcode[0] != "equal"]
    # All offsets refer to the original text. Working backwards also prevents
    # inserted text from being matched again during this replacement pair.
    for offset in reversed(matches):
        for _, old_start, old_end, new_start, new_end in reversed(changes):
            start, end = offset + old_start, offset + old_end
            first = next((i for i, (_, _, stop) in enumerate(ranges)
                          if stop > start), len(ranges) - 1)
            for i in range(first, len(ranges)):
                node, node_start, node_end = ranges[i]
                if i != first and node_start >= end:
                    break
                text = node.text or ""
                left = max(0, start - node_start)
                right = min(node_end, end) - node_start
                node.text = (text[:left]
                             + (replacement[new_start:new_end] if i == first else "")
                             + text[right:])
                if node.text and (node.text[0].isspace() or node.text[-1].isspace()):
                    node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return len(matches)


def iter_paragraphs(document):
    def walk_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    yield from cell.paragraphs
                    yield from walk_tables(cell.tables)

    yield from document.paragraphs
    yield from walk_tables(document.tables)
    for section in document.sections:
        for part in (section.header, section.footer, section.first_page_header,
                     section.first_page_footer, section.even_page_header,
                     section.even_page_footer):
            try:
                yield from part.paragraphs
            except Exception:  # a header/footer that is not linked or not present
                continue


def fail(payload: dict, code: int = 1) -> int:
    kit.emit({"ok": False, **payload})
    return code


def load_spec(path: str) -> list[tuple[str, str]]:
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    items = payload.get("replace", []) if isinstance(payload, dict) else payload
    pairs: list[tuple[str, str]] = []
    for item in items:
        find = str(item.get("find", ""))
        if find:
            pairs.append((find, str(item.get("replace", ""))))
    return pairs


def main() -> int:
    parser = argparse.ArgumentParser(description="Replace text in a .docx copy.")
    parser.add_argument("--input", required=True, help="source .docx (never modified)")
    parser.add_argument("--out", help="output path; default <name>-edited.docx")
    parser.add_argument("--find", action="append", default=[], help="text to find; pair with --replace")
    parser.add_argument("--replace", action="append", default=[], help="replacement for the matching --find")
    parser.add_argument("--spec", help="JSON file with replace pairs")
    parser.add_argument("--require-all", action="store_true", help="exit 1 when a pair matches nothing")
    args = parser.parse_args()

    if args.find and len(args.find) != len(args.replace):
        return fail({"error": "--find and --replace must be given in pairs"}, 2)
    pairs = list(zip(args.find, args.replace))
    if args.spec:
        try:
            pairs.extend(load_spec(args.spec))
        except (OSError, ValueError) as exc:
            return fail({"error": f"cannot read spec: {exc}"}, 2)
    if not pairs:
        return fail({"error": "no replacements requested"}, 2)

    source = Path(args.input).expanduser()
    if not source.is_file():
        return fail({"error": f"input not found: {args.input}"}, 1)
    out = Path(args.out).expanduser() if args.out else source.with_name(f"{source.stem}-edited.docx")
    if out.resolve() == source.resolve():
        return fail({"error": "refusing to overwrite the input; pass a different --out"}, 1)

    from docx import Document

    document = Document(str(source))
    per_pair: list[dict] = []
    total = 0
    for find, replacement in pairs:
        hits = 0
        for paragraph in iter_paragraphs(document):
            hits += replace_in_paragraph(paragraph, find, replacement)
        total += hits
        per_pair.append({"find": find, "replace": replacement, "count": hits})

    out.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(out))
    misses = [item["find"] for item in per_pair if not item["count"]]
    payload = {
        "ok": not (args.require_all and misses),
        "input": str(source),
        "output": str(out),
        "replacements": total,
        "pairs": per_pair,
        "misses": misses,
    }
    if args.require_all and misses:
        payload["error"] = "patterns matched nothing: " + ", ".join(misses)
        return fail(payload, 1)
    kit.emit(payload)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
