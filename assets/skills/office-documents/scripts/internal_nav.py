#!/usr/bin/env python
"""Build deterministic internal navigation: a linked contents list and anchors.

Word's TOC field needs a refresh before its links work.  This script instead
writes static hyperlinks that any renderer follows: bookmarks on every heading,
a contents list at the top (with quick links and back links), and optional
Top/Bottom anchors.  Because the entries are literal text there is nothing to
update before the file is read.  Rerunning on a file this script already
processed strips the previous navigation first, so the output is rebuilt
cleanly instead of stacking a second contents block.

Exit codes: 0 success, 1 no headings, 2 bad input.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
import docx_kit as kit

# Windows consoles default to a legacy code page; every payload is UTF-8 JSON.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

TOP_BOOKMARK = "nav_top"
BOTTOM_BOOKMARK = "nav_bottom"
TOC_BOOKMARK = "nav_toc"


def link_run(text: str, anchor: str):
    link = etree.Element(kit.q("hyperlink"))
    link.set(kit.q("anchor"), anchor)
    link.append(kit.make_run(text))
    return link


def add_bookmark(paragraph, name: str, ident: int) -> None:
    start = etree.Element(kit.q("bookmarkStart"))
    start.set(kit.q("id"), str(ident))
    start.set(kit.q("name"), name)
    end = etree.Element(kit.q("bookmarkEnd"))
    end.set(kit.q("id"), str(ident))
    position = 1 if paragraph.find(kit.q("pPr")) is not None else 0
    paragraph.insert(position, start)
    paragraph.append(end)


def paragraph(style: str | None = None):
    node = etree.Element(kit.q("p"))
    if style:
        props = etree.SubElement(node, kit.q("pPr"))
        etree.SubElement(props, kit.q("pStyle")).set(kit.q("val"), style)
    return node


def _remove_bookmark(node) -> None:
    """Drop one bookmarkStart and its matching bookmarkEnd wherever they live."""
    ident = node.get(kit.q("id"))
    parent = node.getparent()
    if parent is not None:
        parent.remove(node)
    if ident is None:
        return
    root = parent if parent is not None else node
    while root.getparent() is not None:
        root = root.getparent()
    for end in list(root.iter(kit.q("bookmarkEnd"))):
        if end.get(kit.q("id")) == ident:
            end.getparent().remove(end)


def strip_previous_nav(body) -> int:
    """Remove navigation inserted by an earlier run, so reruns rebuild cleanly.

    Without this, a second run stacks a second contents block, re-collects the
    first block's title as a chapter, and duplicates the nav_top/nav_toc/
    nav_bottom bookmark names, which makes every back link resolve to the stale
    block. A paragraph is treated as generated when it links to a nav_ anchor or
    carries a structural nav bookmark; a real heading only loses its nav_h
    bookmarks. Returns the number of paragraphs removed.
    """
    removed = 0
    for paragraph_node in list(body.iter(kit.q("p"))):
        bookmarks = [
            node
            for node in paragraph_node.iter(kit.q("bookmarkStart"))
            if (node.get(kit.q("name")) or "").startswith("nav_")
        ]
        links = [
            node
            for node in paragraph_node.iter(kit.q("hyperlink"))
            if (node.get(kit.q("anchor")) or "").startswith("nav_")
        ]
        if not bookmarks and not links:
            continue
        generated = bool(links) or any(
            node.get(kit.q("name")) in (TOP_BOOKMARK, BOTTOM_BOOKMARK, TOC_BOOKMARK)
            for node in bookmarks
        )
        if generated:
            for node in bookmarks:
                _remove_bookmark(node)
            parent = paragraph_node.getparent()
            if parent is not None:
                parent.remove(paragraph_node)
                removed += 1
        else:
            for node in bookmarks:
                _remove_bookmark(node)
    return removed


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Add bookmark anchors and static internal links to headings.")
    parser.add_argument("input_docx", help="Input .docx")
    parser.add_argument("--out", required=True, help="Output .docx")
    parser.add_argument("--toc_title", default="Table of Contents")
    parser.add_argument("--levels", nargs="+", type=int, default=[1, 2, 3], help="Heading levels to include")
    parser.add_argument("--no_quicklinks", action="store_true", help="Skip the quick links bar")
    parser.add_argument("--no_back_to_toc", action="store_true", help="Skip per-heading back links")
    parser.add_argument("--no_top_bottom", action="store_true", help="Skip Top/Bottom anchors")
    args = parser.parse_args(argv)

    try:
        doc = kit.Docx(args.input_docx)
    except (OSError, KeyError) as exc:
        print("error: %s" % exc, file=sys.stderr)
        return 2

    body = doc.body()
    stripped = strip_previous_nav(body)
    headings = []
    ident = doc.next_id()
    for paragraph_node in list(body.iter(kit.q("p"))):
        level = kit.heading_level(paragraph_node)
        text = kit.paragraph_text(paragraph_node).strip()
        if level in args.levels and text:
            name = "nav_h%d" % ident
            add_bookmark(paragraph_node, name, ident)
            headings.append({"level": level, "text": text, "anchor": name})
            ident += 1

    if not headings:
        print("error: document has no headings at levels %s" % args.levels, file=sys.stderr)
        return 1

    first = body.find(kit.q("p"))
    if first is None:
        first = etree.SubElement(body, kit.q("p"))
    position = body.index(first)

    block = []
    if not args.no_top_bottom:
        top = paragraph()
        add_bookmark(top, TOP_BOOKMARK, ident)
        ident += 1
        block.append(top)

    title = paragraph("Heading1")
    add_bookmark(title, TOC_BOOKMARK, ident)
    ident += 1
    title.append(kit.make_run(args.toc_title))
    block.append(title)

    if not args.no_quicklinks:
        quick = paragraph()
        quick.append(kit.make_run("Quick links: "))
        for index, item in enumerate([entry for entry in headings if entry["level"] == 1] or headings):
            if index:
                quick.append(kit.make_run(" | "))
            quick.append(link_run(item["text"], item["anchor"]))
        block.append(quick)

    for item in headings:
        entry = paragraph("TOC%d" % item["level"] if item["level"] <= 3 else "TOC3")
        props = entry.find(kit.q("pPr"))
        ind = etree.SubElement(props, kit.q("ind"))
        ind.set(kit.q("left"), str((item["level"] - 1) * 360))
        entry.append(link_run(item["text"], item["anchor"]))
        block.append(entry)

    for offset, node in enumerate(block):
        body.insert(position + offset, node)

    if not args.no_back_to_toc:
        anchors = {entry["anchor"] for entry in headings}
        for paragraph_node in list(body.iter(kit.q("p"))):
            start = paragraph_node.find(kit.q("bookmarkStart"))
            if start is not None and start.get(kit.q("name")) in anchors:
                back = paragraph("Caption")
                back.append(link_run("Back to contents", TOC_BOOKMARK))
                paragraph_node.addnext(back)

    if not args.no_top_bottom:
        last = paragraph()
        add_bookmark(last, BOTTOM_BOOKMARK, ident)
        ident += 1
        last.append(link_run("Back to top", TOP_BOOKMARK))
        body.append(last)

    doc.save(args.out)
    kit.emit({"ok": True, "output": str(Path(args.out)), "headings": headings, "entries": len(block), "stripped": stripped})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
