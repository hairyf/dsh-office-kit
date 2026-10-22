#!/usr/bin/env python3
"""DSH Office Suite helper CLI.

Creates, reads, edits, recalculates, renders, and verifies DOCX, PDF, PPTX,
and XLSX files. Every command prints one JSON object on stdout. Failures print
one JSON object on stderr and exit non-zero.

The CLI is deliberately a thin, deterministic layer over well-known Python
libraries. It never calls an external model and never fetches remote services.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import uuid
from pathlib import Path
from typing import Any, Iterable

PACKAGE_ROOT = Path(__file__).resolve().parents[1]


def _package_version() -> str:
    try:
        data = json.loads((PACKAGE_ROOT / "package.json").read_text(encoding="utf-8"))
        return str(data["version"])
    except (OSError, ValueError, KeyError, TypeError):
        return "0.0.0"


VERSION = _package_version()

# Windows consoles default to a legacy code page; every payload is UTF-8 JSON.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass


def emit(value: Any, out: str | None = None) -> None:
    text = json.dumps(value, ensure_ascii=False, indent=2, default=str)
    if out:
        Path(out).parent.mkdir(parents=True, exist_ok=True)
        Path(out).write_text(text + "\n", encoding="utf-8")
    print(text)


def fail(message: str, *, code: int = 1, detail: Any = None) -> "NoReturn":
    payload: dict[str, Any] = {"ok": False, "error": message}
    if detail is not None:
        payload["detail"] = detail
    print(json.dumps(payload, ensure_ascii=False, indent=2, default=str), file=sys.stderr)
    raise SystemExit(code)


def load_json(path: str) -> dict[str, Any]:
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except FileNotFoundError:
        fail(f"specification not found: {path}")
    except json.JSONDecodeError as exc:
        fail(f"invalid JSON specification: {exc}")
    if not isinstance(data, dict):
        fail("specification must be a JSON object")
    return data


def require_file(path: str, what: str = "file") -> Path:
    resolved = Path(path).expanduser()
    if not resolved.is_file():
        fail(f"{what} not found: {path}")
    return resolved


def ensure_parent(path: str) -> Path:
    resolved = Path(path).expanduser()
    resolved.parent.mkdir(parents=True, exist_ok=True)
    return resolved


# ---------------------------------------------------------------------------
# LibreOffice kit (the renderer DSH ships)
# ---------------------------------------------------------------------------

#: Where the kit's CLI lives inside any node_modules tree that hosts this plugin.
KIT_RELATIVE = Path("node_modules") / "@deepseek-ai" / "libreoffice-kit" / "lib" / "cli.js"

#: Actionable explanation used whenever the kit cannot be located.
KIT_MISSING = (
    "the LibreOffice renderer is not available: the DSH LibreOffice kit was not found. "
    "The plugin records the path in dsh-office-kit's config.json when the harness loads it; "
    "for a hand-run helper, set DSH_LIBREOFFICE_KIT_CLI to <kit>/lib/cli.js of "
    "@deepseek-ai/libreoffice-kit"
)


def office_home() -> Path:
    """Everything this plugin owns lives under `~/.dsh/office` (mirrors src/python.ts)."""
    return Path.home() / ".dsh" / "office"


def config_file() -> Path:
    """The config.json the host half and postinstall write."""
    override = os.environ.get("DSH_OFFICE_CONFIG")
    return Path(override).expanduser() if override else office_home() / "config.json"


def legacy_config_files() -> list[Path]:
    """Config files older versions wrote in the OS config directory, newest first."""
    if sys.platform == "darwin":
        primary = Path.home() / "Library" / "Application Support" / "dsh-office-kit" / "config.json"
    elif sys.platform == "win32":
        primary = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming") / "dsh-office-kit" / "config.json"
    else:
        primary = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config") / "dsh-office-kit" / "config.json"
    roaming = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming") / "dsh-office-kit" / "config.json"
    return list(dict.fromkeys([primary, roaming]))


def read_config() -> dict[str, Any]:
    """Merged config: the canonical file wins key by key over older locations."""
    merged: dict[str, Any] = {}
    for file in [*reversed(legacy_config_files()), config_file()]:
        try:
            parsed = json.loads(file.read_text("utf-8"))
        except (OSError, ValueError):
            continue
        if isinstance(parsed, dict):
            merged.update(parsed)
    return merged


def saved_kit() -> tuple[str, str] | None:
    """Return (node, kit CLI) as recorded by the host half, or None.

    This is how a profile install works: the kit lives in the harness tree, which
    no walk-up from the plugin directory can reach, so the host half looks it up
    while it loads and stores the result here for every later helper call.
    """
    kit = read_config().get("libreofficeKit")
    if not isinstance(kit, dict):
        return None
    cli = kit.get("cli")
    if not isinstance(cli, str) or not Path(cli).is_file():
        return None
    node = kit.get("node")
    if not isinstance(node, str) or not node:
        node = os.environ.get("DSH_NODE") or shutil.which("node")
    if not node:
        return None
    return node, cli


def find_kit() -> tuple[str, str] | None:
    """Return (node, kit CLI) for the bundled renderer, or None when it is absent.

    Order: an explicit `DSH_LIBREOFFICE_KIT_CLI`, then the path the host half
    recorded in config.json (the profile-install case), then a walk up from this
    file, which covers a checkout or a plugin-local dependency.
    """
    override = os.environ.get("DSH_LIBREOFFICE_KIT_CLI")
    if override:
        cli = Path(override).expanduser()
        if cli.is_file():
            node = os.environ.get("DSH_NODE") or shutil.which("node")
            return (node, str(cli)) if node else None
    recorded = saved_kit()
    if recorded is not None:
        return recorded
    candidates: list[Path] = []
    for root in (PACKAGE_ROOT, *PACKAGE_ROOT.parents):
        candidates.append(root / KIT_RELATIVE)
    cli = next((candidate for candidate in candidates if candidate.is_file()), None)
    if cli is None:
        return None
    node = os.environ.get("DSH_NODE") or shutil.which("node")
    if node is None:
        return None
    return node, str(cli)


def _kit_env() -> dict[str, str]:
    """Environment for a kit run: the recorded node may be the harness's Electron binary."""
    return {**os.environ, "ELECTRON_RUN_AS_NODE": "1"}


def _kit_call(arguments: list[str], timeout: int = 300) -> dict[str, Any]:
    """Run one kit command and return its JSON report, failing with the kit's own words."""
    kit = find_kit()
    if kit is None:
        fail(KIT_MISSING)
    node, cli = kit
    try:
        completed = subprocess.run(
            [node, cli, *arguments],
            capture_output=True,
            text=True,
            env=_kit_env(),
            # The kit always speaks UTF-8; the locale code page would mangle CJK paths.
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        fail(f"the LibreOffice kit timed out after {timeout}s ({arguments[0]})")
    except OSError as exc:
        fail(f"cannot start the LibreOffice kit: {exc}", detail=KIT_MISSING)
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout).strip()[:4000]
        fail(f"the LibreOffice kit could not run '{arguments[0]}'", detail=detail or None)
    try:
        payload = json.loads(completed.stdout)
    except json.JSONDecodeError:
        fail(
            f"the LibreOffice kit printed no JSON report for '{arguments[0]}'",
            detail=(completed.stdout or completed.stderr).strip()[:800] or None,
        )
    if not isinstance(payload, dict):
        fail(f"the LibreOffice kit printed an unexpected report for '{arguments[0]}'")
    return payload


def kit_report() -> dict[str, Any]:
    """Describe the kit without ever failing: status output must survive a missing runtime."""
    kit = find_kit()
    if kit is None:
        return {"available": False, "remedy": KIT_MISSING}
    node, cli = kit
    report: dict[str, Any] = {"available": True, "node": node, "cli": cli, "backend": None, "version": None}
    try:
        completed = subprocess.run(
            [node, cli, "capabilities", "--json"],
            capture_output=True,
            text=True,
            env=_kit_env(),
            encoding="utf-8",
            errors="replace",
            timeout=120,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        report["error"] = f"{type(exc).__name__}: {exc}"
        return report
    if completed.returncode != 0:
        report["error"] = ((completed.stderr or completed.stdout).strip() or f"capabilities exited {completed.returncode}")[:800]
        return report
    try:
        runtime = (json.loads(completed.stdout) or {}).get("runtime") or {}
    except (json.JSONDecodeError, AttributeError):
        report["error"] = "capabilities printed no JSON"
        return report
    report["backend"] = runtime.get("backend")
    report["version"] = runtime.get("version")
    return report


def glyph_warnings(missing: list[str]) -> list[str]:
    """Turn the kit's missing-font list into one warning a caller can act on."""
    if not missing:
        return []
    return [
        f"{len(missing)} font(s) used by this file are not installed: {', '.join(missing)}. "
        "The file rendered, but that text uses a fallback font and may reflow or show boxes; "
        "a zero exit code only means the renderer ran."
    ]


def report_warnings(warnings: list[str]) -> None:
    for warning in warnings:
        print(f"warning: {warning}", file=sys.stderr)


def export_pdf(input_path: Path, out_dir: Path, dpi: int) -> dict[str, Any]:
    """Convert to PDF through the kit. `dpi` is accepted for signature compatibility."""
    del dpi
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf = out_dir / (input_path.stem + ".pdf")
    report = _kit_call(["convert", "--input", str(input_path), "--output", str(pdf)])
    produced = Path(report.get("outputPath") or pdf)
    if not produced.is_file():
        fail(f"the LibreOffice kit reported a PDF export for {input_path.name} but wrote none")
    return {
        "pdf": str(produced),
        "backend": report.get("backend"),
        "missing_fonts": [str(name) for name in report.get("missingFonts") or []],
    }


def xlsx_render(input_path: Path, out_dir: Path, dpi: int) -> dict[str, Any]:
    """Render a workbook through PDF, because the kit's direct render prints ranges only.

    The direct xlsx render path rasterises each used range and drops charts, images
    and print areas; converting first lays the sheet out the way the app prints it.
    """
    converted = export_pdf(input_path, out_dir, dpi)
    pages = rasterize_pdf(converted["pdf"], out_dir, dpi)
    missing = converted["missing_fonts"]
    warnings = glyph_warnings(missing)
    report_warnings(warnings)
    return {
        "ok": True,
        "kind": "xlsx",
        "command": "render",
        "input": str(input_path),
        "renderer": "libreoffice-kit",
        "engine": converted.get("backend"),
        "page_count": len(pages),
        "pages": pages,
        "images": [
            {"page": index, "sheet": None, "range": None, "path": path}
            for index, path in enumerate(pages, start=1)
        ],
        "pdf": converted["pdf"],
        "missing_fonts": missing,
        "glyph_fallback": bool(missing),
        "warnings": warnings,
    }


def docx_used_fonts(path: Path) -> set[str]:
    """Font names the document actually applies.

    The kit reports every font in the package, including latent styles inherited
    from python-docx's template (e.g. Courier in the unused MacroText style), so
    a document with no Courier text still warns about it. Keep only fonts that
    appear on a run, in a referenced style, in the defaults, or in the theme.
    """
    import zipfile

    from lxml import etree

    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    fonts: set[str] = set()
    styles: dict[str, tuple[set[str], str | None]] = {}
    referenced: set[str] = set()
    theme_fonts: set[str] = set()

    def rfonts(element) -> set[str]:
        found = set()
        for node in element.iter(f"{w}rFonts"):
            for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
                value = node.get(f"{w}{attribute}")
                if value:
                    found.add(value)
        return found

    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())

        def load(name: str):
            return etree.fromstring(archive.read(name))

        if "word/styles.xml" in names:
            root = load("word/styles.xml")
            defaults = root.find(f"{w}docDefaults")
            if defaults is not None:
                fonts.update(rfonts(defaults))
            for style in root.iter(f"{w}style"):
                style_id = style.get(f"{w}styleId")
                based_on = style.find(f"{w}basedOn")
                styles[style_id] = (rfonts(style), based_on.get(f"{w}val") if based_on is not None else None)

        for name in names:
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            if name.startswith(("word/styles", "word/theme/")):
                continue
            root = load(name)
            fonts.update(rfonts(root))
            for tag in ("pStyle", "rStyle", "tblStyle"):
                for element in root.iter(f"{w}{tag}"):
                    value = element.get(f"{w}val")
                    if value:
                        referenced.add(value)

        if "word/theme/theme1.xml" in names:
            theme = load("word/theme/theme1.xml")
            for tag in ("majorFont", "minorFont"):
                for element in theme.iter(tag):
                    latin = element.find("latin")
                    if latin is not None and latin.get("typeface"):
                        theme_fonts.add(latin.get("typeface"))

    pending = list(referenced)
    seen = set()
    while pending:
        style_id = pending.pop()
        if style_id in seen:
            continue
        seen.add(style_id)
        entry = styles.get(style_id)
        if entry is None:
            continue
        style_fonts, based_on = entry
        fonts.update(style_fonts)
        if based_on:
            pending.append(based_on)
    return fonts | theme_fonts


def filter_missing_fonts(input_path: Path, kind: str, missing: list[str]) -> list[str]:
    if kind != "docx" or not missing:
        return missing
    used = {name.lower() for name in docx_used_fonts(input_path)}
    return [name for name in missing if name.lower() in used]


def render_office(input_path: Path, kind: str, out_dir: Path, dpi: int, emit_pdf: bool = False) -> dict[str, Any]:
    """Rasterise an Office file page by page through the kit.

    The kit creates its output directory itself and refuses an existing one, so
    render into a staging directory and move the artifacts into the caller's
    directory, which may be reused across renders.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    staging = out_dir / f".kit-render-{os.getpid()}-{uuid.uuid4().hex[:8]}"
    report = _kit_call(["render", "--input", str(input_path), "--output-dir", str(staging), "--dpi", str(dpi)])
    images = []
    for item in report.get("images") or []:
        if not item.get("path"):
            continue
        source = Path(item["path"])
        target = out_dir / source.name
        os.replace(source, target)
        images.append({
            "page": item.get("page"),
            "sheet": item.get("sheet"),
            "range": item.get("range"),
            "width": item.get("width"),
            "height": item.get("height"),
            "bytes": item.get("byteLength"),
            "path": str(target),
        })
    manifest = staging / "manifest.json"
    if manifest.is_file():
        os.replace(manifest, out_dir / manifest.name)
    shutil.rmtree(staging, ignore_errors=True)
    if not images:
        fail(f"the LibreOffice kit rendered no pages for {input_path.name}")
    missing = filter_missing_fonts(input_path, kind, [str(name) for name in report.get("missingFonts") or []])
    warnings = glyph_warnings(missing)
    report_warnings(warnings)
    result: dict[str, Any] = {
        "ok": True,
        "kind": kind,
        "input": str(input_path),
        "renderer": "libreoffice-kit",
        "engine": report.get("backend"),
        "page_count": len(images),
        "pages": [image["path"] for image in images],
        "images": images,
        "missing_fonts": missing,
        "glyph_fallback": bool(missing),
        "warnings": warnings,
    }
    if emit_pdf:
        converted = export_pdf(input_path, out_dir, dpi)
        result["pdf"] = converted["pdf"]
        converted_missing = filter_missing_fonts(input_path, kind, converted["missing_fonts"])
        if not missing and converted_missing:
            result["missing_fonts"] = converted_missing
            result["glyph_fallback"] = True
            result["warnings"] = glyph_warnings(converted_missing)
            report_warnings(result["warnings"])
    return result


def rasterize_pdf(pdf: str, out_dir: Path, dpi: int = 110) -> list[str]:
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf  # type: ignore[no-redef]
        except ImportError:
            fail("PyMuPDF is not installed; run the plugin postinstall again")
    out_dir.mkdir(parents=True, exist_ok=True)
    pages: list[str] = []
    document = pymupdf.open(pdf)
    try:
        for index, page in enumerate(document, start=1):
            zoom = dpi / 72.0
            pixmap = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
            target = out_dir / f"page-{index}.png"
            pixmap.save(str(target))
            pages.append(str(target))
    finally:
        document.close()
    return pages


# ---------------------------------------------------------------------------
# DOCX
# ---------------------------------------------------------------------------


def _docx_align(name: str | None):
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    return {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }.get((name or "").lower(), WD_ALIGN_PARAGRAPH.LEFT)


def _docx_add_equation(paragraph, latex: str) -> bool:
    try:
        from latex2mathml.converter import convert as latex_to_mathml
        from lxml import etree
        from mathml2omml import convert as mathml_to_omml

        omml = mathml_to_omml(latex_to_mathml(latex))
        if "xmlns:m=" not in omml:
            omml = omml.replace(
                "<m:oMath",
                '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
                1,
            )
        element = etree.fromstring(omml.encode("utf-8"))
        paragraph._p.append(element)
        return True
    except Exception:
        return False


def _docx_mark_header_row(row) -> None:
    """Repeat this row across pages and mark it as a header for assistive tech."""
    from docx.oxml.ns import qn

    properties = row._tr.get_or_add_trPr()
    if properties.find(qn("w:tblHeader")) is None:
        header = properties.makeelement(qn("w:tblHeader"), {})
        properties.append(header)


def _docx_add_block(document, block: dict[str, Any], counters: dict[str, int]) -> None:
    from docx.shared import Inches, Pt

    block_type = str(block.get("type", "paragraph")).lower()
    if block_type == "heading":
        level = int(block.get("level", 1))
        document.add_heading(str(block.get("text", "")), level=max(0, min(level, 9)))
        counters["headings"] += 1
    elif block_type == "paragraph":
        paragraph = document.add_paragraph()
        run = paragraph.add_run(str(block.get("text", "")))
        run.bold = bool(block.get("bold", False))
        run.italic = bool(block.get("italic", False))
        if "size" in block:
            run.font.size = Pt(float(block["size"]))
        if block.get("color"):
            from docx.shared import RGBColor

            raw = str(block["color"]).lstrip("#")
            if len(raw) == 6:
                run.font.color.rgb = RGBColor.from_string(raw.upper())
        paragraph.alignment = _docx_align(block.get("align"))
        if block.get("style"):
            paragraph.style = str(block["style"])
        counters["paragraphs"] += 1
    elif block_type in {"bullets", "numbered"}:
        style = "List Bullet" if block_type == "bullets" else "List Number"
        for item in block.get("items", []):
            paragraph = document.add_paragraph(str(item), style=style)
            if block.get("size"):
                for run in paragraph.runs:
                    run.font.size = Pt(float(block["size"]))
            counters["list_items"] += 1
    elif block_type == "table":
        headers = [str(value) for value in block.get("headers", [])]
        rows = block.get("rows", [])
        columns = max([len(headers), *[len(row) for row in rows]]) if (headers or rows) else 0
        if columns == 0:
            return
        table = document.add_table(rows=0, cols=columns)
        style = str(block.get("style", "Table Grid"))
        try:
            table.style = style
        except Exception:
            table.style = "Table Grid"
        if headers:
            cells = table.add_row().cells
            for index, value in enumerate(headers):
                cells[index].text = value
                for paragraph in cells[index].paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
            _docx_mark_header_row(table.rows[0])
        for row in rows:
            cells = table.add_row().cells
            for index, value in enumerate(row):
                if index < len(cells):
                    cells[index].text = "" if value is None else str(value)
        counters["tables"] += 1
    elif block_type == "image":
        path = require_file(str(block.get("path", "")), "image")
        width = block.get("width")
        if width is None:
            document.add_picture(str(path))
        else:
            document.add_picture(str(path), width=Inches(float(width)))
        counters["images"] += 1
    elif block_type == "equation":
        paragraph = document.add_paragraph()
        if not _docx_add_equation(paragraph, str(block.get("latex", ""))):
            paragraph.add_run(str(block.get("latex", ""))).italic = True
            counters["equation_fallbacks"] += 1
        else:
            counters["equations"] += 1
    elif block_type == "quote":
        document.add_paragraph(str(block.get("text", "")), style="Intense Quote")
        counters["paragraphs"] += 1
    elif block_type == "page_break":
        from docx.enum.text import WD_BREAK

        document.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    elif block_type == "spacer":
        document.add_paragraph("")
    else:
        document.add_paragraph(str(block.get("text", "")))


def _docx_add_blocks(document, blocks: Iterable[dict[str, Any]], counters: dict[str, int]) -> None:
    for block in blocks:
        if isinstance(block, dict):
            _docx_add_block(document, block, counters)


def docx_create(spec: dict[str, Any], out: Path) -> dict[str, Any]:
    from docx import Document
    from docx.shared import Pt

    document = Document()
    style = document.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    section = document.sections[0]
    page_size = str(spec.get("page_size", "Letter")).lower()
    if page_size == "a4":
        section.page_width = int(8.27 * 914400)
        section.page_height = int(11.69 * 914400)
    if str(spec.get("orientation", "portrait")).lower() == "landscape":
        width, height = section.page_width, section.page_height
        section.page_width, section.page_height = height, width
    for side in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        if side in spec:
            setattr(section, side, int(float(spec[side]) * 914400))

    if spec.get("title"):
        document.add_heading(str(spec["title"]), 0)
    if spec.get("subtitle"):
        paragraph = document.add_paragraph(str(spec["subtitle"]))
        paragraph.alignment = _docx_align("center")
    if spec.get("author") or spec.get("title"):
        document.core_properties.title = str(spec.get("title", ""))
        document.core_properties.author = str(spec.get("author", ""))
        # A fresh python-docx document carries the library name in comments; a
        # delivered file should not advertise the tool that wrote it.
        document.core_properties.comments = str(spec.get("comments", ""))
        document.core_properties.last_modified_by = str(spec.get("author", ""))
    if spec.get("header"):
        document.sections[0].header.paragraphs[0].text = str(spec["header"])
    if spec.get("footer"):
        document.sections[0].footer.paragraphs[0].text = str(spec["footer"])

    counters = {key: 0 for key in ("headings", "paragraphs", "list_items", "tables", "images", "equations", "equation_fallbacks")}
    _docx_add_blocks(document, spec.get("blocks", spec.get("sections", [])), counters)
    ensure_parent(str(out))
    document.save(str(out))
    return {"ok": True, "kind": "docx", "command": "create", "output": str(out), "counts": counters}


def _docx_iter_paragraphs(document):
    for paragraph in document.paragraphs:
        yield paragraph
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph
    for section in document.sections:
        for part in (section.header, section.footer):
            for paragraph in part.paragraphs:
                yield paragraph


def _replace_in_paragraph(paragraph, find: str, replacement: str) -> int:
    count = 0
    for run in paragraph.runs:
        if find in run.text:
            run.text = run.text.replace(find, replacement)
            count += 1
    if count == 0 and find in paragraph.text:
        text = paragraph.text.replace(find, replacement)
        for run in list(paragraph.runs):
            run._element.getparent().remove(run._element)
        paragraph.add_run(text)
        count = 1
    return count


def _docx_has_equation(paragraph) -> bool:
    tag = "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"
    return bool(paragraph._p.findall(f".//{tag}"))


def docx_read(input_path: Path) -> dict[str, Any]:
    from docx import Document

    document = Document(str(input_path))
    paragraphs = []
    headings = []
    equations = 0
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text and not _docx_has_equation(paragraph):
            continue
        item = {"text": text, "style": paragraph.style.name if paragraph.style is not None else None}
        if _docx_has_equation(paragraph):
            equations += 1
            item["equation"] = True
        if paragraph.style is not None and paragraph.style.name.lower().startswith("heading"):
            headings.append(text)
        if text:
            paragraphs.append(item)
    tables = []
    for table in document.tables:
        tables.append({"rows": [[cell.text for cell in row.cells] for row in table.rows]})
    section = document.sections[0]
    return {
        "ok": True,
        "kind": "docx",
        "command": "read",
        "input": str(input_path),
        "title": document.core_properties.title or None,
        "author": document.core_properties.author or None,
        "paragraphs": paragraphs,
        "headings": headings,
        "tables": tables,
        "images": len(document.inline_shapes),
        "equations": equations,
        "counts": {
            "paragraphs": len(paragraphs),
            "headings": len(headings),
            "tables": len(tables),
            "images": len(document.inline_shapes),
        },
        "page": {
            "width_inches": round(section.page_width.inches, 2),
            "height_inches": round(section.page_height.inches, 2),
        },
    }


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------


#: TrueType CJK candidates. ReportLab's TTFont only handles glyf outlines, so a
#: CFF/PostScript TTC such as NotoSansCJK raises on registration; it is listed
#: after the TrueType fonts and its failure is not treated as success.
_PDF_CJK_CANDIDATES = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\msyh.ttf",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\simsun.ttc",
    "/System/Library/Fonts/STHeiti Light.ttc",
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/System/Library/Fonts/Supplemental/Songti.ttc",
    "/usr/share/fonts/truetype/winfonts/msyh.ttc",
    "/usr/share/fonts/truetype/winfonts/msyhl.ttc",
    "/usr/share/fonts/truetype/winfonts/simhei.ttf",
    "/usr/share/fonts/truetype/winfonts/NotoSansSC-VF.ttf",
    "/usr/share/fonts/truetype/winfonts/Deng.ttf",
    "/usr/share/fonts/fonts-gb/GB_ST_GB18030.ttf",
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/truetype/arphic/uming.ttc",
    "/usr/share/fonts/truetype/arphic/ukai.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
]

#: Probes used to decide whether a registered face can actually draw CJK text.
_CJK_PROBE_CODEPOINTS = (0x4E2D, 0x6587, 0x9A8C, 0x8BC1)  # 中文验证

#: Fonts that register successfully but carry no CJK glyphs. Using one for CJK
#: text produces blank boxes and a text layer of NUL characters, so it is only
#: ever a last resort for purely Latin content.
_LATIN_FALLBACK_FONTS = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
]


def _font_draws_cjk(face) -> bool:
    """Whether a registered reportlab face maps any probe codepoint to a glyph."""
    glyphs = getattr(face, "charToGlyph", None)
    if not isinstance(glyphs, dict):
        return False
    return all(glyphs.get(codepoint) for codepoint in _CJK_PROBE_CODEPOINTS)


def _register_font(font_name: str, path: str) -> bool:
    """Register one font file under ``font_name``; TTC files use their first face."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    attempts = [{"subfontIndex": 0}, {}] if path.lower().endswith(".ttc") else [{}]
    for extra in attempts:
        try:
            pdfmetrics.registerFont(TTFont(font_name, path, **extra))
            return True
        except Exception:
            continue
    return False


class PdfFontError(RuntimeError):
    """Raised when the content needs glyphs no available font can draw."""


def _pdf_font(required_text: str = "") -> str:
    """Pick a font that can actually draw ``required_text``.

    Registration success alone is not enough: DejaVuSans registers cleanly and
    still renders every CJK character as an empty box, so each candidate is
    probed for real glyph coverage before it is accepted. When the content needs
    CJK and no font supplies it, this fails instead of emitting NUL text.
    """
    from reportlab.pdfbase import pdfmetrics

    needs_cjk = any(ord(character) > 0x2E7F for character in required_text)
    for candidate in _PDF_CJK_CANDIDATES:
        if not Path(candidate).is_file():
            continue
        if not _register_font("OfficeCJK", candidate):
            continue
        if not needs_cjk or _font_draws_cjk(pdfmetrics.getFont("OfficeCJK").face):
            return "OfficeCJK"
    try:
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont

        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        return "STSong-Light"
    except Exception:
        pass
    if needs_cjk:
        raise PdfFontError(
            "no font with CJK glyphs is available: the content contains CJK text but every "
            "candidate font is missing or has no CJK coverage (ReportLab cannot load CFF/PostScript "
            "outlines such as NotoSansCJK-Regular.ttc). Install a TrueType CJK font, for example "
            "fonts-wqy-zenhei or fonts-noto-cjk, and retry."
        )
    for candidate in _LATIN_FALLBACK_FONTS:
        if Path(candidate).is_file() and _register_font("OfficeLatin", candidate):
            return "OfficeLatin"
    return "Helvetica"


def _pdf_blocks_text(blocks: Iterable[dict[str, Any]]) -> str:
    """Collect the literal text of PDF blocks so font choice can match the content."""
    parts: list[str] = []

    def visit(value: Any) -> None:
        if isinstance(value, dict):
            for key in ("text", "title", "subtitle"):
                if isinstance(value.get(key), str):
                    parts.append(value[key])
            for key in ("headers", "items"):
                for item in value.get(key, []) or []:
                    if isinstance(item, str):
                        parts.append(item)
            for row in value.get("rows", []) or []:
                for cell in row if isinstance(row, (list, tuple)) else [row]:
                    if isinstance(cell, str):
                        parts.append(cell)
            for key in ("blocks", "pages"):
                visit(value.get(key, []))
        elif isinstance(value, (list, tuple)):
            for item in value:
                visit(item)

    visit(list(blocks))
    return "".join(parts)


def _pdf_flowables(blocks: Iterable[dict[str, Any]], font: str, counters: dict[str, int]):
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import Image, PageBreak, Paragraph, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape

    styles = getSampleStyleSheet()
    body = ParagraphStyle("OfficeBody", parent=styles["BodyText"], fontName=font, fontSize=10.5, leading=15)
    heading1 = ParagraphStyle("OfficeH1", parent=styles["Heading1"], fontName=font, fontSize=18, leading=24, spaceBefore=16, spaceAfter=8)
    heading2 = ParagraphStyle("OfficeH2", parent=styles["Heading2"], fontName=font, fontSize=14, leading=20, spaceBefore=12, spaceAfter=6)
    title = ParagraphStyle("OfficeTitle", parent=styles["Title"], fontName=font, fontSize=22, leading=28)
    flowables = []
    for block in blocks:
        if not isinstance(block, dict):
            continue
        block_type = str(block.get("type", "paragraph")).lower()
        if block_type == "title":
            flowables.append(Paragraph(escape(str(block.get("text", ""))), title))
            flowables.append(Spacer(1, 0.2 * inch))
        elif block_type == "heading":
            style = heading1 if int(block.get("level", 1)) <= 1 else heading2
            flowables.append(Paragraph(escape(str(block.get("text", ""))), style))
            counters["headings"] += 1
        elif block_type in {"paragraph", "quote"}:
            flowables.append(Paragraph(escape(str(block.get("text", ""))), body))
            counters["paragraphs"] += 1
        elif block_type in {"bullets", "numbered"}:
            for index, item in enumerate(block.get("items", []), start=1):
                prefix = "• " if block_type == "bullets" else f"{index}. "
                flowables.append(Paragraph(prefix + escape(str(item)), body))
                counters["list_items"] += 1
        elif block_type == "table":
            headers = [str(value) for value in block.get("headers", [])]
            rows = [[str(value) for value in row] for row in block.get("rows", [])]
            data = ([headers] if headers else []) + rows
            if data:
                table = Table(data, repeatRows=1 if headers else 0)
                table.setStyle(TableStyle([
                    ("FONTNAME", (0, 0), (-1, -1), font),
                    ("FONTSIZE", (0, 0), (-1, -1), 9),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEF7")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ]))
                flowables.append(table)
                flowables.append(Spacer(1, 0.15 * inch))
                counters["tables"] += 1
        elif block_type == "image":
            path = require_file(str(block.get("path", "")), "image")
            width = float(block.get("width", 4.5)) * inch
            flowables.append(Image(str(path), width=width, height=width * 0.66))
            counters["images"] += 1
        elif block_type == "page_break":
            flowables.append(PageBreak())
        elif block_type == "spacer":
            flowables.append(Spacer(1, float(block.get("height", 0.2)) * inch))
    return flowables


def pdf_create(spec: dict[str, Any], out: Path) -> dict[str, Any]:
    from reportlab.lib.pagesizes import A4, letter
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate

    size = A4 if str(spec.get("page_size", "Letter")).lower() == "a4" else letter
    ensure_parent(str(out))
    counters = {key: 0 for key in ("headings", "paragraphs", "list_items", "tables", "images")}
    following: list[dict[str, Any]] = []
    for page in spec.get("pages", []):
        if isinstance(page, dict):
            following.extend(page.get("blocks", []))
            if page.get("break", True):
                following.append({"type": "page_break"})
    following.extend(spec.get("blocks", []))
    kinds = {str(block.get("type", "")).lower() for block in following if isinstance(block, dict)}
    blocks: list[dict[str, Any]] = []
    # `title` is metadata; only draw it when the blocks do not already carry one.
    if spec.get("title") and "title" not in kinds:
        blocks.append({"type": "title", "text": spec["title"]})
    if spec.get("subtitle"):
        blocks.append({"type": "paragraph", "text": spec["subtitle"]})
    blocks.extend(following)
    # Choose the font against the text that will actually be drawn, so CJK content
    # can never silently fall back to a Latin-only face.
    font = _pdf_font(_pdf_blocks_text(blocks))
    document = SimpleDocTemplate(
        str(out),
        pagesize=size,
        leftMargin=0.85 * inch,
        rightMargin=0.85 * inch,
        topMargin=0.8 * inch,
        bottomMargin=0.8 * inch,
        title=str(spec.get("title", "")),
        author=str(spec.get("author", "")),
    )
    document.build(_pdf_flowables(blocks, font, counters))
    return {"ok": True, "kind": "pdf", "command": "create", "output": str(out), "font": font, "counts": counters}


def pdf_read(input_path: Path) -> dict[str, Any]:
    from pypdf import PdfReader

    reader = PdfReader(str(input_path))
    pages = []
    total_words = 0
    for index, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        total_words += len(text.split())
        pages.append({"page": index, "text": text, "characters": len(text)})
    metadata = {}
    try:
        metadata = {key: str(value) for key, value in (reader.metadata or {}).items() if value is not None}
    except Exception:
        metadata = {}
    fields = []
    try:
        root = reader.trailer["/Root"]
        if "/AcroForm" in root:
            fields = sorted((root["/AcroForm"].get("/Fields") or []))
            fields = [str(getattr(field, "idnum", None) or field) for field in fields]
    except Exception:
        fields = []
    return {
        "ok": True,
        "kind": "pdf",
        "command": "read",
        "input": str(input_path),
        "page_count": len(reader.pages),
        "word_count": total_words,
        "metadata": metadata,
        "pages": pages,
        "form_fields": fields,
        "ocr": {"available": False, "reason": "OCR is not bundled; use a dedicated OCR tool."},
    }


def pdf_render(input_path: Path, out_dir: Path, dpi: int) -> dict[str, Any]:
    pages = rasterize_pdf(str(input_path), out_dir, dpi)
    return {"ok": True, "kind": "pdf", "command": "render", "input": str(input_path), "page_count": len(pages), "pages": pages}


# ---------------------------------------------------------------------------
# PPTX
# ---------------------------------------------------------------------------


def _pptx_layout(prs, name: str):
    mapping = {
        "title": 0,
        "bullets": 1,
        "section": 2,
        "title_only": 5,
        "blank": 6,
    }
    index = mapping.get(name, 5)
    return prs.slide_layouts[min(index, len(prs.slide_layouts) - 1)]


def _pptx_add_chart(slide, spec: dict[str, Any]):
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches

    chart_type = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
    }.get(str(spec.get("type", "bar")).lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)
    data = CategoryChartData()
    data.categories = [str(value) for value in spec.get("categories", [])]
    for series in spec.get("series", []):
        data.add_series(str(series.get("name", "Series")), [float(value) for value in series.get("values", [])])
    frame = slide.shapes.add_chart(
        chart_type,
        Inches(float(spec.get("left", 0.8))),
        Inches(float(spec.get("top", 1.6))),
        Inches(float(spec.get("width", 8.4))),
        Inches(float(spec.get("height", 4.6))),
        data,
    )
    if spec.get("title"):
        frame.chart.has_title = True
        frame.chart.chart_title.text_frame.text = str(spec["title"])
    return frame.chart


def _pptx_add_slide(prs, spec: dict[str, Any], counters: dict[str, int]):
    from pptx.util import Inches, Pt

    kind = str(spec.get("layout", "bullets")).lower()
    slide = prs.slides.add_slide(_pptx_layout(prs, kind))
    title_shape = slide.shapes.title
    if title_shape is not None and spec.get("title") is not None:
        title_shape.text = str(spec["title"])
    elif spec.get("title") is not None:
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(12.1), Inches(0.9))
        title_frame = title_box.text_frame
        title_frame.word_wrap = True
        title_frame.text = str(spec["title"])
        title_frame.paragraphs[0].font.size = Pt(32)
        title_frame.paragraphs[0].font.bold = True
    if kind == "title":
        if len(slide.placeholders) > 1 and spec.get("subtitle") is not None:
            slide.placeholders[1].text = str(spec["subtitle"])
    elif kind == "bullets":
        body = None
        for placeholder in slide.placeholders:
            if placeholder.placeholder_format.idx == 1:
                body = placeholder.text_frame
                break
        if body is not None:
            body.clear()
            items = spec.get("bullets", [])
            for index, item in enumerate(items):
                text = str(item.get("text", "")) if isinstance(item, dict) else str(item)
                paragraph = body.paragraphs[0] if index == 0 else body.add_paragraph()
                paragraph.text = text
                paragraph.level = int(item.get("level", 0)) if isinstance(item, dict) else 0
        counters["bullets"] += len(spec.get("bullets", []))
    else:
        top = 1.5
        if spec.get("bullets"):
            box = slide.shapes.add_textbox(Inches(0.8), Inches(top), Inches(11.6), Inches(0.9))
            frame = box.text_frame
            frame.word_wrap = True
            for index, item in enumerate(spec["bullets"]):
                paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
                paragraph.text = "• " + str(item)
                paragraph.font.size = Pt(18)
            top += 1.0
        if isinstance(spec.get("table"), dict):
            table_spec = spec["table"]
            rows = 1 + len(table_spec.get("rows", []))
            columns = len(table_spec.get("headers", [])) or (len(table_spec["rows"][0]) if table_spec.get("rows") else 0)
            if rows > 0 and columns > 0:
                shape = slide.shapes.add_table(rows, columns, Inches(0.8), Inches(top), Inches(11.6), Inches(0.5 * rows))
                table = shape.table
                for index, value in enumerate(table_spec.get("headers", [])):
                    table.cell(0, index).text = str(value)
                for row_index, row in enumerate(table_spec.get("rows", []), start=1):
                    for column_index, value in enumerate(row):
                        table.cell(row_index, column_index).text = "" if value is None else str(value)
                counters["tables"] += 1
                top += 0.5 * rows + 0.3
        if isinstance(spec.get("chart"), dict):
            chart_spec = dict(spec["chart"])
            chart_spec.setdefault("top", round(top, 2))
            chart_spec.setdefault("height", round(min(4.2, max(1.5, 7.1 - top)), 2))
            _pptx_add_chart(slide, chart_spec)
            counters["charts"] += 1
            top += float(chart_spec["height"]) + 0.2
        if isinstance(spec.get("image"), dict):
            image = spec["image"]
            path = require_file(str(image.get("path", "")), "image")
            slide.shapes.add_picture(
                str(path),
                Inches(float(image.get("left", 0.8))),
                Inches(float(image.get("top", top))),
                width=Inches(float(image.get("width", 3.0))),
            )
            counters["images"] += 1
    counters["slides"] += 1
    return slide


def pptx_create(spec: dict[str, Any], out: Path) -> dict[str, Any]:
    from pptx import Presentation

    prs = Presentation()
    prs.slide_width = int(13.333 * 914400)
    prs.slide_height = int(7.5 * 914400)
    prs.core_properties.title = str(spec.get("title", ""))
    prs.core_properties.author = str(spec.get("author", "DSH Office Suite"))
    # python-pptx leaves its own name in the properties; a delivered deck should
    # carry the author's, not the library's.
    prs.core_properties.comments = str(spec.get("comments", ""))
    prs.core_properties.last_modified_by = str(spec.get("author", "DSH Office Suite"))
    counters = {key: 0 for key in ("slides", "bullets", "tables", "charts", "images")}
    for slide_spec in spec.get("slides", []):
        if isinstance(slide_spec, dict):
            _pptx_add_slide(prs, slide_spec, counters)
    ensure_parent(str(out))
    prs.save(str(out))
    return {"ok": True, "kind": "pptx", "command": "create", "output": str(out), "counts": counters}


def _pptx_iter_text(slide, out: list[str]) -> None:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                text = "".join(run.text for run in paragraph.runs).strip()
                if text:
                    out.append(text)


def pptx_read(input_path: Path) -> dict[str, Any]:
    from pptx import Presentation

    prs = Presentation(str(input_path))
    slides = []
    for index, slide in enumerate(prs.slides, start=1):
        texts: list[str] = []
        _pptx_iter_text(slide, texts)
        tables = []
        charts = []
        images = 0
        for shape in slide.shapes:
            if shape.has_table:
                table = shape.table
                tables.append({"rows": [[cell.text for cell in row.cells] for row in table.rows]})
            if shape.has_chart:
                chart = shape.chart
                categories = list(chart.plots[0].categories) if chart.plots else []
                series = []
                try:
                    for item in chart.series:
                        series.append({"name": item.name, "values": list(item.values)})
                except Exception:
                    series = []
                charts.append({
                    "type": str(chart.chart_type),
                    "title": chart.chart_title.text_frame.text if chart.has_title else None,
                    "categories": [str(value) for value in categories],
                    "series": series,
                })
            if shape.shape_type is not None and "PICTURE" in str(shape.shape_type):
                images += 1
        title = slide.shapes.title.text if slide.shapes.title is not None else None
        slides.append({"slide": index, "title": title, "texts": texts, "tables": tables, "charts": charts, "images": images})
    return {
        "ok": True,
        "kind": "pptx",
        "command": "read",
        "input": str(input_path),
        "slide_count": len(prs.slides),
        "title": prs.core_properties.title or None,
        "slides": slides,
        "counts": {
            "slides": len(prs.slides),
            "tables": sum(len(item["tables"]) for item in slides),
            "charts": sum(len(item["charts"]) for item in slides),
        },
    }


# ---------------------------------------------------------------------------
# XLSX
# ---------------------------------------------------------------------------


def _xlsx_value(value: Any):
    if isinstance(value, str) and value.startswith("="):
        return value
    return value


def _xlsx_cells(worksheet, address: str):
    """Yield every cell in a single address or an A1:B2 range."""
    from openpyxl.utils import range_boundaries

    if ":" not in address:
        yield worksheet[address]
        return
    min_col, min_row, max_col, max_row = range_boundaries(address)
    for row in worksheet.iter_rows(min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row):
        for cell in row:
            yield cell


def _xlsx_apply_style(cell, style: dict[str, Any]) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill

    font = {}
    if style.get("bold"):
        font["bold"] = True
    if style.get("italic"):
        font["italic"] = True
    if style.get("size"):
        font["size"] = float(style["size"])
    if style.get("color"):
        font["color"] = str(style["color"]).lstrip("#").upper()
    if font:
        cell.font = Font(**font)
    if style.get("fill"):
        cell.fill = PatternFill("solid", fgColor=str(style["fill"]).lstrip("#").upper())
    if style.get("number_format"):
        cell.number_format = str(style["number_format"])
    if style.get("align"):
        cell.alignment = Alignment(horizontal=str(style["align"]), vertical="center")


def _xlsx_add_chart(worksheet, spec: dict[str, Any]) -> None:
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference

    chart_type = str(spec.get("type", "bar")).lower()
    chart = {"bar": BarChart, "column": BarChart, "line": LineChart, "pie": PieChart}.get(chart_type, BarChart)()
    chart.title = str(spec.get("title", ""))
    chart.height = float(spec.get("height", 8))
    chart.width = float(spec.get("width", 14))
    data = spec.get("data")
    if isinstance(data, dict):
        min_col = int(data.get("min_col", 2))
        min_row = int(data.get("min_row", 1))
        max_col = int(data.get("max_col", min_col))
        max_row = int(data.get("max_row", worksheet.max_row))
        values = Reference(worksheet, min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row)
        chart.add_data(values, titles_from_data=bool(data.get("titles_from_data", True)))
        categories = data.get("categories")
        if isinstance(categories, dict):
            chart.set_categories(Reference(
                worksheet,
                min_col=int(categories.get("min_col", 1)),
                min_row=int(categories.get("min_row", 2)),
                max_row=int(categories.get("max_row", max_row)),
            ))
    worksheet.add_chart(chart, str(spec.get("anchor", "E2")))


def xlsx_create(spec: dict[str, Any], out: Path) -> dict[str, Any]:
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.drawing.image import Image as XLImage

    workbook = Workbook()
    workbook.remove(workbook.active)
    counters = {"sheets": 0, "charts": 0, "images": 0}
    for sheet_spec in spec.get("sheets", []):
        if not isinstance(sheet_spec, dict):
            continue
        title = str(sheet_spec.get("name", f"Sheet{counters['sheets'] + 1}"))[:31]
        worksheet = workbook.create_sheet(title)
        counters["sheets"] += 1
        for row_index, row in enumerate(sheet_spec.get("rows", []), start=1):
            for column_index, value in enumerate(row, start=1):
                worksheet.cell(row=row_index, column=column_index, value=_xlsx_value(value))
        for address, formula in (sheet_spec.get("formulas") or {}).items():
            worksheet[str(address)] = str(formula)
        for address, style in (sheet_spec.get("styles") or {}).items():
            if isinstance(style, dict):
                for cell in _xlsx_cells(worksheet, str(address)):
                    _xlsx_apply_style(cell, style)
        for column, width in (sheet_spec.get("widths") or {}).items():
            worksheet.column_dimensions[str(column)].width = float(width)
        for row, height in (sheet_spec.get("row_heights") or {}).items():
            worksheet.row_dimensions[int(row)].height = float(height)
        if sheet_spec.get("freeze"):
            worksheet.freeze_panes = str(sheet_spec["freeze"])
        for merged in sheet_spec.get("merged", []):
            worksheet.merge_cells(str(merged))
        for chart_spec in sheet_spec.get("charts", []):
            if isinstance(chart_spec, dict):
                _xlsx_add_chart(worksheet, chart_spec)
                counters["charts"] += 1
        for image_spec in sheet_spec.get("images", []):
            if isinstance(image_spec, dict):
                path = require_file(str(image_spec.get("path", "")), "image")
                image = XLImage(str(path))
                if image_spec.get("width"):
                    image.width = int(image_spec["width"])
                if image_spec.get("height"):
                    image.height = int(image_spec["height"])
                worksheet.add_image(image, str(image_spec.get("anchor", "A1")))
                counters["images"] += 1
        worksheet.sheet_view.showGridLines = bool(sheet_spec.get("gridlines", True))
    if not workbook.sheetnames:
        workbook.create_sheet("Sheet1")
    ensure_parent(str(out))
    workbook.save(str(out))
    return {"ok": True, "kind": "xlsx", "command": "create", "output": str(out), "counts": counters}


def xlsx_read(input_path: Path, data_only: bool = False) -> dict[str, Any]:
    from openpyxl import load_workbook

    workbook = load_workbook(str(input_path), data_only=data_only)
    sheets = []
    for worksheet in workbook.worksheets:
        rows = []
        for row in worksheet.iter_rows():
            values = [cell.value for cell in row]
            while values and values[-1] is None:
                values.pop()
            if values:
                rows.append(values)
        formulas = []
        for row in worksheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formulas.append({"cell": cell.coordinate, "formula": cell.value, "cached": None})
        sheets.append({
            "name": worksheet.title,
            "dimensions": worksheet.dimensions,
            "rows": rows,
            "formulas": formulas,
            "merged": [str(item) for item in worksheet.merged_cells.ranges],
            "charts": len(worksheet._charts),
            "images": len(worksheet._images),
        })
    workbook.close()
    return {
        "ok": True,
        "kind": "xlsx",
        "command": "read",
        "input": str(input_path),
        "data_only": data_only,
        "sheets": sheets,
        "counts": {
            "sheets": len(sheets),
            "formulas": sum(len(sheet["formulas"]) for sheet in sheets),
            "charts": sum(sheet["charts"] for sheet in sheets),
        },
    }


def xlsx_recalc(input_path: Path, out: Path) -> dict[str, Any]:
    """Recalculate formulas through the LibreOffice kit and store the cached values."""
    ensure_parent(str(out))
    staging = out.with_name(f"{out.name}.recalc-{os.getpid()}.xlsx")
    report = _kit_call(["recalculate", "--input", str(input_path), "--output", str(staging)])
    produced = Path(report.get("outputPath") or staging)
    if not produced.is_file():
        fail("the LibreOffice kit reported a recalculation but wrote no workbook")
    os.replace(produced, out)
    cached = 0
    try:
        import openpyxl

        workbook = openpyxl.load_workbook(str(out), data_only=True)
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows():
                cached += sum(1 for cell in row if cell.value is not None)
        workbook.close()
    except Exception:
        cached = 0
    return {
        "ok": True,
        "kind": "xlsx",
        "command": "recalc",
        "input": str(input_path),
        "output": str(out),
        "renderer": "libreoffice-kit",
        "engine": report.get("backend"),
        "executable": find_kit()[1] if find_kit() else None,
        "missing_fonts": [str(name) for name in report.get("missingFonts") or []],
        "cached_values": cached,
    }


# ---------------------------------------------------------------------------
# Status and dispatch
# ---------------------------------------------------------------------------


def status() -> dict[str, Any]:
    import importlib
    import platform

    packages = []
    for name, module_name in [
        ("python-docx", "docx"),
        ("openpyxl", "openpyxl"),
        ("python-pptx", "pptx"),
        ("pypdf", "pypdf"),
        ("reportlab", "reportlab"),
        ("pdfplumber", "pdfplumber"),
        ("pymupdf", "pymupdf"),
        ("pillow", "PIL"),
        ("latex2mathml", "latex2mathml"),
        ("mathml2omml", "mathml2omml"),
    ]:
        try:
            module = importlib.import_module(module_name)
            packages.append({"package": name, "installed": True, "version": str(getattr(module, "__version__", ""))})
        except Exception as exc:
            packages.append({"package": name, "installed": False, "error": f"{type(exc).__name__}: {exc}"})
    installed = {item["package"]: item["installed"] for item in packages}
    kit = kit_report()
    ready = bool(kit.get("available") and kit.get("backend"))
    return {
        "ok": True,
        "command": "status",
        "plugin_version": VERSION,
        "python": {"executable": sys.executable, "version": platform.python_version()},
        "platform": platform.platform(),
        "packages": packages,
        "renderers": {
            "kit": kit,
            "can_render": {
                "docx": ready,
                "pptx": ready,
                "xlsx": ready,
                "pdf": bool(installed.get("pymupdf")),
            },
            "can_recalculate": {"xlsx": ready},
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="DSH Office Suite helper CLI")
    subparsers = parser.add_subparsers(dest="format", required=True)

    for kind in ("docx", "pdf", "pptx", "xlsx"):
        kind_parser = subparsers.add_parser(kind)
        commands = kind_parser.add_subparsers(dest="command", required=True)

        create = commands.add_parser("create")
        create.add_argument("--spec", required=True)
        create.add_argument("--out", required=True)
        if kind in {"docx", "pptx", "xlsx"}:
            create.add_argument("--json-out")

        read = commands.add_parser("read")
        read.add_argument("--input", required=True)
        read.add_argument("--json-out")
        if kind == "xlsx":
            read.add_argument("--data-only", action="store_true")

        render = commands.add_parser("render")
        render.add_argument("--input", required=True)
        render.add_argument("--out-dir", required=True)
        render.add_argument("--dpi", type=int, default=110)
        render.add_argument("--json-out")
        if kind in {"docx", "pptx", "xlsx"}:
            render.add_argument("--pdf", action="store_true", help="also export a PDF next to the page images")

        if kind == "xlsx":
            recalc = commands.add_parser("recalc")
            recalc.add_argument("--input", required=True)
            recalc.add_argument("--out", required=True)
            recalc.add_argument("--json-out")

    status_parser = subparsers.add_parser("status")

    args = parser.parse_args()
    try:
        if args.format == "status":
            emit(status())
            return
        if args.command == "create":
            spec = load_json(args.spec)
            out = Path(args.out)
            if args.format == "docx":
                result = docx_create(spec, out)
            elif args.format == "pdf":
                result = pdf_create(spec, out)
            elif args.format == "pptx":
                result = pptx_create(spec, out)
            else:
                result = xlsx_create(spec, out)
        elif args.command == "read":
            input_path = require_file(args.input)
            if args.format == "docx":
                result = docx_read(input_path)
            elif args.format == "pdf":
                result = pdf_read(input_path)
            elif args.format == "pptx":
                result = pptx_read(input_path)
            else:
                result = xlsx_read(input_path, data_only=bool(getattr(args, "data_only", False)))
        elif args.command == "render":
            input_path = require_file(args.input)
            out_dir = Path(args.out_dir)
            if args.format == "pdf":
                result = pdf_render(input_path, out_dir, args.dpi)
            elif args.format == "xlsx":
                result = xlsx_render(input_path, out_dir, args.dpi)
            else:
                result = render_office(input_path, args.format, out_dir, args.dpi, emit_pdf=bool(getattr(args, "pdf", False)))
        elif args.command == "recalc":
            result = xlsx_recalc(require_file(args.input), Path(args.out))
        else:
            fail(f"unsupported command: {args.command}")
        emit(result, getattr(args, "json_out", None))
    except SystemExit:
        raise
    except Exception as exc:
        fail(f"{type(exc).__name__}: {exc}")


if __name__ == "__main__":
    main()
