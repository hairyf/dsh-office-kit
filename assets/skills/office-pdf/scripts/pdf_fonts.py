"""Font handling shared by the PDF scripts.

reportlab's base-14 fonts cover Latin-1 only, so a Chinese, Japanese, or Korean
stamp or form value needs a real font file. These helpers find one, or report
that the caller must pass --font-file.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

CJK_FONT_CANDIDATES = {
    "win32": [
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/msyhbd.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/simsun.ttc",
    ],
    "linux": [
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
    ],
    "darwin": [
        "/System/Library/Fonts/STHeiti Medium.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/System/Library/Fonts/Supplemental/Songti.ttc",
    ],
}


def has_non_latin(text) -> bool:
    """True when text needs a font beyond the base-14 Latin sets."""
    return isinstance(text, str) and any(ord(character) > 0xFF for character in text)


def find_cjk_font(explicit=None) -> Path | None:
    """Return a usable CJK font path, honouring an explicit path first."""
    if explicit:
        path = Path(explicit).expanduser()
        if not path.is_file():
            raise FileNotFoundError(f"font file not found: {path}")
        return path
    override = os.environ.get("DSH_OFFICE_CJK_FONT")
    if override and Path(override).expanduser().is_file():
        return Path(override).expanduser()
    platform_key = "win32" if sys.platform.startswith("win") else "darwin" if sys.platform == "darwin" else "linux"
    for candidate in CJK_FONT_CANDIDATES[platform_key]:
        path = Path(candidate)
        if path.is_file():
            return path
    return None


def register_cjk_font(path: Path, name: str) -> str:
    """Register a TTF/OTF/TTC file with reportlab and return the resource name."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    kwargs = {}
    if path.suffix.lower() in {".ttc", ".otc", ".dfont"}:
        kwargs["subfontIndex"] = 0
    pdfmetrics.registerFont(TTFont(name, str(path), **kwargs))
    return name
