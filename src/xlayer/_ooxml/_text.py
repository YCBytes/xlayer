"""Shared Excel text decoding for shared strings and inline strings.

Whitespace stripping and Excel ``_xHHHH_`` unescaping live here so the
shared-string table and worksheet inline / ``t="str"`` cells cannot drift.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from xml.etree import ElementTree as ET

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
T_TAG = f"{{{_NS}}}t"
R_TAG = f"{{{_NS}}}r"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_DEFAULT_WS = " \t\r\n"
_EXCEL_ESCAPE_RE = re.compile(r"_x([0-9A-Fa-f]{4})_")

SURROGATE_REASON = "unpaired UTF-16 surrogate in Excel escape"
NESTED_MARKUP_REASON = "unexpected child element in text content"


@dataclass(frozen=True)
class TextError:
    """Why text could not be decoded. Callers turn this into a refusal."""

    reason: str


def inherited_space(elem: ET.Element, parent_space: str) -> str:
    return elem.attrib.get(XML_SPACE, parent_space)


def decode_excel_escapes(text: str) -> str | None:
    """Decode ``_xHHHH_`` escapes in one left-to-right pass.

    Because the pass never re-reads its own output, ``_x005F_x000D_`` yields
    the literal ``_x000D_``. A high surrogate must be immediately followed
    by a low surrogate; the two combine into one character. Any other
    surrogate is malformed and returns ``None``.
    """
    out: list[str] = []
    position = 0
    matches = list(_EXCEL_ESCAPE_RE.finditer(text))
    index = 0
    while index < len(matches):
        match = matches[index]
        out.append(text[position : match.start()])
        code = int(match.group(1), 16)
        if _is_low_surrogate(code):
            return None
        if _is_high_surrogate(code):
            following = matches[index + 1] if index + 1 < len(matches) else None
            if following is None or following.start() != match.end():
                return None
            low = int(following.group(1), 16)
            if not _is_low_surrogate(low):
                return None
            out.append(chr(0x10000 + ((code - 0xD800) << 10) + (low - 0xDC00)))
            position = following.end()
            index += 2
            continue
        out.append(chr(code))
        position = match.end()
        index += 1
    out.append(text[position:])
    return "".join(out)


def decode_rich_text(elem: ET.Element, inherited: str) -> str | TextError:
    """Concatenate ``<t>`` / ``<r><t>`` descendants, excluding phonetic guides."""
    item_space = inherited_space(elem, inherited)
    parts: list[str] = []
    for child in elem:
        if child.tag == T_TAG:
            piece = _text_of(child, item_space)
            if isinstance(piece, TextError):
                return piece
            parts.append(piece)
        elif child.tag == R_TAG:
            run_space = inherited_space(child, item_space)
            for run_child in child:
                if run_child.tag == T_TAG:
                    piece = _text_of(run_child, run_space)
                    if isinstance(piece, TextError):
                        return piece
                    parts.append(piece)
    out: list[str] = []
    for part in parts:
        decoded = decode_excel_escapes(part)
        if decoded is None:
            return TextError(SURROGATE_REASON)
        out.append(decoded)
    return "".join(out)


def _text_of(t_elem: ET.Element, inherited_space: str) -> str | TextError:
    # <t> is simple content: a child element here means the source is
    # malformed, and reading .text alone would silently drop the rest.
    if len(t_elem):
        return TextError(NESTED_MARKUP_REASON)
    space = t_elem.attrib.get(XML_SPACE, inherited_space)
    text = t_elem.text or ""
    if space != "preserve":
        text = text.strip(_DEFAULT_WS)
    return text


def _is_high_surrogate(code: int) -> bool:
    return 0xD800 <= code <= 0xDBFF


def _is_low_surrogate(code: int) -> bool:
    return 0xDC00 <= code <= 0xDFFF
