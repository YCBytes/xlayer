"""Shared string table: displayed text from the workbook string part.

Adapted from ``xlayer-core`` ``parser/strings.py``. Differences from core
are deliberate:

- parts are read through :meth:`WorkbookArchive.parse_xml_part` after
  :meth:`PackageInfo.find_workbook_rel`, so an external or unresolvable
  relationship is not treated as a missing table;
- ``xml:space`` is inherited from ``<sst>`` through ``<si>`` and ``<r>``
  down to ``<t>``; any element may override with an explicit
  ``preserve`` or ``default``;
- default whitespace stripping removes only space/tab/CR/LF, never
  non-breaking spaces;
- Excel ``_xHHHH_`` escapes are decoded in a single non-recursive pass,
  so ``_x005F_`` protects a following literal ``_x…_`` sequence; a valid
  UTF-16 surrogate pair becomes one character, and an unpaired or
  reversed surrogate is refused as ``invalid_part_content`` rather than
  emitted as text that cannot be UTF-8 encoded; and
- phonetic guides (``<rPh>``, ``<phoneticPr>``) remain excluded.
"""

from __future__ import annotations

import re
from xml.etree import ElementTree as ET

from xlayer._errors import Refusal
from xlayer._ooxml.archive import WorkbookArchive
from xlayer._ooxml.package import TARGET_MODE_INTERNAL, PackageInfo

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_SST_ROOT_TAG = f"{{{_NS}}}sst"
_SI_TAG = f"{{{_NS}}}si"
_T_TAG = f"{{{_NS}}}t"
_R_TAG = f"{{{_NS}}}r"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_DEFAULT_WS = " \t\r\n"
_EXCEL_ESCAPE_RE = re.compile(r"_x([0-9A-Fa-f]{4})_")

_RECOVER_TRUSTED_SOURCE = ({"action": "regenerate_workbook_from_trusted_source"},)


def parse_shared_strings(
    archive: WorkbookArchive, package: PackageInfo
) -> tuple[str, ...] | Refusal:
    """Return the shared string table, or a structured refusal."""
    rel = package.find_workbook_rel("sharedStrings")
    if rel is None:
        return ()
    if rel.target_mode != TARGET_MODE_INTERNAL:
        return _invalid_part(
            package.workbook_path,
            "sharedStrings relationship is external",
            target=rel.target,
        )
    part = package.resolve_workbook_rel_by_id(rel.rel_id, "sharedStrings")
    if part is None:
        return _unsafe_target(rel.target)
    if not archive.has_part(part):
        return _missing_part(part, "the sharedStrings relationship points at it")

    root = archive.parse_xml_part(part)
    if isinstance(root, Refusal):
        return root
    if root.tag != _SST_ROOT_TAG:
        return _invalid_part(part, "unexpected root element", root_tag=root.tag)

    # xml:space is inherited from the root down; <si>, <r>, and <t> may
    # each override it.
    root_space = root.attrib.get(_XML_SPACE, "default")
    strings: list[str] = []
    for index, si_elem in enumerate(root.findall(_SI_TAG)):
        text = _extract_text(si_elem, root_space)
        if text is None:
            return _invalid_part(
                part,
                "unpaired UTF-16 surrogate in Excel escape",
                string_index=index,
            )
        strings.append(text)
    return tuple(strings)


def _extract_text(si_elem: ET.Element, inherited: str) -> str | None:
    """Concatenate the item's text, or ``None`` when an escape is malformed."""
    item_space = si_elem.attrib.get(_XML_SPACE, inherited)
    parts: list[str] = []
    for child in si_elem:
        if child.tag == _T_TAG:
            parts.append(_text_of(child, item_space))
        elif child.tag == _R_TAG:
            run_space = child.attrib.get(_XML_SPACE, item_space)
            for run_child in child:
                if run_child.tag == _T_TAG:
                    parts.append(_text_of(run_child, run_space))
    decoded = [_decode_excel_escapes(part) for part in parts]
    if any(text is None for text in decoded):
        return None
    return "".join(text for text in decoded if text is not None)


def _text_of(t_elem: ET.Element, inherited_space: str) -> str:
    space = t_elem.attrib.get(_XML_SPACE, inherited_space)
    text = t_elem.text or ""
    if space != "preserve":
        text = text.strip(_DEFAULT_WS)
    return text


def _decode_excel_escapes(text: str) -> str | None:
    """Decode ``_xHHHH_`` escapes in one left-to-right pass.

    Because the pass never re-reads its own output, ``_x005F_x000D_`` yields
    the literal ``_x000D_``. A high surrogate must be immediately followed
    by a low surrogate; the two combine into one character. Any other
    surrogate is malformed and returns ``None`` so the caller can refuse
    rather than emit text that cannot be UTF-8 encoded.
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


def _is_high_surrogate(code: int) -> bool:
    return 0xD800 <= code <= 0xDBFF


def _is_low_surrogate(code: int) -> bool:
    return 0xDC00 <= code <= 0xDFFF


def _invalid_part(part: str, reason: str, **details: object) -> Refusal:
    return Refusal(
        code="invalid_part_content",
        message=f"Part {part!r} is structurally invalid: {reason}.",
        details={"part": part, "reason": reason, **details},
        recovery_options=_RECOVER_TRUSTED_SOURCE,
    )


def _missing_part(part: str, why: str) -> Refusal:
    return Refusal(
        code="missing_required_part",
        message=f"Required part {part!r} is missing or unusable: {why}.",
        details={"part": part, "reason": why},
        recovery_options=_RECOVER_TRUSTED_SOURCE,
    )


def _unsafe_target(target: str) -> Refusal:
    return Refusal(
        code="unsafe_archive_path",
        message=f"The sharedStrings relationship target {target!r} is not a safe package path.",
        details={"kind": "relationship_target", "target": target},
        recovery_options=_RECOVER_TRUSTED_SOURCE,
    )
