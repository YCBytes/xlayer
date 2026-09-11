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

from xlayer._errors import Refusal
from xlayer._ooxml._text import XML_SPACE, decode_rich_text
from xlayer._ooxml.archive import WorkbookArchive
from xlayer._ooxml.package import TARGET_MODE_INTERNAL, PackageInfo

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_SI_TAG = f"{{{_NS}}}si"
_SST_ROOT_TAG = f"{{{_NS}}}sst"

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

    root_space = root.attrib.get(XML_SPACE, "default")
    strings: list[str] = []
    for index, si_elem in enumerate(root.findall(_SI_TAG)):
        text = decode_rich_text(si_elem, root_space)
        if text is None:
            return _invalid_part(
                part,
                "unpaired UTF-16 surrogate in Excel escape",
                string_index=index,
            )
        strings.append(text)
    return tuple(strings)


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
