"""Workbook registry: sheets, defined names, and workbook-level properties.

Adapted from ``xlayer-core`` ``parser/workbook.py``. Differences from core
are deliberate public architecture, not drift:

- parts are read through :meth:`WorkbookArchive.parse_xml_part` using
  ``package.workbook_path`` (never a hardcoded ``xl/workbook.xml``);
- the root element is matched exactly;
- XML Schema booleans (``true`` / ``false`` / ``1`` / ``0``) are accepted
  for ``date1904`` and defined-name ``hidden``; any other spelling is a
  refusal, not a silent false;
- ``localSheetId`` is a tab index into document order, never ``sheetId``;
- defined-name expressions are kept as raw text — there is no derived
  "broken" flag from a ``#REF!`` substring;
- every place core would raise or paper over (missing attributes, unknown
  sheet state, duplicate names or ids, unresolvable sheet relationships)
  returns a structured refusal;
- ``chartsheet`` and ``dialogsheet`` tabs stay in the registry as metadata
  with paths resolved by exact relationship URI; and
- ``ref_mode`` is omitted: formula XML is always stored A1-style.
"""

from __future__ import annotations

from dataclasses import dataclass
from xml.etree import ElementTree as ET

from xlayer._errors import Refusal
from xlayer._ooxml.archive import (
    WorkbookArchive,
    normalize_package_path,
    validate_opc_target,
)
from xlayer._ooxml.package import (
    SUPPORTED_WORKBOOK_REL_TYPES,
    TARGET_MODE_INTERNAL,
    PackageInfo,
)

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_WORKBOOK_ROOT_TAG = f"{{{_NS}}}workbook"
_SHEETS_TAG = f"{{{_NS}}}sheets"
_SHEET_TAG = f"{{{_NS}}}sheet"
_DEFINED_NAMES_TAG = f"{{{_NS}}}definedNames"
_DEFINED_NAME_TAG = f"{{{_NS}}}definedName"
_BOOK_VIEWS_TAG = f"{{{_NS}}}bookViews"
_WORKBOOK_VIEW_TAG = f"{{{_NS}}}workbookView"
_WORKBOOK_PR_TAG = f"{{{_NS}}}workbookPr"
_EXTERNAL_REFERENCES_TAG = f"{{{_NS}}}externalReferences"
_EXTERNAL_REFERENCE_TAG = f"{{{_NS}}}externalReference"
_REL_ID_ATTR = f"{{{_NS_R}}}id"

_SHEET_KINDS = ("worksheet", "chartsheet", "dialogsheet")
_SHEET_STATES = frozenset({"visible", "hidden", "veryHidden"})
_XML_TRUE = frozenset({"true", "1"})
_XML_FALSE = frozenset({"false", "0"})

# OOXML integer attributes here are xsd:unsignedInt (0..4,294,967,295).
# Length is capped before conversion: str.isdigit() is true for some
# non-ASCII digits that int() rejects, and CPython raises ValueError on
# strings longer than sys.get_int_max_str_digits().
_MAX_UINT = 4_294_967_295
_MAX_UINT_DIGITS = 10

_RECOVER_TRUSTED_SOURCE = ({"action": "regenerate_workbook_from_trusted_source"},)


@dataclass(frozen=True)
class SheetEntry:
    name: str
    sheet_id: int
    rel_id: str
    state: str
    tab_index: int
    kind: str
    part_path: str


@dataclass(frozen=True)
class DefinedName:
    name: str
    text: str
    scope_sheet: str | None
    is_hidden: bool


@dataclass(frozen=True)
class WorkbookRegistry:
    sheets: tuple[SheetEntry, ...]
    defined_names: tuple[DefinedName, ...]
    active_tab: int
    date1904: bool
    external_reference_ids: tuple[str, ...]

    def sheet_by_name(self, name: str) -> SheetEntry | None:
        for sheet in self.sheets:
            if sheet.name == name:
                return sheet
        return None


def parse_workbook_registry(
    archive: WorkbookArchive, package: PackageInfo
) -> WorkbookRegistry | Refusal:
    """Parse the sheet registry from the package's workbook part."""
    part = package.workbook_path
    if not archive.has_part(part):
        return _missing_part(part, "the officeDocument relationship points at it")

    root = archive.parse_xml_part(part)
    if isinstance(root, Refusal):
        return root
    if root.tag != _WORKBOOK_ROOT_TAG:
        return _invalid_part(part, "unexpected root element", root_tag=root.tag)

    inventory = _validate_registry_inventory(root, part)
    if inventory is not None:
        return inventory

    sheets = _parse_sheets(root, archive, package, part)
    if isinstance(sheets, Refusal):
        return sheets

    defined_names = _parse_defined_names(root, sheets, part)
    if isinstance(defined_names, Refusal):
        return defined_names

    active_tab = _parse_active_tab(root, len(sheets), part)
    if isinstance(active_tab, Refusal):
        return active_tab

    date1904 = _parse_date1904(root, part)
    if isinstance(date1904, Refusal):
        return date1904

    external_ids = _parse_external_references(root, part)
    if isinstance(external_ids, Refusal):
        return external_ids

    return WorkbookRegistry(
        sheets=sheets,
        defined_names=defined_names,
        active_tab=active_tab,
        date1904=date1904,
        external_reference_ids=external_ids,
    )


def _validate_registry_inventory(root: ET.Element, part: str) -> Refusal | None:
    """Reject inventories that direct-child parsing would truncate or omit."""
    process = "{http://schemas.openxmlformats.org/markup-compatibility/2006}ProcessContent"
    tags = {_SHEETS_TAG, _SHEET_TAG, _DEFINED_NAMES_TAG, _DEFINED_NAME_TAG}
    admitted = {id(root)}
    for block_tag, item_tag, required in (
        (_SHEETS_TAG, _SHEET_TAG, True),
        (_DEFINED_NAMES_TAG, _DEFINED_NAME_TAG, False),
    ):
        blocks = [child for child in root if child.tag == block_tag]
        if len(blocks) > 1:
            return _invalid_part(part, "duplicate registry inventory block", tag=block_tag)
        if required and not blocks:
            return _invalid_part(part, "workbook has no sheets")
        for block in blocks:
            admitted.add(id(block))
            for item in block:
                if item.tag != item_tag:
                    return _invalid_part(part, "unexpected registry inventory child", tag=item.tag)
                if len(item):
                    return _invalid_part(part, "nested registry inventory payload", tag=item.tag)
                admitted.add(id(item))
    for element in root.iter():
        known = id(element) in admitted
        if element.tag in tags and not known:
            return _invalid_part(part, "displaced registry inventory fragment", tag=element.tag)
        if known and process in element.attrib:
            return _invalid_part(part, "unsupported registry ProcessContent", tag=element.tag)
    return None


def _parse_sheets(
    root: ET.Element, archive: WorkbookArchive, package: PackageInfo, part: str
) -> tuple[SheetEntry, ...] | Refusal:
    sheets_elem = root.find(_SHEETS_TAG)
    if sheets_elem is None:
        return _invalid_part(part, "workbook has no sheets")
    sheet_elems = sheets_elem.findall(_SHEET_TAG)
    if not sheet_elems:
        return _invalid_part(part, "workbook has no sheets")

    sheets: list[SheetEntry] = []
    seen_names: set[str] = set()
    seen_ids: set[int] = set()
    seen_paths: set[str] = set()
    for tab_index, sheet_elem in enumerate(sheet_elems):
        name = sheet_elem.attrib.get("name")
        if not name:
            return _invalid_part(part, "sheet missing name")
        folded = name.casefold()
        if folded in seen_names:
            return _invalid_part(part, "duplicate sheet name", name=name)
        raw_id = sheet_elem.attrib.get("sheetId")
        if not raw_id:
            return _invalid_part(part, "sheet missing sheetId")
        sheet_id = _parse_positive_int(raw_id)
        if sheet_id is None:
            return _invalid_part(part, "invalid sheetId", sheet_id=raw_id)
        if sheet_id in seen_ids:
            return _invalid_part(part, "duplicate sheetId", sheet_id=sheet_id)
        rel_id = sheet_elem.attrib.get(_REL_ID_ATTR)
        if not rel_id:
            return _invalid_part(part, "sheet missing r:id")
        state = sheet_elem.attrib.get("state", "visible")
        if state not in _SHEET_STATES:
            return _invalid_part(part, "invalid sheet state", state=state)

        resolved = _resolve_sheet_part(rel_id, archive, package, part)
        if isinstance(resolved, Refusal):
            return resolved
        kind, part_path = resolved
        if part_path in seen_paths:
            return _invalid_part(part, "duplicate sheet part", part_path=part_path)

        seen_names.add(folded)
        seen_ids.add(sheet_id)
        seen_paths.add(part_path)
        sheets.append(
            SheetEntry(
                name=name,
                sheet_id=sheet_id,
                rel_id=rel_id,
                state=state,
                tab_index=tab_index,
                kind=kind,
                part_path=part_path,
            )
        )
    return tuple(sheets)


def _resolve_sheet_part(
    rel_id: str, archive: WorkbookArchive, package: PackageInfo, part: str
) -> tuple[str, str] | Refusal:
    rel = package.workbook_rels.get(rel_id)
    kind = None if rel is None else _sheet_kind(rel.rel_type)
    if rel is None or kind is None:
        return _invalid_part(part, "sheet r:id is not a sheet relationship", rel_id=rel_id)
    if rel.target_mode != TARGET_MODE_INTERNAL:
        return _unsafe_target(rel.target)
    cleaned = validate_opc_target(rel.target)
    if cleaned is None:
        return _unsafe_target(rel.target)
    candidate = cleaned if cleaned.startswith("/") else f"{package.workbook_dir}/{cleaned}"
    part_path = normalize_package_path(candidate)
    if part_path is None:
        return _unsafe_target(rel.target)
    if not archive.has_part(part_path):
        return _missing_part(part_path, "a sheet relationship points at it")
    return kind, part_path


def _sheet_kind(rel_type: str) -> str | None:
    for kind in _SHEET_KINDS:
        if rel_type == SUPPORTED_WORKBOOK_REL_TYPES[kind]:
            return kind
    return None


def _parse_defined_names(
    root: ET.Element, sheets: tuple[SheetEntry, ...], part: str
) -> tuple[DefinedName, ...] | Refusal:
    names_elem = root.find(_DEFINED_NAMES_TAG)
    if names_elem is None:
        return ()
    names: list[DefinedName] = []
    for name_elem in names_elem.findall(_DEFINED_NAME_TAG):
        name = name_elem.attrib.get("name")
        if not name:
            return _invalid_part(part, "definedName missing name")
        hidden_raw = name_elem.attrib.get("hidden")
        if hidden_raw is None:
            is_hidden = False
        else:
            parsed_hidden = _xml_boolean(hidden_raw)
            if parsed_hidden is None:
                return _invalid_part(part, "invalid hidden", hidden=hidden_raw)
            is_hidden = parsed_hidden
        scope_sheet: str | None = None
        local_id_raw = name_elem.attrib.get("localSheetId")
        if local_id_raw is not None:
            local_id = _parse_non_negative_int(local_id_raw)
            if local_id is None:
                return _invalid_part(part, "invalid localSheetId", localSheetId=local_id_raw)
            if local_id >= len(sheets):
                return _invalid_part(part, "localSheetId out of range", localSheetId=local_id)
            scope_sheet = sheets[local_id].name
        names.append(
            DefinedName(
                name=name,
                text=name_elem.text or "",
                scope_sheet=scope_sheet,
                is_hidden=is_hidden,
            )
        )
    return tuple(names)


def _parse_active_tab(root: ET.Element, sheet_count: int, part: str) -> int | Refusal:
    active_tab = 0
    book_views = root.find(_BOOK_VIEWS_TAG)
    if book_views is not None:
        view = book_views.find(_WORKBOOK_VIEW_TAG)
        if view is not None and "activeTab" in view.attrib:
            raw = view.attrib["activeTab"]
            parsed = _parse_non_negative_int(raw)
            if parsed is None:
                return _invalid_part(part, "invalid activeTab", activeTab=raw)
            active_tab = parsed
    if active_tab >= sheet_count:
        return _invalid_part(part, "activeTab out of range", activeTab=active_tab)
    return active_tab


def _parse_date1904(root: ET.Element, part: str) -> bool | Refusal:
    workbook_pr = root.find(_WORKBOOK_PR_TAG)
    if workbook_pr is None or "date1904" not in workbook_pr.attrib:
        return False
    parsed = _xml_boolean(workbook_pr.attrib["date1904"])
    if parsed is None:
        return _invalid_part(part, "invalid date1904", date1904=workbook_pr.attrib["date1904"])
    return parsed


def _parse_external_references(root: ET.Element, part: str) -> tuple[str, ...] | Refusal:
    refs_elem = root.find(_EXTERNAL_REFERENCES_TAG)
    if refs_elem is None:
        return ()
    ids: list[str] = []
    for ref_elem in refs_elem.findall(_EXTERNAL_REFERENCE_TAG):
        rel_id = ref_elem.attrib.get(_REL_ID_ATTR)
        if not rel_id:
            return _invalid_part(part, "externalReference missing r:id")
        ids.append(rel_id)
    return tuple(ids)


def _xml_boolean(raw: str) -> bool | None:
    if raw in _XML_TRUE:
        return True
    if raw in _XML_FALSE:
        return False
    return None


def _parse_unsigned_int(raw: str) -> int | None:
    if not raw.isascii() or not raw.isdigit() or len(raw) > _MAX_UINT_DIGITS:
        return None
    try:
        value = int(raw)
    except ValueError:
        return None
    if value > _MAX_UINT:
        return None
    return value


def _parse_positive_int(raw: str) -> int | None:
    value = _parse_unsigned_int(raw)
    if value is None or value < 1:
        return None
    return value


def _parse_non_negative_int(raw: str) -> int | None:
    return _parse_unsigned_int(raw)


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
        message=f"The sheet relationship target {target!r} is not a safe package path.",
        details={"kind": "relationship_target", "target": target},
        recovery_options=_RECOVER_TRUSTED_SOURCE,
    )
