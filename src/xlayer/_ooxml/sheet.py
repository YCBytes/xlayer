"""Worksheet cells, formulas, and merges.

Adapted from ``xlayer-core`` ``parser/sheet.py``. Differences are
deliberate:

- the sheet part is the registry's ``part_path``, never a guessed filename;
- a cached formula result is ``stored_value`` and is never treated as a
  verified calculation;
- omitted style index ``s`` resolves ``styles[0]``, which is not assumed
  to be General;
- shared-formula followers keep ``Formula.text is None`` rather than a
  rewritten formula;
- malformed cells refuse the worksheet instead of synthesizing a
  placeholder value;
- ``other_elements`` lists unconsumed top-level qualified names and is
  not an inventory of unsupported features.
"""

from __future__ import annotations

import bisect
import heapq
import math
import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, replace
from types import MappingProxyType
from xml.etree import ElementTree as ET

from xlayer._errors import Refusal
from xlayer._ooxml._text import (
    NESTED_MARKUP_REASON,
    SURROGATE_REASON,
    TextError,
    decode_excel_escapes,
    decode_rich_text,
    inherited_space,
)
from xlayer._ooxml.archive import WorkbookArchive
from xlayer._ooxml.styles import StyleInfo
from xlayer._ooxml.workbook import SheetEntry

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_WORKSHEET_ROOT_TAG = f"{{{_NS}}}worksheet"
_SHEET_DATA_TAG = f"{{{_NS}}}sheetData"
_DIMENSION_TAG = f"{{{_NS}}}dimension"
_ROW_TAG = f"{{{_NS}}}row"
_CELL_TAG = f"{{{_NS}}}c"
_V_TAG = f"{{{_NS}}}v"
_IS_TAG = f"{{{_NS}}}is"
_F_TAG = f"{{{_NS}}}f"
_MERGE_CELLS_TAG = f"{{{_NS}}}mergeCells"
_MERGE_CELL_TAG = f"{{{_NS}}}mergeCell"

_MAX_UINT = 4_294_967_295
_MAX_UINT_DIGITS = 10
_MAX_ROW = 1_048_576
_MAX_COL = 16_384  # XFD

# Explicit bound on merge work. Overlap checking and cell-to-merge lookup are
# both swept rather than compared pairwise, but the count is still capped so
# the cost per worksheet is bounded by something other than the XML element
# limit. Real workbooks in the corpus peak at fewer than a hundred merges.
MAX_MERGE_RANGES = 65_536
_ADDRESS_RE = re.compile(r"^([A-Z]{1,3})([1-9][0-9]{0,6})$")
_NUMBER_RE = re.compile(r"^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$")
_DATE_RE = re.compile(r"^([0-9]{4})-([0-9]{2})-([0-9]{2})$")
_TIME_RE = re.compile(r"^([0-9]{2}):([0-9]{2}):([0-9]{2})(?:\.[0-9]+)?$")
_KNOWN_CELL_TYPES = frozenset({"n", "s", "str", "inlineStr", "b", "e", "d"})
_MONTH_DAYS = (0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
_XML_TRUE = frozenset({"true", "1"})
_XML_FALSE = frozenset({"false", "0"})
# ECMA-376 ST_CellFormulaType. "normal" is the default, and may be either
# omitted or spelled out; both forms mean the same thing.
_FORMULA_KINDS = {
    None: "normal",
    "normal": "normal",
    "shared": "shared",
    "array": "array",
    "dataTable": "data_table",
}
# Kinds whose formula expression is stored as text. A shared follower and a
# data table define themselves through attributes instead.
_FORMULA_KINDS_REQUIRING_TEXT = frozenset({"normal", "array"})

_RECOVER_TRUSTED_SOURCE = ({"action": "regenerate_workbook_from_trusted_source"},)

_CONSUMED_TOP_LEVEL = frozenset({_SHEET_DATA_TAG, _DIMENSION_TAG, _MERGE_CELLS_TAG})


@dataclass(frozen=True)
class Formula:
    text: str | None
    kind: str
    shared_index: int | None
    ref: str | None
    calculate_on_next_recalc: bool


@dataclass(frozen=True)
class SharedFormula:
    master: str
    ref: str
    text: str


@dataclass(frozen=True)
class MergeRange:
    ref: str
    first_cell: str
    last_cell: str
    min_row: int
    min_col: int
    max_row: int
    max_col: int

    def contains(self, row: int, column: int) -> bool:
        return self.min_row <= row <= self.max_row and self.min_col <= column <= self.max_col


@dataclass(frozen=True)
class Cell:
    address: str
    row: int
    column: int
    stored_type: str
    value_kind: str
    raw: str | None
    stored_value: float | bool | str | None
    style_index: int
    style: StyleInfo
    formula: Formula | None
    merge_ref: str | None


@dataclass(frozen=True)
class Worksheet:
    """Parsed worksheet. Mappings are read-only copies."""

    name: str
    part_path: str
    cells: Mapping[str, Cell]
    merges: tuple[MergeRange, ...]
    shared_formulas: Mapping[int, SharedFormula]
    declared_dimension: str | None
    used_range: str | None
    other_elements: tuple[str, ...]

    def __post_init__(self) -> None:
        object.__setattr__(self, "cells", MappingProxyType(dict(self.cells)))
        object.__setattr__(self, "shared_formulas", MappingProxyType(dict(self.shared_formulas)))

    def merge_at(self, address: str) -> MergeRange | None:
        parsed = _parse_address(address)
        if parsed is None:
            return None
        row, column = parsed
        for merge in self.merges:
            if merge.contains(row, column):
                return merge
        return None


def parse_worksheet(
    archive: WorkbookArchive,
    sheet: SheetEntry,
    shared_strings: Sequence[str],
    styles: Sequence[StyleInfo],
) -> Worksheet | Refusal:
    """Return one worksheet's cells, or a structured refusal."""
    if sheet.kind != "worksheet":
        return Refusal(
            code="unsupported_sheet_kind",
            message=(
                f"Sheet {sheet.name!r} is a {sheet.kind}, which this reader does not interpret."
            ),
            details={"kind": sheet.kind, "part": sheet.part_path, "name": sheet.name},
            recovery_options=_RECOVER_TRUSTED_SOURCE,
        )
    part = sheet.part_path
    if isinstance(shared_strings, (str, bytes)):
        raise TypeError("shared_strings must be a sequence of strings")
    if not archive.has_part(part):
        return _missing_part(part, "the sheet relationship points at it")

    root = archive.parse_xml_part(part)
    if isinstance(root, Refusal):
        return root
    if root.tag != _WORKSHEET_ROOT_TAG:
        return _invalid_part(part, "unexpected root element", root_tag=root.tag)

    sheet_data_children = [child for child in root if child.tag == _SHEET_DATA_TAG]
    if not sheet_data_children:
        return _invalid_part(part, "worksheet missing sheetData")
    if len(sheet_data_children) > 1:
        return _invalid_part(part, "duplicate sheetData")
    sheet_data = sheet_data_children[0]

    dimension_elem = root.find(_DIMENSION_TAG)
    declared_dimension: str | None = None
    if dimension_elem is not None:
        declared_dimension = dimension_elem.attrib.get("ref")
        if not declared_dimension:
            return _invalid_part(part, "dimension missing ref")

    other_elements = tuple(child.tag for child in root if child.tag not in _CONSUMED_TOP_LEVEL)

    merges = _parse_merges(root, part)
    if isinstance(merges, Refusal):
        return merges

    cells: dict[str, Cell] = {}
    shared_formulas: dict[int, SharedFormula] = {}
    followers: list[tuple[str, int]] = []
    parsed = _parse_rows(
        sheet_data,
        part,
        styles,
        shared_strings,
        cells,
        inherited_space(root, "default"),
        shared_formulas,
        followers,
    )
    if isinstance(parsed, Refusal):
        return parsed
    checked = _check_shared_followers(part, cells, shared_formulas, followers)
    if isinstance(checked, Refusal):
        return checked
    if merges:
        cells = _assign_merge_refs(cells, merges)

    return Worksheet(
        name=sheet.name,
        part_path=part,
        cells=cells,
        merges=merges,
        shared_formulas=shared_formulas,
        declared_dimension=declared_dimension,
        used_range=_used_range(cells),
        other_elements=other_elements,
    )


def _parse_merges(root: ET.Element, part: str) -> tuple[MergeRange, ...] | Refusal:
    blocks = [child for child in root if child.tag == _MERGE_CELLS_TAG]
    if len(blocks) > 1:
        return _invalid_part(part, "duplicate mergeCells")
    if not blocks:
        return ()
    merge_elems = blocks[0].findall(_MERGE_CELL_TAG)
    if len(merge_elems) > MAX_MERGE_RANGES:
        return _invalid_part(
            part, "too many merge ranges", count=len(merge_elems), limit=MAX_MERGE_RANGES
        )
    ranges: list[MergeRange] = []
    for merge_elem in merge_elems:
        ref = merge_elem.attrib.get("ref")
        if not ref or ":" not in ref:
            return _invalid_part(part, "invalid merge ref", ref=ref)
        bounds = _parse_range_ref(ref)
        if bounds is None:
            return _invalid_part(part, "invalid merge ref", ref=ref)
        min_row, min_col, max_row, max_col = bounds
        if min_row == max_row and min_col == max_col:
            return _invalid_part(part, "invalid merge ref", ref=ref)
        candidate = MergeRange(
            ref=ref,
            first_cell=f"{_column_letter(min_col)}{min_row}",
            last_cell=f"{_column_letter(max_col)}{max_row}",
            min_row=min_row,
            min_col=min_col,
            max_row=max_row,
            max_col=max_col,
        )
        ranges.append(candidate)

    clash = _find_overlapping_pair(
        [(item.min_row, item.min_col, item.max_row, item.max_col) for item in ranges]
    )
    if clash is not None:
        first, second = sorted(clash)
        return _invalid_part(
            part,
            "overlapping merge ranges",
            ref=ranges[second].ref,
            other=ranges[first].ref,
        )
    return tuple(ranges)


def _find_overlapping_pair(
    bounds: Sequence[tuple[int, int, int, int]],
) -> tuple[int, int] | None:
    """Return the indexes of two overlapping rectangles, or ``None``.

    Rectangles are swept in ascending first-row order. Each rectangle still
    active started no later than the candidate and ends no earlier, so every
    active rectangle spans the candidate's first row. While no overlap has
    been found, rectangles sharing a row cannot share a column, so the active
    column spans are disjoint and sorted: only the candidate's two column
    neighbours can overlap it. That keeps the check logarithmic per rectangle
    instead of comparing every pair.
    """
    order = sorted(range(len(bounds)), key=lambda index: bounds[index][0])
    active: list[tuple[int, int, int]] = []  # (min_col, max_col, index), by min_col
    expiring: list[tuple[int, int]] = []  # (max_row, index) min-heap
    for index in order:
        min_row, min_col, max_row, max_col = bounds[index]
        while expiring and expiring[0][0] < min_row:
            _, expired = heapq.heappop(expiring)
            gone = bisect.bisect_left(active, (bounds[expired][1], bounds[expired][3], expired))
            del active[gone]
        # Column spans start at 1, so this lands before any entry sharing
        # min_col and the successor check below still sees it.
        position = bisect.bisect_left(active, (min_col, 0, 0))
        if position < len(active) and active[position][0] <= max_col:
            return index, active[position][2]
        if position > 0 and active[position - 1][1] >= min_col:
            return index, active[position - 1][2]
        active.insert(position, (min_col, max_col, index))
        heapq.heappush(expiring, (max_row, index))
    return None


def _assign_merge_refs(cells: Mapping[str, Cell], merges: Sequence[MergeRange]) -> dict[str, Cell]:
    """Attach each cell's merge ref, sweeping rows rather than rescanning.

    Merges are validated as non-overlapping first, so the merges spanning one
    row have disjoint, sorted column spans and the containing merge is found
    by bisection.
    """
    order = sorted(range(len(merges)), key=lambda index: merges[index].min_row)
    active: list[tuple[int, int, str]] = []  # (min_col, max_col, ref), by min_col
    expiring: list[tuple[int, int, int, str]] = []  # (max_row, *active entry) min-heap
    entered = 0
    assigned: dict[str, Cell] = {}
    for address, cell in sorted(cells.items(), key=lambda item: (item[1].row, item[1].column)):
        # Drop finished merges before admitting new ones, so everything in
        # ``active`` spans this row and their column spans stay disjoint.
        while expiring and expiring[0][0] < cell.row:
            _, min_col, max_col, ref = heapq.heappop(expiring)
            del active[bisect.bisect_left(active, (min_col, max_col, ref))]
        while entered < len(order) and merges[order[entered]].min_row <= cell.row:
            merge = merges[order[entered]]
            entered += 1
            if merge.max_row < cell.row:
                continue  # ended before this row, and rows only increase
            bisect.insort(active, (merge.min_col, merge.max_col, merge.ref))
            heapq.heappush(expiring, (merge.max_row, merge.min_col, merge.max_col, merge.ref))
        position = bisect.bisect_right(active, (cell.column, _MAX_COL + 1, ""))
        merge_ref: str | None = None
        if position > 0 and active[position - 1][1] >= cell.column:
            merge_ref = active[position - 1][2]
        assigned[address] = replace(cell, merge_ref=merge_ref)
    return assigned


def _parse_rows(
    sheet_data: ET.Element,
    part: str,
    styles: Sequence[StyleInfo],
    shared_strings: Sequence[str],
    cells: dict[str, Cell],
    worksheet_space: str,
    shared_formulas: dict[int, SharedFormula],
    followers: list[tuple[str, int]],
) -> Refusal | None:
    data_space = inherited_space(sheet_data, worksheet_space)
    prev_row = 0
    for row_elem in sheet_data.findall(_ROW_TAG):
        raw_row = row_elem.attrib.get("r")
        if raw_row is None:
            row_num = prev_row + 1
        else:
            parsed_row = _parse_row_number(raw_row)
            if parsed_row is None:
                return _invalid_part(part, "invalid row number", row=raw_row)
            row_num = parsed_row
        if row_num <= prev_row:
            return _invalid_part(part, "rows are not strictly increasing", row=row_num)
        if row_num > _MAX_ROW:
            return _invalid_part(part, "row exceeds Excel limit", row=row_num)
        prev_row = row_num

        prev_col = 0
        row_space = inherited_space(row_elem, data_space)
        for cell_elem in row_elem.findall(_CELL_TAG):
            cell = _parse_cell(
                cell_elem,
                part,
                row_num,
                prev_col,
                styles,
                shared_strings,
                row_space,
                shared_formulas,
                followers,
            )
            if isinstance(cell, Refusal):
                return cell
            if cell.address in cells:
                return _invalid_part(part, "duplicate cell address", address=cell.address)
            if cell.column <= prev_col:
                return _invalid_part(
                    part, "columns are not strictly increasing", address=cell.address
                )
            prev_col = cell.column
            cells[cell.address] = cell
    return None


def _parse_cell(
    cell_elem: ET.Element,
    part: str,
    row_num: int,
    prev_col: int,
    styles: Sequence[StyleInfo],
    shared_strings: Sequence[str],
    row_space: str,
    shared_formulas: dict[int, SharedFormula],
    followers: list[tuple[str, int]],
) -> Cell | Refusal:
    raw_ref = cell_elem.attrib.get("r")
    if raw_ref is None:
        column = prev_col + 1
        if column > _MAX_COL:
            return _invalid_part(part, "column exceeds Excel limit", row=row_num)
        address = f"{_column_letter(column)}{row_num}"
        row = row_num
    else:
        parsed = _parse_address(raw_ref)
        if parsed is None:
            return _invalid_part(part, "invalid cell address", address=raw_ref)
        row, column = parsed
        address = f"{_column_letter(column)}{row}"
        if row != row_num:
            return _invalid_part(
                part,
                "cell row does not match enclosing row",
                address=raw_ref,
                row=row_num,
            )

    style_index = _style_index(cell_elem)
    if style_index is None:
        return _invalid_part(
            part, "invalid style index", address=address, s=cell_elem.attrib.get("s")
        )
    if style_index >= len(styles):
        return _invalid_part(
            part, "style index is not in the style table", address=address, s=style_index
        )

    stored_type = cell_elem.attrib.get("t", "n")
    if stored_type not in _KNOWN_CELL_TYPES:
        return _invalid_part(part, "unknown cell type", address=address, stored_type=stored_type)

    payload = _payload_children(cell_elem, part, address)
    if isinstance(payload, Refusal):
        return payload
    v_elem, is_elem, f_elem = payload

    # The formula is read first: whether one exists decides how an empty
    # cache reads, and a malformed formula refuses either way.
    formula: Formula | None = None
    if f_elem is not None:
        parsed_formula = _parse_formula(f_elem, part, address, shared_formulas, followers)
        if isinstance(parsed_formula, Refusal):
            return parsed_formula
        formula = parsed_formula

    decoded = _decode_stored_value(
        stored_type,
        v_elem,
        is_elem,
        shared_strings,
        part,
        address,
        inherited_space(cell_elem, row_space),
        has_formula=formula is not None,
    )
    if isinstance(decoded, Refusal):
        return decoded
    value_kind, raw, stored_value = decoded

    return Cell(
        address=address,
        row=row,
        column=column,
        stored_type=stored_type,
        value_kind=value_kind,
        raw=raw,
        stored_value=stored_value,
        style_index=style_index,
        style=styles[style_index],
        formula=formula,
        merge_ref=None,
    )


def _payload_children(
    cell_elem: ET.Element, part: str, address: str
) -> tuple[ET.Element | None, ET.Element | None, ET.Element | None] | Refusal:
    v_elem = _single_child(cell_elem, _V_TAG, part, address)
    if isinstance(v_elem, Refusal):
        return v_elem
    is_elem = _single_child(cell_elem, _IS_TAG, part, address)
    if isinstance(is_elem, Refusal):
        return is_elem
    f_elem = _single_child(cell_elem, _F_TAG, part, address)
    if isinstance(f_elem, Refusal):
        return f_elem
    for elem, tag in ((v_elem, _V_TAG), (f_elem, _F_TAG)):
        nested = _reject_nested_markup(elem, tag, part, address)
        if nested is not None:
            return nested
    return v_elem, is_elem, f_elem


def _single_child(
    parent: ET.Element, tag: str, part: str, address: str
) -> ET.Element | Refusal | None:
    found = [child for child in parent if child.tag == tag]
    if len(found) > 1:
        local = tag.rsplit("}", 1)[-1]
        return _invalid_part(part, f"duplicate {local} elements", address=address)
    if found:
        return found[0]
    return None


def _reject_nested_markup(
    elem: ET.Element | None, tag: str, part: str, address: str
) -> Refusal | None:
    """Refuse a child element inside simple content such as ``<v>`` or ``<f>``.

    Reading ``.text`` alone would silently drop everything after the child,
    turning malformed source into a plausible partial value.
    """
    if elem is None or not len(elem):
        return None
    return _invalid_part(
        part, NESTED_MARKUP_REASON, address=address, element=tag.rsplit("}", 1)[-1]
    )


def _decode_stored_value(
    stored_type: str,
    v_elem: ET.Element | None,
    is_elem: ET.Element | None,
    shared_strings: Sequence[str],
    part: str,
    address: str,
    cell_space: str,
    *,
    has_formula: bool,
) -> tuple[str, str | None, float | bool | str | None] | Refusal:
    if stored_type == "inlineStr":
        if v_elem is not None:
            return _invalid_part(
                part, "inline string cell has a competing v payload", address=address
            )
        if is_elem is None:
            return "blank", None, None
        text = decode_rich_text(is_elem, cell_space)
        if isinstance(text, TextError):
            return _invalid_part(part, text.reason, address=address)
        return "string", None, text

    if is_elem is not None:
        return _invalid_part(
            part,
            "cell has an inline string payload that its type does not use",
            address=address,
        )

    raw = None if v_elem is None else (v_elem.text or "")
    if raw is None:
        return "blank", None, None
    if raw == "" and has_formula and stored_type != "str":
        # "There is a formula but no saved answer" is not a corrupt cell.
        # The empty raw keeps an explicitly empty cache distinct from an
        # absent <v> and from a saved zero. An empty string is a real
        # result for t="str", so that type keeps it.
        return "blank", "", None

    if stored_type == "n":
        return _decode_number(raw, part, address)
    if stored_type == "b":
        return _decode_boolean(raw, part, address)
    if stored_type == "e":
        return _decode_error(raw, part, address)
    if stored_type == "s":
        return _decode_shared_string(raw, shared_strings, part, address)
    if stored_type == "str":
        decoded = decode_excel_escapes(raw)
        if decoded is None:
            return _invalid_part(part, SURROGATE_REASON, address=address)
        return "string", raw, decoded
    return _decode_iso_date(raw, part, address)


def _decode_number(
    raw: str, part: str, address: str
) -> tuple[str, str | None, float | bool | str | None] | Refusal:
    if raw == "" or _NUMBER_RE.fullmatch(raw) is None:
        return _invalid_part(part, "invalid numeric cell value", address=address, raw=raw)
    try:
        value = float(raw)
    except ValueError:
        return _invalid_part(part, "invalid numeric cell value", address=address, raw=raw)
    if not math.isfinite(value):
        return _invalid_part(part, "non-finite numeric cell value", address=address, raw=raw)
    return "number", raw, value


def _decode_boolean(
    raw: str, part: str, address: str
) -> tuple[str, str | None, float | bool | str | None] | Refusal:
    if raw == "1":
        return "boolean", raw, True
    if raw == "0":
        return "boolean", raw, False
    return _invalid_part(part, "invalid boolean cell value", address=address, raw=raw)


def _decode_error(
    raw: str, part: str, address: str
) -> tuple[str, str | None, float | bool | str | None] | Refusal:
    if raw.startswith("#") and len(raw) > 1:
        return "error", raw, raw
    return _invalid_part(part, "invalid error cell value", address=address, raw=raw)


def _decode_shared_string(
    raw: str,
    shared_strings: Sequence[str],
    part: str,
    address: str,
) -> tuple[str, str | None, float | bool | str | None] | Refusal:
    index = _parse_unsigned_int(raw)
    if index is None:
        return _invalid_part(part, "invalid shared string index", address=address, raw=raw)
    if index >= len(shared_strings):
        return _invalid_part(
            part,
            "shared string index is out of range",
            address=address,
            string_index=index,
        )
    return "string", raw, shared_strings[index]


def _decode_iso_date(
    raw: str, part: str, address: str
) -> tuple[str, str | None, float | bool | str | None] | Refusal:
    if raw != "" and _is_supported_iso_date_text(raw):
        return "iso_date", raw, raw
    return _invalid_part(part, "invalid ISO date cell value", address=address, raw=raw)


def _is_supported_iso_date_text(text: str) -> bool:
    if "T" in text:
        date_part, sep, time_part = text.partition("T")
        return bool(sep) and _is_valid_date(date_part) and _is_valid_time(time_part)
    if "-" in text:
        return _is_valid_date(text)
    return _is_valid_time(text)


def _is_valid_date(text: str) -> bool:
    match = _DATE_RE.fullmatch(text)
    if match is None:
        return False
    year, month, day = (int(match.group(1)), int(match.group(2)), int(match.group(3)))
    if month < 1 or month > 12:
        return False
    max_day = _MONTH_DAYS[month]
    if month == 2 and _is_gregorian_leap(year):
        max_day = 29
    return 1 <= day <= max_day


def _is_gregorian_leap(year: int) -> bool:
    return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)


def _is_valid_time(text: str) -> bool:
    match = _TIME_RE.fullmatch(text)
    if match is None:
        return False
    hour, minute, second = (int(match.group(1)), int(match.group(2)), int(match.group(3)))
    return hour <= 23 and minute <= 59 and second <= 59


def _parse_formula(
    f_elem: ET.Element,
    part: str,
    address: str,
    shared_formulas: dict[int, SharedFormula],
    followers: list[tuple[str, int]],
) -> Formula | Refusal:
    raw_kind = f_elem.attrib.get("t")
    if raw_kind not in _FORMULA_KINDS:
        return _invalid_part(part, "unknown formula kind", address=address, t=raw_kind)
    kind = _FORMULA_KINDS[raw_kind]
    text = f_elem.text
    has_text = text is not None and text.strip() != ""
    if kind in _FORMULA_KINDS_REQUIRING_TEXT and not has_text:
        return _invalid_part(part, f"{kind} formula has no expression", address=address)
    if not has_text:
        # A shared follower and a data table carry no expression of their
        # own. Whitespace is not one either, so text stays absent rather
        # than looking like an expression to a consumer testing presence.
        text = None
    ref = f_elem.attrib.get("ref")

    ca_raw = f_elem.attrib.get("ca")
    if ca_raw is None:
        calculate = False
    else:
        parsed_ca = _xml_boolean(ca_raw)
        if parsed_ca is None:
            return _invalid_part(part, "invalid ca attribute", address=address, ca=ca_raw)
        calculate = parsed_ca

    shared_index: int | None = None
    if kind == "shared":
        raw_si = f_elem.attrib.get("si")
        if raw_si is None:
            return _invalid_part(part, "shared formula missing si", address=address)
        shared_index = _parse_unsigned_int(raw_si)
        if shared_index is None:
            return _invalid_part(part, "invalid shared formula index", address=address, si=raw_si)
        if has_text:
            if ref is None:
                return _invalid_part(part, "shared formula master missing ref", address=address)
            bounds = _parse_range_ref(ref)
            if bounds is None:
                return _invalid_part(part, "invalid shared formula ref", address=address, ref=ref)
            cell_pos = _parse_address(address)
            if cell_pos is None or not _range_contains(bounds, cell_pos[0], cell_pos[1]):
                return _invalid_part(
                    part,
                    "shared formula master is outside its ref",
                    address=address,
                    ref=ref,
                )
            if shared_index in shared_formulas:
                return _invalid_part(
                    part,
                    "duplicate shared formula master",
                    address=address,
                    si=shared_index,
                )
            shared_formulas[shared_index] = SharedFormula(
                master=address, ref=ref, text=text if text is not None else ""
            )
        else:
            followers.append((address, shared_index))
    elif kind in {"array", "data_table"}:
        if ref is None:
            return _invalid_part(part, f"{kind} formula missing ref", address=address)
        bounds = _parse_range_ref(ref)
        if bounds is None:
            return _invalid_part(part, f"invalid {kind} formula ref", address=address, ref=ref)
        cell_pos = _parse_address(address)
        if cell_pos is None or not _range_contains(bounds, cell_pos[0], cell_pos[1]):
            return _invalid_part(
                part,
                f"{kind} formula is outside its ref",
                address=address,
                ref=ref,
            )

    return Formula(
        text=text,
        kind=kind,
        shared_index=shared_index,
        ref=ref,
        calculate_on_next_recalc=calculate,
    )


def _check_shared_followers(
    part: str,
    cells: Mapping[str, Cell],
    shared_formulas: Mapping[int, SharedFormula],
    followers: Sequence[tuple[str, int]],
) -> Refusal | None:
    for address, si in followers:
        master = shared_formulas.get(si)
        if master is None:
            return _invalid_part(
                part, "shared formula follower has no master", address=address, si=si
            )
        bounds = _parse_range_ref(master.ref)
        cell = cells[address]
        if bounds is None or not _range_contains(bounds, cell.row, cell.column):
            return _invalid_part(
                part,
                "shared formula follower is outside the master range",
                address=address,
                si=si,
                ref=master.ref,
            )
    return None


def _parse_range_ref(ref: str) -> tuple[int, int, int, int] | None:
    if ":" in ref:
        start, _sep, end = ref.partition(":")
        if ":" in end:
            return None
        first = _parse_address(start)
        last = _parse_address(end)
        if first is None or last is None:
            return None
        row1, col1 = first
        row2, col2 = last
        return (min(row1, row2), min(col1, col2), max(row1, row2), max(col1, col2))
    parsed = _parse_address(ref)
    if parsed is None:
        return None
    row, column = parsed
    return (row, column, row, column)


def _range_contains(bounds: tuple[int, int, int, int], row: int, column: int) -> bool:
    min_row, min_col, max_row, max_col = bounds
    return min_row <= row <= max_row and min_col <= column <= max_col


def _xml_boolean(raw: str) -> bool | None:
    if raw in _XML_TRUE:
        return True
    if raw in _XML_FALSE:
        return False
    return None


def _style_index(cell_elem: ET.Element) -> int | None:
    raw = cell_elem.attrib.get("s")
    if raw is None:
        return 0
    return _parse_unsigned_int(raw)


def _used_range(cells: Mapping[str, Cell]) -> str | None:
    if not cells:
        return None
    min_row = min(cell.row for cell in cells.values())
    max_row = max(cell.row for cell in cells.values())
    min_col = min(cell.column for cell in cells.values())
    max_col = max(cell.column for cell in cells.values())
    start = f"{_column_letter(min_col)}{min_row}"
    end = f"{_column_letter(max_col)}{max_row}"
    return start if start == end else f"{start}:{end}"


def _parse_address(address: str) -> tuple[int, int] | None:
    match = _ADDRESS_RE.fullmatch(address)
    if match is None:
        return None
    column = _column_index(match.group(1))
    if column is None or column > _MAX_COL:
        return None
    row = _parse_unsigned_int(match.group(2))
    if row is None or row < 1 or row > _MAX_ROW:
        return None
    return row, column


def _parse_row_number(raw: str) -> int | None:
    value = _parse_unsigned_int(raw)
    if value is None or value < 1:
        return None
    return value


def _column_index(letters: str) -> int | None:
    total = 0
    for char in letters:
        total = total * 26 + (ord(char) - 64)
        if total > _MAX_COL:
            return None
    return total


def _column_letter(index: int) -> str:
    chars: list[str] = []
    remaining = index
    while remaining > 0:
        remaining, rem = divmod(remaining - 1, 26)
        chars.append(chr(65 + rem))
    return "".join(reversed(chars))


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
