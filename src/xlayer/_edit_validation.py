"""Positive write-safety proof for ordinary existing scalar cells, not a formatter."""

from __future__ import annotations

import posixpath
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
from typing import TYPE_CHECKING
from xml.etree import ElementTree as ET

from xlayer._approval import refusal
from xlayer._canonical import frozen_mapping
from xlayer._edits import SetValue
from xlayer._errors import Refusal
from xlayer._ooxml.package import PackageInfo, _parse_rels, parse_package
from xlayer._ooxml.patch import NS, XmlSpan, cell_spans
from xlayer._ooxml.sheet import Cell, Worksheet, _parse_range_ref, _range_contains
from xlayer._ooxml.styles import _parse_unsigned_int, classify_number_format

if TYPE_CHECKING:
    from xlayer._workbook import Workbook

_NS = "{" + NS + "}"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_SIGNATURE_BASE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/"
_SIGNATURE_TYPES = frozenset(
    {
        "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
        "application/vnd.openxmlformats-package.digital-signature-origin",
        "application/vnd.openxmlformats-package.digital-signature-certificate",
    }
)


@dataclass(frozen=True)
class EditPlan:
    edit: SetValue
    part: str
    before: Cell
    category: str
    style_proof: Mapping[str, object]
    expected_type: str
    expected_raw: str | None
    expected_value: float | bool | str | None
    effective: bool
    populated_text: bool

    def to_dict(self) -> dict[str, object]:
        return {
            "request": self.edit.to_dict(),
            "part": self.part,
            "target_category": self.category,
            "before": {
                "stored_type": self.before.stored_type,
                "raw": self.before.raw,
                "stored_value": self.before.stored_value,
                "value_kind": self.before.value_kind,
                "style_index": self.before.style_index,
            },
            "expected": {
                "stored_type": self.expected_type,
                "raw": self.expected_raw,
                "stored_value": self.expected_value,
            },
            "style_proof": self.style_proof,
            "effective": self.effective,
            "populated_text_change": self.populated_text,
        }


def date_serial(value: date, date1904: bool) -> int | None:
    epoch = date(1904, 1, 1) if date1904 else date(1900, 1, 1)
    if value < epoch:
        return None
    return (value - epoch).days + (0 if date1904 else 1 + int(value >= date(1900, 3, 1)))


def _source_date(raw: str | None, date1904: bool) -> date | None:
    if raw is None:
        return None
    try:
        number = Decimal(raw)
        maximum = date_serial(date.max, date1904)
        minimum = 0 if date1904 else 1
        if (
            maximum is None
            or not minimum <= number <= maximum
            or number != number.to_integral_value()
        ):
            return None
        integer = int(number)
        if not date1904 and integer == 60:
            return None
        offset = integer if date1904 else integer - 1 - int(integer > 60)
        return (date(1904, 1, 1) if date1904 else date(1900, 1, 1)) + timedelta(days=offset)
    except (InvalidOperation, ValueError, OverflowError):
        return None


@dataclass
class WriteContext:
    package: PackageInfo
    styles_root: ET.Element | None
    custom_formats: dict[int, str]
    roots: dict[str, ET.Element]
    spans: dict[str, dict[str, XmlSpan]]


def write_context(book: Workbook) -> WriteContext | Refusal:
    archive = book._archive
    if archive is None:
        raise RuntimeError("write context requires an open session")
    package = parse_package(archive)
    if isinstance(package, Refusal):
        return package
    if any(
        value in _SIGNATURE_TYPES
        for value in (*package.content_types.values(), *package.default_types.values())
    ) or any(
        rel.rel_type in {_SIGNATURE_BASE + kind for kind in ("origin", "signature", "certificate")}
        for rel in (*package.top_rels.values(), *package.workbook_rels.values())
    ):
        return refusal("signed_package", "editing would invalidate package signatures")
    styles_root = None
    custom: dict[int, str] = {}
    styles_part = package.resolve_workbook_rel("styles")
    if styles_part is not None:
        parsed = archive.parse_xml_part(styles_part)
        if isinstance(parsed, Refusal):
            return parsed
        styles_root = parsed
        for fmt in parsed.findall(f"{_NS}numFmts/{_NS}numFmt"):
            identifier = _parse_unsigned_int(fmt.attrib.get("numFmtId", ""))
            if identifier is not None:
                custom[identifier] = fmt.attrib["formatCode"]
    return WriteContext(package, styles_root, custom, {}, {})


def _format_guard(context: WriteContext, cell: Cell) -> tuple[str, Mapping[str, object]] | Refusal:
    if context.styles_root is None:
        return "non_temporal", frozen_mapping({"basis": "absent_styles_general"})
    xfs = context.styles_root.findall(f"{_NS}cellXfs/{_NS}xf")
    if not xfs:
        return "non_temporal", frozen_mapping({"basis": "absent_cell_xfs_general"})
    xf = xfs[cell.style_index]
    flag = xf.attrib.get("applyNumberFormat")
    if flag not in {None, "0", "1", "false", "true"}:
        return refusal("unsupported_target_style", "invalid number-format application flag")
    direct_id = _parse_unsigned_int(xf.attrib.get("numFmtId", "0"))
    if direct_id is None:
        return refusal("unsupported_target_style", "invalid direct number format")
    direct_kind = classify_number_format(direct_id, context.custom_formats.get(direct_id))
    base_id: int | None = None
    base_kind: str | None = None
    base_flag: str | None = None
    if "xfId" in xf.attrib:
        base_index = _parse_unsigned_int(xf.attrib["xfId"])
        bases = context.styles_root.findall(f"{_NS}cellStyleXfs/{_NS}xf")
        if base_index is None or base_index >= len(bases):
            return refusal("unsupported_target_style", "unresolved base format record")
        base = bases[base_index]
        base_flag = base.attrib.get("applyNumberFormat")
        base_id = _parse_unsigned_int(base.attrib.get("numFmtId", "0"))
        if base_id is None or base_flag not in {None, "0", "1", "false", "true"}:
            return refusal("unsupported_target_style", "invalid base format record")
        if base_flag in {"0", "false"} and base_id != 0:
            return refusal("unsupported_target_style", "disabled base number format is ambiguous")
        base_kind = classify_number_format(base_id, context.custom_formats.get(base_id))
    if flag in {"0", "false"} and base_kind is None:
        return refusal("unsupported_target_style", "disabled direct format has no established base")
    if flag not in {"1", "true"} and base_kind is not None and base_kind != direct_kind:
        return refusal(
            "unsupported_target_style", "direct and inherited temporal categories disagree"
        )
    proof = frozen_mapping(
        {
            "basis": "explicit_direct" if flag in {"1", "true"} else "established_agreement",
            "direct_num_fmt_id": direct_id,
            "base_num_fmt_id": base_id,
            "apply_number_format": flag,
            "base_apply_number_format": base_flag,
            "reader_temporal_kind": cell.style.temporal_kind,
            "temporal_kind": direct_kind,
        }
    )
    return direct_kind, proof


def _worksheet_guard(book: Workbook, context: WriteContext, part: str) -> Refusal | None:
    archive = book._archive
    if archive is None:
        raise RuntimeError("worksheet guard requires open session")
    if part not in context.roots:
        root = archive.parse_xml_part(part)
        if isinstance(root, Refusal):
            return root
        raw = archive.read_part(part)
        if isinstance(raw, Refusal):
            return raw
        spans = cell_spans(raw)
        if isinstance(spans, Refusal):
            return spans
        context.roots[part] = root
        context.spans[part] = spans
    root = context.roots[part]
    protection = root.find(_NS + "sheetProtection")
    if protection is not None and protection.attrib.get("sheet", "0") not in {"0", "false"}:
        return refusal("protected_worksheet", "worksheet protection is enabled or indeterminate")
    if any(child.tag in {_NS + "tableParts", _NS + "dataValidations"} for child in root):
        return refusal("unsupported_target_structure", "table or validation ownership is unmodeled")
    directory, filename = posixpath.split(part)
    rel_part = posixpath.join(directory, "_rels", filename + ".rels")
    if archive.has_part(rel_part):
        rel_root = archive.parse_xml_part(rel_part)
        if isinstance(rel_root, Refusal):
            return rel_root
        rels = _parse_rels(rel_root, rel_part)
        if isinstance(rels, Refusal):
            return rels
        pivot_type = (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
        )
        if any(rel.rel_type == pivot_type for rel in rels.values()):
            return refusal(
                "unsupported_target_structure", "PivotTable display ownership is unmodeled"
            )
    return None


def validate_edit(
    book: Workbook, edit: SetValue, *, context: WriteContext | None = None
) -> EditPlan | Refusal:
    sheet = book.read_sheet(edit.sheet)
    if isinstance(sheet, Refusal):
        return sheet
    if context is None:
        created = write_context(book)
        if isinstance(created, Refusal):
            return created
        context = created
    guard = _worksheet_guard(book, context, sheet.part_path)
    if guard is not None:
        return guard
    if sheet.merge_at(edit.cell) is not None:
        return refusal(
            "merged_cell", "merge roots and children are not writable", address=edit.cell
        )
    cell = sheet.cells.get(edit.cell)
    if cell is None:
        return refusal("missing_cell", "target has no stored cell element", address=edit.cell)
    if cell.formula is not None:
        return refusal("formula_cell", "formula cells are not scalar inputs", address=edit.cell)
    for other in sheet.cells.values():
        if (
            other.formula is not None
            and other.formula.kind in {"array", "data_table"}
            and other.formula.ref
        ):
            bounds = _parse_range_ref(other.formula.ref)
            if bounds is not None and _range_contains(bounds, cell.row, cell.column):
                return refusal(
                    "formula_range_target", "target belongs to a formula-owned output range"
                )
    span = context.spans[sheet.part_path].get(edit.cell)
    if span is None or span.special_markup:
        return refusal(
            "unsupported_target_structure", "target cannot be safely mapped to scalar bytes"
        )
    allowed_attrs = {"r", "s", "t", _XML_SPACE[1:]}
    if any(key not in allowed_attrs for key in span.attrs):
        return refusal("unsupported_target_structure", "target has unmodeled attributes")
    root = context.roots[sheet.part_path]
    cells = root.findall(f"{_NS}sheetData/{_NS}row/{_NS}c")
    source = cells[list(sheet.cells).index(edit.cell)]
    if any(child.tag not in {_NS + "v", _NS + "is"} for child in source):
        return refusal("unsupported_target_structure", "target has unmodeled children")
    rich = source.find(_NS + "is")
    if cell.stored_type == "s":
        archive = book._archive
        if archive is None:
            raise RuntimeError("shared string guard requires an open archive")
        part = context.package.resolve_workbook_rel("sharedStrings")
        if part is not None:
            strings = archive.parse_xml_part(part)
            if isinstance(strings, Refusal):
                return strings
            rich = strings.findall(_NS + "si")[int(cell.raw or "0")]
    if rich is not None and (
        any(child.tag != _NS + "t" for child in rich) or len(rich.findall(_NS + "t")) != 1
    ):
        return refusal("rich_text_cell", "rich or phonetic text cannot be replaced")
    if rich is not None and (
        set(rich.attrib) - {_XML_SPACE}
        or any(set(child.attrib) - {_XML_SPACE} or len(child) for child in rich)
    ):
        return refusal("unsupported_target_structure", "text payload has unmodeled metadata")
    if any(child.attrib for child in source if child.tag == _NS + "v"):
        return refusal("unsupported_target_structure", "scalar payload has unmodeled attributes")
    style = _format_guard(context, cell)
    if isinstance(style, Refusal):
        return style
    temporal, proof = style
    return _scalar_plan(book.registry.date1904, sheet, cell, edit, temporal, proof)


def _scalar_plan(
    date1904: bool,
    sheet: Worksheet,
    cell: Cell,
    edit: SetValue,
    temporal: str,
    proof: Mapping[str, object],
) -> EditPlan | Refusal:
    value = edit.value
    if temporal not in {"date", "non_temporal"}:
        return refusal(
            "unsupported_temporal_target", "time/date-time/unknown formats are not writable"
        )
    category = (
        "unestablished"
        if cell.value_kind == "blank"
        else {
            "number": "numeric",
            "boolean": "boolean",
            "string": "text",
            "iso_date": "date",
        }.get(cell.value_kind)
    )
    if category is None:
        return refusal("invalid_value", "error values are not supported input targets")
    old_date: date | None = None
    if temporal == "date":
        if category in {"text", "boolean"}:
            return refusal(
                "incompatible_date_style", "temporal style cannot establish text/bool as date"
            )
        if cell.value_kind == "number":
            old_date = _source_date(cell.raw, date1904)
            if old_date is None:
                return refusal(
                    "invalid_value", "previous date serial is not a supported integral date"
                )
        if cell.value_kind == "iso_date":
            try:
                old_date = date.fromisoformat(cell.raw or "")
            except ValueError:
                return refusal("unsupported_temporal_target", "ISO target is not date-only")
            if date_serial(old_date, date1904) is None:
                return refusal("invalid_value", "previous ISO date is outside the supported epoch")
        category = "date"
    elif cell.value_kind == "iso_date":
        return refusal("unsupported_temporal_target", "ISO date target requires date-only style")
    kinds: dict[str, set[type[object]]] = {
        "numeric": {int, float},
        "boolean": {bool},
        "text": {str},
        "date": {date},
        "unestablished": {int, float, bool, str},
    }
    if type(value) not in kinds[category]:
        return refusal(
            "type_mismatch",
            "requested scalar does not match established target category",
            target_category=category,
        )
    expected_type = "n"
    expected_raw: str | None
    expected_value: float | bool | str | None
    equal = False
    if type(value) is date:
        serial = date_serial(value, date1904)
        if serial is None:
            return refusal("invalid_value", "requested date precedes the workbook epoch")
        expected_type = "d" if cell.stored_type == "d" else "n"
        expected_raw = value.isoformat() if expected_type == "d" else str(serial)
        expected_value = expected_raw if expected_type == "d" else float(serial)
        equal = old_date == value
    elif type(value) is bool:
        expected_type, expected_raw, expected_value = "b", "1" if value else "0", value
        equal = cell.value_kind == "boolean" and cell.stored_value is value
    elif type(value) is str:
        expected_type, expected_raw, expected_value = "inlineStr", None, value
        equal = cell.value_kind == "string" and cell.stored_value == value
    else:
        expected_raw = str(value) if type(value) is int else repr(value)
        if not isinstance(value, (int, float)):
            raise TypeError("validated numeric scalar required")
        expected_value = float(value)
        if cell.value_kind == "number" and cell.raw is not None:
            try:
                old, new = Decimal(cell.raw), Decimal(expected_raw)
                if not old.is_finite():
                    return refusal(
                        "invalid_value", "previous raw number cannot be compared exactly"
                    )
                equal = old == new and (old != 0 or old.is_signed() == new.is_signed())
            except (InvalidOperation, ValueError, OverflowError):
                return refusal("invalid_value", "previous raw number cannot be compared exactly")
    if equal:
        expected_type, expected_raw, expected_value = cell.stored_type, cell.raw, cell.stored_value
    return EditPlan(
        edit,
        sheet.part_path,
        cell,
        category,
        proof,
        expected_type,
        expected_raw,
        expected_value,
        not equal,
        category == "text" and not equal,
    )
