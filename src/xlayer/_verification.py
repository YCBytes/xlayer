"""Fresh reparse and independent byte/structural comparison before publication.

The byte oracle is deliberately separate from the writer: a small lexical
token walk masks type attributes and scalar payloads, not complete cells. All
other bytes (including target style/spacing) must match. Raw expectations come
from the preview's plans, never from the patcher's claimed success.
"""

from __future__ import annotations

import re
from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass, replace
from pathlib import Path
from typing import TYPE_CHECKING
from xml.etree import ElementTree as ET

from xlayer._approval import refusal
from xlayer._edit_validation import EditPlan
from xlayer._errors import Refusal, WorkbookOpenError

if TYPE_CHECKING:
    from xlayer._workbook import Workbook

_ATTRIBUTE = re.compile(rb"([^\s=<>/]+)\s*=\s*(['\"])(.*?)\2", re.DOTALL)
_TOKEN = re.compile(rb"<!--.*?-->|<\?.*?\?>|<!\[CDATA\[.*?\]\]>|<[^>]*(?:>|$)", re.DOTALL)
_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _masked_sheet(
    raw: bytes, targets: set[str], original_types: Mapping[str, bool], *, output: bool = False
) -> bytes | None:
    """Independent verified-XML token walk, never calls cell_spans/patch helpers."""
    # Attributes may contain >; recover the complete tag with a quote-aware
    # local walk. Comments/PI/CDATA remain exact outside scalar payloads.
    masked: list[bytes] = []
    position = 0
    row, column = 0, 0
    current_target = False
    payload_depth = 0
    stack: list[str] = []
    for match in _TOKEN.finditer(raw):
        if match.start() < position:
            continue
        if not payload_depth:
            masked.append(raw[position : match.start()])
        start = match.start()
        end = match.end()
        token = raw[start:end]
        if token.startswith((b"<!--", b"<?", b"<![CDATA[")):
            if not payload_depth:
                masked.append(token)
            position = end
            continue
        quote = 0
        for offset in range(start + 1, len(raw)):
            char = raw[offset]
            if quote:
                if char == quote:
                    quote = 0
            elif char in (34, 39):
                quote = char
            elif char == 62:
                end = offset + 1
                break
        token = raw[start:end]
        closing = token.startswith(b"</")
        name = re.split(rb"[\s/>]", token[2:] if closing else token[1:], maxsplit=1)[0].decode(
            "utf-8"
        )
        local = name.rsplit(":", 1)[-1]
        self_closing = token.rstrip().endswith(b"/>")
        if closing:
            if payload_depth:
                payload_depth -= 1
            else:
                masked.append(token)
            if stack:
                stack.pop()
            if local == "c":
                current_target = False
        else:
            if local in {"row", "c"}:
                attributes = list(_ATTRIBUTE.finditer(token))
                coordinate_attr = next((attr for attr in attributes if attr[1] == b"r"), None)
                # XML attribute spelling is not its value: numeric references
                # and XML whitespace normalization must agree with the reader.
                coordinate_value = (
                    ET.fromstring(  # noqa: S314 - verified XML, DTD preflight already applied
                        b"<carrier value="
                        + coordinate_attr[2]
                        + coordinate_attr[3]
                        + coordinate_attr[2]
                        + b"/>"
                    ).attrib["value"]
                    if coordinate_attr is not None
                    else None
                )
                if local == "row" and stack and stack[-1] == "sheetData":
                    row = int(coordinate_value if coordinate_value is not None else str(row + 1))
                    column = 0
                if local == "c" and stack and stack[-1] == "row":
                    coordinate = coordinate_value
                    if coordinate:
                        coordinate_match = re.fullmatch(r"([A-Z]+)([0-9]+)", coordinate)
                        if coordinate_match is None:
                            return None
                        column = 0
                        for letter in coordinate_match[1]:
                            column = column * 26 + ord(letter) - 64
                    else:
                        column += 1
                        number, letters = column, ""
                        while number:
                            number, digit = divmod(number - 1, 26)
                            letters = chr(65 + digit) + letters
                        coordinate = f"{letters}{row}"
                    current_target = coordinate in targets
                    if current_target:
                        type_attr = next((attr for attr in attributes if attr[1] == b"t"), None)
                        if original_types[coordinate]:
                            if type_attr is None:
                                return None
                            # Existing type attributes retain their quote and
                            # delimiter spelling; only their value may change.
                            token = token[: type_attr.start(3)] + token[type_attr.end(3) :]
                        elif output:
                            if (
                                type_attr is None
                                or token[type_attr.start() - 1 : type_attr.start()] != b" "
                            ):
                                return None
                            # The writer inserts one delimiter. Keep all older
                            # whitespace, even immediately before that space.
                            token = token[: type_attr.start() - 1] + token[type_attr.end() :]
                        elif type_attr is not None:
                            return None
                        if self_closing:
                            token = token[:-2] + b"></" + name.encode() + b">"
            if payload_depth:
                if not self_closing:
                    payload_depth += 1
            elif current_target and local in {"v", "is"} and stack and stack[-1] == "c":
                if not self_closing:
                    payload_depth = 1
            else:
                masked.append(token)
            if not self_closing:
                stack.append(local)
        position = end
    if payload_depth:
        return None
    masked.append(raw[position:])
    return b"".join(masked)


def _target_cells(raw: bytes) -> Iterator[tuple[str, ET.Element]]:
    root = ET.fromstring(raw)  # noqa: S314 - archive preflight already applied
    row, column = 0, 0
    for row_elem in root.findall(f"{_MAIN}sheetData/{_MAIN}row"):
        row = int(row_elem.attrib.get("r", str(row + 1)))
        column = 0
        for elem in row_elem.findall(_MAIN + "c"):
            coordinate = elem.attrib.get("r")
            if coordinate is None:
                column += 1
                number, letters = column, ""
                while number:
                    number, digit = divmod(number - 1, 26)
                    letters = chr(65 + digit) + letters
                coordinate = f"{letters}{row}"
            else:
                column = 0
                for char in coordinate.rstrip("0123456789"):
                    column = column * 26 + ord(char) - 64
            yield coordinate, elem


def _target_structure(raw: bytes, targets: set[str]) -> dict[str, object] | None:
    result: dict[str, object] = {}
    for coordinate, elem in _target_cells(raw):
        if coordinate in targets:
            # The new text representation must be plain inline text, not
            # arbitrary rich markup that happens to decode to the same value.
            for child in elem:
                if child.tag not in {_MAIN + "v", _MAIN + "is"}:
                    return None
                allowed_attrs = (
                    {"{http://www.w3.org/XML/1998/namespace}space"}
                    if child.tag == _MAIN + "is"
                    else set()
                )
                if set(child.attrib) - allowed_attrs:
                    return None
                if child.tag == _MAIN + "is" and (
                    len(child) != 1
                    or child[0].tag != _MAIN + "t"
                    or len(child[0])
                    or set(child[0].attrib) - {"{http://www.w3.org/XML/1998/namespace}space"}
                ):
                    return None
            result[coordinate] = {key: value for key, value in elem.attrib.items() if key != "t"}
    return result if set(result) == targets else None


@dataclass(frozen=True)
class VerificationReport:
    output_fingerprint: str
    changed_parts: tuple[str, ...]
    checks: tuple[str, ...]

    def to_dict(self) -> dict[str, object]:
        return {
            "all_passed": True,
            "checks": list(self.checks),
            "changed_parts": list(self.changed_parts),
        }


def verify_output(
    book: Workbook, path: Path, plans: Sequence[EditPlan], part: str
) -> VerificationReport | Refusal:
    from xlayer._workbook import Workbook

    original = book._archive
    if original is None:
        raise RuntimeError("verification requires the open source snapshot")
    try:
        output = Workbook.open(path, limits=original._limits)
    except WorkbookOpenError as exc:
        return refusal(
            "verification_failed",
            "fresh output failed package validation",
            cause=exc.refusal.to_dict(),
        )
    with output:
        fresh = output._archive
        if fresh is None:
            raise RuntimeError("fresh output session lost its archive")
        if (
            original.part_names() != fresh.part_names()
            or book.registry != output.registry
            or book._styles != output._styles
            or book._shared_strings != output._shared_strings
        ):
            return refusal(
                "verification_failed", "package inventory or shared infrastructure changed"
            )
        if [i.filename for i in original._zip.infolist()] != [
            i.filename for i in fresh._zip.infolist()
        ] or original._zip.comment != fresh._zip.comment:
            return refusal("verification_failed", "ZIP inventory or archive comment changed")
        for before, after in zip(original._zip.infolist(), fresh._zip.infolist(), strict=True):
            for attr in (
                "date_time",
                "compress_type",
                "comment",
                "extra",
                "create_system",
                "create_version",
                "extract_version",
                "internal_attr",
                "external_attr",
            ):
                if getattr(before, attr) != getattr(after, attr):
                    return refusal(
                        "verification_failed",
                        "ZIP metadata changed",
                        entry=before.filename,
                        field=attr,
                    )
        changed: list[str] = []
        for name in original.part_names():
            old, new = original.read_part(name), fresh.read_part(name)
            if isinstance(old, Refusal) or isinstance(new, Refusal):
                return refusal("verification_failed", "part failed integrity checks", part=name)
            if old != new:
                changed.append(name)
                if name != part:
                    return refusal("verification_failed", "undeclared part changed", part=name)
        if changed != [part]:
            return refusal(
                "verification_failed", "actual touched parts differ from declared one-part change"
            )
        old_bytes, new_bytes = original.read_part(part), fresh.read_part(part)
        if not isinstance(old_bytes, bytes) or not isinstance(new_bytes, bytes):
            return refusal("verification_failed", "target part cannot be compared")
        effective = {plan.edit.cell for plan in plans if plan.effective}
        original_types = {
            coordinate: "t" in elem.attrib
            for coordinate, elem in _target_cells(old_bytes)
            if coordinate in effective
        }
        old_masked, new_masked = (
            _masked_sheet(old_bytes, effective, original_types),
            _masked_sheet(new_bytes, effective, original_types, output=True),
        )
        old_structure = _target_structure(old_bytes, effective)
        new_structure = _target_structure(new_bytes, effective)
        if (
            old_masked is None
            or new_masked is None
            or old_masked != new_masked
            or old_structure is None
            or new_structure is None
            or old_structure != new_structure
        ):
            return refusal(
                "verification_failed", "bytes outside approved type/payload surface changed"
            )
        plan_by_sheet = {(plan.edit.sheet, plan.edit.cell): plan for plan in plans}
        for entry in book.registry.sheets:
            if entry.kind != "worksheet":
                continue
            old_sheet, new_sheet = book.read_sheet(entry.name), output.read_sheet(entry.name)
            if isinstance(old_sheet, Refusal) or isinstance(new_sheet, Refusal):
                return refusal(
                    "verification_failed", "all supported worksheets must reparse", sheet=entry.name
                )
            if (
                old_sheet.cells.keys() != new_sheet.cells.keys()
                or old_sheet.merges != new_sheet.merges
                or old_sheet.shared_formulas != new_sheet.shared_formulas
                or old_sheet.declared_dimension != new_sheet.declared_dimension
                or old_sheet.used_range != new_sheet.used_range
                or old_sheet.other_elements != new_sheet.other_elements
            ):
                return refusal(
                    "verification_failed", "worksheet structure changed", sheet=entry.name
                )
            for address, old_cell in old_sheet.cells.items():
                new_cell = new_sheet.cells[address]
                plan = plan_by_sheet.get((entry.name, address))
                if plan is None or not plan.effective:
                    if new_cell != old_cell:
                        return refusal(
                            "verification_failed",
                            "non-target or no-op cell changed",
                            address=address,
                        )
                    continue
                if (new_cell.stored_type, new_cell.raw, new_cell.stored_value) != (
                    plan.expected_type,
                    plan.expected_raw,
                    plan.expected_value,
                ):
                    return refusal(
                        "verification_failed",
                        "target payload differs from preview",
                        address=address,
                    )
                normalized = replace(
                    new_cell,
                    stored_type=old_cell.stored_type,
                    raw=old_cell.raw,
                    stored_value=old_cell.stored_value,
                    value_kind=old_cell.value_kind,
                )
                if normalized != old_cell:
                    return refusal(
                        "verification_failed", "target's non-value facts changed", address=address
                    )
        return VerificationReport(
            output.source_fingerprint,
            tuple(changed),
            (
                "package_and_zip_integrity",
                "inventory_and_metadata",
                "shared_infrastructure",
                "all_supported_worksheets_reparsed",
                "complete_target_ledger",
                "non_target_facts",
                "untouched_part_byte_identity",
                "independent_allowed_byte_surface",
                "one_part_invariant",
            ),
        )
