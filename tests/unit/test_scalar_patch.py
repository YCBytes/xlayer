"""Raw XML expectations for the deliberately narrow mutation surface."""

from __future__ import annotations

import importlib.util
import zipfile
from datetime import date
from pathlib import Path
from typing import TYPE_CHECKING

import pytest

from tests.unit._dependency_cases import write_workbook_case
from xlayer._edits import SetValue
from xlayer._errors import Refusal
from xlayer._workbook import Workbook

if TYPE_CHECKING:
    from xlayer._edit_validation import EditPlan


def plan(
    tmp_path: Path, xml: str, value: object, *, styles: str = "", date1904: bool = False
) -> EditPlan | Refusal:
    assert importlib.util.find_spec("xlayer._edit_validation") is not None, (
        "write validation missing"
    )
    from xlayer._edit_validation import validate_edit

    path = write_workbook_case(
        tmp_path / "input.xlsx", {"Inputs": f"<sheetData><row>{xml}</row></sheetData>"}
    )
    if styles or date1904:
        with zipfile.ZipFile(path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        if styles:
            parts["xl/styles.xml"] = styles.encode()
            parts["xl/_rels/workbook.xml.rels"] = parts["xl/_rels/workbook.xml.rels"].replace(
                b"</Relationships>",
                b'<Relationship Id="st" '
                b'Type="http://schemas.openxmlformats.org/'
                b'officeDocument/2006/relationships/styles" '
                b'Target="styles.xml"/></Relationships>',
            )
        if date1904:
            parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(
                b"<sheets>", b'<workbookPr date1904="1"/><sheets>'
            )
        with zipfile.ZipFile(path, "w") as archive:
            for name, data in parts.items():
                archive.writestr(name, data)
    with Workbook.open(path) as book:
        return validate_edit(book, SetValue("Inputs", "A1", value))  # type: ignore[arg-type]


def style_xml(xfs: str, base: str = "") -> str:
    return (
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + base
        + f"<cellXfs>{xfs}</cellXfs></styleSheet>"
    )


@pytest.mark.parametrize(
    ("xml", "value", "category", "kind", "raw", "effective"),
    [
        ('<c r="A1"><v>1</v></c>', 2, "numeric", "n", "2", True),
        ('<c r="A1"><v>1e0</v></c>', 1.0, "numeric", "n", "1e0", False),
        ('<c r="A1"><v>0</v></c>', -0.0, "numeric", "n", "-0.0", True),
        (
            '<c r="A1"><v>9007199254740993</v></c>',
            9007199254740992.0,
            "numeric",
            "n",
            "9007199254740992.0",
            True,
        ),
        ('<c r="A1" t="b"><v>0</v></c>', True, "boolean", "b", "1", True),
        ('<c r="A1"/>', "", "unestablished", "inlineStr", None, True),
        ('<c r="A1"/>', 2, "unestablished", "n", "2", True),
        ('<c r="A1"/>', False, "unestablished", "b", "0", True),
        (
            '<c r="A1" t="inlineStr"><is><t>old</t></is></c>',
            "=1+1",
            "text",
            "inlineStr",
            None,
            True,
        ),
    ],
)
def test_scalar_raw_plan(
    tmp_path: Path,
    xml: str,
    value: object,
    category: str,
    kind: str,
    raw: str | None,
    effective: bool,
) -> None:
    result = plan(tmp_path, xml, value)
    assert not isinstance(result, Refusal)
    assert (result.category, result.expected_type, result.expected_raw, result.effective) == (
        category,
        kind,
        raw,
        effective,
    )


@pytest.mark.parametrize(
    ("xml", "value", "code"),
    [
        ('<c r="A1"><v>1</v></c>', True, "type_mismatch"),
        ('<c r="A1" t="b"><v>1</v></c>', 1, "type_mismatch"),
        ('<c r="A1"><f>1+1</f><v>2</v></c>', 3, "formula_cell"),
        ('<c r="A1" t="e"><v>#REF!</v></c>', 3, "invalid_value"),
        ('<c r="A1" cm="1"><v>1</v></c>', 3, "unsupported_target_structure"),
        ('<c r="A1" t="inlineStr"><is><r><t>rich</t></r></is></c>', "plain", "rich_text_cell"),
    ],
)
def test_target_refusals(tmp_path: Path, xml: str, value: object, code: str) -> None:
    result = plan(tmp_path, xml, value)
    assert isinstance(result, Refusal)
    assert result.code == code


@pytest.mark.parametrize(
    ("old", "value", "mode", "expected"),
    [
        ("59", date(1900, 3, 1), False, "61"),
        ("1", date(1900, 2, 28), False, "59"),
        ("0", date(1904, 1, 2), True, "1"),
    ],
)
def test_date_systems(tmp_path: Path, old: str, value: date, mode: bool, expected: str) -> None:
    result = plan(
        tmp_path,
        f'<c r="A1"><v>{old}</v></c>',
        value,
        styles=style_xml('<xf numFmtId="14"/>'),
        date1904=mode,
    )
    assert not isinstance(result, Refusal)
    assert result.category == "date" and result.expected_raw == expected


@pytest.mark.parametrize("old", ["60", "59.5", "-1", "2958466"])
def test_invalid_previous_date_serial(tmp_path: Path, old: str) -> None:
    result = plan(
        tmp_path,
        f'<c r="A1"><v>{old}</v></c>',
        date(2024, 1, 1),
        styles=style_xml('<xf numFmtId="14"/>'),
    )
    assert isinstance(result, Refusal)
    assert result.code == "invalid_value"


@pytest.mark.parametrize(
    ("xf", "base", "allowed"),
    [
        ('<xf numFmtId="14"/>', "", True),
        ('<xf numFmtId="14" xfId="0"/>', '<cellStyleXfs><xf numFmtId="0"/></cellStyleXfs>', False),
        (
            '<xf numFmtId="14" xfId="0" applyNumberFormat="1"/>',
            '<cellStyleXfs><xf numFmtId="0"/></cellStyleXfs>',
            True,
        ),
        ('<xf numFmtId="14" xfId="0"/>', '<cellStyleXfs><xf numFmtId="14"/></cellStyleXfs>', True),
        ('<xf numFmtId="14" applyNumberFormat="0"/>', "", False),
        ('<xf numFmtId="14" xfId="9"/>', '<cellStyleXfs><xf numFmtId="14"/></cellStyleXfs>', False),
    ],
)
def test_positive_format_guard(tmp_path: Path, xf: str, base: str, allowed: bool) -> None:
    result = plan(tmp_path, '<c r="A1"/>', date(2024, 1, 1), styles=style_xml(xf, base))
    assert isinstance(result, Refusal) is not allowed
    if isinstance(result, Refusal):
        assert result.code == "unsupported_target_style"


@pytest.mark.parametrize(
    "text",
    [
        "  spaced  ",
        "\r\n\t",
        "\u00a0😀",
        "_x000D_",
        "_x005F_xD800_",
        "\x00\ufffe",
        '<>&"',
        "=SUM(A1)",
    ],
)
def test_text_writer_decoder_roundtrip(tmp_path: Path, text: str) -> None:
    assert importlib.util.find_spec("xlayer._ooxml.patch") is not None
    from xlayer._ooxml.patch import patch_worksheet

    result = plan(tmp_path, '<c r="A1" s="0"/>', text)
    assert not isinstance(result, Refusal)
    raw = (
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<sheetData><row><c r="A1" s="0"/></row></sheetData></worksheet>'
    )
    patched = patch_worksheet(raw, (result,))
    assert isinstance(patched, bytes)
    from xml.etree import ElementTree as ET

    from xlayer._ooxml._text import TextError, decode_rich_text

    root = ET.fromstring(patched)  # noqa: S314
    elem = root.find(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is")
    assert elem is not None
    decoded = decode_rich_text(elem, "default")
    assert not isinstance(decoded, TextError) and decoded == text
    assert b"<f" not in patched


def test_patch_preserves_prefix_bom_inference_and_outside_bytes(tmp_path: Path) -> None:
    assert importlib.util.find_spec("xlayer._ooxml.patch") is not None
    from xlayer._ooxml.patch import patch_worksheet

    result = plan(tmp_path, "<c><v>1</v></c>", 2)
    assert not isinstance(result, Refusal)
    raw = (
        b'\xef\xbb\xbf<?xml version="1.0" encoding="UTF-8"?>'
        b'<m:worksheet xmlns:m="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b"<!--keep--><m:sheetData><m:row><m:c s = '0'><m:v>1</m:v></m:c>"
        b"</m:row></m:sheetData></m:worksheet>"
    )
    output = patch_worksheet(raw, (result,))
    assert output == raw.replace(b"s = '0'", b"s = '0' t=\"n\"").replace(b">1</m:v>", b">2</m:v>")


def test_utf16_readable_but_not_lexically_editable() -> None:
    from xlayer._ooxml.patch import cell_spans

    xml = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData><row><c r="A1"><v>1</v></c></row></sheetData></worksheet>'
    )
    result = cell_spans(xml.encode("utf-16-le"))
    assert isinstance(result, Refusal) and result.code == "unsupported_write_encoding"
