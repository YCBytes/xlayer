"""Gates for the worksheet reader against crafted archives and cleaned fixtures."""

from __future__ import annotations

import json
import random
import time
import zipfile
from collections.abc import Mapping
from pathlib import Path

import pytest

from xlayer._errors import Refusal
from xlayer._ooxml.archive import WorkbookArchive
from xlayer._ooxml.package import PackageInfo, parse_package
from xlayer._ooxml.sheet import (
    MAX_MERGE_RANGES,
    Cell,
    SharedFormula,
    Worksheet,
    parse_worksheet,
)
from xlayer._ooxml.strings import parse_shared_strings
from xlayer._ooxml.styles import StyleInfo, parse_styles
from xlayer._ooxml.workbook import SheetEntry, parse_workbook_registry

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"

_GENERAL = StyleInfo(
    format_id=0,
    format_code="General",
    is_date=False,
    is_time=False,
    temporal_kind="non_temporal",
)

_CONTENT_TYPES = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels"
    ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
"""

_PACKAGE_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
</Relationships>
"""

_WORKBOOK = (
    b'<?xml version="1.0" encoding="UTF-8"?>'
    b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    b'<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>'
)

_DEFAULT_RELS = f"""<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{_REL}worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
""".encode()


def _letter(column: int) -> str:
    return chr(ord("A") + column - 1)


def _random_rect(rng: random.Random) -> tuple[int, int, int, int]:
    """A (min_row, max_row, min_col, max_col) rectangle covering 2+ cells.

    A single-cell merge is not a merge, and the reader refuses it, so the
    geometry comparisons never generate one.
    """
    while True:
        min_row, max_row = sorted((rng.randint(1, 12), rng.randint(1, 12)))
        min_col, max_col = sorted((rng.randint(1, 6), rng.randint(1, 6)))
        if min_row != max_row or min_col != max_col:
            return min_row, max_row, min_col, max_col


def worksheet_xml(body: str, *, root_attrs: str = "") -> bytes:
    attrs = f" {root_attrs}" if root_attrs else ""
    return (
        f'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{_NS}"{attrs}>{body}</worksheet>'
    ).encode()


def make_archive(tmp_path: Path, parts: Mapping[str, bytes]) -> WorkbookArchive:
    path = tmp_path / "book.xlsx"
    with zipfile.ZipFile(path, "w") as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    archive = WorkbookArchive.load(path)
    assert isinstance(archive, WorkbookArchive)
    return archive


def standard_parts(sheet: bytes, extra: Mapping[str, bytes] | None = None) -> dict[str, bytes]:
    parts = {
        "[Content_Types].xml": _CONTENT_TYPES,
        "_rels/.rels": _PACKAGE_RELS,
        "xl/workbook.xml": _WORKBOOK,
        "xl/_rels/workbook.xml.rels": _DEFAULT_RELS,
        "xl/worksheets/sheet1.xml": sheet,
    }
    if extra:
        parts.update(extra)
    return parts


def sheet_entry(
    *, kind: str = "worksheet", part_path: str = "xl/worksheets/sheet1.xml"
) -> SheetEntry:
    return SheetEntry(
        name="S1",
        sheet_id=1,
        rel_id="rId1",
        state="visible",
        tab_index=0,
        kind=kind,
        part_path=part_path,
    )


def parse_sheet(
    tmp_path: Path,
    sheet_xml: bytes,
    *,
    shared_strings: tuple[str, ...] = (),
    styles: tuple[StyleInfo, ...] = (_GENERAL,),
    entry: SheetEntry | None = None,
    extra: Mapping[str, bytes] | None = None,
) -> Worksheet | Refusal:
    archive = make_archive(tmp_path, standard_parts(sheet_xml, extra))
    package = parse_package(archive)
    assert isinstance(package, PackageInfo), f"package refused: {package!r}"
    return parse_worksheet(
        archive,
        entry or sheet_entry(),
        shared_strings,
        styles,
    )


def expect_sheet(result: Worksheet | Refusal) -> Worksheet:
    assert isinstance(result, Worksheet), f"expected worksheet, got {result!r}"
    return result


def expect_refusal(result: Worksheet | Refusal, code: str) -> Refusal:
    assert isinstance(result, Refusal), f"expected refusal {code!r}, got {result!r}"
    assert result.code == code, f"expected {code!r}, got {result.code!r}: {result.message}"
    assert len(result.recovery_options) > 0, f"refusal {code!r} has no recovery options"
    return result


class TestWorksheetIdentity:
    def test_empty_sheet_has_no_cells_or_used_range(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData/>")
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        assert sheet.name == "S1"
        assert sheet.part_path == "xl/worksheets/sheet1.xml"
        assert dict(sheet.cells) == {}
        assert sheet.used_range is None
        assert sheet.declared_dimension is None
        assert sheet.other_elements == ()
        assert sheet.merges == ()

    def test_declared_dimension_is_recorded_verbatim(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<dimension ref="A1:ZZ999"/><sheetData/>')
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        assert sheet.declared_dimension == "A1:ZZ999"
        assert sheet.used_range is None

    def test_used_range_comes_from_present_cells_not_dimension(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<dimension ref="A1:Z100"/>'
            "<sheetData>"
            '<row r="2"><c r="B2"/></row>'
            '<row r="10"><c r="D10"/></row>'
            "</sheetData>"
        )
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        assert sheet.used_range == "B2:D10"
        assert set(sheet.cells) == {"B2", "D10"}
        assert "A1" not in sheet.cells

    def test_sparse_sheet_does_not_invent_cells(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"/></row><row r="100"><c r="Z100"/></row></sheetData>'
        )
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        assert list(sheet.cells) == ["A1", "Z100"]
        assert sheet.used_range == "A1:Z100"

    def test_omitted_row_and_cell_r_are_inferred(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData><row><c/><c r='C1'/></row><row><c/></row></sheetData>")
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        assert list(sheet.cells) == ["A1", "C1", "A2"]
        assert sheet.cells["A1"].row == 1
        assert sheet.cells["A1"].column == 1
        assert sheet.cells["C1"].column == 3
        assert sheet.cells["A2"].row == 2

    def test_unconsumed_top_level_elements_keep_qualified_names(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            "<sheetViews><sheetView workbookViewId='0'/></sheetViews>"
            "<sheetData/>"
            "<conditionalFormatting sqref='A1'><cfRule type='expression' operator=''>"
            "<formula>1</formula></cfRule></conditionalFormatting>"
        )
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        assert sheet.other_elements == (
            f"{{{_NS}}}sheetViews",
            f"{{{_NS}}}conditionalFormatting",
        )

    def test_empty_other_elements_is_not_a_complete_inventory(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1" ph="1"/></row></sheetData>')
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        assert sheet.other_elements == ()
        assert "A1" in sheet.cells

    def test_chartsheet_is_unsupported_not_corrupt(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData/>")
        refusal = expect_refusal(
            parse_sheet(tmp_path, xml, entry=sheet_entry(kind="chartsheet")),
            "unsupported_sheet_kind",
        )
        assert refusal.details["kind"] == "chartsheet"
        assert refusal.details["part"] == "xl/worksheets/sheet1.xml"

    def test_missing_sheet_part_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData/>")
        refusal = expect_refusal(
            parse_sheet(
                tmp_path,
                xml,
                entry=sheet_entry(part_path="xl/worksheets/missing.xml"),
            ),
            "missing_required_part",
        )
        assert refusal.details["part"] == "xl/worksheets/missing.xml"

    def test_wrong_root_is_refused(self, tmp_path: Path) -> None:
        xml = (
            b'<?xml version="1.0"?><chartsheet xmlns="'
            + _NS.encode()
            + b'"><sheetData/></chartsheet>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_missing_sheet_data_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetViews/>")
        refusal = expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")
        assert refusal.details["reason"] == "worksheet missing sheetData"

    def test_duplicate_sheet_data_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData/><sheetData/>")
        refusal = expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")
        assert refusal.details["reason"] == "duplicate sheetData"

    def test_non_ascii_row_digit_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData><row r='1'><c r='A1\u0662'/></row></sheetData>")
        refusal = expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")
        assert "address" in refusal.details

    def test_column_past_xfd_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData><row r='1'><c r='XFE1'/></row></sheetData>")
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_row_past_excel_limit_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData><row r='1048577'><c r='A1048577'/></row></sheetData>")
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_cell_row_mismatch_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData><row r='1'><c r='A2'/></row></sheetData>")
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_out_of_order_rows_are_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="3"><c r="A3"/></row><row r="1"><c r="A1"/></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_out_of_order_columns_are_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData><row r='1'><c r='C1'/><c r='A1'/></row></sheetData>")
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_duplicate_address_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData><row r='1'><c r='A1'/><c r='A1'/></row></sheetData>")
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_repeated_reads_are_equal(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<dimension ref="A1"/><sheetData><row r="1"><c r="A1"/></row></sheetData>'
        )
        first = expect_sheet(parse_sheet(tmp_path, xml))
        second = expect_sheet(parse_sheet(tmp_path, xml))
        assert first == second

    def test_reading_does_not_change_the_source_file(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData><row r='1'><c r='A1'/></row></sheetData>")
        parts = standard_parts(xml)
        path = tmp_path / "book.xlsx"
        with zipfile.ZipFile(path, "w") as zf:
            for name, data in parts.items():
                zf.writestr(name, data)
        before = path.read_bytes()
        archive = WorkbookArchive.load(path)
        assert isinstance(archive, WorkbookArchive)
        expect_sheet(parse_worksheet(archive, sheet_entry(), (), (_GENERAL,)))
        assert path.read_bytes() == before

    def test_cells_mapping_rejects_mutation(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData><row r='1'><c r='A1'/></row></sheetData>")
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        with pytest.raises(TypeError):
            sheet.cells["B1"] = sheet.cells["A1"]  # type: ignore[index]

    def test_constructor_copies_the_cells_mapping(self) -> None:
        cell = Cell(
            address="A1",
            row=1,
            column=1,
            stored_type="n",
            value_kind="blank",
            raw=None,
            stored_value=None,
            style_index=0,
            style=_GENERAL,
            formula=None,
            merge_ref=None,
        )
        source = {"A1": cell}
        sheet = Worksheet(
            name="S1",
            part_path="xl/worksheets/sheet1.xml",
            cells=source,
            merges=(),
            shared_formulas={},
            declared_dimension=None,
            used_range="A1",
            other_elements=(),
        )
        source["B1"] = cell
        assert "B1" not in sheet.cells
        with pytest.raises(TypeError):
            sheet.cells["C1"] = cell  # type: ignore[index]


class TestStoredValues:
    def test_number_is_decoded_as_float(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<sheetData><row r="1"><c r="B1"><v>999</v></c></row></sheetData>')
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["B1"]
        assert cell.stored_type == "n"
        assert cell.value_kind == "number"
        assert cell.raw == "999"
        assert cell.stored_value == 999.0
        assert cell.formula is None

    def test_scientific_number_is_accepted(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="n"><v>1.5E-3</v></c></row></sheetData>'
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.stored_value == 0.0015

    def test_absent_numeric_cache_is_blank_not_zero(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1" t="n"/></row></sheetData>')
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.value_kind == "blank"
        assert cell.raw is None
        assert cell.stored_value is None

    def test_empty_numeric_cache_is_refused_not_zero(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1" t="n"><v/></c></row></sheetData>')
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_nan_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1"><v>nan</v></c></row></sheetData>')
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_underscore_grouped_number_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1"><v>1_000</v></c></row></sheetData>')
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_boolean_zero_and_one(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            "<sheetData><row r='1'>"
            '<c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>0</v></c>'
            "</row></sheetData>"
        )
        cells = expect_sheet(parse_sheet(tmp_path, xml)).cells
        assert cells["A1"].stored_value is True
        assert cells["B1"].stored_value is False
        assert cells["A1"].value_kind == "boolean"

    def test_boolean_without_cache_is_blank(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1" t="b"/></row></sheetData>')
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.stored_type == "b"
        assert cell.value_kind == "blank"
        assert cell.stored_value is None

    def test_boolean_two_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1" t="b"><v>2</v></c></row></sheetData>')
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_error_value_is_not_a_parser_failure(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c></row></sheetData>'
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.value_kind == "error"
        assert cell.stored_value == "#DIV/0!"
        assert cell.raw == "#DIV/0!"

    def test_error_without_cache_is_blank(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1" t="e"/></row></sheetData>')
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.value_kind == "blank"
        assert cell.stored_value is None

    def test_error_not_starting_with_hash_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="e"><v>DIV0</v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_shared_string_uses_the_table(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1" t="s"><v>1</v></c></row></sheetData>')
        cell = expect_sheet(parse_sheet(tmp_path, xml, shared_strings=("zero", "hello"))).cells[
            "A1"
        ]
        assert cell.value_kind == "string"
        assert cell.raw == "1"
        assert cell.stored_value == "hello"

    def test_out_of_range_shared_string_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="s"><v>500</v></c></row></sheetData>'
        )
        refusal = expect_refusal(
            parse_sheet(tmp_path, xml, shared_strings=("only", "ten") * 5),
            "invalid_part_content",
        )
        assert refusal.details["address"] == "A1"
        assert "500" in str(refusal.details)

    def test_non_ascii_shared_string_index_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="s"><v>\u0661</v></c></row></sheetData>'
        )
        expect_refusal(
            parse_sheet(tmp_path, xml, shared_strings=("a", "b")),
            "invalid_part_content",
        )

    def test_empty_string_is_not_missing(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1" t="str"><v></v></c></row></sheetData>')
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.value_kind == "string"
        assert cell.raw == ""
        assert cell.stored_value == ""

    def test_plain_str_keeps_undecoded_raw(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="str"><v>a_x000D_b</v></c></row></sheetData>'
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.raw == "a_x000D_b"
        assert cell.stored_value == "a\rb"

    def test_iso_date_keeps_validated_string(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="d"><v>1976-11-22T08:30:00</v></c></row></sheetData>'
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.value_kind == "iso_date"
        assert cell.stored_value == "1976-11-22T08:30:00"
        assert cell.raw == "1976-11-22T08:30:00"

    def test_iso_date_banana_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="d"><v>banana</v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_unknown_cell_type_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1" t="q"><v>1</v></c></row></sheetData>')
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_duplicate_v_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><v>1</v><v>2</v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_inline_string_competing_with_v_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="inlineStr">'
            "<is><t>x</t></is><v>0</v></c></row></sheetData>"
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")


class TestInlineStrings:
    def test_sheetdata_preserve_reaches_text(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData xml:space="preserve"><row r="1">'
            '<c r="A1" t="inlineStr"><is><t>  retained  </t></is></c>'
            "</row></sheetData>"
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.stored_value == "  retained  "
        assert cell.value_kind == "string"
        assert cell.raw is None

    def test_row_default_overrides_sheetdata_preserve(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData xml:space="preserve"><row r="1" xml:space="default">'
            '<c r="A1" t="inlineStr"><is><t>  trimmed  </t></is></c>'
            "</row></sheetData>"
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.stored_value == "trimmed"

    def test_rich_runs_concatenate(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="inlineStr"><is>'
            "<r><t>Hel</t></r><r><t>lo</t></r>"
            "</is></c></row></sheetData>"
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.stored_value == "Hello"

    def test_phonetic_guide_is_excluded(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="inlineStr"><is>'
            '<t>base</t><rPh sb="0" eb="1"><t>phon</t></rPh>'
            "</is></c></row></sheetData>"
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.stored_value == "base"

    def test_lone_surrogate_is_refused_with_address(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="inlineStr"><is>'
            "<t>_xD800_</t></is></c></row></sheetData>"
        )
        refusal = expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")
        assert refusal.details["address"] == "A1"
        assert refusal.details["reason"] == "unpaired UTF-16 surrogate in Excel escape"

    def test_str_does_not_apply_rich_text_trimming(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData xml:space="default"><row r="1">'
            '<c r="A1" t="str"><v>  kept  </v></c></row></sheetData>'
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.raw == "  kept  "
        assert cell.stored_value == "  kept  "


class TestStyleResolution:
    def test_omitted_s_uses_styles_zero_not_hardcoded_general(self, tmp_path: Path) -> None:
        date_style = StyleInfo(
            format_id=14,
            format_code="mm-dd-yy",
            is_date=True,
            is_time=False,
            temporal_kind="date",
        )
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1"><v>45747</v></c></row></sheetData>')
        cell = expect_sheet(parse_sheet(tmp_path, xml, styles=(date_style,))).cells["A1"]
        assert cell.style_index == 0
        assert cell.style is date_style
        assert cell.style.temporal_kind == "date"
        assert cell.value_kind == "number"
        assert cell.stored_value == 45747.0

    def test_dangling_style_index_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1" s="3"><v>1</v></c></row></sheetData>')
        expect_refusal(parse_sheet(tmp_path, xml, styles=(_GENERAL,)), "invalid_part_content")

    def test_unknown_format_stays_unknown(self, tmp_path: Path) -> None:
        unknown = StyleInfo(
            format_id=27,
            format_code="",
            is_date=False,
            is_time=False,
            temporal_kind="unknown",
        )
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1" s="0"><v>1</v></c></row></sheetData>')
        cell = expect_sheet(parse_sheet(tmp_path, xml, styles=(unknown,))).cells["A1"]
        assert cell.style.temporal_kind == "unknown"
        assert cell.style.is_date is False
        assert cell.style.is_time is False
        assert cell.value_kind == "number"

    def test_text_cell_with_date_format_stays_string(self, tmp_path: Path) -> None:
        date_style = StyleInfo(
            format_id=14,
            format_code="mm-dd-yy",
            is_date=True,
            is_time=False,
            temporal_kind="date",
        )
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" s="0" t="inlineStr"><is><t>label</t></is></c>'
            "</row></sheetData>"
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml, styles=(date_style,))).cells["A1"]
        assert cell.value_kind == "string"
        assert cell.stored_value == "label"
        assert cell.style.temporal_kind == "date"


class TestScalarPayloads:
    """``<v>``, ``<f>``, and ``<t>`` hold simple text; a child element is malformed.

    Reading ``.text`` alone silently drops everything after the first child,
    which would turn malformed content into a plausible partial value.
    """

    def test_nested_markup_in_numeric_value_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><v>1<extra/>2</v></c></row></sheetData>'
        )
        refusal = expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")
        assert refusal.details["address"] == "A1"

    def test_nested_markup_in_string_value_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="str"><v>abc<b/>def</v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_nested_markup_in_formula_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><f>1+<x/>2</f><v>3</v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_nested_markup_in_inline_text_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="inlineStr"><is>'
            "<t>a<b/>c</t></is></c></row></sheetData>"
        )
        refusal = expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")
        assert refusal.details["address"] == "A1"

    def test_nested_markup_in_inline_run_text_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="inlineStr"><is>'
            "<r><t>a<b/>c</t></r></is></c></row></sheetData>"
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_phonetic_guide_children_are_still_allowed(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="inlineStr"><is>'
            '<t>base</t><rPh sb="0" eb="1"><t>phon</t></rPh>'
            "</is></c></row></sheetData>"
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.stored_value == "base"


class TestFormulas:
    def test_normal_formula_keeps_cached_result_separate(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="12"><c r="D12"><f>B12*C12</f><v>4500</v></c></row></sheetData>'
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["D12"]
        assert cell.formula is not None
        assert cell.formula.text == "B12*C12"
        assert cell.formula.kind == "normal"
        assert cell.formula.calculate_on_next_recalc is False
        assert cell.value_kind == "number"
        assert cell.stored_value == 4500.0
        assert cell.raw == "4500"

    def test_explicit_normal_kind_is_accepted(self, tmp_path: Path) -> None:
        # ECMA-376 ST_CellFormulaType spells the default out as "normal";
        # omitting t and writing t="normal" mean the same thing.
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><f t="normal">1+1</f><v>2</v></c></row></sheetData>'
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.formula is not None
        assert cell.formula.kind == "normal"
        assert cell.formula.text == "1+1"

    def test_empty_normal_formula_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1"><f/><v>1</v></c></row></sheetData>')
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_whitespace_only_normal_formula_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><f> </f><v>1</v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_empty_array_formula_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><f t="array" ref="A1"/><v>1</v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_array_ref_must_contain_the_formula_cell(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1">'
            '<f t="array" ref="B1:B2">SUM(1)</f><v>1</v></c></row></sheetData>'
        )
        refusal = expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")
        assert refusal.details["address"] == "A1"
        assert refusal.details["ref"] == "B1:B2"

    def test_data_table_ref_must_contain_the_formula_cell(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1">'
            '<f t="dataTable" ref="C3:D4"/><v>1</v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_overlapping_master_refs_are_accepted(self, tmp_path: Path) -> None:
        """A master's ``ref`` bounds a group; it does not own those cells.

        Excel writes one master's ``ref`` across cells that a second master
        then claims. Membership comes from each cell's own ``si``, and a cell
        carries at most one ``<f>``, so nothing is ambiguous. This shape is
        invented; the workbook evidence behind the decision is recorded in
        the internal worksheet-reader notes.
        """
        xml = worksheet_xml(
            "<sheetData>"
            '<row r="2"><c r="B2"><f t="shared" ref="B2:B6" si="7">A2*2</f>'
            "<v>2</v></c></row>"
            '<row r="3"><c r="B3"><f t="shared" si="7"/><v>4</v></c></row>'
            '<row r="4"><c r="B4"><f t="shared" ref="B4" si="9">A4*3</f>'
            "<v>9</v></c></row>"
            "</sheetData>"
        )
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        assert sheet.shared_formulas[7].ref == "B2:B6"
        assert sheet.shared_formulas[9].ref == "B4"
        assert sheet.cells["B3"].formula is not None
        assert sheet.cells["B3"].formula.shared_index == 7
        assert sheet.cells["B4"].formula is not None
        assert sheet.cells["B4"].formula.shared_index == 9

    def test_disjoint_shared_masters_are_accepted(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            "<sheetData>"
            '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1</f><v>1</v></c>'
            '<c r="B1"><f t="shared" ref="B1:B2" si="1">2</f><v>2</v></c></row>'
            "</sheetData>"
        )
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        assert sorted(sheet.shared_formulas) == [0, 1]

    def test_master_range_starting_above_a_later_master_is_accepted(self, tmp_path: Path) -> None:
        # The second master's ref begins on row 1, above the first master's
        # own row, so overlap checking cannot assume document order.
        xml = worksheet_xml(
            "<sheetData>"
            '<row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">1</f><v>1</v></c></row>'
            '<row r="11"><c r="B11"><f t="shared" ref="B1:B11" si="1">2</f><v>2</v></c></row>'
            "</sheetData>"
        )
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        assert sheet.shared_formulas[1].ref == "B1:B11"

    def test_formula_without_cache_is_not_malformed(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="b"><f>B1&gt;0</f></c></row></sheetData>'
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.formula is not None
        assert cell.formula.text == "B1>0"
        assert cell.value_kind == "blank"
        assert cell.stored_value is None
        assert cell.raw is None

    def test_empty_formula_cache_is_an_empty_saved_result(self, tmp_path: Path) -> None:
        # A formula with an explicitly empty cache is not a corrupt cell.
        # Generators write this for a formula that has never calculated.
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1"><f>A2</f><v/></c></row></sheetData>')
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.formula is not None
        assert cell.formula.text == "A2"
        assert cell.value_kind == "blank"
        assert cell.raw == ""
        assert cell.stored_value is None

    def test_formula_cache_states_stay_distinct(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            "<sheetData><row r='1'>"
            '<c r="A1"><f>X1</f></c>'
            '<c r="B1"><f>X2</f><v/></c>'
            '<c r="C1"><f>X3</f><v>0</v></c>'
            "</row></sheetData>"
        )
        cells = expect_sheet(parse_sheet(tmp_path, xml)).cells
        assert (cells["A1"].raw, cells["A1"].stored_value) == (None, None)
        assert (cells["B1"].raw, cells["B1"].stored_value) == ("", None)
        assert (cells["C1"].raw, cells["C1"].stored_value) == ("0", 0.0)
        assert cells["C1"].value_kind == "number"

    def test_empty_cache_applies_to_every_scalar_result_type(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            "<sheetData><row r='1'>"
            '<c r="A1" t="b"><f>X1</f><v/></c>'
            '<c r="B1" t="e"><f>X2</f><v/></c>'
            '<c r="C1" t="d"><f>X3</f><v/></c>'
            "</row></sheetData>"
        )
        cells = expect_sheet(parse_sheet(tmp_path, xml)).cells
        for address in ("A1", "B1", "C1"):
            assert cells[address].value_kind == "blank", address
            assert cells[address].raw == "", address
            assert cells[address].stored_value is None, address
            assert cells[address].formula is not None, address

    def test_empty_string_formula_result_stays_an_empty_string(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="str"><f>X1</f><v/></c></row></sheetData>'
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.value_kind == "string"
        assert cell.stored_value == ""

    def test_empty_cache_without_a_formula_is_still_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml('<sheetData><row r="1"><c r="A1"><v/></c></row></sheetData>')
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_malformed_cache_with_a_formula_is_still_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><f>X1</f><v> </v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_whitespace_only_shared_follower_has_no_text(self, tmp_path: Path) -> None:
        # Followers are documented as text=None; a space is not an expression.
        xml = worksheet_xml(
            "<sheetData>"
            '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1</f><v>1</v></c></row>'
            '<row r="2"><c r="A2"><f t="shared" si="0"> </f><v>2</v></c></row>'
            "</sheetData>"
        )
        follower = expect_sheet(parse_sheet(tmp_path, xml)).cells["A2"].formula
        assert follower is not None
        assert follower.kind == "shared"
        assert follower.shared_index == 0
        assert follower.text is None

    def test_shared_master_and_follower(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            "<sheetData>"
            '<row r="8"><c r="B8"><f t="shared" ref="B8:B11" si="0">E8*F8</f>'
            "<v>1000</v></c></row>"
            '<row r="9"><c r="B9"><f t="shared" si="0"/><v>4000</v></c></row>'
            "</sheetData>"
        )
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        master = sheet.cells["B8"].formula
        follower = sheet.cells["B9"].formula
        assert master is not None
        assert master.kind == "shared"
        assert master.text == "E8*F8"
        assert master.shared_index == 0
        assert master.ref == "B8:B11"
        assert follower is not None
        assert follower.text is None
        assert follower.shared_index == 0
        assert sheet.shared_formulas[0] == SharedFormula(master="B8", ref="B8:B11", text="E8*F8")
        assert sheet.cells["B9"].stored_value == 4000.0

    def test_duplicate_shared_master_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            "<sheetData><row r='1'>"
            '<c r="A1"><f t="shared" ref="A1:A2" si="0">1</f><v>1</v></c>'
            '<c r="B1"><f t="shared" ref="B1:B2" si="0">2</f><v>2</v></c>'
            "</row></sheetData>"
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_missing_shared_master_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><f t="shared" si="0"/><v>1</v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_master_missing_ref_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><f t="shared" si="0">A2</f><v>1</v></c>'
            "</row></sheetData>"
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_invalid_shared_index_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1">'
            '<f t="shared" ref="A1" si="x">1</f><v>1</v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_follower_outside_master_range_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            "<sheetData>"
            '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1</f><v>1</v></c></row>'
            '<row r="3"><c r="A3"><f t="shared" si="0"/><v>1</v></c></row>'
            "</sheetData>"
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_array_formula_requires_ref(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><f t="array">SUM(A1)</f><v>1</v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_array_formula_with_ref_is_recorded(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="23"><c r="B23">'
            '<f t="array" ref="B23">SUMPRODUCT(A1:A2)</f><v>9</v></c></row></sheetData>'
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["B23"]
        assert cell.formula is not None
        assert cell.formula.kind == "array"
        assert cell.formula.ref == "B23"
        assert cell.formula.text == "SUMPRODUCT(A1:A2)"

    def test_data_table_requires_ref(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><f t="dataTable"/><v>1</v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_data_table_with_ref_is_recorded(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1">'
            '<f t="dataTable" ref="A1:B2"/><v>1</v></c></row></sheetData>'
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.formula is not None
        assert cell.formula.kind == "data_table"
        assert cell.formula.ref == "A1:B2"
        assert cell.formula.text is None

    def test_unknown_formula_kind_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><f t="mystery">1</f><v>1</v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_calculate_on_next_recalc_accepts_xml_booleans(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><f ca="true">TODAY()</f><v>1</v></c>'
            '<c r="B1"><f ca="1">NOW()</f><v>2</v></c></row></sheetData>'
        )
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        assert sheet.cells["A1"].formula is not None
        assert sheet.cells["A1"].formula.calculate_on_next_recalc is True
        assert sheet.cells["B1"].formula is not None
        assert sheet.cells["B1"].formula.calculate_on_next_recalc is True

    def test_invalid_ca_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><f ca="yes">A2</f><v>1</v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_error_formula_cache_stays_an_error_value(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c></row></sheetData>'
        )
        cell = expect_sheet(parse_sheet(tmp_path, xml)).cells["A1"]
        assert cell.formula is not None
        assert cell.value_kind == "error"
        assert cell.stored_value == "#DIV/0!"

    def test_duplicate_f_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><f>1</f><f>2</f><v>1</v></c></row></sheetData>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_shared_formulas_mapping_is_copied(self) -> None:
        source = {0: SharedFormula(master="A1", ref="A1:A2", text="1")}
        sheet = Worksheet(
            name="S1",
            part_path="xl/worksheets/sheet1.xml",
            cells={},
            merges=(),
            shared_formulas=source,
            declared_dimension=None,
            used_range=None,
            other_elements=(),
        )
        source[1] = SharedFormula(master="B1", ref="B1", text="2")
        assert 1 not in sheet.shared_formulas
        with pytest.raises(TypeError):
            sheet.shared_formulas[2] = source[0]  # type: ignore[index]


class TestMerges:
    def test_merge_at_finds_childless_address(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="10"><c r="A10"><v>10</v></c></row></sheetData>'
            '<mergeCells count="99"><mergeCell ref="B10:C10"/></mergeCells>'
        )
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        found = sheet.merge_at("B10")
        assert found is not None
        assert found.ref == "B10:C10"
        assert found.first_cell == "B10"
        assert found.last_cell == "C10"
        assert "B10" not in sheet.cells
        assert "C10" not in sheet.cells
        assert sheet.merge_at("A10") is None

    def test_existing_cells_record_merge_ref(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c>'
            "</row></sheetData>"
            '<mergeCells><mergeCell ref="A1:B1"/></mergeCells>'
        )
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        assert sheet.cells["A1"].merge_ref == "A1:B1"
        assert sheet.cells["B1"].merge_ref == "A1:B1"
        assert sheet.merges[0].first_cell == "A1"

    def test_inverted_ref_is_normalized(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData/><mergeCells><mergeCell ref='C3:A1'/></mergeCells>")
        merge = expect_sheet(parse_sheet(tmp_path, xml)).merges[0]
        assert merge.first_cell == "A1"
        assert merge.last_cell == "C3"
        assert merge.ref == "C3:A1"

    def test_overlapping_merges_are_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml(
            '<sheetData/><mergeCells><mergeCell ref="A1:B2"/><mergeCell ref="B2:C3"/></mergeCells>'
        )
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_malformed_merge_ref_is_refused(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData/><mergeCells><mergeCell ref='A1'/></mergeCells>")
        expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")

    def test_many_disjoint_merges_stay_within_a_bounded_cost(self, tmp_path: Path) -> None:
        # Comparing every merge with every earlier one made this quadratic:
        # 20,000 merges is ~200 million comparisons, plus one scan per cell.
        count = 20_000
        merges = "".join(f'<mergeCell ref="A{row}:B{row}"/>' for row in range(1, count + 1))
        rows = "".join(
            f'<row r="{row}"><c r="A{row}"><v>{row}</v></c></row>' for row in range(1, count + 1)
        )
        xml = worksheet_xml(
            f'<sheetData>{rows}</sheetData><mergeCells count="{count}">{merges}</mergeCells>'
        )
        started = time.perf_counter()
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        elapsed = time.perf_counter() - started
        assert len(sheet.merges) == count
        assert sheet.cells["A7"].merge_ref == "A7:B7"
        assert sheet.cells[f"A{count}"].merge_ref == f"A{count}:B{count}"
        assert elapsed < 5.0, f"merge handling took {elapsed:.1f}s for {count} merges"

    def test_merge_count_over_the_cap_is_refused(self, tmp_path: Path) -> None:
        count = MAX_MERGE_RANGES + 1
        merges = "".join(f'<mergeCell ref="A{row}:B{row}"/>' for row in range(1, count + 1))
        xml = worksheet_xml(f"<sheetData/><mergeCells>{merges}</mergeCells>")
        refusal = expect_refusal(parse_sheet(tmp_path, xml), "invalid_part_content")
        assert refusal.details["reason"] == "too many merge ranges"
        assert refusal.details["limit"] == MAX_MERGE_RANGES

    def test_overlap_detection_matches_a_pairwise_check(self, tmp_path: Path) -> None:
        # The sweep replaced an every-pair comparison, so compare the two on
        # random rectangles: same accept/refuse answer, every time.
        rng = random.Random(20260912)  # noqa: S311
        for _ in range(40):
            rects = [_random_rect(rng) for _ in range(rng.randint(2, 7))]
            refs = [
                f"{_letter(min_col)}{min_row}:{_letter(max_col)}{max_row}"
                for min_row, max_row, min_col, max_col in rects
            ]
            overlaps = any(
                not (
                    left[1] < right[0]
                    or right[1] < left[0]
                    or left[3] < right[2]
                    or right[3] < left[2]
                )
                for index, left in enumerate(rects)
                for right in rects[:index]
            )
            body = "".join(f'<mergeCell ref="{ref}"/>' for ref in refs)
            result = parse_sheet(
                tmp_path, worksheet_xml(f"<sheetData/><mergeCells>{body}</mergeCells>")
            )
            if overlaps:
                expect_refusal(result, "invalid_part_content")
            else:
                assert len(expect_sheet(result).merges) == len(refs)

    def test_merge_lookup_matches_a_scan_of_every_merge(self, tmp_path: Path) -> None:
        rng = random.Random(902)  # noqa: S311
        for _ in range(40):
            placed: list[tuple[int, int, int, int]] = []
            for _ in range(rng.randint(1, 6)):
                candidate = _random_rect(rng)
                if all(
                    candidate[1] < other[0]
                    or other[1] < candidate[0]
                    or candidate[3] < other[2]
                    or other[3] < candidate[2]
                    for other in placed
                ):
                    placed.append(candidate)
            rng.shuffle(placed)
            refs = [
                f"{_letter(min_col)}{min_row}:{_letter(max_col)}{max_row}"
                for min_row, max_row, min_col, max_col in placed
            ]
            addresses = sorted({(rng.randint(1, 12), rng.randint(1, 6)) for _ in range(10)})
            rows = ""
            for row in sorted({row for row, _ in addresses}):
                cells = "".join(
                    f'<c r="{_letter(col)}{row}"><v>1</v></c>'
                    for candidate_row, col in addresses
                    if candidate_row == row
                )
                rows += f'<row r="{row}">{cells}</row>'
            merges = "".join(f'<mergeCell ref="{ref}"/>' for ref in refs)
            sheet = expect_sheet(
                parse_sheet(
                    tmp_path,
                    worksheet_xml(
                        f"<sheetData>{rows}</sheetData><mergeCells>{merges}</mergeCells>"
                    ),
                )
            )
            for cell in sheet.cells.values():
                scanned = next(
                    (merge.ref for merge in sheet.merges if merge.contains(cell.row, cell.column)),
                    None,
                )
                assert cell.merge_ref == scanned, f"{cell.address} in {refs}"

    def test_merge_cells_is_consumed_not_other_elements(self, tmp_path: Path) -> None:
        xml = worksheet_xml("<sheetData/><mergeCells><mergeCell ref='A1:B1'/></mergeCells>")
        sheet = expect_sheet(parse_sheet(tmp_path, xml))
        assert sheet.other_elements == ()
        assert len(sheet.merges) == 1


def load_fixture_sheet(name: str, sheet_name: str) -> tuple[Path, Worksheet]:
    path = FIXTURES / name
    archive = WorkbookArchive.load(path)
    assert isinstance(archive, WorkbookArchive)
    package = parse_package(archive)
    assert isinstance(package, PackageInfo), f"package refused: {package!r}"
    registry = parse_workbook_registry(archive, package)
    assert not isinstance(registry, Refusal), f"registry refused: {registry!r}"
    strings = parse_shared_strings(archive, package)
    assert not isinstance(strings, Refusal), f"strings refused: {strings!r}"
    styles = parse_styles(archive, package)
    assert not isinstance(styles, Refusal), f"styles refused: {styles!r}"
    entry = registry.sheet_by_name(sheet_name)
    assert entry is not None, f"missing sheet {sheet_name!r}"
    sheet = parse_worksheet(archive, entry, strings, styles)
    assert isinstance(sheet, Worksheet), f"sheet refused: {sheet!r}"
    return path, sheet


def assert_cell_matches(cell: Cell, expected: Mapping[str, object]) -> None:
    if "stored_type" in expected:
        assert cell.stored_type == expected["stored_type"]
    if "value_kind" in expected:
        assert cell.value_kind == expected["value_kind"]
    if "raw" in expected:
        assert cell.raw == expected["raw"]
    if "stored_value" in expected:
        assert cell.stored_value == expected["stored_value"]
    if "style_index" in expected:
        assert cell.style_index == expected["style_index"]
    if "temporal_kind" in expected:
        assert cell.style.temporal_kind == expected["temporal_kind"]
    if "merge_ref" in expected:
        assert cell.merge_ref == expected["merge_ref"]
    if "formula" in expected:
        formula_expected = expected["formula"]
        if formula_expected is None:
            assert cell.formula is None
        else:
            assert cell.formula is not None
            assert isinstance(formula_expected, Mapping)
            for key, value in formula_expected.items():
                assert getattr(cell.formula, key) == value
            assert cell.formula is not None  # stored_value is an unverified cache


class TestWorksheetFixtures:
    def test_edges_fixture(self) -> None:
        expected = json.loads(
            (FIXTURES / "test_workbook_7_write_v11_edges.sheet.json").read_text(encoding="utf-8")
        )
        path, sheet = load_fixture_sheet("test_workbook_7_write_v11_edges.xlsx", expected["sheet"])
        assert sheet.part_path == expected["part_path"]
        assert sheet.declared_dimension == expected["declared_dimension"]
        assert sheet.used_range == expected["used_range"]
        assert list(sheet.other_elements) == expected["other_elements"]
        assert len(sheet.cells) == expected["cell_count"]
        assert [merge.ref for merge in sheet.merges] == [item["ref"] for item in expected["merges"]]
        assert sheet.merge_at("B10") is not None
        assert sheet.merge_at("C10") is not None
        for address, spec in expected["cells"].items():
            assert_cell_matches(sheet.cells[address], spec)
        for address in expected["absent"]:
            assert address not in sheet.cells
        before = path.read_bytes()
        load_fixture_sheet("test_workbook_7_write_v11_edges.xlsx", expected["sheet"])
        assert path.read_bytes() == before

    def test_worksheet_final_selected_sheets(self) -> None:
        expected = json.loads(
            (FIXTURES / "test_workbook_4_worksheet_final.sheet.json").read_text(encoding="utf-8")
        )
        _, types = load_fixture_sheet("test_workbook_4_worksheet_final.xlsx", "cell_types")
        assert types.part_path == expected["cell_types"]["part_path"]
        assert types.declared_dimension == expected["cell_types"]["declared_dimension"]
        assert types.used_range == expected["cell_types"]["used_range"]
        for address, spec in expected["cell_types"]["cells"].items():
            assert_cell_matches(types.cells[address], spec)

        _, formulas = load_fixture_sheet("test_workbook_4_worksheet_final.xlsx", "formulas")
        for address, spec in expected["formulas"]["cells"].items():
            assert_cell_matches(formulas.cells[address], spec)
        shared = expected["formulas"]["shared_formulas"]["0"]
        assert formulas.shared_formulas[0].master == shared["master"]
        assert formulas.shared_formulas[0].ref == shared["ref"]
        assert formulas.shared_formulas[0].text == shared["text"]

        _, merged = load_fixture_sheet("test_workbook_4_worksheet_final.xlsx", "merged_cells")
        assert [merge.ref for merge in merged.merges] == expected["merged_cells"]["merges"]
        for address, spec in expected["merged_cells"]["cells"].items():
            assert_cell_matches(merged.cells[address], spec)

        _, empty = load_fixture_sheet("test_workbook_4_worksheet_final.xlsx", "empty_sheet")
        assert dict(empty.cells) == {}
        assert empty.used_range is None
        assert empty.declared_dimension == expected["empty_sheet"]["declared_dimension"]

        _, single = load_fixture_sheet("test_workbook_4_worksheet_final.xlsx", "single_cell")
        assert single.used_range == "A1"
        assert_cell_matches(single.cells["A1"], expected["single_cell"]["cells"]["A1"])

    def test_registry_visible_one_numeric_and_shared_string(self) -> None:
        _, sheet = load_fixture_sheet("test_workbook_2_registry.xlsx", "USC - Hedging")
        assert sheet.cells["A1"].value_kind == "string"
        assert sheet.cells["A1"].stored_value == "Sheet 2 - Visible, name has spaces and hyphen"
        assert sheet.cells["B1"].value_kind == "number"
        assert sheet.cells["B1"].raw == "999"
        assert sheet.cells["B1"].stored_value == 999.0

    def test_1904_date_serial_keeps_raw_and_temporal_kind(self) -> None:
        _, sheet = load_fixture_sheet("test_workbook_6b_1904.xlsx", "Date 1904 Test")
        cell = sheet.cells["B2"]
        assert cell.value_kind == "number"
        assert cell.raw == "44196"
        assert cell.stored_value == 44196.0
        assert cell.style.temporal_kind == "date"

    def test_formats_shared_strings_and_blank_styled_cell(self) -> None:
        _, sheet1 = load_fixture_sheet("test_workbook_3_formats.xlsx", "Formats and Strings")
        assert sheet1.cells["B9"].stored_value == "normal string before rich text"
        assert sheet1.cells["B10"].stored_value == "  leading spaces preserved  "
        _, sheet2 = load_fixture_sheet("test_workbook_3_formats.xlsx", "Styles Inflation")
        blank = sheet2.cells["C2"]
        assert blank.value_kind == "blank"
        assert blank.raw is None
        assert blank.stored_value is None
        assert blank.stored_type == "n"

    def test_repeated_fixture_reads_are_equal(self) -> None:
        first = load_fixture_sheet("test_workbook_7_write_v11_edges.xlsx", "edges")[1]
        second = load_fixture_sheet("test_workbook_7_write_v11_edges.xlsx", "edges")[1]
        assert first == second
