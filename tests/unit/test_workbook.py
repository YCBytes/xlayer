"""Gates for the workbook registry against crafted archives and cleaned fixtures."""

from __future__ import annotations

import json
import zipfile
from collections.abc import Mapping
from pathlib import Path

import pytest

from xlayer._errors import Refusal
from xlayer._ooxml.archive import WorkbookArchive
from xlayer._ooxml.package import PackageInfo, parse_package
from xlayer._ooxml.workbook import (
    DefinedName,
    SheetEntry,
    WorkbookRegistry,
    parse_workbook_registry,
)

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"

_MINIMAL_SHEET = (
    b'<?xml version="1.0" encoding="UTF-8"?>'
    b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    b"<sheetData/></worksheet>"
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

_DEFAULT_WB_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>
"""

_SHEET1 = '<sheet name="S1" sheetId="1" r:id="rId1"/>'


def _workbook(
    *,
    sheets: str = _SHEET1,
    pr: str = "",
    names: str = "",
    active_tab: str | None = "0",
    extra: str = "",
) -> bytes:
    view = (
        ""
        if active_tab is None
        else f'<bookViews><workbookView activeTab="{active_tab}"/></bookViews>'
    )
    names_xml = f"<definedNames>{names}</definedNames>" if names else ""
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{_NS}" xmlns:r="{_NS_R}">'
        f"{pr}{view}<sheets>{sheets}</sheets>{names_xml}{extra}</workbook>"
    ).encode()


def make_archive(tmp_path: Path, parts: Mapping[str, bytes]) -> WorkbookArchive:
    path = tmp_path / "book.xlsx"
    with zipfile.ZipFile(path, "w") as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    archive = WorkbookArchive.load(path)
    assert isinstance(archive, WorkbookArchive)
    return archive


def standard_parts(
    workbook: bytes,
    *,
    workbook_rels: bytes = _DEFAULT_WB_RELS,
    extra: Mapping[str, bytes] | None = None,
) -> dict[str, bytes]:
    parts = {
        "[Content_Types].xml": _CONTENT_TYPES,
        "_rels/.rels": _PACKAGE_RELS,
        "xl/workbook.xml": workbook,
        "xl/_rels/workbook.xml.rels": workbook_rels,
        "xl/worksheets/sheet1.xml": _MINIMAL_SHEET,
    }
    if extra:
        parts.update(extra)
    return parts


def parse_registry(tmp_path: Path, parts: Mapping[str, bytes]) -> WorkbookRegistry | Refusal:
    archive = make_archive(tmp_path, parts)
    package = parse_package(archive)
    assert isinstance(package, PackageInfo), f"package refused: {package!r}"
    return parse_workbook_registry(archive, package)


def expect_registry(result: WorkbookRegistry | Refusal) -> WorkbookRegistry:
    assert isinstance(result, WorkbookRegistry), f"expected registry, got {result!r}"
    return result


def expect_refusal(result: WorkbookRegistry | Refusal, code: str) -> Refusal:
    assert isinstance(result, Refusal), f"expected refusal {code!r}, got {result!r}"
    assert result.code == code, f"expected {code!r}, got {result.code!r}: {result.message}"
    assert len(result.recovery_options) > 0, f"refusal {code!r} has no recovery options"
    return result


def load_fixture(name: str) -> tuple[WorkbookArchive, PackageInfo, WorkbookRegistry]:
    archive = WorkbookArchive.load(FIXTURES / name)
    assert isinstance(archive, WorkbookArchive)
    package = parse_package(archive)
    assert isinstance(package, PackageInfo), f"package refused: {package!r}"
    registry = parse_workbook_registry(archive, package)
    assert isinstance(registry, WorkbookRegistry), f"registry refused: {registry!r}"
    return archive, package, registry


def assert_matches_expected(registry: WorkbookRegistry, expected_name: str) -> None:
    expected = json.loads((FIXTURES / expected_name).read_text(encoding="utf-8"))
    assert registry.active_tab == expected["active_tab"]
    assert registry.date1904 is expected["date1904"]
    assert list(registry.external_reference_ids) == expected["external_reference_ids"]
    assert [
        {
            "name": sheet.name,
            "sheet_id": sheet.sheet_id,
            "rel_id": sheet.rel_id,
            "state": sheet.state,
            "tab_index": sheet.tab_index,
            "kind": sheet.kind,
            "part_path": sheet.part_path,
        }
        for sheet in registry.sheets
    ] == expected["sheets"]
    assert [
        {
            "name": item.name,
            "text": item.text,
            "scope_sheet": item.scope_sheet,
            "is_hidden": item.is_hidden,
        }
        for item in registry.defined_names
    ] == expected["defined_names"]


class TestRegistryFixture:
    def test_matches_independently_transcribed_xml(self) -> None:
        archive, package, registry = load_fixture("test_workbook_2_registry.xlsx")
        assert package.workbook_path == "xl/workbook.xml"
        assert_matches_expected(registry, "test_workbook_2_registry.expected.json")
        archive.close()

    def test_local_sheet_id_is_tab_index_not_sheet_id(self) -> None:
        # HedgeRates has localSheetId="0". Tab 0 is "USC - Hedging" (sheetId 1);
        # the expression text points at "Visible One" (sheetId 2, tab 4). Scope
        # must follow tab order, not sheetId and not the expression text.
        _archive, _package, registry = load_fixture("test_workbook_2_registry.xlsx")
        hedge = next(item for item in registry.defined_names if item.name == "HedgeRates")
        assert hedge.scope_sheet == "USC - Hedging"
        visible_one = registry.sheet_by_name("Visible One")
        assert visible_one is not None
        assert visible_one.sheet_id == 2
        _archive.close()

    def test_defined_name_text_is_raw_including_ref_token(self) -> None:
        _archive, _package, registry = load_fixture("test_workbook_2_registry.xlsx")
        broken = next(item for item in registry.defined_names if item.name == "BrokenRef")
        assert broken.text == "#REF!"
        assert not hasattr(broken, "is_broken")
        _archive.close()

    def test_sheet_by_name_and_immutability(self) -> None:
        _archive, _package, registry = load_fixture("test_workbook_2_registry.xlsx")
        found = registry.sheet_by_name("Hidden Sheet")
        assert found is not None
        assert found.state == "hidden"
        assert registry.sheet_by_name("missing") is None
        with pytest.raises(AttributeError):
            registry.sheets = ()  # type: ignore[misc]
        _archive.close()


class TestDate1904Fixture:
    def test_matches_independently_transcribed_xml(self) -> None:
        archive, package, registry = load_fixture("test_workbook_6b_1904.xlsx")
        assert package.workbook_path == "xl/workbook.xml"
        assert_matches_expected(registry, "test_workbook_6b_1904.expected.json")
        archive.close()


class TestBooleanSpellings:
    @pytest.mark.parametrize(
        ("raw", "expected"),
        [("1", True), ("0", False), ("true", True), ("false", False)],
    )
    def test_date1904_accepts_xml_schema_booleans(
        self, tmp_path: Path, raw: str, expected: bool
    ) -> None:
        registry = expect_registry(
            parse_registry(
                tmp_path, standard_parts(_workbook(pr=f'<workbookPr date1904="{raw}"/>'))
            )
        )
        assert registry.date1904 is expected

    def test_date1904_defaults_false_when_absent(self, tmp_path: Path) -> None:
        registry = expect_registry(parse_registry(tmp_path, standard_parts(_workbook())))
        assert registry.date1904 is False
        registry = expect_registry(
            parse_registry(tmp_path, standard_parts(_workbook(pr="<workbookPr/>")))
        )
        assert registry.date1904 is False

    @pytest.mark.parametrize("raw", ["on", "TRUE", "yes", ""])
    def test_date1904_refuses_malformed_values(self, tmp_path: Path, raw: str) -> None:
        refusal = expect_refusal(
            parse_registry(
                tmp_path, standard_parts(_workbook(pr=f'<workbookPr date1904="{raw}"/>'))
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "invalid date1904"
        assert refusal.details["part"] == "xl/workbook.xml"

    @pytest.mark.parametrize(
        ("raw", "expected"),
        [("1", True), ("0", False), ("true", True), ("false", False)],
    )
    def test_defined_name_hidden_accepts_xml_schema_booleans(
        self, tmp_path: Path, raw: str, expected: bool
    ) -> None:
        registry = expect_registry(
            parse_registry(
                tmp_path,
                standard_parts(
                    _workbook(names=f'<definedName name="N" hidden="{raw}">A1</definedName>')
                ),
            )
        )
        assert registry.defined_names == (
            DefinedName(name="N", text="A1", scope_sheet=None, is_hidden=expected),
        )

    def test_defined_name_hidden_defaults_false(self, tmp_path: Path) -> None:
        registry = expect_registry(
            parse_registry(
                tmp_path,
                standard_parts(_workbook(names='<definedName name="N">A1</definedName>')),
            )
        )
        assert registry.defined_names[0].is_hidden is False

    @pytest.mark.parametrize("raw", ["on", "TRUE", "2"])
    def test_defined_name_hidden_refuses_malformed_values(self, tmp_path: Path, raw: str) -> None:
        refusal = expect_refusal(
            parse_registry(
                tmp_path,
                standard_parts(
                    _workbook(names=f'<definedName name="N" hidden="{raw}">A1</definedName>')
                ),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "invalid hidden"


class TestDefinedNames:
    def test_quoted_ref_token_is_raw_text_not_a_broken_flag(self, tmp_path: Path) -> None:
        registry = expect_registry(
            parse_registry(
                tmp_path,
                standard_parts(_workbook(names='<definedName name="Quoted">"#REF!"</definedName>')),
            )
        )
        assert registry.defined_names[0].text == '"#REF!"'
        assert not hasattr(registry.defined_names[0], "is_broken")

    def test_empty_text_is_preserved(self, tmp_path: Path) -> None:
        registry = expect_registry(
            parse_registry(
                tmp_path,
                standard_parts(_workbook(names='<definedName name="Empty"/>')),
            )
        )
        assert registry.defined_names[0].text == ""

    def test_local_sheet_id_follows_tab_order(self, tmp_path: Path) -> None:
        # sheetId values are 10 then 1; localSheetId="1" must resolve to tab 1.
        sheets = (
            '<sheet name="First" sheetId="10" r:id="rId1"/>'
            '<sheet name="Second" sheetId="1" r:id="rId2"/>'
        )
        rels = (
            '<?xml version="1.0"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{_REL}worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{_REL}worksheet" Target="worksheets/sheet2.xml"/>'
            "</Relationships>"
        ).encode()
        types = _CONTENT_TYPES.replace(
            b"</Types>",
            b'  <Override PartName="/xl/worksheets/sheet2.xml" '
            b'ContentType="application/vnd.openxmlformats-officedocument.'
            b'spreadsheetml.worksheet+xml"/>\n</Types>',
        )
        registry = expect_registry(
            parse_registry(
                tmp_path,
                {
                    **standard_parts(
                        _workbook(
                            sheets=sheets,
                            names='<definedName name="N" localSheetId="1">A1</definedName>',
                        ),
                        workbook_rels=rels,
                    ),
                    "[Content_Types].xml": types,
                    "xl/worksheets/sheet2.xml": _MINIMAL_SHEET,
                },
            )
        )
        assert registry.defined_names[0].scope_sheet == "Second"
        assert registry.sheets[1].sheet_id == 1

    def test_out_of_range_local_sheet_id_refused(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_registry(
                tmp_path,
                standard_parts(
                    _workbook(names='<definedName name="N" localSheetId="1">A1</definedName>')
                ),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "localSheetId out of range"

    def test_non_integer_local_sheet_id_refused(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_registry(
                tmp_path,
                standard_parts(
                    _workbook(names='<definedName name="N" localSheetId="x">A1</definedName>')
                ),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "invalid localSheetId"

    def test_missing_defined_name_name_refused(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_registry(
                tmp_path,
                standard_parts(_workbook(names="<definedName>A1</definedName>")),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "definedName missing name"


class TestSheetKindsAndPaths:
    def test_chartsheet_and_dialogsheet_are_retained(self, tmp_path: Path) -> None:
        sheets = (
            '<sheet name="Data" sheetId="1" r:id="rId1"/>'
            '<sheet name="Chart" sheetId="2" r:id="rId2"/>'
            '<sheet name="Dialog" sheetId="3" r:id="rId3"/>'
        )
        rels = (
            '<?xml version="1.0"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{_REL}worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{_REL}chartsheet" Target="chartsheets/sheet1.xml"/>'
            f'<Relationship Id="rId3" Type="{_REL}dialogsheet" Target="dialogsheets/sheet1.xml"/>'
            "</Relationships>"
        ).encode()
        registry = expect_registry(
            parse_registry(
                tmp_path,
                {
                    **standard_parts(_workbook(sheets=sheets), workbook_rels=rels),
                    "xl/chartsheets/sheet1.xml": _MINIMAL_SHEET,
                    "xl/dialogsheets/sheet1.xml": _MINIMAL_SHEET,
                },
            )
        )
        assert [sheet.kind for sheet in registry.sheets] == [
            "worksheet",
            "chartsheet",
            "dialogsheet",
        ]
        assert registry.sheets[1].part_path == "xl/chartsheets/sheet1.xml"
        assert registry.sheets[2].part_path == "xl/dialogsheets/sheet1.xml"
        assert registry.sheets[1].tab_index == 1

    def test_workbook_path_is_taken_from_package(self, tmp_path: Path) -> None:
        workbook = _workbook(sheets='<sheet name="" sheetId="1" r:id="rId1"/>')
        types = (
            b'<?xml version="1.0"?>'
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            b'<Default Extension="rels" '
            b'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            b'<Default Extension="xml" ContentType="application/xml"/>'
            b'<Override PartName="/custom/book.xml" ContentType="'
            b'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            b'<Override PartName="/custom/worksheets/sheet1.xml" ContentType="'
            b'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            b"</Types>"
        )
        package_rels = (
            b'<?xml version="1.0"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId1" '
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
            b'officeDocument" Target="custom/book.xml"/>'
            b"</Relationships>"
        )
        workbook_rels = (
            b'<?xml version="1.0"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId1" '
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
            b'worksheet" Target="worksheets/sheet1.xml"/>'
            b"</Relationships>"
        )
        archive = make_archive(
            tmp_path,
            {
                "[Content_Types].xml": types,
                "_rels/.rels": package_rels,
                "custom/book.xml": workbook,
                "custom/_rels/book.xml.rels": workbook_rels,
                "custom/worksheets/sheet1.xml": _MINIMAL_SHEET,
            },
        )
        package = parse_package(archive)
        assert isinstance(package, PackageInfo)
        assert package.workbook_path == "custom/book.xml"
        refusal = expect_refusal(parse_workbook_registry(archive, package), "invalid_part_content")
        assert refusal.details["part"] == "custom/book.xml"
        assert refusal.details["reason"] == "sheet missing name"
        archive.close()

    def test_absolute_worksheet_target_resolves(self, tmp_path: Path) -> None:
        rels = _DEFAULT_WB_RELS.replace(
            b'Target="worksheets/sheet1.xml"', b'Target="/xl/worksheets/sheet1.xml"'
        )
        registry = expect_registry(
            parse_registry(tmp_path, standard_parts(_workbook(), workbook_rels=rels))
        )
        assert registry.sheets[0].part_path == "xl/worksheets/sheet1.xml"


class TestExternalReferences:
    def test_collects_relationship_ids(self, tmp_path: Path) -> None:
        extra = (
            f'<externalReferences><externalReference r:id="rId9"/>'
            f'<externalReference xmlns:r="{_NS_R}" r:id="rId10"/>'
            f"</externalReferences>"
        )
        registry = expect_registry(parse_registry(tmp_path, standard_parts(_workbook(extra=extra))))
        assert registry.external_reference_ids == ("rId9", "rId10")

    def test_missing_external_reference_id_refused(self, tmp_path: Path) -> None:
        extra = "<externalReferences><externalReference/></externalReferences>"
        refusal = expect_refusal(
            parse_registry(tmp_path, standard_parts(_workbook(extra=extra))),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "externalReference missing r:id"


class TestRegistryRefusals:
    def test_unexpected_root_element(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_registry(
                tmp_path,
                standard_parts(
                    b'<?xml version="1.0"?><sst xmlns="'
                    b'http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
                ),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "unexpected root element"
        assert refusal.details["part"] == "xl/workbook.xml"

    def test_malformed_workbook_xml(self, tmp_path: Path) -> None:
        expect_refusal(
            parse_registry(tmp_path, standard_parts(b"<workbook><broken")),
            "malformed_xml",
        )

    def test_zero_sheets(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_registry(tmp_path, standard_parts(_workbook(sheets=""))),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "workbook has no sheets"

    def test_missing_sheet_id(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_registry(
                tmp_path,
                standard_parts(_workbook(sheets='<sheet name="S1" r:id="rId1"/>')),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "sheet missing sheetId"

    def test_non_integer_sheet_id(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_registry(
                tmp_path,
                standard_parts(_workbook(sheets='<sheet name="S1" sheetId="1.5" r:id="rId1"/>')),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "invalid sheetId"

    def test_zero_sheet_id(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_registry(
                tmp_path,
                standard_parts(_workbook(sheets='<sheet name="S1" sheetId="0" r:id="rId1"/>')),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "invalid sheetId"

    @pytest.mark.parametrize(
        ("workbook", "reason"),
        [
            (
                _workbook(sheets='<sheet name="S1" sheetId="²" r:id="rId1"/>'),
                "invalid sheetId",
            ),
            (_workbook(active_tab="²"), "invalid activeTab"),
            (
                _workbook(names='<definedName name="N" localSheetId="²">A1</definedName>'),
                "invalid localSheetId",
            ),
        ],
        ids=["sheetId", "activeTab", "localSheetId"],
    )
    def test_superscript_digits_are_refused_not_raised(
        self, tmp_path: Path, workbook: bytes, reason: str
    ) -> None:
        # "²".isdigit() is True, but int("²") raises ValueError. Attacker-
        # controlled XML must not escape as an unstructured exception.
        refusal = expect_refusal(
            parse_registry(tmp_path, standard_parts(workbook)), "invalid_part_content"
        )
        assert refusal.details["reason"] == reason

    def test_oversized_digit_string_is_refused_not_raised(self, tmp_path: Path) -> None:
        huge = "1" * 5000
        refusal = expect_refusal(
            parse_registry(
                tmp_path,
                standard_parts(
                    _workbook(sheets=f'<sheet name="S1" sheetId="{huge}" r:id="rId1"/>')
                ),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "invalid sheetId"

    def test_missing_relationship_id(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_registry(
                tmp_path,
                standard_parts(_workbook(sheets='<sheet name="S1" sheetId="1"/>')),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "sheet missing r:id"

    def test_unknown_sheet_state(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_registry(
                tmp_path,
                standard_parts(
                    _workbook(sheets='<sheet name="S1" sheetId="1" state="Hidden" r:id="rId1"/>')
                ),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "invalid sheet state"

    def test_duplicate_sheet_names_are_case_insensitive(self, tmp_path: Path) -> None:
        sheets = (
            '<sheet name="Data" sheetId="1" r:id="rId1"/>'
            '<sheet name="data" sheetId="2" r:id="rId2"/>'
        )
        rels = (
            '<?xml version="1.0"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{_REL}worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{_REL}worksheet" Target="worksheets/sheet2.xml"/>'
            "</Relationships>"
        ).encode()
        refusal = expect_refusal(
            parse_registry(
                tmp_path,
                {
                    **standard_parts(_workbook(sheets=sheets), workbook_rels=rels),
                    "xl/worksheets/sheet2.xml": _MINIMAL_SHEET,
                },
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "duplicate sheet name"

    def test_duplicate_sheet_id(self, tmp_path: Path) -> None:
        sheets = (
            '<sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="1" r:id="rId2"/>'
        )
        rels = (
            '<?xml version="1.0"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{_REL}worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{_REL}worksheet" Target="worksheets/sheet2.xml"/>'
            "</Relationships>"
        ).encode()
        refusal = expect_refusal(
            parse_registry(
                tmp_path,
                {
                    **standard_parts(_workbook(sheets=sheets), workbook_rels=rels),
                    "xl/worksheets/sheet2.xml": _MINIMAL_SHEET,
                },
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "duplicate sheetId"

    def test_duplicate_relationship_id_across_tabs(self, tmp_path: Path) -> None:
        sheets = (
            '<sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId1"/>'
        )
        refusal = expect_refusal(
            parse_registry(tmp_path, standard_parts(_workbook(sheets=sheets))),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "duplicate sheet part"
        assert refusal.details["part_path"] == "xl/worksheets/sheet1.xml"

    def test_distinct_relationships_resolving_to_the_same_part(self, tmp_path: Path) -> None:
        sheets = (
            '<sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/>'
        )
        rels = (
            '<?xml version="1.0"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{_REL}worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{_REL}worksheet" Target="/xl/worksheets/sheet1.xml"/>'
            "</Relationships>"
        ).encode()
        refusal = expect_refusal(
            parse_registry(tmp_path, standard_parts(_workbook(sheets=sheets), workbook_rels=rels)),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "duplicate sheet part"
        assert refusal.details["part_path"] == "xl/worksheets/sheet1.xml"

    def test_unknown_relationship_id(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_registry(
                tmp_path,
                standard_parts(_workbook(sheets='<sheet name="S1" sheetId="1" r:id="rId99"/>')),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "sheet r:id is not a sheet relationship"

    def test_non_sheet_relationship_type(self, tmp_path: Path) -> None:
        rels = (
            '<?xml version="1.0"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{_REL}styles" Target="styles.xml"/>'
            "</Relationships>"
        ).encode()
        refusal = expect_refusal(
            parse_registry(tmp_path, standard_parts(_workbook(), workbook_rels=rels)),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "sheet r:id is not a sheet relationship"

    def test_missing_worksheet_part(self, tmp_path: Path) -> None:
        parts = standard_parts(_workbook())
        del parts["xl/worksheets/sheet1.xml"]
        refusal = expect_refusal(parse_registry(tmp_path, parts), "missing_required_part")
        assert refusal.details["part"] == "xl/worksheets/sheet1.xml"

    def test_unsafe_worksheet_target(self, tmp_path: Path) -> None:
        # One leading ".." from xl/ stays inside the package after
        # normalization; two levels escape it and must be refused as unsafe.
        rels = _DEFAULT_WB_RELS.replace(
            b'Target="worksheets/sheet1.xml"', b'Target="../../evil.xml"'
        )
        refusal = expect_refusal(
            parse_registry(tmp_path, standard_parts(_workbook(), workbook_rels=rels)),
            "unsafe_archive_path",
        )
        assert refusal.details["target"] == "../../evil.xml"

    def test_out_of_range_active_tab(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_registry(tmp_path, standard_parts(_workbook(active_tab="1"))),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "activeTab out of range"

    def test_non_integer_active_tab(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_registry(tmp_path, standard_parts(_workbook(active_tab="first"))),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "invalid activeTab"

    def test_active_tab_defaults_to_zero(self, tmp_path: Path) -> None:
        registry = expect_registry(
            parse_registry(tmp_path, standard_parts(_workbook(active_tab=None)))
        )
        assert registry.active_tab == 0


class TestRegistryObjects:
    def test_entries_are_frozen(self) -> None:
        sheet = SheetEntry(
            name="S",
            sheet_id=1,
            rel_id="rId1",
            state="visible",
            tab_index=0,
            kind="worksheet",
            part_path="xl/worksheets/sheet1.xml",
        )
        with pytest.raises(AttributeError):
            sheet.name = "other"  # type: ignore[misc]
        name = DefinedName(name="N", text="A1", scope_sheet=None, is_hidden=False)
        with pytest.raises(AttributeError):
            name.is_hidden = True  # type: ignore[misc]
        assert sheet.kind == "worksheet"
