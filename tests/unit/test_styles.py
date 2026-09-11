"""Gates for number-format classification and styles-part parsing."""

from __future__ import annotations

import json
import zipfile
from collections.abc import Mapping
from pathlib import Path

from xlayer._errors import Refusal
from xlayer._ooxml.archive import WorkbookArchive
from xlayer._ooxml.package import PackageInfo, parse_package
from xlayer._ooxml.styles import (
    StyleInfo,
    classify_number_format,
    parse_styles,
    split_format_sections,
)

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
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
  <Override PartName="/xl/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
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
  <Relationship Id="rId2" Type="{_REL}styles" Target="styles.xml"/>
</Relationships>
""".encode()


def make_archive(tmp_path: Path, parts: Mapping[str, bytes]) -> WorkbookArchive:
    path = tmp_path / "book.xlsx"
    with zipfile.ZipFile(path, "w") as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    archive = WorkbookArchive.load(path)
    assert isinstance(archive, WorkbookArchive)
    return archive


def standard_parts(
    styles: bytes,
    *,
    workbook_rels: bytes = _DEFAULT_RELS,
) -> dict[str, bytes]:
    return {
        "[Content_Types].xml": _CONTENT_TYPES,
        "_rels/.rels": _PACKAGE_RELS,
        "xl/workbook.xml": _WORKBOOK,
        "xl/_rels/workbook.xml.rels": workbook_rels,
        "xl/worksheets/sheet1.xml": _MINIMAL_SHEET,
        "xl/styles.xml": styles,
    }


def parse_style_table(
    tmp_path: Path, parts: Mapping[str, bytes]
) -> tuple[StyleInfo, ...] | Refusal:
    archive = make_archive(tmp_path, parts)
    package = parse_package(archive)
    assert isinstance(package, PackageInfo), f"package refused: {package!r}"
    return parse_styles(archive, package)


def expect_styles(result: tuple[StyleInfo, ...] | Refusal) -> tuple[StyleInfo, ...]:
    assert isinstance(result, tuple), f"expected styles, got {result!r}"
    return result


def expect_refusal(result: tuple[StyleInfo, ...] | Refusal, code: str) -> Refusal:
    assert isinstance(result, Refusal), f"expected refusal {code!r}, got {result!r}"
    assert result.code == code
    assert len(result.recovery_options) > 0
    return result


def styles_xml(*, num_fmts: str = "", xfs: str = '<xf numFmtId="0"/>') -> bytes:
    fmts = f"<numFmts>{num_fmts}</numFmts>" if num_fmts else ""
    return (
        f'<?xml version="1.0" encoding="UTF-8"?>'
        f'<styleSheet xmlns="{_NS}">{fmts}<cellXfs>{xfs}</cellXfs></styleSheet>'
    ).encode()


class TestClassifyNumberFormat:
    """Independently justified answers — not a copy of the classifier."""

    def test_quoted_elapsed_token_is_literal_text(self) -> None:
        # 0 "[h]" displays the characters [h] after a zero. The brackets are
        # quoted, so they are not an elapsed-time token.
        assert classify_number_format(164, '0 "[h]"') == "non_temporal"

    def test_month_token_plus_time_is_date_time(self) -> None:
        # mmm is a month name. h:mm is a clock time. Together they are date-time.
        assert classify_number_format(164, "mmm h:mm") == "date-time"

    def test_unquoted_elapsed_hours_are_time(self) -> None:
        assert classify_number_format(46, "[h]:mm:ss") == "time"

    def test_minutes_adjacent_to_hours_are_time(self) -> None:
        assert classify_number_format(20, "h:mm") == "time"

    def test_minutes_adjacent_to_seconds_are_time(self) -> None:
        assert classify_number_format(45, "mm:ss") == "time"

    def test_month_without_time_context_is_date(self) -> None:
        assert classify_number_format(17, "mmm-yy") == "date"

    def test_mixed_conditional_sections_are_unknown(self) -> None:
        # First branch is a number; second is a date. Neither is the format.
        assert classify_number_format(164, "[>=45000]0;yyyy-mm-dd") == "unknown"

    def test_agreeing_sections_keep_the_shared_kind(self) -> None:
        assert classify_number_format(164, "yyyy-mm-dd;yyyy-mm-dd") == "date"
        assert classify_number_format(164, "0.00;(0.00)") == "non_temporal"

    def test_trailing_empty_section_does_not_create_disagreement(self) -> None:
        assert classify_number_format(164, "yyyy-mm-dd;") == "date"

    def test_locale_prefix_does_not_hide_a_date(self) -> None:
        assert classify_number_format(171, "[$-409]DD/MM/YYYY") == "date"

    def test_general_and_text_are_non_temporal(self) -> None:
        assert classify_number_format(0, "General") == "non_temporal"
        assert classify_number_format(49, "@") == "non_temporal"

    def test_builtin_ids_without_code(self) -> None:
        assert classify_number_format(14, None) == "date"
        assert classify_number_format(18, None) == "time"
        assert classify_number_format(22, None) == "date-time"
        assert classify_number_format(1, None) == "non_temporal"

    def test_excluded_locale_id_without_code_is_unknown(self) -> None:
        assert classify_number_format(27, None) == "unknown"
        assert classify_number_format(50, None) == "unknown"

    def test_unresolved_format_id_is_unknown(self) -> None:
        assert classify_number_format(99, None) == "unknown"

    def test_custom_override_of_builtin_id_uses_the_code(self) -> None:
        # A numFmt entry can replace the built-in meaning of ID 14.
        assert classify_number_format(14, "0.00") == "non_temporal"

    def test_locale_id_with_explicit_code_is_classified(self) -> None:
        assert classify_number_format(27, "yyyy-mm-dd") == "date"


class TestElapsedTimeContext:
    """Defect 1: an elapsed-hour token must remain hour context for a following mm.

    Microsoft documents ``[h]:mm`` as elapsed hours and minutes. Dropping the
    bracket and re-reading ``mm`` as a month produced ``date-time``.
    """

    def test_elapsed_hours_then_minutes_is_time(self) -> None:
        assert classify_number_format(164, "[h]:mm") == "time"

    def test_elapsed_double_hours_then_minutes_is_time(self) -> None:
        assert classify_number_format(164, "[hh]:mm") == "time"

    def test_elapsed_hours_minutes_seconds_is_time(self) -> None:
        assert classify_number_format(164, "[h]:mm:ss") == "time"

    def test_elapsed_minutes_then_seconds_is_time(self) -> None:
        assert classify_number_format(164, "[mm]:ss") == "time"

    def test_month_plus_clock_time_is_still_date_time(self) -> None:
        assert classify_number_format(164, "mmm h:mm") == "date-time"

    def test_quoted_elapsed_token_is_still_literal(self) -> None:
        assert classify_number_format(164, '0 "[h]"') == "non_temporal"

    def test_parser_reports_elapsed_format_as_time(self, tmp_path: Path) -> None:
        xml = styles_xml(
            num_fmts='<numFmt numFmtId="164" formatCode="[h]:mm"/>',
            xfs='<xf numFmtId="164"/>',
        )
        styles = expect_styles(parse_style_table(tmp_path, standard_parts(xml)))
        assert styles == (
            StyleInfo(
                format_id=164,
                format_code="[h]:mm",
                is_date=False,
                is_time=True,
                temporal_kind="time",
            ),
        )


class TestFormatEscapes:
    """Defect 2: a backslash makes the next character literal.

    ``\\"`` displays a double quote; it does not open a quoted block. Treating
    it as one swallowed the date tokens that followed.
    """

    def test_escaped_quote_before_date_tokens_is_date(self) -> None:
        assert classify_number_format(164, r"\"yyyy-mm-dd") == "date"

    def test_escaped_quote_before_time_tokens_is_time(self) -> None:
        assert classify_number_format(164, r"\"h:mm") == "time"

    def test_quoted_date_looking_text_is_not_a_date(self) -> None:
        assert classify_number_format(164, '"yyyy-mm-dd" 0') == "non_temporal"

    def test_quoted_elapsed_bracket_is_not_elapsed_time(self) -> None:
        assert classify_number_format(164, '"[h]" 0') == "non_temporal"

    def test_genuine_elapsed_token_still_works(self) -> None:
        assert classify_number_format(164, "[h]:mm:ss") == "time"

    def test_mixed_numeric_and_date_sections_still_unknown(self) -> None:
        assert classify_number_format(164, "[>=45000]0;yyyy-mm-dd") == "unknown"

    def test_escaped_backslash_then_semicolon_splits(self) -> None:
        # "\\" is an escaped backslash (literal), so the following ";" is a
        # real separator.
        assert split_format_sections(r"0\\;yyyy") == ["0\\\\", "yyyy"]

    def test_escaped_semicolon_is_not_a_separator(self) -> None:
        assert split_format_sections(r"0\;0;yyyy") == ["0\\;0", "yyyy"]

    def test_quoted_semicolon_is_not_a_separator(self) -> None:
        assert split_format_sections('"a;b";0') == ['"a;b"', "0"]

    def test_escaped_quote_does_not_open_a_quoted_block(self) -> None:
        # The \" is literal, so the ";" after it is a real separator.
        assert split_format_sections(r"\"x;yyyy") == ['\\"x', "yyyy"]

    def test_bracket_metadata_separates_normally(self) -> None:
        assert split_format_sections("[Red]0;[Blue]0") == ["[Red]0", "[Blue]0"]


class TestParseStyles:
    def test_cell_xfs_index_and_temporal_kind(self, tmp_path: Path) -> None:
        xml = styles_xml(
            num_fmts='<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>'
            '<numFmt numFmtId="165" formatCode="hh:mm:ss"/>'
            '<numFmt numFmtId="166" formatCode="yyyy-mm-dd hh:mm:ss"/>',
            xfs='<xf numFmtId="0"/><xf numFmtId="164"/><xf numFmtId="165"/><xf numFmtId="166"/>',
        )
        styles = expect_styles(parse_style_table(tmp_path, standard_parts(xml)))
        assert [item.temporal_kind for item in styles] == [
            "non_temporal",
            "date",
            "time",
            "date-time",
        ]
        assert styles[1].is_date is True and styles[1].is_time is False
        assert styles[2].is_date is False and styles[2].is_time is True
        assert styles[3].is_date is True and styles[3].is_time is True
        assert styles[0].is_date is False and styles[0].is_time is False

    def test_unknown_kind_does_not_claim_date_or_time(self, tmp_path: Path) -> None:
        xml = styles_xml(xfs='<xf numFmtId="27"/>')
        styles = expect_styles(parse_style_table(tmp_path, standard_parts(xml)))
        assert styles[0].temporal_kind == "unknown"
        assert styles[0].is_date is False
        assert styles[0].is_time is False

    def test_absent_relationship_defaults_to_general(self, tmp_path: Path) -> None:
        rels = f"""<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{_REL}worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
""".encode()
        parts = standard_parts(styles_xml(), workbook_rels=rels)
        del parts["xl/styles.xml"]
        styles = expect_styles(parse_style_table(tmp_path, parts))
        assert styles == (
            StyleInfo(
                format_id=0,
                format_code="General",
                is_date=False,
                is_time=False,
                temporal_kind="non_temporal",
            ),
        )

    def test_external_relationship_is_refused(self, tmp_path: Path) -> None:
        rels = f"""<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{_REL}worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="{_REL}styles"
    Target="https://example.com/styles.xml" TargetMode="External"/>
</Relationships>
""".encode()
        refusal = expect_refusal(
            parse_style_table(tmp_path, standard_parts(styles_xml(), workbook_rels=rels)),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "styles relationship is external"

    def test_unresolvable_relationship_is_refused(self, tmp_path: Path) -> None:
        rels = f"""<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{_REL}worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="{_REL}styles" Target="../../evil.xml"/>
</Relationships>
""".encode()
        refusal = expect_refusal(
            parse_style_table(tmp_path, standard_parts(styles_xml(), workbook_rels=rels)),
            "unsafe_archive_path",
        )
        assert refusal.details["target"] == "../../evil.xml"

    def test_missing_part_is_refused(self, tmp_path: Path) -> None:
        parts = standard_parts(styles_xml())
        del parts["xl/styles.xml"]
        refusal = expect_refusal(parse_style_table(tmp_path, parts), "missing_required_part")
        assert refusal.details["part"] == "xl/styles.xml"

    def test_unexpected_root_is_refused(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_style_table(
                tmp_path,
                standard_parts(f'<?xml version="1.0"?><workbook xmlns="{_NS}"/>'.encode()),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "unexpected root element"

    def test_missing_format_code_is_refused(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_style_table(
                tmp_path, standard_parts(styles_xml(num_fmts='<numFmt numFmtId="164"/>'))
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "numFmt missing formatCode"

    def test_duplicate_num_fmt_id_is_refused(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_style_table(
                tmp_path,
                standard_parts(
                    styles_xml(
                        num_fmts='<numFmt numFmtId="164" formatCode="0"/>'
                        '<numFmt numFmtId="164" formatCode="0.00"/>'
                    )
                ),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "duplicate numFmtId"

    def test_malformed_num_fmt_id_is_refused(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_style_table(
                tmp_path,
                standard_parts(
                    styles_xml(xfs='<xf numFmtId="²"/>'),
                ),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "invalid numFmtId"

    def test_missing_cell_xfs_defaults_to_general(self, tmp_path: Path) -> None:
        xml = f'<?xml version="1.0"?><styleSheet xmlns="{_NS}"/>'.encode()
        styles = expect_styles(parse_style_table(tmp_path, standard_parts(xml)))
        assert styles[0].temporal_kind == "non_temporal"
        assert styles[0].format_code == "General"


class TestStyleFixtures:
    def test_1904_fixture_custom_date_time_kinds(self) -> None:
        archive = WorkbookArchive.load(FIXTURES / "test_workbook_6b_1904.xlsx")
        assert isinstance(archive, WorkbookArchive)
        package = parse_package(archive)
        assert isinstance(package, PackageInfo)
        styles = parse_styles(archive, package)
        assert isinstance(styles, tuple)
        expected = json.loads(
            (FIXTURES / "test_workbook_6b_1904.styles.json").read_text(encoding="utf-8")
        )
        assert [
            {
                "index": index,
                "format_id": item.format_id,
                "format_code": item.format_code,
                "temporal_kind": item.temporal_kind,
            }
            for index, item in enumerate(styles)
        ] == expected
        archive.close()

    def test_formats_fixture_selected_indexes(self) -> None:
        archive = WorkbookArchive.load(FIXTURES / "test_workbook_3_formats.xlsx")
        assert isinstance(archive, WorkbookArchive)
        package = parse_package(archive)
        assert isinstance(package, PackageInfo)
        styles = parse_styles(archive, package)
        assert isinstance(styles, tuple)
        expected = json.loads(
            (FIXTURES / "test_workbook_3_formats.styles.json").read_text(encoding="utf-8")
        )
        assert len(styles) == expected["cell_xf_count"]
        for row in expected["indexes"]:
            item = styles[row["index"]]
            assert item.format_id == row["format_id"]
            assert item.format_code == row["format_code"]
            assert item.temporal_kind == row["temporal_kind"]
        archive.close()
