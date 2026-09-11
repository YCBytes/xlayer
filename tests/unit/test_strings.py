"""Gates for shared-string decoding against crafted archives and cleaned fixtures."""

from __future__ import annotations

import json
import zipfile
from collections.abc import Mapping
from pathlib import Path

from xlayer._errors import Refusal
from xlayer._ooxml.archive import WorkbookArchive
from xlayer._ooxml.package import PackageInfo, parse_package
from xlayer._ooxml.strings import parse_shared_strings

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
  <Override PartName="/xl/sharedStrings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
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
  <Relationship Id="rId2" Type="{_REL}sharedStrings" Target="sharedStrings.xml"/>
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
    sst: bytes,
    *,
    workbook_rels: bytes = _DEFAULT_RELS,
    extra: Mapping[str, bytes] | None = None,
) -> dict[str, bytes]:
    parts = {
        "[Content_Types].xml": _CONTENT_TYPES,
        "_rels/.rels": _PACKAGE_RELS,
        "xl/workbook.xml": _WORKBOOK,
        "xl/_rels/workbook.xml.rels": workbook_rels,
        "xl/worksheets/sheet1.xml": _MINIMAL_SHEET,
        "xl/sharedStrings.xml": sst,
    }
    if extra:
        parts.update(extra)
    return parts


def parse_strings(tmp_path: Path, parts: Mapping[str, bytes]) -> tuple[str, ...] | Refusal:
    archive = make_archive(tmp_path, parts)
    package = parse_package(archive)
    assert isinstance(package, PackageInfo), f"package refused: {package!r}"
    return parse_shared_strings(archive, package)


def expect_strings(result: tuple[str, ...] | Refusal) -> tuple[str, ...]:
    assert isinstance(result, tuple), f"expected strings, got {result!r}"
    return result


def expect_refusal(result: tuple[str, ...] | Refusal, code: str) -> Refusal:
    assert isinstance(result, Refusal), f"expected refusal {code!r}, got {result!r}"
    assert result.code == code
    assert len(result.recovery_options) > 0
    return result


def sst(*items: str, root_attrs: str = "") -> bytes:
    body = "".join(items)
    attrs = f" {root_attrs}" if root_attrs else ""
    return (
        f'<?xml version="1.0" encoding="UTF-8"?><sst xmlns="{_NS}"{attrs}>{body}</sst>'
    ).encode()


class TestSharedStringDecoding:
    def test_simple_and_rich_text_concatenate(self, tmp_path: Path) -> None:
        xml = sst(
            "<si><t>foo</t></si>",
            "<si><r><t>Hel</t></r><r><t>lo</t></r></si>",
        )
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == (
            "foo",
            "Hello",
        )

    def test_preserve_on_text_node_keeps_regular_spaces(self, tmp_path: Path) -> None:
        xml = sst('<si><t xml:space="preserve">  kept  </t></si>')
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("  kept  ",)

    def test_preserve_on_parent_is_inherited(self, tmp_path: Path) -> None:
        xml = sst('<si xml:space="preserve"><t>  from-si  </t></si>')
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("  from-si  ",)

    def test_preserve_on_run_is_inherited_by_text(self, tmp_path: Path) -> None:
        xml = sst('<si><r xml:space="preserve"><t>  from-run  </t></r></si>')
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("  from-run  ",)

    def test_child_default_overrides_parent_preserve(self, tmp_path: Path) -> None:
        xml = sst('<si xml:space="preserve"><t xml:space="default">  trimmed  </t></si>')
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("trimmed",)

    def test_default_whitespace_does_not_strip_nbsp(self, tmp_path: Path) -> None:
        xml = sst("<si><t>\u00a0hello\u00a0</t></si>")
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == (
            "\u00a0hello\u00a0",
        )

    def test_rich_text_preserve_boundary(self, tmp_path: Path) -> None:
        xml = sst('<si><r><t xml:space="preserve">Hello </t></r><r><t>World</t></r></si>')
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("Hello World",)

    def test_phonetic_guide_is_excluded(self, tmp_path: Path) -> None:
        xml = sst(
            "<si><r><t>漢字</t></r>"
            '<rPh sb="0" eb="2"><t>かんじ</t></rPh>'
            '<phoneticPr fontId="1"/></si>'
        )
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("漢字",)

    def test_excel_escaped_carriage_return(self, tmp_path: Path) -> None:
        xml = sst("<si><t>line_x000D_break</t></si>")
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("line\rbreak",)

    def test_protected_literal_escape_sequence(self, tmp_path: Path) -> None:
        xml = sst("<si><t>_x005F_x000D_</t></si>")
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("_x000D_",)

    def test_empty_si_is_empty_string(self, tmp_path: Path) -> None:
        xml = sst("<si/>")
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("",)


class TestWhitespaceInheritanceFromRoot:
    """Defect 3: xml:space is inherited from <sst>, not only from <si>.

    XML applies a whitespace instruction to all descendants until overridden.
    Inheritance that started at <si> ignored a preserve on the root.
    """

    def test_root_preserve_reaches_simple_text(self, tmp_path: Path) -> None:
        xml = sst("<si><t>  retained  </t></si>", root_attrs='xml:space="preserve"')
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("  retained  ",)

    def test_root_preserve_reaches_rich_text_runs(self, tmp_path: Path) -> None:
        xml = sst(
            "<si><r><t> a </t></r><r><t> b </t></r></si>",
            root_attrs='xml:space="preserve"',
        )
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == (" a  b ",)

    def test_item_default_overrides_root_preserve(self, tmp_path: Path) -> None:
        xml = sst(
            '<si xml:space="default"><t>  trimmed  </t></si>',
            root_attrs='xml:space="preserve"',
        )
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("trimmed",)

    def test_run_default_overrides_inherited_preserve(self, tmp_path: Path) -> None:
        xml = sst(
            '<si xml:space="preserve"><r xml:space="default"><t>  run  </t></r>'
            "<r><t>  kept  </t></r></si>"
        )
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("run  kept  ",)

    def test_text_default_overrides_inherited_run_preserve(self, tmp_path: Path) -> None:
        xml = sst('<si><r xml:space="preserve"><t xml:space="default">  t  </t></r></si>')
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("t",)

    def test_nbsp_survives_default_under_root_default(self, tmp_path: Path) -> None:
        xml = sst("<si><t>\u00a0x\u00a0</t></si>", root_attrs='xml:space="default"')
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("\u00a0x\u00a0",)


class TestExcelEscapeUnicode:
    """Defect 4: decoded text must be valid Unicode or a structured refusal.

    ``_xD83D__xDE00_`` is one character encoded as a UTF-16 pair; emitting two
    surrogate code points produced a string that cannot be UTF-8 encoded.
    """

    def test_carriage_return_escape(self, tmp_path: Path) -> None:
        xml = sst("<si><t>a_x000D_b</t></si>")
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("a\rb",)

    def test_protected_literal_is_not_decoded_twice(self, tmp_path: Path) -> None:
        xml = sst("<si><t>_x005F_x000D_</t></si>")
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("_x000D_",)

    def test_surrogate_pair_combines_into_one_character(self, tmp_path: Path) -> None:
        xml = sst("<si><t>_xD83D__xDE00_</t></si>")
        strings = expect_strings(parse_strings(tmp_path, standard_parts(xml)))
        assert strings == ("\U0001f600",)
        assert strings[0].encode("utf-8") == b"\xf0\x9f\x98\x80"

    def test_ordinary_astral_text_is_unchanged(self, tmp_path: Path) -> None:
        xml = sst("<si><t>\U0001f600 ok</t></si>")
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("\U0001f600 ok",)

    def test_lone_high_surrogate_is_refused(self, tmp_path: Path) -> None:
        xml = sst("<si><t>ok</t></si>", "<si><t>_xD800_</t></si>")
        refusal = expect_refusal(
            parse_strings(tmp_path, standard_parts(xml)), "invalid_part_content"
        )
        assert refusal.details["part"] == "xl/sharedStrings.xml"
        assert refusal.details["reason"] == "unpaired UTF-16 surrogate in Excel escape"
        assert refusal.details["string_index"] == 1

    def test_lone_low_surrogate_is_refused(self, tmp_path: Path) -> None:
        xml = sst("<si><t>_xDC00_</t></si>")
        refusal = expect_refusal(
            parse_strings(tmp_path, standard_parts(xml)), "invalid_part_content"
        )
        assert refusal.details["reason"] == "unpaired UTF-16 surrogate in Excel escape"
        assert refusal.details["string_index"] == 0

    def test_reversed_surrogate_pair_is_refused(self, tmp_path: Path) -> None:
        xml = sst("<si><t>_xDE00__xD83D_</t></si>")
        expect_refusal(parse_strings(tmp_path, standard_parts(xml)), "invalid_part_content")

    def test_high_surrogate_not_immediately_followed_is_refused(self, tmp_path: Path) -> None:
        xml = sst("<si><t>_xD83D_x_xDE00_</t></si>")
        expect_refusal(parse_strings(tmp_path, standard_parts(xml)), "invalid_part_content")

    def test_protected_surrogate_literals_stay_text(self, tmp_path: Path) -> None:
        xml = sst("<si><t>_x005F_xD83D__x005F_xDE00_</t></si>")
        assert expect_strings(parse_strings(tmp_path, standard_parts(xml))) == ("_xD83D__xDE00_",)


class TestSharedStringRelationships:
    def test_absent_relationship_is_empty_table(self, tmp_path: Path) -> None:
        rels = f"""<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{_REL}worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
""".encode()
        parts = standard_parts(sst("<si><t>x</t></si>"), workbook_rels=rels)
        del parts["xl/sharedStrings.xml"]
        assert expect_strings(parse_strings(tmp_path, parts)) == ()

    def test_external_relationship_is_refused(self, tmp_path: Path) -> None:
        rels = f"""<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{_REL}worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="{_REL}sharedStrings"
    Target="https://example.com/sst.xml" TargetMode="External"/>
</Relationships>
""".encode()
        parts = standard_parts(sst("<si><t>x</t></si>"), workbook_rels=rels)
        refusal = expect_refusal(parse_strings(tmp_path, parts), "invalid_part_content")
        assert refusal.details["reason"] == "sharedStrings relationship is external"

    def test_unresolvable_relationship_is_refused(self, tmp_path: Path) -> None:
        rels = f"""<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{_REL}worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="{_REL}sharedStrings" Target="../../evil.xml"/>
</Relationships>
""".encode()
        parts = standard_parts(sst("<si><t>x</t></si>"), workbook_rels=rels)
        refusal = expect_refusal(parse_strings(tmp_path, parts), "unsafe_archive_path")
        assert refusal.details["target"] == "../../evil.xml"

    def test_missing_part_is_refused(self, tmp_path: Path) -> None:
        parts = standard_parts(sst("<si><t>x</t></si>"))
        del parts["xl/sharedStrings.xml"]
        refusal = expect_refusal(parse_strings(tmp_path, parts), "missing_required_part")
        assert refusal.details["part"] == "xl/sharedStrings.xml"

    def test_unexpected_root_is_refused(self, tmp_path: Path) -> None:
        refusal = expect_refusal(
            parse_strings(
                tmp_path,
                standard_parts(f'<?xml version="1.0"?><workbook xmlns="{_NS}"/>'.encode()),
            ),
            "invalid_part_content",
        )
        assert refusal.details["reason"] == "unexpected root element"


class TestSharedStringFixtures:
    def test_registry_fixture_plain_strings(self) -> None:
        archive = WorkbookArchive.load(FIXTURES / "test_workbook_2_registry.xlsx")
        assert isinstance(archive, WorkbookArchive)
        package = parse_package(archive)
        assert isinstance(package, PackageInfo)
        strings = parse_shared_strings(archive, package)
        assert isinstance(strings, tuple)
        expected = json.loads(
            (FIXTURES / "test_workbook_2_registry.strings.json").read_text(encoding="utf-8")
        )
        assert list(strings) == expected
        archive.close()

    def test_formats_fixture_plain_rich_and_preserved(self) -> None:
        archive = WorkbookArchive.load(FIXTURES / "test_workbook_3_formats.xlsx")
        assert isinstance(archive, WorkbookArchive)
        package = parse_package(archive)
        assert isinstance(package, PackageInfo)
        strings = parse_shared_strings(archive, package)
        assert isinstance(strings, tuple)
        expected = json.loads(
            (FIXTURES / "test_workbook_3_formats.strings.json").read_text(encoding="utf-8")
        )
        assert list(strings) == expected
        archive.close()
