"""Gates for OOXML package resolution against crafted archives."""

from __future__ import annotations

import zipfile
from collections.abc import Mapping
from pathlib import Path
from typing import cast

import pytest

from xlayer._errors import Refusal
from xlayer._ooxml.archive import WorkbookArchive
from xlayer._ooxml.package import (
    SUPPORTED_WORKBOOK_CONTENT_TYPE,
    PackageInfo,
    parse_package,
)

_CONTENT_TYPES = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels"
    ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
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

_WORKBOOK = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
"""

_WORKBOOK_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId9"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="https://example.com/report" TargetMode="External"/>
</Relationships>
"""

_MINIMAL_PARTS: dict[str, bytes] = {
    "[Content_Types].xml": _CONTENT_TYPES,
    "_rels/.rels": _PACKAGE_RELS,
    "xl/workbook.xml": _WORKBOOK,
    "xl/_rels/workbook.xml.rels": _WORKBOOK_RELS,
}


def make_archive(tmp_path: Path, parts: Mapping[str, bytes]) -> WorkbookArchive:
    path = tmp_path / "book.xlsx"
    with zipfile.ZipFile(path, "w") as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    archive = WorkbookArchive.load(path)
    assert isinstance(archive, WorkbookArchive)
    return archive


def expect_package(result: PackageInfo | Refusal) -> PackageInfo:
    assert isinstance(result, PackageInfo), f"expected PackageInfo, got refusal: {result!r}"
    return result


def expect_refusal(result: PackageInfo | Refusal, code: str) -> Refusal:
    assert isinstance(result, Refusal), f"expected refusal {code!r}, got {result!r}"
    assert result.code == code, f"expected {code!r}, got {result.code!r}: {result.message}"
    # Contract: every refusal carries at least one structured recovery option.
    assert len(result.recovery_options) > 0, f"refusal {code!r} has no recovery options"
    return result


def parts_with_workbook_content_type(content_type: bytes) -> dict[str, bytes]:
    parts = dict(_MINIMAL_PARTS)
    parts["[Content_Types].xml"] = _CONTENT_TYPES.replace(
        b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        content_type,
    )
    return parts


class TestHappyPath:
    def test_minimal_package_resolves(self, tmp_path: Path) -> None:
        package = expect_package(parse_package(make_archive(tmp_path, _MINIMAL_PARTS)))
        assert package.workbook_path == "xl/workbook.xml"
        assert package.workbook_dir == "xl"
        assert package.content_type_of("xl/workbook.xml") == SUPPORTED_WORKBOOK_CONTENT_TYPE
        assert package.default_types["rels"].endswith("relationships+xml")

    def test_resolves_rel_by_supported_kind(self, tmp_path: Path) -> None:
        package = expect_package(parse_package(make_archive(tmp_path, _MINIMAL_PARTS)))
        assert package.resolve_workbook_rel("styles") == "xl/styles.xml"
        assert package.resolve_workbook_rel("sharedStrings") is None

    def test_find_workbook_rel_distinguishes_absence_from_unusable(self, tmp_path: Path) -> None:
        package = expect_package(parse_package(make_archive(tmp_path, _MINIMAL_PARTS)))
        styles = package.find_workbook_rel("styles")
        assert styles is not None
        assert styles.target == "styles.xml"
        assert package.find_workbook_rel("sharedStrings") is None

        parts = dict(_MINIMAL_PARTS)
        parts["xl/_rels/workbook.xml.rels"] = _WORKBOOK_RELS.replace(
            b'Target="styles.xml"', b'Target="../../evil.xml"'
        )
        broken_dir = tmp_path / "broken"
        broken_dir.mkdir()
        broken = expect_package(parse_package(make_archive(broken_dir, parts)))
        found = broken.find_workbook_rel("styles")
        assert found is not None
        assert found.target == "../../evil.xml"
        assert broken.resolve_workbook_rel("styles") is None

    def test_unknown_rel_kind_is_a_programming_error(self, tmp_path: Path) -> None:
        package = expect_package(parse_package(make_archive(tmp_path, _MINIMAL_PARTS)))
        with pytest.raises(KeyError):
            package.resolve_workbook_rel("hyperlink")

    def test_resolves_rel_by_id_with_expected_kind(self, tmp_path: Path) -> None:
        package = expect_package(parse_package(make_archive(tmp_path, _MINIMAL_PARTS)))
        assert package.resolve_workbook_rel_by_id("rId1", "worksheet") == (
            "xl/worksheets/sheet1.xml"
        )
        assert package.resolve_workbook_rel_by_id("rId99", "worksheet") is None

    def test_resolve_by_id_rejects_kind_mismatch(self, tmp_path: Path) -> None:
        # rId1 is a worksheet relationship; asking for it as styles must fail
        # rather than resolve, so a crafted rels part cannot redirect lookups.
        package = expect_package(parse_package(make_archive(tmp_path, _MINIMAL_PARTS)))
        assert package.resolve_workbook_rel_by_id("rId1", "styles") is None

    def test_package_mappings_are_read_only(self, tmp_path: Path) -> None:
        package = expect_package(parse_package(make_archive(tmp_path, _MINIMAL_PARTS)))
        mutable = cast("dict[str, str]", package.content_types)
        with pytest.raises(TypeError):
            mutable["/evil.xml"] = "application/xml"

    def test_absolute_office_document_target(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["_rels/.rels"] = _PACKAGE_RELS.replace(
            b'Target="xl/workbook.xml"', b'Target="/xl/workbook.xml"'
        )
        package = expect_package(parse_package(make_archive(tmp_path, parts)))
        assert package.workbook_path == "xl/workbook.xml"

    def test_missing_workbook_rels_is_tolerated(self, tmp_path: Path) -> None:
        parts = {k: v for k, v in _MINIMAL_PARTS.items() if k != "xl/_rels/workbook.xml.rels"}
        package = expect_package(parse_package(make_archive(tmp_path, parts)))
        assert package.workbook_rels == {}
        assert package.resolve_workbook_rel("styles") is None


class TestExternalRelationships:
    def test_external_rel_is_preserved_as_metadata(self, tmp_path: Path) -> None:
        package = expect_package(parse_package(make_archive(tmp_path, _MINIMAL_PARTS)))
        hyperlink = package.workbook_rels["rId9"]
        assert hyperlink.target_mode == "External"
        assert hyperlink.target == "https://example.com/report"

    def test_external_rel_never_resolves_to_a_part(self, tmp_path: Path) -> None:
        package = expect_package(parse_package(make_archive(tmp_path, _MINIMAL_PARTS)))
        assert package.resolve_workbook_rel_by_id("rId9", "worksheet") is None

    def test_external_office_document_is_refused(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["_rels/.rels"] = _PACKAGE_RELS.replace(
            b'Target="xl/workbook.xml"',
            b'Target="https://example.com/wb.xlsx" TargetMode="External"',
        )
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "invalid_part_content"
        )
        assert refusal.details["reason"] == "officeDocument relationship is external"

    def test_exact_type_matching_rejects_lookalike_uri(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["xl/_rels/workbook.xml.rels"] = _WORKBOOK_RELS.replace(
            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            b"http://evil.example/officeDocument/2006/relationships/styles",
        )
        package = expect_package(parse_package(make_archive(tmp_path, parts)))
        assert package.resolve_workbook_rel("styles") is None

    def test_scheme_target_rejected_before_base_joining(self, tmp_path: Path) -> None:
        # An Internal-mode target carrying a URI scheme must be rejected on
        # the raw target; joining "xl/" first would mask the scheme check.
        parts = dict(_MINIMAL_PARTS)
        parts["xl/_rels/workbook.xml.rels"] = _WORKBOOK_RELS.replace(
            b'Target="styles.xml"', b'Target="http://evil.example/styles.xml"'
        )
        package = expect_package(parse_package(make_archive(tmp_path, parts)))
        assert package.resolve_workbook_rel("styles") is None

    def test_authority_target_rejected_before_base_joining(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["xl/_rels/workbook.xml.rels"] = _WORKBOOK_RELS.replace(
            b'Target="styles.xml"', b'Target="//host/styles.xml"'
        )
        package = expect_package(parse_package(make_archive(tmp_path, parts)))
        assert package.resolve_workbook_rel("styles") is None

    def test_encoded_query_target_rejected_after_decoding(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["xl/_rels/workbook.xml.rels"] = _WORKBOOK_RELS.replace(
            b'Target="styles.xml"', b'Target="styles.xml%3Fx=1"'
        )
        package = expect_package(parse_package(make_archive(tmp_path, parts)))
        assert package.resolve_workbook_rel("styles") is None


class TestPackageKindValidation:
    def test_word_document_refused(self, tmp_path: Path) -> None:
        parts = parts_with_workbook_content_type(
            b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        )
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "unsupported_package_kind"
        )
        assert refusal.details["kind"] == "word_document"

    def test_presentation_refused(self, tmp_path: Path) -> None:
        parts = parts_with_workbook_content_type(
            b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
        )
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "unsupported_package_kind"
        )
        assert refusal.details["kind"] == "presentation"

    def test_macro_enabled_workbook_refused_with_recovery(self, tmp_path: Path) -> None:
        parts = parts_with_workbook_content_type(
            b"application/vnd.ms-excel.sheet.macroEnabled.main+xml"
        )
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "unsupported_package_kind"
        )
        assert refusal.details["kind"] == "macro_enabled_workbook"
        assert len(refusal.recovery_options) > 0

    def test_spreadsheet_template_refused(self, tmp_path: Path) -> None:
        parts = parts_with_workbook_content_type(
            b"application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
        )
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "unsupported_package_kind"
        )
        assert refusal.details["kind"] == "spreadsheet_template"

    def test_unknown_content_type_refused(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        # Drop the workbook Override so lookup falls back to the xml Default.
        parts["[Content_Types].xml"] = _CONTENT_TYPES.replace(
            b'  <Override PartName="/xl/workbook.xml"\n'
            b'    ContentType="application/vnd.openxmlformats-officedocument'
            b'.spreadsheetml.sheet.main+xml"/>\n',
            b"",
        )
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "unsupported_package_kind"
        )
        assert refusal.details["kind"] == "unknown"
        assert refusal.details["content_type"] == "application/xml"

    def test_strict_ooxml_refused(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["_rels/.rels"] = _PACKAGE_RELS.replace(
            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            b"http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
        )
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "unsupported_package_kind"
        )
        assert refusal.details["kind"] == "strict_ooxml_workbook"


class TestSchemaRefusals:
    def test_wrong_root_namespace_in_content_types(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["[Content_Types].xml"] = _CONTENT_TYPES.replace(
            b"http://schemas.openxmlformats.org/package/2006/content-types",
            b"http://evil.example/content-types",
        )
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "invalid_part_content"
        )
        assert refusal.details["reason"] == "unexpected root element"

    def test_wrong_root_namespace_in_rels(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["_rels/.rels"] = _PACKAGE_RELS.replace(
            b"http://schemas.openxmlformats.org/package/2006/relationships",
            b"http://evil.example/relationships",
        )
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "invalid_part_content"
        )
        assert refusal.details["reason"] == "unexpected root element"

    def test_duplicate_override_part_name(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        duplicate = (
            b'  <Override PartName="/xl/styles.xml"\n    ContentType="application/xml"/>\n</Types>'
        )
        parts["[Content_Types].xml"] = _CONTENT_TYPES.replace(b"</Types>", duplicate)
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "invalid_part_content"
        )
        assert refusal.details["reason"] == "duplicate Override PartName"

    def test_duplicate_default_extension(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        duplicate = b'  <Default Extension="xml" ContentType="text/xml"/>\n</Types>'
        parts["[Content_Types].xml"] = _CONTENT_TYPES.replace(b"</Types>", duplicate)
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "invalid_part_content"
        )
        assert refusal.details["reason"] == "duplicate Default Extension"

    def test_multiple_office_document_relationships(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        second = (
            b'  <Relationship Id="rId2"\n'
            b'    Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            b'relationships/officeDocument"\n'
            b'    Target="xl/workbook.xml"/>\n</Relationships>'
        )
        parts["_rels/.rels"] = _PACKAGE_RELS.replace(b"</Relationships>", second)
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "invalid_part_content"
        )
        assert refusal.details["reason"] == "multiple officeDocument relationships"

    def test_mixed_transitional_and_strict_office_document(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        strict = (
            b'  <Relationship Id="rId2"\n'
            b'    Type="http://purl.oclc.org/ooxml/officeDocument/relationships/'
            b'officeDocument"\n'
            b'    Target="xl/workbook.xml"/>\n</Relationships>'
        )
        parts["_rels/.rels"] = _PACKAGE_RELS.replace(b"</Relationships>", strict)
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "invalid_part_content"
        )
        assert refusal.details["reason"] == (
            "mixed transitional and strict officeDocument relationships"
        )

    def test_invalid_target_mode(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["xl/_rels/workbook.xml.rels"] = _WORKBOOK_RELS.replace(
            b'TargetMode="External"', b'TargetMode="Both"'
        )
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "invalid_part_content"
        )
        assert refusal.details["reason"] == "invalid TargetMode"
        assert refusal.details["target_mode"] == "Both"

    def test_empty_relationship_id(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["xl/_rels/workbook.xml.rels"] = _WORKBOOK_RELS.replace(b'Id="rId2"', b'Id=""')
        expect_refusal(parse_package(make_archive(tmp_path, parts)), "invalid_part_content")

    @pytest.mark.parametrize(
        ("kind", "target"),
        [
            ("styles", "styles.xml"),
            ("sharedStrings", "sharedStrings.xml"),
            ("calcChain", "calcChain.xml"),
            ("theme", "theme/theme1.xml"),
        ],
    )
    def test_duplicate_singleton_relationship_rejected(
        self, tmp_path: Path, kind: str, target: str
    ) -> None:
        type_uri = f"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}"
        extra = (
            f'  <Relationship Id="rId7" Type="{type_uri}" Target="{target}"/>\n'
            f'  <Relationship Id="rId8" Type="{type_uri}" Target="{target}"/>\n'
            f"</Relationships>"
        ).encode()
        parts = dict(_MINIMAL_PARTS)
        parts["xl/_rels/workbook.xml.rels"] = _WORKBOOK_RELS.replace(b"</Relationships>", extra)
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "invalid_part_content"
        )
        assert refusal.details["reason"] == f"duplicate {kind} relationship"
        assert refusal.details["kind"] == kind

    def test_duplicate_relationship_id(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["xl/_rels/workbook.xml.rels"] = _WORKBOOK_RELS.replace(b'Id="rId2"', b'Id="rId1"')
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "invalid_part_content"
        )
        assert refusal.details["reason"] == "duplicate relationship ID"
        assert refusal.details["rel_id"] == "rId1"

    def test_relationship_missing_target(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["xl/_rels/workbook.xml.rels"] = _WORKBOOK_RELS.replace(b'Target="styles.xml"', b"")
        expect_refusal(parse_package(make_archive(tmp_path, parts)), "invalid_part_content")

    def test_override_missing_part_name(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["[Content_Types].xml"] = _CONTENT_TYPES.replace(b'PartName="/xl/styles.xml"', b"")
        expect_refusal(parse_package(make_archive(tmp_path, parts)), "invalid_part_content")


class TestMissingPartRefusals:
    def test_missing_content_types(self, tmp_path: Path) -> None:
        parts = {k: v for k, v in _MINIMAL_PARTS.items() if k != "[Content_Types].xml"}
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "missing_required_part"
        )
        assert refusal.details["part"] == "[Content_Types].xml"

    def test_missing_package_rels(self, tmp_path: Path) -> None:
        parts = {k: v for k, v in _MINIMAL_PARTS.items() if k != "_rels/.rels"}
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "missing_required_part"
        )
        assert refusal.details["part"] == "_rels/.rels"

    def test_no_office_document_relationship(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["_rels/.rels"] = _PACKAGE_RELS.replace(
            b"officeDocument/2006/relationships/officeDocument",
            b"officeDocument/2006/relationships/somethingElse",
        )
        expect_refusal(parse_package(make_archive(tmp_path, parts)), "missing_required_part")

    def test_workbook_part_absent(self, tmp_path: Path) -> None:
        parts = {k: v for k, v in _MINIMAL_PARTS.items() if k != "xl/workbook.xml"}
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "missing_required_part"
        )
        assert refusal.details["part"] == "xl/workbook.xml"

    def test_escaping_office_document_target(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["_rels/.rels"] = _PACKAGE_RELS.replace(
            b'Target="xl/workbook.xml"', b'Target="../evil.xml"'
        )
        refusal = expect_refusal(
            parse_package(make_archive(tmp_path, parts)), "unsafe_archive_path"
        )
        assert refusal.details["kind"] == "relationship_target"

    def test_escaping_workbook_rel_target_resolves_to_none(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["xl/_rels/workbook.xml.rels"] = _WORKBOOK_RELS.replace(
            b'Target="styles.xml"', b'Target="../../styles.xml"'
        )
        package = expect_package(parse_package(make_archive(tmp_path, parts)))
        assert package.resolve_workbook_rel("styles") is None

    def test_malformed_content_types(self, tmp_path: Path) -> None:
        parts = dict(_MINIMAL_PARTS)
        parts["[Content_Types].xml"] = b"<Types><broken"
        expect_refusal(parse_package(make_archive(tmp_path, parts)), "malformed_xml")
