"""Gates for the safe archive layer.

Every refusal code in the Part 1A contract is exercised with a small crafted
archive; limits are lowered through ``ArchiveLimits`` rather than by creating
genuinely large files.
"""

from __future__ import annotations

import os
import sys
import warnings
import zipfile
from collections.abc import Mapping
from pathlib import Path

import pytest

from xlayer._errors import Refusal
from xlayer._ooxml.archive import (
    ArchiveLimits,
    WorkbookArchive,
    normalize_package_path,
    validate_opc_target,
)

_XML = b'<?xml version="1.0"?><root/>'


def build_zip(path: Path, entries: Mapping[str, bytes]) -> Path:
    with zipfile.ZipFile(path, "w") as zf:
        for name, data in entries.items():
            zf.writestr(name, data)
    return path


def load(path: Path, limits: ArchiveLimits | None = None) -> WorkbookArchive | Refusal:
    return WorkbookArchive.load(path, limits or ArchiveLimits())


def expect_refusal(result: WorkbookArchive | Refusal, code: str) -> Refusal:
    assert isinstance(result, Refusal), f"expected refusal {code!r}, got {result!r}"
    assert result.code == code, f"expected {code!r}, got {result.code!r}: {result.message}"
    # Contract: every refusal carries at least one structured recovery option.
    assert len(result.recovery_options) > 0, f"refusal {code!r} has no recovery options"
    return result


def expect_archive(result: WorkbookArchive | Refusal) -> WorkbookArchive:
    assert isinstance(result, WorkbookArchive), f"expected archive, got refusal: {result!r}"
    return result


class TestLoadHappyPath:
    def test_lists_parts_in_archive_order(self, tmp_path: Path) -> None:
        path = build_zip(tmp_path / "ok.xlsx", {"b.xml": _XML, "a.xml": _XML})
        archive = expect_archive(load(path))
        assert archive.part_names() == ("b.xml", "a.xml")
        assert archive.has_part("a.xml")
        assert not archive.has_part("missing.xml")
        archive.close()

    def test_read_part_returns_bytes_and_caches(self, tmp_path: Path) -> None:
        path = build_zip(tmp_path / "ok.xlsx", {"a.xml": _XML})
        archive = expect_archive(load(path))
        first = archive.read_part("a.xml")
        second = archive.read_part("a.xml")
        assert first == _XML
        assert first is second  # cached
        archive.close()

    def test_reads_deflated_entries(self, tmp_path: Path) -> None:
        payload = b"<r>" + b"<c/>" * 200 + b"</r>"
        path = tmp_path / "deflated.xlsx"
        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("a.xml", payload)
        archive = expect_archive(load(path))
        assert archive.read_part("a.xml") == payload
        archive.close()

    def test_deterministic_across_loads(self, tmp_path: Path) -> None:
        path = build_zip(tmp_path / "ok.xlsx", {"x/y.xml": _XML, "a.xml": _XML})
        one = expect_archive(load(path))
        two = expect_archive(load(path))
        assert one.part_names() == two.part_names()
        assert one.read_part("a.xml") == two.read_part("a.xml")
        one.close()
        two.close()

    def test_close_is_idempotent(self, tmp_path: Path) -> None:
        archive = expect_archive(load(build_zip(tmp_path / "ok.xlsx", {"a.xml": _XML})))
        archive.close()
        archive.close()
        assert archive.closed


class TestContainerRefusals:
    def test_missing_file(self, tmp_path: Path) -> None:
        expect_refusal(load(tmp_path / "absent.xlsx"), "invalid_path")

    def test_directory(self, tmp_path: Path) -> None:
        expect_refusal(load(tmp_path), "invalid_path")

    @pytest.mark.skipif(
        sys.platform == "win32" or os.geteuid() == 0,
        reason="permission bits are not enforced on Windows or for root",
    )
    def test_unreadable_file(self, tmp_path: Path) -> None:
        path = build_zip(tmp_path / "locked.xlsx", {"a.xml": _XML})
        path.chmod(0o000)
        try:
            expect_refusal(load(path), "unreadable_file")
        finally:
            path.chmod(0o644)

    def test_oversized_file(self, tmp_path: Path) -> None:
        path = build_zip(tmp_path / "big.xlsx", {"a.xml": _XML})
        refusal = expect_refusal(load(path, ArchiveLimits(max_file_bytes=10)), "archive_too_large")
        assert refusal.details["limit_bytes"] == 10

    def test_cfb_container(self, tmp_path: Path) -> None:
        path = tmp_path / "legacy.xls"
        path.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 64)
        refusal = expect_refusal(load(path), "unsupported_container")
        assert refusal.details["container"] == "cfb"
        assert len(refusal.recovery_options) > 0

    def test_not_a_zip(self, tmp_path: Path) -> None:
        path = tmp_path / "plain.txt"
        path.write_bytes(b"just text, no archive here")
        expect_refusal(load(path), "not_a_zip")

    def test_zip_magic_but_corrupt(self, tmp_path: Path) -> None:
        path = tmp_path / "corrupt.xlsx"
        path.write_bytes(b"PK\x03\x04" + b"\x00" * 32)
        expect_refusal(load(path), "not_a_zip")

    def test_zip64_entry_count_preflight(self, tmp_path: Path) -> None:
        # Patch the end-of-central-directory entry counts (offsets 8 and 10)
        # to the ZIP64 sentinel; the preflight must refuse before ZipFile
        # constructs anything.
        path = build_zip(tmp_path / "huge.xlsx", {"a.xml": _XML})
        raw = bytearray(path.read_bytes())
        eocd = raw.rfind(b"PK\x05\x06")
        assert eocd != -1
        raw[eocd + 8 : eocd + 12] = b"\xff\xff\xff\xff"
        path.write_bytes(bytes(raw))
        expect_refusal(load(path), "archive_too_many_parts")

    def test_understated_entry_count_is_rejected(self, tmp_path: Path) -> None:
        # Declare fewer entries than the central directory actually holds:
        # the stdlib walks the directory by size, so trusting the declared
        # count alone would let a hostile archive allocate unbounded ZipInfo
        # objects. The walk must catch the mismatch before ZipFile exists.
        path = build_zip(tmp_path / "lying.xlsx", {"a.xml": _XML, "b.xml": _XML, "c.xml": _XML})
        raw = bytearray(path.read_bytes())
        eocd = raw.rfind(b"PK\x05\x06")
        assert eocd != -1
        raw[eocd + 8 : eocd + 10] = (1).to_bytes(2, "little")
        raw[eocd + 10 : eocd + 12] = (1).to_bytes(2, "little")
        path.write_bytes(bytes(raw))
        refusal = expect_refusal(load(path), "malformed_archive")
        assert refusal.details["actual_records"] == 3
        assert refusal.details["declared_entries"] == 1

    def test_central_directory_out_of_bounds(self, tmp_path: Path) -> None:
        # Point the central directory past the end of the file.
        path = build_zip(tmp_path / "oob.xlsx", {"a.xml": _XML})
        raw = bytearray(path.read_bytes())
        eocd = raw.rfind(b"PK\x05\x06")
        assert eocd != -1
        raw[eocd + 16 : eocd + 20] = (len(raw)).to_bytes(4, "little")
        path.write_bytes(bytes(raw))
        expect_refusal(load(path), "malformed_archive")


class TestConcatenationRefusals:
    def test_trailing_data_after_eocd_rejected(self, tmp_path: Path) -> None:
        path = build_zip(tmp_path / "trail.xlsx", {"a.xml": _XML})
        path.write_bytes(path.read_bytes() + b"JUNK-APPENDED-AFTER-EOCD")
        refusal = expect_refusal(load(path), "malformed_archive")
        assert "trailing data" in str(refusal.details["reason"])

    def test_concatenated_zips_rejected(self, tmp_path: Path) -> None:
        first = build_zip(tmp_path / "first.xlsx", {"a.xml": _XML})
        second = build_zip(tmp_path / "second.xlsx", {"b.xml": _XML})
        combined = tmp_path / "combined.xlsx"
        combined.write_bytes(first.read_bytes() + second.read_bytes())
        expect_refusal(load(combined), "malformed_archive")

    def test_prefixed_zip_rejected(self, tmp_path: Path) -> None:
        clean = build_zip(tmp_path / "clean.xlsx", {"a.xml": _XML})
        prefixed = tmp_path / "prefixed.xlsx"
        prefixed.write_bytes(b"GARBAGE-PREFIX" + clean.read_bytes())
        expect_refusal(load(prefixed), "not_a_zip")

    def test_nonzero_first_local_header_offset_rejected(self, tmp_path: Path) -> None:
        # Patch the sole central-directory record's local-header offset field
        # (offset 42 within the record) so it no longer points at byte zero.
        path = build_zip(tmp_path / "shifted.xlsx", {"a.xml": _XML})
        raw = bytearray(path.read_bytes())
        record = raw.find(b"PK\x01\x02")
        assert record != -1
        raw[record + 42 : record + 46] = (1).to_bytes(4, "little")
        path.write_bytes(bytes(raw))
        refusal = expect_refusal(load(path), "malformed_archive")
        assert refusal.details["first_local_offset"] == 1


class TestArchiveLimitsValidation:
    def test_negative_or_zero_limits_raise(self) -> None:
        with pytest.raises(ValueError, match="max_file_bytes"):
            ArchiveLimits(max_file_bytes=0)
        with pytest.raises(ValueError, match="max_part_count"):
            ArchiveLimits(max_part_count=-1)
        with pytest.raises(ValueError, match="max_xml_depth"):
            ArchiveLimits(max_xml_depth=0)


class TestEntryRefusals:
    def test_too_many_parts(self, tmp_path: Path) -> None:
        entries = {f"part{i}.xml": _XML for i in range(5)}
        path = build_zip(tmp_path / "many.xlsx", entries)
        expect_refusal(load(path, ArchiveLimits(max_part_count=4)), "archive_too_many_parts")

    def test_part_too_large(self, tmp_path: Path) -> None:
        path = build_zip(tmp_path / "fat.xlsx", {"a.xml": b"x" * 100})
        expect_refusal(load(path, ArchiveLimits(max_part_bytes=99)), "archive_part_too_large")

    def test_total_expansion_too_large(self, tmp_path: Path) -> None:
        path = build_zip(tmp_path / "sum.xlsx", {"a.xml": b"x" * 60, "b.xml": b"y" * 60})
        expect_refusal(
            load(path, ArchiveLimits(max_part_bytes=80, max_total_bytes=100)),
            "archive_expansion_too_large",
        )

    def test_unsupported_compression(self, tmp_path: Path) -> None:
        path = tmp_path / "bz2.xlsx"
        with zipfile.ZipFile(path, "w") as zf:
            zf.writestr("a.xml", _XML, compress_type=zipfile.ZIP_BZIP2)
        refusal = expect_refusal(load(path), "unsupported_compression")
        assert refusal.details["compress_type"] == zipfile.ZIP_BZIP2

    def test_crc_failure_returns_refusal(self, tmp_path: Path) -> None:
        # Corrupt the stored payload so the declared CRC no longer matches;
        # the failure must surface as a refusal, not an exception.
        payload = b"<root>payload-for-crc-corruption</root>"
        path = build_zip(tmp_path / "crc.xlsx", {"a.xml": payload})
        raw = bytearray(path.read_bytes())
        index = raw.find(payload)
        assert index != -1
        raw[index + 10] ^= 0xFF
        path.write_bytes(bytes(raw))
        archive = expect_archive(load(path))
        result = archive.read_part("a.xml")
        assert isinstance(result, Refusal)
        assert result.code == "malformed_archive"
        archive.close()

    def test_traversal_entry_name(self, tmp_path: Path) -> None:
        path = build_zip(tmp_path / "trav.xlsx", {"../evil.xml": _XML})
        refusal = expect_refusal(load(path), "unsafe_archive_path")
        assert refusal.details["entry"] == "../evil.xml"

    def test_absolute_entry_name(self, tmp_path: Path) -> None:
        path = build_zip(tmp_path / "abs.xlsx", {"/abs.xml": _XML})
        expect_refusal(load(path), "unsafe_archive_path")

    def test_backslash_entry_name(self, tmp_path: Path) -> None:
        path = build_zip(tmp_path / "back.xlsx", {"bad\\name.xml": _XML})
        expect_refusal(load(path), "unsafe_archive_path")

    def test_drive_prefixed_entry_name(self, tmp_path: Path) -> None:
        path = build_zip(tmp_path / "drive.xlsx", {"C:evil.xml": _XML})
        expect_refusal(load(path), "unsafe_archive_path")

    def test_encrypted_entry(self, tmp_path: Path) -> None:
        # zipfile sanitizes flag bits when writing, so set the encryption bit
        # by patching the central-directory record (offset 8 after PK\x01\x02).
        path = build_zip(tmp_path / "enc.xlsx", {"xl/secret.xml": _XML})
        raw = bytearray(path.read_bytes())
        central_dir = raw.find(b"PK\x01\x02")
        assert central_dir != -1
        raw[central_dir + 8] |= 0x1
        path.write_bytes(bytes(raw))
        refusal = expect_refusal(load(path), "encrypted_archive_entry")
        assert len(refusal.recovery_options) > 0

    def test_duplicate_entry(self, tmp_path: Path) -> None:
        path = tmp_path / "dup.xlsx"
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")  # zipfile warns about duplicate names
            with zipfile.ZipFile(path, "w") as zf:
                zf.writestr("a.xml", _XML)
                zf.writestr("a.xml", _XML)
        expect_refusal(load(path), "duplicate_archive_entry")


class TestXmlParsing:
    def _parse(self, tmp_path: Path, data: bytes, limits: ArchiveLimits | None = None) -> object:
        archive = expect_archive(load(build_zip(tmp_path / "x.xlsx", {"a.xml": data}), limits))
        result = archive.parse_xml_part("a.xml")
        archive.close()
        return result

    def test_parses_wellformed_xml(self, tmp_path: Path) -> None:
        root = self._parse(tmp_path, _XML)
        assert not isinstance(root, Refusal)

    def test_doctype_blocked_utf8(self, tmp_path: Path) -> None:
        evil = b'<?xml version="1.0"?><!DOCTYPE r [<!ENTITY a "b">]><r>&a;</r>'
        result = self._parse(tmp_path, evil)
        assert isinstance(result, Refusal)
        assert result.code == "dtd_blocked"

    def test_doctype_blocked_utf16_le(self, tmp_path: Path) -> None:
        evil = '<?xml version="1.0" encoding="UTF-16"?><!DOCTYPE r><r/>'.encode("utf-16")
        result = self._parse(tmp_path, evil)
        assert isinstance(result, Refusal)
        assert result.code == "dtd_blocked"

    def test_doctype_blocked_utf16_be(self, tmp_path: Path) -> None:
        evil = b"\xfe\xff" + '<?xml version="1.0"?><!DOCTYPE r><r/>'.encode("utf-16-be")
        result = self._parse(tmp_path, evil)
        assert isinstance(result, Refusal)
        assert result.code == "dtd_blocked"

    def test_valid_utf16_parses(self, tmp_path: Path) -> None:
        data = '<?xml version="1.0" encoding="UTF-16"?><root/>'.encode("utf-16")
        result = self._parse(tmp_path, data)
        assert not isinstance(result, Refusal)

    def test_lowercase_doctype_is_still_refused(self, tmp_path: Path) -> None:
        # A lowercase doctype keyword is not legal XML; it must refuse either
        # way, and expat classifies it as malformed.
        evil = b'<?xml version="1.0"?><!doctype r><r/>'
        result = self._parse(tmp_path, evil)
        assert isinstance(result, Refusal)
        assert result.code == "malformed_xml"

    def test_malformed_xml(self, tmp_path: Path) -> None:
        result = self._parse(tmp_path, b"<root><unclosed>")
        assert isinstance(result, Refusal)
        assert result.code == "malformed_xml"
        assert result.details["part"] == "a.xml"

    def test_depth_limit(self, tmp_path: Path) -> None:
        deep = b"<a>" * 10 + b"</a>" * 10
        result = self._parse(tmp_path, deep, ArchiveLimits(max_xml_depth=5))
        assert isinstance(result, Refusal)
        assert result.code == "xml_limits_exceeded"
        assert result.details["kind"] == "depth"

    def test_element_count_limit(self, tmp_path: Path) -> None:
        wide = b"<r>" + b"<c/>" * 10 + b"</r>"
        result = self._parse(tmp_path, wide, ArchiveLimits(max_xml_elements=5))
        assert isinstance(result, Refusal)
        assert result.code == "xml_limits_exceeded"
        assert result.details["kind"] == "elements"


class TestPathNormalization:
    def test_normalizes_dot_segments(self) -> None:
        assert normalize_package_path("xl/./worksheets/sheet1.xml") == "xl/worksheets/sheet1.xml"

    def test_strips_leading_slash(self) -> None:
        assert normalize_package_path("/xl/workbook.xml") == "xl/workbook.xml"

    def test_refuses_escape(self) -> None:
        assert normalize_package_path("../outside.xml") is None
        assert normalize_package_path("xl/../../outside.xml") is None

    def test_refuses_empty(self) -> None:
        assert normalize_package_path("") is None
        assert normalize_package_path("/") is None

    def test_refuses_uri_schemes(self) -> None:
        assert normalize_package_path("http://evil.example/x.xml") is None
        assert normalize_package_path("file:///etc/passwd") is None
        assert normalize_package_path("mailto:someone@example.com") is None
        assert normalize_package_path("C:evil.xml") is None

    def test_refuses_authority(self) -> None:
        assert normalize_package_path("//host/share/x.xml") is None
        assert normalize_package_path("///triple.xml") is None

    def test_refuses_query_and_fragment(self) -> None:
        assert normalize_package_path("xl/workbook.xml?x=1") is None
        assert normalize_package_path("xl/workbook.xml#frag") is None

    def test_refuses_backslash(self) -> None:
        assert normalize_package_path("xl\\workbook.xml") is None


class TestTargetValidation:
    """Two-phase OPC target validation: raw checks, decode, re-check."""

    def test_plain_target_passes_unchanged(self) -> None:
        assert validate_opc_target("worksheets/sheet1.xml") == "worksheets/sheet1.xml"

    def test_decodes_safe_percent_encoding(self) -> None:
        assert validate_opc_target("media/image%201.png") == "media/image 1.png"

    def test_rejects_unsafe_percent_encodings(self) -> None:
        assert validate_opc_target("%2e%2e/outside.xml") is None
        assert validate_opc_target("xl%2fworkbook.xml") is None
        assert validate_opc_target("xl%5Cworkbook.xml") is None
        assert validate_opc_target("xl/%252e%252e/x.xml") is None  # double-encoded

    def test_post_decode_query_rejected(self) -> None:
        # "?" is checked on the raw string; %3F must not reintroduce it.
        assert validate_opc_target("styles.xml%3Fx=1") is None

    def test_post_decode_fragment_rejected(self) -> None:
        assert validate_opc_target("styles.xml%23frag") is None

    def test_post_decode_scheme_rejected(self) -> None:
        assert validate_opc_target("http%3A//evil.example/x.xml") is None

    def test_post_decode_control_character_rejected(self) -> None:
        assert validate_opc_target("styles%01.xml") is None

    def test_raw_scheme_and_authority_rejected(self) -> None:
        assert validate_opc_target("http://evil.example/x.xml") is None
        assert validate_opc_target("//host/share/x.xml") is None
        assert validate_opc_target("C:evil.xml") is None
