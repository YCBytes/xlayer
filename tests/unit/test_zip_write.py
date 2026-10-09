"""ZIP32 preservation, using independently constructed headers and descriptors."""

from __future__ import annotations

import struct
import zipfile
import zlib
from pathlib import Path
from typing import BinaryIO, cast

import pytest

from tests.unit.test_transaction import host, prepared, workbook
from xlayer._edits import SetValue
from xlayer._errors import Refusal
from xlayer._ooxml.archive import WorkbookArchive
from xlayer._receipt import Receipt
from xlayer._workbook import Workbook


def variant(
    path: Path,
    *,
    flags: int = 0,
    compression: int = 8,
    signed: bool = True,
    local_extra: bytes = b"",
    central_extra: bytes = b"",
    reverse_directory: bool = False,
    name_encoding: str = "utf-8",
) -> None:
    """Small independent ZIP encoder; inputs are synthetic public unit fixtures."""
    with zipfile.ZipFile(path) as source:
        parts = [(info.filename, source.read(info)) for info in source.infolist()]
    parts += [("opaque/café.xml", b"opaque payload"), ("empty/", b"")]
    local, records = bytearray(), []
    for name, data in parts:
        raw_name = name.encode(name_encoding)
        crc = zlib.crc32(data)
        compressor = zlib.compressobj(6, zlib.DEFLATED, -15)
        content = compressor.compress(data) + compressor.flush() if compression == 8 else data
        offset = len(local)
        values = (0, 0, 0) if flags & 8 else (crc, len(content), len(data))
        local += (
            struct.pack(
                "<4s5H3I2H",
                b"PK\x03\x04",
                20,
                flags,
                compression,
                123,
                456,
                *values,
                len(raw_name),
                len(local_extra),
            )
            + raw_name
            + local_extra
            + content
        )
        if flags & 8:
            local += (b"PK\x07\x08" if signed else b"") + struct.pack(
                "<III", crc, len(content), len(data)
            )
        comment = b"entry comment"
        records.append(
            struct.pack(
                "<4s6H3I5H2I",
                b"PK\x01\x02",
                45,
                20,
                flags,
                compression,
                123,
                456,
                crc,
                len(content),
                len(data),
                len(raw_name),
                len(central_extra),
                len(comment),
                0,
                0,
                0,
                offset,
            )
            + raw_name
            + central_extra
            + comment
        )
    directory = b"".join(reversed(records) if reverse_directory else records)
    comment = b"archive comment"
    path.write_bytes(
        local
        + directory
        + struct.pack(
            "<4s4H2IH",
            b"PK\x05\x06",
            0,
            0,
            len(records),
            len(records),
            len(directory),
            len(local),
            len(comment),
        )
        + comment
    )


def records(path: Path) -> dict[str, tuple[bytes, bytes, bytes]]:
    """Independent framing oracle; no production admission/writer helpers."""
    raw = path.read_bytes()
    with zipfile.ZipFile(path) as source:
        infos = sorted(source.infolist(), key=lambda i: i.header_offset)
        result = {}
        for index, info in enumerate(infos):
            start = info.header_offset
            n, e = struct.unpack_from("<HH", raw, start + 26)
            content_start = start + 30 + n + e
            content_end = content_start + info.compress_size
            end = infos[index + 1].header_offset if index + 1 < len(infos) else source.start_dir
            result[info.filename] = (
                raw[start:content_start],
                raw[content_start:content_end],
                raw[content_end:end],
            )
        return result


@pytest.mark.parametrize("mode", [0, 2, 4, 6])
@pytest.mark.parametrize("descriptor", [None, False, True])
def test_deflate_options_and_descriptor_records_preserved(
    tmp_path: Path, mode: int, descriptor: bool | None
) -> None:
    source = workbook(tmp_path)
    flags = 0x800 | mode | (8 if descriptor is not None else 0)
    variant(source, flags=flags, signed=descriptor is not False)
    before = source.read_bytes()
    output = tmp_path / "out.xlsx"
    with Workbook.open(source) as book:
        proposal, preview, approval = prepared(book, output)
        assert not preview.evidence["blocked"]
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Receipt), result
        assert result.evidence["recalculation_status"] == "required_not_run"
    assert source.read_bytes() == before
    old, new = records(source), records(output)
    assert list(old) == list(new)
    for name in old:
        if name != "xl/worksheets/1.xml":
            assert new[name] == old[name]  # includes original compressed stream
        else:
            assert new[name][0][:14] == old[name][0][:14]
            assert new[name][0][26:] == old[name][0][26:]
            assert len(new[name][2]) == len(old[name][2])
            with zipfile.ZipFile(output) as archive:
                info = archive.getinfo(name)
                assert info.flag_bits == flags
                if descriptor is not None:
                    assert struct.unpack("<III", new[name][2][-12:]) == (
                        info.CRC,
                        info.compress_size,
                        info.file_size,
                    )
                assert b"<v>120</v>" in archive.read(name)


@pytest.mark.parametrize("signed", [False, True])
def test_stored_descriptors_and_cp437_names(tmp_path: Path, signed: bool) -> None:
    source = workbook(tmp_path)
    variant(source, flags=8, compression=0, signed=signed, name_encoding="cp437")
    output = tmp_path / "out.xlsx"
    with Workbook.open(source) as book:
        proposal, _, approval = prepared(book, output)
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Receipt), result
    assert records(output)["opaque/café.xml"] == records(source)["opaque/café.xml"]


@pytest.mark.parametrize("central_padding", [False, True])
def test_growth_hint_and_physical_order_are_preserved(
    tmp_path: Path, central_padding: bool
) -> None:
    source = workbook(tmp_path)
    padding = struct.pack("<HHHH", 0xA220, 20, 0xA028, 512) + bytes(16)
    variant(
        source,
        flags=0x806,
        local_extra=padding,
        central_extra=padding if central_padding else b"",
        reverse_directory=True,
    )
    output = tmp_path / "out.xlsx"
    with Workbook.open(source) as book:
        proposal, _, approval = prepared(book, output)
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Receipt), result
    old, new = records(source), records(output)
    assert list(old) == list(new)
    for name in old:
        assert new[name][0][26:] == old[name][0][26:]
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(output) as after:
        assert before.namelist() == after.namelist()
        assert before.comment == after.comment
        for prior, actual in zip(before.infolist(), after.infolist(), strict=True):
            assert actual.flag_bits == prior.flag_bits
            assert actual.extra == prior.extra
            assert actual.external_attr == prior.external_attr == 0


@pytest.mark.parametrize(
    "issue", ["flag", "extra", "padding_signature", "padding_data", "stored_option", "local_extra"]
)
def test_unsupported_metadata_blocks_preview_and_atomic_apply(tmp_path: Path, issue: str) -> None:
    source = workbook(tmp_path, '<c r="A1"><v>100</v></c><c r="B1"><v>200</v></c>')
    padding = struct.pack("<HHHH", 0xA220, 5, 0xA028, 1) + b"x"
    variant(
        source,
        flags=0x810 if issue == "flag" else 2 if issue == "stored_option" else 0x800,
        compression=0 if issue == "stored_option" else 8,
        central_extra=b"\x0d\xf0\x00\x00" if issue == "extra" else b"",
        local_extra=b"\x0d\xf0\x00\x00"
        if issue == "local_extra"
        else padding
        if issue == "padding_data"
        else struct.pack("<HHHH", 0xA220, 4, 1, 0)
        if issue == "padding_signature"
        else b"",
    )
    before = source.read_bytes()
    output = tmp_path / "out.xlsx"
    with Workbook.open(source) as book:
        proposal, preview, approval = prepared(
            book, output, edits=[SetValue("Inputs", "A1", 120), SetValue("Inputs", "B1", 220)]
        )
        assert preview.evidence["blocked"] is True
        assert any(
            f["refusal"]["code"] == "unsupported_zip_metadata"
            for f in cast("list[dict[str, dict[str, object]]]", preview.to_dict()["findings"])
        )
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal)
    assert not output.exists() and source.read_bytes() == before
    assert not list(tmp_path.glob(".xlayer-*"))


@pytest.mark.parametrize("issue", ["descriptor_crc", "local_flags", "local_size"])
def test_conflicting_local_framing_blocks_preview(tmp_path: Path, issue: str) -> None:
    source = workbook(tmp_path)
    variant(source, flags=0x808 if issue == "descriptor_crc" else 0x800)
    with zipfile.ZipFile(source) as archive:
        info = archive.getinfo("opaque/café.xml")
    raw = bytearray(source.read_bytes())
    n, e = struct.unpack_from("<HH", raw, info.header_offset + 26)
    if issue == "descriptor_crc":
        raw[info.header_offset + 30 + n + e + info.compress_size + 4] ^= 1
    elif issue == "local_flags":
        raw[info.header_offset + 6] ^= 2
    else:
        raw[info.header_offset + 18] ^= 1
    source.write_bytes(raw)
    with Workbook.open(source) as book:
        _, preview, _ = prepared(book, tmp_path / "out.xlsx")
        assert preview.evidence["blocked"] is True
        assert any(
            f["refusal"]["code"] == "malformed_archive"
            for f in cast("list[dict[str, dict[str, object]]]", preview.to_dict()["findings"])
        )


@pytest.mark.parametrize("fault", ["flags", "local_padding", "descriptor"])
def test_verifier_rejects_writer_metadata_or_framing_drift(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, fault: str
) -> None:
    import xlayer._transaction as transaction

    source = workbook(tmp_path)
    padding = struct.pack("<HHHH", 0xA220, 4, 0xA028, 0)
    variant(source, flags=0x808, local_extra=padding)
    real = transaction.write_archive

    def corrupt(
        archive: WorkbookArchive, handle: BinaryIO, part: str, patched: bytes
    ) -> Refusal | None:
        result = real(archive, handle, part, patched)
        assert result is None
        handle.seek(0)
        raw = bytearray(handle.read())
        with zipfile.ZipFile(handle) as written:
            info = written.getinfo("xl/worksheets/1.xml")
        n, e = struct.unpack_from("<HH", raw, info.header_offset + 26)
        if fault == "local_padding":
            # Change the historical initial-padding value; still a valid hint.
            raw[info.header_offset + 30 + n + 6] ^= 1
        elif fault == "descriptor":
            raw[info.header_offset + 30 + n + e + info.compress_size + 4] ^= 1
        else:
            # ASCII target: removing the UTF-8 bit does not change decoded XML.
            raw[info.header_offset + 7] ^= 8
            for pos in range(written.start_dir, len(raw) - 46):
                if (
                    raw[pos : pos + 4] == b"PK\x01\x02"
                    and raw[pos + 46 : pos + 46 + n] == part.encode()
                ):
                    raw[pos + 9] ^= 8
                    break
        handle.seek(0)
        handle.write(raw)
        return None

    monkeypatch.setattr(transaction, "write_archive", corrupt)
    output = tmp_path / "out.xlsx"
    with Workbook.open(source) as book:
        proposal, _, approval = prepared(book, output)
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "verification_failed"
    assert not output.exists()
    assert not list(tmp_path.glob(".xlayer-*"))
