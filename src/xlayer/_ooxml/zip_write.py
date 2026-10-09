"""Bounded ZIP32 admission and one-part preservation; never extracts to disk.

Untouched local records are copied verbatim. Only the target's stream, CRC,
sizes and descriptor, central offsets, and EOCD derived fields may change.
Unknown flags/extras refuse even if a raw copy would be mechanically possible.
"""

from __future__ import annotations

import struct
import zipfile
import zlib
from dataclasses import dataclass
from typing import IO, BinaryIO, NoReturn

from xlayer._approval import refusal
from xlayer._errors import Refusal
from xlayer._ooxml.archive import WorkbookArchive

_LOCAL = struct.Struct("<4s5H3I2H")
_CENTRAL = struct.Struct("<4s6H3I5H2I")
_EOCD = struct.Struct("<4s4H2IH")
_VALUES = struct.Struct("<III")
_U32_MAX = 0xFFFFFFFF


class _ZipAdmissionError(Exception):
    def __init__(self, code: str, reason: str, entry: str | None = None) -> None:
        self.refusal = refusal(code, reason, **({"entry": entry} if entry else {}))


def _malformed(reason: str, entry: str | None = None) -> NoReturn:
    raise _ZipAdmissionError("malformed_archive", reason, entry)


def _unsupported(reason: str, entry: str | None = None) -> NoReturn:
    raise _ZipAdmissionError("unsupported_zip_metadata", reason, entry)


def _read_at(stream: IO[bytes], offset: int, length: int) -> bytes:
    if offset < 0 or length < 0:
        _malformed("negative ZIP record bounds")
    stream.seek(offset)
    data = stream.read(length)
    if len(data) != length:
        _malformed("truncated ZIP record")
    return data


def _extras(raw: bytes, name: str) -> None:
    position, seen = 0, False
    while position < len(raw):
        if len(raw) - position < 4:
            _malformed("truncated ZIP extra field", name)
        key, size = struct.unpack_from("<HH", raw, position)
        end = position + 4 + size
        if end > len(raw):
            _malformed("truncated ZIP extra payload", name)
        value = raw[position + 4 : end]
        if key != 0xA220 or seen or size < 4 or value[:2] != b"\x28\xa0" or any(value[4:]):
            _unsupported("unreviewed ZIP extra field or growth hint", name)
        seen, position = True, end


@dataclass(frozen=True)
class ZipEntry:
    info: zipfile.ZipInfo
    central: bytes
    header: bytes
    start: int
    content_start: int
    end: int
    descriptor: bytes


@dataclass(frozen=True)
class ZipLayout:
    # Central order need not equal physical order.
    entries: tuple[ZipEntry, ...]
    eocd: bytes


def _layout(archive: WorkbookArchive, stream: IO[bytes]) -> ZipLayout:
    infos = archive._zip.infolist()
    physical = sorted(infos, key=lambda info: info.header_offset)
    bounds = {
        info.filename: physical[index + 1].header_offset
        if index + 1 < len(physical)
        else archive._zip.start_dir
        for index, info in enumerate(physical)
    }
    if not physical or physical[0].header_offset != 0:
        _unsupported("ZIP local records must start at offset zero")
    position = archive._zip.start_dir
    entries = []
    for info in infos:
        name = info.filename
        if (
            info.compress_type not in {zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED}
            or info.flag_bits & ~0x80E
            or (info.compress_type == zipfile.ZIP_STORED and info.flag_bits & 6)
            or info.volume
            or info.extract_version > 20
            or info.reserved
            or (info.is_dir() and info.file_size)
        ):
            _unsupported("ZIP metadata cannot be safely preserved", name)
        fixed = _read_at(stream, position, _CENTRAL.size)
        fields = _CENTRAL.unpack(fixed)
        if fields[0] != b"PK\x01\x02":
            _malformed("invalid central record signature", name)
        n, e, c = fields[10:13]
        variable = _read_at(stream, position + _CENTRAL.size, n + e + c)
        central = fixed + variable
        _extras(variable[n : n + e], name)
        expected = (
            info.create_version | info.create_system << 8,
            info.extract_version,
            info.flag_bits,
            info.compress_type,
            info.CRC,
            info.compress_size,
            info.file_size,
            info.volume,
            info.internal_attr,
            info.external_attr,
            info.header_offset,
        )
        actual = (
            fields[1],
            fields[2],
            fields[3],
            fields[4],
            *fields[7:10],
            fields[13],
            fields[14],
            fields[15],
            fields[16],
        )
        if (
            actual != expected
            or variable[n : n + e] != info.extra
            or variable[n + e :] != info.comment
            or variable[:n].decode("utf-8" if info.flag_bits & 0x800 else "cp437") != name
        ):
            _malformed("central record differs from loaded entry metadata", name)
        position += len(central)
        start, end = info.header_offset, bounds[name]
        if start + _LOCAL.size > end:
            _malformed("overlapping or truncated local header", name)
        local_fixed = _read_at(stream, start, _LOCAL.size)
        local = _LOCAL.unpack(local_fixed)
        ln, le = local[9:11]
        content_start = start + _LOCAL.size + ln + le
        content_end = content_start + info.compress_size
        if content_end > end:
            _malformed("overlapping or truncated local payload", name)
        local_variable = _read_at(stream, start + _LOCAL.size, ln + le)
        if (
            local[0] != b"PK\x03\x04"
            or local[1:6] != fields[2:7]
            or local_variable[:ln] != variable[:n]
        ):
            _malformed("local and central headers disagree", name)
        _extras(local_variable[ln:], name)
        values = (info.CRC, info.compress_size, info.file_size)
        descriptor_length = end - content_end
        if info.flag_bits & 8:
            if local[6:9] not in {(0, 0, 0), values}:
                _malformed("local descriptor fields conflict with central metadata", name)
            if descriptor_length not in {12, 16}:
                _unsupported("unreviewed ZIP descriptor layout", name)
            descriptor = _read_at(stream, content_end, descriptor_length)
            if (descriptor_length == 16 and descriptor[:4] != b"PK\x07\x08") or _VALUES.unpack(
                descriptor[-12:]
            ) != values:
                _malformed("ZIP descriptor disagrees with central metadata", name)
        else:
            if local[6:9] != values:
                _malformed("local sizes or CRC disagree with central metadata", name)
            if descriptor_length:
                _unsupported("unreviewed bytes between ZIP local records", name)
            descriptor = b""
        entries.append(
            ZipEntry(
                info, central, local_fixed + local_variable, start, content_start, end, descriptor
            )
        )
    fixed_end = _read_at(stream, position, _EOCD.size)
    end_fields = _EOCD.unpack(fixed_end)
    if (
        end_fields[0] != b"PK\x05\x06"
        or end_fields[1:5] != (0, 0, len(infos), len(infos))
        or end_fields[5:7] != (position - archive._zip.start_dir, archive._zip.start_dir)
    ):
        _malformed("central directory and end record disagree")
    comment = _read_at(stream, position + _EOCD.size, end_fields[7])
    stream.seek(0, 2)
    if stream.tell() != position + _EOCD.size + len(comment) or comment != archive._zip.comment:
        _malformed("end record or archive comment changed")
    return ZipLayout(tuple(entries), fixed_end + comment)


def inspect_write_layout(archive: WorkbookArchive) -> ZipLayout | Refusal:
    stream = archive._zip.fp
    if stream is None:
        raise RuntimeError("write admission requires an open archive")
    position = stream.tell()
    try:
        return _layout(archive, stream)
    except _ZipAdmissionError as exc:
        return exc.refusal
    finally:
        stream.seek(position)


def check_write_support(archive: WorkbookArchive) -> Refusal | None:
    result = inspect_write_layout(archive)
    return result if isinstance(result, Refusal) else None


def _copy_range(source: IO[bytes], output: BinaryIO, start: int, end: int) -> None:
    source.seek(start)
    remaining = end - start
    while remaining:
        data = source.read(min(remaining, 1024 * 1024))
        if not data:
            _malformed("source local record truncated during copy")
        output.write(data)
        remaining -= len(data)


def write_preserved_archive(
    archive: WorkbookArchive, handle: BinaryIO, part: str, patched: bytes
) -> Refusal | None:
    layout = inspect_write_layout(archive)
    if isinstance(layout, Refusal):
        return layout
    # Keep all original CRC checks, including opaque parts and empty directories.
    for entry in layout.entries:
        if entry.info.is_dir():
            try:
                content = archive._zip.read(entry.info)
            except (zipfile.BadZipFile, zlib.error) as exc:
                return refusal(
                    "malformed_archive", "directory integrity check failed", error=str(exc)
                )
            if content:
                return refusal("unsupported_zip_metadata", "directory entry carries opaque payload")
        else:
            checked = archive.read_part(entry.info.filename)
            if isinstance(checked, Refusal):
                return checked
    target = next((entry for entry in layout.entries if entry.info.filename == part), None)
    if target is None or target.info.is_dir():
        raise RuntimeError("declared target part is absent from the write layout")
    compressed = patched
    if target.info.compress_type == zipfile.ZIP_DEFLATED:
        level = {0: 6, 2: 9, 4: 3, 6: 1}[target.info.flag_bits & 6]
        encoder = zlib.compressobj(level, zlib.DEFLATED, -15)
        compressed = encoder.compress(patched) + encoder.flush()
    values = (zlib.crc32(patched), len(compressed), len(patched))
    if max(values[1:]) >= _U32_MAX or len(layout.entries) >= 0xFFFF:
        return refusal("unsupported_zip_metadata", "write would require ZIP64")
    stream = archive._zip.fp
    if stream is None:
        raise RuntimeError("write lost the open source archive")
    offsets = {}
    for entry in sorted(layout.entries, key=lambda entry: entry.start):
        offsets[entry.info.filename] = handle.tell()
        if entry is not target:
            _copy_range(stream, handle, entry.start, entry.end)
            continue
        header = bytearray(entry.header)
        if not entry.descriptor or header[14:26] != bytes(12):
            _VALUES.pack_into(header, 14, *values)
        handle.write(header)
        handle.write(compressed)
        if entry.descriptor:
            handle.write(entry.descriptor[:-12] + _VALUES.pack(*values))
    start_dir = handle.tell()
    if start_dir >= _U32_MAX:
        return refusal("unsupported_zip_metadata", "write offsets would require ZIP64")
    for entry in layout.entries:
        central = bytearray(entry.central)
        struct.pack_into("<I", central, 42, offsets[entry.info.filename])
        if entry is target:
            _VALUES.pack_into(central, 16, *values)
        handle.write(central)
    end = bytearray(layout.eocd)
    struct.pack_into("<II", end, 12, handle.tell() - start_dir, start_dir)
    handle.write(end)
    return None
