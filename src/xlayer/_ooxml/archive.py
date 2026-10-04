"""Safe opening of untrusted ``.xlsx`` ZIP containers.

This layer is public-only hardening with no counterpart in ``xlayer-core``,
whose parsers read an already-extracted directory and assume trusted input.
Here, resource limits are validated before decompression, the central
directory is walked and reconciled against the end-of-central-directory
record before ``ZipFile`` constructs anything, all parts stay in memory,
nothing is ever extracted to disk, and no I/O happens beyond one bounded read
of the source file's bytes.

XML safety is enforced by a streaming expat preflight of every part before a
tree is built: DOCTYPE declarations are rejected through expat's own handler
(and therefore independent of text encoding), and element-count and nesting
limits bound memory without allocating a tree. Modern expat additionally
ships billion-laughs amplification limits, and ElementTree performs no
external entity or network I/O.

Every failure is returned as a structured :class:`~xlayer._errors.Refusal`
carrying at least one structured recovery option; the public
``Workbook.open`` boundary decides whether to raise.
"""

from __future__ import annotations

import hashlib
import io
import posixpath
import re
import struct
import urllib.parse
import xml.parsers.expat
import zipfile
import zlib
from dataclasses import dataclass, fields
from pathlib import Path
from xml.etree import ElementTree as ET

from xlayer._errors import Refusal

CONTENT_TYPES_PART = "[Content_Types].xml"

# Compound File Binary magic: both legacy .xls and encrypted .xlsx use this
# container, and neither is supported.
_CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_ZIP_MAGIC_PREFIXES = (b"PK\x03\x04", b"PK\x05\x06")

_ENCRYPTED_FLAG_BIT = 0x1
# General-purpose flag bit 11: entry name and comment are UTF-8 (else cp437).
_UTF8_NAME_FLAG_BIT = 0x800
_SUPPORTED_COMPRESSION = (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED)

# End-of-central-directory record: fixed 22 bytes plus up to 65535 comment
# bytes. Fields at offset 4: this disk (H), disk of the central directory (H),
# entries on this disk (H), total entries (H), central directory size (I),
# central directory offset (I).
_EOCD_SIGNATURE = b"PK\x05\x06"
_EOCD_FIXED_SIZE = 22
_EOCD_MAX_SCAN = _EOCD_FIXED_SIZE + 65535
_ZIP64_U16 = 0xFFFF
_ZIP64_U32 = 0xFFFFFFFF

_CENTRAL_HEADER_SIGNATURE = b"PK\x01\x02"
_CENTRAL_HEADER_FIXED_SIZE = 46

# RFC 3986 scheme (also catches Windows drive prefixes such as "C:").
_URI_SCHEME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9+.\-]*:")

# Percent-encodings that could smuggle traversal or separators past the
# checks below ("."/"/"/"\\"), plus "%25" so double-encoding cannot hide them.
_UNSAFE_PERCENT_ENCODING_RE = re.compile(r"%2e|%2f|%5c|%25", re.IGNORECASE)

_RECOVER_TRUSTED_SOURCE = ({"action": "regenerate_workbook_from_trusted_source"},)
_RECOVER_VALID_XLSX = ({"action": "provide_valid_xlsx_file"},)
_RECOVER_SMALLER = ({"action": "provide_smaller_workbook"},)
_RECOVER_UNENCRYPTED = ({"action": "resave_as_unencrypted_xlsx"},)


@dataclass(frozen=True)
class ArchiveLimits:
    """Resource ceilings enforced before decompression and tree building.

    Module-level defaults are the public contract; construction with lowered
    values is intended only for security tests, which use small crafted
    archives instead of genuinely large ones.
    """

    max_file_bytes: int = 100 * 1024 * 1024
    max_part_count: int = 2_000
    max_part_bytes: int = 200 * 1024 * 1024
    max_total_bytes: int = 1024 * 1024 * 1024
    max_xml_elements: int = 1_000_000
    max_xml_depth: int = 100

    def __post_init__(self) -> None:
        for spec in fields(self):
            value = getattr(self, spec.name)
            if not isinstance(value, int) or value <= 0:
                raise ValueError(f"{spec.name} must be a positive integer, got {value!r}")


DEFAULT_LIMITS = ArchiveLimits()


class _DoctypeFoundError(Exception):
    pass


class _XmlLimitError(Exception):
    def __init__(self, kind: str, value: int, limit: int) -> None:
        super().__init__(kind)
        self.kind = kind
        self.value = value
        self.limit = limit


def _preflight_xml(data: bytes, part: str, limits: ArchiveLimits) -> Refusal | None:
    """Stream-parse a part with expat before any tree is built.

    Rejects DOCTYPE via expat's own handler (encoding-independent), enforces
    element-count and depth limits with no allocation proportional to the
    document, and surfaces malformed XML.
    """
    parser = xml.parsers.expat.ParserCreate()
    depth = 0
    count = 0

    def start_doctype(
        _doctype_name: str,
        _system_id: str | None,
        _public_id: str | None,
        _has_internal_subset: bool,
    ) -> None:
        raise _DoctypeFoundError

    def start_element(_name: str, _attrs: dict[str, str]) -> None:
        nonlocal depth, count
        depth += 1
        count += 1
        if depth > limits.max_xml_depth:
            raise _XmlLimitError("depth", depth, limits.max_xml_depth)
        if count > limits.max_xml_elements:
            raise _XmlLimitError("elements", count, limits.max_xml_elements)

    def end_element(_name: str) -> None:
        nonlocal depth
        depth -= 1

    parser.StartDoctypeDeclHandler = start_doctype
    parser.StartElementHandler = start_element
    parser.EndElementHandler = end_element

    try:
        parser.Parse(data, True)
    except _DoctypeFoundError:
        return Refusal(
            code="dtd_blocked",
            message=f"Part {part!r} contains a DOCTYPE declaration, which no legitimate "
            "OOXML part uses.",
            details={"part": part},
            recovery_options=_RECOVER_TRUSTED_SOURCE,
        )
    except _XmlLimitError as exc:
        return Refusal(
            code="xml_limits_exceeded",
            message=f"Part {part!r} exceeds the XML {exc.kind} limit of {exc.limit}.",
            details={"part": part, "kind": exc.kind, "value": exc.value, "limit": exc.limit},
            recovery_options=_RECOVER_TRUSTED_SOURCE,
        )
    except xml.parsers.expat.ExpatError as exc:
        return Refusal(
            code="malformed_xml",
            message=f"Part {part!r} is not well-formed XML: {exc}.",
            details={"part": part, "error": str(exc)},
            recovery_options=_RECOVER_TRUSTED_SOURCE,
        )
    return None


def _validate_entry_name(name: str) -> Refusal | None:
    unsafe_reason: str | None = None
    if not name:
        unsafe_reason = "empty entry name"
    elif any(ord(ch) < 0x20 or ch == "\x7f" for ch in name):
        # ZipInfo truncates names at the first NUL, so "a.xml\x00../x" would
        # otherwise be seen by the stdlib as "a.xml" while carrying hidden
        # bytes in the archive.
        unsafe_reason = "control character in path"
    elif name.startswith("/"):
        unsafe_reason = "absolute path"
    elif "\\" in name:
        unsafe_reason = "backslash in path"
    elif len(name) >= 2 and name[1] == ":":
        unsafe_reason = "drive-prefixed path"
    elif any(segment == ".." for segment in name.split("/")):
        unsafe_reason = "parent-directory traversal"

    if unsafe_reason is None:
        return None
    return Refusal(
        code="unsafe_archive_path",
        message=f"Archive entry {name!r} is unsafe: {unsafe_reason}.",
        details={"entry": name, "reason": unsafe_reason},
        recovery_options=_RECOVER_TRUSTED_SOURCE,
    )


def _malformed_archive(reason: str, **details: object) -> Refusal:
    return Refusal(
        code="malformed_archive",
        message=f"The archive's central directory is inconsistent: {reason}.",
        details={"reason": reason, **details},
        recovery_options=_RECOVER_TRUSTED_SOURCE,
    )


def _close_on_refusal(zip_file: zipfile.ZipFile, refusal: Refusal) -> Refusal:
    """Close a ZipFile that was constructed but will not be owned by an archive."""
    zip_file.close()
    return refusal


def _decode_entry_name(raw_name: bytes, flags: int) -> str | Refusal:
    """Decode a central-directory name exactly as the stdlib would.

    Done here, on the raw bytes, for two reasons. First, ``ZipInfo.__init__``
    post-processes names (``os.sep`` becomes ``/``, the name is cut at the
    first NUL), so validating ``ZipInfo.filename`` is platform-dependent: a
    ``bad\\name.xml`` entry is refused on POSIX and silently accepted as
    ``bad/name.xml`` on Windows. Second, an invalid UTF-8 name with the UTF-8
    flag set makes ``ZipFile`` raise ``UnicodeDecodeError``, which is not a
    ``BadZipFile`` and would escape as an unstructured exception.
    """
    encoding = "utf-8" if flags & _UTF8_NAME_FLAG_BIT else "cp437"
    try:
        return raw_name.decode(encoding)
    except UnicodeDecodeError:
        return _malformed_archive(
            "entry name is not valid in its declared encoding",
            encoding=encoding,
            raw_name=raw_name.hex(),
        )


def _preflight_central_directory(raw: bytes, limits: ArchiveLimits) -> tuple[str, ...] | Refusal:
    """Walk the actual central-directory records before ``ZipFile`` exists.

    The stdlib parses the central directory by *size*, not by the declared
    entry count, so a hostile archive declaring one entry while shipping a
    directory full of records would otherwise make ``ZipFile`` allocate an
    unbounded object graph. Records are counted here, capped against the
    limit, and reconciled against the end-of-central-directory metadata.

    Concatenated and prefixed archives are rejected structurally: the central
    directory must end exactly where the EOCD record begins, nothing may
    follow the EOCD beyond its declared comment, and the first local file
    header must sit at byte zero.

    Multi-disk and spanned metadata is refused before ``ZipFile`` exists:
    either disk number non-zero, or the per-disk entry count disagreeing with
    the total. ZIP64 sentinels are classified as ``archive_too_many_parts``
    first, so a ZIP64 archive is not misreported as multi-disk.

    Entry names are decoded and validated from the raw record bytes so the
    verdict is identical on every platform (see :func:`_decode_entry_name`).
    Returns the raw names in record order; the caller reconciles them against
    what ``ZipFile`` reports.
    """
    tail = raw[-_EOCD_MAX_SCAN:]
    position = tail.rfind(_EOCD_SIGNATURE)
    if position == -1 or len(tail) - position < _EOCD_FIXED_SIZE:
        return Refusal(
            code="not_a_zip",
            message="No usable end-of-central-directory record was found.",
            details={"reason": "missing_or_truncated_eocd"},
            recovery_options=_RECOVER_VALID_XLSX,
        )
    eocd_absolute = len(raw) - len(tail) + position
    this_disk, cd_disk, disk_entries, total_entries, cd_size, cd_offset = struct.unpack_from(
        "<HHHHII", tail, position + 4
    )
    (comment_len,) = struct.unpack_from("<H", tail, position + 20)
    if eocd_absolute + _EOCD_FIXED_SIZE + comment_len != len(raw):
        return _malformed_archive(
            "trailing data after the end-of-central-directory comment",
            expected_end=eocd_absolute + _EOCD_FIXED_SIZE + comment_len,
            file_size=len(raw),
        )
    if total_entries == _ZIP64_U16 or cd_size == _ZIP64_U32 or cd_offset == _ZIP64_U32:
        return Refusal(
            code="archive_too_many_parts",
            message="The archive uses ZIP64 metadata, which exceeds every supported limit.",
            details={"declared": "zip64"},
            recovery_options=_RECOVER_SMALLER,
        )
    if this_disk != 0 or cd_disk != 0:
        return _malformed_archive(
            "archive spans multiple disks",
            this_disk=this_disk,
            central_directory_disk=cd_disk,
        )
    if disk_entries != total_entries:
        return _malformed_archive(
            "disk entry count disagrees with the total",
            disk_entries=disk_entries,
            total_entries=total_entries,
        )
    if total_entries > limits.max_part_count:
        return Refusal(
            code="archive_too_many_parts",
            message=f"Archive declares {total_entries} entries; the limit is "
            f"{limits.max_part_count}.",
            details={"declared_entries": total_entries, "limit": limits.max_part_count},
            recovery_options=_RECOVER_SMALLER,
        )
    if cd_offset + cd_size != eocd_absolute:
        return _malformed_archive(
            "central directory does not end at the end-of-central-directory record",
            cd_end=cd_offset + cd_size,
            eocd_position=eocd_absolute,
        )

    names: list[str] = []
    offset = cd_offset
    end = cd_offset + cd_size
    min_local_offset: int | None = None
    while offset < end:
        if raw[offset : offset + 4] != _CENTRAL_HEADER_SIGNATURE:
            return _malformed_archive("record signature mismatch", at_offset=offset)
        if end - offset < _CENTRAL_HEADER_FIXED_SIZE:
            return _malformed_archive("truncated central-directory record", at_offset=offset)
        (flags,) = struct.unpack_from("<H", raw, offset + 8)
        name_len, extra_len, record_comment_len = struct.unpack_from("<HHH", raw, offset + 28)
        (local_offset,) = struct.unpack_from("<I", raw, offset + 42)
        name_start = offset + _CENTRAL_HEADER_FIXED_SIZE
        if name_start + name_len > end:
            return _malformed_archive("entry name overruns the central directory", at_offset=offset)
        name = _decode_entry_name(raw[name_start : name_start + name_len], flags)
        if isinstance(name, Refusal):
            return name
        name_refusal = _validate_entry_name(name)
        if name_refusal is not None:
            return name_refusal
        if min_local_offset is None or local_offset < min_local_offset:
            min_local_offset = local_offset
        offset = name_start + name_len + extra_len + record_comment_len
        names.append(name)
        if len(names) > limits.max_part_count:
            return Refusal(
                code="archive_too_many_parts",
                message=f"Archive contains more than {limits.max_part_count} "
                "central-directory records.",
                details={"limit": limits.max_part_count},
                recovery_options=_RECOVER_SMALLER,
            )
    if offset != end:
        return _malformed_archive("trailing bytes inside the central directory")
    if len(names) != total_entries:
        return _malformed_archive(
            "record count does not match the declared entry count",
            actual_records=len(names),
            declared_entries=total_entries,
        )
    if names and min_local_offset != 0:
        return _malformed_archive(
            "data precedes the first local file header (prefixed or concatenated archive)",
            first_local_offset=min_local_offset,
        )
    return tuple(names)


class WorkbookArchive:
    """A validated, read-only, in-memory view of one ``.xlsx`` package.

    Construct via :meth:`load`. The source file handle is closed before
    :meth:`load` returns; the archive operates on an in-memory snapshot, so
    later modification of the source file cannot affect this session. The
    SHA-256 of those snapshot bytes is computed at load and retained as a
    digest only.
    """

    def __init__(
        self,
        source_path: Path,
        zip_file: zipfile.ZipFile,
        part_order: tuple[str, ...],
        limits: ArchiveLimits,
        source_fingerprint: str,
    ) -> None:
        self._source_path = source_path
        self._zip = zip_file
        self._part_order = part_order
        self._part_set = frozenset(part_order)
        self._limits = limits
        self._source_fingerprint = source_fingerprint
        self._cache: dict[str, bytes] = {}
        self._closed = False

    @classmethod
    def load(cls, path: Path, limits: ArchiveLimits = DEFAULT_LIMITS) -> WorkbookArchive | Refusal:
        if not path.is_file():
            return Refusal(
                code="invalid_path",
                message=f"{str(path)!r} does not exist or is not a regular file.",
                details={"path": str(path)},
                recovery_options=({"action": "verify_file_path"},),
            )

        # One open and one bounded read: the size is enforced on the bytes
        # actually read, so the file growing after a stat cannot bypass the
        # limit (TOCTOU), and filesystem errors become refusals.
        try:
            with path.open("rb") as handle:
                raw = handle.read(limits.max_file_bytes + 1)
        except OSError as exc:
            return Refusal(
                code="unreadable_file",
                message=f"The file could not be read: {exc}.",
                details={"path": str(path), "error": str(exc)},
                recovery_options=({"action": "fix_file_permissions"},),
            )
        if len(raw) > limits.max_file_bytes:
            return Refusal(
                code="archive_too_large",
                message=f"File exceeds the {limits.max_file_bytes}-byte limit.",
                details={"limit_bytes": limits.max_file_bytes},
                recovery_options=_RECOVER_SMALLER,
            )

        if raw.startswith(_CFB_MAGIC):
            return Refusal(
                code="unsupported_container",
                message=(
                    "This is a Compound File Binary container: either a legacy .xls "
                    "workbook or an encrypted .xlsx. Neither is supported."
                ),
                details={"container": "cfb"},
                recovery_options=_RECOVER_UNENCRYPTED,
            )
        if not raw.startswith(_ZIP_MAGIC_PREFIXES):
            return Refusal(
                code="not_a_zip",
                message="The file does not begin with a ZIP local-file or archive signature.",
                details={"path": str(path)},
                recovery_options=_RECOVER_VALID_XLSX,
            )

        raw_names = _preflight_central_directory(raw, limits)
        if isinstance(raw_names, Refusal):
            return raw_names

        zip_file: zipfile.ZipFile | None = None
        try:
            try:
                zip_file = zipfile.ZipFile(io.BytesIO(raw))
            except zipfile.BadZipFile as exc:
                # zip_file was never assigned; there is nothing to close.
                return Refusal(
                    code="not_a_zip",
                    message=f"The ZIP structure could not be read: {exc}.",
                    details={"error": str(exc)},
                    recovery_options=_RECOVER_VALID_XLSX,
                )
            infos = zip_file.infolist()
            if len(infos) != len(raw_names):
                return _close_on_refusal(
                    zip_file,
                    _malformed_archive(
                        "the stdlib parsed a different entry count than the central-directory walk",
                        stdlib_entries=len(infos),
                        walked_records=len(raw_names),
                    ),
                )
            seen: set[str] = set()
            order: list[str] = []
            declared_total = 0
            for info, raw_name in zip(infos, raw_names, strict=True):
                name = info.filename
                if name != raw_name:
                    # The raw names were already validated; the stdlib must not have
                    # rewritten any of them, or downstream code would operate on a
                    # different name than the one the archive actually carries.
                    return _close_on_refusal(
                        zip_file,
                        _malformed_archive(
                            "the stdlib reported an entry name that differs from the raw record",
                            raw_name=raw_name,
                            stdlib_name=name,
                        ),
                    )
                if info.flag_bits & _ENCRYPTED_FLAG_BIT:
                    return _close_on_refusal(
                        zip_file,
                        Refusal(
                            code="encrypted_archive_entry",
                            message=f"Archive entry {name!r} is encrypted.",
                            details={"entry": name},
                            recovery_options=_RECOVER_UNENCRYPTED,
                        ),
                    )
                if info.is_dir():
                    continue
                if info.compress_type not in _SUPPORTED_COMPRESSION:
                    return _close_on_refusal(
                        zip_file,
                        Refusal(
                            code="unsupported_compression",
                            message=f"Archive entry {name!r} uses compression method "
                            f"{info.compress_type}; only stored and deflated entries "
                            "are supported.",
                            details={"entry": name, "compress_type": info.compress_type},
                            recovery_options=_RECOVER_VALID_XLSX,
                        ),
                    )
                if name in seen:
                    # Duplicate names are a smuggling vector: ZIP consumers disagree
                    # about which copy wins, so the archive is rejected outright.
                    return _close_on_refusal(
                        zip_file,
                        Refusal(
                            code="duplicate_archive_entry",
                            message=f"Archive entry {name!r} appears more than once.",
                            details={"entry": name},
                            recovery_options=_RECOVER_TRUSTED_SOURCE,
                        ),
                    )
                seen.add(name)
                order.append(name)

                if info.file_size > limits.max_part_bytes:
                    return _close_on_refusal(
                        zip_file,
                        Refusal(
                            code="archive_part_too_large",
                            message=(
                                f"Part {name!r} declares {info.file_size} uncompressed bytes; "
                                f"the limit is {limits.max_part_bytes}."
                            ),
                            details={
                                "entry": name,
                                "declared_bytes": info.file_size,
                                "limit_bytes": limits.max_part_bytes,
                            },
                            recovery_options=_RECOVER_SMALLER,
                        ),
                    )
                declared_total += info.file_size
                if declared_total > limits.max_total_bytes:
                    return _close_on_refusal(
                        zip_file,
                        Refusal(
                            code="archive_expansion_too_large",
                            message=(
                                f"Total declared uncompressed size exceeds "
                                f"{limits.max_total_bytes} bytes."
                            ),
                            details={
                                "declared_total_bytes": declared_total,
                                "limit_bytes": limits.max_total_bytes,
                            },
                            recovery_options=_RECOVER_SMALLER,
                        ),
                    )

            return cls(
                source_path=path,
                zip_file=zip_file,
                part_order=tuple(order),
                limits=limits,
                source_fingerprint=f"sha256:{hashlib.sha256(raw).hexdigest()}",
            )
        except BaseException:
            # Cleanup also covers interruption/exit, without swallowing it.
            # Refusal paths already closed; every other failure propagates
            # unchanged after releasing the handle we still own.
            if zip_file is not None:
                zip_file.close()
            raise

    @property
    def source_path(self) -> Path:
        return self._source_path

    @property
    def source_fingerprint(self) -> str:
        """``sha256:`` plus the lowercase hex digest of the loaded snapshot."""
        return self._source_fingerprint

    def part_names(self) -> tuple[str, ...]:
        """All file parts, in archive order (deterministic for identical bytes)."""
        return self._part_order

    def has_part(self, name: str) -> bool:
        return name in self._part_set

    def read_part(self, name: str) -> bytes | Refusal:
        """Return a part's bytes, decompressing on first access.

        Declared sizes were validated at load time; ``ZipFile.read`` enforces
        the declared size and CRC during decompression, so a part cannot
        silently exceed the validated limits, and integrity failures become
        refusals rather than exceptions.
        """
        if name not in self._part_set:
            raise KeyError(name)
        cached = self._cache.get(name)
        if cached is not None:
            return cached
        try:
            data = self._zip.read(name)
        except (zipfile.BadZipFile, zlib.error, EOFError, OSError) as exc:
            return Refusal(
                code="malformed_archive",
                message=f"Part {name!r} failed integrity checks during decompression: {exc}.",
                details={"entry": name, "error": str(exc)},
                recovery_options=_RECOVER_TRUSTED_SOURCE,
            )
        self._cache[name] = data
        return data

    def parse_xml_part(self, name: str) -> ET.Element | Refusal:
        data = self.read_part(name)
        if isinstance(data, Refusal):
            return data
        preflight = _preflight_xml(data, name, self._limits)
        if preflight is not None:
            return preflight
        try:
            # Preflight above already rejected DOCTYPE and enforced limits.
            return ET.fromstring(data)  # noqa: S314
        except ET.ParseError as exc:
            return Refusal(
                code="malformed_xml",
                message=f"Part {name!r} is not well-formed XML: {exc}.",
                details={"part": name, "error": str(exc)},
                recovery_options=_RECOVER_TRUSTED_SOURCE,
            )

    def close(self) -> None:
        """Release the in-memory ZIP. Idempotent."""
        if not self._closed:
            self._zip.close()
            self._cache.clear()
            self._closed = True

    @property
    def closed(self) -> bool:
        return self._closed


def _structurally_unsafe_reference(value: str) -> bool:
    if (
        not value
        or "\\" in value
        or "?" in value
        or "#" in value
        or any(ord(character) < 0x20 for character in value)
    ):
        return True
    # The scheme check is anchored and the authority check needs a leading
    # "//", so a package-root slash would hide both from the raw string while
    # normalize_package_path strips that slash and leaves the prefix intact
    # ("/C:/evil.xml" -> "C:/evil.xml"). Check the stripped form too.
    return any(
        not form or form.startswith("//") or _URI_SCHEME_RE.match(form) is not None
        for form in (value, value.lstrip("/"))
    )


def validate_opc_target(target: str) -> str | None:
    """Two-phase validation of a raw OPC relationship target.

    OPC targets are URI references, not filesystem paths. Phase one applies
    structural checks to the raw string: schemes (which also catches drive
    prefixes), authorities, queries, fragments, backslashes, control
    characters, and percent-encodings that could smuggle a separator or dot
    segment (``%2e``, ``%2f``, ``%5c``, and ``%25`` against double-encoding).
    Phase two percent-decodes the remainder and applies the same structural
    checks again, so an encoding cannot reintroduce a rejected construct —
    for example ``%3F`` decoding to ``?``.

    This must run on the target *before* any base-path joining; joining first
    would mask the scheme and authority checks behind the base prefix.
    Returns the decoded target, or ``None`` when unsafe.
    """
    if _structurally_unsafe_reference(target):
        return None
    if _UNSAFE_PERCENT_ENCODING_RE.search(target):
        return None
    decoded = urllib.parse.unquote(target) if "%" in target else target
    if _structurally_unsafe_reference(decoded):
        return None
    return decoded


def normalize_package_path(candidate: str) -> str | None:
    """Normalize an already-validated, base-joined path to a package path.

    Retains its own structural guards as defense in depth, then resolves dot
    segments. Returns ``None`` when the path is unsafe or would escape the
    package root, so callers treat hostile paths as unresolvable rather than
    following them. Percent-decoding is deliberately not performed here; that
    happens once, pre-join, in :func:`validate_opc_target`.
    """
    if _structurally_unsafe_reference(candidate):
        return None
    normalized = posixpath.normpath(candidate.lstrip("/"))
    if normalized in {"", "."} or normalized.startswith("../"):
        return None
    return normalized
