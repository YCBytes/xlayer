"""Number-format styles: date, time, non-temporal, or unknown.

Adapted from ``xlayer-core`` ``parser/styles.py``. The built-in ID tables
and metadata-stripping approach are retained; the classifier is not a
verbatim copy:

- one scanner (``_scan``) decides what is syntax and what is display text,
  so section splitting, literal removal, and elapsed-token detection cannot
  disagree: a backslash escapes the next character (``\\"yyyy-mm-dd`` is a
  date, not an unterminated quote), quoted text is inert (``0 "[h]"`` is
  literal text, not elapsed hours), and ``;`` separates sections only
  outside literals and brackets;
- ``m`` / ``mm`` are minutes only when adjacent to an hour or second
  token, so ``mmm h:mm`` is date-time rather than time-only; elapsed
  brackets keep their position, so ``[h]:mm`` is time, not date-time;
- every unquoted format section is classified, and disagreement
  (``[>=45000]0;yyyy-mm-dd``) is ``unknown``, not the first section;
- unresolved IDs and excluded locale IDs (27-36, 50-58) without an
  explicit ``formatCode`` are ``unknown``, never silently ``non_temporal``.

A ``non_temporal`` result means the format is known not to be a date or
time display. It does not prove the cell value is a number.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from xml.etree import ElementTree as ET

from xlayer._errors import Refusal
from xlayer._ooxml.archive import WorkbookArchive
from xlayer._ooxml.package import TARGET_MODE_INTERNAL, PackageInfo

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_STYLES_ROOT_TAG = f"{{{_NS}}}styleSheet"
_NUM_FMTS_TAG = f"{{{_NS}}}numFmts"
_NUM_FMT_TAG = f"{{{_NS}}}numFmt"
_CELL_XFS_TAG = f"{{{_NS}}}cellXfs"
_XF_TAG = f"{{{_NS}}}xf"

KIND_DATE = "date"
KIND_TIME = "time"
KIND_DATE_TIME = "date-time"
KIND_NON_TEMPORAL = "non_temporal"
KIND_UNKNOWN = "unknown"

BUILTIN_DATE_FORMAT_IDS = frozenset({14, 15, 16, 17})
BUILTIN_TIME_FORMAT_IDS = frozenset({18, 19, 20, 21, 45, 46, 47})
BUILTIN_DATETIME_FORMAT_IDS = frozenset({22})
LOCALE_FORMAT_IDS = frozenset(range(27, 37)) | frozenset(range(50, 59))

BUILTIN_FORMATS: dict[int, str] = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
}

_MAX_UINT = 4_294_967_295
_MAX_UINT_DIGITS = 10
_DATETIME_TOKEN_RE = re.compile(r"am/pm|a/p|y+|m+|d+|h+|s+", re.IGNORECASE)
_ELAPSED_INNER_RE = re.compile(r"^(h{1,2}|m{1,2}|s{1,2})$", re.IGNORECASE)

_RECOVER_TRUSTED_SOURCE = ({"action": "regenerate_workbook_from_trusted_source"},)


@dataclass(frozen=True)
class StyleInfo:
    format_id: int
    format_code: str
    is_date: bool
    is_time: bool
    temporal_kind: str


_GENERAL = StyleInfo(
    format_id=0,
    format_code="General",
    is_date=False,
    is_time=False,
    temporal_kind=KIND_NON_TEMPORAL,
)


def classify_number_format(format_id: int, format_code: str | None) -> str:
    """Return the temporal kind for one number format.

    When ``format_code`` is present it is authoritative, including a custom
    override of a built-in id. When it is absent, only known built-in ids
    are classified; locale-specific and unresolved ids are ``unknown``.
    """
    if format_code is not None:
        return _classify_format_code(format_code)
    if format_id in BUILTIN_DATE_FORMAT_IDS:
        return KIND_DATE
    if format_id in BUILTIN_TIME_FORMAT_IDS:
        return KIND_TIME
    if format_id in BUILTIN_DATETIME_FORMAT_IDS:
        return KIND_DATE_TIME
    if format_id in BUILTIN_FORMATS:
        return _classify_format_code(BUILTIN_FORMATS[format_id])
    return KIND_UNKNOWN


def parse_styles(archive: WorkbookArchive, package: PackageInfo) -> tuple[StyleInfo, ...] | Refusal:
    """Return the cellXf style table, or a structured refusal."""
    rel = package.find_workbook_rel("styles")
    if rel is None:
        return (_GENERAL,)
    if rel.target_mode != TARGET_MODE_INTERNAL:
        return _invalid_part(
            package.workbook_path, "styles relationship is external", target=rel.target
        )
    part = package.resolve_workbook_rel_by_id(rel.rel_id, "styles")
    if part is None:
        return _unsafe_target(rel.target)
    if not archive.has_part(part):
        return _missing_part(part, "the styles relationship points at it")

    root = archive.parse_xml_part(part)
    if isinstance(root, Refusal):
        return root
    if root.tag != _STYLES_ROOT_TAG:
        return _invalid_part(part, "unexpected root element", root_tag=root.tag)

    custom_formats = _parse_num_fmts(root, part)
    if isinstance(custom_formats, Refusal):
        return custom_formats

    cell_xfs = root.find(_CELL_XFS_TAG)
    if cell_xfs is None:
        return (_GENERAL,)

    styles: list[StyleInfo] = []
    for xf in cell_xfs.findall(_XF_TAG):
        raw_id = xf.attrib.get("numFmtId", "0")
        format_id = _parse_unsigned_int(raw_id)
        if format_id is None:
            return _invalid_part(part, "invalid numFmtId", numFmtId=raw_id)
        if format_id in custom_formats:
            format_code: str | None = custom_formats[format_id]
        elif format_id in BUILTIN_FORMATS:
            format_code = BUILTIN_FORMATS[format_id]
        else:
            format_code = None
        kind = classify_number_format(format_id, format_code)
        styles.append(
            StyleInfo(
                format_id=format_id,
                format_code=format_code if format_code is not None else "",
                is_date=kind in {KIND_DATE, KIND_DATE_TIME},
                is_time=kind in {KIND_TIME, KIND_DATE_TIME},
                temporal_kind=kind,
            )
        )
    if not styles:
        return (_GENERAL,)
    return tuple(styles)


def _parse_num_fmts(root: ET.Element, part: str) -> dict[int, str] | Refusal:
    custom: dict[int, str] = {}
    num_fmts = root.find(_NUM_FMTS_TAG)
    if num_fmts is None:
        return custom
    for num_fmt in num_fmts.findall(_NUM_FMT_TAG):
        raw_id = num_fmt.attrib.get("numFmtId")
        if not raw_id:
            return _invalid_part(part, "numFmt missing numFmtId")
        format_id = _parse_unsigned_int(raw_id)
        if format_id is None:
            return _invalid_part(part, "invalid numFmtId", numFmtId=raw_id)
        format_code = num_fmt.attrib.get("formatCode")
        if not format_code:
            return _invalid_part(part, "numFmt missing formatCode")
        if format_id in custom:
            return _invalid_part(part, "duplicate numFmtId", numFmtId=format_id)
        custom[format_id] = format_code
    return custom


def _classify_format_code(code: str) -> str:
    kinds: set[str] = set()
    for section in _token_sections(_scan(code)):
        kind = _classify_section(section)
        if kind is not None:
            kinds.add(kind)
    if not kinds:
        return KIND_NON_TEMPORAL
    if len(kinds) > 1:
        return KIND_UNKNOWN
    return next(iter(kinds))


def split_format_sections(code: str) -> list[str]:
    """Split a format code on real section separators.

    A ``;`` inside quotes, inside brackets, or preceded by a backslash is
    part of the section, not a separator. Each returned section is the
    exact source text between separators.
    """
    return ["".join(raw for _, raw, _ in section) for section in _token_sections(_scan(code))]


# Scanner token kinds. Every helper that needs to know whether a character is
# syntax or display text goes through ``_scan`` so they cannot disagree.
_LITERAL = "literal"  # quoted text, backslash-escaped char, ``_x`` / ``*x`` fill
_BRACKET = "bracket"  # ``[...]`` colour, condition, locale, or elapsed token
_SEPARATOR = "separator"
_TEXT = "text"  # a single format character

_Token = tuple[str, str, str]  # (kind, raw source text, payload)


def _scan(code: str) -> list[_Token]:
    tokens: list[_Token] = []
    index = 0
    length = len(code)
    while index < length:
        char = code[index]
        if char in "\\_*":
            raw = code[index : index + 2]
            tokens.append((_LITERAL, raw, raw[1:]))
            index += 2
        elif char == '"':
            end = code.find('"', index + 1)
            if end == -1:
                tokens.append((_LITERAL, code[index:], code[index + 1 :]))
                index = length
            else:
                tokens.append((_LITERAL, code[index : end + 1], code[index + 1 : end]))
                index = end + 1
        elif char == "[":
            end = code.find("]", index + 1)
            if end == -1:
                # Unterminated bracket: Excel rejects the code; keep the
                # character inert rather than swallowing the rest.
                tokens.append((_LITERAL, char, char))
                index += 1
            else:
                tokens.append((_BRACKET, code[index : end + 1], code[index + 1 : end]))
                index = end + 1
        elif char == ";":
            tokens.append((_SEPARATOR, char, char))
            index += 1
        else:
            tokens.append((_TEXT, char, char))
            index += 1
    return tokens


def _token_sections(tokens: list[_Token]) -> list[list[_Token]]:
    sections: list[list[_Token]] = [[]]
    for token in tokens:
        if token[0] == _SEPARATOR:
            sections.append([])
        else:
            sections[-1].append(token)
    return sections


def _classify_section(tokens: list[_Token]) -> str | None:
    """Classify one section, or ``None`` when it displays no format text."""
    # Ordered stream of date/time tokens. Elapsed brackets are kept in
    # position, prefixed with "[", so ``[h]`` is hour context for ``mm``.
    stream: list[str] = []
    text_chars: list[str] = []
    pending: list[str] = []

    def flush() -> None:
        run = "".join(pending).lower()
        stream.extend(match.group(0) for match in _DATETIME_TOKEN_RE.finditer(run))
        pending.clear()

    for kind, _raw, payload in tokens:
        if kind == _TEXT:
            pending.append(payload)
            text_chars.append(payload)
            continue
        flush()
        if kind == _BRACKET and _ELAPSED_INNER_RE.match(payload.strip()):
            stream.append("[" + payload.strip().lower())
    flush()

    if not stream:
        return None if "".join(text_chars).strip() == "" else KIND_NON_TEMPORAL

    has_date = False
    has_time = False
    for index, token in enumerate(stream):
        if token[0] == "[":
            has_time = True
        elif token[0] in {"y", "d"}:
            has_date = True
        elif token[0] in {"h", "s"} or token in {"am/pm", "a/p"}:
            has_time = True
        elif len(token) >= 3 or not _is_minute_token(stream, index):
            has_date = True
        else:
            has_time = True

    if has_date and has_time:
        return KIND_DATE_TIME
    if has_date:
        return KIND_DATE
    return KIND_TIME


def _is_minute_token(stream: list[str], index: int) -> bool:
    previous = stream[index - 1].lstrip("[") if index > 0 else ""
    following = stream[index + 1].lstrip("[") if index + 1 < len(stream) else ""
    return previous.startswith("h") or following.startswith("s")


def _parse_unsigned_int(raw: str) -> int | None:
    if not raw.isascii() or not raw.isdigit() or len(raw) > _MAX_UINT_DIGITS:
        return None
    try:
        value = int(raw)
    except ValueError:
        return None
    if value > _MAX_UINT:
        return None
    return value


def _invalid_part(part: str, reason: str, **details: object) -> Refusal:
    return Refusal(
        code="invalid_part_content",
        message=f"Part {part!r} is structurally invalid: {reason}.",
        details={"part": part, "reason": reason, **details},
        recovery_options=_RECOVER_TRUSTED_SOURCE,
    )


def _missing_part(part: str, why: str) -> Refusal:
    return Refusal(
        code="missing_required_part",
        message=f"Required part {part!r} is missing or unusable: {why}.",
        details={"part": part, "reason": why},
        recovery_options=_RECOVER_TRUSTED_SOURCE,
    )


def _unsafe_target(target: str) -> Refusal:
    return Refusal(
        code="unsafe_archive_path",
        message=f"The styles relationship target {target!r} is not a safe package path.",
        details={"kind": "relationship_target", "target": target},
        recovery_options=_RECOVER_TRUSTED_SOURCE,
    )
