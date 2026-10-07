"""Namespace-aware UTF-8 byte spans and narrow scalar payload patches.

Expat supplies structural boundaries; byte scans only locate lexical tag/attribute
ends inside those boundaries. This is not a regex XML parser or tree serializer.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import TYPE_CHECKING
from xml.parsers import expat
from xml.sax.saxutils import escape

from xlayer._approval import refusal
from xlayer._errors import Refusal
from xlayer._ooxml.sheet import _column_letter, _parse_address

if TYPE_CHECKING:
    from xlayer._edit_validation import EditPlan

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_ATTR = re.compile(rb"([^\s=<>/]+)\s*=\s*([\"'])(.*?)\2", re.DOTALL)
_EXCEL_ESCAPE = re.compile(r"_(?=x[0-9a-fA-F]{4}_)")


@dataclass
class XmlSpan:
    tag: str
    qname: bytes
    attrs: dict[str, str]
    start: int
    open_end: int
    end: int = 0
    close_start: int = 0
    self_closing: bool = False
    special_markup: bool = False
    children: list[XmlSpan] = field(default_factory=list)


def tag_end(raw: bytes, start: int) -> int:
    quote = 0
    for index in range(start, len(raw)):
        char = raw[index]
        if quote:
            if char == quote:
                quote = 0
        elif char in (34, 39):
            quote = char
        elif char == 62:
            return index + 1
    raise ValueError("unterminated tag in validated XML")


def cell_spans(raw: bytes) -> dict[str, XmlSpan] | Refusal:
    if b"\x00" in raw or raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return refusal("unsupported_write_encoding", "worksheet editing requires UTF-8")
    try:
        raw.decode("utf-8")
    except UnicodeDecodeError:
        return refusal("unsupported_write_encoding", "worksheet editing requires UTF-8")
    declaration = re.match(rb"(?:\xef\xbb\xbf)?<\?xml\b(.*?)\?>", raw, re.DOTALL)
    if declaration is not None:
        encoding = re.search(rb"encoding\s*=\s*([\"'])(.*?)\1", declaration[1])
        if encoding is not None and encoding[2].lower() not in {b"utf-8", b"utf8"}:
            return refusal("unsupported_write_encoding", "worksheet declares a non-UTF-8 encoding")
    parser = expat.ParserCreate(namespace_separator="}")
    stack: list[XmlSpan] = []
    result: dict[str, XmlSpan] = {}
    row_number = 0
    column = 0

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal row_number, column
        offset = parser.CurrentByteIndex
        end = tag_end(raw, offset)
        lexical = raw[offset:end]
        qname = re.split(rb"[\s/>]", lexical[1:], maxsplit=1)[0]
        span = XmlSpan(
            name, qname, attrs, offset, end, self_closing=lexical.rstrip().endswith(b"/>")
        )
        if stack:
            stack[-1].children.append(span)
        if name == NS + "}row" and stack and stack[-1].tag == NS + "}sheetData":
            row_number = int(attrs.get("r", str(row_number + 1)))
            column = 0
        if name == NS + "}c" and stack and stack[-1].tag == NS + "}row":
            address = attrs.get("r", f"{_column_letter(column + 1)}{row_number}")
            coordinate = _parse_address(address)
            if coordinate is None or address in result:
                raise ValueError("cell identity does not map uniquely")
            column = coordinate[1]
            result[address] = span
        stack.append(span)

    def end(_name: str) -> None:
        span = stack.pop()
        span.close_start = span.open_end if span.self_closing else parser.CurrentByteIndex
        span.end = span.open_end if span.self_closing else tag_end(raw, span.close_start)

    def special(*_args: object) -> None:
        for span in stack:
            span.special_markup = True

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CommentHandler = special
    parser.ProcessingInstructionHandler = special
    try:
        parser.Parse(raw, True)
    except (expat.ExpatError, ValueError):
        return refusal(
            "unsupported_target_structure", "cannot map validated worksheet to lexical cells"
        )
    return result


def encode_text(value: str) -> str:
    """Protect escape-looking literals before escaping XML-forbidden scalars/CR."""
    # Lookahead also protects escape starts sharing the preceding escape's
    # closing underscore; a non-overlapping whole-escape regex misses those.
    protected = _EXCEL_ESCAPE.sub("_x005F_", value)
    pieces: list[str] = []
    for char in protected:
        number = ord(char)
        if char == "\r" or (number < 32 and char not in "\n\t") or number in {0xFFFE, 0xFFFF}:
            pieces.append(f"_x{number:04X}_")
        else:
            pieces.append(escape(char))
    return "".join(pieces)


def _replacement(raw: bytes, span: XmlSpan, plan: EditPlan) -> bytes | Refusal:
    if span.special_markup:
        return refusal(
            "unsupported_target_structure", "comment or processing instruction inside target"
        )
    opening = raw[span.start : span.open_end]
    attribute = next((m for m in _ATTR.finditer(opening) if m[1] == b"t"), None)
    if attribute is not None:
        opening = (
            opening[: attribute.start(3)]
            + plan.expected_type.encode()
            + opening[attribute.end(3) :]
        )
    else:
        offset = len(opening) - (2 if span.self_closing else 1)
        opening = opening[:offset] + b' t="' + plan.expected_type.encode() + b'"' + opening[offset:]
    prefix = span.qname.rpartition(b":")[0]
    prefix = prefix + b":" if prefix else b""
    if plan.expected_type == "inlineStr":
        text = encode_text(str(plan.expected_value)).encode("utf-8")
        payload = (
            b"<"
            + prefix
            + b"is><"
            + prefix
            + b't xml:space="preserve">'
            + text
            + b"</"
            + prefix
            + b"t></"
            + prefix
            + b"is>"
        )
    else:
        payload = (
            b"<" + prefix + b"v>" + str(plan.expected_raw).encode("ascii") + b"</" + prefix + b"v>"
        )
    if span.self_closing:
        return opening[:-2] + b">" + payload + b"</" + span.qname + b">"
    body = raw[span.open_end : span.close_start]
    old_payload = [child for child in span.children if child.tag in {NS + "}v", NS + "}is"}]
    if len(old_payload) > 1 or len(span.children) != len(old_payload):
        return refusal("unsupported_target_structure", "competing or unsupported target payload")
    if old_payload:
        old = old_payload[0]
        body = raw[span.open_end : old.start] + payload + raw[old.end : span.close_start]
    else:
        body += payload
    return opening + body + raw[span.close_start : span.end]


def patch_worksheet(raw: bytes, plans: Sequence[EditPlan]) -> bytes | Refusal:
    spans = cell_spans(raw)
    if isinstance(spans, Refusal):
        return spans
    replacements: list[tuple[int, int, bytes]] = []
    for plan in plans:
        if not plan.effective:
            continue
        span = spans.get(plan.edit.cell)
        if span is None:
            return refusal(
                "unsupported_target_structure", "target lacks a unique lexical cell span"
            )
        replacement = _replacement(raw, span, plan)
        if isinstance(replacement, Refusal):
            return replacement
        replacements.append((span.start, span.end, replacement))
    for start, end, replacement in sorted(replacements, reverse=True):
        raw = raw[:start] + replacement + raw[end:]
    return raw
