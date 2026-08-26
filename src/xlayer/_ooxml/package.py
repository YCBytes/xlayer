"""OOXML package resolution: content types and relationships.

Adapted from ``xlayer-core`` ``parser/package.py``. Differences from core are
deliberate public architecture, not drift:

- parts are read through a :class:`~xlayer._ooxml.archive.WorkbookArchive`
  instead of an extracted directory;
- root elements and their namespaces are matched exactly, and relationship
  types are matched against exact supported URIs, never by suffix or local
  name;
- ``TargetMode`` is modeled and validated; external relationships are
  preserved as inert metadata but never resolved as package parts and never
  fetched;
- the package must actually be a transitional spreadsheet workbook: Word,
  PowerPoint, macro-enabled, template, and Strict OOXML packages receive
  honest structured refusals; and
- schema-level problems in attacker-controlled XML (missing or empty
  attributes, duplicate relationship IDs, duplicate content-type keys,
  multiple or mixed officeDocument relationships) return refusals instead of
  raising or silently picking a winner.

No file path is hardcoded; all paths are resolved from relationships.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from types import MappingProxyType
from typing import TYPE_CHECKING

from xlayer._errors import Refusal
from xlayer._ooxml.archive import (
    CONTENT_TYPES_PART,
    WorkbookArchive,
    normalize_package_path,
    validate_opc_target,
)

if TYPE_CHECKING:
    from xml.etree import ElementTree as ET

_NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
_NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships"

_CONTENT_TYPES_ROOT_TAG = f"{{{_NS_CONTENT_TYPES}}}Types"
_CONTENT_TYPES_OVERRIDE_TAG = f"{{{_NS_CONTENT_TYPES}}}Override"
_CONTENT_TYPES_DEFAULT_TAG = f"{{{_NS_CONTENT_TYPES}}}Default"
_RELATIONSHIPS_ROOT_TAG = f"{{{_NS_RELATIONSHIPS}}}Relationships"
_RELATIONSHIP_TAG = f"{{{_NS_RELATIONSHIPS}}}Relationship"

_OFFICE_REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"

REL_TYPE_OFFICE_DOCUMENT = _OFFICE_REL_BASE + "officeDocument"

# Strict OOXML uses a different relationship namespace; detected explicitly so
# the refusal is honest rather than a misleading "missing part".
_STRICT_REL_TYPE_OFFICE_DOCUMENT = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
)

# The only package kind slice 1 supports: a transitional .xlsx workbook.
SUPPORTED_WORKBOOK_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
)

_KNOWN_UNSUPPORTED_MAIN_CONTENT_TYPES = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml": (
        "word_document"
    ),
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml": (
        "presentation"
    ),
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml": "macro_enabled_workbook",
    "application/vnd.ms-excel.template.macroEnabled.main+xml": "macro_enabled_template",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml": (
        "spreadsheet_template"
    ),
}

# Exact relationship-type URIs this package will resolve from the workbook's
# relationships. Matching is by full URI equality, never by suffix.
SUPPORTED_WORKBOOK_REL_TYPES = {
    "worksheet": _OFFICE_REL_BASE + "worksheet",
    "styles": _OFFICE_REL_BASE + "styles",
    "sharedStrings": _OFFICE_REL_BASE + "sharedStrings",
    "calcChain": _OFFICE_REL_BASE + "calcChain",
    "theme": _OFFICE_REL_BASE + "theme",
}

_PACKAGE_RELS_PART = "_rels/.rels"

TARGET_MODE_INTERNAL = "Internal"
TARGET_MODE_EXTERNAL = "External"
_VALID_TARGET_MODES = frozenset({TARGET_MODE_INTERNAL, TARGET_MODE_EXTERNAL})

_RECOVER_TRUSTED_SOURCE = ({"action": "regenerate_workbook_from_trusted_source"},)


@dataclass(frozen=True)
class Relationship:
    rel_id: str
    rel_type: str
    target: str  # as declared; resolve through PackageInfo before use
    target_mode: str = TARGET_MODE_INTERNAL


@dataclass(frozen=True)
class PackageInfo:
    """Resolved package structure. All mappings are read-only copies."""

    workbook_path: str  # e.g. "xl/workbook.xml"
    workbook_dir: str  # e.g. "xl"
    # part path with leading slash, as declared -> content type
    content_types: Mapping[str, str] = field(default_factory=dict)
    # file extension -> default content type
    default_types: Mapping[str, str] = field(default_factory=dict)
    # r:id -> Relationship from _rels/.rels
    top_rels: Mapping[str, Relationship] = field(default_factory=dict)
    # r:id -> Relationship from xl/_rels/workbook.xml.rels
    workbook_rels: Mapping[str, Relationship] = field(default_factory=dict)

    def __post_init__(self) -> None:
        for name in ("content_types", "default_types", "top_rels", "workbook_rels"):
            object.__setattr__(self, name, MappingProxyType(dict(getattr(self, name))))

    def _resolve_target(self, rel: Relationship) -> str | None:
        # External targets are URIs, not package parts. They are preserved as
        # metadata on the Relationship but are never resolved and never
        # fetched.
        if rel.target_mode != TARGET_MODE_INTERNAL:
            return None
        # The raw target is validated as a URI reference BEFORE base joining;
        # joining first would mask scheme/authority checks behind the base
        # prefix. validate_opc_target also re-checks after percent-decoding.
        cleaned = validate_opc_target(rel.target)
        if cleaned is None:
            return None
        candidate = cleaned if cleaned.startswith("/") else f"{self.workbook_dir}/{cleaned}"
        return normalize_package_path(candidate)

    def resolve_workbook_rel(self, rel_kind: str) -> str | None:
        """Resolve a supported workbook relationship kind (e.g. ``"styles"``).

        ``rel_kind`` must be a key of :data:`SUPPORTED_WORKBOOK_REL_TYPES`;
        anything else is a programming error. Returns a package-relative
        path, or ``None`` when the relationship is absent, external, or its
        target is unresolvable.
        """
        exact_type = SUPPORTED_WORKBOOK_REL_TYPES[rel_kind]
        for rel in self.workbook_rels.values():
            if rel.rel_type == exact_type:
                return self._resolve_target(rel)
        return None

    def resolve_workbook_rel_by_id(self, rel_id: str, expected_kind: str) -> str | None:
        """Resolve a workbook relationship by r:id, requiring an exact kind.

        The relationship's type must equal the exact URI for
        ``expected_kind``; a crafted rels part cannot redirect a worksheet
        lookup to some other relationship type. ``expected_kind`` must be a
        key of :data:`SUPPORTED_WORKBOOK_REL_TYPES`.
        """
        exact_type = SUPPORTED_WORKBOOK_REL_TYPES[expected_kind]
        rel = self.workbook_rels.get(rel_id)
        if rel is None or rel.rel_type != exact_type:
            return None
        return self._resolve_target(rel)

    def content_type_of(self, part_path: str) -> str | None:
        override = self.content_types.get("/" + part_path)
        if override is not None:
            return override
        _, _, extension = part_path.rpartition(".")
        return self.default_types.get(extension)


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


def _parse_content_types(
    root: ET.Element, part: str
) -> tuple[dict[str, str], dict[str, str]] | Refusal:
    if root.tag != _CONTENT_TYPES_ROOT_TAG:
        return _invalid_part(part, "unexpected root element", root_tag=root.tag)
    overrides: dict[str, str] = {}
    defaults: dict[str, str] = {}
    for child in root:
        if child.tag == _CONTENT_TYPES_OVERRIDE_TAG:
            part_name = child.attrib.get("PartName")
            content_type = child.attrib.get("ContentType")
            if not part_name or not content_type:
                return _invalid_part(
                    part, "Override element with missing or empty PartName or ContentType"
                )
            if part_name in overrides:
                return _invalid_part(part, "duplicate Override PartName", part_name=part_name)
            overrides[part_name] = content_type
        elif child.tag == _CONTENT_TYPES_DEFAULT_TAG:
            extension = child.attrib.get("Extension")
            content_type = child.attrib.get("ContentType")
            if not extension or not content_type:
                return _invalid_part(
                    part, "Default element with missing or empty Extension or ContentType"
                )
            if extension in defaults:
                return _invalid_part(part, "duplicate Default Extension", extension=extension)
            defaults[extension] = content_type
    return overrides, defaults


def _parse_rels(root: ET.Element, part: str) -> dict[str, Relationship] | Refusal:
    if root.tag != _RELATIONSHIPS_ROOT_TAG:
        return _invalid_part(part, "unexpected root element", root_tag=root.tag)
    rels: dict[str, Relationship] = {}
    for child in root:
        if child.tag != _RELATIONSHIP_TAG:
            continue
        rel_id = child.attrib.get("Id")
        rel_type = child.attrib.get("Type")
        target = child.attrib.get("Target")
        if not rel_id or not rel_type or not target:
            return _invalid_part(
                part, "Relationship element with missing or empty Id, Type, or Target"
            )
        target_mode = child.attrib.get("TargetMode", TARGET_MODE_INTERNAL)
        if target_mode not in _VALID_TARGET_MODES:
            return _invalid_part(part, "invalid TargetMode", rel_id=rel_id, target_mode=target_mode)
        if rel_id in rels:
            # Two relationships with one ID are ambiguous, and consumers that
            # disagree about which wins can be steered to different parts.
            return _invalid_part(part, "duplicate relationship ID", rel_id=rel_id)
        rels[rel_id] = Relationship(
            rel_id=rel_id,
            rel_type=rel_type,
            target=target,
            target_mode=target_mode,
        )
    return rels


def _unsupported_package_kind(kind: str, content_type: str | None) -> Refusal:
    return Refusal(
        code="unsupported_package_kind",
        message=f"This OOXML package is not a supported spreadsheet workbook (detected: {kind}).",
        details={"kind": kind, "content_type": content_type},
        recovery_options=({"action": "resave_as_transitional_xlsx"},),
    )


def parse_package(archive: WorkbookArchive) -> PackageInfo | Refusal:
    """Resolve and validate the package structure of a ``.xlsx`` archive."""
    if not archive.has_part(CONTENT_TYPES_PART):
        return _missing_part(CONTENT_TYPES_PART, "not present in the archive")
    content_types_root = archive.parse_xml_part(CONTENT_TYPES_PART)
    if isinstance(content_types_root, Refusal):
        return content_types_root
    parsed_types = _parse_content_types(content_types_root, CONTENT_TYPES_PART)
    if isinstance(parsed_types, Refusal):
        return parsed_types
    content_types, default_types = parsed_types

    if not archive.has_part(_PACKAGE_RELS_PART):
        return _missing_part(_PACKAGE_RELS_PART, "not present in the archive")
    top_rels_root = archive.parse_xml_part(_PACKAGE_RELS_PART)
    if isinstance(top_rels_root, Refusal):
        return top_rels_root
    top_rels = _parse_rels(top_rels_root, _PACKAGE_RELS_PART)
    if isinstance(top_rels, Refusal):
        return top_rels

    transitional = [r for r in top_rels.values() if r.rel_type == REL_TYPE_OFFICE_DOCUMENT]
    strict = [r for r in top_rels.values() if r.rel_type == _STRICT_REL_TYPE_OFFICE_DOCUMENT]
    if transitional and strict:
        return _invalid_part(
            _PACKAGE_RELS_PART, "mixed transitional and strict officeDocument relationships"
        )
    if len(transitional) > 1:
        return _invalid_part(
            _PACKAGE_RELS_PART,
            "multiple officeDocument relationships",
            count=len(transitional),
        )
    if strict:
        return _unsupported_package_kind("strict_ooxml_workbook", None)
    if not transitional:
        return _missing_part(
            _PACKAGE_RELS_PART, "no officeDocument relationship points at a workbook"
        )

    office_rel = transitional[0]
    if office_rel.target_mode != TARGET_MODE_INTERNAL:
        return _invalid_part(
            _PACKAGE_RELS_PART, "officeDocument relationship is external", target=office_rel.target
        )

    cleaned_target = validate_opc_target(office_rel.target)
    workbook_path = None if cleaned_target is None else normalize_package_path(cleaned_target)
    if workbook_path is None:
        return Refusal(
            code="unsafe_archive_path",
            message=f"The officeDocument relationship target {office_rel.target!r} is not a "
            "safe package path.",
            details={"kind": "relationship_target", "target": office_rel.target},
            recovery_options=_RECOVER_TRUSTED_SOURCE,
        )

    if not archive.has_part(workbook_path):
        return _missing_part(workbook_path, "the officeDocument relationship points at it")

    workbook_dir, _, workbook_name = workbook_path.rpartition("/")
    rels_prefix = f"{workbook_dir}/_rels" if workbook_dir else "_rels"
    workbook_rels_part = f"{rels_prefix}/{workbook_name}.rels"

    workbook_rels: dict[str, Relationship] = {}
    if archive.has_part(workbook_rels_part):
        workbook_rels_root = archive.parse_xml_part(workbook_rels_part)
        if isinstance(workbook_rels_root, Refusal):
            return workbook_rels_root
        parsed_rels = _parse_rels(workbook_rels_root, workbook_rels_part)
        if isinstance(parsed_rels, Refusal):
            return parsed_rels
        workbook_rels = parsed_rels

    # Worksheet relationships are legitimately plural (one per sheet); these
    # four are singletons, and a second one is ambiguous and steerable.
    for singleton_kind in ("styles", "sharedStrings", "calcChain", "theme"):
        exact_type = SUPPORTED_WORKBOOK_REL_TYPES[singleton_kind]
        matches = sum(1 for rel in workbook_rels.values() if rel.rel_type == exact_type)
        if matches > 1:
            return _invalid_part(
                workbook_rels_part,
                f"duplicate {singleton_kind} relationship",
                kind=singleton_kind,
                count=matches,
            )

    package = PackageInfo(
        workbook_path=workbook_path,
        workbook_dir=workbook_dir,
        content_types=content_types,
        default_types=default_types,
        top_rels=top_rels,
        workbook_rels=workbook_rels,
    )

    workbook_content_type = package.content_type_of(workbook_path)
    if workbook_content_type != SUPPORTED_WORKBOOK_CONTENT_TYPE:
        kind = _KNOWN_UNSUPPORTED_MAIN_CONTENT_TYPES.get(workbook_content_type or "", "unknown")
        return _unsupported_package_kind(kind, workbook_content_type)

    return package
