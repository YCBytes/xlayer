"""Private bounded potential-reference analysis and detached evidence records.

References are not calculated effects, approval, or edit safety. Serialized
evidence may be large; nothing here automatically logs or exports formulas.
"""

from __future__ import annotations

import heapq
import json
import math
import re
from collections import deque
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, fields
from types import MappingProxyType
from typing import Literal, TypeAlias, TypeVar

from xlayer._errors import Refusal
from xlayer._ooxml.formula import (
    FormulaProblem,
    ParsedFormula,
    ReferenceKind,
    ResolvedReference,
    _endpoint,
    _valid_name,
    build_name_index,
    build_reference_context,
    parse_formula,
    resolve_references,
)
from xlayer._ooxml.sheet import Cell, Formula, Worksheet, _column_letter
from xlayer._ooxml.workbook import SheetEntry, WorkbookRegistry

JsonValue: TypeAlias = (
    bool | int | float | str | Sequence["JsonValue"] | Mapping[str, "JsonValue"] | None
)
_T = TypeVar("_T")
_INT_MIN, _INT_MAX = -(2**63), 2**63 - 1
_QUERY = re.compile(r"[A-Za-z]{1,3}[1-9][0-9]{0,6}\Z")
_FINGERPRINT = re.compile(r"sha256:[0-9a-f]{64}\Z")


def _string(value: object, name: str) -> str:
    if not isinstance(value, str):
        raise TypeError(f"{name} must be a string")
    if any(0xD800 <= ord(char) <= 0xDFFF for char in value):
        raise ValueError(f"{name} contains a lone surrogate")
    return value


def _integer(value: object, name: str, minimum: int = 0, maximum: int = _INT_MAX) -> int:
    if type(value) is not int:
        raise TypeError(f"{name} must be an integer, not bool")
    if not minimum <= value <= maximum:
        raise ValueError(f"{name} outside {minimum}..{maximum}")
    return value


def _require(value: object, kind: type[_T], name: str) -> _T:
    if not isinstance(value, kind):
        raise TypeError(f"{name} must be {kind.__name__}")
    return value


def _records(value: object, kind: type[_T], name: str) -> tuple[_T, ...]:
    if not isinstance(value, Sequence) or isinstance(value, (str, bytes, bytearray)):
        raise TypeError(f"{name} must be a sequence")
    return tuple(_require(item, kind, name) for item in value)


def _pair(value: object, name: str, minimum: int = 0) -> tuple[int, int]:
    if not isinstance(value, Sequence) or isinstance(value, (str, bytes)):
        raise TypeError(f"{name} must be an integer pair")
    if len(value) != 2:
        raise ValueError(f"{name} must contain two integers")
    return (_integer(value[0], name, minimum), _integer(value[1], name, minimum))


def _freeze_json(value: object, *, depth: int = 0, _ancestry: set[int] | None = None) -> JsonValue:
    """Copy and validate a bounded JSON domain; shared acyclic containers are legal."""
    if depth > 64:
        raise ValueError("diagnostic nesting exceeds 64")
    if value is None or isinstance(value, bool):
        return value
    if isinstance(value, str):
        return _string(value, "diagnostic string")
    if isinstance(value, int):
        return _integer(value, "diagnostic integer", _INT_MIN)
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError("diagnostic float must be finite")
        return value
    if not isinstance(value, (Mapping, Sequence)) or isinstance(value, (bytes, bytearray)):
        raise TypeError("diagnostic value outside JSON domain")
    ancestry = set() if _ancestry is None else _ancestry
    identity = id(value)
    if identity in ancestry:
        raise ValueError("cyclic diagnostic container")
    ancestry.add(identity)
    try:
        if isinstance(value, Mapping):
            result: dict[str, JsonValue] = {}
            for key, item in value.items():
                result[_string(key, "diagnostic key")] = _freeze_json(
                    item, depth=depth + 1, _ancestry=ancestry
                )
            return MappingProxyType(result)
        return tuple(_freeze_json(item, depth=depth + 1, _ancestry=ancestry) for item in value)
    finally:
        ancestry.remove(identity)


def _thaw_json(value: JsonValue) -> JsonValue:
    if isinstance(value, Mapping):
        return {key: _thaw_json(item) for key, item in value.items()}
    if isinstance(value, Sequence) and not isinstance(value, str):
        return [_thaw_json(item) for item in value]
    return value


def _json_mapping(value: object) -> Mapping[str, JsonValue]:
    if not isinstance(value, Mapping):
        raise TypeError("diagnostic details/option must be a mapping")
    frozen = _freeze_json(value)
    if not isinstance(frozen, Mapping):
        raise TypeError("diagnostic mapping required")
    return frozen


def _canonical_address(address: str) -> str:
    if _QUERY.fullmatch(address) is None or _endpoint(address) is None:
        raise ValueError("address must be an unadorned in-grid ASCII A1 coordinate")
    return address.upper()


@dataclass(frozen=True)
class ImpactLimits:
    max_sheets: int = 256
    max_scanned_cells: int = 250000
    max_defined_names: int = 10000
    max_formula_cells: int = 25000
    max_formula_chars: int = 8192
    max_formula_nesting: int = 64
    max_references: int = 100000
    max_membership_checks: int = 1000000
    max_visited_nodes: int = 10000
    max_evidence_edges: int = 50000

    def __post_init__(self) -> None:
        for field in fields(self):
            value = getattr(self, field.name)
            if type(value) is not int or not 1 <= value <= _INT_MAX:
                raise ValueError(f"{field.name} must be an integer in 1..2**63-1, not bool")


DEFAULT_IMPACT_LIMITS = ImpactLimits()


@dataclass(frozen=True)
class CellRef:
    sheet: str
    address: str

    def __post_init__(self) -> None:
        if not _string(self.sheet, "sheet"):
            raise ValueError("empty sheet identity")
        if _canonical_address(_string(self.address, "address")) != self.address:
            raise ValueError("CellRef address must be canonical uppercase")


@dataclass(frozen=True)
class RangeRef:
    sheet: str
    min_row: int
    min_column: int
    max_row: int
    max_column: int

    def __post_init__(self) -> None:
        if not _string(self.sheet, "sheet"):
            raise ValueError("empty sheet identity")
        for name, ceiling in (
            ("min_row", 1048576),
            ("max_row", 1048576),
            ("min_column", 16384),
            ("max_column", 16384),
        ):
            _integer(getattr(self, name), name, 1, ceiling)
        if self.min_row > self.max_row or self.min_column > self.max_column:
            raise ValueError("range bounds are not normalized")

    def contains(self, cell: CellRef) -> bool:
        endpoint = _endpoint(cell.address)
        return (
            cell.sheet == self.sheet
            and endpoint is not None
            and self.min_row <= endpoint.row <= self.max_row
            and self.min_column <= endpoint.column <= self.max_column
        )


@dataclass(frozen=True)
class NameEvidence:
    name: str
    scope_sheet: str | None
    text: str
    is_hidden: bool
    normalized_reference: RangeRef

    def __post_init__(self) -> None:
        _string(self.name, "name")
        _string(self.text, "text")
        if self.scope_sheet is not None:
            _string(self.scope_sheet, "scope_sheet")
        _require(self.is_hidden, bool, "is_hidden")
        _require(self.normalized_reference, RangeRef, "normalized_reference")


@dataclass(frozen=True)
class ReferenceEvidence:
    formula_cell: CellRef
    formula_text: str
    source_span: tuple[int, int]
    occurrence_index: int
    reference_text: str
    normalized_reference: RangeRef
    reference_kind: ReferenceKind
    basis: tuple[str, ...]
    name_definition: NameEvidence | None
    shared_master: CellRef | None
    shared_offset: tuple[int, int] | None

    def __post_init__(self) -> None:
        _require(self.formula_cell, CellRef, "formula_cell")
        _string(self.formula_text, "formula_text")
        _string(self.reference_text, "reference_text")
        span = _pair(self.source_span, "source_span")
        if not 0 <= span[0] < span[1] <= len(self.formula_text):
            raise ValueError("source_span outside formula")
        if self.formula_text[slice(*span)] != self.reference_text:
            raise ValueError("reference_text is not exact source slice")
        object.__setattr__(self, "source_span", span)
        _integer(self.occurrence_index, "occurrence_index")
        bounds = _require(self.normalized_reference, RangeRef, "normalized_reference")
        if self.reference_kind not in {"cell", "range", "defined_name"}:
            raise ValueError("invalid reference_kind")
        if self.name_definition is not None:
            _require(self.name_definition, NameEvidence, "name_definition")
        if (self.reference_kind == "defined_name") != (self.name_definition is not None):
            raise ValueError("name evidence required iff defined_name")
        if self.name_definition is not None and self.name_definition.normalized_reference != bounds:
            raise ValueError("name bounds do not match reference")
        if (self.shared_master is None) != (self.shared_offset is None):
            raise ValueError("shared master and offset must coexist")
        expected = ["direct_ooxml"]
        if self.shared_master is not None:
            _require(self.shared_master, CellRef, "shared_master")
            offset = _pair(self.shared_offset, "shared_offset", _INT_MIN)
            object.__setattr__(self, "shared_offset", offset)
            expected.append("shared_translation")
        if self.name_definition is not None:
            expected.append("defined_name_resolution")
        if bounds.min_row != bounds.max_row or bounds.min_column != bounds.max_column:
            expected.append("range_membership")
        basis = _records(self.basis, str, "basis")
        if basis != tuple(expected):
            raise ValueError("basis is not ordered exact proof steps")
        object.__setattr__(self, "basis", basis)


@dataclass(frozen=True)
class DependencyEdge:
    precedent: CellRef
    dependent: CellRef
    evidence: ReferenceEvidence

    def __post_init__(self) -> None:
        _require(self.precedent, CellRef, "precedent")
        _require(self.dependent, CellRef, "dependent")
        evidence = _require(self.evidence, ReferenceEvidence, "evidence")
        if evidence.formula_cell != self.dependent or not evidence.normalized_reference.contains(
            self.precedent
        ):
            raise ValueError("edge identity/reference inconsistent")


@dataclass(frozen=True)
class AnalysisIssue:
    code: str
    sheet: str | None
    formula_cell: CellRef | None
    source_span: tuple[int, int] | None
    details: Mapping[str, JsonValue]
    recovery_options: tuple[Mapping[str, JsonValue], ...]

    def __post_init__(self) -> None:
        _string(self.code, "code")
        if self.sheet is not None:
            _string(self.sheet, "sheet")
        if self.formula_cell is not None:
            _require(self.formula_cell, CellRef, "formula_cell")
        if self.source_span is not None:
            span = _pair(self.source_span, "source_span")
            if span[0] > span[1]:
                raise ValueError("invalid issue source_span")
            object.__setattr__(self, "source_span", span)
        object.__setattr__(self, "details", _json_mapping(self.details))
        if not isinstance(self.recovery_options, Sequence):
            raise TypeError("recovery_options must be a sequence")
        options = tuple(_json_mapping(option) for option in self.recovery_options)
        if not options or not all(options):
            raise ValueError("recovery_options must contain nonempty option mappings")
        object.__setattr__(self, "recovery_options", options)


@dataclass(frozen=True)
class InventoryCursor:
    sheet: str
    tab_index: int
    phase: Literal["unread_sheet", "scan_cells", "admit_formula"]
    next_cell: CellRef | None

    def __post_init__(self) -> None:
        _string(self.sheet, "sheet")
        _integer(self.tab_index, "tab_index")
        if self.phase not in {"unread_sheet", "scan_cells", "admit_formula"}:
            raise ValueError("invalid inventory phase")
        if self.next_cell is not None:
            _require(self.next_cell, CellRef, "next_cell")
            if self.next_cell.sheet != self.sheet:
                raise ValueError("inventory cursor sheet mismatch")
        if (self.phase == "unread_sheet") != (self.next_cell is None):
            raise ValueError("next_cell None only for unread_sheet")


@dataclass(frozen=True)
class ReferenceCursor:
    dependent: CellRef
    source_span: tuple[int, int]
    occurrence_index: int
    reference_kind: ReferenceKind

    def __post_init__(self) -> None:
        _require(self.dependent, CellRef, "dependent")
        span = _pair(self.source_span, "source_span")
        if span[0] >= span[1]:
            raise ValueError("empty reference span")
        object.__setattr__(self, "source_span", span)
        _integer(self.occurrence_index, "occurrence_index")
        if self.reference_kind not in {"cell", "range", "defined_name"}:
            raise ValueError("invalid reference kind")


@dataclass(frozen=True)
class TraversalCursor:
    precedent: CellRef
    depth: int
    phase: Literal["queued", "match_candidate", "admit_edge"]
    next_candidate: ReferenceCursor | None

    def __post_init__(self) -> None:
        _require(self.precedent, CellRef, "precedent")
        _integer(self.depth, "depth")
        if self.phase not in {"queued", "match_candidate", "admit_edge"}:
            raise ValueError("invalid traversal phase")
        if self.next_candidate is not None:
            _require(self.next_candidate, ReferenceCursor, "next_candidate")
        if (self.phase == "queued") != (self.next_candidate is None):
            raise ValueError("next_candidate None only for queued")


@dataclass(frozen=True)
class AnalysisWork:
    sheets_attempted: int = 0
    trusted_sheets: int = 0
    scanned_cells: int = 0
    defined_names_admitted: int = 0
    formula_cells_attempted: int = 0
    max_formula_chars_observed: int = 0
    max_formula_nesting_observed: int = 0
    references_admitted: int = 0
    membership_checks: int = 0
    visited_nodes: int = 0
    evidence_edges_admitted: int = 0

    def __post_init__(self) -> None:
        for field in fields(self):
            _integer(getattr(self, field.name), field.name)


@dataclass(frozen=True)
class DependencyImpact:
    schema_version: Literal["1.0"]
    dependency_contract_version: Literal["1.0"]
    source_fingerprint: str
    root: CellRef
    scope: Literal["worksheet_cell_formulas"]
    analysis_status: Literal["complete", "partial", "unknown", "not_evaluated"]
    direct_dependents: tuple[CellRef, ...]
    transitive_dependents: tuple[CellRef, ...]
    known_direct_count: int | None
    known_transitive_count: int | None
    exact_total_count: int | None
    evidence_edges: tuple[DependencyEdge, ...]
    predecessor_edges: Mapping[CellRef, DependencyEdge]
    known_cycles: tuple[tuple[CellRef, ...], ...]
    issues: tuple[AnalysisIssue, ...]
    inventory_frontier: tuple[InventoryCursor, ...]
    traversal_frontier: tuple[TraversalCursor, ...]
    excluded_tabs: tuple[SheetEntry, ...]
    limits: ImpactLimits
    work: AnalysisWork
    alternate_paths_enumerated: Literal[False]

    def __post_init__(self) -> None:
        if (
            _string(self.schema_version, "schema_version") != "1.0"
            or _string(self.dependency_contract_version, "dependency_contract_version") != "1.0"
        ):
            raise ValueError("unsupported dependency schema/contract version")
        if _FINGERPRINT.fullmatch(_string(self.source_fingerprint, "source_fingerprint")) is None:
            raise ValueError("invalid snapshot fingerprint")
        _require(self.root, CellRef, "root")
        _require(self.limits, ImpactLimits, "limits")
        _require(self.work, AnalysisWork, "work")
        if _string(
            self.scope, "scope"
        ) != "worksheet_cell_formulas" or self.analysis_status not in {
            "complete",
            "partial",
            "unknown",
            "not_evaluated",
        }:
            raise ValueError("invalid analysis scope/status")
        if self.alternate_paths_enumerated is not False:
            raise ValueError("alternative paths are not enumerated")
        for name, kind in (
            ("direct_dependents", CellRef),
            ("transitive_dependents", CellRef),
            ("evidence_edges", DependencyEdge),
            ("issues", AnalysisIssue),
            ("inventory_frontier", InventoryCursor),
            ("traversal_frontier", TraversalCursor),
            ("excluded_tabs", SheetEntry),
        ):
            object.__setattr__(self, name, _records(getattr(self, name), kind, name))
        cycles = tuple(_records(cycle, CellRef, "cycle") for cycle in self.known_cycles)
        object.__setattr__(self, "known_cycles", cycles)
        if not isinstance(self.predecessor_edges, Mapping):
            raise TypeError("predecessor_edges must be a mapping")
        predecessors = {
            _require(key, CellRef, "predecessor key"): _require(
                value, DependencyEdge, "predecessor edge"
            )
            for key, value in self.predecessor_edges.items()
        }
        object.__setattr__(self, "predecessor_edges", MappingProxyType(predecessors))
        direct, transitive = set(self.direct_dependents), set(self.transitive_dependents)
        if (
            len(direct) != len(self.direct_dependents)
            or len(transitive) != len(self.transitive_dependents)
            or direct & transitive
            or self.root in direct | transitive
        ):
            raise ValueError("dependent lists must be unique/disjoint and exclude root")
        retained = set(self.evidence_edges)
        if len(retained) != len(self.evidence_edges):
            raise ValueError("duplicate edge occurrence")
        if set(predecessors) != direct | transitive or any(
            edge.dependent != cell or edge not in retained for cell, edge in predecessors.items()
        ):
            raise ValueError("every dependent requires one retained predecessor")
        for name in ("known_direct_count", "known_transitive_count", "exact_total_count"):
            value = getattr(self, name)
            if value is not None:
                _integer(value, name)
        if self.analysis_status in {"unknown", "not_evaluated"}:
            if (
                direct
                or transitive
                or retained
                or any(
                    value is not None
                    for value in (
                        self.known_direct_count,
                        self.known_transitive_count,
                        self.exact_total_count,
                    )
                )
            ):
                raise ValueError("unknown/skipped result cannot expose proof/counts")
        elif self.known_direct_count != len(direct) or self.known_transitive_count != len(
            transitive
        ):
            raise ValueError("known counts must match known lists")
        if self.analysis_status == "complete":
            if (
                self.exact_total_count != len(direct | transitive)
                or self.issues
                or self.inventory_frontier
                or self.traversal_frontier
            ):
                raise ValueError("complete result needs exact counts and no unfinished work/issues")
        elif self.exact_total_count is not None:
            raise ValueError("only complete analysis has an exact total")

    def path_to(self, ref: CellRef) -> tuple[DependencyEdge, ...]:
        _require(ref, CellRef, "ref")
        if ref == self.root:
            return ()
        if ref not in self.predecessor_edges:
            raise ValueError("cell is not a known dependent")
        path: list[DependencyEdge] = []
        current = ref
        for _ in range(len(self.predecessor_edges)):
            edge = self.predecessor_edges.get(current)
            if edge is None:
                break
            path.append(edge)
            current = edge.precedent
            if current == self.root:
                return tuple(reversed(path))
        raise ValueError("inconsistent predecessor path")

    def to_dict(self) -> dict[str, JsonValue]:
        indices = {edge: index for index, edge in enumerate(self.evidence_edges)}
        result = {
            field.name: _to_json(getattr(self, field.name))
            for field in fields(self)
            if field.name != "predecessor_edges"
        }
        result["predecessor_edges"] = [
            {"cell": _to_json(cell), "edge_index": indices[edge]}
            for cell, edge in self.predecessor_edges.items()
        ]
        return result


def _to_json(value: object) -> JsonValue:
    if isinstance(
        value,
        (
            CellRef,
            RangeRef,
            NameEvidence,
            ReferenceEvidence,
            DependencyEdge,
            AnalysisIssue,
            InventoryCursor,
            ReferenceCursor,
            TraversalCursor,
            AnalysisWork,
            ImpactLimits,
            SheetEntry,
        ),
    ):
        return {field.name: _to_json(getattr(value, field.name)) for field in fields(value)}
    if isinstance(value, Mapping):
        return {_string(key, "JSON key"): _to_json(item) for key, item in value.items()}
    if isinstance(value, Sequence) and not isinstance(value, (str, bytes, bytearray)):
        return [_to_json(item) for item in value]
    return _thaw_json(_freeze_json(value))


@dataclass(frozen=True)
class _ReferenceIndex:
    source_fingerprint: str
    sheet_order: Mapping[str, int]
    points: Mapping[CellRef, tuple[ReferenceEvidence, ...]]
    ranges: Mapping[str, tuple[ReferenceEvidence, ...]]
    issues: tuple[AnalysisIssue, ...]
    inventory_frontier: tuple[InventoryCursor, ...]
    excluded_tabs: tuple[SheetEntry, ...]
    work: AnalysisWork


@dataclass(frozen=True)
class _Traversal:
    edges: tuple[DependencyEdge, ...]
    predecessors: Mapping[CellRef, DependencyEdge]
    depths: Mapping[CellRef, int]
    issues: tuple[AnalysisIssue, ...]
    frontier: tuple[TraversalCursor, ...]
    work: AnalysisWork


_RECOVERY: tuple[Mapping[str, JsonValue], ...] = ({"action": "inspect_formula_or_review_in_excel"},)
_BUDGET_RECOVERY: tuple[Mapping[str, JsonValue], ...] = (
    {"action": "review_outside_bounded_analyzer"},
)


def _cell_key(cell: CellRef, order: Mapping[str, int]) -> tuple[int, int, int]:
    endpoint = _endpoint(cell.address)
    if endpoint is None:
        raise ValueError("invalid canonical cell identity")
    return order[cell.sheet], endpoint.row, endpoint.column


def _reference_key(
    evidence: ReferenceEvidence, order: Mapping[str, int]
) -> tuple[tuple[int, int, int], tuple[int, int], int]:
    return _cell_key(evidence.formula_cell, order), evidence.source_span, evidence.occurrence_index


def _area(bounds: RangeRef) -> bool:
    return bounds.min_row != bounds.max_row or bounds.min_column != bounds.max_column


def _point(bounds: RangeRef) -> CellRef:
    return CellRef(bounds.sheet, f"{_column_letter(bounds.min_column)}{bounds.min_row}")


def _issue(
    code: str,
    sheet: str | None,
    cell: CellRef | None,
    span: tuple[int, int] | None,
    details: Mapping[str, JsonValue],
    *,
    recovery: tuple[Mapping[str, JsonValue], ...] = _RECOVERY,
) -> AnalysisIssue:
    return AnalysisIssue(code, sheet, cell, span, details, recovery)


def _evidence(
    resolved: ResolvedReference,
    cell: CellRef,
    text: str,
    master: CellRef | None,
    offset: tuple[int, int] | None,
) -> ReferenceEvidence:
    token = resolved.token
    bounds = RangeRef(
        resolved.sheet, resolved.min_row, resolved.min_column, resolved.max_row, resolved.max_column
    )
    basis = ["direct_ooxml"]
    if master is not None:
        basis.append("shared_translation")
    name = None
    if resolved.definition is not None:
        definition = resolved.definition
        name = NameEvidence(
            definition.name, definition.scope_sheet, definition.text, definition.is_hidden, bounds
        )
        basis.append("defined_name_resolution")
    if _area(bounds):
        basis.append("range_membership")
    return ReferenceEvidence(
        cell,
        text,
        token.source_span,
        token.occurrence_index,
        text[slice(*token.source_span)],
        bounds,
        token.kind,
        tuple(basis),
        name,
        master,
        offset,
    )


def _source(
    cell: Cell,
    sheet: Worksheet,
) -> tuple[str, CellRef | None, tuple[int, int] | None] | None:
    formula = cell.formula
    if formula is None:
        return None
    if formula.kind != "shared":
        return formula.text or "", None, None
    shared = (
        sheet.shared_formulas.get(formula.shared_index)
        if formula.shared_index is not None
        else None
    )
    if shared is None:
        return None
    master_cell = sheet.cells.get(shared.master)
    if (
        master_cell is None
        or master_cell.formula is None
        or master_cell.formula.shared_index != formula.shared_index
        or master_cell.formula.kind != "shared"
        or master_cell.formula.text != shared.text
    ):
        return None
    master = CellRef(sheet.name, shared.master)
    return shared.text, master, (cell.row - master_cell.row, cell.column - master_cell.column)


def _build_index(
    registry: WorkbookRegistry,
    read_sheet: Callable[[str], Worksheet | Refusal],
    source_fingerprint: str,
    limits: ImpactLimits,
) -> _ReferenceIndex:
    order = {entry.name: entry.tab_index for entry in registry.sheets}
    worksheets = tuple(entry for entry in registry.sheets if entry.kind == "worksheet")
    excluded = tuple(entry for entry in registry.sheets if entry.kind != "worksheet")
    points: dict[CellRef, list[ReferenceEvidence]] = {}
    ranges: dict[str, list[ReferenceEvidence]] = {}
    issues: list[AnalysisIssue] = []
    frontier: list[InventoryCursor] = []
    work = {field.name: 0 for field in fields(AnalysisWork)}
    context = build_reference_context(registry, names=None)
    names_checked = False
    names_unavailable = False

    def stop(
        entry_index: int,
        phase: Literal["unread_sheet", "scan_cells", "admit_formula"],
        cell: CellRef | None,
        budget: str,
        consumed: int,
        limit: int,
    ) -> None:
        entry = worksheets[entry_index]
        frontier.append(InventoryCursor(entry.name, entry.tab_index, phase, cell))
        frontier.extend(
            InventoryCursor(e.name, e.tab_index, "unread_sheet", None)
            for e in worksheets[entry_index + 1 :]
        )
        issues.append(
            _issue(
                "inventory_budget_exceeded",
                entry.name,
                cell,
                None,
                {"budget": budget, "consumed": consumed, "limit": limit},
                recovery=_BUDGET_RECOVERY,
            )
        )

    aborted = False
    for entry_index, entry in enumerate(worksheets):
        if work["sheets_attempted"] >= limits.max_sheets:
            stop(
                entry_index,
                "unread_sheet",
                None,
                "max_sheets",
                work["sheets_attempted"],
                limits.max_sheets,
            )
            break
        work["sheets_attempted"] += 1
        sheet = read_sheet(entry.name)
        if isinstance(sheet, Refusal):
            recovery = tuple(_json_mapping(option) for option in sheet.recovery_options)
            issues.append(
                _issue(
                    "worksheet_unreadable",
                    entry.name,
                    None,
                    None,
                    {"refusal": _freeze_json(sheet.to_dict())},
                    recovery=recovery,
                )
            )
            continue
        if not isinstance(sheet, Worksheet):
            raise TypeError("read_sheet must return a Worksheet or Refusal")
        if sheet.name != entry.name or sheet.part_path != entry.part_path:
            issues.append(
                _issue(
                    "dependency_invariant_failure",
                    entry.name,
                    None,
                    None,
                    {"reason": "worksheet source identity mismatch"},
                )
            )
            break
        work["trusted_sheets"] += 1
        previous = (0, 0)
        for address, cell in sheet.cells.items():
            if work["scanned_cells"] >= limits.max_scanned_cells:
                stop(
                    entry_index,
                    "scan_cells",
                    CellRef(entry.name, address),
                    "max_scanned_cells",
                    work["scanned_cells"],
                    limits.max_scanned_cells,
                )
                aborted = True
                break
            endpoint = _endpoint(address)
            supplied_cell: object = cell
            if (
                not isinstance(supplied_cell, Cell)
                or endpoint is None
                or address != address.upper()
                or address != cell.address
                or (endpoint.row, endpoint.column) != (cell.row, cell.column)
                or (cell.row, cell.column) <= previous
            ):
                issues.append(
                    _issue(
                        "dependency_invariant_failure",
                        entry.name,
                        None,
                        None,
                        {"reason": "cell identity/order mismatch"},
                    )
                )
                aborted = True
                break
            previous = cell.row, cell.column
            ref = CellRef(entry.name, address)
            work["scanned_cells"] += 1
            formula = cell.formula
            if formula is None:
                continue
            if not isinstance(formula, Formula):
                raise TypeError("cell formula must be Formula")
            if work["formula_cells_attempted"] >= limits.max_formula_cells:
                stop(
                    entry_index,
                    "admit_formula",
                    ref,
                    "max_formula_cells",
                    work["formula_cells_attempted"],
                    limits.max_formula_cells,
                )
                aborted = True
                break
            work["formula_cells_attempted"] += 1
            source = _source(cell, sheet)
            if source is None:
                issues.append(
                    _issue(
                        "dependency_invariant_failure",
                        entry.name,
                        ref,
                        None,
                        {"reason": "inconsistent shared master"},
                    )
                )
                aborted = True
                break
            text, master, offset = source
            work["max_formula_chars_observed"] = max(work["max_formula_chars_observed"], len(text))
            if formula.kind not in {"normal", "shared"}:
                issues.append(
                    _issue(
                        "unsupported_formula_kind",
                        entry.name,
                        ref,
                        None,
                        {"kind": formula.kind, "formula_text": text, "ref": formula.ref},
                    )
                )
                continue
            parsed = parse_formula(
                text, max_chars=limits.max_formula_chars, max_nesting=limits.max_formula_nesting
            )
            if isinstance(parsed, ParsedFormula):
                work["max_formula_nesting_observed"] = max(
                    work["max_formula_nesting_observed"], parsed.max_nesting
                )
                uses_names = any(token.kind == "defined_name" for token in parsed.references)
                valid_names = all(
                    token.name is not None and _valid_name(token.name)
                    for token in parsed.references
                    if token.kind == "defined_name"
                )
                if uses_names and valid_names:
                    if not names_checked:
                        names_checked = True
                        if len(registry.defined_names) > limits.max_defined_names:
                            names_unavailable = True
                        else:
                            context = build_reference_context(
                                registry, names=build_name_index(registry.defined_names)
                            )
                            work["defined_names_admitted"] = len(registry.defined_names)
                    if names_unavailable:
                        token = next(
                            token for token in parsed.references if token.kind == "defined_name"
                        )
                        issues.append(
                            _issue(
                                "name_inventory_limit_exceeded",
                                entry.name,
                                ref,
                                token.source_span,
                                {
                                    "limit": limits.max_defined_names,
                                    "inventory_count": len(registry.defined_names),
                                },
                                recovery=_BUDGET_RECOVERY,
                            )
                        )
                        continue
                resolved = resolve_references(
                    parsed,
                    context,
                    formula_sheet=entry.name,
                    row_offset=0 if offset is None else offset[0],
                    column_offset=0 if offset is None else offset[1],
                    max_chars=limits.max_formula_chars,
                )
            else:
                resolved = parsed
            if isinstance(resolved, FormulaProblem):
                information = dict(resolved.information)
                observed_nesting = information.get("observed_nesting", 0)
                if isinstance(observed_nesting, int):
                    work["max_formula_nesting_observed"] = max(
                        work["max_formula_nesting_observed"], observed_nesting
                    )
                definition_nesting = information.get("observed_definition_nesting", 0)
                if isinstance(definition_nesting, int):
                    work["max_formula_nesting_observed"] = max(
                        work["max_formula_nesting_observed"], definition_nesting
                    )
                observed_chars = information.get("observed_definition_chars", 0)
                if isinstance(observed_chars, int):
                    work["max_formula_chars_observed"] = max(
                        work["max_formula_chars_observed"], observed_chars
                    )
                issues.append(
                    _issue(
                        resolved.code,
                        entry.name,
                        ref,
                        resolved.source_span,
                        {
                            "reason": resolved.reason,
                            "formula_text": text,
                            "information": information,
                        },
                    )
                )
                continue
            for item in resolved:
                if item.definition is not None:
                    work["max_formula_chars_observed"] = max(
                        work["max_formula_chars_observed"], len(item.definition.text)
                    )
            if work["references_admitted"] + len(resolved) > limits.max_references:
                stop(
                    entry_index,
                    "admit_formula",
                    ref,
                    "max_references",
                    work["references_admitted"],
                    limits.max_references,
                )
                aborted = True
                break
            work["references_admitted"] += len(resolved)
            for item in resolved:
                evidence = _evidence(item, ref, text, master, offset)
                bounds = evidence.normalized_reference
                if _area(bounds):
                    ranges.setdefault(bounds.sheet, []).append(evidence)
                else:
                    points.setdefault(_point(bounds), []).append(evidence)
        if aborted:
            break
    return _ReferenceIndex(
        source_fingerprint,
        MappingProxyType(order),
        MappingProxyType(
            {
                point: tuple(sorted(bucket, key=lambda e: _reference_key(e, order)))
                for point, bucket in points.items()
            }
        ),
        MappingProxyType(
            {
                sheet: tuple(sorted(bucket, key=lambda e: _reference_key(e, order)))
                for sheet, bucket in ranges.items()
            }
        ),
        tuple(issues),
        tuple(frontier),
        excluded,
        AnalysisWork(**work),
    )


def _index_is_consistent(
    index: _ReferenceIndex,
    registry: WorkbookRegistry,
    source_fingerprint: str,
) -> bool:
    expected = {sheet.name: sheet.tab_index for sheet in registry.sheets}
    worksheets = {sheet.name for sheet in registry.sheets if sheet.kind == "worksheet"}
    if (
        index.source_fingerprint != source_fingerprint
        or dict(index.sheet_order) != expected
        or len(expected) != len(registry.sheets)
        or any(issue.code == "dependency_invariant_failure" for issue in index.issues)
    ):
        return False
    count = 0
    occurrences: set[tuple[CellRef, int]] = set()
    for point, bucket in index.points.items():
        if point.sheet not in worksheets:
            return False
        for evidence in bucket:
            identity = evidence.formula_cell, evidence.occurrence_index
            if identity in occurrences:
                return False
            occurrences.add(identity)
            if (
                evidence.formula_cell.sheet not in worksheets
                or _area(evidence.normalized_reference)
                or _point(evidence.normalized_reference) != point
            ):
                return False
            count += 1
    for sheet, bucket in index.ranges.items():
        if sheet not in worksheets:
            return False
        for evidence in bucket:
            identity = evidence.formula_cell, evidence.occurrence_index
            if identity in occurrences:
                return False
            occurrences.add(identity)
            if (
                evidence.formula_cell.sheet not in worksheets
                or not _area(evidence.normalized_reference)
                or evidence.normalized_reference.sheet != sheet
            ):
                return False
            count += 1
    return count == index.work.references_admitted


def _traversal_is_consistent(
    traversed: _Traversal,
    index: _ReferenceIndex,
    root: CellRef,
    limits: ImpactLimits,
) -> bool:
    """Check explicit proof structure before SCCs/classification, not catch bugs."""
    depths = traversed.depths
    retained = set(traversed.edges)
    if (
        depths.get(root) != 0
        or len(retained) != len(traversed.edges)
        or set(traversed.predecessors) != set(depths) - {root}
        or traversed.work.visited_nodes != len(depths)
        or traversed.work.evidence_edges_admitted != len(retained)
        or len(depths) > limits.max_visited_nodes
        or len(retained) > limits.max_evidence_edges
        or traversed.work.membership_checks > limits.max_membership_checks
    ):
        return False
    for cell, depth in depths.items():
        if cell.sheet not in index.sheet_order or type(depth) is not int:
            return False
        if cell != root and depth < 1:
            return False
    admitted = {e for bucket in index.points.values() for e in bucket}
    admitted.update(e for bucket in index.ranges.values() for e in bucket)
    for edge in traversed.edges:
        if (
            edge.precedent not in depths
            or edge.dependent not in depths
            or edge.evidence not in admitted
        ):
            return False
        if depths[edge.dependent] > depths[edge.precedent] + 1:
            return False
    for cell, edge in traversed.predecessors.items():
        if edge not in retained or edge.dependent != cell:
            return False
        if depths[cell] != depths[edge.precedent] + 1:
            return False
    for field in fields(index.work):
        if field.name not in {
            "visited_nodes",
            "evidence_edges_admitted",
            "membership_checks",
        } and getattr(traversed.work, field.name) != getattr(index.work, field.name):
            return False
    return all(
        cursor.precedent in depths and cursor.depth == depths[cursor.precedent]
        for cursor in traversed.frontier
    )


@dataclass
class _TraversalBudget:
    checks: int = 0
    nodes: int = 0
    edges: int = 0


def _traverse(
    index: _ReferenceIndex,
    root: CellRef,
    limits: ImpactLimits,
    *,
    shared_budget: _TraversalBudget | None = None,
) -> _Traversal:
    if shared_budget is not None:
        if shared_budget.nodes >= limits.max_visited_nodes:
            raise ValueError("batch must check root admission before starting traversal")
        shared_budget.nodes += 1
    queue = deque([root])
    depths = {root: 0}
    predecessors: dict[CellRef, DependencyEdge] = {}
    edges: list[DependencyEdge] = []
    issues: list[AnalysisIssue] = []
    frontier: list[TraversalCursor] = []
    work = {field.name: getattr(index.work, field.name) for field in fields(index.work)}
    halted = False
    while queue:
        current = queue.popleft()
        candidates = heapq.merge(
            index.points.get(current, ()),
            index.ranges.get(current.sheet, ()),
            key=lambda e: _reference_key(e, index.sheet_order),
        )
        for candidate in candidates:
            phase: Literal["match_candidate", "admit_edge"] = "admit_edge"
            budget: str | None = None
            if _area(candidate.normalized_reference):
                checks = (
                    work["membership_checks"] if shared_budget is None else shared_budget.checks
                )
                if checks >= limits.max_membership_checks:
                    phase, budget = "match_candidate", "max_membership_checks"
                else:
                    work["membership_checks"] += 1
                    if shared_budget is not None:
                        shared_budget.checks += 1
                    if not candidate.normalized_reference.contains(current):
                        continue
            if budget is None:
                edge_count = len(edges) if shared_budget is None else shared_budget.edges
                node_count = len(depths) if shared_budget is None else shared_budget.nodes
                if edge_count >= limits.max_evidence_edges:
                    budget = "max_evidence_edges"
                elif (
                    candidate.formula_cell not in depths and node_count >= limits.max_visited_nodes
                ):
                    budget = "max_visited_nodes"
            if budget is not None:
                cursor = ReferenceCursor(
                    candidate.formula_cell,
                    candidate.source_span,
                    candidate.occurrence_index,
                    candidate.reference_kind,
                )
                frontier.append(TraversalCursor(current, depths[current], phase, cursor))
                frontier.extend(
                    TraversalCursor(cell, depths[cell], "queued", None) for cell in queue
                )
                consumed = {
                    "max_membership_checks": work["membership_checks"]
                    if shared_budget is None
                    else shared_budget.checks,
                    "max_evidence_edges": len(edges)
                    if shared_budget is None
                    else shared_budget.edges,
                    "max_visited_nodes": len(depths)
                    if shared_budget is None
                    else shared_budget.nodes,
                }[budget]
                issues.append(
                    _issue(
                        "traversal_budget_exceeded",
                        candidate.formula_cell.sheet,
                        candidate.formula_cell,
                        candidate.source_span,
                        {"budget": budget, "consumed": consumed, "limit": getattr(limits, budget)},
                        recovery=_BUDGET_RECOVERY,
                    )
                )
                halted = True
                break
            edge = DependencyEdge(current, candidate.formula_cell, candidate)
            edges.append(edge)
            if shared_budget is not None:
                shared_budget.edges += 1
            if edge.dependent not in depths:
                if shared_budget is not None:
                    shared_budget.nodes += 1
                depths[edge.dependent] = depths[current] + 1
                predecessors[edge.dependent] = edge
                queue.append(edge.dependent)
        if halted:
            break
    work["visited_nodes"], work["evidence_edges_admitted"] = len(depths), len(edges)
    return _Traversal(
        tuple(edges),
        MappingProxyType(predecessors),
        MappingProxyType(depths),
        tuple(issues),
        tuple(frontier),
        AnalysisWork(**work),
    )


def _observed_cycles(
    nodes: tuple[CellRef, ...],
    edges: tuple[DependencyEdge, ...],
    sheet_order: Mapping[str, int],
) -> tuple[tuple[CellRef, ...], ...]:
    """Iterative Kosaraju on retained proof; does not claim unknown full-graph SCCs."""
    forward: dict[CellRef, set[CellRef]] = {node: set() for node in nodes}
    reverse: dict[CellRef, set[CellRef]] = {node: set() for node in nodes}
    for edge in edges:
        forward[edge.precedent].add(edge.dependent)
        reverse[edge.dependent].add(edge.precedent)
    seen: set[CellRef] = set()
    finished: list[CellRef] = []
    for node in sorted(nodes, key=lambda cell: _cell_key(cell, sheet_order)):
        if node in seen:
            continue
        stack = [(node, False)]
        while stack:
            current, exiting = stack.pop()
            if exiting:
                finished.append(current)
            elif current not in seen:
                seen.add(current)
                stack.append((current, True))
                stack.extend(
                    (child, False)
                    for child in sorted(
                        forward[current],
                        key=lambda cell: _cell_key(cell, sheet_order),
                        reverse=True,
                    )
                    if child not in seen
                )
    assigned: set[CellRef] = set()
    cycles: list[tuple[CellRef, ...]] = []
    for node in reversed(finished):
        if node in assigned:
            continue
        component: list[CellRef] = []
        pending = [node]
        assigned.add(node)
        while pending:
            current = pending.pop()
            component.append(current)
            for parent in reverse[current]:
                if parent not in assigned:
                    assigned.add(parent)
                    pending.append(parent)
        if len(component) > 1 or node in forward[node]:
            cycles.append(tuple(sorted(component, key=lambda cell: _cell_key(cell, sheet_order))))
    return tuple(sorted(cycles, key=lambda cycle: _cell_key(cycle[0], sheet_order)))


def _issue_key(
    issue: AnalysisIssue,
    order: Mapping[str, int],
) -> tuple[int, tuple[int, int, int], tuple[int, int], str, str]:
    return (
        order.get(issue.sheet, -1) if issue.sheet is not None else -1,
        (-1, -1, -1) if issue.formula_cell is None else _cell_key(issue.formula_cell, order),
        (-1, -1) if issue.source_span is None else issue.source_span,
        issue.code,
        json.dumps(
            _to_json(issue.details),
            sort_keys=True,
            ensure_ascii=False,
            separators=(",", ":"),
            allow_nan=False,
        ),
    )


def _unknown_impact(
    root: CellRef,
    source_fingerprint: str,
    limits: ImpactLimits,
    index: _ReferenceIndex,
    issues: Sequence[AnalysisIssue],
    *,
    work: AnalysisWork | None = None,
    traversal_frontier: tuple[TraversalCursor, ...] = (),
) -> DependencyImpact:
    return DependencyImpact(
        "1.0",
        "1.0",
        source_fingerprint,
        root,
        "worksheet_cell_formulas",
        "unknown",
        (),
        (),
        None,
        None,
        None,
        (),
        {},
        (),
        tuple(sorted(issues, key=lambda issue: _issue_key(issue, index.sheet_order))),
        index.inventory_frontier,
        traversal_frontier,
        index.excluded_tabs,
        limits,
        index.work if work is None else work,
        False,
    )


def analyse_dependencies(
    registry: WorkbookRegistry,
    read_sheet: Callable[[str], Worksheet | Refusal],
    root: CellRef,
    source_fingerprint: str,
    limits: ImpactLimits,
) -> DependencyImpact:
    """Build an ephemeral workbook inventory and bounded reverse closure."""
    index = _build_index(registry, read_sheet, source_fingerprint, limits)
    return _impact_from_index(registry, index, root, source_fingerprint, limits)


def _impact_from_index(
    registry: WorkbookRegistry,
    index: _ReferenceIndex,
    root: CellRef,
    source_fingerprint: str,
    limits: ImpactLimits,
    *,
    shared_budget: _TraversalBudget | None = None,
) -> DependencyImpact:
    """Compose evidence from one inventory; the standalone contract is unchanged."""
    issues = list(index.issues)
    if root.sheet not in index.sheet_order or not _index_is_consistent(
        index, registry, source_fingerprint
    ):
        if not any(issue.code == "dependency_invariant_failure" for issue in issues):
            issues.append(
                _issue(
                    "dependency_invariant_failure",
                    None,
                    None,
                    None,
                    {"reason": "inconsistent index identity/reference inventory"},
                )
            )
        trusted = False
    else:
        trusted = index.work.trusted_sheets > 0
    if not trusted:
        return _unknown_impact(root, source_fingerprint, limits, index, issues)
    traversed = (
        _traverse(index, root, limits)
        if shared_budget is None
        else _traverse(index, root, limits, shared_budget=shared_budget)
    )
    issues.extend(traversed.issues)
    if not _traversal_is_consistent(traversed, index, root, limits):
        issues.append(
            _issue(
                "dependency_invariant_failure",
                None,
                None,
                None,
                {"reason": "inconsistent traversal nodes/edges/depths"},
            )
        )
        return _unknown_impact(
            root,
            source_fingerprint,
            limits,
            index,
            issues,
            work=traversed.work,
            traversal_frontier=traversed.frontier,
        )
    status: Literal["complete", "partial"] = (
        "partial" if issues or index.inventory_frontier or traversed.frontier else "complete"
    )
    known = [cell for cell in traversed.depths if cell != root]
    ordered = sorted(known, key=lambda cell: _cell_key(cell, index.sheet_order))
    direct = tuple(cell for cell in ordered if traversed.depths[cell] == 1)
    transitive = tuple(cell for cell in ordered if traversed.depths[cell] >= 2)
    edges = tuple(
        sorted(
            traversed.edges,
            key=lambda edge: (
                _cell_key(edge.precedent, index.sheet_order),
                _cell_key(edge.dependent, index.sheet_order),
                edge.evidence.source_span,
                edge.evidence.occurrence_index,
            ),
        )
    )
    predecessors = {cell: traversed.predecessors[cell] for cell in ordered}
    cycles = _observed_cycles(tuple(traversed.depths), edges, index.sheet_order)
    return DependencyImpact(
        "1.0",
        "1.0",
        source_fingerprint,
        root,
        "worksheet_cell_formulas",
        status,
        direct,
        transitive,
        len(direct),
        len(transitive),
        len(known) if status == "complete" else None,
        edges,
        predecessors,
        cycles,
        tuple(sorted(issues, key=lambda issue: _issue_key(issue, index.sheet_order))),
        index.inventory_frontier,
        traversed.frontier,
        index.excluded_tabs,
        limits,
        traversed.work,
        False,
    )
