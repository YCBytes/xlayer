"""Detached, bounded observations. No calculation, logging, authority or graph export."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field, fields, replace
from typing import cast

from xlayer._batch_impact import BatchImpact, RootImpact
from xlayer._canonical import canonical_json, frozen_mapping, thaw_json
from xlayer._dependencies import AnalysisWork, DependencyImpact, _cell_key, _to_json
from xlayer._errors import Refusal
from xlayer._impact_summary import summarize_batch
from xlayer._ooxml.sheet import Worksheet

MAX_CELLS = 128
MAX_RESPONSE_BYTES = 1_048_576
MAX_SUMMARY_ISSUE_BYTES = 1_048_576


@dataclass(frozen=True)
class ReadLimits:
    """Positive response ceilings, not model-token or worksheet-parser limits."""

    max_cells: int = MAX_CELLS
    max_response_bytes: int = MAX_RESPONSE_BYTES

    def __post_init__(self) -> None:
        for name, maximum in (("max_cells", MAX_CELLS), ("max_response_bytes", MAX_RESPONSE_BYTES)):
            value = getattr(self, name)
            if type(value) is not int or not 1 <= value <= maximum:
                raise ValueError(f"{name} must be an exact integer in 1..{maximum}")


DEFAULT_READ_LIMITS = ReadLimits()


@dataclass(frozen=True)
class Inspection:
    """Library-produced immutable evidence from one loaded snapshot.

    Serialization returns detached JSON-friendly objects or canonical bytes.
    There is no approval identity or freshness claim associated with a read.
    """

    evidence: Mapping[str, object]
    _encoded: bytes = field(init=False, repr=False, compare=False)

    def __post_init__(self) -> None:
        evidence = frozen_mapping(self.evidence)
        object.__setattr__(self, "evidence", evidence)
        object.__setattr__(self, "_encoded", canonical_json(evidence))

    def to_dict(self) -> dict[str, object]:
        return cast("dict[str, object]", thaw_json(self.evidence))

    def canonical_json(self) -> bytes:
        return self._encoded


def check_limits(value: object, default: object, name: str) -> None:
    """Only public boundaries cap existing limit records; private ceilings stay usable."""
    if type(value) is not type(default):
        raise TypeError(f"{name} must be an exact {type(default).__name__}")
    for spec in fields(default):  # type: ignore[arg-type]
        actual, maximum = getattr(value, spec.name), getattr(default, spec.name)
        if type(actual) is not int or not 1 <= actual <= maximum:
            raise ValueError(f"{name}.{spec.name} must be an exact integer in 1..{maximum}")


def minimum_json_bytes(value: object, ceiling: int) -> int:
    """Lower bound only: repeated occurrences count, escaping is left to the encoder.

    Stop once an overflow is proven, before copying or encoding large text.
    The canonical domain itself is still validated by the existing encoder.
    """
    total = 0

    def visit(item: object, depth: int) -> None:
        nonlocal total
        if total > ceiling:
            return
        if depth > 64:
            raise ValueError("JSON nesting exceeds 64")
        if item is None:
            total += 4
        elif type(item) is bool:
            total += 4 if item else 5
        elif type(item) is str:
            total += len(item) + 2
        elif type(item) in {int, float}:
            total += len(str(item))
        elif isinstance(item, Mapping):
            total += 2 + max(0, len(item) - 1) + len(item)
            for key, child in item.items():
                visit(key, depth + 1)
                visit(child, depth + 1)
                if total > ceiling:
                    break
        elif isinstance(item, Sequence) and not isinstance(item, (str, bytes, bytearray)):
            total += 2 + max(0, len(item) - 1)
            for child in item:
                visit(child, depth + 1)
                if total > ceiling:
                    break
        else:
            raise TypeError("value outside the canonical JSON domain")

    visit(value, 0)
    return total


def limit_refusal(
    operation: str, fingerprint: str, budget: str, limit: int, **observed: object
) -> Refusal:
    return Refusal(
        code="inspection_limit_exceeded",
        message="The complete inspection cannot fit the supported response bounds.",
        operation=operation,
        details={"source_fingerprint": fingerprint, "budget": budget, "limit": limit, **observed},
        recovery_options=({"action": "request_fewer_cells_or_narrower_inspection"},),
    )


def make_inspection(
    kind: str, fingerprint: str, data: Mapping[str, object], limits: ReadLimits, operation: str
) -> Inspection | Refusal:
    evidence = {
        "schema_version": "1.0",
        "inspection_contract_version": "1.0",
        "kind": kind,
        "source_fingerprint": fingerprint,
        "data": data,
    }
    minimum = minimum_json_bytes(evidence, limits.max_response_bytes)
    if minimum > limits.max_response_bytes:
        return limit_refusal(
            operation,
            fingerprint,
            "max_response_bytes",
            limits.max_response_bytes,
            observed_minimum_bytes=minimum,
        )
    result = Inspection(evidence)
    size = len(result.canonical_json())
    if size > limits.max_response_bytes:
        return limit_refusal(
            operation,
            fingerprint,
            "max_response_bytes",
            limits.max_response_bytes,
            observed_bytes=size,
        )
    return result


def observe_cell(sheet: Worksheet, address: str) -> dict[str, object]:
    cell = sheet.cells.get(address)
    merge = sheet.merge_at(address)
    formula: dict[str, object] | None = None
    cell_format: dict[str, object] | None = None
    if cell is not None:
        style = cell.style
        cell_format = {
            "style_index": cell.style_index,
            "format_id": style.format_id,
            "format_code": style.format_code,
            "temporal_kind": style.temporal_kind,
            "basis": "cell_xf_number_format",
        }
        if cell.formula is not None:
            stored = cell.formula
            master = (
                sheet.shared_formulas.get(stored.shared_index)
                if stored.shared_index is not None
                else None
            )
            formula = {
                "text": stored.text,
                "kind": stored.kind,
                "shared_index": stored.shared_index,
                "ref": stored.ref,
                "calculate_on_next_recalc": stored.calculate_on_next_recalc,
                "shared_master": None
                if master is None
                else {
                    "address": master.master,
                    "ref": master.ref,
                    "text": master.text,
                },
            }
    has_formula = cell is not None and cell.formula is not None
    return {
        "address": address,
        "presence": "missing" if cell is None else "present",
        "stored_type": None if cell is None else cell.stored_type,
        "value_kind": None if cell is None else cell.value_kind,
        "raw": None if cell is None else cell.raw,
        "stored_value": None if cell is None else cell.stored_value,
        "value_source": "not_present"
        if cell is None
        else "saved_formula_cache"
        if has_formula
        else "stored_literal",
        "formula_result_status": "not_verified" if has_formula else None,
        "cell_format": cell_format,
        "formula": formula,
        "merge": None
        if merge is None
        else {
            "ref": merge.ref,
            "root": merge.first_cell,
            "role": "root" if address == merge.first_cell else "child",
        },
    }


def single_impact_summary(
    impact: DependencyImpact, sheet_order: Mapping[str, int]
) -> Mapping[str, object] | Refusal:
    """Reuse retained one-root proof, guarding diagnostics before summary hashing."""
    for issue in impact.issues:
        identity = _to_json(issue)
        minimum = minimum_json_bytes(identity, MAX_SUMMARY_ISSUE_BYTES)
        if minimum > MAX_SUMMARY_ISSUE_BYTES:
            return limit_refusal(
                "dependency_impact",
                impact.source_fingerprint,
                "max_summary_issue_bytes",
                MAX_SUMMARY_ISSUE_BYTES,
                observed_minimum_bytes=minimum,
            )
        size = len(canonical_json(identity))
        if size > MAX_SUMMARY_ISSUE_BYTES:
            return limit_refusal(
                "dependency_impact",
                impact.source_fingerprint,
                "max_summary_issue_bytes",
                MAX_SUMMARY_ISSUE_BYTES,
                observed_bytes=size,
            )
    work = impact.work
    traversal = AnalysisWork(
        membership_checks=work.membership_checks,
        visited_nodes=work.visited_nodes,
        evidence_edges_admitted=work.evidence_edges_admitted,
    )
    inventory = replace(work, membership_checks=0, visited_nodes=0, evidence_edges_admitted=0)
    known = set(impact.direct_dependents) | set(impact.transitive_dependents)
    batch = BatchImpact(
        (RootImpact(impact.root, "started", replace(impact, work=traversal), None, 0),),
        inventory,
        frozen_mapping(
            {
                "membership_checks": work.membership_checks,
                "visited_nodes": work.visited_nodes,
                "evidence_edges_admitted": work.evidence_edges_admitted,
            }
        ),
        tuple(sorted(known, key=lambda ref: _cell_key(ref, sheet_order))),
        impact.transitive_dependents,
    )
    return summarize_batch(
        batch, source_fingerprint=impact.source_fingerprint, sheet_order=sheet_order
    )
