"""Ephemeral batch evidence with one inventory and shared traversal ceilings."""

from __future__ import annotations

from collections.abc import Callable, Mapping
from dataclasses import dataclass, replace

from xlayer._canonical import frozen_mapping
from xlayer._dependencies import (
    AnalysisWork,
    CellRef,
    DependencyImpact,
    ImpactLimits,
    _build_index,
    _cell_key,
    _impact_from_index,
    _to_json,
    _TraversalBudget,
)
from xlayer._errors import Refusal
from xlayer._ooxml.sheet import Worksheet
from xlayer._ooxml.workbook import WorkbookRegistry


@dataclass(frozen=True)
class RootImpact:
    root: CellRef
    state: str
    impact: DependencyImpact | None
    blocked_budget: str | None
    root_cursor: int

    def to_dict(self) -> dict[str, object]:
        return {
            "root": {"sheet": self.root.sheet, "address": self.root.address},
            "state": self.state,
            "impact": self.impact.to_dict() if self.impact else None,
            "blocked_budget": self.blocked_budget,
            "root_cursor": self.root_cursor,
        }


@dataclass(frozen=True)
class BatchImpact:
    roots: tuple[RootImpact, ...]
    inventory_work: AnalysisWork
    work: Mapping[str, object]
    known_union: tuple[CellRef, ...]
    transitive_union: tuple[CellRef, ...]

    def to_dict(self) -> dict[str, object]:
        return {
            "roots": [root.to_dict() for root in self.roots],
            "inventory_work": _to_json(self.inventory_work),
            "traversal_work": self.work,
            "known_union": [
                {"sheet": ref.sheet, "address": ref.address} for ref in self.known_union
            ],
            "known_transitive_union": [
                {"sheet": ref.sheet, "address": ref.address} for ref in self.transitive_union
            ],
            "scope": "worksheet_cell_formulas",
            "numerical_effects_evaluated": False,
        }


def analyse_batch(
    registry: WorkbookRegistry,
    read_sheet: Callable[[str], Worksheet | Refusal],
    roots: tuple[CellRef, ...],
    fingerprint: str,
    limits: ImpactLimits,
) -> BatchImpact:
    index = _build_index(registry, read_sheet, fingerprint, limits)
    budget = _TraversalBudget()
    results: list[RootImpact] = []
    union: set[CellRef] = set()
    transitive: set[CellRef] = set()
    for cursor, root in enumerate(roots):
        if budget.nodes >= limits.max_visited_nodes:
            results.append(RootImpact(root, "not_started", None, "max_visited_nodes", cursor))
            continue
        previous_nodes = budget.nodes
        result = _impact_from_index(
            registry, index, root, fingerprint, limits, shared_budget=budget
        )
        # Inventory is represented once. A root admitted before untrustworthy
        # inventory short-circuits traversal still consumes the global ceiling.
        root_only = budget.nodes == previous_nodes
        if root_only:
            budget.nodes += 1
        local = AnalysisWork(
            membership_checks=result.work.membership_checks,
            visited_nodes=1 if root_only else result.work.visited_nodes,
            evidence_edges_admitted=result.work.evidence_edges_admitted,
        )
        result = replace(result, work=local)
        results.append(RootImpact(root, "started", result, None, cursor))
        union.update(result.direct_dependents)
        union.update(result.transitive_dependents)
        transitive.update(result.transitive_dependents)

    def order(ref: CellRef) -> tuple[int, int, int]:
        return _cell_key(ref, index.sheet_order)

    return BatchImpact(
        tuple(results),
        index.work,
        frozen_mapping(
            {
                "membership_checks": budget.checks,
                "visited_nodes": budget.nodes,
                "evidence_edges_admitted": budget.edges,
            }
        ),
        tuple(sorted(union, key=order)),
        tuple(sorted(transitive, key=order)),
    )
