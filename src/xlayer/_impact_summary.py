"""Bounded presentation of retained proof; never analysis, calculation or authority."""

from __future__ import annotations

import re
from collections import Counter
from collections.abc import Mapping
from typing import Literal

from xlayer._batch_impact import BatchImpact, RootImpact
from xlayer._canonical import digest, fingerprint, frozen_mapping
from xlayer._dependencies import (
    DEPENDENCY_CONTRACT_VERSION,
    AnalysisIssue,
    CellRef,
    DependencyImpact,
    _cell_key,
    _to_json,
)

IMPACT_SUMMARY_VERSION: Literal["1.0"] = "1.0"
_SAMPLE_LIMITS = frozen_mapping(
    {
        "direct_dependents": 10,
        "transitive_dependents": 10,
        "affected_sheets": 10,
        "issue_groups": 20,
        "issue_examples_per_group": 2,
        "path_examples_per_kind": 1,
        "path_hops": 12,
    }
)
_FUNCTION = re.compile(r"[A-Za-z_\\][A-Za-z0-9_.\\]*\Z")


def _function(issue: AnalysisIssue) -> str | None:
    if issue.code not in {"unsupported_function", "dynamic_reference_function"}:
        return None
    text, span = issue.details.get("formula_text"), issue.source_span
    if not isinstance(text, str) or span is None:
        return None
    start, end = span
    if not 0 <= start < end <= len(text):
        return None
    identifier = text[start:end]
    if _FUNCTION.fullmatch(identifier) is None:
        return None
    if start and (text[start - 1].isalnum() or text[start - 1] in "_.\\$?"):
        return None
    if not text[end:].lstrip(" \t\r\n").startswith("("):
        return None
    return identifier.upper()


def _issues(batch: BatchImpact) -> Mapping[str, object]:
    retained: dict[str, AnalysisIssue] = {}
    summed = 0
    for root in batch.roots:
        if root.impact is not None:
            summed += len(root.impact.issues)
            for issue in root.impact.issues:
                retained.setdefault(digest(_to_json(issue)), issue)
    buckets: dict[tuple[str, str | None], list[AnalysisIssue]] = {}
    for issue in retained.values():
        buckets.setdefault((issue.code, _function(issue)), []).append(issue)
    ordered = sorted(buckets, key=lambda key: (-len(buckets[key]), key[0], key[1] or ""))
    groups: list[object] = []
    for code, function in ordered[:20]:
        issues = buckets[(code, function)]
        groups.append(
            {
                "code": code,
                "function": function,
                "count": len(issues),
                "examples": [
                    {
                        "sheet": i.sheet,
                        "formula_cell": _to_json(i.formula_cell),
                        "source_span": _to_json(i.source_span),
                    }
                    for i in issues[:2]
                ],
                "examples_truncated": len(issues) > 2,
            }
        )
    return {
        "retained_issue_count": len(retained),
        "sum_per_root_issue_count": summed,
        "retained_group_count": len(buckets),
        "omitted_group_count": max(0, len(buckets) - 20),
        "groups": groups,
    }


def _path(impact: DependencyImpact, target: CellRef, kind: str) -> Mapping[str, object]:
    path = impact.path_to(target)
    return {
        "kind": kind,
        "target": _to_json(target),
        "total_hops": len(path),
        "path_truncated": len(path) > 12,
        "hops": [
            {
                "precedent": _to_json(edge.precedent),
                "dependent": _to_json(edge.dependent),
                "reference_text": edge.evidence.reference_text,
                "source_span": _to_json(edge.evidence.source_span),
                "normalized_reference": _to_json(edge.evidence.normalized_reference),
                "basis": _to_json(edge.evidence.basis),
            }
            for edge in path[:12]
        ],
    }


def _root(root: RootImpact, order: Mapping[str, int]) -> Mapping[str, object]:
    impact = root.impact
    direct = (
        ()
        if impact is None
        else tuple(sorted(impact.direct_dependents, key=lambda ref: _cell_key(ref, order)))
    )
    transitive = (
        ()
        if impact is None
        else tuple(sorted(impact.transitive_dependents, key=lambda ref: _cell_key(ref, order)))
    )
    paths: list[object] = []
    if impact is not None:
        for kind, cells in (("direct", direct), ("transitive", transitive)):
            if cells:
                paths.append(_path(impact, cells[0], kind))
    return {
        "root": _to_json(root.root),
        "state": root.state,
        "analysis_status": impact.analysis_status if impact else "not_evaluated",
        "known_direct_count": impact.known_direct_count if impact else None,
        "known_transitive_count": impact.known_transitive_count if impact else None,
        "exact_total_count": impact.exact_total_count if impact else None,
        "blocked_budget": root.blocked_budget,
        "root_cursor": root.root_cursor,
        "direct_dependents_sample": [_to_json(ref) for ref in direct[:10]],
        "transitive_dependents_sample": [_to_json(ref) for ref in transitive[:10]],
        "direct_sample_truncated": len(direct) > 10,
        "transitive_sample_truncated": len(transitive) > 10,
        "inventory_frontier": _to_json(impact.inventory_frontier) if impact else (),
        "traversal_frontier": _to_json(impact.traversal_frontier) if impact else (),
        "path_examples": paths,
    }


def summarize_batch(
    batch: BatchImpact,
    *,
    source_fingerprint: str,
    sheet_order: Mapping[str, int],
) -> Mapping[str, object]:
    """Project the same snapshot and proofs without additional workbook work."""
    fingerprint(source_fingerprint)
    order = dict(sheet_order)
    if any(type(index) is not int or index < 0 for index in order.values()):
        raise ValueError("invalid sheet order index")
    if len(set(order.values())) != len(order):
        raise ValueError("duplicate sheet order index")
    cells = set(batch.known_union) | set(batch.transitive_union)
    trustworthy = False
    for root in batch.roots:
        cells.add(root.root)
        impact = root.impact
        if (root.state == "started") != (impact is not None) or root.state not in {
            "started",
            "not_started",
        }:
            raise ValueError("invalid root admission state")
        if impact is not None:
            if impact.source_fingerprint != source_fingerprint:
                raise ValueError("summary fingerprint differs from retained impact")
            if impact.root != root.root:
                raise ValueError("impact root differs from admission root")
            trustworthy |= impact.analysis_status in {"complete", "partial"}
            cells.update(impact.direct_dependents)
            cells.update(impact.transitive_dependents)
            for edge in impact.evidence_edges:
                cells.update((edge.precedent, edge.dependent))
    if any(ref.sheet not in order for ref in cells):
        raise ValueError("sheet order missing a retained cell's sheet")
    counts = Counter(ref.sheet for ref in set(batch.known_union))
    affected = [
        {"sheet": sheet, "known_affected_cell_count": counts[sheet]}
        for sheet in sorted(counts, key=order.__getitem__)
    ]
    return frozen_mapping(
        {
            "impact_summary_version": IMPACT_SUMMARY_VERSION,
            "dependency_contract_version": DEPENDENCY_CONTRACT_VERSION,
            "source_fingerprint": source_fingerprint,
            "scope": "worksheet_cell_formulas",
            "numerical_effects_evaluated": False,
            "sample_limits": _SAMPLE_LIMITS,
            "roots": [_root(root, order) for root in batch.roots],
            "known_union_count": len(batch.known_union) if trustworthy else None,
            "known_transitive_union_count": len(batch.transitive_union) if trustworthy else None,
            "known_affected_sheet_count": len(counts) if trustworthy else None,
            "affected_sheets": affected[:10],
            "affected_sheets_truncated": len(affected) > 10,
            "issues": _issues(batch),
        }
    )
