"""Snapshot-owned verification of a finite set of saved-workbook claims."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, fields
from typing import TYPE_CHECKING, cast

from xlayer._canonical import canonical_json, digest, frozen_mapping, thaw_json
from xlayer._dependencies import DEFAULT_IMPACT_LIMITS, ImpactLimits
from xlayer._errors import Refusal
from xlayer._explanation_contract import (
    DEFAULT_EXPLANATION_LIMITS,
    ExplanationLimits,
    ExplanationRequest,
    ExplanationScope,
    ExplanationSizeError,
    array,
    authorized,
    bounded_bytes,
    check_explanation_limits,
    explanation_refusal,
    object_dict,
)
from xlayer._explanation_render import LIMITATIONS, RENDERER_VERSION, render_report
from xlayer._inspection import DEFAULT_READ_LIMITS, ReadLimits, check_limits

if TYPE_CHECKING:
    from xlayer._public_workbook import Workbook


@dataclass(frozen=True, init=False)
class CheckedExplanation:
    """A detached checked report, produced only by Workbook.check_explanation()."""

    evidence: Mapping[str, object]
    _encoded: bytes

    def __init__(self) -> None:
        raise TypeError("use Workbook.check_explanation")

    @classmethod
    def _create(cls, body: dict[str, object], max_bytes: int) -> CheckedExplanation:
        bounded_bytes({**body, "report_digest": "sha256:" + "0" * 64}, max_bytes)
        complete = {**body, "report_digest": digest(body)}
        result = object.__new__(cls)
        object.__setattr__(result, "evidence", frozen_mapping(complete))
        object.__setattr__(result, "_encoded", bounded_bytes(complete, max_bytes))
        return result

    def to_dict(self) -> dict[str, object]:
        return object_dict(thaw_json(self.evidence))

    def canonical_json(self) -> bytes:
        return self._encoded

    @property
    def report_digest(self) -> str:
        return cast("str", self.evidence["report_digest"])

    def render_text(self) -> str:
        return cast("str", self.evidence["rendered_text"])


def claim_selectors(claim: dict[str, object]) -> list[dict[str, object]]:
    if claim["kind"] == "fact":
        return [object_dict(claim["selector"])]
    if claim["kind"] == "all_equal":
        return [
            {"family": "cell", "field": claim["field"], **object_dict(member)}
            for member in array(claim["cells"])
        ]
    # Used for scope authorization only; membership covers no required fact.
    return [{"family": "impact", "field": "analysis_status"}]


def project(
    observations: dict[str, object], item: dict[str, object]
) -> tuple[bool, object, list[object]]:
    family, name = item["family"], cast("str", item["field"])
    if family == "cell":
        for i, inspection in enumerate(array(observations["cells"])):
            data = object_dict(object_dict(inspection)["data"])
            if data["sheet"] != item["sheet"]:
                continue
            for j, value in enumerate(array(data["cells"])):
                cell = object_dict(value)
                if cell["address"] != item["address"]:
                    continue
                formula = cell["formula"]
                if name == "has_formula":
                    observed: object = formula is not None
                elif name in {"formula_text", "formula_kind", "shared_master_text"}:
                    if formula is None:
                        observed = None
                    elif name == "shared_master_text":
                        master = object_dict(formula)["shared_master"]
                        observed = None if master is None else object_dict(master)["text"]
                    else:
                        observed = object_dict(formula)[
                            "text" if name == "formula_text" else "kind"
                        ]
                elif name in {"format_temporal_kind", "merge_role"}:
                    value = cell["cell_format" if name == "format_temporal_kind" else "merge"]
                    observed = (
                        None
                        if value is None
                        else object_dict(value)[
                            "temporal_kind" if name == "format_temporal_kind" else "role"
                        ]
                    )
                else:
                    observed = cell[name]
                return True, observed, ["observations", "cells", i, "data", "cells", j]
    elif family in {"inventory", "sheet"} and observations["inventory"] is not None:
        data = object_dict(object_dict(observations["inventory"])["data"])
        if family == "inventory":
            value = len(array(data["sheets"])) if name == "sheet_count" else data[name]
            return True, value, ["observations", "inventory", "data"]
        for i, value in enumerate(array(data["sheets"])):
            tab = object_dict(value)
            if tab["name"] == item["sheet"]:
                return True, tab[name], ["observations", "inventory", "data", "sheets", i]
    elif family == "impact" and observations["impact"] is not None:
        summary = object_dict(
            object_dict(object_dict(observations["impact"])["data"])["impact_summary"]
        )
        if name in {"scope", "numerical_effects_evaluated", "known_union_count"}:
            return True, summary[name], ["observations", "impact", "data", "impact_summary"]
        root = object_dict(array(summary["roots"])[0])
        return True, root[name], ["observations", "impact", "data", "impact_summary", "roots", 0]
    return False, None, []


def equal(actual: object, expected: object, field: str) -> bool:
    if field == "stored_value" and type(actual) in {int, float} and type(expected) in {int, float}:
        # Python's mixed comparison preserves large-integer distinctions.
        return actual == expected
    if type(actual) is not type(expected):
        return False
    if isinstance(actual, dict):
        other = object_dict(expected)
        return set(actual) == set(other) and all(equal(v, other[k], "") for k, v in actual.items())
    if isinstance(actual, list):
        members = array(expected)
        return len(actual) == len(members) and all(
            equal(a, b, "") for a, b in zip(actual, members, strict=True)
        )
    return actual == expected


def check_explanation(
    book: Workbook,
    request: ExplanationRequest,
    *,
    scope: ExplanationScope,
    limits: ExplanationLimits = DEFAULT_EXPLANATION_LIMITS,
    read_limits: ReadLimits = DEFAULT_READ_LIMITS,
    impact_limits: ImpactLimits = DEFAULT_IMPACT_LIMITS,
) -> CheckedExplanation | Refusal:
    book._ensure_open()
    if type(request) is not ExplanationRequest or type(scope) is not ExplanationScope:
        raise TypeError("exact ExplanationRequest and ExplanationScope records are required")
    check_explanation_limits(limits)
    check_limits(read_limits, DEFAULT_READ_LIMITS, "read_limits")
    check_limits(impact_limits, DEFAULT_IMPACT_LIMITS, "impact_limits")

    def refuse(code: str, **details: object) -> Refusal:
        return explanation_refusal(code, details=details, limits=limits)

    def capture_failure(cause: Refusal, stage: str, subject: object) -> Refusal:
        # Existing operational refusals may contain an inventory, unselected cell
        # or source path. Only fresh, authorized context crosses this boundary.
        return refuse(
            cause.code,
            cause_code=cause.code,
            stage=stage,
            subject=subject,
            source_fingerprint=book.source_fingerprint,
            scope_digest=scope.scope_digest,
        )

    try:
        checked = ExplanationRequest.from_dict(request.to_dict(), limits=limits)
        if isinstance(checked, Refusal):
            return explanation_refusal(checked.code, details=checked.details, limits=limits)
        bounded_bytes(scope.to_dict(), limits.max_request_bytes)
        if sum(len(v) for v in scope.cells.values()) > limits.max_cells:
            return refuse("explanation_limit_exceeded", reason="max_cells")
        if len(scope.required_facts) > limits.max_required_facts:
            return refuse("explanation_limit_exceeded", reason="max_required_facts")
        data = checked.to_dict()
        if (
            data["source_fingerprint"] != book.source_fingerprint
            or data["scope_digest"] != scope.scope_digest
        ):
            return refuse(
                "explanation_binding_mismatch",
                source_fingerprint=book.source_fingerprint,
                scope_digest=scope.scope_digest,
            )
        claims = [object_dict(c) for c in array(data["claims"])]
        for claim in claims:
            for item in claim_selectors(claim):
                if not authorized(item, scope):
                    return refuse("explanation_scope_exceeded", claim_id=claim["id"])
        observations: dict[str, object] = {"inventory": None, "cells": [], "impact": None}
        if scope.include_inventory:
            inventory = book.list_sheets(limits=read_limits)
            if isinstance(inventory, Refusal):
                return capture_failure(inventory, "inventory", None)
            observations["inventory"] = inventory.to_dict()
        bounded_bytes(observations, limits.max_evidence_bytes)
        order = {tab.name: tab.tab_index for tab in book._book.registry.sheets}
        for name in sorted(scope.cells, key=lambda value: (order.get(value, len(order)), value)):
            cells = book.read_cells(name, scope.cells[name], limits=read_limits)
            if isinstance(cells, Refusal):
                return capture_failure(
                    cells, "cells", {"sheet": name, "addresses": list(scope.cells[name])}
                )
            array(observations["cells"]).append(cells.to_dict())
            bounded_bytes(observations, limits.max_evidence_bytes)
        if scope.impact_root is not None:
            impact = book.dependency_impact(
                *scope.impact_root, limits=impact_limits, read_limits=read_limits
            )
            if isinstance(impact, Refusal):
                return capture_failure(
                    impact,
                    "impact",
                    {"sheet": scope.impact_root[0], "address": scope.impact_root[1]},
                )
            observations["impact"] = impact.to_dict()
        bounded_bytes(observations, limits.max_evidence_bytes)
        covered: set[bytes] = set()
        witnesses: list[object] = []
        problems: list[dict[str, object]] = []
        total = 0

        def failure(
            code: str,
            claim: dict[str, object],
            *,
            item: object = None,
            actual: object = None,
            path: list[object] | None = None,
        ) -> None:
            nonlocal total
            total += 1
            if len(problems) >= limits.max_diagnostics:
                return
            detail: dict[str, object] = {"code": code, "claim_id": claim["id"]}
            if item is not None:
                detail["selector"] = item
            if path is not None:
                detail.update(witness_path=path, observed_digest=digest(actual))
                try:
                    bounded_bytes(actual, 256)
                except ExplanationSizeError:
                    pass
                else:
                    detail["observed"] = actual
            problems.append(detail)

        for claim in claims:
            if claim["kind"] == "known_dependent":
                summary = object_dict(
                    object_dict(object_dict(observations["impact"])["data"])["impact_summary"]
                )
                relation = cast("str", claim["relation"])
                key = relation + "_dependents_sample"
                sample = array(object_dict(array(summary["roots"])[0])[key])
                index = next(
                    (i for i, target in enumerate(sample) if equal(target, claim["target"], "")),
                    None,
                )
                if index is None:
                    failure("explanation_evidence_missing", claim)
                else:
                    witnesses.append(
                        {
                            "claim_id": claim["id"],
                            "target": claim["target"],
                            "relation": relation,
                            "path": [
                                "observations",
                                "impact",
                                "data",
                                "impact_summary",
                                "roots",
                                0,
                                key,
                                index,
                            ],
                            "observed_digest": digest(sample[index]),
                        }
                    )
                continue
            for item in claim_selectors(claim):
                available, actual, path = project(observations, item)
                if not available:
                    failure("explanation_evidence_missing", claim, item=item)
                elif not equal(actual, claim["expected"], cast("str", item["field"])):
                    failure("explanation_contradicted", claim, item=item, actual=actual, path=path)
                else:
                    covered.add(canonical_json(item))
                    witnesses.append(
                        {
                            "claim_id": claim["id"],
                            "selector": item,
                            "path": path,
                            "observed_digest": digest(actual),
                        }
                    )
        if problems:
            return refuse(
                cast("str", problems[0]["code"]),
                diagnostics=problems,
                total_known=total,
                diagnostics_truncated=total > len(problems),
            )
        unmet = [thaw_json(s) for s in scope.required_facts if canonical_json(s) not in covered]
        if unmet:
            return refuse(
                "explanation_requirements_unmet",
                missing_selectors=unmet[: limits.max_diagnostics],
                total_known=len(unmet),
                diagnostics_truncated=len(unmet) > limits.max_diagnostics,
            )
        body: dict[str, object] = {
            "schema_version": "1.0",
            "explanation_contract_version": "1.0",
            "renderer_version": RENDERER_VERSION,
            "source_fingerprint": book.source_fingerprint,
            "scope": scope.to_dict(),
            "scope_digest": scope.scope_digest,
            "request_digest": digest(data),
            "evidence_digest": digest(observations),
            "observations": observations,
            "accepted_claims": claims,
            "witnesses": witnesses,
            "required_coverage": thaw_json(scope.required_facts),
            "limits": {
                "explanation": limits.to_dict(),
                "read": {f.name: getattr(read_limits, f.name) for f in fields(read_limits)},
                "impact": {f.name: getattr(impact_limits, f.name) for f in fields(impact_limits)},
            },
            "limitations": LIMITATIONS,
        }
        bounded_bytes(body, limits.max_report_bytes)
        body["rendered_text"] = render_report(body, limits.max_report_bytes)
        return CheckedExplanation._create(body, limits.max_report_bytes)
    except ExplanationSizeError as exc:
        return refuse("explanation_limit_exceeded", reason=str(exc))
    except BaseException:
        book.close()
        raise
