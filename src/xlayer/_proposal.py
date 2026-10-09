"""Private immutable intent and deterministic preview over an open read snapshot."""

from __future__ import annotations

import os
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING

from xlayer import __version__
from xlayer._approval import Approval, VerifiedApproval, refusal
from xlayer._batch_impact import analyse_batch
from xlayer._canonical import EvidenceTooLarge, digest, frozen_mapping
from xlayer._dependencies import DEFAULT_IMPACT_LIMITS, CellRef, ImpactLimits, _to_json
from xlayer._edit_validation import EditPlan, validate_edit, write_context
from xlayer._edits import SetValue
from xlayer._errors import ClosedWorkbookError, Refusal
from xlayer._ooxml.zip_write import check_write_support
from xlayer._output import OutputSpec, bind_output
from xlayer._preview import Preview
from xlayer._receipt import Receipt

if TYPE_CHECKING:
    from xlayer._workbook import Workbook

VERSIONS = {
    "schema_version": "1.0",
    "transaction_contract_version": "1.1",
    "package_version": __version__,
    "mutation_version": "1.1",
    "verification_version": "1.1",
    "policy_version": "1.1",
}
THRESHOLDS = {"known_transitive_union_gt": 100, "known_affected_sheets_gt": 3, "batch_size_gte": 50}


@dataclass(frozen=True)
class Proposal:
    _book: Workbook = field(repr=False, compare=False)
    edits: tuple[SetValue, ...]
    output: OutputSpec
    limits: ImpactLimits
    annotations: Mapping[str, object]
    proposal_digest: str = field(init=False)
    _cached_preview: Preview | None = field(default=None, init=False, repr=False, compare=False)
    _receipt: Receipt | None = field(default=None, init=False, repr=False, compare=False)
    _applying: bool = field(default=False, init=False, repr=False, compare=False)

    def __post_init__(self) -> None:
        object.__setattr__(self, "annotations", frozen_mapping(self.annotations))
        object.__setattr__(self, "proposal_digest", digest(self.identity()))

    def identity(self) -> dict[str, object]:
        return {
            **VERSIONS,
            "source_fingerprint": self._book.source_fingerprint,
            "source_binding": str(self._book._source_binding),
            "source_resolved": str(self._book._source_resolved),
            "output": self.output.to_dict(),
            "edits": [edit.to_dict() for edit in self.edits],
            "limits": _to_json(self.limits),
        }

    def _ensure_open(self) -> None:
        if self._book.closed:
            raise ClosedWorkbookError("The proposal's workbook is closed.")

    def _bind_refusal(self, result: Refusal) -> Refusal:
        """Correlate operational failures with expected, never claimed, identities."""
        details = {
            **dict(result.details),
            "source_fingerprint": self._book.source_fingerprint,
            "proposal_digest": self.proposal_digest,
        }
        if self._cached_preview is not None:
            details["preview_digest"] = self._cached_preview.preview_digest
        return replace(result, details=details)

    def preview(self) -> Preview | Refusal:
        self._ensure_open()
        try:
            if self._cached_preview is not None:
                return self._cached_preview
            built = self._build_preview()
            if isinstance(built, Refusal):
                return self._bind_refusal(built)
            preview, _plans = built
            object.__setattr__(self, "_cached_preview", preview)
            return preview
        except EvidenceTooLarge:
            return self._bind_refusal(
                refusal("transaction_limit_exceeded", "preview exceeds canonical evidence ceiling")
            )
        except BaseException:
            self._book.close()
            raise

    def _build_preview(self) -> tuple[Preview, tuple[EditPlan, ...]] | Refusal:
        self._ensure_open()
        incumbent = self.output.inspect(self._book)
        if isinstance(incumbent, Refusal):
            return incumbent
        context = write_context(self._book)
        if isinstance(context, Refusal):
            return context
        findings: list[dict[str, object]] = []
        plans: list[EditPlan] = []
        missing_parents: set[str] = set()
        for cursor, edit in enumerate(self.edits):
            if edit.sheet in missing_parents:
                continue
            result = validate_edit(self._book, edit, context=context)
            if isinstance(result, Refusal):
                findings.append(
                    {
                        "edit_index": cursor,
                        "sheet": edit.sheet,
                        "cell": edit.cell,
                        "refusal": result.to_dict(),
                    }
                )
                if result.code in {
                    "sheet_not_found",
                    "unsupported_sheet_kind",
                    "invalid_part_content",
                    "malformed_xml",
                    "missing_required_part",
                }:
                    missing_parents.add(edit.sheet)
            else:
                plans.append(result)
        parts = {plan.part for plan in plans}
        if len(parts) > 1:
            findings.append(
                {
                    "refusal": refusal(
                        "multiple_target_parts", "one worksheet per transaction"
                    ).to_dict()
                }
            )
        effective = tuple(plan for plan in plans if plan.effective)
        if not effective and not findings:
            return refusal("no_changes", "all validated edits preserve their existing payloads")
        if effective:
            archive = self._book._archive
            if archive is None:
                raise RuntimeError("preview requires an open source archive")
            write_failure = check_write_support(archive)
            if write_failure is not None:
                findings.append({"refusal": write_failure.to_dict()})
        required = {"apply_set_value"}
        if any(plan.populated_text for plan in effective):
            required.add("change_populated_text")
        analysis = (
            analyse_batch(
                self._book.registry,
                self._book.read_sheet,
                tuple(CellRef(plan.edit.sheet, plan.edit.cell) for plan in effective),
                self._book.source_fingerprint,
                self.limits,
            )
            if effective
            else None
        )
        if analysis is not None:
            for root in analysis.roots:
                if root.impact is None:
                    findings.append(
                        {
                            "refusal": refusal(
                                "transaction_limit_exceeded",
                                "root admission budget exhausted",
                                root_cursor=root.root_cursor,
                                budget=root.blocked_budget,
                            ).to_dict()
                        }
                    )
                elif root.impact.analysis_status in {"unknown", "not_evaluated"}:
                    findings.append(
                        {
                            "refusal": refusal(
                                "dependency_analysis_untrustworthy",
                                "dependency evidence cannot authorize this slice",
                            ).to_dict()
                        }
                    )
                elif root.impact.analysis_status == "partial":
                    required.add("accept_partial_dependency_coverage")
        union = analysis.known_union if analysis else ()
        has_proof = analysis is not None and any(
            root.impact is not None and root.impact.analysis_status in {"complete", "partial"}
            for root in analysis.roots
        )
        known_transitive = len(analysis.transitive_union) if analysis and has_proof else None
        known_sheets = len({ref.sheet for ref in union}) if has_proof else None
        measured = {
            "known_transitive_union": known_transitive,
            "known_affected_sheets": known_sheets,
            "batch_size": len(self.edits),
        }
        broad = (
            (
                known_transitive is not None
                and known_transitive > THRESHOLDS["known_transitive_union_gt"]
            )
            or (known_sheets is not None and known_sheets > THRESHOLDS["known_affected_sheets_gt"])
            or len(self.edits) >= THRESHOLDS["batch_size_gte"]
        )
        next_action = (
            "revise_proposal"
            if findings
            else (
                "obtain_external_approval"
                if len(required) > 1
                else "inspect_known_impact"
                if broad
                else "obtain_approval"
            )
        )
        evidence = {
            **VERSIONS,
            "source_fingerprint": self._book.source_fingerprint,
            "proposal_digest": self.proposal_digest,
            "output": self.output.to_dict(),
            "output_incumbent": incumbent,
            "edits": [plan.to_dict() for plan in plans],
            "requested_edits": [edit.to_dict() for edit in self.edits],
            "findings": findings,
            "blocked": bool(findings),
            "dependencies": analysis.to_dict() if analysis else None,
            "limits": _to_json(self.limits),
            "thresholds": THRESHOLDS,
            "measurements": measured,
            "breadth_advisory": broad,
            "required_capabilities": sorted(required),
            "next_action": next_action,
            "expected_parts": sorted(parts),
            "recalculation_status": "required_not_run" if union else "unknown",
            "dependency_trigger": "known_dependents"
            if union
            else "no_dependency_trigger_found"
            if analysis
            and all(r.impact and r.impact.analysis_status == "complete" for r in analysis.roots)
            else "unknown",
            "visual_status": "not_evaluated",
            "guidance": (
                "Potential references are not calculated effects. Obtain host full recalculation "
                "before trusting downstream answers; visual rendering was not evaluated."
            ),
        }
        preview = Preview(evidence, {"caller_annotations": self.annotations})
        preview.canonical_json()
        return preview, tuple(plans)

    def apply(
        self, *, approval: Approval | None, verify_approval: Callable[[Approval], VerifiedApproval]
    ) -> Receipt | Refusal:
        self._ensure_open()
        if approval is not None and type(approval) is not Approval:
            raise TypeError("approval must be an Approval or None")
        if not callable(verify_approval):
            raise TypeError("a trusted host verifier callable is required")
        if self._applying:
            raise RuntimeError("reentrant proposal use is unsupported")
        from xlayer._transaction import apply_proposal

        object.__setattr__(self, "_applying", True)
        try:
            result = apply_proposal(self, approval, verify_approval)
            return self._bind_refusal(result) if isinstance(result, Refusal) else result
        except EvidenceTooLarge:
            return self._bind_refusal(
                refusal("transaction_limit_exceeded", "canonical evidence exceeds 16 MiB")
            )
        except BaseException:
            self._book.close()
            raise
        finally:
            object.__setattr__(self, "_applying", False)


def create_proposal(
    book: Workbook,
    edits: Sequence[SetValue],
    output_path: str | os.PathLike[str],
    overwrite: bool = False,
    *,
    limits: ImpactLimits = DEFAULT_IMPACT_LIMITS,
    annotations: Mapping[str, object] | None = None,
) -> Proposal | Refusal:
    if book.closed:
        raise ClosedWorkbookError("The workbook is closed.")
    candidate: object = edits
    if not isinstance(candidate, Sequence) or isinstance(candidate, (str, bytes)):
        raise TypeError("edits must be a sequence of SetValue")
    if not 1 <= len(edits) <= 100:
        raise ValueError("a transaction contains 1..100 edits")
    if any(type(edit) is not SetValue for edit in edits) or type(limits) is not ImpactLimits:
        raise TypeError("exact SetValue records and ImpactLimits are required")
    captured = tuple(edits)
    if len({(edit.sheet, edit.cell) for edit in captured}) != len(captured):
        return refusal(
            "duplicate_target",
            "duplicate canonical target in ordered batch",
            source_fingerprint=book.source_fingerprint,
        )
    output = bind_output(book, output_path, overwrite)
    if isinstance(output, Refusal):
        return replace(
            output, details={**dict(output.details), "source_fingerprint": book.source_fingerprint}
        )
    return Proposal(book, captured, output, limits, {} if annotations is None else annotations)
