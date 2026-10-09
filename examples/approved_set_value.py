"""Provider-free application integration with an exact-preview approval store.

This in-memory store is trusted demonstration application code, not user
authentication or a production approval service. Only pass permitted=True after
an independent application decision about this exact preview, never a model's
own assertion. The example grants only ordinary apply_set_value, not elevated
populated-text or partial-coverage permissions. It stores no receipt automatically.
"""

from __future__ import annotations

import os
from datetime import UTC, datetime
from uuid import uuid4

from xlayer import (
    Approval,
    Preview,
    Proposal,
    Receipt,
    Refusal,
    SetValue,
    VerifiedApproval,
    Workbook,
)


class ExactPreviewApprovalStore:
    """Remember explicitly permitted ordinary approvals; deny unrecorded claims."""

    def __init__(self) -> None:
        self._records: dict[str, Approval] = {}

    def record(self, proposal: Proposal, preview: Preview, *, permitted: bool) -> Approval | None:
        if type(permitted) is not bool:
            raise TypeError("permitted must be an explicit application bool")
        if (
            not permitted
            or preview.evidence["blocked"]
            or preview.evidence["required_capabilities"] != ("apply_set_value",)
            or preview.evidence["proposal_digest"] != proposal.proposal_digest
            or preview.evidence["source_fingerprint"] != proposal.source_fingerprint
        ):
            return None
        approval = Approval(
            schema_version="1.0",
            approval_id=str(uuid4()),
            issuer="example-application",
            actor_kind="service",
            source_fingerprint=proposal.source_fingerprint,
            proposal_digest=proposal.proposal_digest,
            preview_digest=preview.preview_digest,
            issued_at=datetime.now(UTC).isoformat(),
        )
        self._records[approval.approval_id] = approval
        return approval

    def verify(self, claim: Approval) -> VerifiedApproval:
        authorized = self._records.get(claim.approval_id) == claim
        return VerifiedApproval(claim, authorized, ("apply_set_value",) if authorized else ())


def edit_one_cell(
    source_path: str | os.PathLike[str],
    output_path: str | os.PathLike[str],
    edit: SetValue,
    *,
    permitted: bool,
) -> Receipt | Refusal:
    """Run the ordinary workflow; WorkbookOpenError is left to the caller."""
    with Workbook.open(source_path) as book:
        proposal = book.propose([edit], output_path=output_path)
        if isinstance(proposal, Refusal):
            return proposal
        preview = proposal.preview()
        if isinstance(preview, Refusal):
            return preview
        store = ExactPreviewApprovalStore()
        approval = store.record(proposal, preview, permitted=permitted)
        return proposal.apply(approval=approval, verify_approval=store.verify)
