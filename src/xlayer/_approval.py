"""Authority comes from the application's required verifier, not actor strings."""

from __future__ import annotations

from dataclasses import dataclass

from xlayer._canonical import fingerprint, string
from xlayer._errors import Refusal


def refusal(code: str, reason: str, **details: object) -> Refusal:
    return Refusal(
        code=code,
        message=f"SetValue refused: {reason}.",
        operation="set_value",
        details={"reason": reason, **details},
        recovery_options=({"action": "revise_or_reapprove_transaction", "reason": reason},),
    )


@dataclass(frozen=True)
class Approval:
    schema_version: str
    approval_id: str
    issuer: str
    actor_kind: str
    source_fingerprint: str
    proposal_digest: str
    preview_digest: str
    issued_at: str

    def __post_init__(self) -> None:
        if self.schema_version != "1.0":
            raise ValueError("unsupported approval schema")
        for name in ("approval_id", "issuer", "actor_kind", "issued_at"):
            string(getattr(self, name), name, nonempty=True)
        if self.actor_kind not in {"human", "service", "automation", "model"}:
            raise ValueError("unsupported approval actor kind")
        for name in ("source_fingerprint", "proposal_digest", "preview_digest"):
            fingerprint(getattr(self, name))

    def to_dict(self) -> dict[str, object]:
        return {name: getattr(self, name) for name in self.__dataclass_fields__}


@dataclass(frozen=True)
class VerifiedApproval:
    approval: Approval
    authorized: bool
    capabilities: tuple[str, ...]

    def __post_init__(self) -> None:
        if type(self.approval) is not Approval or type(self.authorized) is not bool:
            raise TypeError("host decision requires Approval and an exact bool")
        if not isinstance(self.capabilities, (tuple, list)):
            raise TypeError("capabilities must be a sequence of strings")
        capabilities = tuple(
            sorted({string(c, "capability", nonempty=True) for c in self.capabilities})
        )
        object.__setattr__(self, "capabilities", capabilities)


def check_approval(
    approval: Approval,
    decision: VerifiedApproval,
    source: str,
    proposal: str,
    preview: str,
    required: tuple[str, ...],
) -> Refusal | None:
    if type(decision) is not VerifiedApproval:
        raise TypeError("verifier must return VerifiedApproval")
    if (approval.source_fingerprint, approval.proposal_digest, approval.preview_digest) != (
        source,
        proposal,
        preview,
    ):
        return refusal("preview_digest_mismatch", "approval does not match this exact preview")
    if decision.approval != approval or not decision.authorized:
        return refusal("approval_rejected", "host denied or mismatched the approval identity")
    if not set(required) <= set(decision.capabilities):
        return refusal(
            "approval_rejected", "host has not granted required capabilities", required=required
        )
    if len(required) > 1 and decision.approval.actor_kind not in {"human", "service"}:
        return refusal(
            "approval_rejected", "elevated changes require independent external authority"
        )
    return None
