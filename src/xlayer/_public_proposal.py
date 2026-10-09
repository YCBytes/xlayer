"""Public proposal access; approval, verification and publication remain engine-owned."""

from __future__ import annotations

from collections.abc import Callable

from xlayer._approval import Approval, VerifiedApproval
from xlayer._errors import Refusal
from xlayer._preview import Preview
from xlayer._proposal import Proposal as _EngineProposal
from xlayer._public_edits import SetValue
from xlayer._receipt import Receipt


class Proposal:
    """An ordered whole-batch intent produced by Workbook.propose().

    Direct construction and durable reconstruction are not supported. The host
    authenticates an exact approval; the model cannot authorize itself.
    """

    __slots__ = ("_edits", "_proposal")

    def __init__(self, proposal: _EngineProposal, edits: tuple[SetValue, ...]) -> None:
        self._proposal = proposal
        self._edits = edits

    @property
    def proposal_digest(self) -> str:
        return self._proposal.proposal_digest

    @property
    def source_fingerprint(self) -> str:
        return self._proposal._book.source_fingerprint

    @property
    def edits(self) -> tuple[SetValue, ...]:
        return self._edits

    def preview(self) -> Preview | Refusal:
        return self._proposal.preview()

    def apply(
        self,
        *,
        approval: Approval | None,
        verify_approval: Callable[[Approval], VerifiedApproval],
    ) -> Receipt | Refusal:
        """Revalidate and publish once; a verifier is required, never inferred."""
        return self._proposal.apply(approval=approval, verify_approval=verify_approval)
