"""Private immutable preview evidence; facts are not approval or recalculation."""

from __future__ import annotations

from dataclasses import dataclass

from xlayer._canonical import EvidenceRecord, canonical_json


@dataclass(frozen=True)
class Preview(EvidenceRecord):
    @property
    def preview_digest(self) -> str:
        return self._digest

    def to_dict(self) -> dict[str, object]:
        return self._dict("preview_digest")

    def canonical_json(self) -> bytes:
        return canonical_json(self.to_dict())
