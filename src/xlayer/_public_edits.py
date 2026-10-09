"""Promoted scalar intent; validation and serialization stay in the existing engine."""

from __future__ import annotations

from dataclasses import dataclass, field

from xlayer._edits import SetValue as _EngineEdit


@dataclass(frozen=True)
class SetValue:
    """Propose an existing-cell scalar change, without authorizing it.

    Exact int, finite float, bool and str are supported. Date writes and scalar
    subclasses are deliberately outside this alpha's public support boundary.
    """

    sheet: str
    cell: str
    value: int | float | bool | str
    _edit: _EngineEdit = field(init=False, repr=False, compare=False)

    def __post_init__(self) -> None:
        if type(self.value) not in {int, float, bool, str}:
            raise TypeError("SetValue accepts exact int, finite float, bool or str")
        edit = _EngineEdit(self.sheet, self.cell, self.value)
        object.__setattr__(self, "cell", edit.cell)
        object.__setattr__(self, "_edit", edit)

    def to_dict(self) -> dict[str, object]:
        return self._edit.to_dict()
