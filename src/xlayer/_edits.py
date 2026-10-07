"""Private typed SetValue intent; construction never authorizes mutation."""

from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import date
from typing import TypeAlias

from xlayer._canonical import string
from xlayer._dependencies import _canonical_address

Scalar: TypeAlias = int | float | bool | str | date


def scalar_record(value: Scalar) -> dict[str, object]:
    return {
        "kind": type(value).__name__,
        "value": value.isoformat() if type(value) is date else value,
    }


@dataclass(frozen=True)
class SetValue:
    sheet: str
    cell: str
    value: Scalar

    def __post_init__(self) -> None:
        string(self.sheet, "sheet", nonempty=True)
        string(self.cell, "cell", nonempty=True)
        object.__setattr__(self, "cell", _canonical_address(self.cell))
        if type(self.value) not in {int, float, bool, str, date}:
            raise TypeError("SetValue accepts exact int, float, bool, str or datetime.date")
        if (
            type(self.value) is int
            and not -999_999_999_999_999 <= self.value <= 999_999_999_999_999
        ):
            raise ValueError("integer exceeds the supported 15-digit range")
        if type(self.value) is float and not math.isfinite(self.value):
            raise ValueError("float must be finite")
        if type(self.value) is str:
            string(self.value, "value")
            if len(self.value.encode("utf-16-le")) // 2 > 32767 or self.value.count("\n") > 253:
                raise ValueError("text exceeds UTF-16 length or line-feed limit")

    def to_dict(self) -> dict[str, object]:
        return {"sheet": self.sheet, "cell": self.cell, "value": scalar_record(self.value)}
