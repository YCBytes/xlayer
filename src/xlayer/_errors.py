"""Structured refusals and the exception boundary.

Private until slice 1 ships: the module is named ``_errors`` so the public
namespace stays empty ahead of the promotion gates, and will be renamed to
``errors`` as a deliberate export decision when the slice completes.

The rule, from the slice 1 contract: an exception is raised only when no
session or result object can exist — opening a workbook, or misusing the API
itself. Every operation on an existing session returns a structured
:class:`Refusal` instead of raising, so a caller can branch on stable codes
rather than parse exception strings.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from types import MappingProxyType
from typing import cast

REFUSAL_SCHEMA_VERSION = "1.0"


def _freeze(value: object) -> object:
    """Recursively convert to read-only structures, decoupled from the input."""
    if isinstance(value, Mapping):
        return MappingProxyType({str(key): _freeze(item) for key, item in value.items()})
    if isinstance(value, list | tuple):
        return tuple(_freeze(item) for item in value)
    if isinstance(value, set | frozenset):
        return tuple(sorted((_freeze(item) for item in value), key=repr))
    return value


def _thaw(value: object) -> object:
    """Recursively convert frozen structures back to plain JSON-friendly ones."""
    if isinstance(value, Mapping):
        return {key: _thaw(item) for key, item in value.items()}
    if isinstance(value, tuple):
        return [_thaw(item) for item in value]
    return value


@dataclass(frozen=True)
class Refusal:
    """A machine-readable operational refusal.

    ``code`` values are a stable public contract; changing or removing one is
    a breaking change. ``recovery_options`` entries are structured actions a
    caller can take, not prose; construction rejects a refusal without at
    least one non-empty option, so the guarantee holds by type rather than by
    review.

    All nested data is recursively frozen at construction and decoupled from
    the caller's structures, so neither input aliasing nor mutation of any
    nested value can change a refusal after the fact.
    """

    code: str
    message: str
    operation: str | None = None
    target: str | None = None
    details: Mapping[str, object] = field(default_factory=dict)
    recovery_options: tuple[Mapping[str, object], ...] = ()

    def __post_init__(self) -> None:
        if not self.recovery_options:
            raise ValueError(
                f"Refusal {self.code!r} must carry at least one structured recovery option"
            )
        frozen_options: list[Mapping[str, object]] = []
        for option in self.recovery_options:
            if not option:
                raise ValueError(
                    f"Refusal {self.code!r} has an empty recovery option; options must be "
                    "non-empty mappings"
                )
            frozen_options.append(cast("Mapping[str, object]", _freeze(dict(option))))
        object.__setattr__(
            self, "details", cast("Mapping[str, object]", _freeze(dict(self.details)))
        )
        object.__setattr__(self, "recovery_options", tuple(frozen_options))

    def to_dict(self) -> dict[str, object]:
        return {
            "refusal_schema_version": REFUSAL_SCHEMA_VERSION,
            "code": self.code,
            "message": self.message,
            "operation": self.operation,
            "target": self.target,
            "details": cast("dict[str, object]", _thaw(self.details)),
            "recovery_options": [
                cast("dict[str, object]", _thaw(option)) for option in self.recovery_options
            ],
        }


class XlayerError(Exception):
    """Base class for every exception raised by xlayer."""


class WorkbookOpenError(XlayerError):
    """Raised when a workbook cannot be opened at all.

    Opening is the one operational surface that raises rather than returns,
    because no session object can exist on failure and a union return would
    break ``with Workbook.open(...)``. The full structured refusal is carried
    on :attr:`refusal` and serializes identically to returned refusals.
    """

    def __init__(self, refusal: Refusal) -> None:
        super().__init__(refusal.message)
        self.refusal = refusal


class ClosedWorkbookError(XlayerError):
    """Raised when a method is called on a closed workbook (programming error)."""
