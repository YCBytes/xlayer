"""The alpha input adapter preserves exact scalar kinds, not implicit coercions."""

from __future__ import annotations

from dataclasses import FrozenInstanceError
from datetime import date, datetime

import pytest

import xlayer


class IntSubclass(int):
    pass


class TextSubclass(str):
    pass


@pytest.mark.parametrize("value", [0, -999_999_999_999_999, 1.25, -0.0, True, False, "", "😀"])
def test_public_edit_preserves_scalar_and_canonical_address(
    value: int | float | bool | str,
) -> None:
    edit = xlayer.SetValue("Inputs", "a1", value)
    assert edit.sheet == "Inputs" and edit.cell == "A1"
    assert type(edit.value) is type(value)
    assert edit.to_dict() == {
        "sheet": "Inputs",
        "cell": "A1",
        "value": {"kind": type(value).__name__, "value": value},
    }
    with pytest.raises(FrozenInstanceError):
        edit.value = 5  # type: ignore[misc]


@pytest.mark.parametrize(
    "value", [date(2024, 1, 1), datetime(2024, 1, 1), IntSubclass(1), TextSubclass("x"), None, b"x"]
)
def test_public_edit_does_not_promote_dates_or_subclasses(value: object) -> None:
    with pytest.raises(TypeError):
        xlayer.SetValue("Inputs", "A1", value)  # type: ignore[arg-type]


@pytest.mark.parametrize(
    "value", [float("inf"), float("nan"), 10**15, "\ud800", "x" * 32768, "\n" * 254]
)
def test_public_edit_uses_existing_value_guards(value: int | float | str) -> None:
    with pytest.raises(ValueError):
        xlayer.SetValue("Inputs", "A1", value)


@pytest.mark.parametrize("address", ["$A$1", "XFE1", "A1048577", "A1٢", " A1", "S!A1", "A1:A2"])
def test_public_edit_uses_existing_address_guards(address: str) -> None:
    with pytest.raises(ValueError):
        xlayer.SetValue("Inputs", address, 1)


@pytest.mark.parametrize("value", [True, False, 1.0, 0, -1])
def test_archive_limit_requires_exact_positive_integer(value: object) -> None:
    from xlayer._ooxml.archive import ArchiveLimits

    with pytest.raises(ValueError, match="max_file_bytes"):
        ArchiveLimits(max_file_bytes=value)  # type: ignore[arg-type]
