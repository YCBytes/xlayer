"""Direct gates for the refusal envelope and exception boundary."""

from __future__ import annotations

from typing import cast

import pytest

from xlayer._errors import (
    REFUSAL_SCHEMA_VERSION,
    ClosedWorkbookError,
    Refusal,
    WorkbookOpenError,
    XlayerError,
)

_RECOVERY = ({"action": "retry"},)


def test_to_dict_with_minimal_fields() -> None:
    refusal = Refusal(code="some_code", message="Refused.", recovery_options=_RECOVERY)
    assert refusal.to_dict() == {
        "refusal_schema_version": REFUSAL_SCHEMA_VERSION,
        "code": "some_code",
        "message": "Refused.",
        "operation": None,
        "target": None,
        "details": {},
        "recovery_options": [{"action": "retry"}],
    }


def test_to_dict_with_all_fields() -> None:
    refusal = Refusal(
        code="formula_cell",
        message="Target holds a formula.",
        operation="set_value",
        target="Forecast!D12",
        details={"formula_kind": "normal"},
        recovery_options=({"action": "choose_different_target"},),
    )
    serialized = refusal.to_dict()
    assert serialized["operation"] == "set_value"
    assert serialized["target"] == "Forecast!D12"
    assert serialized["details"] == {"formula_kind": "normal"}
    assert serialized["recovery_options"] == [{"action": "choose_different_target"}]


def test_empty_recovery_options_rejected() -> None:
    with pytest.raises(ValueError, match="at least one structured recovery option"):
        Refusal(code="c", message="m")
    with pytest.raises(ValueError, match="at least one structured recovery option"):
        Refusal(code="c", message="m", recovery_options=())


def test_empty_recovery_option_mapping_rejected() -> None:
    with pytest.raises(ValueError, match="non-empty mappings"):
        Refusal(code="c", message="m", recovery_options=({},))


def test_nested_details_are_recursively_frozen() -> None:
    refusal = Refusal(
        code="c",
        message="m",
        details={"outer": {"nested": [1, 2]}},
        recovery_options=_RECOVERY,
    )
    outer = refusal.details["outer"]
    # Nested mapping is read-only, nested list has become a tuple.
    mutable = cast("dict[str, object]", outer)
    with pytest.raises(TypeError):
        mutable["nested"] = "mutated"
    assert isinstance(outer, dict) is False
    nested = cast("dict[str, object]", outer)["nested"]
    assert nested == (1, 2)
    assert isinstance(nested, tuple)


def test_details_are_decoupled_from_caller_input() -> None:
    source: dict[str, object] = {"outer": {"nested": ["a"]}}
    refusal = Refusal(code="c", message="m", details=source, recovery_options=_RECOVERY)
    inner = source["outer"]
    assert isinstance(inner, dict)
    inner["nested"] = ["mutated"]
    outer = cast("dict[str, object]", refusal.details["outer"])
    assert outer["nested"] == ("a",)


def test_details_mapping_is_read_only() -> None:
    refusal = Refusal(code="c", message="m", details={"k": "v"}, recovery_options=_RECOVERY)
    mutable = cast("dict[str, object]", refusal.details)
    with pytest.raises(TypeError):
        mutable["k"] = "mutated"


def test_recovery_options_are_read_only() -> None:
    refusal = Refusal(code="c", message="m", recovery_options=_RECOVERY)
    mutable = cast("dict[str, object]", refusal.recovery_options[0])
    with pytest.raises(TypeError):
        mutable["action"] = "mutated"


def test_to_dict_returns_plain_json_friendly_copies() -> None:
    refusal = Refusal(
        code="c",
        message="m",
        details={"outer": {"nested": [1]}},
        recovery_options=_RECOVERY,
    )
    serialized = refusal.to_dict()
    details = serialized["details"]
    assert isinstance(details, dict)
    inner = details["outer"]
    assert isinstance(inner, dict)
    nested = inner["nested"]
    assert isinstance(nested, list)  # thawed back to JSON-friendly list
    nested.append(2)
    fresh = refusal.to_dict()["details"]
    assert isinstance(fresh, dict)
    assert fresh["outer"] == {"nested": [1]}


def test_workbook_open_error_wraps_refusal() -> None:
    refusal = Refusal(code="not_a_zip", message="Not a ZIP archive.", recovery_options=_RECOVERY)
    error = WorkbookOpenError(refusal)
    assert error.refusal is refusal
    assert str(error) == "Not a ZIP archive."
    assert isinstance(error, XlayerError)
    assert error.refusal.to_dict()["code"] == "not_a_zip"


def test_closed_workbook_error_hierarchy() -> None:
    with pytest.raises(XlayerError):
        raise ClosedWorkbookError("workbook is closed")
