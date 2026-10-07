"""Literal contract expectations for intent, authority and canonical evidence."""

from __future__ import annotations

import importlib.util
import json
from datetime import date, datetime
from typing import TYPE_CHECKING

import pytest

if TYPE_CHECKING:
    from xlayer._edits import SetValue


def edit(value: object, cell: str = "a1") -> SetValue:
    assert importlib.util.find_spec("xlayer._edits") is not None, "typed intent is missing"
    from xlayer._edits import SetValue

    return SetValue("Inputs", cell, value)  # type: ignore[arg-type]


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (1, {"kind": "int", "value": 1}),
        (1.0, {"kind": "float", "value": 1.0}),
        (True, {"kind": "bool", "value": True}),
        ("", {"kind": "str", "value": ""}),
        (date(2024, 2, 29), {"kind": "date", "value": "2024-02-29"}),
    ],
)
def test_exact_scalar_kind_and_canonical_coordinate(value: object, expected: object) -> None:
    result = edit(value)
    assert result.cell == "A1"
    assert result.to_dict() == {"sheet": "Inputs", "cell": "A1", "value": expected}


@pytest.mark.parametrize("value", [None, datetime(2024, 1, 1), complex(1), [], {}])
def test_no_implicit_coercion(value: object) -> None:
    with pytest.raises(TypeError):
        edit(value)


@pytest.mark.parametrize(
    "value", [float("nan"), float("inf"), 10**15, -(10**15), "\ud800", "a" * 32768, "\n" * 254]
)
def test_value_bounds(value: object) -> None:
    with pytest.raises(ValueError):
        edit(value)


def test_utf16_bound_counts_supplementary_characters() -> None:
    assert edit("😀" * 16383).value == "😀" * 16383
    with pytest.raises(ValueError):
        edit("😀" * 16384)


@pytest.mark.parametrize("cell", ["$A$1", "A0", "A1٢", "XFE1", "A1048577", "A1:B2"])
def test_address_validation(cell: str) -> None:
    with pytest.raises(ValueError):
        edit(1, cell)


def test_canonical_json_is_literal_sorted_ascii_and_signed_zero() -> None:
    assert importlib.util.find_spec("xlayer._canonical") is not None
    from xlayer._canonical import canonical_json, freeze_json

    source: dict[str, object] = {"z": ["😀", -0.0], "a": {"b": True}}
    frozen = freeze_json(source)
    source["z"] = []
    assert canonical_json(frozen) == b'{"a":{"b":true},"z":["\\ud83d\\ude00",-0.0]}'


def test_json_domain_refuses_cycles_nonstring_keys_and_nonfinite_values() -> None:
    assert importlib.util.find_spec("xlayer._canonical") is not None
    from xlayer._canonical import freeze_json

    cycle: list[object] = []
    cycle.append(cycle)
    for value in (cycle, {1: "bad"}, {"x": float("nan")}, {"x": date(2024, 1, 1)}):
        with pytest.raises((TypeError, ValueError)):
            freeze_json(value)


def test_preview_copies_and_freezes_nested_evidence() -> None:
    assert importlib.util.find_spec("xlayer._preview") is not None
    from xlayer._preview import Preview

    data: dict[str, object] = {"schema_version": "1.0", "edits": [{"raw": "1"}]}
    result = Preview(data)
    data["edits"] = []
    returned = result.to_dict()
    assert returned["edits"] == [{"raw": "1"}]
    returned["edits"] = []
    assert result.to_dict()["edits"] == [{"raw": "1"}]
    assert json.loads(result.canonical_json())["preview_digest"] == result.preview_digest
    with pytest.raises(TypeError):
        result.evidence["edits"] = []  # type: ignore[index]


def test_receipt_audit_metadata_does_not_change_evidence_digest() -> None:
    assert importlib.util.find_spec("xlayer._receipt") is not None
    from xlayer._receipt import Receipt

    a = Receipt({"schema_version": "1.0", "verified": True}, {"timestamp": "first"})
    b = Receipt({"verified": True, "schema_version": "1.0"}, {"timestamp": "second"})
    assert a.receipt_digest == b.receipt_digest
    assert a.canonical_json() != b.canonical_json()


def test_canonical_evidence_ceiling_is_enforced() -> None:
    assert importlib.util.find_spec("xlayer._canonical") is not None
    from xlayer._canonical import EvidenceTooLarge, canonical_json

    with pytest.raises(EvidenceTooLarge):
        canonical_json({"x": "a" * (16 * 1024 * 1024)})


def test_approval_claims_do_not_supply_authority() -> None:
    assert importlib.util.find_spec("xlayer._approval") is not None
    from xlayer._approval import Approval, VerifiedApproval, check_approval

    digest = "sha256:" + "a" * 64
    artifact = Approval("1.0", "id", "host", "human", digest, digest, digest, "now")
    denied = VerifiedApproval(artifact, False, ())
    assert (
        check_approval(artifact, denied, digest, digest, digest, ("apply_set_value",)) is not None
    )
    accepted = VerifiedApproval(artifact, True, ("apply_set_value",))
    assert check_approval(artifact, accepted, digest, digest, digest, ("apply_set_value",)) is None
    assert (
        check_approval(
            artifact, accepted, digest, digest, digest, ("apply_set_value", "change_populated_text")
        )
        is not None
    )


def test_model_cannot_hold_external_authority_grants() -> None:
    assert importlib.util.find_spec("xlayer._approval") is not None
    from xlayer._approval import Approval, VerifiedApproval, check_approval

    digest = "sha256:" + "a" * 64
    artifact = Approval("1.0", "id", "host", "model", digest, digest, digest, "now")
    decision = VerifiedApproval(artifact, True, ("apply_set_value", "change_populated_text"))
    assert (
        check_approval(
            artifact, decision, digest, digest, digest, ("apply_set_value", "change_populated_text")
        )
        is not None
    )
