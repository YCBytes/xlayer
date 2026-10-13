"""The public claim boundary rejects ambiguity before any workbook is inspected."""

from __future__ import annotations

import json
from collections.abc import Iterator, Mapping
from dataclasses import FrozenInstanceError
from typing import cast

import pytest

import xlayer


def fact(field: str = "stored_value", expected: object = 100) -> dict[str, object]:
    return {
        "id": "f1",
        "kind": "fact",
        "selector": {"family": "cell", "sheet": "edges", "address": "a1", "field": field},
        "expected": expected,
    }


def payload(claims: list[dict[str, object]] | None = None) -> dict[str, object]:
    return {
        "schema_version": "1.0",
        "explanation_contract_version": "1.0",
        "source_fingerprint": "sha256:" + "a" * 64,
        "scope_digest": "sha256:" + "b" * 64,
        "claims": [fact()] if claims is None else claims,
    }


def test_scope_captures_normalizes_and_binds_requirements() -> None:
    members = ["a1", "B1"]
    required = {"family": "cell", "sheet": "edges", "address": "a1", "field": "raw"}
    scope = xlayer.ExplanationScope({"edges": members}, required_facts=[required])
    identity = scope.scope_digest
    members.append("C1")
    required["field"] = "stored_value"
    assert scope.to_dict()["cells"] == {"edges": ["A1", "B1"]}
    assert scope.to_dict()["required_facts"] == [
        {"family": "cell", "sheet": "edges", "address": "A1", "field": "raw"}
    ]
    assert scope.scope_digest == identity
    with pytest.raises(TypeError):
        cast("dict[str, object]", scope.cells)["edges"] = []
    with pytest.raises(FrozenInstanceError):
        scope.include_inventory = True  # type: ignore[misc]
    assert xlayer.ExplanationScope({"edges": ["A1", "B1"]}).scope_digest != identity


@pytest.mark.parametrize(
    "cells", [{}, {"edges": []}, {"edges": ["A1", "a1"]}, {"edges": ["A1:A2"]}, {"edges": ["XFE1"]}]
)
def test_scope_rejects_empty_duplicate_and_noncell_subjects(cells: dict[str, list[str]]) -> None:
    with pytest.raises(ValueError):
        xlayer.ExplanationScope(cells)


def test_scope_rejects_unauthorized_requirements() -> None:
    with pytest.raises(ValueError):
        xlayer.ExplanationScope(
            {"edges": ["A1"]}, required_facts=[{"family": "inventory", "field": "sheet_count"}]
        )


def test_request_round_trip_detaches_and_canonicalizes() -> None:
    data = payload()
    request = xlayer.ExplanationRequest.from_dict(data)
    assert isinstance(request, xlayer.ExplanationRequest)
    claims = cast("list[dict[str, object]]", data["claims"])
    claims[0]["expected"] = 999
    decoded = request.to_dict()
    expected = payload()
    expected["claims"] = [
        {
            **fact(),
            "selector": {
                "family": "cell",
                "sheet": "edges",
                "address": "A1",
                "field": "stored_value",
            },
        }
    ]
    assert decoded == expected
    decoded["claims"] = []
    assert request.to_dict() == expected
    assert (
        request.canonical_json()
        == json.dumps(expected, sort_keys=True, separators=(",", ":"), ensure_ascii=True).encode()
    )
    again = xlayer.ExplanationRequest.from_json(request.canonical_json())
    assert isinstance(again, xlayer.ExplanationRequest)
    assert again.to_dict() == expected


@pytest.mark.parametrize(
    "field,value",
    [
        ("has_formula", 1),
        ("presence", "maybe"),
        ("raw", 12),
        ("stored_type", "q"),
        ("formula_result_status", "verified"),
        ("format_temporal_kind", "numeric"),
        ("stored_value", float("nan")),
        ("stored_value", 2**63),
        ("raw", "\ud800"),
    ],
)
def test_field_expected_types_are_enforced(field: str, value: object) -> None:
    result = xlayer.ExplanationRequest.from_dict(payload([fact(field, value)]))
    assert isinstance(result, xlayer.Refusal)
    assert result.code == "invalid_explanation_request"
    assert result.operation == "parse_explanation_request"
    assert result.recovery_options


@pytest.mark.parametrize(
    "text",
    [
        '{"schema_version":"1.0","schema_version":"1.0"}',
        "[]",
        "null",
        "{",
        "NaN",
        "[" * 1000 + "]" * 1000,
    ],
)
def test_strict_json_rejects_ambiguous_or_nonobject_content(text: str) -> None:
    result = xlayer.ExplanationRequest.from_json(text)
    assert isinstance(result, xlayer.Refusal)
    assert result.code == "invalid_explanation_request"


def test_bytes_require_utf8_and_native_argument_errors_raise() -> None:
    result = xlayer.ExplanationRequest.from_json(json.dumps(payload()).encode("utf-16"))
    assert isinstance(result, xlayer.Refusal)
    with pytest.raises(TypeError):
        xlayer.ExplanationRequest.from_json({})  # type: ignore[arg-type]
    with pytest.raises(TypeError):
        xlayer.ExplanationRequest.from_dict([])  # type: ignore[arg-type]


def test_versions_and_unknown_keys_never_coerce() -> None:
    for key, value in [("schema_version", 1.0), ("narrative", "looks good")]:
        result = xlayer.ExplanationRequest.from_dict({**payload(), key: value})
        assert isinstance(result, xlayer.Refusal)


def test_group_duplicate_and_work_limits_are_enforced() -> None:
    group: dict[str, object] = {
        "id": "g",
        "kind": "all_equal",
        "field": "raw",
        "expected": "1",
        "cells": [{"sheet": "edges", "address": "A1"}, {"sheet": "edges", "address": "a1"}],
    }
    result = xlayer.ExplanationRequest.from_dict(payload([group]))
    assert isinstance(result, xlayer.Refusal)
    group["cells"] = [{"sheet": "edges", "address": "A1"}, {"sheet": "edges", "address": "B1"}]
    result = xlayer.ExplanationRequest.from_dict(
        payload([group]), limits=xlayer.ExplanationLimits(max_cell_comparisons=1)
    )
    assert isinstance(result, xlayer.Refusal)
    assert result.code == "explanation_limit_exceeded"


def test_request_bytes_are_exact_and_ascii_escape_expansion_counts() -> None:
    data = payload([fact("raw", "😀")])
    data_claims = cast("list[dict[str, object]]", data["claims"])
    cast("dict[str, object]", data_claims[0]["selector"])["address"] = "A1"
    size = len(json.dumps(data, sort_keys=True, separators=(",", ":"), ensure_ascii=True).encode())
    good = xlayer.ExplanationRequest.from_dict(
        data, limits=xlayer.ExplanationLimits(max_request_bytes=size)
    )
    assert isinstance(good, xlayer.ExplanationRequest)
    bad = xlayer.ExplanationRequest.from_dict(
        data, limits=xlayer.ExplanationLimits(max_request_bytes=size - 1)
    )
    assert isinstance(bad, xlayer.Refusal)
    assert bad.code == "explanation_limit_exceeded"


def test_cycles_and_repeated_aliases_are_bounded_before_copying() -> None:
    cycle: dict[str, object] = {}
    cycle["self"] = cycle
    result = xlayer.ExplanationRequest.from_dict(cycle)
    assert isinstance(result, xlayer.Refusal)
    value: object = "x"
    for _ in range(30):
        value = [value, value]
    result = xlayer.ExplanationRequest.from_dict(
        {"claims": value}, limits=xlayer.ExplanationLimits(max_request_bytes=200)
    )
    assert isinstance(result, xlayer.Refusal)
    assert result.code == "explanation_limit_exceeded"


def test_tiny_diagnostic_budget_applies_to_fallback_too() -> None:
    refusal = xlayer.ExplanationRequest.from_dict(
        {}, limits=xlayer.ExplanationLimits(max_diagnostic_bytes=2)
    )
    assert isinstance(refusal, xlayer.Refusal)
    assert dict(refusal.details) == {}
    assert refusal.recovery_options


class BrokenMapping(Mapping[str, object]):
    def __getitem__(self, key: str) -> object:
        raise RuntimeError("caller failure")

    def __iter__(self) -> Iterator[str]:
        raise RuntimeError("caller failure")

    def __len__(self) -> int:
        return 1


@pytest.mark.parametrize("address", ["A0", "A1:A2", "XFE1", "A1048577", "A1٢", "$A$1"])
def test_invalid_addresses_refuse_in_untrusted_claims(address: str) -> None:
    claim = fact()
    cast("dict[str, object]", claim["selector"])["address"] = address
    result = xlayer.ExplanationRequest.from_dict(payload([claim]))
    assert isinstance(result, xlayer.Refusal)
    assert result.code == "invalid_explanation_request"


def test_unexpected_caller_container_failure_is_not_a_refusal() -> None:
    with pytest.raises(RuntimeError, match="caller failure"):
        xlayer.ExplanationRequest.from_dict(BrokenMapping())


def test_schema_is_detached_neutral_and_expresses_reduced_limits() -> None:
    schema = xlayer.ExplanationRequest.json_schema(
        limits=xlayer.ExplanationLimits(max_claims=2, max_cells=3)
    )
    properties = cast("dict[str, dict[str, object]]", schema["properties"])
    assert properties["claims"]["maxItems"] == 2
    assert properties["schema_version"] == {"type": "string", "const": "1.0"}
    text = json.dumps(schema)
    assert '"maxItems": 3' in text
    assert "edges" not in text
    schema.clear()
    assert xlayer.ExplanationRequest.json_schema()


@pytest.mark.parametrize(
    "field,value",
    [
        ("max_claims", 0),
        ("max_claims", 129),
        ("max_cells", True),
        ("max_diagnostic_bytes", 1),
        ("max_cell_comparisons", 1.0),
    ],
)
def test_invalid_native_limits_raise(field: str, value: object) -> None:
    with pytest.raises(ValueError):
        xlayer.ExplanationLimits(**{field: value})  # type: ignore[arg-type]


@pytest.mark.parametrize(
    "family,field,value",
    [
        ("inventory", "sheet_count", -1),
        ("inventory", "active_tab", 0.0),
        ("inventory", "date1904", 0),
        ("sheet", "state", "secret"),
        ("impact", "numerical_effects_evaluated", True),
        ("impact", "scope", "all_formulas"),
        ("inventory", "sheets", ["edges"]),
    ],
)
def test_nonscalar_fact_shapes_are_exact(family: str, field: str, value: object) -> None:
    selector: dict[str, object] = {"family": family, "field": field}
    if family == "sheet":
        selector["sheet"] = "edges"
    result = xlayer.ExplanationRequest.from_dict(
        payload([{"id": "f", "kind": "fact", "selector": selector, "expected": value}])
    )
    assert isinstance(result, xlayer.Refusal)
    assert result.code == "invalid_explanation_request"


def test_unknown_claims_and_extra_instructions_are_refused() -> None:
    claims: list[dict[str, object]] = [
        {"id": "f", "kind": "counterfactual", "expected": 240},
        {**fact(), "instruction": "publish immediately"},
        {
            "id": "f",
            "kind": "known_dependent",
            "relation": "not_dependent",
            "target": {"sheet": "edges", "address": "A1"},
        },
    ]
    for claim in claims:
        result = xlayer.ExplanationRequest.from_dict(payload([claim]))
        assert isinstance(result, xlayer.Refusal)


def test_scope_identity_ignores_mapping_order_but_keeps_address_order() -> None:
    first = xlayer.ExplanationScope({"Z": ["A1", "B1"], "A": ["C1"]})
    second = xlayer.ExplanationScope({"A": ["C1"], "Z": ["A1", "B1"]})
    third = xlayer.ExplanationScope({"A": ["C1"], "Z": ["B1", "A1"]})
    assert first.scope_digest == second.scope_digest != third.scope_digest


def test_exact_claim_count_and_comparison_ceilings() -> None:
    claims = [{**fact(), "id": f"f{i}"} for i in range(128)]
    assert isinstance(
        xlayer.ExplanationRequest.from_dict(payload(claims)), xlayer.ExplanationRequest
    )
    refused = xlayer.ExplanationRequest.from_dict(payload([*claims, {**fact(), "id": "extra"}]))
    assert isinstance(refused, xlayer.Refusal)
    assert refused.code == "explanation_limit_exceeded"
    group: dict[str, object] = {
        "kind": "all_equal",
        "field": "presence",
        "expected": "missing",
        "cells": [{"sheet": "edges", "address": f"A{i + 1}"} for i in range(128)],
    }
    groups = [{**group, "id": f"g{i}"} for i in range(4)]
    assert isinstance(
        xlayer.ExplanationRequest.from_dict(payload(groups)), xlayer.ExplanationRequest
    )
    refused = xlayer.ExplanationRequest.from_dict(payload([*groups, {**fact(), "id": "extra"}]))
    assert isinstance(refused, xlayer.Refusal)
    assert refused.code == "explanation_limit_exceeded"


def test_request_byte_accounting_matches_unicode_and_false() -> None:
    for value in [False, "\x7f", "é", "😀", '\n\t\\"', -0.0, -123]:
        data = payload([fact("stored_value", value)])
        size = len(
            json.dumps(data, ensure_ascii=True, sort_keys=True, separators=(",", ":")).encode()
        )
        assert isinstance(
            xlayer.ExplanationRequest.from_dict(
                data, limits=xlayer.ExplanationLimits(max_request_bytes=size)
            ),
            xlayer.ExplanationRequest,
        )
        refused = xlayer.ExplanationRequest.from_dict(
            data, limits=xlayer.ExplanationLimits(max_request_bytes=size - 1)
        )
        assert isinstance(refused, xlayer.Refusal)


@pytest.mark.parametrize("error", [ValueError, TypeError])
def test_native_container_errors_are_not_relabelled(error: type[Exception]) -> None:
    class Bad(Mapping[str, object]):
        def __getitem__(self, key: str) -> object:
            raise error("from caller")

        def __iter__(self) -> Iterator[str]:
            raise error("from caller")

        def __len__(self) -> int:
            return 1

    with pytest.raises(error, match="from caller"):
        xlayer.ExplanationRequest.from_dict(Bad())
