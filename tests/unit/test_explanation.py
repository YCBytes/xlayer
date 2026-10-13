"""Saved XML facts, bounded evidence, disclosure and ownership through the root API."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import cast

import pytest

import xlayer
from tests.unit._dependency_cases import write_workbook_case
from tests.unit.test_public_inspections import source

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"
EDGES = FIXTURES / "test_workbook_7_write_v11_edges.xlsx"


def selected(address: str, field: str, sheet: str = "edges") -> dict[str, object]:
    return {"family": "cell", "sheet": sheet, "address": address, "field": field}


def fact(
    address: str, field: str, expected: object, *, sheet: str = "edges", ident: str = "f"
) -> dict[str, object]:
    return {
        "id": ident,
        "kind": "fact",
        "selector": selected(address, field, sheet),
        "expected": expected,
    }


def request(
    book: xlayer.Workbook, scope: xlayer.ExplanationScope, claims: list[dict[str, object]]
) -> xlayer.ExplanationRequest:
    value = xlayer.ExplanationRequest.from_dict(
        {
            "schema_version": "1.0",
            "explanation_contract_version": "1.0",
            "source_fingerprint": book.source_fingerprint,
            "scope_digest": scope.scope_digest,
            "claims": claims,
        }
    )
    assert isinstance(value, xlayer.ExplanationRequest), value
    return value


def obj(value: object) -> dict[str, object]:
    assert isinstance(value, dict)
    return cast("dict[str, object]", value)


def items(value: object) -> list[object]:
    assert isinstance(value, list)
    return value


def digest(value: object) -> str:
    return (
        "sha256:"
        + hashlib.sha256(
            json.dumps(
                value,
                sort_keys=True,
                separators=(",", ":"),
                ensure_ascii=True,
                allow_nan=False,
            ).encode()
        ).hexdigest()
    )


def expected_answer(path: Path, scope: xlayer.ExplanationScope, body: str, claims: int) -> str:
    """Hand-authored complete text oracle; never use the production renderer."""
    return (
        "Checked saved-workbook facts\n"
        f"Source: sha256:{hashlib.sha256(path.read_bytes()).hexdigest()}\n"
        f"Scope: {digest(scope.to_dict())}\n"
        + body
        + f"Accepted claims={claims}; required facts covered=0.\n"
        "Calculation, visual appearance, business correctness "
        "and edit approval were not established.\n"
        "This checks declared saved facts, not arbitrary model prose or whole-task completion.\n"
        "Decoded numbers are convenience values; raw text preserves saved digits.\n"
        "Saved format classification is not a converted date, domain type or edit permission.\n"
        "The report describes its loaded snapshot; current on-disk freshness was not checked.\n"
    )


def test_real_fixture_facts_report_and_witness_bindings_are_independent() -> None:
    original = EDGES.read_bytes()
    required = [selected("A1", "raw"), selected("C10", "presence")]
    scope = xlayer.ExplanationScope({"edges": ["A1", "C10", "D1"]}, required_facts=required)
    with xlayer.Workbook.open(EDGES) as book:
        claims: list[dict[str, object]] = [
            fact("A1", "raw", "100"),
            fact("C10", "presence", "missing", ident="missing"),
            fact("C10", "merge_role", "child", ident="merge"),
            fact("D1", "format_temporal_kind", "date", ident="style"),
        ]
        submitted = request(book, scope, claims)
        report = book.check_explanation(submitted, scope=scope)
        assert isinstance(report, xlayer.CheckedExplanation)
        repeated = book.check_explanation(submitted, scope=scope)
        assert isinstance(repeated, xlayer.CheckedExplanation)
        assert repeated.canonical_json() == report.canonical_json()
    data = report.to_dict()
    assert EDGES.read_bytes() == original
    assert data["source_fingerprint"] == "sha256:" + hashlib.sha256(original).hexdigest()
    assert data["required_coverage"] == required
    assert data["request_digest"] == digest(submitted.to_dict())
    assert data["evidence_digest"] == digest(data["observations"])
    assert data.pop("report_digest") == digest(data)
    assert report.report_digest == report.evidence["report_digest"]
    assert report.render_text() == report.evidence["rendered_text"]
    assert obj(data["observations"])["inventory"] is None
    paths = items(data["witnesses"])
    expected = [("A1", "100"), ("C10", "missing"), ("C10", "child"), ("D1", "date")]
    for witness, (address, value) in zip(paths, expected, strict=True):
        record: object = data
        for step in items(obj(witness)["path"]):
            record = (
                obj(record)[step] if isinstance(step, str) else items(record)[cast("int", step)]
            )
        assert obj(record)["address"] == address
        assert obj(witness)["observed_digest"] == digest(value)
    data["observations"] = {}
    assert report.to_dict()["observations"] != {}


def test_late_group_mismatch_refuses_everything() -> None:
    scope = xlayer.ExplanationScope({"edges": ["A1", "F1"]})
    with xlayer.Workbook.open(EDGES) as book:
        result = book.check_explanation(
            request(
                book,
                scope,
                [
                    {
                        "id": "g",
                        "kind": "all_equal",
                        "field": "stored_value",
                        "expected": 100,
                        "cells": [
                            {"sheet": "edges", "address": "A1"},
                            {"sheet": "edges", "address": "F1"},
                        ],
                    }
                ],
            ),
            scope=scope,
        )
        assert isinstance(result, xlayer.Refusal)
        assert result.code == "explanation_contradicted"
        assert "accepted_claims" not in result.to_dict()
        assert not book.closed


def test_successful_group_covers_each_required_selector() -> None:
    required = [selected("B10", "presence"), selected("C10", "presence")]
    scope = xlayer.ExplanationScope({"edges": ["B10", "C10"]}, required_facts=required)
    with xlayer.Workbook.open(EDGES) as book:
        result = book.check_explanation(
            request(
                book,
                scope,
                [
                    {
                        "id": "g",
                        "kind": "all_equal",
                        "field": "presence",
                        "expected": "missing",
                        "cells": [
                            {"sheet": "edges", "address": "B10"},
                            {"sheet": "edges", "address": "C10"},
                        ],
                    }
                ],
            ),
            scope=scope,
        )
    assert isinstance(result, xlayer.CheckedExplanation)
    assert result.to_dict()["required_coverage"] == required


def test_truth_does_not_replace_missing_required_facts() -> None:
    scope = xlayer.ExplanationScope({"edges": ["A1"]}, required_facts=[selected("A1", "raw")])
    with xlayer.Workbook.open(EDGES) as book:
        result = book.check_explanation(
            request(book, scope, [fact("A1", "stored_value", 100)]), scope=scope
        )
    assert isinstance(result, xlayer.Refusal)
    assert result.code == "explanation_requirements_unmet"


def no_read(*_args: object, **_kwargs: object) -> xlayer.Inspection:
    raise AssertionError("an unauthorized inspection occurred")


def test_all_claims_authorized_before_any_inspection(monkeypatch: pytest.MonkeyPatch) -> None:
    scope = xlayer.ExplanationScope({"edges": ["A1"]})
    with xlayer.Workbook.open(EDGES) as book:
        monkeypatch.setattr(xlayer.Workbook, "read_cells", no_read)
        result = book.check_explanation(
            request(
                book,
                scope,
                [fact("A1", "raw", "100"), fact("B1", "raw", "secret", ident="outside")],
            ),
            scope=scope,
        )
        assert isinstance(result, xlayer.Refusal)
        assert result.code == "explanation_scope_exceeded"
        assert "secret" not in json.dumps(result.to_dict())
        assert not book.closed


def test_cell_only_check_never_calls_inventory(monkeypatch: pytest.MonkeyPatch) -> None:
    scope = xlayer.ExplanationScope({"edges": ["A1"]})
    with xlayer.Workbook.open(EDGES) as book:
        monkeypatch.setattr(xlayer.Workbook, "list_sheets", no_read)
        result = book.check_explanation(
            request(book, scope, [fact("A1", "raw", "100")]), scope=scope
        )
    assert isinstance(result, xlayer.CheckedExplanation)


def test_missing_sheet_refusal_does_not_disclose_inventory() -> None:
    scope = xlayer.ExplanationScope({"missing": ["A1"]})
    with xlayer.Workbook.open(EDGES) as book:
        result = book.check_explanation(
            request(book, scope, [fact("A1", "raw", None, sheet="missing")]), scope=scope
        )
    assert isinstance(result, xlayer.Refusal)
    assert result.code == "sheet_not_found"
    assert "edges" not in json.dumps(result.to_dict())
    assert "available_names" not in json.dumps(result.to_dict())
    assert str(EDGES) not in json.dumps(result.to_dict())


def test_unselected_corrupt_cell_is_not_disclosed(tmp_path: Path) -> None:
    path = source(tmp_path, '<c r="A1"><v>1</v></c><c r="Z1"><v>secret-invalid-value</v></c>')
    scope = xlayer.ExplanationScope({"Inputs": ["A1"]})
    with xlayer.Workbook.open(path) as book:
        result = book.check_explanation(
            request(book, scope, [fact("A1", "raw", "1", sheet="Inputs")]), scope=scope
        )
    assert isinstance(result, xlayer.Refusal)
    assert result.code == "invalid_part_content"
    assert "Z1" not in json.dumps(result.to_dict())
    assert "secret-invalid-value" not in json.dumps(result.to_dict())


def test_binding_and_reduced_limits_precede_reads(monkeypatch: pytest.MonkeyPatch) -> None:
    scope = xlayer.ExplanationScope({"edges": ["A1", "B1"]})
    with xlayer.Workbook.open(EDGES) as book:
        submitted = request(book, scope, [fact("A1", "raw", "100")])
        monkeypatch.setattr(xlayer.Workbook, "read_cells", no_read)
        reduced = book.check_explanation(
            submitted, scope=scope, limits=xlayer.ExplanationLimits(max_cells=1)
        )
        assert isinstance(reduced, xlayer.Refusal)
        assert reduced.code == "explanation_limit_exceeded"
        changed = book.check_explanation(
            submitted, scope=xlayer.ExplanationScope({"edges": ["A1"]})
        )
        assert isinstance(changed, xlayer.Refusal)
        assert changed.code == "explanation_binding_mismatch"


def test_numeric_comparison_excludes_bool_and_does_not_round_integer(tmp_path: Path) -> None:
    path = source(tmp_path, '<c r="A1"><v>1</v></c><c r="B1"><v>9007199254740992</v></c>')
    scope = xlayer.ExplanationScope({"Inputs": ["A1", "B1"]})
    with xlayer.Workbook.open(path) as book:
        for address, value in [("A1", True), ("B1", 9007199254740993)]:
            result = book.check_explanation(
                request(book, scope, [fact(address, "stored_value", value, sheet="Inputs")]),
                scope=scope,
            )
            assert isinstance(result, xlayer.Refusal)
            assert result.code == "explanation_contradicted"


def test_inventory_and_dependency_witnesses_match_published_records() -> None:
    path = FIXTURES / "test_workbook_4_worksheet_final.xlsx"
    scope = xlayer.ExplanationScope({}, include_inventory=True, impact_root=("formulas", "E8"))
    claims: list[dict[str, object]] = [
        {
            "id": "count",
            "kind": "fact",
            "selector": {"family": "inventory", "field": "sheet_count"},
            "expected": 13,
        },
        {
            "id": "coverage",
            "kind": "fact",
            "selector": {"family": "impact", "field": "analysis_status"},
            "expected": "partial",
        },
        {
            "id": "member",
            "kind": "known_dependent",
            "target": {"sheet": "formulas", "address": "B8"},
            "relation": "direct",
        },
    ]
    with xlayer.Workbook.open(path) as book:
        report = book.check_explanation(request(book, scope, claims), scope=scope)
        assert isinstance(report, xlayer.CheckedExplanation)
        claims[-1]["target"] = {"sheet": "formulas", "address": "XFD1048576"}
        absent = book.check_explanation(request(book, scope, claims), scope=scope)
    assert isinstance(absent, xlayer.Refusal)
    assert absent.code == "explanation_evidence_missing"
    data = report.to_dict()
    witness = obj(items(data["witnesses"])[2])
    record: object = data
    for step in items(witness["path"]):
        record = obj(record)[step] if isinstance(step, str) else items(record)[cast("int", step)]
    assert record == {"sheet": "formulas", "address": "B8"}
    assert witness["observed_digest"] == digest(record)
    assert "partial" in report.render_text()
    assert "lower bounds" in report.render_text()


def test_report_bound_includes_rendering_and_final_digest() -> None:
    scope = xlayer.ExplanationScope({"edges": ["A1"]})
    with xlayer.Workbook.open(EDGES) as book:
        submitted = request(book, scope, [fact("A1", "raw", "100")])
        report = book.check_explanation(submitted, scope=scope)
        assert isinstance(report, xlayer.CheckedExplanation)
        result = book.check_explanation(
            submitted, scope=scope, limits=xlayer.ExplanationLimits(max_report_bytes=100)
        )
        assert isinstance(result, xlayer.Refusal)
        assert result.code == "explanation_limit_exceeded"
        assert not book.closed


def test_loaded_snapshot_survives_source_replacement_and_report_survives_close(
    tmp_path: Path,
) -> None:
    path = tmp_path / "snapshot.xlsx"
    path.write_bytes(EDGES.read_bytes())
    scope = xlayer.ExplanationScope({"edges": ["A1"]})
    book = xlayer.Workbook.open(path)
    submitted = request(book, scope, [fact("A1", "raw", "100")])
    path.write_bytes(b"replaced")
    report = book.check_explanation(submitted, scope=scope)
    assert isinstance(report, xlayer.CheckedExplanation)
    book.close()
    assert report.render_text()
    with pytest.raises(xlayer.ClosedWorkbookError):
        book.check_explanation(submitted, scope=scope)


def test_unexpected_capture_error_closes_but_native_error_does_not(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    scope = xlayer.ExplanationScope({"edges": ["A1"]})
    book = xlayer.Workbook.open(EDGES)
    submitted = request(book, scope, [fact("A1", "raw", "100")])
    with pytest.raises(TypeError):
        book.check_explanation({}, scope=scope)  # type: ignore[arg-type]
    assert not book.closed
    monkeypatch.setattr(xlayer.Workbook, "read_cells", no_read)
    with pytest.raises(AssertionError, match="unauthorized inspection"):
        book.check_explanation(submitted, scope=scope)
    assert book.closed


def test_capture_order_uses_tabs_then_requested_addresses(tmp_path: Path) -> None:
    path = write_workbook_case(
        tmp_path / "ordered.xlsx",
        {
            "Zed": '<sheetData><row><c r="A1"><v>1</v></c></row></sheetData>',
            "Alpha": '<sheetData><row><c r="A1"><v>2</v></c></row></sheetData>',
        },
    )
    scope = xlayer.ExplanationScope({"Alpha": ["B1", "A1"], "Zed": ["A1"]})
    with xlayer.Workbook.open(path) as book:
        result = book.check_explanation(
            request(book, scope, [fact("A1", "raw", "1", sheet="Zed")]), scope=scope
        )
    assert isinstance(result, xlayer.CheckedExplanation)
    observed = items(obj(result.to_dict()["observations"])["cells"])
    assert [obj(obj(v)["data"])["sheet"] for v in observed] == ["Zed", "Alpha"]
    assert obj(obj(observed[1])["data"])["requested_addresses"] == ["B1", "A1"]


def test_full_missing_cell_render_is_honest() -> None:
    scope = xlayer.ExplanationScope({"edges": ["C10"]})
    with xlayer.Workbook.open(EDGES) as book:
        report = book.check_explanation(
            request(book, scope, [fact("C10", "presence", "missing")]), scope=scope
        )
        assert isinstance(report, xlayer.CheckedExplanation)
        expected = (
            "Checked saved-workbook facts\n"
            f"Source: {book.source_fingerprint}\nScope: {scope.scope_digest}\n"
            '"edges"!"C10": physically missing; no cell is stored here; '
            "not a saved zero or a present blank.\n"
            '  Merge ref="B10:C10"; root="B10"; merge role="child"; '
            "no root value is copied to a missing child.\n"
            "Accepted claims=1; required facts covered=0.\n"
            "Calculation, visual appearance, business correctness "
            "and edit approval were not established.\n"
            "This checks declared saved facts, not arbitrary model prose "
            "or whole-task completion.\n"
            "Decoded numbers are convenience values; raw text preserves saved digits.\n"
            "Saved format classification is not a converted date, domain type or edit permission.\n"
            "The report describes its loaded snapshot; current on-disk freshness was not checked.\n"
        )
        assert report.render_text() == expected


def test_all_saved_cache_states_and_shared_followers_remain_explicit(tmp_path: Path) -> None:
    path = source(
        tmp_path,
        '<c r="A1"/><c r="B1" t="inlineStr"><is><t/></is></c>'
        '<c r="C1"><f>1</f></c><c r="D1"><f>1</f><v/></c>'
        '<c r="E1"><f>0</f><v>0</v></c><c r="F1" t="b"><f>TRUE()</f></c>'
        '<c r="G1" t="str"><f>""</f><v/></c>'
        '<c r="H1" t="e"><f>1/0</f><v>#DIV/0!</v></c>'
        '<c r="I1"><f t="shared" si="7" ref="I1:J1">A1*2</f><v>0</v></c>'
        '<c r="J1"><f t="shared" si="7"/><v>0</v></c>',
    )
    addresses = [f"{letter}1" for letter in "ABCDEFGHIJK"]
    scope = xlayer.ExplanationScope({"Inputs": addresses})
    raw = [None, None, None, "", "0", None, "", "#DIV/0!", "0", "0", None]
    claims = [
        fact(address, "raw", value, sheet="Inputs", ident=f"raw{i}")
        for i, (address, value) in enumerate(zip(addresses, raw, strict=True))
    ]
    claims += [
        fact("J1", "formula_text", None, sheet="Inputs", ident="own"),
        fact("J1", "has_formula", True, sheet="Inputs", ident="has"),
        fact("J1", "shared_master_text", "A1*2", sheet="Inputs", ident="master"),
    ]
    with xlayer.Workbook.open(path) as book:
        report = book.check_explanation(request(book, scope, claims), scope=scope)
    assert isinstance(report, xlayer.CheckedExplanation)
    data = report.to_dict()
    text = report.render_text()
    for phrase in [
        "present blank literal",
        "stored literal empty text",
        "absent cache element",
        "empty cache element",
        "Saved empty-text cache",
        "physically missing",
        "follower expression not expanded",
        'raw="0"; decoded=0.0',
        'raw="#DIV/0!"; decoded="#DIV/0!"',
    ]:
        assert phrase in text
    assert text.count("unverified saved formula cache") == 8
    # These answers come from the XML above. Full-text equality also guards
    # omissions/extra claims, which phrase checks alone cannot catch.
    general = (
        '  Saved format id=0; code="General"; temporal kind="non_temporal"; '
        'basis="cell_xf_number_format".\n'
    )
    blocks = [
        '"Inputs"!"A1": present; present blank literal; raw=null; decoded=null; '
        'stored type="n"; value kind="blank".\n',
        '"Inputs"!"B1": present; stored literal empty text; raw=null; decoded=""; '
        'stored type="inlineStr"; value kind="string".\n',
        '"Inputs"!"C1": present; unverified saved formula cache; raw=null; decoded=null; '
        'stored type="n"; value kind="blank".\n'
        '  Formula kind="normal"; own formula text="1"; result status="not_verified"; '
        "not recalculated or verified.\n"
        "  absent cache element; formula remains a formula, not an editable blank input.\n",
        '"Inputs"!"D1": present; unverified saved formula cache; raw=""; decoded=null; '
        'stored type="n"; value kind="blank".\n'
        '  Formula kind="normal"; own formula text="1"; result status="not_verified"; '
        "not recalculated or verified.\n"
        "  empty cache element; formula remains a formula, not an editable blank input.\n",
        '"Inputs"!"E1": present; unverified saved formula cache; raw="0"; decoded=0.0; '
        'stored type="n"; value kind="number".\n'
        '  Formula kind="normal"; own formula text="0"; result status="not_verified"; '
        "not recalculated or verified.\n",
        '"Inputs"!"F1": present; unverified saved formula cache; raw=null; decoded=null; '
        'stored type="b"; value kind="blank".\n'
        '  Formula kind="normal"; own formula text="TRUE()"; result status="not_verified"; '
        "not recalculated or verified.\n"
        "  absent cache element; formula remains a formula, not an editable blank input.\n",
        '"Inputs"!"G1": present; unverified saved formula cache; raw=""; decoded=""; '
        'stored type="str"; value kind="string".\n'
        '  Formula kind="normal"; own formula text="\\"\\""; result status="not_verified"; '
        "not recalculated or verified.\n"
        "  Saved empty-text cache; not an absent cache or numeric zero.\n",
        '"Inputs"!"H1": present; unverified saved formula cache; raw="#DIV/0!"; decoded="#DIV/0!"; '
        'stored type="e"; value kind="error".\n'
        '  Formula kind="normal"; own formula text="1/0"; result status="not_verified"; '
        "not recalculated or verified.\n",
        '"Inputs"!"I1": present; unverified saved formula cache; raw="0"; decoded=0.0; '
        'stored type="n"; value kind="number".\n'
        '  Formula kind="shared"; own formula text="A1*2"; result status="not_verified"; '
        "not recalculated or verified.\n"
        '  Shared master address="I1"; master ref="I1:J1"; master text="A1*2"; '
        "follower expression not expanded.\n",
        '"Inputs"!"J1": present; unverified saved formula cache; raw="0"; decoded=0.0; '
        'stored type="n"; value kind="number".\n'
        '  Formula kind="shared"; own formula text=null; result status="not_verified"; '
        "not recalculated or verified.\n"
        '  Shared master address="I1"; master ref="I1:J1"; master text="A1*2"; '
        "follower expression not expanded.\n",
    ]
    expected_body = "".join(block + general for block in blocks) + (
        '"Inputs"!"K1": physically missing; no cell is stored here; '
        "not a saved zero or a present blank.\n"
    )
    assert text == expected_answer(path, scope, expected_body, 14)
    assert obj(data["limitations"]) == {
        "calculation": "not_evaluated",
        "formula_cache_freshness": "not_verified",
        "business_correctness": "not_evaluated",
        "visual_appearance": "not_evaluated",
        "edit_authority": "not_granted",
        "arbitrary_model_prose": "not_verified",
        "whole_task": "not_certified",
        "current_file_freshness": "not_checked",
    }


def test_complete_analysis_sample_absence_still_is_not_negative_proof(tmp_path: Path) -> None:
    xml = (
        '<sheetData><row><c r="A1"><v>1</v></c>'
        + "".join(f'<c r="{letter}1"><f>A1*2</f><v>2</v></c>' for letter in "BCDEFGHIJKLM")
        + "</row></sheetData>"
    )
    path = write_workbook_case(tmp_path / "sample.xlsx", {"Inputs": xml})
    required = {"family": "impact", "field": "analysis_status"}
    scope = xlayer.ExplanationScope({}, impact_root=("Inputs", "A1"), required_facts=[required])
    member: dict[str, object] = {
        "id": "member",
        "kind": "known_dependent",
        "relation": "direct",
        "target": {"sheet": "Inputs", "address": "B1"},
    }
    with xlayer.Workbook.open(path) as book:
        missing = book.check_explanation(request(book, scope, [member]), scope=scope)
        assert isinstance(missing, xlayer.Refusal)
        assert missing.code == "explanation_requirements_unmet"
        claims: list[dict[str, object]] = [
            member,
            {"id": "coverage", "kind": "fact", "selector": required, "expected": "complete"},
        ]
        report = book.check_explanation(request(book, scope, claims), scope=scope)
        assert isinstance(report, xlayer.CheckedExplanation)
        member["target"] = {"sheet": "Inputs", "address": "M1"}
        refused = book.check_explanation(request(book, scope, claims), scope=scope)
        assert isinstance(refused, xlayer.Refusal)
        assert refused.code == "explanation_evidence_missing"
        assert "sample truncated=true" in report.render_text()


def test_unknown_dependency_counts_are_null_and_not_zero(tmp_path: Path) -> None:
    path = source(tmp_path, '<c r="A1"><v>nan</v></c>')
    scope = xlayer.ExplanationScope({}, impact_root=("Inputs", "A1"))
    claims: list[dict[str, object]] = [
        {
            "id": "status",
            "kind": "fact",
            "selector": {"family": "impact", "field": "analysis_status"},
            "expected": "unknown",
        },
        {
            "id": "count",
            "kind": "fact",
            "selector": {"family": "impact", "field": "known_direct_count"},
            "expected": None,
        },
    ]
    with xlayer.Workbook.open(path) as book:
        report = book.check_explanation(request(book, scope, claims), scope=scope)
        assert isinstance(report, xlayer.CheckedExplanation)
    assert "known direct=null" in report.render_text()


@pytest.mark.parametrize("status", ["complete", "partial", "unknown"])
def test_full_impact_answer_preserves_coverage_and_numerical_limits(
    tmp_path: Path, status: str
) -> None:
    cells = '<c r="A1"><v>1</v></c><c r="B1"><f>A1</f><v>1</v></c>'
    if status == "partial":
        cells += '<c r="C1"><f>INDIRECT("A1")</f><v>1</v></c>'
    elif status == "unknown":
        cells = '<c r="A1"><v>nan</v></c>'
    path = source(tmp_path, cells)
    scope = xlayer.ExplanationScope({}, impact_root=("Inputs", "A1"))
    claims: list[dict[str, object]] = [
        {
            "id": "coverage",
            "kind": "fact",
            "selector": {"family": "impact", "field": "analysis_status"},
            "expected": status,
        }
    ]
    with xlayer.Workbook.open(path) as book:
        report = book.check_explanation(request(book, scope, claims), scope=scope)
    assert isinstance(report, xlayer.CheckedExplanation)
    direct = union = "null" if status == "unknown" else "1"
    transitive = "null" if status == "unknown" else "0"
    exact = "1" if status == "complete" else "null"
    direct_sample = "[]" if status == "unknown" else '[{"address":"B1","sheet":"Inputs"}]'
    expected = (
        'Structural impact for {"address":"A1","sheet":"Inputs"}: '
        f"{status}; known direct={direct}; known transitive={transitive}; "
        f"known union={union}; exact total={exact}.\n"
        '  scope="worksheet_cell_formulas"; numerical_effects_evaluated=false.\n'
        "  Partial known counts are lower bounds; null means unknown/not evaluated, never zero. "
        "These are potential structural references, not predictions of numerical effects.\n"
        f"  direct sample truncated=false; published sample={direct_sample}.\n"
        "  transitive sample truncated=false; published sample=[].\n"
        "  Sample absence is not proof of non-dependence; reference paths and analysis limits "
        "remain in the structured evidence.\n"
    )
    assert report.render_text() == expected_answer(path, scope, expected, 1)


def test_complete_report_and_evidence_size_boundaries_are_exact() -> None:
    scope = xlayer.ExplanationScope({"edges": ["A1"]})
    with xlayer.Workbook.open(EDGES) as book:
        submitted = request(book, scope, [fact("A1", "raw", "100")])
        first = book.check_explanation(submitted, scope=scope)
        assert isinstance(first, xlayer.CheckedExplanation)
        evidence_size = len(
            json.dumps(
                first.to_dict()["observations"],
                ensure_ascii=True,
                sort_keys=True,
                separators=(",", ":"),
            ).encode()
        )
        good = book.check_explanation(
            submitted,
            scope=scope,
            limits=xlayer.ExplanationLimits(max_evidence_bytes=evidence_size),
        )
        assert isinstance(good, xlayer.CheckedExplanation)
        bad = book.check_explanation(
            submitted,
            scope=scope,
            limits=xlayer.ExplanationLimits(max_evidence_bytes=evidence_size - 1),
        )
        assert isinstance(bad, xlayer.Refusal)
        # Changing max_report_bytes changes one number in the report itself.
        size = len(first.canonical_json())
        body = first.to_dict()
        obj(obj(body["limits"])["explanation"])["max_report_bytes"] = size
        size = len(json.dumps(body, sort_keys=True, separators=(",", ":")).encode())
        exact = book.check_explanation(
            submitted, scope=scope, limits=xlayer.ExplanationLimits(max_report_bytes=size)
        )
        assert isinstance(exact, xlayer.CheckedExplanation)
        assert len(exact.canonical_json()) == size
        over = book.check_explanation(
            submitted, scope=scope, limits=xlayer.ExplanationLimits(max_report_bytes=size - 1)
        )
        assert isinstance(over, xlayer.Refusal)


def test_diagnostic_examples_are_capped_and_tiny_details_still_refuse() -> None:
    scope = xlayer.ExplanationScope({"edges": ["A1"]})
    with xlayer.Workbook.open(EDGES) as book:
        submitted = request(
            book, scope, [fact("A1", "raw", "wrong", ident=f"f{i}") for i in range(5)]
        )
        result = book.check_explanation(
            submitted, scope=scope, limits=xlayer.ExplanationLimits(max_diagnostics=2)
        )
        assert isinstance(result, xlayer.Refusal)
        details = obj(result.to_dict()["details"])
        assert len(items(details["diagnostics"])) == 2
        assert details["total_known"] == 5
        assert details["diagnostics_truncated"] is True
        tiny = book.check_explanation(
            submitted, scope=scope, limits=xlayer.ExplanationLimits(max_diagnostic_bytes=2)
        )
        assert isinstance(tiny, xlayer.Refusal)
        assert dict(tiny.details) == {}


@pytest.mark.parametrize(
    "claim_kind", ["counterfactual", "business_conclusion", "formula_equivalence"]
)
def test_even_true_unsupported_claims_never_reach_workbook(claim_kind: str) -> None:
    result = xlayer.ExplanationRequest.from_dict(
        {
            "schema_version": "1.0",
            "explanation_contract_version": "1.0",
            "source_fingerprint": "sha256:" + "0" * 64,
            "scope_digest": "sha256:" + "1" * 64,
            "claims": [{"id": "f", "kind": claim_kind, "expected": True}],
        }
    )
    assert isinstance(result, xlayer.Refusal)
    assert result.code == "unsupported_explanation_claim"
