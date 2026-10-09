"""Bounded public presentation must not improve or rescan retained dependency proof."""

from __future__ import annotations

from dataclasses import Field, fields, replace
from pathlib import Path
from typing import cast

import pytest

from tests.unit._dependency_cases import write_workbook_case
from tests.unit.test_public_inspections import data, source
from xlayer import ClosedWorkbookError, ImpactLimits, Inspection, ReadLimits, Refusal, Workbook
from xlayer._batch_impact import BatchImpact, RootImpact
from xlayer._canonical import thaw_json
from xlayer._dependencies import AnalysisWork, DependencyImpact, _to_json, analyse_dependencies
from xlayer._impact_summary import summarize_batch
from xlayer._workbook import Workbook as EngineWorkbook


def summary(result: Inspection | Refusal) -> dict[str, object]:
    value = data(result)["impact_summary"]
    assert isinstance(value, dict)
    return cast("dict[str, object]", value)


def root(result: Inspection | Refusal) -> dict[str, object]:
    roots = summary(result)["roots"]
    assert isinstance(roots, list) and len(roots) == 1
    return cast("dict[str, object]", roots[0])


def test_complete_summary_is_same_pure_projection_of_single_query(tmp_path: Path) -> None:
    path = source(
        tmp_path,
        '<c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c><c r="C1"><f>B1</f><v>2</v></c>',
    )
    with Workbook.open(path) as book, EngineWorkbook.open(path) as engine:
        actual = book.dependency_impact("Inputs", "a1")
        retained = engine.dependency_impact("Inputs", "A1")
        assert isinstance(retained, DependencyImpact)
        # Equivalence oracle only: numerical/graph truth is asserted separately below.
        batch = BatchImpact(
            (RootImpact(retained.root, "started", retained, None, 0),),
            AnalysisWork(),
            {},
            retained.direct_dependents + retained.transitive_dependents,
            retained.transitive_dependents,
        )
        assert data(actual)["impact_summary"] == thaw_json(
            summarize_batch(
                batch, source_fingerprint=book.source_fingerprint, sheet_order={"Inputs": 0}
            )
        )
        assert data(actual)["work"] == _to_json(retained.work)
        assert data(actual)["limits"] == _to_json(retained.limits)
        observed = root(actual)
        assert observed["analysis_status"] == "complete"
        assert observed["known_direct_count"] == 1 and observed["known_transitive_count"] == 1
        assert observed["exact_total_count"] == 2
        assert observed["direct_dependents_sample"] == [{"sheet": "Inputs", "address": "B1"}]
        assert observed["transitive_dependents_sample"] == [{"sheet": "Inputs", "address": "C1"}]
        assert summary(actual)["numerical_effects_evaluated"] is False


def test_only_one_analysis_and_no_full_graph_serialization(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from xlayer import _workbook

    original = analyse_dependencies
    calls = 0

    def counted(*args: object, **kwargs: object) -> DependencyImpact:
        nonlocal calls
        calls += 1
        return original(*args, **kwargs)  # type: ignore[arg-type]

    def forbidden(_self: DependencyImpact) -> dict[str, object]:
        pytest.fail("public presentation serialized the raw dependency graph")

    monkeypatch.setattr(_workbook, "analyse_dependencies", counted)
    monkeypatch.setattr(DependencyImpact, "to_dict", forbidden)
    with Workbook.open(
        source(tmp_path, '<c r="A1"><v>1</v></c><c r="B1"><f>A1</f><v>1</v></c>')
    ) as book:
        result = book.dependency_impact("Inputs", "A1")
        assert isinstance(result, Inspection) and calls == 1
        work = data(result)["work"]
        assert isinstance(work, dict) and work["sheets_attempted"] == 1


def test_presentation_samples_do_not_reduce_counts_or_coverage(tmp_path: Path) -> None:
    payload = '<c r="A1"><v>1</v></c>' + "".join(
        f'<c r="{chr(66 + i)}1"><f>A1</f><v>1</v></c>' for i in range(15)
    )
    with Workbook.open(source(tmp_path, payload)) as book:
        result = book.dependency_impact("Inputs", "A1")
    observed = root(result)
    assert observed["analysis_status"] == "complete" and observed["known_direct_count"] == 15
    assert observed["direct_sample_truncated"] is True
    assert len(cast("list[object]", observed["direct_dependents_sample"])) == 10
    assert summary(result)["known_union_count"] == 15


def test_partial_preserves_lower_bounds_issues_and_traversal_frontiers(tmp_path: Path) -> None:
    with Workbook.open(
        source(
            tmp_path,
            '<c r="A1"><v>1</v></c><c r="B1"><f>A1</f><v>1</v></c>'
            '<c r="C1"><f>B1</f><v>1</v></c>'
            '<c r="D1"><f>INDIRECT("A1")</f><v>1</v></c>',
        )
    ) as book:
        result = book.dependency_impact("Inputs", "A1", limits=ImpactLimits(max_evidence_edges=1))
    observed = root(result)
    assert observed["analysis_status"] == "partial"
    assert observed["known_direct_count"] == 1 and observed["exact_total_count"] is None
    assert observed["traversal_frontier"]
    issues = summary(result)["issues"]
    assert isinstance(issues, dict) and issues["retained_issue_count"]
    assert data(result)["limits"] == _to_json(ImpactLimits(max_evidence_edges=1))


def test_unknown_is_null_not_zero_and_exhaustion_is_not_a_resume_api(tmp_path: Path) -> None:
    path = write_workbook_case(
        tmp_path / "s.xlsx",
        {
            "Inputs": '<sheetData><row><c r="A1"><v>1</v></c></row></sheetData>',
            "Bad": '<sheetData><row><c r="A1"><v>nan</v></c></row></sheetData>',
        },
    )
    with Workbook.open(path) as book:
        result = book.dependency_impact("Inputs", "A1", limits=ImpactLimits(max_sheets=1))
    observed = root(result)
    assert observed["analysis_status"] in {"partial", "unknown"}
    assert observed["inventory_frontier"] and observed["exact_total_count"] is None
    # Unknown is about the entire evidence inventory, not just the requested root.
    # One trusted sheet in the first input yields partial even for its bad neighbour.
    path = write_workbook_case(
        path, {"Bad": '<sheetData><row><c r="A1"><v>nan</v></c></row></sheetData>'}
    )
    with Workbook.open(path) as book:
        unknown = book.dependency_impact("Bad", "A1")
    observed = root(unknown)
    assert observed["analysis_status"] == "unknown"
    assert observed["known_direct_count"] is None and summary(unknown)["known_union_count"] is None


def test_complete_zero_does_not_mean_numerical_validation(tmp_path: Path) -> None:
    with Workbook.open(source(tmp_path, '<c r="A1"><v>1</v></c>')) as book:
        result = book.dependency_impact("Inputs", "A1")
    observed = root(result)
    assert observed["analysis_status"] == "complete" and observed["exact_total_count"] == 0
    assert summary(result)["numerical_effects_evaluated"] is False


@pytest.mark.parametrize("kind", ["formula", "definition"])
def test_oversized_diagnostic_is_bounded_before_summary_hashing(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, kind: str
) -> None:
    from xlayer import _impact_summary, _inspection

    huge = "A1" + " " * 1_048_576
    path = write_workbook_case(
        tmp_path / "s.xlsx",
        {
            "Inputs": '<sheetData><row><c r="A1"><v>1</v></c>'
            f'<c r="B1"><f>{huge if kind == "formula" else "LargeName"}</f>'
            "<v>1</v></c></row></sheetData>"
        },
        names_xml=f'<definedName name="LargeName">{huge}</definedName>'
        if kind == "definition"
        else "",
    )

    def forbidden(*_args: object, **_kwargs: object) -> bytes:
        pytest.fail("oversized issue was encoded or hashed before the fixed projection bound")

    monkeypatch.setattr(_impact_summary, "digest", forbidden)
    monkeypatch.setattr(_inspection, "canonical_json", forbidden)
    with Workbook.open(path) as book:
        result = book.dependency_impact("Inputs", "A1")
        assert isinstance(result, Refusal) and result.code == "inspection_limit_exceeded"
        assert result.details["budget"] == "max_summary_issue_bytes"
        assert result.details["limit"] == 1_048_576
        assert "observed_minimum_bytes" in result.details
        assert not book.closed and huge not in repr(result.to_dict())


def test_summary_response_byte_overflow_returns_no_improved_partial(tmp_path: Path) -> None:
    with Workbook.open(
        source(tmp_path, '<c r="A1"><v>1</v></c><c r="B1"><f>INDIRECT("A1")</f><v>1</v></c>')
    ) as book:
        full = book.dependency_impact("Inputs", "A1")
        assert isinstance(full, Inspection)
        limit = len(full.canonical_json())
        assert isinstance(
            book.dependency_impact(
                "Inputs", "A1", read_limits=ReadLimits(max_response_bytes=limit)
            ),
            Inspection,
        )
        refused = book.dependency_impact(
            "Inputs", "A1", read_limits=ReadLimits(max_response_bytes=limit - 1)
        )
        assert isinstance(refused, Refusal) and refused.details["budget"] == "max_response_bytes"
        assert "impact_summary" not in refused.details


@pytest.mark.parametrize("spec", fields(ImpactLimits()))
def test_public_analysis_cannot_raise_engine_ceiling(tmp_path: Path, spec: Field[int]) -> None:
    name = spec.name
    high = replace(ImpactLimits(), **{name: getattr(ImpactLimits(), name) + 1})
    with Workbook.open(source(tmp_path, "")) as book:
        with pytest.raises(ValueError, match=name):
            book.dependency_impact("Inputs", "A1", limits=high)
        assert not book.closed


def test_query_closed_missing_and_invalid_arguments(tmp_path: Path) -> None:
    book = Workbook.open(source(tmp_path, ""))
    result = book.dependency_impact("inputs", "a1")
    assert isinstance(result, Refusal) and result.code == "sheet_not_found"
    with pytest.raises(ValueError):
        book.dependency_impact("Inputs", "$A$1")
    assert not book.closed
    book.close()
    with pytest.raises(ClosedWorkbookError):
        book.dependency_impact("Inputs", "A1")


def test_cycle_retains_bounded_paths_without_numerical_claim(tmp_path: Path) -> None:
    with Workbook.open(
        source(tmp_path, '<c r="A1"><f>B1</f><v>1</v></c><c r="B1"><f>A1</f><v>1</v></c>')
    ) as book:
        result = book.dependency_impact("Inputs", "A1")
    observed = root(result)
    assert observed["known_direct_count"] == 1
    assert observed["path_examples"]
    assert summary(result)["numerical_effects_evaluated"] is False
    assert "known_cycles" not in observed  # this summary is not the full retained proof


def test_unexpected_summary_failure_closes_session(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from xlayer import _public_workbook

    failure = ValueError("summary invariant defect")

    def broken(*_args: object, **_kwargs: object) -> None:
        raise failure

    book = Workbook.open(source(tmp_path, '<c r="A1"><v>1</v></c>'))
    monkeypatch.setattr(_public_workbook, "single_impact_summary", broken)
    with pytest.raises(ValueError) as caught:
        book.dependency_impact("Inputs", "A1")
    assert caught.value is failure and book.closed


@pytest.mark.parametrize("payload", ["x", "😀"])
def test_diagnostic_exact_byte_n_and_n_plus_one(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, payload: str
) -> None:
    from xlayer import _inspection
    from xlayer._canonical import canonical_json
    from xlayer._dependencies import AnalysisIssue

    with EngineWorkbook.open(source(tmp_path, '<c r="A1"><v>1</v></c>')) as engine:
        retained = engine.dependency_impact("Inputs", "A1")
        assert isinstance(retained, DependencyImpact)
        issue = AnalysisIssue(
            "unsupported_formula",
            "Inputs",
            None,
            None,
            {"text": payload * 100},
            ({"action": "inspect"},),
        )
        size = len(canonical_json(_to_json(issue)))
        retained = replace(
            retained, issues=(issue,), analysis_status="partial", exact_total_count=None
        )
        monkeypatch.setattr(_inspection, "MAX_SUMMARY_ISSUE_BYTES", size)
        assert not isinstance(_inspection.single_impact_summary(retained, {"Inputs": 0}), Refusal)
        monkeypatch.setattr(_inspection, "MAX_SUMMARY_ISSUE_BYTES", size - 1)
        refused = _inspection.single_impact_summary(retained, {"Inputs": 0})
        assert (
            isinstance(refused, Refusal) and refused.details["budget"] == "max_summary_issue_bytes"
        )


@pytest.mark.parametrize("argument", ["limits", "read_limits"])
def test_wrong_analysis_limits_type_is_programming_error(tmp_path: Path, argument: str) -> None:
    with Workbook.open(source(tmp_path, "")) as book:
        with pytest.raises(TypeError):
            book.dependency_impact("Inputs", "A1", **{argument: None})  # type: ignore[arg-type]
        assert not book.closed
