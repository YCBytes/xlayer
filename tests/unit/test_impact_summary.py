"""Independent counts, selected proofs and disclosure at projection boundaries."""

from __future__ import annotations

import importlib.util
from collections.abc import Mapping
from dataclasses import replace
from pathlib import Path
from typing import cast

import pytest

from tests.unit._dependency_cases import write_workbook_case
from xlayer._batch_impact import BatchImpact, analyse_batch
from xlayer._canonical import canonical_json, thaw_json
from xlayer._dependencies import (
    DEFAULT_IMPACT_LIMITS,
    AnalysisIssue,
    CellRef,
    DependencyImpact,
    ImpactLimits,
)
from xlayer._workbook import Workbook


def analyze(
    tmp_path: Path,
    sheets: Mapping[str, str],
    *,
    roots: tuple[CellRef, ...] = (CellRef("S", "A1"),),
    limits: ImpactLimits = DEFAULT_IMPACT_LIMITS,
) -> tuple[BatchImpact, str, dict[str, int]]:
    path = write_workbook_case(tmp_path / "summary.xlsx", sheets)
    with Workbook.open(path) as book:
        return (
            analyse_batch(book.registry, book.read_sheet, roots, book.source_fingerprint, limits),
            book.source_fingerprint,
            {s.name: s.tab_index for s in book.registry.sheets},
        )


def project(batch: BatchImpact, fingerprint: str, order: Mapping[str, int]) -> dict[str, object]:
    assert importlib.util.find_spec("xlayer._impact_summary") is not None, "impact summary missing"
    from xlayer._impact_summary import summarize_batch

    return cast(
        "dict[str, object]",
        thaw_json(summarize_batch(batch, source_fingerprint=fingerprint, sheet_order=order)),
    )


def rows(summary: dict[str, object]) -> list[dict[str, object]]:
    return cast("list[dict[str, object]]", summary["roots"])


def test_overlap_counts_are_union_not_sum_and_proofs_are_selected(tmp_path: Path) -> None:
    batch, fp, order = analyze(
        tmp_path,
        {
            "S": '<sheetData><row><c r="A1"><v>1</v></c>'
            '<c r="B1"><v>2</v></c><c r="C1"><f>A1+B1</f></c><c r="D1"><f>C1</f></c>'
            "</row></sheetData>"
        },
        roots=(CellRef("S", "A1"), CellRef("S", "B1")),
    )
    before = batch.to_dict()
    summary = project(batch, fp, order)
    assert summary["known_union_count"] == 2
    assert summary["known_transitive_union_count"] == 1
    assert summary["affected_sheets"] == [{"sheet": "S", "known_affected_cell_count": 2}]
    assert summary["known_affected_sheet_count"] == 1
    assert summary["scope"] == "worksheet_cell_formulas"
    assert summary["numerical_effects_evaluated"] is False
    assert summary["dependency_contract_version"] == "1.1"
    assert summary["impact_summary_version"] == "1.0"
    assert summary["source_fingerprint"] == fp
    assert cast(dict[str, object], summary["issues"])["retained_issue_count"] == 0
    for root in rows(summary):
        assert root["known_direct_count"] == root["known_transitive_count"] == 1
        assert root["exact_total_count"] == 2
        paths = cast(list[dict[str, object]], root["path_examples"])
        assert [p["total_hops"] for p in paths] == [1, 2]
        assert [p["kind"] for p in paths] == ["direct", "transitive"]
        for path in paths:
            for hop in cast(list[dict[str, object]], path["hops"]):
                assert hop["reference_text"] in {"A1", "B1", "C1"}
                assert hop["basis"] == ["direct_ooxml"]
    assert batch.to_dict() == before
    assert canonical_json(summary) == canonical_json(
        project(batch, fp, dict(reversed(tuple(order.items()))))
    )


@pytest.mark.parametrize(
    "body,status,count",
    [
        ("<sheetData/>", "complete", 0),
        ('<sheetData><row><c r="B1"><f>TABLE(A1)</f></c></row></sheetData>', "partial", 0),
        ("<sheetData><wrapper/></sheetData>", "unknown", None),
    ],
)
def test_empty_samples_do_not_invent_independence(
    tmp_path: Path,
    body: str,
    status: str,
    count: int | None,
) -> None:
    batch, fp, order = analyze(tmp_path, {"S": body})
    summary = project(batch, fp, order)
    root = rows(summary)[0]
    assert root["analysis_status"] == status and root["known_direct_count"] == count
    assert summary["known_union_count"] == count
    assert root["exact_total_count"] == (0 if status == "complete" else None)
    assert root["direct_dependents_sample"] == root["path_examples"] == []


def test_not_started_root_discloses_null_counts_and_actual_budget(tmp_path: Path) -> None:
    batch, fp, order = analyze(
        tmp_path,
        {"S": "<sheetData/>"},
        roots=(CellRef("S", "A1"), CellRef("S", "B1")),
        limits=ImpactLimits(max_visited_nodes=1),
    )
    root = rows(project(batch, fp, order))[1]
    assert root["state"] == "not_started" and root["analysis_status"] == "not_evaluated"
    assert root["blocked_budget"] == "max_visited_nodes" and root["root_cursor"] == 1
    assert (
        root["known_direct_count"]
        is root["known_transitive_count"]
        is root["exact_total_count"]
        is None
    )


@pytest.mark.parametrize("count", [9, 10, 11])
def test_direct_sample_boundary_keeps_total_and_truncation(tmp_path: Path, count: int) -> None:
    body = (
        "<sheetData><row>"
        + "".join(f'<c r="{chr(66 + i)}1"><f>A1</f></c>' for i in range(count))
        + "</row></sheetData>"
    )
    batch, fp, order = analyze(tmp_path, {"S": body})
    root = rows(project(batch, fp, order))[0]
    assert root["known_direct_count"] == count
    assert len(cast(list[object], root["direct_dependents_sample"])) == min(count, 10)
    assert root["direct_sample_truncated"] is (count > 10)


@pytest.mark.parametrize("count", [9, 10, 11])
def test_transitive_sample_boundary_keeps_total(tmp_path: Path, count: int) -> None:
    body = (
        '<sheetData><row><c r="B1"><f>A1</f></c></row>'
        + "".join(
            f'<row r="{i}"><c r="A{i}"><f>{"B1" if i == 2 else f"A{i - 1}"}</f></c></row>'
            for i in range(2, count + 2)
        )
        + "</sheetData>"
    )
    batch, fp, order = analyze(tmp_path, {"S": body})
    root = rows(project(batch, fp, order))[0]
    assert root["known_transitive_count"] == count
    assert len(cast(list[object], root["transitive_dependents_sample"])) == min(count, 10)
    assert root["transitive_sample_truncated"] is (count > 10)


@pytest.mark.parametrize("count", [9, 10, 11])
def test_sheet_samples_use_tab_order_not_alphabetical_order(tmp_path: Path, count: int) -> None:
    sheets = {"S": "<sheetData/>"}
    for i in range(count, 0, -1):
        sheets[f"Tab{i}"] = '<sheetData><row><c r="A1"><f>S!A1</f></c></row></sheetData>'
    batch, fp, order = analyze(tmp_path, sheets)
    summary = project(batch, fp, order)
    assert summary["known_affected_sheet_count"] == count
    assert summary["affected_sheets_truncated"] is (count > 10)
    retained = cast(list[dict[str, object]], summary["affected_sheets"])
    assert [s["sheet"] for s in retained] == [
        f"Tab{i}" for i in range(count, max(count - 10, 0), -1)
    ]
    assert all(s["known_affected_cell_count"] == 1 for s in retained)


@pytest.mark.parametrize("hops", [11, 12, 13])
def test_path_prefix_is_explicit_and_does_not_enumerate_all_paths(
    tmp_path: Path,
    hops: int,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    body = f'<sheetData><row><c r="A1"><f>Z{hops - 1}</f></c><c r="B1"><v>1</v></c>'
    body += '<c r="Z1"><f>B1</f></c></row>'
    body += (
        "".join(f'<row r="{i}"><c r="Z{i}"><f>Z{i - 1}</f></c></row>' for i in range(2, hops))
        + "</sheetData>"
    )
    batch, fp, order = analyze(tmp_path, {"S": body}, roots=(CellRef("S", "B1"),))
    calls: list[CellRef] = []
    original = DependencyImpact.path_to

    def tracked(self: DependencyImpact, ref: CellRef) -> tuple[object, ...]:
        calls.append(ref)
        return original(self, ref)

    monkeypatch.setattr(DependencyImpact, "path_to", tracked)
    paths = cast(list[dict[str, object]], rows(project(batch, fp, order))[0]["path_examples"])
    path = paths[1]
    assert calls == [CellRef("S", "Z1"), CellRef("S", "A1")]
    assert path["target"] == {"sheet": "S", "address": "A1"}
    assert path["total_hops"] == hops and path["path_truncated"] is (hops > 12)
    assert len(cast(list[object], path["hops"])) == min(hops, 12)


@pytest.mark.parametrize("count", [19, 20, 21])
def test_issue_group_cap_and_shared_inventory_dedup(tmp_path: Path, count: int) -> None:
    body = (
        "<sheetData>"
        + "".join(
            f'<row r="{i + 1}"><c r="C{i + 1}"><f>Fn{i:02d}(A1)</f></c></row>' for i in range(count)
        )
        + "</sheetData>"
    )
    batch, fp, order = analyze(
        tmp_path, {"S": body}, roots=(CellRef("S", "A1"), CellRef("S", "B1"))
    )
    issues = cast(dict[str, object], project(batch, fp, order)["issues"])
    assert issues["retained_issue_count"] == count
    assert issues["sum_per_root_issue_count"] == count * 2
    assert issues["retained_group_count"] == count
    assert issues["omitted_group_count"] == max(count - 20, 0)
    groups = cast(list[dict[str, object]], issues["groups"])
    assert [g["function"] for g in groups] == [f"FN{i:02d}" for i in range(min(count, 20))]


def test_function_labels_require_own_exact_valid_span_and_examples_are_bounded(
    tmp_path: Path,
) -> None:
    batch, fp, order = analyze(tmp_path, {"S": "<sheetData/>"})
    impact = batch.roots[0].impact
    assert impact is not None
    issues = tuple(
        AnalysisIssue(
            "unsupported_function",
            "S",
            CellRef("S", f"C{i + 1}"),
            span,
            {"formula_text": text, "reason": "TABLE isn't admitted"},
            ({"action": "inspect_formula"},),
        )
        for i, (text, span) in enumerate(
            [
                ("Bad(A1)", (0, 3)),
                ("Bad(B1)", (0, 3)),
                ("Bad(C1)", (0, 3)),
                ("Other(A1)", None),
                ("Other(A1)", (0, 99)),
                ("Other(A1)", (0, 6)),
            ]
        )
    )
    changed = replace(
        batch,
        roots=(
            replace(
                batch.roots[0],
                impact=replace(
                    impact, analysis_status="partial", exact_total_count=None, issues=issues
                ),
            ),
        ),
    )
    grouped = cast(dict[str, object], project(changed, fp, order)["issues"])
    groups = cast(list[dict[str, object]], grouped["groups"])
    assert [(g["function"], g["count"]) for g in groups] == [(None, 3), ("BAD", 3)]
    assert all(
        g["examples_truncated"] is True and len(cast(list[object], g["examples"])) == 2
        for g in groups
    )


def test_summary_identity_and_recursive_freeze_are_not_caller_mutable(tmp_path: Path) -> None:
    batch, fp, order = analyze(
        tmp_path, {"S": '<sheetData><row><c r="B1"><f>A1</f></c></row></sheetData>'}
    )
    project(batch, fp, order)
    from xlayer._impact_summary import summarize_batch

    frozen = summarize_batch(batch, source_fingerprint=fp, sheet_order=order)
    detached = cast(dict[str, object], thaw_json(frozen))
    rows(detached).clear()
    order.clear()
    assert len(cast(tuple[object, ...], frozen["roots"])) == 1
    with pytest.raises(TypeError):
        cast(dict[str, object], frozen)["known_union_count"] = 99
    nested = cast(tuple[Mapping[str, object], ...], frozen["roots"])[0]
    with pytest.raises(TypeError):
        cast(dict[str, object], nested)["analysis_status"] = "complete"
    with pytest.raises(ValueError, match="fingerprint"):
        summarize_batch(batch, source_fingerprint="sha256:" + "0" * 64, sheet_order={"S": 0})
    with pytest.raises(ValueError, match="sheet"):
        summarize_batch(batch, source_fingerprint=fp, sheet_order={})


def test_cycles_and_incomplete_frontiers_keep_authoritative_facts(tmp_path: Path) -> None:
    batch, fp, order = analyze(
        tmp_path,
        {"S": '<sheetData><row><c r="A1"><f>B1</f></c><c r="B1"><f>A1</f></c></row></sheetData>'},
    )
    root = rows(project(batch, fp, order))[0]
    assert root["known_direct_count"] == 1 and root["known_transitive_count"] == 0
    assert root["exact_total_count"] == 1
    assert len(cast(list[object], root["path_examples"])) == 1
    partial, fp, order = analyze(
        tmp_path,
        {"S": '<sheetData><row><c r="B1"><f>A1</f></c><c r="C1"><f>B1</f></c></row></sheetData>'},
        limits=ImpactLimits(max_visited_nodes=1),
    )
    root = rows(project(partial, fp, order))[0]
    assert root["analysis_status"] == "partial" and root["exact_total_count"] is None
    impact = partial.roots[0].impact
    assert impact is not None and impact.traversal_frontier
    from xlayer._dependencies import _to_json

    assert root["traversal_frontier"] == _to_json(impact.traversal_frontier)


def test_issue_identity_includes_details_not_just_group_or_site(tmp_path: Path) -> None:
    batch, fp, order = analyze(
        tmp_path, {"S": "<sheetData/>"}, roots=(CellRef("S", "A1"), CellRef("S", "B1"))
    )
    common = AnalysisIssue(
        "unsupported_function",
        "S",
        CellRef("S", "C1"),
        (0, 3),
        {"formula_text": "Bad(A1)"},
        ({"action": "inspect_formula"},),
    )
    changed_roots = []
    for i, root in enumerate(batch.roots):
        assert root.impact is not None
        traversal = AnalysisIssue(
            "analysis_limit",
            "S",
            None,
            None,
            {"root_cursor": i},
            ({"action": "increase_explicit_limit"},),
        )
        changed_roots.append(
            replace(
                root,
                impact=replace(
                    root.impact,
                    analysis_status="partial",
                    exact_total_count=None,
                    issues=(common, traversal),
                ),
            )
        )
    summary = project(replace(batch, roots=tuple(changed_roots)), fp, order)
    issues = cast(dict[str, object], summary["issues"])
    assert issues["retained_issue_count"] == 3 and issues["sum_per_root_issue_count"] == 4
    groups = cast(list[dict[str, object]], issues["groups"])
    assert [(g["code"], g["function"], g["count"]) for g in groups] == [
        ("analysis_limit", None, 2),
        ("unsupported_function", "BAD", 1),
    ]


def test_distinct_real_root_traversal_cutoffs_are_not_deduplicated(tmp_path: Path) -> None:
    batch, fp, order = analyze(
        tmp_path,
        {
            "S": '<sheetData><row><c r="A1"><v>1</v></c>'
            '<c r="C1"><f>SUM(A:A)</f></c></row>'
            '<row r="2"><c r="A2"><v>2</v></c></row></sheetData>'
        },
        roots=(CellRef("S", "A1"), CellRef("S", "A2")),
        limits=ImpactLimits(max_membership_checks=1),
    )
    summary = project(batch, fp, order)
    issues = cast(dict[str, object], summary["issues"])
    assert issues["retained_issue_count"] == issues["sum_per_root_issue_count"] == 2
    groups = cast(list[dict[str, object]], issues["groups"])
    assert [(g["code"], g["count"]) for g in groups] == [("traversal_budget_exceeded", 2)]
    for index, expected_address in enumerate(("A1", "A2")):
        impact = batch.roots[index].impact
        assert impact is not None and impact.analysis_status == "partial"
        assert impact.issues[0].details["root"] == {
            "sheet": "S",
            "address": expected_address,
        }
    first, second = batch.roots[0].impact, batch.roots[1].impact
    assert first is not None and second is not None
    assert first.traversal_frontier[0].precedent == CellRef("S", "C1")
    assert second.traversal_frontier[0].precedent == CellRef("S", "A2")
