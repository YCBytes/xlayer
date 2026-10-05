"""Independent result-schema and bounded graph assertions."""

from __future__ import annotations

import hashlib
import json
import random
from collections.abc import Callable, Mapping
from dataclasses import fields, replace
from pathlib import Path
from typing import cast

import pytest

from tests.unit._dependency_cases import write_workbook_case
from xlayer._dependencies import (
    DEFAULT_IMPACT_LIMITS,
    AnalysisIssue,
    AnalysisWork,
    CellRef,
    DependencyEdge,
    DependencyImpact,
    ImpactLimits,
    InventoryCursor,
    NameEvidence,
    RangeRef,
    ReferenceEvidence,
    TraversalCursor,
    _freeze_json,
    _thaw_json,
    analyse_dependencies,
)
from xlayer._errors import Refusal
from xlayer._ooxml.sheet import Worksheet
from xlayer._ooxml.workbook import WorkbookRegistry
from xlayer._workbook import Workbook

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"
FINGERPRINT = "sha256:" + "0" * 64


def example() -> DependencyImpact:
    root, dependent = CellRef("S", "A1"), CellRef("S", "B1")
    evidence = ReferenceEvidence(
        dependent,
        "A1",
        (0, 2),
        0,
        "A1",
        RangeRef("S", 1, 1, 1, 1),
        "cell",
        ("direct_ooxml",),
        None,
        None,
        None,
    )
    edge = DependencyEdge(root, dependent, evidence)
    work = AnalysisWork(1, 1, 1, 0, 1, 2, 0, 1, 0, 2, 1)
    return DependencyImpact(
        "1.0",
        "1.0",
        FINGERPRINT,
        root,
        "worksheet_cell_formulas",
        "complete",
        (dependent,),
        (),
        1,
        0,
        1,
        (edge,),
        {dependent: edge},
        (),
        (),
        (),
        (),
        (),
        ImpactLimits(),
        work,
        False,
    )


def test_limits_defaults_and_domain() -> None:
    limits = ImpactLimits()
    assert [getattr(limits, f.name) for f in fields(limits)] == [
        256,
        250000,
        10000,
        25000,
        8192,
        64,
        100000,
        1000000,
        10000,
        50000,
    ]
    for field in fields(limits):
        for bad in (True, 1.0, 0, -1, 2**63):
            with pytest.raises(ValueError, match=field.name):
                replace(limits, **{field.name: cast(int, bad)})
        assert getattr(replace(limits, **{field.name: 1}), field.name) == 1
        assert getattr(replace(limits, **{field.name: 2**63 - 1}), field.name) == 2**63 - 1


def test_frozen_record_construction_isolated_from_inputs() -> None:
    original = example()
    dependents = list(original.direct_dependents)
    cycles = [[original.root]]
    predecessor = dict(original.predecessor_edges)
    result = replace(
        original,
        direct_dependents=cast(tuple[CellRef, ...], dependents),
        known_cycles=cast(tuple[tuple[CellRef, ...], ...], cycles),
        predecessor_edges=predecessor,
    )
    dependents.clear()
    cycles[0].clear()
    predecessor.clear()
    assert result.direct_dependents == (CellRef("S", "B1"),)
    assert result.known_cycles == ((CellRef("S", "A1"),),)
    assert result.path_to(CellRef("S", "B1")) == original.evidence_edges
    with pytest.raises(TypeError):
        replace(original, root=cast(CellRef, {"sheet": "S", "address": "A1"}))


def test_exact_complete_json() -> None:
    expected = json.loads((FIXTURES / "dependency_complete.expected.json").read_text())
    result = example()
    assert result.to_dict() == expected
    out = result.to_dict()
    edges = cast(list[dict[str, object]], out["evidence_edges"])
    edges.clear()
    assert result.to_dict() == expected
    assert json.dumps(result.to_dict(), sort_keys=True, ensure_ascii=False, allow_nan=False)


@pytest.mark.parametrize("value", [2**63, -(2**63) - 1, float("nan"), float("inf"), "\ud800"])
def test_json_invalid_scalars(value: object) -> None:
    with pytest.raises(ValueError):
        _freeze_json(value)


@pytest.mark.parametrize("value", [object(), b"bytes", {1: "bad"}, {1, 2}])
def test_json_wrong_domain(value: object) -> None:
    with pytest.raises(TypeError):
        _freeze_json(value)


def test_json_scalar_and_container_domain() -> None:
    shared: list[object] = [None, True, 2**63 - 1, -(2**63), 1.5, "😀\ufffe"]
    data = {"a": shared, "b": shared}
    frozen = _freeze_json(data)
    shared.clear()
    assert _thaw_json(frozen) == {
        "a": [None, True, 2**63 - 1, -(2**63), 1.5, "😀\ufffe"],
        "b": [None, True, 2**63 - 1, -(2**63), 1.5, "😀\ufffe"],
    }
    cycle: list[object] = []
    cycle.append(cycle)
    with pytest.raises(ValueError):
        _freeze_json(cycle)
    nested: object = 0
    for _ in range(65):
        nested = [nested]
    with pytest.raises(ValueError):
        _freeze_json(nested)


def test_issue_is_deeply_frozen_and_recovery_required() -> None:
    details = {"outer": {"list": [1]}}
    issue = AnalysisIssue(
        "formula_parse_failure",
        "S",
        CellRef("S", "B1"),
        (0, 2),
        details,
        ({"action": "inspect_formula"},),
    )
    details["outer"]["list"].append(2)
    outer = cast(Mapping[str, object], issue.details["outer"])
    assert outer["list"] == (1,)
    with pytest.raises(ValueError):
        replace(issue, recovery_options=())
    with pytest.raises(ValueError):
        replace(issue, recovery_options=({},))


def test_compound_provenance_and_optional_fields() -> None:
    ref = RangeRef("S", 1, 1, 2, 2)
    name = NameEvidence("ScopedRate", None, "S!$A$1:$B$2", True, ref)
    cell = CellRef("S", "B3")
    evidence = ReferenceEvidence(
        cell,
        "SUM(ScopedRate)",
        (4, 14),
        0,
        "ScopedRate",
        ref,
        "defined_name",
        ("direct_ooxml", "shared_translation", "defined_name_resolution", "range_membership"),
        name,
        CellRef("S", "A3"),
        (0, 1),
    )
    assert evidence.shared_offset == (0, 1)
    assert evidence.name_definition == name
    with pytest.raises(ValueError):
        replace(evidence, reference_text="not the source slice")
    with pytest.raises(ValueError):
        replace(evidence, shared_master=None)


def test_paths_use_observed_shortest_predecessor_tree() -> None:
    result = example()
    assert result.path_to(result.root) == ()
    assert result.path_to(CellRef("S", "B1")) == result.evidence_edges
    with pytest.raises(ValueError):
        result.path_to(CellRef("S", "C1"))


def analyze(
    tmp_path: Path,
    body: str,
    *,
    root: str = "A1",
    limits: ImpactLimits = DEFAULT_IMPACT_LIMITS,
    other: str | None = None,
    names: str = "",
) -> DependencyImpact:
    sheets = {"S": body}
    if other is not None:
        sheets["T"] = other
    path = write_workbook_case(tmp_path / "case.xlsx", sheets, names_xml=names)
    with Workbook.open(path) as book:
        return analyse_dependencies(
            book.registry, book.read_sheet, CellRef("S", root), book.source_fingerprint, limits
        )


def grid(cells: str) -> str:
    return '<sheetData><row r="1">' + cells + "</row></sheetData>"


@pytest.mark.parametrize("root_cell", ['<c r="C12"><v>10</v></c>', '<c r="C12"/>', ""])
def test_value_root_chain(tmp_path: Path, root_cell: str) -> None:
    body = (
        '<sheetData><row r="12">'
        + root_cell
        + (
            '<c r="D12"><f>C12*2</f><v>20</v></c>'
            '<c r="E12"><f>D12+5</f><v>25</v></c></row></sheetData>'
        )
    )
    result = analyze(tmp_path, body, root="C12")
    assert result.analysis_status == "complete"
    assert result.direct_dependents == (CellRef("S", "D12"),)
    assert result.transitive_dependents == (CellRef("S", "E12"),)
    assert result.exact_total_count == 2
    assert [edge.dependent.address for edge in result.path_to(CellRef("S", "E12"))] == [
        "D12",
        "E12",
    ]


def test_repeated_occurrences_and_direct_transitive_do_not_double_count(tmp_path: Path) -> None:
    result = analyze(tmp_path, grid('<c r="B1"><f>A1+A1</f></c><c r="C1"><f>A1+B1</f></c>'))
    assert result.direct_dependents == (CellRef("S", "B1"), CellRef("S", "C1"))
    assert result.transitive_dependents == ()
    assert len(result.evidence_edges) == 4
    assert result.path_to(CellRef("S", "C1"))[0].precedent == result.root


@pytest.mark.parametrize("root", ["A1", "XFD1048576", "C12"])
def test_sparse_geometric_range_never_expands_grid(tmp_path: Path, root: str) -> None:
    result = analyze(tmp_path, grid('<c r="B1"><f>SUM(A1:XFD1048576)</f></c>'), root=root)
    assert result.work.references_admitted == 1
    assert result.work.scanned_cells == 1
    assert result.direct_dependents == (CellRef("S", "B1"),)
    assert result.work.membership_checks == 2  # root and discovered B1


@pytest.mark.parametrize("root,count", [("A1", 1), ("A2", 1), ("A3", 0), ("B1", 0)])
def test_range_boundary_and_outside(tmp_path: Path, root: str, count: int) -> None:
    result = analyze(tmp_path, grid('<c r="C1"><f>SUM(A1:A2)</f></c>'), root=root)
    assert result.known_direct_count == count


@pytest.mark.parametrize(
    "bad",
    [
        'INDIRECT("A1")',
        "A1+OFFSET(B1,1,1)",
        "INDEX(A1,1)",
        "Missing!A1",
        "#REF!",
        "A:A",
        "A1 B1",
        "Table[Col]",
        "_xlfn.SUM(A1)",
        "A1#",
        "[1]S!A1",
        "S:T!A1",
        "{1,2}",
        "NoName",
        "SUM(ScopedRate)",
    ],
)
def test_unsupported_site_is_global_partial_not_guessed_edge(tmp_path: Path, bad: str) -> None:
    from xml.sax.saxutils import escape

    result = analyze(tmp_path, grid(f'<c r="B1"><f>A1</f></c><c r="C1"><f>{escape(bad)}</f></c>'))
    assert result.analysis_status == "partial" and result.exact_total_count is None
    assert result.direct_dependents == (CellRef("S", "B1"),)
    assert [edge.dependent.address for edge in result.evidence_edges] == ["B1"]
    assert result.issues[0].formula_cell == CellRef("S", "C1")
    assert result.issues[0].recovery_options


def test_guarded_sheet_failure_preserves_lower_bound(tmp_path: Path) -> None:
    bad = '<sheetData><wrapper><row><c r="B1"><f>A1</f></c></row></wrapper></sheetData>'
    result = analyze(tmp_path, grid('<c r="B1"><f>A1</f></c>'), other=bad)
    assert result.analysis_status == "partial" and result.known_direct_count == 1
    assert result.issues[0].code == "worksheet_unreadable"
    refusal = cast(Mapping[str, object], result.issues[0].details["refusal"])
    assert refusal["code"] == "invalid_part_content"
    unknown = analyze(tmp_path, bad, other=bad)
    assert unknown.analysis_status == "unknown"
    assert unknown.known_direct_count is None and unknown.exact_total_count is None
    assert unknown.evidence_edges == ()


def test_unused_bad_names_do_not_reduce_coverage(tmp_path: Path) -> None:
    result = analyze(
        tmp_path,
        grid('<c r="B1"><f>A1</f></c>'),
        names='<definedName name="Invalid">#REF!</definedName>',
    )
    assert result.analysis_status == "complete" and result.work.defined_names_admitted == 0


def test_shared_members_and_combined_evidence(tmp_path: Path) -> None:
    body = '<sheetData><row r="3"><c r="B3">'
    body += '<f t="shared" si="7" ref="B3:B6">SUM(ScopedRate)</f></c></row>'
    body += '<row r="4"><c r="B4"><f t="shared" si="9" ref="B4">A1</f></c></row>'
    body += '<row r="6"><c r="B6"><f t="shared" si="7"/></c></row></sheetData>'
    result = analyze(
        tmp_path, body, names='<definedName name="ScopedRate">S!$A$1:$A$2</definedName>'
    )
    assert [r.address for r in result.direct_dependents] == ["B3", "B4", "B6"]
    assert CellRef("S", "B5") not in result.direct_dependents
    shared = next(edge.evidence for edge in result.evidence_edges if edge.dependent.address == "B6")
    assert shared.basis == (
        "direct_ooxml",
        "shared_translation",
        "defined_name_resolution",
        "range_membership",
    )
    assert shared.formula_text == "SUM(ScopedRate)" and shared.shared_offset == (3, 0)
    assert shared.shared_master == CellRef("S", "B3")
    assert shared.normalized_reference == RangeRef("S", 1, 1, 2, 1)


@pytest.mark.parametrize(
    "body,cycles",
    [
        ('<c r="A1"><f>A1</f></c>', (("A1",),)),
        ('<c r="A1"><f>B1</f></c><c r="B1"><f>A1</f></c>', (("A1", "B1"),)),
        ('<c r="B1"><f>A1+C1</f></c><c r="C1"><f>B1</f></c>', (("B1", "C1"),)),
    ],
)
def test_observed_cycles_and_shortest_paths(
    tmp_path: Path, body: str, cycles: tuple[tuple[str, ...], ...]
) -> None:
    result = analyze(tmp_path, grid(body))
    assert tuple(tuple(c.address for c in cycle) for cycle in result.known_cycles) == cycles
    assert result.root not in result.direct_dependents + result.transitive_dependents
    assert result.path_to(result.root) == ()


def test_index_identity_failure_discards_proof(tmp_path: Path) -> None:
    path = write_workbook_case(tmp_path / "case.xlsx", {"S": grid('<c r="B1"><f>A1</f></c>')})
    with Workbook.open(path) as book:
        sheet = book.read_sheet("S")
        assert isinstance(sheet, Worksheet)

        def corrupted(_name: str) -> Worksheet | Refusal:
            return replace(sheet, name="Wrong")

        result = analyse_dependencies(
            book.registry, corrupted, CellRef("S", "A1"), book.source_fingerprint, ImpactLimits()
        )
        assert result.analysis_status == "unknown" and result.known_direct_count is None
        assert result.issues[0].code == "dependency_invariant_failure"


def test_declared_dependency_case_helper_imports() -> None:
    assert write_workbook_case.__module__ == "tests.unit._dependency_cases"


@pytest.mark.parametrize(
    "budget", ["max_sheets", "max_scanned_cells", "max_formula_cells", "max_references"]
)
def test_inventory_limit_frontiers(tmp_path: Path, budget: str) -> None:
    body = grid('<c r="B1"><f>A1</f></c><c r="C1"><f>A1+B1</f></c>')
    result = analyze(
        tmp_path, body, limits=replace(ImpactLimits(), **{budget: 1}), other="<sheetData/>"
    )
    assert result.analysis_status == "partial" and result.inventory_frontier
    assert result.issues[-1].code == "inventory_budget_exceeded"
    assert result.work.sheets_attempted <= result.limits.max_sheets
    assert result.work.scanned_cells <= result.limits.max_scanned_cells
    assert result.work.formula_cells_attempted <= result.limits.max_formula_cells
    assert result.work.references_admitted <= result.limits.max_references
    if budget == "max_references":
        assert result.work.references_admitted == 1  # never admit a prefix of C1
        assert result.inventory_frontier[0].phase == "admit_formula"


def test_name_inventory_limit_is_full_and_lazy(tmp_path: Path) -> None:
    names = '<definedName name="ScopedRate">S!$A$1</definedName>' * 2
    limits = replace(ImpactLimits(), max_defined_names=1)
    unused = analyze(tmp_path, grid('<c r="B1"><f>A1</f></c>'), names=names, limits=limits)
    assert unused.analysis_status == "complete" and unused.work.defined_names_admitted == 0
    used = analyze(tmp_path, grid('<c r="B1"><f>ScopedRate</f></c>'), names=names, limits=limits)
    assert used.analysis_status == "partial" and used.evidence_edges == ()
    assert used.work.defined_names_admitted == 0
    assert used.issues[0].code == "name_inventory_limit_exceeded"
    admitted = analyze(
        tmp_path,
        grid('<c r="B1"><f>ScopedRate</f></c>'),
        names='<definedName name="ScopedRate">S!$A$1</definedName>',
        limits=limits,
    )
    assert admitted.analysis_status == "complete" and admitted.work.defined_names_admitted == 1


def test_formula_limits_are_local_not_global_abort(tmp_path: Path) -> None:
    body = grid('<c r="B1"><f>' + "A1" + " " * 11 + '</f></c><c r="C1"><f>A1</f></c>')
    result = analyze(tmp_path, body, limits=replace(ImpactLimits(), max_formula_chars=12))
    assert result.analysis_status == "partial"
    assert result.direct_dependents == (CellRef("S", "C1"),)
    assert result.work.max_formula_chars_observed == 13
    deep = analyze(
        tmp_path,
        grid('<c r="B1"><f>((A1))</f></c><c r="C1"><f>A1</f></c>'),
        limits=replace(ImpactLimits(), max_formula_nesting=1),
    )
    assert deep.direct_dependents == (CellRef("S", "C1"),)
    assert deep.work.max_formula_nesting_observed == 2


def test_mixed_candidate_edge_cutoff_is_canonical(tmp_path: Path) -> None:
    body = grid('<c r="B1"><f>SUM(A1:A2)</f></c><c r="C1"><f>A1</f></c>')
    result = analyze(tmp_path, body, limits=replace(ImpactLimits(), max_evidence_edges=1))
    assert result.direct_dependents == (CellRef("S", "B1"),)
    assert result.work.membership_checks == 1
    current, queued = result.traversal_frontier
    assert (current.precedent.address, current.phase) == ("A1", "admit_edge")
    assert current.next_candidate is not None and current.next_candidate.dependent.address == "C1"
    assert (queued.precedent.address, queued.phase) == ("B1", "queued")


def test_membership_cutoff_retains_unchecked_candidate(tmp_path: Path) -> None:
    body = grid('<c r="B1"><f>SUM(A1:A2)</f></c><c r="C1"><f>SUM(A1:A2)</f></c>')
    result = analyze(tmp_path, body, limits=replace(ImpactLimits(), max_membership_checks=1))
    assert result.direct_dependents == (CellRef("S", "B1"),)
    current = result.traversal_frontier[0]
    assert current.phase == "match_candidate" and current.next_candidate is not None
    assert current.next_candidate.dependent.address == "C1"
    assert result.work.membership_checks == 1


def test_node_budget_and_exact_ceiling(tmp_path: Path) -> None:
    body = grid('<c r="B1"><f>A1</f></c>')
    blocked = analyze(tmp_path, body, limits=replace(ImpactLimits(), max_visited_nodes=1))
    assert blocked.direct_dependents == () and blocked.work.visited_nodes == 1
    assert blocked.traversal_frontier[0].phase == "admit_edge"
    exact = analyze(
        tmp_path,
        body,
        limits=replace(
            ImpactLimits(),
            max_visited_nodes=2,
            max_evidence_edges=1,
            max_references=1,
            max_formula_cells=1,
            max_scanned_cells=1,
        ),
    )
    assert exact.analysis_status == "complete" and exact.exact_total_count == 1


@pytest.mark.parametrize("cutoff", [False, True])
def test_analyzed_complete_and_cutoff_json_are_handwritten(tmp_path: Path, cutoff: bool) -> None:
    body = (
        '<c r="B1"><f>SUM(A1:A2)</f></c><c r="C1"><f>A1</f></c>'
        if cutoff
        else ('<c r="B1"><f>A1</f></c>')
    )
    result = analyze(
        tmp_path,
        grid(body),
        limits=replace(ImpactLimits(), max_evidence_edges=1) if cutoff else ImpactLimits(),
    )
    filename = (
        "dependency_edge_cutoff.expected.json" if cutoff else "dependency_complete.expected.json"
    )
    expected = json.loads((FIXTURES / filename).read_text())
    # Only the byte identity is variable. All graph/schema/work/cursor answers
    # in these files were written independently, never exported by the analyzer.
    expected["source_fingerprint"] = (
        "sha256:" + hashlib.sha256((tmp_path / "case.xlsx").read_bytes()).hexdigest()
    )
    assert result.to_dict() == expected


def test_json_depth64_and_cursor_validation() -> None:
    nested: object = 0
    for _ in range(64):
        nested = [nested]
    assert _freeze_json(nested) is not None
    with pytest.raises(ValueError, match="nesting"):
        _freeze_json([nested])
    with pytest.raises(ValueError):
        InventoryCursor("S", 0, "scan_cells", None)
    with pytest.raises(ValueError):
        TraversalCursor(CellRef("S", "A1"), 0, "admit_edge", None)
    with pytest.raises(ValueError):
        replace(example(), source_fingerprint="sha256:not-a-hash")
    with pytest.raises(ValueError):
        replace(example(), analysis_status="unknown")


@pytest.mark.parametrize("kind", ["array", "dataTable"])
def test_array_and_data_table_are_not_guessed(tmp_path: Path, kind: str) -> None:
    text = "A1*2" if kind == "array" else ""
    body = grid(f'<c r="B1"><f t="{kind}" ref="B1:B2">{text}</f></c>')
    result = analyze(tmp_path, body)
    assert result.analysis_status == "partial" and result.direct_dependents == ()
    assert result.issues[0].code == "unsupported_formula_kind"
    assert result.issues[0].details["ref"] == "B1:B2"
    assert result.work.formula_cells_attempted == 1


def test_budget_pairs_and_work_accounting(tmp_path: Path) -> None:
    one = grid('<c r="B1"><f>A1</f></c>')
    two = grid('<c r="B1"><f>A1</f></c><c r="C1"><f>A1</f></c>')
    for budget, counter in (
        ("max_scanned_cells", "scanned_cells"),
        ("max_formula_cells", "formula_cells_attempted"),
        ("max_references", "references_admitted"),
        ("max_evidence_edges", "evidence_edges_admitted"),
    ):
        limits = replace(ImpactLimits(), **{budget: 1})
        at = analyze(tmp_path, one, limits=limits)
        over = analyze(tmp_path, two, limits=limits)
        assert at.analysis_status == "complete" and over.analysis_status == "partial", budget
        assert getattr(at.work, counter) == getattr(over.work, counter) == 1
    limits = replace(ImpactLimits(), max_sheets=1)
    assert analyze(tmp_path, one, limits=limits).analysis_status == "complete"
    over = analyze(tmp_path, one, other="<sheetData/>", limits=limits)
    assert over.work.sheets_attempted == 1 and over.analysis_status == "partial"
    assert over.inventory_frontier == (InventoryCursor("T", 1, "unread_sheet", None),)
    for length in (2, 3):
        result = analyze(
            tmp_path,
            grid(f'<c r="B1"><f>{"A1" + " " * (length - 2)}</f></c>'),
            limits=replace(ImpactLimits(), max_formula_chars=2),
        )
        assert result.work.max_formula_chars_observed == length
        assert result.analysis_status == ("complete" if length == 2 else "partial")
    for depth in (1, 2):
        result = analyze(
            tmp_path,
            grid(f'<c r="B1"><f>{"(" * depth}A1{")" * depth}</f></c>'),
            limits=replace(ImpactLimits(), max_formula_nesting=1),
        )
        assert result.work.max_formula_nesting_observed == depth
        assert result.analysis_status == ("complete" if depth == 1 else "partial")
    # Only one rectangle check is required here: the root is outside its bounds
    # and no new node is reached, so an exact ceiling is still complete.
    at = analyze(
        tmp_path,
        grid('<c r="B1"><f>SUM(C1:C2)</f></c>'),
        limits=replace(ImpactLimits(), max_membership_checks=1),
    )
    over = analyze(
        tmp_path,
        grid('<c r="B1"><f>SUM(C1:C2)</f></c><c r="C1"><f>SUM(D1:D2)</f></c>'),
        limits=replace(ImpactLimits(), max_membership_checks=1),
    )
    assert at.work.membership_checks == over.work.membership_checks == 1
    assert at.analysis_status == "complete" and over.analysis_status == "partial"
    # Names and node ceilings have separate N/N+1 tests above; verify every
    # admitted counter against its corresponding bound on a mixed cutoff.
    for result in (at, over):
        for budget, counter in (
            ("max_sheets", "sheets_attempted"),
            ("max_scanned_cells", "scanned_cells"),
            ("max_formula_cells", "formula_cells_attempted"),
            ("max_defined_names", "defined_names_admitted"),
            ("max_references", "references_admitted"),
            ("max_membership_checks", "membership_checks"),
            ("max_visited_nodes", "visited_nodes"),
            ("max_evidence_edges", "evidence_edges_admitted"),
        ):
            assert getattr(result.work, counter) <= getattr(result.limits, budget)


def test_present_constants_count_and_atomic_reference_frontier(tmp_path: Path) -> None:
    body = grid('<c r="A1"><v>1</v></c><c r="B1"/><c r="C1"><f>A1+B1</f></c>')
    cutoff = analyze(tmp_path, body, limits=replace(ImpactLimits(), max_scanned_cells=2))
    assert cutoff.work.scanned_cells == 2 and cutoff.work.formula_cells_attempted == 0
    assert cutoff.inventory_frontier[0] == InventoryCursor("S", 0, "scan_cells", CellRef("S", "C1"))
    atomic = analyze(tmp_path, body, limits=replace(ImpactLimits(), max_references=1))
    assert atomic.work.references_admitted == 0 and atomic.evidence_edges == ()
    assert atomic.inventory_frontier[0].phase == "admit_formula"


def test_large_used_name_and_duplicate_tail_quarantine(tmp_path: Path) -> None:
    names = '<definedName name="ScopedRate">S!$A$1' + " " * 20 + "</definedName>"
    limited = analyze(
        tmp_path,
        grid('<c r="B1"><f>ScopedRate</f></c>'),
        names=names,
        limits=replace(ImpactLimits(), max_formula_chars=10),
    )
    assert limited.analysis_status == "partial" and limited.direct_dependents == ()
    assert limited.work.max_formula_chars_observed == 26
    duplicate = names.replace(" " * 20, "") + (
        '<definedName name="scopedrate" localSheetId="0">S!$A$2</definedName>'
    )
    withheld = analyze(
        tmp_path,
        grid('<c r="B1"><f>ScopedRate</f></c>'),
        names=duplicate,
        limits=replace(ImpactLimits(), max_defined_names=1),
    )
    assert withheld.work.defined_names_admitted == 0
    assert withheld.direct_dependents == ()
    full = analyze(tmp_path, grid('<c r="B1"><f>ScopedRate</f></c>'), names=duplicate, root="A2")
    assert full.direct_dependents == (CellRef("S", "B1"),)


def test_shift_mixed_locks_before_range_normalization_in_grid(tmp_path: Path) -> None:
    body = '<sheetData><row r="2"><c r="C2">'
    body += '<f t="shared" si="0" ref="C2:D3">SUM($B2:A$1)</f></c></row>'
    body += '<row r="3"><c r="D3"><f t="shared" si="0"/></c></row></sheetData>'
    result = analyze(tmp_path, body, root="B3")
    assert result.direct_dependents == (CellRef("S", "D3"),)
    evidence = result.evidence_edges[0].evidence
    assert evidence.normalized_reference == RangeRef("S", 1, 2, 3, 2)
    assert evidence.formula_text == "SUM($B2:A$1)" and evidence.reference_text == "$B2:A$1"
    assert evidence.shared_offset == (1, 1) and evidence.source_span == (4, 11)


def test_5000_cell_chain_is_iterative_and_has_linear_proof(tmp_path: Path) -> None:
    body = (
        "<sheetData>"
        + "".join(
            f'<row r="{row}"><c r="A{row}"><f>A{row - 1}</f></c></row>' for row in range(2, 5002)
        )
        + "</sheetData>"
    )
    result = analyze(tmp_path, body)
    assert result.analysis_status == "complete" and result.exact_total_count == 5000
    assert result.direct_dependents == (CellRef("S", "A2"),)
    assert len(result.transitive_dependents) == 4999
    assert result.work.visited_nodes == 5001 and result.work.references_admitted == 5000
    assert len(result.evidence_edges) == len(result.predecessor_edges) == 5000
    assert len(result.path_to(CellRef("S", "A5001"))) == 5000
    assert result.known_cycles == ()


def test_seeded_small_graphs_against_independent_adjacency(tmp_path: Path) -> None:
    rng = random.Random(80231)  # noqa: S311 — reproducible test inputs, not security randomness
    nodes = [f"{chr(65 + i)}1" for i in range(6)]
    for _ in range(20):
        adjacency = {node: set[str]() for node in nodes}
        formulas = []
        for dependent in nodes:
            sources = [node for node in nodes if rng.randrange(4) == 0]
            for source in sources:
                adjacency[source].add(dependent)
            text = "+".join(sources) or "1"
            formulas.append(f'<c r="{dependent}"><f>{text}</f></c>')
        result = analyze(tmp_path, grid("".join(formulas)))
        # Deliberately tiny fixed-point oracle, independent of production BFS/SCC.
        closure = {node: {node} for node in nodes}
        for _round in nodes:
            for node in nodes:
                closure[node].update(
                    child for reachable in tuple(closure[node]) for child in adjacency[reachable]
                )
        depths = {nodes[0]: 0}
        for _round in nodes:
            for source in tuple(depths):
                for target in adjacency[source]:
                    depths[target] = min(depths.get(target, 99), depths[source] + 1)
        assert result.direct_dependents == tuple(
            CellRef("S", n) for n in nodes if depths.get(n) == 1
        )
        assert result.transitive_dependents == tuple(
            CellRef("S", n) for n in nodes if depths.get(n, -1) >= 2
        )
        expected_edges = {(s, d) for s in closure[nodes[0]] for d in adjacency[s]}
        assert {
            (e.precedent.address, e.dependent.address) for e in result.evidence_edges
        } == expected_edges
        components = {
            tuple(n for n in nodes if n in closure[member] and member in closure[n])
            for member in closure[nodes[0]]
        }
        cycles = sorted(c for c in components if len(c) > 1 or c[0] in adjacency[c[0]])
        assert result.known_cycles == tuple(tuple(CellRef("S", n) for n in c) for c in cycles)
        for node, depth in depths.items():
            assert len(result.path_to(CellRef("S", node))) == depth


def test_seeded_rectangles_against_tiny_enumerated_oracle(tmp_path: Path) -> None:
    rng = random.Random(226)  # noqa: S311 — reproducible test inputs, not security randomness
    rectangles = []
    rows = []
    for row in range(1, 25):
        r1, r2 = sorted([rng.randint(1, 4), rng.randint(1, 4)])
        c1, c2 = sorted([rng.randint(1, 4), rng.randint(1, 4)])
        rectangles.append(
            {f"{chr(64 + c)}{r}" for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)}
        )
        rows.append(
            f'<row r="{row}"><c r="F{row}"><f>'
            f"SUM({chr(64 + c1)}{r1}:{chr(64 + c2)}{r2})</f></c></row>"
        )
    for r in range(1, 5):
        for c in range(1, 5):
            root = f"{chr(64 + c)}{r}"
            result = analyze(tmp_path, "<sheetData>" + "".join(rows) + "</sheetData>", root=root)
            expected = tuple(
                CellRef("S", f"F{i}") for i, cells in enumerate(rectangles, 1) if root in cells
            )
            assert result.direct_dependents == expected and result.transitive_dependents == ()
            assert result.exact_total_count == len(expected)


@pytest.mark.parametrize("defect", ["fingerprint", "sheet_order", "point_identity"])
def test_explicit_index_inconsistency_discards_all_proof(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    defect: str,
) -> None:
    import xlayer._dependencies as dependencies

    original = dependencies._build_index

    def inconsistent(
        registry: WorkbookRegistry,
        read_sheet: Callable[[str], Worksheet | Refusal],
        fingerprint: str,
        limits: ImpactLimits,
    ) -> dependencies._ReferenceIndex:
        index = original(registry, read_sheet, fingerprint, limits)
        if defect == "fingerprint":
            return replace(index, source_fingerprint=FINGERPRINT)
        if defect == "sheet_order":
            return replace(index, sheet_order={"S": 9})
        return replace(index, points={CellRef("S", "A2"): index.points[CellRef("S", "A1")]})

    monkeypatch.setattr(dependencies, "_build_index", inconsistent)
    result = analyze(tmp_path, grid('<c r="B1"><f>A1</f></c>'))
    assert result.analysis_status == "unknown" and result.evidence_edges == ()
    assert result.direct_dependents == result.transitive_dependents == ()
    assert (
        result.known_direct_count
        is result.known_transitive_count
        is result.exact_total_count
        is None
    )
    assert result.issues[0].code == "dependency_invariant_failure"


def test_cell_budget_does_not_examine_unprocessed_cell_payload(tmp_path: Path) -> None:
    body = '<sheetData><row r="1"><c r="A1"/></row><row r="2"><c r="A2"/></row></sheetData>'
    path = write_workbook_case(tmp_path / "case.xlsx", {"S": body})
    with Workbook.open(path) as book:
        worksheet = book.read_sheet("S")
        assert isinstance(worksheet, Worksheet)
        unprocessed = replace(worksheet.cells["A2"], row=1)
        injected = replace(worksheet, cells={"A1": worksheet.cells["A1"], "A2": unprocessed})
        impact = analyse_dependencies(
            book.registry,
            lambda _name: injected,
            CellRef("S", "A1"),
            book.source_fingerprint,
            replace(ImpactLimits(), max_scanned_cells=1),
        )
    assert impact.analysis_status == "partial" and impact.work.scanned_cells == 1
    assert impact.inventory_frontier == (InventoryCursor("S", 0, "scan_cells", CellRef("S", "A2")),)
    assert [issue.code for issue in impact.issues] == ["inventory_budget_exceeded"]


@pytest.mark.parametrize("formula", ["''!A1", "''!ScopedRate", "ScopedRate"])
def test_empty_sheet_qualifier_never_fabricates_local_edge(tmp_path: Path, formula: str) -> None:
    definition = "''!$A$1" if formula == "ScopedRate" else "S!$A$1"
    impact = analyze(
        tmp_path,
        grid(f'<c r="B1"><f>{formula}</f></c><c r="C1"><f>A1</f></c>'),
        names=f'<definedName name="ScopedRate">{definition}</definedName>',
    )
    assert impact.analysis_status == "partial"
    assert impact.direct_dependents == (CellRef("S", "C1"),)
    assert all(edge.dependent.address != "B1" for edge in impact.evidence_edges)


def test_rejected_fixed_definition_still_records_observed_chars(tmp_path: Path) -> None:
    impact = analyze(
        tmp_path,
        grid('<c r="B1"><f>ScopedRate</f></c>'),
        names='<definedName name="ScopedRate">SUM(S!$A$1)</definedName>',
    )
    assert impact.analysis_status == "partial" and impact.evidence_edges == ()
    assert impact.work.max_formula_chars_observed == 11


def test_later_name_failure_preserves_earlier_definition_work(tmp_path: Path) -> None:
    impact = analyze(
        tmp_path,
        grid('<c r="B1"><f>ScopedRate+Absent</f></c>'),
        names='<definedName name="ScopedRate">S!$A$1' + " " * 100 + "</definedName>",
    )
    assert impact.analysis_status == "partial" and impact.evidence_edges == ()
    assert impact.work.max_formula_chars_observed == 106


@pytest.mark.parametrize(
    "formula,depth", [("(((A1+INDEX(A1,1))))", 3), ("((A1+", 2), ("((A1))?", 2)]
)
def test_rejected_expression_preserves_observed_nesting(
    tmp_path: Path, formula: str, depth: int
) -> None:
    impact = analyze(tmp_path, grid(f'<c r="B1"><f>{formula}</f></c>'))
    assert impact.analysis_status == "partial" and impact.evidence_edges == ()
    assert impact.work.max_formula_nesting_observed == depth


@pytest.mark.parametrize("definition,depth", [("SUM(S!$A$1)", 1), ("SUM(SUM(S!$A$1))", 2)])
def test_rejected_definition_preserves_observed_nesting(
    tmp_path: Path,
    definition: str,
    depth: int,
) -> None:
    impact = analyze(
        tmp_path,
        grid('<c r="B1"><f>ScopedRate</f></c>'),
        names=f'<definedName name="ScopedRate">{definition}</definedName>',
    )
    assert impact.analysis_status == "partial" and impact.evidence_edges == ()
    assert impact.work.max_formula_nesting_observed == depth


def test_duplicate_index_occurrence_is_unknown_not_constructor_error(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import xlayer._dependencies as dependencies

    original = dependencies._build_index

    def duplicate(
        registry: WorkbookRegistry,
        read_sheet: Callable[[str], Worksheet | Refusal],
        fingerprint: str,
        limits: ImpactLimits,
    ) -> dependencies._ReferenceIndex:
        index = original(registry, read_sheet, fingerprint, limits)
        root = CellRef("S", "A1")
        return replace(
            index,
            points={root: index.points[root] * 2},
            work=replace(index.work, references_admitted=2),
        )

    monkeypatch.setattr(dependencies, "_build_index", duplicate)
    impact = analyze(tmp_path, grid('<c r="B1"><f>A1</f></c>'))
    assert impact.analysis_status == "unknown" and impact.evidence_edges == ()
    assert impact.exact_total_count is impact.known_direct_count is None
    assert impact.issues[-1].code == "dependency_invariant_failure"


@pytest.mark.parametrize("defect", ["wrong_depth", "outside_node_edge", "missing_retained_edge"])
def test_inconsistent_traversal_is_unknown_before_cycles_or_counts(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    defect: str,
) -> None:
    import xlayer._dependencies as dependencies

    original = dependencies._traverse

    def inconsistent(
        index: dependencies._ReferenceIndex,
        root: CellRef,
        limits: ImpactLimits,
    ) -> dependencies._Traversal:
        traversal = original(index, root, limits)
        dependent = CellRef("S", "B1")
        if defect == "wrong_depth":
            return replace(traversal, depths={root: 0, dependent: 2})
        if defect == "missing_retained_edge":
            return replace(traversal, edges=())
        edge = DependencyEdge(CellRef("S", "A2"), dependent, traversal.edges[0].evidence)
        return replace(traversal, edges=(edge,), predecessors={dependent: edge})

    monkeypatch.setattr(dependencies, "_traverse", inconsistent)
    impact = analyze(tmp_path, grid('<c r="B1"><f>SUM(A1:A2)</f></c>'))
    assert impact.analysis_status == "unknown" and impact.evidence_edges == ()
    assert impact.direct_dependents == impact.transitive_dependents == ()
    assert (
        impact.known_direct_count
        is impact.known_transitive_count
        is impact.exact_total_count
        is None
    )
    assert impact.issues[-1].code == "dependency_invariant_failure"
