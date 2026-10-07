"""Shared transaction ceilings must not multiply by the number of edited roots."""

from __future__ import annotations

import importlib.util
from pathlib import Path
from typing import TYPE_CHECKING

from tests.unit._dependency_cases import write_workbook_case
from xlayer._dependencies import CellRef, ImpactLimits
from xlayer._workbook import Workbook

if TYPE_CHECKING:
    from xlayer._batch_impact import BatchImpact


def batch(tmp_path: Path, limits: ImpactLimits) -> BatchImpact:
    assert importlib.util.find_spec("xlayer._batch_impact") is not None, "batch analysis missing"
    from xlayer._batch_impact import analyse_batch

    path = write_workbook_case(
        tmp_path / "chain.xlsx",
        {
            "Inputs": '<sheetData><row><c r="A1"><v>1</v></c>'
            '<c r="B1"><f>A1</f><v>1</v></c><c r="C1"><f>B1</f><v>1</v></c></row></sheetData>'
        },
    )
    with Workbook.open(path) as book:
        return analyse_batch(
            book.registry,
            book.read_sheet,
            (CellRef("Inputs", "A1"), CellRef("Inputs", "B1")),
            book.source_fingerprint,
            limits,
        )


def test_one_inventory_local_work_and_deduplicated_union(tmp_path: Path) -> None:
    result = batch(tmp_path, ImpactLimits())
    assert result.inventory_work.scanned_cells == 3
    assert result.inventory_work.formula_cells_attempted == 2
    assert result.work["visited_nodes"] == 5
    assert result.work["evidence_edges_admitted"] == 3
    assert result.known_union == (CellRef("Inputs", "B1"), CellRef("Inputs", "C1"))
    assert result.roots[0].impact is not None
    assert result.roots[0].impact.work.scanned_cells == 0
    assert result.roots[0].impact.work.visited_nodes == 3
    assert result.roots[1].impact is not None
    assert result.roots[1].impact.work.visited_nodes == 2


def test_exhausted_initial_admission_is_not_started_not_complete_zero(tmp_path: Path) -> None:
    result = batch(tmp_path, ImpactLimits(max_visited_nodes=3))
    assert result.roots[0].impact is not None
    assert result.roots[0].impact.analysis_status == "complete"
    assert result.roots[1].state == "not_started"
    assert result.roots[1].impact is None
    assert result.roots[1].blocked_budget == "max_visited_nodes"
    assert result.roots[1].root_cursor == 1


def test_started_traversal_retains_partial_lower_bound(tmp_path: Path) -> None:
    result = batch(tmp_path, ImpactLimits(max_visited_nodes=4))
    assert result.roots[1].state == "started"
    impact = result.roots[1].impact
    assert impact is not None and impact.analysis_status == "partial"
    assert impact.known_direct_count == 0 and impact.exact_total_count is None
    assert impact.traversal_frontier
    assert result.work["visited_nodes"] == 4


def test_exactly_finished_work_is_complete(tmp_path: Path) -> None:
    result = batch(tmp_path, ImpactLimits(max_visited_nodes=5, max_evidence_edges=3))
    assert all(
        root.impact is not None and root.impact.analysis_status == "complete"
        for root in result.roots
    )
