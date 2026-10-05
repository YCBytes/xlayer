"""Private dependency smoke, not a supported public-API usage example."""

import os
from pathlib import Path

import pytest

import xlayer
from xlayer import _dependencies, _workbook
from xlayer._dependencies import CellRef, DependencyImpact
from xlayer._errors import ClosedWorkbookError
from xlayer._ooxml import formula
from xlayer._workbook import Workbook

FIXTURE = Path(__file__).resolve().parents[1] / "fixtures" / "test_workbook_4_worksheet_final.xlsx"


def test_installed_origin_enforced_when_requested() -> None:
    if os.environ.get("XLAYER_REQUIRE_INSTALLED") == "1":
        for module in (xlayer, formula, _dependencies, _workbook):
            assert module.__file__ is not None
            origin = Path(module.__file__).resolve()
            assert "site-packages" in origin.parts, (module.__name__, origin)


def test_installed_private_dependency_smoke() -> None:
    before = FIXTURE.read_bytes()
    with Workbook.open(FIXTURE) as book:
        result = book.dependency_impact("formulas", "E8")
        assert isinstance(result, DependencyImpact)
        assert result.analysis_status == "partial" and result.exact_total_count is None
        assert result.direct_dependents == (CellRef("formulas", "B8"),)
        evidence = result.evidence_edges[0].evidence
        assert evidence.formula_text == "E8*F8" and evidence.source_span == (0, 2)
        assert evidence.shared_master == CellRef("formulas", "B8")
        assert evidence.shared_offset == (0, 0)
        detached = result.to_dict()
    assert result.to_dict() == detached
    assert len(result.path_to(CellRef("formulas", "B8"))) == 1
    with pytest.raises(ClosedWorkbookError):
        book.dependency_impact("formulas", "E8")
    assert FIXTURE.read_bytes() == before
