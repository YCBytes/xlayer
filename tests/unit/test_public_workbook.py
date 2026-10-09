"""Public operations exercise raw input cases and detached observations."""

from __future__ import annotations

from pathlib import Path

import pytest

import xlayer
from tests.unit._dependency_cases import write_workbook_case


def test_public_inventory_and_read_are_bounded_snapshot_facts(tmp_path: Path) -> None:
    source = write_workbook_case(
        tmp_path / "source.xlsx",
        {
            "Inputs": '<sheetData><row><c r="A1"><v>12</v></c>'
            '<c r="B1"><f>A1*2</f><v>24</v></c></row></sheetData>'
        },
    )
    with xlayer.Workbook.open(source) as book:
        inventory = book.list_sheets()
        assert isinstance(inventory, xlayer.Inspection)
        assert inventory.to_dict()["data"] == {
            "sheets": [{"name": "Inputs", "state": "visible", "kind": "worksheet", "tab_index": 0}],
            "active_tab": 0,
            "date1904": False,
            "defined_name_count": 0,
        }
        result = book.read_cells("Inputs", ["b1", "A1", "C1"])
        assert isinstance(result, xlayer.Inspection)
        data = result.to_dict()["data"]
        assert isinstance(data, dict)
        assert data["requested_addresses"] == ["B1", "A1", "C1"]
        cells = data["cells"]
        assert isinstance(cells, list)
        assert cells[0]["stored_value"] == 24.0
        assert cells[0]["formula_result_status"] == "not_verified"
        assert cells[0]["value_source"] == "saved_formula_cache"
        assert cells[1]["raw"] == "12"
        assert cells[2]["presence"] == "missing"
        assert not hasattr(book, "registry") and not hasattr(book, "read_sheet")
        assert not hasattr(book, "archive")
    assert book.closed and result.to_dict()["source_fingerprint"] == book.source_fingerprint
    with pytest.raises(xlayer.ClosedWorkbookError):
        book.read_cells("Inputs", ["A1"])


def test_public_proposal_only_retains_public_edits(tmp_path: Path) -> None:
    source = write_workbook_case(
        tmp_path / "s.xlsx", {"Inputs": '<sheetData><row><c r="A1"><v>1</v></c></row></sheetData>'}
    )
    with xlayer.Workbook.open(source) as book:
        edit = xlayer.SetValue("Inputs", "A1", 2)
        proposal = book.propose([edit], output_path=tmp_path / "out.xlsx")
        assert isinstance(proposal, xlayer.Proposal)
        assert proposal.edits == (edit,)
        assert proposal.source_fingerprint == book.source_fingerprint
        assert isinstance(proposal.preview(), xlayer.Preview)
        assert not (tmp_path / "out.xlsx").exists()
