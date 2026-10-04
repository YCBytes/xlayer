"""Internal lifecycle smoke for both clean-installed distribution formats.

This is not a public API example or promotion: imports stay underscore-private.
The existing sanitized fixture ships in the sdist/checkout, never the wheel.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from xlayer._errors import ClosedWorkbookError, Refusal
from xlayer._ooxml.sheet import Worksheet
from xlayer._workbook import Workbook

FIXTURE = Path(__file__).resolve().parents[1] / "fixtures" / "test_workbook_7_write_v11_edges.xlsx"


def test_installed_internal_workbook_opens_reads_and_closes() -> None:
    before = FIXTURE.read_bytes()
    with Workbook.open(FIXTURE) as book:
        assert [entry.name for entry in book.registry.sheets] == ["edges"]
        sheet = book.read_sheet("edges")
        assert isinstance(sheet, Worksheet)
        assert sheet.cells["A1"].raw == "100"
        assert sheet.cells["B1"].stored_value == "label"
        assert sheet.cells["D1"].style.temporal_kind == "date"
        assert "C10" not in sheet.cells
        merge = sheet.merge_at("C10")
        assert merge is not None
        assert merge.ref == "B10:C10"
        assert book.read_sheet("edges") is sheet
        refused = book.read_sheet("missing")
        assert isinstance(refused, Refusal)
        assert refused.code == "sheet_not_found"
        assert refused.recovery_options
        fingerprint = book.source_fingerprint
    assert book.closed
    assert book.source_fingerprint == fingerprint
    assert sheet.cells["A1"].raw == "100"
    with pytest.raises(ClosedWorkbookError):
        book.read_sheet("edges")
    with pytest.raises(ClosedWorkbookError):
        _ = book.registry
    book.close()
    assert FIXTURE.read_bytes() == before
