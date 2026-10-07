"""Private installed-artifact transaction smoke, not a supported public example."""

from __future__ import annotations

import json
from pathlib import Path

from xlayer._approval import Approval, VerifiedApproval
from xlayer._edits import SetValue
from xlayer._errors import Refusal
from xlayer._preview import Preview
from xlayer._receipt import Receipt
from xlayer._workbook import Workbook

FIXTURE = Path(__file__).resolve().parents[1] / "fixtures" / "test_workbook_7_write_v11_edges.xlsx"


def test_installed_private_transaction_previews_applies_reopens_returns_receipt(
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(FIXTURE.read_bytes())
    before = source.read_bytes()
    output = tmp_path / "output.xlsx"
    with Workbook.open(source) as book:
        proposal = book.propose([SetValue("edges", "A1", 120)], output_path=output)
        assert not isinstance(proposal, Refusal)
        preview = proposal.preview()
        assert isinstance(preview, Preview) and not preview.evidence["blocked"]
        approval = Approval(
            "1.0",
            "approval-id",
            "test-host",
            "human",
            book.source_fingerprint,
            proposal.proposal_digest,
            preview.preview_digest,
            "2026-10-06T00:00:00Z",
        )

        def verifier(claim: Approval) -> VerifiedApproval:
            return VerifiedApproval(
                claim, True, ("apply_set_value", "accept_partial_dependency_coverage")
            )

        receipt = proposal.apply(approval=approval, verify_approval=verifier)
        assert isinstance(receipt, Receipt)
        assert receipt.evidence["actual_changed_parts"] == ("xl/worksheets/sheet1.xml",)
        assert receipt.evidence["visual_status"] == "not_evaluated"
        payload = json.loads(receipt.canonical_json())
        assert payload["verification"]["all_passed"] is True
        assert payload["preview_digest_reference"] == preview.preview_digest
    with Workbook.open(output) as book:
        sheet = book.read_sheet("edges")
        assert not isinstance(sheet, Refusal) and sheet.cells["A1"].raw == "120"
        assert receipt.evidence["output_fingerprint"] == book.source_fingerprint
    assert source.read_bytes() == before
    assert not list(tmp_path.glob(".xlayer-*"))
