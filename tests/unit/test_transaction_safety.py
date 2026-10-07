"""Target ownership, filesystem identity, fidelity and mixed-batch boundaries."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from tests.unit._dependency_cases import write_workbook_case
from tests.unit.test_transaction import host, prepared, workbook
from xlayer._approval import Approval, VerifiedApproval
from xlayer._dependencies import ImpactLimits
from xlayer._edits import SetValue
from xlayer._errors import Refusal
from xlayer._preview import Preview
from xlayer._receipt import Receipt
from xlayer._workbook import Workbook

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
RELS = "http://schemas.openxmlformats.org/package/2006/relationships"


def rewrite(path: Path, parts: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path) as source:
        payload = {name: source.read(name) for name in source.namelist()}
    payload.update(parts)
    with zipfile.ZipFile(path, "w") as result:
        for name, data in payload.items():
            result.writestr(name, data)


@pytest.mark.parametrize(
    ("cell", "extras", "code"),
    [
        ("A1", '<mergeCells><mergeCell ref="A1:B1"/></mergeCells>', "merged_cell"),
        ("B1", '<mergeCells><mergeCell ref="A1:B1"/></mergeCells>', "merged_cell"),
        ("A1", '<sheetProtection sheet="1"/>', "protected_worksheet"),
        ("A1", "<dataValidations/>", "unsupported_target_structure"),
        ("A1", "<tableParts/>", "unsupported_target_structure"),
    ],
)
def test_unmodeled_ownership_blocks_even_empty_or_missing_cells(
    tmp_path: Path, cell: str, extras: str, code: str
) -> None:
    path = write_workbook_case(
        tmp_path / "input.xlsx",
        {"Inputs": '<sheetData><row><c r="A1"/></row></sheetData>' + extras},
    )
    with Workbook.open(path) as book:
        proposal = book.propose([SetValue("Inputs", cell, 2)], output_path=tmp_path / "out.xlsx")
        assert not isinstance(proposal, Refusal)
        preview = proposal.preview()
        assert isinstance(preview, Preview) and preview.evidence["blocked"]
        assert preview.to_dict()["findings"][0]["refusal"]["code"] == code  # type: ignore[index]


@pytest.mark.parametrize("kind", ["array", "dataTable"])
def test_formula_owned_non_master_scalar_refuses(tmp_path: Path, kind: str) -> None:
    expression = "1+1" if kind == "array" else ""
    path = workbook(
        tmp_path,
        f'<c r="A1"><f t="{kind}" ref="A1:B1">{expression}</f><v>2</v></c><c r="B1"><v>2</v></c>',
    )
    with Workbook.open(path) as book:
        proposal, preview, approval = prepared(
            book, tmp_path / "out.xlsx", edits=[SetValue("Inputs", "B1", 3)]
        )
        assert preview.evidence["blocked"]
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "formula_range_target"


def test_pivot_ownership_uses_sheet_relationship_part(tmp_path: Path) -> None:
    path = workbook(tmp_path)
    rewrite(
        path,
        {
            "xl/worksheets/_rels/1.xml.rels": (
                f'<Relationships xmlns="{RELS}"><Relationship Id="p" '
                f'Type="{REL}pivotTable" Target="../pivotTables/pivot.xml"/></Relationships>'
            ).encode()
        },
    )
    with Workbook.open(path) as book:
        proposal, preview, approval = prepared(book, tmp_path / "out.xlsx")
        assert preview.evidence["blocked"]
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "unsupported_target_structure"


def test_signed_package_surface_refuses_without_guessing_signature_paths(tmp_path: Path) -> None:
    path = workbook(tmp_path)
    with zipfile.ZipFile(path) as archive:
        content = archive.read("[Content_Types].xml")
    rewrite(
        path,
        {
            "[Content_Types].xml": content.replace(
                b"</Types>",
                b'<Override PartName="/custom/s.xml" '
                b'ContentType="application/vnd.openxmlformats-package.'
                b'digital-signature-xmlsignature+xml"/></Types>',
            )
        },
    )
    with Workbook.open(path) as book:
        proposal = book.propose([SetValue("Inputs", "A1", 2)], output_path=tmp_path / "out.xlsx")
        assert not isinstance(proposal, Refusal)
        result = proposal.preview()
        assert isinstance(result, Refusal) and result.code == "signed_package"


def test_shared_string_noop_keeps_index_when_another_cell_changes(tmp_path: Path) -> None:
    path = workbook(tmp_path, '<c r="A1"><v>100</v></c><c r="B1" t="s"><v>0</v></c>')
    with zipfile.ZipFile(path) as archive:
        rels = archive.read("xl/_rels/workbook.xml.rels")
    rewrite(
        path,
        {
            "xl/sharedStrings.xml": f'<sst xmlns="{MAIN}"><si><t>same</t></si></sst>'.encode(),
            "xl/_rels/workbook.xml.rels": rels.replace(
                b"</Relationships>",
                (
                    f'<Relationship Id="ss" Type="{REL}sharedStrings" '
                    'Target="sharedStrings.xml"/></Relationships>'
                ).encode(),
            ),
        },
    )
    with Workbook.open(path) as book:
        proposal, preview, approval = prepared(
            book,
            tmp_path / "out.xlsx",
            edits=[SetValue("Inputs", "A1", 120), SetValue("Inputs", "B1", "same")],
        )
        assert preview.to_dict()["edits"][1]["expected"]["stored_type"] == "s"  # type: ignore[index]
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Receipt)
    with zipfile.ZipFile(tmp_path / "out.xlsx") as output:
        assert b'<c r="B1" t="s"><v>0</v></c>' in output.read("xl/worksheets/1.xml")


def test_populated_text_change_requires_external_grants(tmp_path: Path) -> None:
    with Workbook.open(workbook(tmp_path, '<c r="A1" t="str"><v>old</v></c>')) as book:
        proposal, preview, approval = prepared(
            book, tmp_path / "out.xlsx", edits=[SetValue("Inputs", "A1", "")]
        )
        assert preview.evidence["next_action"] == "obtain_external_approval"
        denied = proposal.apply(
            approval=approval,
            verify_approval=lambda claim: VerifiedApproval(claim, True, ("apply_set_value",)),
        )
        assert isinstance(denied, Refusal) and denied.code == "approval_rejected"
        assert isinstance(proposal.apply(approval=approval, verify_approval=host), Receipt)


@pytest.mark.parametrize("value", [1, 1.0, True, "=1+1", "\r\n😀_x000D_", ""])
def test_existing_blank_roundtrip_retains_unestablished_receipt(
    tmp_path: Path, value: object
) -> None:
    with Workbook.open(workbook(tmp_path, '<c r="A1"/>')) as book:
        proposal, _, approval = prepared(
            book,
            tmp_path / "out.xlsx",
            edits=[SetValue("Inputs", "A1", value)],  # type: ignore[arg-type]
        )
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Receipt)
        assert result.to_dict()["edits"][0]["target_category"] == "unestablished"  # type: ignore[index]
    with Workbook.open(tmp_path / "out.xlsx") as book:
        sheet = book.read_sheet("Inputs")
        assert not isinstance(sheet, Refusal) and sheet.cells["A1"].stored_value == value
        assert sheet.cells["A1"].formula is None


def test_final_source_recheck_follows_fresh_host_decision(tmp_path: Path) -> None:
    path = workbook(tmp_path)
    calls = 0

    def verifier(claim: Approval) -> VerifiedApproval:
        nonlocal calls
        calls += 1
        if calls == 2:
            path.write_bytes(b"source replaced during host check")
        return host(claim)

    with Workbook.open(path) as book:
        proposal, _, approval = prepared(book, tmp_path / "out.xlsx")
        result = proposal.apply(approval=approval, verify_approval=verifier)
        assert isinstance(result, Refusal) and result.code == "stale_source"
    assert not (tmp_path / "out.xlsx").exists() and not list(tmp_path.glob(".xlayer-*"))


def test_changed_parent_directory_object_invalidates_preview(tmp_path: Path) -> None:
    output_parent = tmp_path / "outputs"
    output_parent.mkdir()
    with Workbook.open(workbook(tmp_path)) as book:
        proposal, _, approval = prepared(book, output_parent / "out.xlsx")
        output_parent.rename(tmp_path / "old-output-directory")
        output_parent.mkdir()
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "stale_output"
    assert not (output_parent / "out.xlsx").exists()


def test_missing_supported_worksheet_cannot_be_approved_away(tmp_path: Path) -> None:
    path = write_workbook_case(
        tmp_path / "input.xlsx",
        {
            "Inputs": '<sheetData><row><c r="A1"><v>1</v></c></row></sheetData>',
            "Broken": '<sheetData><row><c r="A1"><v>banana</v></c></row></sheetData>',
        },
    )
    with Workbook.open(path) as book:
        proposal, preview, approval = prepared(
            book, tmp_path / "out.xlsx", edits=[SetValue("Inputs", "A1", 2)]
        )
        assert "accept_partial_dependency_coverage" in preview.evidence["required_capabilities"]  # type: ignore[operator]
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "verification_failed"
    assert not (tmp_path / "out.xlsx").exists()


def test_multiple_sheet_batch_and_advisory_are_not_safety_overrides(tmp_path: Path) -> None:
    # Each cell's enclosing row must match; use one cell per row.
    data = (
        "<sheetData>"
        + "".join(f'<row r="{i}"><c r="A{i}"><v>1</v></c></row>' for i in range(1, 51))
        + "</sheetData>"
    )
    path = write_workbook_case(
        tmp_path / "input.xlsx",
        {"Inputs": data, "Other": '<sheetData><row><c r="A1"><v>1</v></c></row></sheetData>'},
    )
    with Workbook.open(path) as book:
        proposal = book.propose(
            [SetValue("Inputs", f"A{i}", 2) for i in range(1, 51)],
            output_path=tmp_path / "out.xlsx",
        )
        assert not isinstance(proposal, Refusal)
        preview = proposal.preview()
        assert isinstance(preview, Preview) and not preview.evidence["blocked"]
        assert (
            preview.evidence["breadth_advisory"]
            and preview.evidence["next_action"] == "inspect_known_impact"
        )
        assert preview.evidence["recalculation_status"] == "unknown"
        multiple = book.propose(
            [SetValue("Inputs", "A1", 2), SetValue("Other", "A1", 2)],
            output_path=tmp_path / "other.xlsx",
        )
        assert not isinstance(multiple, Refusal)
        result = multiple.preview()
        assert isinstance(result, Preview) and result.evidence["blocked"]


def test_partial_dependency_can_proceed_only_with_external_grant(tmp_path: Path) -> None:
    path = workbook(
        tmp_path, '<c r="A1"><v>100</v></c><c r="B1"><f>A1</f></c><c r="C1"><f>B1</f></c>'
    )
    with Workbook.open(path) as book:
        proposal = book.propose(
            [SetValue("Inputs", "A1", 2)],
            output_path=tmp_path / "out.xlsx",
            limits=ImpactLimits(max_visited_nodes=2),
        )
        assert not isinstance(proposal, Refusal)
        preview = proposal.preview()
        assert isinstance(preview, Preview) and not preview.evidence["blocked"]
        assert preview.evidence["next_action"] == "obtain_external_approval"
        approval = Approval(
            "1.0",
            "id",
            "host",
            "human",
            book.source_fingerprint,
            proposal.proposal_digest,
            preview.preview_digest,
            "now",
        )
        assert isinstance(proposal.apply(approval=approval, verify_approval=host), Receipt)
