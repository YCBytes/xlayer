"""Demonstrated review findings, using independent input and corruption probes."""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.unit.test_transaction import host, prepared, workbook
from xlayer._dependencies import ImpactLimits
from xlayer._edits import SetValue
from xlayer._errors import Refusal, WorkbookOpenError
from xlayer._preview import Preview
from xlayer._receipt import Receipt
from xlayer._workbook import Workbook


@pytest.mark.parametrize(
    "cell",
    [
        '<c r="A1" ><v>100</v></c>',
        '<c r="A1" />',
        '<c r="&#65;1"><v>100</v></c>',
        "<c r=\"A1\"  t = 'n'  ><v>100</v></c>",
        '<c r="A1" xmlns:unused="urn:example:t=\'not-an-attribute\'" ><v>100</v></c>',
    ],
)
def test_valid_lexical_coordinates_and_spacing_complete(tmp_path: Path, cell: str) -> None:
    source = workbook(tmp_path, cell)
    before = source.read_bytes()
    with Workbook.open(source) as book:
        proposal, preview, approval = prepared(
            book, tmp_path / "out.xlsx", edits=[SetValue("Inputs", "A1", 2)]
        )
        assert not preview.evidence["blocked"]
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Receipt)
    assert source.read_bytes() == before
    with Workbook.open(tmp_path / "out.xlsx") as output:
        sheet = output.read_sheet("Inputs")
        assert not isinstance(sheet, Refusal) and sheet.cells["A1"].raw == "2"


def test_raw_decimal_outside_supported_comparison_is_an_operational_refusal(
    tmp_path: Path,
) -> None:
    with Workbook.open(workbook(tmp_path, '<c r="A1"><v>1e-9223372036854775809</v></c>')) as book:
        sheet = book.read_sheet("Inputs")
        assert not isinstance(sheet, Refusal) and sheet.cells["A1"].stored_value == 0.0
        proposal = book.propose([SetValue("Inputs", "A1", 2)], output_path=tmp_path / "out.xlsx")
        assert not isinstance(proposal, Refusal)
        preview = proposal.preview()
        assert isinstance(preview, Preview) and preview.evidence["blocked"]
        assert preview.to_dict()["findings"][0]["refusal"]["code"] == "invalid_value"  # type: ignore[index]
        assert not book.closed
    assert not (tmp_path / "out.xlsx").exists()


def test_symlink_loop_retains_workbook_open_error_boundary(tmp_path: Path) -> None:
    loop = tmp_path / "loop.xlsx"
    try:
        loop.symlink_to(loop.name)
    except OSError:
        pytest.skip("symlinks unavailable")
    with pytest.raises(WorkbookOpenError) as caught:
        Workbook.open(loop)
    assert caught.value.refusal.code == "invalid_path"


@pytest.mark.parametrize("space_owner", ["is", "t"])
def test_existing_space_metadata_cannot_hide_new_rich_payload(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, space_owner: str
) -> None:
    import xlayer._transaction as transaction

    original = transaction.write_archive

    def corrupt(*args: object) -> object:
        values = list(args)
        payload = values[-1]
        assert isinstance(payload, bytes)
        plain = b'<t xml:space="preserve">new</t>'
        assert plain in payload
        values[-1] = payload.replace(plain, b"<r><rPr><b/></rPr><t>new</t></r>")
        return original(*values)  # type: ignore[arg-type]

    inner = '<t xml:space="preserve">old</t>' if space_owner == "t" else "<t>old</t>"
    attributes = ' xml:space="preserve"' if space_owner == "is" else ""
    source = workbook(tmp_path, f'<c r="A1" t="inlineStr"><is{attributes}>{inner}</is></c>')
    before = source.read_bytes()
    with Workbook.open(source) as book:
        proposal, preview, approval = prepared(
            book, tmp_path / "out.xlsx", edits=[SetValue("Inputs", "A1", "new")]
        )
        assert not preview.evidence["blocked"]
        monkeypatch.setattr(transaction, "write_archive", corrupt)
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "verification_failed"
    assert source.read_bytes() == before and not (tmp_path / "out.xlsx").exists()
    assert not list(tmp_path.glob(".xlayer-*"))


@pytest.mark.parametrize(
    "payload",
    ['<is custom="original"><t>old</t></is>', '<is><t custom="original">old</t></is>'],
)
def test_unmodeled_text_attributes_are_not_writable(tmp_path: Path, payload: str) -> None:
    with Workbook.open(workbook(tmp_path, f'<c r="A1" t="inlineStr">{payload}</c>')) as book:
        proposal, preview, approval = prepared(
            book, tmp_path / "out.xlsx", edits=[SetValue("Inputs", "A1", "new")]
        )
        assert preview.evidence["blocked"]
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "unsupported_target_structure"


def test_unknown_inventory_counts_root_admission_without_inventing_proof(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from dataclasses import replace

    import xlayer._batch_impact as module
    from xlayer._dependencies import _build_index

    def inconsistent(*args: object) -> object:
        index = _build_index(*args)  # type: ignore[arg-type]
        return replace(index, source_fingerprint="sha256:" + "0" * 64)

    monkeypatch.setattr(module, "_build_index", inconsistent)
    with Workbook.open(workbook(tmp_path, '<c r="A1"><v>1</v></c><c r="B1"><v>1</v></c>')) as book:
        proposal = book.propose(
            [SetValue("Inputs", "A1", 2), SetValue("Inputs", "B1", 2)],
            output_path=tmp_path / "out.xlsx",
            limits=ImpactLimits(max_visited_nodes=1),
        )
        assert not isinstance(proposal, Refusal)
        preview = proposal.preview()
        assert isinstance(preview, Preview) and preview.evidence["blocked"]
        body = preview.to_dict()
        dependency = body["dependencies"]
        assert dependency["traversal_work"]["visited_nodes"] == 1  # type: ignore[index]
        roots = dependency["roots"]  # type: ignore[index]
        assert roots[0]["impact"]["known_direct_count"] is None
        assert roots[0]["impact"]["work"]["visited_nodes"] == 1
        assert roots[1]["state"] == "not_started" and roots[1]["impact"] is None
        assert body["measurements"]["known_transitive_union"] is None  # type: ignore[index]
        assert body["measurements"]["known_affected_sheets"] is None  # type: ignore[index]


def test_breadth_policy_uses_the_thresholds_it_records(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import xlayer._proposal as module

    monkeypatch.setattr(
        module,
        "THRESHOLDS",
        {
            "known_transitive_union_gt": 0,
            "known_affected_sheets_gt": 3,
            "batch_size_gte": 50,
        },
    )
    path = workbook(
        tmp_path, '<c r="A1"><v>1</v></c><c r="B1"><f>A1</f></c><c r="C1"><f>B1</f></c>'
    )
    with Workbook.open(path) as book:
        _, preview, _ = prepared(book, tmp_path / "out.xlsx")
        assert preview.evidence["thresholds"]["known_transitive_union_gt"] == 0  # type: ignore[index]
        assert preview.evidence["breadth_advisory"] is True
        assert preview.evidence["next_action"] == "inspect_known_impact"


@pytest.mark.parametrize("field", ["version", "threshold"])
def test_changed_policy_cannot_reuse_prior_approval(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, field: str
) -> None:
    import xlayer._proposal as module

    with Workbook.open(workbook(tmp_path)) as book:
        proposal, preview, approval = prepared(book, tmp_path / "out.xlsx")
        if field == "version":
            monkeypatch.setattr(module, "VERSIONS", {**module.VERSIONS, "policy_version": "2.0"})
        else:
            monkeypatch.setattr(module, "THRESHOLDS", {**module.THRESHOLDS, "batch_size_gte": 49})
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "preview_digest_mismatch"
        assert result.details["preview_digest"] == preview.preview_digest
        assert not (tmp_path / "out.xlsx").exists()
