"""Real ZIP end-to-end mutation with independent raw/structural oracles."""

from __future__ import annotations

import os
import stat
import zipfile
from dataclasses import replace
from pathlib import Path
from typing import TYPE_CHECKING
from xml.etree import ElementTree as ET

import pytest

from tests.unit._dependency_cases import write_workbook_case
from xlayer._approval import Approval, VerifiedApproval
from xlayer._edits import SetValue
from xlayer._errors import Refusal
from xlayer._preview import Preview
from xlayer._receipt import Receipt
from xlayer._workbook import Workbook

if TYPE_CHECKING:
    from xlayer._proposal import Proposal

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def workbook(
    tmp_path: Path, body: str = '<c r="A1"><v>100</v></c><c r="B1"><f>A1*2</f><v>200</v></c>'
) -> Path:
    return write_workbook_case(
        tmp_path / "source.xlsx", {"Inputs": f"<sheetData><row>{body}</row></sheetData>"}
    )


def prepared(
    book: Workbook, output: Path, *, edits: list[SetValue] | None = None, overwrite: bool = False
) -> tuple[Proposal, Preview, Approval]:
    assert hasattr(book, "propose")
    proposal = book.propose(
        edits or [SetValue("Inputs", "A1", 120)], output_path=output, overwrite=overwrite
    )
    assert not isinstance(proposal, Refusal)
    preview = proposal.preview()
    assert isinstance(preview, Preview)
    approval = Approval(
        "1.0",
        "approved-id",
        "trusted-host",
        "human",
        book.source_fingerprint,
        proposal.proposal_digest,
        preview.preview_digest,
        "2026-10-06T00:00:00Z",
    )
    return proposal, preview, approval


def host(approval: Approval) -> VerifiedApproval:
    return VerifiedApproval(
        approval,
        True,
        ("apply_set_value", "change_populated_text", "accept_partial_dependency_coverage"),
    )


def test_first_edit_full_verified_receipt_and_untouched_parts(tmp_path: Path) -> None:
    source = workbook(tmp_path)
    before = source.read_bytes()
    output = tmp_path / "out.xlsx"
    with Workbook.open(source) as book:
        proposal, preview, approval = prepared(book, output)
        assert not preview.evidence["blocked"]
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Receipt)
        assert result.evidence["recalculation_status"] == "required_not_run"
        assert result.evidence["visual_status"] == "not_evaluated"
        assert result.evidence["verification"]["all_passed"] is True  # type: ignore[index]
        assert result.canonical_json()
        sheet = book.read_sheet("Inputs")
        assert not isinstance(sheet, Refusal) and sheet.cells["A1"].raw == "100"
        replay = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(replay, Refusal) and replay.code == "proposal_already_applied"
    with zipfile.ZipFile(source) as old, zipfile.ZipFile(output) as new:
        assert old.namelist() == new.namelist()
        for part in old.namelist():
            if part == "xl/worksheets/1.xml":
                assert new.read(part) == old.read(part).replace(
                    b'<c r="A1">', b'<c r="A1" t="n">'
                ).replace(b"<v>100</v>", b"<v>120</v>")
            else:
                assert old.read(part) == new.read(part)
        root = ET.fromstring(new.read("xl/worksheets/1.xml"))  # noqa: S314
        assert root.find(f'.//{NS}c[@r="B1"]/{NS}v').text == "200"  # type: ignore[union-attr]
    assert source.read_bytes() == before
    if os.name == "posix":
        assert stat.S_IMODE(output.stat().st_mode) == 0o600
    assert not list(tmp_path.glob(".xlayer-*"))


@pytest.mark.parametrize("change", ["deleted", "changed", "replaced"])
def test_stale_source_never_publishes(tmp_path: Path, change: str) -> None:
    source = workbook(tmp_path)
    with Workbook.open(source) as book:
        proposal, _, approval = prepared(book, tmp_path / "out.xlsx")
        if change == "deleted":
            source.unlink()
        else:
            if change == "replaced":
                source.unlink()
            source.write_bytes(b"changed")
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "stale_source"
    assert not (tmp_path / "out.xlsx").exists()
    assert not list(tmp_path.glob(".xlayer-*"))


def test_invalid_batch_cannot_apply_safe_subset(tmp_path: Path) -> None:
    source = workbook(tmp_path)
    before = source.read_bytes()
    with Workbook.open(source) as book:
        proposal, _, approval = prepared(
            book,
            tmp_path / "out.xlsx",
            edits=[SetValue("Inputs", "A1", 120), SetValue("Inputs", "B1", 1)],
        )
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "formula_cell"
    assert source.read_bytes() == before and not (tmp_path / "out.xlsx").exists()


@pytest.mark.parametrize(
    "kind", ["missing", "wrong_digest", "denied", "model_text", "revoked", "changed_grants"]
)
def test_approval_cannot_be_forged_or_revoked_during_staging(tmp_path: Path, kind: str) -> None:
    source = workbook(tmp_path, '<c r="A1" t="inlineStr"><is><t>old</t></is></c>')
    with Workbook.open(source) as book:
        proposal, _, approval = prepared(
            book, tmp_path / "out.xlsx", edits=[SetValue("Inputs", "A1", "new")]
        )
        artifact = None if kind == "missing" else approval
        if kind == "wrong_digest":
            artifact = replace(approval, preview_digest="sha256:" + "0" * 64)
        if kind == "model_text":
            artifact = replace(approval, actor_kind="model")
        calls = 0

        def verifier(claim: Approval) -> VerifiedApproval:
            nonlocal calls
            calls += 1
            return VerifiedApproval(
                claim,
                not (kind == "denied" or (kind == "revoked" and calls == 2)),
                ("apply_set_value", "change_populated_text")
                if kind != "changed_grants" or calls == 1
                else ("apply_set_value", "change_populated_text", "extra"),
            )

        result = proposal.apply(approval=artifact, verify_approval=verifier)
        assert isinstance(result, Refusal)
        assert result.code in {"approval_required", "preview_digest_mismatch", "approval_rejected"}
    assert not (tmp_path / "out.xlsx").exists() and not list(tmp_path.glob(".xlayer-*"))


def test_unexpected_verifier_error_propagates_and_closes_session(tmp_path: Path) -> None:
    book = Workbook.open(workbook(tmp_path))
    proposal, _, approval = prepared(book, tmp_path / "out.xlsx")

    def explode(_approval: Approval) -> VerifiedApproval:
        raise RuntimeError("host failed")

    with pytest.raises(RuntimeError, match="host failed"):
        proposal.apply(approval=approval, verify_approval=explode)
    assert book.closed and not list(tmp_path.glob(".xlayer-*"))


def test_fresh_verifier_oserror_is_not_a_filesystem_refusal(tmp_path: Path) -> None:
    book = Workbook.open(workbook(tmp_path))
    proposal, _, approval = prepared(book, tmp_path / "out.xlsx")
    calls = 0

    def verifier(claim: Approval) -> VerifiedApproval:
        nonlocal calls
        calls += 1
        if calls == 2:
            raise OSError("approval store unexpectedly failed")
        return host(claim)

    with pytest.raises(OSError, match="approval store"):
        proposal.apply(approval=approval, verify_approval=verifier)
    assert book.closed and not (tmp_path / "out.xlsx").exists()
    assert not list(tmp_path.glob(".xlayer-*"))


def test_explicit_overwrite_and_changed_incumbent(tmp_path: Path) -> None:
    source = workbook(tmp_path)
    output = tmp_path / "out.xlsx"
    output.write_bytes(b"incumbent")
    with Workbook.open(source) as book:
        proposal, _, approval = prepared(book, output, overwrite=True)
        output.write_bytes(b"changed")
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "stale_output"
        assert output.read_bytes() == b"changed"
        proposal, _, approval = prepared(book, output, overwrite=True)
        assert isinstance(proposal.apply(approval=approval, verify_approval=host), Receipt)


def test_exclusive_publication_race_cannot_overwrite(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import xlayer._transaction as transaction

    output = tmp_path / "out.xlsx"
    real_link = os.link

    def race(src: str | os.PathLike[str], dst: str | os.PathLike[str]) -> None:
        output.write_bytes(b"competitor")
        real_link(src, dst)

    monkeypatch.setattr(transaction.os, "link", race)
    with Workbook.open(workbook(tmp_path)) as book:
        proposal, _, approval = prepared(book, output)
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "publication_failed"
    assert output.read_bytes() == b"competitor" and not list(tmp_path.glob(".xlayer-*"))


@pytest.mark.parametrize("phase", ["write", "fsync", "verify", "receipt", "link"])
def test_precommit_failure_leaves_source_and_incumbent_untouched(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, phase: str
) -> None:
    import xlayer._transaction as transaction

    source = workbook(tmp_path)
    before = source.read_bytes()
    output = tmp_path / "out.xlsx"
    if phase in {"write", "fsync", "link"}:
        target = transaction.os if phase in {"fsync", "link"} else transaction
        name = phase if phase in {"fsync", "link"} else "write_archive"

        def fail(*_args: object, **_kwargs: object) -> None:
            raise OSError("injected failure")

        monkeypatch.setattr(target, name, fail)
    elif phase == "verify":
        monkeypatch.setattr(
            transaction,
            "verify_output",
            lambda *_args: Refusal(
                "verification_failed", "injected", recovery_options=({"action": "inspect"},)
            ),
        )
    else:

        def bad_receipt(*_args: object, **_kwargs: object) -> Receipt:
            raise RuntimeError("receipt failed")

        monkeypatch.setattr(transaction, "Receipt", bad_receipt)
    book = Workbook.open(source)
    proposal, _, approval = prepared(book, output)
    if phase == "receipt":
        with pytest.raises(RuntimeError, match="receipt failed"):
            proposal.apply(approval=approval, verify_approval=host)
        assert book.closed
    else:
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal)
        book.close()
    assert (
        source.read_bytes() == before
        and not output.exists()
        and not list(tmp_path.glob(".xlayer-*"))
    )


def test_postcommit_temp_cleanup_failure_returns_committed_receipt(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import xlayer._transaction as transaction

    real_unlink = transaction.os.unlink

    def fail_temp(path: str | os.PathLike[str]) -> None:
        if Path(path).name.startswith(".xlayer-"):
            raise OSError("cleanup failed")
        real_unlink(path)

    with Workbook.open(workbook(tmp_path)) as book:
        proposal, _, approval = prepared(book, tmp_path / "out.xlsx")
        monkeypatch.setattr(transaction.os, "unlink", fail_temp)
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Receipt) and (tmp_path / "out.xlsx").exists()
        assert result.audit_metadata["cleanup_warning"]
        assert isinstance(proposal.apply(approval=approval, verify_approval=host), Refusal)
    for temp in tmp_path.glob(".xlayer-*"):
        real_unlink(temp)
