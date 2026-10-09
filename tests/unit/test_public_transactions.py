"""Public-only end-to-end writes and independent bytes/XML/authority oracles."""

from __future__ import annotations

import os
import zipfile
from collections.abc import Callable, Iterator, Mapping, Sequence
from dataclasses import fields, replace
from pathlib import Path
from typing import overload
from xml.etree import ElementTree as ET

import pytest

from tests.unit._dependency_cases import write_workbook_case
from tests.unit.test_public_inspections import cells, source
from xlayer import (
    Approval,
    ClosedWorkbookError,
    ImpactLimits,
    Preview,
    Proposal,
    Receipt,
    Refusal,
    SetValue,
    VerifiedApproval,
    Workbook,
)

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def prepared(
    book: Workbook, output: Path, edits: list[SetValue], *, overwrite: bool = False
) -> tuple[Proposal, Preview, Approval]:
    proposed = book.propose(edits, output_path=output, overwrite=overwrite)
    assert isinstance(proposed, Proposal)
    preview = proposed.preview()
    assert isinstance(preview, Preview)
    approved = Approval(
        "1.0",
        "approval-1",
        "test-host",
        "human",
        book.source_fingerprint,
        proposed.proposal_digest,
        preview.preview_digest,
        "2026-01-01T00:00:00Z",
    )
    return proposed, preview, approved


def verifier(
    approved: Approval, capabilities: tuple[str, ...] = ("apply_set_value",)
) -> Callable[[Approval], VerifiedApproval]:
    def verify(claim: Approval) -> VerifiedApproval:
        return VerifiedApproval(claim, claim == approved, capabilities if claim == approved else ())

    return verify


@pytest.mark.parametrize(
    "value", [2, -1.25, True, False, "", "=1+1", "😀", "_x000D_", "a\x00b", " leading "]
)
def test_public_scalar_round_trip_exact_literal_and_one_part(
    tmp_path: Path, value: int | float | bool | str
) -> None:
    path = source(tmp_path, '<c r="A1"/>')
    before = path.read_bytes()
    output = tmp_path / "out.xlsx"
    with Workbook.open(path) as book:
        proposal, preview, approved = prepared(book, output, [SetValue("Inputs", "A1", value)])
        result = proposal.apply(approval=approved, verify_approval=verifier(approved))
        assert isinstance(result, Receipt), result
        assert result.evidence["visual_status"] == "not_evaluated"
        assert result.evidence["package_version"] == "0.1.0a1"
        assert result.evidence["transaction_contract_version"] == "1.2"
        assert preview.preview_digest == result.evidence["preview_digest_reference"]
        assert cells(book.read_cells("Inputs", ["A1"]))[0]["stored_value"] is None
        repeated = proposal.apply(approval=approved, verify_approval=verifier(approved))
        assert isinstance(repeated, Refusal) and repeated.code == "proposal_already_applied"
    with Workbook.open(output) as result_book:
        actual = cells(result_book.read_cells("Inputs", ["A1"]))[0]
        assert actual["stored_value"] == value
        if type(value) is bool:
            assert type(actual["stored_value"]) is bool
        if type(value) is str:
            assert actual["stored_type"] == "inlineStr" and actual["formula"] is None
    with zipfile.ZipFile(path) as old, zipfile.ZipFile(output) as new:
        assert old.namelist() == new.namelist()
        assert [n for n in old.namelist() if old.read(n) != new.read(n)] == ["xl/worksheets/1.xml"]
        root = ET.fromstring(new.read("xl/worksheets/1.xml"))  # noqa: S314
        cell = root.find(f'.//{NS}c[@r="A1"]')
        assert cell is not None and cell.find(f"{NS}f") is None
        if type(value) is bool:
            assert cell.attrib["t"] == "b" and cell.find(f"{NS}v").text == ("1" if value else "0")  # type: ignore[union-attr]
        elif type(value) is str:
            assert cell.attrib["t"] == "inlineStr" and cell.find(f"{NS}is/{NS}t") is not None
        else:
            assert cell.attrib["t"] == "n" and float(cell.find(f"{NS}v").text or "") == value  # type: ignore[union-attr]
    assert path.read_bytes() == before and not list(tmp_path.glob(".xlayer-*"))


@pytest.mark.parametrize("value", [120, 120.5, True, "=literal text"])
def test_public_and_private_workbook_output_bytes_are_identical(
    tmp_path: Path, value: int | float | bool | str
) -> None:
    from xlayer._edits import SetValue as EngineEdit
    from xlayer._workbook import Workbook as EngineWorkbook

    input_cell = (
        '<c r="A1"/>'
        if type(value) is str
        else '<c r="A1" t="b"><v>0</v></c>'
        if type(value) is bool
        else '<c r="A1"><v>100</v></c>'
    )
    path = source(tmp_path, input_cell + '<c r="B1"><f>A1*2</f><v>200</v></c>')
    private_out, public_out = tmp_path / "private.xlsx", tmp_path / "public.xlsx"
    with EngineWorkbook.open(path) as private:
        proposed = private.propose([EngineEdit("Inputs", "A1", value)], output_path=private_out)
        assert not isinstance(proposed, Refusal)
        preview = proposed.preview()
        assert isinstance(preview, Preview)
        approval = Approval(
            "1.0",
            "private",
            "test-host",
            "human",
            private.source_fingerprint,
            proposed.proposal_digest,
            preview.preview_digest,
            "2026-01-01T00:00:00Z",
        )
        assert isinstance(
            proposed.apply(approval=approval, verify_approval=verifier(approval)), Receipt
        )
    with Workbook.open(path) as public:
        public_proposal, preview, approval = prepared(
            public, public_out, [SetValue("Inputs", "A1", value)]
        )
        result = public_proposal.apply(approval=approval, verify_approval=verifier(approval))
        assert isinstance(result, Receipt)
        assert result.evidence["recalculation_status"] == "required_not_run"
        assert result.evidence["visual_status"] == "not_evaluated"
    assert public_out.read_bytes() == private_out.read_bytes()


@pytest.mark.parametrize(
    "target,extra,expected",
    [
        ("B1", "", "formula_cell"),
        ("C1", "", "missing_cell"),
        ("A1", '<mergeCells><mergeCell ref="A1:C1"/></mergeCells>', "merged_cell"),
        ("A1", '<sheetProtection sheet="1"/>', "protected_worksheet"),
        ("A1", "<dataValidations/>", "unsupported_target_structure"),
    ],
)
def test_invalid_member_cannot_apply_safe_subset(
    tmp_path: Path, target: str, extra: str, expected: str
) -> None:
    path = source(
        tmp_path,
        '<c r="A1"><v>1</v></c><c r="B1"><f>A1</f><v>1</v></c><c r="D1"><v>1</v></c>',
        extra=extra,
    )
    before = path.read_bytes()
    with Workbook.open(path) as book:
        proposal, preview, approved = prepared(
            book,
            tmp_path / "out.xlsx",
            [SetValue("Inputs", "D1", 2), SetValue("Inputs", target, 2)],
        )
        assert preview.evidence["blocked"] is True
        result = proposal.apply(approval=approved, verify_approval=verifier(approved))
        assert isinstance(result, Refusal) and result.code == expected
    assert before == path.read_bytes() and not (tmp_path / "out.xlsx").exists()


def test_duplicate_multi_sheet_and_noop_remain_refusals(tmp_path: Path) -> None:
    path = write_workbook_case(
        tmp_path / "s.xlsx",
        {
            "Inputs": '<sheetData><row><c r="A1"><v>1</v></c></row></sheetData>',
            "Other": '<sheetData><row><c r="A1"><v>1</v></c></row></sheetData>',
        },
    )
    with Workbook.open(path) as book:
        duplicate = book.propose(
            [SetValue("Inputs", "a1", 2), SetValue("Inputs", "A1", 3)],
            output_path=tmp_path / "out.xlsx",
        )
        assert isinstance(duplicate, Refusal) and duplicate.code == "duplicate_target"
        proposed = book.propose([SetValue("Inputs", "A1", 1)], output_path=tmp_path / "out.xlsx")
        assert isinstance(proposed, Proposal)
        result = proposed.preview()
        assert isinstance(result, Refusal) and result.code == "no_changes"
        proposed = book.propose(
            [SetValue("Inputs", "A1", 2), SetValue("Other", "A1", 2)],
            output_path=tmp_path / "out.xlsx",
        )
        assert isinstance(proposed, Proposal)
        result = proposed.preview()
        assert isinstance(result, Preview) and result.evidence["blocked"] is True
    assert not (tmp_path / "out.xlsx").exists()


@pytest.mark.parametrize("mode", ["missing", "digest", "denied", "revoked", "changed_grants"])
def test_public_apply_requires_exact_fresh_host_authority(tmp_path: Path, mode: str) -> None:
    path = source(tmp_path, '<c r="A1"><v>1</v></c>')
    before = path.read_bytes()
    with Workbook.open(path) as book:
        proposal, _, approved = prepared(book, tmp_path / "out.xlsx", [SetValue("Inputs", "A1", 2)])
        calls = 0

        def verify(claim: Approval) -> VerifiedApproval:
            nonlocal calls
            calls += 1
            return VerifiedApproval(
                claim,
                claim == approved and mode != "denied" and not (mode == "revoked" and calls == 2),
                ("apply_set_value", "unexpected")
                if mode == "changed_grants" and calls == 2
                else ("apply_set_value",),
            )

        claim = (
            None
            if mode == "missing"
            else replace(approved, preview_digest="sha256:" + "0" * 64)
            if mode == "digest"
            else approved
        )
        result = proposal.apply(approval=claim, verify_approval=verify)
        assert isinstance(result, Refusal) and result.code in {
            "approval_required",
            "preview_digest_mismatch",
            "approval_rejected",
        }
        assert not book.closed
    assert (
        path.read_bytes() == before
        and not (tmp_path / "out.xlsx").exists()
        and not list(tmp_path.glob(".xlayer-*"))
    )


@pytest.mark.parametrize("kind", ["text", "partial"])
def test_public_changes_do_not_infer_elevated_grants(tmp_path: Path, kind: str) -> None:
    text = '<c r="A1" t="inlineStr"><is><t>old</t></is></c>'
    numeric = '<c r="A1"><v>1</v></c><c r="B1"><f>INDIRECT("A1")</f><v>1</v></c>'
    with Workbook.open(source(tmp_path, text if kind == "text" else numeric)) as book:
        proposal, preview, approved = prepared(
            book, tmp_path / "out.xlsx", [SetValue("Inputs", "A1", "new" if kind == "text" else 2)]
        )
        required = (
            "change_populated_text" if kind == "text" else "accept_partial_dependency_coverage"
        )
        assert required in preview.evidence["required_capabilities"]  # type: ignore[operator]
        result = proposal.apply(approval=approved, verify_approval=verifier(approved))
        assert isinstance(result, Refusal) and result.code == "approval_rejected"
        actor = replace(approved, actor_kind="model")
        result = proposal.apply(
            approval=actor, verify_approval=verifier(actor, ("apply_set_value", required))
        )
        assert isinstance(result, Refusal) and result.code == "approval_rejected"
    assert not (tmp_path / "out.xlsx").exists()


@pytest.mark.parametrize("change", ["source", "output", "parent"])
def test_public_bound_identity_change_never_publishes(tmp_path: Path, change: str) -> None:
    path = source(tmp_path, '<c r="A1"><v>1</v></c>')
    parent = tmp_path / "outputs"
    parent.mkdir()
    output = parent / "out.xlsx"
    output.write_bytes(b"incumbent")
    with Workbook.open(path) as book:
        proposal, _, approved = prepared(
            book, output, [SetValue("Inputs", "A1", 2)], overwrite=True
        )
        if change == "source":
            path.write_bytes(b"replacement")
        elif change == "output":
            output.write_bytes(b"changed")
        else:
            parent.rename(tmp_path / "old")
            parent.mkdir()
        result = proposal.apply(approval=approved, verify_approval=verifier(approved))
        assert isinstance(result, Refusal) and result.code == (
            "stale_source" if change == "source" else "stale_output"
        )
    assert not list(parent.glob(".xlayer-*"))
    if change != "parent":
        assert output.read_bytes() == (b"changed" if change == "output" else b"incumbent")
    else:
        assert not output.exists()


def test_public_output_binding_never_allows_source_alias_or_silent_overwrite(
    tmp_path: Path,
) -> None:
    path = source(tmp_path, '<c r="A1"><v>1</v></c>')
    output = tmp_path / "out.xlsx"
    output.write_bytes(b"incumbent")
    with Workbook.open(path) as book:
        for destination, code in [
            (path, "unsafe_output_path"),
            (output, "output_exists"),
            (tmp_path / "missing" / "out.xlsx", "output_unavailable"),
        ]:
            result = book.propose([SetValue("Inputs", "A1", 2)], output_path=destination)
            assert isinstance(result, Refusal) and result.code == code
    assert output.read_bytes() == b"incumbent" and not (tmp_path / "missing").exists()


@pytest.mark.parametrize(
    "failure", [RuntimeError("host failure"), ValueError("host failure"), OSError("host failure")]
)
def test_public_verifier_error_closes_and_propagates(
    tmp_path: Path, failure: BaseException
) -> None:
    book = Workbook.open(source(tmp_path, '<c r="A1"><v>1</v></c>'))
    proposed, _, approved = prepared(book, tmp_path / "out.xlsx", [SetValue("Inputs", "A1", 2)])

    def broken(_claim: Approval) -> VerifiedApproval:
        raise failure

    with pytest.raises(type(failure)) as caught:
        proposed.apply(approval=approved, verify_approval=broken)
    assert caught.value is failure and book.closed
    assert not (tmp_path / "out.xlsx").exists() and not list(tmp_path.glob(".xlayer-*"))


@pytest.mark.parametrize("phase", ["verify", "publish"])
def test_public_faults_keep_incumbent_and_cleanup(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, phase: str
) -> None:
    from xlayer import _transaction

    path = source(tmp_path, '<c r="A1"><v>1</v></c>')
    output = tmp_path / "out.xlsx"
    output.write_bytes(b"incumbent")
    before = path.read_bytes()
    if phase == "verify":

        def refuse(*_args: object) -> Refusal:
            return Refusal(
                "verification_failed", "injected", recovery_options=({"action": "inspect"},)
            )

        monkeypatch.setattr(_transaction, "verify_output", refuse)
    else:

        def fail(*_args: object) -> None:
            raise OSError("injected publication failure")

        monkeypatch.setattr(_transaction.os, "replace", fail)
    with Workbook.open(path) as book:
        proposed, _, approved = prepared(
            book, output, [SetValue("Inputs", "A1", 2)], overwrite=True
        )
        result = proposed.apply(approval=approved, verify_approval=verifier(approved))
        assert isinstance(result, Refusal) and result.code == (
            "verification_failed" if phase == "verify" else "publication_failed"
        )
    assert (
        path.read_bytes() == before
        and output.read_bytes() == b"incumbent"
        and not list(tmp_path.glob(".xlayer-*"))
    )


@pytest.mark.parametrize(
    "failure",
    [RuntimeError("engine defect"), TypeError("engine defect"), ValueError("engine defect")],
)
def test_unexpected_proposal_construction_failure_closes(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, failure: BaseException
) -> None:
    book = Workbook.open(source(tmp_path, ""))

    def broken(*_args: object, **_kwargs: object) -> None:
        raise failure

    monkeypatch.setattr(book._book, "propose", broken)
    with pytest.raises(type(failure)) as caught:
        book.propose([SetValue("Inputs", "A1", 2)], output_path=tmp_path / "out.xlsx")
    assert caught.value is failure and book.closed


def test_huge_transaction_identity_returns_existing_bound_refusal(tmp_path: Path) -> None:
    with Workbook.open(source(tmp_path, "")) as book:
        # Each scalar is within its own bound; the aggregate canonical identity is not.
        edits = [SetValue("Inputs", f"A{i}", "😀" * 16383) for i in range(1, 101)]
        result = book.propose(edits, output_path=tmp_path / "out.xlsx")
        assert isinstance(result, Refusal) and result.code == "transaction_limit_exceeded"
        assert not book.closed


def test_public_proposal_captures_inputs_and_closed_cache_fails(tmp_path: Path) -> None:
    book = Workbook.open(source(tmp_path, '<c r="A1"><v>1</v></c>'))
    edits = [SetValue("Inputs", "A1", 2)]
    annotations: dict[str, object] = {"host": ["before"]}
    proposed = book.propose(edits, output_path=tmp_path / "out.xlsx", annotations=annotations)
    assert isinstance(proposed, Proposal)
    edits.clear()
    annotations["host"] = ["after"]
    assert proposed.edits == (SetValue("Inputs", "A1", 2),)
    preview = proposed.preview()
    assert isinstance(preview, Preview)
    book.close()
    with pytest.raises(ClosedWorkbookError):
        proposed.preview()
    with pytest.raises(ClosedWorkbookError):
        proposed.apply(
            approval=None,
            verify_approval=verifier(
                Approval(
                    "1.0",
                    "dummy",
                    "host",
                    "human",
                    book.source_fingerprint,
                    proposed.proposal_digest,
                    preview.preview_digest,
                    "now",
                )
            ),
        )


def test_public_propose_argument_errors_keep_session_healthy(tmp_path: Path) -> None:
    from xlayer._edits import SetValue as EngineEdit

    with Workbook.open(source(tmp_path, "")) as book:
        with pytest.raises(TypeError):
            book.propose([EngineEdit("Inputs", "A1", 2)], output_path=tmp_path / "out.xlsx")  # type: ignore[list-item]
        with pytest.raises(ValueError):
            book.propose([], output_path=tmp_path / "out.xlsx")
        with pytest.raises(TypeError):
            book.propose(
                [SetValue("Inputs", "A1", 2)],
                output_path=tmp_path / "out.xlsx",
                overwrite=1,  # type: ignore[arg-type]
            )
        assert not book.closed
        for spec in fields(ImpactLimits()):
            raised = replace(ImpactLimits(), **{spec.name: getattr(ImpactLimits(), spec.name) + 1})
            with pytest.raises(ValueError, match=spec.name):
                book.propose(
                    [SetValue("Inputs", "A1", 2)], output_path=tmp_path / "out.xlsx", limits=raised
                )
        assert not book.closed


class BrokenPath(os.PathLike[str]):
    def __fspath__(self) -> str:
        raise RuntimeError("path capture failed")


class BrokenMetadata(Mapping[str, object]):
    def __init__(
        self, failure: BaseException | None = None, *, fail_during: str = "iteration"
    ) -> None:
        self.failure = RuntimeError("metadata capture failed") if failure is None else failure
        self.fail_during = fail_during

    def __len__(self) -> int:
        return 1

    def __iter__(self) -> Iterator[str]:
        if self.fail_during == "iteration":
            raise self.failure
        return iter(("note",))

    def __getitem__(self, key: str) -> object:
        raise self.failure


class BrokenMetadataSequence(Sequence[object]):
    def __init__(self, failure: BaseException) -> None:
        self.failure = failure

    def __len__(self) -> int:
        return 1

    @overload
    def __getitem__(self, index: int) -> object: ...

    @overload
    def __getitem__(self, index: slice) -> Sequence[object]: ...

    def __getitem__(self, index: int | slice) -> object:
        raise self.failure


@pytest.mark.parametrize("boundary", ["path", "metadata"])
def test_unexpected_host_input_capture_failure_closes(tmp_path: Path, boundary: str) -> None:
    book = Workbook.open(source(tmp_path, ""))
    with pytest.raises(RuntimeError, match="capture failed"):
        book.propose(
            [SetValue("Inputs", "A1", 2)],
            output_path=BrokenPath() if boundary == "path" else tmp_path / "out.xlsx",
            annotations=BrokenMetadata() if boundary == "metadata" else None,
        )
    assert book.closed


@pytest.mark.parametrize("error_type", [TypeError, ValueError])
@pytest.mark.parametrize("fail_during", ["iteration", "lookup"])
@pytest.mark.parametrize("location", ["root", "nested_mapping", "nested_list"])
def test_unexpected_annotation_type_or_value_error_closes_unchanged(
    tmp_path: Path, error_type: type[Exception], fail_during: str, location: str
) -> None:
    failure = error_type("application annotation capture failed")
    broken = BrokenMetadata(failure, fail_during=fail_during)
    metadata = (
        broken
        if location == "root"
        else {"nested": broken}
        if location == "nested_mapping"
        else {"nested": [broken]}
    )
    path = source(tmp_path, '<c r="A1"/>')
    original = path.read_bytes()
    output = tmp_path / "out.xlsx"
    book = Workbook.open(path)
    with pytest.raises(error_type) as caught:
        book.propose([SetValue("Inputs", "A1", 2)], output_path=output, annotations=metadata)
    assert caught.value is failure
    assert book.closed
    assert path.read_bytes() == original and not output.exists()
    assert not list(tmp_path.glob(".xlayer-*"))


@pytest.mark.parametrize("error_type", [TypeError, ValueError])
def test_unexpected_nested_annotation_sequence_error_closes_unchanged(
    tmp_path: Path, error_type: type[Exception]
) -> None:
    failure = error_type("application sequence capture failed")
    book = Workbook.open(source(tmp_path, '<c r="A1"/>'))
    output = tmp_path / "out.xlsx"
    with pytest.raises(error_type) as caught:
        book.propose(
            [SetValue("Inputs", "A1", 2)],
            output_path=output,
            annotations={"nested": BrokenMetadataSequence(failure)},
        )
    assert caught.value is failure
    assert book.closed and not output.exists()


@pytest.mark.parametrize(
    ("metadata", "error_type"),
    [
        (True, TypeError),
        ({"value": object()}, TypeError),
        ({1: "value"}, TypeError),
        ({"value": b"bytes"}, TypeError),
        ({"value": "\ud800"}, ValueError),
        ({"value": float("inf")}, ValueError),
        ({"value": 2**63}, ValueError),
    ],
)
def test_deliberate_annotation_domain_error_keeps_session_healthy(
    tmp_path: Path, metadata: object, error_type: type[Exception]
) -> None:
    output = tmp_path / "out.xlsx"
    with Workbook.open(source(tmp_path, '<c r="A1"/>')) as book:
        with pytest.raises(error_type):
            book.propose(
                [SetValue("Inputs", "A1", 2)],
                output_path=output,
                annotations=metadata,  # type: ignore[arg-type]
            )
        assert not book.closed and not output.exists()
        assert cells(book.read_cells("Inputs", ["A1"]))[0]["presence"] == "present"
