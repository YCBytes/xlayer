"""Retain every owned handle so garbage collection cannot counterfeit cleanup."""

from __future__ import annotations

import os
import tempfile
import zipfile
from pathlib import Path
from types import TracebackType
from typing import BinaryIO, cast

import pytest

from tests.unit.test_transaction import host, prepared, workbook
from xlayer._approval import Approval, VerifiedApproval
from xlayer._errors import Refusal
from xlayer._ooxml.archive import DEFAULT_LIMITS, ArchiveLimits, WorkbookArchive
from xlayer._receipt import Receipt
from xlayer._workbook import Workbook


class _FlushFailure:
    def __init__(self, stream: BinaryIO) -> None:
        self.stream = stream

    def __getattr__(self, name: str) -> object:
        return getattr(self.stream, name)

    def __enter__(self) -> _FlushFailure:
        return self

    def __exit__(
        self,
        _kind: type[BaseException] | None,
        _error: BaseException | None,
        _traceback: TracebackType | None,
    ) -> None:
        self.stream.close()

    def flush(self) -> None:
        raise OSError("flush failed")


@pytest.mark.parametrize(
    "phase",
    [
        "success",
        "mkstemp",
        "fdopen",
        "write",
        "flush",
        "fsync",
        "reopen",
        "compare",
        "receipt",
        "final_host",
        "publication",
    ],
)
def test_retained_handles_close_at_every_transaction_boundary(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, phase: str
) -> None:
    import xlayer._transaction as transaction
    import xlayer._verification as verification
    import xlayer._workbook as coordinator

    source = workbook(tmp_path)
    before = source.read_bytes()
    book = Workbook.open(source)
    proposal, _, approval = prepared(book, tmp_path / "out.xlsx")
    original_archive = book._archive
    assert original_archive is not None
    source_zip = original_archive._zip
    descriptors: list[int] = []
    streams: list[BinaryIO] = []
    reopened: list[zipfile.ZipFile] = []
    real_mkstemp, real_fdopen = tempfile.mkstemp, os.fdopen
    real_load = WorkbookArchive.load

    def tracked_temp(*, prefix: str, suffix: str, dir: Path) -> tuple[int, str]:
        if phase == "mkstemp":
            raise OSError("staging failed")
        result = real_mkstemp(prefix=prefix, suffix=suffix, dir=dir)
        descriptors.append(result[0])
        return result

    def tracked_stream(descriptor: int, mode: str) -> BinaryIO:
        if phase == "fdopen":
            raise OSError("wrapping failed")
        stream = cast("BinaryIO", real_fdopen(descriptor, mode))
        streams.append(stream)
        return cast("BinaryIO", _FlushFailure(stream)) if phase == "flush" else stream

    def tracked_load(
        _cls: type[WorkbookArchive], path: Path, limits: ArchiveLimits = DEFAULT_LIMITS
    ) -> WorkbookArchive | Refusal:
        result = real_load(path, limits)
        if isinstance(result, WorkbookArchive):
            reopened.append(result._zip)
        return result

    def unexpected(*_args: object, **_kwargs: object) -> None:
        raise RuntimeError(f"{phase} failed")

    def filesystem_failure(*_args: object, **_kwargs: object) -> None:
        raise OSError(f"{phase} failed")

    monkeypatch.setattr(tempfile, "mkstemp", tracked_temp)
    monkeypatch.setattr(transaction.os, "fdopen", tracked_stream)
    monkeypatch.setattr(WorkbookArchive, "load", classmethod(tracked_load))
    if phase == "write":
        monkeypatch.setattr(transaction, "write_archive", unexpected)
    elif phase == "fsync":
        monkeypatch.setattr(transaction.os, "fsync", filesystem_failure)
    elif phase == "reopen":
        monkeypatch.setattr(
            coordinator,
            "parse_styles",
            lambda *_args: Refusal(
                "malformed_xml",
                "injected reopen failure",
                recovery_options=({"action": "inspect"},),
            ),
        )
    elif phase == "compare":
        monkeypatch.setattr(verification, "_target_structure", unexpected)
    elif phase == "receipt":
        monkeypatch.setattr(transaction, "Receipt", unexpected)
    elif phase == "publication":
        monkeypatch.setattr(transaction.os, "link", filesystem_failure)
    calls = 0

    def verifier(claim: Approval) -> VerifiedApproval:
        nonlocal calls
        calls += 1
        if phase == "final_host" and calls == 2:
            raise RuntimeError("final_host failed")
        return host(claim)

    try:
        if phase in {"write", "compare", "receipt", "final_host"}:
            with pytest.raises(RuntimeError, match=f"{phase} failed"):
                proposal.apply(approval=approval, verify_approval=verifier)
            assert book.closed and source_zip.fp is None
        else:
            result = proposal.apply(approval=approval, verify_approval=verifier)
            assert isinstance(result, Receipt if phase == "success" else Refusal)
            assert not book.closed and source_zip.fp is not None
        assert all(stream.closed for stream in streams)
        assert all(handle.fp is None for handle in reopened)
        for descriptor in descriptors:
            with pytest.raises(OSError):
                os.fstat(descriptor)
        if phase in {"reopen", "compare", "receipt", "final_host", "publication", "success"}:
            assert reopened  # a retained fresh ZIP, not just a vacuous all()
        assert source.read_bytes() == before
        assert (tmp_path / "out.xlsx").exists() == (phase == "success")
        assert not list(tmp_path.glob(".xlayer-*"))
    finally:
        book.close()
    assert source_zip.fp is None


def test_failed_cleanup_annotation_cannot_reverse_a_commit(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import xlayer._transaction as transaction

    original_receipt, original_unlink = Receipt, os.unlink
    receipt_calls = 0

    def receipt(*args: object, **kwargs: object) -> Receipt:
        nonlocal receipt_calls
        receipt_calls += 1
        if receipt_calls == 2:
            raise RuntimeError("audit annotation failed")
        return original_receipt(*args, **kwargs)  # type: ignore[arg-type]

    def no_unlink(path: str | os.PathLike[str]) -> None:
        if Path(path).name.startswith(".xlayer-"):
            raise OSError("cleanup failed")
        original_unlink(path)

    with Workbook.open(workbook(tmp_path)) as book:
        proposal, _, approval = prepared(book, tmp_path / "out.xlsx")
        monkeypatch.setattr(transaction, "Receipt", receipt)
        monkeypatch.setattr(transaction.os, "unlink", no_unlink)
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Receipt) and receipt_calls == 2
        assert result is proposal._receipt and (tmp_path / "out.xlsx").exists()
        repeat = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(repeat, Refusal) and repeat.code == "proposal_already_applied"
        assert repeat.details["receipt_digest"] == result.receipt_digest
    for temp in tmp_path.glob(".xlayer-*"):
        original_unlink(temp)


def test_precommit_cleanup_failure_does_not_mask_the_authoritative_refusal(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import xlayer._transaction as transaction

    original_unlink, original_fdopen = os.unlink, os.fdopen
    streams: list[BinaryIO] = []

    def tracked_stream(descriptor: int, mode: str) -> BinaryIO:
        stream = cast("BinaryIO", original_fdopen(descriptor, mode))
        streams.append(stream)
        return stream

    def no_unlink(_path: object) -> None:
        raise OSError("cleanup failed")

    with Workbook.open(workbook(tmp_path)) as book:
        proposal, _, approval = prepared(book, tmp_path / "out.xlsx")
        monkeypatch.setattr(transaction.os, "fdopen", tracked_stream)
        monkeypatch.setattr(transaction.os, "unlink", no_unlink)
        monkeypatch.setattr(
            transaction,
            "verify_output",
            lambda *_args: Refusal(
                "verification_failed",
                "authoritative failure",
                recovery_options=({"action": "inspect"},),
            ),
        )
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "verification_failed"
        assert not (tmp_path / "out.xlsx").exists()
        assert streams and all(stream.closed for stream in streams)
    for temp in tmp_path.glob(".xlayer-*"):
        original_unlink(temp)
