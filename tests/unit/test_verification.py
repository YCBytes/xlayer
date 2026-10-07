"""Independent verifier probes: successful writer assertions are not an oracle."""

from __future__ import annotations

import copy
import zipfile
from pathlib import Path
from typing import BinaryIO, cast

import pytest

from tests.unit.test_transaction import host, prepared, workbook
from tests.unit.test_transaction_safety import rewrite
from xlayer._edits import SetValue
from xlayer._errors import Refusal
from xlayer._workbook import Workbook


@pytest.mark.parametrize(
    ("old", "new"),
    [
        (b"<v>200</v>", b"<v>201</v>"),
        (b"A1*2", b"A1*3"),
        (b'<c r="B1">', b'<c r="B1" s="0">'),
        (b"<row>", b'<row hidden="1">'),
        (b'<c r="A1" t="n">', b'<c r="A1" t="n" s="0">'),
        (b"<v>120</v>", b"<v>121</v>"),
        (b"<v>120</v>", b'<v>120</v><evil:v xmlns:evil="urn:unmodeled"/>'),
    ],
)
def test_independent_checks_refuse_undeclared_changes(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, old: bytes, new: bytes
) -> None:
    import xlayer._transaction as transaction

    original = transaction.write_archive

    def corrupt(*args: object) -> object:
        captured = list(args)
        patched = captured[-1]
        assert isinstance(patched, bytes) and old in patched
        captured[-1] = patched.replace(old, new)
        return original(*captured)  # type: ignore[arg-type]

    source = workbook(tmp_path)
    before = source.read_bytes()
    with Workbook.open(source) as book:
        proposal, _, approval = prepared(book, tmp_path / "out.xlsx")
        monkeypatch.setattr(transaction, "write_archive", corrupt)
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "verification_failed"
    assert source.read_bytes() == before and not (tmp_path / "out.xlsx").exists()
    assert not list(tmp_path.glob(".xlayer-*"))


def test_same_decoded_text_with_new_rich_markup_is_not_verified(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import xlayer._transaction as transaction

    original = transaction.write_archive

    def corrupt(*args: object) -> object:
        captured = list(args)
        payload = captured[-1]
        assert isinstance(payload, bytes)
        captured[-1] = payload.replace(
            b'<t xml:space="preserve">new</t>', b"<r><rPr><b/></rPr><t>new</t></r>"
        )
        return original(*captured)  # type: ignore[arg-type]

    with Workbook.open(
        workbook(tmp_path, '<c r="A1" t="inlineStr"><is><t>old</t></is></c>')
    ) as book:
        proposal, _, approval = prepared(
            book, tmp_path / "out.xlsx", edits=[SetValue("Inputs", "A1", "new")]
        )
        monkeypatch.setattr(transaction, "write_archive", corrupt)
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "verification_failed"
    assert not (tmp_path / "out.xlsx").exists()


@pytest.mark.parametrize("part", ["xl/worksheets/1.xml", "xl/_rels/workbook.xml.rels"])
def test_non_target_merge_and_relationship_corruption_refuse(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, part: str
) -> None:
    import xlayer._transaction as transaction

    source = workbook(tmp_path)
    with zipfile.ZipFile(source) as archive:
        worksheet = archive.read("xl/worksheets/1.xml")
    rewrite(
        source,
        {
            "xl/worksheets/1.xml": worksheet.replace(
                b"</worksheet>", b'<mergeCells><mergeCell ref="C1:D1"/></mergeCells></worksheet>'
            )
        },
    )
    before = source.read_bytes()
    original = transaction.write_archive

    def corrupt(*args: object) -> object:
        result = original(*args)  # type: ignore[arg-type]
        handle = cast("BinaryIO", args[1])
        with zipfile.ZipFile(handle) as archive:
            payloads = [(info, archive.read(info)) for info in archive.infolist()]
        handle.seek(0)
        handle.truncate(0)
        with zipfile.ZipFile(handle, "w", allowZip64=False) as archive:
            for info, payload in payloads:
                if info.filename == part:
                    payload = (
                        payload.replace(b"C1:D1", b"C1:E1")
                        if part.endswith("1.xml")
                        else payload.replace(
                            b"<Relationship ", b'<Relationship TargetMode="Internal" ', 1
                        )
                    )
                archive.writestr(copy.copy(info), payload)
        return result

    with Workbook.open(source) as book:
        proposal, _, approval = prepared(book, tmp_path / "out.xlsx")
        monkeypatch.setattr(transaction, "write_archive", corrupt)
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "verification_failed"
    assert source.read_bytes() == before and not (tmp_path / "out.xlsx").exists()
    assert not list(tmp_path.glob(".xlayer-*"))
