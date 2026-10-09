"""Existing public workbooks, reproducible evidence and publication faults."""

from __future__ import annotations

import hashlib
import os
import subprocess
import sys
import zipfile
from datetime import date
from pathlib import Path
from typing import BinaryIO

import pytest

from tests.unit.test_scalar_patch import style_xml
from tests.unit.test_transaction import host, prepared, workbook
from tests.unit.test_transaction_safety import REL, rewrite
from xlayer._dependencies import CellRef, ImpactLimits
from xlayer._edits import SetValue
from xlayer._errors import Refusal
from xlayer._ooxml.archive import WorkbookArchive
from xlayer._receipt import Receipt
from xlayer._workbook import Workbook

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


@pytest.mark.parametrize(
    ("fixture", "sheet", "cell", "value"),
    [
        ("test_workbook_2_registry.xlsx", "USC - Hedging", "B1", 1000),
        ("test_workbook_3_formats.xlsx", "Formats and Strings", "A1", "updated label"),
        ("test_workbook_4_worksheet_final.xlsx", "single_cell", "A1", 43),
        ("test_workbook_6b_1904.xlsx", "Date 1904 Test", "A1", "updated label"),
        ("test_workbook_7_write_v11_edges.xlsx", "edges", "A1", 120),
    ],
)
def test_all_existing_public_workbook_fixtures_roundtrip(
    tmp_path: Path, fixture: str, sheet: str, cell: str, value: object
) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes((FIXTURES / fixture).read_bytes())
    before = source.read_bytes()
    with Workbook.open(source) as book:
        proposal, preview, approval = prepared(
            book,
            tmp_path / "out.xlsx",
            edits=[SetValue(sheet, cell, value)],  # type: ignore[arg-type]
        )
        assert not preview.evidence["blocked"]
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Receipt)
        assert len(result.evidence["actual_changed_parts"]) == 1  # type: ignore[arg-type]
    with Workbook.open(tmp_path / "out.xlsx") as book:
        inspected = book.read_sheet(sheet)
        assert not isinstance(inspected, Refusal) and inspected.cells[cell].stored_value == value
    assert source.read_bytes() == before


@pytest.mark.parametrize(
    ("old", "requested", "date1904", "stored_type", "expected"),
    [
        ("59", date(1900, 3, 1), False, "n", "61"),
        ("0", date(1904, 1, 2), True, "n", "1"),
        ("2024-01-01", date(2024, 2, 29), False, "d", "2024-02-29"),
    ],
)
def test_date_writes_verify_exact_serial_or_iso(
    tmp_path: Path, old: str, requested: date, date1904: bool, stored_type: str, expected: str
) -> None:
    path = workbook(tmp_path, f'<c r="A1" t="{stored_type}"><v>{old}</v></c>')
    with zipfile.ZipFile(path) as archive:
        rels = archive.read("xl/_rels/workbook.xml.rels")
        registry = archive.read("xl/workbook.xml")
    rewrite(
        path,
        {
            "xl/styles.xml": style_xml('<xf numFmtId="14"/>').encode(),
            "xl/_rels/workbook.xml.rels": rels.replace(
                b"</Relationships>",
                (
                    f'<Relationship Id="s" Type="{REL}styles" Target="styles.xml"/></Relationships>'
                ).encode(),
            ),
            "xl/workbook.xml": registry.replace(b"<sheets>", b'<workbookPr date1904="1"/><sheets>')
            if date1904
            else registry,
        },
    )
    with Workbook.open(path) as book:
        proposal, _, approval = prepared(
            book, tmp_path / "out.xlsx", edits=[SetValue("Inputs", "A1", requested)]
        )
        assert isinstance(proposal.apply(approval=approval, verify_approval=host), Receipt)
    with zipfile.ZipFile(tmp_path / "out.xlsx") as output:
        assert f"<v>{expected}</v>".encode() in output.read("xl/worksheets/1.xml")


def test_canonical_receipt_reproducible_in_fresh_processes(tmp_path: Path) -> None:
    source = workbook(tmp_path)
    output = tmp_path / "out.xlsx"
    script = """
import hashlib,sys
from pathlib import Path
from xlayer._workbook import Workbook
from xlayer._edits import SetValue
from xlayer._approval import Approval,VerifiedApproval
from xlayer._receipt import Receipt
with Workbook.open(sys.argv[1]) as book:
 p=book.propose([SetValue("Inputs","A1",120)],output_path=sys.argv[2],
  annotations={k:k for k in {"z","a"}})
 v=p.preview()
 a=Approval("1.0","id","host","human",book.source_fingerprint,
  p.proposal_digest,v.preview_digest,"now")
 r=p.apply(approval=a,verify_approval=lambda a:VerifiedApproval(a,True,("apply_set_value",)))
 assert isinstance(r,Receipt),r
 print(hashlib.sha256(r.canonical_json()).hexdigest())
Path(sys.argv[2]).unlink()
"""
    hashes = []
    for seed in ("1", "17", "999"):
        run = subprocess.run(  # noqa: S603 - fixed interpreter and test program
            [sys.executable, "-c", script, str(source), str(output)],
            env={**os.environ, "PYTHONHASHSEED": seed},
            capture_output=True,
            text=True,
            check=True,
        )
        hashes.append(run.stdout.strip())
    assert len(set(hashes)) == 1
    assert not output.exists()


def test_apply_after_cwd_change_uses_captured_absolute_bindings(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    workbook(tmp_path)
    monkeypatch.chdir(tmp_path)
    with Workbook.open("source.xlsx") as book:
        proposal, _, approval = prepared(book, Path("out.xlsx"))
        monkeypatch.chdir(tmp_path.parent)
        assert isinstance(proposal.apply(approval=approval, verify_approval=host), Receipt)
    assert (tmp_path / "out.xlsx").exists()


def test_replace_failure_preserves_incumbent_and_retry_remains_possible(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import xlayer._transaction as transaction

    source = workbook(tmp_path)
    output = tmp_path / "out.xlsx"
    output.write_bytes(b"incumbent")
    real_replace = os.replace

    def fail(_source: str | os.PathLike[str], _destination: str | os.PathLike[str]) -> None:
        raise OSError("replace failed")

    with Workbook.open(source) as book:
        proposal, _, approval = prepared(book, output, overwrite=True)
        monkeypatch.setattr(transaction.os, "replace", fail)
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and output.read_bytes() == b"incumbent"
        monkeypatch.setattr(transaction.os, "replace", real_replace)
        assert isinstance(proposal.apply(approval=approval, verify_approval=host), Receipt)
    assert not list(tmp_path.glob(".xlayer-*"))


def test_opaque_part_corruption_is_detected_independently(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import xlayer._transaction as transaction

    path = workbook(tmp_path)
    rewrite(path, {"custom/opaque.bin": b"preserve these exact bytes"})
    original = transaction.write_archive

    def corrupt(
        archive: WorkbookArchive, handle: BinaryIO, part: str, patched: bytes
    ) -> Refusal | None:
        result = original(archive, handle, part, patched)
        # The preservation writer no longer writes opaque content from the
        # decompressed cache. Inject actual staged-output corruption instead.
        handle.seek(0)
        with zipfile.ZipFile(handle) as output:
            entries = [(info, output.read(info)) for info in output.infolist()]
        handle.seek(0)
        handle.truncate()
        with zipfile.ZipFile(handle, "w") as output:
            for info, content in entries:
                output.writestr(
                    info, b"changed" if info.filename == "custom/opaque.bin" else content
                )
        return result

    with Workbook.open(path) as book:
        proposal, _, approval = prepared(book, tmp_path / "out.xlsx")
        monkeypatch.setattr(transaction, "write_archive", corrupt)
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "verification_failed"
    assert not (tmp_path / "out.xlsx").exists()


def test_shared_range_membership_work_is_charged_per_root(tmp_path: Path) -> None:
    from xlayer._batch_impact import analyse_batch

    source = workbook(
        tmp_path, '<c r="A1"><v>1</v></c><c r="B1"><v>1</v></c><c r="C1"><f>SUM(A1:B1)</f></c>'
    )
    with Workbook.open(source) as book:
        result = analyse_batch(
            book.registry,
            book.read_sheet,
            (CellRef("Inputs", "A1"), CellRef("Inputs", "B1")),
            book.source_fingerprint,
            ImpactLimits(max_membership_checks=3),
        )
        assert result.work["membership_checks"] == 3
        assert (
            result.roots[0].impact is not None
            and result.roots[0].impact.analysis_status == "complete"
        )
        assert (
            result.roots[1].impact is not None
            and result.roots[1].impact.analysis_status == "partial"
        )
        assert result.known_union == (CellRef("Inputs", "C1"),)


def test_fifo_output_is_refused_without_reading_or_blocking(tmp_path: Path) -> None:
    if not hasattr(os, "mkfifo"):
        pytest.skip("FIFO unavailable")
    output = tmp_path / "out.xlsx"
    os.mkfifo(output, 0o600)
    with Workbook.open(workbook(tmp_path)) as book:
        result = book.propose([SetValue("Inputs", "A1", 2)], output_path=output, overwrite=True)
        assert isinstance(result, Refusal) and result.code == "unsafe_output_path"


def test_no_dependency_scope_claim_of_global_recalculation_success(tmp_path: Path) -> None:
    with Workbook.open(workbook(tmp_path, '<c r="A1"><v>100</v></c>')) as book:
        proposal, preview, approval = prepared(book, tmp_path / "out.xlsx")
        assert preview.evidence["dependency_trigger"] == "no_dependency_trigger_found"
        assert preview.evidence["recalculation_status"] == "unknown"
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Receipt) and result.evidence["recalculation_status"] == "unknown"
        assert hashlib.sha256((tmp_path / "out.xlsx").read_bytes()).hexdigest() in str(
            result.evidence["output_fingerprint"]
        )
