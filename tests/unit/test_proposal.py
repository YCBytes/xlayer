"""Proposal/preview identity and path-binding tests, without performing a write."""

from __future__ import annotations

import os
from pathlib import Path
from typing import TYPE_CHECKING

import pytest

from tests.unit._dependency_cases import write_workbook_case
from xlayer._dependencies import ImpactLimits
from xlayer._edits import SetValue
from xlayer._errors import ClosedWorkbookError, Refusal
from xlayer._workbook import Workbook

if TYPE_CHECKING:
    from xlayer._proposal import Proposal


def propose(
    book: Workbook, edits: list[SetValue], output: Path, **kwargs: object
) -> Proposal | Refusal:
    assert hasattr(book, "propose"), "proposal entry point missing"
    return book.propose(edits, output_path=output, **kwargs)  # type: ignore[arg-type]


def source(tmp_path: Path) -> Path:
    return write_workbook_case(
        tmp_path / "input.xlsx",
        {
            "Inputs": '<sheetData><row><c r="A1"><v>100</v></c><c r="B1"><v>2</v></c>'
            '<c r="C1"><f>A1*B1</f><v>200</v></c></row></sheetData>'
        },
    )


def test_preview_exact_facts_and_approval_requirements(tmp_path: Path) -> None:
    path = source(tmp_path)
    with Workbook.open(path) as book:
        result = propose(book, [SetValue("Inputs", "A1", 120)], tmp_path / "out.xlsx")
        assert not isinstance(result, Refusal)
        preview = result.preview()
        assert not isinstance(preview, Refusal)
        data = preview.to_dict()
        assert data["recalculation_status"] == "required_not_run"
        assert data["visual_status"] == "not_evaluated"
        assert data["required_capabilities"] == ["apply_set_value"]
        assert data["next_action"] == "obtain_approval"
        assert data["expected_parts"] == ["xl/worksheets/1.xml"]
        assert not (tmp_path / "out.xlsx").exists()


def test_invalid_second_member_blocks_whole_batch_and_retains_findings(tmp_path: Path) -> None:
    path = source(tmp_path)
    before = path.read_bytes()
    with Workbook.open(path) as book:
        result = propose(
            book,
            [SetValue("Inputs", "A1", 120), SetValue("Inputs", "Z1", 3)],
            tmp_path / "out.xlsx",
        )
        assert not isinstance(result, Refusal)
        preview = result.preview()
        assert not isinstance(preview, Refusal)
        assert preview.evidence["blocked"] is True
        assert preview.evidence["next_action"] == "revise_proposal"
    assert path.read_bytes() == before and not (tmp_path / "out.xlsx").exists()


def test_no_changes_and_duplicate_target_refuse(tmp_path: Path) -> None:
    with Workbook.open(source(tmp_path)) as book:
        result = propose(book, [SetValue("Inputs", "A1", 100)], tmp_path / "out.xlsx")
        assert not isinstance(result, Refusal)
        preview = result.preview()
        assert isinstance(preview, Refusal) and preview.code == "no_changes"
        repeated = propose(
            book, [SetValue("Inputs", "a1", 2), SetValue("Inputs", "A1", 3)], tmp_path / "out.xlsx"
        )
        assert isinstance(repeated, Refusal) and repeated.code == "duplicate_target"


def test_closed_proposal_cannot_use_cached_preview(tmp_path: Path) -> None:
    book = Workbook.open(source(tmp_path))
    result = propose(book, [SetValue("Inputs", "A1", 2)], tmp_path / "out.xlsx")
    assert not isinstance(result, Refusal)
    preview = result.preview()
    assert not isinstance(preview, Refusal)
    book.close()
    with pytest.raises(ClosedWorkbookError):
        result.preview()
    assert preview.to_dict()["source_fingerprint"] == book.source_fingerprint


@pytest.mark.parametrize("kind", ["source", "hardlink", "symlink", "dangling", "directory"])
def test_unsafe_output_paths_refuse(tmp_path: Path, kind: str) -> None:
    path = source(tmp_path)
    output = tmp_path / "out.xlsx"
    if kind == "source":
        output = path
    elif kind == "hardlink":
        os.link(path, output)
    elif kind in {"symlink", "dangling"}:
        try:
            output.symlink_to(path if kind == "symlink" else tmp_path / "missing")
        except OSError:
            pytest.skip("symlinks unavailable")
    else:
        output.mkdir()
    with Workbook.open(path) as book:
        result = propose(book, [SetValue("Inputs", "A1", 2)], output, overwrite=True)
        assert isinstance(result, Refusal) and result.code == "unsafe_output_path"


def test_explicit_overwrite_identity_is_previewed(tmp_path: Path) -> None:
    output = tmp_path / "out.xlsx"
    output.write_bytes(b"existing")
    with Workbook.open(source(tmp_path)) as book:
        refused = propose(book, [SetValue("Inputs", "A1", 2)], output)
        assert isinstance(refused, Refusal) and refused.code == "output_exists"
        result = propose(book, [SetValue("Inputs", "A1", 2)], output, overwrite=True)
        assert not isinstance(result, Refusal)
        preview = result.preview()
        assert not isinstance(preview, Refusal)
        assert preview.evidence["output_incumbent"] is not None


def test_cwd_changes_do_not_redirect_source_or_output(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source(tmp_path)
    monkeypatch.chdir(tmp_path)
    with Workbook.open("input.xlsx") as book:
        result = propose(book, [SetValue("Inputs", "A1", 2)], Path("out.xlsx"))
        assert not isinstance(result, Refusal)
        monkeypatch.chdir(tmp_path.parent)
        preview = result.preview()
        assert not isinstance(preview, Refusal)
        assert book.source_path == Path("input.xlsx")
        assert result.output.path == tmp_path / "out.xlsx"


def test_changed_limits_and_order_change_proposal_identity(tmp_path: Path) -> None:
    with Workbook.open(source(tmp_path)) as book:
        edits = [SetValue("Inputs", "A1", 2), SetValue("Inputs", "B1", 3)]
        a = propose(book, edits, tmp_path / "out.xlsx")
        b = propose(book, edits[::-1], tmp_path / "out.xlsx")
        c = propose(book, edits, tmp_path / "out.xlsx", limits=ImpactLimits(max_visited_nodes=1))
        assert (
            not isinstance(a, Refusal) and not isinstance(b, Refusal) and not isinstance(c, Refusal)
        )
        assert len({a.proposal_digest, b.proposal_digest, c.proposal_digest}) == 3
        preview = c.preview()
        assert not isinstance(preview, Refusal) and preview.evidence["blocked"] is True


def test_read_session_does_not_require_path_after_bytes_loaded(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import xlayer._workbook as module

    path = source(tmp_path)
    from xlayer._ooxml.styles import parse_styles as original

    def remove_path(*args: object) -> object:
        result = original(*args)  # type: ignore[arg-type]
        path.unlink()
        return result

    monkeypatch.setattr(module, "parse_styles", remove_path)
    with Workbook.open(path) as book:
        sheet = book.read_sheet("Inputs")
        assert not isinstance(sheet, Refusal) and sheet.cells["A1"].raw == "100"
