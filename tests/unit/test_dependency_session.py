"""Private snapshot/ownership integration and independent raw-XML fixture facts."""

from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
import zipfile
from dataclasses import replace
from pathlib import Path
from typing import cast

import pytest

from tests.unit._dependency_cases import write_workbook_case
from xlayer import _workbook
from xlayer._dependencies import CellRef, DependencyImpact, ImpactLimits
from xlayer._errors import ClosedWorkbookError, Refusal
from xlayer._ooxml.workbook import WorkbookRegistry
from xlayer._workbook import Workbook

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"
WB4 = FIXTURES / "test_workbook_4_worksheet_final.xlsx"
CHAIN = '<sheetData><row r="1"><c r="B1"><f>A1</f></c></row></sheetData>'


def result(book: Workbook, sheet: str = "S", address: str = "A1") -> DependencyImpact:
    impact = book.dependency_impact(sheet, address)
    assert isinstance(impact, DependencyImpact)
    return impact


def test_dependency_closed_check_precedes_arguments(tmp_path: Path) -> None:
    book = Workbook.open(write_workbook_case(tmp_path / "case.xlsx", {"S": CHAIN}))
    book.close()
    with pytest.raises(ClosedWorkbookError):
        book.dependency_impact(cast(str, None), cast(str, 1), limits=cast(ImpactLimits, None))


@pytest.mark.parametrize(
    ("name", "address", "limits"),
    [(None, "A1", ImpactLimits()), ("S", 1, ImpactLimits()), ("S", "A1", None)],
)
def test_dependency_argument_types_leave_session_usable(
    tmp_path: Path,
    name: object,
    address: object,
    limits: object,
) -> None:
    with Workbook.open(write_workbook_case(tmp_path / "case.xlsx", {"S": CHAIN})) as book:
        with pytest.raises(TypeError):
            book.dependency_impact(
                cast(str, name), cast(str, address), limits=cast(ImpactLimits, limits)
            )
        assert not book.closed
        assert result(book).exact_total_count == 1


@pytest.mark.parametrize(
    "address", ["", "$A$1", " A1", "A1 ", "S!A1", "A0", "XFE1", "A1048577", "A1٢"]
)
def test_dependency_argument_coordinates_leave_session_usable(tmp_path: Path, address: str) -> None:
    with Workbook.open(write_workbook_case(tmp_path / "case.xlsx", {"S": CHAIN})) as book:
        with pytest.raises(ValueError):
            book.dependency_impact("S", address)
        assert not book.closed and result(book).exact_total_count == 1
        assert result(book, address="a1").root == CellRef("S", "A1")


def test_dependency_unknown_root_uses_exact_lookup(tmp_path: Path) -> None:
    path = write_workbook_case(tmp_path / "case.xlsx", {"S": CHAIN, "T": "<sheetData/>"})
    with Workbook.open(path) as book:
        refused = book.dependency_impact("s", "A1")
        assert isinstance(refused, Refusal)
        assert (refused.code, refused.operation) == ("sheet_not_found", "dependency_impact")
        assert refused.details["available_names"] == ("S", "T")
        assert refused.recovery_options and not book.closed


@pytest.mark.parametrize("kind", ["chartsheet", "dialogsheet"])
def test_dependency_unsupported_root_and_excluded_scope(tmp_path: Path, kind: str) -> None:
    path = write_workbook_case(tmp_path / "case.xlsx", {"S": CHAIN, "Visual": "<sheetData/>"})
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    rels = "xl/_rels/workbook.xml.rels"
    parts[rels] = parts[rels].replace(
        b'relationships/worksheet" Target="worksheets/2.xml"',
        f'relationships/{kind}" Target="worksheets/2.xml"'.encode(),
    )
    with zipfile.ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    with Workbook.open(path) as book:
        refused = book.dependency_impact("Visual", "A1")
        assert isinstance(refused, Refusal)
        assert (refused.code, refused.operation) == ("unsupported_sheet_kind", "dependency_impact")
        impact = result(book)
        assert impact.analysis_status == "complete" and impact.work.sheets_attempted == 1
        assert [(entry.name, entry.kind) for entry in impact.excluded_tabs] == [("Visual", kind)]


@pytest.mark.parametrize("stage", ["lookup", "acquisition", "validation", "analyzer"])
@pytest.mark.parametrize(
    "failure",
    [RuntimeError("bug"), TypeError("bug"), ValueError("bug"), KeyboardInterrupt(), SystemExit(2)],
)
def test_dependency_unexpected_failure_closes_and_reraises(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    stage: str,
    failure: BaseException,
) -> None:
    book = Workbook.open(write_workbook_case(tmp_path / "case.xlsx", {"S": CHAIN}))
    archive = book._archive  # Retain ownership object: GC cannot fake close.
    assert archive is not None

    def fail(*_args: object, **_kwargs: object) -> object:
        raise failure

    if stage == "lookup":
        monkeypatch.setattr(WorkbookRegistry, "sheet_by_name", fail)
    elif stage == "acquisition":
        monkeypatch.setattr(Workbook, "registry", property(fail))
    elif stage == "validation":
        monkeypatch.setattr(_workbook, "_canonical_address", fail, raising=False)
    else:
        monkeypatch.setattr(_workbook, "analyse_dependencies", fail, raising=False)
    with pytest.raises(type(failure)) as caught:
        book.dependency_impact("S", "A1")
    assert caught.value is failure
    assert book.closed and archive.closed and archive._zip.fp is None


def test_dependency_result_survives_close(tmp_path: Path) -> None:
    book = Workbook.open(write_workbook_case(tmp_path / "case.xlsx", {"S": CHAIN}))
    impact = result(book)
    before = impact.to_dict()
    book.close()
    assert impact.to_dict() == before
    assert len(impact.path_to(CellRef("S", "B1"))) == 1
    with pytest.raises(ClosedWorkbookError):
        book.dependency_impact("S", "A1")


def test_dependency_snapshot_survives_source_replace_or_delete(tmp_path: Path) -> None:
    path = write_workbook_case(tmp_path / "case.xlsx", {"S": CHAIN})
    with Workbook.open(path) as book:
        before = path.read_bytes()
        original = result(book)
        assert path.read_bytes() == before
        write_workbook_case(path, {"S": "<sheetData/>"})
        assert result(book) == original
        with Workbook.open(path) as fresh:
            assert result(fresh).exact_total_count == 0
            assert fresh.source_fingerprint != book.source_fingerprint
        path.unlink()
        assert result(book) == original


def test_cold_warm_repeated_queries_have_identical_work(tmp_path: Path) -> None:
    path = write_workbook_case(tmp_path / "case.xlsx", {"S": CHAIN})
    with Workbook.open(path) as cold, Workbook.open(path) as warm:
        warm.read_sheet("S")
        keys = set(cold.__dict__)
        first = result(cold)
        assert result(cold) == first == result(warm)
        assert set(cold.__dict__) == keys  # no session-lifetime dependency index
        assert first.work.scanned_cells == first.work.formula_cells_attempted == 1
        assert first.work.sheets_attempted == 1


def test_public_fixture_dependencies_match_raw_xml() -> None:
    answers = json.loads(
        (FIXTURES / "test_workbook_4_worksheet_final.dependencies.json").read_text()
    )
    with Workbook.open(WB4) as book:
        for query in answers["queries"]:
            impact = result(book, query["sheet"], query["address"])
            assert impact.analysis_status == answers["analysis_status"]
            assert [(c.sheet, c.address) for c in impact.direct_dependents] == [
                tuple(c) for c in query["direct"]
            ]
            assert [(c.sheet, c.address) for c in impact.transitive_dependents] == [
                tuple(c) for c in query["transitive"]
            ]
            assert impact.exact_total_count is None
            if "master" in query:
                evidence = impact.evidence_edges[0].evidence
                assert evidence.shared_master == CellRef(*query["master"])
                assert evidence.formula_text == query["formula_text"]
                assert evidence.source_span == tuple(query["span"])
                assert evidence.shared_offset == (0, 0)
            if "cycles" in query:
                assert impact.known_cycles == tuple(
                    tuple(CellRef(*c) for c in group) for group in query["cycles"]
                )
            sites = {
                (i.formula_cell.sheet, i.formula_cell.address)
                for i in impact.issues
                if i.formula_cell
            }
            assert {tuple(c) for c in answers["required_issue_sites"]} <= sites
        name_impact = result(book, "ref_data", "A2")
        names = {
            edge.evidence.name_definition.name
            for edge in name_impact.evidence_edges
            if edge.evidence.name_definition
        }
        assert names == {"PriceData"}
        total = result(book, "ref_data", "B12")
        assert total.evidence_edges[0].evidence.name_definition is not None
        assert total.evidence_edges[0].evidence.name_definition.name == "VolumeTotal"


@pytest.mark.parametrize("column", ["B", "C", "D", "E", "F", "G", "H"])
def test_horizontal_shared_follower_matches_its_own_column(column: str) -> None:
    answers = json.loads(
        (FIXTURES / "test_workbook_4_worksheet_final.dependencies.json").read_text()
    )["shared_horizontal"]
    with Workbook.open(WB4) as book:
        impact = result(book, answers["sheet"], f"{column}{answers['root_row']}")
        assert impact.direct_dependents == (CellRef("merged_cells", f"{column}11"),)
        assert impact.transitive_dependents == ()
        evidence = impact.evidence_edges[0].evidence
        assert evidence.shared_master == CellRef(*answers["master"])
        assert evidence.shared_offset == (0, ord(column) - ord("B"))
        assert evidence.source_span == tuple(answers["span"])
        assert evidence.formula_text == answers["formula_text"]
        assert evidence.normalized_reference.min_column == ord(column) - ord("A") + 1


@pytest.mark.parametrize("fixture", sorted(FIXTURES.glob("*.xlsx")), ids=lambda p: p.name)
def test_all_public_fixtures_have_a_trusted_dependency_inventory(fixture: Path) -> None:
    with Workbook.open(fixture) as book:
        first = next(entry for entry in book.registry.sheets if entry.kind == "worksheet")
        impact = result(book, first.name, "A1")
        assert impact.analysis_status in {"complete", "partial"}
        assert impact.work.trusted_sheets == sum(
            entry.kind == "worksheet" for entry in book.registry.sheets
        )
        assert all(issue.code != "worksheet_unreadable" for issue in impact.issues)


def test_formula_free_fixture_has_exact_zero() -> None:
    with Workbook.open(FIXTURES / "test_workbook_7_write_v11_edges.xlsx") as book:
        impact = result(book, "edges", "C10")
        assert impact.analysis_status == "complete" and impact.exact_total_count == 0
        assert impact.work.formula_cells_attempted == 0
        assert impact.evidence_edges == ()
        assert impact.known_cycles == ()


def test_same_bytes_and_changed_zip_order_have_correct_identity(tmp_path: Path) -> None:
    first = write_workbook_case(tmp_path / "first.xlsx", {"S": CHAIN})
    with zipfile.ZipFile(first) as archive:
        order = tuple(reversed(archive.namelist()))
    second = write_workbook_case(tmp_path / "second.xlsx", {"S": CHAIN}, zip_order=order)
    with Workbook.open(first) as left, Workbook.open(second) as right:
        a, b = result(left).to_dict(), result(right).to_dict()
    for path, output in ((first, a), (second, b)):
        assert (
            output["source_fingerprint"]
            == "sha256:" + hashlib.sha256(path.read_bytes()).hexdigest()
        )
    assert a.pop("source_fingerprint") != b.pop("source_fingerprint")
    assert a == b


def test_hash_seed_and_locale_do_not_change_canonical_work(tmp_path: Path) -> None:
    path = write_workbook_case(tmp_path / "case.xlsx", {"S": CHAIN})
    script = """
import json, locale, sys
from xlayer._workbook import Workbook
if sys.argv[2] == "C":
    effective = locale.setlocale(locale.LC_CTYPE, "C")
else:
    for candidate in ("C.UTF-8", "en_US.UTF-8", ".UTF-8", "UTF-8"):
        try:
            effective = locale.setlocale(locale.LC_CTYPE, candidate)
            break
        except locale.Error:
            pass
    else:
        raise RuntimeError("No second effective locale available for determinism proof")
with Workbook.open(sys.argv[1]) as book:
    impact = book.dependency_impact("S", "A1")
    print(json.dumps({"locale": effective, "impact": impact.to_dict()}, sort_keys=True))
"""
    outputs = []
    for seed, requested in (
        (seed, requested) for seed in ("1", "441", "random") for requested in ("C", "UTF-8")
    ):
        env = dict(os.environ, PYTHONHASHSEED=seed, LC_ALL="C", LANG="C")
        outputs.append(
            subprocess.check_output(  # noqa: S603 — fixed script, interpreter and test path
                [sys.executable, "-c", script, str(path), requested],
                env=env,
            )
        )
    reports = [json.loads(output) for output in outputs]
    assert len({report["locale"] for report in reports}) >= 2
    assert all(report["impact"] == reports[0]["impact"] for report in reports)


def test_refusal_remains_operational_not_session_failure(tmp_path: Path) -> None:
    path = write_workbook_case(
        tmp_path / "case.xlsx",
        {"S": CHAIN, "T": "<sheetData><row><c><v>bad</v></c></row></sheetData>"},
    )
    with Workbook.open(path) as book:
        impact = book.dependency_impact("S", "A1", limits=replace(ImpactLimits(), max_references=1))
        assert isinstance(impact, DependencyImpact) and impact.analysis_status == "partial"
        assert impact.direct_dependents == (CellRef("S", "B1"),)
        assert not book.closed
