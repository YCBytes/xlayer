"""Managed snapshot ownership and wiring over the existing real readers."""

from __future__ import annotations

import hashlib
import json
import sys
import zipfile
from collections.abc import Callable, Iterator, Mapping, Sequence
from pathlib import Path
from types import FrameType, MethodType
from typing import TypeAlias, cast
from xml.etree import ElementTree as ET

import pytest

from xlayer import _workbook
from xlayer._errors import ClosedWorkbookError, Refusal, WorkbookOpenError
from xlayer._ooxml.archive import ArchiveLimits, WorkbookArchive
from xlayer._ooxml.sheet import Cell, Worksheet, parse_worksheet
from xlayer._ooxml.styles import StyleInfo
from xlayer._ooxml.workbook import SheetEntry, WorkbookRegistry
from xlayer._workbook import Workbook

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"
EDGES = FIXTURES / "test_workbook_7_write_v11_edges.xlsx"
REGISTRY = FIXTURES / "test_workbook_2_registry.xlsx"
TraceHook: TypeAlias = Callable[[FrameType, str, object], "TraceHook | None"]


def retain_archives(monkeypatch: pytest.MonkeyPatch) -> list[WorkbookArchive]:
    """Keep real instances alive so destructor cleanup cannot satisfy a test."""
    retained: list[WorkbookArchive] = []
    real_init = WorkbookArchive.__init__

    def track(self: WorkbookArchive, *args: object, **kwargs: object) -> None:
        real_init(self, *args, **kwargs)  # type: ignore[arg-type]
        retained.append(self)

    monkeypatch.setattr(WorkbookArchive, "__init__", track)
    return retained


def rewrite_fixture(path: Path, changes: Mapping[str, bytes]) -> Path:
    """Only refusal probes rewrite a sanitized fixture into a temporary ZIP."""
    with zipfile.ZipFile(REGISTRY) as original, zipfile.ZipFile(path, "w") as output:
        for name in original.namelist():
            output.writestr(name, changes.get(name, original.read(name)))
    return path


def assert_archives_closed(archives: list[WorkbookArchive]) -> None:
    assert archives, "the failure did not reach an owned archive"
    assert all(archive.closed and archive._zip.fp is None for archive in archives)


def assert_cell_matches(cell: Cell, expected: Mapping[str, object]) -> None:
    """Compare session wiring with written XML facts, not parser-derived output."""
    for key, value in expected.items():
        if key == "temporal_kind":
            assert cell.style.temporal_kind == value
        elif key == "formula" and value is not None:
            assert cell.formula is not None
            assert isinstance(value, Mapping)
            for attribute, fact in value.items():
                assert getattr(cell.formula, attribute) == fact
        else:
            assert getattr(cell, key) == value


@pytest.fixture
def session() -> Iterator[Workbook]:
    book = Workbook.open(REGISTRY)
    try:
        yield book
    finally:
        book.close()


def spy_sheet_parser(monkeypatch: pytest.MonkeyPatch) -> list[str]:
    calls: list[str] = []

    def track(
        archive: WorkbookArchive,
        sheet: SheetEntry,
        strings: Sequence[str],
        styles: Sequence[StyleInfo],
    ) -> Worksheet | Refusal:
        calls.append(sheet.name)
        return parse_worksheet(archive, sheet, strings, styles)

    monkeypatch.setattr(_workbook, "parse_worksheet", track, raising=False)
    return calls


@pytest.mark.parametrize("as_string", [False, True])
def test_open_loads_registry_and_identity_from_exact_snapshot(as_string: bool) -> None:
    supplied = str(EDGES) if as_string else EDGES
    book = Workbook.open(supplied)
    try:
        assert not book.closed
        assert book.source_path == EDGES
        assert book.source_fingerprint == f"sha256:{hashlib.sha256(EDGES.read_bytes()).hexdigest()}"
        assert [sheet.name for sheet in book.registry.sheets] == ["edges"]
        assert book.registry.date1904 is False
    finally:
        book.close()
    assert book.closed


def test_relative_source_path_is_not_resolved(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.chdir(FIXTURES)
    supplied = Path(EDGES.name)
    book = Workbook.open(supplied)
    try:
        assert book.source_path == supplied
        assert not book.source_path.is_absolute()
    finally:
        book.close()


def test_real_registry_keeps_tab_order_and_scoped_names() -> None:
    book = Workbook.open(REGISTRY)
    try:
        assert [entry.sheet_id for entry in book.registry.sheets] == [1, 5, 3, 4, 2]
        hedge = next(name for name in book.registry.defined_names if name.name == "HedgeRates")
        assert hedge.scope_sheet == "USC - Hedging"
    finally:
        book.close()


def test_open_forwards_limits_to_archive() -> None:
    with pytest.raises(WorkbookOpenError) as caught:
        Workbook.open(EDGES, limits=ArchiveLimits(max_file_bytes=1))
    assert caught.value.refusal.code == "archive_too_large"
    assert caught.value.refusal.details["limit_bytes"] == 1


def test_missing_file_is_operational_open_error(tmp_path: Path) -> None:
    with pytest.raises(WorkbookOpenError) as caught:
        Workbook.open(tmp_path / "missing.xlsx")
    assert caught.value.refusal.code == "invalid_path"
    assert caught.value.refusal.recovery_options


@pytest.mark.parametrize("path", [None, 1, b"book.xlsx"])
def test_bad_path_type_is_programming_error(path: object) -> None:
    with pytest.raises(TypeError):
        Workbook.open(path)  # type: ignore[arg-type]


def test_bad_limits_type_is_programming_error() -> None:
    with pytest.raises(TypeError):
        Workbook.open(EDGES, limits=None)  # type: ignore[arg-type]


@pytest.mark.parametrize(
    "stage", ["parse_package", "parse_workbook_registry", "parse_shared_strings", "parse_styles"]
)
def test_each_open_stage_preserves_full_refusal_and_closes(
    stage: str, monkeypatch: pytest.MonkeyPatch
) -> None:
    archives = retain_archives(monkeypatch)
    refusal = Refusal(
        code="invalid_part_content",
        message="specific parser evidence",
        details={"part": "some.xml", "nested": {"reason": "broken"}},
        recovery_options=({"action": "regenerate_workbook_from_trusted_source"},),
    )

    def refuse(*_args: object) -> Refusal:
        return refusal

    monkeypatch.setattr(_workbook, stage, refuse)
    with pytest.raises(WorkbookOpenError) as caught:
        Workbook.open(REGISTRY)
    assert caught.value.refusal is refusal
    assert caught.value.refusal.to_dict() == refusal.to_dict()
    assert_archives_closed(archives)


@pytest.mark.parametrize(
    "stage", ["parse_package", "parse_workbook_registry", "parse_shared_strings", "parse_styles"]
)
@pytest.mark.parametrize(
    "failure", [RuntimeError("bug"), KeyboardInterrupt("cancelled"), SystemExit("stopped")]
)
def test_each_open_stage_unexpected_failure_closes_and_propagates(
    stage: str, failure: BaseException, monkeypatch: pytest.MonkeyPatch
) -> None:
    archives = retain_archives(monkeypatch)

    def fail(*_args: object) -> object:
        raise failure

    monkeypatch.setattr(_workbook, stage, fail)
    with pytest.raises(type(failure)) as caught:
        Workbook.open(REGISTRY)
    assert caught.value is failure
    assert_archives_closed(archives)


def test_session_construction_failure_does_not_lose_archive(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    archives = retain_archives(monkeypatch)
    failure = RuntimeError("session construction failed")

    def fail(_self: Workbook, *_args: object, **_kwargs: object) -> None:
        raise failure

    monkeypatch.setattr(Workbook, "__init__", fail)
    with pytest.raises(RuntimeError) as caught:
        Workbook.open(EDGES)
    assert caught.value is failure
    assert_archives_closed(archives)


@pytest.mark.parametrize("operation", ["open", "read"])
def test_interruption_at_first_owned_line_is_inside_cleanup_guard(
    operation: str, monkeypatch: pytest.MonkeyPatch
) -> None:
    archives = retain_archives(monkeypatch)
    book = Workbook.open(EDGES) if operation == "read" else None
    target = (
        cast("MethodType", Workbook.open).__func__.__code__
        if operation == "open"
        else Workbook.read_sheet.__code__
    )
    failure = KeyboardInterrupt("interrupted after archive acquisition")
    fired = False

    def interrupt(frame: FrameType, event: str, _arg: object) -> TraceHook | None:
        nonlocal fired
        if (
            event == "line"
            and frame.f_code is target
            and isinstance(frame.f_locals.get("archive"), WorkbookArchive)
        ):
            fired = True
            sys.settrace(None)
            raise failure
        return interrupt

    previous = sys.gettrace()
    try:
        sys.settrace(interrupt)
        with pytest.raises(KeyboardInterrupt) as caught:
            if book is None:
                Workbook.open(EDGES)
            else:
                book.read_sheet("edges")
        assert fired, "probe never reached successful archive acquisition"
        assert caught.value is failure
        assert_archives_closed(archives)
        if book is not None:
            assert book.closed
            assert book._sheet_cache == {}
    finally:
        sys.settrace(previous)
        if book is not None:
            book.close()
        for archive in archives:
            archive.close()


def test_archive_stage_refusal_identity_is_preserved(monkeypatch: pytest.MonkeyPatch) -> None:
    refusal = Refusal(
        code="invalid_path",
        message="specific archive evidence",
        details={"path": "missing.xlsx", "nested": {"reason": "absent"}},
        recovery_options=({"action": "verify_file_path"},),
    )

    def refuse(_cls: type[WorkbookArchive], _path: Path, _limits: ArchiveLimits) -> Refusal:
        return refusal

    archives = retain_archives(monkeypatch)
    monkeypatch.setattr(WorkbookArchive, "load", classmethod(refuse))
    with pytest.raises(WorkbookOpenError) as caught:
        Workbook.open(EDGES)
    assert caught.value.refusal is refusal
    assert caught.value.refusal.to_dict() == refusal.to_dict()
    assert archives == []


def test_open_stages_run_in_locked_order(monkeypatch: pytest.MonkeyPatch) -> None:
    stages: list[str] = []
    real_load = WorkbookArchive.load

    def load(
        _cls: type[WorkbookArchive], path: Path, limits: ArchiveLimits
    ) -> WorkbookArchive | Refusal:
        stages.append("archive")
        return real_load(path, limits)

    def wrap(parser: Callable[..., object], stage: str) -> Callable[..., object]:
        def track(*args: object) -> object:
            stages.append(stage)
            return parser(*args)

        return track

    monkeypatch.setattr(WorkbookArchive, "load", classmethod(load))
    expected = ["parse_package", "parse_workbook_registry", "parse_shared_strings", "parse_styles"]
    for name in expected:
        monkeypatch.setattr(_workbook, name, wrap(getattr(_workbook, name), name))
    with Workbook.open(REGISTRY):
        assert stages == ["archive", *expected]


@pytest.mark.parametrize(
    "part", ["[Content_Types].xml", "xl/workbook.xml", "xl/sharedStrings.xml", "xl/styles.xml"]
)
def test_actual_malformed_shared_infrastructure_refuses_at_open(
    part: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = rewrite_fixture(tmp_path / "bad.xlsx", {part: b"<wrong/>"})
    archives = retain_archives(monkeypatch)
    with pytest.raises(WorkbookOpenError) as caught:
        Workbook.open(path)
    assert caught.value.refusal.code == "invalid_part_content"
    assert caught.value.refusal.details["part"] == part
    assert caught.value.refusal.details["reason"] == "unexpected root element"
    assert_archives_closed(archives)


def test_open_does_not_parse_any_worksheet(monkeypatch: pytest.MonkeyPatch) -> None:
    parts: list[str] = []
    parse_xml = WorkbookArchive.parse_xml_part

    def track(self: WorkbookArchive, name: str) -> ET.Element | Refusal:
        parts.append(name)
        return parse_xml(self, name)

    monkeypatch.setattr(WorkbookArchive, "parse_xml_part", track)
    book = Workbook.open(REGISTRY)
    book.close()
    assert "xl/workbook.xml" in parts
    assert "xl/sharedStrings.xml" in parts
    assert "xl/styles.xml" in parts
    assert not any("worksheets/" in part for part in parts)


def test_read_sheet_uses_exact_registry_lookup_and_saved_values(session: Workbook) -> None:
    sheet = session.read_sheet("USC - Hedging")
    assert isinstance(sheet, Worksheet)
    assert sheet.part_path == "xl/worksheets/sheet1.xml"
    assert sheet.cells["B1"].raw == "999"
    assert sheet.cells["B1"].stored_value == 999.0
    assert sheet.cells["A1"].stored_value == "Sheet 2 - Visible, name has spaces and hyphen"


@pytest.mark.parametrize("name", ["not a sheet", "usc - hedging", "USC - Hedging ", ""])
def test_unknown_sheet_has_ordered_names_and_structured_recovery(
    session: Workbook, name: str
) -> None:
    refusal = session.read_sheet(name)
    assert isinstance(refusal, Refusal)
    assert refusal.code == "sheet_not_found"
    assert refusal.operation == "read_sheet"
    assert refusal.target == name
    assert refusal.to_dict()["details"] == {
        "requested_name": name,
        "available_names": [
            "USC - Hedging",
            "Visible Two",
            "Hidden Sheet",
            "Very Hidden",
            "Visible One",
        ],
    }
    assert refusal.recovery_options[0]["action"] == "choose_existing_sheet"
    assert refusal.recovery_options[0]["available_names"] == refusal.details["available_names"]
    assert not session.closed


@pytest.mark.parametrize("name", [None, 1, b"edges", ["edges"]])
def test_bad_sheet_name_type_is_programming_error(session: Workbook, name: object) -> None:
    with pytest.raises(TypeError):
        session.read_sheet(name)  # type: ignore[arg-type]
    assert not session.closed


def test_repeated_good_sheet_parses_once_and_cache_is_session_local(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls = spy_sheet_parser(monkeypatch)
    first = Workbook.open(EDGES)
    second = Workbook.open(EDGES)
    try:
        assert calls == []
        read = first.read_sheet("edges")
        assert isinstance(read, Worksheet)
        assert first.read_sheet("edges") is read
        assert calls == ["edges"]
        other = second.read_sheet("edges")
        assert other == read
        assert other is not read
        assert calls == ["edges", "edges"]
    finally:
        first.close()
        second.close()


def test_bad_sheet_is_lazy_cached_and_does_not_block_a_good_sheet(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = rewrite_fixture(
        tmp_path / "one-bad-sheet.xlsx", {"xl/worksheets/sheet1.xml": b"<malformed"}
    )
    calls = spy_sheet_parser(monkeypatch)
    book = Workbook.open(path)
    try:
        assert calls == []
        refused = book.read_sheet("USC - Hedging")
        assert isinstance(refused, Refusal)
        assert refused.code == "malformed_xml"
        assert refused.details["part"] == "xl/worksheets/sheet1.xml"
        assert book.read_sheet("USC - Hedging") is refused
        assert calls == ["USC - Hedging"]
        good = book.read_sheet("Visible Two")
        assert isinstance(good, Worksheet)
        assert good.cells["B1"].raw == "42"  # independently checked in the fixture XML
        assert good.cells["B1"].stored_value == 42.0
        assert calls == ["USC - Hedging", "Visible Two"]
        assert not book.closed
    finally:
        book.close()


@pytest.mark.parametrize("kind", ["chartsheet", "dialogsheet"])
def test_unsupported_registered_sheet_keeps_existing_refusal(
    kind: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    with zipfile.ZipFile(REGISTRY) as original:
        rels = original.read("xl/_rels/workbook.xml.rels")
    root = ET.fromstring(rels)  # noqa: S314 - trusted, sanitized fixture
    rel = root.find(
        '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship[@Id="rId1"]'
    )
    assert rel is not None
    rel.set("Type", f"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}")
    changed = ET.tostring(root)
    path = rewrite_fixture(tmp_path / "unsupported.xlsx", {"xl/_rels/workbook.xml.rels": changed})
    calls = spy_sheet_parser(monkeypatch)
    book = Workbook.open(path)
    try:
        refusal = book.read_sheet("USC - Hedging")
        assert isinstance(refusal, Refusal)
        assert refusal.code == "unsupported_sheet_kind"
        assert book.read_sheet("USC - Hedging") is refusal
        assert calls == ["USC - Hedging"]
        assert isinstance(book.read_sheet("Visible Two"), Worksheet)
        assert not book.closed
    finally:
        book.close()


@pytest.mark.parametrize(
    "failure", [RuntimeError("parser bug"), KeyboardInterrupt("cancelled"), SystemExit("stopped")]
)
def test_unexpected_sheet_failure_closes_session_and_preserves_exception(
    failure: BaseException, monkeypatch: pytest.MonkeyPatch
) -> None:
    archives = retain_archives(monkeypatch)
    book = Workbook.open(EDGES)

    def fail(*_args: object) -> object:
        raise failure

    monkeypatch.setattr(_workbook, "parse_worksheet", fail, raising=False)
    try:
        with pytest.raises(type(failure)) as caught:
            book.read_sheet("edges")
        assert caught.value is failure
        assert book.closed
        assert_archives_closed(archives)
    finally:
        book.close()


@pytest.mark.parametrize("failure", [RuntimeError("lookup bug"), KeyboardInterrupt("cancelled")])
def test_unexpected_lookup_failure_also_closes_session(
    failure: BaseException, monkeypatch: pytest.MonkeyPatch
) -> None:
    archives = retain_archives(monkeypatch)
    book = Workbook.open(EDGES)

    def fail(_self: WorkbookRegistry, _name: str) -> SheetEntry | None:
        raise failure

    monkeypatch.setattr(WorkbookRegistry, "sheet_by_name", fail)
    try:
        with pytest.raises(type(failure)) as caught:
            book.read_sheet("edges")
        assert caught.value is failure
        assert book.closed
        assert_archives_closed(archives)
    finally:
        book.close()


def test_unknown_names_never_grow_sheet_cache(session: Workbook) -> None:
    for index in range(100):
        assert isinstance(session.read_sheet(f"unknown {index}"), Refusal)
    assert session._sheet_cache == {}


def test_close_blocks_all_data_access_but_retains_identity_and_returned_records(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    archives = retain_archives(monkeypatch)
    book = Workbook.open(REGISTRY)
    registry = book.registry
    sheet = book.read_sheet("USC - Hedging")
    assert isinstance(sheet, Worksheet)
    path, fingerprint = book.source_path, book.source_fingerprint
    book.close()
    book.close()
    assert book.closed
    assert book.source_path == path
    assert book.source_fingerprint == fingerprint
    assert_archives_closed(archives)
    for name in ("USC - Hedging", "Visible Two", "missing", None):
        with pytest.raises(ClosedWorkbookError):
            book.read_sheet(name)  # type: ignore[arg-type]
    with pytest.raises(ClosedWorkbookError):
        _ = book.registry
    assert registry.sheets[0].name == "USC - Hedging"
    assert sheet.cells["B1"].raw == "999"
    with pytest.raises(TypeError):
        sheet.cells["C999"] = sheet.cells["B1"]  # type: ignore[index]
    assert book._archive is None
    assert book._registry is None
    assert book._shared_strings == ()
    assert book._styles == ()
    assert book._sheet_cache == {}


def test_close_calls_archive_once(monkeypatch: pytest.MonkeyPatch) -> None:
    book = Workbook.open(EDGES)
    calls: list[WorkbookArchive] = []
    real_close = WorkbookArchive.close

    def track(archive: WorkbookArchive) -> None:
        calls.append(archive)
        real_close(archive)

    monkeypatch.setattr(WorkbookArchive, "close", track)
    book.close()
    book.close()
    assert len(calls) == 1


def test_context_returns_same_session_and_closes_normally() -> None:
    book = Workbook.open(EDGES)
    try:
        with book as entered:
            assert entered is book
            assert isinstance(entered.read_sheet("edges"), Worksheet)
        assert book.closed
        with pytest.raises(ClosedWorkbookError), book:
            pytest.fail("closed context must not enter")
    finally:
        book.close()


@pytest.mark.parametrize(
    "failure", [RuntimeError("body bug"), KeyboardInterrupt("cancelled"), SystemExit("stopped")]
)
def test_context_body_failure_is_not_suppressed_and_archive_closes(
    failure: BaseException, monkeypatch: pytest.MonkeyPatch
) -> None:
    archives = retain_archives(monkeypatch)
    book = Workbook.open(EDGES)
    try:
        with pytest.raises(type(failure)) as caught, book:
            raise failure
        assert caught.value is failure
        assert book.closed
        assert_archives_closed(archives)
    finally:
        book.close()


@pytest.mark.parametrize("change", ["replace", "delete"])
def test_snapshot_reads_and_fingerprint_survive_source_change(change: str, tmp_path: Path) -> None:
    path = tmp_path / "snapshot.xlsx"
    original = EDGES.read_bytes()
    path.write_bytes(original)
    book = Workbook.open(path)
    try:
        # No sheet read before changing the file: this proves the archive is
        # a snapshot rather than a cache of previously requested worksheets.
        if change == "replace":
            path.write_bytes(REGISTRY.read_bytes())
            fresh = Workbook.open(path)
            try:
                assert fresh.registry.sheets[0].name == "USC - Hedging"
                assert fresh.source_fingerprint != book.source_fingerprint
            finally:
                fresh.close()
        else:
            path.unlink()
            with pytest.raises(WorkbookOpenError) as caught:
                Workbook.open(path)
            assert caught.value.refusal.code == "invalid_path"
        sheet = book.read_sheet("edges")
        assert isinstance(sheet, Worksheet)
        assert sheet.cells["A1"].raw == "100"
        assert book.source_fingerprint == f"sha256:{hashlib.sha256(original).hexdigest()}"
        assert book.source_path == path
    finally:
        book.close()


@pytest.mark.parametrize(
    "filename, sheet_count",
    [
        ("test_workbook_2_registry.xlsx", 5),
        ("test_workbook_6b_1904.xlsx", 1),
        ("test_workbook_3_formats.xlsx", 2),
        ("test_workbook_4_worksheet_final.xlsx", 13),
        ("test_workbook_7_write_v11_edges.xlsx", 1),
    ],
)
def test_all_fixture_sheets_read_equal_in_separate_sessions_without_source_mutation(
    filename: str, sheet_count: int
) -> None:
    path = FIXTURES / filename
    before = path.read_bytes()
    with Workbook.open(path) as first, Workbook.open(path) as second:
        assert len(first.registry.sheets) == sheet_count
        for entry in first.registry.sheets:
            left = first.read_sheet(entry.name)
            right = second.read_sheet(entry.name)
            assert isinstance(left, Worksheet)
            assert isinstance(right, Worksheet)
            assert left == right
    assert path.read_bytes() == before


def test_edges_session_matches_existing_handwritten_expectations() -> None:
    expected = json.loads(
        (FIXTURES / "test_workbook_7_write_v11_edges.sheet.json").read_text(encoding="utf-8")
    )
    with Workbook.open(EDGES) as book:
        sheet = book.read_sheet(expected["sheet"])
        assert isinstance(sheet, Worksheet)
        assert sheet.part_path == expected["part_path"]
        assert sheet.declared_dimension == expected["declared_dimension"]
        assert sheet.used_range == expected["used_range"]
        assert len(sheet.cells) == expected["cell_count"]
        assert list(sheet.other_elements) == expected["other_elements"]
        assert [merge.ref for merge in sheet.merges] == [item["ref"] for item in expected["merges"]]
        for address, spec in expected["cells"].items():
            assert_cell_matches(sheet.cells[address], spec)
        for address in expected["absent"]:
            assert address not in sheet.cells
        assert sheet.merge_at("C10") is not None


def test_formula_fixture_matches_existing_handwritten_expectations() -> None:
    expected = json.loads(
        (FIXTURES / "test_workbook_4_worksheet_final.sheet.json").read_text(encoding="utf-8")
    )
    with Workbook.open(FIXTURES / "test_workbook_4_worksheet_final.xlsx") as book:
        for name, spec in expected.items():
            sheet = book.read_sheet(name)
            assert isinstance(sheet, Worksheet)
            for key in ("part_path", "declared_dimension", "used_range"):
                if key in spec:
                    assert getattr(sheet, key) == spec[key]
            for address, cell_spec in spec.get("cells", {}).items():
                assert_cell_matches(sheet.cells[address], cell_spec)
            if "merges" in spec:
                assert [merge.ref for merge in sheet.merges] == spec["merges"]
            for index, shared in spec.get("shared_formulas", {}).items():
                master = sheet.shared_formulas[int(index)]
                assert master.master == shared["master"]
                assert master.ref == shared["ref"]
                assert master.text == shared["text"]


def test_1904_and_text_styles_pass_through_without_datetime_conversion() -> None:
    with Workbook.open(FIXTURES / "test_workbook_6b_1904.xlsx") as book:
        assert book.registry.date1904 is True
        sheet = book.read_sheet("Date 1904 Test")
        assert isinstance(sheet, Worksheet)
        assert sheet.cells["B2"].raw == "44196"
        assert sheet.cells["B2"].stored_value == 44196.0
        assert sheet.cells["B2"].style.temporal_kind == "date"
    with Workbook.open(FIXTURES / "test_workbook_3_formats.xlsx") as book:
        text = book.read_sheet("Formats and Strings")
        assert isinstance(text, Worksheet)
        assert text.cells["B10"].stored_value == "  leading spaces preserved  "
        blanks = book.read_sheet("Styles Inflation")
        assert isinstance(blanks, Worksheet)
        assert blanks.cells["C2"].value_kind == "blank"
        assert blanks.cells["C2"].stored_value is None
