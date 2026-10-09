"""Independent observation, bound and lifecycle oracles for the supported facade."""

from __future__ import annotations

import hashlib
import json
import zipfile
from collections.abc import Sequence
from dataclasses import FrozenInstanceError, fields, replace
from pathlib import Path
from typing import cast, overload

import pytest

from tests.unit._dependency_cases import write_workbook_case
from xlayer import (
    ArchiveLimits,
    ClosedWorkbookError,
    Inspection,
    ReadLimits,
    Refusal,
    Workbook,
    WorkbookOpenError,
)

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def source(tmp_path: Path, cells: str, *, extra: str = "", sheet_attrs: str = "") -> Path:
    path = write_workbook_case(
        tmp_path / "source.xlsx", {"Inputs": f"<sheetData><row>{cells}</row></sheetData>{extra}"}
    )
    if sheet_attrs:
        rewrite(
            path,
            {
                "xl/worksheets/1.xml": (
                    f'<worksheet xmlns="{NS}" {sheet_attrs}><sheetData><row>'
                    f"{cells}</row></sheetData>{extra}</worksheet>"
                ).encode()
            },
        )
    return path


def rewrite(path: Path, changes: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path) as zf:
        parts = {name: zf.read(name) for name in zf.namelist()}
    parts.update(changes)
    with zipfile.ZipFile(path, "w") as zf:
        for name, content in parts.items():
            zf.writestr(name, content)


def data(result: Inspection | Refusal) -> dict[str, object]:
    assert isinstance(result, Inspection)
    value = result.to_dict()["data"]
    assert isinstance(value, dict)
    return cast("dict[str, object]", value)


def cells(result: Inspection | Refusal) -> list[dict[str, object]]:
    value = data(result)["cells"]
    assert isinstance(value, list)
    return cast("list[dict[str, object]]", value)


def test_all_presence_and_cache_states_are_distinct(tmp_path: Path) -> None:
    path = source(
        tmp_path,
        '<c r="A1"/><c r="B1" t="inlineStr"><is><t/></is></c>'
        '<c r="C1"><f>1</f></c><c r="D1"><f>1</f><v/></c>'
        '<c r="E1"><f>0</f><v>0</v></c><c r="F1" t="b"><f>TRUE()</f></c>'
        '<c r="G1" t="str"><f>""</f><v/></c><c r="H1" t="e"><f>1/0</f><v>#DIV/0!</v></c>',
    )
    with Workbook.open(path) as book:
        observed = cells(
            book.read_cells("Inputs", ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1"])
        )
    assert [(c["presence"], c["value_kind"], c["raw"], c["stored_value"]) for c in observed] == [
        ("present", "blank", None, None),
        ("present", "string", None, ""),
        ("present", "blank", None, None),
        ("present", "blank", "", None),
        ("present", "number", "0", 0.0),
        ("present", "blank", None, None),
        ("present", "string", "", ""),
        ("present", "error", "#DIV/0!", "#DIV/0!"),
        ("missing", None, None, None),
    ]
    assert [c["formula_result_status"] for c in observed] == [
        None,
        None,
        *(["not_verified"] * 6),
        None,
    ]
    assert all(set(c) == set(observed[0]) for c in observed)
    assert observed[-1] == {
        "address": "I1",
        "presence": "missing",
        "stored_type": None,
        "value_kind": None,
        "raw": None,
        "stored_value": None,
        "value_source": "not_present",
        "formula_result_status": None,
        "cell_format": None,
        "formula": None,
        "merge": None,
    }


def test_inline_and_formula_string_use_shared_unicode_decoder(tmp_path: Path) -> None:
    path = source(
        tmp_path,
        '<c r="A1" t="inlineStr"><is><r><t> x_x000D_</t></r>'
        '<rPh sb="0" eb="1"><t>ignore</t></rPh>'
        "<r><t>_xD83D__xDE00_</t></r></is></c>"
        '<c r="B1" t="str"><f>"x"</f><v>_x005F_x000D_</v></c>',
        sheet_attrs='xml:space="preserve"',
    )
    with Workbook.open(path) as book:
        observed = cells(book.read_cells("Inputs", ["A1", "B1"]))
    assert observed[0]["stored_value"] == " x\r😀"
    assert observed[1]["raw"] == "_x005F_x000D_"
    assert observed[1]["stored_value"] == "_x000D_"


def test_shared_formula_master_and_childless_merge_are_not_fabricated(tmp_path: Path) -> None:
    path = source(
        tmp_path,
        '<c r="A1"><f t="shared" si="7" ref="A1:B1" ca="1">C1*2</f><v>4</v></c>'
        '<c r="B1"><f t="shared" si="7"/><v>6</v></c>',
        extra='<mergeCells><mergeCell ref="C1:D1"/></mergeCells>',
    )
    with Workbook.open(path) as book:
        observed = cells(book.read_cells("Inputs", ["A1", "B1", "C1", "D1"]))
    master = {"address": "A1", "ref": "A1:B1", "text": "C1*2"}
    assert observed[0]["formula"] == {
        "text": "C1*2",
        "kind": "shared",
        "shared_index": 7,
        "ref": "A1:B1",
        "calculate_on_next_recalc": True,
        "shared_master": master,
    }
    assert observed[1]["formula"] == {
        "text": None,
        "kind": "shared",
        "shared_index": 7,
        "ref": None,
        "calculate_on_next_recalc": False,
        "shared_master": master,
    }
    assert observed[2]["merge"] == {"ref": "C1:D1", "root": "C1", "role": "root"}
    assert observed[3]["merge"] == {"ref": "C1:D1", "root": "C1", "role": "child"}
    assert observed[2]["presence"] == observed[3]["presence"] == "missing"
    assert observed[3]["stored_value"] is None


def test_style_zero_and_temporal_facts_do_not_convert_values(tmp_path: Path) -> None:
    path = source(
        tmp_path,
        '<c r="A1"><v>45658</v></c><c r="B1" t="d"><v>2025-01-01</v></c>'
        '<c r="C1" t="inlineStr"><is><t>text</t></is></c>',
    )
    with zipfile.ZipFile(path) as zf:
        relationships = zf.read("xl/_rels/workbook.xml.rels").decode()
    rewrite(
        path,
        {
            "xl/_rels/workbook.xml.rels": relationships.replace(
                "</Relationships>",
                '<Relationship Id="styles" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
                'Target="styles.xml"/></Relationships>',
            ).encode(),
            "xl/styles.xml": (
                f'<styleSheet xmlns="{NS}"><cellXfs><xf numFmtId="14"/></cellXfs></styleSheet>'
            ).encode(),
        },
    )
    with Workbook.open(path) as book:
        observed = cells(book.read_cells("Inputs", ["A1", "B1", "C1"]))
    assert [c["stored_value"] for c in observed] == [45658.0, "2025-01-01", "text"]
    assert [c["value_kind"] for c in observed] == ["number", "iso_date", "string"]
    assert all(
        c["cell_format"]
        == {
            "style_index": 0,
            "format_id": 14,
            "format_code": "mm-dd-yy",
            "temporal_kind": "date",
            "basis": "cell_xf_number_format",
        }
        for c in observed
    )


@pytest.mark.parametrize(
    "name,tab,addresses,values",
    [
        (
            "test_workbook_7_write_v11_edges.xlsx",
            "edges",
            ["A1", "D1", "C10"],
            [100.0, 45658.0, None],
        ),
        ("test_workbook_6b_1904.xlsx", "Date 1904 Test", ["A1"], ["Label"]),
    ],
)
def test_real_fixture_observations(
    name: str, tab: str, addresses: list[str], values: list[object]
) -> None:
    with Workbook.open(FIXTURES / name) as book:
        assert [c["stored_value"] for c in cells(book.read_cells(tab, addresses))] == values


def test_registry_fixture_inventory_includes_hidden_tabs() -> None:
    with Workbook.open(FIXTURES / "test_workbook_2_registry.xlsx") as book:
        inventory = data(book.list_sheets())
    assert inventory == {
        "sheets": [
            {"name": "USC - Hedging", "state": "visible", "kind": "worksheet", "tab_index": 0},
            {"name": "Visible Two", "state": "visible", "kind": "worksheet", "tab_index": 1},
            {"name": "Hidden Sheet", "state": "hidden", "kind": "worksheet", "tab_index": 2},
            {"name": "Very Hidden", "state": "veryHidden", "kind": "worksheet", "tab_index": 3},
            {"name": "Visible One", "state": "visible", "kind": "worksheet", "tab_index": 4},
        ],
        "active_tab": 4,
        "date1904": False,
        "defined_name_count": 4,
    }


@pytest.mark.parametrize(
    "addresses",
    [
        [],
        ["A1", "a1"],
        ["A1٢"],
        ["XFE1"],
        ["A1048577"],
        [" A1"],
        ["$A$1"],
        ["Inputs!A1"],
        ["A1:B1"],
    ],
)
def test_invalid_request_preserves_healthy_session(tmp_path: Path, addresses: list[str]) -> None:
    with Workbook.open(source(tmp_path, "")) as book:
        with pytest.raises(ValueError):
            book.read_cells("Inputs", addresses)
        assert not book.closed and isinstance(book.list_sheets(), Inspection)


@pytest.mark.parametrize("addresses", ["A1", b"A1", None, iter(["A1"]), [1]])
def test_wrong_request_types_are_programming_errors(tmp_path: Path, addresses: object) -> None:
    with Workbook.open(source(tmp_path, "")) as book:
        with pytest.raises(TypeError):
            book.read_cells("Inputs", addresses)  # type: ignore[arg-type]
        assert not book.closed


def test_address_grid_edges_and_requested_order(tmp_path: Path) -> None:
    with Workbook.open(source(tmp_path, "")) as book:
        result = data(book.read_cells("Inputs", ["xfd1048576", "a1", "Z999999"]))
    assert result["requested_addresses"] == ["XFD1048576", "A1", "Z999999"]


def test_cell_count_n_and_n_plus_one_before_parsing(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    with Workbook.open(source(tmp_path, "")) as book:
        assert isinstance(
            book.read_cells("Inputs", ["A1", "B1"], limits=ReadLimits(max_cells=2)), Inspection
        )

        def forbidden(_name: str) -> None:
            pytest.fail("over-limit request parsed a sheet")

        monkeypatch.setattr(book._book, "read_sheet", forbidden)
        result = book.read_cells("Inputs", ["A1", "B1", "C1"], limits=ReadLimits(max_cells=2))
        assert isinstance(result, Refusal) and result.code == "inspection_limit_exceeded"
        assert result.details == {
            "source_fingerprint": book.source_fingerprint,
            "budget": "max_cells",
            "limit": 2,
            "requested_cells": 3,
        }


def test_exact_response_bytes_n_and_n_plus_one(tmp_path: Path) -> None:
    with Workbook.open(
        source(tmp_path, '<c r="A1" t="inlineStr"><is><t>😀"\t</t></is></c>')
    ) as book:
        result = book.read_cells("Inputs", ["A1"])
        assert isinstance(result, Inspection)
        size = len(result.canonical_json())
        assert isinstance(
            book.read_cells("Inputs", ["A1"], limits=ReadLimits(max_response_bytes=size)),
            Inspection,
        )
        refused = book.read_cells("Inputs", ["A1"], limits=ReadLimits(max_response_bytes=size - 1))
        assert isinstance(refused, Refusal) and refused.details["budget"] == "max_response_bytes"
        assert refused.details["observed_bytes"] == size
        assert not book.closed and "cells" not in refused.details


def test_repeated_large_values_refuse_before_encoding(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from xlayer import _inspection

    # Repeated occurrences include raw + decoded text in distinct real cells.
    payload = "".join(f'<c r="{chr(65 + i)}1" t="str"><v>{"x" * 32767}</v></c>' for i in range(20))
    path = source(tmp_path, payload)
    with Workbook.open(path) as book:

        def forbidden(_value: object) -> bytes:
            pytest.fail("large repeated response was encoded before preflight")

        monkeypatch.setattr(_inspection, "canonical_json", forbidden)
        result = book.read_cells("Inputs", [f"{chr(65 + i)}1" for i in range(20)])
        assert isinstance(result, Refusal) and result.details["budget"] == "max_response_bytes"
        assert "observed_minimum_bytes" in result.details


def test_unicode_expansion_is_checked_not_silently_cropped(tmp_path: Path) -> None:
    payload = "".join(
        f'<c r="{chr(65 + i)}1" t="inlineStr"><is><t>{"😀" * 16383}</t></is></c>' for i in range(6)
    )
    with Workbook.open(source(tmp_path, payload)) as book:
        result = book.read_cells("Inputs", [f"{chr(65 + i)}1" for i in range(6)])
        assert isinstance(result, Refusal)
        size = result.details["observed_bytes"]
        assert isinstance(size, int) and size > 1_048_576
        assert not book.closed


def test_frozen_detached_result_and_canonical_envelope(tmp_path: Path) -> None:
    with Workbook.open(source(tmp_path, '<c r="A1"><v>2</v></c>')) as book:
        addresses = ["a1"]
        result = book.read_cells("Inputs", addresses)
        assert isinstance(result, Inspection)
        addresses[0] = "B1"
        encoded = result.canonical_json()
        output = result.to_dict()
        mutable = cast("dict[str, object]", output["data"])
        mutable["cells"] = []
        assert result.canonical_json() == encoded and len(cells(result)) == 1
        with pytest.raises(TypeError):
            result.evidence["data"] = {}  # type: ignore[index]
        with pytest.raises(FrozenInstanceError):
            result.evidence = {}  # type: ignore[misc]
    assert json.loads(encoded) == result.to_dict()
    assert set(result.to_dict()) == {
        "schema_version",
        "inspection_contract_version",
        "kind",
        "source_fingerprint",
        "data",
    }


@pytest.mark.parametrize("change", ["delete", "replace"])
def test_read_snapshot_survives_source_change(tmp_path: Path, change: str) -> None:
    path = source(tmp_path, '<c r="A1"><v>9</v></c>')
    original = path.read_bytes()
    with Workbook.open(path) as book:
        if change == "delete":
            path.unlink()
        else:
            path.write_bytes(b"replacement")
        assert cells(book.read_cells("Inputs", ["A1"]))[0]["stored_value"] == 9.0
        assert book.source_fingerprint == "sha256:" + hashlib.sha256(original).hexdigest()


def test_malformed_sheet_and_missing_sheet_are_not_missing_cells(tmp_path: Path) -> None:
    with Workbook.open(source(tmp_path, '<c r="A1"><v>nan</v></c>')) as book:
        malformed = book.read_cells("Inputs", ["Z1"])
        assert isinstance(malformed, Refusal) and malformed.code == "invalid_part_content"
        missing = book.read_cells("inputs", ["A1"])
        assert isinstance(missing, Refusal) and missing.code == "sheet_not_found"
        assert not book.closed


@pytest.mark.parametrize("method", ["list_sheets", "read_cells"])
def test_unexpected_projection_error_closes_and_propagates(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, method: str
) -> None:
    from xlayer import _public_workbook

    book = Workbook.open(source(tmp_path, '<c r="A1"><v>2</v></c>'))
    failure = ValueError("unexpected projection defect")

    def broken(*_args: object, **_kwargs: object) -> None:
        raise failure

    monkeypatch.setattr(_public_workbook, "make_inspection", broken)
    with pytest.raises(ValueError) as caught:
        book.list_sheets() if method == "list_sheets" else book.read_cells("Inputs", ["A1"])
    assert caught.value is failure and book.closed


class BrokenSequence(Sequence[str]):
    def __len__(self) -> int:
        return 1

    @overload
    def __getitem__(self, index: int) -> str: ...
    @overload
    def __getitem__(self, index: slice) -> Sequence[str]: ...
    def __getitem__(self, index: int | slice) -> str | Sequence[str]:
        raise RuntimeError("broken request sequence")


def test_unexpected_request_capture_error_closes_session(tmp_path: Path) -> None:
    book = Workbook.open(source(tmp_path, ""))
    with pytest.raises(RuntimeError, match="broken request sequence"):
        book.read_cells("Inputs", BrokenSequence())
    assert book.closed


@pytest.mark.parametrize("field", ["max_cells", "max_response_bytes"])
@pytest.mark.parametrize("value", [True, False, 0, -1, 1.0, "1", 2**63])
def test_read_limit_fields_are_exact_bounded_integers(field: str, value: object) -> None:
    with pytest.raises(ValueError, match=field):
        ReadLimits(**{field: value})  # type: ignore[arg-type]


def test_public_archive_ceiling_cannot_be_raised(tmp_path: Path) -> None:
    for spec in fields(ArchiveLimits()):
        too_high = replace(ArchiveLimits(), **{spec.name: getattr(ArchiveLimits(), spec.name) + 1})
        with pytest.raises(ValueError, match=spec.name):
            Workbook.open(source(tmp_path, ""), limits=too_high)


def test_public_open_failure_is_structured(tmp_path: Path) -> None:
    with pytest.raises(WorkbookOpenError) as caught:
        Workbook.open(tmp_path / "absent.xlsx")
    assert caught.value.refusal.recovery_options


def test_closed_check_precedes_validation_and_cache(tmp_path: Path) -> None:
    book = Workbook.open(source(tmp_path, ""))
    assert isinstance(book.list_sheets(), Inspection)
    book.close()
    book.close()
    with pytest.raises(ClosedWorkbookError):
        book.read_cells("Inputs", [])
    with pytest.raises(ClosedWorkbookError):
        book.list_sheets(limits=None)  # type: ignore[arg-type]
    with pytest.raises(ClosedWorkbookError):
        book.__enter__()
    assert book.source_path.name == "source.xlsx" and book.source_fingerprint.startswith("sha256:")


def test_unsupported_chart_tab_is_inventoried_but_not_read(tmp_path: Path) -> None:
    path = source(tmp_path, "")
    with zipfile.ZipFile(path) as zf:
        rels = zf.read("xl/_rels/workbook.xml.rels")
    rewrite(
        path,
        {
            "xl/_rels/workbook.xml.rels": rels.replace(
                b"relationships/worksheet", b"relationships/chartsheet"
            )
        },
    )
    with Workbook.open(path) as book:
        assert data(book.list_sheets())["sheets"] == [
            {"name": "Inputs", "state": "visible", "kind": "chartsheet", "tab_index": 0}
        ]
        result = book.read_cells("Inputs", ["A1"])
        assert isinstance(result, Refusal) and result.code == "unsupported_sheet_kind"
        assert not book.closed


def test_inventory_has_exact_byte_bound_without_partial_tab_list(tmp_path: Path) -> None:
    with Workbook.open(source(tmp_path, "")) as book:
        result = book.list_sheets()
        assert isinstance(result, Inspection)
        size = len(result.canonical_json())
        assert isinstance(book.list_sheets(limits=ReadLimits(max_response_bytes=size)), Inspection)
        result = book.list_sheets(limits=ReadLimits(max_response_bytes=size - 1))
        assert isinstance(result, Refusal) and result.code == "inspection_limit_exceeded"
        assert "sheets" not in result.details


def test_unknown_format_is_not_non_temporal_or_effective_style(tmp_path: Path) -> None:
    path = source(tmp_path, '<c r="A1"><v>5</v></c>')
    with zipfile.ZipFile(path) as zf:
        rels = zf.read("xl/_rels/workbook.xml.rels").decode()
    rewrite(
        path,
        {
            "xl/_rels/workbook.xml.rels": rels.replace(
                "</Relationships>",
                '<Relationship Id="styles" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
                'Target="styles.xml"/></Relationships>',
            ).encode(),
            "xl/styles.xml": (
                f'<styleSheet xmlns="{NS}"><cellXfs><xf numFmtId="300"/></cellXfs></styleSheet>'
            ).encode(),
        },
    )
    with Workbook.open(path) as book:
        value = cells(book.read_cells("Inputs", ["A1"]))[0]
    assert value["cell_format"] == {
        "style_index": 0,
        "format_id": 300,
        "format_code": "",
        "temporal_kind": "unknown",
        "basis": "cell_xf_number_format",
    }
    assert value["stored_value"] == 5.0


def test_public_facade_construction_failure_releases_engine(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from xlayer import _public_workbook
    from xlayer._workbook import Workbook as EngineWorkbook

    captured: list[EngineWorkbook] = []
    failure = RuntimeError("facade construction defect")

    def broken(_self: Workbook, book: EngineWorkbook) -> None:
        captured.append(book)
        raise failure

    monkeypatch.setattr(_public_workbook.Workbook, "__init__", broken)
    with pytest.raises(RuntimeError) as caught:
        Workbook.open(source(tmp_path, ""))
    assert caught.value is failure and len(captured) == 1
    assert captured[0].closed is True


@pytest.mark.parametrize(
    "value",
    [None, True, False, 0, -123, 1.25, '😀\n"\\', ["same", "same"], {"😀": [1, {"key": "\x00"}]}],
)
def test_json_preflight_is_a_lower_bound_including_repeated_values(value: object) -> None:
    from xlayer._canonical import canonical_json
    from xlayer._inspection import minimum_json_bytes

    assert minimum_json_bytes(value, 10000) <= len(canonical_json(value))


@pytest.mark.parametrize("operation", ["list", "read"])
def test_wrong_limits_type_is_argument_error_not_session_failure(
    tmp_path: Path, operation: str
) -> None:
    with Workbook.open(source(tmp_path, "")) as book:
        with pytest.raises(TypeError):
            if operation == "list":
                book.list_sheets(limits=ArchiveLimits())  # type: ignore[arg-type]
            else:
                book.read_cells("Inputs", ["A1"], limits=None)  # type: ignore[arg-type]
        assert not book.closed
