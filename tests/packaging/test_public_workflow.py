"""The supported root workflow and the actual provider-free application example."""

from __future__ import annotations

import importlib.util
import os
import sys
from dataclasses import replace
from pathlib import Path
from types import ModuleType
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

from xlayer import Approval, Inspection, Preview, Proposal, Receipt, Refusal, SetValue, Workbook

ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "tests/fixtures/test_workbook_7_write_v11_edges.xlsx"


def example() -> ModuleType:
    path = ROOT / "examples/approved_set_value.py"
    assert path.is_file(), "the real runnable application example must ship with the sdist"
    spec = importlib.util.spec_from_file_location("approved_set_value_example", path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_installed_public_workflow_real_fixture_and_returned_receipt(tmp_path: Path) -> None:
    if os.environ.get("XLAYER_REQUIRE_INSTALLED") == "1":
        import xlayer

        assert "site-packages" in Path(xlayer.__file__).resolve().parts
        assert "xlayer_core" not in sys.modules
    before = FIXTURE.read_bytes()
    output = tmp_path / "updated.xlsx"
    module = example()
    result = module.edit_one_cell(FIXTURE, output, SetValue("edges", "F1", 44), permitted=True)
    assert isinstance(result, Receipt), result
    assert result.evidence["visual_status"] == "not_evaluated"
    assert result.evidence["verification"]["all_passed"] is True  # type: ignore[index]
    with Workbook.open(output) as book:
        read = book.read_cells("edges", ["F1"])
        assert isinstance(read, Inspection)
        assert read.to_dict()["data"]["cells"][0]["raw"] == "44"  # type: ignore[index]
    with ZipFile(output) as edited, ZipFile(FIXTURE) as original:
        changed = [n for n in original.namelist() if edited.read(n) != original.read(n)]
        assert changed == ["xl/worksheets/sheet1.xml"]
        root = ET.fromstring(edited.read(changed[0]))  # noqa: S314
        node = root.find('.//{*}c[@r="F1"]/{*}v')
        assert node is not None and node.text == "44"
    assert FIXTURE.read_bytes() == before and not list(tmp_path.glob(".xlayer-*"))


def test_example_host_denial_writes_nothing(tmp_path: Path) -> None:
    result = example().edit_one_cell(
        FIXTURE, tmp_path / "out.xlsx", SetValue("edges", "F1", 44), permitted=False
    )
    assert isinstance(result, Refusal) and result.code == "approval_required"
    assert not (tmp_path / "out.xlsx").exists() and not list(tmp_path.glob(".xlayer-*"))


def test_example_store_denies_unrecorded_or_changed_exact_preview(tmp_path: Path) -> None:
    store = example().ExactPreviewApprovalStore()
    with Workbook.open(FIXTURE) as book:
        proposed = book.propose([SetValue("edges", "F1", 44)], output_path=tmp_path / "out.xlsx")
        assert isinstance(proposed, Proposal)
        preview = proposed.preview()
        assert isinstance(preview, Preview)
        unrecorded = Approval(
            "1.0",
            "invented",
            "demo-host",
            "service",
            book.source_fingerprint,
            proposed.proposal_digest,
            preview.preview_digest,
            "now",
        )
        assert store.verify(unrecorded).authorized is False
        recorded = store.record(proposed, preview, permitted=True)
        assert isinstance(recorded, Approval) and store.verify(recorded).authorized is True
        tampered = replace(recorded, preview_digest="sha256:" + "0" * 64)
        assert store.verify(tampered).authorized is False
        result = proposed.apply(approval=tampered, verify_approval=store.verify)
        assert isinstance(result, Refusal) and result.code == "preview_digest_mismatch"
        assert not (tmp_path / "out.xlsx").exists()


def test_example_does_not_auto_grant_populated_text_or_partial_coverage(tmp_path: Path) -> None:
    store = example().ExactPreviewApprovalStore()
    with Workbook.open(FIXTURE) as book:
        proposed = book.propose(
            [SetValue("edges", "B1", "replacement")], output_path=tmp_path / "out.xlsx"
        )
        assert isinstance(proposed, Proposal)
        preview = proposed.preview()
        assert isinstance(preview, Preview)
        assert store.record(proposed, preview, permitted=True) is None
        result = proposed.apply(approval=None, verify_approval=store.verify)
        assert isinstance(result, Refusal)
    assert not (tmp_path / "out.xlsx").exists()


def test_example_rejects_preview_from_another_transaction(tmp_path: Path) -> None:
    store = example().ExactPreviewApprovalStore()
    with Workbook.open(FIXTURE) as book:
        first = book.propose([SetValue("edges", "F1", 44)], output_path=tmp_path / "one.xlsx")
        second = book.propose([SetValue("edges", "F1", 45)], output_path=tmp_path / "two.xlsx")
        assert isinstance(first, Proposal) and isinstance(second, Proposal)
        preview = second.preview()
        assert isinstance(preview, Preview)
        assert store.record(first, preview, permitted=True) is None


@pytest.mark.parametrize(
    "relative", ["README.md", "docs/api.md", "docs/support.md", "CONTRIBUTING.md", "CHANGELOG.md"]
)
def test_deliberate_public_docs_exist_and_do_not_link_private_notes(relative: str) -> None:
    path = ROOT / relative
    assert path.is_file()
    text = path.read_text(encoding="utf-8")
    assert "internal-docs/" not in text and "internal-evals/" not in text


def test_readme_and_api_describe_alpha_root_surface_honestly() -> None:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    api = (ROOT / "docs/api.md").read_text(encoding="utf-8")
    assert "0.1.0a1" in readme and "pip install xlayer" not in readme
    assert "no supported workbook API" not in readme.lower()
    for name in (
        "Workbook",
        "SetValue",
        "Proposal",
        "Inspection",
        "ReadLimits",
        "Approval",
        "VerifiedApproval",
        "ArchiveLimits",
        "ImpactLimits",
    ):
        assert name in api
    assert "saved_formula_cache" in api and "not_verified" in api
