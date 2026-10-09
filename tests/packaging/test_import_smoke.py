"""Exact supported root imports, installed typing, and zero runtime dependencies."""

from __future__ import annotations

import importlib
import importlib.metadata
import importlib.util
from pathlib import Path

import xlayer


def test_package_imports_and_reports_version() -> None:
    assert xlayer.__version__ == "0.1.0a1"
    assert importlib.metadata.version("xlayer") == xlayer.__version__


def test_declares_no_runtime_dependencies() -> None:
    declared = importlib.metadata.requires("xlayer") or []
    runtime = [req for req in declared if "extra ==" not in req]
    assert runtime == []


def test_xlayer_core_is_not_importable() -> None:
    assert importlib.util.find_spec("xlayer_core") is None


def test_exact_alpha_surface_after_internal_imports() -> None:
    expected = {
        "Workbook",
        "SetValue",
        "Proposal",
        "Inspection",
        "Preview",
        "Receipt",
        "Approval",
        "VerifiedApproval",
        "Refusal",
        "XlayerError",
        "WorkbookOpenError",
        "ClosedWorkbookError",
        "ArchiveLimits",
        "ImpactLimits",
        "ReadLimits",
    }
    for name in ("_workbook", "_proposal", "_ooxml.archive", "_ooxml.sheet", "_dependencies"):
        importlib.import_module(f"xlayer.{name}")
    assert set(xlayer.__all__) == {"__version__", *expected}
    assert len(xlayer.__all__) == len(expected) + 1
    assert {name for name in vars(xlayer) if not name.startswith("_")} == expected


def test_distribution_is_typed() -> None:
    # Checked against the installed package directory rather than distribution
    # metadata, because an editable install reports only its import shim.
    package_dir = Path(xlayer.__file__).parent
    assert (package_dir / "py.typed").is_file()
