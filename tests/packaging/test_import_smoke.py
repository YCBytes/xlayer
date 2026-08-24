"""Packaging gates for the pre-release skeleton.

These prove the distribution imports in a clean environment, declares no
runtime dependency on the ``xlayer-core`` research repository, and claims no
workbook capability.
"""

from __future__ import annotations

import importlib.metadata
import importlib.util
from pathlib import Path

import xlayer


def test_package_imports_and_reports_version() -> None:
    assert xlayer.__version__ == "0.1.0.dev0"


def test_declares_no_runtime_dependencies() -> None:
    declared = importlib.metadata.requires("xlayer") or []
    runtime = [req for req in declared if "extra ==" not in req]
    assert runtime == []


def test_xlayer_core_is_not_importable() -> None:
    assert importlib.util.find_spec("xlayer_core") is None


def test_no_workbook_capability_is_exported() -> None:
    assert xlayer.__all__ == ["__version__"]
    public = [name for name in vars(xlayer) if not name.startswith("_")]
    assert public == []


def test_distribution_is_typed() -> None:
    # Checked against the installed package directory rather than distribution
    # metadata, because an editable install reports only its import shim.
    package_dir = Path(xlayer.__file__).parent
    assert (package_dir / "py.typed").is_file()
