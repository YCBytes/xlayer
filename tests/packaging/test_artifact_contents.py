"""Artifact content gates.

The wheel and sdist must contain only intended files. This is a permanent gate:
a ``docs/`` glob once picked up an internal design note that would have shipped
inside the public sdist, so artifact contents are asserted rather than assumed.
"""

from __future__ import annotations

import tarfile
import zipfile
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[2]
DIST_DIR = REPO_ROOT / "dist"

# Everything the wheel may contain outside its .dist-info metadata directory.
# Slice 1 adds the refusal envelope and the internal archive, package,
# and workbook-registry layers; all are underscore-prefixed and nothing is
# exported publicly until the full slice passes its gates.
WHEEL_PAYLOAD = frozenset(
    {
        "xlayer/__init__.py",
        "xlayer/py.typed",
        "xlayer/_errors.py",
        "xlayer/_ooxml/__init__.py",
        "xlayer/_ooxml/archive.py",
        "xlayer/_ooxml/package.py",
        "xlayer/_ooxml/workbook.py",
    }
)

# What the sdist may contain, after the top-level version directory is
# stripped. Exact filenames and directory prefixes are distinguished so that,
# for example, "LICENSE" cannot accidentally admit "LICENSE2". Nothing under
# docs/ is listed: shipping documentation is a deliberate decision, so new
# docs must be added here before the gate will let them through.
SDIST_ALLOWED_FILES = frozenset(
    {
        "LICENSE",
        "README.md",
        "pyproject.toml",
        "PKG-INFO",
        ".gitignore",
    }
)
SDIST_ALLOWED_DIR_PREFIXES = ("src/", "tests/")


def _built_artifact(pattern: str) -> Path:
    matches = sorted(DIST_DIR.glob(pattern)) if DIST_DIR.is_dir() else []
    if not matches:
        pytest.skip(f"no built artifact matching {pattern!r}; run `python -m build` first")
    return matches[-1]


def test_wheel_contains_only_intended_files() -> None:
    with zipfile.ZipFile(_built_artifact("*.whl")) as archive:
        names = set(archive.namelist())

    payload = {name for name in names if ".dist-info/" not in name}
    assert payload == set(WHEEL_PAYLOAD)


def test_wheel_ships_the_license() -> None:
    with zipfile.ZipFile(_built_artifact("*.whl")) as archive:
        names = archive.namelist()

    assert any(name.endswith(".dist-info/licenses/LICENSE") for name in names)


# Files that must be present in every sdist, exactly. The complete source
# payload is derived from WHEEL_PAYLOAD so the sdist cannot silently drop a
# module the wheel ships (hatchling ignores missing only-include entries
# rather than erroring).
REQUIRED_SDIST_FILES = (
    frozenset(
        {
            "LICENSE",
            "README.md",
            "pyproject.toml",
            "PKG-INFO",
        }
    )
    | {f"src/{payload_path}" for payload_path in WHEEL_PAYLOAD}
    | {
        "tests/fixtures/test_workbook_2_registry.xlsx",
        "tests/fixtures/test_workbook_2_registry.expected.json",
        "tests/fixtures/test_workbook_6b_1904.xlsx",
        "tests/fixtures/test_workbook_6b_1904.expected.json",
    }
)


def _sdist_relative_files() -> list[str]:
    with tarfile.open(_built_artifact("*.tar.gz")) as archive:
        members = [member.name for member in archive.getmembers() if member.isfile()]
    # Strip the leading "<name>-<version>/" directory every sdist carries.
    return [name.split("/", 1)[1] for name in members if "/" in name]


def test_sdist_contains_only_intended_files() -> None:
    relative = _sdist_relative_files()
    unexpected = [
        name
        for name in relative
        if name not in SDIST_ALLOWED_FILES and not name.startswith(SDIST_ALLOWED_DIR_PREFIXES)
    ]
    assert unexpected == []


def test_sdist_contains_required_files() -> None:
    relative = set(_sdist_relative_files())
    missing = REQUIRED_SDIST_FILES - relative
    assert not missing, f"sdist is missing required files: {sorted(missing)}"
