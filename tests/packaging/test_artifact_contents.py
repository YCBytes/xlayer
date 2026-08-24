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
WHEEL_PAYLOAD = frozenset({"xlayer/__init__.py", "xlayer/py.typed"})

# Path prefixes the sdist may contain, after the top-level version directory is
# stripped. Nothing under docs/ is listed: shipping documentation is a deliberate
# decision, so new docs must be added here before the gate will let them through.
SDIST_ALLOWED_PREFIXES = (
    "src/",
    "tests/",
    "LICENSE",
    "README.md",
    "pyproject.toml",
    "PKG-INFO",
    ".gitignore",
)


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


def test_sdist_contains_only_intended_files() -> None:
    with tarfile.open(_built_artifact("*.tar.gz")) as archive:
        members = [member.name for member in archive.getmembers() if member.isfile()]

    # Strip the leading "<name>-<version>/" directory every sdist carries.
    relative = [name.split("/", 1)[1] for name in members if "/" in name]
    unexpected = [name for name in relative if not name.startswith(SDIST_ALLOWED_PREFIXES)]
    assert unexpected == []
