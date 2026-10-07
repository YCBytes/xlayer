"""Artifact content gates.

The wheel and sdist must contain only intended files. This is a permanent gate:
a ``docs/`` glob once picked up an internal design note that would have shipped
inside the public sdist, so artifact contents are asserted rather than assumed.
"""

from __future__ import annotations

import io
import os
import re
import tarfile
import zipfile
from collections.abc import Collection, Mapping
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[2]
DIST_DIR = REPO_ROOT / "dist"
APPROVED_XLSX = frozenset(
    {
        "tests/fixtures/test_workbook_2_registry.xlsx",
        "tests/fixtures/test_workbook_3_formats.xlsx",
        "tests/fixtures/test_workbook_4_worksheet_final.xlsx",
        "tests/fixtures/test_workbook_6b_1904.xlsx",
        "tests/fixtures/test_workbook_7_write_v11_edges.xlsx",
    }
)
_HOME_PATH = re.compile(
    rb"/(?:Users|home)/[^/\s<>\"']+/|[A-Za-z]:[\\/]Users[\\/][^\\/\s<>\"']+[\\/]"
)

# Everything the wheel may contain outside its .dist-info metadata directory.
# Slice 1 adds the refusal envelope and the internal archive, package,
# workbook-registry, shared-string, styles, text-decode, and worksheet
# layers plus the managed read-only coordinator; all are underscore-prefixed
# and nothing is exported publicly
# until the full slice passes its gates.
WHEEL_PAYLOAD = frozenset(
    {
        "xlayer/__init__.py",
        "xlayer/py.typed",
        "xlayer/_errors.py",
        "xlayer/_workbook.py",
        "xlayer/_dependencies.py",
        "xlayer/_canonical.py",
        "xlayer/_edits.py",
        "xlayer/_approval.py",
        "xlayer/_preview.py",
        "xlayer/_receipt.py",
        "xlayer/_batch_impact.py",
        "xlayer/_edit_validation.py",
        "xlayer/_ooxml/__init__.py",
        "xlayer/_ooxml/archive.py",
        "xlayer/_ooxml/package.py",
        "xlayer/_ooxml/workbook.py",
        "xlayer/_ooxml/strings.py",
        "xlayer/_ooxml/styles.py",
        "xlayer/_ooxml/_text.py",
        "xlayer/_ooxml/sheet.py",
        "xlayer/_ooxml/formula.py",
        "xlayer/_ooxml/patch.py",
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
    selected = os.environ.get("XLAYER_ARTIFACT_DIR")
    directory = Path(selected) if selected is not None else DIST_DIR
    matches = sorted(directory.glob(pattern)) if directory.is_dir() else []
    if not matches:
        if selected is not None:
            pytest.fail(f"missing fresh artifact matching {pattern!r} in {directory}")
        pytest.skip(f"no built artifact matching {pattern!r}; run `python -m build` first")
    assert len(matches) == 1, f"artifact selection must be unambiguous: {matches}"
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
        "tests/fixtures/test_workbook_2_registry.strings.json",
        "tests/fixtures/test_workbook_6b_1904.xlsx",
        "tests/fixtures/test_workbook_6b_1904.expected.json",
        "tests/fixtures/test_workbook_6b_1904.styles.json",
        "tests/fixtures/test_workbook_3_formats.xlsx",
        "tests/fixtures/test_workbook_3_formats.strings.json",
        "tests/fixtures/test_workbook_3_formats.styles.json",
        "tests/fixtures/test_workbook_4_worksheet_final.xlsx",
        "tests/fixtures/test_workbook_4_worksheet_final.sheet.json",
        "tests/fixtures/test_workbook_7_write_v11_edges.xlsx",
        "tests/fixtures/test_workbook_7_write_v11_edges.sheet.json",
        "tests/packaging/test_workbook_smoke.py",
        "tests/packaging/test_dependency_smoke.py",
        "tests/__init__.py",
        "tests/unit/__init__.py",
        "tests/unit/_dependency_cases.py",
        "tests/unit/test_formula.py",
        "tests/unit/test_dependencies.py",
        "tests/unit/test_dependency_session.py",
        "tests/fixtures/dependency_complete.expected.json",
        "tests/fixtures/dependency_edge_cutoff.expected.json",
        "tests/fixtures/test_workbook_4_worksheet_final.dependencies.json",
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
    _assert_required_sdist(relative)


def _assert_required_sdist(relative: Collection[str]) -> None:
    missing = REQUIRED_SDIST_FILES - set(relative)
    assert not missing, f"sdist is missing required files: {sorted(missing)}"


def _wheel_contents(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _sdist_contents(path: Path) -> dict[str, bytes]:
    result = {}
    with tarfile.open(path) as archive:
        for member in archive.getmembers():
            if member.isfile() and "/" in member.name:
                data = archive.extractfile(member)
                assert data is not None
                result[member.name.split("/", 1)[1]] = data.read()
    return result


def _assert_current_sources(wheel: Mapping[str, bytes], sdist: Mapping[str, bytes]) -> None:
    for name in WHEEL_PAYLOAD:
        current = (REPO_ROOT / "src" / name).read_bytes()
        assert wheel.get(name) == current, f"wheel source bytes differ: {name}"
        assert sdist.get(f"src/{name}") == current, f"sdist source bytes differ: {name}"


def _assert_published_payload(payload: Mapping[str, bytes]) -> None:
    for name, data in payload.items():
        assert not {"internal-docs", "internal-notes"} & set(name.split("/")), (
            f"private document in payload: {name}"
        )
        if name.endswith(".xlsx"):
            assert name in APPROVED_XLSX, f"unapproved binary fixture: {name}"
            with zipfile.ZipFile(io.BytesIO(data)) as workbook:
                for part in workbook.namelist():
                    if part.endswith((".xml", ".rels")):
                        assert not _HOME_PATH.search(workbook.read(part)), (
                            f"personal home path in fixture: {name}:{part}"
                        )
        else:
            assert not _HOME_PATH.search(data), f"personal home path in text: {name}"


def test_explicit_fresh_artifact_directory_cannot_silently_skip(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("XLAYER_ARTIFACT_DIR", str(tmp_path))
    with pytest.raises(pytest.fail.Exception, match="fresh artifact"):
        _built_artifact("*.whl")


def test_artifact_sources_match_current_checkout() -> None:
    _assert_current_sources(
        _wheel_contents(_built_artifact("*.whl")),
        _sdist_contents(_built_artifact("*.tar.gz")),
    )


def test_changed_artifact_byte_is_detected() -> None:
    wheel = {name: (REPO_ROOT / "src" / name).read_bytes() for name in WHEEL_PAYLOAD}
    sdist = {f"src/{name}": payload for name, payload in wheel.items()}
    wheel["xlayer/_dependencies.py"] += b"\n# changed in a test-temp payload\n"
    with pytest.raises(AssertionError, match="source bytes"):
        _assert_current_sources(wheel, sdist)
    wheel["xlayer/_dependencies.py"] = (REPO_ROOT / "src/xlayer/_dependencies.py").read_bytes()
    sdist["src/xlayer/_ooxml/formula.py"] += b"\n# changed\n"
    with pytest.raises(AssertionError, match="source bytes"):
        _assert_current_sources(wheel, sdist)


@pytest.mark.parametrize(
    "missing", ["tests/__init__.py", "tests/unit/__init__.py", "tests/unit/_dependency_cases.py"]
)
def test_required_dependency_helper_and_package_markers_cannot_drop(missing: str) -> None:
    # The gate must reject loss of any helper import boundary, not merely a module filename.
    files = set(REQUIRED_SDIST_FILES)
    files.discard(missing)
    with pytest.raises(AssertionError, match="missing required files"):
        _assert_required_sdist(files)


def test_published_payload_has_no_private_content() -> None:
    wheel = _wheel_contents(_built_artifact("*.whl"))
    sdist = _sdist_contents(_built_artifact("*.tar.gz"))
    _assert_published_payload(wheel)
    _assert_published_payload(sdist)
    assert {name for name in sdist if name.endswith(".xlsx")} == APPROVED_XLSX


def test_privacy_guard_detects_neutral_test_probes() -> None:
    # Construct invented home paths at runtime so no personal path is shipped
    # inside this very test's source. No private corpus strings are used.
    probe = ("/" + "Users" + "/" + "example-person" + "/notes").encode()
    with pytest.raises(AssertionError, match="personal home path"):
        _assert_published_payload({"tests/probe.txt": probe})
    with pytest.raises(AssertionError, match="private document"):
        _assert_published_payload({"internal-docs/probe.md": b"neutral marker"})
