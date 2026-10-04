"""Private, read-only ownership boundary over the OOXML readers.

Opening validates shared infrastructure, not every worksheet. Sheet records
are loaded on demand from one in-memory snapshot. Formula values are saved
caches, not verified calculation; this session makes no edit-safety claim.
Use a context manager or explicit close. No thread-safety guarantee is made.
"""

from __future__ import annotations

import os
from pathlib import Path
from types import TracebackType

from xlayer._errors import ClosedWorkbookError, Refusal, WorkbookOpenError
from xlayer._ooxml.archive import DEFAULT_LIMITS, ArchiveLimits, WorkbookArchive
from xlayer._ooxml.package import parse_package
from xlayer._ooxml.sheet import Worksheet, parse_worksheet
from xlayer._ooxml.strings import parse_shared_strings
from xlayer._ooxml.styles import StyleInfo, parse_styles
from xlayer._ooxml.workbook import WorkbookRegistry, parse_workbook_registry


class Workbook:
    """An internal managed snapshot; not a supported public workbook API."""

    def __init__(
        self,
        archive: WorkbookArchive,
        registry: WorkbookRegistry,
        shared_strings: tuple[str, ...],
        styles: tuple[StyleInfo, ...],
    ) -> None:
        self._source_path = archive.source_path
        self._source_fingerprint = archive.source_fingerprint
        self._archive: WorkbookArchive | None = archive
        self._registry: WorkbookRegistry | None = registry
        self._shared_strings = shared_strings
        self._styles = styles
        self._sheet_cache: dict[str, Worksheet | Refusal] = {}

    @classmethod
    def open(
        cls, path: str | os.PathLike[str], *, limits: ArchiveLimits = DEFAULT_LIMITS
    ) -> Workbook:
        """Load shared infrastructure; own no resource on an unsuccessful exit."""
        if not isinstance(limits, ArchiveLimits):
            raise TypeError("limits must be an ArchiveLimits instance")
        source = Path(path)
        archive: WorkbookArchive | Refusal | None = None
        try:
            archive = WorkbookArchive.load(source, limits)
            if isinstance(archive, Refusal):
                raise WorkbookOpenError(archive)
            package = parse_package(archive)
            if isinstance(package, Refusal):
                raise WorkbookOpenError(package)
            registry = parse_workbook_registry(archive, package)
            if isinstance(registry, Refusal):
                raise WorkbookOpenError(registry)
            shared_strings = parse_shared_strings(archive, package)
            if isinstance(shared_strings, Refusal):
                raise WorkbookOpenError(shared_strings)
            styles = parse_styles(archive, package)
            if isinstance(styles, Refusal):
                raise WorkbookOpenError(styles)
            return cls(archive, registry, shared_strings, styles)
        except BaseException:
            # Cleanup only: operational refusals and unexpected interruptions
            # keep their original exception/evidence after releasing ownership.
            if isinstance(archive, WorkbookArchive):
                archive.close()
            raise

    @property
    def source_path(self) -> Path:
        return self._source_path

    @property
    def source_fingerprint(self) -> str:
        """Identity of loaded bytes, never a second read of the source path."""
        return self._source_fingerprint

    @property
    def closed(self) -> bool:
        return self._archive is None

    @property
    def registry(self) -> WorkbookRegistry:
        if self._registry is None:
            raise ClosedWorkbookError("The workbook is closed.")
        return self._registry

    def read_sheet(self, name: str) -> Worksheet | Refusal:
        """Read one registered tab; cache its snapshot-derived outcome."""
        invalid_name = False
        try:
            archive = self._archive
            if archive is None:
                raise ClosedWorkbookError("The workbook is closed.")
            valid_name = isinstance(name, str)
            if not valid_name:
                invalid_name = True
                raise TypeError("sheet name must be a string")
            cached = self._sheet_cache.get(name)
            if cached is not None:
                return cached
            registry = self.registry
            sheet = registry.sheet_by_name(name)
            if sheet is None:
                names = [entry.name for entry in registry.sheets]
                return Refusal(
                    code="sheet_not_found",
                    message=f"No sheet named {name!r} exists in this workbook.",
                    operation="read_sheet",
                    target=name,
                    details={"requested_name": name, "available_names": names},
                    recovery_options=(
                        {"action": "choose_existing_sheet", "available_names": names},
                    ),
                )
            result = parse_worksheet(archive, sheet, self._shared_strings, self._styles)
            self._sheet_cache[name] = result
            return result
        except BaseException as exc:
            # A parser refusal is a normal result; unexpected failures leave
            # no partially usable session. Only our deliberate argument error
            # leaves an open session usable; interruption during validation
            # and unexpected TypeErrors in the reader still close.
            if not invalid_name or not isinstance(exc, TypeError):
                self.close()
            raise

    def close(self) -> None:
        """Release owned data; records already returned remain usable."""
        archive = self._archive
        if archive is None:
            return
        try:
            archive.close()
        finally:
            self._archive = None
            self._registry = None
            self._shared_strings = ()
            self._styles = ()
            self._sheet_cache.clear()

    def __enter__(self) -> Workbook:
        if self.closed:
            raise ClosedWorkbookError("The workbook is closed.")
        return self

    def __exit__(
        self,
        _exc_type: type[BaseException] | None,
        _exc: BaseException | None,
        _traceback: TracebackType | None,
    ) -> None:
        self.close()
