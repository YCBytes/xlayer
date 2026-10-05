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

from xlayer._dependencies import (
    _QUERY,
    DEFAULT_IMPACT_LIMITS,
    CellRef,
    DependencyImpact,
    ImpactLimits,
    _canonical_address,
    analyse_dependencies,
)
from xlayer._errors import ClosedWorkbookError, Refusal, WorkbookOpenError
from xlayer._ooxml.archive import DEFAULT_LIMITS, ArchiveLimits, WorkbookArchive
from xlayer._ooxml.formula import _endpoint
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

    def dependency_impact(
        self, sheet_name: str, address: str, *, limits: ImpactLimits = DEFAULT_IMPACT_LIMITS
    ) -> DependencyImpact | Refusal:
        """Analyze bounded potential references, not recalculation or edit safety.

        Rebuild the index from this snapshot on each call. Ordinary worksheet
        refusals become coverage issues; unexpected failures end the session.
        Returned evidence is detached and remains usable after close.
        """
        argument_error: BaseException | None = None
        try:
            if self._archive is None:
                raise ClosedWorkbookError("The workbook is closed.")
            valid_types = (
                isinstance(sheet_name, str),
                isinstance(address, str),
                isinstance(limits, ImpactLimits),
            )
            if not all(valid_types):
                argument_error = TypeError(
                    "sheet_name/address must be strings; limits must be ImpactLimits"
                )
                raise argument_error
            valid_coordinate = (
                _QUERY.fullmatch(address) is not None and _endpoint(address) is not None
            )
            if not valid_coordinate:
                argument_error = ValueError(
                    "address must be an unadorned in-grid ASCII A1 coordinate"
                )
                raise argument_error
            canonical = _canonical_address(address)
            registry = self.registry
            sheet = registry.sheet_by_name(sheet_name)
            if sheet is None:
                names = [entry.name for entry in registry.sheets]
                return Refusal(
                    code="sheet_not_found",
                    message=f"No sheet named {sheet_name!r} exists in this workbook.",
                    operation="dependency_impact",
                    target=sheet_name,
                    details={"requested_name": sheet_name, "available_names": names},
                    recovery_options=(
                        {"action": "choose_existing_sheet", "available_names": names},
                    ),
                )
            if sheet.kind != "worksheet":
                names = [entry.name for entry in registry.sheets if entry.kind == "worksheet"]
                return Refusal(
                    code="unsupported_sheet_kind",
                    message=f"Dependency analysis does not read {sheet.kind} tabs.",
                    operation="dependency_impact",
                    target=sheet_name,
                    details={"requested_name": sheet_name, "kind": sheet.kind},
                    recovery_options=({"action": "choose_worksheet", "available_names": names},),
                )
            return analyse_dependencies(
                registry,
                self.read_sheet,
                CellRef(sheet.name, canonical),
                self.source_fingerprint,
                limits,
            )
        except BaseException as exc:
            # Exempt only the exact error we deliberately constructed. A reader,
            # lookup or validation helper's TypeError/ValueError still closes.
            if exc is not argument_error:
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
