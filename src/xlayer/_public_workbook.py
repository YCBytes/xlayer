"""Supported operations composed over the existing managed snapshot/transaction engine."""

from __future__ import annotations

import os
from collections.abc import Mapping, Sequence
from pathlib import Path
from types import TracebackType
from typing import TypeVar

from xlayer._approval import refusal as _transaction_refusal
from xlayer._canonical import (
    CanonicalTypeError,
    CanonicalValueError,
    EvidenceTooLarge,
    frozen_mapping,
    string,
)
from xlayer._dependencies import DEFAULT_IMPACT_LIMITS, ImpactLimits, _canonical_address, _to_json
from xlayer._errors import ClosedWorkbookError, Refusal
from xlayer._inspection import (
    DEFAULT_READ_LIMITS,
    Inspection,
    ReadLimits,
    check_limits,
    limit_refusal,
    make_inspection,
    observe_cell,
    single_impact_summary,
)
from xlayer._ooxml.archive import DEFAULT_LIMITS, ArchiveLimits
from xlayer._public_edits import SetValue
from xlayer._public_proposal import Proposal
from xlayer._workbook import Workbook as _EngineWorkbook

_T = TypeVar("_T")


class Workbook:
    """A managed local .xlsx snapshot. Use open() and a context manager.

    Opening admits shared infrastructure, not every worksheet or every edit.
    Saved formula caches are never recalculated or verified by an inspection.
    """

    __slots__ = ("_book",)

    def __init__(self, book: _EngineWorkbook) -> None:
        self._book = book

    @classmethod
    def open(
        cls, path: str | os.PathLike[str], *, limits: ArchiveLimits = DEFAULT_LIMITS
    ) -> Workbook:
        check_limits(limits, DEFAULT_LIMITS, "limits")
        book = _EngineWorkbook.open(path, limits=limits)
        try:
            return cls(book)
        except BaseException:
            book.close()
            raise

    @property
    def source_path(self) -> Path:
        return self._book.source_path

    @property
    def source_fingerprint(self) -> str:
        return self._book.source_fingerprint

    @property
    def closed(self) -> bool:
        return self._book.closed

    def close(self) -> None:
        self._book.close()

    def _ensure_open(self) -> None:
        if self.closed:
            raise ClosedWorkbookError("The workbook is closed.")

    def _capture(self, values: Sequence[_T]) -> tuple[_T, ...]:
        try:
            return tuple(values)
        except BaseException:
            # A user-defined sequence may fail unexpectedly while captured.
            # This is not one of our deliberate argument-validation errors.
            self.close()
            raise

    def list_sheets(self, *, limits: ReadLimits = DEFAULT_READ_LIMITS) -> Inspection | Refusal:
        self._ensure_open()
        check_limits(limits, DEFAULT_READ_LIMITS, "limits")
        try:
            registry = self._book.registry
            return make_inspection(
                "sheets",
                self.source_fingerprint,
                {
                    "sheets": [
                        {"name": s.name, "state": s.state, "kind": s.kind, "tab_index": s.tab_index}
                        for s in registry.sheets
                    ],
                    "active_tab": registry.active_tab,
                    "date1904": registry.date1904,
                    "defined_name_count": len(registry.defined_names),
                },
                limits,
                "list_sheets",
            )
        except BaseException:
            self.close()
            raise

    def read_cells(
        self, sheet_name: str, addresses: Sequence[str], *, limits: ReadLimits = DEFAULT_READ_LIMITS
    ) -> Inspection | Refusal:
        self._ensure_open()
        check_limits(limits, DEFAULT_READ_LIMITS, "limits")
        string(sheet_name, "sheet_name", nonempty=True)
        candidate: object = addresses
        if not isinstance(candidate, Sequence) or isinstance(candidate, (str, bytes, bytearray)):
            raise TypeError("addresses must be a non-string sequence")
        captured = self._capture(addresses)
        if not captured:
            raise ValueError("addresses must not be empty")
        canonical = tuple(_canonical_address(string(a, "address", nonempty=True)) for a in captured)
        if len(set(canonical)) != len(canonical):
            raise ValueError("duplicate canonical address")
        if len(canonical) > limits.max_cells:
            return limit_refusal(
                "read_cells",
                self.source_fingerprint,
                "max_cells",
                limits.max_cells,
                requested_cells=len(canonical),
            )
        try:
            sheet = self._book.read_sheet(sheet_name)
            if isinstance(sheet, Refusal):
                return sheet
            return make_inspection(
                "cells",
                self.source_fingerprint,
                {
                    "sheet": sheet_name,
                    "requested_addresses": canonical,
                    "cells": [observe_cell(sheet, address) for address in canonical],
                },
                limits,
                "read_cells",
            )
        except BaseException:
            self.close()
            raise

    def dependency_impact(
        self,
        sheet_name: str,
        address: str,
        *,
        limits: ImpactLimits = DEFAULT_IMPACT_LIMITS,
        read_limits: ReadLimits = DEFAULT_READ_LIMITS,
    ) -> Inspection | Refusal:
        self._ensure_open()
        check_limits(limits, DEFAULT_IMPACT_LIMITS, "limits")
        check_limits(read_limits, DEFAULT_READ_LIMITS, "read_limits")
        string(sheet_name, "sheet_name", nonempty=True)
        canonical = _canonical_address(string(address, "address", nonempty=True))
        try:
            impact = self._book.dependency_impact(sheet_name, canonical, limits=limits)
            if isinstance(impact, Refusal):
                return impact
            order = {entry.name: entry.tab_index for entry in self._book.registry.sheets}
            summary = single_impact_summary(impact, order)
            if isinstance(summary, Refusal):
                return summary
            return make_inspection(
                "dependency_impact",
                self.source_fingerprint,
                {
                    "impact_summary": summary,
                    "limits": _to_json(impact.limits),
                    "work": _to_json(impact.work),
                },
                read_limits,
                "dependency_impact",
            )
        except BaseException:
            self.close()
            raise

    def propose(
        self,
        edits: Sequence[SetValue],
        *,
        output_path: str | os.PathLike[str],
        overwrite: bool = False,
        limits: ImpactLimits = DEFAULT_IMPACT_LIMITS,
        annotations: Mapping[str, object] | None = None,
    ) -> Proposal | Refusal:
        self._ensure_open()
        check_limits(limits, DEFAULT_IMPACT_LIMITS, "limits")
        candidate: object = edits
        if not isinstance(candidate, Sequence) or isinstance(candidate, (str, bytes, bytearray)):
            raise TypeError("edits must be a non-string sequence of public SetValue")
        captured = self._capture(edits)
        if not 1 <= len(captured) <= 100:
            raise ValueError("a transaction contains 1..100 edits")
        if any(type(edit) is not SetValue for edit in captured):
            raise TypeError("exact public SetValue records are required")
        if type(overwrite) is not bool:
            raise TypeError("overwrite must be an exact bool")
        path_argument: object = output_path
        if isinstance(path_argument, str):
            path_text = path_argument
        elif isinstance(path_argument, os.PathLike):
            try:
                path_text = path_argument.__fspath__()
            except BaseException:
                self.close()
                raise
            if not isinstance(path_text, str):
                raise TypeError("output_path must resolve to str, not bytes")
        else:
            raise TypeError("output_path must be str or os.PathLike[str]")
        destination = Path(path_text)
        string(str(destination), "output_path")
        if "\x00" in str(destination):
            raise ValueError("output_path must not contain NUL")
        try:
            metadata = frozen_mapping({} if annotations is None else annotations)
        except (CanonicalTypeError, CanonicalValueError):
            # Deliberate canonical-domain argument errors do not poison a session.
            # Caller-container TypeError/ValueError instead follow owned cleanup.
            raise
        except BaseException:
            self.close()
            raise
        try:
            proposed = self._book.propose(
                tuple(edit._edit for edit in captured),
                output_path=destination,
                overwrite=overwrite,
                limits=limits,
                annotations=metadata,
            )
            return proposed if isinstance(proposed, Refusal) else Proposal(proposed, captured)
        except EvidenceTooLarge:
            return _transaction_refusal(
                "transaction_limit_exceeded",
                "proposal exceeds canonical evidence ceiling",
                source_fingerprint=self.source_fingerprint,
            )
        except BaseException:
            self.close()
            raise

    def __enter__(self) -> Workbook:
        self._ensure_open()
        return self

    def __exit__(
        self,
        _exc_type: type[BaseException] | None,
        _exc: BaseException | None,
        _traceback: TracebackType | None,
    ) -> None:
        self.close()
