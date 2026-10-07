"""Captured local filesystem bindings; no in-place writes or concurrent-writer CAS."""

from __future__ import annotations

import hashlib
import os
import stat
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

from xlayer._approval import refusal
from xlayer._errors import Refusal

if TYPE_CHECKING:
    from xlayer._workbook import Workbook


def hash_file(path: Path, maximum: int, code: str) -> str | Refusal:
    try:
        if not stat.S_ISREG(path.stat().st_mode):
            return refusal(code, "bound file is not a regular file")
        with path.open("rb") as handle:
            data = handle.read(maximum + 1)
        if len(data) > maximum:
            return refusal(code, "bound file exceeds loaded archive limits")
        return "sha256:" + hashlib.sha256(data).hexdigest()
    except OSError as exc:
        return refusal(code, "bound file cannot be read", error=str(exc))


def check_source(book: Workbook) -> Refusal | None:
    archive = book._archive
    if archive is None:
        raise RuntimeError("source check requires an open session")
    for path in {book._source_binding, book._source_resolved}:
        identity = hash_file(path, archive._limits.max_file_bytes, "stale_source")
        if isinstance(identity, Refusal):
            return identity
        if identity != book.source_fingerprint:
            return refusal("stale_source", "source path no longer contains the loaded snapshot")
    return None


@dataclass(frozen=True)
class OutputSpec:
    supplied_parent: Path
    parent: Path
    parent_identity: tuple[int, int]
    filename: str
    overwrite: bool

    @property
    def path(self) -> Path:
        return self.parent / self.filename

    def to_dict(self) -> dict[str, object]:
        return {
            "output_path": str(self.path),
            "supplied_parent": str(self.supplied_parent),
            "parent_identity": list(self.parent_identity),
            "overwrite": self.overwrite,
            "publication_mode": "replace" if self.overwrite else "exclusive_link",
            "receipt_persistence": "returned_host_saves",
            "posix_file_mode": "0600",
        }

    def inspect(self, book: Workbook) -> str | Refusal | None:
        try:
            resolved = self.supplied_parent.resolve(strict=True)
            info = resolved.stat()
            if resolved != self.parent or (info.st_dev, info.st_ino) != self.parent_identity:
                return refusal("stale_output", "destination directory binding changed")
            if not stat.S_ISDIR(info.st_mode):
                return refusal("output_unavailable", "destination parent is not a directory")
            try:
                incumbent = self.path.lstat()
            except FileNotFoundError:
                return None
            if not stat.S_ISREG(incumbent.st_mode):
                return refusal("unsafe_output_path", "output is a symlink or non-regular file")
            identity = incumbent.st_dev, incumbent.st_ino
            if identity == book._source_file_identity:
                return refusal("unsafe_output_path", "output aliases the originally loaded source")
            for source in {book._source_binding, book._source_resolved}:
                try:
                    source_info = source.stat()
                except FileNotFoundError:
                    continue
                if identity == (source_info.st_dev, source_info.st_ino):
                    return refusal("unsafe_output_path", "output aliases a source location")
            if not self.overwrite:
                return refusal("output_exists", "overwrite was not authorized")
            archive = book._archive
            if archive is None:
                raise RuntimeError("destination check requires an open session")
            return hash_file(self.path, archive._limits.max_file_bytes, "output_unavailable")
        except (OSError, RuntimeError) as exc:
            # Path.resolve reports symlink loops as RuntimeError on older Python.
            if isinstance(exc, RuntimeError) and "symlink" not in str(exc).lower():
                raise
            return refusal("output_unavailable", "destination cannot be inspected", error=str(exc))


def bind_output(
    book: Workbook, path: str | os.PathLike[str], overwrite: bool
) -> OutputSpec | Refusal:
    if type(overwrite) is not bool:
        raise TypeError("overwrite must be an exact bool")
    supplied = Path(path).absolute()
    if supplied.suffix.lower() != ".xlsx" or supplied.name in {"", ".xlsx"}:
        return refusal("unsafe_output_path", "a separate .xlsx output filename is required")
    if supplied in {book._source_binding, book._source_resolved}:
        return refusal("unsafe_output_path", "in-place mutation is not supported")
    try:
        parent = supplied.parent.resolve(strict=True)
        info = parent.stat()
        if not stat.S_ISDIR(info.st_mode) or not info.st_ino or not os.access(parent, os.W_OK):
            return refusal("output_unavailable", "parent lacks usable identity or write access")
        output = OutputSpec(
            supplied.parent, parent, (info.st_dev, info.st_ino), supplied.name, overwrite
        )
        inspected = output.inspect(book)
        return inspected if isinstance(inspected, Refusal) else output
    except (OSError, RuntimeError) as exc:
        return refusal("output_unavailable", "output parent cannot be bound", error=str(exc))
