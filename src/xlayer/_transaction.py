"""Prepare and verify before the single publication point; return audit evidence."""

from __future__ import annotations

import copy
import os as os
import tempfile
import zipfile
import zlib
from collections.abc import Callable
from pathlib import Path
from typing import TYPE_CHECKING, BinaryIO, cast

from xlayer._approval import Approval, VerifiedApproval, check_approval, refusal
from xlayer._canonical import digest
from xlayer._errors import Refusal
from xlayer._ooxml.archive import WorkbookArchive
from xlayer._ooxml.patch import patch_worksheet
from xlayer._output import check_source
from xlayer._preview import Preview
from xlayer._receipt import Receipt
from xlayer._verification import verify_output

if TYPE_CHECKING:
    from xlayer._proposal import Proposal


def write_archive(
    archive: WorkbookArchive, handle: BinaryIO, part: str, patched: bytes
) -> Refusal | None:
    infos = archive._zip.infolist()
    # These flags/extras would be rewritten or carry offsets/checksums whose
    # preservation cannot be claimed. Reject instead of silently stripping.
    if any(info.flag_bits & ~0x800 or info.extra or info.volume for info in infos):
        return refusal("unsupported_zip_metadata", "ZIP metadata cannot be safely preserved")
    payloads: list[tuple[zipfile.ZipInfo, bytes]] = []
    for info in infos:
        if info.is_dir():
            if info.file_size or info.compress_type not in {
                zipfile.ZIP_STORED,
                zipfile.ZIP_DEFLATED,
            }:
                return refusal(
                    "unsupported_zip_metadata", "directory entry carries unsupported payload"
                )
            try:
                data = archive._zip.read(info)
            except (zipfile.BadZipFile, zlib.error) as exc:
                return refusal(
                    "malformed_archive", "directory integrity check failed", error=str(exc)
                )
            if data:
                return refusal("unsupported_zip_metadata", "directory entry carries opaque payload")
        else:
            content = archive.read_part(info.filename)
            if isinstance(content, Refusal):
                return content
            data = patched if info.filename == part else content
        payloads.append((info, data))
    with zipfile.ZipFile(handle, "w", allowZip64=False) as output:
        output.comment = archive._zip.comment
        for info, data in payloads:
            output.writestr(copy.copy(info), data)
    return None


def _identities(proposal: Proposal, preview: Preview) -> dict[str, object]:
    return {
        "source_fingerprint": proposal._book.source_fingerprint,
        "proposal_digest": proposal.proposal_digest,
        "preview_digest": preview.preview_digest,
    }


def _output_recheck(proposal: Proposal, preview: Preview) -> Refusal | None:
    current = proposal.output.inspect(proposal._book)
    if isinstance(current, Refusal) or current != preview.evidence["output_incumbent"]:
        return refusal("stale_output", "output binding or incumbent changed after preview")
    return None


def apply_proposal(
    proposal: Proposal, approval: Approval | None, verifier: Callable[[Approval], VerifiedApproval]
) -> Receipt | Refusal:
    if proposal._receipt is not None:
        return refusal(
            "proposal_already_applied",
            "this proposal already committed",
            receipt_digest=proposal._receipt.receipt_digest,
        )
    preview = proposal.preview()
    if isinstance(preview, Refusal):
        return preview
    if preview.evidence["blocked"]:
        findings = cast("tuple[dict[str, object], ...]", preview.evidence["findings"])
        first = cast("dict[str, object]", findings[0]["refusal"])
        return refusal(
            str(first["code"]),
            "validation blocks the whole batch",
            findings=preview.evidence["findings"],
            **_identities(proposal, preview),
        )
    if approval is None:
        return refusal(
            "approval_required", "trusted approval is required", **_identities(proposal, preview)
        )
    if digest(proposal.identity()) != proposal.proposal_digest:
        return refusal("preview_digest_mismatch", "proposal identity changed")
    required = cast("tuple[str, ...]", preview.evidence["required_capabilities"])
    decision = verifier(approval)
    refused = check_approval(
        approval,
        decision,
        proposal._book.source_fingerprint,
        proposal.proposal_digest,
        preview.preview_digest,
        required,
    )
    if refused is not None:
        return refused
    source_check = check_source(proposal._book)
    if source_check is not None:
        return source_check
    output_check = _output_recheck(proposal, preview)
    if output_check is not None:
        return output_check
    refreshed = proposal._build_preview()
    if isinstance(refreshed, Refusal):
        return refreshed
    current_preview, plans = refreshed
    if current_preview.preview_digest != preview.preview_digest:
        return refusal("preview_digest_mismatch", "validation or policy changed after approval")
    archive = proposal._book._archive
    if archive is None:
        raise RuntimeError("transaction lost its source session")
    part = plans[0].part
    source_bytes = archive.read_part(part)
    if isinstance(source_bytes, Refusal):
        return source_bytes
    patched = patch_worksheet(source_bytes, plans)
    if isinstance(patched, Refusal):
        return patched
    temp: Path | None = None
    committed = False
    prepared_receipt: Receipt | None = None
    filesystem_operation = True
    try:
        descriptor, name = tempfile.mkstemp(
            prefix=".xlayer-", suffix=".xlsx", dir=proposal.output.parent
        )
        temp = Path(name)
        try:
            handle = os.fdopen(descriptor, "w+b")
        except BaseException:
            os.close(descriptor)
            raise
        with handle:
            refused = write_archive(archive, handle, part, patched)
            if refused is not None:
                return refused
            handle.flush()
            os.fsync(handle.fileno())
        filesystem_operation = False
        report = verify_output(proposal._book, temp, plans, part)
        if isinstance(report, Refusal):
            return report
        evidence = {
            **dict(preview.evidence),
            "preview_digest_reference": preview.preview_digest,
            "output_fingerprint": report.output_fingerprint,
            "approval": approval.to_dict(),
            "verified_capabilities": list(decision.capabilities),
            "verification": report.to_dict(),
            "actual_changed_parts": list(report.changed_parts),
            "publication_mode": "replace" if proposal.output.overwrite else "exclusive_link",
        }
        prepared_receipt = Receipt(evidence, {"caller_annotations": proposal.annotations})
        prepared_receipt.canonical_json()
        fresh = verifier(approval)
        refused = check_approval(
            approval,
            fresh,
            proposal._book.source_fingerprint,
            proposal.proposal_digest,
            preview.preview_digest,
            required,
        )
        if refused is not None:
            return refused
        if fresh.capabilities != decision.capabilities:
            return refusal("approval_rejected", "fresh host grants differ from prepared receipt")
        # A host callback may take time or affect filesystem state. Perform
        # file identity checks after it, immediately before publication.
        source_check = check_source(proposal._book)
        if source_check is not None:
            return source_check
        output_check = _output_recheck(proposal, preview)
        if output_check is not None:
            return output_check
        # The ONLY publication point. No fallback that could clobber a rival.
        filesystem_operation = True
        if proposal.output.overwrite:
            os.replace(temp, proposal.output.path)  # noqa: PTH105 - explicit atomic primitive
        else:
            os.link(temp, proposal.output.path)
        committed = True
        object.__setattr__(proposal, "_receipt", prepared_receipt)
        if not proposal.output.overwrite:
            try:
                os.unlink(temp)  # noqa: PTH108 - narrow OS failure-injection boundary
            except OSError:
                # Audit-only annotation is best effort. Evidence was already
                # constructed and the workbook is committed; never roll back.
                try:
                    annotated = Receipt(
                        prepared_receipt.evidence,
                        {
                            **dict(prepared_receipt.audit_metadata),
                            "cleanup_warning": "owned temporary link could not be removed",
                        },
                    )
                    annotated.canonical_json()
                    prepared_receipt = annotated
                except Exception:  # noqa: S110 - no automatic export of private audit data
                    pass
        return prepared_receipt
    except OSError as exc:
        if not filesystem_operation:
            raise
        if committed:
            if prepared_receipt is None:
                raise
            return prepared_receipt
        return refusal(
            "publication_failed",
            "prepublication filesystem operation failed",
            error=str(exc),
            **_identities(proposal, preview),
        )
    finally:
        if temp is not None and not committed:
            try:  # noqa: SIM105 - original failure must remain authoritative
                os.unlink(temp)  # noqa: PTH108
            except OSError:
                pass
