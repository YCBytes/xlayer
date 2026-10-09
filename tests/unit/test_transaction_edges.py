"""Additional boundary probes from the contract closure checklist."""

from __future__ import annotations

import zipfile
from dataclasses import replace
from pathlib import Path
from typing import cast

import pytest

from tests.unit.test_scalar_patch import plan, style_xml
from tests.unit.test_transaction import host, prepared, workbook
from tests.unit.test_transaction_safety import MAIN, REL, rewrite
from xlayer._approval import Approval, VerifiedApproval
from xlayer._dependencies import ImpactLimits
from xlayer._edits import SetValue
from xlayer._errors import Refusal
from xlayer._preview import Preview
from xlayer._workbook import Workbook


@pytest.mark.parametrize("identifier", [18, 22, 27, 999])
def test_time_datetime_locale_and_unknown_target_formats_refuse(
    tmp_path: Path, identifier: int
) -> None:
    result = plan(
        tmp_path, '<c r="A1"><v>1</v></c>', 2, styles=style_xml(f'<xf numFmtId="{identifier}"/>')
    )
    assert isinstance(result, Refusal) and result.code == "unsupported_temporal_target"


@pytest.mark.parametrize(
    "payload", ["<r><t>rich</t></r>", '<t>plain</t><rPh sb="0" eb="1"><t>guide</t></rPh>']
)
def test_shared_string_rich_or_phonetic_origin_refuses(tmp_path: Path, payload: str) -> None:
    path = workbook(tmp_path, '<c r="A1" t="s"><v>0</v></c>')
    with zipfile.ZipFile(path) as archive:
        rels = archive.read("xl/_rels/workbook.xml.rels")
    rewrite(
        path,
        {
            "xl/sharedStrings.xml": f'<sst xmlns="{MAIN}"><si>{payload}</si></sst>'.encode(),
            "xl/_rels/workbook.xml.rels": rels.replace(
                b"</Relationships>",
                (
                    f'<Relationship Id="ss" Type="{REL}sharedStrings" '
                    'Target="sharedStrings.xml"/></Relationships>'
                ).encode(),
            ),
        },
    )
    with Workbook.open(path) as book:
        proposal, preview, approval = prepared(
            book, tmp_path / "out.xlsx", edits=[SetValue("Inputs", "A1", "new")]
        )
        assert preview.evidence["blocked"]
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "rich_text_cell"


def test_utf16_session_reads_but_preview_refuses_edit_encoding(tmp_path: Path) -> None:
    path = workbook(tmp_path)
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("xl/worksheets/1.xml").decode()
    rewrite(path, {"xl/worksheets/1.xml": xml.encode("utf-16")})
    with Workbook.open(path) as book:
        sheet = book.read_sheet("Inputs")
        assert not isinstance(sheet, Refusal) and sheet.cells["A1"].raw == "100"
        proposal = book.propose([SetValue("Inputs", "A1", 120)], output_path=tmp_path / "out.xlsx")
        assert not isinstance(proposal, Refusal)
        preview = proposal.preview()
        assert isinstance(preview, Preview) and preview.evidence["blocked"]
        assert preview.to_dict()["findings"][0]["refusal"]["code"] == "unsupported_write_encoding"  # type: ignore[index]


def test_unknown_dependency_evidence_never_becomes_approvable_zero(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import xlayer._batch_impact as module
    from xlayer._dependencies import _build_index

    def inconsistent(*args: object) -> object:
        index = _build_index(*args)  # type: ignore[arg-type]
        return replace(index, source_fingerprint="sha256:" + "0" * 64)

    monkeypatch.setattr(module, "_build_index", inconsistent)
    with Workbook.open(workbook(tmp_path)) as book:
        proposal, preview, approval = prepared(book, tmp_path / "out.xlsx")
        assert preview.evidence["blocked"]
        dependencies = preview.to_dict()["dependencies"]
        assert dependencies["roots"][0]["impact"]["known_direct_count"] is None  # type: ignore[index]
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "dependency_analysis_untrustworthy"


@pytest.mark.parametrize("mode", ["deny", "mismatch"])
def test_refused_apply_still_carries_the_expected_identity_chain(tmp_path: Path, mode: str) -> None:
    with Workbook.open(workbook(tmp_path)) as book:
        proposal, preview, approval = prepared(book, tmp_path / "out.xlsx")
        if mode == "mismatch":
            approval = replace(approval, preview_digest="sha256:" + "0" * 64)
        result = proposal.apply(
            approval=approval, verify_approval=lambda claim: VerifiedApproval(claim, False, ())
        )
        assert isinstance(result, Refusal)
        assert result.details["source_fingerprint"] == book.source_fingerprint
        assert result.details["proposal_digest"] == proposal.proposal_digest
        assert result.details["preview_digest"] == preview.preview_digest


@pytest.mark.parametrize("fault", ["duplicate", "destination"])
def test_early_proposal_refusal_carries_only_the_known_source_identity(
    tmp_path: Path, fault: str
) -> None:
    with Workbook.open(workbook(tmp_path)) as book:
        edits = [SetValue("Inputs", "A1", 120)]
        if fault == "duplicate":
            edits.append(SetValue("Inputs", "A1", 121))
        result = book.propose(edits, output_path=tmp_path / "missing" / "out.xlsx")
        assert isinstance(result, Refusal)
        assert result.details["source_fingerprint"] == book.source_fingerprint
        assert "proposal_digest" not in result.details and "preview_digest" not in result.details


def test_caller_mutation_cannot_change_prepared_intent_or_metadata(tmp_path: Path) -> None:
    values = [SetValue("Inputs", "A1", 120)]
    metadata: dict[str, object] = {"nested": ["original"]}
    with Workbook.open(workbook(tmp_path)) as book:
        proposal = book.propose(values, output_path=tmp_path / "out.xlsx", annotations=metadata)
        assert not isinstance(proposal, Refusal)
        identity = proposal.proposal_digest
        values[0] = SetValue("Inputs", "A1", 999)
        metadata["nested"] = []
        assert proposal.edits[0].value == 120 and proposal.proposal_digest == identity
        assert proposal.annotations["nested"] == ("original",)


def test_missing_parent_is_not_created(tmp_path: Path) -> None:
    with Workbook.open(workbook(tmp_path)) as book:
        result = book.propose(
            [SetValue("Inputs", "A1", 2)], output_path=tmp_path / "missing" / "out.xlsx"
        )
        assert isinstance(result, Refusal) and result.code == "output_unavailable"
    assert not (tmp_path / "missing").exists()


def test_retargeted_parent_alias_invalidates_destination(tmp_path: Path) -> None:
    first, second, alias = tmp_path / "first", tmp_path / "second", tmp_path / "alias"
    first.mkdir()
    second.mkdir()
    try:
        alias.symlink_to(first, target_is_directory=True)
    except OSError:
        pytest.skip("symlinks unavailable")
    with Workbook.open(workbook(tmp_path)) as book:
        proposal, _, approval = prepared(book, alias / "out.xlsx")
        alias.unlink()
        alias.symlink_to(second, target_is_directory=True)
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "stale_output"
    assert not (first / "out.xlsx").exists() and not (second / "out.xlsx").exists()


def test_supported_empty_zip_directory_and_comment_preserved(tmp_path: Path) -> None:
    path = workbook(tmp_path)
    with zipfile.ZipFile(path, "a") as archive:
        archive.comment = b"opaque archive comment"
        archive.writestr("empty-directory/", b"")
    with Workbook.open(path) as book:
        proposal, _, approval = prepared(book, tmp_path / "out.xlsx")
        result = proposal.apply(approval=approval, verify_approval=host)
        assert not isinstance(result, Refusal)
    with zipfile.ZipFile(tmp_path / "out.xlsx") as archive:
        assert archive.comment == b"opaque archive comment"
        assert archive.read("empty-directory/") == b""


def test_directory_payload_is_rejected_before_decompression(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = workbook(tmp_path)
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr("not-a-payload-part/", b"unmodeled payload")
    with Workbook.open(path) as book:
        proposal, _, approval = prepared(book, tmp_path / "out.xlsx")
        validated_archive = book._archive
        assert validated_archive is not None
        original = validated_archive._zip.read

        def guarded(name: str | zipfile.ZipInfo, *args: object, **kwargs: object) -> bytes:
            if isinstance(name, zipfile.ZipInfo) and name.is_dir():
                raise RuntimeError("directory payload was decompressed")
            return original(name, *args, **kwargs)  # type: ignore[arg-type]

        monkeypatch.setattr(validated_archive._zip, "read", guarded)
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "unsupported_zip_metadata"
    assert not (tmp_path / "out.xlsx").exists()


@pytest.mark.parametrize("fault", ["crc", "compression"])
def test_empty_directory_integrity_failures_are_structured(tmp_path: Path, fault: str) -> None:
    path = workbook(tmp_path)
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr("empty-directory/", b"")
    with Workbook.open(path) as book:
        proposal, _, approval = prepared(book, tmp_path / "out.xlsx")
        source_archive = book._archive
        assert source_archive is not None
        entry = source_archive._zip.getinfo("empty-directory/")
        if fault == "crc":
            entry.CRC = 1
        else:
            entry.compress_type = 99
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal)
        # Framing admission is now in preview: the injected mutation changes
        # that evidence, so the older approval is rejected BEFORE writing.
        assert result.code == "preview_digest_mismatch"
        _, fresh, _ = prepared(book, tmp_path / "out.xlsx")
        assert any(
            f["refusal"]["code"]
            == ("malformed_archive" if fault == "crc" else "unsupported_zip_metadata")
            for f in cast("list[dict[str, dict[str, object]]]", fresh.to_dict()["findings"])
        )
        assert not book.closed
    assert not (tmp_path / "out.xlsx").exists() and not list(tmp_path.glob(".xlayer-*"))


def test_unavailable_exclusive_link_primitive_has_no_copy_fallback(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import xlayer._transaction as module

    def unsupported(*_args: object, **_kwargs: object) -> None:
        raise OSError("filesystem does not support hard links")

    monkeypatch.setattr(module.os, "link", unsupported)
    with Workbook.open(workbook(tmp_path)) as book:
        proposal, _, approval = prepared(book, tmp_path / "out.xlsx")
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "publication_failed"
    assert not (tmp_path / "out.xlsx").exists() and not list(tmp_path.glob(".xlayer-*"))


def test_exact_global_budget_unstarted_root_blocks_partial_override(tmp_path: Path) -> None:
    with Workbook.open(
        workbook(tmp_path, '<c r="A1"><v>100</v></c><c r="B1"><v>2</v></c>')
    ) as book:
        proposal = book.propose(
            [SetValue("Inputs", "A1", 120), SetValue("Inputs", "B1", 3)],
            output_path=tmp_path / "out.xlsx",
            limits=ImpactLimits(max_visited_nodes=1),
        )
        assert not isinstance(proposal, Refusal)
        preview = proposal.preview()
        assert isinstance(preview, Preview) and preview.evidence["blocked"]
        approval = Approval(
            "1.0",
            "id",
            "host",
            "human",
            book.source_fingerprint,
            proposal.proposal_digest,
            preview.preview_digest,
            "now",
        )
        result = proposal.apply(approval=approval, verify_approval=host)
        assert isinstance(result, Refusal) and result.code == "transaction_limit_exceeded"
