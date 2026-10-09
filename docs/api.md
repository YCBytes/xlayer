# Alpha API

This is the experimental `0.1.0a1` interface. Supported imports are from the
package root only:

```python
from xlayer import (
    Workbook,
    SetValue,
    Proposal,
    Inspection,
    Preview,
    Receipt,
    Approval,
    VerifiedApproval,
    Refusal,
    XlayerError,
    WorkbookOpenError,
    ClosedWorkbookError,
    ArchiveLimits,
    ImpactLimits,
    ReadLimits,
    __version__,
)
```

Underscore modules/parser records are implementation details. Construct edits,
approvals and limits directly. Obtain sessions/proposals through `open()`/
`propose()` and results through operations. Direct ownership/result construction
and loading proposals from JSON are unsupported. This is not a Python sandbox.

## Sessions and failures

```text
Workbook.open(path, *, limits=ArchiveLimits()) -> Workbook
book.list_sheets(*, limits=ReadLimits()) -> Inspection | Refusal
book.read_cells(sheet_name, addresses, *, limits=ReadLimits()) -> Inspection | Refusal
book.dependency_impact(
    sheet_name, address, *, limits=ImpactLimits(), read_limits=ReadLimits()
) -> Inspection | Refusal
book.propose(
    edits, *, output_path, overwrite=False, limits=ImpactLimits(), annotations=None
) -> Proposal | Refusal
```

Paths accept `str` or `os.PathLike[str]`. Use a context manager or `book.close()`;
close is idempotent. `source_path` is a `pathlib.Path` retaining caller spelling.
`source_fingerprint` identifies loaded bytes as `sha256:<64 lowercase hex>`, not
a later read of that path. `closed` is a bool.

Opening validates shared infrastructure, not all worksheets/edit safety.
Operational opening failure raises `WorkbookOpenError`; inspect `.refusal`.
Other operational failures return `Refusal`. Invalid Python arguments raise
`TypeError`/`ValueError`; closed use raises `ClosedWorkbookError` before cached
work. Opening/closed-use exceptions derive from `XlayerError`; ordinary argument
errors retain their Python types.

Unexpected parsing/projection/callback errors release the session and propagate
unchanged, not as a guessed refusal. Argument errors and ordinary refusals leave
a healthy session open. After close, only identity properties and `closed` remain
available; detached results still work, proposals cannot preview/apply.
Concurrent/reentrant session use is unsupported.

Reads use the loaded snapshot after source replacement/deletion. Apply rechecks
the bound source. A session never refreshes itself to its output; open it separately.

Sheet lookup is exact. Read/impact addresses are unadorned ASCII A1, either case,
normalized uppercase within A1–XFD1048576. No ranges, `$`, whitespace, sheet
prefixes or Unicode digits. `read_cells` takes a non-string sequence, captures
it before parsing, rejects empty/duplicate canonical addresses, and keeps request
order. Malformed worksheets refuse the whole read, never fabricate missing cells.

## Inspection

`.evidence` is recursively frozen/detached. `to_dict()` returns an independent
dictionary/list tree; `canonical_json()` returns bytes with sorted keys, compact
separators, ASCII escaping and finite numbers. No timestamp/logging/approval digest.
The envelope has exactly:

```text
schema_version: "1.0"
inspection_contract_version: "1.0"
kind: "sheets" | "cells" | "dependency_impact"
source_fingerprint: "sha256:<64 lowercase hex>"
data: kind-specific object
```

### Sheet inventory

`data` is `{sheets, active_tab, date1904, defined_name_count}`. Each tab has
`{name, state, kind, tab_index}`, in tab order. States: `visible`, `hidden`,
`veryHidden`; kinds: `worksheet`, `chartsheet`, `dialogsheet`. Chart/dialog cells
are not read. `active_tab` is a tab index, not a sheet ID. Saved `date1904` and
defined-name count are facts, not conversion or proof that names resolve.

### Cells

`data` is `{sheet, requested_addresses, cells}`, with canonical addresses and
request-ordered observations. Every observation has these keys, including nulls:

```text
address: str
presence: "present" | "missing"
stored_type: str | null
value_kind: "number" | "boolean" | "string" | "error" | "iso_date" | "blank" | null
raw: str | null
stored_value: float | bool | str | null
value_source: "stored_literal" | "saved_formula_cache" | "not_present"
formula_result_status: "not_verified" | null
cell_format: {style_index, format_id, format_code, temporal_kind, basis} | null
formula: {text, kind, shared_index, ref, calculate_on_next_recalc, shared_master} | null
merge: {ref, root, role} | null
```

`raw` preserves `<v>` text, not rich-text XML. Numeric floats are convenience
decodes, not digit-preserving evidence. Date/time serials stay numeric; `t="d"`
stays validated saved ISO text. Absent `t` means stored type `n`. Text uses the
shared Unicode/Excel-escape decoder; phonetic guides are not the cell value.

Missing element, existing blank, empty text, formula without `<v>`, formula with
`<v/>`, and saved zero stay distinct. Every formula has `saved_formula_cache` and
`not_verified`, including blank/error/string caches. Nothing is recalculated.

`cell_format.basis` is `cell_xf_number_format`; `temporal_kind` is `date`, `time`,
`date-time`, `non_temporal` or `unknown`. This is not effective inherited/conditional
display style or edit permission. Text carrying a date format stays text.

Formula kinds: `normal`, `shared`, `array`, `data_table`. A shared follower's
own `text` is null; its `shared_master` contains validated `{address, ref, text}`.
Other formulas have null `shared_master`. Expressions are not rewritten/expanded.
Array/data-table ownership does not invent cells. Merge context is geometric:
`{ref, root, role: "root" | "child"}`. A child without `<c>` stays missing with
null stored value, not a copy of its root.

### Dependency impact

`data` is `{impact_summary, limits, work}`. Summary `1.0`, dependency contract
`1.1`, describes one started root in `worksheet_cell_formulas`: retained potential
references, not calculated effects. Presentation does not rescan the workbook.

Root fields include `root`, `state`, `analysis_status`, `known_direct_count`,
`known_transitive_count`, `exact_total_count`, `blocked_budget`, `root_cursor`,
`direct_dependents_sample`, `transitive_dependents_sample`, truncation flags,
`inventory_frontier`, `traversal_frontier`, and bounded `path_examples`.
Complete has an exact total; partial has lower-bound known counts and null exact
total; unknown has null counts, not zero.

Totals include known union/transitive counts and affected-sheet counts with
bounded affected-sheet samples and grouped diagnostic examples. Sample limits:
10 direct, 10 transitive, 10 affected sheets, 20 issue groups, 2 examples per
group, one path per kind of at most 12 hops. Truncation cannot improve coverage.
Paths are reference proofs; frontiers are facts, not a resume API. This is not
a raw graph/full-formula export or full cycle list.

`limits` reports actual ceilings. `work` contains actual `sheets_attempted`,
`trusted_sheets`, `scanned_cells`, `defined_names_admitted`,
`formula_cells_attempted`, `max_formula_chars_observed`,
`max_formula_nesting_observed`, `references_admitted`, `membership_checks`,
`visited_nodes` and `evidence_edges_admitted`.

## Limits

`ReadLimits(max_cells=128, max_response_bytes=1_048_576)` accepts exact positive
integers up to these defaults. Cell ceiling: explicit reads; byte ceiling: all
complete inspections, not model tokens/time. Over-limit returns
`inspection_limit_exceeded` with budget/limit/count or observed-byte evidence,
never cropped facts. A fixed 1 MiB per-diagnostic identity guard runs before
summary hashing; its budget is `max_summary_issue_bytes`, not response size.

Existing archive/impact records can be constructed with higher values for private
use, but public operations accept defaults or lower positive ceilings only.
Wrong limits object raises `TypeError`; invalid fields/raised public ceilings
raise `ValueError`. Default tables are in [Support](support.md).

## Propose, approve, apply

Frozen `SetValue(sheet, cell, value)` preserves exact type, canonical address and
ordered intent. `to_dict()` returns
`{sheet, cell, value: {kind: "int"|"float"|"bool"|"str", value}}`.
Date/datetime and subclasses are rejected; construction never authorizes a write.

`propose` captures 1–100 exact public edits, rejects duplicate targets, and requires
a separate `.xlsx` in an existing directory. `overwrite` is an exact bool, false
by default. The batch targets one worksheet. `annotations` is detached JSON audit
metadata, not authority/business truth. `Proposal.edits` is a public edit tuple;
`source_fingerprint` and `proposal_digest` identify retained source/intent.

`preview() -> Preview | Refusal` writes nothing and may be cached. Preview can be
blocked: inspect `evidence["blocked"]`, findings and `required_capabilities`.
If every validated payload is unchanged, the result is `no_changes`.

Preview schema `1.0` retains versions, source/proposal identity, output binding/
incumbent, requested/effective edits, findings, full bounded dependencies plus
summary, limits, advisory measurements, grants, expected parts, next action,
recalculation and visual status. `preview_digest` binds exact evidence;
`to_dict()`/`canonical_json()` add the digest and detached `audit_metadata`.
Caller annotations are outside the approval identity.

Required `Approval` fields:

```text
schema_version="1.0", approval_id, issuer, actor_kind,
source_fingerprint, proposal_digest, preview_digest, issued_at
```

Actor kinds: `human`, `service`, `automation`, `model`. Strings are identity
claims, not authority. The required host verifier authenticates the exact Approval
and returns `VerifiedApproval(approval, authorized, capabilities)`. Capabilities
come only from that trusted decision. Ordinary writes require `apply_set_value`;
populated text/partial coverage additionally requires `change_populated_text`/
`accept_partial_dependency_coverage` and independent human/service authority.
Unknown/not-evaluated evidence cannot be overridden.

`proposal.apply(*, approval, verify_approval) -> Receipt | Refusal` has no default
verifier/direct-write shortcut. Missing/denied/mismatched approval writes nothing.
After staging, the engine checks fresh host authority, source/output/parent
bindings and preview identity again. Invalid members block the entire batch.

Receipt retains preview evidence and adds `preview_digest_reference`, output
fingerprint, approval/grants, verification, actual changed parts and publication
mode. `receipt_digest`, `to_dict()` and `canonical_json()` work like preview
counterparts. Persistence is application-owned, not an atomic workbook/receipt
bundle. Replay refuses.

Known dependents retain `required_not_run`; absent proof may retain `unknown`.
Visual status is `not_evaluated`. Structural success is not verified business
correctness, recalculation or rendered fidelity. See [Support](support.md) for
interruption and delivery boundaries.
