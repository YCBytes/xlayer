# Support and limitations

`0.1.0a1` is experimental, not a general spreadsheet editor, calculation engine,
AI agent, security sandbox or production certificate. See [API](api.md).

## Workbook and edit scope

Reads admit local transitional SpreadsheetML `.xlsx` ZIP packages subject to
archive/XML/relationship checks. Strict OOXML, macro-enabled files, templates,
encrypted/legacy containers and other document kinds refuse. Stored/deflated
entries are supported, not arbitrary compression/ZIP64. External relationships
are inert metadata, never fetched.

Reads include plain/shared/rich/inline text, booleans, finite numbers, saved
error/ISO-date values, formula metadata/caches and merges. Hidden sheets are
readable; chart/dialog tabs are inventory only. Reading does not prove editability
or effective display style.

Writes are narrower: existing supported scalar cells in one worksheet per batch.
No formula overwrite, missing-cell creation, rich-text replacement, merged-cell
writes (root or child), protected-sheet editing, table/validation ownership or
unsupported target metadata. Unsupported number-format inheritance remains a
refusal, not a guess from read format. Signed packages refuse editing. Some
read-admitted ZIP metadata/layouts or worksheet encodings are not writable.

Values: exact int in ±999,999,999,999,999, finite float, bool and str. Text allows
32,767 UTF-16 code units and 253 LF characters, with valid Unicode scalars.
Literal `=A1` is text, not a formula. Existing numeric/boolean/text categories
are preserved; an existing non-date blank is `unestablished`. Date-formatted
targets retain safety checks and cannot be filled by numbers. Date/datetime
writes are not public; time/date-time/unknown-format targets refuse writing.

No formula/style/row/sheet/table/pivot operations, multi-worksheet batches,
live Excel control, calculation/rendering adapters, pagination, CLI/MCP or durable
proposal reconstruction. Excel is the selected host; other hosts are not certified.

## Uncertainty and verification

Dependencies are bounded potential-reference proofs, not numerical effects.
Supported static references/functions, fixed names and axis ranges are analyzed
within explicit ceilings; dynamic/unsupported syntax, name issues, excluded
scope and unfinished work are surfaced. Complete, partial and unknown differ.
Samples are not full dependent lists; grouped issues are not all full diagnostics
or alternate paths.

Verification freshly parses staged output, checks declared supported changes,
non-target scalar/structural facts, part inventory, unmodified bytes and reviewed
ZIP-record preservation. It is not a benchmark golden, accounting check, formula
calculation or rendering. Saved downstream caches stay unverified. Obtain
appropriate Excel full recalculation and human/business checks before trusting
answers. Visual status remains `not_evaluated` even when structural checks pass.

## Default ceilings

Public callers can lower ceilings, not raise them.

| Archive limit | Default |
| --- | ---: |
| `max_file_bytes` | 100 MiB |
| `max_part_count` | 2,000 |
| `max_part_bytes` | 200 MiB |
| `max_total_bytes` | 1 GiB |
| `max_xml_elements` | 1,000,000 per parsed part |
| `max_xml_depth` | 100 |

| Impact limit | Default |
| --- | ---: |
| `max_sheets` | 256 |
| `max_scanned_cells` | 250,000 |
| `max_defined_names` | 10,000 |
| `max_formula_cells` | 25,000 |
| `max_formula_chars` | 8,192 |
| `max_formula_nesting` | 64 |
| `max_references` | 100,000 |
| `max_membership_checks` | 1,000,000 |
| `max_visited_nodes` | 10,000 |
| `max_evidence_edges` | 50,000 |

Reads default to 128 requested cells and 1 MiB canonical response. Impact
presentation bounds each diagnostic identity to 1 MiB before hashing.
Canonical transaction evidence is bounded to 16 MiB, with 1–100 edits.
These are admission/work/response limits, not peak-memory/CPU-time/simultaneous
budget/denial-of-service/model-token guarantees. A bounded read can materialize
the whole worksheet. Refusal never silently raises caps or invokes another editor.

## Refusals

Handle `Refusal.code`, not prose. `to_dict()` contains `refusal_schema_version="1.0"`,
`code`, `message`, optional `operation`/`target`, `details`, and nonempty structured
`recovery_options`. Unknown future codes are refusals, not permission to retry.

Current families:

- Source/container: `invalid_path`, `unreadable_file`, `not_a_zip`,
  `unsupported_container`, `unsupported_package_kind`, `encrypted_archive_entry`.
- Archive/XML: `malformed_archive`, `malformed_xml`, `dtd_blocked`,
  `unsafe_archive_path`, `duplicate_archive_entry`, `unsupported_compression`,
  `archive_too_large`, `archive_too_many_parts`, `archive_part_too_large`,
  `archive_expansion_too_large`, `xml_limits_exceeded`.
- Package/sheet: `missing_required_part`, `invalid_part_content`,
  `sheet_not_found`, `unsupported_sheet_kind`.
- Intent/target: `duplicate_target`, `multiple_target_parts`, `missing_cell`,
  `formula_cell`, `formula_range_target`, `merged_cell`, `protected_worksheet`,
  `signed_package`, `rich_text_cell`, `unsupported_target_structure`,
  `unsupported_target_style`, `unsupported_temporal_target`, `incompatible_date_style`,
  `type_mismatch`, `invalid_value`, `unsupported_write_encoding`,
  `unsupported_zip_metadata`, `no_changes`.
- Bounds/coverage: `inspection_limit_exceeded`, `transaction_limit_exceeded`,
  `dependency_analysis_untrustworthy`. Dependency issues/frontiers are separate
  evidence, not automatically parser refusals or complete coverage.
- Authority/delivery: `approval_required`, `approval_rejected`,
  `preview_digest_mismatch`, `unsafe_output_path`, `output_exists`,
  `output_unavailable`, `stale_source`, `stale_output`, `verification_failed`,
  `publication_failed`, `proposal_already_applied`.

Findings can block a Preview instead of returning Refusal. Recovery actions are
suggestions, not automatic repair/permission. Changed intent/output/schema/
limits/source/policy requires a new preview/approval. Never drop blocked members
or approve unknown evidence.

## Security, privacy and delivery

Archive names are validated from raw directory records; package targets cannot
escape the package. DTDs refuse and resource checks precede decompression/tree
building. Parts are in memory, never extracted to disk. This is not a sandbox:
isolate untrusted workloads and apply external operational resource controls.

The application owns identity, UI/store, organization policy, intent, business
checks and evidence storage. A model cannot authenticate its own permission.
Elevated changes require independent human/service authority, not inferred grants.

Output is a separate `.xlsx` in an existing directory. Source aliases,
symlink/non-regular destinations, implicit overwrite and changed bindings refuse.
Preparation/verification precede one publication point; there is no safe-subset
delivery. This is not a concurrent-writer lock or protection against hostile
ancestor races, arbitrary trusted-code tampering, every filesystem guarantee or
every power-loss scenario.

Interruption near publication can leave an output without a returned receipt.
Do not infer output absence from every exception. After a committed cleanup
warning, inspect the receipt and application-owned temporary-file handling,
not blind replay. Receipt persistence is not automatically paired with the workbook.

No implicit telemetry, network, log files or receipt export. Explicit evidence
can contain sensitive cells/formulas/local paths. The application controls access
and where that information goes.

## Compatibility

The alpha may change in a separately versioned/documented revision. Meanings and
approval identity must not silently drift within a version. Package `0.1.0a1`
uses transaction `1.2`, dependency `1.1`, summary/inspection `1.0`, approval/
refusal/evidence schemas `1.0`, and engine policy `1.1`. Older package-bound
approvals are not migrated/relabeled. Local build is not publication, production
certification or a remote CI result.
