# xlayer

An open-source Python transaction layer for safe, auditable changes to `.xlsx`
workbooks.

> **Status: pre-release. No supported workbook API exists yet.**
> Internal parsing infrastructure is under construction as private modules, but
> nothing is exposed: installing this package gives you a version number and no
> supported way to read, inspect, or modify a workbook. It is published in this
> state so that packaging, typing, and release gates can be reviewed while
> capability is promoted one tested slice at a time. Everything under
> "Intended design" below is a target, not current behaviour.

## The problem

Most Python spreadsheet libraries expose mutation primitives and leave the caller
to decide what is safe. That is the right trade-off for a deterministic script
written by someone who already knows the workbook.

It becomes risky when the caller is a general application or an AI system working
on an unfamiliar file, where a half-applied change, a silently overwritten
formula, or a stale cached value is unacceptable. The missing piece is not file
access — it is a trustworthy change protocol.

## Intended design

xlayer is intended to wrap a workbook change in a controlled lifecycle:

```text
understand -> propose -> preview -> approve or refuse
           -> apply the complete batch atomically -> verify -> receipt
```

The design commitments that lifecycle exists to serve, none of which are
implemented yet:

- **Whole-batch atomicity.** A proposal applies completely or writes nothing.
  There is no automatic "apply the safe subset", because that produces a
  partially edited workbook while reporting that the plan failed.
- **Determinism.** The same workbook bytes and the same proposal produce the same
  parse, preview, and mutation decisions. No AI calls sit inside parsing or
  mutation.
- **Structured refusals.** Unsupported or ambiguous cases return machine-readable
  refusals carrying recovery information, rather than best-effort guesses or
  exception strings a caller has to parse.
- **Honest status reporting.** When a change leaves formulas needing a
  spreadsheet host to recalculate, the result says so. Verification never implies
  that formula results were recomputed when nothing recomputed them.
- **Source preservation.** Writes go to an explicit output path. The input
  workbook is left unchanged.
- **Structural, not domain, semantics.** xlayer can derive structural facts such
  as regions, dependencies, and header-like cells. It does not claim to know that
  a cell is a covenant, a tax rule, or a correct business assumption.
- **No provider dependency.** The base install requires no AI SDK and works
  entirely offline.

Capabilities are promoted into this package from a separate research repository
one tested vertical slice at a time, each arriving with its own tests, refusal
boundaries, and fidelity evidence. Nothing is described here as supported until
it has passed those gates.

## What you get today

```bash
pip install xlayer
```

```python
import xlayer

xlayer.__version__  # "0.1.0.dev0"
```

There is no public API beyond that version string, and a test in this repository
enforces it. Internal parsing infrastructure is being built inside the package
as private modules, but none of it is exposed or supported yet.

## Scope boundaries

xlayer is not intended to be:

- a replacement for `openpyxl`, `pandas`, or `XlsxWriter` for general workbook
  reading, analysis, or generation;
- a spreadsheet calculation engine;
- an AI model, prompt framework, or natural-language planner;
- a source of financial, accounting, or industry-specific truth;
- a live Microsoft Excel controller;
- a promise of safe arbitrary OOXML mutation; or
- a mechanism for partially applying a rejected change.

For broad workbook generation or direct manipulation, existing libraries remain
the right tool. xlayer is for workflows where preview, refusal, atomicity,
verification, and auditability matter more than breadth.

## Requirements

Python 3.11 or newer. No runtime dependencies.

## Development

```bash
git clone https://github.com/YCBytes/xlayer.git
cd xlayer
python -m venv .venv
.venv/bin/python -m pip install -e ".[dev]"
```

```bash
.venv/bin/ruff check .          # lint
.venv/bin/ruff format --check . # formatting
.venv/bin/mypy                  # strict type checking
.venv/bin/pytest                # tests
```

## License

Apache License 2.0. See [LICENSE](LICENSE).
