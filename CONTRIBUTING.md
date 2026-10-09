# Contributing

Work on one bounded capability at a time. Start with supported behavior,
refusal boundaries and verification evidence; do not broaden a writer or API
as a side effect of cleanup.

## Development checks

Use Python 3.11 or newer and an isolated environment. From the repository root:

```bash
python -m venv .venv
python -m pip install -e ".[dev]"
python -m pip install build twine
ruff check .
ruff format --check .
mypy
python -m build --outdir dist
twine check --strict dist/*
XLAYER_ARTIFACT_DIR=dist pytest
```

Activate the environment or use its executables explicitly (`.venv/bin/` on
POSIX; `.venv\Scripts\` on Windows). Use a fresh artifact directory: explicit
artifact selection rejects missing/ambiguous files rather than silently skipping.
Without built artifacts, some artifact checks skip; that is not a packaging pass.
Windows PowerShell can set `$env:XLAYER_ARTIFACT_DIR = "dist"` before `pytest`.

CI retains Python 3.11–3.14 across Linux/macOS/Windows, plus lint and packaging.
Packaging installs both wheel and sdist into separate clean environments and
runs the full suite against installed imports, not `src/`. Local tests do not
prove that a pending remote matrix passed.

## Required invariants

- Supported imports are exactly the root allowlist; underscore modules are private.
  Add an API/export/artifact gate with a support change, not an exemption.
- Zero runtime dependencies; no research repository or provider SDK imports.
- Preserve exact-preview authority, whole-batch refusal, explicit separate output,
  snapshot identity and owned-resource cleanup. No safe-subset fallback.
- Saved caches are not fresh calculation, reference proofs are not numerical
  effects, and structural verification is not rendered/business correctness.
- Every support claim and public schema/version must match tested behavior.

## Tests and publishing hygiene

Write a regression first, observe its intended failure, then fix the owning
boundary. Keep the full existing suite. Semantic expected answers use handwritten
inputs/raw OOXML or independent evidence, not production parser/writer output.
Cross-implementation equality is an equivalence check, not a correctness oracle.
Never regenerate a golden merely to make a test green.

Use included sanitized workbooks for real-file integration and small crafted
archives for adversarial limits. Do not publish private workbooks, extracted
formulas/values, personal paths, credentials, model/provider settings, evaluator
answers, internal working notes or unreviewed fixture provenance. Check sdists
too: they contain tests and approved fixtures.

Docs/examples ship by exact filename. Wheels contain only package code, typing
and metadata; sdists include approved source/test/docs/example files. Update
required-file and source-byte gates in the same change. Do not add broad doc
globs. Building an artifact does not authorize pushing a release.
