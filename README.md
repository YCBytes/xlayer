# xlayer

An experimental Python transaction layer for controlled changes to `.xlsx`
workbooks. Inspect stored facts, preview a change, require exact-preview approval,
verify the produced file, and receive a receipt. Xlayer runs offline with no
runtime dependencies or AI provider built in.

**Status: public alpha `0.1.0a1`.** This is not a stable release or production
certification. Building this version does not publish it to PyPI.

## What works

- Bounded sheet inventory, explicit cell reads and potential-reference impact
  summaries with coverage limitations intact.
- `SetValue` for existing supported inputs: exact Python `int`, finite `float`,
  `bool` or `str`, in one worksheet per transaction.
- Application-owned approval bound to the exact source, proposal and preview.
- Whole-batch editing to an explicit separate output, staged verification and
  returned evidence. No automatic safe-subset fallback.

Formula results are saved caches, never verified calculations. Xlayer does not
render workbooks or establish business correctness. Date writes, formula writes,
missing-cell creation, structural editing and live Excel control are unsupported.
Excel is the selected host; other spreadsheet hosts are not certified.

## Install

Requires Python 3.11 or newer. From a checkout of this repository:

```bash
python -m pip install .
```

Or build a wheel with `python -m build` and install that local wheel. Use an
isolated environment for evaluation. See [Contributing](CONTRIBUTING.md).

## Inspect, then propose

```python
from xlayer import Refusal, SetValue, Workbook

with Workbook.open("inputs.xlsx") as book:
    observation = book.read_cells("Inputs", ["C12", "D12"])
    if isinstance(observation, Refusal):
        print(observation.to_dict())
    else:
        print(observation.to_dict())  # saved formula results are unverified
        proposal = book.propose([SetValue("Inputs", "C12", 120)], output_path="updated.xlsx")
        if isinstance(proposal, Refusal):
            print(proposal.to_dict())
        else:
            print(proposal.preview().to_dict())  # preview creates no output
```

Opening can raise `WorkbookOpenError` carrying a structured `.refusal`. This
example stops before approval; it does not modify a workbook. The application
decides what is permitted, authenticates an `Approval` through its required
verifier, and handles the returned receipt/refusal. A model's claim is not authority.

See [the tested approval-store example](examples/approved_set_value.py) for the
complete provider-free flow. From a checkout with Xlayer installed, using the
included disposable fixture:

```python
from examples.approved_set_value import edit_one_cell
from xlayer import SetValue

result = edit_one_cell(
    "tests/fixtures/test_workbook_7_write_v11_edges.xlsx",
    "updated.xlsx",  # must not already exist
    SetValue("edges", "F1", 44),
    permitted=True,  # explicit application permission for this demonstration only
)
print(result.to_dict())
```

The example store is trusted demonstration code, not user authentication. It
denies unrecorded/mismatched claims and does not auto-grant elevated permissions.
In a real application, show the preview and obtain an independent decision
before recording approval. Receipt persistence belongs to the application.

## Documentation

- [API](docs/api.md): supported imports, bounded facts, lifecycle and approval.
- [Support and limitations](docs/support.md): targets, refusals, resource/security
  boundaries and what verification does not prove.
- [Changelog](CHANGELOG.md): versioned support changes.

Serialized evidence can contain workbook content and local paths. The application
owns access control and persistence. Nothing is logged or uploaded implicitly.

## License

Apache License 2.0. See [LICENSE](LICENSE).
