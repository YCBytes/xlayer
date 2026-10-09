# Changelog

## 0.1.0a1 — experimental public API

- Supported root imports for managed sessions, bounded stored-fact reads,
  dependency summaries and approval-bound `SetValue` transactions.
- Immutable versioned inspections and explicit response/work ceilings.
- Public scalar inputs restricted to exact int, finite float, bool and str.
  Date/datetime writes remain unsupported in this alpha.
- Provider-free approval-store example and installed-artifact workflow tests;
  focused API/support/contributor documentation.
- Existing whole-batch atomicity, source preservation, structural verification
  and returned receipts retained through thin facades.

This is the first supported experimental interface, not a stable `0.1.0` release
or a statement that this artifact has been published. Formula caches remain
unverified and visual status remains `not_evaluated`. Transaction contract `1.2`,
dependency `1.1`, impact summary `1.0` and engine policy `1.1` remain unchanged;
inspection contract starts at `1.0`. New package-bound approvals are required.
