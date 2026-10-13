"""Check JSON claims before displaying a saved-fact report.

Run: python examples/explain_saved_facts.py workbook.xlsx Sheet1 A1 claims.json
The application chooses scope. An AI may generate the claims JSON using
ExplanationRequest.json_schema() and the current source/scope identities.
Workbook data and rendered output are untrusted plain text; escape for your UI.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from xlayer import (
    CheckedExplanation,
    ExplanationRequest,
    ExplanationScope,
    Refusal,
    Workbook,
)


class ExplanationPublisher:
    """Example application state; one candidate for the current submission."""

    def __init__(self) -> None:
        self.current: CheckedExplanation | None = None

    def submit(
        self, book: Workbook, scope: ExplanationScope, payload: str | bytes
    ) -> CheckedExplanation | Refusal:
        # Clear before parsing too: malformed replacement input must not leave
        # a previous success looking like approval of the new answer.
        self.current = None
        request = ExplanationRequest.from_json(payload)
        if isinstance(request, Refusal):
            return request
        result = book.check_explanation(request, scope=scope)
        if isinstance(result, CheckedExplanation):
            self.current = result
        return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    parser.add_argument("sheet")
    parser.add_argument("address")
    parser.add_argument("claims", type=Path)
    args = parser.parse_args()
    scope = ExplanationScope({args.sheet: [args.address]})
    publisher = ExplanationPublisher()
    # The library also enforces the request limit; bound the application's file read.
    with args.claims.open("rb") as stream:
        payload = stream.read(1_048_577)
    with Workbook.open(args.workbook) as book:
        result = publisher.submit(book, scope, payload)
    if isinstance(result, Refusal):
        print(f"No checked report: {result.code}")
    else:
        print(result.render_text(), end="")


if __name__ == "__main__":
    main()
