"""Exercise saved-fact checking and the publication example after installation."""

from __future__ import annotations

import json
import runpy
from pathlib import Path

from xlayer import CheckedExplanation, ExplanationRequest, ExplanationScope, Refusal, Workbook

ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "tests/fixtures/test_workbook_7_write_v11_edges.xlsx"


def test_installed_root_api_checks_saved_xml_fact() -> None:
    scope = ExplanationScope({"edges": ["A1"]})
    with Workbook.open(FIXTURE) as book:
        request = ExplanationRequest.from_dict(
            {
                "schema_version": "1.0",
                "explanation_contract_version": "1.0",
                "source_fingerprint": book.source_fingerprint,
                "scope_digest": scope.scope_digest,
                "claims": [
                    {
                        "id": "saved",
                        "kind": "fact",
                        "selector": {
                            "family": "cell",
                            "sheet": "edges",
                            "address": "A1",
                            "field": "raw",
                        },
                        "expected": "100",
                    }
                ],
            }
        )
        assert isinstance(request, ExplanationRequest)
        report = book.check_explanation(request, scope=scope)
        assert isinstance(report, CheckedExplanation)
    assert report.report_digest.startswith("sha256:")
    assert 'raw="100"; decoded=100.0' in report.render_text()
    assert report.to_dict()["limitations"]["calculation"] == "not_evaluated"  # type: ignore[index]


def test_example_clears_old_report_before_malformed_or_wrong_submission() -> None:
    example = runpy.run_path(str(ROOT / "examples/explain_saved_facts.py"))
    publisher = example["ExplanationPublisher"]()
    scope = ExplanationScope({"edges": ["A1"]})
    with Workbook.open(FIXTURE) as book:
        payload = {
            "schema_version": "1.0",
            "explanation_contract_version": "1.0",
            "source_fingerprint": book.source_fingerprint,
            "scope_digest": scope.scope_digest,
            "claims": [
                {
                    "id": "saved",
                    "kind": "fact",
                    "selector": {
                        "family": "cell",
                        "sheet": "edges",
                        "address": "A1",
                        "field": "raw",
                    },
                    "expected": "100",
                }
            ],
        }
        good = json.dumps(payload)
        assert isinstance(publisher.submit(book, scope, good), CheckedExplanation)
        historical = publisher.current
        assert historical is not None
        assert isinstance(publisher.submit(book, scope, "{"), Refusal)
        assert publisher.current is None
        assert historical.render_text()
        assert isinstance(publisher.submit(book, scope, good), CheckedExplanation)
        assert isinstance(publisher.submit(book, scope, good.replace('"100"', '"999"')), Refusal)
        assert publisher.current is None
        assert isinstance(publisher.submit(book, scope, good), CheckedExplanation)
        changed_scope = ExplanationScope({"edges": ["A1", "B1"]})
        assert isinstance(publisher.submit(book, changed_scope, good), Refusal)
        assert publisher.current is None
