"""Reproducible synthetic catalog scorecard. NOT real-world fitment accuracy."""
from __future__ import annotations

from collections import Counter
from pathlib import Path
import json
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import parser as catalog
from tests.synthetic_catalog_fixture import CASES, expected_records, generated_pdf


def main():
    with tempfile.TemporaryDirectory() as d:
        pdf = Path(d) / "synthetic-catalog.pdf"
        pdf.write_bytes(generated_pdf())
        parsed = catalog.parse_catalog_pdf(pdf)
    expected = Counter(expected_records())
    actual = Counter((str(r["Code"]), str(r["Brand"]), str(r["Model"]), str(r["Year"]))
                     for r in parsed.application_rows)
    correct = sum((expected & actual).values())
    gold_count = sum(expected.values())
    output_count = sum(actual.values())
    precision = correct / output_count if output_count else 0
    recall = correct / gold_count if gold_count else 0
    report = {
        "benchmark": "synthetic-text-PDF-only",
        "expected_applications": gold_count,
        "extracted_catalog_lines": len(parsed.raw_rows),
        "output_applications": output_count,
        "exact_record_matches": correct,
        "exact_record_precision": round(precision, 4),
        "exact_record_recall": round(recall, 4),
        "rows_flagged_needs_review": parsed.metrics["needs_review"],
        "unique_lineage_ids": len({r["_Row Key"] for r in parsed.application_rows}),
        "warning": "Synthetic normalization agreement, NOT fitment validation or real-world accuracy",
    }
    print(json.dumps(report, indent=2))
    if correct != gold_count or output_count != gold_count:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
