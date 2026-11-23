"""End-to-end accuracy gate for a synthetic PDF (not real-world fitment)."""
from __future__ import annotations

import io
import tempfile
from pathlib import Path
import unittest

import parser as catalog
from tests.synthetic_catalog_fixture import CASES, expected_records, generated_pdf


class EndToEndSyntheticTests(unittest.TestCase):
    def test_text_pdf_normalizes_all_expected_records(self):
        with tempfile.TemporaryDirectory() as d:
            filename = Path(d) / "synthetic-catalog.pdf"
            filename.write_bytes(generated_pdf())
            parsed = catalog.parse_catalog_pdf(filename)

        self.assertEqual(len(parsed.raw_rows), len(CASES))
        actual = [(str(r["Code"]), str(r["Brand"]), str(r["Model"]), str(r["Year"]))
                  for r in parsed.application_rows]
        self.assertCountEqual(actual, expected_records())
        self.assertEqual(len({r["_Row Key"] for r in parsed.application_rows}),
                         len(parsed.application_rows))
        evidence_keys = {e["Row Key"] for e in parsed.evidence_rows}
        for row in parsed.application_rows:
            self.assertIn(row["_Row Key"], evidence_keys)
            self.assertEqual(row["Page"], 1)
            self.assertTrue(row["Source Line"].startswith(str(row["Code"])))

    def test_unresolved_brand_stays_in_review(self):
        with tempfile.TemporaryDirectory() as d:
            filename = Path(d) / "synthetic-catalog.pdf"
            filename.write_bytes(generated_pdf())
            parsed = catalog.parse_catalog_pdf(filename)
        row = next(r for r in parsed.application_rows if r["Code"] == "VK-104")
        self.assertEqual(row["Brand"], "")
        self.assertEqual(row["Review Status"], "Needs Review")

    def test_export_keeps_source_key_on_application_and_evidence_tabs(self):
        with tempfile.TemporaryDirectory() as d:
            filename = Path(d) / "synthetic-catalog.pdf"
            filename.write_bytes(generated_pdf())
            parsed = catalog.parse_catalog_pdf(filename)
        df = catalog.applications_dataframe(parsed.application_rows)
        wb = catalog.build_workbook(
            df,
            catalog.raw_rows_dataframe(parsed.raw_rows),
            df[df["Review Status"] != "Auto-Accepted"],
            catalog.conflict_dataframe(parsed.conflict_rows),
            catalog.evidence_dataframe(parsed.evidence_rows),
        )
        from openpyxl import load_workbook
        doc = load_workbook(io.BytesIO(catalog.workbook_to_bytes(wb)), data_only=False)
        apps = doc["Applications"]
        app_hdr = [c.value for c in apps[1]]
        self.assertIn("Row Key", app_hdr)
        col_idx = app_hdr.index("Row Key") + 1
        exported = {apps.cell(i, col_idx).value for i in range(2, apps.max_row + 1)}
        evidence = doc["Evidence Log"]
        ev_header = [c.value for c in evidence[1]]
        self.assertIn("Row Key", ev_header)
        ev_idx = ev_header.index("Row Key") + 1
        self.assertTrue(exported.issubset({
            evidence.cell(i, ev_idx).value for i in range(2, evidence.max_row + 1)
        }))


if __name__ == "__main__":
    unittest.main()
