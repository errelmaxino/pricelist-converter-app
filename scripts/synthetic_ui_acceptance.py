"""Browser acceptance for the existing Streamlit converter, using invented data only.

No real supplier files, external AI calls, credentials, or production services.
Run after starting Streamlit locally on port 8501.
"""
from __future__ import annotations

import os
import re
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

from tests.synthetic_catalog_fixture import expected_records


def main() -> None:
    screenshots = Path(os.environ.get("UI_ACCEPTANCE_OUT", "artifacts/synthetic-ui"))
    screenshots.mkdir(parents=True, exist_ok=True)

    with TemporaryDirectory(prefix="synthetic-converter-") as tmp:
        pdf_path = Path(tmp) / "invented-catalog.pdf"
        xlsx_path = Path(tmp) / "invented-review.xlsx"

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(
                viewport={"width": 1366, "height": 900},
                accept_downloads=True,
            )
            page.set_default_timeout(60000)
            page.goto("http://127.0.0.1:8501", wait_until="domcontentloaded")
            page.get_by_role(
                "heading", name="Automotive Parts Data Quality Lab"
            ).wait_for()

            supplier = page.get_by_role("textbox", name="Supplier name")
            supplier.fill("INVENTED DATA — DEMO ONLY")
            supplier.press("Tab")
            page.get_by_role(
                "button", name="Download synthetic sample PDF"
            ).wait_for()

            with page.expect_download() as pdf_download:
                page.get_by_role(
                    "button", name="Download synthetic sample PDF"
                ).click()
            pdf_download.value.save_as(pdf_path)
            assert pdf_path.read_bytes().startswith(b"%PDF-"), "Invalid sample PDF"

            inputs = page.locator('input[type="file"]')
            assert inputs.count() >= 2, "PDF uploader not found"
            inputs.last.set_input_files(str(pdf_path))
            page.get_by_text("Conversion summary").wait_for()
            metrics = page.locator('[data-testid="stMetric"]')
            catalog = metrics.filter(has_text="Catalog rows").inner_text()
            applications = metrics.filter(has_text="Application rows").inner_text()
            assert re.search(r"\b8\b", catalog), f"Unexpected catalog rows: {catalog}"
            assert re.search(r"\b9\b", applications), (
                f"Unexpected application rows: {applications}"
            )
            page.get_by_role("tab", name="7) Export").wait_for()
            page.screenshot(
                path=str(screenshots / "synthetic-summary.png"), full_page=True
            )

            page.get_by_role("tab", name="7) Export").click()
            page.get_by_role(
                "button", name="Download Excel workbook"
            ).wait_for()
            page.get_by_text("Workbook sheets").wait_for()
            page.get_by_text("Evidence Log", exact=True).last.wait_for()
            page.screenshot(
                path=str(screenshots / "synthetic-export.png"), full_page=True
            )

            with page.expect_download() as excel_download:
                page.get_by_role(
                    "button", name="Download Excel workbook"
                ).click()
            excel_download.value.save_as(xlsx_path)
            browser.close()

        wb = load_workbook(xlsx_path, data_only=False, read_only=True)
        expected_tabs = {
            "Applications", "Catalog Rows", "Review Queue",
            "Conflict Log", "Evidence Log",
        }
        assert expected_tabs.issubset(set(wb.sheetnames)), wb.sheetnames

        rows = wb["Applications"].iter_rows(values_only=True)
        headers = next(rows)
        names = [str(v) if v is not None else "" for v in headers]
        for name in ("Code", "Brand", "Model", "Year", "Review Status", "Row Key"):
            assert name in names, f"Missing output field: {name}"
        values = list(rows)
        assert len(values) == len(expected_records()), (
            f"Unexpected exported applications: {len(values)}"
        )
        idx = {name: names.index(name) for name in names}
        normalized = [
            tuple(str(row[idx[key]] or "") for key in ("Code", "Brand", "Model", "Year"))
            for row in values
        ]
        assert sorted(normalized) == sorted(expected_records()), (
            "Downloaded workbook differs from invented expected records"
        )
        unresolved = [
            row for row in values if row[idx["Code"]] == "VK-104"
        ]
        assert len(unresolved) == 1, "Invented ambiguous row missing"
        assert not unresolved[0][idx["Brand"]], (
            "Ambiguous brand was incorrectly accepted"
        )
        assert unresolved[0][idx["Review Status"]] == "Needs Review", (
            "Invented ambiguous row should remain for human review"
        )
        evidence = wb["Evidence Log"]
        evidence_rows = evidence.iter_rows(values_only=True)
        evidence_headers = next(evidence_rows)
        assert "Row Key" in evidence_headers
        evidence_key = list(evidence_headers).index("Row Key")
        keys = {row[evidence_key] for row in evidence_rows}
        assert all(row[idx["Row Key"]] in keys for row in values), (
            "Some exported rows lack matching source evidence"
        )
        wb.close()

    print("PASS: Browser opened converter and downloaded invented PDF")
    print("PASS: Browser uploaded PDF and showed 8 catalog rows / 9 applications")
    print("PASS: Exported Excel has 5 expected tabs and 9 exact invented records")
    print("PASS: Ambiguous brand remains in human review")
    print("PASS: Every exported application is linked to evidence")
    print("PASS: Screenshots captured from invented data only")


if __name__ == "__main__":
    main()
