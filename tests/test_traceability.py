"""Traceability: synthetic duplicate codes must never attach another row's evidence."""
from __future__ import annotations
import ast
from pathlib import Path
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import parser as catalog


def synthetic_row(description, code="VK-REPEAT", page=1):
    return catalog.CatalogRow(
        page=page, supplier_brand="SYNTHETIC", category="Brake Pad",
        code=code, axle="Front", description=description, price=100.0,
        source_line=f"{code} Front {description} 100.00",
    )


def ui_key_for_row(df, idx):
    """Extract the actual UI helper without executing the Streamlit app."""
    app = ast.parse((ROOT / "streamlit_app.py").read_text(encoding="utf-8"))
    fn = next(x for x in app.body
              if isinstance(x, ast.FunctionDef) and x.name == "_row_key_for_idx")
    ns = {}
    exec(compile(ast.Module(body=[fn], type_ignores=[]), "<ui-key-helper>", "exec"), ns)
    return ns["_row_key_for_idx"](df, idx)


class TraceabilityTests(unittest.TestCase):
    def setUp(self):
        self.rows, self.review, self.conflicts, self.evidence = catalog.expand_catalog_rows([
            synthetic_row("TOY COROLLA '07", code="VK-DUP"),
            synthetic_row("MIT LANCER '08", code="VK-DUP"),
            synthetic_row("MAZ BONGO 2000", code="VK-OTHER"),
        ])

    def test_duplicate_codes_get_distinct_source_keys(self):
        self.assertEqual(len(self.rows), 3)
        self.assertEqual(len({row["_Row Key"] for row in self.rows}), 3)

    def test_evidence_joins_to_exact_application_row(self):
        df = catalog.applications_dataframe(self.rows)
        self.assertIn("Row Key", df.columns)
        self.assertEqual(len(df), 3)
        for i, row in enumerate(self.rows):
            key = row["_Row Key"]
            self.assertEqual(df.iloc[i]["Row Key"], key)
            self.assertEqual(ui_key_for_row(df, i), key)
            own = [item for item in self.evidence if item["Row Key"] == key]
            self.assertTrue(own)
            self.assertEqual(own[0]["Code"], row["Code"])
            self.assertEqual(own[0]["Page"], row["Page"])
            self.assertTrue(all(item["Row Key"] == key for item in own))
        self.assertEqual({item["Row Key"] for item in self.evidence},
                         {row["_Row Key"] for row in self.rows})


if __name__ == "__main__":
    unittest.main()
