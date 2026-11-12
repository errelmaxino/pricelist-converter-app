"""Synthetic regression examples for automotive catalog normalization.

All inputs are invented. Tests check parsing and review behavior, NOT real
part-to-vehicle compatibility or business performance.
"""
from __future__ import annotations

from pathlib import Path
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import parser as catalog


def synthetic_row(description: str, code: str = "VK-DEMO") -> catalog.CatalogRow:
    return catalog.CatalogRow(
        page=1,
        supplier_brand="SYNTHETIC",
        category="Brake Pad",
        code=code,
        axle="Front",
        description=description,
        price=100.0,
        source_line=f"{code} Front {description}",
    )


class CatalogQualityTests(unittest.TestCase):
    def test_short_year_is_normalized_without_claiming_real_fitment(self):
        rows, queue, conflicts, evidence = catalog.expand_catalog_rows(
            [synthetic_row("TOY COROLLA '07")]
        )
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Brand"], "TOYOTA")
        self.assertEqual(rows[0]["Year"], "2007")
        self.assertEqual(rows[0]["Model"], "COROLLA")
        self.assertEqual(rows[0]["Page"], 1)
        self.assertTrue(any(item["Field"] == "Year" for item in evidence))

    def test_slash_separated_applications_keep_source_linkage(self):
        source = synthetic_row("TOY COROLLA '07 / MIT LANCER '08")
        rows, _queue, _conflicts, evidence = catalog.expand_catalog_rows([source])
        self.assertEqual(len(rows), 2)
        self.assertEqual(
            {(r["Brand"], r["Model"], r["Year"]) for r in rows},
            {("TOYOTA", "COROLLA", "2007"), ("MITSUBISHI", "LANCER", "2008")},
        )
        self.assertTrue(all(r["Source Line"] == source.source_line for r in rows))
        self.assertEqual(len({r["_Row Key"] for r in rows}), 2)
        self.assertTrue({r["_Row Key"] for r in rows}.issubset({item["Row Key"] for item in evidence}))

    def test_model_name_bongo_2000_is_not_mistaken_for_year(self):
        rows, _queue, _conflicts, _evidence = catalog.expand_catalog_rows(
            [synthetic_row("MAZ BONGO 2000")]
        )
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Brand"], "MAZDA")
        self.assertEqual(rows[0]["Model"], "BONGO 2000")
        self.assertEqual(rows[0]["Year"], "")

    def test_unknown_brand_stays_reviewable_without_inventing_brand(self):
        rows, queue, _conflicts, _evidence = catalog.expand_catalog_rows(
            [synthetic_row("COROLLA 2007")]
        )
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Brand"], "")
        self.assertEqual(rows[0]["Review Status"], "Needs Review")
        self.assertEqual(len(queue), 1)


if __name__ == "__main__":
    unittest.main()
