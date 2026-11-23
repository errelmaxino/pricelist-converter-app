"""Invented PDFs and text fragments, reproducing bugs seen in private catalog QA.

Never commit real supplier text, part IDs, or prices from the source PDFs.
"""
from __future__ import annotations

import tempfile
from pathlib import Path
import unittest

import parser
from tests.synthetic_catalog_fixture import pdf_with_lines


class PriceColumnTests(unittest.TestCase):
    def test_right_edge_split_price_overrides_zero_tail_and_removes_digit_from_description(self):
        # Invented line with a price printed in two separate PDF text fragments.
        line = {
            "text": "VK-9000 front TOYOTA PRACTICE 2011 7 00.00",
            "words": [
                {"text": "VK-9000", "x0": 28.2, "x1": 80},
                {"text": "front", "x0": 111.0, "x1": 143.0},
                {"text": "TOYOTA", "x0": 149.0, "x1": 187.0},
                {"text": "PRACTICE", "x0": 190.0, "x1": 237.0},
                {"text": "2011", "x0": 245.0, "x1": 260.0},
                {"text": "7", "x0": 522.8, "x1": 528.0},
                {"text": "00.00", "x0": 528.4, "x1": 553.0},
            ],
        }
        row = parser._extract_row_from_line(line, 612.0)
        self.assertIsNotNone(row)
        self.assertEqual(row["price"], 700.0)
        self.assertEqual(row["description"], "TOYOTA PRACTICE 2011")

    def test_right_edge_three_part_price_is_reassembled(self):
        line = {
            "text": "VK-9001 rear SUZUKI PRACTICE 7 0 0.00",
            "words": [
                {"text": "VK-9001", "x0": 28, "x1": 85},
                {"text": "rear", "x0": 110, "x1": 135},
                {"text": "SUZUKI", "x0": 150, "x1": 190},
                {"text": "PRACTICE", "x0": 200, "x1": 260},
                {"text": "7", "x0": 514, "x1": 520},
                {"text": "0", "x0": 521, "x1": 526},
                {"text": "0.00", "x0": 527, "x1": 551},
            ],
        }
        row = parser._extract_row_from_line(line, 612.0)
        self.assertEqual(row["price"], 700.0)
        self.assertEqual(row["description"], "SUZUKI PRACTICE")

    def test_normal_price_remains_unchanged(self):
        line = {
            "text": "VK-9002 rear TOYOTA PRACTICE 650.00",
            "words": [
                {"text": "VK-9002", "x0": 28, "x1": 75},
                {"text": "rear", "x0": 109, "x1": 130},
                {"text": "TOYOTA", "x0": 149, "x1": 190},
                {"text": "PRACTICE", "x0": 195, "x1": 260},
                {"text": "650.00", "x0": 522, "x1": 552},
            ],
        }
        row = parser._extract_row_from_line(line, 612.0)
        self.assertEqual(row["price"], 650.0)
        self.assertEqual(row["description"], "TOYOTA PRACTICE")


class SectionChangeTests(unittest.TestCase):
    def test_section_heading_does_not_drop_product_before_it(self):
        lines = [
            "BRAKE PAD",
            "VK-9010 rear TOYOTA PRACTICE 920.00",
            "HAND BRAKE SHOE",
            "VK-9011 rear SUZUKI PRACTICE 460.00",
        ]
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "invented-sections.pdf"
            p.write_bytes(pdf_with_lines(lines))
            raw, warnings = parser.extract_catalog_rows(p)
        self.assertEqual([r.code for r in raw], ["VK-9010", "VK-9011"])
        self.assertEqual([r.category for r in raw], ["Brake Pad", "Hand Brake Shoe"])
        self.assertEqual(warnings, [])


if __name__ == "__main__":
    unittest.main()
