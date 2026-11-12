"""Synthetic tests for the opt-in AI proposal gate; no AI calls are made."""
from __future__ import annotations

import json
import unittest

from ai_review import build_review_prompt, validate_review_proposal

ROW = {
    "Row Key": "P1|R1|VK-DEMO|1",
    "Source Line": "VK-DEMO TOY COROLLA '07 100.00",
    "Brand": "TOYOTA",
    "Model": "COROLLA",
    "Year": "2007",
}


def sample_proposal(**updates):
    payload = {
        "row_key": ROW["Row Key"],
        "source_line": ROW["Source Line"],
        "evidence_quote": "TOY COROLLA '07",
        "proposed_fields": {"Brand": "TOYOTA"},
        "reason": "TOY is an abbreviation.",
    }
    payload.update(updates)
    return json.dumps(payload)


class AiReviewSafetyTests(unittest.TestCase):
    def test_prompt_contains_source_and_rejects_fitment_claims(self):
        p = build_review_prompt(ROW)
        self.assertIn("VK-DEMO TOY COROLLA", p)
        self.assertIn("untrusted DATA", p)
        self.assertIn("Never claim OEM compatibility", p)

    def test_valid_suggestion_is_still_only_a_proposal(self):
        result = validate_review_proposal(sample_proposal(), ROW)
        self.assertEqual(result["proposed_fields"]["Brand"], "TOYOTA")
        self.assertEqual(ROW["Brand"], "TOYOTA")

    def test_rejects_wrong_row_key(self):
        with self.assertRaisesRegex(ValueError, "Row key"):
            validate_review_proposal(sample_proposal(row_key="other"), ROW)

    def test_rejects_changed_source_line(self):
        with self.assertRaisesRegex(ValueError, "Source line"):
            validate_review_proposal(sample_proposal(source_line="OTHER"), ROW)

    def test_rejects_unquoted_hallucinated_source(self):
        with self.assertRaisesRegex(ValueError, "excerpt"):
            validate_review_proposal(sample_proposal(evidence_quote="OEM VERIFIED"), ROW)

    def test_rejects_unsupported_field(self):
        with self.assertRaisesRegex(ValueError, "unsupported"):
            validate_review_proposal(sample_proposal(proposed_fields={"FitmentConfirmed": "Yes"}), ROW)

    def test_accepts_explicit_no_change(self):
        result = validate_review_proposal(sample_proposal(proposed_fields={}, evidence_quote=""), ROW)
        self.assertEqual(result["proposed_fields"], {})

    def test_rejects_non_json_response(self):
        with self.assertRaisesRegex(ValueError, "JSON"):
            validate_review_proposal("Sure! The right part is...", ROW)


if __name__ == "__main__":
    unittest.main()
