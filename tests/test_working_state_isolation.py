"""Working state must track file contents, not merely a filename and byte count."""
from __future__ import annotations

import ast
import hashlib
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[1]


def get_real_signature_function():
    ast_module = ast.parse((ROOT / "streamlit_app.py").read_text(encoding="utf-8"))
    func = next(node for node in ast_module.body
                if isinstance(node, ast.FunctionDef) and node.name == "_signature")
    scope = {"hashlib": hashlib}
    exec(compile(ast.Module(body=[func], type_ignores=[]), "<signature>", "exec"), scope)
    return scope["_signature"]


class WorkingStateIsolationTests(unittest.TestCase):
    def test_same_filename_and_size_but_different_bytes_reset_state(self):
        signature = get_real_signature_function()
        s1 = signature("catalog.pdf", b"AAA", "Synthetic", "", {}, [])
        s2 = signature("catalog.pdf", b"BBB", "Synthetic", "", {}, [])
        self.assertNotEqual(s1, s2, "Different PDFs must not share edited working state.")

    def test_same_content_and_parameters_remains_stable(self):
        signature = get_real_signature_function()
        s1 = signature("catalog.pdf", b"AAA", "Synthetic", "", {"TOY": "TOYOTA"}, [])
        s2 = signature("catalog.pdf", b"AAA", "Synthetic", "", {"TOY": "TOYOTA"}, [])
        self.assertEqual(s1, s2)


if __name__ == "__main__":
    unittest.main()
