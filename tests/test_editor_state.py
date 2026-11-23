"""Invented rows: edits in the grid must not overwrite newer review-form changes."""
from __future__ import annotations

import ast
from pathlib import Path
from types import SimpleNamespace
import unittest

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]


def load_state_functions(state):
    module = ast.parse((ROOT / "streamlit_app.py").read_text())
    names = {"_working_df", "_set_working_df", "_merge_editor_subset"}
    nodes = [n for n in module.body if isinstance(n, ast.FunctionDef) and n.name in names]
    assert len(nodes) == 3
    scope = {"st": SimpleNamespace(session_state=state), "pd": pd,
             "EDITABLE_FIELDS": ["Brand", "Model", "Year", "Your Price (PHP)"]}
    exec(compile(ast.Module(body=nodes, type_ignores=[]), "<ui-state>", "exec"), scope)
    return scope


class EditorStateTests(unittest.TestCase):
    def test_unedited_preview_cannot_erase_form_correction(self):
        state = {"working_df": pd.DataFrame([{"Brand": "FORM_CORRECTION", "Model": "TEST"}]),
                 "editor_revision": 1}
        fn = load_state_functions(state)["_merge_editor_subset"]
        stale = pd.DataFrame([{"Brand": "", "Model": "TEST"}])
        fn(stale, {})
        self.assertEqual(state["working_df"].at[0, "Brand"], "FORM_CORRECTION")

    def test_grid_edit_changes_only_the_explicit_field(self):
        state = {"working_df": pd.DataFrame([{"Brand": "", "Model": "FORM_MODEL"}]),
                 "editor_revision": 0}
        fn = load_state_functions(state)["_merge_editor_subset"]
        preview = pd.DataFrame([{"Brand": "DEMO_BRAND", "Model": "OLD_MODEL"}])
        fn(preview, {0: {"Brand": "DEMO_BRAND"}})
        self.assertEqual(state["working_df"].at[0, "Brand"], "DEMO_BRAND")
        self.assertEqual(state["working_df"].at[0, "Model"], "FORM_MODEL")
        self.assertEqual(state["editor_revision"], 1)

    def test_external_edit_invalidates_previous_grid_version(self):
        state = {"working_df": pd.DataFrame([{"Brand": ""}]),
                 "editor_revision": 3}
        fn = load_state_functions(state)["_set_working_df"]
        fn(pd.DataFrame([{"Brand": "REVIEWED"}]))
        self.assertEqual(state["editor_revision"], 4)


if __name__ == "__main__":
    unittest.main()
