from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from typing import Any, Dict, List, Optional

APP_DIR = Path(__file__).resolve().parent / "app_data"
APP_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = APP_DIR / "templates.db"


def _conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_name TEXT NOT NULL,
            header_fingerprint TEXT,
            mode TEXT NOT NULL,
            preserve_unknown INTEGER NOT NULL,
            mapping_json TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    return conn


def save_template(payload: Dict[str, Any]) -> None:
    supplier_name = payload["supplier_name"].strip()
    header_fingerprint = payload.get("header_fingerprint", "")
    mode = payload.get("mode", "normalized")
    preserve_unknown = 1 if payload.get("preserve_unknown", True) else 0
    mapping_json = json.dumps(payload.get("mapping", {}), ensure_ascii=False, sort_keys=True)

    with _conn() as conn:
        existing = conn.execute(
            "SELECT id FROM templates WHERE supplier_name = ? AND header_fingerprint = ?",
            (supplier_name, header_fingerprint),
        ).fetchone()
        if existing:
            conn.execute(
                """
                UPDATE templates
                SET mode = ?, preserve_unknown = ?, mapping_json = ?, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
                """,
                (mode, preserve_unknown, mapping_json, existing["id"]),
            )
        else:
            conn.execute(
                """
                INSERT INTO templates (supplier_name, header_fingerprint, mode, preserve_unknown, mapping_json)
                VALUES (?, ?, ?, ?, ?)
                """,
                (supplier_name, header_fingerprint, mode, preserve_unknown, mapping_json),
            )


def list_templates() -> List[Dict[str, Any]]:
    with _conn() as conn:
        rows = conn.execute(
            "SELECT supplier_name, header_fingerprint, mode, preserve_unknown, mapping_json, updated_at FROM templates ORDER BY updated_at DESC"
        ).fetchall()
    return [
        {
            "supplier_name": row["supplier_name"],
            "header_fingerprint": row["header_fingerprint"],
            "mode": row["mode"],
            "preserve_unknown": bool(row["preserve_unknown"]),
            "mapping": json.loads(row["mapping_json"]),
            "updated_at": row["updated_at"],
        }
        for row in rows
    ]


def find_template(supplier_name: str = "", header_fingerprint: str = "") -> Optional[Dict[str, Any]]:
    supplier_name = supplier_name.strip()
    with _conn() as conn:
        row = None
        if supplier_name and header_fingerprint:
            row = conn.execute(
                "SELECT supplier_name, header_fingerprint, mode, preserve_unknown, mapping_json, updated_at FROM templates WHERE supplier_name = ? AND header_fingerprint = ? ORDER BY updated_at DESC LIMIT 1",
                (supplier_name, header_fingerprint),
            ).fetchone()
        if row is None and supplier_name:
            row = conn.execute(
                "SELECT supplier_name, header_fingerprint, mode, preserve_unknown, mapping_json, updated_at FROM templates WHERE supplier_name = ? ORDER BY updated_at DESC LIMIT 1",
                (supplier_name,),
            ).fetchone()
        if row is None and header_fingerprint:
            row = conn.execute(
                "SELECT supplier_name, header_fingerprint, mode, preserve_unknown, mapping_json, updated_at FROM templates WHERE header_fingerprint = ? ORDER BY updated_at DESC LIMIT 1",
                (header_fingerprint,),
            ).fetchone()

    if row is None:
        return None
    return {
        "supplier_name": row["supplier_name"],
        "header_fingerprint": row["header_fingerprint"],
        "mode": row["mode"],
        "preserve_unknown": bool(row["preserve_unknown"]),
        "mapping": json.loads(row["mapping_json"]),
        "updated_at": row["updated_at"],
    }


def export_templates_json() -> str:
    return json.dumps(list_templates(), ensure_ascii=False, indent=2)


def import_templates_json(text: str) -> int:
    payload = json.loads(text)
    if not isinstance(payload, list):
        raise ValueError("Template import must be a JSON list.")
    count = 0
    for item in payload:
        if not isinstance(item, dict) or "supplier_name" not in item or "mapping" not in item:
            continue
        save_template(item)
        count += 1
    return count
