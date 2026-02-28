from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from typing import Any, Dict, List, Optional

APP_DIR = Path(__file__).resolve().parent / 'app_data'
APP_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = APP_DIR / 'profiles.db'


def _conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute(
        '''
        CREATE TABLE IF NOT EXISTS supplier_profiles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_name TEXT NOT NULL UNIQUE,
            notes TEXT DEFAULT '',
            custom_aliases_json TEXT DEFAULT '{}',
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        '''
    )
    return conn


def list_profiles() -> List[Dict[str, Any]]:
    with _conn() as conn:
        rows = conn.execute(
            'SELECT supplier_name, notes, custom_aliases_json, updated_at FROM supplier_profiles ORDER BY updated_at DESC'
        ).fetchall()
    return [
        {
            'supplier_name': row['supplier_name'],
            'notes': row['notes'] or '',
            'custom_aliases': json.loads(row['custom_aliases_json'] or '{}'),
            'updated_at': row['updated_at'],
        }
        for row in rows
    ]


def get_profile(supplier_name: str) -> Optional[Dict[str, Any]]:
    supplier_name = supplier_name.strip()
    if not supplier_name:
        return None
    with _conn() as conn:
        row = conn.execute(
            'SELECT supplier_name, notes, custom_aliases_json, updated_at FROM supplier_profiles WHERE supplier_name = ?',
            (supplier_name,),
        ).fetchone()
    if row is None:
        return None
    return {
        'supplier_name': row['supplier_name'],
        'notes': row['notes'] or '',
        'custom_aliases': json.loads(row['custom_aliases_json'] or '{}'),
        'updated_at': row['updated_at'],
    }


def save_profile(supplier_name: str, notes: str = '', custom_aliases: Optional[Dict[str, str]] = None) -> None:
    supplier_name = supplier_name.strip()
    if not supplier_name:
        raise ValueError('Supplier name is required.')
    alias_json = json.dumps(custom_aliases or {}, ensure_ascii=False, sort_keys=True)
    with _conn() as conn:
        existing = conn.execute('SELECT id FROM supplier_profiles WHERE supplier_name = ?', (supplier_name,)).fetchone()
        if existing:
            conn.execute(
                '''
                UPDATE supplier_profiles
                SET notes = ?, custom_aliases_json = ?, updated_at = CURRENT_TIMESTAMP
                WHERE supplier_name = ?
                ''',
                (notes or '', alias_json, supplier_name),
            )
        else:
            conn.execute(
                '''
                INSERT INTO supplier_profiles (supplier_name, notes, custom_aliases_json)
                VALUES (?, ?, ?)
                ''',
                (supplier_name, notes or '', alias_json),
            )


def export_profiles_json() -> str:
    return json.dumps(list_profiles(), ensure_ascii=False, indent=2)


def import_profiles_json(text: str) -> int:
    payload = json.loads(text)
    if not isinstance(payload, list):
        raise ValueError('Imported profile backup must be a JSON list.')
    count = 0
    for item in payload:
        if not isinstance(item, dict) or 'supplier_name' not in item:
            continue
        save_profile(
            supplier_name=item.get('supplier_name', ''),
            notes=item.get('notes', ''),
            custom_aliases=item.get('custom_aliases', {}),
        )
        count += 1
    return count
