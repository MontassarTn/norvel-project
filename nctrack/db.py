"""
db — SQLite storage for the NC Tracker input data.

The legacy workbook imported each CSV into a sheet; the modern version imports
them into a SQLite database instead and computes every report from it.

    from nctrack.db import build_database, load_database

    build_database("data", "nctrack.db")   # recreate the database from data/
    ds = load_database("nctrack.db")       # Dataset for the report modules

Values are stored as TEXT, exactly as they appear in the CSVs, so a Dataset
read from the database is identical to one read from the CSVs.
"""

from __future__ import annotations

import csv
import sqlite3
from pathlib import Path

from nctrack.loader import Dataset

TABLES = ("defect_log", "production_log", "defect_types", "parts", "parameters")


def build_database(data_dir: str | Path, db_path: str | Path) -> None:
    """(Re)create *db_path* with one table per input CSV in *data_dir*."""
    db_path = Path(db_path)
    db_path.unlink(missing_ok=True)
    with sqlite3.connect(db_path) as conn:
        for table in TABLES:
            with open(Path(data_dir) / f"{table}.csv", newline="", encoding="utf-8") as fh:
                reader = csv.reader(fh)
                header = next(reader)
                rows = list(reader)
            columns = ", ".join(f'"{c}" TEXT' for c in header)
            conn.execute(f'CREATE TABLE "{table}" ({columns})')
            placeholders = ", ".join("?" for _ in header)
            conn.executemany(f'INSERT INTO "{table}" VALUES ({placeholders})', rows)
    conn.close()


def load_database(db_path: str | Path) -> Dataset:
    """Read every input table from *db_path*, in import order, into a Dataset."""
    conn = sqlite3.connect(Path(db_path))
    conn.row_factory = sqlite3.Row
    try:
        tables = {
            table: [dict(r) for r in conn.execute(f'SELECT * FROM "{table}" ORDER BY rowid')]
            for table in TABLES
        }
    finally:
        conn.close()
    return Dataset(**tables)
