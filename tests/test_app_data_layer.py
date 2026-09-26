"""
Tests that the Streamlit app's data layer returns the same reports as nctrack.

The app loads CSVs into SQLite via app._build_db(), then computes reports
through nctrack — exactly what happens at startup. We verify that the reports
produced via that path are identical to those produced by calling nctrack directly.
"""

from __future__ import annotations

import sqlite3
import csv as _csv
from pathlib import Path
from typing import Any

import pytest

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"

# ---------------------------------------------------------------------------
# Import app helpers without triggering Streamlit bootstrap
# ---------------------------------------------------------------------------

import importlib
import sys

# Prevent Streamlit from running at import time
sys.modules.setdefault("streamlit", __import__("unittest.mock", fromlist=["MagicMock"]).MagicMock())

# Now safe to import the functions we want to test
from nctrack import alertes, mens_global, mens_lignes, pareto, rap_hebdo
from nctrack.config import LegacyConfig, CorrectedConfig
from nctrack.loader import load, Dataset


# ---------------------------------------------------------------------------
# Fixture: SQLite DB built the same way as app._build_db()
# ---------------------------------------------------------------------------


def _build_test_db(tmp_path: Path) -> sqlite3.Connection:
    """Replicate app._build_db() without importing the Streamlit module."""
    db_path = tmp_path / "nctrack_test.db"
    db = sqlite3.connect(str(db_path), check_same_thread=False)
    tables = [
        "defect_log",
        "production_log",
        "defect_types",
        "parts",
        "parameters",
    ]
    for table in tables:
        path = DATA_DIR / f"{table}.csv"
        with open(path, newline="", encoding="utf-8") as fh:
            reader = _csv.DictReader(fh)
            rows = list(reader)
        if not rows:
            continue
        cols = list(rows[0].keys())
        col_defs = ", ".join(f'"{c}" TEXT' for c in cols)
        db.execute(f'CREATE TABLE IF NOT EXISTS "{table}" ({col_defs})')
        placeholders = ", ".join("?" for _ in cols)
        db.executemany(
            f'INSERT INTO "{table}" VALUES ({placeholders})',
            [[r[c] for c in cols] for r in rows],
        )
    db.commit()
    return db


def _read_table(db: sqlite3.Connection, table: str) -> list[dict[str, str]]:
    """Read a full table from SQLite and return list of dicts."""
    cursor = db.execute(f'SELECT * FROM "{table}"')
    cols = [d[0] for d in cursor.description]
    return [dict(zip(cols, row)) for row in cursor.fetchall()]


def _dataset_from_db(db: sqlite3.Connection) -> Dataset:
    """Build a Dataset from the SQLite tables (mirrors loader.load())."""
    return Dataset(
        defect_log=_read_table(db, "defect_log"),
        production_log=_read_table(db, "production_log"),
        defect_types=_read_table(db, "defect_types"),
        parts=_read_table(db, "parts"),
        parameters=_read_table(db, "parameters"),
    )


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


@pytest.fixture(scope="module")
def db(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("db")
    return _build_test_db(tmp)


@pytest.fixture(scope="module")
def ds_direct():
    """Dataset loaded directly from CSVs (nctrack.loader.load)."""
    return load(DATA_DIR)


@pytest.fixture(scope="module")
def ds_from_db(db):
    """Dataset rebuilt from the SQLite DB that app._build_db() creates."""
    return _dataset_from_db(db)


# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------


def _norm(rows: list[dict[str, Any]]) -> list[dict[str, str]]:
    """Normalise a list of report rows to comparable string dicts."""
    return [{k: str(v) for k, v in row.items()} for row in rows]


# ---------------------------------------------------------------------------
# Legacy mode parity tests
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("mode", ["legacy", "corrected"])
def test_rap_hebdo_data_layer_parity(ds_direct, ds_from_db, mode):
    cfg = LegacyConfig() if mode == "legacy" else CorrectedConfig()
    direct = _norm(rap_hebdo.compute(ds_direct, cfg))
    via_db = _norm(rap_hebdo.compute(ds_from_db, cfg))
    assert direct == via_db, f"rap_hebdo differs in {mode} mode"


@pytest.mark.parametrize("mode", ["legacy", "corrected"])
def test_pareto_data_layer_parity(ds_direct, ds_from_db, mode):
    cfg = LegacyConfig() if mode == "legacy" else CorrectedConfig()
    direct = _norm(pareto.compute(ds_direct, cfg))
    via_db = _norm(pareto.compute(ds_from_db, cfg))
    assert direct == via_db, f"pareto differs in {mode} mode"


@pytest.mark.parametrize("mode", ["legacy", "corrected"])
def test_alertes_data_layer_parity(ds_direct, ds_from_db, mode):
    cfg = LegacyConfig() if mode == "legacy" else CorrectedConfig()
    direct = _norm(alertes.compute(ds_direct, cfg))
    via_db = _norm(alertes.compute(ds_from_db, cfg))
    assert direct == via_db, f"alertes differs in {mode} mode"


@pytest.mark.parametrize("mode", ["legacy", "corrected"])
def test_mens_lignes_data_layer_parity(ds_direct, ds_from_db, mode):
    cfg = LegacyConfig() if mode == "legacy" else CorrectedConfig()
    al_direct = alertes.compute(ds_direct, cfg)
    al_db = alertes.compute(ds_from_db, cfg)
    direct = _norm(mens_lignes.compute(ds_direct, cfg, al_direct))
    via_db = _norm(mens_lignes.compute(ds_from_db, cfg, al_db))
    assert direct == via_db, f"mens_lignes differs in {mode} mode"


@pytest.mark.parametrize("mode", ["legacy", "corrected"])
def test_mens_global_data_layer_parity(ds_direct, ds_from_db, mode):
    cfg = LegacyConfig() if mode == "legacy" else CorrectedConfig()
    al_direct = alertes.compute(ds_direct, cfg)
    al_db = alertes.compute(ds_from_db, cfg)
    direct = _norm(mens_global.compute(ds_direct, cfg, al_direct))
    via_db = _norm(mens_global.compute(ds_from_db, cfg, al_db))
    assert direct == via_db, f"mens_global differs in {mode} mode"


# ---------------------------------------------------------------------------
# SQLite table completeness
# ---------------------------------------------------------------------------


def test_db_tables_populated(db):
    for table in ["defect_log", "production_log", "defect_types", "parts", "parameters"]:
        rows = db.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0]
        assert rows > 0, f"Table {table} is empty in the test DB"
