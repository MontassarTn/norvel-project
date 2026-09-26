"""
Tests for the Streamlit app's data layer and pages.

- The app imports data/ into SQLite (nctrack.db.build_database) and computes
  every report from the database (nctrack.db.load_database).  The reports must
  be identical to those computed straight from the CSVs, in both modes.
- The app itself is run headless with Streamlit's AppTest: every page must
  render without an exception in both modes.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from nctrack import alertes, mens_global, mens_lignes, pareto, rap_hebdo
from nctrack.config import CorrectedConfig, LegacyConfig
from nctrack.db import build_database, load_database
from nctrack.loader import load

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"


def _all_reports(ds, cfg):
    al_rows = alertes.compute(ds, cfg)
    return {
        "rap_hebdo": rap_hebdo.compute(ds, cfg),
        "pareto": pareto.compute(ds, cfg),
        "alertes": al_rows,
        "mens_lignes": mens_lignes.compute(ds, cfg, al_rows),
        "mens_global": mens_global.compute(ds, cfg, al_rows),
    }


@pytest.fixture(scope="module")
def db_dataset(tmp_path_factory):
    db_path = tmp_path_factory.mktemp("db") / "nctrack.db"
    build_database(DATA_DIR, db_path)
    return load_database(db_path)


def test_database_holds_every_input_row(db_dataset):
    csv_ds = load(DATA_DIR)
    for table in ("defect_log", "production_log", "defect_types", "parts", "parameters"):
        assert getattr(db_dataset, table) == getattr(csv_ds, table), table


@pytest.mark.parametrize("make_cfg", [LegacyConfig, CorrectedConfig], ids=["legacy", "corrected"])
def test_reports_from_database_match_csv(db_dataset, make_cfg):
    cfg = make_cfg()
    assert _all_reports(db_dataset, cfg) == _all_reports(load(DATA_DIR), cfg)


PAGES = ["weekly", "pareto", "alerts", "monthly", "what_changed"]


@pytest.mark.parametrize("mode", ["legacy", "corrected"])
@pytest.mark.parametrize("page", PAGES)
def test_app_page_renders(page, mode):
    from streamlit.testing.v1 import AppTest

    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=60)
    at.run()
    at.sidebar.radio[0].set_value(mode)
    at.sidebar.radio[1].set_value(page)
    at.run()
    assert not at.exception, [e.message for e in at.exception]
    assert at.header, "page rendered no header"
