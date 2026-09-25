"""
Characterization tests for nctrack v1-parity (legacy mode).

Each test runs one report against the matching golden file in sorties_legacy/
and compares field-by-field, ignoring only line endings.

Outputs are written to a temporary folder; sorties_legacy/ is NEVER touched.
"""

from __future__ import annotations

import csv
import io
from pathlib import Path

import pytest

from nctrack import alertes, mens_global, mens_lignes, pareto, rap_hebdo
from nctrack.config import LegacyConfig
from nctrack.loader import load

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
GOLDEN_DIR = ROOT / "sorties_legacy"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _parse_csv(text: str) -> list[dict[str, str]]:
    """Parse CSV text into list of dicts, stripping whitespace from values."""
    reader = csv.DictReader(io.StringIO(text.replace("\r\n", "\n").replace("\r", "\n")))
    return [dict(row) for row in reader]


def _golden(name: str) -> list[dict[str, str]]:
    return _parse_csv((GOLDEN_DIR / name).read_text(encoding="utf-8"))


def _computed(text: str) -> list[dict[str, str]]:
    return _parse_csv(text)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


@pytest.fixture(scope="session")
def ds():
    return load(DATA_DIR)


@pytest.fixture(scope="session")
def cfg():
    return LegacyConfig()


@pytest.fixture(scope="session")
def al_rows(ds, cfg):
    return alertes.compute(ds, cfg)


# ---------------------------------------------------------------------------
# rap_hebdo
# ---------------------------------------------------------------------------


def test_rap_hebdo_row_count(ds, cfg):
    rows = rap_hebdo.compute(ds, cfg)
    golden = _golden("rap_hebdo.csv")
    assert len(rows) == len(golden), (
        f"rap_hebdo: got {len(rows)} rows, expected {len(golden)}"
    )


def test_rap_hebdo_fields(ds, cfg, tmp_path):
    rows = rap_hebdo.compute(ds, cfg)
    csv_text = rap_hebdo.to_csv(rows)
    (tmp_path / "rap_hebdo.csv").write_text(csv_text, encoding="utf-8")

    computed = _computed(csv_text)
    golden = _golden("rap_hebdo.csv")

    assert len(computed) == len(golden)
    for i, (got, exp) in enumerate(zip(computed, golden)):
        for field in exp:
            assert got[field] == exp[field], (
                f"rap_hebdo row {i+1} field '{field}': "
                f"got {got[field]!r}, expected {exp[field]!r}"
            )


# ---------------------------------------------------------------------------
# pareto
# ---------------------------------------------------------------------------


def test_pareto_row_count(ds, cfg):
    rows = pareto.compute(ds, cfg)
    golden = _golden("pareto.csv")
    assert len(rows) == len(golden), (
        f"pareto: got {len(rows)} rows, expected {len(golden)}"
    )


def test_pareto_fields(ds, cfg, tmp_path):
    rows = pareto.compute(ds, cfg)
    csv_text = pareto.to_csv(rows)
    (tmp_path / "pareto.csv").write_text(csv_text, encoding="utf-8")

    computed = _computed(csv_text)
    golden = _golden("pareto.csv")

    assert len(computed) == len(golden)
    for i, (got, exp) in enumerate(zip(computed, golden)):
        for field in exp:
            assert got[field] == exp[field], (
                f"pareto row {i+1} field '{field}': "
                f"got {got[field]!r}, expected {exp[field]!r}"
            )


# ---------------------------------------------------------------------------
# alertes
# ---------------------------------------------------------------------------


def test_alertes_row_count(ds, cfg):
    rows = alertes.compute(ds, cfg)
    golden = _golden("alertes.csv")
    assert len(rows) == len(golden), (
        f"alertes: got {len(rows)} rows, expected {len(golden)}"
    )


def test_alertes_fields(ds, cfg, tmp_path):
    rows = alertes.compute(ds, cfg)
    csv_text = alertes.to_csv(rows)
    (tmp_path / "alertes.csv").write_text(csv_text, encoding="utf-8")

    computed = _computed(csv_text)
    golden = _golden("alertes.csv")

    assert len(computed) == len(golden)
    for i, (got, exp) in enumerate(zip(computed, golden)):
        for field in exp:
            assert got[field] == exp[field], (
                f"alertes row {i+1} field '{field}': "
                f"got {got[field]!r}, expected {exp[field]!r}"
            )


# ---------------------------------------------------------------------------
# mens_lignes
# ---------------------------------------------------------------------------


def test_mens_lignes_row_count(ds, cfg, al_rows):
    rows = mens_lignes.compute(ds, cfg, al_rows)
    golden = _golden("mens_lignes.csv")
    assert len(rows) == len(golden), (
        f"mens_lignes: got {len(rows)} rows, expected {len(golden)}"
    )


def test_mens_lignes_fields(ds, cfg, al_rows, tmp_path):
    rows = mens_lignes.compute(ds, cfg, al_rows)
    csv_text = mens_lignes.to_csv(rows)
    (tmp_path / "mens_lignes.csv").write_text(csv_text, encoding="utf-8")

    computed = _computed(csv_text)
    golden = _golden("mens_lignes.csv")

    assert len(computed) == len(golden)
    for i, (got, exp) in enumerate(zip(computed, golden)):
        for field in exp:
            assert got[field] == exp[field], (
                f"mens_lignes row {i+1} field '{field}': "
                f"got {got[field]!r}, expected {exp[field]!r}"
            )


# ---------------------------------------------------------------------------
# mens_global
# ---------------------------------------------------------------------------


def test_mens_global_row_count(ds, cfg, al_rows):
    rows = mens_global.compute(ds, cfg, al_rows)
    golden = _golden("mens_global.csv")
    assert len(rows) == len(golden), (
        f"mens_global: got {len(rows)} rows, expected {len(golden)}"
    )


def test_mens_global_fields(ds, cfg, al_rows, tmp_path):
    rows = mens_global.compute(ds, cfg, al_rows)
    csv_text = mens_global.to_csv(rows)
    (tmp_path / "mens_global.csv").write_text(csv_text, encoding="utf-8")

    computed = _computed(csv_text)
    golden = _golden("mens_global.csv")

    assert len(computed) == len(golden)
    for i, (got, exp) in enumerate(zip(computed, golden)):
        for field in exp:
            assert got[field] == exp[field], (
                f"mens_global row {i+1} field '{field}': "
                f"got {got[field]!r}, expected {exp[field]!r}"
            )
