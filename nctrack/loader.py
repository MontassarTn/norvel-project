"""
Data loader for nctrack.

Reads all input CSVs from a *data_dir* and returns plain Python
data-structures (dicts / lists).  No pandas — keeps dependencies minimal
and mirrors the VBA's row-by-row processing style.
"""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Any


def _read_csv(path: Path) -> list[dict[str, str]]:
    """Read a CSV file and return a list of dicts (header as keys)."""
    with open(path, newline="", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        return [row for row in reader]


class Dataset:
    """Holds the five input tables as lists of dicts."""

    def __init__(
        self,
        defect_log: list[dict[str, str]],
        production_log: list[dict[str, str]],
        defect_types: list[dict[str, str]],
        parts: list[dict[str, str]],
        parameters: list[dict[str, str]],
    ) -> None:
        self.defect_log = defect_log
        self.production_log = production_log
        self.defect_types = defect_types
        self.parts = parts
        self.parameters = parameters

        # --- convenience look-ups built once ---

        # defect_types: code → {label, base_severity, rework_hours_per_part}
        self.defect_map: dict[str, dict[str, Any]] = {}
        for row in defect_types:
            hours_str = row["rework_hours_per_part"].strip()
            self.defect_map[row["code"]] = {
                "label": row["label"],
                "base_severity": row["base_severity"],
                "rework_hours": float(hours_str) if hours_str else None,
            }

        # parts: part_ref → unit_cost_eur (float)
        self.parts_map: dict[str, float] = {
            r["part_ref"]: float(r["unit_cost_eur"]) for r in parts
        }

        # parameters: parameter → value (str)
        self.params_map: dict[str, str] = {
            r["parameter"]: r["value"] for r in parameters
        }

        # production look-up: (date_str, line) → qty_produced (float)
        self.prod_day: dict[tuple[str, str], float] = {}
        for r in production_log:
            self.prod_day[(r["date"], r["line"])] = float(r["qty_produced"])


def load(data_dir: str | Path) -> Dataset:
    """Load all input CSVs from *data_dir* and return a :class:`Dataset`."""
    p = Path(data_dir)
    return Dataset(
        defect_log=_read_csv(p / "defect_log.csv"),
        production_log=_read_csv(p / "production_log.csv"),
        defect_types=_read_csv(p / "defect_types.csv"),
        parts=_read_csv(p / "parts.csv"),
        parameters=_read_csv(p / "parameters.csv"),
    )
