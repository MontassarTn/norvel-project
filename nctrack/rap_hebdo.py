"""
rap_hebdo — Weekly report (CalcHebdo equivalent).

Legacy divergences reproduced when active in cfg:
  D1  L4 defect qty halved in nd accumulator
  D2  tx_rebut >= 3 → ROUGE  (vs strict > 3)
  D6  records on zero-production days silently skipped

Output columns:
  semaine,ligne,produit,nb_def,tx_def,rebut,tx_rebut,statut,cnq_eur
"""

from __future__ import annotations

import datetime
import math
from collections import defaultdict
from typing import Any

from nctrack.config import Config
from nctrack.loader import Dataset


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _iso_week(date: datetime.date) -> int:
    """Return ISO week number (1–53)."""
    return date.isocalendar()[1]


def _iso_year(date: datetime.date) -> int:
    return date.isocalendar()[0]


def _week_label(date: datetime.date) -> str:
    """Format week as '2026-W23'."""
    y, w, _ = date.isocalendar()
    return f"{y}-W{w:02d}"


def _cost_nc(
    row: dict[str, str],
    ds: Dataset,
    labour_rate: float,
) -> int:
    """Cost of one defect record (CoutNC equivalent).

    The VBA accumulates costs as Long integers, converting each CoutNC result
    via CLng() (banker's rounding) before adding to the accumulator.
    Python round() uses the same banker's rounding convention.
    """
    qty = float(row["qty"])
    code = row["defect_code"]
    disposition = row["disposition"]
    part_ref = row["part_ref"]

    part_cost = ds.parts_map.get(part_ref, 0.0)
    defect_info = ds.defect_map.get(code, {})

    if disposition == "SCRAP":
        return round(qty * part_cost)
    if disposition == "REWORK":
        if code == "D06":
            # delamination rework is costed as scrap (VBA + rule agree)
            return round(qty * part_cost)
        hours = defect_info.get("rework_hours") or 0.0
        return round(qty * hours * labour_rate)
    # ACCEPT → 0
    return 0


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------


def compute(ds: Dataset, cfg: Config) -> list[dict[str, Any]]:
    """
    Compute the weekly report rows.

    Returns a list of dicts with keys:
      semaine, ligne, produit, nb_def, tx_def, rebut, tx_rebut, statut, cnq_eur
    """
    # --- thresholds ---
    if cfg.d7_hardcoded_params:
        labour_rate = 45.0
        red_threshold = 3.0
        orange_threshold = 2.0
    else:
        labour_rate = float(ds.params_map["rework_rate_eur_per_hour"])
        red_threshold = float(ds.params_map["scrap_threshold_red_pct"])
        orange_threshold = float(ds.params_map["scrap_threshold_orange_pct"])

    # --- accumulators: (week_label, line) → values ---
    prod: dict[tuple[str, str], float] = defaultdict(float)
    nd: dict[tuple[str, str], float] = defaultdict(float)   # defect qty
    rb: dict[tuple[str, str], float] = defaultdict(float)   # scrap qty
    cq: dict[tuple[str, str], float] = defaultdict(float)   # cnq

    # Pass 1 — production
    for row in ds.production_log:
        d = datetime.date.fromisoformat(row["date"])
        wlabel = _week_label(d)
        line = row["line"]
        prod[(wlabel, line)] += float(row["qty_produced"])

    # Pass 2 — defects
    for row in ds.defect_log:
        d = datetime.date.fromisoformat(row["date"])
        line = row["line"]

        # D6: skip records on zero-production days
        if cfg.d6_skip_zero_prod:
            day_prod = ds.prod_day.get((row["date"], line), 0.0)
            if day_prod == 0.0:
                continue

        wlabel = _week_label(d)
        qty = float(row["qty"])
        key = (wlabel, line)

        # D1: L4 defect qty halved
        if cfg.d1_l4_halving and line == "L4":
            nd[key] += qty / 2.0
        else:
            nd[key] += qty

        if row["disposition"] == "SCRAP":
            rb[key] += qty  # scrap always uses full qty

        cq[key] += _cost_nc(row, ds, labour_rate)

    # Pass 3 — build output rows, sorted by week then line
    rows: list[dict[str, Any]] = []
    for key in sorted(prod):
        wlabel, line = key
        p = prod[key]
        if p <= 0:
            continue
        n = nd[key]
        r = rb[key]
        tx_def = round(n / p * 100, 2)
        tx_rebut = round(r / p * 100, 2)

        # D2: >= vs >
        if cfg.d2_rouge_gte:
            if tx_rebut >= red_threshold:
                statut = "ROUGE"
            elif tx_rebut > orange_threshold:
                statut = "ORANGE"
            else:
                statut = "VERT"
        else:
            if tx_rebut > red_threshold:
                statut = "ROUGE"
            elif tx_rebut > orange_threshold:
                statut = "ORANGE"
            else:
                statut = "VERT"

        rows.append(
            {
                "semaine": wlabel,
                "ligne": line,
                "produit": int(p),
                "nb_def": n,
                "tx_def": tx_def,
                "rebut": int(r),
                "tx_rebut": tx_rebut,
                "statut": statut,
                "cnq_eur": cq[key],
            }
        )
    return rows


# ---------------------------------------------------------------------------
# CSV writer
# ---------------------------------------------------------------------------

COLUMNS = ["semaine", "ligne", "produit", "nb_def", "tx_def",
           "rebut", "tx_rebut", "statut", "cnq_eur"]


def _fmt_field(key: str, value: Any) -> str:
    """Format a single field the same way the VBA ExportSorties does."""
    if key in ("nb_def", "tx_def", "tx_rebut"):
        # Numbers: strip trailing .0 only when integer, keep decimals otherwise
        # VBA Txt() / Fmt2() uses the raw float repr without trailing zeros for
        # nb_def (e.g. 4.5 stays 4.5, 14 stays 14), and 2dp for tx_ fields.
        if key == "nb_def":
            # Use minimal repr: 4.5 → "4.5", 14.0 → "14"
            if value == int(value):
                return str(int(value))
            return str(value)
        else:
            # tx_ fields: always 2 decimal places
            return f"{value:.2f}"
    if key == "cnq_eur":
        return str(int(value))
    return str(value)


def to_csv(rows: list[dict[str, Any]]) -> str:
    """Render rows to CSV string (no trailing newline after last row)."""
    lines = [",".join(COLUMNS)]
    for row in rows:
        lines.append(",".join(_fmt_field(k, row[k]) for k in COLUMNS))
    return "\n".join(lines) + "\n"
