"""
mens_lignes — Monthly per-line report (CalcMensuel → MENS_LIGNES equivalent).

Legacy divergences reproduced when active in cfg:
  D1  L4 defect qty halved in nd accumulator
  D4  ACCEPT excluded from monthly nd accumulator
  D6  records on zero-production days silently skipped

Output columns:
  mois,ligne,produit,defauts,tx_def,rebut,tx_rebut,statut,cnq_eur,tendance
"""

from __future__ import annotations

import datetime
import math
from collections import defaultdict
from typing import Any

from nctrack.config import Config
from nctrack.loader import Dataset


def _month_label(d: datetime.date) -> str:
    return f"{d.year}-{d.month:02d}"


def _cost_nc(
    row: dict[str, str],
    ds: Dataset,
    labour_rate: float,
) -> int:
    """Cost of one defect record (per-record banker's rounding, matching VBA CLng)."""
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
            return round(qty * part_cost)
        hours = defect_info.get("rework_hours") or 0.0
        return round(qty * hours * labour_rate)
    return 0


def compute(
    ds: Dataset,
    cfg: Config,
    alert_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    Compute monthly per-line rows.

    Parameters
    ----------
    ds          : loaded dataset
    cfg         : configuration flags
    alert_rows  : output of alertes.compute() (needed for alert counts)

    Returns list of dicts with keys:
      mois, ligne, produit, defauts, tx_def, rebut, tx_rebut,
      statut, cnq_eur, tendance
    """
    # --- thresholds ---
    if cfg.d7_hardcoded_params:
        labour_rate = 45.0
        red_threshold = 3.0
        orange_threshold = 2.0
        trend_band = 0.5
    else:
        labour_rate = float(ds.params_map["rework_rate_eur_per_hour"])
        red_threshold = float(ds.params_map["scrap_threshold_red_pct"])
        orange_threshold = float(ds.params_map["scrap_threshold_orange_pct"])
        trend_band = float(ds.params_map["trend_stable_band_pts"])

    prod: dict[tuple[str, str], float] = defaultdict(float)
    nd: dict[tuple[str, str], float] = defaultdict(float)
    rb: dict[tuple[str, str], float] = defaultdict(float)
    cq: dict[tuple[str, str], float] = defaultdict(float)

    # Pass 1 — production
    for row in ds.production_log:
        d = datetime.date.fromisoformat(row["date"])
        m = _month_label(d)
        line = row["line"]
        prod[(m, line)] += float(row["qty_produced"])

    # Pass 2 — defects
    for row in ds.defect_log:
        d = datetime.date.fromisoformat(row["date"])
        line = row["line"]

        # D6: skip records on zero-production days
        if cfg.d6_skip_zero_prod:
            day_prod = ds.prod_day.get((row["date"], line), 0.0)
            if day_prod == 0.0:
                continue

        m = _month_label(d)
        qty = float(row["qty"])
        key = (m, line)
        disposition = row["disposition"]

        # D4: ACCEPT excluded from monthly nd
        if not (cfg.d4_accept_excluded_monthly and disposition == "ACCEPT"):
            # D1: L4 defect qty halved
            if cfg.d1_l4_halving and line == "L4":
                nd[key] += qty / 2.0
            else:
                nd[key] += qty

        if disposition == "SCRAP":
            rb[key] += qty

        cq[key] += _cost_nc(row, ds, labour_rate)

    # Collect ordered months and lines
    all_months = sorted({m for m, _ in prod})
    lines_order = ["L1", "L2", "L3", "L4"]

    # Previous month tx_def per line (for trend)
    prev_tx_def: dict[str, float | None] = {l: None for l in lines_order}

    rows: list[dict[str, Any]] = []

    for m in all_months:
        for line in lines_order:
            key = (m, line)
            p = prod.get(key, 0.0)
            if p <= 0:
                continue
            n = nd.get(key, 0.0)
            r = rb.get(key, 0.0)
            tx_def = round(n / p * 100, 2)
            tx_rebut = round(r / p * 100, 2)

            # Status: VBA CalcMensuel uses fr = rb/p (fraction, not percent)
            # then compares fr > 0.03 and fr*100 > 2
            fr = r / p
            if fr > 0.03:
                statut = "ROUGE"
            elif fr * 100 > orange_threshold:
                statut = "ORANGE"
            else:
                statut = "VERT"

            # Trend
            prev = prev_tx_def.get(line)
            if prev is None:
                tendance = "N/A"
            else:
                delta = tx_def - prev
                if delta > trend_band:
                    tendance = "HAUSSE"
                elif delta < -trend_band:
                    tendance = "BAISSE"
                else:
                    tendance = "STABLE"

            prev_tx_def[line] = tx_def

            rows.append(
                {
                    "mois": m,
                    "ligne": line,
                    "produit": int(p),
                    "defauts": n,
                    "tx_def": tx_def,
                    "rebut": int(r),
                    "tx_rebut": tx_rebut,
                    "statut": statut,
                    "cnq_eur": cq.get(key, 0),
                    "tendance": tendance,
                }
            )
    return rows


# ---------------------------------------------------------------------------
# CSV writer
# ---------------------------------------------------------------------------

COLUMNS = [
    "mois", "ligne", "produit", "defauts", "tx_def",
    "rebut", "tx_rebut", "statut", "cnq_eur", "tendance",
]


def _fmt_defauts(value: float) -> str:
    if value == int(value):
        return str(int(value))
    return str(value)


def to_csv(rows: list[dict[str, Any]]) -> str:
    lines_out = [",".join(COLUMNS)]
    for row in rows:
        parts = [
            row["mois"],
            row["ligne"],
            str(row["produit"]),
            _fmt_defauts(row["defauts"]),
            f"{row['tx_def']:.2f}",
            str(row["rebut"]),
            f"{row['tx_rebut']:.2f}",
            row["statut"],
            str(int(row["cnq_eur"])),
            row["tendance"],
        ]
        lines_out.append(",".join(parts))
    return "\n".join(lines_out) + "\n"
