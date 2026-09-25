"""
mens_global — Monthly global summary (CalcMensuel → MENS_GLOBAL equivalent).

Legacy divergences reproduced when active in cfg:
  D1  L4 defect qty multiplied by d1_l4_factor (legacy: 0.5, corrected: 1.0)
  D4  ACCEPT excluded from monthly nd accumulator
  D6  records on zero-production days silently skipped
  D7  thresholds hardcoded instead of read from parameters.csv
  D9  CNQ rounded per record (legacy CLng) vs rounding only the final total

Output columns:
  mois,produit,defauts,cnq_eur,top3,nb_critiques,nb_recurrences
"""

from __future__ import annotations

import datetime
from collections import defaultdict
from typing import Any

from nctrack.config import Config
from nctrack.loader import Dataset


def _month_label(d: datetime.date) -> str:
    return f"{d.year}-{d.month:02d}"


def _cost_nc_exact(
    row: dict[str, str],
    ds: Dataset,
    labour_rate: float,
) -> float:
    """Exact (unrounded) cost of one defect record."""
    qty = float(row["qty"])
    code = row["defect_code"]
    disposition = row["disposition"]
    part_ref = row["part_ref"]
    part_cost = ds.parts_map.get(part_ref, 0.0)
    defect_info = ds.defect_map.get(code, {})
    if disposition == "SCRAP":
        return qty * part_cost
    if disposition == "REWORK":
        if code == "D06":
            return qty * part_cost
        hours = defect_info.get("rework_hours") or 0.0
        return qty * hours * labour_rate
    return 0.0


def _cost_nc_rounded(
    row: dict[str, str],
    ds: Dataset,
    labour_rate: float,
) -> int:
    """Per-record rounded cost (VBA CLng behaviour, D9 on)."""
    return round(_cost_nc_exact(row, ds, labour_rate))


def compute(
    ds: Dataset,
    cfg: Config,
    alert_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    Compute monthly global summary rows.

    Returns list of dicts with keys:
      mois, produit, defauts, cnq_eur, top3, nb_critiques, nb_recurrences
    """
    # --- thresholds ---
    if cfg.d7_hardcoded_params:
        labour_rate = 45.0
    else:
        labour_rate = float(ds.params_map["rework_rate_eur_per_hour"])

    prod: dict[str, float] = defaultdict(float)
    nd: dict[str, float] = defaultdict(float)
    cq: dict[str, float] = defaultdict(float)
    # top3: qty per (month, defect_code) — all dispositions per VBA (tq array)
    tq: dict[tuple[str, str], float] = defaultdict(float)

    # Pass 1 — production
    for row in ds.production_log:
        d = datetime.date.fromisoformat(row["date"])
        m = _month_label(d)
        prod[m] += float(row["qty_produced"])

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
        code = row["defect_code"]
        disposition = row["disposition"]

        # D4: ACCEPT excluded from monthly nd
        if not (cfg.d4_accept_excluded_monthly and disposition == "ACCEPT"):
            if cfg.d1_l4_halving and line == "L4":
                nd[m] += qty * cfg.d1_l4_factor
            else:
                nd[m] += qty

        # D9: per-record rounding (legacy) vs exact accumulation (corrected)
        if cfg.d9_round_per_record:
            cq[m] += _cost_nc_rounded(row, ds, labour_rate)
        else:
            cq[m] += _cost_nc_exact(row, ds, labour_rate)

        # tq: all dispositions contribute to top3 ranking
        tq[(m, code)] += qty

    # Pass 3 — alert counts
    crit_count: dict[str, int] = defaultdict(int)
    recur_count: dict[str, int] = defaultdict(int)
    for alert in alert_rows:
        m = _month_label(datetime.date.fromisoformat(alert["date"]))
        if alert["type"] == "CRITIQUE":
            crit_count[m] += 1
        elif alert["type"] == "RECURRENCE":
            recur_count[m] += 1

    # Build output rows
    rows: list[dict[str, Any]] = []
    for m in sorted(prod):
        p = prod[m]
        if p <= 0:
            continue

        # top3 by quantity (descending), tie-break by code
        month_codes = {
            code: qty
            for (mo, code), qty in tq.items()
            if mo == m
        }
        ranked = sorted(month_codes.items(), key=lambda x: (-x[1], x[0]))
        top3_codes = [c for c, _ in ranked[:3]]
        top3 = " / ".join(top3_codes)

        n = nd[m]
        rows.append(
            {
                "mois": m,
                "produit": int(p),
                "defauts": n,
                "cnq_eur": round(cq[m]),  # D9: round the final total
                "top3": top3,
                "nb_critiques": crit_count.get(m, 0),
                "nb_recurrences": recur_count.get(m, 0),
            }
        )
    return rows


# ---------------------------------------------------------------------------
# CSV writer
# ---------------------------------------------------------------------------

COLUMNS = [
    "mois", "produit", "defauts", "cnq_eur", "top3",
    "nb_critiques", "nb_recurrences",
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
            str(row["produit"]),
            _fmt_defauts(row["defauts"]),
            str(int(row["cnq_eur"])),
            row["top3"],
            str(row["nb_critiques"]),
            str(row["nb_recurrences"]),
        ]
        lines_out.append(",".join(parts))
    return "\n".join(lines_out) + "\n"
