"""
pareto — Weekly Pareto ranking (CalcPareto equivalent).

Legacy divergences reproduced when active in cfg:
  D5  Week anchor = preceding Sunday (Sun-to-Sat buckets)
      vs ISO Monday-to-Sunday

Output columns:
  semaine_du,rang,code,libelle,qte,pct,cumul_pct,prioritaire

The week label is always the Sunday that starts (legacy) or ends (corrected)
the bucket — the legacy output uses the Sunday of the Sun-to-Sat week.
"""

from __future__ import annotations

import datetime
from collections import defaultdict
from typing import Any

from nctrack.config import Config
from nctrack.loader import Dataset

# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

_ORIGIN_LEGACY = datetime.date(2025, 12, 28)  # Sunday before ISO W01 2026


def _sunday_of_week(d: datetime.date) -> datetime.date:
    """Return the preceding Sunday (VBA logic: roll back to vbSunday)."""
    # weekday(): Mon=0 … Sun=6
    # vbSunday roll-back: subtract (d.weekday()+1) % 7 days
    offset = (d.weekday() + 1) % 7  # Sun→0, Mon→1, …, Sat→6
    return d - datetime.timedelta(days=offset)


def _iso_week_sunday(d: datetime.date) -> datetime.date:
    """Return the Sunday of the ISO week containing *d* (Mon–Sun)."""
    # ISO weekday: Mon=1 … Sun=7
    iso_wd = d.isoweekday()  # Mon=1, Sun=7
    # days to add to reach Sunday
    days_to_sunday = 7 - iso_wd
    return d + datetime.timedelta(days=days_to_sunday)


def _week_key(d: datetime.date, sunday_anchor: bool) -> datetime.date:
    """Return the Sunday date used as bucket key."""
    if sunday_anchor:
        return _sunday_of_week(d)
    else:
        return _iso_week_sunday(d)


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------


def compute(ds: Dataset, cfg: Config) -> list[dict[str, Any]]:
    """
    Compute Pareto rows.

    Returns list of dicts with keys:
      semaine_du, rang, code, libelle, qte, pct, cumul_pct, prioritaire
    """
    # Accumulate qty per (week_sunday, defect_code)
    week_code_qty: dict[tuple[datetime.date, str], float] = defaultdict(float)

    for row in ds.defect_log:
        d = datetime.date.fromisoformat(row["date"])
        wk = _week_key(d, cfg.d5_sunday_anchor)
        code = row["defect_code"]
        week_code_qty[(wk, code)] += float(row["qty"])

    # Group by week
    weeks: dict[datetime.date, dict[str, float]] = defaultdict(dict)
    for (wk, code), qty in week_code_qty.items():
        weeks[wk][code] = qty

    rows: list[dict[str, Any]] = []

    for wk in sorted(weeks):
        code_qty = {c: q for c, q in weeks[wk].items() if q > 0}
        if not code_qty:
            continue

        total = sum(code_qty.values())
        # Sort: descending qty; tie-break by code (D01 < D02 < …)
        ranked = sorted(code_qty.items(), key=lambda x: (-x[1], x[0]))

        cumul_raw = 0.0
        for rank, (code, qty) in enumerate(ranked, start=1):
            pct_raw = qty / total * 100
            pct = round(pct_raw, 2)
            # Legacy prioritaire logic: mark "X" if cumul_raw < 80 BEFORE adding
            prioritaire = "X" if cumul_raw < 80.0 else ""
            cumul_raw += pct_raw
            cumul = round(cumul_raw, 2)
            label = ds.defect_map.get(code, {}).get("label", "")
            rows.append(
                {
                    "semaine_du": wk.isoformat(),
                    "rang": rank,
                    "code": code,
                    "libelle": label,
                    "qte": int(qty),
                    "pct": pct,
                    "cumul_pct": cumul,
                    "prioritaire": prioritaire,
                }
            )
    return rows


# ---------------------------------------------------------------------------
# CSV writer
# ---------------------------------------------------------------------------

COLUMNS = [
    "semaine_du", "rang", "code", "libelle", "qte",
    "pct", "cumul_pct", "prioritaire",
]


def to_csv(rows: list[dict[str, Any]]) -> str:
    lines = [",".join(COLUMNS)]
    for row in rows:
        parts = [
            row["semaine_du"],
            str(row["rang"]),
            row["code"],
            row["libelle"],
            str(row["qte"]),
            f"{row['pct']:.2f}",
            f"{row['cumul_pct']:.2f}",
            row["prioritaire"],
        ]
        lines.append(",".join(parts))
    return "\n".join(lines) + "\n"
