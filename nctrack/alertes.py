"""
alertes — Alert computation (CalcAlertes equivalent).

Legacy divergences reproduced when active in cfg:
  D7  thresholds hardcoded instead of read from parameters.csv
  D8  recurrence forward-scan (per-anchor, may duplicate on dense data)
      vs proper sliding-window (one flag per qualifying window, per R6)

Output columns:
  type,id,date,ligne,piece,code,qte,message
"""

from __future__ import annotations

import datetime
from collections import defaultdict
from typing import Any

from nctrack.config import Config
from nctrack.loader import Dataset

_MSG_CRITIQUE = "defaut critique - prevenir resp. qualite"
_MSG_RECURRENCE = "recurrence 7j - ouvrir 8D"


def _effective_severity(base: str, qty: float, escalation_qty: float) -> str:
    """Apply escalation rule (R5) and return effective severity."""
    if qty < escalation_qty:
        return base
    mapping = {"MINOR": "MAJOR", "MAJOR": "CRITICAL", "CRITICAL": "CRITICAL"}
    return mapping.get(base, base)


def _recurrence_legacy(
    line_rows: list[dict[str, str]],
    recurrence_window: int,
    recurrence_min: int,
) -> list[dict[str, Any]]:
    """D8 forward-scan: legacy per-anchor algorithm.

    For each anchor row, count matching records within the window starting at
    the anchor date. Each qualifying anchor emits a separate RECURRENCE row.
    """
    n = len(line_rows)
    rows = []
    for i in range(n):
        anchor = line_rows[i]
        d1 = datetime.date.fromisoformat(anchor["date"])
        cnt = 0
        for j in range(i, n):
            candidate = line_rows[j]
            if (
                candidate["part_ref"] == anchor["part_ref"]
                and candidate["defect_code"] == anchor["defect_code"]
            ):
                d2 = datetime.date.fromisoformat(candidate["date"])
                if (d2 - d1).days <= recurrence_window:
                    cnt += 1
        if cnt >= recurrence_min:
            rows.append(
                {
                    "type": "RECURRENCE",
                    "id": anchor["id"],
                    "date": anchor["date"],
                    "ligne": anchor["line"],
                    "piece": anchor["part_ref"],
                    "code": anchor["defect_code"],
                    "qte": cnt,
                    "message": _MSG_RECURRENCE,
                }
            )
    return rows


def _recurrence_sliding(
    defect_log: list[dict[str, str]],
    recurrence_window_days: int,
    recurrence_min: int,
) -> list[dict[str, Any]]:
    """D8 fix: proper sliding-window algorithm (R6).

    Emits exactly one RECURRENCE flag per qualifying window, keyed on the
    earliest occurrence in the window.  Matches rules_check.py compute_alerts.
    """
    by_triplet: dict[tuple[str, str, str], list[tuple[datetime.date, str]]] = (
        defaultdict(list)
    )
    for rec in defect_log:
        key = (rec["line"], rec["part_ref"], rec["defect_code"])
        by_triplet[key].append((datetime.date.fromisoformat(rec["date"]), rec["id"]))

    seen_windows: set[tuple[Any, datetime.date]] = set()
    rows: list[dict[str, Any]] = []

    for key, occurrences in by_triplet.items():
        occurrences.sort()
        for d_start, id_start in occurrences:
            window_end = d_start + datetime.timedelta(days=recurrence_window_days - 1)
            count = sum(1 for d2, _ in occurrences if d_start <= d2 <= window_end)
            if count >= recurrence_min:
                wk = (key, d_start)
                if wk not in seen_windows:
                    seen_windows.add(wk)
                    rows.append(
                        {
                            "type": "RECURRENCE",
                            "id": id_start,
                            "date": d_start.isoformat(),
                            "ligne": key[0],
                            "piece": key[1],
                            "code": key[2],
                            "qte": count,
                            "message": _MSG_RECURRENCE,
                        }
                    )
    return rows


def compute(ds: Dataset, cfg: Config) -> list[dict[str, Any]]:
    """
    Compute alert rows.

    Returns list of dicts with keys:
      type, id, date, ligne, piece, code, qte, message
    """
    # --- thresholds ---
    if cfg.d7_hardcoded_params:
        escalation_qty = 20.0
        recurrence_window = 6  # d2 - d1 <= 6 means a 7-day span
        recurrence_min = 3
        recurrence_window_days = 7  # for sliding-window path
    else:
        escalation_qty = float(ds.params_map["severity_escalation_qty"])
        recurrence_window = int(ds.params_map["recurrence_window_days"]) - 1
        recurrence_min = int(ds.params_map["recurrence_min_count"])
        recurrence_window_days = int(ds.params_map["recurrence_window_days"])

    rows: list[dict[str, Any]] = []

    # --- Part A: CRITICAL alerts ---
    for row in ds.defect_log:
        code = row["defect_code"]
        qty = float(row["qty"])
        defect_info = ds.defect_map.get(code, {})
        base_sev = defect_info.get("base_severity", "MINOR")
        eff_sev = _effective_severity(base_sev, qty, escalation_qty)
        if eff_sev == "CRITICAL":
            rows.append(
                {
                    "type": "CRITIQUE",
                    "id": row["id"],
                    "date": row["date"],
                    "ligne": row["line"],
                    "piece": row["part_ref"],
                    "code": code,
                    "qte": int(qty),
                    "message": _MSG_CRITIQUE,
                }
            )

    # --- Part B: RECURRENCE alerts ---
    if cfg.d8_forward_scan:
        # Legacy D8: per-anchor forward scan
        lines = ["L1", "L2", "L3", "L4"]
        for line in lines:
            line_rows = [r for r in ds.defect_log if r["line"] == line]
            rows.extend(
                _recurrence_legacy(line_rows, recurrence_window, recurrence_min)
            )
    else:
        # Corrected D8: proper sliding window (R6)
        rows.extend(
            _recurrence_sliding(ds.defect_log, recurrence_window_days, recurrence_min)
        )

    return rows


# ---------------------------------------------------------------------------
# CSV writer
# ---------------------------------------------------------------------------

COLUMNS = ["type", "id", "date", "ligne", "piece", "code", "qte", "message"]


def to_csv(rows: list[dict[str, Any]]) -> str:
    lines = [",".join(COLUMNS)]
    for row in rows:
        parts = [
            row["type"],
            row["id"],
            row["date"],
            row["ligne"],
            row["piece"],
            row["code"],
            str(row["qte"]),
            row["message"],
        ]
        lines.append(",".join(parts))
    return "\n".join(lines) + "\n"
