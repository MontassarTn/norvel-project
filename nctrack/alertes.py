"""
alertes — Alert computation (CalcAlertes equivalent).

No divergences directly from D1–D8 affect the alert output on current data.
Parameters D7 (hardcoded vs csv) affects escalation qty and recurrence window;
D8 (forward-scan recurrence) is reproduced exactly to match legacy output.

Output columns:
  type,id,date,ligne,piece,code,qte,message
"""

from __future__ import annotations

import datetime
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
    else:
        escalation_qty = float(ds.params_map["severity_escalation_qty"])
        recurrence_window = int(ds.params_map["recurrence_window_days"]) - 1
        recurrence_min = int(ds.params_map["recurrence_min_count"])

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

    # --- Part B: RECURRENCE alerts (D8 forward-scan, per line) ---
    lines = ["L1", "L2", "L3", "L4"]
    for line in lines:
        line_rows = [r for r in ds.defect_log if r["line"] == line]
        n = len(line_rows)
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
                        "ligne": line,
                        "piece": anchor["part_ref"],
                        "code": anchor["defect_code"],
                        "qte": cnt,
                        "message": _MSG_RECURRENCE,
                    }
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
