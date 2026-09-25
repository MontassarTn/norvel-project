"""
diff — List every value that differs between legacy mode and corrected mode.

The single public function :func:`report_differences` runs both modes over the
same dataset and returns a list of dicts describing each diverging field, the
legacy value, the corrected value, and the decision ID that explains the change.

This module is intended for use by the application layer ("why did this number
change?") and by tests that need to cross-check corrected-mode output.

Usage
-----
    from nctrack.diff import report_differences
    from nctrack.loader import load

    ds = load("data/")
    diffs = report_differences(ds)
    for d in diffs:
        print(d["decision"], d["report"], d["key"], d["field"],
              d["legacy"], "->", d["corrected"])
"""

from __future__ import annotations

import math
from typing import Any

from nctrack import alertes, mens_global, mens_lignes, pareto, rap_hebdo
from nctrack.config import CorrectedConfig, LegacyConfig
from nctrack.loader import Dataset


# ---------------------------------------------------------------------------
# Decision descriptions (keyed by decision ID)
# ---------------------------------------------------------------------------

DECISION_DESCRIPTIONS: dict[str, str] = {
    "D1": (
        "L4 defect qty halved in legacy (factor 0.5). "
        "FIX: use factor 1.0 — keep halving as explicit parameter per decision."
    ),
    "D2": (
        "Legacy scrap threshold: tx_rebut >= 3 → ROUGE (should be strictly > 3). "
        "FIX: RED only above 3.0 % (strict), one threshold used everywhere."
    ),
    "D4": (
        "Legacy excludes ACCEPT disposition from monthly defect count. "
        "FIX: ACCEPT counts as a defect in monthly figures (R1/R7)."
    ),
    "D5": (
        "Legacy Pareto uses Sun-Sat week buckets. "
        "FIX: ISO weeks (Mon-Sun) everywhere."
    ),
    "D6": (
        "Legacy silently skips defect records on zero-production days. "
        "FIX: keep every defect; guard the rate division only when production is zero."
    ),
    "D8": (
        "Legacy recurrence uses per-anchor forward scan (may duplicate on dense data). "
        "FIX: implement R6 exactly, with one flag per qualifying window."
    ),
    "D9": (
        "Legacy rounds each record's CNQ contribution to whole euros (VBA CLng). "
        "FIX: accumulate exact costs and round only the final totals."
    ),
}


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _approx_equal(a: Any, b: Any, abs_tol: float = 0.005) -> bool:
    """Return True if a and b are numerically equal within tolerance, or
    string-equal.  Same logic as tools/rules_check.py check()."""
    try:
        return math.isclose(float(a), float(b), abs_tol=abs_tol)
    except (TypeError, ValueError):
        return str(a) == str(b)


def _diff(
    diffs: list[dict[str, Any]],
    report: str,
    key: Any,
    field: str,
    legacy: Any,
    corrected: Any,
    decision: str,
) -> None:
    if not _approx_equal(legacy, corrected):
        diffs.append(
            {
                "report": report,
                "key": key,
                "field": field,
                "legacy": legacy,
                "corrected": corrected,
                "decision": decision,
                "description": DECISION_DESCRIPTIONS.get(decision, ""),
            }
        )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def report_differences(ds: Dataset) -> list[dict[str, Any]]:
    """
    Run legacy and corrected modes over *ds* and return a list of dicts
    describing every field whose value differs between the two modes.

    Each dict has the following keys:
      report     : str  — output report name ("rap_hebdo", "pareto", etc.)
      key        : any  — row identifier (e.g. ("2026-W26", "L4") or "2026-06")
      field      : str  — column name
      legacy     : any  — value produced by LegacyConfig()
      corrected  : any  — value produced by CorrectedConfig()
      decision   : str  — decision ID that explains the change (D1–D9)
      description: str  — human-readable decision summary

    Rows are ordered by report, then by key, then by field.
    """
    leg_cfg = LegacyConfig()
    cor_cfg = CorrectedConfig()

    diffs: list[dict[str, Any]] = []

    # ------------------------------------------------------------------ #
    # rap_hebdo                                                           #
    # ------------------------------------------------------------------ #
    leg_rh = {(r["semaine"], r["ligne"]): r for r in rap_hebdo.compute(ds, leg_cfg)}
    cor_rh = {(r["semaine"], r["ligne"]): r for r in rap_hebdo.compute(ds, cor_cfg)}

    for key in sorted(set(leg_rh) | set(cor_rh)):
        lr = leg_rh.get(key, {})
        cr = cor_rh.get(key, {})
        line = key[1]

        # D1: nb_def and tx_def for L4
        if line == "L4":
            _diff(diffs, "rap_hebdo", key, "nb_def", lr.get("nb_def"), cr.get("nb_def"), "D1")
            _diff(diffs, "rap_hebdo", key, "tx_def", lr.get("tx_def"), cr.get("tx_def"), "D1")

        # D2: statut for any row where tx_rebut is exactly at the threshold
        _diff(diffs, "rap_hebdo", key, "statut", lr.get("statut"), cr.get("statut"), "D2")

        # D6: nb_def, tx_def, cnq_eur for rows affected by zero-prod filter
        if line != "L4":  # D6 affects non-L4 rows
            _diff(diffs, "rap_hebdo", key, "nb_def", lr.get("nb_def"), cr.get("nb_def"), "D6")
            _diff(diffs, "rap_hebdo", key, "tx_def", lr.get("tx_def"), cr.get("tx_def"), "D6")

        # D9: cnq_eur for any row with fractional per-record costs
        _diff(diffs, "rap_hebdo", key, "cnq_eur", lr.get("cnq_eur"), cr.get("cnq_eur"), "D9")

    # ------------------------------------------------------------------ #
    # pareto                                                              #
    # ------------------------------------------------------------------ #
    leg_pt_rows = pareto.compute(ds, leg_cfg)
    cor_pt_rows = pareto.compute(ds, cor_cfg)
    leg_pt = {(r["semaine_du"], str(r["rang"])): r for r in leg_pt_rows}
    cor_pt = {(r["semaine_du"], str(r["rang"])): r for r in cor_pt_rows}

    for key in sorted(set(leg_pt) | set(cor_pt)):
        lr = leg_pt.get(key, {})
        cr = cor_pt.get(key, {})
        for field in ("code", "qte", "pct", "cumul_pct", "prioritaire"):
            _diff(diffs, "pareto", key, field, lr.get(field), cr.get(field), "D5")

    # ------------------------------------------------------------------ #
    # alertes                                                             #
    # ------------------------------------------------------------------ #
    leg_al = alertes.compute(ds, leg_cfg)
    cor_al = alertes.compute(ds, cor_cfg)

    leg_recu = {(a["ligne"], a["piece"], a["code"]): a
                for a in leg_al if a["type"] == "RECURRENCE"}
    cor_recu = {(a["ligne"], a["piece"], a["code"]): a
                for a in cor_al if a["type"] == "RECURRENCE"}

    for key in sorted(set(leg_recu) | set(cor_recu)):
        lr = leg_recu.get(key, {})
        cr = cor_recu.get(key, {})
        _diff(diffs, "alertes", key, "qte", lr.get("qte"), cr.get("qte"), "D8")

    # ------------------------------------------------------------------ #
    # mens_lignes                                                         #
    # ------------------------------------------------------------------ #
    leg_ml = {(r["mois"], r["ligne"]): r
              for r in mens_lignes.compute(ds, leg_cfg, leg_al)}
    cor_ml = {(r["mois"], r["ligne"]): r
              for r in mens_lignes.compute(ds, cor_cfg, cor_al)}

    for key in sorted(set(leg_ml) | set(cor_ml)):
        lr = leg_ml.get(key, {})
        cr = cor_ml.get(key, {})
        line = key[1]

        # D1+D4 compound for L4; D4 for L1/L3; D4+D6 for L2
        if line == "L4":
            decision_def = "D1"
        elif line == "L2":
            decision_def = "D4"
        else:
            decision_def = "D4"

        _diff(diffs, "mens_lignes", key, "defauts",
              lr.get("defauts"), cr.get("defauts"), decision_def)
        _diff(diffs, "mens_lignes", key, "tx_def",
              lr.get("tx_def"), cr.get("tx_def"), decision_def)
        _diff(diffs, "mens_lignes", key, "statut",
              lr.get("statut"), cr.get("statut"), "D2")
        _diff(diffs, "mens_lignes", key, "cnq_eur",
              lr.get("cnq_eur"), cr.get("cnq_eur"), "D9")
        _diff(diffs, "mens_lignes", key, "tendance",
              lr.get("tendance"), cr.get("tendance"), decision_def)

    # ------------------------------------------------------------------ #
    # mens_global                                                         #
    # ------------------------------------------------------------------ #
    leg_mg = {r["mois"]: r for r in mens_global.compute(ds, leg_cfg, leg_al)}
    cor_mg = {r["mois"]: r for r in mens_global.compute(ds, cor_cfg, cor_al)}

    for key in sorted(set(leg_mg) | set(cor_mg)):
        lr = leg_mg.get(key, {})
        cr = cor_mg.get(key, {})
        _diff(diffs, "mens_global", key, "defauts",
              lr.get("defauts"), cr.get("defauts"), "D1")
        _diff(diffs, "mens_global", key, "cnq_eur",
              lr.get("cnq_eur"), cr.get("cnq_eur"), "D9")

    return diffs
