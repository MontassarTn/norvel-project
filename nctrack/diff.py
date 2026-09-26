"""
diff — List every value that differs between legacy mode and corrected mode.

The single public function :func:`report_differences` runs both modes over the
same dataset and returns a list of dicts describing each diverging field, the
legacy value, the corrected value, and the decision(s) that explain the change.

Attribution is computed, not hard-coded: for each decision, the reports are
recomputed in legacy mode with only that decision's setting switched to its
corrected value.  A field is attributed to every decision that changes it on
its own; a field changed by several decisions gets a compound tag such as
"D4+D6".  A field that only changes when decisions are combined is tagged
"combined".

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

import dataclasses
import math
from typing import Any

from nctrack import alertes, mens_global, mens_lignes, pareto, rap_hebdo
from nctrack.config import Config, CorrectedConfig, LegacyConfig
from nctrack.loader import Dataset


# ---------------------------------------------------------------------------
# Decision descriptions (keyed by decision ID)
# ---------------------------------------------------------------------------

DECISION_DESCRIPTIONS: dict[str, str] = {
    "D1": (
        "Legacy halves L4 defect quantities (hidden rule since 2011). "
        "KEEP: same 0.5 factor in corrected mode, now an explicit, documented parameter."
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
    "D7": (
        "Legacy hard-codes every threshold and never reads parameters.csv. "
        "FIX: read all thresholds from parameters.csv."
    ),
    "D8": (
        "Legacy recurrence: per-anchor forward scan, same production line only. "
        "FIX: implement R6 exactly (same code + same part on any line), one flag per window."
    ),
    "D9": (
        "Legacy rounds each record's CNQ contribution to whole euros (VBA CLng). "
        "FIX: accumulate exact costs and round only the final totals."
    ),
}

# Columns that identify a row in each report
_ROW_KEYS: dict[str, tuple[str, ...]] = {
    "rap_hebdo": ("semaine", "ligne"),
    "pareto": ("semaine_du", "rang"),
    "alertes": ("type", "id"),
    "mens_lignes": ("mois", "ligne"),
    "mens_global": ("mois",),
}

_REPORT_ORDER = list(_ROW_KEYS)


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


def _run(ds: Dataset, cfg: Config) -> dict[str, list[dict[str, Any]]]:
    """Compute all five reports for one configuration."""
    al_rows = alertes.compute(ds, cfg)
    return {
        "rap_hebdo": rap_hebdo.compute(ds, cfg),
        "pareto": pareto.compute(ds, cfg),
        "alertes": al_rows,
        "mens_lignes": mens_lignes.compute(ds, cfg, al_rows),
        "mens_global": mens_global.compute(ds, cfg, al_rows),
    }


def _flatten(reports: dict[str, list[dict[str, Any]]]) -> dict[tuple, Any]:
    """Map (report, row key, field) → value for every non-key field."""
    flat: dict[tuple, Any] = {}
    for report, rows in reports.items():
        key_cols = _ROW_KEYS[report]
        for row in rows:
            key_parts = tuple(str(row[c]) for c in key_cols)
            key: Any = key_parts[0] if len(key_parts) == 1 else key_parts
            for field, value in row.items():
                if field not in key_cols:
                    flat[(report, key, field)] = value
    return flat


def _changed_fields(base: dict[tuple, Any], other: dict[tuple, Any]) -> set[tuple]:
    return {
        k for k in set(base) | set(other)
        if not _approx_equal(base.get(k), other.get(k))
    }


def _decision_id(field_name: str) -> str:
    """Config field "d4_accept_excluded_monthly" → decision "D4"."""
    return field_name.split("_", 1)[0].upper()


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
      legacy     : any  — value produced by LegacyConfig() (None if no such row)
      corrected  : any  — value produced by CorrectedConfig() (None if no such row)
      decision   : str  — decision(s) that explain the change, e.g. "D6" or "D4+D6"
      description: str  — human-readable decision summary

    Rows are ordered by report, then by key, then by field.
    """
    leg_cfg = LegacyConfig()
    cor_cfg = CorrectedConfig()

    legacy = _flatten(_run(ds, leg_cfg))
    corrected = _flatten(_run(ds, cor_cfg))

    # Which fields does each decision change on its own?
    changed_by: dict[str, set[tuple]] = {}
    for f in dataclasses.fields(Config):
        leg_value = getattr(leg_cfg, f.name)
        cor_value = getattr(cor_cfg, f.name)
        if leg_value == cor_value:
            continue  # e.g. D1: kept, same setting in both modes
        variant = dataclasses.replace(leg_cfg, **{f.name: cor_value})
        changed = _changed_fields(legacy, _flatten(_run(ds, variant)))
        changed_by.setdefault(_decision_id(f.name), set()).update(changed)

    decisions_in_order = sorted(changed_by, key=lambda d: int(d[1:]))

    diffs: list[dict[str, Any]] = []
    for k in _changed_fields(legacy, corrected):
        report, key, field = k
        ids = [d for d in decisions_in_order if k in changed_by[d]]
        diffs.append(
            {
                "report": report,
                "key": key,
                "field": field,
                "legacy": legacy.get(k),
                "corrected": corrected.get(k),
                "decision": "+".join(ids) if ids else "combined",
                "description": (
                    " | ".join(DECISION_DESCRIPTIONS.get(d, d) for d in ids)
                    or "Changes only when several decisions apply together."
                ),
            }
        )

    diffs.sort(key=lambda d: (_REPORT_ORDER.index(d["report"]), str(d["key"]), d["field"]))
    return diffs
