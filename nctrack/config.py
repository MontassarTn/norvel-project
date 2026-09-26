"""
Configuration for nctrack.

Each divergence between the legacy VBA and the official rules (D1–D9)
is represented as a flag here.  In *legacy mode* every flag reproduces
the VBA behaviour exactly.  In *corrected mode* every flag follows the
decision recorded in docs/DECISIONS.md.

New parameters (not in data/parameters.csv)
--------------------------------------------
L4_FACTOR : float
    D1 — the factor applied to L4 defective quantity (0.5: each L4 part
    counts as half a defect).  Decision D1 is KEEP: the halving stays in
    corrected mode, but as this explicit, documented parameter instead of
    a hidden rule in the code.

Usage
-----
from nctrack.config import LegacyConfig, CorrectedConfig
cfg = LegacyConfig()       # all divergences active
cfg = CorrectedConfig()    # decisions from docs/DECISIONS.md applied
"""

from dataclasses import dataclass, field


# ---------------------------------------------------------------------------
# L4 halving factor — stored here (nctrack config), never in data/parameters.csv
# ---------------------------------------------------------------------------

L4_FACTOR = 0.5  # D1: KEEP decision — explicit parameter (was hidden in the VBA since 2011)


@dataclass
class Config:
    # D1 — L4 defect qty multiplied by d1_l4_factor in weekly AND monthly defect counts
    d1_l4_halving: bool = True
    # D1 explicit factor — only used when d1_l4_halving is True
    d1_l4_factor: float = L4_FACTOR
    # D2 — scrap threshold uses >= 3 (True) instead of > 3 (False)
    d2_rouge_gte: bool = False
    # D4 — ACCEPT excluded from monthly nd accumulator
    d4_accept_excluded_monthly: bool = False
    # D5 — Pareto week anchor: True=Sunday-to-Saturday, False=ISO Mon–Sun
    d5_sunday_anchor: bool = False
    # D6 — defect records on zero-production days silently skipped
    d6_skip_zero_prod: bool = False
    # D7 — thresholds hardcoded (True) instead of read from parameters.csv
    #       In practice both sets of values are identical on current data,
    #       so this flag only affects *where* the values come from.
    d7_hardcoded_params: bool = False
    # D8 — recurrence algorithm: True=legacy per-anchor forward scan (may emit
    #       duplicate alerts on dense data), False=proper sliding-window (one flag
    #       per qualifying window, per R6)
    d8_forward_scan: bool = False
    # D9 — CNQ rounded per record (True=legacy CLng behaviour)
    #       vs accumulating exact fractional costs and rounding only the final total
    d9_round_per_record: bool = False


def LegacyConfig() -> Config:  # noqa: N802  (acts like a constructor alias)
    """Return a Config with every legacy divergence active."""
    return Config(
        d1_l4_halving=True,
        d1_l4_factor=L4_FACTOR,
        d2_rouge_gte=True,
        d4_accept_excluded_monthly=True,
        d5_sunday_anchor=True,
        d6_skip_zero_prod=True,
        d7_hardcoded_params=True,
        d8_forward_scan=True,
        d9_round_per_record=True,
    )


def CorrectedConfig() -> Config:  # noqa: N802
    """Return a Config that applies every decision in docs/DECISIONS.md
    (D1 kept as an explicit parameter, D2–D9 fixed)."""
    return Config()
