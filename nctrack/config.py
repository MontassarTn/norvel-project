"""
Configuration for nctrack.

Each divergence between the legacy VBA and the official rules (D1–D8)
is represented as a flag here.  In *legacy mode* every flag reproduces
the VBA behaviour exactly.  Corrected mode is left for a future sprint
(all flags will be True / rule-compliant).

Usage
-----
from nctrack.config import LegacyConfig
cfg = LegacyConfig()          # all divergences active
"""

from dataclasses import dataclass, field


@dataclass
class Config:
    # D1 — L4 defect qty halved in weekly AND monthly defect accumulator
    d1_l4_halving: bool = False
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


def LegacyConfig() -> Config:  # noqa: N802  (acts like a constructor alias)
    """Return a Config with every legacy divergence active."""
    return Config(
        d1_l4_halving=True,
        d2_rouge_gte=True,
        d4_accept_excluded_monthly=True,
        d5_sunday_anchor=True,
        d6_skip_zero_prod=True,
        d7_hardcoded_params=True,
    )


def CorrectedConfig() -> Config:  # noqa: N802
    """Return a Config with every divergence corrected (future use)."""
    return Config()
