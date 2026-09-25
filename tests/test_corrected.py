"""
Corrected-mode tests for nctrack.

Each test corresponds to one or more decision entries (D1–D9) from
docs/DECISIONS.md. Numbers are taken directly from that document and from
tools/rules_check.py v4 output (which computes the correct rule value).

Where a test compares output that is also checked by tools/rules_check.py,
a comment marks it "cross-check vs rules_check.py".

Legacy-mode tests in tests/test_parity.py remain unchanged and must
continue to pass alongside these.
"""

from __future__ import annotations

import math
from pathlib import Path

import pytest

from nctrack import alertes, mens_global, mens_lignes, pareto, rap_hebdo
from nctrack.config import CorrectedConfig, LegacyConfig
from nctrack.diff import report_differences
from nctrack.loader import load

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"


# ---------------------------------------------------------------------------
# Session-scoped fixtures
# ---------------------------------------------------------------------------


@pytest.fixture(scope="session")
def ds():
    return load(DATA_DIR)


@pytest.fixture(scope="session")
def cfg():
    return CorrectedConfig()


@pytest.fixture(scope="session")
def leg_cfg():
    return LegacyConfig()


@pytest.fixture(scope="session")
def rh(ds, cfg):
    return rap_hebdo.compute(ds, cfg)


@pytest.fixture(scope="session")
def rh_legacy(ds, leg_cfg):
    return rap_hebdo.compute(ds, leg_cfg)


@pytest.fixture(scope="session")
def al(ds, cfg):
    return alertes.compute(ds, cfg)


@pytest.fixture(scope="session")
def al_legacy(ds, leg_cfg):
    return alertes.compute(ds, leg_cfg)


@pytest.fixture(scope="session")
def pt(ds, cfg):
    return pareto.compute(ds, cfg)


@pytest.fixture(scope="session")
def ml(ds, cfg, al):
    return mens_lignes.compute(ds, cfg, al)


@pytest.fixture(scope="session")
def mg(ds, cfg, al):
    return mens_global.compute(ds, cfg, al)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _rh_row(rows, semaine, ligne):
    return next(
        (r for r in rows if r["semaine"] == semaine and r["ligne"] == ligne), None
    )


def _ml_row(rows, mois, ligne):
    return next(
        (r for r in rows if r["mois"] == mois and r["ligne"] == ligne), None
    )


def _mg_row(rows, mois):
    return next((r for r in rows if r["mois"] == mois), None)


def _pt_row(rows, semaine_du, rang):
    return next(
        (r for r in rows if r["semaine_du"] == semaine_du and r["rang"] == rang), None
    )


def _approx(a, b, tol=0.005):
    try:
        return math.isclose(float(a), float(b), abs_tol=tol)
    except (TypeError, ValueError):
        return str(a) == str(b)


# ===========================================================================
# D1 — L4 defect qty halved (KEEP with explicit factor 0.5 → corrected = 1.0)
# Decision: "KEEP the L4 halving, but as an explicit, documented parameter
#            (L4 factor = 0.5), not a hidden rule."
# In corrected mode (CorrectedConfig) d1_l4_halving=False, so factor = 1.0.
# Numbers from DECISIONS.md and cross-checked against rules_check.py v4.
# ===========================================================================


def test_d1_l4_nb_def_w26(rh):
    """D1: W26 L4 nb_def = 9 (rule); legacy = 4.5.
    Cross-check vs rules_check.py: D1 mismatch W26 L4 rule=9.0."""
    row = _rh_row(rh, "2026-W26", "L4")
    assert row is not None
    assert row["nb_def"] == 9.0


def test_d1_l4_tx_def_w26(rh):
    """D1: W26 L4 tx_def = 1.18 (rule); legacy = 0.59.
    Cross-check vs rules_check.py."""
    row = _rh_row(rh, "2026-W26", "L4")
    assert row is not None
    assert _approx(row["tx_def"], 1.18)


def test_d1_l4_nb_def_w28(rh):
    """D1: W28 L4 nb_def = 28 (rule); legacy = 14.
    Cross-check vs rules_check.py: DECISIONS.md example."""
    row = _rh_row(rh, "2026-W28", "L4")
    assert row is not None
    assert row["nb_def"] == 28.0


def test_d1_l4_nb_def_w29(rh):
    """D1: W29 L4 nb_def = 40 (rule); legacy = 20.
    Cross-check vs rules_check.py."""
    row = _rh_row(rh, "2026-W29", "L4")
    assert row is not None
    assert row["nb_def"] == 40.0


@pytest.mark.parametrize("semaine,expected_nb_def", [
    ("2026-W23", 8.0),
    ("2026-W24", 12.0),
    ("2026-W25", 4.0),
    ("2026-W26", 9.0),
    ("2026-W27", 6.0),
    ("2026-W28", 28.0),
    ("2026-W29", 40.0),
    ("2026-W30", 3.0),
])
def test_d1_all_l4_weekly_nb_def(rh, semaine, expected_nb_def):
    """D1: all L4 weekly nb_def values match rules_check.py expectations."""
    row = _rh_row(rh, semaine, "L4")
    assert row is not None, f"Missing row {semaine} L4"
    assert row["nb_def"] == expected_nb_def, (
        f"{semaine} L4: nb_def={row['nb_def']!r}, expected {expected_nb_def}"
    )


def test_d1_monthly_june_l4_defauts(ml):
    """D1+D4: June L4 defauts = 34 (rule); legacy = 16.5.
    Cross-check vs rules_check.py: D1+D4 mismatch 2026-06 L4."""
    row = _ml_row(ml, "2026-06", "L4")
    assert row is not None
    assert row["defauts"] == 34.0


def test_d1_monthly_july_l4_defauts(ml):
    """D1+D4: July L4 defauts = 76 (rule); legacy = 35.
    Cross-check vs rules_check.py."""
    row = _ml_row(ml, "2026-07", "L4")
    assert row is not None
    assert row["defauts"] == 76.0


# ===========================================================================
# D2 — ROUGE threshold: >= 3 fixed to > 3 (strict)
# Decision: "FIX: RED only above 3.0 % (strict), with one threshold used
#            everywhere."
# Only row in current data: W27 L3 tx_rebut = 3.00 %.
# Cross-check vs rules_check.py: D2 mismatch W27 L3 statut ROUGE→ORANGE.
# ===========================================================================


def test_d2_w27_l3_statut_is_orange(rh):
    """D2: W27 L3 tx_rebut = 3.00 % must be ORANGE (not ROUGE).
    Cross-check vs rules_check.py."""
    row = _rh_row(rh, "2026-W27", "L3")
    assert row is not None
    assert _approx(row["tx_rebut"], 3.00)
    assert row["statut"] == "ORANGE", (
        f"W27 L3 tx_rebut={row['tx_rebut']}: expected ORANGE, got {row['statut']}"
    )


def test_d2_legacy_was_rouge(rh_legacy):
    """D2: Legacy W27 L3 statut must be ROUGE (confirming the fix matters)."""
    row = _rh_row(rh_legacy, "2026-W27", "L3")
    assert row is not None
    assert row["statut"] == "ROUGE"


def test_d2_above_threshold_still_rouge(rh):
    """D2: Any row with tx_rebut > 3 must still be ROUGE in corrected mode."""
    for row in rh:
        if row["tx_rebut"] > 3.0:
            assert row["statut"] == "ROUGE", (
                f"{row['semaine']} {row['ligne']}: tx_rebut={row['tx_rebut']} "
                f"expected ROUGE, got {row['statut']}"
            )


def test_d2_exactly_three_is_orange(rh):
    """D2: No row with tx_rebut == 3.0 may be ROUGE in corrected mode."""
    for row in rh:
        if _approx(row["tx_rebut"], 3.00):
            assert row["statut"] != "ROUGE", (
                f"{row['semaine']} {row['ligne']}: tx_rebut=3.00 must not be ROUGE"
            )


# ===========================================================================
# D4 — ACCEPT excluded from monthly nd (fix: include ACCEPT)
# Decision: "FIX: ACCEPT counts as a defect in the monthly figures too (R1/R7)."
# Numbers from DECISIONS.md §D4 and rules_check.py v4.
# ===========================================================================


def test_d4_june_l1_defauts(ml):
    """D4: June L1 defauts = 60 (rule); legacy = 51 (9 ACCEPT excluded).
    Cross-check vs rules_check.py."""
    row = _ml_row(ml, "2026-06", "L1")
    assert row is not None
    assert row["defauts"] == 60.0


def test_d4_june_l3_defauts(ml):
    """D4: June L3 defauts = 35 (rule); legacy = 29.
    Cross-check vs rules_check.py."""
    row = _ml_row(ml, "2026-06", "L3")
    assert row is not None
    assert row["defauts"] == 35.0


def test_d4_july_l1_defauts(ml):
    """D4: July L1 defauts = 40 (rule); legacy = 32.
    Cross-check vs rules_check.py."""
    row = _ml_row(ml, "2026-07", "L1")
    assert row is not None
    assert row["defauts"] == 40.0


def test_d4_july_l3_defauts(ml):
    """D4: July L3 defauts = 46 (rule); legacy = 42.
    Cross-check vs rules_check.py."""
    row = _ml_row(ml, "2026-07", "L3")
    assert row is not None
    assert row["defauts"] == 46.0


def test_d4_june_l1_tx_def(ml):
    """D4: June L1 tx_def = 4.88 (rule); legacy = 4.15.
    Cross-check vs rules_check.py."""
    row = _ml_row(ml, "2026-06", "L1")
    assert row is not None
    assert _approx(row["tx_def"], 4.88)


# ===========================================================================
# D5 — Pareto ISO weeks (Mon–Sun) instead of Sun–Sat
# Decision: "FIX: ISO weeks (Mon–Sun) everywhere."
# DECISIONS.md example: ISO week ending 2026-06-07 ranks D01 first (qty 15).
# Cross-check vs rules_check.py compute_pareto output.
# ===========================================================================


def test_d5_pareto_week_2026_06_07_rank1_is_d01(pt):
    """D5: ISO week 2026-06-07 rank 1 = D01, qty 15.
    Legacy had D05 rank 1, qty 9.
    Cross-check vs rules_check.py."""
    row = _pt_row(pt, "2026-06-07", 1)
    assert row is not None
    assert row["code"] == "D01", f"Expected D01, got {row['code']}"
    assert row["qte"] == 15, f"Expected 15, got {row['qte']}"


def test_d5_pareto_no_legacy_week_2026_05_31(pt):
    """D5: ISO mode must not produce a 2026-05-31 week (that was a Sun-Sat artefact)."""
    week_labels = {r["semaine_du"] for r in pt}
    assert "2026-05-31" not in week_labels, (
        "Found spurious Sun-Sat week 2026-05-31 in corrected Pareto"
    )


def test_d5_pareto_has_week_2026_07_26(pt):
    """D5: ISO mode must include week 2026-07-26 (Mon 2026-07-20 to Sun 2026-07-26)."""
    week_labels = {r["semaine_du"] for r in pt}
    assert "2026-07-26" in week_labels, (
        "Missing ISO week 2026-07-26 in corrected Pareto"
    )


def test_d5_pareto_week_count(pt):
    """D5: Corrected Pareto has 8 ISO week buckets (legacy had 8 Sun-Sat buckets,
    but different ones — legacy has 2026-05-31 instead of 2026-07-26)."""
    week_labels = sorted(set(r["semaine_du"] for r in pt))
    assert len(week_labels) == 8, (
        f"Expected 8 ISO week buckets, got {len(week_labels)}: {week_labels}"
    )


# ===========================================================================
# D6 — Zero-production-day records kept (fix: remove ProdJour = 0 guard)
# Decision: "FIX: keep every defect; guard the rate division only when
#            production is zero."
# Confirmed instance: NC-0040 L2 2026-06-17 REWORK qty=2.
# Effects from DECISIONS.md: W25 L2 nb_def legacy=2 → rule=4.
# Cross-check vs rules_check.py.
# ===========================================================================


def test_d6_w25_l2_nb_def(rh):
    """D6: W25 L2 nb_def = 4 (rule); legacy = 2 (NC-0040 excluded).
    Cross-check vs rules_check.py."""
    row = _rh_row(rh, "2026-W25", "L2")
    assert row is not None
    assert row["nb_def"] == 4.0


def test_d6_w25_l2_tx_def(rh):
    """D6: W25 L2 tx_def = 2.03 (rule); legacy = 1.02.
    Cross-check vs rules_check.py."""
    row = _rh_row(rh, "2026-W25", "L2")
    assert row is not None
    assert _approx(row["tx_def"], 2.03)


def test_d6_w25_l2_cnq_eur(rh):
    """D6: W25 L2 cnq_eur = 1206 (rule, NC-0040 adds 2×0.5h×45€=45€);
    legacy = 1161.
    Cross-check vs rules_check.py."""
    row = _rh_row(rh, "2026-W25", "L2")
    assert row is not None
    assert row["cnq_eur"] == 1206


def test_d6_june_l2_defauts(ml):
    """D4+D6: June L2 defauts = 31 (rule); legacy = 25.
    Cross-check vs rules_check.py."""
    row = _ml_row(ml, "2026-06", "L2")
    assert row is not None
    assert row["defauts"] == 31.0


def test_d6_june_l2_tx_def(ml):
    """D4+D6: June L2 tx_def = 3.09 (rule); legacy = 2.49.
    Cross-check vs rules_check.py."""
    row = _ml_row(ml, "2026-06", "L2")
    assert row is not None
    assert _approx(row["tx_def"], 3.09)


# ===========================================================================
# D7 — Thresholds read from parameters.csv (no hardcoding)
# Decision: "FIX: read all thresholds from parameters.csv, with no hard-coded
#            values."
# No numerical change on current data (hardcoded = CSV values).
# Test that corrected mode reads from params_map (d7_hardcoded_params=False).
# ===========================================================================


def test_d7_corrected_config_reads_params(cfg):
    """D7: CorrectedConfig has d7_hardcoded_params=False."""
    assert cfg.d7_hardcoded_params is False


def test_d7_rh_cnq_consistent_with_params(ds, cfg, rh):
    """D7: Weekly cnq values are the same as when params == hardcoded values
    (since data/parameters.csv matches VBA hardcoded values on current data)."""
    # If D7 fix breaks something, cnq values will differ from legacy for D7 reason.
    # Since the CSV values equal the hardcoded ones, corrected cnq should match
    # corrected-D9 legacy (i.e. only D9 matters for cnq delta).
    # At minimum: all cnq_eur values are non-negative integers.
    for row in rh:
        assert isinstance(row["cnq_eur"], int), (
            f"{row['semaine']} {row['ligne']}: cnq_eur is not int: {row['cnq_eur']!r}"
        )
        assert row["cnq_eur"] >= 0


# ===========================================================================
# D8 — Recurrence: proper sliding window instead of per-anchor forward scan
# Decision: "FIX: implement R6 exactly as written, with one flag per
#            qualifying window."
# Current data: one qualifying triplet (L2/P-2001/D03, span 4 days).
# Both legacy and corrected produce 1 RECURRENCE alert for current data
# (the structural fix doesn't change the count but is architecturally correct).
# Cross-check vs rules_check.py compute_alerts.
# ===========================================================================


def test_d8_recurrence_count_unchanged(al, al_legacy):
    """D8: current data has one qualifying window — count unchanged (1 each)."""
    leg_recu = [a for a in al_legacy if a["type"] == "RECURRENCE"]
    cor_recu = [a for a in al if a["type"] == "RECURRENCE"]
    assert len(leg_recu) == 1
    assert len(cor_recu) == 1


def test_d8_recurrence_triplet(al):
    """D8: the single RECURRENCE alert is for L2/P-2001/D03 (NC-0085, 2026-07-06).
    Cross-check vs rules_check.py."""
    recu = [a for a in al if a["type"] == "RECURRENCE"]
    assert len(recu) == 1
    a = recu[0]
    assert a["ligne"] == "L2"
    assert a["piece"] == "P-2001"
    assert a["code"] == "D03"
    assert a["qte"] == 3


def test_d8_corrected_config_uses_sliding_window(cfg):
    """D8: CorrectedConfig has d8_forward_scan=False (sliding window path)."""
    assert cfg.d8_forward_scan is False


def test_d8_legacy_config_uses_forward_scan(leg_cfg):
    """D8: LegacyConfig has d8_forward_scan=True."""
    assert leg_cfg.d8_forward_scan is True


# ===========================================================================
# D9 — CNQ rounded per record fixed to round only final totals
# Decision: "FIX: round only the final totals."
# The per-record rounding causes a net −9 € error over the full dataset
# (108 067 € per-record-rounded vs 108 076 € exact-then-round).
# ===========================================================================


def test_d9_total_global_cnq(mg):
    """D9: total CNQ across both months = 108 076 € in corrected mode.
    Legacy total = 108 067 € (−9 € due to per-record CLng rounding)."""
    total = sum(r["cnq_eur"] for r in mg)
    assert total == 108076, (
        f"Expected total global CNQ = 108076, got {total}"
    )


def test_d9_legacy_total_cnq(ds, leg_cfg, al_legacy):
    """D9: legacy total CNQ = 108 022 € (combines per-record CLng rounding error
    with D1/D6 exclusions; corrected total = 108 076 €)."""
    mg_leg = mens_global.compute(ds, leg_cfg, al_legacy)
    total = sum(r["cnq_eur"] for r in mg_leg)
    assert total == 108022, (
        f"Expected legacy total CNQ = 108022, got {total}"
    )


def test_d9_w25_l2_cnq_corrected(rh):
    """D9: W25 L2 cnq_eur in corrected mode = 1206 (exact accumulation).
    Note: D6 also affects this value (+45€ from NC-0040), so this tests
    both D6 and D9 together."""
    row = _rh_row(rh, "2026-W25", "L2")
    assert row is not None
    assert row["cnq_eur"] == 1206


@pytest.mark.parametrize("semaine,ligne,expected_cnq", [
    # Values from running corrected mode — exact accumulation + final rounding
    ("2026-W23", "L1", 594),
    ("2026-W23", "L2", 1193),
    ("2026-W23", "L3", 952),
    ("2026-W23", "L4", 704),
    ("2026-W25", "L2", 1206),  # also D6 (+45 from NC-0040)
    ("2026-W26", "L2", 1416),
    ("2026-W26", "L4", 4715),
    ("2026-W27", "L1", 68),
    ("2026-W28", "L4", 16530),
    ("2026-W29", "L4", 24048),
])
def test_d9_weekly_cnq_spot_checks(rh, semaine, ligne, expected_cnq):
    """D9: spot-check corrected weekly CNQ values (exact accumulation + round)."""
    row = _rh_row(rh, semaine, ligne)
    assert row is not None, f"Missing row {semaine} {ligne}"
    assert row["cnq_eur"] == expected_cnq, (
        f"{semaine} {ligne}: cnq_eur={row['cnq_eur']}, expected {expected_cnq}"
    )


def test_d9_monthly_l4_june_cnq(ml):
    """D9: June L4 monthly cnq_eur in corrected mode = 12198."""
    row = _ml_row(ml, "2026-06", "L4")
    assert row is not None
    assert row["cnq_eur"] == 12198


def test_d9_monthly_global_june_cnq(mg):
    """D9: June global cnq_eur in corrected mode = 33793."""
    row = _mg_row(mg, "2026-06")
    assert row is not None
    assert row["cnq_eur"] == 33793


def test_d9_monthly_global_july_cnq(mg):
    """D9: July global cnq_eur in corrected mode = 74283."""
    row = _mg_row(mg, "2026-07")
    assert row is not None
    assert row["cnq_eur"] == 74283


# ===========================================================================
# Global D1+D4+D6 compound: monthly global defauts
# Decision: combine all three fixes — defauts for June and July should match
# rules_check.py compute_monthly_global output.
# ===========================================================================


def test_compound_june_global_defauts(mg):
    """D1+D4+D6: June global defauts = 160 (rule); legacy = 121.5.
    Cross-check vs rules_check.py."""
    row = _mg_row(mg, "2026-06")
    assert row is not None
    assert row["defauts"] == 160.0


def test_compound_july_global_defauts(mg):
    """D1+D4+D6: July global defauts = 205 (rule); legacy = 143.
    Cross-check vs rules_check.py."""
    row = _mg_row(mg, "2026-07")
    assert row is not None
    assert row["defauts"] == 205.0


# ===========================================================================
# report_differences() — the diff function
# Verify it finds the expected number of differences and tags them correctly.
# ===========================================================================


@pytest.fixture(scope="session")
def diffs(ds):
    return report_differences(ds)


def test_diff_returns_list(diffs):
    """diff: report_differences returns a non-empty list."""
    assert isinstance(diffs, list)
    assert len(diffs) > 0


def test_diff_has_required_keys(diffs):
    """diff: every entry has the required keys."""
    required = {"report", "key", "field", "legacy", "corrected", "decision", "description"}
    for d in diffs:
        assert required.issubset(d.keys()), (
            f"Missing keys in diff entry: {required - d.keys()}"
        )


def test_diff_d1_present(diffs):
    """diff: D1 differences are reported for L4 nb_def in rap_hebdo."""
    d1 = [d for d in diffs if d["decision"] == "D1" and d["report"] == "rap_hebdo"
          and d["field"] == "nb_def"]
    assert len(d1) > 0, "No D1/rap_hebdo/nb_def differences found"


def test_diff_d2_present(diffs):
    """diff: D2 differences are reported for W27 L3 statut."""
    d2 = [d for d in diffs if d["decision"] == "D2"
          and d["key"] == ("2026-W27", "L3") and d["field"] == "statut"]
    assert len(d2) == 1
    assert d2[0]["legacy"] == "ROUGE"
    assert d2[0]["corrected"] == "ORANGE"


def test_diff_d4_present(diffs):
    """diff: D4 differences are reported for mens_lignes defauts."""
    d4 = [d for d in diffs if d["decision"] in ("D1", "D4")
          and d["report"] == "mens_lignes" and d["field"] == "defauts"]
    assert len(d4) > 0


def test_diff_d5_present(diffs):
    """diff: D5 differences are reported for pareto."""
    d5 = [d for d in diffs if d["decision"] == "D5"]
    assert len(d5) > 0


def test_diff_d9_present(diffs):
    """diff: D9 differences are reported for cnq_eur fields."""
    d9 = [d for d in diffs if d["decision"] == "D9"]
    assert len(d9) > 0


def test_diff_decision_descriptions_non_empty(diffs):
    """diff: all decisions have a non-empty description string."""
    for d in diffs:
        assert isinstance(d["description"], str)
        assert len(d["description"]) > 0, (
            f"Empty description for decision {d['decision']}"
        )


def test_diff_w27_l3_statut_legacy_rouge_corrected_orange(diffs):
    """diff: W27 L3 statut: legacy=ROUGE, corrected=ORANGE (D2)."""
    entry = next(
        (d for d in diffs
         if d["report"] == "rap_hebdo"
         and d["key"] == ("2026-W27", "L3")
         and d["field"] == "statut"),
        None,
    )
    assert entry is not None
    assert entry["legacy"] == "ROUGE"
    assert entry["corrected"] == "ORANGE"
    assert entry["decision"] == "D2"


def test_diff_june_global_defauts_change(diffs):
    """diff: June global defauts changes from 121.5 to 160.0."""
    entry = next(
        (d for d in diffs
         if d["report"] == "mens_global"
         and d["key"] == "2026-06"
         and d["field"] == "defauts"),
        None,
    )
    assert entry is not None
    assert _approx(entry["legacy"], 121.5)
    assert _approx(entry["corrected"], 160.0)
