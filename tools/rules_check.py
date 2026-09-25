"""
tools/rules_check.py  (v3 – rule-only implementation)
=======================================================
Every computation in this file is derived solely from BUSINESS_RULES.md.
No VBA behaviour is copied, mirrored, or approximated.

Above each function the exact BUSINESS_RULES.md sentence(s) that define it
are quoted verbatim.  Where the rules are silent on an implementation detail,
the assumption is stated explicitly and collected in ASSUMPTIONS at the bottom
of this module (also echoed in docs/LEGACY_ANALYSIS.md §11).

Run from the repository root:
    python tools/rules_check.py

Output: mismatch list grouped by divergence row (D1 … Dn), plus a summary
count table.
"""

import csv
import math
import os
from collections import defaultdict
from datetime import date, timedelta

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA   = os.path.join(ROOT, "data")
LEGACY = os.path.join(ROOT, "sorties_legacy")

# ---------------------------------------------------------------------------
# ASSUMPTION LOG
# Every entry: (id, rule, assumption text)
# Printed at run-time so reviewers can see what the rules leave unspecified.
# ---------------------------------------------------------------------------
ASSUMPTIONS = [
    ("A1", "R3",
     "Weeks are ISO weeks (Mon–Sun) per §1 of BUSINESS_RULES.md.  The legacy "
     "output labels each Pareto week by the Sunday date of that week.  This "
     "checker uses the same Sunday-of-ISO-week label so that rows can be "
     "matched against sorties_legacy/pareto.csv.  The rule does not specify "
     "the label format; the Sunday label is an *output formatting* choice, "
     "not a computational one."),
    ("A2", "R3",
     "When two defect types have equal quantity in a week, this checker ranks "
     "the lower-numbered code first (D01 before D02, etc.).  BUSINESS_RULES.md "
     "does not specify a tie-break order.  The legacy VBA applies the same "
     "ordering (its arrays are indexed 1–8 in code order)."),
    ("A3", "R1/R7",
     "BUSINESS_RULES.md does not say to exclude defect records that fall on "
     "days for which no production entry exists.  This checker includes all "
     "records in the defect totals.  Production totals come only from "
     "production_log.csv.  If a (line, week/month) has defect records but "
     "zero production, a rate cannot be computed and the row is omitted from "
     "the report — consistent with R1's formula requiring a non-zero denominator."),
    ("A4", "R3",
     "R3 ranks by 'total defective quantity'.  If a defect_log record carries "
     "a negative qty (a correction entry such as NC-0125), it is summed "
     "algebraically.  A code whose net quantity is zero or negative contributes "
     "no defects for that week and is excluded from the Pareto."),
    ("A5", "R7",
     "R7 says 'top 3 defect types by quantity'.  This checker uses total "
     "defective quantity across all dispositions for the month (consistent "
     "with R1).  Tie-breaking follows A2."),
    ("A6", "R2/R7",
     "The status column in sorties_legacy/ uses French labels (VERT/ORANGE/ROUGE). "
     "R2 defines the thresholds in English.  This checker maps GREEN->VERT, "
     "ORANGE->ORANGE, RED->ROUGE to allow direct comparison."),
]

# ---------------------------------------------------------------------------
# Load data
# ---------------------------------------------------------------------------
def load_parameters():
    """Read every threshold from data/parameters.csv.
    BUSINESS_RULES.md §2 'Parameters' table is the canonical source."""
    params = {}
    with open(os.path.join(DATA, "parameters.csv"), newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            params[row["parameter"]] = float(row["value"])
    return params

def load_defect_types():
    """Returns dict code -> {label, base_severity, rework_hours (float or None)}."""
    dt = {}
    with open(os.path.join(DATA, "defect_types.csv"), newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            rh = row["rework_hours_per_part"].strip()
            dt[row["code"]] = {
                "label":         row["label"],
                "base_severity": row["base_severity"],
                "rework_hours":  float(rh) if rh else None,
            }
    return dt

def load_parts():
    """Returns dict part_ref -> unit_cost_eur (float)."""
    parts = {}
    with open(os.path.join(DATA, "parts.csv"), newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            parts[row["part_ref"]] = float(row["unit_cost_eur"])
    return parts

def load_defect_log():
    records = []
    with open(os.path.join(DATA, "defect_log.csv"), newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            records.append({
                "id":          row["id"],
                "date":        date.fromisoformat(row["date"]),
                "line":        row["line"],
                "part_ref":    row["part_ref"],
                "defect_code": row["defect_code"],
                "qty":         int(row["qty"]),
                "disposition": row["disposition"],
            })
    return records

def load_production_log():
    """Returns dict (date, line) -> qty_produced."""
    prod = {}
    with open(os.path.join(DATA, "production_log.csv"), newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            d = date.fromisoformat(row["date"])
            prod[(d, row["line"])] = int(row["qty_produced"])
    return prod

def load_legacy(name):
    path = os.path.join(LEGACY, name)
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))

# ---------------------------------------------------------------------------
# Week helpers
# R1/R2/R3/R7: "Weeks are ISO weeks (Monday to Sunday) throughout."
# ---------------------------------------------------------------------------
def iso_year_week(d: date) -> str:
    """ISO 8601 week label, e.g. '2026-W23'."""
    iso = d.isocalendar()
    return f"{iso[0]}-W{iso[1]:02d}"

def sunday_of_iso_week(d: date) -> date:
    """Return the Sunday that ends the ISO week containing d.
    Used only for labelling the Pareto output to match sorties_legacy/ format.
    See assumption A1."""
    # ISO weekday: Mon=1 … Sun=7
    return d + timedelta(days=7 - d.isoweekday())

# ---------------------------------------------------------------------------
# R4 – Cost of non-quality
# "For each record:
#   SCRAP  → qty × part unit cost
#   REWORK → qty × rework hours of the defect type × 45 €
#   ACCEPT → 0 € (the record still counts as a defect)
# A D06 (delamination) record declared as REWORK is costed as SCRAP,
# because delamination is not reworkable."
# ---------------------------------------------------------------------------
def r4_cnq(rec, defect_types, parts, labour_rate):
    qty  = rec["qty"]
    code = rec["defect_code"]
    disp = rec["disposition"]
    if disp == "SCRAP":
        return round(qty * parts[rec["part_ref"]], 0)
    if disp == "REWORK":
        if code == "D06":
            # "D06 … declared as REWORK is costed as SCRAP" (R4)
            return round(qty * parts[rec["part_ref"]], 0)
        rh = defect_types[code]["rework_hours"]
        return round(qty * rh * labour_rate, 0)
    # ACCEPT → 0
    return 0.0

# ---------------------------------------------------------------------------
# R2 – Scrap rate status
# "scrap rate > 3.0 % → RED
#  scrap rate > 2.0 % and ≤ 3.0 % → ORANGE
#  scrap rate ≤ 2.0 % → GREEN"
# (See A6 for French label mapping.)
# ---------------------------------------------------------------------------
def r2_status(scrap_pct, red_t, org_t):
    if scrap_pct > red_t:
        return "ROUGE"
    if scrap_pct > org_t:
        return "ORANGE"
    return "VERT"

# ---------------------------------------------------------------------------
# R5 – Severity escalation
# "Each record takes the base severity of its defect type.
#  If the record quantity is 20 or more, the severity is escalated by one
#  level (MINOR → MAJOR, MAJOR → CRITICAL). CRITICAL stays CRITICAL."
# ---------------------------------------------------------------------------
_SEV = ["MINOR", "MAJOR", "CRITICAL"]

def r5_severity(base_severity, qty, esc_qty):
    idx = _SEV.index(base_severity)
    if qty >= esc_qty and idx < 2:
        idx += 1
    return _SEV[idx]

# ---------------------------------------------------------------------------
# R1 + R2 + R4 — Weekly indicators
# R1: "defect rate (%) = total defective quantity / total quantity produced × 100
#      All dispositions (SCRAP, REWORK, ACCEPT) count as defects."
# R2: "scrap rate (%) = scrapped quantity / total quantity produced × 100"
# R4: CNQ summed per line per week.
# See A3 for records on zero-production days.
# ---------------------------------------------------------------------------
def compute_weekly(records, prod, defect_types, parts, params):
    labour_rate = params["rework_rate_eur_per_hour"]
    red_t = params["scrap_threshold_red_pct"]
    org_t = params["scrap_threshold_orange_pct"]

    prod_wk    = defaultdict(float)   # (week_label, line) -> produced qty
    defect_wk  = defaultdict(float)   # (week_label, line) -> defective qty (R1: all disp.)
    scrap_wk   = defaultdict(float)   # (week_label, line) -> scrapped qty  (R2)
    cnq_wk     = defaultdict(float)   # (week_label, line) -> CNQ euros     (R4)

    # Sum production per ISO week/line
    for (d, line), qty in prod.items():
        prod_wk[(iso_year_week(d), line)] += qty

    # Sum defects per ISO week/line — A3: include ALL records, no day-level filter
    for rec in records:
        key = (iso_year_week(rec["date"]), rec["line"])
        defect_wk[key] += rec["qty"]
        if rec["disposition"] == "SCRAP":
            scrap_wk[key] += rec["qty"]
        cnq_wk[key] += r4_cnq(rec, defect_types, parts, labour_rate)

    rows = []
    for key in sorted(prod_wk):
        wk, line = key
        p = prod_wk[key]
        if p <= 0:
            continue   # no rate computable (A3)
        nd  = defect_wk.get(key, 0.0)
        rb  = scrap_wk.get(key, 0.0)
        cq  = cnq_wk.get(key, 0.0)
        tx_def    = nd / p * 100
        tx_rebut  = rb / p * 100
        rows.append({
            "semaine":  wk,
            "ligne":    line,
            "produit":  p,
            "nb_def":   nd,
            "tx_def":   round(tx_def, 2),
            "rebut":    rb,
            "tx_rebut": round(tx_rebut, 2),
            "statut":   r2_status(tx_rebut, red_t, org_t),
            "cnq_eur":  round(cq, 0),
        })
    return rows

# ---------------------------------------------------------------------------
# R3 – Weekly Pareto
# "For each week, defect types are ranked by total defective quantity
#  (descending), with each type's share and the cumulative percentage.
#  The types needed to reach 80 % cumulative are marked as priorities
#  (the type that crosses 80 % is included)."
# See A1 (week label), A2 (tie-break), A4 (negative corrections).
# ---------------------------------------------------------------------------
def compute_pareto(records, defect_types):
    # Group by ISO week (A1: label as Sunday date of that week for CSV matching)
    qty_by_week = defaultdict(lambda: defaultdict(int))
    for rec in records:
        sun = sunday_of_iso_week(rec["date"])
        qty_by_week[sun][rec["defect_code"]] += rec["qty"]   # algebraic sum (A4)

    rows = []
    for sun in sorted(qty_by_week):
        week_data = qty_by_week[sun]
        # A4: exclude codes with net qty ≤ 0
        active = [(c, q) for c, q in week_data.items() if q > 0]
        if not active:
            continue
        total = sum(q for _, q in active)
        # Sort descending by qty; tie-break: lower code number first (A2)
        active.sort(key=lambda x: (-x[1], int(x[0][1:])))
        cumul = 0.0
        for rank, (code, qty) in enumerate(active, 1):
            pct    = qty / total * 100
            cumul += pct
            # R3: "the type that crosses 80 % is included"
            # A type is needed iff the cumulative BEFORE adding it was < 80 %
            priority = "X" if (cumul - pct) < 80.0 else ""
            rows.append({
                "semaine_du":  sun.isoformat(),
                "rang":        rank,
                "code":        code,
                "libelle":     defect_types[code]["label"],
                "qte":         qty,
                "pct":         round(pct, 2),
                "cumul_pct":   round(cumul, 2),
                "prioritaire": priority,
            })
    return rows

# ---------------------------------------------------------------------------
# R5 + R6 – Alerts
# R5: "Every CRITICAL record generates an alert in the weekly and monthly
#      summaries."
# R6: "If the same defect code is recorded on the same part reference 3 or
#      more times within any 7-day window (first and third occurrence at most
#      6 days apart), the tool raises a RECURRENCE flag."
# ---------------------------------------------------------------------------
def compute_alerts(records, defect_types, params):
    esc_qty = int(params["severity_escalation_qty"])
    rec_min = int(params["recurrence_min_count"])
    rec_win = int(params["recurrence_window_days"])

    alerts = []

    # ---- R5: CRITICAL alerts ------------------------------------------
    # "Each record takes the base severity of its defect type."
    # R5 applies to ALL records (no production-day filter is stated).
    for rec in records:
        base = defect_types[rec["defect_code"]]["base_severity"]
        eff  = r5_severity(base, rec["qty"], esc_qty)
        if eff == "CRITICAL":
            alerts.append({
                "type":    "CRITICAL",
                "id":      rec["id"],
                "date":    rec["date"].isoformat(),
                "ligne":   rec["line"],
                "piece":   rec["part_ref"],
                "code":    rec["defect_code"],
                "qte":     rec["qty"],
                "message": "defaut critique - prevenir resp. qualite",
            })

    # ---- R6: Recurrence -----------------------------------------------
    # "same defect code … same part reference … 3 or more times within any
    #  7-day window (first and third occurrence at most 6 days apart)"
    # Plain reading: sliding window of rec_win days; raise ONE flag per
    # qualifying window (keyed on the earliest date in each window).
    by_triplet = defaultdict(list)
    for rec in records:
        key = (rec["line"], rec["part_ref"], rec["defect_code"])
        by_triplet[key].append((rec["date"], rec["id"]))

    seen_windows = set()
    for key, occurrences in by_triplet.items():
        occurrences.sort()
        for i, (d_start, id_start) in enumerate(occurrences):
            window_end = d_start + timedelta(days=rec_win - 1)
            count = sum(1 for d2, _ in occurrences if d_start <= d2 <= window_end)
            if count >= rec_min:
                wk = (key, d_start)
                if wk not in seen_windows:
                    seen_windows.add(wk)
                    alerts.append({
                        "type":    "RECURRENCE",
                        "id":      id_start,
                        "date":    d_start.isoformat(),
                        "ligne":   key[0],
                        "piece":   key[1],
                        "code":    key[2],
                        "qte":     count,
                        "message": "recurrence 7j - ouvrir 8D",
                    })

    crit = [a for a in alerts if a["type"] == "CRITICAL"]
    recu = [a for a in alerts if a["type"] == "RECURRENCE"]
    return crit + recu

# ---------------------------------------------------------------------------
# R7 – Monthly summary (per-line)
# "defect rate and scrap rate per line (same definitions as R1 and R2,
#  over the month)"
# "total CNQ, per line and overall"
# "trend versus the previous month, per line: ↑ if the defect rate
#  increased by more than 0.5 points, ↓ if it decreased by more than
#  0.5 points, otherwise 'stable'"
# See A3 for records on zero-production days.
# ---------------------------------------------------------------------------
def compute_monthly_lines(records, prod, defect_types, parts, params):
    labour_rate = params["rework_rate_eur_per_hour"]
    red_t = params["scrap_threshold_red_pct"]
    org_t = params["scrap_threshold_orange_pct"]
    trend_band = params["trend_stable_band_pts"]

    prod_mo   = defaultdict(float)
    defect_mo = defaultdict(float)
    scrap_mo  = defaultdict(float)
    cnq_mo    = defaultdict(float)

    for (d, line), qty in prod.items():
        prod_mo[(d.year, d.month, line)] += qty

    # R7 = R1 over the month: ALL dispositions count as defects (A3: no day filter)
    for rec in records:
        key = (rec["date"].year, rec["date"].month, rec["line"])
        defect_mo[key] += rec["qty"]
        if rec["disposition"] == "SCRAP":
            scrap_mo[key] += rec["qty"]
        cnq_mo[key] += r4_cnq(rec, defect_types, parts, labour_rate)

    all_keys = sorted(k for k in prod_mo if prod_mo[k] > 0)
    rows = []
    prev_tx = {}   # (year, month, line) -> tx_def, for trend calculation
    for key in all_keys:
        yr, mo, line = key
        p  = prod_mo[key]
        nd = defect_mo.get(key, 0.0)
        rb = scrap_mo.get(key, 0.0)
        cq = cnq_mo.get(key, 0.0)
        tx_def   = nd / p * 100
        tx_rebut = rb / p * 100
        # Trend: compare to previous calendar month (R7 says "previous month")
        prev_key = (yr, mo - 1, line) if mo > 1 else (yr - 1, 12, line)
        if prev_key in prev_tx:
            diff = tx_def - prev_tx[prev_key]
            if diff > trend_band:
                trend = "HAUSSE"
            elif diff < -trend_band:
                trend = "BAISSE"
            else:
                trend = "STABLE"
        else:
            trend = "N/A"
        prev_tx[key] = tx_def
        rows.append({
            "mois":     f"{yr}-{mo:02d}",
            "ligne":    line,
            "produit":  p,
            "defauts":  nd,
            "tx_def":   round(tx_def, 2),
            "rebut":    rb,
            "tx_rebut": round(tx_rebut, 2),
            "statut":   r2_status(tx_rebut, red_t, org_t),
            "cnq_eur":  round(cq, 0),
            "tendance": trend,
        })
    return rows

# ---------------------------------------------------------------------------
# R7 – Monthly global summary
# "total quantity produced and total defective quantity"
# "total CNQ, per line and overall"
# "top 3 defect types by quantity"
# "number of CRITICAL alerts and RECURRENCE flags"
# ---------------------------------------------------------------------------
def compute_monthly_global(records, prod, defect_types, parts, params, alert_rows):
    labour_rate = params["rework_rate_eur_per_hour"]

    prod_mo   = defaultdict(float)
    defect_mo = defaultdict(float)
    cnq_mo    = defaultdict(float)
    tq_mo     = defaultdict(lambda: defaultdict(int))  # (yr,mo) -> code -> qty
    nbc = defaultdict(int)   # CRITICAL count per month
    nbr = defaultdict(int)   # RECURRENCE count per month

    for (d, line), qty in prod.items():
        prod_mo[(d.year, d.month)] += qty

    # A3: no day-level filter; A5: all dispositions contribute to top-3 qty
    for rec in records:
        mo  = (rec["date"].year, rec["date"].month)
        defect_mo[mo] += rec["qty"]
        cnq_mo[mo]    += r4_cnq(rec, defect_types, parts, labour_rate)
        tq_mo[mo][rec["defect_code"]] += rec["qty"]

    for a in alert_rows:
        d  = date.fromisoformat(a["date"])
        mo = (d.year, d.month)
        if a["type"] == "CRITICAL":
            nbc[mo] += 1
        else:
            nbr[mo] += 1

    rows = []
    for mo in sorted(k for k in prod_mo if prod_mo[k] > 0):
        yr, m = mo
        # A5: top-3 by qty, tie-break A2
        top3 = sorted(tq_mo[mo].items(), key=lambda x: (-x[1], int(x[0][1:])))[:3]
        rows.append({
            "mois":            f"{yr}-{m:02d}",
            "produit":         prod_mo[mo],
            "defauts":         defect_mo[mo],
            "cnq_eur":         round(cnq_mo[mo], 0),
            "top3":            " / ".join(c for c, _ in top3),
            "nb_critiques":    nbc[mo],
            "nb_recurrences":  nbr[mo],
        })
    return rows

# ---------------------------------------------------------------------------
# Comparison helpers
# ---------------------------------------------------------------------------
MISMATCHES = []

def _tag(m, divergence_id):
    """Attach a divergence row tag to a mismatch dict."""
    m["div"] = divergence_id
    MISMATCHES.append(m)

def check(section, key, field, expected, actual, div="?"):
    es, as_ = str(expected).strip(), str(actual).strip()
    try:
        match = math.isclose(float(es), float(as_), abs_tol=0.005)
    except ValueError:
        match = es == as_
    if not match:
        _tag({"section": section, "key": key, "field": field,
              "rule_says": es, "legacy_has": as_}, div)

def _f2(x):
    return f"{float(x):.2f}"

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    params       = load_parameters()
    defect_types = load_defect_types()
    parts        = load_parts()
    records      = load_defect_log()
    prod         = load_production_log()

    print("=" * 70)
    print("rules_check.py v3 -- R1-R7 vs sorties_legacy/")
    print("=" * 70)
    print("\nParameters (from data/parameters.csv):")
    for k, v in params.items():
        print(f"  {k} = {v}")

    print("\nAssumptions (rules are silent on these points):")
    for aid, rule, text in ASSUMPTIONS:
        print(f"  [{aid}/{rule}] {text}")

    # ------------------------------------------------------------------ #
    # RAP_HEBDO  (R1, R2, R4)                                            #
    # ------------------------------------------------------------------ #
    print("\n--- RAP_HEBDO (R1, R2, R4) ---")
    rule_h = {(r["semaine"], r["ligne"]): r
              for r in compute_weekly(records, prod, defect_types, parts, params)}
    leg_h  = {(r["semaine"], r["ligne"]): r
              for r in load_legacy("rap_hebdo.csv")}

    for key in sorted(rule_h):
        rr = rule_h[key]
        lr = leg_h.get(key)
        if lr is None:
            _tag({"section": "RAP_HEBDO", "key": key, "field": "row",
                  "rule_says": "present", "legacy_has": "MISSING"}, "?")
            continue
        lbl = f"RAP_HEBDO {key[0]} {key[1]}"
        check(lbl, key, "nb_def",   rr["nb_def"],  lr["nb_def"],   "D1")
        check(lbl, key, "tx_def",   _f2(rr["tx_def"]),   _f2(float(lr["tx_def"])),   "D1")
        check(lbl, key, "rebut",    rr["rebut"],   lr["rebut"],    "D1")
        check(lbl, key, "tx_rebut", _f2(rr["tx_rebut"]), _f2(float(lr["tx_rebut"])), "D1")
        check(lbl, key, "statut",   rr["statut"],  lr["statut"],   "D2")
        check(lbl, key, "cnq_eur",  rr["cnq_eur"], lr["cnq_eur"],  "D1")
    for key in sorted(leg_h):
        if key not in rule_h:
            _tag({"section": "RAP_HEBDO", "key": key, "field": "row",
                  "rule_says": "EXTRA in legacy", "legacy_has": str(leg_h[key])}, "?")

    # ------------------------------------------------------------------ #
    # PARETO  (R3)                                                        #
    # ------------------------------------------------------------------ #
    print("--- PARETO (R3) ---")
    rule_p = {(r["semaine_du"], str(r["rang"])): r
              for r in compute_pareto(records, defect_types)}
    leg_p  = {(r["semaine_du"], r["rang"]): r
              for r in load_legacy("pareto.csv")}

    for key in sorted(rule_p):
        rr = rule_p[key]
        lr = leg_p.get(key)
        if lr is None:
            _tag({"section": "PARETO", "key": key, "field": "row",
                  "rule_says": "present", "legacy_has": "MISSING"}, "D5")
            continue
        lbl = f"PARETO {key[0]} rank {key[1]}"
        check(lbl, key, "code",        rr["code"],      lr["code"],           "D5")
        check(lbl, key, "qte",         rr["qte"],       lr["qte"],            "D5")
        check(lbl, key, "pct",         _f2(rr["pct"]),  _f2(float(lr["pct"])), "D5")
        check(lbl, key, "cumul_pct",   _f2(rr["cumul_pct"]), _f2(float(lr["cumul_pct"])), "D5")
        check(lbl, key, "prioritaire", rr["prioritaire"], lr["prioritaire"],  "D5")
    for key in sorted(leg_p):
        if key not in rule_p:
            _tag({"section": "PARETO", "key": key, "field": "row",
                  "rule_says": "EXTRA in legacy", "legacy_has": str(leg_p[key])}, "D5")

    # ------------------------------------------------------------------ #
    # ALERTES  (R5, R6)                                                  #
    # ------------------------------------------------------------------ #
    print("--- ALERTES (R5, R6) ---")
    rule_alerts  = compute_alerts(records, defect_types, params)
    legacy_alerts = load_legacy("alertes.csv")

    leg_crit  = {r["id"]: r for r in legacy_alerts if r["type"] == "CRITIQUE"}
    leg_recu  = {r["id"]: r for r in legacy_alerts if r["type"] == "RECURRENCE"}
    rule_crit = {r["id"]: r for r in rule_alerts   if r["type"] == "CRITICAL"}
    rule_recu = {r["id"]: r for r in rule_alerts   if r["type"] == "RECURRENCE"}

    for id_ in sorted(rule_crit):
        if id_ not in leg_crit:
            r = rule_crit[id_]
            _tag({"section": "ALERTES/CRITIQUE", "key": id_, "field": "presence",
                  "rule_says": f"CRITICAL ({r['code']} qty={r['qte']})",
                  "legacy_has": "ABSENT"}, "D-new")
    for id_ in sorted(leg_crit):
        if id_ not in rule_crit:
            lr = leg_crit[id_]
            _tag({"section": "ALERTES/CRITIQUE", "key": id_, "field": "presence",
                  "rule_says": "not CRITICAL",
                  "legacy_has": f"CRITIQUE ({lr['code']} qty={lr['qte']})"}, "D-new")

    def _recu_sig(r):
        return (r.get("ligne"), r.get("piece"), r.get("code"))

    rule_rsigs = {_recu_sig(r) for r in rule_recu.values()}
    leg_rsigs  = {_recu_sig(r) for r in leg_recu.values()}
    for k in sorted(rule_rsigs - leg_rsigs):
        _tag({"section": "ALERTES/RECURRENCE", "key": k, "field": "presence",
              "rule_says": "flag expected", "legacy_has": "ABSENT"}, "D7")
    for k in sorted(leg_rsigs - rule_rsigs):
        _tag({"section": "ALERTES/RECURRENCE", "key": k, "field": "presence",
              "rule_says": "no flag", "legacy_has": "present"}, "D7")

    # ------------------------------------------------------------------ #
    # MENS_LIGNES  (R7)                                                  #
    # ------------------------------------------------------------------ #
    print("--- MENS_LIGNES (R7) ---")
    rule_ml = {(r["mois"], r["ligne"]): r
               for r in compute_monthly_lines(records, prod, defect_types, parts, params)}
    leg_ml  = {(r["mois"], r["ligne"]): r
               for r in load_legacy("mens_lignes.csv")}

    for key in sorted(rule_ml):
        rr = rule_ml[key]
        lr = leg_ml.get(key)
        if lr is None:
            _tag({"section": "MENS_LIGNES", "key": key, "field": "row",
                  "rule_says": "present", "legacy_has": "MISSING"}, "?")
            continue
        lbl = f"MENS_LIGNES {key[0]} {key[1]}"
        check(lbl, key, "defauts",  rr["defauts"],  lr["defauts"],  "D1+D4")
        check(lbl, key, "tx_def",   _f2(rr["tx_def"]),  _f2(float(lr["tx_def"])),  "D1+D4")
        check(lbl, key, "rebut",    rr["rebut"],    lr["rebut"],    "D1")
        check(lbl, key, "tx_rebut", _f2(rr["tx_rebut"]), _f2(float(lr["tx_rebut"])), "D1")
        check(lbl, key, "statut",   rr["statut"],   lr["statut"],   "D2")
        check(lbl, key, "cnq_eur",  rr["cnq_eur"],  lr["cnq_eur"],  "D1+D4")
        check(lbl, key, "tendance", rr["tendance"], lr["tendance"], "D1+D4")

    # ------------------------------------------------------------------ #
    # MENS_GLOBAL  (R7)                                                  #
    # ------------------------------------------------------------------ #
    print("--- MENS_GLOBAL (R7) ---")
    rule_mg = {r["mois"]: r
               for r in compute_monthly_global(records, prod, defect_types,
                                               parts, params, rule_alerts)}
    leg_mg  = {r["mois"]: r for r in load_legacy("mens_global.csv")}

    for key in sorted(rule_mg):
        rr = rule_mg[key]
        lr = leg_mg.get(key)
        if lr is None:
            _tag({"section": "MENS_GLOBAL", "key": key, "field": "row",
                  "rule_says": "present", "legacy_has": "MISSING"}, "?")
            continue
        lbl = f"MENS_GLOBAL {key}"
        check(lbl, key, "produit",        rr["produit"],       lr["produit"],       "?")
        check(lbl, key, "defauts",        rr["defauts"],       lr["defauts"],       "D1+D4")
        check(lbl, key, "cnq_eur",        rr["cnq_eur"],       lr["cnq_eur"],       "D1+D4")
        check(lbl, key, "top3",           rr["top3"],          lr["top3"],          "D1+D4")
        check(lbl, key, "nb_critiques",   rr["nb_critiques"],  lr["nb_critiques"],  "D7")
        check(lbl, key, "nb_recurrences", rr["nb_recurrences"], lr["nb_recurrences"], "D7")

    # ------------------------------------------------------------------ #
    # Print results                                                       #
    # ------------------------------------------------------------------ #
    print("\n" + "=" * 70)
    if not MISMATCHES:
        print("No mismatches found.")
        return

    # Group by divergence tag
    by_div = defaultdict(list)
    for m in MISMATCHES:
        by_div[m["div"]].append(m)

    print(f"\nTotal mismatches: {len(MISMATCHES)}\n")
    for div_id in sorted(by_div):
        grp = by_div[div_id]
        print(f"\n[{div_id}] — {len(grp)} mismatch(es)")
        for m in grp:
            print(f"  {m['section']}  key={m['key']}  field={m['field']}")
            print(f"    rule  : {m['rule_says']}")
            print(f"    legacy: {m['legacy_has']}")

    # Summary count table
    print("\n--- SUMMARY: mismatches per divergence row ---")
    print(f"  {'Div':<8} {'Count':>5}")
    print(f"  {'-'*8} {'-'*5}")
    for div_id in sorted(by_div):
        print(f"  {div_id:<8} {len(by_div[div_id]):>5}")
    print(f"  {'TOTAL':<8} {len(MISMATCHES):>5}")

    # CSV dump
    print("\n--- MISMATCH CSV ---")
    print("div,section,key,field,rule_says,legacy_has")
    for m in MISMATCHES:
        def esc(s):
            s = str(s).replace('"', '""')
            return f'"{s}"' if ("," in s or '"' in s) else s
        print(",".join(esc(m[k]) for k in
                       ["div", "section", "key", "field", "rule_says", "legacy_has"]))

if __name__ == "__main__":
    main()
