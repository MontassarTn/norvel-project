"""
tools/rules_check.py
====================
Implements rules R1–R7 exactly as BUSINESS_RULES.md states them, independently
of the VBA code.  Reads data/ CSVs, computes every indicator, then compares the
results with sorties_legacy/ and prints a full diff report.

Run from the repository root:
    python tools/rules_check.py
"""

import csv
import math
import os
import sys
from collections import defaultdict
from datetime import date, timedelta

# ---------------------------------------------------------------------------
# Paths (relative to repo root)
# ---------------------------------------------------------------------------
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")
LEGACY = os.path.join(ROOT, "sorties_legacy")

# ---------------------------------------------------------------------------
# Parameters – read from data/parameters.csv (as BUSINESS_RULES.md requires)
# ---------------------------------------------------------------------------
def load_parameters():
    params = {}
    with open(os.path.join(DATA, "parameters.csv"), newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            params[row["parameter"]] = float(row["value"])
    return params

# ---------------------------------------------------------------------------
# Load reference tables
# ---------------------------------------------------------------------------
def load_defect_types():
    """Returns dict: code -> {label, base_severity, rework_hours}"""
    dt = {}
    with open(os.path.join(DATA, "defect_types.csv"), newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            dt[row["code"]] = {
                "label": row["label"],
                "base_severity": row["base_severity"],
                "rework_hours": float(row["rework_hours_per_part"]) if row["rework_hours_per_part"].strip() else None,
            }
    return dt

def load_parts():
    """Returns dict: part_ref -> unit_cost_eur"""
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
                "id": row["id"],
                "date": date.fromisoformat(row["date"]),
                "line": row["line"],
                "part_ref": row["part_ref"],
                "defect_code": row["defect_code"],
                "qty": int(row["qty"]),
                "disposition": row["disposition"],
                "operator": row["operator"],
                "comment": row["comment"],
            })
    return records

def load_production_log():
    prod = {}  # (date, line) -> qty_produced
    with open(os.path.join(DATA, "production_log.csv"), newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            d = date.fromisoformat(row["date"])
            prod[(d, row["line"])] = int(row["qty_produced"])
    return prod

# ---------------------------------------------------------------------------
# ISO week helper
# ---------------------------------------------------------------------------
def iso_week(d: date) -> int:
    return d.isocalendar()[1]

def iso_year_week(d: date) -> str:
    return f"{d.isocalendar()[0]}-W{d.isocalendar()[1]:02d}"

# ---------------------------------------------------------------------------
# R1 – Defect rate
# R2 – Scrap rate and alert status
# R4 – Cost of non-quality (CNQ)
# => Computed together per (week, line) and (month, line)
# ---------------------------------------------------------------------------
def r4_cnq(rec, defect_types, parts, labour_rate):
    """Cost of non-quality for one defect record, per R4."""
    qty = rec["qty"]
    code = rec["defect_code"]
    disp = rec["disposition"]
    if disp == "SCRAP":
        return round(qty * parts[rec["part_ref"]], 0)
    elif disp == "REWORK":
        if code == "D06":
            # D06 not reworkable → cost as SCRAP
            return round(qty * parts[rec["part_ref"]], 0)
        else:
            rh = defect_types[code]["rework_hours"]
            return round(qty * rh * labour_rate, 0)
    else:  # ACCEPT
        return 0.0

def r2_status(scrap_rate_pct, red_threshold, orange_threshold):
    """R2: strictly >3 = ROUGE, strictly >2 = ORANGE, else VERT.
    Uses French column names to match sorties_legacy/ output.
    """
    if scrap_rate_pct > red_threshold:
        return "ROUGE"
    elif scrap_rate_pct > orange_threshold:
        return "ORANGE"
    else:
        return "VERT"

# ---------------------------------------------------------------------------
# R5 – Severity escalation
# ---------------------------------------------------------------------------
SEVERITY_ORDER = ["MINOR", "MAJOR", "CRITICAL"]

def r5_severity(base_severity, qty, escalation_qty):
    idx = SEVERITY_ORDER.index(base_severity)
    if qty >= escalation_qty and idx < len(SEVERITY_ORDER) - 1:
        idx += 1
    return SEVERITY_ORDER[idx]

# ---------------------------------------------------------------------------
# Weekly computations (R1, R2, R3, R4, R5)
# ---------------------------------------------------------------------------
def compute_weekly(records, prod, defect_types, parts, params):
    labour_rate = params["rework_rate_eur_per_hour"]
    esc_qty = int(params["severity_escalation_qty"])
    red_t = params["scrap_threshold_red_pct"]
    org_t = params["scrap_threshold_orange_pct"]

    # Aggregation per (week_label, line)
    prod_week = defaultdict(float)       # total produced
    defect_week = defaultdict(float)     # total defective qty (R1: all dispositions)
    scrap_week = defaultdict(float)      # total scrapped qty (R2)
    cnq_week = defaultdict(float)        # CNQ (R4)

    # Production per ISO week/line
    for (d, line), qty in prod.items():
        key = (iso_year_week(d), line)
        prod_week[key] += qty

    # Defects
    for rec in records:
        d = rec["date"]
        line = rec["line"]
        key = (iso_year_week(d), line)
        if prod.get((d, line), 0) == 0:
            continue  # no production on that day → skip (matches VBA ProdJour check)
        qty = rec["qty"]
        prod_week[key]  # ensure key exists even if 0 prod on week – already done above
        defect_week[key] += qty              # R1: all dispositions
        if rec["disposition"] == "SCRAP":
            scrap_week[key] += qty
        cnq_week[key] += r4_cnq(rec, defect_types, parts, labour_rate)

    # Build rows sorted by (week, line)
    all_week_keys = sorted(
        {k for k in prod_week if prod_week[k] > 0},
        key=lambda k: (k[0], k[1])
    )
    rows = []
    for key in all_week_keys:
        wk, line = key
        p = prod_week[key]
        nd = defect_week[key]
        rb = scrap_week[key]
        cq = cnq_week[key]
        tx_def = nd / p * 100
        tx_rebut = rb / p * 100
        status = r2_status(tx_rebut, red_t, org_t)
        rows.append({
            "semaine": wk,
            "ligne": line,
            "produit": p,
            "nb_def": nd,
            "tx_def": round(tx_def, 2),
            "rebut": rb,
            "tx_rebut": round(tx_rebut, 2),
            "statut": status,
            "cnq_eur": round(cq, 0),
        })
    return rows

# ---------------------------------------------------------------------------
# R3 – Weekly Pareto
# ---------------------------------------------------------------------------
def compute_pareto(records, defect_types, prod):
    # Group by ISO week (Sunday-start label matching VBA output)
    # VBA uses the Sunday of the week as semaine_du
    def sunday_of(d: date) -> date:
        return d - timedelta(days=d.weekday() + 1) if d.weekday() != 6 else d

    qty_per_week = defaultdict(lambda: defaultdict(int))
    for rec in records:
        # All dispositions (R3 ranks by defective qty, no disposition filter)
        sun = sunday_of(rec["date"])
        qty_per_week[sun][rec["defect_code"]] += rec["qty"]

    rows = []
    for sun in sorted(qty_per_week.keys()):
        week_data = qty_per_week[sun]
        # total uses only positive quantities (negative corrections may zero a code)
        total = sum(q for q in week_data.values() if q > 0)
        if total == 0:
            continue
        # VBA only includes codes where q > 0; negative-qty corrections can zero a code
        # Tie-break: lower code number first (matches VBA bubble sort on 1..8 array)
        sorted_codes = sorted(
            [(c, q) for c, q in week_data.items() if q > 0],
            key=lambda x: (-x[1], int(x[0][1:]))
        )
        cumul = 0.0
        for rank, (code, qty) in enumerate(sorted_codes, 1):
            pct = qty / total * 100
            cumul += pct
            # R3: the type that CROSSES 80% is INCLUDED → flag is set after adding
            priority = "X" if cumul <= 80 or (cumul > 80 and (cumul - pct) < 80) else ""
            rows.append({
                "semaine_du": sun.isoformat(),
                "rang": rank,
                "code": code,
                "libelle": defect_types[code]["label"],
                "qte": qty,
                "pct": round(pct, 2),
                "cumul_pct": round(cumul, 2),
                "prioritaire": priority,
            })
    return rows

# ---------------------------------------------------------------------------
# R5 + R6 – Alerts
# ---------------------------------------------------------------------------
def compute_alerts(records, defect_types, params):
    esc_qty = int(params["severity_escalation_qty"])
    rec_min = int(params["recurrence_min_count"])
    rec_win = int(params["recurrence_window_days"])

    alerts = []

    # R5 – CRITICAL alerts
    for rec in records:
        base_sev = defect_types[rec["defect_code"]]["base_severity"]
        eff_sev = r5_severity(base_sev, rec["qty"], esc_qty)
        if eff_sev == "CRITICAL":
            alerts.append({
                "type": "CRITICAL",
                "id": rec["id"],
                "date": rec["date"].isoformat(),
                "ligne": rec["line"],
                "piece": rec["part_ref"],
                "code": rec["defect_code"],
                "qte": rec["qty"],
                "message": "defaut critique - prevenir resp. qualite",
            })

    # R6 – Recurrence: sliding window
    # Group by (line, part_ref, defect_code)
    triplet_dates = defaultdict(list)
    for rec in records:
        key = (rec["line"], rec["part_ref"], rec["defect_code"])
        triplet_dates[key].append((rec["date"], rec["id"]))

    flagged_triplets = set()  # avoid duplicate alerts for same triplet/window
    for key, occurrences in triplet_dates.items():
        occurrences.sort()
        for i, (d_start, id_start) in enumerate(occurrences):
            # Count how many fall within [d_start, d_start + rec_win - 1]
            window_end = d_start + timedelta(days=rec_win - 1)
            count = sum(1 for (d2, _) in occurrences if d_start <= d2 <= window_end)
            if count >= rec_min:
                flag_key = (key, d_start)
                if flag_key not in flagged_triplets:
                    flagged_triplets.add(flag_key)
                    alerts.append({
                        "type": "RECURRENCE",
                        "id": id_start,
                        "date": d_start.isoformat(),
                        "ligne": key[0],
                        "piece": key[1],
                        "code": key[2],
                        "qte": count,
                        "message": "recurrence 7j - ouvrir 8D",
                    })

    # Sort: CRITICAL first (in data order), then RECURRENCE
    crit = [a for a in alerts if a["type"] == "CRITICAL"]
    recu = [a for a in alerts if a["type"] == "RECURRENCE"]
    return crit + recu

# ---------------------------------------------------------------------------
# R7 – Monthly summary
# ---------------------------------------------------------------------------
def compute_monthly_lines(records, prod, defect_types, parts, params):
    labour_rate = params["rework_rate_eur_per_hour"]
    red_t = params["scrap_threshold_red_pct"]
    org_t = params["scrap_threshold_orange_pct"]

    prod_month = defaultdict(float)
    defect_month = defaultdict(float)
    scrap_month = defaultdict(float)
    cnq_month = defaultdict(float)

    for (d, line), qty in prod.items():
        key = (d.year, d.month, line)
        prod_month[key] += qty

    for rec in records:
        d = rec["date"]
        line = rec["line"]
        if prod.get((d, line), 0) == 0:
            continue
        key = (d.year, d.month, line)
        qty = rec["qty"]
        defect_month[key] += qty          # R1/R7: ALL dispositions
        if rec["disposition"] == "SCRAP":
            scrap_month[key] += qty
        cnq_month[key] += r4_cnq(rec, defect_types, parts, labour_rate)

    # Sort keys
    all_keys = sorted({k for k in prod_month if prod_month[k] > 0})
    rows = []
    prev_tx_def = {}  # (year, month-1, line) -> tx_def
    for key in all_keys:
        yr, mo, line = key
        p = prod_month[key]
        nd = defect_month[key]
        rb = scrap_month[key]
        cq = cnq_month[key]
        tx_def = nd / p * 100
        tx_rebut = rb / p * 100
        status = r2_status(tx_rebut, red_t, org_t)
        # Trend vs previous month
        prev_key = (yr, mo - 1, line) if mo > 1 else (yr - 1, 12, line)
        if prev_key in prev_tx_def:
            diff = tx_def - prev_tx_def[prev_key]
            if diff > 0.5:
                trend = "HAUSSE"
            elif diff < -0.5:
                trend = "BAISSE"
            else:
                trend = "STABLE"
        else:
            trend = "N/A"
        prev_tx_def[key] = tx_def
        rows.append({
            "mois": f"{yr}-{mo:02d}",
            "ligne": line,
            "produit": p,
            "defauts": nd,
            "tx_def": round(tx_def, 2),
            "rebut": rb,
            "tx_rebut": round(tx_rebut, 2),
            "statut": status,
            "cnq_eur": round(cq, 0),
            "tendance": trend,
        })
    return rows

def compute_monthly_global(records, prod, defect_types, parts, params, alert_rows):
    labour_rate = params["rework_rate_eur_per_hour"]

    prod_month = defaultdict(float)
    defect_month = defaultdict(float)
    cnq_month = defaultdict(float)
    tq_month = defaultdict(lambda: defaultdict(int))  # month -> code -> qty
    nbc = defaultdict(int)
    nbr = defaultdict(int)

    for (d, line), qty in prod.items():
        prod_month[(d.year, d.month)] += qty

    for rec in records:
        d = rec["date"]
        line = rec["line"]
        if prod.get((d, line), 0) == 0:
            continue
        mo = (d.year, d.month)
        qty = rec["qty"]
        defect_month[mo] += qty
        cnq_month[mo] += r4_cnq(rec, defect_types, parts, labour_rate)
        tq_month[mo][rec["defect_code"]] += qty

    for a in alert_rows:
        d = date.fromisoformat(a["date"])
        mo = (d.year, d.month)
        if a["type"] == "CRITICAL":
            nbc[mo] += 1
        else:
            nbr[mo] += 1

    rows = []
    for mo in sorted({k for k in prod_month if prod_month[k] > 0}):
        yr, m = mo
        top3_sorted = sorted(tq_month[mo].items(), key=lambda x: -x[1])[:3]
        top3 = " / ".join(c for c, _ in top3_sorted)
        rows.append({
            "mois": f"{yr}-{m:02d}",
            "produit": prod_month[mo],
            "defauts": defect_month[mo],
            "cnq_eur": round(cnq_month[mo], 0),
            "top3": top3,
            "nb_critiques": nbc[mo],
            "nb_recurrences": nbr[mo],
        })
    return rows

# ---------------------------------------------------------------------------
# Load legacy CSV files
# ---------------------------------------------------------------------------
def load_legacy(name):
    path = os.path.join(LEGACY, name)
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))

# ---------------------------------------------------------------------------
# Comparison helpers
# ---------------------------------------------------------------------------
def fmtnum(x):
    """Format a float the same way the VBA Fmt2/Txt does: 2 dp or integer."""
    if x == int(x):
        return str(int(x))
    return f"{x:.2f}"

def _f2(x):
    return f"{float(x):.2f}"

MISMATCHES = []

def check(section, key, field, expected, actual):
    exp_s = str(expected).strip()
    act_s = str(actual).strip()
    # Normalise numeric comparison: strip trailing zeros
    try:
        exp_f = float(exp_s)
        act_f = float(act_s)
        match = math.isclose(exp_f, act_f, abs_tol=0.005)
    except ValueError:
        match = exp_s == act_s
    if not match:
        MISMATCHES.append({
            "section": section,
            "key": key,
            "field": field,
            "rule_says": exp_s,
            "legacy_has": act_s,
        })

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    params = load_parameters()
    defect_types = load_defect_types()
    parts = load_parts()
    records = load_defect_log()
    prod = load_production_log()

    print("=== rules_check.py — R1–R7 vs sorties_legacy/ ===\n")
    print(f"Parameters loaded: {params}\n")

    # ---- Weekly report ------------------------------------------------
    print("--- RAP_HEBDO (R1, R2, R4) ---")
    rule_hebdo = compute_weekly(records, prod, defect_types, parts, params)
    legacy_hebdo = load_legacy("rap_hebdo.csv")

    # Index legacy by (semaine, ligne)
    leg_h = {(r["semaine"], r["ligne"]): r for r in legacy_hebdo}
    rule_h = {(r["semaine"], r["ligne"]): r for r in rule_hebdo}

    for key in sorted(rule_h.keys()):
        rr = rule_h[key]
        lr = leg_h.get(key)
        if lr is None:
            MISMATCHES.append({"section": "RAP_HEBDO", "key": str(key), "field": "ALL",
                                "rule_says": "row present", "legacy_has": "MISSING"})
            continue
        label = f"RAP_HEBDO {key[0]} {key[1]}"
        check(label, key, "nb_def",    rr["nb_def"],    lr["nb_def"])
        check(label, key, "tx_def",    _f2(rr["tx_def"]), _f2(float(lr["tx_def"])))
        check(label, key, "rebut",     rr["rebut"],     lr["rebut"])
        check(label, key, "tx_rebut",  _f2(rr["tx_rebut"]), _f2(float(lr["tx_rebut"])))
        check(label, key, "statut",    rr["statut"],    lr["statut"])
        check(label, key, "cnq_eur",   rr["cnq_eur"],   lr["cnq_eur"])

    for key in sorted(leg_h.keys()):
        if key not in rule_h:
            MISMATCHES.append({"section": "RAP_HEBDO", "key": str(key), "field": "ALL",
                                "rule_says": "EXTRA in legacy", "legacy_has": str(leg_h[key])})

    # ---- Pareto -------------------------------------------------------
    print("--- PARETO (R3) ---")
    rule_pareto = compute_pareto(records, defect_types, prod)
    legacy_pareto = load_legacy("pareto.csv")

    # Index by (semaine_du, rang)
    leg_p = {(r["semaine_du"], r["rang"]): r for r in legacy_pareto}
    rule_p = {(r["semaine_du"], str(r["rang"])): r for r in rule_pareto}

    for key in sorted(rule_p.keys()):
        rr = rule_p[key]
        lr = leg_p.get(key)
        if lr is None:
            MISMATCHES.append({"section": "PARETO", "key": str(key), "field": "ALL",
                                "rule_says": "row present", "legacy_has": "MISSING"})
            continue
        label = f"PARETO {key[0]} rank {key[1]}"
        check(label, key, "code",       rr["code"],      lr["code"])
        check(label, key, "qte",        rr["qte"],       lr["qte"])
        check(label, key, "pct",        _f2(rr["pct"]),  _f2(float(lr["pct"])))
        check(label, key, "cumul_pct",  _f2(rr["cumul_pct"]), _f2(float(lr["cumul_pct"])))
        check(label, key, "prioritaire", rr["prioritaire"], lr["prioritaire"])

    # ---- Alerts -------------------------------------------------------
    print("--- ALERTES (R5, R6) ---")
    rule_alerts = compute_alerts(records, defect_types, params)
    legacy_alerts = load_legacy("alertes.csv")

    # Compare by id and type
    leg_a_crit = {r["id"]: r for r in legacy_alerts if r["type"] == "CRITIQUE"}
    leg_a_recu = {r["id"]: r for r in legacy_alerts if r["type"] == "RECURRENCE"}
    rule_a_crit = {r["id"]: r for r in rule_alerts if r["type"] == "CRITICAL"}
    rule_a_recu = {r["id"]: r for r in rule_alerts if r["type"] == "RECURRENCE"}

    # CRITICAL alerts present in rule but not legacy
    for id_ in sorted(rule_a_crit):
        if id_ not in leg_a_crit:
            r = rule_a_crit[id_]
            MISMATCHES.append({"section": "ALERTES/CRITIQUE", "key": id_, "field": "presence",
                                "rule_says": f"CRITICAL alert for {id_} ({r['code']} qty={r['qte']})",
                                "legacy_has": "ABSENT"})
    # CRITICAL in legacy but not in rule
    for id_ in sorted(leg_a_crit):
        if id_ not in rule_a_crit:
            lr = leg_a_crit[id_]
            MISMATCHES.append({"section": "ALERTES/CRITIQUE", "key": id_, "field": "presence",
                                "rule_says": "NOT a CRITICAL alert",
                                "legacy_has": f"CRITIQUE alert present ({lr['code']} qty={lr['qte']})"})

    # RECURRENCE: compare by (ligne, piece, code)
    def recu_key(r):
        return (r.get("ligne", r.get("ligne")), r.get("piece", r.get("piece")), r.get("code", r.get("code")))

    rule_recu_keys = {recu_key(r) for r in rule_a_recu.values()}
    leg_recu_keys  = {recu_key(r) for r in leg_a_recu.values()}

    for k in sorted(rule_recu_keys - leg_recu_keys):
        MISMATCHES.append({"section": "ALERTES/RECURRENCE", "key": str(k), "field": "presence",
                            "rule_says": "RECURRENCE flag expected", "legacy_has": "ABSENT"})
    for k in sorted(leg_recu_keys - rule_recu_keys):
        MISMATCHES.append({"section": "ALERTES/RECURRENCE", "key": str(k), "field": "presence",
                            "rule_says": "NO recurrence flag", "legacy_has": "RECURRENCE present"})

    # ---- Monthly lines ------------------------------------------------
    print("--- MENS_LIGNES (R7) ---")
    rule_ml = compute_monthly_lines(records, prod, defect_types, parts, params)
    legacy_ml = load_legacy("mens_lignes.csv")

    leg_ml = {(r["mois"], r["ligne"]): r for r in legacy_ml}
    rule_ml_d = {(r["mois"], r["ligne"]): r for r in rule_ml}

    for key in sorted(rule_ml_d.keys()):
        rr = rule_ml_d[key]
        lr = leg_ml.get(key)
        if lr is None:
            MISMATCHES.append({"section": "MENS_LIGNES", "key": str(key), "field": "ALL",
                                "rule_says": "row present", "legacy_has": "MISSING"})
            continue
        label = f"MENS_LIGNES {key[0]} {key[1]}"
        check(label, key, "defauts",   rr["defauts"],   lr["defauts"])
        check(label, key, "tx_def",    _f2(rr["tx_def"]), _f2(float(lr["tx_def"])))
        check(label, key, "rebut",     rr["rebut"],     lr["rebut"])
        check(label, key, "tx_rebut",  _f2(rr["tx_rebut"]), _f2(float(lr["tx_rebut"])))
        check(label, key, "statut",    rr["statut"],    lr["statut"])
        check(label, key, "cnq_eur",   rr["cnq_eur"],   lr["cnq_eur"])
        check(label, key, "tendance",  rr["tendance"],  lr["tendance"])

    # ---- Monthly global -----------------------------------------------
    print("--- MENS_GLOBAL (R7) ---")
    rule_alerts_for_global = compute_alerts(records, defect_types, params)
    rule_mg = compute_monthly_global(records, prod, defect_types, parts, params,
                                     rule_alerts_for_global)
    legacy_mg = load_legacy("mens_global.csv")

    leg_mg = {r["mois"]: r for r in legacy_mg}
    rule_mg_d = {r["mois"]: r for r in rule_mg}

    for key in sorted(rule_mg_d.keys()):
        rr = rule_mg_d[key]
        lr = leg_mg.get(key)
        if lr is None:
            MISMATCHES.append({"section": "MENS_GLOBAL", "key": key, "field": "ALL",
                                "rule_says": "row present", "legacy_has": "MISSING"})
            continue
        label = f"MENS_GLOBAL {key}"
        check(label, key, "produit",        rr["produit"],      lr["produit"])
        check(label, key, "defauts",        rr["defauts"],      lr["defauts"])
        check(label, key, "cnq_eur",        rr["cnq_eur"],      lr["cnq_eur"])
        check(label, key, "top3",           rr["top3"],         lr["top3"])
        check(label, key, "nb_critiques",   rr["nb_critiques"], lr["nb_critiques"])
        check(label, key, "nb_recurrences", rr["nb_recurrences"], lr["nb_recurrences"])

    # ---- Print results ------------------------------------------------
    print("\n" + "=" * 70)
    if not MISMATCHES:
        print("No mismatches found.")
        return

    print(f"Found {len(MISMATCHES)} mismatch(es):\n")
    current_section = None
    for m in MISMATCHES:
        if m["section"] != current_section:
            current_section = m["section"]
            print(f"\n[{current_section}]")
        print(f"  key={m['key']}  field={m['field']}")
        print(f"    rule  : {m['rule_says']}")
        print(f"    legacy: {m['legacy_has']}")

    print("\n--- MISMATCH SUMMARY (CSV-friendly) ---")
    print("section,key,field,rule_says,legacy_has")
    for m in MISMATCHES:
        def esc(s):
            s = str(s).replace('"', '""')
            return f'"{s}"' if "," in s or '"' in s else s
        print(",".join(esc(m[k]) for k in ["section", "key", "field", "rule_says", "legacy_has"]))

if __name__ == "__main__":
    main()
