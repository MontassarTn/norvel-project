# Norvel Composites — Non-Conformance Tracker
## Official business rules (Quality Procedure QP-NC-014, rev. C)

> **Note:** Norvel Composites is a fictional company. This tracker and its data are a synthetic
> reproduction of a typical factory quality tool, built for the IBM Bob 2.0 hackathon.

### 1. Context

Norvel Composites manufactures composite interior panels for aircraft cabins on four production
lines (L1–L4). Line L4 runs a double shift, 7 days a week; L1–L3 run Monday to Friday.

The Non-Conformance Tracker was introduced around 2009 to replace paper defect forms. It records
every non-conformance, computes quality indicators per line, and produces the weekly and monthly
figures reviewed in the quality meetings.

This document describes what the tool is **supposed** to do according to the quality procedure.

### 2. Input data

| File | Content |
|---|---|
| `data/defect_log.csv` | One row per non-conformance: `id, date, line, part_ref, defect_code, qty, disposition, operator, comment` |
| `data/production_log.csv` | One row per line per production day: `date, line, qty_produced` |
| `data/defect_types.csv` | Defect catalogue: code, label, base severity, rework hours per part |
| `data/parts.csv` | Part catalogue: reference, description, unit cost (€) |
| `data/parameters.csv` | Thresholds and rates used by the rules below |

**Disposition values**

- `SCRAP` — part is scrapped
- `REWORK` — part is repaired and returned to production
- `ACCEPT` — part is used as-is under a concession (still a non-conformance)

**Defect catalogue**

| Code | Defect type | Base severity | Rework hours / part |
|---|---|---|---|
| D01 | Surface scratch | Minor | 0.25 |
| D02 | Dent | Minor | 0.5 |
| D03 | Missing fastener | Major | 0.5 |
| D04 | Wrong torque | Major | 0.25 |
| D05 | Dimensional out of tolerance | Major | 1.5 |
| D06 | Delamination | Critical | not reworkable |
| D07 | Contamination (FOD) | Critical | 1.0 |
| D08 | Label / marking error | Minor | 0.1 |

**Part catalogue**

| Part ref | Description | Unit cost (€) |
|---|---|---|
| P-1001 | Sidewall panel | 420 |
| P-1002 | Ceiling panel | 380 |
| P-2001 | Bin door | 610 |
| P-2002 | Bin housing | 1,150 |
| P-3001 | Partition panel | 890 |

**Parameters**

| Parameter | Value |
|---|---|
| Scrap threshold — red | 3.0 % |
| Scrap threshold — orange | 2.0 % |
| Rework labour rate | 45 €/hour |
| Severity escalation quantity | 20 parts |
| Recurrence — minimum occurrences | 3 |
| Recurrence — time window | 7 days |
| Trend — "stable" band | ± 0.5 percentage points |

Weeks are ISO weeks (Monday to Sunday) throughout.

### 3. Business rules

**R1 — Defect rate**
For each line and week:
`defect rate (%) = total defective quantity / total quantity produced × 100`
All dispositions (SCRAP, REWORK, ACCEPT) count as defects.

**R2 — Scrap rate and alert status**
For each line and week:
`scrap rate (%) = scrapped quantity / total quantity produced × 100`

| Scrap rate | Status |
|---|---|
| > 3.0 % | RED |
| > 2.0 % and ≤ 3.0 % | ORANGE |
| ≤ 2.0 % | GREEN |

**R3 — Weekly Pareto**
For each week, defect types are ranked by total defective quantity (descending), with each type's
share and the cumulative percentage. The types needed to reach 80 % cumulative are marked as
priorities (the type that crosses 80 % is included).

**R4 — Cost of non-quality (CNQ)**
For each record:

| Disposition | Cost |
|---|---|
| SCRAP | qty × part unit cost |
| REWORK | qty × rework hours of the defect type × 45 € |
| ACCEPT | 0 € (the record still counts as a defect) |

A D06 (delamination) record declared as REWORK is costed as SCRAP, because delamination is not
reworkable. CNQ is reported per line, per week and per month, in euros.

**R5 — Severity**
Each record takes the base severity of its defect type. If the record quantity is **20 or more**,
the severity is escalated by one level (MINOR → MAJOR, MAJOR → CRITICAL). CRITICAL stays CRITICAL.
Every CRITICAL record generates an alert in the weekly and monthly summaries.

**R6 — Recurrence**
If the same defect code is recorded on the same part reference **3 or more times within any
7-day window** (first and third occurrence at most 6 days apart), the tool raises a RECURRENCE
flag and recommends opening an 8D problem-solving report.

**R7 — Monthly summary**
For each calendar month:

- total quantity produced and total defective quantity
- defect rate and scrap rate per line (same definitions as R1 and R2, over the month)
- total CNQ, per line and overall
- top 3 defect types by quantity
- number of CRITICAL alerts and RECURRENCE flags
- trend versus the previous month, per line: ↑ if the defect rate increased by more than
  0.5 points, ↓ if it decreased by more than 0.5 points, otherwise "stable"

### 4. Outputs

- Weekly report: R1, R2, R3, R4, R5 alerts, R6 flags, per line
- Monthly summary: R7
