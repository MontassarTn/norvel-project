# Divergence Decision Log

> **Purpose:** For each divergence identified between the legacy VBA and rules R1–R7,
> this document records the legacy behaviour, the quantified business impact (from
> `tools/rules_check.py` v4, 218 total mismatches), a recommendation, and a blank
> **Decision** column for the product owner to fill in before modernisation begins.
>
> Source analysis: `docs/LEGACY_ANALYSIS.md §9`  
> Checker: `tools/rules_check.py` v4  
> Golden master: `sorties_legacy/`

---

## Decision Table

| ID | Rule(s) | Legacy behaviour | Business impact (checker mismatches) | Recommendation | Decision |
|:---:|:---:|---|---:|---|---|
| D1 | R1 | `CalcHebdo` (line 85) and `CalcMensuel` (line 267) divide defective quantity by 2 before accumulating into `nd` for **line L4 only**. The scrap accumulator (`rb`) always uses the full quantity, so only `nb_def`, `tx_def`, and derived monthly `defauts`/`tx_def` are affected. Marked `' ne pas toucher` since 2011. | **16 weekly + 4 monthly + 2 global = 22** fields (tagged D1, D1+D4, D1+D4+D6). Every L4 `nb_def` and `tx_def` in `rap_hebdo.csv` is exactly half the rule value (e.g. W26 L4: legacy `nb_def=4.5` vs rule `9.0`). | **Business decision required** before coding. If L4 registers each part twice (double-shift or double-scan), halving normalises the count and the legacy behaviour should be preserved as a documented exception. If parts are registered once, halving is an error and must be removed. The `' ne pas toucher` comment must not be treated as a substitute for a documented decision. | |
| D2 | R2 | `CalcHebdo` (line 102) classifies scrap rate **`>= 3 %`** as ROUGE. R2 specifies **`> 3 %`** (strictly greater). At exactly 3.00 % the legacy reports ROUGE where the rule requires ORANGE. | **1 field** — `rap_hebdo.csv` W27 L3 `statut`. `rebut=9`, `produit=300`, `tx_rebut=3.00 %`. Legacy: `ROUGE`; rule: `ORANGE`. | **Fix**: change the threshold in the modern version to `tx_rebut > 3.0` (strict). Impact is narrow (one row in current data) but is a clear non-conformance with the declared quality procedure. | |
| D3 | R2 | `Colorer` (Module1.bas line 74) uses `> 3` (strict) for the RED background, while `CalcHebdo` uses `>= 3` for the text `"ROUGE"`. At exactly 3.00 % the cell text and cell background disagree inside the open workbook. | **0 CSV mismatches** — the text value exported by `ExportSorties` is the one from `CalcHebdo` (affected by D2). This is a workbook-only visual inconsistency. | **Fix as part of D2**: the modern version has a single threshold (`> 3.0`) applied uniformly to both the status label and any colour-coding. No separate code change from D2. | |
| D4 | R1, R7 | `CalcMensuel` (line 266) adds an `If dp <> "ACCEPT"` guard before accumulating `nd`, so ACCEPT dispositions are **excluded from the monthly defect count**. `CalcHebdo` has no such guard — ACCEPT is included weekly. The asymmetry causes the monthly `defauts` and `tx_def` to be systematically lower than the weekly path implies. | **8 monthly-only + 4 L4 compound + 4 L2 compound = 16** field mismatches (tagged D4, D1+D4, D4+D6). Example: June L1 `defauts=51` (legacy) vs `60` (rule); difference = 9 ACCEPT units excluded (NC-0006 qty 6, NC-0042 qty 1, NC-0051 qty 2). CNQ is unaffected (ACCEPT cost = 0). | **Fix**: remove the `dp <> "ACCEPT"` guard from `CalcMensuel`'s `nd` accumulator. All three dispositions must count as defects in both weekly and monthly paths per R1. | |
| D5 | R3 | `CalcPareto` (lines 130–136) rolls each date back to the preceding **Sunday** (`Weekday(d, vbSunday)`), creating **Sunday-to-Saturday** week buckets. ISO weeks run Monday–Sunday. Every week boundary is one day earlier than the rule requires, producing a spurious `2026-05-31` week in legacy output and missing the `2026-07-26` week that the rule produces. | **178 field mismatches** — the largest single divergence, affecting every row and field of `pareto.csv`. Week labels, defect code groupings, quantities, ranks, percentages, and the `prioritaire` flag all differ. Example: legacy week `2026-06-07` ranks D05 first (qty 9, records from that Sunday only); rule ISO week ending `2026-06-07` (Mon 2026-06-01–Sun 2026-06-07) ranks D01 first (qty 15). | **Fix**: group defect records by ISO week (Monday–Sunday). Use `isocalendar()` or equivalent. Every historical Pareto figure will change — stakeholders relying on the legacy trend should be informed before go-live. | |
| D6 | R1, R4 | `CalcHebdo` (line 79) and `CalcMensuel` (line 265) call `ProdJour(date, line)` and skip any defect record whose date has no production entry **or a production qty of 0**. The rule does not authorise this filter. | **5 field mismatches** (tagged D6 and as part of D4+D6 / D1+D4+D6 compounds). Confirmed instance: NC-0040 (L2, 2026-06-17, D02 REWORK qty=2; `production_log.csv` has L2 qty=0 that day). Effects: W25 L2 `nb_def`: legacy `2` vs rule `4`; `tx_def`: `1.02` vs `2.03`; `cnq_eur` short by 45 € (= 2 × 0.5 h × 45 €). June L2 and June global `cnq_eur` similarly short by 45 €. | **Fix**: remove the `ProdJour = 0` guard in both paths. All records in `defect_log.csv` are valid defect events regardless of the production log entry for that day. If production truly was zero and parts were still scrapped or reworked, that is exactly the situation the quality team needs to see. | |
| D7 | R4, all | `data/parameters.csv` is **never loaded** (`Module1.bas` line 31, commented out). Every threshold — labour rate (45 €/h), red threshold (3 %), orange threshold (2 %), escalation qty (20), recurrence window (6 days) — is hardcoded in `Module2.bas`. Changing `parameters.csv` has no effect on the macro. | **0 numerical mismatches** on current data: the hardcoded values happen to equal `parameters.csv` exactly. The risk is architectural: any threshold change requires a VBA source edit, which bypasses the declared configuration mechanism and may go unnoticed in audits. | **Fix**: read all thresholds from `parameters.csv` at startup. No numerical change on current data; eliminates future configuration drift. | |
| D8 | R6 | `CalcAlertes` recurrence loop (lines 211–238) scans forward from each anchor row `r` and counts occurrences of the same `(line, part_ref, defect_code)` triplet within `d2 - d1 <= 6`. Every qualifying anchor writes a separate `RECURRENCE` row, which could produce duplicate alerts for the same window on denser datasets. | **0 mismatches** on current data. The single qualifying triplet (L2 / P-2001 / D03: NC-0085 2026-07-06, NC-0095 2026-07-08, NC-0106 2026-07-10, span = 4 days) produces exactly one alert from both legacy and rule. Duplicate risk not triggered. | **Monitor / Fix in modernisation**: replace the per-anchor forward scan with a proper sliding-window algorithm that emits exactly one flag per qualifying window, regardless of data density. No immediate impact on existing reports. | |

---

## Notes on Zero-Mismatch Divergences

**D3** (Colorer threshold inconsistency) and **D7** (parameters.csv not loaded) produce
**0 CSV mismatches** on the current dataset, but both are genuine non-conformances:

- D3 is invisible in CSVs because `ExportSorties` exports the text value from
  `CalcHebdo`, not the cell background colour. The workbook shows a conflicting visual
  state at W27 L3.
- D7 poses a configuration-governance risk: a threshold change in `parameters.csv` would
  be silently ignored.

**D8** (recurrence forward-scan) also produces 0 mismatches because the current dataset
has only one qualifying triplet, which does not expose the duplicate-alert path.

---

## Compound Tags Explained

Some mismatches in `tools/rules_check.py` v4 output carry compound tags where two or
three bugs interact on the same field:

| Compound tag | Meaning |
|---|---|
| `D1+D4` | L4 monthly `defauts`/`tx_def`: halving (D1) and ACCEPT exclusion (D4) both reduce the count below the rule value |
| `D4+D6` | L2 monthly `defauts`/`tx_def`: ACCEPT exclusion (D4) and NC-0040 zero-prod exclusion (D6) both reduce the count |
| `D1+D4+D6` | Global monthly `defauts`: all three bugs (L4 halving + ACCEPT exclusion + NC-0040 exclusion) converge on the same field |

Each compound-tag row must have **all named bugs fixed together** before the field will
match the rule value.

---

*Document generated from `tools/rules_check.py` v4 output.  
Checker assumptions (A1–A6) are documented in `docs/LEGACY_ANALYSIS.md §11`.*
