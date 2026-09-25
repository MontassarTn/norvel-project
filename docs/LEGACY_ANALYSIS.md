# Legacy VBA Analysis — `LANCER_TOUT`

> **Scope:** This document explains exactly what the legacy Excel macro `LANCER_TOUT` does —
> step by step, from input to output — and catalogues every place where its behaviour
> diverges from the official quality rules R1–R7 defined in `BUSINESS_RULES.md`.
>
> No existing file has been modified. The legacy code and all reference outputs in
> `sorties_legacy/` remain frozen.

---

## 1. Entry Point — `LANCER_TOUT` (Module1.bas, line 12)

`LANCER_TOUT` is the single public macro invoked by the user. It turns off screen
updates, calls six subroutines in fixed order, restores screen updates, and shows a
completion message box.

```
LANCER_TOUT
 ├─ ImportDonnees   — load CSVs into raw sheets
 ├─ CalcHebdo       — compute weekly indicators per line
 ├─ CalcPareto      — compute weekly Pareto ranking
 ├─ CalcAlertes     — emit CRITICAL and RECURRENCE alerts
 ├─ CalcMensuel     — compute monthly per-line and global summaries
 ├─ Colorer         — colour-code the RAP_HEBDO sheet cells
 └─ ExportSorties   — write five output CSVs to sorties_legacy/
```

---

## 2. Step 1 — `ImportDonnees` / `ImportCSV` (Module1.bas, lines 25–52)

### What it does

`ImportDonnees` builds the path `<workbook folder>\data\` and calls `ImportCSV` four
times to load the reference and transactional data into worksheet tabs.

| Call order | Source file              | Target sheet |
|:---:|---|---|
| 1 | `data/defect_log.csv`    | `SAISIE`     |
| 2 | `data/production_log.csv`| `PROD`       |
| 3 | `data/defect_types.csv`  | `REF_DEF`    |
| 4 | `data/parts.csv`         | `REF_PCE`    |

**`parameters.csv` is not loaded.** The import call is present but commented out:

```vba
'ImportCSV chem & "parameters.csv", "PARAM"   ' pas utilise pour l'instant
```

All threshold values (labour rate, escalation quantity, recurrence window, scrap
thresholds) are hardcoded in `Module2.bas` instead of being read from
`data/parameters.csv`.

### How `ImportCSV` works

1. Calls `FeuilleVide(nom)` to obtain a cleared (or newly created) worksheet with all
   cells formatted as text (`@`).
2. Reads the file line by line with `Line Input`.
3. Splits each non-empty line on `,` and writes each field into the corresponding cell.
   The header row becomes row 1; data starts at row 2.

> **Note:** Because every cell is formatted as text (`@`), all values — including dates
> and numbers — are stored as strings. Helper functions `DT()`, `Val()`, and `NumL()`
> convert them at use time.

### Helper functions (Module2.bas, lines 8–21)

| Function | Purpose |
|---|---|
| `DT(s)` | Parses `"YYYY-MM-DD"` strings into a VBA `Date`. |
| `NumL(l)` | Extracts the line number from `"L1"…"L4"` as an integer (1–4). |
| `Txt(x)` | Converts a number to string, replacing `,` with `.` (locale-safe). |
| `Fmt2(x)` | Formats a number to 2 decimal places, replacing `,` with `.`. |

---

## 3. Step 2 — `CalcHebdo` (Module2.bas, lines 63–122)

### Inputs

- Sheet `PROD` — production quantities per line per day.
- Sheet `SAISIE` — defect log (all dispositions).
- Sheet `REF_DEF` — not used directly here; costs delegated to `CoutNC`.

### Internal arrays (indexed `[week 1..53][line 1..4]`)

| Array | Meaning |
|---|---|
| `p(s,l)` | Total parts produced on line `l` in ISO week `s`. |
| `nd(s,l)` | Total defective quantity counted for line `l` in week `s`. |
| `rb(s,l)` | Scrapped quantity for line `l` in week `s`. |
| `cq(s,l)` | Cost of non-quality (€) for line `l` in week `s`. |

### Pass 1 — Accumulate production

Iterates `PROD` from row 2 downward. For each row: parses the date, computes the ISO
week number using `DatePart("ww", d, vbMonday, vbFirstFourDays)`, extracts the line
number, and adds `qty_produced` to `p(s,l)`.

### Pass 2 — Accumulate defects

Iterates `SAISIE` from row 2 downward. For each defect record:

1. If `ProdJour(date, line) = 0` (no production entry for that line on that day), the
   record is **silently skipped** (`GoTo suivant`).
2. Computes ISO week `s` and line `l`.
3. Reads `qty = ws.Cells(r, 6)`.
4. **Line 4 special case:** adds `q / 2` to `nd(s, l)` instead of `q`.
5. If disposition is `SCRAP`, adds `q` (full quantity, not halved) to `rb(s, l)`.
6. Adds `CoutNC(r)` to `cq(s, l)`.

### `CoutNC(r)` — cost of one defect record (Module2.bas, lines 47–61)

Reads `qty`, `defect_code`, and `disposition` from row `r` of `SAISIE`.

| Condition | Cost formula |
|---|---|
| `SCRAP` | `qty × PrixPce(part_ref)` |
| `REWORK` AND `defect_code = "D06"` | `qty × PrixPce(part_ref)` (costed as scrap) |
| `REWORK` (other) | `qty × rework_hours × 45` (hardcoded labour rate) |
| `ACCEPT` | 0 |

`PrixPce` looks up the unit cost in `REF_PCE`; the row offset is computed as
`Val(Mid(c, 2)) + 1`, mapping `"D01"` → row 2, …, `"D08"` → row 9, which assumes the
defect types are loaded in code order starting at row 2 (matching `defect_types.csv`).

### Pass 3 — Write `RAP_HEBDO`

For each `(s, l)` where `p(s, l) > 0`:

- `tx_def = nd(s,l) / p(s,l) × 100`
- `tx_rebut = rb(s,l) / p(s,l) × 100`
- Status:
  - `tx_rebut >= 3` → `"ROUGE"`
  - `tx_rebut > 2` → `"ORANGE"`
  - otherwise → `"VERT"`
- Writes a row: `semaine` (`"2026-Wnn"`), `ligne`, `produit`, `nb_def`, `tx_def`,
  `rebut`, `tx_rebut`, `statut`, `cnq_eur`.

Output written to sheet `RAP_HEBDO`; exported to `sorties_legacy/rap_hebdo.csv`.

---

## 4. Step 3 — `CalcPareto` (Module2.bas, lines 124–177)

### Inputs

- Sheet `SAISIE` — defect log.
- Sheet `REF_DEF` — defect type labels.

### Logic

Uses a fixed-origin week numbering anchored at **2025-12-28** (Sunday before ISO W01
2026):

```vba
d = d - Weekday(d, vbSunday) + 1   ' roll back to Sunday
k = (d - DateSerial(2025, 12, 28)) / 7 + 1
```

Array `q(k, c)` accumulates defective quantity per week-slot `k` and defect code `c`
(1–8). All dispositions contribute.

For each week with at least one defect:
1. Bubble-sort defect codes by descending quantity.
2. For each ranked entry, compute individual percentage and cumulative percentage.
3. If the cumulative **before** adding the current entry is `< 80`, mark as `"X"`.
4. Then add the entry's percentage to the cumulative.

Output: sheet `PARETO` → `sorties_legacy/pareto.csv`.
Columns: `semaine_du` (Sunday date), `rang`, `code`, `libelle`, `qte`, `pct`,
`cumul_pct`, `prioritaire`.

---

## 5. Step 4 — `CalcAlertes` (Module2.bas, lines 179–240)

### Part A — CRITICAL defects

Iterates every row of `SAISIE`. For each record:

1. Reads base severity from `REF_DEF` row `Val(Mid(defect_code, 2)) + 1`, column 3.
2. If `qty >= 20`: escalates MINOR → MAJOR, MAJOR → CRITICAL (CRITICAL stays).
3. If resulting severity = `"CRITICAL"`, writes a `CRITIQUE` alert row to `ALERTES`.

Fields written: `type="CRITIQUE"`, `id`, `date`, `ligne`, `piece`, `code`, `qte`,
`message="defaut critique - prevenir resp. qualite"`.

**No filter on production day.** Records with no matching production entry are still
evaluated for alerts.

### Part B — Recurrence (added 2013, Module2.bas, lines 211–239)

For each line L1–L4, iterates every anchor row `r` that belongs to line `li`:

For each anchor `r`, scans all rows `r2` from `r` onward:
- Checks same line, same `part_ref`, same `defect_code`.
- If `DT(date_r2) - DT(date_r) <= 6`, increments `cnt`.
- If `cnt >= 3`, writes a `RECURRENCE` alert.

The anchor itself is counted (r2 starts at r, so `d2 - d1 = 0 ≤ 6` always counts).
This means a record qualifies when it is the first of at least 3 occurrences of the
same `(line, part_ref, defect_code)` triplet within 7 calendar days (days 0–6 inclusive
= a span of up to 6 days, i.e. a 7-day window).

Output: sheet `ALERTES` → `sorties_legacy/alertes.csv`.

---

## 6. Step 5 — `CalcMensuel` (Module2.bas, lines 242–374)

### Inputs

- Sheet `PROD`, `SAISIE`, `ALERTES` (already populated by `CalcAlertes`).

### Pass 1 — Production by month

Sums `qty_produced` into `p(m, l)` by calendar month.

### Pass 2 — Defects by month

For each SAISIE row, skipping rows where `ProdJour = 0`:
- **Only REWORK and SCRAP** dispositions contribute to `nd(m, l)`. ACCEPT is excluded:
  `If dp <> "ACCEPT" Then nd(m,l) = ...`.
- Line 4: `nd` gets `q / 2` (same halving as weekly).
- Only SCRAP adds to `rb(m, l)`.
- `CoutNC(r)` added to `cq(m, l)` for all dispositions.
- All dispositions contribute to `tq(m, c)` (quantity per defect code, used for top-3).

### Pass 3 — Count alerts

Iterates `ALERTES`, counts `CRITIQUE` and `RECURRENCE` rows per month from the `date`
column.

### Write `MENS_LIGNES`

For each `(m, l)` with production:
- `tx_def = nd(m,l) / p(m,l) × 100`
- `tx_rebut = rb(m,l) / p(m,l) × 100` (note: internally computed as `fr = rb/p`, then
  `fr * 100` for the output and status comparison uses `fr > 0.03` / `fr * 100 > 2`).
- Status: `fr > 0.03` → ROUGE, `fr * 100 > 2` → ORANGE, else VERT.
- Trend vs previous month: if previous month had production, compares `tx_def` of this
  month vs previous. `>0.5 pts` → `"HAUSSE"`, `< -0.5 pts` → `"BAISSE"`, else
  `"STABLE"`. If no previous month data: `"N/A"`.

Output: `sorties_legacy/mens_lignes.csv`.

### Write `MENS_GLOBAL`

For each month with any production, sums across lines for `produit`, `defauts`, `cnq`.
Sorts defect codes by quantity and takes the top 3 as `"D01 / D02 / D03"` format.
Appends alert counts from `nbc(m)` and `nbr(m)`.

Output: `sorties_legacy/mens_global.csv`.

---

## 7. Step 6 — `Colorer` (Module1.bas, lines 69–83)

Reads `tx_rebut` (column 7) of each `RAP_HEBDO` row and sets the background colour of
column 8 (the `statut` cell):

- `> 3` → red (`RGB(255,0,0)`)
- `> 2` → orange (`RGB(255,165,0)`)
- else → green (`RGB(0,176,80)`)

This uses **strict greater-than** for the red threshold, while `CalcHebdo` uses
**`>= 3`** when writing the `statut` text. The cell background and the text status text
can therefore disagree at exactly 3.00 %.

---

## 8. Step 7 — `ExportSorties` (Module1.bas, lines 85–110)

Writes five sheets to `sorties_legacy/` as CSV:
`rap_hebdo`, `pareto`, `alertes`, `mens_lignes`, `mens_global`.

Column count is determined by scanning row 1 for non-empty cells. Row count is
determined by scanning column A for non-empty values. Values are written as-is (no
quoting), joined with `,`, and written with `Print #f` (adds a newline).

---

## 9. Divergence Table — Code vs Rules R1–R7

> **Methodology (v3):** `tools/rules_check.py` (v3) implements every rule purely from
> `BUSINESS_RULES.md`, reading all thresholds from `data/parameters.csv`, with no VBA
> behaviour copied.  It labels every mismatch with the divergence row responsible and
> prints a summary count.  The table below is built entirely from that output.
>
> **Checker summary (218 mismatches):**
>
> | Div tag | Count | Output files affected |
> |:---:|---:|---|
> | D1 | 19 | `rap_hebdo.csv` (all L4 rows) |
> | D1+D4 | 20 | `mens_lignes.csv` (all rows), `mens_global.csv` |
> | D2 | 1 | `rap_hebdo.csv` W27 L3 |
> | D5 | 178 | `pareto.csv` (all weeks) |
> | **Total** | **218** | |

| # | Rule | VBA location | Code behaviour | Rule requirement | Verified proof — exact numbers from `rules_check.py` v3 |
|:---:|:---:|---|---|---|---|
| D1 | R1 | `Module2.bas` line 85 (`CalcHebdo`)<br>`nd(s,l) = nd(s,l) + q/2`<br>and line 267 (`CalcMensuel`) | Defective quantity for **line L4 only** is divided by 2 before accumulation in both the weekly and monthly paths. The scrap accumulator `rb` uses the full quantity (line 89/273), so only the defect count, defect rate, and derived tx\_def are affected — not the scrap rate, scrap count, or CNQ. The 2011 modification is marked `' ne pas toucher`. | R1: *"defect rate (%) = total defective quantity / total quantity produced × 100 — All dispositions … count as defects."* No per-line adjustment to defective quantity is stated. | **Weekly (19 field mismatches):** Every L4 `nb_def` and `tx_def` in `rap_hebdo.csv` is exactly half the rule value. Example: W26 L4 `nb_def = 4.5` (legacy) vs `9.0` (rule); W28 L4 `nb_def = 14` vs `28`; W29 L4 `nb_def = 20` vs `40`. The fractional `4.5` is proof of halving. **Monthly (subset of D1+D4 tag):** June L4 `defauts = 16.5` (legacy) vs `34.0` (rule); July L4 `defauts = 35` vs `76`. |
| D2 | R2 | `Module2.bas` line 102<br>`If txr >= 3 Then st = "ROUGE"` | Scrap rate of **exactly 3.00 %** is classified `ROUGE`. | R2: *"scrap rate > 3.0 % → RED"* — strictly greater than. At exactly 3.00 % the correct status is ORANGE. | **1 field mismatch.** `rap_hebdo.csv` W27 L3: `rebut = 9`, `produit = 300`, `tx_rebut = 3.00` (= 9 / 300 × 100). Legacy `statut = ROUGE`; `rules_check.py` computes `ORANGE`. This is the only row in the dataset where `tx_rebut` lands exactly on the threshold. |
| D3 | R2 | `Module1.bas` line 74<br>`If v > 3 Then ... ROUGE` | The `Colorer` sub colours the `statut` cell background RED only when `tx_rebut > 3` (strict), while `CalcHebdo` writes the text `ROUGE` when `txr >= 3`. At exactly 3.00 % the text and the cell background disagree. | R2 is consistent: one threshold, one status. Having two different thresholds inside the same tool for the same indicator is non-conformant. | **Not detectable in CSVs** — `ExportSorties` exports the text value, not the cell colour. Conflict visible only in the open workbook on W27 L3. No additional checker mismatch beyond D2. |
| D4 | R1 / R7 | `Module2.bas` line 266<br>`If dp <> "ACCEPT" Then nd(m,l) = ...` | In `CalcMensuel`, ACCEPT dispositions are **excluded** from the monthly defect count `nd`. In `CalcHebdo` (line 87–88) there is no such filter — ACCEPT is included. | R1: *"All dispositions (SCRAP, REWORK, ACCEPT) count as defects."* R7 references the same R1 definition. Both weekly and monthly paths must include ACCEPT. | **Part of D1+D4 tag (20 field mismatches).** June L1 ACCEPT records excluded from monthly count: NC-0006 qty 6, NC-0042 qty 1, NC-0051 qty 2 → total 9 missing. `mens_lignes.csv` June L1: `defauts = 51` (legacy) vs `60` (rule), difference = 9. June L2: missing 4 (NC-0053 qty 1, NC-0058 qty 3); June L3: missing 6. `mens_global.csv` June `defauts = 121.5` (legacy) vs `160` (rule). |
| D5 | R3 | `Module2.bas` lines 130–136 (`CalcPareto`)<br>`d = d - Weekday(d, vbSunday) + 1` | The VBA rolls each date back to the **preceding Sunday** using `Weekday(d, vbSunday)` (where vbSunday = 1 makes Sunday = day 1). This creates **Sunday-to-Saturday** weeks. The Pareto bucket for records dated `2026-06-01..06-06` is labelled `2026-05-31` (the Sunday before). Records dated `2026-06-07..06-13` land in `2026-06-07`. Every week boundary is shifted one day earlier than ISO. | R3 references ISO weeks throughout: `BUSINESS_RULES.md §1` states *"Weeks are ISO weeks (Monday to Sunday) throughout."* The Pareto must be grouped by ISO week (Mon–Sun), not Sun–Sat. | **178 field mismatches — the largest divergence.** Rule checker groups by ISO week (Mon–Sun) labelled with the week's Sunday. Legacy output contains a `2026-05-31` week (Sun-to-Sat ending 2026-06-06) that the rule has no equivalent for; rule output has a `2026-07-26` week (Mon-to-Sun 2026-07-20..26) not present in legacy. Every week shows different quantities and rankings. Example: week `2026-06-07` (legacy) ranks D05 first (qty 9) from records 2026-06-07 only; rule's ISO week ending 2026-06-07 (Mon 2026-06-01..Sun 2026-06-07) ranks D01 first (qty 15) from all six days' records. |
| D6 | R1 / R4 | `Module2.bas` line 79 (`CalcHebdo`)<br>`If ProdJour(d, …) = 0 Then GoTo suivant`<br>and line 265 (`CalcMensuel`) | Defect records whose date has **no matching production entry, or a production entry of qty = 0**, are silently skipped in both `CalcHebdo` and `CalcMensuel`. | R1 defines the defect rate as a ratio over all defective parts. It does not say to exclude records on zero-production days. The rules mention no such filter. | **3 field mismatches (tagged D1 in checker — now correctly attributed here).** `production_log.csv` row: `2026-06-17,L2,0`. NC-0040 (L2, 2026-06-17, D02 REWORK, qty 2) is excluded by the VBA. Effect: W25 L2 `nb_def = 2` (legacy) vs `4` (rule); `tx_def = 1.02` vs `2.03`; `cnq_eur = 1161` vs `1206` (difference = 45 = 2 × 0.5 h × 45 €). `mens_lignes.csv` June L2 similarly affected (included in D1+D4 tag). |
| D7 | R4 / all | `Module1.bas` line 31 (comment)<br>`'ImportCSV chem & "parameters.csv", "PARAM"` | `data/parameters.csv` is **never loaded**. Every threshold is hardcoded: labour rate `45` (Module2.bas line 56), red threshold `3` / orange `2` (lines 102–104), escalation qty `20` (line 191), recurrence window `6 days` (line 221). | The parameters file is the declared single source of truth for all thresholds. A change to `data/parameters.csv` has no effect on the macro. | **0 numerical mismatches** in current dataset — the hardcoded values happen to equal `parameters.csv` exactly. The divergence is an architectural maintainability issue: changing a threshold requires editing VBA source, not the CSV. |
| D8 | R6 | `Module2.bas` lines 211–238 (recurrence loop) | The recurrence algorithm scans forward from each anchor row `r` and counts occurrences of the same `(line, part_ref, defect_code)` within `d2 - d1 <= 6`. The anchor itself counts (r2 starts at r). Behaviour when multiple anchors qualify: each qualifying anchor writes a separate RECURRENCE row. | R6: flag when ≥ 3 occurrences of the same triplet fall within any 7-day window. Plain reading: a sliding window producing exactly one flag per qualifying window. | **0 mismatches** in current dataset. The single qualifying triplet (L2 / P-2001 / D03: NC-0085/07-06, NC-0095/07-08, NC-0106/07-10, span = 4 days) triggers exactly one alert from both rule and VBA. The structural risk of duplicate alerts is not triggered by this data. |

---

## 10. Summary of Decisions for Modernisation

> Ordered by impact (D5 is the most pervasive; D1 and D4 affect all monthly figures;
> D6 is newly confirmed as a separate data-exclusion bug).

| # | Divergence | checker mismatches | Recommended action |
|:---:|---|---:|---|
| D1 | L4 defect qty halved in weekly path | 19 | **Business decision required.** The 2011 comment (`' ne pas toucher`) implies intent. If L4's double-shift causes each part to be logged twice, halving normalises the count. If not, remove the halving. The modern version must document the chosen interpretation before coding. |
| D2 | `>= 3` instead of `> 3` for ROUGE threshold | 1 | **Fix**: use strict `> 3.0` per R2. Changes exactly one row (W27 L3: ROUGE → ORANGE). |
| D3 | `Colorer` cell-background threshold inconsistent with text | 0 (CSV) | **Fix**: align background threshold to match the corrected text threshold (`> 3.0`). No CSV impact; workbook-only. |
| D4 | ACCEPT excluded from monthly defect count | part of 20 | **Fix**: include all dispositions in `CalcMensuel`'s `nd` accumulator, matching `CalcHebdo` and R1/R7. Changes every monthly `defauts` and `tx_def`. |
| D5 | Pareto grouped by Sun–Sat instead of ISO Mon–Sun | 178 | **Fix**: group defect records by ISO week (Monday–Sunday). Every week's quantities and rankings change. This is the largest structural divergence in the tool. |
| D6 | Records on zero-production days silently excluded | 3 | **Fix**: remove the `ProdJour = 0` guard in both `CalcHebdo` and `CalcMensuel`. All records in `defect_log.csv` count as defects per R1, regardless of the production entry for that day. |
| D7 | `parameters.csv` not loaded; thresholds hardcoded | 0 | **Fix**: read all thresholds from `parameters.csv` at startup. No numerical change on current data. |
| D8 | Recurrence forward-scan (structural) | 0 | **Monitor**: no output error on current data. Use a proper sliding window in the modern version to prevent duplicate alerts on denser datasets. |

---

## 11. Checker Assumptions (rules silent on these details)

These assumptions are stated in `tools/rules_check.py` and printed at every run.
Where the modern version makes a different choice, the assumption should be revisited.

| ID | Rule(s) | Assumption |
|:---:|:---:|---|
| A1 | R3 | The rule does not specify how to label a Pareto week in the output. The checker uses the ISO week's Sunday date (matching `sorties_legacy/pareto.csv` format) as a purely formatting choice. |
| A2 | R3, R7 | When two defect types have equal quantity, the lower-numbered code ranks first (D01 before D02, etc.). `BUSINESS_RULES.md` does not specify a tie-break. |
| A3 | R1, R7 | All defect records are included in totals regardless of whether a production entry exists for that day. If a (line, period) has defect records but zero total production, no rate row is emitted (rate undefined). |
| A4 | R3 | A negative `qty` in `defect_log.csv` (e.g. NC-0125, a correction entry) is summed algebraically. A code whose net week quantity is ≤ 0 is excluded from the Pareto for that week. |
| A5 | R7 | "Top 3 defect types by quantity" uses total defective quantity across all dispositions for the month (consistent with R1). |
| A6 | R2, R7 | `sorties_legacy/` uses French status labels (`VERT`/`ORANGE`/`ROUGE`). The checker maps R2's English terms to French for direct comparison. |
