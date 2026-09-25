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

> **Methodology:** Every row in this table was verified by running `tools/rules_check.py`,
> which re-implements R1–R7 independently from the VBA (reading `data/parameters.csv`
> for all thresholds) and compares its output against every field in `sorties_legacy/`.
> Rows D1 and D2 produce measurable output differences; D3 is a visual-only divergence
> in the workbook (not in the CSVs); D5 and D6 are structural faults that do not produce
> numerical differences in this specific dataset.

| # | Rule | VBA location | Code behaviour | Rule requirement | Verified proof (numbers from data/ and sorties\_legacy/) |
|:---:|:---:|---|---|---|---|
| D1 | R1 | `Module2.bas` line 85<br>`nd(s,l) = nd(s,l) + q/2` | Defective quantity for **line L4** is divided by 2 before being added to the defect count — in both weekly (`CalcHebdo`) and monthly (`CalcMensuel`, line 267) paths. The same halving is applied in the scrap rate denominator for L4 ACCEPT records when they flow through `nd`, but the scrap accumulator `rb` uses the full quantity (line 89). | R1 requires `total defective quantity / total quantity produced` with no per-line adjustment. The division by 2 is an undocumented modification introduced in 2011 (comment: `' modif JLB 2011 - ne pas toucher`). | **Weekly:** `rap_hebdo.csv` W26 L4: `nb_def = 4.5` — the fractional value is conclusive proof of halving. `rules_check.py` computes nb\_def = 9.0 (full qty). All eight L4 weekly nb\_def values differ by exactly a factor of 2. **Monthly:** `mens_lignes.csv` June L4: `defauts = 16.5` vs rule = 34.0. July L4: `defauts = 35` vs rule = 76 (halving plus ACCEPT exclusion — see D4). `mens_global.csv` June: `defauts = 121.5` (fractional) vs rule = 158. |
| D2 | R2 | `Module2.bas` line 102<br>`If txr >= 3 Then st = "ROUGE"` | Scrap rate **equal to exactly 3.00 %** is classified `ROUGE`. | R2: RED status requires scrap rate **strictly greater than** 3.0 %. At exactly 3.00 % the correct status is ORANGE. | `rap_hebdo.csv` W27 L3: `rebut = 9`, `produit = 300`, `tx_rebut = 3.00`, `statut = ROUGE`. Calculation: 9 / 300 × 100 = 3.00 %. Under R2 (> 3.0 %), this must be ORANGE. `rules_check.py` reports ORANGE for this row; legacy has ROUGE. This is the only row in the dataset where `tx_rebut` falls exactly on the 3 % boundary. |
| D3 | R2 | `Module1.bas` line 74<br>`If v > 3 Then ... ROUGE` | The `Colorer` subroutine sets the cell **background** of the `statut` column to red only when `tx_rebut > 3` (strict). `CalcHebdo` sets the **text** of the same cell using `>= 3`. | The two representations of status in the workbook are inconsistent. The CSV export captures the text value (affected by D2); the cell background colour (affected by D3) is not exported to CSV. | Not detectable in `sorties_legacy/` CSVs — the exported text comes from `CalcHebdo`. The conflict is visible only in the open workbook: W27 L3 text = ROUGE, background = orange (because 3.00 % is not > 3). No output file mismatch. |
| D4 | R1 / R7 | `Module2.bas` line 266<br>`If dp <> "ACCEPT" Then nd(m,l) = ...` | In `CalcMensuel`, ACCEPT records are **excluded** from the monthly defect count `nd`. In `CalcHebdo` (line 81–88) ACCEPT records are included — there is no equivalent filter. This inconsistency means the weekly defect rate and the monthly defect rate use different definitions. | R1 and R7: *"All dispositions (SCRAP, REWORK, ACCEPT) count as defects."* ACCEPT must be included in both weekly and monthly paths. | For each line in June, `rules_check.py` identifies the exact ACCEPT quantities missing from the legacy monthly count: **L1** +9 (NC-0006 qty 6, NC-0042 qty 1, NC-0051 qty 2); **L2** +4 (NC-0053 qty 1, NC-0058 qty 3); **L3** +6 (NC-0003 qty 2, NC-0009 qty 1, NC-0010 qty 2, NC-0043 qty 1). `mens_lignes.csv` June L1: `defauts = 51`; rule = 60 (difference = 9 = ACCEPT qty for L1). The weekly `rap_hebdo.csv` W23 L1 `nb_def = 11` does include ACCEPT, confirming the inconsistency between weekly and monthly paths. |
| D5 | R3 | `Module2.bas` lines 170–171<br>`If cum < 80 Then ... "X"`<br>then `cum = cum + ...` | The `prioritaire` flag is set **before** the current item's percentage is added to `cum`. A code is marked priority if the cumulative *before it* is < 80 %, so the code that pushes the cumulative past 80 % is marked. This coincidentally produces the correct result for most entries — but the condition `cum < 80` (pre-addition) is equivalent to the correct condition `(cum - pct) < 80`, which means the algorithm is correct as long as no type starts exactly at or above 80 % before being processed. | R3: *"The types needed to reach 80 % cumulative are marked as priorities (the type that crosses 80 % is included)."* The crossing type must be flagged. | `rules_check.py` found **no mismatch** between rule and legacy for any Pareto row in this dataset — the off-by-one is structurally wrong but produces identical results here because no type in this data starts at cumul ≥ 80 % before it is processed. The bug would manifest when the cumulative just reaches or exceeds 80 % across a boundary case (e.g. two codes tied at exactly half the 80 % threshold). Retained as a code-quality issue requiring a fix in the modern version. |
| D6 | R4 / all | `Module1.bas` line 31 (comment)<br>`'ImportCSV … "parameters.csv"` | `data/parameters.csv` is never loaded. Every threshold used in computations is hardcoded in `Module2.bas`: labour rate 45 € (line 56), scrap red threshold 3 / orange 2 (lines 102–104), escalation quantity 20 (line 191), recurrence window 6 days (line 221). | The parameters file is the declared single source of truth. A change to `data/parameters.csv` has no effect on the macro. | `rules_check.py` loaded `parameters.csv` and used its values. Since the hardcoded values happen to match `parameters.csv` exactly, this produces no numerical difference in the current dataset. The divergence is an architectural maintainability issue: updating parameters requires editing VBA source, not the CSV. |
| D7 | R6 | `Module2.bas` lines 211–238 (recurrence loop) | The recurrence algorithm scans forward from each anchor row and counts occurrences of the same `(line, part_ref, defect_code)` triplet within `d2 - d1 <= 6` days. The anchor itself is always counted (r2 starts at r). This is a forward-only sliding window anchored at each candidate first occurrence. | R6 requires a flag when the same triplet appears ≥ 3 times within any 7-day window. For a first occurrence that is followed by ≥ 2 more within 6 days, the forward-scan and the sliding-window give the same result. Potential over-counting (duplicate alerts for the same window) can arise when a second occurrence itself is also followed by ≥ 2 more — but in practice the second anchor would only start a new window if the first event has already rolled outside 7 days. | `rules_check.py` (R6 sliding window) and legacy `alertes.csv` both produce exactly one RECURRENCE alert: NC-0085 (L2, P-2001, D03, 2026-07-06). The three records NC-0085 (07-06), NC-0095 (07-08), NC-0106 (07-10) are within a 4-day span. NC-0095 as a second anchor sees only NC-0095 + NC-0106 = 2 occurrences → no duplicate. **No output difference in this dataset.** |

---

## 10. Summary of Decisions for Modernisation

> Rows are ordered by impact: D1 and D4 produce large measurable differences;
> D2 produces one status flip; D3 is workbook-only; D5 and D6 are latent.

| # | Divergence | Recommended action |
|:---:|---|---|
| D1 | L4 defect qty halved | **Requires a business decision.** The 2011 modification comment (`' ne pas toucher`) suggests it was intentional. If L4's double shift means each physical part is logged twice, halving normalises the rate. If not, the full qty is correct per R1. The modern version must document and preserve whichever interpretation is confirmed. |
| D2 | `>= 3` vs `> 3` for RED alert | **Fix** in the modern version: use `> 3.0` (strict) per R2. This changes exactly one row in the current dataset (W27 L3: ROUGE → ORANGE). |
| D3 | Colorer threshold vs text status mismatch | **Fix**: align the visual indicator to the same `> 3.0` threshold used for the corrected text status. No CSV output change; visual only. |
| D4 | ACCEPT excluded from monthly defect count | **Fix** in the modern version: include all dispositions (SCRAP, REWORK, ACCEPT) in the monthly defect count, matching the weekly path and R1/R7. This changes `defauts` and `tx_def` for every line/month combination. |
| D5 | Pareto priority flag set before cumulative update | **Fix**: evaluate the crossing condition after adding the item's percentage. No output change on current data but the algorithm is incorrect by construction. |
| D6 | `parameters.csv` not loaded; thresholds hardcoded | **Fix** in the modern version: read all threshold values from `parameters.csv` at startup. Hardcoded values currently match the file, so no numerical change. |
| D7 | Recurrence forward-scan (structural) | **Monitor**: in the current dataset the forward-scan and the sliding window give identical results. The modern version should use a proper sliding window to be safe, but this is not a confirmed output error. |
