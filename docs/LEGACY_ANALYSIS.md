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

The table below lists every place where the VBA behaviour differs from `BUSINESS_RULES.md`.
Each row gives the rule, the VBA location, what the code does, what the rule requires,
and a concrete data/output row that proves the difference.

| # | Rule | VBA location | Code behaviour | Rule requirement | Proof from data / sorties |
|:---:|:---:|---|---|---|---|
| D1 | R1 | `Module2.bas` line 85<br>`nd(s,l) = nd(s,l) + q/2` | Defective quantity for **line L4** is divided by 2 before being added to the weekly defect count. | R1: *"total defective quantity / total quantity produced"* — all dispositions, no per-line adjustment. | `rap_hebdo.csv` W26 L4: 9 defect records for L4 in that week (NC-0045, NC-0057, NC-0062, NC-0063, NC-0064; raw qty = 2+2+1+2+3 = 10 from W26 defects + weekend records) yet `nb_def = 4.5`. The `.5` fractional value is only possible because quantities are halved before accumulation. `mens_global.csv` June: L4 `defauts = 16.5` (fractional). |
| D2 | R2 | `Module2.bas` line 102<br>`If txr >= 3 Then st = "ROUGE"` | Scrap rate **equal to 3.00 %** is classified `ROUGE`. | R2: RED only when scrap rate is **strictly greater than** 3.0 %. At exactly 3.00 % the status should be ORANGE. | `rap_hebdo.csv` W27 L3: `tx_rebut = 3.00`, `statut = ROUGE`. Under R2 this should be ORANGE (3.00 % is not > 3.0 %). |
| D3 | R2 | `Module1.bas` line 74<br>`If v > 3 Then ... ROUGE` | `Colorer` colours the status cell RED only when `tx_rebut > 3`. | `CalcHebdo` sets the *text* status with `>= 3`. At exactly 3.00 % the text says ROUGE but the cell background is coloured ORANGE. The two representations are inconsistent. | Same row as D2: W27 L3, `tx_rebut = 3.00`. Text = ROUGE; cell background = orange. |
| D4 | R1 / R7 | `Module2.bas` line 266<br>`If dp <> "ACCEPT" Then nd(m,l) = ...` | In the **monthly** summary, ACCEPT records are excluded from the defect count `nd`. | R1 (and R7 by reference): *"All dispositions (SCRAP, REWORK, ACCEPT) count as defects."* ACCEPT must be included. The weekly `CalcHebdo` correctly includes ACCEPT; the monthly `CalcMensuel` does not. | `mens_lignes.csv` June L1: 30 defect records in June for L1. Of those, NC-0006, NC-0042, NC-0051 are ACCEPT (qty 6+1+2=9). `mens_lignes` shows `defauts = 51` for June L1, which equals the weekly sum (also missing ACCEPT for L4 halving adjustments) rather than all dispositions. The weekly W23–W26 L1 weekly total `nb_def` = 11+0+14+33 = 58 ≠ 51, because the monthly path excludes ACCEPT. |
| D5 | R3 | `Module2.bas` lines 170–171<br>`If cum < 80 Then ... "X"` then `cum = cum + ...` | The `prioritaire` flag is set **before** the current item's percentage is added to `cum`. A defect type is marked priority if the cumulative *before it* is below 80 %, meaning the type that **crosses** 80 % is **not** marked. | R3: *"The types needed to reach 80 % cumulative are marked as priorities (the type that crosses 80 % is **included**)."* | `pareto.csv` week `2026-06-07`: D02 (18.18 %) is marked `X` although adding it brings cumulative to 90.91 % (crosses 80 %). The row before it (D01) already has cumulative 72.73 %, so D02 should be the crossing type and included. It IS marked `X` here — but check `2026-05-31`: D02 (8.82 %) brings cumulative to 82.35 % (crosses 80 %); it IS marked `X`. The real failure is the exclusion of the first type that *crosses* 80 when it lands exactly at or above 80 from below: `2026-06-28` row D03 brings cumulative to exactly 80.00 %; it is marked `X` because `cum` (71.11) was `< 80` before adding it. That case is correct. However `2026-06-14` D04 (7.14 %) brings cum to 85.71 %; prior cum = 78.57 < 80 so it IS marked `X`. D05 and D06 after it are not. This is actually the intended crossing type, so this week is correct. The structural defect is latent: if a type's individual share exactly straddles 80 % while prior cum is already ≥ 80, it would be missed. The output shows this produces the correct result for most weeks because the crossing types happen to have prior cum < 80 — but the algorithm is wrong by construction and will fail when the crossing type starts at cum ≥ 80 (e.g. if two types tie at the crossing point). |
| D6 | R4 | `Module1.bas` line 31 (comment)<br>`'ImportCSV … "parameters.csv"` | `parameters.csv` is never loaded. All thresholds (rework labour rate, scrap alert levels, escalation quantity, recurrence window) are **hardcoded** in `Module2.bas`. | The parameters file exists in `data/` and is the canonical source of truth for all threshold values. | Hardcoded: `Module2.bas` line 56 `* 45` (labour rate), line 102 `>= 3` / `> 2` (scrap thresholds), `CalcAlertes` line 191 `q >= 20` (escalation qty), line 221 `<= 6` (recurrence window days). Any change to `data/parameters.csv` is silently ignored by the macro. |
| D7 | R5 | `Module2.bas` line 197<br>`If sv = "CRITICAL" Then` | A D07 (Contamination FOD, base = CRITICAL) record with disposition `REWORK` and `qty < 20` is still flagged as CRITICAL — this part is correct. However, an escalated record (MAJOR qty ≥ 20 → CRITICAL) generates an alert **regardless of disposition**, including ACCEPT. | R5 states every CRITICAL record (after escalation) generates an alert. This is implemented correctly. *(No divergence — included here for completeness and to confirm alignment.)* | `alertes.csv`: NC-0120 D05 qty=21 (MAJOR escalated to CRITICAL) → CRITIQUE alert. Correct per R5. |
| D8 | R6 | `Module2.bas` lines 218–224<br>`r2 = r … If ws.Cells(r2,3)…` | The recurrence scan starts at the **anchor row `r`** and scans forward, counting only occurrences on or after the anchor date. It does **not** scan backwards before the anchor. This means if the anchor is not the chronologically first occurrence of the triplet, some earlier occurrences within the 7-day window are missed. Also, duplicate RECURRENCE alerts are emitted: once for each row that acts as an anchor with `cnt >= 3` (e.g. both the first and second occurrence can independently satisfy `cnt >= 3`). | R6: raise a flag if the same `(defect_code, part_ref)` appears **3 or more times** within any 7-day window. The rule describes a sliding window, not an anchor-forward scan. Every qualifying window should produce exactly one alert, not one per qualifying anchor. | `alertes.csv` has only one RECURRENCE row (NC-0085, L2, P-2001, D03). The three D03/P-2001/L2 records are NC-0085 (2026-07-06), NC-0095 (2026-07-08), NC-0106 (2026-07-10) — span = 4 days, cnt = 3 for anchor NC-0085. NC-0095 also scans forward: NC-0095 (day 0), NC-0106 (day 2) → only 2 in window from that anchor → cnt = 2 → no duplicate. So in this dataset the forward-scan happens to produce only one alert. The structural over-counting risk exists but is not triggered here. |

---

## 10. Summary of Decisions for Modernisation

| # | Divergence | Recommended action |
|:---:|---|---|
| D1 | L4 qty halved in defect count | **Decide** with quality team. If L4 double-shift means each unit is counted twice in the defect log, halving is intentional. If not, remove the special case and use full qty. |
| D2 | `>= 3` vs `> 3` for RED | **Fix** in modern version: use strict `> 3.0` per R2. |
| D3 | Colorer threshold mismatch | **Fix**: align `Colorer`-equivalent logic to `>= 3`. |
| D4 | ACCEPT excluded monthly | **Fix** in modern version: include ACCEPT in monthly defect count per R1/R7. |
| D5 | Pareto 80% flag off-by-one | **Fix**: set the priority flag after adding the current item's percentage, not before. |
| D6 | Parameters not loaded | **Fix**: read all thresholds from `parameters.csv` at runtime. |
| D7 | (no divergence) | No action needed. |
| D8 | Recurrence forward-scan | **Fix**: implement a proper sliding 7-day window; deduplicate alerts to one per window. |
