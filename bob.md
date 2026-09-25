# Project context — Norvel NC Tracker modernization

## What this is

Entry for the IBM Bob 2.0 hackathon (lablab.ai, online, 25–27 September 2026). The project demonstrates modernizing a legacy factory quality tool with an AI development partner: analyse → document → lock behaviour with tests → migrate → verify → improve.

Norvel Composites is a fictional company. The legacy tool and all data are synthetic, built for the hackathon to reproduce a typical factory quality tool. No real company code or data is used.

## Repository layout

- `BUSINESS_RULES.md` — the official quality procedure (rules R1–R7): what the tool is SUPPOSED to do
- `data/` — input CSVs: defect log, production log, defect types, parts, parameters
- `legacy_vba/Module1.bas`, `Module2.bas` — the legacy Excel VBA tool (entry macro: `LANCER_TOUT`)
- `NCTRACK.xlsm` — Excel workbook containing the imported VBA modules
- `sorties_legacy/` — outputs of the legacy tool = golden master for characterization tests

## How the legacy tool runs

Open `NCTRACK.xlsm` (saved in the project root, next to `data/`), run macro `LANCER_TOUT`. It imports the CSVs into sheets, computes weekly rates, Pareto, alerts and monthly summaries, and exports 5 CSVs to `sorties_legacy/`: `rap_hebdo`, `pareto`, `alertes`, `mens_lignes`, `mens_global`.

## Current status

- [x] Business rules defined, sample data generated
- [x] Legacy VBA written
- [x] Run the macro in Excel and confirm `sorties_legacy/` is produced
- [x] Commit and tag `v0-legacy`
- [ ] Modernization phase (to be done with IBM Bob for the hackathon demo)

## Target stack for the modern version

Python, Streamlit, SQLite, pytest.

## Important

- The legacy behaviour may differ from `BUSINESS_RULES.md` in places. Do not "fix" the legacy code: it must stay frozen as the reference. Differences are to be discovered, documented and decided on during modernization.
- `data/`, `legacy_vba/`, `NCTRACK.xlsm` and `sorties_legacy/` are the frozen reference: never edit or regenerate them. New code writes its outputs elsewhere.
- The modernization itself is the hackathon demo and is done with IBM Bob.
- The owner's working language is French; code comments in the legacy tool are in French.
