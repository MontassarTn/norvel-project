# Norvel NC Tracker — modernizing a legacy factory quality tool with IBM Bob

Factories still run their quality meetings on Excel/VBA tools written a decade ago. Nobody knows exactly what they compute, so rewriting them is slow, and a rewrite silently changes the numbers people rely on.

This project shows a repeatable workflow, run with **IBM Bob 2.0**, that modernizes such a tool **without losing or silently changing a single number**. Every difference between the old tool and the official procedure is found, documented, decided on by the product owner, and applied deliberately.

It is demonstrated on the *NC Tracker* of Norvel Composites, a non-conformance tracker written in Excel VBA (2009–2013), rebuilt as a tested Python package with a Streamlit + SQLite app.

**Live app:** https://norvel-project-brssjw3ji76l6mcfmwzdge.streamlit.app

> Norvel Composites is fictional. The legacy tool, its data and the business rules are synthetic, built for the IBM Bob 2.0 hackathon to reproduce a typical factory quality tool. Its deviations from the procedure were planted on purpose, so the discovery could be measured.

## The workflow

| Step | What happens | Output |
|---|---|---|
| 0. Freeze | Run the legacy macro once; its outputs become the reference | [`sorties_legacy/`](sorties_legacy/), tag `v0-legacy` |
| 1. Analyse | Bob explains the VBA step by step and lists where it differs from the rules | [`docs/LEGACY_ANALYSIS.md`](docs/LEGACY_ANALYSIS.md) |
| 2. Verify | Bob implements rules R1–R7 as written, independently of the VBA, and diffs every value against the legacy outputs | [`tools/rules_check.py`](tools/rules_check.py) |
| 3. Decide | The product owner decides, per difference, to keep the legacy behaviour or follow the rule | [`docs/DECISIONS.md`](docs/DECISIONS.md) |
| 4. Lock | Bob ports the tool to Python; characterization tests require byte-identical outputs | [`nctrack/`](nctrack/), [`tests/test_parity.py`](tests/test_parity.py), tag `v1-parity` |
| 5. Correct | Each decision becomes a switch; corrected mode applies them, each with tests | [`nctrack/config.py`](nctrack/config.py), [`tests/test_corrected.py`](tests/test_corrected.py), tag `v2-corrected` |
| 6. Ship | A Streamlit app shows both modes and every number that changed, with the reason | [`app.py`](app.py), tag `v3-app` |

The step-by-step playbook, with the prompts used and the lessons learned, is in [`docs/WORKFLOW.md`](docs/WORKFLOW.md).

## Results

- **9 divergences** between the legacy tool and the procedure, each pinned to a VBA line and proven with data ([analysis §9](docs/LEGACY_ANALYSIS.md)). Examples: line L4's defects silently halved since 2011, a 3.00 % scrap rate labelled RED instead of ORANGE, ACCEPT dispositions missing from monthly figures, Pareto weeks starting on Sunday instead of Monday.
- **Legacy mode reproduces all five legacy reports exactly**: every field of `rap_hebdo`, `pareto`, `alertes`, `mens_lignes` and `mens_global`, checked by 10 characterization tests.
- **Corrected mode** applies the product owner's decisions: one KEEP (the L4 factor, now an explicit parameter) and the rest FIX. **358 values** change, and the app's *What changed* page traces each one to the decision that causes it.
- **90 automated tests** (`pytest`), including a headless run of every app page in both modes.

## Run it

```bash
pip install -r requirements.txt
pytest                                    # 90 tests
streamlit run app.py                      # the app
python -m nctrack.runner --data data --output out   # legacy-mode CSVs
```

The app loads `data/` into a SQLite database (`nctrack.db`, rebuilt at start) and computes every report from it. A sidebar switch selects **Legacy** (identical to the old tool) or **Corrected** (decisions applied).

## Who did what

| Contributor | Work |
|---|---|
| **IBM Bob 2.0** (4 sessions, 40 Bobcoins; [screenshots](docs/bob-sessions/)) | Analysis of the VBA and divergence table; the independent rules checker (several iterations); draft of the decision log; the Python port with parity tests; corrected mode, the D9 finding and the diff function; the first version of the Streamlit app |
| **Product owner** (Montassar Jaziri) | The keep/fix decision for every divergence |
| **Claude** (claude.ai) | Generated the synthetic legacy tool, data and business rules (`v0-legacy`) |
| **Claude Code** | Repository setup; independent re-checks of each Bob result; after Bob's budget ran out: applying decisions D1 and D8 as written, computed diff attribution, the SQLite data path, the app tests and this documentation |

## Repository layout

```
BUSINESS_RULES.md     official quality procedure (rules R1–R7)
data/                 input CSVs (frozen)
legacy_vba/           legacy VBA modules (frozen), also imported in NCTRACK.xlsm
sorties_legacy/       outputs of the legacy macro = golden master (frozen)
tools/rules_check.py  rules R1–R7 implemented independently of the VBA
nctrack/              Python package: one module per report, config, diff, SQLite
tests/                parity, corrected-mode and app tests
app.py                Streamlit app
docs/                 analysis, decisions, workflow, Bob session screenshots
bob.md                project context for the AI assistant
```

## License

[MIT](LICENSE)
