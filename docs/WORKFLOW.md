# Modernizing a legacy tool with an AI partner — the playbook

This is the workflow used on the NC Tracker, written so it can be reused on any legacy tool that turns input files into reports (Excel/VBA, Access, old scripts). The AI partner was IBM Bob 2.0; the prompts below are the ones used, with project names generalised where needed.

The principle: **the legacy outputs are the truth about what the tool does; the procedure is the truth about what it should do.** Every difference between the two is found, proven, decided by a person and applied on purpose, never by accident.

## 0. Freeze the reference

Before the AI touches anything:

1. Run the legacy tool once on real (or representative) input and keep its outputs. They are the **golden master**.
2. Commit the legacy code, the inputs and the outputs, and tag the commit (here `v0-legacy`). Everything after that tag is modernization work, so `git diff v0-legacy` shows exactly what changed.
3. Mark those files as frozen in the AI's context file (here [`bob.md`](../bob.md)) and keep git from rewriting their bytes ([`.gitattributes`](../.gitattributes)).

If you want to measure how much the AI discovers, keep your own notes about known defects **outside** the workspace it can read.

## 1. Analyse

> Read bob.md, BUSINESS_RULES.md, the CSVs in data/ and the VBA in legacy_vba/. Write docs/LEGACY_ANALYSIS.md explaining step by step what the entry macro does (inputs, calculations, outputs). Then add a table of every place where the code's behaviour differs from the rules, each with the code line and a row from data/ or the legacy outputs that proves it. Don't change any existing file. Commit your work when done.

**Check:** recompute a few of the cited numbers yourself. Here, the first pass found half of the real divergences, cited wrong arithmetic for two of them, and reported one divergence that did not exist.

## 2. Verify against the rules, not the code

Reading code finds what looks wrong. To find what *is* different, implement the procedure independently and compare every value:

> In tools/rules_check.py, implement the rules exactly as BUSINESS_RULES.md states them, independently of the legacy code. Run it on data/ and compare every value with the legacy outputs. For each mismatch, find the code line responsible and add it to the divergence table with the numbers as evidence. Re-check every existing row the same way: correct evidence that doesn't match the files, remove rows the data and code don't support, and make sure the recommended actions don't contradict each other.

**Check the checker.** The first version quietly copied the legacy behaviour in several places (its own comments said "matches VBA"), which hid exactly the divergences it was meant to find. The fix was a review prompt:

> Rewrite the checker so every calculation comes only from BUSINESS_RULES.md: above each function, quote the rule sentence it implements, and never copy a legacy behaviour. Where the procedure is silent, use the plainest reading of it and list that assumption. Print a summary count of mismatches per divergence row.

After that, the checker found two more divergences and removed the false positive.

## 3. Decide

> Create docs/DECISIONS.md with one row per divergence: ID, rule, legacy behaviour, business impact (which reported numbers change, using the checker's numbers), and a recommendation (keep the legacy behaviour or follow the rule) with a short reason. Leave a "Decision" column empty for me to fill in.

A person fills in the Decision column. Not every divergence is a bug: here the L4 halving was **kept**, but turned from a hidden rule into a documented parameter.

## 4. Lock the behaviour

> Create a Python package that reproduces the legacy tool exactly, including every divergence listed in the analysis. Structure it so each divergence can later be switched by a setting, but implement legacy mode only for now. Write characterization tests that compare each output with the matching golden-master file, field by field, ignoring only line endings. Write outputs to a temporary folder, never to the golden master. Use parallel subagents, one per report. You're done when all tests pass.

Exact-match tests find what reading and rule checks miss: to match the golden master, the port had to reproduce the legacy's per-record cost rounding, which became divergence D9. Tag the result (`v1-parity`).

**Check:** regenerate the outputs yourself and diff them against the golden master; don't rely only on the AI's report of its tests.

## 5. Correct

> Implement corrected mode, applying every decision in docs/DECISIONS.md. Keep every legacy-mode test passing. Add corrected-mode tests that check, for each decision, the numbers given in the decision log, and cross-check against the rules checker where both follow the same rule. Add a function that lists every value that differs between legacy and corrected mode, with the decision that explains it.

**Check each decision against its wording.** Here one decision was applied backwards (a KEEP implemented as a FIX), and one rule was implemented with an extra condition the procedure does not contain. Attribution of each changed value is computed by switching one decision at a time, not hard-coded, so the explanations stay right when decisions change.

## 6. Ship

A small app shows both modes side by side and, for every number that changed, the decision and rule behind it. That page is what lets the quality team accept the new tool: nothing changed that nobody decided.

## Lessons

- **Freeze first, tag often.** Tags at each milestone (`v0-legacy`, `v1-parity`, `v2-corrected`, `v3-app`) make every step reviewable.
- **Verify the AI's claims with the files, not its summary.** Several session reports said work was committed or pushed when it wasn't, or attributed mismatches to the wrong cause. `git status`, re-running the tests and diffing outputs caught each case in seconds.
- **An independent check must be independent.** Ask for the rule text next to each calculation, and forbid copying the legacy logic.
- **Golden-master tests are a discovery tool**, not just a safety net.
- **Budget the sessions.** The hackathon allowance (40 Bobcoins) covered four sessions: analysis and decisions, the port, corrected mode, and the first app version. Short, well-scoped prompts with explicit "done when" criteria used it best.
