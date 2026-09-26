"""
NCTrack Streamlit application.

At startup:
  - Loads the CSVs from data/ into a SQLite database (nctrack.db, recreated
    once per server start, not committed).
  - Computes all five reports from that database via the nctrack package.

Pages (via sidebar navigation):
  1. Weekly Report   – rap_hebdo
  2. Pareto           – pareto
  3. Alerts           – alertes
  4. Monthly Summary  – mens_lignes + mens_global
  5. What Changed     – diff between legacy and corrected mode

Sidebar toggle selects legacy or corrected mode on every page.
Scrap status (VERT / ORANGE / ROUGE) is colour-coded as in the workbook.
Each table has a CSV download button.
"""

from __future__ import annotations

import io
from pathlib import Path
from typing import Any

import streamlit as st

from nctrack import alertes, mens_global, mens_lignes, pareto, rap_hebdo
from nctrack.config import CorrectedConfig, LegacyConfig
from nctrack.diff import DECISION_DESCRIPTIONS, report_differences
from nctrack.db import build_database, load_database

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DATA_DIR = Path(__file__).parent / "data"
DB_PATH = Path(__file__).parent / "nctrack.db"

STATUS_COLOURS = {
    "ROUGE": "#ff4b4b",
    "ORANGE": "#ffa500",
    "VERT": "#21c354",
}

# ---------------------------------------------------------------------------
# Database
# ---------------------------------------------------------------------------


@st.cache_resource(show_spinner="Loading data into SQLite…")
def _database() -> str:
    """Recreate nctrack.db from data/ once per server process; return its path."""
    build_database(DATA_DIR, DB_PATH)
    return str(DB_PATH)


# ---------------------------------------------------------------------------
# Report computation (cached per mode)
# ---------------------------------------------------------------------------


@st.cache_data(show_spinner=False)
def _compute_all(mode: str) -> dict[str, list[dict[str, Any]]]:
    """Compute all five reports for the given mode ('legacy' or 'corrected')."""
    cfg = LegacyConfig() if mode == "legacy" else CorrectedConfig()
    ds = load_database(_database())
    al_rows = alertes.compute(ds, cfg)
    return {
        "rap_hebdo": rap_hebdo.compute(ds, cfg),
        "pareto": pareto.compute(ds, cfg),
        "alertes": al_rows,
        "mens_lignes": mens_lignes.compute(ds, cfg, al_rows),
        "mens_global": mens_global.compute(ds, cfg, al_rows),
    }


@st.cache_data(show_spinner=False)
def _compute_diffs() -> list[dict[str, Any]]:
    ds = load_database(_database())
    return report_differences(ds)


# ---------------------------------------------------------------------------
# Rendering helpers
# ---------------------------------------------------------------------------


def _rows_to_csv(rows: list[dict[str, Any]]) -> bytes:
    """Serialise a list of dicts to CSV bytes (UTF-8 with BOM for Excel)."""
    if not rows:
        return b""
    buf = io.StringIO()
    import csv as _csv

    writer = _csv.DictWriter(buf, fieldnames=list(rows[0].keys()))
    writer.writeheader()
    writer.writerows(rows)
    return buf.getvalue().encode("utf-8-sig")


def _status_badge(statut: str) -> str:
    colour = STATUS_COLOURS.get(statut, "#cccccc")
    return f'<span style="background:{colour};color:#fff;padding:2px 8px;border-radius:4px;font-weight:bold">{statut}</span>'


def _highlight_status(rows: list[dict[str, Any]], status_col: str = "statut") -> None:
    """Render a table with coloured status badges using st.write HTML."""
    if not rows:
        st.info("No data.")
        return
    cols = list(rows[0].keys())
    # Build HTML table
    header = "".join(f"<th>{c}</th>" for c in cols)
    body_rows = []
    for row in rows:
        cells = []
        for c in cols:
            val = row[c]
            if c == status_col and val in STATUS_COLOURS:
                cells.append(f"<td>{_status_badge(str(val))}</td>")
            else:
                cells.append(f"<td>{val}</td>")
        body_rows.append("<tr>" + "".join(cells) + "</tr>")
    table_html = (
        "<div style='overflow-x:auto'>"
        "<table style='border-collapse:collapse;font-size:13px;width:100%'>"
        f"<thead><tr style='background:#f0f2f6'>{header}</tr></thead>"
        "<tbody>" + "".join(body_rows) + "</tbody>"
        "</table></div>"
    )
    st.write(table_html, unsafe_allow_html=True)


def _plain_table(rows: list[dict[str, Any]]) -> None:
    if not rows:
        st.info("No data.")
        return
    st.dataframe(rows, use_container_width=True)


def _download_btn(rows: list[dict[str, Any]], filename: str, key: str) -> None:
    if rows:
        st.download_button(
            label=f"⬇ Download {filename}",
            data=_rows_to_csv(rows),
            file_name=filename,
            mime="text/csv",
            key=key,
        )


# ---------------------------------------------------------------------------
# Page: Weekly Report
# ---------------------------------------------------------------------------


def page_weekly(reports: dict[str, list[dict[str, Any]]]) -> None:
    st.header("📅 Weekly Report — rap_hebdo")
    rows = reports["rap_hebdo"]
    _highlight_status(rows)
    _download_btn(rows, "rap_hebdo.csv", "dl_rap_hebdo")


# ---------------------------------------------------------------------------
# Page: Pareto
# ---------------------------------------------------------------------------


def page_pareto(reports: dict[str, list[dict[str, Any]]]) -> None:
    st.header("📊 Pareto — top defects by week")
    rows = reports["pareto"]
    _plain_table(rows)
    _download_btn(rows, "pareto.csv", "dl_pareto")


# ---------------------------------------------------------------------------
# Page: Alerts
# ---------------------------------------------------------------------------


def page_alerts(reports: dict[str, list[dict[str, Any]]]) -> None:
    st.header("🚨 Alerts — alertes")
    rows = reports["alertes"]
    critique = [r for r in rows if r["type"] == "CRITIQUE"]
    recurrence = [r for r in rows if r["type"] == "RECURRENCE"]

    col1, col2 = st.columns(2)
    col1.metric("CRITIQUE alerts", len(critique))
    col2.metric("RECURRENCE alerts", len(recurrence))

    st.subheader("CRITIQUE")
    _plain_table(critique)
    st.subheader("RECURRENCE")
    _plain_table(recurrence)

    _download_btn(rows, "alertes.csv", "dl_alertes")


# ---------------------------------------------------------------------------
# Page: Monthly Summary
# ---------------------------------------------------------------------------


def page_monthly(reports: dict[str, list[dict[str, Any]]]) -> None:
    st.header("📆 Monthly Summary")

    st.subheader("Per-line (mens_lignes)")
    ml = reports["mens_lignes"]
    _highlight_status(ml)
    _download_btn(ml, "mens_lignes.csv", "dl_mens_lignes")

    st.subheader("Global (mens_global)")
    mg = reports["mens_global"]
    _plain_table(mg)
    _download_btn(mg, "mens_global.csv", "dl_mens_global")


# ---------------------------------------------------------------------------
# Page: What Changed
# ---------------------------------------------------------------------------


def page_what_changed() -> None:
    st.header("🔄 What Changed — legacy vs corrected")
    st.caption(
        "Every value that differs between legacy mode and corrected mode, "
        "with the decision ID and business rule that explain the change."
    )

    diffs = _compute_diffs()

    if not diffs:
        st.success("No differences found between legacy and corrected mode.")
        return

    # Decision filter
    all_decisions = sorted({d["decision"] for d in diffs})
    selected = st.multiselect(
        "Filter by decision",
        options=all_decisions,
        default=all_decisions,
        format_func=lambda k: f"{k} — {DECISION_DESCRIPTIONS.get(k, '')[:60]}",
    )
    filtered = [d for d in diffs if d["decision"] in selected] if selected else diffs

    # Report filter
    all_reports = sorted({d["report"] for d in filtered})
    sel_report = st.selectbox("Filter by report", ["(all)"] + all_reports)
    if sel_report != "(all)":
        filtered = [d for d in filtered if d["report"] == sel_report]

    st.metric("Differences shown", len(filtered))

    # Render table
    display_rows = []
    for d in filtered:
        display_rows.append(
            {
                "report": d["report"],
                "key": str(d["key"]),
                "field": d["field"],
                "legacy": d["legacy"],
                "corrected": d["corrected"],
                "decision": d["decision"],
                "rule": d["description"],
            }
        )

    if display_rows:
        st.dataframe(display_rows, use_container_width=True)
        _download_btn(display_rows, "what_changed.csv", "dl_what_changed")

    # Decision legend
    with st.expander("Decision descriptions"):
        for k, v in DECISION_DESCRIPTIONS.items():
            st.markdown(f"**{k}** — {v}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> None:
    st.set_page_config(
        page_title="NCTrack",
        page_icon="🔧",
        layout="wide",
    )

    # --- Sidebar ---
    st.sidebar.title("NCTrack")
    mode = st.sidebar.radio(
        "Mode",
        options=["legacy", "corrected"],
        format_func=lambda m: "Legacy (VBA parity)" if m == "legacy" else "Corrected (rule-compliant)",
        index=0,
    )
    st.sidebar.markdown("---")
    page = st.sidebar.radio(
        "Page",
        options=["weekly", "pareto", "alerts", "monthly", "what_changed"],
        format_func=lambda p: {
            "weekly": "📅 Weekly Report",
            "pareto": "📊 Pareto",
            "alerts": "🚨 Alerts",
            "monthly": "📆 Monthly Summary",
            "what_changed": "🔄 What Changed",
        }[p],
    )

    # Mode badge
    if mode == "legacy":
        st.sidebar.info("⚙️ Legacy mode: VBA parity")
    else:
        st.sidebar.success("✅ Corrected mode: rule-compliant")

    # --- Compute reports ---
    reports = _compute_all(mode)

    # --- Route to page ---
    if page == "weekly":
        page_weekly(reports)
    elif page == "pareto":
        page_pareto(reports)
    elif page == "alerts":
        page_alerts(reports)
    elif page == "monthly":
        page_monthly(reports)
    elif page == "what_changed":
        page_what_changed()


if __name__ == "__main__":
    main()
