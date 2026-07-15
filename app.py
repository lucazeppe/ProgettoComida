"""App Streamlit per il controllo dei pasti aziendali (fornitori Amati e Zippi)."""
from __future__ import annotations

import calendar
from datetime import date, datetime

import pandas as pd
import streamlit as st

from business import (
    ANOMALY_DOUBLE_LABEL,
    ANOMALY_HOURS_LABEL,
    apply_overrides,
    build_master,
    check_weekend_activity,
    supplier_recap,
)
from exports import build_reminder_export, build_summary_export
from parsers import parse_amati, parse_hours, parse_zippi

st.set_page_config(page_title="Controllo Pasti Aziendali", layout="wide")

MONTHS_IT = [
    "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
    "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre",
]

if "master_df" not in st.session_state:
    st.session_state.master_df = None
if "overrides" not in st.session_state:
    st.session_state.overrides = {}
if "warnings" not in st.session_state:
    st.session_state.warnings = []
if "period" not in st.session_state:
    st.session_state.period = None

st.title("Controllo Pasti Aziendali — Amati / Zippi")

with st.sidebar:
    st.header("Periodo di riferimento")
    today = date.today()
    year = st.number_input("Anno", min_value=2020, max_value=2100, value=today.year, step=1)
    month = st.selectbox("Mese", options=list(range(1, 13)), format_func=lambda m: MONTHS_IT[m - 1], index=today.month - 1)
    exclude_weekends = st.checkbox(
        "Escludi sabato/domenica dal calcolo (segnala se contengono dati)", value=True
    )

    st.header("File di input")
    zippi_file = st.file_uploader("Ordini Zippi (.xlsx)", type=["xlsx"])
    amati_file = st.file_uploader("Ordini Amati (.xlsx)", type=["xlsx"])
    hours_file = st.file_uploader("Ore lavorate (.xlsx)", type=["xlsx"])

    process = st.button("Processa", type="primary", use_container_width=True)

if process:
    if not (zippi_file and amati_file and hours_file):
        st.error("Carica tutti e 3 i file prima di procedere.")
        st.stop()

    all_warnings: list[str] = []
    try:
        zippi_df, w = parse_zippi(zippi_file, year, month, warn_weekends=exclude_weekends)
        all_warnings += w
        amati_df, w = parse_amati(amati_file, year, month)
        all_warnings += w
        hours_df, w = parse_hours(hours_file, year, month, warn_weekends=exclude_weekends)
        all_warnings += w
    except ValueError as e:
        st.error(str(e))
        st.stop()

    master_df = build_master(zippi_df, amati_df, hours_df, year, month, exclude_weekends=exclude_weekends)
    if exclude_weekends:
        all_warnings += check_weekend_activity(master_df)

    st.session_state.master_df = master_df
    st.session_state.overrides = {}
    st.session_state.warnings = all_warnings
    st.session_state.period = (year, month)

if st.session_state.master_df is None:
    st.info("Carica i 3 file e premi **Processa** per iniziare.")
    st.stop()

for w in st.session_state.warnings:
    st.warning(w)

master_df = st.session_state.master_df
# Vista editabile: non esclude le righe annullate da entrambi i fornitori, così
# restano visibili/correggibili finché non si riprocessa da capo.
edit_df = apply_overrides(master_df, st.session_state.overrides, drop_fully_cancelled=False)
# Vista effettiva per recap ed export: le righe senza più alcun ordine sono escluse.
effective_df = apply_overrides(master_df, st.session_state.overrides, drop_fully_cancelled=True)

st.subheader("Vista dipendenti — ordini e anomalie")

view = edit_df.copy()
view["Anomalia"] = view.apply(
    lambda r: ", ".join(
        filter(None, [
            ANOMALY_DOUBLE_LABEL if r["anomaly_double"] else None,
            ANOMALY_HOURS_LABEL if r["anomaly_hours"] else None,
        ])
    ),
    axis=1,
)
view = view.sort_values(["employee_id", "date"])
view["Rimuovi Amati"] = view.apply(lambda r: bool(st.session_state.overrides.get((r["employee_id"], r["date"]), {}).get("remove_amati")), axis=1)
view["Rimuovi Zippi"] = view.apply(lambda r: bool(st.session_state.overrides.get((r["employee_id"], r["date"]), {}).get("remove_zippi")), axis=1)
view["Abbona vincolo ore"] = view.apply(lambda r: bool(st.session_state.overrides.get((r["employee_id"], r["date"]), {}).get("waive_hours")), axis=1)

display_cols = [
    "employee_id", "nombre", "email", "date", "amati", "zippi", "hours",
    "eligible", "Anomalia", "Rimuovi Amati", "Rimuovi Zippi", "Abbona vincolo ore",
]
edited = st.data_editor(
    view[display_cols],
    column_config={
        "employee_id": st.column_config.TextColumn("ID Dipendente", disabled=True),
        "nombre": st.column_config.TextColumn("Nome", disabled=True),
        "email": st.column_config.TextColumn("Email", disabled=True),
        "date": st.column_config.DateColumn("Data", disabled=True),
        "amati": st.column_config.CheckboxColumn("Amati", disabled=True),
        "zippi": st.column_config.CheckboxColumn("Zippi", disabled=True),
        "hours": st.column_config.NumberColumn("Ore", disabled=True, format="%.1f"),
        "eligible": st.column_config.CheckboxColumn("Diritto (>=6h)", disabled=True),
        "Anomalia": st.column_config.TextColumn("Anomalia", disabled=True),
    },
    disabled=False,
    hide_index=True,
    use_container_width=True,
    key="employee_editor",
)

if st.button("Applica forzature"):
    new_overrides = dict(st.session_state.overrides)
    for _, r in edited.iterrows():
        key = (r["employee_id"], r["date"])
        ov = {
            "remove_amati": bool(r["Rimuovi Amati"]),
            "remove_zippi": bool(r["Rimuovi Zippi"]),
            "waive_hours": bool(r["Abbona vincolo ore"]),
        }
        if any(ov.values()):
            new_overrides[key] = ov
        else:
            new_overrides.pop(key, None)
    st.session_state.overrides = new_overrides
    st.rerun()

st.subheader("Recap fornitore — validazione fatture")
st.caption("Conteggio pasti forniti (nessun importo: i costi 23,46/165 riguardano l'addebito al dipendente).")
recap = supplier_recap(effective_df)
st.dataframe(recap, hide_index=True, use_container_width=True)

st.subheader("Export")
col1, col2 = st.columns(2)
year_sel, month_sel = st.session_state.period

with col1:
    reminder_bytes = build_reminder_export(effective_df)
    st.download_button(
        "Scarica export solleciti (anomalie)",
        data=reminder_bytes,
        file_name=f"anomalie_pasti_{year_sel}-{month_sel:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with col2:
    summary_bytes = build_summary_export(effective_df, year_sel, month_sel)
    st.download_button(
        "Scarica riepilogo colorato",
        data=summary_bytes,
        file_name=f"riepilogo_pasti_{year_sel}-{month_sel:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
