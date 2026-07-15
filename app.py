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
    is_intern,
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
if "missing_hours_employees" not in st.session_state:
    st.session_state.missing_hours_employees = []
if "missing_hours_dismissed" not in st.session_state:
    st.session_state.missing_hours_dismissed = False

st.title("Controllo Pasti Aziendali — Amati / Zippi")

with st.sidebar:
    st.header("Periodo di riferimento")
    today = date.today()
    # Default: mese precedente a quello corrente (il mese "chiuso" da verificare).
    default_year, default_month = (today.year, today.month - 1) if today.month > 1 else (today.year - 1, 12)
    year = st.number_input("Anno", min_value=2020, max_value=2100, value=default_year, step=1)
    month = st.selectbox("Mese", options=list(range(1, 13)), format_func=lambda m: MONTHS_IT[m - 1], index=default_month - 1)
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
        master_df = build_master(zippi_df, amati_df, hours_df, year, month, exclude_weekends=exclude_weekends)
    except ValueError as e:
        st.error(str(e))
        st.stop()

    # Dipendenti con almeno un ordine ma assenti dal file ore: trattati come 0
    # ore (quindi anomalia su tutti i loro ordini) senza bloccare il
    # processamento, ma segnalati con un avviso dedicato e richiudibile.
    # Calcolato su master_df (post exclude_weekends) e non sui file grezzi:
    # un ordine di sabato/domenica escluso dal calcolo non deve generare un
    # falso avviso su un dipendente il cui unico ordine non viene mai processato.
    ordering_ids = set(master_df.loc[master_df["amati"] | master_df["zippi"], "employee_id"])
    hours_ids = set(hours_df["employee_id"])
    missing_ids = sorted(ordering_ids - hours_ids)
    missing_employees = []
    for emp_id in missing_ids:
        row = master_df.loc[master_df["employee_id"] == emp_id]
        nombre = row["nombre"].iloc[0] if not row.empty and row["nombre"].iloc[0] else ""
        email = row["email"].iloc[0] if not row.empty else None
        label = f"{emp_id} ({nombre})" if nombre else emp_id
        missing_employees.append({"label": label, "is_intern": is_intern(email)})

    st.session_state.master_df = master_df
    st.session_state.overrides = {}
    st.session_state.warnings = all_warnings
    st.session_state.period = (year, month)
    st.session_state.missing_hours_employees = missing_employees
    st.session_state.missing_hours_dismissed = False

if st.session_state.master_df is None:
    st.info("Carica i 3 file e premi **Processa** per iniziare.")
    st.stop()

for w in st.session_state.warnings:
    st.warning(w)

if st.session_state.missing_hours_employees and not st.session_state.missing_hours_dismissed:
    col_w1, col_w2 = st.columns([12, 1])
    with col_w1:
        missing = st.session_state.missing_hours_employees
        n_interns = sum(1 for m in missing if m["is_intern"])
        n_regular = len(missing) - n_interns

        subject_parts = []
        if n_interns:
            subject_parts.append(f"{n_interns} praticante" if n_interns == 1 else f"{n_interns} praticanti")
        if n_regular:
            subject_parts.append(f"{n_regular} dipendente" if n_regular == 1 else f"{n_regular} dipendenti")
        subject = " e ".join(subject_parts)
        verb = "ha" if len(missing) == 1 else "hanno"

        names = ", ".join(m["label"] for m in missing[:10])
        more = f" e altri {len(missing) - 10}" if len(missing) > 10 else ""
        st.warning(
            f"{subject} {verb} ordini ma non risultano nel file ore "
            f"(trattati come 0 ore, quindi anomalia su tutti i loro ordini): {names}{more}."
        )
    with col_w2:
        if st.button("✕", key="dismiss_missing_hours", help="Chiudi avviso"):
            st.session_state.missing_hours_dismissed = True
            st.rerun()

master_df = st.session_state.master_df

st.subheader("Vista dipendenti — ordini e anomalie")

col_t1, col_t2 = st.columns(2)
with col_t1:
    show_only_anomalies = st.toggle("Mostra solo anomalie", value=True)
with col_t2:
    waive_intern_hours = st.toggle(
        "Abbona di default ore insufficienti praticanti (.intern@e80group.com)", value=True
    )

# Le forzature esplicite (st.session_state.overrides) restano quelle salvate
# dall'utente con "Applica forzature". L'abbono praticanti è un layer separato,
# calcolato ad ogni render in base al toggle: si somma alle forzature esplicite
# solo per il calcolo di viste/export, senza scriverle in session_state (così
# disattivare il toggle annulla subito l'effetto, senza lasciare residui).
combined_overrides = dict(st.session_state.overrides)
auto_waived_keys: set[tuple[str, date]] = set()
if waive_intern_hours:
    base_eff = apply_overrides(master_df, st.session_state.overrides, drop_fully_cancelled=False)
    intern_hours_anomaly = base_eff.apply(
        lambda r: is_intern(r["email"]) and r["anomaly_hours"], axis=1
    )
    for _, r in base_eff[intern_hours_anomaly].iterrows():
        key = (r["employee_id"], r["date"])
        if key not in st.session_state.overrides:
            auto_waived_keys.add(key)
        ov = dict(combined_overrides.get(key, {}))
        ov["waive_hours"] = True
        combined_overrides[key] = ov

# Vista editabile: non esclude le righe annullate da entrambi i fornitori, così
# restano visibili/correggibili finché non si riprocessa da capo.
edit_df = apply_overrides(master_df, combined_overrides, drop_fully_cancelled=False)
# Vista effettiva per recap ed export: le righe senza più alcun ordine sono escluse.
effective_df = apply_overrides(master_df, combined_overrides, drop_fully_cancelled=True)

with st.expander("Filtri"):
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        search_employee = st.text_input("Cerca dipendente (ID o nome)", value="")
    with col_f2:
        first_day = date(year, month, 1)
        last_day = date(year, month, calendar.monthrange(year, month)[1])
        date_range = st.date_input(
            "Intervallo date", value=(first_day, last_day),
            min_value=first_day, max_value=last_day,
        )

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

n_total = len(view)

if show_only_anomalies:
    view = view[view["Anomalia"] != ""]

if search_employee.strip():
    q = search_employee.strip().lower()
    view = view[
        view["employee_id"].str.lower().str.contains(q, na=False)
        | view["nombre"].str.lower().str.contains(q, na=False)
    ]

if isinstance(date_range, tuple) and len(date_range) == 2:
    d_from, d_to = date_range
    view = view[(view["date"] >= d_from) & (view["date"] <= d_to)]
elif isinstance(date_range, tuple) and len(date_range) == 1:
    # Range parzialmente selezionato (solo data di inizio): filtra da lì in poi.
    view = view[view["date"] >= date_range[0]]
elif isinstance(date_range, date):
    view = view[view["date"] == date_range]

st.caption(f"{len(view)} righe mostrate su {n_total} totali.")

view["Rimuovi Amati"] = view.apply(lambda r: bool(combined_overrides.get((r["employee_id"], r["date"]), {}).get("remove_amati")), axis=1)
view["Rimuovi Zippi"] = view.apply(lambda r: bool(combined_overrides.get((r["employee_id"], r["date"]), {}).get("remove_zippi")), axis=1)
view["Abbona vincolo ore"] = view.apply(lambda r: bool(combined_overrides.get((r["employee_id"], r["date"]), {}).get("waive_hours")), axis=1)

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
        # Una riga abbonata solo dal toggle praticanti (non toccata dall'utente)
        # non va resa una forzatura esplicita permanente: deve restare governata
        # dal toggle, altrimenti disattivarlo in seguito non avrebbe più effetto.
        if key in auto_waived_keys and ov == {"remove_amati": False, "remove_zippi": False, "waive_hours": True}:
            continue
        if any(ov.values()):
            new_overrides[key] = ov
        else:
            new_overrides.pop(key, None)
    st.session_state.overrides = new_overrides
    st.rerun()

st.subheader("Vista fornitori — validazione fatture")
st.caption("Conteggio pasti forniti")
recap = supplier_recap(effective_df)

kpi1, kpi2, kpi3 = st.columns(3)
total_amati = int(recap["Amati"].sum()) if not recap.empty else 0
total_zippi = int(recap["Zippi"].sum()) if not recap.empty else 0
kpi1.metric("Totale pasti Amati", total_amati)
kpi2.metric("Totale pasti Zippi", total_zippi)
kpi3.metric("Totale pasti mese", total_amati + total_zippi)

st.dataframe(recap, hide_index=True, use_container_width=True)

st.subheader("Export")
col1, col2 = st.columns(2)
year_sel, month_sel = st.session_state.period

with col1:
    reminder_bytes = build_reminder_export(effective_df)
    st.download_button(
        "Scarica export solleciti (anomalie)",
        data=reminder_bytes,
        file_name=f"meal_anomalies_{year_sel}-{month_sel:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with col2:
    summary_bytes = build_summary_export(effective_df, year_sel, month_sel)
    st.download_button(
        "Scarica riepilogo",
        data=summary_bytes,
        file_name=f"meal_summary_{year_sel}-{month_sel:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
