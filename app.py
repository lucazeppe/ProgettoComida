"""App Streamlit per il controllo dei pasti aziendali (fornitori Amati e Zippi)."""
from __future__ import annotations

import calendar
from datetime import date, datetime

import pandas as pd
import streamlit as st

import i18n
from business import (
    apply_overrides,
    build_master,
    find_non_company_email_ids,
    is_intern,
    supplier_recap,
)
from exports import build_mail_reminder_export, build_reminder_export, build_summary_export
from parsers import parse_amati, parse_hours, parse_zippi

if "lang" not in st.session_state:
    st.session_state.lang = "it"

st.set_page_config(page_title=i18n.t(st.session_state.lang, "page_title"), layout="wide")

with st.sidebar:
    st.radio(
        "🌐",
        options=["it", "es"],
        format_func=lambda code: i18n.FLAG_LABELS[code],
        horizontal=True,
        key="lang",
        label_visibility="collapsed",
    )

lang = st.session_state.lang

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
if "non_company_email_employees" not in st.session_state:
    st.session_state.non_company_email_employees = []
if "non_company_email_dismissed" not in st.session_state:
    st.session_state.non_company_email_dismissed = False

st.title(i18n.t(lang, "app_title"))

with st.sidebar:
    st.header(i18n.t(lang, "sidebar_period_header"))
    today = date.today()
    # Default: mese precedente a quello corrente (il mese "chiuso" da verificare).
    default_year, default_month = (today.year, today.month - 1) if today.month > 1 else (today.year - 1, 12)
    year = st.number_input(i18n.t(lang, "label_year"), min_value=2020, max_value=2100, value=default_year, step=1)
    month = st.selectbox(
        i18n.t(lang, "label_month"), options=list(range(1, 13)),
        format_func=lambda m: i18n.MONTHS[lang][m - 1], index=default_month - 1,
    )
    exclude_weekends = st.checkbox(i18n.t(lang, "checkbox_exclude_weekends"), value=True)

    st.header(i18n.t(lang, "sidebar_files_header"))
    zippi_file = st.file_uploader(i18n.t(lang, "uploader_zippi"), type=["xlsx"])
    amati_file = st.file_uploader(i18n.t(lang, "uploader_amati"), type=["xlsx"])
    hours_file = st.file_uploader(i18n.t(lang, "uploader_hours"), type=["xlsx"])

    process = st.button(i18n.t(lang, "button_process"), type="primary", use_container_width=True)

if process:
    if not (zippi_file and amati_file and hours_file):
        st.error(i18n.t(lang, "error_missing_files"))
        st.stop()

    all_warnings: list[str] = []
    try:
        zippi_df, w = parse_zippi(zippi_file, year, month, warn_weekends=exclude_weekends, lang=lang)
        all_warnings += w
        amati_df, w = parse_amati(amati_file, year, month, lang=lang)
        all_warnings += w
        hours_df, w = parse_hours(hours_file, year, month, warn_weekends=exclude_weekends, lang=lang)
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

    # Dipendenti con almeno un ordine effettuato con una email non aziendale
    # (calcolato sugli stessi file grezzi con lo stesso filtro weekend usato da
    # build_master, e poi ristretto ai soli ID con ordini effettivamente
    # processati, per lo stesso motivo del controllo sopra).
    non_company_ids = find_non_company_email_ids(zippi_df, amati_df, exclude_weekends=exclude_weekends)
    non_company_ids &= ordering_ids
    non_company_employees = []
    for emp_id in sorted(non_company_ids):
        row = master_df.loc[master_df["employee_id"] == emp_id]
        nombre = row["nombre"].iloc[0] if not row.empty and row["nombre"].iloc[0] else ""
        non_company_employees.append(f"{emp_id} ({nombre})" if nombre else emp_id)

    st.session_state.master_df = master_df
    st.session_state.overrides = {}
    st.session_state.warnings = all_warnings
    st.session_state.period = (year, month)
    st.session_state.missing_hours_employees = missing_employees
    st.session_state.missing_hours_dismissed = False
    st.session_state.non_company_email_employees = non_company_employees
    st.session_state.non_company_email_dismissed = False

if st.session_state.master_df is None:
    st.info(i18n.t(lang, "info_upload_prompt"))
    st.stop()

for w in st.session_state.warnings:
    st.warning(w)

if st.session_state.missing_hours_employees and not st.session_state.missing_hours_dismissed:
    col_w1, col_w2 = st.columns([12, 1])
    with col_w1:
        st.warning(i18n.missing_hours_warning(lang, st.session_state.missing_hours_employees))
    with col_w2:
        if st.button("✕", key="dismiss_missing_hours", help=i18n.t(lang, "dismiss_help")):
            st.session_state.missing_hours_dismissed = True
            st.rerun()

if st.session_state.non_company_email_employees and not st.session_state.non_company_email_dismissed:
    col_nc1, col_nc2 = st.columns([12, 1])
    with col_nc1:
        st.warning(i18n.non_company_email_warning(lang, st.session_state.non_company_email_employees))
    with col_nc2:
        if st.button("✕", key="dismiss_non_company_email", help=i18n.t(lang, "dismiss_help")):
            st.session_state.non_company_email_dismissed = True
            st.rerun()

master_df = st.session_state.master_df

st.subheader(i18n.t(lang, "subheader_employee_view"))

col_t1, col_t2 = st.columns(2)
with col_t1:
    show_only_anomalies = st.toggle(i18n.t(lang, "toggle_show_only_anomalies"), value=True)
with col_t2:
    waive_intern_hours = st.toggle(i18n.t(lang, "toggle_waive_intern_hours"), value=True)

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

with st.expander(i18n.t(lang, "expander_filters")):
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        search_employee = st.text_input(i18n.t(lang, "label_search_employee"), value="")
    with col_f2:
        first_day = date(year, month, 1)
        last_day = date(year, month, calendar.monthrange(year, month)[1])
        date_range = st.date_input(
            i18n.t(lang, "label_date_range"), value=(first_day, last_day),
            min_value=first_day, max_value=last_day,
        )

view = edit_df.copy()
view["Anomalia"] = view.apply(
    lambda r: ", ".join(
        filter(None, [
            i18n.t(lang, "anomaly_double") if r["anomaly_double"] else None,
            i18n.t(lang, "anomaly_hours") if r["anomaly_hours"] else None,
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

st.caption(i18n.t(lang, "caption_rows_shown", n=len(view), m=n_total))

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
        "employee_id": st.column_config.TextColumn(i18n.t(lang, "col_employee_id"), disabled=True),
        "nombre": st.column_config.TextColumn(i18n.t(lang, "col_name"), disabled=True),
        "email": st.column_config.TextColumn(i18n.t(lang, "col_email"), disabled=True),
        "date": st.column_config.DateColumn(i18n.t(lang, "col_date"), disabled=True),
        "amati": st.column_config.CheckboxColumn(i18n.t(lang, "col_amati"), disabled=True),
        "zippi": st.column_config.CheckboxColumn(i18n.t(lang, "col_zippi"), disabled=True),
        "hours": st.column_config.NumberColumn(i18n.t(lang, "col_hours"), disabled=True, format="%.1f"),
        "eligible": st.column_config.CheckboxColumn(i18n.t(lang, "col_eligible"), disabled=True),
        "Anomalia": st.column_config.TextColumn(i18n.t(lang, "col_anomaly"), disabled=True),
        "Rimuovi Amati": st.column_config.CheckboxColumn(i18n.t(lang, "col_remove_amati")),
        "Rimuovi Zippi": st.column_config.CheckboxColumn(i18n.t(lang, "col_remove_zippi")),
        "Abbona vincolo ore": st.column_config.CheckboxColumn(i18n.t(lang, "col_waive_hours")),
    },
    disabled=False,
    hide_index=True,
    use_container_width=True,
    key="employee_editor",
)

if st.button(i18n.t(lang, "button_apply_overrides")):
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

st.subheader(i18n.t(lang, "subheader_supplier_view"))
st.caption(i18n.t(lang, "caption_meal_count"))
recap = supplier_recap(effective_df)

kpi1, kpi2, kpi3 = st.columns(3)
total_amati = int(recap["Amati"].sum()) if not recap.empty else 0
total_zippi = int(recap["Zippi"].sum()) if not recap.empty else 0
kpi1.metric(i18n.t(lang, "metric_total_amati"), total_amati)
kpi2.metric(i18n.t(lang, "metric_total_zippi"), total_zippi)
kpi3.metric(i18n.t(lang, "metric_total_month"), total_amati + total_zippi)

st.dataframe(recap, hide_index=True, use_container_width=True)

st.subheader(i18n.t(lang, "subheader_export"))
col1, col2, col3 = st.columns(3)
year_sel, month_sel = st.session_state.period

with col1:
    reminder_bytes = build_reminder_export(effective_df)
    st.download_button(
        i18n.t(lang, "button_export_reminders"),
        data=reminder_bytes,
        file_name=f"meal_anomalies_{year_sel}-{month_sel:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with col2:
    summary_bytes = build_summary_export(effective_df, year_sel, month_sel)
    st.download_button(
        i18n.t(lang, "button_export_summary"),
        data=summary_bytes,
        file_name=f"meal_summary_{year_sel}-{month_sel:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with col3:
    mail_reminder_bytes = build_mail_reminder_export(effective_df)
    st.download_button(
        i18n.t(lang, "button_export_mail_reminder"),
        data=mail_reminder_bytes,
        file_name="mail_reminder.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
