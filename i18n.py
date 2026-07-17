"""Traduzioni IT/ES per il frontend (app.py) e i messaggi utente che vi arrivano
(errori/warning di parsers.py, etichette anomalia di business.py).

Gli export Excel (exports.py) restano inglese fisso e NON usano questo modulo.
"""
from __future__ import annotations

DEFAULT_LANG = "it"

FLAG_LABELS = {"it": "🇮🇹 Italiano", "es": "🇪🇸 Español"}

MONTHS = {
    "it": ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
           "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"],
    "es": ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
           "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"],
}

STRINGS = {
    "it": {
        "page_title": "Controllo Pasti Aziendali",
        "app_title": "Controllo Pasti Aziendali — Amati / Zippi",
        "sidebar_period_header": "Periodo di riferimento",
        "label_year": "Anno",
        "label_month": "Mese",
        "checkbox_exclude_weekends": "Escludi sabato/domenica dal calcolo (segnala se contengono dati)",
        "sidebar_files_header": "File di input",
        "uploader_zippi": "Ordini Zippi (.xlsx)",
        "uploader_amati": "Ordini Amati (.xlsx)",
        "uploader_hours": "Ore lavorate (.xlsx)",
        "button_process": "Processa",
        "error_missing_files": "Carica tutti e 3 i file prima di procedere.",
        "info_upload_prompt": "Carica i 3 file e premi **Processa** per iniziare.",
        "dismiss_help": "Chiudi avviso",
        "subheader_employee_view": "Vista dipendenti — ordini e anomalie",
        "toggle_show_only_anomalies": "Mostra solo anomalie",
        "toggle_waive_intern_hours": "Abbona di default ore insufficienti praticanti (.intern@e80group.com)",
        "expander_filters": "Filtri",
        "label_search_employee": "Cerca dipendente (ID o nome)",
        "label_date_range": "Intervallo date",
        "caption_rows_shown": "{n} righe mostrate su {m} totali.",
        "col_employee_id": "ID Dipendente",
        "col_name": "Nome",
        "col_email": "Email",
        "col_date": "Data",
        "col_amati": "Amati",
        "col_zippi": "Zippi",
        "col_hours": "Ore",
        "col_eligible": "Diritto (>4h)",
        "col_anomaly": "Anomalia",
        "col_remove_amati": "Rimuovi Amati",
        "col_remove_zippi": "Rimuovi Zippi",
        "col_waive_hours": "Abbona vincolo ore",
        "button_apply_overrides": "Applica forzature",
        "subheader_supplier_view": "Vista fornitori — validazione fatture",
        "caption_meal_count": "Conteggio pasti forniti",
        "metric_total_amati": "Totale pasti Amati",
        "metric_total_zippi": "Totale pasti Zippi",
        "metric_total_month": "Totale pasti mese",
        "subheader_export": "Export",
        "button_export_reminders": "Scarica anomalie",
        "button_export_summary": "Scarica riepilogo",
        "button_export_mail_reminder": "Scarica sollecito mail",
        "anomaly_double": "Doppia prenotazione",
        "anomaly_hours": "Ore ufficio insufficienti",
        "on_date_prefix": "il",
        "err_zippi_no_header": "File Zippi: non trovo la riga di intestazione (colonna 'Nombre') nelle prime 15 righe del foglio.",
        "err_zippi_no_id_col": "File Zippi: colonna 'Numero de empleado' non trovata.",
        "err_zippi_no_email_col": "File Zippi: colonna 'Correo' non trovata.",
        "err_zippi_out_of_month": "File Zippi: trovate date fuori dal mese scelto ({year}-{month:02d}), es: {examples}. Verifica di aver caricato il file del mese corretto.",
        "warn_zippi_weekend_orders": "File Zippi: trovati {n} ordini nel weekend (sabato/domenica), inattesi.",
        "err_amati_sheet_week_mismatch": "File Amati: trovati {n} fogli ma il mese {year}-{month:02d} ha {m} settimane lavorative attese. Verifica il file/mese scelto.",
        "err_amati_bad_header": "File Amati, foglio '{sheet_name}': intestazione inattesa ({e}).",
        "warn_amati_sheet_order": "File Amati: i fogli non sembrano in ordine cronologico in base al numero nel nome ({sheet_names}). Le settimane potrebbero essere attribuite ai fogli sbagliati: verifica l'ordine delle schede nel file.",
        "warn_amati_order_out_of_month": "File Amati, foglio '{sheet_name}': ordine per {nombre} {apellidos} in un giorno ({weekday}) fuori dal mese {year}-{month:02d}, ignorato.",
        "err_hours_no_id_col": "File Ore: non trovo una colonna ID dipendente riconosciuta (cercate: {candidates}) nelle prime righe del file.",
        "err_hours_no_date_header": "File Ore: nessuna colonna con intestazione data trovata.",
        "err_hours_no_month_date": "File Ore: nessuna colonna con data del mese scelto ({year}-{month:02d}) trovata. Verifica di aver caricato il file del mese corretto.",
        "err_hours_bad_value": "File Ore: valore ore non riconosciuto ('{raw_hours}') per il dipendente {emp_id} il {date}. Atteso un numero (es. 8 o 8,5).",
        "err_hours_duplicate_ids": "File Ore: trovati ID dipendente duplicati (righe diverse per lo stesso ID): {examples}. Verifica il file, ogni dipendente deve avere una sola riga.",
        "warn_hours_missing_columns": "File Ore: mancano colonne per {n} giorno/i lavorativo/i di {year}-{month:02d} (es: {examples}) — quei giorni saranno trattati come 0 ore.",
        "warn_hours_weekend_rows": "File Ore: trovate {n} righe con ore nel weekend (sabato/domenica), inattese.",
        "mh_intern_singular": "praticante",
        "mh_intern_plural": "praticanti",
        "mh_regular_singular": "dipendente",
        "mh_regular_plural": "dipendenti",
        "mh_verb_singular": "ha",
        "mh_verb_plural": "hanno",
        "mh_and": "e",
        "mh_body": "{subject} {verb} ordini ma non risultano nel file ore (trattati come 0 ore, quindi anomalia su tutti i loro ordini): {names}{more}.",
        "mh_more": " e altri {count}",
    },
    "es": {
        "page_title": "Control de Comidas de la Empresa",
        "app_title": "Control de Comidas de la Empresa — Amati / Zippi",
        "sidebar_period_header": "Período de referencia",
        "label_year": "Año",
        "label_month": "Mes",
        "checkbox_exclude_weekends": "Excluir sábado/domingo del cálculo (avisa si contienen datos)",
        "sidebar_files_header": "Archivos de entrada",
        "uploader_zippi": "Pedidos Zippi (.xlsx)",
        "uploader_amati": "Pedidos Amati (.xlsx)",
        "uploader_hours": "Horas trabajadas (.xlsx)",
        "button_process": "Procesar",
        "error_missing_files": "Carga los 3 archivos antes de continuar.",
        "info_upload_prompt": "Carga los 3 archivos y presiona **Procesar** para comenzar.",
        "dismiss_help": "Cerrar aviso",
        "subheader_employee_view": "Vista de empleados — pedidos y anomalías",
        "toggle_show_only_anomalies": "Mostrar solo anomalías",
        "toggle_waive_intern_hours": "Perdonar por defecto las horas insuficientes de los practicantes (.intern@e80group.com)",
        "expander_filters": "Filtros",
        "label_search_employee": "Buscar empleado (ID o nombre)",
        "label_date_range": "Rango de fechas",
        "caption_rows_shown": "{n} filas mostradas de {m} totales.",
        "col_employee_id": "ID Empleado",
        "col_name": "Nombre",
        "col_email": "Email",
        "col_date": "Fecha",
        "col_amati": "Amati",
        "col_zippi": "Zippi",
        "col_hours": "Horas",
        "col_eligible": "Derecho (>4h)",
        "col_anomaly": "Anomalía",
        "col_remove_amati": "Quitar Amati",
        "col_remove_zippi": "Quitar Zippi",
        "col_waive_hours": "Perdonar requisito de horas",
        "button_apply_overrides": "Aplicar excepciones manuales",
        "subheader_supplier_view": "Vista de proveedores — validación de facturas",
        "caption_meal_count": "Recuento de comidas suministradas",
        "metric_total_amati": "Total comidas Amati",
        "metric_total_zippi": "Total comidas Zippi",
        "metric_total_month": "Total comidas del mes",
        "subheader_export": "Exportar",
        "button_export_reminders": "Descargar anomalías",
        "button_export_summary": "Descargar resumen",
        "button_export_mail_reminder": "Descargar aviso por correo",
        "anomaly_double": "Doble reserva",
        "anomaly_hours": "Horas de oficina insuficientes",
        "on_date_prefix": "el",
        "err_zippi_no_header": "Archivo Zippi: no se encuentra la fila de encabezado (columna 'Nombre') en las primeras 15 filas de la hoja.",
        "err_zippi_no_id_col": "Archivo Zippi: no se encontró la columna 'Numero de empleado'.",
        "err_zippi_no_email_col": "Archivo Zippi: no se encontró la columna 'Correo'.",
        "err_zippi_out_of_month": "Archivo Zippi: se encontraron fechas fuera del mes seleccionado ({year}-{month:02d}), ej: {examples}. Verifica haber cargado el archivo del mes correcto.",
        "warn_zippi_weekend_orders": "Archivo Zippi: se encontraron {n} pedidos en el fin de semana (sábado/domingo), inesperados.",
        "err_amati_sheet_week_mismatch": "Archivo Amati: se encontraron {n} hojas pero el mes {year}-{month:02d} tiene {m} semanas laborales esperadas. Verifica el archivo/mes seleccionado.",
        "err_amati_bad_header": "Archivo Amati, hoja '{sheet_name}': encabezado inesperado ({e}).",
        "warn_amati_sheet_order": "Archivo Amati: las hojas no parecen estar en orden cronológico según el número en el nombre ({sheet_names}). Las semanas podrían atribuirse a las hojas equivocadas: verifica el orden de las pestañas en el archivo.",
        "warn_amati_order_out_of_month": "Archivo Amati, hoja '{sheet_name}': pedido de {nombre} {apellidos} en un día ({weekday}) fuera del mes {year}-{month:02d}, ignorado.",
        "err_hours_no_id_col": "Archivo de Horas: no se encuentra una columna de ID de empleado reconocida (buscadas: {candidates}) en las primeras filas del archivo.",
        "err_hours_no_date_header": "Archivo de Horas: no se encontró ninguna columna con encabezado de fecha.",
        "err_hours_no_month_date": "Archivo de Horas: no se encontró ninguna columna con fecha del mes seleccionado ({year}-{month:02d}). Verifica haber cargado el archivo del mes correcto.",
        "err_hours_bad_value": "Archivo de Horas: valor de horas no reconocido ('{raw_hours}') para el empleado {emp_id} el {date}. Se esperaba un número (ej. 8 u 8,5).",
        "err_hours_duplicate_ids": "Archivo de Horas: se encontraron ID de empleado duplicados (filas distintas para el mismo ID): {examples}. Verifica el archivo, cada empleado debe tener una sola fila.",
        "warn_hours_missing_columns": "Archivo de Horas: faltan columnas para {n} día(s) laborable(s) de {year}-{month:02d} (ej: {examples}) — esos días se tratarán como 0 horas.",
        "warn_hours_weekend_rows": "Archivo de Horas: se encontraron {n} filas con horas en el fin de semana (sábado/domingo), inesperadas.",
        "mh_intern_singular": "practicante",
        "mh_intern_plural": "practicantes",
        "mh_regular_singular": "empleado",
        "mh_regular_plural": "empleados",
        "mh_verb_singular": "tiene",
        "mh_verb_plural": "tienen",
        "mh_and": "y",
        "mh_body": "{subject} {verb} pedidos pero no aparecen en el archivo de horas (tratados como 0 horas, por lo tanto anomalía en todos sus pedidos): {names}{more}.",
        "mh_more": " y otros {count}",
    },
}


def t(lang: str, key: str, **kwargs) -> str:
    """Cerca STRINGS[lang][key] e applica .format(**kwargs).
    Fallback su italiano se manca la lingua o la chiave (mai un crash)."""
    lang_strings = STRINGS.get(lang, STRINGS[DEFAULT_LANG])
    template = lang_strings.get(key, STRINGS[DEFAULT_LANG].get(key, key))
    return template.format(**kwargs)


def missing_hours_warning(lang: str, missing: list[dict]) -> str:
    """Compone la frase 'dipendenti con ordini ma assenti dal file ore'.
    Concordanza singolare/plurale su 2 variabili (praticanti/dipendenti),
    non riducibile a un singolo template — per questo è una funzione dedicata."""
    n_interns = sum(1 for m in missing if m["is_intern"])
    n_regular = len(missing) - n_interns

    subject_parts = []
    if n_interns:
        key = "mh_intern_singular" if n_interns == 1 else "mh_intern_plural"
        subject_parts.append(f"{n_interns} {t(lang, key)}")
    if n_regular:
        key = "mh_regular_singular" if n_regular == 1 else "mh_regular_plural"
        subject_parts.append(f"{n_regular} {t(lang, key)}")
    subject = f" {t(lang, 'mh_and')} ".join(subject_parts)

    verb = t(lang, "mh_verb_singular" if len(missing) == 1 else "mh_verb_plural")
    names = ", ".join(m["label"] for m in missing[:10])
    more = t(lang, "mh_more", count=len(missing) - 10) if len(missing) > 10 else ""

    return t(lang, "mh_body", subject=subject, verb=verb, names=names, more=more)
