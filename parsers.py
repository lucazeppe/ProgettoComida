"""Parser per i 3 file di input: Zippi, Amati, Ore lavorate.

Ogni parser ritorna (DataFrame, warnings: list[str]) e solleva ValueError
con messaggio descrittivo per le incoerenze bloccanti.
"""
from __future__ import annotations

import calendar
import re
from datetime import date

import openpyxl
import pandas as pd

import i18n

# Ordine di priorità: "id" (colonna con l'Employee ID 2000xxx aziendale) deve
# vincere su "dipendente"/"empleado" quando entrambe compaiono nella stessa
# intestazione, perché quest'ultime sono spesso un numero dipendente interno
# più corto (es. 100) e NON la chiave di match richiesta.
ID_COLUMN_CANDIDATES_PRIORITY = ["id", "employee id", "id dipendente", "cod ssff"]
ID_COLUMN_CANDIDATES_FALLBACK = ["dipendente", "empleado", "numero de empleado", "employee"]
ID_COLUMN_CANDIDATES = ID_COLUMN_CANDIDATES_PRIORITY + ID_COLUMN_CANDIDATES_FALLBACK


def title_case(s: str) -> str:
    """Ogni parola con sola iniziale maiuscola (gestisce nomi/cognomi con più parole,
    e trattini/apostrofi come separatori, es. "GARCIA-LOPEZ" -> "Garcia-Lopez")."""
    def _cap_word(w: str) -> str:
        parts = re.split(r"([-'])", w)
        return "".join(p if p in ("-", "'") else (p[:1].upper() + p[1:].lower()) for p in parts)
    return " ".join(_cap_word(w) for w in s.split() if w)


def _normalize_id(value) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, float):
        return str(int(value))
    s = str(value).strip()
    if s == "" or s.lower() == "none":
        return None
    if s.endswith(".0"):
        s = s[:-2]
    return s


def month_business_days(year: int, month: int) -> list[date]:
    """Tutti i giorni Lun-Ven del mese scelto, in ordine."""
    n_days = calendar.monthrange(year, month)[1]
    return [
        d for d in (date(year, month, day) for day in range(1, n_days + 1))
        if d.weekday() < 5
    ]


def month_all_days(year: int, month: int) -> list[date]:
    n_days = calendar.monthrange(year, month)[1]
    return [date(year, month, day) for day in range(1, n_days + 1)]


def _in_month(d: date, year: int, month: int) -> bool:
    return d.year == year and d.month == month


def _parse_hours_value(value) -> float:
    """Converte il valore ore di una cella in float. Gestisce None, numeri,
    stringhe con virgola decimale (es. '8,5') e valori orari (datetime.time)."""
    if value is None:
        return 0.0
    if isinstance(value, str):
        s = value.strip().replace(",", ".")
        if s == "":
            return 0.0
        return float(s)
    if hasattr(value, "hour") and hasattr(value, "minute"):
        return value.hour + value.minute / 60 + getattr(value, "second", 0) / 3600
    return float(value)


# ---------------------------------------------------------------------------
# ZIPPI: righe = dipendente, colonne D..fine = date di ordine (valore datetime)
# ---------------------------------------------------------------------------

def parse_zippi(file, year: int, month: int, warn_weekends: bool = True, lang: str = "it") -> tuple[pd.DataFrame, list[str]]:
    warnings: list[str] = []
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb[wb.sheetnames[0]]

    header_row_idx = None
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), start=1):
        if row and row[0] and "nombre" in str(row[0]).strip().lower():
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError(i18n.t(lang, "err_zippi_no_header"))

    header = [c for c in ws[header_row_idx]]
    header_values = [c.value for c in header]
    lower_header = [str(v).strip().lower() if v else "" for v in header_values]

    try:
        col_id = lower_header.index("numero de empleado")
    except ValueError:
        try:
            col_id = next(i for i, v in enumerate(lower_header) if "numero de empleado" in v)
        except StopIteration:
            raise ValueError(i18n.t(lang, "err_zippi_no_id_col"))
    try:
        col_email = lower_header.index("correo")
    except ValueError:
        raise ValueError(i18n.t(lang, "err_zippi_no_email_col"))

    date_col_start = max(col_id, col_email) + 1

    records = []
    out_of_month_dates = []
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        nombre = row[0]
        if nombre is None or str(nombre).strip() == "":
            continue  # riga vuota residua (es. "PRIMERA SEMANA"): salta, non è fine dati
        emp_id = _normalize_id(row[col_id])
        email = row[col_email]
        if emp_id is None:
            continue
        for cell in row[date_col_start:]:
            if cell is None:
                continue
            if hasattr(cell, "date"):
                d = cell.date() if hasattr(cell, "date") else cell
            else:
                continue
            if not _in_month(d, year, month):
                out_of_month_dates.append((emp_id, str(nombre), d))
                continue
            records.append({
                "employee_id": emp_id,
                "nombre": str(nombre).strip(),
                "email": str(email).strip() if email else None,
                "date": d,
            })

    if out_of_month_dates:
        prefix = i18n.t(lang, "on_date_prefix")
        examples = ", ".join(f"{n} {prefix} {d.isoformat()}" for _, n, d in out_of_month_dates[:5])
        raise ValueError(i18n.t(lang, "err_zippi_out_of_month", year=year, month=month, examples=examples))

    df = pd.DataFrame(records, columns=["employee_id", "nombre", "email", "date"])

    if warn_weekends:
        weekend_rows = df[df["date"].apply(lambda d: d.weekday() >= 5)]
        if not weekend_rows.empty:
            warnings.append(i18n.t(lang, "warn_zippi_weekend_orders", n=len(weekend_rows)))

    return df, warnings


# ---------------------------------------------------------------------------
# AMATI: un foglio per settimana, colonne Lunes..Viernes con 0/1
# ---------------------------------------------------------------------------

WEEKDAY_COLS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]


def parse_amati(file, year: int, month: int, lang: str = "it") -> tuple[pd.DataFrame, list[str]]:
    warnings: list[str] = []
    wb = openpyxl.load_workbook(file, data_only=True)

    business_days = month_business_days(year, month)
    weeks: list[list[date]] = []
    current_week: list[date] = []
    for d in business_days:
        if current_week and d.weekday() == 0:
            weeks.append(current_week)
            current_week = []
        current_week.append(d)
    if current_week:
        weeks.append(current_week)
    # Pad ogni settimana a Lun..Ven (5 slot) allineando sul weekday
    padded_weeks = []
    for week in weeks:
        slots: list[date | None] = [None] * 5
        for d in week:
            slots[d.weekday()] = d
        padded_weeks.append(slots)

    sheet_names = wb.sheetnames
    if len(sheet_names) != len(padded_weeks):
        raise ValueError(
            i18n.t(lang, "err_amati_sheet_week_mismatch", n=len(sheet_names), year=year, month=month, m=len(padded_weeks))
        )

    # I fogli vengono abbinati alle settimane in base al solo ordine delle tab
    # nel file. Se il nome del foglio contiene un numero (es. "Hoja 3"), verifica
    # che sia crescente nell'ordine delle tab: altrimenti avvisa (non blocca),
    # perché uno scambio di tab produrrebbe un'attribuzione settimana errata
    # senza nessun altro segnale visibile.
    sheet_numbers = [re.search(r"\d+", name) for name in sheet_names]
    if all(sheet_numbers):
        numbers = [int(m.group()) for m in sheet_numbers]
        if numbers != sorted(numbers):
            warnings.append(i18n.t(lang, "warn_amati_sheet_order", sheet_names=", ".join(sheet_names)))

    records = []
    for sheet_name, week_slots in zip(sheet_names, padded_weeks):
        ws = wb[sheet_name]
        header = [c.value for c in ws[1]]
        lower_header = [str(v).strip().lower() if v else "" for v in header]
        try:
            col_nombre = lower_header.index("nombre")
            col_apellidos = lower_header.index("apellidos")
            col_id = lower_header.index("empleado")
            col_email = lower_header.index("correo")
        except ValueError as e:
            raise ValueError(i18n.t(lang, "err_amati_bad_header", sheet_name=sheet_name, e=e))

        day_cols = {}
        for wd_idx, wd_name in enumerate(WEEKDAY_COLS):
            try:
                day_cols[wd_idx] = lower_header.index(wd_name.lower())
            except ValueError:
                day_cols[wd_idx] = None

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[col_nombre] is None or str(row[col_nombre]).strip() == "":
                continue
            emp_id = _normalize_id(row[col_id])
            if emp_id is None:
                continue
            nombre = str(row[col_nombre]).strip()
            apellidos = str(row[col_apellidos]).strip() if row[col_apellidos] else ""
            email = row[col_email]
            for wd_idx, col_idx in day_cols.items():
                if col_idx is None:
                    continue
                value = row[col_idx]
                if value is None:
                    continue
                if isinstance(value, str):
                    if value.strip() in ("", "0"):
                        continue
                elif not value:
                    continue
                target_date = week_slots[wd_idx]
                if target_date is None:
                    # Giorno del foglio settimanale fuori dal mese scelto (sconfinamento fisiologico)
                    warnings.append(i18n.t(
                        lang, "warn_amati_order_out_of_month",
                        sheet_name=sheet_name, nombre=nombre, apellidos=apellidos,
                        weekday=WEEKDAY_COLS[wd_idx], year=year, month=month,
                    ))
                    continue
                records.append({
                    "employee_id": emp_id,
                    "nombre": nombre,
                    "apellidos": apellidos,
                    "email": str(email).strip() if email else None,
                    "date": target_date,
                })

    df = pd.DataFrame(records, columns=["employee_id", "nombre", "apellidos", "email", "date"])
    return df, warnings


# ---------------------------------------------------------------------------
# ORE LAVORATE: colonna ID + colonne datetime con ore
# ---------------------------------------------------------------------------

def parse_hours(file, year: int, month: int, warn_weekends: bool = True, lang: str = "it") -> tuple[pd.DataFrame, list[str]]:
    warnings: list[str] = []
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb[wb.sheetnames[0]]

    # L'export tipico ha righe di intestazione separate: una riga "titolo" con le
    # date (una per colonna giorno) e un'altra riga con i nomi campo (ID, Nome...).
    # Non si può assumere che siano la stessa riga: si cercano indipendentemente
    # nelle prime righe del file.
    id_row_idx = None
    id_col_idx = None
    for i in range(1, 11):
        row = [c.value for c in ws[i]]
        lower_row = [str(v).strip().lower() if v else "" for v in row]
        # Prima cerca una colonna "ID" prioritaria; solo se assente in questa
        # riga, ripiega su "Dipendente"/"Empleado" (numero interno, non ideale).
        match_j = next((j for j, v in enumerate(lower_row) if v in ID_COLUMN_CANDIDATES_PRIORITY), None)
        if match_j is None:
            match_j = next((j for j, v in enumerate(lower_row) if v in ID_COLUMN_CANDIDATES_FALLBACK), None)
        if match_j is not None:
            id_row_idx = i
            id_col_idx = match_j
            break

    if id_row_idx is None:
        raise ValueError(i18n.t(lang, "err_hours_no_id_col", candidates=ID_COLUMN_CANDIDATES))

    # Nome/Cognome (se presenti) sono nella stessa riga della colonna ID.
    id_header_row = [c.value for c in ws[id_row_idx]]
    lower_id_header_row = [str(v).strip().lower() if v else "" for v in id_header_row]
    nome_col_idx = lower_id_header_row.index("nome") if "nome" in lower_id_header_row else None
    cognome_col_idx = lower_id_header_row.index("cognome") if "cognome" in lower_id_header_row else None

    date_row_idx = None
    best_count = 0
    for i in range(1, 11):
        row = [c.value for c in ws[i]]
        count = sum(1 for v in row if hasattr(v, "date"))
        if count > best_count:
            best_count = count
            date_row_idx = i
    if date_row_idx is None or best_count == 0:
        raise ValueError(i18n.t(lang, "err_hours_no_date_header"))

    date_row = [c.value for c in ws[date_row_idx]]
    all_date_cols = {
        j: v.date() for j, v in enumerate(date_row)
        if hasattr(v, "date")
    }

    header_row_idx = max(id_row_idx, date_row_idx)

    # Il file ore è solo di supporto (serve unicamente a leggere le ore dei
    # giorni del mese scelto): eventuali colonne di altri mesi (es. l'export
    # sconfina nel mese successivo) vengono ignorate silenziosamente, non
    # generano errore né warning.
    date_cols = {j: d for j, d in all_date_cols.items() if _in_month(d, year, month)}
    if not date_cols:
        raise ValueError(i18n.t(lang, "err_hours_no_month_date", year=year, month=month))

    missing_days = sorted(set(month_business_days(year, month)) - set(date_cols.values()))
    if missing_days:
        examples = ", ".join(d.isoformat() for d in missing_days[:5])
        warnings.append(i18n.t(lang, "warn_hours_missing_columns", n=len(missing_days), year=year, month=month, examples=examples))

    records = []
    seen_ids = set()
    duplicate_ids = set()
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        if row[id_col_idx] is None:
            continue
        emp_id = _normalize_id(row[id_col_idx])
        if emp_id is None:
            continue
        if not re.fullmatch(r"\d{4,10}", emp_id):
            # Formato ID non riconosciuto: riga ignorata silenziosamente (il file
            # ore è di supporto, non blocca il processamento). L'unico errore
            # bloccante legato all'ID è a valle: un dipendente con ordine ma
            # nessuna corrispondenza nel file ore (vedi business.build_master).
            continue
        if emp_id in seen_ids:
            duplicate_ids.add(emp_id)
        seen_ids.add(emp_id)

        nombre_ore = None
        if nome_col_idx is not None or cognome_col_idx is not None:
            nome_val = row[nome_col_idx] if nome_col_idx is not None and nome_col_idx < len(row) else None
            cognome_val = row[cognome_col_idx] if cognome_col_idx is not None and cognome_col_idx < len(row) else None
            parts = [
                title_case(str(v).strip()) for v in (nome_val, cognome_val)
                if v is not None and str(v).strip() != ""
            ]
            nombre_ore = " ".join(parts) if parts else None

        for j, d in date_cols.items():
            raw_hours = row[j] if j < len(row) else None
            try:
                hours = _parse_hours_value(raw_hours)
            except (ValueError, TypeError):
                raise ValueError(i18n.t(lang, "err_hours_bad_value", raw_hours=raw_hours, emp_id=emp_id, date=d.isoformat()))
            records.append({"employee_id": emp_id, "date": d, "hours": hours, "nombre_ore": nombre_ore})

    if duplicate_ids:
        examples = ", ".join(sorted(duplicate_ids)[:10])
        raise ValueError(i18n.t(lang, "err_hours_duplicate_ids", examples=examples))

    df = pd.DataFrame(records, columns=["employee_id", "date", "hours", "nombre_ore"])

    if warn_weekends:
        weekend_rows = df[(df["hours"] > 0) & (df["date"].apply(lambda d: d.weekday() >= 5))]
        if not weekend_rows.empty:
            warnings.append(i18n.t(lang, "warn_hours_weekend_rows", n=len(weekend_rows)))

    return df, warnings
