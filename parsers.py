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

ID_COLUMN_CANDIDATES = [
    "dipendente", "empleado", "employee id", "numero de empleado", "id",
    "employee", "id dipendente",
]


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

def parse_zippi(file, year: int, month: int, warn_weekends: bool = True) -> tuple[pd.DataFrame, list[str]]:
    warnings: list[str] = []
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb[wb.sheetnames[0]]

    header_row_idx = None
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), start=1):
        if row and row[0] and "nombre" in str(row[0]).strip().lower():
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError(
            "File Zippi: non trovo la riga di intestazione (colonna 'Nombre') "
            "nelle prime 15 righe del foglio."
        )

    header = [c for c in ws[header_row_idx]]
    header_values = [c.value for c in header]
    lower_header = [str(v).strip().lower() if v else "" for v in header_values]

    try:
        col_id = lower_header.index("numero de empleado")
    except ValueError:
        try:
            col_id = next(i for i, v in enumerate(lower_header) if "numero de empleado" in v)
        except StopIteration:
            raise ValueError("File Zippi: colonna 'Numero de empleado' non trovata.")
    try:
        col_email = lower_header.index("correo")
    except ValueError:
        raise ValueError("File Zippi: colonna 'Correo' non trovata.")

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
        examples = ", ".join(f"{n} il {d.isoformat()}" for _, n, d in out_of_month_dates[:5])
        raise ValueError(
            f"File Zippi: trovate date fuori dal mese scelto ({year}-{month:02d}), "
            f"es: {examples}. Verifica di aver caricato il file del mese corretto."
        )

    df = pd.DataFrame(records, columns=["employee_id", "nombre", "email", "date"])

    if warn_weekends:
        weekend_rows = df[df["date"].apply(lambda d: d.weekday() >= 5)]
        if not weekend_rows.empty:
            warnings.append(
                f"File Zippi: trovati {len(weekend_rows)} ordini nel weekend (sabato/domenica), inattesi."
            )

    return df, warnings


# ---------------------------------------------------------------------------
# AMATI: un foglio per settimana, colonne Lunes..Viernes con 0/1
# ---------------------------------------------------------------------------

WEEKDAY_COLS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]


def parse_amati(file, year: int, month: int) -> tuple[pd.DataFrame, list[str]]:
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
            f"File Amati: trovati {len(sheet_names)} fogli ma il mese {year}-{month:02d} "
            f"ha {len(padded_weeks)} settimane lavorative attese. Verifica il file/mese scelto."
        )

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
            raise ValueError(f"File Amati, foglio '{sheet_name}': intestazione inattesa ({e}).")

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
                    warnings.append(
                        f"File Amati, foglio '{sheet_name}': ordine per {nombre} {apellidos} "
                        f"in un giorno ({WEEKDAY_COLS[wd_idx]}) fuori dal mese {year}-{month:02d}, ignorato."
                    )
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

def parse_hours(file, year: int, month: int, warn_weekends: bool = True) -> tuple[pd.DataFrame, list[str]]:
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
        for j, v in enumerate(row):
            if v and str(v).strip().lower() in ID_COLUMN_CANDIDATES:
                id_row_idx = i
                id_col_idx = j
                break
        if id_row_idx is not None:
            break

    if id_row_idx is None:
        raise ValueError(
            "File Ore: non trovo una colonna ID dipendente riconosciuta "
            f"(cercate: {ID_COLUMN_CANDIDATES}) nelle prime righe del file."
        )

    date_row_idx = None
    best_count = 0
    for i in range(1, 11):
        row = [c.value for c in ws[i]]
        count = sum(1 for v in row if hasattr(v, "date"))
        if count > best_count:
            best_count = count
            date_row_idx = i
    if date_row_idx is None or best_count == 0:
        raise ValueError("File Ore: nessuna colonna con intestazione data trovata.")

    date_row = [c.value for c in ws[date_row_idx]]
    date_cols = {
        j: v.date() for j, v in enumerate(date_row)
        if hasattr(v, "date")
    }

    header_row_idx = max(id_row_idx, date_row_idx)

    out_of_month = [d for d in date_cols.values() if not _in_month(d, year, month)]
    if out_of_month:
        examples = ", ".join(sorted(d.isoformat() for d in set(out_of_month))[:5])
        raise ValueError(
            f"File Ore: trovate colonne con date fuori dal mese scelto ({year}-{month:02d}), "
            f"es: {examples}. Verifica di aver caricato il file del mese corretto."
        )

    records = []
    bad_ids = []
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        if row[id_col_idx] is None:
            continue
        emp_id = _normalize_id(row[id_col_idx])
        if emp_id is None:
            continue
        if not re.fullmatch(r"\d{4,10}", emp_id):
            bad_ids.append(emp_id)
            continue
        for j, d in date_cols.items():
            raw_hours = row[j] if j < len(row) else None
            try:
                hours = _parse_hours_value(raw_hours)
            except (ValueError, TypeError):
                raise ValueError(
                    f"File Ore: valore ore non riconosciuto ('{raw_hours}') per il dipendente "
                    f"{emp_id} il {d.isoformat()}. Atteso un numero (es. 8 o 8,5)."
                )
            records.append({"employee_id": emp_id, "date": d, "hours": hours})

    if bad_ids:
        raise ValueError(
            "File Ore: formato ID dipendente non riconosciuto (atteso numerico 2000xxx), "
            f"trovati es: {bad_ids[:5]}. Verifica di aver caricato l'export ore con ID aziendale corretto."
        )

    df = pd.DataFrame(records, columns=["employee_id", "date", "hours"])

    if warn_weekends:
        weekend_rows = df[(df["hours"] > 0) & (df["date"].apply(lambda d: d.weekday() >= 5))]
        if not weekend_rows.empty:
            warnings.append(
                f"File Ore: trovate {len(weekend_rows)} righe con ore nel weekend (sabato/domenica), inattese."
            )

    return df, warnings
