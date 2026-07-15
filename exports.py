"""Generazione dei 2 export Excel: solleciti anomalie e riepilogo colorato per fornitore."""
from __future__ import annotations

import calendar
from datetime import date
from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from business import ANOMALY_DOUBLE_LABEL, ANOMALY_HOURS_LABEL

GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
HEADER_FONT = Font(bold=True)


def build_reminder_export(master_df: pd.DataFrame) -> bytes:
    """ID dipendente, nome, email, giorno anomalia, ragione anomalia.

    Un giorno con entrambe le anomalie genera 2 record distinti.
    master_df deve essere già lo stato effettivo (dopo eventuali forzature).
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Anomalie"
    headers = ["ID Dipendente", "Nome", "Email", "Giorno anomalia", "Ragione anomalia"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = HEADER_FONT

    for _, row in master_df.sort_values(["employee_id", "date"]).iterrows():
        email = row["email"] if pd.notna(row["email"]) else ""
        if row["anomaly_double"]:
            ws.append([row["employee_id"], row["nombre"], email, row["date"], ANOMALY_DOUBLE_LABEL])
        if row["anomaly_hours"]:
            ws.append([row["employee_id"], row["nombre"], email, row["date"], ANOMALY_HOURS_LABEL])

    for col_cells in ws.columns:
        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = max(12, min(40, max_len + 2))

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_summary_export(master_df: pd.DataFrame, year: int, month: int) -> bytes:
    """Righe = dipendenti (ID, nome), colonne = giorni del mese.

    Cella = A / Z / AZ, colorata di verde se eligible quel giorno, rosso altrimenti.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Riepilogo"

    n_days = calendar.monthrange(year, month)[1]
    days = [date(year, month, d) for d in range(1, n_days + 1)]

    ws.append(["ID Dipendente", "Nome"] + [d.strftime("%d/%m (%a)") for d in days])
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")

    employees = (
        master_df[["employee_id", "nombre"]]
        .drop_duplicates(subset=["employee_id"])
        .sort_values("nombre")
        .to_dict("records")
    )

    lookup = {
        (row["employee_id"], row["date"]): row
        for _, row in master_df.iterrows()
    }

    for emp in employees:
        emp_id = emp["employee_id"]
        line = [emp_id, emp["nombre"]]
        ws.append(line)
        excel_row = ws.max_row
        for col_idx, d in enumerate(days, start=3):
            row = lookup.get((emp_id, d))
            if row is None:
                continue
            if row["amati"] and row["zippi"]:
                value = "AZ"
            elif row["amati"]:
                value = "A"
            elif row["zippi"]:
                value = "Z"
            else:
                continue
            cell = ws.cell(row=excel_row, column=col_idx, value=value)
            cell.fill = GREEN_FILL if row["eligible"] else RED_FILL
            cell.alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 28
    for col_idx in range(3, 3 + len(days)):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = 10
    ws.freeze_panes = "C2"

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
