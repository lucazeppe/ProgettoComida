"""Logica di business: unione dati, calcolo costi/anomalie, applicazione eccezioni forzate."""
from __future__ import annotations

from datetime import date

import pandas as pd

HOURS_THRESHOLD = 6
COST_LOW = 23.46
COST_HIGH = 165.0

ANOMALY_DOUBLE_LABEL = "Doppia prenotazione"
ANOMALY_HOURS_LABEL = "Ore ufficio insufficienti"

MASTER_COLUMNS = [
    "employee_id", "nombre", "email", "date",
    "amati", "zippi", "hours", "eligible",
    "anomaly_double", "anomaly_hours",
    "cost_amati", "cost_zippi",
]


def _build_employee_registry(zippi_df: pd.DataFrame, amati_df: pd.DataFrame, hours_df: pd.DataFrame) -> pd.DataFrame:
    records = {}

    for _, r in amati_df.iterrows():
        full_name = f"{r['nombre']} {r['apellidos']}".strip()
        records.setdefault(r["employee_id"], {"nombre": full_name, "email": r["email"]})

    for _, r in zippi_df.iterrows():
        entry = records.setdefault(r["employee_id"], {"nombre": r["nombre"], "email": r["email"]})
        if not entry.get("email"):
            entry["email"] = r["email"]
        if not entry.get("nombre"):
            entry["nombre"] = r["nombre"]

    for emp_id in hours_df["employee_id"].unique():
        records.setdefault(emp_id, {"nombre": None, "email": None})

    reg = pd.DataFrame(
        [{"employee_id": k, "nombre": v["nombre"], "email": v["email"]} for k, v in records.items()],
        columns=["employee_id", "nombre", "email"],
    )
    # Evita che i valori mancanti diventino NaN float (celle Excel non valide negli export):
    # meglio una stringa vuota, sia per i DataFrame vuoti che per join successivi.
    reg["nombre"] = reg["nombre"].fillna("")
    reg["email"] = reg["email"].fillna("")
    return reg


def _is_weekday(dates: pd.Series) -> pd.Series:
    """Maschera booleana Lun-Ven, con dtype bool esplicito anche su Series vuote
    (necessario perché Series([]).apply(...) restituisce dtype object, e
    l'indicizzazione booleana con dtype non-bool su un DataFrame vuoto ne
    azzera anche le colonne)."""
    return pd.Series([d.weekday() < 5 for d in dates], index=dates.index, dtype=bool)


def build_master(
    zippi_df: pd.DataFrame,
    amati_df: pd.DataFrame,
    hours_df: pd.DataFrame,
    year: int,
    month: int,
    exclude_weekends: bool = True,
) -> pd.DataFrame:
    """Costruisce il DataFrame long-format: una riga per (employee_id, date) con almeno un ordine.

    Se exclude_weekends è True, righe con data di sabato/domenica non entrano nel
    calcolo di costi/anomalie (i giorni non lavorativi non danno diritto al pasto).
    """
    registry = _build_employee_registry(zippi_df, amati_df, hours_df)

    if exclude_weekends:
        zippi_df = zippi_df[_is_weekday(zippi_df["date"])]
        amati_df = amati_df[_is_weekday(amati_df["date"])]
        hours_df = hours_df[_is_weekday(hours_df["date"])]

    zippi_orders = zippi_df[["employee_id", "date"]].drop_duplicates().assign(zippi=True)
    amati_orders = amati_df[["employee_id", "date"]].drop_duplicates().assign(amati=True)

    orders = pd.merge(
        zippi_orders, amati_orders, on=["employee_id", "date"], how="outer"
    )
    if orders.empty:
        orders = pd.DataFrame(columns=["employee_id", "date", "zippi", "amati"])
    orders["zippi"] = orders["zippi"].fillna(False)
    orders["amati"] = orders["amati"].fillna(False)

    orders = orders.merge(hours_df, on=["employee_id", "date"], how="left")
    orders["hours"] = orders["hours"].fillna(0.0)

    orders = orders.merge(registry, on="employee_id", how="left")

    orders = _compute_costs_anomalies(orders)

    return orders[MASTER_COLUMNS]


def _compute_costs_anomalies(df: pd.DataFrame, force_eligible: pd.Series | None = None) -> pd.DataFrame:
    df = df.copy()
    df["eligible"] = df["hours"] >= HOURS_THRESHOLD
    if force_eligible is not None:
        df["eligible"] = df["eligible"] | force_eligible.reindex(df.index, fill_value=False)
    df["anomaly_double"] = df["amati"] & df["zippi"]
    df["anomaly_hours"] = ~df["eligible"] & (df["amati"] | df["zippi"])

    cost_amati = []
    cost_zippi = []
    for _, row in df.iterrows():
        ca, cz = None, None
        if row["amati"] and row["zippi"]:
            if row["eligible"]:
                ca, cz = COST_LOW, COST_HIGH  # Amati assorbe la tariffa scontata (ordine alfabetico)
            else:
                ca, cz = COST_HIGH, COST_HIGH
        elif row["amati"]:
            ca = COST_LOW if row["eligible"] else COST_HIGH
        elif row["zippi"]:
            cz = COST_LOW if row["eligible"] else COST_HIGH
        cost_amati.append(ca)
        cost_zippi.append(cz)
    df["cost_amati"] = cost_amati
    df["cost_zippi"] = cost_zippi
    return df


def apply_overrides(
    master_df: pd.DataFrame,
    overrides: dict,
    drop_fully_cancelled: bool = True,
) -> pd.DataFrame:
    """Applica le eccezioni forzate dall'utente e ricalcola costi/anomalie.

    overrides: {(employee_id, date): {"remove_amati": bool, "remove_zippi": bool, "waive_hours": bool}}

    "waive_hours" forza l'idoneità al pasto senza alterare le ore realmente
    registrate (il valore "hours" mostrato resta quello vero).

    Se drop_fully_cancelled è True (default, usato per export/recap), le righe
    rimaste senza alcun ordine dopo le forzature vengono escluse. Se False
    (usato per la vista editabile), restano visibili così l'utente può
    correggere/annullare la forzatura anche dopo aver rimosso entrambi gli
    ordini di un giorno.
    """
    df = master_df.copy()
    if not overrides:
        return _compute_costs_anomalies(df)

    force_eligible = pd.Series(False, index=df.index)
    for idx, row in df.iterrows():
        ov = overrides.get((row["employee_id"], row["date"]))
        if not ov:
            continue
        if ov.get("remove_amati"):
            df.at[idx, "amati"] = False
        if ov.get("remove_zippi"):
            df.at[idx, "zippi"] = False
        if ov.get("waive_hours"):
            force_eligible.at[idx] = True

    if drop_fully_cancelled:
        keep_mask = df["amati"] | df["zippi"]
        df = df[keep_mask].reset_index(drop=True)
        force_eligible = force_eligible[keep_mask].reset_index(drop=True)

    return _compute_costs_anomalies(df, force_eligible=force_eligible)


def check_weekend_activity(master_df: pd.DataFrame) -> list[str]:
    weekend = master_df[~_is_weekday(master_df["date"])]
    if weekend.empty:
        return []
    return [f"Trovati {len(weekend)} ordini registrati nel weekend (sabato/domenica), inattesi."]


def supplier_recap(master_df: pd.DataFrame) -> pd.DataFrame:
    """Conteggio pasti per fornitore/giorno, per validare le fatture (nessun importo)."""
    rows = []
    for d, group in master_df.groupby("date"):
        rows.append({
            "date": d,
            "Amati": int(group["amati"].sum()),
            "Zippi": int(group["zippi"].sum()),
        })
    recap = pd.DataFrame(rows, columns=["date", "Amati", "Zippi"]).sort_values("date")
    total_row = pd.DataFrame([{
        "date": "TOTALE",
        "Amati": recap["Amati"].sum() if not recap.empty else 0,
        "Zippi": recap["Zippi"].sum() if not recap.empty else 0,
    }])
    return pd.concat([recap, total_row], ignore_index=True)
