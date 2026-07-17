"""Logica di business: unione dati, calcolo costi/anomalie, applicazione eccezioni forzate."""
from __future__ import annotations

from datetime import date

import pandas as pd

from parsers import title_case

# Soglia di ore stretta: diritto al pasto solo se le ore lavorate sono
# STRETTAMENTE superiori a HOURS_THRESHOLD (> non >=).
HOURS_THRESHOLD = 4
COST_LOW = 23.46
COST_HIGH = 165.0

# Le etichette anomalia a video seguono la lingua scelta (vedi i18n.py,
# chiavi "anomaly_double"/"anomaly_hours"). Gli export Excel invece usano
# sempre inglese (lingua neutra), indipendentemente dalla lingua
# dell'interfaccia — così eventuali automazioni a valle (es. Power Automate)
# non dipendono dalla lingua scelta a video.
ANOMALY_DOUBLE_LABEL_EXPORT = "Double Booking"
ANOMALY_HOURS_LABEL_EXPORT = "Insufficient Office Hours"

INTERN_EMAIL_SUFFIX = ".intern@e80group.com"


def is_intern(email) -> bool:
    return isinstance(email, str) and email.strip().lower().endswith(INTERN_EMAIL_SUFFIX)

MASTER_COLUMNS = [
    "employee_id", "nombre", "email", "date",
    "amati", "zippi", "hours", "eligible",
    "anomaly_double", "anomaly_hours",
    "cost_amati", "cost_zippi",
]


def _build_employee_registry(zippi_df: pd.DataFrame, amati_df: pd.DataFrame, hours_df: pd.DataFrame) -> pd.DataFrame:
    records = {}

    for _, r in amati_df.iterrows():
        full_name = title_case(f"{r['nombre']} {r['apellidos']}".strip())
        records.setdefault(r["employee_id"], {"nombre": full_name, "email": r["email"]})

    for _, r in zippi_df.iterrows():
        entry = records.setdefault(r["employee_id"], {"nombre": title_case(r["nombre"]), "email": r["email"]})
        if not entry.get("email"):
            entry["email"] = r["email"]
        if not entry.get("nombre"):
            entry["nombre"] = title_case(r["nombre"])

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

    # Un dipendente con ordine ma assente dal file ore non genera più errore:
    # viene trattato come 0 ore quel giorno (quindi anomalia "ore insufficienti"
    # su tutti i suoi ordini), esattamente come un giorno con ore=0 registrate.
    zippi_orders = zippi_df[["employee_id", "date"]].drop_duplicates().assign(zippi=True)
    amati_orders = amati_df[["employee_id", "date"]].drop_duplicates().assign(amati=True)

    orders = pd.merge(
        zippi_orders, amati_orders, on=["employee_id", "date"], how="outer"
    )
    if orders.empty:
        orders = pd.DataFrame(columns=["employee_id", "date", "zippi", "amati"])
    orders["zippi"] = orders["zippi"].fillna(False)
    orders["amati"] = orders["amati"].fillna(False)

    orders = orders.merge(hours_df[["employee_id", "date", "hours"]], on=["employee_id", "date"], how="left")
    orders["hours"] = orders["hours"].fillna(0.0)

    # Nome/Cognome dal file ore: attributo per dipendente (non per singolo
    # giorno), così resta lo stesso su tutte le righe anche se un giorno
    # specifico non ha un corrispettivo nel file ore.
    hours_names = (
        hours_df[["employee_id", "nombre_ore"]]
        .dropna(subset=["nombre_ore"])
        .drop_duplicates(subset=["employee_id"])
    )
    orders = orders.merge(hours_names, on="employee_id", how="left")
    orders["nombre_ore"] = orders["nombre_ore"].fillna("")

    orders = orders.merge(registry, on="employee_id", how="left")

    # Nome dipendente uniforme ovunque nell'app: il file ore è la fonte più
    # affidabile (Nome/Cognome dedicati), quindi ha priorità; solo se assente
    # (dipendente non presente nel file ore) si usa il nome da Amati/Zippi.
    orders["nombre"] = orders["nombre_ore"].where(orders["nombre_ore"].astype(bool), orders["nombre"])
    orders = orders.drop(columns=["nombre_ore"])

    orders = _compute_costs_anomalies(orders)

    return orders[MASTER_COLUMNS]


def _compute_costs_anomalies(df: pd.DataFrame, force_eligible: pd.Series | None = None) -> pd.DataFrame:
    df = df.copy()
    df["eligible"] = df["hours"] > HOURS_THRESHOLD
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


def supplier_recap(master_df: pd.DataFrame) -> pd.DataFrame:
    """Conteggio pasti per fornitore/giorno, per validare le fatture (nessun importo).

    Solo il dettaglio giornaliero: i totali del mese si calcolano a parte
    (es. sommando le colonne Amati/Zippi) per essere mostrati come KPI.
    """
    rows = []
    for d, group in master_df.groupby("date"):
        rows.append({
            "date": d,
            "Amati": int(group["amati"].sum()),
            "Zippi": int(group["zippi"].sum()),
        })
    return pd.DataFrame(rows, columns=["date", "Amati", "Zippi"]).sort_values("date").reset_index(drop=True)
