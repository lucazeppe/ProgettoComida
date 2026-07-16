# Controllo Pasti Aziendali — Amati / Zippi

App Streamlit per il controllo mensile dei buoni pasto aziendali, con incrocio tra gli ordini di due fornitori (Amati e Zippi) e le ore lavorate dei dipendenti. Nessun database o persistenza: ogni sessione parte da zero caricando i file del mese.

## Regole di business

- Un dipendente ha diritto al pasto a **23,46** se ha lavorato almeno **4 ore** quel giorno; altrimenti il pasto costa **165**.
- In caso di doppia prenotazione (ordine sia Amati che Zippi lo stesso giorno) con diritto al pasto: un ordine viene addebitato a 23,46 e l'altro a 165 (mai 23,46 due volte). Senza diritto: entrambi a 165.
- Il recap per fornitore (usato per validare le fatture) conta solo il numero di pasti forniti, senza importi: una doppia prenotazione conta 1 pasto per Amati **e** 1 per Zippi.
- Le anomalie rilevate sono due: **doppia prenotazione** e **ordine in giorno con ore insufficienti**. Sono forzabili/correggibili a video (rimozione di un ordine, abbono del vincolo ore) prima di generare gli export.

## Requisiti

```
pip install -r requirements.txt
```

## Avvio

```
streamlit run app.py
```
oppure, se `streamlit` non è nel PATH:
```
python3 -m streamlit run app.py
```

## File di input richiesti

1. **Ordini Zippi** (.xlsx) — una riga per dipendente, con le date di ordine elencate nelle colonne successive a `Correo`.
2. **Ordini Amati** (.xlsx) — un foglio per settimana, colonne `Lunes..Viernes` con 0/1. Le date non sono esplicite: vengono dedotte automaticamente in base al mese scelto (un foglio = una settimana consecutiva).
3. **Ore lavorate** (.xlsx) — una riga per dipendente con **ID formato 2000xxx**, colonne datetime (una per giorno) con le ore lavorate.

La chiave di match tra i file è l'**Employee ID** in formato `2000xxx`.

**Coerenza col mese scelto**: sui file ordini (Zippi e Amati) eventuali giorni fuori dal mese selezionato bloccano il processamento con un errore descrittivo — sugli ordini la correttezza del mese è essenziale. Sul file ore invece eventuali colonne di altri mesi (es. l'export sconfina nel mese successivo, come tipicamente accade) vengono **ignorate silenziosamente**: servono solo come supporto per leggere le ore dei giorni del mese scelto.

## Output

- **Export solleciti anomalie**: Excel con ID dipendente, nome, email, giorno e ragione dell'anomalia — pensato come base per un flusso Power Automate di sollecito via email.
- **Export riepilogo colorato**: Excel con dipendenti in riga e giorni del mese in colonna; cella `A`/`Z`/`AZ` colorata di verde (diritto al pasto) o rosso (nessun diritto).

## Struttura del codice

- `parsers.py` — lettura e validazione dei 3 file di input.
- `business.py` — unione dati, calcolo costi/anomalie, applicazione delle forzature.
- `exports.py` — generazione dei due file Excel di export.
- `app.py` — interfaccia Streamlit.
