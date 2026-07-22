# Controllo Pasti Aziendali — Amati / Zippi

App Streamlit per il controllo mensile dei buoni pasto aziendali, con incrocio tra gli ordini di due fornitori (Amati e Zippi) e le ore lavorate dei dipendenti. Nessun database o persistenza: ogni sessione parte da zero caricando i file del mese. Interfaccia disponibile in italiano (default) e spagnolo, con toggle a bandiera in alto nella sidebar.

## Regole di business

- Un dipendente ha diritto al pasto a **23,46** se ha lavorato più di **4 ore** (soglia stretta: 4 ore esatte NON danno diritto) quel giorno; altrimenti il pasto costa **165** (Amati) o **160** (Zippi) — le due tariffe "senza diritto" sono diverse per fornitore.
- In caso di doppia prenotazione (ordine sia Amati che Zippi lo stesso giorno) con diritto al pasto: un ordine viene addebitato a 23,46 e l'altro a 165 (mai 23,46 due volte). Senza diritto: 165 + 160 (325 totali).
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

1. **Ordini Zippi** (.xlsx) — una riga per dipendente, con le date di ordine elencate nelle colonne successive a `Correo`. La mail in questa colonna non viene più usata (vedi punto 4).
2. **Ordini Amati** (.xlsx) — un foglio per settimana, colonne `Lunes..Viernes` con 0/1. Le date non sono esplicite: vengono dedotte automaticamente in base al mese scelto. I fogli si identificano dal **numero di settimana ISO contenuto nel nome del foglio** (es. `"23"`), non dall'ordine delle schede nel file: devono esistere tutti i fogli delle settimane del mese scelto (fogli extra vengono ignorati), ciascuno con tutte e 5 le colonne giorno e almeno una riga dipendente, altrimenti il caricamento viene bloccato con un errore descrittivo.
3. **Ore lavorate** (.xlsx) — una riga per dipendente con **ID formato 2000xxx**, colonne datetime (una per giorno) con le ore lavorate.
4. **Anagrafica dipendenti** (.xlsx) — export HR con una riga per dipendente; si usano solo le colonne `User/Employee ID` ed `Business Email Information Email Address` (tutte le altre, es. reparto/manager/date assunzione, vengono ignorate). È l'unica fonte dell'email usata nell'app e negli export: se un dipendente non compare in questo file, o vi compare con il campo email vuoto, l'email resta vuota (nessun errore, ma un avviso dedicato e richiudibile a video elenca questi dipendenti).

La chiave di match tra i file è l'**Employee ID** in formato `2000xxx`.

**Coerenza col mese scelto**: sui file ordini (Zippi e Amati) eventuali giorni fuori dal mese selezionato bloccano il processamento con un errore descrittivo — sugli ordini la correttezza del mese è essenziale. Sul file ore invece eventuali colonne di altri mesi (es. l'export sconfina nel mese successivo, come tipicamente accade) vengono **ignorate silenziosamente**: servono solo come supporto per leggere le ore dei giorni del mese scelto. Un dipendente con ordini ma assente dal file ore non blocca il processamento (trattato come 0 ore quel giorno) ma genera un avviso dedicato, richiudibile, a video.

## Funzionalità principali

- **Vista dipendenti**: tabella con ordini/anomalie per giorno, filtri (solo anomalie, ricerca per ID/nome, intervallo date), e forzature manuali (rimuovi ordine Amati/Zippi, abbona vincolo ore) applicabili prima di generare gli export.
- **Abbono automatico praticanti**: toggle che abbona di default l'anomalia "ore insufficienti" per i dipendenti riconosciuti come praticanti in base al dominio email (mai la doppia prenotazione), senza renderlo un vincolo permanente se poi disattivato.
- **Vista fornitori**: conteggio pasti per Amati/Zippi per giorno, per validare le fatture (nessun importo).

## Output

Tutti gli export Excel sono **sempre in inglese**, indipendentemente dalla lingua scelta nell'interfaccia — così eventuali automazioni a valle (es. Power Automate) non dipendono dal toggle di lingua.

- **Scarica anomalie**: un record per ogni singola anomalia (ID dipendente, nome, email, giorno, ragione) — un giorno con entrambe le anomalie genera 2 record.
- **Scarica riepilogo**: dipendenti in riga, giorni del mese in colonna; cella `A`/`Z`/`AZ` colorata di verde (diritto al pasto) o rosso (nessun diritto), con colonne di totale per dipendente e due righe di totale giornaliero per fornitore in fondo. La colonna `Note` segnala i dipendenti le cui anomalie ore sono dovute a un record assente nel file ore (non a ore insufficienti realmente registrate), da verificare a parte.
- **Scarica sollecito mail**: un record per dipendente con almeno un'anomalia (non uno per anomalia), con le date di ciascun tipo di anomalia elencate come lista testuale `GG/MM/AAAA, GG/MM/AAAA` in due colonne separate (ore insufficienti / doppia prenotazione). Foglio "Emails" formattato come vera **Tabella Excel** (non un semplice range), richiesto dal flusso Power Automate di invio solleciti. Colonne aggiuntive per quel flusso: `RecordID` (intero progressivo da 1, valore fisso non formula), `Object` (fisso `"Comida"`), `Allegato`/`NameAllegato`/`ProcessingState`/`Att1..Att30` (vuote, strutturali per il template).

## Struttura del codice

- `parsers.py` — lettura e validazione dei 4 file di input.
- `business.py` — unione dati, calcolo costi/anomalie, applicazione delle forzature.
- `exports.py` — generazione dei tre file Excel di export (sempre in inglese).
- `i18n.py` — traduzioni IT/ES dell'interfaccia e dei messaggi di errore/warning dei parser.
- `app.py` — interfaccia Streamlit.
