# Google Maps — Luoghi Visitati

App web per visualizzare la cronologia posizioni di Google Maps.

## Come ottenere i dati

1. Vai su [Google Takeout](https://takeout.google.com)
2. Seleziona solo **"Mappe"** → **"Cronologia posizioni"**
3. Scarica l'archivio ed estrai i file JSON dalla cartella:
   - `Semantic Location History/YYYY/YYYY_MESE.json` (consigliato)
   - oppure `Records.json` (dati GPS grezzi)

## Avvio

```bash
cd maps_viewer
pip install -r requirements.txt
python app.py
```

Apri il browser su `http://localhost:5000`

## Funzionalità

- Caricamento multiplo di file JSON (drag & drop o selezione)
- Tabella con: **nome luogo**, **indirizzo**, **data**, **orario**, **tempo di permanenza**
- Filtri per testo e intervallo di date
- Paginazione (50 record per pagina)
- Statistiche riassuntive (totale visite, giorni, permanenza media/massima)
- Vista **mappa interattiva** (OpenStreetMap + Leaflet)
- Clic su un luogo → dettaglio completo + link Google Maps
- Supporta sia i file mensili (`timelineObjects`) che `Records.json`
