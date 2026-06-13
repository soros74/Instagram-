#!/usr/bin/env python3
"""
Google Maps Location History Analyzer
======================================
Parses records.json (Google Takeout) e produce file Excel organizzati per:
  output/<anno>/<anno>-<mm>_<Mese>.xlsx
  ogni file ha un foglio per settimana (gg 1-7, 8-14, 15-21, 22-28, 29-31)

Utilizzo:
  python analyze_location_history.py [records.json] [--output-dir DIR]

Requisiti:
  pip install openpyxl requests
"""

import json
import sys
import math
import time
import datetime
import logging
import argparse
from pathlib import Path
from collections import defaultdict

try:
    from zoneinfo import ZoneInfo
except ImportError:
    from backports.zoneinfo import ZoneInfo  # pip install backports.zoneinfo

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Configurazione ─────────────────────────────────────────────────────────────

ITALY_TZ = ZoneInfo("Europe/Rome")

# Soglie accuratezza (metri): punti PEGGIORI di questi vengono scartati
ACCURACY_LIMITS: dict[str, int] = {
    "GPS":     50,
    "WIFI":   200,
    "NETWORK": 200,
    "CELL":   150,   # celle telefoniche → soglia più severa
}

# Rilevamento soste
VISIT_RADIUS_M  = 150  # raggio max (m) per considerare lo stesso luogo
VISIT_MIN_MIN   = 5    # durata minima (min) per contare come visita

# Geocoding
GEOCODE_DELAY   = 1.2  # secondi tra le richieste a Nominatim
OVERPASS_DELAY  = 2.0  # secondi tra le richieste a Overpass
MAX_ALTERNATIVES = 3   # numero massimo di luoghi alternativi da mostrare

# File di cache geocoding su disco (evita di rifetchiare)
GEOCODE_CACHE_FILE = Path(".geocode_cache.json")
OVERPASS_CACHE_FILE = Path(".overpass_cache.json")

_geocode_cache: dict[str, dict]  = {}
_overpass_cache: dict[str, list] = {}

# ── Festività italiane ─────────────────────────────────────────────────────────

def _easter(year: int) -> datetime.date:
    """Calcola la data di Pasqua (algoritmo anonimo gregoriano)."""
    a = year % 19
    b, c = divmod(year, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    ll = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * ll) // 451
    month = (h + ll - 7 * m + 114) // 31
    day   = (h + ll - 7 * m + 114) % 31 + 1
    return datetime.date(year, month, day)

_holiday_cache: dict[int, set[datetime.date]] = {}

def italian_holidays(year: int) -> set[datetime.date]:
    if year in _holiday_cache:
        return _holiday_cache[year]
    fixed = [
        (1, 1),  (1, 6),  (4, 25), (5, 1),  (6, 2),
        (8, 15), (11, 1), (12, 8), (12, 25), (12, 26),
    ]
    h = {datetime.date(year, m, d) for m, d in fixed}
    easter = _easter(year)
    h.add(easter)
    h.add(easter + datetime.timedelta(days=1))  # Pasquetta
    _holiday_cache[year] = h
    return h

def is_working_day(dt: datetime.datetime) -> bool:
    """True solo se il giorno è un giorno lavorativo (no weekend, no festività)."""
    d = dt.date()
    if d.weekday() >= 5:  # sabato=5, domenica=6
        return False
    return d not in italian_holidays(d.year)

# ── Affidabilità posizione ─────────────────────────────────────────────────────

def is_reliable(loc: dict) -> bool:
    source   = loc.get("source", "")
    accuracy = loc.get("accuracy", 99999)
    limit    = ACCURACY_LIMITS.get(source, 0)
    return limit > 0 and accuracy <= limit

# ── Distanza Haversine ─────────────────────────────────────────────────────────

def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Distanza in metri tra due punti lat/lon."""
    R = 6_371_000
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

# ── Cache su disco ─────────────────────────────────────────────────────────────

def _load_cache(path: Path, store: dict) -> None:
    if path.exists():
        try:
            store.update(json.loads(path.read_text("utf-8")))
        except Exception:
            pass

def _save_cache(path: Path, store: dict) -> None:
    try:
        path.write_text(json.dumps(store, ensure_ascii=False, indent=2), "utf-8")
    except Exception:
        pass

# ── Geocoding Nominatim ────────────────────────────────────────────────────────

_NOMINATIM_HEADERS = {
    "User-Agent": "LocationHistoryAnalyzer/2.0 (personal-analysis)",
    "Accept-Language": "it,en;q=0.8",
}

def _geo_key(lat: float, lon: float) -> str:
    return f"{lat:.5f},{lon:.5f}"

def reverse_geocode(lat: float, lon: float) -> dict:
    """Restituisce {'name': str, 'address': str} per le coordinate date."""
    key = _geo_key(lat, lon)
    if key in _geocode_cache:
        return _geocode_cache[key]

    time.sleep(GEOCODE_DELAY)
    result = {"name": "", "address": ""}
    try:
        r = requests.get(
            "https://nominatim.openstreetmap.org/reverse",
            params={"lat": lat, "lon": lon, "format": "jsonv2",
                    "addressdetails": 1, "namedetails": 1, "zoom": 18},
            headers=_NOMINATIM_HEADERS,
            timeout=12,
        )
        if r.status_code == 200:
            data = r.json()
            if not isinstance(data, dict):
                raise ValueError(f"Risposta non valida: {data!r}")
            name = (data.get("name")
                    or (data.get("namedetails") or {}).get("name", "")
                    or "")
            addr = data.get("display_name", "")
            # Pulisce l'indirizzo: prende i primi 4 segmenti (no nazione)
            parts = [p.strip() for p in addr.split(",")]
            addr_clean = ", ".join(parts[:4]) if len(parts) >= 4 else addr
            result = {"name": name, "address": addr_clean}
    except Exception as exc:
        logging.warning("Geocoding fallito per %s,%s: %s", lat, lon, exc)

    _geocode_cache[key] = result
    _save_cache(GEOCODE_CACHE_FILE, _geocode_cache)
    return result

# ── Luoghi alternativi (Overpass API) ─────────────────────────────────────────

def _overpass_key(lat: float, lon: float, radius: float) -> str:
    return f"{lat:.4f},{lon:.4f},{radius:.0f}"

def nearby_places(lat: float, lon: float, radius_m: float) -> list[dict]:
    """
    Restituisce fino a MAX_ALTERNATIVES luoghi nominati nel raggio dato,
    ordinati per distanza, con probabilità calcolata per distanza inversa.
    Ogni elemento: {'name': str, 'address': str, 'distance_m': float, 'probability': float}
    """
    radius_m = max(radius_m, 30)  # minimo 30m
    key = _overpass_key(lat, lon, radius_m)
    if key in _overpass_cache:
        return _overpass_cache[key]

    time.sleep(OVERPASS_DELAY)
    places: list[dict] = []

    query = (
        f"[out:json][timeout:15];"
        f"("
        f"node(around:{radius_m:.0f},{lat},{lon})[name];"
        f"way(around:{radius_m:.0f},{lat},{lon})[name];"
        f");"
        f"out center {MAX_ALTERNATIVES * 4};"
    )

    try:
        r = requests.post(
            "https://overpass-api.de/api/interpreter",
            data=query,
            timeout=20,
        )
        if r.status_code == 200:
            for el in r.json().get("elements", []):
                tags = el.get("tags", {})
                name = tags.get("name", "")
                if not name:
                    continue
                if el["type"] == "node":
                    elat, elon = el["lat"], el["lon"]
                elif el["type"] == "way":
                    c = el.get("center", {})
                    elat, elon = c.get("lat", lat), c.get("lon", lon)
                else:
                    continue

                dist = haversine(lat, lon, elat, elon)
                # Indirizzo dal tag OSM se disponibile
                addr_parts = []
                if tags.get("addr:street"):
                    num = tags.get("addr:housenumber", "")
                    addr_parts.append(
                        tags["addr:street"] + (" " + num if num else "")
                    )
                if tags.get("addr:city"):
                    addr_parts.append(tags["addr:city"])

                places.append({
                    "name": name,
                    "address": ", ".join(addr_parts),
                    "distance_m": dist,
                })

        # Ordina per distanza, prendi i più vicini
        places.sort(key=lambda x: x["distance_m"])
        places = places[:MAX_ALTERNATIVES]

        # Probabilità = punteggio inverso alla distanza
        if places:
            scores = [1.0 / (p["distance_m"] + 1) for p in places]
            total  = sum(scores)
            for p, s in zip(places, scores):
                p["probability"] = round((s / total) * 100, 1)
        else:
            # nessun luogo trovato → il geocode principale è l'unico candidato
            pass

    except Exception as exc:
        logging.warning("Overpass fallito per %s,%s: %s", lat, lon, exc)

    _overpass_cache[key] = places
    _save_cache(OVERPASS_CACHE_FILE, _overpass_cache)
    return places

# ── Parsing punti ──────────────────────────────────────────────────────────────

def parse_point(loc: dict) -> dict | None:
    """
    Converte un record raw in un punto pulito.
    Restituisce None se il punto non è affidabile o cade in un giorno non lavorativo.
    """
    if not is_reliable(loc):
        return None

    try:
        ts_utc = datetime.datetime.fromisoformat(
            loc["timestamp"].replace("Z", "+00:00")
        )
    except (KeyError, ValueError):
        return None

    ts_local = ts_utc.astimezone(ITALY_TZ)
    if not is_working_day(ts_local):
        return None

    activities: list[dict] = []
    for grp in loc.get("activity", []):
        for act in grp.get("activity", []):
            t = act.get("type", "UNKNOWN")
            c = act.get("confidence", 0)
            activities.append({"type": t, "confidence": c})

    return {
        "lat":        loc["latitudeE7"]  / 1e7,
        "lon":        loc["longitudeE7"] / 1e7,
        "accuracy":   loc.get("accuracy", 9999),
        "source":     loc.get("source", ""),
        "timestamp":  ts_local,
        "activities": activities,
    }

# ── Rilevamento soste (stay-point detection) ───────────────────────────────────

def detect_visits(points: list[dict]) -> list[dict]:
    """
    Algoritmo di Li et al. per rilevare le soste (stay points).
    Restituisce una lista di visite con centroide, orari e attività aggregate.
    """
    visits = []
    n = len(points)
    i = 0

    while i < n:
        j = i + 1
        while j < n:
            dist = haversine(
                points[i]["lat"], points[i]["lon"],
                points[j]["lat"], points[j]["lon"],
            )
            if dist > VISIT_RADIUS_M:
                break
            j += 1

        cluster = points[i:j]
        duration_min = (
            cluster[-1]["timestamp"] - cluster[0]["timestamp"]
        ).total_seconds() / 60.0

        if duration_min >= VISIT_MIN_MIN:
            centroid_lat = sum(p["lat"] for p in cluster) / len(cluster)
            centroid_lon = sum(p["lon"] for p in cluster) / len(cluster)
            avg_accuracy = sum(p["accuracy"] for p in cluster) / len(cluster)

            all_acts: list[dict] = []
            for p in cluster:
                all_acts.extend(p["activities"])

            visits.append({
                "arrival":       cluster[0]["timestamp"],
                "departure":     cluster[-1]["timestamp"],
                "duration_min":  duration_min,
                "lat":           centroid_lat,
                "lon":           centroid_lon,
                "accuracy_m":    avg_accuracy,
                "activities":    all_acts,
                "n_points":      len(cluster),
                # salviamo indice del primo punto in arrivo per il percorso
                "point_index_start": i,
            })
            i = j
        else:
            i += 1

    return visits

# ── Analisi percorso ───────────────────────────────────────────────────────────

_ACT_LABELS: dict[str, str] = {
    "IN_VEHICLE": "In veicolo",
    "ON_FOOT":    "A piedi",
    "WALKING":    "A piedi",
    "RUNNING":    "Correndo",
    "ON_BICYCLE": "In bicicletta",
    "STILL":      "Fermo",
    "TILTING":    "In movimento",
    "UNKNOWN":    "Sconosciuto",
}

def _aggregate_activities(activities: list[dict]) -> list[tuple[str, float]]:
    """Aggrega i confidence per tipo, normalizza, restituisce lista (tipo, %) desc."""
    totals: dict[str, float] = defaultdict(float)
    for act in activities:
        totals[act["type"]] += act["confidence"]
    grand = sum(totals.values())
    if grand == 0:
        return []
    return sorted(
        ((t, (s / grand) * 100) for t, s in totals.items()),
        key=lambda x: -x[1],
    )

def route_description(transit_activities: list[dict]) -> str:
    """
    Descrive il percorso verso la destinazione filtrando le attività banali.
    Esempio: "In veicolo (73%) → A piedi (21%)"
    """
    relevant = [
        a for a in transit_activities
        if a["type"] not in ("STILL", "UNKNOWN", "TILTING")
    ]
    ranked = _aggregate_activities(relevant)
    if not ranked:
        return "N/D"
    parts = []
    for typ, pct in ranked[:3]:
        if pct >= 8:
            parts.append(f"{_ACT_LABELS.get(typ, typ)} ({pct:.0f}%)")
    return " → ".join(parts) if parts else "N/D"

def activities_description(activities: list[dict]) -> str:
    """
    Descrive le attività rilevate durante la sosta (per la colonna 'Alternative').
    Usato per indicare lo stato fisico nell'ora di permanenza.
    """
    ranked = _aggregate_activities(activities)
    parts = []
    for typ, pct in ranked[:4]:
        if pct >= 5:
            parts.append(f"{_ACT_LABELS.get(typ, typ)}: {pct:.0f}%")
    return "\n".join(parts)

# ── Settimana nel mese ─────────────────────────────────────────────────────────

def week_of_month(d: datetime.date) -> int:
    """Settimana 1-5 all'interno del mese (blocchi di 7 giorni: 1-7, 8-14, …)."""
    return (d.day - 1) // 7 + 1

def week_date_range(year: int, month: int, week: int) -> tuple[datetime.date, datetime.date]:
    """Primo e ultimo giorno del blocco settimanale nel mese."""
    import calendar
    last_day = calendar.monthrange(year, month)[1]
    start = datetime.date(year, month, (week - 1) * 7 + 1)
    end   = datetime.date(year, month, min(week * 7, last_day))
    return start, end

# ── Formattazione ──────────────────────────────────────────────────────────────

_MONTH_IT = [
    "", "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
    "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre",
]
_DAY_IT = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"]

def fmt_duration(minutes: float) -> str:
    h = int(minutes // 60)
    m = int(minutes % 60)
    return f"{h}h {m:02d}m" if h else f"{m}m"

def fmt_alternatives(places: list[dict]) -> str:
    """Formatta luoghi alternativi: 'Nome - Indirizzo: XX%'"""
    if not places:
        return ""
    lines = []
    for p in places:
        name = p.get("name", "")
        addr = p.get("address", "")
        pct  = p.get("probability", 0)
        loc_str = f"{name} — {addr}" if addr else name
        lines.append(f"{loc_str}: {pct:.0f}%")
    return "\n".join(lines)

# ── Excel styling ──────────────────────────────────────────────────────────────

_C_HEADER_BG  = "1F4E79"
_C_TITLE_BG   = "2E75B6"
_C_ALT_ROW_BG = "D9E2F3"
_C_WHITE       = "FFFFFF"

_THIN = Side(style="thin", color="AAAAAA")
_CELL_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_COLUMNS = [
    ("Giorno",              14),
    ("Nome Luogo",          28),
    ("Indirizzo",           40),
    ("Ora Arrivo",          11),
    ("Ora Partenza",        12),
    ("Durata",              10),
    ("Come Arrivato",       32),
    ("Attività Rilevata",   28),
    ("Luoghi Alternativi\n(con probabilità)", 40),
]

def _header_cell(ws, row: int, col: int, value: str) -> None:
    cell = ws.cell(row=row, column=col, value=value)
    cell.font      = Font(color=_C_WHITE, bold=True, size=10)
    cell.fill      = PatternFill("solid", fgColor=_C_HEADER_BG)
    cell.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
    cell.border    = _CELL_BORDER

def _write_week_sheet(
    ws,
    week_num: int,
    week_start: datetime.date,
    week_end: datetime.date,
    rows: list[dict],
) -> None:
    ws.title = f"Sett.{week_num} {week_start.strftime('%d.%m')}-{week_end.strftime('%d.%m')}"

    ncols = len(_COLUMNS)

    # ── Riga titolo ──────────────────────────────────────────────────────────
    ws.merge_cells(f"A1:{get_column_letter(ncols)}1")
    tc = ws["A1"]
    tc.value     = (f"Settimana {week_num}  ·  "
                    f"{week_start.day} {_MONTH_IT[week_start.month]} {week_start.year}"
                    f" – {week_end.day} {_MONTH_IT[week_end.month]} {week_end.year}")
    tc.font      = Font(bold=True, size=12, color=_C_WHITE)
    tc.fill      = PatternFill("solid", fgColor=_C_TITLE_BG)
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # ── Intestazioni colonne ─────────────────────────────────────────────────
    for ci, (header, width) in enumerate(_COLUMNS, start=1):
        _header_cell(ws, 2, ci, header)
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[2].height = 38
    ws.freeze_panes = "A3"

    # ── Righe dati ───────────────────────────────────────────────────────────
    for ri, row in enumerate(rows, start=3):
        alt_lines = row.get("alternatives", "").count("\n") + 1
        act_lines = row.get("activity_desc", "").count("\n") + 1
        route_lines = row.get("route", "").count("\n") + 1
        row_height = max(30, 15 * max(alt_lines, act_lines, route_lines))
        ws.row_dimensions[ri].height = row_height

        use_fill = (ri % 2 == 0)
        fill = PatternFill("solid", fgColor=_C_ALT_ROW_BG) if use_fill else None

        values = [
            row.get("day_label", ""),
            row.get("name", ""),
            row.get("address", ""),
            row.get("arrival_str", ""),
            row.get("departure_str", ""),
            row.get("duration_str", ""),
            row.get("route", ""),
            row.get("activity_desc", ""),
            row.get("alternatives", ""),
        ]
        for ci, val in enumerate(values, start=1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.border    = _CELL_BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if fill:
                cell.fill = fill

# ── Main ───────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Analizza records.json Google Maps → Excel organizzato per anno/mese/settimana"
    )
    parser.add_argument(
        "input",
        nargs="?",
        default="records.json",
        help="Percorso del file records.json (default: records.json)",
    )
    parser.add_argument(
        "--output-dir",
        default="location_history_excel",
        help="Cartella di output (default: location_history_excel)",
    )
    parser.add_argument(
        "--no-alternatives",
        action="store_true",
        help="Salta la ricerca di luoghi alternativi (più veloce)",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING"],
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S",
    )
    log = logging.getLogger("lha")

    output_dir = Path(args.output_dir)
    input_path = Path(args.input)

    if not input_path.exists():
        log.error("File non trovato: %s", input_path)
        sys.exit(1)

    # Carica cache geocoding da disco
    _load_cache(GEOCODE_CACHE_FILE, _geocode_cache)
    _load_cache(OVERPASS_CACHE_FILE, _overpass_cache)
    log.info("Cache geocoding: %d voci  |  Cache Overpass: %d voci",
             len(_geocode_cache), len(_overpass_cache))

    # ── Carica e filtra ───────────────────────────────────────────────────────
    log.info("Caricamento %s ...", input_path)
    with open(input_path, "r", encoding="utf-8") as f:
        raw_data = json.load(f)
    raw_locs = raw_data.get("locations", [])
    log.info("Punti grezzi: %d", len(raw_locs))

    log.info("Filtraggio punti (sorgente inaffidabile, weekend, festività) ...")
    points: list[dict] = []
    skipped_unreliable = 0
    skipped_nonworking = 0

    for loc in raw_locs:
        if not is_reliable(loc):
            skipped_unreliable += 1
            continue
        p = parse_point(loc)
        if p is None:
            skipped_nonworking += 1
            continue
        points.append(p)

    log.info(
        "Rimasti: %d punti  (scartati per inaffidabilità: %d  |  weekend/festivi: %d)",
        len(points), skipped_unreliable, skipped_nonworking,
    )

    if not points:
        log.error("Nessun punto affidabile nei giorni lavorativi. Controllare il file.")
        sys.exit(1)

    points.sort(key=lambda p: p["timestamp"])

    # ── Rilevamento visite ────────────────────────────────────────────────────
    log.info("Rilevamento soste (stay-point detection) ...")
    visits_raw = detect_visits(points)
    log.info("Soste trovate: %d", len(visits_raw))

    if not visits_raw:
        log.error("Nessuna sosta rilevata. Prova ad aumentare VISIT_RADIUS_M o VISIT_MIN_MIN.")
        sys.exit(1)

    # ── Arricchimento (geocoding + percorso) ──────────────────────────────────
    log.info("Geocoding e analisi percorsi (potrebbe richiedere tempo) ...")
    visits_enriched: list[dict] = []

    for idx, visit in enumerate(visits_raw):
        lat, lon = visit["lat"], visit["lon"]

        # Geocode posizione principale
        geo = reverse_geocode(lat, lon)

        # Luoghi alternativi nel raggio di incertezza
        if args.no_alternatives:
            alt_places: list[dict] = []
        else:
            alt_places = nearby_places(lat, lon, visit["accuracy_m"])

        # Percorso: attività tra la visita precedente e questa
        if idx > 0:
            prev_dep = visits_raw[idx - 1]["departure"]
            transit_pts = [
                p for p in points
                if prev_dep <= p["timestamp"] <= visit["arrival"]
            ]
            transit_acts: list[dict] = []
            for tp in transit_pts:
                transit_acts.extend(tp["activities"])
        else:
            transit_acts = []

        dt = visit["arrival"]
        day_label = f"{_DAY_IT[dt.weekday()]} {dt.strftime('%d/%m/%Y')}"

        visits_enriched.append({
            "arrival":       visit["arrival"],
            "departure":     visit["departure"],
            "arrival_str":   visit["arrival"].strftime("%H:%M"),
            "departure_str": visit["departure"].strftime("%H:%M"),
            "duration_str":  fmt_duration(visit["duration_min"]),
            "day_label":     day_label,
            "name":          geo.get("name", ""),
            "address":       geo.get("address", ""),
            "route":         route_description(transit_acts),
            "activity_desc": activities_description(visit["activities"]),
            "alternatives":  fmt_alternatives(alt_places),
            "year":          dt.year,
            "month":         dt.month,
            "week":          week_of_month(dt.date()),
        })

        if (idx + 1) % 25 == 0 or (idx + 1) == len(visits_raw):
            log.info("  → geocodificati %d / %d", idx + 1, len(visits_raw))

    # ── Raggruppa per anno → mese → settimana ─────────────────────────────────
    tree: dict = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
    for v in visits_enriched:
        tree[v["year"]][v["month"]][v["week"]].append(v)

    # ── Scrittura Excel ───────────────────────────────────────────────────────
    output_dir.mkdir(parents=True, exist_ok=True)
    files_written = 0

    for year, months in sorted(tree.items()):
        year_dir = output_dir / str(year)
        year_dir.mkdir(exist_ok=True)

        for month, weeks in sorted(months.items()):
            month_name = _MONTH_IT[month]
            filename = year_dir / f"{year}-{month:02d}_{month_name}.xlsx"

            wb = Workbook()
            wb.remove(wb.active)  # rimuove il foglio vuoto predefinito

            for week_num, week_visits in sorted(weeks.items()):
                week_start, week_end = week_date_range(year, month, week_num)
                ws = wb.create_sheet()
                _write_week_sheet(ws, week_num, week_start, week_end, week_visits)

            wb.save(filename)
            log.info("Salvato: %s  (%d visite)", filename,
                     sum(len(w) for w in weeks.values()))
            files_written += 1

    log.info("Completato! %d file Excel in: %s", files_written, output_dir)


if __name__ == "__main__":
    main()
