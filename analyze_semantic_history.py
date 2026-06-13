#!/usr/bin/env python3
"""
Google Maps Semantic Location History Analyzer
==============================================
Legge la cartella "Semantic Location History" del Google Takeout e produce
file Excel organizzati per anno/mese/settimana.

I file Semantic contengono già nome luogo, indirizzo, percorso con distanza
e mezzo di trasporto: NON richiede geocoding né connessione internet.

Utilizzo:
  python analyze_semantic_history.py [CARTELLA] [--output-dir DIR]

Esempio:
  python analyze_semantic_history.py \
    "/storage/emulated/0/Download/takeout/Semantic Location History" \
    --output-dir "/storage/emulated/0/foto a69/records Claude"
"""

import json
import sys
import datetime
import argparse
import logging
from pathlib import Path
from collections import defaultdict

try:
    from zoneinfo import ZoneInfo
except ImportError:
    from backports.zoneinfo import ZoneInfo

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Configurazione ─────────────────────────────────────────────────────────────

ITALY_TZ = ZoneInfo("Europe/Rome")

# Nomi mesi inglesi → numero (Google usa nomi inglesi nei filename)
_MONTH_EN_TO_NUM: dict[str, int] = {
    "january": 1,  "february": 2,  "march": 3,    "april": 4,
    "may": 5,      "june": 6,      "july": 7,     "august": 8,
    "september": 9,"october": 10,  "november": 11, "december": 12,
    # anche italiano per sicurezza
    "gennaio": 1,  "febbraio": 2,  "marzo": 3,    "aprile": 4,
    "maggio": 5,   "giugno": 6,    "luglio": 7,   "agosto": 8,
    "settembre": 9,"ottobre": 10,  "novembre": 11, "dicembre": 12,
}

_MONTH_IT = [
    "", "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
    "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre",
]
_DAY_IT = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"]

# ── Festività italiane ─────────────────────────────────────────────────────────

def _easter(year: int) -> datetime.date:
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
        (1, 1), (1, 6), (4, 25), (5, 1), (6, 2),
        (8, 15), (11, 1), (12, 8), (12, 25), (12, 26),
    ]
    h = {datetime.date(year, m, d) for m, d in fixed}
    easter = _easter(year)
    h.add(easter)
    h.add(easter + datetime.timedelta(days=1))
    _holiday_cache[year] = h
    return h

def is_working_day(dt: datetime.datetime) -> bool:
    d = dt.date()
    if d.weekday() >= 5:
        return False
    return d not in italian_holidays(d.year)

# ── Parsing timestamp ──────────────────────────────────────────────────────────

def parse_ts(s: str) -> datetime.datetime:
    """Converte timestamp ISO 8601 in datetime locale Italia."""
    return datetime.datetime.fromisoformat(
        s.replace("Z", "+00:00")
    ).astimezone(ITALY_TZ)

# ── Etichette attività ─────────────────────────────────────────────────────────

_ACT_LABELS: dict[str, str] = {
    "IN_PASSENGER_VEHICLE": "In auto",
    "IN_VEHICLE":           "In veicolo",
    "WALKING":              "A piedi",
    "CYCLING":              "In bicicletta",
    "IN_BUS":               "In autobus",
    "IN_TRAIN":             "In treno",
    "IN_SUBWAY":            "In metro",
    "IN_TRAM":              "In tram",
    "IN_FERRY":             "In traghetto",
    "FLYING":               "In aereo",
    "RUNNING":              "Correndo",
    "MOTORCYCLING":         "In moto",
    "SKIING":               "Sci",
    "SAILING":              "In barca",
    "IN_TAXI":              "In taxi",
    "STILL":                "Fermo",
    "UNKNOWN_ACTIVITY_TYPE":"Sconosciuto",
}

_CONF_LABELS: dict[str, str] = {
    "HIGH_CONFIDENCE":   "Alta",
    "MEDIUM_CONFIDENCE": "Media",
    "LOW_CONFIDENCE":    "Bassa",
    "USER_CONFIRMED":    "Confermata",
    "INFERRED":          "Dedotta",
}

def act_label(code: str) -> str:
    return _ACT_LABELS.get(code, code.replace("_", " ").title())

def conf_label(code: str) -> str:
    return _CONF_LABELS.get(code, code)

# ── Formattazione ──────────────────────────────────────────────────────────────

def fmt_duration(seconds: float) -> str:
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    return f"{h}h {m:02d}m" if h else f"{m}m"

def fmt_distance(meters: int | None) -> str:
    if meters is None:
        return ""
    if meters >= 1000:
        return f"{meters / 1000:.1f} km"
    return f"{meters} m"

def clean_address(addr: str) -> str:
    """Rimuove 'Italia' finale e codici postali ridondanti."""
    if not addr:
        return ""
    parts = [p.strip() for p in addr.split(",")]
    # Rimuove "Italia" in fondo
    if parts and parts[-1].strip().lower() in ("italia", "italy"):
        parts = parts[:-1]
    return ", ".join(parts)

def fmt_alternatives(candidates: list[dict]) -> str:
    """
    Formatta i candidati alternativi di Google.
    Ogni candidato: {'name': str, 'address': str, 'locationConfidence': float}
    """
    if not candidates:
        return ""
    lines = []
    for c in candidates:
        name = c.get("name", "")
        addr = clean_address(c.get("address", ""))
        conf = c.get("locationConfidence", 0)
        loc  = f"{name} — {addr}" if addr else name
        lines.append(f"{loc}: {conf:.0f}%")
    return "\n".join(lines)

def fmt_route(seg: dict | None) -> str:
    """
    Descrive il percorso da un activitySegment.
    Formato: "In auto (85%) / A piedi (10%)  ·  3,2 km  ·  18m"
    """
    if not seg:
        return "N/D"

    # Modalità principale
    main_type = seg.get("activityType", "")
    main_label = act_label(main_type) if main_type else "Sconosciuto"

    # Alternative con probabilità
    activities = seg.get("activities", [])
    mode_parts = []
    for act in sorted(activities, key=lambda x: -x.get("probability", 0))[:3]:
        atype = act.get("activityType", "")
        prob  = act.get("probability", 0)
        if prob >= 5:
            mode_parts.append(f"{act_label(atype)} ({prob:.0f}%)")

    modes_str = " / ".join(mode_parts) if mode_parts else main_label

    # Distanza e durata
    dist = fmt_distance(seg.get("distance"))
    start_ts = seg.get("_start_ts")
    end_ts   = seg.get("_end_ts")
    dur_str  = ""
    if start_ts and end_ts:
        dur_str = fmt_duration((end_ts - start_ts).total_seconds())

    extras = "  ·  ".join(filter(None, [dist, dur_str]))
    return f"{modes_str}  ·  {extras}" if extras else modes_str

# ── Scanning dei file ──────────────────────────────────────────────────────────

def scan_files(root: Path) -> list[tuple[int, int, Path]]:
    """
    Scansiona la cartella Semantic Location History e restituisce
    lista di (anno, mese, path) ordinata cronologicamente.
    """
    results = []
    for json_file in sorted(root.rglob("*.json")):
        # Deduce anno dalla cartella parent (es. .../2015/...)
        year = None
        for part in json_file.parts:
            if part.isdigit() and 2000 <= int(part) <= 2100:
                year = int(part)
                break

        # Deduce mese dal nome file
        stem = json_file.stem.lower()
        month = None

        # Formato "2015_march" o "2015_MARCH"
        for sep in ("_", "-", " "):
            if sep in stem:
                parts = stem.split(sep)
                for part in parts:
                    if part.isdigit() and 2000 <= int(part) <= 2100:
                        year = int(part)
                    elif part in _MONTH_EN_TO_NUM:
                        month = _MONTH_EN_TO_NUM[part]

        # Formato solo "march"
        if month is None and stem in _MONTH_EN_TO_NUM:
            month = _MONTH_EN_TO_NUM[stem]

        if year and month:
            results.append((year, month, json_file))
        else:
            logging.warning("File ignorato (anno/mese non riconosciuti): %s", json_file)

    results.sort()
    return results

# ── Parsing di un file mensile ─────────────────────────────────────────────────

def parse_month_file(path: Path) -> list[dict]:
    """
    Legge un file JSON mensile e restituisce lista di timelineObjects già
    tipizzati come 'placeVisit' o 'activitySegment'.
    """
    try:
        data = json.loads(path.read_text("utf-8"))
    except Exception as e:
        logging.error("Errore lettura %s: %s", path, e)
        return []
    return data.get("timelineObjects", [])

# ── Estrazione visite ──────────────────────────────────────────────────────────

def extract_visits(
    objects: list[dict],
    date_from: datetime.date | None,
    date_to:   datetime.date | None,
) -> list[dict]:
    """
    Estrae le visite (placeVisit) accoppiando ogni visita con il segmento
    di percorso (activitySegment) che la precede immediatamente.
    Applica i filtri: intervallo date, weekend, festività italiane.
    """
    visits  = []
    last_seg: dict | None = None  # activitySegment precedente

    for obj in objects:
        if "activitySegment" in obj:
            seg = obj["activitySegment"]
            dur = seg.get("duration", {})
            try:
                seg["_start_ts"] = parse_ts(dur["startTimestamp"])
                seg["_end_ts"]   = parse_ts(dur["endTimestamp"])
            except (KeyError, ValueError):
                pass
            last_seg = seg
            continue

        if "placeVisit" not in obj:
            continue

        pv  = obj["placeVisit"]
        dur = pv.get("duration", {})

        try:
            arrival   = parse_ts(dur["startTimestamp"])
            departure = parse_ts(dur["endTimestamp"])
        except (KeyError, ValueError):
            last_seg = None
            continue

        # Filtro intervallo date
        d = arrival.date()
        if date_from and d < date_from:
            last_seg = None
            continue
        if date_to and d > date_to:
            last_seg = None
            continue

        # Filtro weekend e festività
        if not is_working_day(arrival):
            last_seg = None
            continue

        loc  = pv.get("location", {})
        name = loc.get("name", "")
        addr = clean_address(loc.get("address", ""))
        loc_conf   = loc.get("locationConfidence", 0)
        place_conf = pv.get("placeConfidence", "")
        visit_conf = pv.get("visitConfidence", 0)

        # Candidati alternativi (già forniti da Google)
        other = pv.get("otherCandidateLocations", [])

        # Durata sosta
        duration_sec = (departure - arrival).total_seconds()

        day_label = f"{_DAY_IT[arrival.weekday()]} {arrival.strftime('%d/%m/%Y')}"

        visits.append({
            "arrival":      arrival,
            "departure":    departure,
            "arrival_str":  arrival.strftime("%H:%M"),
            "departure_str":departure.strftime("%H:%M"),
            "duration_str": fmt_duration(duration_sec),
            "day_label":    day_label,
            "name":         name,
            "address":      addr,
            "confidence":   f"{conf_label(place_conf)} ({loc_conf:.0f}%)" if place_conf else f"{loc_conf:.0f}%",
            "route":        fmt_route(last_seg),
            "alternatives": fmt_alternatives(other),
            "year":         arrival.year,
            "month":        arrival.month,
            "week":         (arrival.day - 1) // 7 + 1,
        })

        last_seg = None  # consumato

    return visits

# ── Settimana nel mese ─────────────────────────────────────────────────────────

def week_date_range(year: int, month: int, week: int):
    import calendar
    last_day = calendar.monthrange(year, month)[1]
    start = datetime.date(year, month, (week - 1) * 7 + 1)
    end   = datetime.date(year, month, min(week * 7, last_day))
    return start, end

# ── Excel styling ──────────────────────────────────────────────────────────────

_THIN = Side(style="thin", color="AAAAAA")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_COLUMNS = [
    ("Giorno",              14),
    ("Nome Luogo",          28),
    ("Indirizzo",           40),
    ("Ora Arrivo",          11),
    ("Ora Partenza",        12),
    ("Durata",              10),
    ("Come Arrivato\n(mezzo · distanza · tempo)", 42),
    ("Confidenza Google",   18),
    ("Luoghi Alternativi\n(con probabilità)", 40),
]

def _write_week_sheet(ws, week_num, week_start, week_end, rows):
    ncols = len(_COLUMNS)
    ws.title = f"Sett.{week_num} {week_start.strftime('%d.%m')}-{week_end.strftime('%d.%m')}"

    # Titolo
    ws.merge_cells(f"A1:{get_column_letter(ncols)}1")
    tc = ws["A1"]
    tc.value     = (f"Settimana {week_num}  ·  "
                    f"{week_start.day} {_MONTH_IT[week_start.month]} {week_start.year}"
                    f" – {week_end.day} {_MONTH_IT[week_end.month]} {week_end.year}")
    tc.font      = Font(bold=True, size=12, color="FFFFFF")
    tc.fill      = PatternFill("solid", fgColor="2E75B6")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Intestazioni
    for ci, (header, width) in enumerate(_COLUMNS, 1):
        cell = ws.cell(row=2, column=ci, value=header)
        cell.font      = Font(color="FFFFFF", bold=True, size=10)
        cell.fill      = PatternFill("solid", fgColor="1F4E79")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = _BORDER
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[2].height = 38
    ws.freeze_panes = "A3"

    # Dati
    for ri, row in enumerate(rows, 3):
        alt_lines   = row.get("alternatives", "").count("\n") + 1
        route_lines = row.get("route", "").count("\n") + 1
        ws.row_dimensions[ri].height = max(30, 15 * max(alt_lines, route_lines))

        fill = PatternFill("solid", fgColor="D9E2F3") if ri % 2 == 0 else None

        values = [
            row.get("day_label", ""),
            row.get("name", ""),
            row.get("address", ""),
            row.get("arrival_str", ""),
            row.get("departure_str", ""),
            row.get("duration_str", ""),
            row.get("route", ""),
            row.get("confidence", ""),
            row.get("alternatives", ""),
        ]
        for ci, val in enumerate(values, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.border    = _BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if fill:
                cell.fill = fill

# ── Richiesta interattiva date ─────────────────────────────────────────────────

def _ask_date(prompt: str) -> datetime.date | None:
    while True:
        raw = input(f"{prompt} [lascia vuoto = nessun limite]: ").strip()
        if not raw:
            return None
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return datetime.datetime.strptime(raw, fmt).date()
            except ValueError:
                continue
        print("  ✗ Formato non valido. Usa GG/MM/AAAA (es. 01/03/2019)")

# ── Main ───────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Analizza Semantic Location History Google → Excel per anno/mese/settimana"
    )
    parser.add_argument(
        "input_dir",
        nargs="?",
        default=".",
        help="Cartella 'Semantic Location History' (default: cartella corrente)",
    )
    parser.add_argument("--output-dir", default="semantic_history_excel")
    parser.add_argument("--date-from", default=None, help="Data inizio GG/MM/AAAA")
    parser.add_argument("--date-to",   default=None, help="Data fine GG/MM/AAAA")
    parser.add_argument("--log-level", default="INFO",
                        choices=["DEBUG", "INFO", "WARNING"])
    args = parser.parse_args()

    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S",
    )
    log = logging.getLogger("sha")

    root = Path(args.input_dir)
    if not root.exists():
        log.error("Cartella non trovata: %s", root)
        sys.exit(1)

    # ── Selezione date ────────────────────────────────────────────────────────
    print()
    print("═" * 54)
    print("  ANALISI SEMANTIC LOCATION HISTORY — GOOGLE MAPS")
    print("═" * 54)

    def parse_date_arg(s):
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return datetime.datetime.strptime(s, fmt).date()
            except ValueError:
                continue
        log.error("Data non valida: %s", s)
        sys.exit(1)

    if args.date_from:
        date_from = parse_date_arg(args.date_from)
    else:
        print("\n  Specifica l'intervallo di date da estrarre.")
        print("  Premi INVIO per non impostare un limite.\n")
        date_from = _ask_date("  Data INIZIO (GG/MM/AAAA)")

    if args.date_to:
        date_to = parse_date_arg(args.date_to)
    else:
        date_to = _ask_date("  Data FINE   (GG/MM/AAAA)")

    from_str = date_from.strftime("%d/%m/%Y") if date_from else "inizio archivio"
    to_str   = date_to.strftime("%d/%m/%Y")   if date_to   else "fine archivio"
    print(f"\n  Estrazione: {from_str}  →  {to_str}")
    print("═" * 54)
    print()

    # ── Scansione file ────────────────────────────────────────────────────────
    files = scan_files(root)
    if not files:
        log.error("Nessun file JSON trovato in: %s", root)
        sys.exit(1)

    # Filtra file fuori dall'intervallo (ottimizzazione: salta i mesi inutili)
    if date_from:
        files = [(y, m, p) for y, m, p in files
                 if datetime.date(y, m, 1) >= date_from.replace(day=1)]
    if date_to:
        import calendar
        files = [(y, m, p) for y, m, p in files
                 if datetime.date(y, m, 1) <= date_to.replace(
                     day=calendar.monthrange(date_to.year, date_to.month)[1])]

    log.info("File da elaborare: %d", len(files))

    # ── Elaborazione ──────────────────────────────────────────────────────────
    all_visits: list[dict] = []
    for year, month, path in files:
        log.info("Leggo: %s", path.name)
        objects = parse_month_file(path)
        visits  = extract_visits(objects, date_from, date_to)
        log.info("  → %d visite trovate (giorni lavorativi)", len(visits))
        all_visits.extend(visits)

    log.info("Totale visite: %d", len(all_visits))
    if not all_visits:
        log.error("Nessuna visita trovata nell'intervallo specificato.")
        sys.exit(1)

    # ── Raggruppa per anno → mese → settimana ─────────────────────────────────
    tree: dict = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
    for v in all_visits:
        tree[v["year"]][v["month"]][v["week"]].append(v)

    # ── Scrittura Excel ───────────────────────────────────────────────────────
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    files_written = 0

    for year, months in sorted(tree.items()):
        year_dir = output_dir / str(year)
        year_dir.mkdir(exist_ok=True)

        for month, weeks in sorted(months.items()):
            month_name = _MONTH_IT[month]
            filename   = year_dir / f"{year}-{month:02d}_{month_name}.xlsx"

            wb = Workbook()
            wb.remove(wb.active)

            for week_num, week_visits in sorted(weeks.items()):
                ws_start, ws_end = week_date_range(year, month, week_num)
                ws = wb.create_sheet()
                _write_week_sheet(ws, week_num, ws_start, ws_end, week_visits)

            wb.save(filename)
            n = sum(len(w) for w in weeks.values())
            log.info("Salvato: %s  (%d visite)", filename.name, n)
            files_written += 1

    log.info("Completato! %d file Excel in: %s", files_written, output_dir)


if __name__ == "__main__":
    main()
