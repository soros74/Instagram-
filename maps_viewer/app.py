import json
import os
from datetime import datetime, timezone
from flask import Flask, render_template, request, jsonify

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500 MB


def parse_timestamp(ts):
    """Parse ISO 8601 timestamp to datetime object."""
    if not ts:
        return None
    # Handle both formats: "2023-01-15T10:30:00Z" and "2023-01-15T10:30:00.000Z"
    ts = ts.replace('Z', '+00:00')
    try:
        return datetime.fromisoformat(ts)
    except ValueError:
        return None


def format_duration(start_dt, end_dt):
    """Return human-readable duration string."""
    if not start_dt or not end_dt:
        return "—"
    delta = end_dt - start_dt
    total = int(delta.total_seconds())
    if total < 0:
        return "—"
    hours, remainder = divmod(total, 3600)
    minutes = remainder // 60
    if hours > 0:
        return f"{hours}h {minutes}m"
    return f"{minutes}m"


def parse_timeline_json(data):
    """Parse Semantic Location History monthly JSON files."""
    visits = []
    objects = data.get("timelineObjects", [])
    for obj in objects:
        pv = obj.get("placeVisit")
        if not pv:
            continue

        location = pv.get("location", {})
        duration = pv.get("duration", {})

        name = location.get("name", "")
        address = location.get("address", "")
        lat = location.get("latitudeE7", 0) / 1e7
        lng = location.get("longitudeE7", 0) / 1e7

        start_dt = parse_timestamp(duration.get("startTimestamp"))
        end_dt = parse_timestamp(duration.get("endTimestamp"))

        if not start_dt:
            continue

        local_start = start_dt.astimezone()
        visits.append({
            "name": name or "Luogo sconosciuto",
            "address": address or "—",
            "date": local_start.strftime("%d/%m/%Y"),
            "time": local_start.strftime("%H:%M"),
            "date_sort": local_start.strftime("%Y%m%d%H%M"),
            "duration": format_duration(start_dt, end_dt),
            "duration_minutes": int((end_dt - start_dt).total_seconds() // 60) if end_dt else 0,
            "lat": round(lat, 6),
            "lng": round(lng, 6),
            "confidence": pv.get("visitConfidence", 0),
        })
    return visits


def parse_records_json(data):
    """Parse Records.json (raw location history) — groups nearby points into visits."""
    visits = []
    locations = data.get("locations", [])

    # Sort by timestamp
    locations.sort(key=lambda x: x.get("timestampMs", x.get("timestamp", "0")))

    for loc in locations:
        ts = loc.get("timestamp") or ""
        if not ts:
            ts_ms = int(loc.get("timestampMs", 0))
            dt = datetime.fromtimestamp(ts_ms / 1000, tz=timezone.utc)
        else:
            dt = parse_timestamp(ts)
        if not dt:
            continue

        lat = loc.get("latitudeE7", 0) / 1e7
        lng = loc.get("longitudeE7", 0) / 1e7
        local_dt = dt.astimezone()

        visits.append({
            "name": "Posizione GPS",
            "address": f"{round(lat,5)}, {round(lng,5)}",
            "date": local_dt.strftime("%d/%m/%Y"),
            "time": local_dt.strftime("%H:%M"),
            "date_sort": local_dt.strftime("%Y%m%d%H%M"),
            "duration": "—",
            "duration_minutes": 0,
            "lat": round(lat, 6),
            "lng": round(lng, 6),
            "confidence": loc.get("accuracy", 0),
        })

    return visits


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "Nessun file caricato"}), 400

    all_visits = []
    errors = []

    for f in files:
        if not f.filename.endswith(".json"):
            errors.append(f"{f.filename}: non è un file JSON")
            continue
        try:
            data = json.load(f)
        except json.JSONDecodeError as e:
            errors.append(f"{f.filename}: JSON non valido ({e})")
            continue

        if "timelineObjects" in data:
            visits = parse_timeline_json(data)
        elif "locations" in data:
            visits = parse_records_json(data)
        else:
            errors.append(f"{f.filename}: formato non riconosciuto")
            continue

        all_visits.extend(visits)

    # Sort by date descending
    all_visits.sort(key=lambda x: x["date_sort"], reverse=True)

    return jsonify({
        "visits": all_visits,
        "total": len(all_visits),
        "errors": errors,
    })


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
