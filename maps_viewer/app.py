import json
import os
from datetime import datetime, timezone
from math import radians, sin, cos, sqrt, atan2
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


def haversine_m(lat1, lng1, lat2, lng2):
    """Distance in metres between two lat/lng points."""
    R = 6_371_000
    p1, p2 = radians(lat1), radians(lat2)
    dp = radians(lat2 - lat1)
    dl = radians(lng2 - lng1)
    a = sin(dp / 2) ** 2 + cos(p1) * cos(p2) * sin(dl / 2) ** 2
    return R * 2 * atan2(sqrt(a), sqrt(1 - a))


def detect_stays(points, radius_m=150, min_minutes=5):
    """
    Cluster raw GPS points into stays (places where the user stopped).
    Uses a centroid-expansion approach: O(n), works on millions of records.
    """
    stays = []
    if not points:
        return stays

    c_lat = points[0]["lat"]
    c_lng = points[0]["lng"]
    c_count = 1
    c_start = 0

    for i in range(1, len(points)):
        d = haversine_m(c_lat, c_lng, points[i]["lat"], points[i]["lng"])
        if d <= radius_m:
            # Expand cluster with running average
            c_lat = (c_lat * c_count + points[i]["lat"]) / (c_count + 1)
            c_lng = (c_lng * c_count + points[i]["lng"]) / (c_count + 1)
            c_count += 1
        else:
            dur = (points[i - 1]["dt"] - points[c_start]["dt"]).total_seconds() / 60
            if dur >= min_minutes and c_count >= 2:
                stays.append({
                    "lat": round(c_lat, 6),
                    "lng": round(c_lng, 6),
                    "start": points[c_start]["dt"],
                    "end": points[i - 1]["dt"],
                    "dur_min": int(dur),
                })
            c_lat, c_lng = points[i]["lat"], points[i]["lng"]
            c_count = 1
            c_start = i

    # Last cluster
    if c_count >= 2:
        dur = (points[-1]["dt"] - points[c_start]["dt"]).total_seconds() / 60
        if dur >= min_minutes:
            stays.append({
                "lat": round(c_lat, 6),
                "lng": round(c_lng, 6),
                "start": points[c_start]["dt"],
                "end": points[-1]["dt"],
                "dur_min": int(dur),
            })
    return stays


def parse_records_json(data):
    """
    Parse Records.json — extracts meaningful stays via stay detection.
    Returns one entry per detected stop (not one per raw GPS point).
    """
    raw = data.get("locations", [])

    # Build normalised list sorted by time
    points = []
    for loc in raw:
        ts = loc.get("timestamp") or ""
        if ts:
            dt = parse_timestamp(ts)
        else:
            ts_ms = loc.get("timestampMs")
            if not ts_ms:
                continue
            dt = datetime.fromtimestamp(int(ts_ms) / 1000, tz=timezone.utc)
        if not dt:
            continue
        lat = loc.get("latitudeE7", 0) / 1e7
        lng = loc.get("longitudeE7", 0) / 1e7
        if lat == 0 and lng == 0:
            continue
        points.append({"lat": lat, "lng": lng, "dt": dt})

    points.sort(key=lambda p: p["dt"])

    stays = detect_stays(points)

    visits = []
    for s in stays:
        local_start = s["start"].astimezone()
        h, m = divmod(s["dur_min"], 60)
        dur_str = f"{h}h {m}m" if h else f"{m}m"
        visits.append({
            "name": f"Sosta GPS ({s['lat']:.4f}, {s['lng']:.4f})",
            "address": f"{s['lat']:.5f}, {s['lng']:.5f}",
            "date": local_start.strftime("%d/%m/%Y"),
            "time": local_start.strftime("%H:%M"),
            "date_sort": local_start.strftime("%Y%m%d%H%M"),
            "duration": dur_str,
            "duration_minutes": s["dur_min"],
            "lat": s["lat"],
            "lng": s["lng"],
            "confidence": 0,
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
