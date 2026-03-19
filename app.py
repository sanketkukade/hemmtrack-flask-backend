"""
HemmTrack Pro — Flask Backend
Deploy on Railway. Provides /get_stats endpoint for dashboard.
"""

from flask import Flask, jsonify, request
from flask_cors import CORS
from datetime import datetime
import os

app = Flask(__name__)

# Allow your GitHub Pages frontend + localhost dev
CORS(app, origins=[
    "https://sanketkukade.github.io",
    "http://127.0.0.1:5500",
    "http://localhost:5500",
    "http://localhost:3000",
])

# ──────────────────────────────────────────────
# In-memory store (replace with DB later)
# ──────────────────────────────────────────────
defect_records = [
    {"station": "ST-100", "defect": "Open Hem",  "shift": "A", "time": "06:30", "status": "High",   "inspector": "Rajesh",  "date": "2026-03-20"},
    {"station": "ST-120", "defect": "Dent",       "shift": "A", "time": "07:15", "status": "Medium", "inspector": "Suresh",  "date": "2026-03-20"},
    {"station": "ST-150", "defect": "Open Hem",  "shift": "A", "time": "07:45", "status": "High",   "inspector": "Amit",    "date": "2026-03-20"},
    {"station": "ST-200", "defect": "Scratch",    "shift": "B", "time": "14:10", "status": "Low",    "inspector": "Priya",   "date": "2026-03-20"},
    {"station": "ST-100", "defect": "Open Hem",  "shift": "B", "time": "14:30", "status": "High",   "inspector": "Rajesh",  "date": "2026-03-20"},
    {"station": "ST-220", "defect": "Burr",       "shift": "B", "time": "15:00", "status": "Medium", "inspector": "Nilesh",  "date": "2026-03-20"},
    {"station": "ST-150", "defect": "Scratch",    "shift": "C", "time": "22:10", "status": "Low",    "inspector": "Pooja",   "date": "2026-03-20"},
    {"station": "ST-100", "defect": "Open Hem",  "shift": "C", "time": "22:45", "status": "High",   "inspector": "Amit",    "date": "2026-03-20"},
    {"station": "ST-200", "defect": "Dent",       "shift": "A", "time": "08:20", "status": "Medium", "inspector": "Suresh",  "date": "2026-03-20"},
    {"station": "ST-120", "defect": "Open Hem",  "shift": "A", "time": "09:00", "status": "High",   "inspector": "Rajesh",  "date": "2026-03-20"},
    {"station": "ST-220", "defect": "Gap Issue",  "shift": "B", "time": "16:30", "status": "High",   "inspector": "Priya",   "date": "2026-03-20"},
    {"station": "ST-150", "defect": "Open Hem",  "shift": "C", "time": "23:15", "status": "High",   "inspector": "Nilesh",  "date": "2026-03-20"},
]


def compute_stats(records):
    """Aggregate defect records into dashboard stats."""
    total = len(records)
    high = sum(1 for r in records if r["status"] == "High")

    # Unique stations
    stations = list(set(r["station"] for r in records))

    # Alerts = High + Medium
    alerts = sum(1 for r in records if r["status"] in ("High", "Medium"))

    # By station
    by_station = {}
    for r in records:
        by_station[r["station"]] = by_station.get(r["station"], 0) + 1

    # By defect type
    defect_types = {}
    for r in records:
        defect_types[r["defect"]] = defect_types.get(r["defect"], 0) + 1

    # By shift
    by_shift = {}
    for r in records:
        by_shift[r["shift"]] = by_shift.get(r["shift"], 0) + 1

    # Recent 10 (newest first by time)
    recent = sorted(records, key=lambda x: x["time"], reverse=True)[:10]

    return {
        "total_defects": total,
        "high_defects": high,
        "stations": len(stations),
        "alerts": alerts,
        "by_station": by_station,
        "defect_types": defect_types,
        "by_shift": by_shift,
        "recent": recent,
        "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


# ──────────────────────────────────────────────
# ROUTES
# ──────────────────────────────────────────────

@app.route("/", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "service": "HemmTrack Pro Backend",
        "version": "2.0",
        "timestamp": datetime.now().isoformat(),
    })


@app.route("/get_stats", methods=["GET"])
def get_stats():
    """Main dashboard endpoint — returns aggregated defect stats."""
    try:
        stats = compute_stats(defect_records)
        return jsonify(stats), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/add_defect", methods=["POST"])
def add_defect():
    """Add a new defect record (for future form integration)."""
    try:
        data = request.get_json()
        required = ["station", "defect", "shift", "status"]
        for field in required:
            if field not in data:
                return jsonify({"error": f"Missing field: {field}"}), 400

        record = {
            "station":   data["station"],
            "defect":    data["defect"],
            "shift":     data["shift"],
            "time":      data.get("time", datetime.now().strftime("%H:%M")),
            "status":    data["status"],
            "inspector": data.get("inspector", "—"),
            "date":      data.get("date", datetime.now().strftime("%Y-%m-%d")),
        }
        defect_records.append(record)
        return jsonify({"message": "Defect added", "total": len(defect_records)}), 201
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ──────────────────────────────────────────────
# RUN
# ──────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
