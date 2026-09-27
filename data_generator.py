from pathlib import Path
from datetime import datetime, timedelta
import math
import csv
import json
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent

def load_project():
    with open(BASE_DIR / "config" / "projects.json", encoding="utf-8") as f:
        return json.load(f)["project"]

def sensor_value(sensor_type, sec, sensor_index):
    # Synthetic one-second telemetry. This is intentionally deterministic.
    phase = sensor_index * 0.7
    minute = sec / 60.0

    # Simulated load stages for a static-load-test style dataset.
    if minute < 20:
        load = 0
    elif minute < 40:
        load = 25
    elif minute < 60:
        load = 50
    elif minute < 80:
        load = 75
    elif minute < 100:
        load = 100
    else:
        load = 50

    if sensor_type == "Temperature":
        return 27 + 1.2 * math.sin(sec / 900) + 0.10 * math.sin(sec / 13 + phase)

    if sensor_type == "Deflection":
        return 0.02 * load + 0.035 * math.sin(sec / 9 + phase) + 0.012 * math.sin(sec / 37)

    if sensor_type == "Vibration":
        return 1.05 + 0.006 * load + 0.16 * math.sin(sec / 4 + phase) + 0.07 * math.sin(sec / 13)

    return 0.0

def ensure_synthetic_csv(hours=2):
    project = load_project()
    path = BASE_DIR / "data" / "synthetic_sensor_data.csv"

    # Regenerate at every launch so the date is always current.
    start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([
            "timestamp", "sensor_id", "sensor_type",
            "location", "unit", "value", "load_percent"
        ])

        for sec in range(hours * 3600):
            ts = start + timedelta(seconds=sec)
            minute = sec / 60.0

            if minute < 20:
                load = 0
            elif minute < 40:
                load = 25
            elif minute < 60:
                load = 50
            elif minute < 80:
                load = 75
            elif minute < 100:
                load = 100
            else:
                load = 50

            for idx, sensor in enumerate(project["sensors"]):
                value = sensor_value(sensor["type"], sec, idx)
                writer.writerow([
                    ts.strftime("%Y-%m-%d %H:%M:%S"),
                    sensor["sensor_id"],
                    sensor["type"],
                    sensor["location"],
                    sensor["unit"],
                    round(value, 5),
                    load
                ])

    return path, start, start + timedelta(seconds=hours * 3600 - 1)

def load_data():
    path = BASE_DIR / "data" / "synthetic_sensor_data.csv"
    return pd.read_csv(path, parse_dates=["timestamp"])
