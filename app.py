from pathlib import Path
from datetime import datetime
import json
import math
import time

import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify

from data_generator import load_project, ensure_synthetic_csv, load_data

from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

BASE_DIR = Path(__file__).resolve().parent
app = Flask(__name__)

PROJECT = load_project()
DATA_PATH, DATA_START, DATA_END = ensure_synthetic_csv(hours=2)

# Live stream starts from this moment. Each request advances the synthetic clock.
LIVE_STARTED = time.time()

def get_sensor(sensor_id):
    return next(s for s in PROJECT["sensors"] if s["sensor_id"] == sensor_id)

def live_value(sensor, elapsed):
    # Same family of signal as the historical generator.
    sec = int(elapsed)
    idx = PROJECT["sensors"].index(sensor)
    sensor_type = sensor["type"]
    phase = idx * 0.7

    minute = (sec % 7200) / 60
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
        value = 27 + 1.2 * math.sin(sec / 900) + 0.10 * math.sin(sec / 13 + phase)
    elif sensor_type == "Deflection":
        value = 0.02 * load + 0.035 * math.sin(sec / 9 + phase) + 0.012 * math.sin(sec / 37)
    else:
        value = 1.05 + 0.006 * load + 0.16 * math.sin(sec / 4 + phase) + 0.07 * math.sin(sec / 13)

    return round(value, 5)

@app.route("/")
def dashboard():
    return render_template(
        "not_scope.html",
        title="Dashboard",
        message="Not In Scope"
    )

@app.route("/devices")
@app.route("/sensors")
@app.route("/mqtt")
@app.route("/projects")
@app.route("/structures")
@app.route("/users")
@app.route("/audit-log")
@app.route("/settings")
def not_scope():
    return render_template(
        "not_scope.html",
        title="Not In Scope",
        message="Not In Scope"
    )

@app.route("/data-download", methods=["GET", "POST"])
def data_download():
    if request.method == "POST":
        selected = request.form.getlist("sensor_ids")
        start = request.form.get("start")
        end = request.form.get("end")

        df = load_data()

        if selected:
            df = df[df["sensor_id"].isin(selected)]

        if start:
            df = df[df["timestamp"] >= pd.to_datetime(start)]
        if end:
            df = df[df["timestamp"] <= pd.to_datetime(end)]

        if df.empty:
            return render_template(
                "data_download.html",
                project=PROJECT,
                sensors=PROJECT["sensors"],
                data_start=DATA_START.strftime("%Y-%m-%d"),
                data_end=DATA_END.strftime("%Y-%m-%d"),
                error="No sensor data found for the selected range."
            )

        output_dir = BASE_DIR / "outputs" / "excel"
        output_dir.mkdir(parents=True, exist_ok=True)
        output = output_dir / "SHM_Sensor_Data.xlsx"

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            project_df = pd.DataFrame([
                ["Project ID", PROJECT["project_id"]],
                ["Project Name", PROJECT["project_name"]],
                ["Client", PROJECT["client"]],
                ["Site Location", PROJECT["site_location"]],
                ["Structure", PROJECT["structure_type"]],
                ["Test Type", PROJECT["test_type"]],
                ["Data Start", str(df["timestamp"].min())],
                ["Data End", str(df["timestamp"].max())],
                ["Sampling", "1 second"],
                ["Records", len(df)]
            ], columns=["Parameter", "Value"])
            project_df.to_excel(writer, sheet_name="Project Information", index=False)

            pd.DataFrame(PROJECT["sensors"]).to_excel(
                writer, sheet_name="Sensor Mapping", index=False
            )

            df.to_excel(writer, sheet_name="All Sensor Data", index=False)

            for sensor_id, sensor_df in df.groupby("sensor_id"):
                sensor_df.to_excel(
                    writer,
                    sheet_name=sensor_id[:31],
                    index=False
                )

        wb = load_workbook(output)

        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")

            for col in range(1, ws.max_column + 1):
                max_len = 12
                for row in range(1, min(ws.max_row, 250) + 1):
                    value = ws.cell(row, col).value
                    if value is not None:
                        max_len = max(max_len, len(str(value)))
                ws.column_dimensions[get_column_letter(col)].width = min(max_len + 2, 35)

        # Add chart to every sensor sheet.
        for sensor in PROJECT["sensors"]:
            if sensor["sensor_id"] not in wb.sheetnames:
                continue

            ws = wb[sensor["sensor_id"]]
            chart = LineChart()
            chart.title = f'{sensor["sensor_id"]} - {sensor["type"]}'
            chart.y_axis.title = sensor["unit"]
            chart.x_axis.title = "Time"

            # Value is column F.
            values = Reference(
                ws,
                min_col=6,
                min_row=1,
                max_row=ws.max_row
            )
            chart.add_data(values, titles_from_data=True)
            chart.height = 7
            chart.width = 14
            ws.add_chart(chart, "I2")

        wb.save(output)
        return send_file(output, as_attachment=True)

    return render_template(
        "data_download.html",
        project=PROJECT,
        sensors=PROJECT["sensors"],
        data_start=DATA_START.strftime("%Y-%m-%d"),
        data_end=DATA_END.strftime("%Y-%m-%d"),
        error=None
    )

@app.route("/reports")
def reports():
    return render_template(
        "reports.html",
        project=PROJECT,
        sensors=PROJECT["sensors"],
        data_start=DATA_START.strftime("%Y-%m-%d"),
        data_end=DATA_END.strftime("%Y-%m-%d")
    )

@app.route("/generate-report", methods=["POST"])
def generate_report():
    df = load_data()

    output_dir = BASE_DIR / "outputs" / "reports"
    output_dir.mkdir(parents=True, exist_ok=True)

    chart_dir = output_dir / "charts"
    chart_dir.mkdir(parents=True, exist_ok=True)

    chart_paths = []

    for sensor in PROJECT["sensors"]:
        sdf = df[df["sensor_id"] == sensor["sensor_id"]]

        # For the Word report, plot one point per 10 seconds to keep the document practical.
        sdf = sdf.iloc[::10]

        plt.figure(figsize=(9, 4))
        plt.plot(
            sdf["timestamp"],
            sdf["value"],
            linewidth=1
        )
        plt.title(f'{sensor["sensor_id"]} - {sensor["type"]}')
        plt.xlabel("Time")
        plt.ylabel(sensor["unit"])
        plt.xticks(rotation=25)
        plt.tight_layout()

        path = chart_dir / f'{sensor["sensor_id"]}.png'
        plt.savefig(path, dpi=160)
        plt.close()
        chart_paths.append(path)

    summary = (
        df.groupby(
            ["sensor_id", "sensor_type", "location", "unit"]
        )["value"]
        .agg(["count", "min", "max", "mean"])
        .reset_index()
    )

    output = output_dir / "SHM_Automated_Report.docx"

    doc = Document()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(PROJECT["report_head"])
    r.bold = True
    r.font.size = Pt(18)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(PROJECT["project_name"]).bold = True

    doc.add_heading("1. Project Information", level=1)

    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Parameter"
    table.rows[0].cells[1].text = "Value"

    project_info = [
        ("Project ID", PROJECT["project_id"]),
        ("Client", PROJECT["client"]),
        ("Site Location", PROJECT["site_location"]),
        ("Structure", PROJECT["structure_type"]),
        ("Test Type", PROJECT["test_type"]),
        ("Sampling", "1 second"),
        ("Data Start", str(df["timestamp"].min())),
        ("Data End", str(df["timestamp"].max()))
    ]

    for key, value in project_info:
        cells = table.add_row().cells
        cells[0].text = key
        cells[1].text = str(value)

    doc.add_heading("2. Sensor Summary", level=1)

    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"

    headers = [
        "Sensor", "Type", "Location",
        "Minimum", "Maximum", "Average"
    ]

    for i, header in enumerate(headers):
        table.rows[0].cells[i].text = header

    for _, row in summary.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(row["sensor_id"])
        cells[1].text = str(row["sensor_type"])
        cells[2].text = str(row["location"])
        cells[3].text = f'{row["min"]:.4f}'
        cells[4].text = f'{row["max"]:.4f}'
        cells[5].text = f'{row["mean"]:.4f}'

    doc.add_heading("3. Sensor Trend Charts", level=1)

    for path in chart_paths:
        doc.add_paragraph(path.stem)
        doc.add_picture(str(path), width=Inches(6.2))

    doc.add_heading("4. Observation", level=1)
    doc.add_paragraph(
        "This prototype report is generated automatically from one-second "
        "synthetic sensor telemetry. Final design values, acceptance criteria, "
        "temperature correction, recovery calculations and engineering "
        "interpretation will be configured from the approved project methodology."
    )

    doc.save(output)

    return send_file(output, as_attachment=True)

@app.route("/api/live")
def api_live():
    elapsed = time.time() - LIVE_STARTED
    now = datetime.now().strftime("%H:%M:%S")

    readings = []

    for sensor in PROJECT["sensors"]:
        readings.append({
            "sensor_id": sensor["sensor_id"],
            "type": sensor["type"],
            "location": sensor["location"],
            "unit": sensor["unit"],
            "timestamp": now,
            "value": live_value(sensor, elapsed)
        })

    return jsonify({
        "timestamp": now,
        "readings": readings
    })

if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
