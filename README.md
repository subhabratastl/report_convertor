# JMeter Report Generator — Setup Guide

## Files
```
app.py              ← Flask API server  (runs the Python script)
jmeter_core.py      ← Your original script  (core logic, no changes needed)
index.html          ← Frontend UI  (open in any browser)
requirements.txt    ← Python dependencies
```

## Quick Start

### 1. Install dependencies
```bash
pip install flask
```

### 2. Start the API server
```bash
python app.py
```
You'll see:
```
  JMeter Report API  →  http://localhost:5000
  Health check       →  http://localhost:5000/health
  Generate endpoint  →  POST http://localhost:5000/generate
```

### 3. Open the UI
Open `index.html` in your browser (double-click or `file:///path/to/index.html`).

### 4. Generate a report
1. Set the API URL (default: `http://localhost:5000`)
2. Click **Check** to verify the server is online
3. Browse and select your `.jtl` file(s)
4. Enter a run label per file
5. Set the output filename and APDEX threshold
6. Click **Generate Report** — the HTML report downloads automatically

## API Reference

### POST /generate
Multipart form fields:
- `jtl_0`, `jtl_1`, …   — JTL files
- `label_0`, `label_1`, … — Run labels (must match file count)
- `apdex_t`              — APDEX threshold in ms (default: 500)
- `output_name`          — Output filename (default: performance_report)

Returns: `text/html` file download

### GET /health
Returns: `{"status": "ok", "version": "3.0"}`

## Notes
- Multiple JTL files = comparison report with trend arrows
- TC- rows are excluded from all KPI calculations (view only)
- Report is fully self-contained HTML with embedded charts
