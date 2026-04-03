"""
JMeter Report Generator - Flask API Server
Run: python app.py
API: POST /generate  →  returns HTML report
"""

import os
import sys
import tempfile
import traceback
from flask import Flask, request, Response, jsonify

# ── bootstrap: make jmeter_core importable ──────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
import importlib.util, types

def _load_core():
    spec = importlib.util.spec_from_file_location(
        "jmeter_core",
        os.path.join(os.path.dirname(__file__), "jmeter_core.py")
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod

jmeter = _load_core()

# ── Flask app ────────────────────────────────────────────────────────
app = Flask(__name__)

# def _add_cors(response):
#     response.headers["Access-Control-Allow-Origin"]  = "*"
#     response.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
#     response.headers["Access-Control-Allow-Headers"] = "Content-Type"
#     return response

def _add_cors(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "*"
    return response

@app.after_request
def after_request(response):
    return _add_cors(response)

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "version": "3.0"})

@app.route("/generate", methods=["OPTIONS"])
def generate_options():
    return _add_cors(Response("", 204))

@app.route("/generate", methods=["POST"])
def generate():
    """
    Expects multipart/form-data:
      - jtl_0, jtl_1, ...  : JTL files
      - label_0, label_1, ...: Run labels (matching count)
      - apdex_t             : int (optional, default 500)
      - output_name         : string (optional, default 'performance_report')
    Returns: HTML report as download
    """
    try:
        # ── collect files ────────────────────────────────────────────
        jtl_files  = []
        labels     = []
        i = 0
        while f"jtl_{i}" in request.files:
            jtl_files.append(request.files[f"jtl_{i}"])
            labels.append(request.form.get(f"label_{i}", f"Run-{i+1}").strip())
            i += 1

        if not jtl_files:
            return jsonify({"error": "No JTL files received. Send jtl_0, jtl_1, … fields."}), 400

        if len(labels) != len(jtl_files):
            return jsonify({"error": f"Label count ({len(labels)}) must match JTL count ({len(jtl_files)})."}), 400

        apdex_t     = int(request.form.get("apdex_t", 500))
        output_name = (request.form.get("output_name", "performance_report") or "performance_report").strip()
        if not output_name.endswith(".html"):
            output_name += ".html"

        # ── write JTLs to temp dir, parse & compute ──────────────────
        with tempfile.TemporaryDirectory() as tmpdir:
            jtl_paths = []
            for idx, f in enumerate(jtl_files):
                safe_name = f"run_{idx}_{f.filename.replace(os.sep, '_')}"
                dest = os.path.join(tmpdir, safe_name)
                f.save(dest)
                jtl_paths.append(dest)

            runs = []
            for path in jtl_paths:
                rows = jmeter.parse_jtl(path)
                runs.append(jmeter.compute(rows, T=apdex_t))

            html = jmeter.render(runs, labels, jtl_paths)

        # ── return as downloadable HTML file ─────────────────────────
        return Response(
            html,
            status=200,
            mimetype="text/html",
            headers={
                "Content-Disposition": f'attachment; filename="{output_name}"',
                "Content-Type": "text/html; charset=utf-8",
            }
        )

    except SystemExit as e:
        # jmeter_core calls sys.exit() on validation errors
        msg = str(e).replace("[ERROR] ", "")
        return jsonify({"error": msg}), 422

    except Exception:
        tb = traceback.format_exc()
        print(tb, file=sys.stderr)
        return jsonify({"error": "Internal server error", "detail": tb}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"\n  JMeter Report API  →  http://localhost:{port}")
    print(f"  Health check       →  http://localhost:{port}/health")
    print(f"  Generate endpoint  →  POST http://localhost:{port}/generate\n")
    app.run(host="0.0.0.0", port=port, debug=False)
