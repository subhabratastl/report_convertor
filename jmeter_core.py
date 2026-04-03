#!/usr/bin/env python3
# JMeter Professional HTML Report Generator v3.0
# TC- rows: displayed in separate "Transaction Controllers" view only
# All metrics/calculations use API rows only (non-TC)
#
# USAGE:
#   python jmeter_report_v3.py --jtl load_results.jtl --labels "Load Test" --output report.html
#
# MULTI-RUN COMPARISON:
#   python jmeter_report_v3.py --jtl run1.jtl run2.jtl --labels "Load" "Stress" --output report.html

import argparse, csv, json, os, sys
from collections import defaultdict
from datetime import datetime, timezone

REQUIRED = {"timeStamp", "elapsed", "label", "success"}

def is_tc(label):
    """Returns True if the label is a Transaction Controller row (starts with TC -)"""
    return label.strip().upper().startswith("TC -") or label.strip().upper().startswith("TC-")

# ══════════════════════════════════════════════════════════
# 1. PARSE
# ══════════════════════════════════════════════════════════
def parse_jtl(filepath):
    if not os.path.exists(filepath):
        sys.exit("[ERROR] Not found: " + filepath)
    rows = []
    with open(filepath, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        miss = REQUIRED - set(reader.fieldnames or [])
        if miss:
            sys.exit("[ERROR] Missing columns " + str(miss) + " in " + filepath)
        for row in reader:
            try:
                row["_e"]  = int(row["elapsed"])
                row["_t"]  = int(row["timeStamp"])
                row["_ok"] = row["success"].strip().lower() == "true"
                row["_lb"] = row.get("label", "Unknown").strip()
                row["_rc"] = row.get("responseCode", "").strip()
                row["_rm"] = row.get("responseMessage", "").strip()
                row["_fm"] = row.get("failureMessage", "").strip()
                row["_by"] = int(row.get("bytes", 0) or 0)
                row["_tc"] = is_tc(row["_lb"])   # flag TC rows
                rows.append(row)
            except Exception:
                continue
    if not rows:
        sys.exit("[ERROR] No valid rows in " + filepath)
    rows.sort(key=lambda r: r["_t"])
    api_rows = [r for r in rows if not r["_tc"]]
    tc_rows  = [r for r in rows if r["_tc"]]
    print("[INFO] {:,} total samples  ({:,} API rows, {:,} TC rows)  <- {}".format(
        len(rows), len(api_rows), len(tc_rows), filepath))
    return rows  # return all; compute() will split internally

# ══════════════════════════════════════════════════════════
# 2. HELPERS
# ══════════════════════════════════════════════════════════
def ptile(sv, p):
    if not sv: return 0
    return sv[min(int(len(sv) * p / 100), len(sv) - 1)]

def apdex(elapsed_list, T):
    sat = sum(1 for e in elapsed_list if e <= T)
    tol = sum(1 for e in elapsed_list if T < e <= 4 * T)
    fru = sum(1 for e in elapsed_list if e > 4 * T)
    n   = len(elapsed_list)
    sc  = round((sat + tol / 2) / n, 3) if n else 0
    if   sc >= 0.94: r, c = "Excellent",    "#22c55e"
    elif sc >= 0.85: r, c = "Good",         "#84cc16"
    elif sc >= 0.70: r, c = "Fair",         "#eab308"
    elif sc >= 0.50: r, c = "Poor",         "#f97316"
    else:            r, c = "Unacceptable", "#ef4444"
    return {
        "score": sc, "rating": r, "color": c,
        "sat": sat, "tol": tol, "fru": fru, "n": n,
        "sat_p": round(sat / n * 100, 1) if n else 0,
        "tol_p": round(tol / n * 100, 1) if n else 0,
        "fru_p": round(fru / n * 100, 1) if n else 0,
    }

def score_color(err_pct):
    if err_pct == 0:  return "#22c55e"
    if err_pct < 1:   return "#84cc16"
    if err_pct < 5:   return "#f97316"
    return "#ef4444"

def trend_icon(val, prev, lower_is_better=True):
    if prev is None: return ""
    diff = val - prev
    if abs(diff) < 0.001:
        return '<span style="color:#94a3b8">&#9644;</span>'
    better = diff < 0 if lower_is_better else diff > 0
    col    = "#22c55e" if better else "#ef4444"
    arrow  = "&#9650;" if diff > 0 else "&#9660;"
    return '<span style="color:' + col + ';font-size:11px"> ' + arrow + '</span>'

def fmt_duration(seconds):
    seconds = int(seconds)
    if seconds < 60: return str(seconds) + "s"
    m = seconds // 60; s = seconds % 60
    if m < 60: return "{}m {}s".format(m, s)
    h = m // 60; m = m % 60
    return "{}h {}m {}s".format(h, m, s)

def build_per_label(label_rows, T):
    """Build per-label metrics from a list of rows for that label."""
    el  = sorted(r["_e"] for r in label_rows)
    ec  = sum(1 for r in label_rows if not r["_ok"])
    n   = len(label_rows)
    d   = max(1, (max(r["_t"] + r["_e"] for r in label_rows) -
                  min(r["_t"] for r in label_rows)) / 1000)
    ap  = apdex([r["_e"] for r in label_rows], T)
    ap["label"]   = label_rows[0]["_lb"]
    ap["samples"] = n
    return {
        "label":   label_rows[0]["_lb"],
        "samples": n,
        "errors":  ec,
        "err_pct": round(ec / n * 100, 2),
        "avg":     round(sum(el) / n, 1),
        "min":     el[0],
        "max":     el[-1],
        "median":  ptile(el, 50),
        "p90":     ptile(el, 90),
        "p95":     ptile(el, 95),
        "p99":     ptile(el, 99),
        "tps":     round(n / d, 2),
        "avg_kb":  round(sum(r["_by"] for r in label_rows) / n / 1024, 2),
        "apdex":   ap,
    }

# ══════════════════════════════════════════════════════════
# 3. COMPUTE METRICS  (API rows only for all KPIs)
# ══════════════════════════════════════════════════════════
def compute(all_rows, T=500):
    api_rows = [r for r in all_rows if not r["_tc"]]
    tc_rows  = [r for r in all_rows if r["_tc"]]

    if not api_rows:
        sys.exit("[ERROR] No API (non-TC) rows found in JTL.")

    ts  = [r["_t"] for r in api_rows]
    t0  = min(ts)
    t1  = max(r["_t"] + r["_e"] for r in api_rows)
    dur = (t1 - t0) / 1000.0

    start_dt    = datetime.fromtimestamp(t0 / 1000, tz=timezone.utc)
    end_dt      = datetime.fromtimestamp(t1 / 1000, tz=timezone.utc)
    start_str   = start_dt.strftime("%Y-%m-%d %H:%M:%S UTC")
    end_str     = end_dt.strftime("%Y-%m-%d %H:%M:%S UTC")
    start_local = start_dt.strftime("%d %b %Y, %I:%M:%S %p")
    end_local   = end_dt.strftime("%d %b %Y, %I:%M:%S %p")
    dur_str     = fmt_duration(dur)

    # ── per-label (API only) ──
    by_label = defaultdict(list)
    for r in api_rows:
        by_label[r["_lb"]].append(r)

    per_label = []
    for lb in sorted(by_label.keys()):
        per_label.append(build_per_label(by_label[lb], T))

    # ── per-label TC (for display only) ──
    by_tc = defaultdict(list)
    for r in tc_rows:
        by_tc[r["_lb"]].append(r)

    per_tc = []
    for lb in sorted(by_tc.keys()):
        per_tc.append(build_per_label(by_tc[lb], T))

    # ── throughput timeline (API rows only) ──
    bkt = defaultdict(lambda: {"t": 0, "e": 0})
    for r in api_rows:
        s = r["_t"] // 1000
        bkt[s]["t"] += 1
        if not r["_ok"]: bkt[s]["e"] += 1

    tl, tv, ev = [], [], []
    for s in sorted(bkt):
        tl.append(datetime.fromtimestamp(s, tz=timezone.utc).strftime("%H:%M:%S"))
        tv.append(bkt[s]["t"])
        ev.append(bkt[s]["e"])

    # ── error reasons (API rows only) ──
    rm = defaultdict(lambda: {"n": 0, "lbs": set()})
    for r in api_rows:
        if not r["_ok"]:
            k = r["_fm"] or r["_rm"] or r["_rc"] or "Unknown"
            rm[k]["n"] += 1
            rm[k]["lbs"].add(r["_lb"])
    reasons = sorted(
        [{"reason": k, "count": v["n"], "labels": ",".join(sorted(v["lbs"]))}
         for k, v in rm.items()],
        key=lambda x: -x["count"]
    )

    codes = defaultdict(int)
    for r in api_rows:
        codes[r["_rc"] or "N/A"] += 1

    all_e = sorted(r["_e"] for r in api_rows)
    errs  = sum(1 for r in api_rows if not r["_ok"])
    ov_ap = apdex(all_e, T)

    return {
        "total":        len(api_rows),        # ← API rows only
        "tc_total":     len(tc_rows),          # ← TC rows (display info only)
        "errors":       errs,
        "err_pct":      round(errs / len(api_rows) * 100, 2),
        "avg":          round(sum(all_e) / len(all_e), 1),
        "min":          all_e[0],
        "max":          all_e[-1],
        "p90":          ptile(all_e, 90),
        "p95":          ptile(all_e, 95),
        "p99":          ptile(all_e, 99),
        "tps":          round(len(api_rows) / dur, 2),
        "dur":          round(dur, 1),
        "dur_str":      dur_str,
        "start":        start_str,
        "end":          end_str,
        "start_local":  start_local,
        "end_local":    end_local,
        "per_label":    per_label,   # API only
        "per_tc":       per_tc,      # TC only (view only)
        "tps_labels":   tl,
        "tps_vals":     tv,
        "tps_err_vals": ev,
        "reasons":      reasons,
        "codes":        dict(codes),
        "apdex":        ov_ap,
        "apdex_t":      T,
        "peak_tps":     max(tv) if tv else 0,
    }

# ══════════════════════════════════════════════════════════
# 4. HTML BUILDERS
# ══════════════════════════════════════════════════════════
PALETTE = ["#6366f1","#22d3ee","#22c55e","#f97316","#a855f7",
           "#ec4899","#fbbf24","#34d399","#60a5fa","#fb923c"]

def kpi(label, value, sub="", color=""):
    vs = ' style="color:' + color + '"' if color else ""
    return (
        '<div class="kpi-card">'
        '<div class="kpi-lbl">' + label + '</div>'
        '<div class="kpi-val"' + vs + '>' + str(value) + '</div>'
        '<div class="kpi-sub">' + sub + '</div>'
        '</div>'
    )

def badge(text, color="#6366f1"):
    return ('<span class="badge" style="background:' + color + '22;color:' + color +
            ';border:1px solid ' + color + '44">' + text + '</span>')

def err_badge(pct):
    if pct == 0: return badge("0%", "#22c55e")
    if pct < 5:  return badge(str(pct) + "%", "#f97316")
    return badge(str(pct) + "%", "#ef4444")

def apdex_badge(ap):
    return badge(ap["rating"] + " " + str(ap["score"]), ap["color"])

def time_info_card(run):
    return (
        '<div class="time-card">'
        '<div class="time-row">'
        '<div class="time-item">'
        '<div class="time-lbl">&#128197; Test Start Time</div>'
        '<div class="time-val">' + run["start_local"] + '</div>'
        '<div class="time-sub">' + run["start"] + '</div>'
        '</div>'
        '<div class="time-sep">&#8594;</div>'
        '<div class="time-item">'
        '<div class="time-lbl">&#128197; Test End Time</div>'
        '<div class="time-val">' + run["end_local"] + '</div>'
        '<div class="time-sub">' + run["end"] + '</div>'
        '</div>'
        '<div class="time-sep">&#9202;</div>'
        '<div class="time-item">'
        '<div class="time-lbl">&#9200; Total Duration</div>'
        '<div class="time-val" style="color:#22d3ee">' + run["dur_str"] + '</div>'
        '<div class="time-sub">' + str(run["dur"]) + ' seconds | '
        + '{:,}'.format(run["total"]) + ' API samples (excl. '
        + '{:,}'.format(run["tc_total"]) + ' TC rows)</div>'
        '</div>'
        '</div>'
        '</div>'
    )

def api_summary_table(run):
    """Table for API rows only."""
    rows_html = ""
    for r in run["per_label"]:
        rc = "row-danger" if r["err_pct"] >= 10 else ("row-warn" if r["err_pct"] >= 1 else "")
        rows_html += (
            '<tr class="' + rc + '">'
            '<td class="tc">' + r["label"] + '</td>'
            '<td>' + '{:,}'.format(r["samples"]) + '</td>'
            '<td>' + str(r["avg"]) + '</td>'
            '<td>' + str(r["min"]) + '</td>'
            '<td>' + str(r["max"]) + '</td>'
            '<td>' + str(r["median"]) + '</td>'
            '<td class="hi">' + str(r["p90"]) + '</td>'
            '<td class="hi">' + str(r["p95"]) + '</td>'
            '<td class="hi">' + str(r["p99"]) + '</td>'
            '<td>' + str(r["tps"]) + '</td>'
            '<td>' + str(r["avg_kb"]) + ' KB</td>'
            '<td>' + '{:,}'.format(r["errors"]) + '</td>'
            '<td>' + err_badge(r["err_pct"]) + '</td>'
            '<td>' + apdex_badge(r["apdex"]) + '</td>'
            '</tr>'
        )
    return (
        '<div class="notice-bar">&#9432; Sample counts below are <b>API requests only</b> — Transaction Controller (TC) rows are excluded from all calculations.</div>'
        '<div class="tbl-wrap"><table>'
        '<thead><tr>'
        '<th>API / Request</th><th>Samples</th><th>Avg(ms)</th>'
        '<th>Min</th><th>Max</th><th>Median</th>'
        '<th>P90</th><th>P95</th><th>P99</th>'
        '<th>TPS</th><th>Avg Size</th><th>Errors</th><th>Err%</th><th>APDEX</th>'
        '</tr></thead>'
        '<tbody>' + rows_html + '</tbody>'
        '</table></div>'
    )

def tc_summary_table(run):
    """Table for TC rows (view only, no metrics impact)."""
    if not run["per_tc"]:
        return '<p style="color:var(--muted);font-size:13px;padding:16px">No Transaction Controller rows found in this JTL.</p>'
    rows_html = ""
    for r in run["per_tc"]:
        rc = "row-danger" if r["err_pct"] >= 10 else ("row-warn" if r["err_pct"] >= 1 else "")
        rows_html += (
            '<tr class="' + rc + '">'
            '<td class="tc" style="color:#a855f7">' + r["label"] + '</td>'
            '<td>' + '{:,}'.format(r["samples"]) + '</td>'
            '<td>' + str(r["avg"]) + '</td>'
            '<td>' + str(r["min"]) + '</td>'
            '<td>' + str(r["max"]) + '</td>'
            '<td>' + str(r["median"]) + '</td>'
            '<td class="hi">' + str(r["p90"]) + '</td>'
            '<td class="hi">' + str(r["p95"]) + '</td>'
            '<td class="hi">' + str(r["p99"]) + '</td>'
            '<td>' + str(r["tps"]) + '</td>'
            '<td>' + str(r["avg_kb"]) + ' KB</td>'
            '<td>' + '{:,}'.format(r["errors"]) + '</td>'
            '<td>' + err_badge(r["err_pct"]) + '</td>'
            '<td>' + apdex_badge(r["apdex"]) + '</td>'
            '</tr>'
        )
    return (
        '<div class="notice-bar tc-notice">&#9432; Transaction Controllers are <b>view only</b> — they aggregate the API calls within them. These rows are <b>excluded</b> from all KPI calculations above.</div>'
        '<div class="tbl-wrap"><table>'
        '<thead><tr>'
        '<th>Transaction Controller</th><th>Samples</th><th>Avg(ms)</th>'
        '<th>Min</th><th>Max</th><th>Median</th>'
        '<th>P90</th><th>P95</th><th>P99</th>'
        '<th>TPS</th><th>Avg Size</th><th>Errors</th><th>Err%</th><th>APDEX</th>'
        '</tr></thead>'
        '<tbody>' + rows_html + '</tbody>'
        '</table></div>'
    )

def pct_table(run):
    rows_html = ""
    for r in run["per_label"]:
        rows_html += (
            '<tr>'
            '<td class="tc">' + r["label"] + '</td>'
            '<td>' + '{:,}'.format(r["samples"]) + '</td>'
            '<td>' + str(r["avg"]) + '</td>'
            '<td>' + str(r["median"]) + '</td>'
            '<td class="hi">' + str(r["p90"]) + '</td>'
            '<td class="hi">' + str(r["p95"]) + '</td>'
            '<td class="hi">' + str(r["p99"]) + '</td>'
            '<td>' + str(r["max"]) + '</td>'
            '</tr>'
        )
    return (
        '<div class="tbl-wrap"><table>'
        '<thead><tr><th>API / Request</th><th>Samples</th><th>Avg(ms)</th>'
        '<th>Median</th><th>P90(ms)</th><th>P95(ms)</th><th>P99(ms)</th><th>Max(ms)</th>'
        '</tr></thead><tbody>' + rows_html + '</tbody></table></div>'
    )

def apdex_table(run):
    rows_html = ""
    for r in run["per_label"]:
        ap   = r["apdex"]
        fill = int(ap["score"] * 100)
        rows_html += (
            '<tr>'
            '<td class="tc">' + r["label"] + '</td>'
            '<td>' + '{:,}'.format(r["samples"]) + '</td>'
            '<td>'
            '<div style="display:flex;align-items:center;gap:8px">'
            '<div style="flex:1;background:#2e3148;border-radius:999px;height:7px;min-width:70px">'
            '<div style="width:' + str(fill) + '%;background:' + ap["color"] + ';border-radius:999px;height:7px"></div>'
            '</div>'
            '<b style="color:' + ap["color"] + '">' + str(ap["score"]) + '</b>'
            '</div></td>'
            '<td>' + badge(ap["rating"], ap["color"]) + '</td>'
            '<td style="color:#22c55e">' + '{:,}'.format(ap["sat"]) + ' <small>(' + str(ap["sat_p"]) + '%)</small></td>'
            '<td style="color:#eab308">' + '{:,}'.format(ap["tol"]) + ' <small>(' + str(ap["tol_p"]) + '%)</small></td>'
            '<td style="color:#ef4444">' + '{:,}'.format(ap["fru"]) + ' <small>(' + str(ap["fru_p"]) + '%)</small></td>'
            '</tr>'
        )
    return (
        '<div class="tbl-wrap"><table>'
        '<thead><tr><th>API / Request</th><th>Samples</th><th>Score</th>'
        '<th>Rating</th><th>Satisfied</th><th>Tolerating</th><th>Frustrated</th>'
        '</tr></thead><tbody>' + rows_html + '</tbody></table></div>'
    )

def reasons_table(run):
    if not run["reasons"]:
        return '<div class="tbl-wrap"><table><tbody><tr><td colspan="4" style="text-align:center;padding:24px;color:#22c55e">&#10003; No errors found</td></tr></tbody></table></div>'
    rows_html = ""
    for i, e in enumerate(run["reasons"], 1):
        rows_html += (
            '<tr>'
            '<td style="text-align:center;font-weight:700;color:#94a3b8">' + str(i) + '</td>'
            '<td style="font-family:monospace;font-size:12px;color:#f97316;word-break:break-word;max-width:400px">' + str(e["reason"]) + '</td>'
            '<td>' + '{:,}'.format(e["count"]) + '</td>'
            '<td style="font-size:12px;color:#94a3b8;max-width:280px;word-break:break-word">' + e["labels"] + '</td>'
            '</tr>'
        )
    return (
        '<div class="tbl-wrap"><table>'
        '<thead><tr><th>#</th><th>Error / Failure Message</th><th>Count</th><th>Affected APIs</th></tr></thead>'
        '<tbody>' + rows_html + '</tbody></table></div>'
    )

# ══════════════════════════════════════════════════════════
# 5. COMPARISON TABLES
# ══════════════════════════════════════════════════════════
def comparison_tables(runs, labels):
    all_labels = []
    seen = set()
    for run in runs:
        for r in run["per_label"]:
            if r["label"] not in seen:
                all_labels.append(r["label"])
                seen.add(r["label"])

    def get(run, lb, key):
        for r in run["per_label"]:
            if r["label"] == lb: return r.get(key, "-")
        return "-"

    thead = "<tr><th>API / Request</th>"
    for i, lb in enumerate(labels):
        col = PALETTE[i % len(PALETTE)]
        thead += '<th style="color:' + col + '">' + lb + '</th>'
    thead += "</tr>"

    metrics_list = [
        ("avg",     "Avg (ms)", True),
        ("p90",     "P90 (ms)", True),
        ("p95",     "P95 (ms)", True),
        ("p99",     "P99 (ms)", True),
        ("tps",     "TPS",      False),
        ("err_pct", "Err %",    True),
    ]

    tables_html = ""
    for key, title, lower_better in metrics_list:
        tbody = ""
        for lb in all_labels:
            vals = [get(run, lb, key) for run in runs]
            tbody += "<tr><td class='tc'>" + lb + "</td>"
            for idx, v in enumerate(vals):
                prev  = vals[idx - 1] if idx > 0 else None
                trend = ""
                if prev is not None and isinstance(v, (int, float)) and isinstance(prev, (int, float)):
                    trend = trend_icon(v, prev, lower_better)
                tbody += "<td>" + str(v) + trend + "</td>"
            tbody += "</tr>"
        tables_html += (
            '<div class="section-sub-title">' + title + '</div>'
            '<div class="tbl-wrap" style="margin-bottom:24px"><table>'
            '<thead>' + thead + '</thead>'
            '<tbody>' + tbody + '</tbody>'
            '</table></div>'
        )
    return tables_html

def exec_comparison_table(runs, labels):
    rows_html = ""
    metrics = [
        ("total",   "API Samples (excl. TC)", False),
        ("avg",     "Avg Response (ms)",      True),
        ("p90",     "P90 (ms)",               True),
        ("p95",     "P95 (ms)",               True),
        ("tps",     "Throughput (TPS)",       False),
        ("err_pct", "Error Rate (%)",         True),
        ("dur",     "Duration (s)",           True),
    ]
    for key, title, lower_better in metrics:
        row  = "<tr><td class='tc'>" + title + "</td>"
        vals = [run[key] for run in runs]
        for idx, v in enumerate(vals):
            prev  = vals[idx - 1] if idx > 0 else None
            trend = ""
            if prev is not None:
                trend = trend_icon(v, prev, lower_better)
            col = ""
            if key == "err_pct":
                col = "color:" + score_color(v)
            row += '<td style="font-weight:600;' + col + '">' + str(v) + trend + "</td>"
        row += "</tr>"
        rows_html += row

    row = "<tr><td class='tc'>APDEX Score</td>"
    for run in runs:
        ap = run["apdex"]
        row += '<td>' + badge(ap["rating"] + " " + str(ap["score"]), ap["color"]) + '</td>'
    row += "</tr>"
    rows_html += row

    thead = "<tr><th>Metric</th>"
    for i, lb in enumerate(labels):
        col = PALETTE[i % len(PALETTE)]
        thead += '<th style="color:' + col + '">' + lb + '</th>'
    thead += "</tr>"

    return (
        '<div class="tbl-wrap"><table>'
        '<thead>' + thead + '</thead>'
        '<tbody>' + rows_html + '</tbody>'
        '</table></div>'
    )

# ══════════════════════════════════════════════════════════
# 6. DOWNLOAD DOCX JAVASCRIPT (embedded in HTML)
# ══════════════════════════════════════════════════════════
DOCX_JS = r"""
/* ════════════════════════════════════════════════════════
   DOWNLOAD AS WORD DOCUMENT
   Uses raw Open XML + JSZip — zero dependency on docx.js
   TC rows excluded from all Word doc metrics (same as HTML)
   ════════════════════════════════════════════════════════ */

function xe(tag, attrs, ...children) {
  let a = '';
  for (const [k,v] of Object.entries(attrs||{})) a += ` ${k}="${v}"`;
  const inner = children.join('');
  return inner ? `<${tag}${a}>${inner}</${tag}>` : `<${tag}${a}/>`;
}
function esc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

function rPr(opts={}){
  let x='';
  if(opts.bold)   x+=xe('w:b',{});
  if(opts.italic) x+=xe('w:i',{});
  if(opts.color)  x+=xe('w:color',{'w:val':opts.color.replace('#','').toUpperCase()});
  if(opts.size)   x+=xe('w:sz',{'w:val':String(opts.size)})+xe('w:szCs',{'w:val':String(opts.size)});
  if(opts.font)   x+=xe('w:rFonts',{'w:ascii':opts.font,'w:hAnsi':opts.font,'w:cs':opts.font});
  return x ? xe('w:rPr',{},x) : '';
}
function run(text, opts={}){
  return xe('w:r',{}, rPr(opts), xe('w:t',{'xml:space':'preserve'}, esc(text)));
}
function pPrXml(opts={}){
  let x='';
  if(opts.align && opts.align!=='left') x+=xe('w:jc',{'w:val':opts.align});
  if(opts.spaceBefore || opts.spaceAfter){
    const a={};
    if(opts.spaceBefore) a['w:before']=String(opts.spaceBefore);
    if(opts.spaceAfter)  a['w:after'] =String(opts.spaceAfter);
    x+=xe('w:spacing',a);
  }
  if(opts.borderBottom){
    x+=xe('w:pBdr',{},xe('w:bottom',{'w:val':'single','w:sz':'12','w:space':'1','w:color':opts.borderBottom}));
  }
  if(opts.numId){
    x+=xe('w:numPr',{},xe('w:ilvl',{'w:val':'0'}),xe('w:numId',{'w:val':opts.numId}));
  }
  return x ? xe('w:pPr',{},x) : '';
}
function para(runs_xml, opts={}){
  return xe('w:p',{}, pPrXml(opts), runs_xml);
}
function paraBreak(){ return xe('w:p',{}, xe('w:r',{}, xe('w:br',{'w:type':'page'}))); }

function tcPrXml(opts={}){
  let x='';
  if(opts.fill) x+=xe('w:shd',{'w:val':'clear','w:color':'auto','w:fill':opts.fill.replace('#','').toUpperCase()});
  if(opts.borders!==false){
    const bc = (opts.borderColor||'D1D5DB').replace('#','').toUpperCase();
    const bs = xe('w:top',{'w:val':'single','w:sz':'4','w:space':'0','w:color':bc})
              +xe('w:left',{'w:val':'single','w:sz':'4','w:space':'0','w:color':bc})
              +xe('w:bottom',{'w:val':'single','w:sz':'4','w:space':'0','w:color':bc})
              +xe('w:right',{'w:val':'single','w:sz':'4','w:space':'0','w:color':bc});
    x+=xe('w:tcBorders',{},bs);
  }
  if(opts.widthTwips) x+=xe('w:tcW',{'w:w':String(opts.widthTwips),'w:type':'dxa'});
  if(opts.vAlign)     x+=xe('w:vAlign',{'w:val':opts.vAlign});
  return x ? xe('w:tcPr',{},x) : '';
}
function cell(content_xml, opts={}){
  const align = opts.align||'left';
  const spacingXml = xe('w:spacing',{'w:before':'40','w:after':'40'});
  const jcXml = align!=='left'?xe('w:jc',{'w:val':align}):'';
  const pPr = (spacingXml||jcXml)?xe('w:pPr',{},spacingXml,jcXml):'';
  return xe('w:tc',{}, tcPrXml(opts), xe('w:p',{}, pPr, content_xml));
}

const PAGE_W = 9638;
function pct2twip(pct){ return Math.round(PAGE_W * pct / 100); }

function hdrRow(headers, widths_pct, bgHex){
  const bg = (bgHex||'1B2A4A').replace('#','').toUpperCase();
  const cells = headers.map((h,i)=>{
    const tw = pct2twip(widths_pct[i]);
    const runXml = run(String(h),{bold:true,color:'FFFFFF',size:17,font:'Arial'});
    return cell(runXml,{fill:bg,borderColor:bg,widthTwips:tw,vAlign:'center',align:'center'});
  });
  return xe('w:tr',{},cells.join(''));
}

function dataRow(rowData, widths_pct, rowIdx){
  const fill = rowIdx%2===0 ? 'F8FAFC' : 'FFFFFF';
  const cells = rowData.map((v,ci)=>{
    const tw = pct2twip(widths_pct[ci]);
    let text='', color='0F172A', bold=ci===0, bg=fill, align=ci===0?'left':'center';
    if(typeof v==='object'&&v!==null){
      text  = String(v.text||'');
      color = (v.color||'0F172A').replace('#','');
      bold  = v.bold!==undefined ? v.bold : ci===0;
      bg    = v.bg   || fill;
      align = ci===0?'left':'center';
    } else {
      text = String(v);
      color = ci===0 ? '1B2A4A' : '0F172A';
    }
    const runXml = run(text,{bold,color,size:17,font:'Arial'});
    return cell(runXml,{fill:bg,widthTwips:tw,vAlign:'center',align});
  });
  return xe('w:tr',{},cells.join(''));
}

function makeTable(headers, rows, widths){
  const total = widths.reduce((a,b)=>a+b,0);
  const pcts  = widths.map(w=>w/total*100);
  const tblPr = xe('w:tblPr',{},
    xe('w:tblStyle',{'w:val':'TableGrid'}),
    xe('w:tblW',{'w:w':'5000','w:type':'pct'}),
    xe('w:tblBorders',{},
      xe('w:insideH',{'w:val':'single','w:sz':'4','w:color':'D1D5DB'}),
      xe('w:insideV',{'w:val':'single','w:sz':'4','w:color':'D1D5DB'})
    )
  );
  let rowsXml = '';
  if(headers && headers.length) rowsXml += hdrRow(headers, pcts);
  rows.forEach((r,ri)=>{ rowsXml += dataRow(r, pcts, ri); });
  return xe('w:tbl',{}, tblPr, rowsXml);
}

function sectionTitle(num, title, subtitle){
  let runsXml = '';
  if(num) runsXml += run(num+'  ',{bold:true,size:28,color:'2563EB',font:'Arial'});
  runsXml += run(title,{bold:true,size:28,color:'1B2A4A',font:'Arial'});
  let xml = para(runsXml,{spaceBefore:200,spaceAfter:40,borderBottom:'2563EB'});
  if(subtitle) xml += para(run(subtitle,{size:17,color:'64748B',font:'Arial'}),{spaceBefore:40,spaceAfter:120});
  return xml;
}
function subTitle(text){
  return para(run(text,{bold:true,size:21,color:'0E7490',font:'Arial'}),{spaceBefore:160,spaceAfter:60});
}
function normalPara(text,opts={}){
  const {bold=false,size=18,color='0F172A',align='left',italic=false}=opts;
  return para(run(text,{bold,size,color:color.replace('#',''),italic,font:'Arial'}),
    {spaceBefore:40,spaceAfter:80,align});
}
function bulletPara(text,numId='1'){
  return para(run(text,{size:18,color:'0F172A',font:'Arial'}),{numId,spaceAfter:20,spaceBefore:20});
}

function errCol(p)   { return p===0?'158030':(p<5?'B45309':'B91C1C'); }
function apdexCol(s) { return s>=0.85?'158030':(s>=0.70?'B45309':'B91C1C'); }
function stCol(s)    { return s==='PASS'?'158030':(s==='WARN'?'B45309':'B91C1C'); }
function stBg(s)     { return s==='PASS'?'D1FAE5':(s==='WARN'?'FEF3C7':'FEE2E2'); }
function ovSt(ep)    { return ep<1?'PASS':(ep<5?'WARN':'FAIL'); }

function wrapDocument(bodyXml, meta){
  const date = meta.gen_time ? meta.gen_time.substring(0,10) : '';
  const hdrXml =
    xe('w:p',{},
      xe('w:pPr',{},
        xe('w:pBdr',{},xe('w:bottom',{'w:val':'single','w:sz':'8','w:space':'1','w:color':'2563EB'})),
        xe('w:spacing',{'w:after':'60'})
      ),
      run(meta.project+' — Performance Test Report  |  API Samples Only (TC rows excluded)',{bold:true,size:15,color:'1B2A4A',font:'Arial'}),
      run('  CONFIDENTIAL',{italic:true,size:14,color:'64748B',font:'Arial'})
    );
  const ftrXml =
    xe('w:p',{},
      xe('w:pPr',{},xe('w:jc',{'w:val':'center'})),
      run(date+' | '+meta.team+' | Page ',{size:15,color:'64748B',font:'Arial'}),
      xe('w:r',{},xe('w:rPr',{},xe('w:sz',{'w:val':'15'}),xe('w:color',{'w:val':'64748B'})),xe('w:fldChar',{'w:fldCharType':'begin'})),
      xe('w:r',{},xe('w:rPr',{},xe('w:sz',{'w:val':'15'}),xe('w:color',{'w:val':'64748B'})),xe('w:instrText',{'xml:space':'preserve'},' PAGE ')),
      xe('w:r',{},xe('w:rPr',{},xe('w:sz',{'w:val':'15'}),xe('w:color',{'w:val':'64748B'})),xe('w:fldChar',{'w:fldCharType':'end'})),
      run(' of ',{size:15,color:'64748B',font:'Arial'}),
      xe('w:r',{},xe('w:rPr',{},xe('w:sz',{'w:val':'15'}),xe('w:color',{'w:val':'64748B'})),xe('w:fldChar',{'w:fldCharType':'begin'})),
      xe('w:r',{},xe('w:rPr',{},xe('w:sz',{'w:val':'15'}),xe('w:color',{'w:val':'64748B'})),xe('w:instrText',{'xml:space':'preserve'},' NUMPAGES ')),
      xe('w:r',{},xe('w:rPr',{},xe('w:sz',{'w:val':'15'}),xe('w:color',{'w:val':'64748B'})),xe('w:fldChar',{'w:fldCharType':'end'}))
    );
  const sectPr =
    xe('w:sectPr',{},
      xe('w:headerReference',{'w:type':'default','r:id':'rId1'}),
      xe('w:footerReference',{'w:type':'default','r:id':'rId2'}),
      xe('w:pgSz',{'w:w':'11906','w:h':'16838'}),
      xe('w:pgMar',{'w:top':'1134','w:right':'1134','w:bottom':'907','w:left':'1134','w:header':'709','w:footer':'709'})
    );
  const docXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
    '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '+
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '+
    'xmlns:o="urn:schemas-microsoft-com:office:office" '+
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '+
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '+
    'xmlns:v="urn:schemas-microsoft-com:vml" '+
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '+
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '+
    'xmlns:w10="urn:schemas-microsoft-com:office:word" '+
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '+
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '+
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '+
    'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '+
    'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '+
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">\n'+
    '<w:body>'+bodyXml+sectPr+'</w:body></w:document>';
  const hdrFile = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
    '<w:hdr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '+
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'+hdrXml+'</w:hdr>';
  const ftrFile = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
    '<w:ftr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '+
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'+ftrXml+'</w:ftr>';
  return { docXml, hdrFile, ftrFile };
}

async function packDocx(bodyXml, meta){
  const {docXml, hdrFile, ftrFile} = wrapDocument(bodyXml, meta);
  const rels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'+
    '</Relationships>';
  const docRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'+
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'+
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'+
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'+
    '</Relationships>';
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+
    '<Default Extension="xml" ContentType="application/xml"/>'+
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'+
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'+
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'+
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'+
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'+
    '</Types>';
  const stylesXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'+
    '<w:docDefaults><w:rPrDefault><w:rPr>'+
    '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'+
    '<w:sz w:val="20"/><w:szCs w:val="20"/>'+
    '</w:rPr></w:rPrDefault></w:docDefaults>'+
    '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>'+
    '<w:tblPr><w:tblBorders>'+
    '<w:top w:val="single" w:sz="4" w:color="D1D5DB"/>'+
    '<w:left w:val="single" w:sz="4" w:color="D1D5DB"/>'+
    '<w:bottom w:val="single" w:sz="4" w:color="D1D5DB"/>'+
    '<w:right w:val="single" w:sz="4" w:color="D1D5DB"/>'+
    '<w:insideH w:val="single" w:sz="4" w:color="D1D5DB"/>'+
    '<w:insideV w:val="single" w:sz="4" w:color="D1D5DB"/>'+
    '</w:tblBorders></w:tblPr></w:style></w:styles>';
  const numberingXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'+
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'+
    '<w:lvlText w:val="•"/><w:lvlJc w:val="left"/>'+
    '<w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>'+
    '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>'+
    '</w:lvl></w:abstractNum>'+
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypes);
  zip.file('_rels/.rels', rels);
  zip.file('word/document.xml', docXml);
  zip.file('word/_rels/document.xml.rels', docRels);
  zip.file('word/styles.xml', stylesXml);
  zip.file('word/numbering.xml', numberingXml);
  zip.file('word/header1.xml', hdrFile);
  zip.file('word/footer1.xml', ftrFile);
  return await zip.generateAsync({type:'blob', mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
}

async function downloadDocx() {
  const btn = document.getElementById('dl-btn');
  btn.disabled = true;
  btn.textContent = '⏳ Generating...';
  try {
    const raw = document.getElementById('report-data');
    if (!raw) { alert('No report data found.'); btn.disabled=false; btn.innerHTML='&#128196; Download Word Doc'; return; }
    const D    = JSON.parse(raw.textContent);
    const meta = D.meta;
    const runs = D.runs;
    const lbls = D.labels;
    const multi= runs.length > 1;

    let body = '';

    /* ── COVER PAGE ── */
    body += para(run('PERFORMANCE TEST REPORT',{bold:true,size:56,color:'1B2A4A',font:'Arial'}),{align:'center',spaceBefore:600,spaceAfter:60});
    body += para(run(meta.project,{bold:true,size:36,color:'2563EB',font:'Arial'}),{align:'center',spaceAfter:40});
    body += para(run('Version '+meta.version,{size:22,color:'64748B',font:'Arial'}),{align:'center',spaceAfter:80});
    body += para(run('★ All metrics calculated from API requests only — TC rows excluded',{italic:true,size:18,color:'6366f1',font:'Arial'}),{align:'center',spaceAfter:200});
    const coverData = [
      [{text:'Prepared By',bold:true,color:'1B2A4A'}, {text:meta.team,color:'0F172A'}],
      [{text:'Report Date',bold:true,color:'1B2A4A'}, {text:new Date().toLocaleDateString('en-GB',{day:'numeric',month:'long',year:'numeric'}),color:'0F172A'}],
      [{text:'Generated',bold:true,color:'1B2A4A'},   {text:meta.gen_time,color:'0F172A'}],
      [{text:'Environment',bold:true,color:'1B2A4A'}, {text:'Staging / Pre-Production',color:'0F172A'}],
      [{text:'Test Tool',bold:true,color:'1B2A4A'},   {text:'Apache JMeter v5.6.3',color:'0F172A'}],
      [{text:'APDEX T',bold:true,color:'1B2A4A'},     {text:meta.apdex_t+' ms',color:'0F172A'}],
      [{text:'Runs',bold:true,color:'1B2A4A'},        {text:runs.length+'  ('+lbls.join(', ')+')',color:'0F172A'}],
      [{text:'Source Files',bold:true,color:'1B2A4A'},{text:(meta.jtl_files||[]).join(', '),color:'0F172A'}],
      [{text:'Note',bold:true,color:'6366F1'},        {text:'TC (Transaction Controller) rows excluded from all sample counts, KPIs, and error rates.',color:'6366F1'}],
    ];
    body += makeTable([],coverData,[35,65]);
    body += para(run('CONFIDENTIAL — For Internal Use Only',{italic:true,size:16,color:'64748B',font:'Arial'}),{align:'center',spaceBefore:200,spaceAfter:40});
    body += paraBreak();

    /* ── 1. EXECUTIVE SUMMARY ── */
    body += sectionTitle('1.','Executive Summary','High-level performance overview — API requests only (TC rows excluded)');
    body += subTitle('1.1  Test Runs Overview');
    const ovH = ['Run','API Samples','Avg(ms)','P90','P95','P99','TPS','Err%','APDEX','Status'];
    const ovW = [18,10,8,7,7,7,7,7,13,9];
    const ovD = runs.map((r,ri)=>{
      const ap=r.apdex, st=ovSt(r.err_pct);
      return [{text:lbls[ri],bold:true,color:'1B2A4A'},
        r.total.toLocaleString()+' API'+(r.tc_total?' (+'+r.tc_total+' TC excl.)':''),
        String(r.avg),
        {text:String(r.p90),color:'0E7490'},{text:String(r.p95),color:'0E7490'},{text:String(r.p99),color:'0E7490'},
        String(r.tps),{text:r.err_pct+'%',color:errCol(r.err_pct),bold:true},
        {text:ap.score+' ('+ap.rating+')',color:apdexCol(ap.score),bold:true},
        {text:st,bold:true,color:stCol(st),bg:stBg(st)}];
    });
    body += makeTable(ovH,ovD,ovW);

    body += subTitle('1.2  Key Findings');
    if(multi){
      const avgChg=((runs[runs.length-1].avg-runs[0].avg)/Math.max(runs[0].avg,1)*100).toFixed(1);
      const errChg=(runs[runs.length-1].err_pct-runs[0].err_pct).toFixed(2);
      const tpsChg=((runs[runs.length-1].tps-runs[0].tps)/Math.max(runs[0].tps,0.1)*100).toFixed(1);
      body += bulletPara('Response time changed '+avgChg+'%: '+runs[0].avg+'ms → '+runs[runs.length-1].avg+'ms');
      body += bulletPara('Error rate changed '+errChg+' pp: '+runs[0].err_pct+'% → '+runs[runs.length-1].err_pct+'%');
      body += bulletPara('Throughput changed '+tpsChg+'%: '+runs[0].tps+' TPS → '+runs[runs.length-1].tps+' TPS');
    }
    runs.forEach((r,ri)=>{
      const worst=[...r.per_label].sort((a,b)=>b.avg-a.avg);
      if(worst.length) body+=bulletPara('['+lbls[ri]+'] Slowest API: '+worst[0].label+' @ '+worst[0].avg+'ms avg');
      if(r.reasons&&r.reasons.length) body+=bulletPara('['+lbls[ri]+'] Top error: '+r.reasons[0].reason.substring(0,80)+' ('+r.reasons[0].count+'x)');
    });

    body += subTitle('1.3  Overall Verdict');
    const lastRun=runs[runs.length-1], lastSt=ovSt(lastRun.err_pct);
    body += normalPara('Status: '+lastSt+'   |   APDEX: '+lastRun.apdex.score+' — '+lastRun.apdex.rating,{bold:true,size:20,color:stCol(lastSt)});
    body += normalPara('Performance testing of '+meta.project+' '+meta.version+' completed across '+runs.length+' run(s). Latest run ('+lbls[lbls.length-1]+') recorded '+lastRun.total.toLocaleString()+' API samples with '+lastRun.avg+'ms average response, '+lastRun.err_pct+'% error rate, and '+lastRun.apdex.score+' APDEX score. TC rows were excluded from all calculations.',{size:18});
    body += paraBreak();

    /* ── 2+. PER-RUN RESULTS ── */
    runs.forEach((r,ri)=>{
      const ap=r.apdex, st=ovSt(r.err_pct), sn=(2+ri), jtlFile=(meta.jtl_files||[])[ri]||'';
      body += sectionTitle(sn+'.', lbls[ri], jtlFile?'Source: '+jtlFile+' | API rows only':'');
      body += subTitle(sn+'.1  Key Performance Indicators  (API requests only)');
      body += makeTable([],[
        [{text:'API Samples',bold:true,color:'1B2A4A'},    {text:r.total.toLocaleString()+' requests',color:'2563EB'}],
        [{text:'TC Rows (excluded)',bold:true,color:'6366F1'},{text:r.tc_total.toLocaleString()+' rows — view only, not counted',color:'6366F1'}],
        [{text:'Error Rate',bold:true,color:'1B2A4A'},     {text:r.err_pct+'%',color:errCol(r.err_pct)}],
        [{text:'Avg Response',bold:true,color:'1B2A4A'},   {text:r.avg+' ms',color:'0F172A'}],
        [{text:'Throughput',bold:true,color:'1B2A4A'},     {text:r.tps+' TPS',color:'158030'}],
        [{text:'P90 Latency',bold:true,color:'1B2A4A'},    {text:r.p90+' ms',color:'0E7490'}],
        [{text:'P95 Latency',bold:true,color:'1B2A4A'},    {text:r.p95+' ms',color:'0E7490'}],
        [{text:'P99 Latency',bold:true,color:'1B2A4A'},    {text:r.p99+' ms',color:r.p99>4000?'B91C1C':'B45309'}],
        [{text:'APDEX Score',bold:true,color:'1B2A4A'},    {text:ap.score+' / 1.0',color:apdexCol(ap.score)}],
        [{text:'Overall Status',bold:true,color:'1B2A4A'},{text:st,bold:true,color:stCol(st),bg:stBg(st)}],
      ],[40,60]);

      body += subTitle(sn+'.2  Response Time by API');
      body += makeTable(['API / Request','Samples','Avg','Min','Max','Median','P90','P95','P99','TPS','Err%','APDEX'],
        r.per_label.map(tx=>{const a2=tx.apdex;return[
          {text:tx.label,bold:true,color:'1B2A4A'},tx.samples.toLocaleString(),
          {text:String(tx.avg),color:tx.avg>2000?'B91C1C':tx.avg>1000?'B45309':'0F172A'},
          String(tx.min),String(tx.max),String(tx.median),
          {text:String(tx.p90),color:'0E7490'},{text:String(tx.p95),color:'0E7490'},{text:String(tx.p99),color:'0E7490'},
          String(tx.tps),{text:tx.err_pct+'%',color:errCol(tx.err_pct),bold:true},
          {text:a2.score+' '+a2.rating,color:apdexCol(a2.score),bold:true}];}),
        [20,8,7,6,6,7,7,7,7,7,7,8]);

      const T=r.apdex_t;
      body += subTitle(sn+'.3  APDEX Breakdown  (T='+T+'ms)');
      body += makeTable(['Zone','Condition','Count','% of Total','Meaning'],[
        [{text:'Satisfied',bold:true,color:'158030'},'Response <= '+T+'ms',ap.sat.toLocaleString(),ap.sat_p+'%','User fully satisfied'],
        [{text:'Tolerating',bold:true,color:'B45309'},'Response <= '+(T*4)+'ms',ap.tol.toLocaleString(),ap.tol_p+'%','User accepts performance'],
        [{text:'Frustrated',bold:true,color:'B91C1C'},'Response > '+(T*4)+'ms',ap.fru.toLocaleString(),ap.fru_p+'%','User likely to abandon'],
      ],[15,25,12,12,36]);
      body += normalPara('APDEX Score: '+ap.score+'  —  '+ap.rating,{bold:true,size:20,color:apdexCol(ap.score)});

      if(r.reasons&&r.reasons.length){
        body += subTitle(sn+'.4  Error Reasons');
        body += makeTable(['#','Error / Failure Message','Count','Affected APIs'],
          r.reasons.slice(0,15).map((e,idx)=>[String(idx+1),{text:String(e.reason).substring(0,120),color:'C2410C'},{text:String(e.count),color:'B91C1C',bold:true},{text:String(e.labels).substring(0,60),color:'64748B'}]),
          [5,50,10,35]);
      }

      /* TC rows — view only section */
      if(r.per_tc && r.per_tc.length){
        body += subTitle(sn+'.5  Transaction Controllers (View Only — Not Counted in Metrics)');
        body += normalPara('The following TC rows appear in the JTL for reference. They aggregate the API calls above and are excluded from all sample counts, KPIs, and error rates to avoid double-counting.',{size:16,color:'6366F1',italic:true});
        body += makeTable(['Transaction Controller','Samples','Avg(ms)','P90','P95','Err%'],
          r.per_tc.map(tc=>[
            {text:tc.label,bold:true,color:'7C3AED'},
            tc.samples.toLocaleString(),String(tc.avg),
            {text:String(tc.p90),color:'0E7490'},{text:String(tc.p95),color:'0E7490'},
            {text:tc.err_pct+'%',color:errCol(tc.err_pct),bold:true}
          ]),
          [30,12,12,12,12,12]);
      }

      if(ri<runs.length-1) body+=paraBreak();
    });
    body += paraBreak();

    /* ── COMPARISON (multi only) ── */
    if(multi){
      const cmpN=2+runs.length;
      body += sectionTitle(cmpN+'.','Run Comparison','Side-by-side API metrics  |  TC rows excluded from all values');
      body += subTitle(cmpN+'.1  Overall Metrics Comparison');
      const cmpH=['Metric',...lbls], cmpW=[28,...lbls.map(()=>Math.round(72/lbls.length))];
      const crow=(lbl,vals,colorFn,lb=true)=>{
        const row=[{text:lbl,bold:true,color:'1B2A4A'}];
        vals.forEach((v,i)=>{
          const prev=i>0?vals[i-1]:null;
          let t='';
          if(prev!=null){const d=v-prev;if(Math.abs(d)>0.001)t=(lb?d<0:d>0)?' (+)':' (-)';}
          row.push({text:String(v)+t,color:colorFn?colorFn(v,i):'0F172A',bold:i>0});
        });
        return row;
      };
      body += makeTable(cmpH,[
        crow('API Samples (excl TC)', runs.map(r=>r.total),     ()=>'2563EB',false),
        crow('Avg Response(ms)',      runs.map(r=>r.avg),        (_,i)=>errCol(runs[i].err_pct)),
        crow('P90 (ms)',              runs.map(r=>r.p90),        (_,i)=>apdexCol(runs[i].apdex.score)),
        crow('P95 (ms)',              runs.map(r=>r.p95),        (_,i)=>apdexCol(runs[i].apdex.score)),
        crow('P99 (ms)',              runs.map(r=>r.p99),        (_,i)=>apdexCol(runs[i].apdex.score)),
        crow('Throughput(TPS)',       runs.map(r=>r.tps),        ()=>'158030',false),
        crow('Error Rate(%)',         runs.map(r=>r.err_pct),    v=>errCol(v)),
        crow('Total Errors',         runs.map(r=>r.errors),     v=>v>0?'B91C1C':'158030'),
        crow('APDEX Score',          runs.map(r=>r.apdex.score),v=>apdexCol(v),false),
        [{text:'APDEX Rating',bold:true,color:'1B2A4A'},...runs.map(r=>({text:r.apdex.rating,color:apdexCol(r.apdex.score),bold:true}))],
        [{text:'Overall Status',bold:true,color:'1B2A4A'},...runs.map(r=>{const s=ovSt(r.err_pct);return{text:s,bold:true,color:stCol(s),bg:stBg(s)};})],
        crow('Duration (s)',         runs.map(r=>r.dur),         ()=>'0F172A'),
      ],cmpW);
      body += paraBreak();
    }

    /* ── CONCLUSION & SLA ── */
    const slaN=3+runs.length, lr=runs[runs.length-1];
    body += sectionTitle(slaN+'.','Conclusion & SLA Verdict','Final pass/fail assessment — API requests only');
    body += subTitle('SLA Criteria Verdict');
    const slaD=[
      ['Avg Response Time','< 1,000 ms',  lr.avg+' ms',    {text:lr.avg<1000?'PASS':'FAIL',   bold:true,color:lr.avg<1000?'158030':'B91C1C',   bg:stBg(lr.avg<1000?'PASS':'FAIL')}],
      ['P90 Response Time','< 2,000 ms',  lr.p90+' ms',    {text:lr.p90<2000?'PASS':'FAIL',   bold:true,color:lr.p90<2000?'158030':'B91C1C',   bg:stBg(lr.p90<2000?'PASS':'FAIL')}],
      ['P99 Response Time','< 5,000 ms',  lr.p99+' ms',    {text:lr.p99<5000?'PASS':'FAIL',   bold:true,color:lr.p99<5000?'158030':'B91C1C',   bg:stBg(lr.p99<5000?'PASS':'FAIL')}],
      ['Error Rate',       '< 1%',        lr.err_pct+'%',  {text:lr.err_pct<1?'PASS':'FAIL',  bold:true,color:lr.err_pct<1?'158030':'B91C1C',  bg:stBg(lr.err_pct<1?'PASS':'FAIL')}],
      ['Throughput',       '> 10 TPS',    lr.tps+' TPS',   {text:lr.tps>=10?'PASS':'FAIL',    bold:true,color:lr.tps>=10?'158030':'B91C1C',    bg:stBg(lr.tps>=10?'PASS':'FAIL')}],
      ['APDEX Score',      '> 0.70',      String(lr.apdex.score),{text:lr.apdex.score>=0.70?'PASS':'FAIL',bold:true,color:lr.apdex.score>=0.70?'158030':'B91C1C',bg:stBg(lr.apdex.score>=0.70?'PASS':'FAIL')}],
    ];
    body += makeTable(['SLA Criteria','Target','Actual (Latest Run)','Verdict'],slaD,[28,16,24,14]);
    const passed=slaD.filter(r=>r[3].text==='PASS').length;
    body += subTitle('Summary Statement');
    body += normalPara('Performance testing of '+meta.project+' '+meta.version+' completed across '+runs.length+' run(s). The latest run ('+lbls[lbls.length-1]+') passed '+passed+' of '+slaD.length+' SLA criteria. Overall status: '+ovSt(lr.err_pct)+'. All figures based on API requests only — TC rows excluded.',{size:18});
    body += subTitle('Sign-off');
    body += makeTable([],[
      [{text:'Prepared By',bold:true,color:'1B2A4A'},{text:'Reviewed By',bold:true,color:'1B2A4A'},{text:'Approved By',bold:true,color:'1B2A4A'}],
      [{text:meta.team+'\n_______________________',color:'64748B'},{text:'Engineering Lead\n_______________________',color:'64748B'},{text:'Project Manager\n_______________________',color:'64748B'}],
    ],[34,33,33]);

    /* ── PACK & DOWNLOAD ── */
    const blob = await packDocx(body, meta);
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = 'performance_report_v3.docx';
    a.click();
    URL.revokeObjectURL(url);
    btn.disabled = false;
    btn.innerHTML = '&#10003; Downloaded!';
    setTimeout(()=>{ btn.innerHTML='&#128196; Download Word Doc'; }, 3000);
  } catch(err) {
    console.error(err);
    alert('Error generating document: '+err.message);
    btn.disabled = false;
    btn.innerHTML = '&#128196; Download Word Doc';
  }
}
"""

# ══════════════════════════════════════════════════════════
# 7. RENDER FULL HTML
# ══════════════════════════════════════════════════════════
def render(runs, labels, jtl_paths):
    multi = len(runs) > 1
    gen   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    year  = datetime.now().year
    T     = runs[0]["apdex_t"]

    p = []

    p.append("""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Performance Test Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#080b14;--surf:#0f1422;--surf2:#161c2e;--surf3:#1d2540;
  --border:#232d45;--border2:#2d3a56;
  --accent:#6366f1;--teal:#22d3ee;--green:#22c55e;--red:#ef4444;
  --orange:#f97316;--text:#e8edf5;--muted:#8896b0;--dim:#4a5568;
  --r:10px;--r2:14px;
}
body{background:var(--bg);color:var(--text);font-family:'Segoe UI',system-ui,sans-serif;font-size:14px;line-height:1.5}
.header{background:linear-gradient(135deg,#0e0b2e,#0f1422,#0a1628);border-bottom:1px solid var(--border2);padding:0}
.header-inner{max-width:1500px;margin:0 auto;padding:30px 48px;display:flex;align-items:center;justify-content:space-between;gap:20px;flex-wrap:wrap}
.logo-block h1{font-size:24px;font-weight:800;color:#fff;letter-spacing:-.3px}
.logo-block h1 em{color:var(--accent);font-style:normal}
.logo-block .sub{font-size:12px;color:var(--muted);margin-top:3px}
.meta-pills{display:flex;flex-wrap:wrap;gap:8px;justify-content:flex-end;align-items:center}
.pill{padding:5px 14px;border-radius:999px;font-size:11px;font-weight:600;border:1px solid var(--border2);background:var(--surf2);color:var(--muted);display:flex;align-items:center;gap:5px}
.pill b{color:var(--text)}
.dl-btn{display:inline-flex;align-items:center;gap:8px;padding:10px 22px;border-radius:10px;border:none;cursor:pointer;background:linear-gradient(135deg,#2563EB,#1B2A4A);color:#fff;font-size:13px;font-weight:700;letter-spacing:.3px;box-shadow:0 4px 18px #2563eb44;transition:all .2s;white-space:nowrap}
.dl-btn:hover{background:linear-gradient(135deg,#1d4ed8,#0f1a35);box-shadow:0 6px 24px #2563eb66;transform:translateY(-1px)}
.dl-btn:active{transform:translateY(0)}
.dl-btn:disabled{opacity:.6;cursor:not-allowed;transform:none}
.nav-bar{background:var(--surf);border-bottom:1px solid var(--border);position:sticky;top:0;z-index:100}
.nav-inner{max-width:1500px;margin:0 auto;padding:0 48px;display:flex;flex-wrap:wrap;gap:2px}
.nav-btn{padding:13px 18px;border:none;background:none;cursor:pointer;color:var(--muted);font-size:12.5px;font-weight:600;border-bottom:2px solid transparent;transition:all .18s;white-space:nowrap}
.nav-btn:hover{color:var(--text);background:var(--surf2)}
.nav-btn.active{color:var(--accent);border-bottom-color:var(--accent)}
.nav-btn.tc-nav{color:#a855f7 !important}
.nav-btn.tc-nav:hover{background:var(--surf2)}
.nav-btn.tc-nav.active{color:#a855f7;border-bottom-color:#a855f7}
.page{max-width:1500px;margin:0 auto;padding:36px 48px}
.tab{display:none}.tab.on{display:block}
.sec-title{font-size:18px;font-weight:700;color:#fff;margin-bottom:6px;display:flex;align-items:center;gap:10px}
.sec-desc{font-size:13px;color:var(--muted);margin-bottom:28px}
.section-sub-title{font-size:13px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;margin:24px 0 12px}
/* Notice bars */
.notice-bar{background:#6366f122;border:1px solid #6366f144;border-left:4px solid #6366f1;border-radius:var(--r);padding:10px 16px;margin-bottom:14px;font-size:12.5px;color:#a5b4fc}
.tc-notice{background:#a855f722;border-color:#a855f744;border-left-color:#a855f7;color:#d8b4fe}
/* TC section header */
.tc-section-hdr{background:linear-gradient(90deg,#a855f722,transparent);border:1px solid #a855f733;border-radius:var(--r2);padding:16px 20px;margin:28px 0 16px;display:flex;align-items:center;gap:12px}
.tc-section-hdr .icon{font-size:20px}
.tc-section-hdr .title{font-size:14px;font-weight:700;color:#c084fc}
.tc-section-hdr .sub{font-size:12px;color:var(--muted);margin-top:2px}
.time-card{background:var(--surf);border:1px solid var(--border2);border-radius:var(--r2);padding:20px 28px;margin-bottom:22px;border-left:4px solid var(--teal)}
.time-row{display:flex;align-items:center;gap:24px;flex-wrap:wrap}
.time-item{flex:1;min-width:180px}
.time-lbl{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.7px;font-weight:600;margin-bottom:5px}
.time-val{font-size:16px;font-weight:700;color:#fff;margin-bottom:3px}
.time-sub{font-size:11px;color:var(--dim)}
.time-sep{font-size:22px;color:var(--teal);font-weight:700;flex-shrink:0}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(175px,1fr));gap:14px;margin-bottom:28px}
.kpi-card{background:var(--surf);border:1px solid var(--border);border-radius:var(--r2);padding:20px 22px;transition:border-color .2s,transform .15s;cursor:default}
.kpi-card:hover{border-color:var(--accent);transform:translateY(-2px)}
.kpi-lbl{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.7px;font-weight:600}
.kpi-val{font-size:28px;font-weight:800;color:#fff;margin:7px 0 3px;line-height:1}
.kpi-sub{font-size:11px;color:var(--muted)}
.chart-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:22px}
.cbox{background:var(--surf);border:1px solid var(--border);border-radius:var(--r2);padding:22px}
.cbox.span2{grid-column:1/-1}
.cbox-title{font-size:11.5px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;margin-bottom:18px}
.cwrap{position:relative;height:268px}
.cwrap.tall{height:320px}
.tbl-wrap{overflow-x:auto;border-radius:var(--r2);border:1px solid var(--border);margin-bottom:8px}
table{width:100%;border-collapse:collapse}
thead tr{background:var(--surf3)}
thead th{padding:11px 14px;text-align:left;font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid var(--border2);white-space:nowrap}
tbody tr{border-bottom:1px solid var(--border);transition:background .12s}
tbody tr:last-child{border-bottom:none}
tbody tr:hover{background:var(--surf2)}
td{padding:11px 14px;vertical-align:middle}
.tc{font-weight:600;color:#fff;max-width:240px;word-break:break-word}
.hi{color:var(--teal);font-weight:700}
.row-danger{border-left:3px solid var(--red)}
.row-warn{border-left:3px solid var(--orange)}
.badge{display:inline-block;padding:2px 9px;border-radius:999px;font-size:11px;font-weight:700;white-space:nowrap}
.exec-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:16px;margin-bottom:28px}
.exec-card{background:var(--surf);border:1px solid var(--border);border-radius:var(--r2);padding:22px 24px;position:relative;overflow:hidden}
.exec-card::before{content:"";position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,var(--accent),var(--teal))}
.exec-card.good::before{background:linear-gradient(90deg,#22c55e,#84cc16)}
.exec-card.warn::before{background:linear-gradient(90deg,#f97316,#eab308)}
.exec-card.bad::before{background:linear-gradient(90deg,#ef4444,#f97316)}
.exec-label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.7px;font-weight:600}
.exec-val{font-size:32px;font-weight:800;margin:8px 0 4px;line-height:1}
.exec-sub{font-size:12px;color:var(--muted)}
.apdex-hero{background:var(--surf);border:1px solid var(--border);border-radius:var(--r2);padding:32px 36px;margin-bottom:28px;display:flex;align-items:center;gap:40px;flex-wrap:wrap}
.dial-wrap{position:relative;width:170px;height:110px;flex-shrink:0}
.dial-wrap svg{width:170px;height:110px}
.dial-center{position:absolute;bottom:0;left:50%;transform:translateX(-50%);text-align:center;width:140px}
.dial-score{font-size:34px;font-weight:900;line-height:1}
.dial-label{font-size:11px;color:var(--muted);letter-spacing:.5px;margin-top:2px}
.apdex-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;flex:1;min-width:280px}
.as-card{background:var(--surf2);border:1px solid var(--border);border-radius:var(--r);padding:16px;text-align:center}
.as-val{font-size:24px;font-weight:800;margin-bottom:3px}
.as-lbl{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}
.as-sub{font-size:11px;color:var(--dim);margin-top:2px}
.formula-box{background:var(--surf2);border:1px solid var(--border);border-radius:var(--r);padding:16px 20px;margin-bottom:24px;font-size:13px;color:var(--muted);line-height:2}
.formula-box b{color:var(--text)}
.run-legend{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:24px}
.run-pill{display:flex;align-items:center;gap:7px;background:var(--surf);border:1px solid var(--border);border-radius:999px;padding:6px 14px;font-size:12px;font-weight:600;color:var(--text)}
.run-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
.divider{border:none;border-top:1px solid var(--border);margin:32px 0}
.footer{text-align:center;padding:28px;color:var(--dim);font-size:12px;border-top:1px solid var(--border);margin-top:48px}
.footer b{color:var(--muted)}
@media(max-width:960px){
  .chart-grid,.apdex-stats{grid-template-columns:1fr}
  .page,.header-inner,.nav-inner{padding-left:20px;padding-right:20px}
  .apdex-hero,.time-row{flex-direction:column;align-items:flex-start}
  .time-sep{display:none}
}
</style>
</head>
<body>
""")

    # ── HEADER ──
    run_pills = ""
    for i, (lb, run) in enumerate(zip(labels, runs)):
        col = PALETTE[i % len(PALETTE)]
        run_pills += (
            '<div class="pill">'
            '<span style="width:8px;height:8px;border-radius:50%;background:' + col + ';display:inline-block"></span>'
            ' <b>' + lb + '</b> &nbsp; '
            + '{:,}'.format(run["total"]) + ' API samples'
            + (' &nbsp;+&nbsp; ' + str(run["tc_total"]) + ' TC' if run["tc_total"] else '')
            + '</div>'
        )
    p.append('<div class="header"><div class="header-inner">')
    p.append('<div class="logo-block"><h1>&#9889; Performance <em>Test Report</em></h1>'
             '<div class="sub">JMeter Analysis &nbsp;&#8226;&nbsp; Generated ' + gen
             + ' &nbsp;&#8226;&nbsp; TC rows excluded from all metrics</div></div>')
    p.append(
        '<div class="meta-pills">'
        + run_pills +
        '<button id="dl-btn" class="dl-btn" onclick="downloadDocx()">'
        '&#128196; Download Word Doc'
        '</button>'
        '</div>'
    )
    p.append('</div></div>')

    # ── NAV ──
    cmp_btn = '<button class="nav-btn" onclick="ST(\'cmp\',this)">&#128260; Comparison</button>' if multi else ""
    p.append(
        '<div class="nav-bar"><div class="nav-inner">'
        '<button class="nav-btn active" onclick="ST(\'exec\',this)">&#128203; Executive Summary</button>'
        '<button class="nav-btn" onclick="ST(\'resp\',this)">&#128202; Response Time</button>'
        '<button class="nav-btn" onclick="ST(\'tput\',this)">&#128640; Throughput</button>'
        '<button class="nav-btn" onclick="ST(\'pct\',this)">&#128208; Percentiles</button>'
        '<button class="nav-btn" onclick="ST(\'apdex\',this)">&#127919; APDEX</button>'
        '<button class="nav-btn" onclick="ST(\'err\',this)">&#128293; Error Analysis</button>'
        '<button class="nav-btn" onclick="ST(\'why\',this)">&#129322; Error Reasons</button>'
        '<button class="nav-btn tc-nav" onclick="ST(\'tc\',this)">&#128260; TC View (Info Only)</button>'
        + cmp_btn +
        '</div></div>'
    )
    p.append('<div class="page">')

    # ══════════════════════════════════
    # TAB: EXECUTIVE SUMMARY
    # ══════════════════════════════════
    p.append('<div id="tab-exec" class="tab on">')
    p.append('<div class="sec-title">&#128203; Executive Summary</div>')
    p.append('<div class="sec-desc">High-level performance overview &nbsp;&#8226;&nbsp; <b>All metrics calculated from API requests only</b> — TC rows excluded</div>')

    if multi:
        p.append('<div class="section-sub-title">Run Overview</div>')
        p.append(exec_comparison_table(runs, labels))
        p.append('<hr class="divider"/>')

    for i, (lb, run) in enumerate(zip(labels, runs)):
        col = PALETTE[i % len(PALETTE)]
        ap  = run["apdex"]
        ec  = score_color(run["err_pct"])
        if multi:
            p.append('<div style="font-size:13px;font-weight:700;color:' + col
                     + ';margin:20px 0 12px;display:flex;align-items:center;gap:8px">'
                     '<span style="width:12px;height:12px;border-radius:50%;background:' + col
                     + ';display:inline-block"></span> ' + lb + '</div>')
        p.append(time_info_card(run))
        overall_st = "PASS" if run["err_pct"] < 1 else ("WARN" if run["err_pct"] < 5 else "FAIL")
        sc_class   = "good" if overall_st == "PASS" else ("warn" if overall_st == "WARN" else "bad")
        sc_col     = "#22c55e" if overall_st == "PASS" else ("#f97316" if overall_st == "WARN" else "#ef4444")
        sc_bg      = "#22c55e22" if overall_st == "PASS" else ("#f9731622" if overall_st == "WARN" else "#ef444422")
        p.append('<div class="exec-grid">')
        p.append('<div class="exec-card ' + sc_class + '">'
                 '<div class="exec-label">Overall Status</div>'
                 '<div class="exec-val" style="color:' + sc_col + '">' + overall_st + '</div>'
                 '<div class="exec-sub">'
                 + '{:,}'.format(run["total"]) + ' API samples &nbsp;&#8226;&nbsp; ' + run["dur_str"]
                 + '</div>'
                 '<div style="margin-top:10px;font-size:11px;font-weight:700;padding:3px 10px;border-radius:999px;display:inline-block;background:' + sc_bg + ';color:' + sc_col + '">'
                 + ap["rating"] + ' APDEX ' + str(ap["score"]) + '</div>'
                 '</div>')
        p.append('<div class="exec-card"><div class="exec-label">Avg Response Time</div>'
                 '<div class="exec-val">' + str(run["avg"]) + '<span style="font-size:16px;color:var(--muted)"> ms</span></div>'
                 '<div class="exec-sub">Min ' + str(run["min"]) + 'ms &nbsp;/&nbsp; Max ' + str(run["max"]) + 'ms</div></div>')
        p.append('<div class="exec-card"><div class="exec-label">Throughput</div>'
                 '<div class="exec-val" style="color:var(--green)">' + str(run["tps"]) + '<span style="font-size:16px;color:var(--muted)"> TPS</span></div>'
                 '<div class="exec-sub">Peak: ' + str(run["peak_tps"]) + ' req/sec</div></div>')
        p.append('<div class="exec-card"><div class="exec-label">Error Rate</div>'
                 '<div class="exec-val" style="color:' + ec + '">' + str(run["err_pct"]) + '<span style="font-size:16px;color:var(--muted)">%</span></div>'
                 '<div class="exec-sub">' + '{:,}'.format(run["errors"]) + ' errors / ' + '{:,}'.format(run["total"]) + ' API requests</div></div>')
        p.append('<div class="exec-card"><div class="exec-label">P90 Response</div>'
                 '<div class="exec-val" style="color:var(--teal)">' + str(run["p90"]) + '<span style="font-size:16px;color:var(--muted)"> ms</span></div>'
                 '<div class="exec-sub">P95: ' + str(run["p95"]) + 'ms &nbsp;/&nbsp; P99: ' + str(run["p99"]) + 'ms</div></div>')
        p.append('<div class="exec-card"><div class="exec-label">Test Duration</div>'
                 '<div class="exec-val" style="color:#22d3ee">' + run["dur_str"] + '</div>'
                 '<div class="exec-sub">' + str(run["dur"]) + ' seconds</div></div>')
        p.append('</div>')
        if i < len(runs) - 1: p.append('<hr class="divider"/>')

    p.append('<hr class="divider"/>')
    p.append('<div class="chart-grid"><div class="cbox span2">'
             '<div class="cbox-title">Average Response Time by API (ms) — TC rows excluded</div>'
             '<div class="cwrap"><canvas id="ec_avg"></canvas></div></div></div>')
    p.append('</div>')  # end tab-exec

    # ══════════════════════════════════
    # TAB: RESPONSE TIME
    # ══════════════════════════════════
    p.append('<div id="tab-resp" class="tab">')
    p.append('<div class="sec-title">&#128202; Response Time Summary</div>')
    p.append('<div class="sec-desc">Detailed latency breakdown per API &nbsp;&#8226;&nbsp; TC rows shown separately in the TC View tab</div>')
    for i, (lb, run) in enumerate(zip(labels, runs)):
        col = PALETTE[i % len(PALETTE)]
        if multi: p.append('<div class="section-sub-title" style="color:' + col + '">&#9679; ' + lb + '</div>')
        p.append(time_info_card(run))
        p.append('<div class="kpi-grid">')
        p.append(kpi("API Samples",  '{:,}'.format(run["total"]),    "TC rows excluded",          "#6366f1"))
        p.append(kpi("Avg Response", str(run["avg"]) + " ms",        "mean latency"))
        p.append(kpi("Min / Max",    str(run["min"]) + " / " + str(run["max"]), "ms"))
        p.append(kpi("Error Rate",   str(run["err_pct"]) + "%",      '{:,}'.format(run["errors"]) + " failed", score_color(run["err_pct"])))
        p.append(kpi("Throughput",   str(run["tps"]) + " TPS",       "req/sec",                  "#22c55e"))
        p.append(kpi("Duration",     run["dur_str"],                  str(run["dur"]) + "s",      "#22d3ee"))
        p.append('</div>')
        p.append(api_summary_table(run))
        if i < len(runs) - 1: p.append('<hr class="divider"/>')
    p.append('</div>')

    # ══════════════════════════════════
    # TAB: THROUGHPUT
    # ══════════════════════════════════
    p.append('<div id="tab-tput" class="tab">')
    p.append('<div class="sec-title">&#128640; Throughput / Requests Per Second</div>')
    p.append('<div class="sec-desc">Request volume and error rate over time &nbsp;&#8226;&nbsp; API rows only</div>')
    for i, (lb, run) in enumerate(zip(labels, runs)):
        col = PALETTE[i % len(PALETTE)]
        if multi: p.append('<div class="section-sub-title" style="color:' + col + '">&#9679; ' + lb + '</div>')
        p.append(time_info_card(run))
        p.append('<div class="kpi-grid">')
        p.append(kpi("Peak TPS",     run["peak_tps"],                 "max in 1-sec window",       "#22c55e"))
        p.append(kpi("Avg TPS",      run["tps"],                      "overall",                   "#6366f1"))
        p.append(kpi("Total Errors", '{:,}'.format(run["errors"]),    "failed API requests",       "#ef4444"))
        p.append(kpi("Error Rate",   str(run["err_pct"]) + "%",       "overall",                   score_color(run["err_pct"])))
        p.append('</div>')
        p.append('<div class="chart-grid">')
        p.append('<div class="cbox span2"><div class="cbox-title">Requests &amp; Errors Per Second Over Time</div>'
                 '<div class="cwrap tall"><canvas id="tc_tps' + str(i) + '"></canvas></div></div>')
        p.append('<div class="cbox span2"><div class="cbox-title">Throughput per API (TPS)</div>'
                 '<div class="cwrap"><canvas id="tc_lbl' + str(i) + '"></canvas></div></div>')
        p.append('</div>')
        if i < len(runs) - 1: p.append('<hr class="divider"/>')
    p.append('</div>')

    # ══════════════════════════════════
    # TAB: PERCENTILES
    # ══════════════════════════════════
    p.append('<div id="tab-pct" class="tab">')
    p.append('<div class="sec-title">&#128208; Percentile Response Times</div>')
    p.append('<div class="sec-desc">P90 / P95 / P99 latency distribution &nbsp;&#8226;&nbsp; API rows only</div>')
    for i, (lb, run) in enumerate(zip(labels, runs)):
        col = PALETTE[i % len(PALETTE)]
        if multi: p.append('<div class="section-sub-title" style="color:' + col + '">&#9679; ' + lb + '</div>')
        p.append('<div class="kpi-grid">')
        p.append(kpi("P90", str(run["p90"]) + " ms", "90th percentile", "#6366f1"))
        p.append(kpi("P95", str(run["p95"]) + " ms", "95th percentile", "#f97316"))
        p.append(kpi("P99", str(run["p99"]) + " ms", "99th percentile", "#ef4444"))
        p.append(kpi("Max", str(run["max"]) + " ms", "worst response"))
        p.append('</div>')
        p.append('<div class="chart-grid"><div class="cbox span2">'
                 '<div class="cbox-title">P90 / P95 / P99 per API</div>'
                 '<div class="cwrap"><canvas id="pc_bar' + str(i) + '"></canvas></div></div></div>')
        p.append(pct_table(run))
        if i < len(runs) - 1: p.append('<hr class="divider"/>')
    p.append('</div>')

    # ══════════════════════════════════
    # TAB: APDEX
    # ══════════════════════════════════
    p.append('<div id="tab-apdex" class="tab">')
    p.append('<div class="sec-title">&#127919; APDEX — Application Performance Index</div>')
    p.append('<div class="sec-desc">Standardised user satisfaction score (0.0 to 1.0) &nbsp;&#8226;&nbsp; API rows only</div>')
    p.append('<div class="formula-box">'
             '<b>Formula:</b> &nbsp; Score = (Satisfied + Tolerating / 2) / Total &nbsp;&nbsp;'
             '&#8226;&nbsp; <b>Threshold T = ' + str(T) + ' ms</b><br/>'
             '<span style="color:#22c55e;font-weight:700">Satisfied</span> &le; T ms &nbsp; | &nbsp;'
             '<span style="color:#eab308;font-weight:700">Tolerating</span> &le; ' + str(T * 4) + ' ms &nbsp; | &nbsp;'
             '<span style="color:#ef4444;font-weight:700">Frustrated</span> &gt; ' + str(T * 4) + ' ms &nbsp;&nbsp;&#8226;&nbsp;&nbsp;'
             '<b>Ratings:</b> &nbsp;'
             '<span style="color:#22c55e">Excellent &ge;0.94</span> &nbsp;'
             '<span style="color:#84cc16">Good &ge;0.85</span> &nbsp;'
             '<span style="color:#eab308">Fair &ge;0.70</span> &nbsp;'
             '<span style="color:#f97316">Poor &ge;0.50</span> &nbsp;'
             '<span style="color:#ef4444">Unacceptable &lt;0.50</span>'
             '</div>')
    for i, (lb, run) in enumerate(zip(labels, runs)):
        col  = PALETTE[i % len(PALETTE)]
        ap   = run["apdex"]
        if multi: p.append('<div class="section-sub-title" style="color:' + col + '">&#9679; ' + lb + '</div>')
        p.append(time_info_card(run))
        p.append('<div class="apdex-hero">')
        p.append('<div class="dial-wrap">'
                 '<svg viewBox="0 0 170 110" xmlns="http://www.w3.org/2000/svg">'
                 '<path d="M20,95 A65,65 0 0,1 150,95" fill="none" stroke="#1d2540" stroke-width="13" stroke-linecap="round"/>'
                 '<path d="M20,95 A65,65 0 0,1 150,95" fill="none" stroke="' + ap["color"] + '" stroke-width="13" stroke-linecap="round"'
                 ' stroke-dasharray="' + str(int(204 * ap["score"])) + ' 204"/>'
                 '</svg>'
                 '<div class="dial-center">'
                 '<div class="dial-score" style="color:' + ap["color"] + '">' + str(ap["score"]) + '</div>'
                 '<div class="dial-label">APDEX</div>'
                 '</div></div>')
        p.append('<div style="flex:1">')
        p.append('<div style="font-size:26px;font-weight:800;color:' + ap["color"] + ';margin-bottom:4px">' + ap["rating"] + '</div>')
        p.append('<div style="color:var(--muted);font-size:13px;margin-bottom:20px">'
                 'Score ' + str(ap["score"]) + ' &nbsp;&#8226;&nbsp; T=' + str(T) + 'ms &nbsp;&#8226;&nbsp; '
                 + '{:,}'.format(ap["n"]) + ' API samples</div>')
        p.append('<div class="apdex-stats">')
        p.append('<div class="as-card"><div class="as-val" style="color:#22c55e">' + '{:,}'.format(ap["sat"]) + '</div>'
                 '<div class="as-lbl">Satisfied</div><div class="as-sub">' + str(ap["sat_p"]) + '% &le;' + str(T) + 'ms</div></div>')
        p.append('<div class="as-card"><div class="as-val" style="color:#eab308">' + '{:,}'.format(ap["tol"]) + '</div>'
                 '<div class="as-lbl">Tolerating</div><div class="as-sub">' + str(ap["tol_p"]) + '% &le;' + str(T * 4) + 'ms</div></div>')
        p.append('<div class="as-card"><div class="as-val" style="color:#ef4444">' + '{:,}'.format(ap["fru"]) + '</div>'
                 '<div class="as-lbl">Frustrated</div><div class="as-sub">' + str(ap["fru_p"]) + '% &gt;' + str(T * 4) + 'ms</div></div>')
        p.append('</div></div></div>')
        p.append('<div class="chart-grid"><div class="cbox span2">'
                 '<div class="cbox-title">APDEX Score per API</div>'
                 '<div class="cwrap"><canvas id="ap_bar' + str(i) + '"></canvas></div></div></div>')
        p.append(apdex_table(run))
        if i < len(runs) - 1: p.append('<hr class="divider"/>')
    p.append('</div>')

    # ══════════════════════════════════
    # TAB: ERROR ANALYSIS
    # ══════════════════════════════════
    p.append('<div id="tab-err" class="tab">')
    p.append('<div class="sec-title">&#128293; Error Analysis</div>')
    p.append('<div class="sec-desc">HTTP response code distribution and per-API error rates &nbsp;&#8226;&nbsp; API rows only</div>')
    for i, (lb, run) in enumerate(zip(labels, runs)):
        col = PALETTE[i % len(PALETTE)]
        if multi: p.append('<div class="section-sub-title" style="color:' + col + '">&#9679; ' + lb + '</div>')
        p.append('<div class="kpi-grid">')
        p.append(kpi("Total Errors",     '{:,}'.format(run["errors"]),        "failed API requests",    "#ef4444"))
        p.append(kpi("Error Rate",       str(run["err_pct"]) + "%",           "overall",                score_color(run["err_pct"])))
        p.append(kpi("Distinct Reasons", str(len(run["reasons"])),            "unique failure types",   "#f97316"))
        p.append(kpi("Success Rate",     str(round(100 - run["err_pct"], 2)) + "%", "passed",           "#22c55e"))
        p.append('</div>')
        p.append('<div class="chart-grid">'
                 '<div class="cbox"><div class="cbox-title">Response Code Distribution</div>'
                 '<div class="cwrap"><canvas id="er_code' + str(i) + '"></canvas></div></div>'
                 '<div class="cbox"><div class="cbox-title">Error % per API</div>'
                 '<div class="cwrap"><canvas id="er_pct' + str(i) + '"></canvas></div></div>'
                 '</div>')
        if i < len(runs) - 1: p.append('<hr class="divider"/>')
    p.append('</div>')

    # ══════════════════════════════════
    # TAB: ERROR REASONS
    # ══════════════════════════════════
    p.append('<div id="tab-why" class="tab">')
    p.append('<div class="sec-title">&#129322; Error Reasons</div>')
    p.append('<div class="sec-desc">Full failure messages, assertion errors, and HTTP status descriptions &nbsp;&#8226;&nbsp; API rows only</div>')
    for i, (lb, run) in enumerate(zip(labels, runs)):
        col = PALETTE[i % len(PALETTE)]
        if multi: p.append('<div class="section-sub-title" style="color:' + col + '">&#9679; ' + lb + '</div>')
        p.append(reasons_table(run))
        if i < len(runs) - 1: p.append('<hr class="divider"/>')
    p.append('</div>')

    # ══════════════════════════════════
    # TAB: TC VIEW (Info Only)
    # ══════════════════════════════════
    p.append('<div id="tab-tc" class="tab">')
    p.append('<div class="sec-title" style="color:#c084fc">&#128260; Transaction Controller View</div>')
    p.append('<div class="sec-desc">'
             'TC rows are <b>displayed here for reference only</b>. '
             'They are <b>completely excluded</b> from all KPIs, sample counts, error rates, APDEX, and charts on all other tabs.'
             '</div>')
    p.append('<div class="notice-bar tc-notice" style="margin-bottom:24px;font-size:13px">'
             '&#9888;&nbsp; <b>Why TC rows are excluded:</b> Transaction Controllers in JMeter aggregate the API calls within them. '
             'Including TC rows in sample counts would <b>double-count</b> requests and skew all metrics. '
             'The individual API rows underneath each TC already contain the real data.'
             '</div>')
    for i, (lb, run) in enumerate(zip(labels, runs)):
        col = PALETTE[i % len(PALETTE)]
        if multi: p.append('<div class="section-sub-title" style="color:' + col + '">&#9679; ' + lb + '</div>')
        p.append('<div class="tc-section-hdr">'
                 '<div class="icon">&#128260;</div>'
                 '<div><div class="title">Transaction Controllers — ' + lb + '</div>'
                 '<div class="sub">'
                 + str(len(run["per_tc"])) + ' TCs &nbsp;&#8226;&nbsp; '
                 + '{:,}'.format(run["tc_total"]) + ' TC-level rows in JTL &nbsp;&#8226;&nbsp; View only, not counted in any metric'
                 + '</div></div>'
                 '</div>')
        p.append(tc_summary_table(run))
        if i < len(runs) - 1: p.append('<hr class="divider"/>')
    p.append('</div>')

    # ══════════════════════════════════
    # TAB: COMPARISON (multi only)
    # ══════════════════════════════════
    if multi:
        p.append('<div id="tab-cmp" class="tab">')
        p.append('<div class="sec-title">&#128260; Run Comparison</div>')
        p.append('<div class="sec-desc">Side-by-side metrics across all test runs &nbsp;&#8226;&nbsp; API rows only</div>')
        p.append('<div class="run-legend">')
        for i, lb in enumerate(labels):
            col = PALETTE[i % len(PALETTE)]
            p.append('<div class="run-pill"><span class="run-dot" style="background:' + col + '"></span>' + lb + '</div>')
        p.append('</div>')
        p.append(exec_comparison_table(runs, labels))
        p.append('<div class="chart-grid">'
                 '<div class="cbox span2"><div class="cbox-title">Avg Response Time Comparison (ms)</div>'
                 '<div class="cwrap tall"><canvas id="cmp_avg"></canvas></div></div>'
                 '<div class="cbox span2"><div class="cbox-title">P90 / P95 / P99 Comparison (ms)</div>'
                 '<div class="cwrap tall"><canvas id="cmp_pct"></canvas></div></div>'
                 '<div class="cbox"><div class="cbox-title">Throughput Comparison (TPS)</div>'
                 '<div class="cwrap"><canvas id="cmp_tps"></canvas></div></div>'
                 '<div class="cbox"><div class="cbox-title">Error Rate Comparison (%)</div>'
                 '<div class="cwrap"><canvas id="cmp_err"></canvas></div></div>'
                 '<div class="cbox span2"><div class="cbox-title">APDEX Score Comparison</div>'
                 '<div class="cwrap"><canvas id="cmp_apdex"></canvas></div></div>'
                 '</div>')
        p.append('<hr class="divider"/>')
        p.append(comparison_tables(runs, labels))
        p.append('</div>')

    p.append('</div>')  # end .page

    # ── FOOTER ──
    src_list = " &nbsp;&#8226;&nbsp; ".join(os.path.basename(path) for path in jtl_paths)
    p.append('<div class="footer">&#9889; <b>JMeter Performance Report v3.0</b>'
             ' &nbsp;&#8226;&nbsp; ' + src_list +
             ' &nbsp;&#8226;&nbsp; APDEX T=' + str(T) + 'ms'
             ' &nbsp;&#8226;&nbsp; TC rows excluded from metrics'
             ' &nbsp;&#8226;&nbsp; Generated ' + gen +
             ' &nbsp;&#8226;&nbsp; &copy; ' + str(year) + '</div>')

    # ── JAVASCRIPT ──
    p.append('<script>')
    p.append("const PAL=" + json.dumps(PALETTE) + ";")
    p.append("""
Chart.defaults.color='#8896b0';
Chart.defaults.borderColor='#232d45';
Chart.defaults.font.family="'Segoe UI',system-ui,sans-serif";
Chart.defaults.font.size=12;
function ST(name,btn){
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));
  document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('tab-'+name).classList.add('on');
  btn.classList.add('active');
}
function barChart(id,labels,datasets,yLabel,max){
  const ctx=document.getElementById(id);if(!ctx)return;
  new Chart(ctx,{type:'bar',data:{labels,datasets},options:{responsive:true,maintainAspectRatio:false,
    plugins:{legend:{position:'top',labels:{boxWidth:12}}},
    scales:{x:{ticks:{maxRotation:40,font:{size:11}}},
            y:{beginAtZero:true,max:max,title:{display:!!yLabel,text:yLabel},ticks:{font:{size:11}}}}}});
}
function lineChart(id,labels,datasets){
  const ctx=document.getElementById(id);if(!ctx)return;
  new Chart(ctx,{type:'line',data:{labels,datasets},options:{responsive:true,maintainAspectRatio:false,
    plugins:{legend:{position:'top',labels:{boxWidth:12}}},
    scales:{x:{ticks:{maxTicksLimit:14,maxRotation:30,font:{size:11}}},y:{beginAtZero:true,ticks:{font:{size:11}}}}}});
}
function doughnut(id,labels,data){
  const ctx=document.getElementById(id);if(!ctx)return;
  new Chart(ctx,{type:'doughnut',data:{labels,datasets:[{data,
    backgroundColor:labels.map((_,i)=>PAL[i%PAL.length]+'cc'),
    borderColor:labels.map((_,i)=>PAL[i%PAL.length]),borderWidth:1}]},
  options:{responsive:true,maintainAspectRatio:false,cutout:'60%',
    plugins:{legend:{position:'right',labels:{boxWidth:12,font:{size:11}}},
      tooltip:{callbacks:{label:ctx=>ctx.label+': '+ctx.raw.toLocaleString()}}}}});
}
""")

    for i, run in enumerate(runs):
        lbs  = json.dumps([r["label"]          for r in run["per_label"]])
        avgs = json.dumps([r["avg"]            for r in run["per_label"]])
        tpss = json.dumps([r["tps"]            for r in run["per_label"]])
        p90s = json.dumps([r["p90"]            for r in run["per_label"]])
        p95s = json.dumps([r["p95"]            for r in run["per_label"]])
        p99s = json.dumps([r["p99"]            for r in run["per_label"]])
        eps  = json.dumps([r["err_pct"]        for r in run["per_label"]])
        aps  = json.dumps([r["apdex"]["score"] for r in run["per_label"]])
        apc  = json.dumps([r["apdex"]["color"] for r in run["per_label"]])
        tl   = json.dumps(run["tps_labels"])
        tv   = json.dumps(run["tps_vals"])
        ev   = json.dumps(run["tps_err_vals"])
        cl   = json.dumps(list(run["codes"].keys()))
        cv   = json.dumps(list(run["codes"].values()))
        col  = PALETTE[i % len(PALETTE)]
        p.append("(function(){")
        p.append("const lbs=" + lbs + ",avgs=" + avgs + ",tpss=" + tpss + ";")
        p.append("const p90=" + p90s + ",p95=" + p95s + ",p99=" + p99s + ";")
        p.append("const eps=" + eps + ",aps=" + aps + ",apc=" + apc + ";")
        p.append("const tl=" + tl + ",tv=" + tv + ",ev=" + ev + ";")
        p.append("const cl=" + cl + ",cv=" + cv + ";")
        p.append("const col='" + col + "';")
        if i == 0:
            p.append("barChart('ec_avg',lbs,[{label:'Avg (ms)',data:avgs,"
                      "backgroundColor:lbs.map((_,i)=>PAL[i%PAL.length]+'bb'),"
                      "borderColor:lbs.map((_,i)=>PAL[i%PAL.length]),borderWidth:1,borderRadius:5}],'ms');")
        p.append("lineChart('tc_tps" + str(i) + "',tl,["
                 "{label:'Req/sec',data:tv,borderColor:'#22c55e',backgroundColor:'#22c55e18',fill:true,tension:.35,pointRadius:0,borderWidth:2},"
                 "{label:'Err/sec',data:ev,borderColor:'#ef4444',backgroundColor:'#ef444418',fill:true,tension:.35,pointRadius:0,borderWidth:2}"
                 "]);")
        p.append("barChart('tc_lbl" + str(i) + "',lbs,[{label:'TPS',data:tpss,"
                 "backgroundColor:col+'99',borderColor:col,borderWidth:1,borderRadius:5}],'req/sec');")
        p.append("barChart('pc_bar" + str(i) + "',lbs,["
                 "{label:'P90',data:p90,backgroundColor:'#6366f1bb',borderColor:'#6366f1',borderWidth:1,borderRadius:3},"
                 "{label:'P95',data:p95,backgroundColor:'#f97316bb',borderColor:'#f97316',borderWidth:1,borderRadius:3},"
                 "{label:'P99',data:p99,backgroundColor:'#ef4444bb',borderColor:'#ef4444',borderWidth:1,borderRadius:3}"
                 "],'ms');")
        p.append("barChart('ap_bar" + str(i) + "',lbs,[{label:'APDEX',data:aps,"
                 "backgroundColor:apc.map(c=>c+'bb'),borderColor:apc,borderWidth:1,borderRadius:5}],null,1);")
        p.append("doughnut('er_code" + str(i) + "',cl,cv);")
        p.append("barChart('er_pct" + str(i) + "',lbs,[{label:'Error %',data:eps,"
                 "backgroundColor:eps.map(v=>v>=10?'#ef444499':v>0?'#f9731699':'#22c55e44'),"
                 "borderColor:eps.map(v=>v>=10?'#ef4444':v>0?'#f97316':'#22c55e'),"
                 "borderWidth:1,borderRadius:5}],'%',100);")
        p.append("})();")

    if multi:
        all_lbs_set, seen_set = [], set()
        for run in runs:
            for r in run["per_label"]:
                if r["label"] not in seen_set:
                    all_lbs_set.append(r["label"])
                    seen_set.add(r["label"])
        def get_val(run, lb, key):
            for r in run["per_label"]:
                if r["label"] == lb: return r.get(key, 0)
            return 0
        cmp_lbs = json.dumps(all_lbs_set)
        p.append("const cmpLbs=" + cmp_lbs + ";")
        avg_ds = [{"label": lb, "data": [get_val(run, l, "avg") for l in all_lbs_set],
                   "backgroundColor": PALETTE[i % len(PALETTE)] + "99",
                   "borderColor": PALETTE[i % len(PALETTE)], "borderWidth": 1, "borderRadius": 4}
                  for i, (lb, run) in enumerate(zip(labels, runs))]
        p.append("barChart('cmp_avg',cmpLbs," + json.dumps(avg_ds) + ",'ms');")
        p95_ds = []
        for i, (lb, run) in enumerate(zip(labels, runs)):
            col = PALETTE[i % len(PALETTE)]
            p95_ds.append({"label": lb + " P90", "data": [get_val(run, l, "p90") for l in all_lbs_set],
                           "backgroundColor": col + "66", "borderColor": col, "borderWidth": 1, "borderRadius": 3})
            p95_ds.append({"label": lb + " P95", "data": [get_val(run, l, "p95") for l in all_lbs_set],
                           "backgroundColor": col + "99", "borderColor": col, "borderWidth": 1, "borderRadius": 3})
        p.append("barChart('cmp_pct',cmpLbs," + json.dumps(p95_ds) + ",'ms');")
        tps_ds = [{"label": lb, "data": [get_val(run, l, "tps") for l in all_lbs_set],
                   "backgroundColor": PALETTE[i % len(PALETTE)] + "99",
                   "borderColor": PALETTE[i % len(PALETTE)], "borderWidth": 1, "borderRadius": 4}
                  for i, (lb, run) in enumerate(zip(labels, runs))]
        p.append("barChart('cmp_tps',cmpLbs," + json.dumps(tps_ds) + ",'TPS');")
        err_ds = [{"label": lb, "data": [get_val(run, l, "err_pct") for l in all_lbs_set],
                   "backgroundColor": PALETTE[i % len(PALETTE)] + "99",
                   "borderColor": PALETTE[i % len(PALETTE)], "borderWidth": 1, "borderRadius": 4}
                  for i, (lb, run) in enumerate(zip(labels, runs))]
        p.append("barChart('cmp_err',cmpLbs," + json.dumps(err_ds) + ",'%',100);")
        ap_ds = [{"label": lb,
                  "data": [{r["label"]: r["apdex"]["score"] for r in run["per_label"]}.get(l, 0) for l in all_lbs_set],
                  "backgroundColor": PALETTE[i % len(PALETTE)] + "99",
                  "borderColor": PALETTE[i % len(PALETTE)], "borderWidth": 1, "borderRadius": 4}
                 for i, (lb, run) in enumerate(zip(labels, runs))]
        p.append("barChart('cmp_apdex',cmpLbs," + json.dumps(ap_ds) + ",null,1);")

    p.append('</script>')

    # ══ HIDDEN JSON DATA BLOCK — used by downloadDocx() ══
    report_data = {
        "meta": {
            "project":   "Performance Test",
            "version":   "v1.0",
            "team":      "QA Performance Team",
            "gen_time":  gen,
            "apdex_t":   T,
            "jtl_files": [os.path.basename(f) for f in jtl_paths],
        },
        "labels": labels,
        "runs":   runs,
    }
    p.append('<script id="report-data" type="application/json">')
    p.append(json.dumps(report_data, ensure_ascii=False))
    p.append('</script>')

    # ══ DOWNLOAD DOCX JAVASCRIPT ══
    p.append('<script>')
    p.append(DOCX_JS)
    p.append('</script>')

    p.append('</body></html>')
    return "".join(p)

# ══════════════════════════════════════════════════════════
# 7. MAIN
# ══════════════════════════════════════════════════════════
def main():
    parser = argparse.ArgumentParser(description="JMeter HTML Report Generator v3.0 — TC rows excluded from metrics")
    parser.add_argument("--jtl",     nargs="+", required=True,  help="One or more JTL files")
    parser.add_argument("--labels",  nargs="+",                 help="Run labels (must match JTL count)")
    parser.add_argument("--output",  default="performance_report.html")
    parser.add_argument("--apdex-t", type=int, default=500,     help="APDEX threshold in ms (default 500)")
    args = parser.parse_args()

    jtl_paths = args.jtl
    labels    = args.labels or [os.path.splitext(os.path.basename(f))[0] for f in jtl_paths]
    if len(labels) != len(jtl_paths):
        sys.exit("[ERROR] --labels count must match --jtl count")

    runs = []
    for path in jtl_paths:
        rows = parse_jtl(path)
        runs.append(compute(rows, T=args.apdex_t))

    html = render(runs, labels, jtl_paths)
    with open(args.output, "w", encoding="utf-8") as f:
        f.write(html)

    print("\n[OK] Report saved -> " + args.output)
    for lb, run in zip(labels, runs):
        print("  [{}]  {} API samples  ({} TC rows excluded)  |  APDEX: {} ({})  |  Errors: {}%".format(
            lb, run["total"], run["tc_total"],
            run["apdex"]["score"], run["apdex"]["rating"], run["err_pct"]))

