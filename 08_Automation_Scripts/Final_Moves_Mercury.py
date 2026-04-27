# -*- coding: utf-8 -*-
"""
Final_Moves_Mercury.py
=====================
Executive final moves report for MERCURY / VOYAGE 1.

Rule requested:
- Disch_MERCURY / VOYAGE 1 is displayed as DISCHARGE.
- Gate-in / Baplie actual load is displayed as LOAD.
"""

from __future__ import annotations

import base64
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path


LOCAL_DIR = Path(__file__).resolve().parent
APP_DIR = LOCAL_DIR.parent
VESSEL_DIR = Path(r"C:\capacity_planner\booking_reports\Vessel\MERCURY / VOYAGE 1")
CLOUD_DIR = Path(r"C:\Users\Portfolio_User\Portfolio Workspace\demo service SERVICE - Database\Arquivos_HTML")

OUTPUT_NAME = "Report_Final_Moves_Mercury.html"
OUTPUT = VESSEL_DIR / OUTPUT_NAME
APP_OUTPUT = LOCAL_DIR / OUTPUT_NAME
CLOUD_OUTPUT = CLOUD_DIR / OUTPUT_NAME

sys.path.insert(0, str(LOCAL_DIR))
sys.path.insert(0, str(VESSEL_DIR))
import Comparison_Dashboard_Mercury as comp  # noqa: E402
import Comparison_Disch_Load_Mercury as disch_mod  # noqa: E402


CAP_UNITS = comp.CAP_UNITS
CAP_TEUS = comp.CAP_TEUS
CAP_TONS = comp.CAP_TONS
CAP_PLUGS = comp.CAP_RE
FINAL_POLS = comp.FINAL_POLS
CNTR_COLS = comp.CNTR_COLS
DISPLAY_COLS = comp.DISPLAY_COLS


def fmt(value):
    try:
        return f"{int(float(value)):,}".replace(",", ".")
    except Exception:
        return "0"


def pct(value, limit):
    if not limit:
        return 0.0
    return min(float(value) * 100 / float(limit), 100.0)


def diff_class(discharge, load):
    if load < discharge:
        return "short"
    if load > discharge:
        return "extra"
    return "match"


def load_logo_b64():
    for path in [
        VESSEL_DIR / "logo_experience.png",
        VESSEL_DIR / "logo_experience.png",
        APP_DIR / "assets" / "logo_experience.png",
        APP_DIR / "logo_experience.png",
        APP_DIR / "logo_experience.png",
    ]:
        if path.exists():
            return "../assets/logo_experience.png"
    return ""


def active_columns(discharge_totals, load_totals):
    cols = [c for c in CNTR_COLS if discharge_totals.get(c, 0) > 0 or load_totals.get(c, 0) > 0]
    return cols or ["DC20", "HC40", "RH40"]


def totals_by_port(data_by_port):
    totals = {c: 0 for c in CNTR_COLS + ["Units", "TEUs", "Tons", "Plugs"]}
    for port in FINAL_POLS:
        data = data_by_port.get(port, {})
        for key in totals:
            totals[key] += data.get(key, 0)
    return totals


def render_rows(discharge_by_port, load_by_port, discharge_totals, load_totals, active_cols):
    empty = {c: 0 for c in CNTR_COLS + ["Units", "TEUs", "Tons", "Plugs"]}
    rows = []
    for port in FINAL_POLS:
        discharge = discharge_by_port.get(port, empty)
        load = load_by_port.get(port, empty)

        rows.append(f"<tr><td rowspan='2' class='port-cell'>{port}</td><td class='src discharge'>DISCHARGE</td>")
        for col in active_cols:
            rows.append(f"<td>{fmt(discharge.get(col, 0))}</td>")
        rows.append(
            f"<td class='metric'>{fmt(discharge.get('Units', 0))}</td>"
            f"<td class='metric'>{fmt(discharge.get('TEUs', 0))}</td>"
            f"<td class='metric'>{fmt(discharge.get('Tons', 0))}</td>"
            f"<td class='metric'>{fmt(discharge.get('Plugs', 0))}</td></tr>"
        )

        rows.append("<tr class='load-row'><td class='src load'>LOAD</td>")
        for col in active_cols:
            rows.append(f"<td class='{diff_class(discharge.get(col, 0), load.get(col, 0))}'>{fmt(load.get(col, 0))}</td>")
        rows.append(
            f"<td class='metric {diff_class(discharge.get('Units', 0), load.get('Units', 0))}'>{fmt(load.get('Units', 0))}</td>"
            f"<td class='metric {diff_class(discharge.get('TEUs', 0), load.get('TEUs', 0))}'>{fmt(load.get('TEUs', 0))}</td>"
            f"<td class='metric {diff_class(discharge.get('Tons', 0), load.get('Tons', 0))}'>{fmt(load.get('Tons', 0))}</td>"
            f"<td class='metric {diff_class(discharge.get('Plugs', 0), load.get('Plugs', 0))}'>{fmt(load.get('Plugs', 0))}</td></tr>"
        )

    rows.append("<tr class='total-row'><td colspan='2'>TOTAL FINAL</td>")
    for col in active_cols:
        rows.append(f"<td>{fmt(discharge_totals.get(col, 0))} / {fmt(load_totals.get(col, 0))}</td>")
    rows.append(
        f"<td>{fmt(discharge_totals.get('Units', 0))} / {fmt(load_totals.get('Units', 0))}</td>"
        f"<td>{fmt(discharge_totals.get('TEUs', 0))} / {fmt(load_totals.get('TEUs', 0))}</td>"
        f"<td>{fmt(discharge_totals.get('Tons', 0))} / {fmt(load_totals.get('Tons', 0))}</td>"
        f"<td>{fmt(discharge_totals.get('Plugs', 0))} / {fmt(load_totals.get('Plugs', 0))}</td></tr>"
    )
    return "".join(rows)


def render_container_mix(discharge_by_port, load_by_port, active_cols):
    panels = []
    for port in FINAL_POLS:
        discharge = discharge_by_port.get(port, {})
        load = load_by_port.get(port, {})
        lines = []
        for col in active_cols:
            d_value = int(discharge.get(col, 0))
            l_value = int(load.get(col, 0))
            delta = l_value - d_value
            lines.append(
                f"<tr>"
                f"<td>{DISPLAY_COLS.get(col, col)}</td>"
                f"<td>{fmt(d_value)}</td>"
                f"<td>{fmt(l_value)}</td>"
                f"<td class='{diff_class(d_value, l_value)}'>{delta:+d}</td>"
                f"</tr>"
            )
        panels.append(
            f"""
            <div class="type-card">
              <div class="type-head">
                <span>{port}</span>
                <small>{fmt(discharge.get('Units', 0))} / {fmt(load.get('Units', 0))} units</small>
              </div>
              <table class="type-table">
                <thead><tr><th>Type</th><th>Discharge</th><th>Loading</th><th>Diff</th></tr></thead>
                <tbody>{''.join(lines)}</tbody>
              </table>
            </div>
            """
        )
    return "\n".join(panels)


def build_html():
    discharge_by_port = disch_mod.get_discharge_data()
    discharge_totals = totals_by_port(discharge_by_port)
    load_by_port, load_totals = comp.get_gatein_totals()
    active_cols = active_columns(discharge_totals, load_totals)
    rows = render_rows(discharge_by_port, load_by_port, discharge_totals, load_totals, active_cols)
    container_mix = render_container_mix(discharge_by_port, load_by_port, active_cols)
    generated = datetime.now().strftime("%d/%m/%Y %H:%M")
    logo = load_logo_b64()
    logo_html = f"<img src='{logo}' alt='Portfolio Logistics'>" if logo else "<span>Demo Service</span>"

    units_delta = int(load_totals.get("Units", 0) - CAP_UNITS)
    teus_delta = int(load_totals.get("TEUs", 0) - CAP_TEUS)
    tons_delta = int(load_totals.get("Tons", 0)) - int(CAP_TONS)
    plugs_delta = int(load_totals.get("Plugs", 0) - CAP_PLUGS)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Final Moves Mercury</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
<style>
:root {{
  --teal:#174A7C;
  --gold:#C59A2F;
  --ink:#12323a;
  --muted:#5f6d72;
  --panel:#f7f9f8;
  --line:#cfd8d8;
  --short:#b71c1c;
  --extra:#0d47a1;
  --match:#1b5e20;
}}
body {{ margin:0; background:#eef2f1; color:var(--ink); font-family:'Segoe UI', Arial, sans-serif; font-size:11px; }}
.page {{ max-width:1220px; margin:0 auto; padding:16px; }}
.sheet {{ background:#fff; border:1px solid #d6dddd; border-radius:8px; box-shadow:0 12px 30px rgba(1,77,78,.12); overflow:hidden; }}
.hero {{ display:flex; align-items:center; gap:14px; padding:16px 18px 12px; border-bottom:5px solid var(--gold); background:linear-gradient(90deg,#ffffff 0%,#f8fbfa 65%,#edf5f3 100%); }}
.logo-box {{ width:56px; height:48px; background:var(--teal); border:2px solid var(--gold); border-radius:8px; display:flex; align-items:center; justify-content:center; color:#fff; font-weight:900; }}
.logo-box img {{ max-height:36px; max-width:120px; object-fit:contain; }}
h1 {{ margin:0; color:var(--teal); font-size:22px; font-weight:900; letter-spacing:.2px; }}
.subtitle {{ color:var(--muted); font-size:11px; font-weight:700; margin-top:2px; }}
.pill {{ margin-left:auto; border:1px solid var(--teal); color:var(--teal); border-radius:6px; padding:8px 14px; font-weight:900; background:#fff; }}
.cards {{ display:grid; grid-template-columns:repeat(4,1fr); gap:10px; padding:12px 18px; }}
.card-stat {{ background:var(--teal); color:#fff; border-radius:6px; border-bottom:5px solid var(--gold); padding:10px; text-align:center; min-height:84px; }}
.card-stat .label {{ color:var(--gold); text-transform:uppercase; font-size:9px; font-weight:900; }}
.card-stat .value {{ font-size:23px; font-weight:900; line-height:1.2; }}
.card-stat .limit {{ font-size:10px; opacity:.9; font-weight:700; }}
.delta {{ display:inline-block; margin-left:8px; padding:3px 8px; border-radius:6px; background:#fff; border:2px solid currentColor; font-size:13px; font-weight:1000; line-height:1; box-shadow:0 2px 8px rgba(0,0,0,.18); vertical-align:middle; }}
.delta.short {{ color:#b71c1c !important; background:#fff0f0; }}
.delta.extra {{ color:#004f9e !important; background:#eef6ff; }}
.delta.match {{ color:#1b5e20 !important; background:#effaf0; }}
.short {{ color:var(--short) !important; }}
.extra {{ color:var(--extra) !important; }}
.match {{ color:var(--match) !important; }}
.section-title {{ margin:8px 18px 0; background:var(--teal); color:#fff; border-left:10px solid var(--gold); padding:9px 12px; font-weight:900; font-size:13px; letter-spacing:.5px; text-transform:uppercase; }}
.table-wrap {{ padding:10px 18px 4px; }}
table {{ width:100%; border-collapse:collapse; table-layout:fixed; text-align:center; border:2px solid #263535; }}
thead th {{ background:var(--teal); color:#fff; border:1px solid rgba(255,255,255,.18); padding:7px 4px; font-size:11px; font-weight:900; }}
thead th.gold {{ background:var(--gold); }}
tbody td {{ border:1px solid #b8c1c1; padding:6px 4px; font-size:11px; font-weight:800; color:#071f24; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
tbody tr:nth-of-type(4n+1), tbody tr:nth-of-type(4n+2) {{ background:#f7f8f7; }}
tbody tr:nth-of-type(4n+3), tbody tr:nth-of-type(4n+4) {{ background:#e9eceb; }}
.port-cell {{ background:#fff !important; color:#000; border-right:2px solid #263535; font-weight:900; }}
.src {{ border-right:2px solid #263535; font-size:10px; font-weight:900; }}
.src.discharge {{ background:#e8f0ef; color:var(--teal); }}
.src.load {{ background:#fff7dc; color:#7c5d0d; }}
.metric {{ font-weight:900; }}
.total-row td {{ background:var(--gold) !important; color:#fff !important; border-color:#8a6f1f; font-weight:900; }}
.chart-container {{ margin:14px 18px 0; padding:10px; background:#f8f9fa; border-radius:8px; border:1px solid #ddd; position:relative; height:280px; }}
.executive {{ display:grid; grid-template-columns:1fr; gap:12px; padding:12px 18px 18px; }}
.panel {{ background:var(--panel); border:1px solid #d6dddd; border-radius:8px; padding:12px; min-height:210px; }}
.panel h2 {{ color:var(--teal); font-size:14px; font-weight:900; margin:0 0 10px; text-transform:uppercase; letter-spacing:.6px; }}
.type-grid {{ display:grid; grid-template-columns:repeat(4,1fr); gap:10px; }}
.type-card {{ background:#fff; border:1px solid #d6dddd; border-radius:8px; overflow:hidden; }}
.type-head {{ display:flex; align-items:center; justify-content:space-between; gap:8px; background:var(--teal); color:#fff; padding:8px 10px; border-bottom:4px solid var(--gold); }}
.type-head span {{ font-size:14px; font-weight:1000; }}
.type-head small {{ color:#f0d88a; font-weight:900; white-space:nowrap; }}
.type-table {{ border:0; table-layout:auto; }}
.type-table th {{ background:#eef3f2; color:var(--teal); border-color:#d6dddd; padding:6px 4px; font-size:10px; }}
.type-table td {{ padding:6px 4px; font-size:11px; border-color:#e0e6e6; background:#fff; }}
.type-table td:first-child {{ color:var(--teal); font-weight:1000; text-align:left; padding-left:8px; }}
.type-table td:last-child {{ font-weight:1000; font-size:12px; }}
.note {{ padding:0 18px 16px; color:#58686c; font-size:10.5px; font-style:italic; line-height:1.5; }}
@media (max-width:1100px) {{ .type-grid {{ grid-template-columns:repeat(2,1fr); }} }}
@media (max-width:900px) {{ .cards, .executive, .type-grid {{ grid-template-columns:1fr; }} .hero {{ flex-wrap:wrap; }} .pill {{ margin-left:0; }} }}
</style>
</head>
<body>
<div class="page">
  <div class="sheet">
    <header class="hero">
      <div class="logo-box">{logo_html}</div>
      <div>
        <h1>FINAL MOVES: MERCURY / VOYAGE 1</h1>
        <div class="subtitle">Discharge = Disch_MERCURY / VOYAGE 1 | Loading = Gate-in / Baplie actual | Generated: {generated}</div>
      </div>
      <button class="pill" onclick="if (history.length > 1) {{ history.back(); }} else {{ location.href='/navios'; }}">BACK</button>
    </header>

    <div class="cards">
      <div class="card-stat"><div class="label">Units (Discharge / Loading)</div><div class="value">{fmt(discharge_totals.get('Units',0))} / {fmt(load_totals.get('Units',0))}<span class="delta {diff_class(CAP_UNITS, load_totals.get('Units',0))}">{units_delta:+d}</span></div><div class="limit">Limit Load {fmt(CAP_UNITS)}</div></div>
      <div class="card-stat"><div class="label">TEUs (Discharge / Loading)</div><div class="value">{fmt(discharge_totals.get('TEUs',0))} / {fmt(load_totals.get('TEUs',0))}<span class="delta {diff_class(CAP_TEUS, load_totals.get('TEUs',0))}">{teus_delta:+d}</span></div><div class="limit">Limit Load {fmt(CAP_TEUS)}</div></div>
      <div class="card-stat"><div class="label">Tons (Discharge / Loading)</div><div class="value">{fmt(discharge_totals.get('Tons',0))} / {fmt(load_totals.get('Tons',0))}<span class="delta {diff_class(CAP_TONS, load_totals.get('Tons',0))}">{tons_delta:+d}</span></div><div class="limit">Limit Load {fmt(CAP_TONS)}</div></div>
      <div class="card-stat"><div class="label">Plugs (Discharge / Loading)</div><div class="value">{fmt(discharge_totals.get('Plugs',0))} / {fmt(load_totals.get('Plugs',0))}<span class="delta {diff_class(CAP_PLUGS, load_totals.get('Plugs',0))}">{plugs_delta:+d}</span></div><div class="limit">Limit Load {fmt(CAP_PLUGS)}</div></div>
    </div>

    <div class="section-title">Final Moves by Port - Discharge x Loading</div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th style="width:72px;">POL</th>
            <th style="width:86px;">MOVE</th>
            {"".join(f"<th>{DISPLAY_COLS.get(col, col)}</th>" for col in active_cols)}
            <th class="gold">Units</th>
            <th class="gold">TEUs</th>
            <th class="gold">Tons</th>
            <th class="gold">Plugs</th>
          </tr>
        </thead>
        <tbody>{rows}</tbody>
      </table>
    </div>

    <div class="chart-container">
      <canvas id="compChart"></canvas>
    </div>

    <div class="executive">
      <div class="panel">
        <h2>Container Types by Port</h2>
        <div class="type-grid">
          {container_mix}
        </div>
      </div>
    </div>

    <div class="note">
      * Discharge uses spreadsheet Disch_MERCURY / VOYAGE 1. Loading uses the Baplie/Gate-in actual load. Red means loading below discharge; blue means loading above; blue-gray means match.<br>
      * Operational limits preserved from the comparison report: Units {fmt(CAP_UNITS)}, TEUs {fmt(CAP_TEUS)}, Tons {fmt(CAP_TONS)}, Plugs {fmt(CAP_PLUGS)}.
    </div>
  </div>
</div>
<script>
const ctx = document.getElementById('compChart').getContext('2d');
new Chart(ctx, {{
  type: 'bar',
  plugins: [ChartDataLabels],
  data: {{
    labels: {list(FINAL_POLS)},
    datasets: [
      {{
        label: 'DISCHARGE (Units)',
        data: {[discharge_by_port.get(p, {}).get('Units', 0) for p in FINAL_POLS]},
        backgroundColor: 'rgba(1, 77, 78, 0.7)',
        borderColor: '#174A7C',
        borderWidth: 1,
        breakdowns: {[{k: discharge_by_port.get(p, {}).get(k, 0) for k in active_cols} for p in FINAL_POLS]}
      }},
      {{
        label: 'LOAD (Units)',
        data: {[load_by_port.get(p, {}).get('Units', 0) for p in FINAL_POLS]},
        backgroundColor: 'rgba(185, 147, 47, 0.7)',
        borderColor: '#C59A2F',
        borderWidth: 1,
        breakdowns: {[{k: load_by_port.get(p, {}).get(k, 0) for k in active_cols} for p in FINAL_POLS]}
      }}
    ]
  }},
  options: {{
    responsive: true,
    maintainAspectRatio: false,
    layout: {{ padding: {{ top: 25 }} }},
    plugins: {{
      legend: {{ position: 'top', labels: {{ font: {{ weight: 'bold', size: 11 }} }} }},
      title: {{ display: true, text: 'COMPARISON BY PORT: DISCHARGE VS LOAD (UNITS)', color: '#174A7C', font: {{ size: 14, weight: '900' }} }},
      datalabels: {{
        anchor: 'end',
        align: 'top',
        offset: 2,
        font: {{ weight: 'bold', size: 11 }},
        color: (ctx) => ctx.dataset.borderColor,
        formatter: (val) => val > 0 ? val : ''
      }},
      tooltip: {{
        backgroundColor: 'rgba(255, 255, 255, 0.95)',
        titleColor: '#333',
        bodyColor: '#000',
        borderColor: '#ddd',
        borderWidth: 1,
        padding: 10,
        callbacks: {{
          label: function(context) {{
            let label = context.dataset.label + ': ' + context.parsed.y;
            const brk = context.dataset.breakdowns[context.dataIndex];
            let lines = [label, ''];
            for (const [k, v] of Object.entries(brk)) {{
              if (v > 0) lines.push('  - ' + k + ': ' + v);
            }}
            return lines;
          }}
        }}
      }}
    }},
    scales: {{
      y: {{ beginAtZero: true, grid: {{ display: true, color: '#eee' }} }},
      x: {{ grid: {{ display: false }} }}
    }}
  }}
}});
</script>
</body>
</html>
"""


def write_outputs():
    content = build_html()
    OUTPUT.write_text(content, encoding="utf-8")
    APP_OUTPUT.write_text(content, encoding="utf-8")
    try:
        CLOUD_DIR.mkdir(parents=True, exist_ok=True)
        CLOUD_OUTPUT.write_text(content, encoding="utf-8")
    except PermissionError:
        shutil.copy2(OUTPUT, CLOUD_OUTPUT)
    print(f"[OK] Generated: {OUTPUT}")
    print(f"[OK] Generated: {APP_OUTPUT}")
    print(f"[OK] Generated: {CLOUD_OUTPUT}")


if __name__ == "__main__":
    write_outputs()






