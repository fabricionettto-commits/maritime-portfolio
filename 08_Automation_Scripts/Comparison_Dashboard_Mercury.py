# -*- coding: utf-8 -*-
"""
Comparison_Dashboard_Mercury.py
==============================
Reads the SAME data pipelines used by Capacity and Gate-In reports,
then produces a side-by-side comparison HTML.
"""
import sys, os, base64, re, glob
import pandas as pd, io
from datetime import datetime

VESSEL_DIR = r"C:\capacity_planner\booking_reports\Vessel\MERCURY / VOYAGE 1"
OUTPUT     = os.path.join(VESSEL_DIR, "Report_Comparison_Mercury.html")
APP_DIR    = r"C:\Users\Portfolio_User\Portfolio Workspace\demo service SERVICE - Database\APP\MERCURY / VOYAGE 1"

# ”€ Import logic directly from source scripts ”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€
sys.path.insert(0, VESSEL_DIR)
import Capacity_Mercury_Voyage_1 as cap_mod
import Gate_in_Mercury_Voyage_1 as gi_mod

CNTR_COLS  = ["DC20", "HC40", "HC40_MTY", "RH40", "FR40", "OT40"]
BKG_COLS   = ["DC20", "HC40", "RH40", "FR40", "OT40"]
DISPLAY_COLS = {
    "DC20":"DC20", 
    "HC40":"HC40 Full", 
    "HC40_MTY":"HC40 Empty", 
    "RH40":"RH40", 
    "FR40":"FR40", 
    "OT40":"OT40"
}
FINAL_POLS = ["USNYC", "USHOU", "USNFK", "USNFK"]
CAP_UNITS, CAP_TEUS, CAP_TONS, CAP_RE = 587, 947, 15718, 185

def fmt(v):
    try: return f"{int(v):,}".replace(",", ".")
    except: return "0"

def load_logo_b64():
    for n in ["logo_experience.png", "logo_experience.png", "logo.png", "logo_portfolio.jpg"]:
        p = os.path.join(VESSEL_DIR, n)
        if os.path.exists(p):
            return "../assets/logo_experience.png"
    return ""

# ”€ BOOKING TOTALS: reuse Capacity logic exactly ”€”€”€”€”€”€”€”€”€”€”€”€”€
def get_booking_totals():
    """Run the exact same pipeline as Capacity_Mercury_Voyage_1.generate_report()"""
    files = glob.glob(os.path.join(VESSEL_DIR, "BookingReport*.xls"))
    if not files: return {}, {}
    latest = max(files, key=os.path.getmtime)
    try:
        with open(latest, 'r', encoding='utf-8', errors='ignore') as f: html = f.read()
        df_raw = max(pd.read_html(io.StringIO(html)), key=len)
    except:
        df_raw = pd.read_excel(latest, header=None)
    
    df_clean = cap_mod.scavenger_booking_df(df_raw)
    if df_clean.empty: return {}, {}
    if "ts" not in df_clean.columns: df_clean["ts"] = ""
    for c in ["b20", "b40"]:
        if c in df_clean.columns: df_clean[c] = pd.to_numeric(df_clean[c], errors="coerce").fillna(0)
        else: df_clean[c] = 0
    df_clean["v"]  = df_clean["v"].fillna("").astype(str)
    df_clean["ts"] = df_clean["ts"].fillna("").astype(str)
    
    def norm(x): return "".join(re.findall(r'[A-Z0-9]', str(x).upper()))
    df_clean["_key_v"]  = df_clean["v"].apply(norm)
    df_clean["_key_ts"] = df_clean["ts"].apply(norm)
    
    mask_gh   = df_clean["v"].str.contains(cap_mod.VESSEL_NAME_SPECIFIC, na=False) | df_clean["ts"].str.contains(cap_mod.VESSEL_NAME_SPECIFIC, na=False)
    mask_map  = df_clean["_key_v"].isin(cap_mod.FEEDER_MAP_LIST) | df_clean["_key_ts"].isin(cap_mod.FEEDER_MAP_LIST)
    mask_excl = df_clean["v"].apply(cap_mod._contem_vessel_excluido) | df_clean["ts"].apply(cap_mod._contem_vessel_excluido)
    valid_alp = df_clean[(mask_gh | mask_map) & ~mask_excl].copy()
    
    if not valid_alp.empty:
        valid_alp["POL_FINAL"] = valid_alp["pol"].map(cap_mod.PORT_CONSOLIDATION).fillna(valid_alp["pol"])
        def check_plug_bkg(row):
            ct = str(row.get("ct", "")).upper()
            nor = str(row.get("is_nor", "")).upper()
            temp = str(row.get("temp", "")).strip()
            return ("RH" in ct and "YES" not in nor) or (temp != "" and temp.upper() != "NAN")
        valid_alp["is_plug"] = valid_alp.apply(check_plug_bkg, axis=1)
    
    mil_df = cap_mod.read_mil()
    
    # Build mrows exactly like the Capacity report
    mrows = []
    for pol in FINAL_POLS:
        for pod in cap_mod.ALL_PODS:
            if valid_alp.empty: continue
            sub = valid_alp[(valid_alp["POL_FINAL"] == pol) & (valid_alp["pod"].astype(str).str.contains(pod, na=False, case=False))]
            if sub.empty: continue
            e = {"Line": "ART", "POL": pol, "POD": pod, "Unit Full": 0, "Teu Full": 0, "Tons": 0, "Plugs": 0, "HC40_MTY": 0}
            for ct in BKG_COLS:
                cs = sub[sub["ct"].astype(str).str.contains(ct[:2], na=False, case=False)]
                u = int(cs["b20"].sum() + cs["b40"].sum())
                t = int(cs["b20"].sum() + cs["b40"].sum() * 2)
                e[ct] = u; e["Unit Full"] += u; e["Teu Full"] += t
            e["Plugs"] = int(sub[sub["is_plug"] == True][["b20","b40"]].sum().sum())
            e["Tons"] = e["Teu Full"] * cap_mod.TONS_PER_TEU_ART
            mrows.append(e)
    
    if not mil_df.empty:
        for pol in FINAL_POLS:
            psub = mil_df[mil_df["POL"].astype(str).str.contains(pol, na=False, case=False)]
            if psub.empty: continue
            e = {"Line": "MAR", "POL": pol, "POD": "CATOR", "Unit Full": 0, "Teu Full": 0, "Tons": 0, "Plugs": 0, "HC40_MTY": 0}
            for ct in BKG_COLS:
                ps = psub[psub["cntr_type"].astype(str).str.contains(ct[:2], na=False, case=False)]
                u = int(ps["bkg_20"].sum() + ps["bkg_40"].sum())
                t = int(ps["bkg_20"].sum() + ps["bkg_40"].sum() * 2)
                e[ct] = u; e["Unit Full"] += u; e["Teu Full"] += t
            e["Tons"] = e["Teu Full"] * cap_mod.TONS_PER_TEU_MAR
            mrows.append(e)
    
    # Per-POL summary (combined ART+MAR)
    pol_data = {}
    for pol in FINAL_POLS:
        pr = [r for r in mrows if r["POL"] == pol]
        if not pr: continue
        pol_data[pol] = {c: sum(r.get(c,0) for r in pr) for c in CNTR_COLS}
        pol_data[pol]["Units"] = sum(r["Unit Full"] for r in pr)
        pol_data[pol]["TEUs"]  = sum(r["Teu Full"] for r in pr)
        pol_data[pol]["Tons"]  = sum(r["Tons"] for r in pr)
        pol_data[pol]["Plugs"] = sum(r.get("Plugs",0) for r in pr)
    
    # Grand totals
    grand = {c: sum(r.get(c,0) for r in mrows) for c in CNTR_COLS}
    grand["Units"] = sum(r["Unit Full"] for r in mrows)
    grand["TEUs"]  = sum(r["Teu Full"] for r in mrows)
    grand["Tons"]  = sum(r["Tons"] for r in mrows)
    grand["Plugs"] = sum(r.get("Plugs",0) for r in mrows)
    
    return pol_data, grand

# ”€ GATE-IN TOTALS: load Baplie directly ”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€
def get_gatein_totals():
    """Load Baplie_Mercury_Voyage_1.xlsx and compute totals identically to Gate_in_Mercury_Voyage_1.py"""
    import re as _re
    
    f_path = os.path.join(VESSEL_DIR, "Baplie_Mercury_Voyage_1.xlsx")
    if not os.path.exists(f_path):
        files = (
            glob.glob(os.path.join(VESSEL_DIR, "Baplie_MERCURY / VOYAGE 1*.xlsx"))
            + glob.glob(os.path.join(VESSEL_DIR, "Baplie*.xlsx"))
        )
        if not files: return {}, {}
        f_path = max(files, key=os.path.getmtime)
    
    xl = pd.ExcelFile(f_path)
    # Find correct sheet: look for LOAD or BAPLIE, fallback to first
    sheet = xl.sheet_names[0]
    for s in xl.sheet_names:
        if "LOAD" in s.upper() or "BAPLIE" in s.upper():
            sheet = s; break
    
    df_raw = xl.parse(sheet, header=None)
    
    # Detect header row
    h_idx = 0
    for i in range(min(50, len(df_raw))):
        vals = [str(x).upper() for x in df_raw.iloc[i].fillna("").values]
        if any(k in v for k in ["POL","TYPE","WEIGHT","CONTAINER"] for v in vals):
            h_idx = i; break
    
    df = df_raw.iloc[h_idx + 1:].copy()
    df.columns = [str(x).strip().upper() for x in df_raw.iloc[h_idx].values]
    
    # Map columns
    col_map = {}
    for c in df.columns:
        cu = c.upper()
        if _re.search(r"OWNER|CARRIER|LINE", cu): col_map[c] = "carrier"
        elif _re.search(r"CNTR|CONTAINER.*ID", cu): col_map[c] = "cntr_no"
        elif _re.search(r"TYPE|ISO", cu): col_map[c] = "ct"
        elif _re.search(r"^POL$", cu): col_map[c] = "pol"
        elif _re.search(r"^POD$", cu): col_map[c] = "pod"
        elif _re.search(r"WEIGHT|GROSS", cu): col_map[c] = "weight"
        elif _re.search(r"SETTING|TEMP", cu): col_map[c] = "temp"
        elif _re.search(r"FULL|EMPTY|FE", cu): col_map[c] = "fe"
    
    df = df.rename(columns=col_map)
    df = df.loc[:, ~df.columns.duplicated()]
    for col in ("ct","pol","pod","temp","carrier","weight","fe"):
        if col not in df.columns: df[col] = "F"
    
    df["weight"] = pd.to_numeric(df["weight"], errors="coerce").fillna(0.0) / 1000.0
    
    # Status normalization
    def norm_fe(x):
        x = str(x).upper()
        if any(k in x for k in ["MT", "E", "EMPTY"]): return "E"
        return "F"
    df["status"] = df["fe"].apply(norm_fe)

    df["ct_norm"] = df["ct"].apply(gi_mod.normalize_cntr_type)
    df["pol_final"] = df["pol"].map(gi_mod.PORT_CONSOLIDATION).fillna(df["pol"]).str.upper()
    df["is_plug"] = df.apply(lambda r: (r["ct_norm"]=="RH40") or (str(r["temp"]).strip() != "" and str(r["temp"]).upper() != "NAN"), axis=1)
    
    # Drop any fully-empty rows
    df = df.dropna(subset=["ct"], how="all")
    df = df[df["ct"].astype(str).str.strip() != ""]
    
    # Per-POL summary
    pol_data = {}
    for pol in FINAL_POLS:
        sub = df[df["pol_final"] == pol]
        if sub.empty: continue
        pol_data[pol] = {
            "DC20": sum(1 for idx, r in sub.iterrows() if r["ct_norm"] == "DC20"),
            "HC40": sum(1 for idx, r in sub.iterrows() if r["ct_norm"] == "HC40" and r["status"] == "F"),
            "HC40_MTY": sum(1 for idx, r in sub.iterrows() if r["ct_norm"] == "HC40" and r["status"] == "E"),
            "RH40": sum(1 for idx, r in sub.iterrows() if r["ct_norm"] == "RH40"),
            "FR40": sum(1 for idx, r in sub.iterrows() if r["ct_norm"] == "FR40"),
            "OT40": sum(1 for idx, r in sub.iterrows() if r["ct_norm"] == "OT40"),
        }
        pol_data[pol]["Units"] = len(sub)
        pol_data[pol]["TEUs"]  = sum(2 if "40" in t else 1 for t in sub["ct_norm"])
        pol_data[pol]["Tons"]  = sub["weight"].sum()
        pol_data[pol]["Plugs"] = int(sub["is_plug"].sum())
    
    # Grand totals
    grand = {
        "DC20": sum(1 for idx, r in df.iterrows() if r["ct_norm"] == "DC20"),
        "HC40": sum(1 for idx, r in df.iterrows() if r["ct_norm"] == "HC40" and r["status"] == "F"),
        "HC40_MTY": sum(1 for idx, r in df.iterrows() if r["ct_norm"] == "HC40" and r["status"] == "E"),
        "RH40": sum(1 for idx, r in df.iterrows() if r["ct_norm"] == "RH40"),
        "FR40": sum(1 for idx, r in df.iterrows() if r["ct_norm"] == "FR40"),
        "OT40": sum(1 for idx, r in df.iterrows() if r["ct_norm"] == "OT40"),
    }
    grand["Units"] = len(df)
    grand["TEUs"]  = sum(2 if "40" in t else 1 for t in df["ct_norm"])
    grand["Tons"]  = df["weight"].sum()
    grand["Plugs"] = int(df["is_plug"].sum())
    
    return pol_data, grand

# ”€ BUILD HTML ”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€”€
def generate_comparison():
    bk_pol, bk_tot = get_booking_totals()
    gi_pol, gi_tot = get_gatein_totals()
    
    # Dynamically hide columns with zero cargo (like HC40 Empty if empty)
    active_cols = []
    for c in CNTR_COLS:
        if bk_tot.get(c, 0) > 0 or gi_tot.get(c, 0) > 0:
            active_cols.append(c)
    
    # Fallback to standard columns if everything is zero for some reason
    if not active_cols: active_cols = ["DC20", "HC40", "RH40"]

    empty = {c:0 for c in CNTR_COLS + ["Units","TEUs","Tons","Plugs"]}
    
    def diff_cls(b, g):
        if g == 0: return ""
        if g < b: return "text-danger"
        if g > b: return "text-primary"
        return "text-success"
    
    rows = ""
    for pol in FINAL_POLS:
        b = bk_pol.get(pol, empty)
        g = gi_pol.get(pol, empty)
        
        rows += f"<tr><td rowspan='2' class='pol-cell'>{pol}</td><td class='src-book'>BOOK</td>"
        for c in active_cols: rows += f"<td>{b.get(c,0)}</td>"
        rows += f"<td class='fw-bold'>{b.get('Units',0)}</td><td class='fw-bold'>{b.get('TEUs',0)}</td><td class='fw-bold'>{fmt(b.get('Tons',0))}</td><td class='fw-bold'>{b.get('Plugs',0)}</td></tr>"
        
        rows += f"<tr class='gate-row'><td class='src-gate'>GATE</td>"
        for c in active_cols: rows += f"<td class='{diff_cls(b.get(c,0), g.get(c,0))}'>{g.get(c,0)}</td>"
        rows += f"<td class='fw-bold {diff_cls(b.get('Units',0), g.get('Units',0))}'>{g.get('Units',0)}</td>"
        rows += f"<td class='fw-bold {diff_cls(b.get('TEUs',0), g.get('TEUs',0))}'>{g.get('TEUs',0)}</td>"
        rows += f"<td class='fw-bold {diff_cls(b.get('Tons',0), g.get('Tons',0))}'>{fmt(g.get('Tons',0))}</td>"
        rows += f"<td class='fw-bold {diff_cls(b.get('Plugs',0), g.get('Plugs',0))}'>{g.get('Plugs',0)}</td></tr>"
    
    # GRAND TOTAL (gold)
    rows += f"<tr class='total-row'><td colspan='2'>GRAND TOTAL</td>"
    for c in active_cols: rows += f"<td>{bk_tot.get(c,0)} <small>/</small> {gi_tot.get(c,0)}</td>"
    rows += f"<td>{bk_tot.get('Units',0)} <small>/</small> {gi_tot.get('Units',0)}</td>"
    rows += f"<td>{bk_tot.get('TEUs',0)} <small>/</small> {gi_tot.get('TEUs',0)}</td>"
    rows += f"<td>{fmt(bk_tot.get('Tons',0))} <small>/</small> {fmt(gi_tot.get('Tons',0))}</td>"
    rows += f"<td>{bk_tot.get('Plugs',0)} <small>/</small> {gi_tot.get('Plugs',0)}</td></tr>"
    
    gi_units_pct = gi_tot.get('Units',0)*100/CAP_UNITS if CAP_UNITS else 0
    gi_teus_pct  = gi_tot.get('TEUs',0)*100/CAP_TEUS if CAP_TEUS else 0
    gi_tons_pct  = gi_tot.get('Tons',0)*100/CAP_TONS if CAP_TONS else 0
    gi_plugs_pct = gi_tot.get('Plugs',0)*100/CAP_RE if CAP_RE else 0
    
    html = f"""<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Mercury Comparison</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
<style>
:root {{ --teal: #1D4E89; --gold: #C59A2F; }}
body {{ font-family:'Segoe UI',Arial; font-size:10.5px; background:#222; padding:8px; }}
.main {{ max-width:1100px; margin:0 auto; background:#fff; padding:12px; border-radius:8px; }}
.header-bar {{ display: flex; align-items: center; gap: 12px; margin-bottom: 5px; }}
.logo-box {{ background: var(--teal); padding: 5px; border-radius: 8px; border: 2px solid var(--gold); }}
.logo-box img {{ height: 35px; }}
.header-bar h1 {{ font-size: 18px; font-weight: 900; color: var(--teal); margin: 0; }}
.cards {{ display:grid; grid-template-columns:repeat(4,1fr); gap:8px; margin-bottom:10px; }}
.card-stat {{ background: var(--teal); color: #fff; padding: 6px; border-radius: 6px; border-bottom: 4px solid var(--gold); text-align: center; }}
.card-stat .value {{ font-size: 18px; font-weight: 900; }}
.card-stat .label {{ font-size: 8.5px; color: var(--gold); font-weight: 700; }}
.stat-cap {{ font-size: 9.5px; opacity:.8; }}
table {{ width:100%; border-collapse:collapse; text-align:center; border:2px solid #333; table-layout: fixed; }}
thead th {{ background:var(--teal) !important; color:#fff !important; padding:4px 2px; font-size:10.5px; border:1px solid rgba(255,255,255,0.1); }}
tbody td {{ padding:2px 3px; border:1px solid #bbb; font-weight:700; font-size:10.5px; color:#000; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
.chart-container {{ margin-top: 20px; padding: 10px; background: #f8f9fa; border-radius: 8px; border: 1px solid #ddd; position: relative; height: 280px; }}
tbody tr:nth-of-type(4n+1), tbody tr:nth-of-type(4n+2) {{ background:#f2f2f2; }}
tbody tr:nth-of-type(4n+3), tbody tr:nth-of-type(4n+4) {{ background:#e0e0e0; }}
.pol-cell {{ background:#f9f9f9 !important; border-right:2px solid #333 !important; }}
.src-book {{ background:#fffde7 !important; color:#856404; font-size:9px; font-weight:800; border-right:2px solid #333 !important; }}
.src-gate {{ background:#e8f5e9 !important; color:#1b5e20; font-size:9px; font-weight:800; border-right:2px solid #333 !important; }}
.total-row {{ background:var(--gold) !important; color:#fff !important; font-weight:900 !important; }}
.total-row td {{ color:#fff !important; border-color:#8a6f1f !important; }}
.text-danger {{ color:#b71c1c !important; }}
.text-primary {{ color:#0d47a1 !important; }}
.text-success {{ color:#1b5e20 !important; }}
.gold-th {{ background:var(--gold) !important; color:#fff !important; }}
</style></head><body>
<div class="main">
    <div class="header-bar">
        <div class="logo-box"><img src="{load_logo_b64() or 'logo.png'}"></div>
        <div style="flex:1;">
            <h1>COMPARISON: MERCURY / VOYAGE 1</h1>
            <div style="font-size:10px;color:#666;">BookingReport.xls +  vs Baplie_MERCURY_VOYAGE_1 (LOAD) | Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
        </div>
        <div><button class="btn btn-sm btn-outline-secondary px-3" style="font-size:10px; font-weight:800;" onclick="if (history.length > 1) {{ history.back(); }} else {{ location.href='/navios'; }}">BACK</button></div>
    </div>
    <div class="cards mt-2">
        <div class="card-stat"><div class="label">UNITS (BOOK / GATE)</div><div class="value">{bk_tot.get('Units',0)} <small>/</small> {gi_tot.get('Units',0)}</div><div class="stat-cap">Limit: {CAP_UNITS} (Util: {gi_units_pct:.1f}%)</div></div>
        <div class="card-stat"><div class="label">TEUs (BOOK / GATE)</div><div class="value">{bk_tot.get('TEUs',0)} <small>/</small> {gi_tot.get('TEUs',0)}</div><div class="stat-cap">Limit: {CAP_TEUS} (Util: {gi_teus_pct:.1f}%)</div></div>
        <div class="card-stat"><div class="label">TONS (BOOK / GATE)</div><div class="value">{fmt(bk_tot.get('Tons',0))} <small>/</small> {fmt(gi_tot.get('Tons',0))}</div><div class="stat-cap">Limit: {fmt(CAP_TONS)} (Util: {gi_tons_pct:.1f}%)</div></div>
        <div class="card-stat"><div class="label">PLUGS (BOOK / GATE)</div><div class="value">{bk_tot.get('Plugs',0)} <small>/</small> {gi_tot.get('Plugs',0)}</div><div class="stat-cap">Limit: {CAP_RE} (Util: {gi_plugs_pct:.1f}%)</div></div>
    </div>
    <table class="mt-2 text-center">
        <thead><tr>
            <th style="width:65px;">POL</th><th style="width:45px;">SRC</th>
            {"".join(f"<th style='width:65px;'>{DISPLAY_COLS.get(c,c)}</th>" for c in active_cols)}
            <th class="gold-th" style="width:65px;">Units</th>
            <th class="gold-th" style="width:65px;">TEUs</th>
            <th class="gold-th" style="width:75px;">Tons</th>
            <th class="gold-th" style="width:65px;">Plugs</th>
        </tr></thead>
        <tbody>{rows}</tbody>
    </table>
    
    <div class="row mt-3">
        <div class="col-12">
            <div class="chart-container">
                <canvas id="compChart"></canvas>
            </div>
        </div>
    </div>

    <div style="margin-top:8px;font-size:10px;color:#555;font-style:italic;line-height:1.4;">
        * BOOK = Capacity (BookingReport + ) | GATE = Baplie LOAD actual. Red = missing, Blue = extra.<br>
        * There are 50 x HC40 empty units in USHOU, so the <strong>HC40 Empty</strong> column is shown separately from HC40 Full. Ports without empty units remain at 0.
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
                label: 'BOOKED (Units)',
                data: {[bk_pol.get(p, {}).get('Units', 0) for p in FINAL_POLS]},
                backgroundColor: 'rgba(1, 77, 78, 0.7)',
                borderColor: '#174A7C',
                borderWidth: 1,
                breakdowns: {[ {k: bk_pol.get(p, {}).get(k, 0) for k in active_cols} for p in FINAL_POLS ]}
            }},
            {{
                label: 'GATE-IN (Units)',
                data: {[gi_pol.get(p, {}).get('Units', 0) for p in FINAL_POLS]},
                backgroundColor: 'rgba(185, 147, 47, 0.7)',
                borderColor: '#C59A2F',
                borderWidth: 1,
                breakdowns: {[ {k: gi_pol.get(p, {}).get(k, 0) for k in active_cols} for p in FINAL_POLS ]}
            }}
        ]
    }},
    options: {{
        responsive: true, maintainAspectRatio: false,
        layout: {{ padding: {{ top: 25 }} }},
        plugins: {{
            legend: {{ position: 'top', labels: {{ font: {{ weight: 'bold', size: 11 }} }} }},
            title: {{ display: true, text: 'COMPARISON BY PORT: BOOKED VS GATE-IN (UNITS)', color: '#174A7C', font: {{ size: 14, weight: '900' }} }},
            datalabels: {{
                anchor: 'end', align: 'top', offset: 2,
                font: {{ weight: 'bold', size: 11 }},
                color: (ctx) => ctx.dataset.borderColor,
                formatter: (val) => val > 0 ? val : ''
            }},
            tooltip: {{
                backgroundColor: 'rgba(255, 255, 255, 0.95)',
                titleColor: '#333', bodyColor: '#000', borderColor: '#ddd', borderWidth: 1,
                padding: 10,
                callbacks: {{
                    label: function(context) {{
                        let label = context.dataset.label + ': ' + context.parsed.y;
                        const brk = context.dataset.breakdowns[context.dataIndex];
                        let lines = [label, ''];
                        for (const [k, v] of Object.entries(brk)) {{
                            if (v > 0) lines.push('  -¢ ' + k + ': ' + v);
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
</body></html>"""
    
    with open(OUTPUT, "w", encoding="utf-8") as f: f.write(html)
    print(f"[OK] Comparison Dashboard gerado: {OUTPUT}")

    try:
        os.makedirs(APP_DIR, exist_ok=True)
        app_path = os.path.join(APP_DIR, os.path.basename(OUTPUT))
        with open(app_path, "w", encoding="utf-8") as f: f.write(html)
        print(f"[OK] Copia salva no APP: {app_path}")
    except Exception as e:
        print(f"[!] Erro ao salvar APP: {e}")

    # Secondary save to cloud folder
    CLOUD_DIR = r"C:\Users\Portfolio_User\Portfolio Workspace\demo service SERVICE - Database\Arquivos_HTML"
    try:
        os.makedirs(CLOUD_DIR, exist_ok=True)
        c_path = os.path.join(CLOUD_DIR, os.path.basename(OUTPUT))
        with open(c_path, "w", encoding="utf-8") as f: f.write(html)
        print(f"[OK] Copia salva na Nuvem: {c_path}")
    except Exception as e:
        print(f"[!] Erro ao salvar cloud: {e}")

if __name__ == "__main__":
    generate_comparison()





