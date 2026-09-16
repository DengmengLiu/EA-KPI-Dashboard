#!/usr/bin/env python3
"""
excel_to_json.py  –  Converts the EA RISE Impact KPI Excel workbook to data.json.

COLUMN MAPPING (1-indexed, adjust if your Excel layout differs):
  Col 1: Date        (YYYY-MM-DD string or Excel date)
  Col 2: Region      (ANZ / G.China / India / Japan / Korea / South East Asia)
  Col 3: Account ID
  Col 4: Account Name
  Col 5: EA Name
  Col 6: compliance_check   (1/0)
  Col 7: qbr_govnc          (1/0)
  Col 8: s4_upgrade baseline count
  Col 9: s4_upgrade          (1/0)
  Col 10: ai_adoption_plan   (1/0)
  Col 11: AI tools baseline (ai_agent / joule / ai_btp shared)
  Col 12: ai_agent           (1/0)
  Col 13: joule              (1/0)
  Col 14: ai_btp             (1/0)
  Col 15: risk_assessment    (1/0)
  Col 16: churn_mitigation   (1/0)
  Col 17: lead_created       (1/0)
  Col 18: rwsm_dibo          (1/0)
  Col 19: cc_princ_score     (1/0)
  Col 20: calm_tenant        (1/0)
  Col 21: signavio baseline
  Col 22: signavio           (1/0)
  Col 23: leanix baseline
  Col 24: leanix             (1/0)
  Col 25: tricentis          (1/0)
  Col 26: walkme baseline
  Col 27: walkme             (1/0)
"""

import openpyxl
import json
import os
from collections import defaultdict
from datetime import date as date_type, datetime

# ── Config ────────────────────────────────────────────────────────────────────
EXCEL_PATH = "data/latest.xlsx"
OUT_PATH   = "data/data.json"

# Column indices (0-based for Python)
C_DATE   = 0
C_REGION = 1
C_ACCID  = 2
C_NAME   = 3
C_EA     = 4

MEASURES = [
    {"id": "compliance_check",  "name": "Compliance Check",       "short": "Compliance", "bc": None, "col": 6},
    {"id": "qbr_govnc",         "name": "QBR Governance",          "short": "QBR Gvnc",  "bc": None, "col": 7},
    {"id": "s4_upgrade",        "name": "S4 Upgrade Commitment",   "short": "S4 Upgrd",  "bc": 8,    "col": 9},
    {"id": "ai_adoption_plan",  "name": "AI Adoption Plan",        "short": "AI Adopt",  "bc": None, "col": 10},
    {"id": "ai_agent",          "name": "AI Agent Activated",      "short": "AI Agent",  "bc": 11,   "col": 12},
    {"id": "ai_btp",            "name": "AI BTP Services",         "short": "AI BTP",    "bc": 11,   "col": 14},
    {"id": "calm_tenant",       "name": "CALM Tenant",             "short": "CALM",      "bc": None, "col": 20},
    {"id": "churn_mitigation",  "name": "Churn Mitigation",        "short": "Churn",     "bc": None, "col": 16},
    {"id": "joule",             "name": "Joule Activated",         "short": "Joule",     "bc": 11,   "col": 13},
    {"id": "lead_created",      "name": "Lead Created",            "short": "Lead",      "bc": None, "col": 17},
    {"id": "leanix",            "name": "LeanIX Usage",            "short": "LeanIX",    "bc": 23,   "col": 24},
    {"id": "risk_assessment",   "name": "Risk Assessment",         "short": "Risk Asmt", "bc": None, "col": 15},
    {"id": "rwsm_dibo",         "name": "RwSM DiBO Activated",     "short": "RwSM DiBO", "bc": None, "col": 18},
    {"id": "signavio",          "name": "Signavio Usage",          "short": "Signavio",  "bc": 21,   "col": 22},
    {"id": "tricentis",         "name": "Tricentis",               "short": "Tricentis", "bc": None, "col": 25},
    {"id": "walkme",            "name": "WalkMe Usage",            "short": "WalkMe",    "bc": 26,   "col": 27},
    {"id": "cc_princ_score",    "name": "CC Princ Score",          "short": "CC Score",  "bc": None, "col": 19},
]

BC_COLS = [8, 11, 21, 23, 26]   # 1-indexed baseline columns

CATS = [
    {"id": "ai",        "label": "AI",         "ids": ["ai_adoption_plan","ai_agent","ai_btp","joule"]},
    {"id": "cleancore", "label": "Clean Core",  "ids": ["cc_princ_score","rwsm_dibo"]},
    {"id": "eakpi",     "label": "EA KPI",      "ids": ["churn_mitigation","compliance_check","lead_created","qbr_govnc"]},
    {"id": "rise",      "label": "RISE",        "ids": ["risk_assessment","s4_upgrade"]},
    {"id": "toolchain", "label": "Toolchain",   "ids": ["calm_tenant","leanix","signavio","tricentis","walkme"]},
]

SUB_REGIONS = ["ANZ", "G.China", "India", "Japan", "Korea", "South East Asia"]

# ── Helpers ───────────────────────────────────────────────────────────────────
def to_int(v):
    if v is None: return 0
    if isinstance(v, (int, float)): return int(v)
    s = str(v).strip().lower()
    return 1 if s in ("1", "yes", "y", "true", "x") else 0

def to_str(v, default=""):
    if v is None: return default
    return str(v).strip()

def fmt_date(v):
    if v is None: return None
    if isinstance(v, (date_type, datetime)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y%m%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return s

def rate(count, baseline):
    if baseline == 0: return 0.0
    return round(count / baseline * 100, 1)

# ── Load workbook ─────────────────────────────────────────────────────────────
print(f"Loading {EXCEL_PATH} ...")
wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

ws = None
preferred = ["Data","data","KPI","kpi","Sheet1","Sheet 1","Ops","ops"]
for name in preferred:
    if name in wb.sheetnames:
        ws = wb[name]; break
if ws is None:
    ws = wb.active
print(f"  Sheet: {ws.title!r}  ({ws.max_row} rows, {ws.max_column} cols)")
print(f"  All sheets in workbook: {wb.sheetnames}")

# Print first row (headers) so you can verify column mapping
header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
if header_row:
    print("  Column headers (1-indexed):")
    for i, h in enumerate(header_row, 1):
        print(f"    Col {i:2d}: {h}")

# Print first data row as a sanity check
first_data = next(ws.iter_rows(min_row=2, max_row=2, values_only=True), None)
if first_data:
    print(f"  First data row sample: Date={first_data[0]}  Region={first_data[1]}  AccID={first_data[2]}")

# ── Parse rows ────────────────────────────────────────────────────────────────
rows_by_date_region = defaultdict(list)
all_dates = set()

region_map = {
    "anz":"ANZ","australia":"ANZ","new zealand":"ANZ","nz":"ANZ",
    "china":"G.China","greater china":"G.China","gchina":"G.China","g.china":"G.China",
    "sea":"South East Asia","southeast asia":"South East Asia","south east asia":"South East Asia",
    "japan":"Japan","korea":"Korea","india":"India",
}

for row in ws.iter_rows(min_row=2, values_only=True):
    if not row or row[C_DATE] is None:
        continue
    d = fmt_date(row[C_DATE])
    if not d:
        continue
    region_raw = to_str(row[C_REGION])
    region = region_map.get(region_raw.lower(), region_raw)
    if region not in SUB_REGIONS:
        continue

    acc_id   = to_str(row[C_ACCID])
    acc_name = to_str(row[C_NAME])
    ea_name  = to_str(row[C_EA])

    m = {m_["id"]: to_int(row[m_["col"] - 1]) for m_ in MEASURES}
    b = {str(c): to_int(row[c - 1]) for c in BC_COLS}

    rows_by_date_region[(d, region)].append({
        "id": acc_id, "name": acc_name, "ea": ea_name, "m": m, "b": b
    })
    all_dates.add(d)

all_dates = sorted(all_dates)
print(f"  Dates found: {all_dates}")
if not all_dates:
    print("ERROR: No data rows found. Check column mapping.")
    raise SystemExit(1)

latest  = all_dates[-1]
prev_wk = all_dates[-2] if len(all_dates) >= 2 else latest

# prev_mo: find the date closest to 28 days before latest
from datetime import timedelta
latest_dt = datetime.strptime(latest, "%Y-%m-%d")
target_mo = latest_dt - timedelta(days=28)
prev_mo   = min(all_dates, key=lambda d: abs((datetime.strptime(d, "%Y-%m-%d") - target_mo).days))
# Make sure prev_mo is not the same as latest or prev_wk
if prev_mo == latest and len(all_dates) >= 3:
    prev_mo = all_dates[-3]

print(f"  Latest={latest}  PrevWk={prev_wk}  PrevMo={prev_mo}")

# ── Compute KPI summary for one region + date ─────────────────────────────────
def compute(region, date_str):
    rows = rows_by_date_region.get((date_str, region), [])
    total = len(rows)
    out = {}
    for m_ in MEASURES:
        mid = m_["id"]; bc = m_["bc"]
        if bc is None:
            bl  = total
            cnt = sum(r["m"][mid] for r in rows)
        else:
            bl  = sum(r["b"][str(bc)] for r in rows)
            cnt = sum(r["m"][mid] for r in rows if r["b"][str(bc)] == 1)
        out[mid] = {"rate": rate(cnt, bl), "count": cnt, "baseline": bl}
    return out

def agg_compute(date_str):
    rows = []
    for sr in SUB_REGIONS:
        rows += rows_by_date_region.get((date_str, sr), [])
    total = len(rows)
    out = {}
    for m_ in MEASURES:
        mid = m_["id"]; bc = m_["bc"]
        if bc is None:
            bl  = total
            cnt = sum(r["m"][mid] for r in rows)
        else:
            bl  = sum(r["b"][str(bc)] for r in rows)
            cnt = sum(r["m"][mid] for r in rows if r["b"][str(bc)] == 1)
        out[mid] = {"rate": rate(cnt, bl), "count": cnt, "baseline": bl}
    return out

# ── Build kpis dict ───────────────────────────────────────────────────────────
kpis = {}
for region in SUB_REGIONS + ["APAC"]:
    kpis[region] = {}
    for m_ in MEASURES:
        mid = m_["id"]
        if region == "APAC":
            curr = agg_compute(latest)[mid]
            p_wk = agg_compute(prev_wk)[mid]
            p_mo = agg_compute(prev_mo)[mid]
        else:
            curr = compute(region, latest)[mid]
            p_wk = compute(region, prev_wk)[mid]
            p_mo = compute(region, prev_mo)[mid]
        wd = round(curr["rate"] - p_wk["rate"], 1)
        md = round(curr["rate"] - p_mo["rate"], 1)
        if   abs(wd) >= 1.5: trend = "improved" if wd > 0 else "declined"
        elif abs(md) >= 1.5: trend = "improved" if md > 0 else "declined"
        else:                trend = "flat"
        kpis[region][mid] = {"current": curr, "prev_wk": p_wk, "prev_mo": p_mo,
                              "wd": wd, "md": md, "trend": trend}

# ── Sums ──────────────────────────────────────────────────────────────────────
def build_sums(rk):
    total = len(rk)
    imp = sum(1 for v in rk.values() if v["trend"] == "improved")
    dec = sum(1 for v in rk.values() if v["trend"] == "declined")
    return {"total": total, "improved": imp, "declined": dec, "flat": total-imp-dec, "na": 0}

sums = {r: build_sums(kpis[r]) for r in ["APAC"] + SUB_REGIONS}

# ── Accounts ──────────────────────────────────────────────────────────────────
accs = {}
for region in SUB_REGIONS:
    accs[region] = {}
    for d in all_dates:
        accs[region][d] = {}
        for row in rows_by_date_region.get((d, region), []):
            if row["id"]:
                accs[region][d][row["id"]] = {
                    "name": row["name"], "ea": row["ea"],
                    "m": row["m"], "b": row["b"]
                }

# ── Assemble & write ──────────────────────────────────────────────────────────
# IMPORTANT: sums must be at the TOP LEVEL of DATA (not inside fixed)
# The dashboard JS reads DATA.sums, DATA.fixed.kpis, DATA.fixed.ts separately
DATA = {
    "measures": MEASURES,
    "cats":     CATS,
    "regions":  ["APAC"] + SUB_REGIONS,
    "fixed": {
        "ts":   {"latest": latest, "prev_wk": prev_wk, "prev_mo": prev_mo, "all": all_dates},
        "kpis": kpis,
    },
    "sums": sums,   # top-level, NOT inside fixed
    "accs": accs,
}

os.makedirs("data", exist_ok=True)
with open(OUT_PATH, "w", encoding="utf-8") as f:
    json.dump(DATA, f, separators=(",", ":"), ensure_ascii=False)
print(f"Written: {OUT_PATH}  ({os.path.getsize(OUT_PATH):,} bytes)")
