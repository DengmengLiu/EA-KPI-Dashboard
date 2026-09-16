#!/usr/bin/env python3
"""
setup.py  –  ONE-TIME SETUP SCRIPT
=====================================
Run this once on your local machine to prepare the GitHub-ready zip package.

Usage:
    python scripts/setup.py  "path/to/EA_RISE_Impact_KPI_Dashboard.html"

What it does:
  1. Reads your HTML dashboard file
  2. Extracts the embedded DATA JSON  →  saves as data/data.json
  3. Modifies index.html to add section markers (needed for future auto-refresh)
  4. Packages everything into  ea-kpi-dashboard.zip  (ready to upload to GitHub)

Requirements:  Python 3.8+  (no extra packages needed)
"""

import sys, re, json, os, zipfile, shutil

# ── Argument check ────────────────────────────────────────────────────────────
if len(sys.argv) < 2:
    # Try to auto-detect HTML in common locations
    candidates = [
        "EA_RISE_Impact_KPI_Dashboard_v60_v9__2_.html",
        "EA_RISE_Impact_KPI_Dashboard.html",
        "dashboard.html",
        "index.html",
    ]
    html_src = None
    for c in candidates:
        if os.path.exists(c):
            html_src = c
            break
    if html_src is None:
        print("Usage: python scripts/setup.py  path/to/your_dashboard.html")
        print("\nCould not auto-detect HTML file. Please pass it as an argument.")
        sys.exit(1)
    print(f"Auto-detected HTML: {html_src}")
else:
    html_src = sys.argv[1]

if not os.path.exists(html_src):
    print(f"ERROR: File not found: {html_src}")
    sys.exit(1)

# ── Read HTML ─────────────────────────────────────────────────────────────────
print(f"\n1. Reading HTML: {html_src}")
with open(html_src, "r", encoding="utf-8", errors="replace") as f:
    html = f.read()
print(f"   Size: {len(html):,} chars")

# ── Extract DATA JSON ─────────────────────────────────────────────────────────
print("\n2. Extracting DATA JSON...")
marker = "const DATA="
if marker not in html:
    print(f"ERROR: Could not find '{marker}' in the HTML.")
    print("Make sure this is the correct dashboard HTML file.")
    sys.exit(1)

start_idx = html.index(marker) + len(marker)

# Walk forward counting braces to find the matching closing }
depth = 0
i = start_idx
while i < len(html):
    ch = html[i]
    if ch == '{':
        depth += 1
    elif ch == '}':
        depth -= 1
        if depth == 0:
            end_idx = i + 1
            break
    i += 1
else:
    print("ERROR: Could not find end of DATA object.")
    sys.exit(1)

data_json_str = html[start_idx:end_idx]
try:
    data_obj = json.loads(data_json_str)
except json.JSONDecodeError as e:
    print(f"ERROR: Failed to parse DATA JSON: {e}")
    sys.exit(1)
print(f"   Extracted OK  ({len(data_json_str):,} chars)")
latest_date = data_obj.get("fixed", {}).get("ts", {}).get("latest", "unknown")
print(f"   Latest date in data: {latest_date}")

# ── Save data/data.json ───────────────────────────────────────────────────────
print("\n3. Saving data/data.json...")
os.makedirs("data", exist_ok=True)
data_compact = json.dumps(data_obj, separators=(",", ":"), ensure_ascii=False)
with open("data/data.json", "w", encoding="utf-8") as f:
    f.write(data_compact)
print(f"   Saved: data/data.json  ({os.path.getsize('data/data.json'):,} bytes)")

# ── Create modified index.html with markers ───────────────────────────────────
print("\n4. Creating index.html with data-section markers...")
old_block = f"const DATA={data_json_str};"
new_block = f"/* __DATA_START__ */const DATA={data_json_str};/* __DATA_END__ */"
if old_block not in html:
    print("WARNING: Could not inject markers precisely. Using regex fallback.")
    # Fallback: rebuild from data object
    new_block = f"/* __DATA_START__ */const DATA={data_compact};/* __DATA_END__ */"
    html_out = re.sub(
        r"const DATA=\{",
        "/* __DATA_START__ */const DATA={",
        html, count=1
    )
    # Now find the end marker position and insert the end marker
    # This is tricky — just use the original HTML as-is if fallback fails
    html_out = html  # keep original; inject_data.py won't work but dashboard still works
    print("   Using original HTML (markers not injected - auto-refresh will be limited)")
else:
    html_out = html.replace(old_block, new_block, 1)
    print("   Markers injected OK")

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_out)
print(f"   Saved: index.html  ({os.path.getsize('index.html'):,} bytes)")

# ── Copy Excel to data/ if present ───────────────────────────────────────────
xlsx_src = None
for candidate in [
    "AFO_EA_Impact_KPI___cACV_Territory_Analyzer_Ops_updated__3_.xlsx",
    "latest.xlsx",
    "data/latest.xlsx",
]:
    if os.path.exists(candidate):
        xlsx_src = candidate
        break

if xlsx_src and xlsx_src != "data/latest.xlsx":
    shutil.copy2(xlsx_src, "data/latest.xlsx")
    print(f"\n   Copied Excel → data/latest.xlsx")

# ── Build the zip ─────────────────────────────────────────────────────────────
print("\n5. Building ea-kpi-dashboard.zip...")
ZIP_PATH = "ea-kpi-dashboard.zip"

required_files = [
    ("index.html",                        "index.html"),
    ("data/data.json",                    "data/data.json"),
    ("scripts/excel_to_json.py",          "scripts/excel_to_json.py"),
    ("scripts/inject_data.py",            "scripts/inject_data.py"),
    ("scripts/setup.py",                  "scripts/setup.py"),
    (".github/workflows/refresh.yml",     ".github/workflows/refresh.yml"),
    ("README.md",                         "README.md"),
]
optional_files = [
    ("data/latest.xlsx", "data/latest.xlsx"),
]

with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
    for src, dest in required_files:
        if os.path.exists(src):
            zf.write(src, dest)
            size = os.path.getsize(src)
            print(f"   + {dest:50s}  {size:>10,} B")
        else:
            print(f"   ! MISSING: {src}")
    for src, dest in optional_files:
        if os.path.exists(src):
            zf.write(src, dest)
            size = os.path.getsize(src)
            print(f"   + {dest:50s}  {size:>10,} B  (optional)")

zip_size = os.path.getsize(ZIP_PATH)
print(f"\n✅ DONE!")
print(f"   Zip: {ZIP_PATH}  ({zip_size / 1024 / 1024:.1f} MB)")
print(f"\nNext steps:")
print(f"  1. Create a GitHub repo named  ea-kpi-dashboard  (set to Public)")
print(f"  2. Upload ALL contents of {ZIP_PATH} to the repo root")
print(f"  3. Enable GitHub Pages: Settings → Pages → Deploy from branch main / root")
print(f"  4. Your dashboard will be live at:  https://YOUR-USERNAME.github.io/ea-kpi-dashboard/")
print(f"\nFor future data refreshes:")
print(f"  - Upload your new Excel as  data/latest.xlsx  in GitHub")
print(f"  - GitHub Actions will automatically rebuild index.html")
