#!/usr/bin/env python3
"""
inject_data.py  –  Reads data/data.json and splices it into index.html,
replacing the block between /* __DATA_START__ */ and /* __DATA_END__ */.
Run this after excel_to_json.py.
"""
import json, re, os

DATA_PATH = "data/data.json"
HTML_PATH = "index.html"

print(f"Loading {DATA_PATH} ...")
with open(DATA_PATH, "r", encoding="utf-8") as f:
    data_obj = json.load(f)

data_compact = json.dumps(data_obj, separators=(",", ":"), ensure_ascii=False)

print(f"Reading {HTML_PATH} ...")
with open(HTML_PATH, "r", encoding="utf-8") as f:
    html = f.read()

if "/* __DATA_START__ */" not in html:
    print("ERROR: markers not found in index.html.")
    print("Run setup.py first to prepare the HTML file.")
    raise SystemExit(1)

PATTERN     = r"/\* __DATA_START__ \*/const DATA=.*?;/\* __DATA_END__ \*/"
REPLACEMENT = f"/* __DATA_START__ */const DATA={data_compact};/* __DATA_END__ */"

new_html = re.sub(PATTERN, REPLACEMENT, html, flags=re.DOTALL, count=1)

with open(HTML_PATH, "w", encoding="utf-8") as f:
    f.write(new_html)

latest = data_obj.get("fixed", {}).get("ts", {}).get("latest", "unknown")
print(f"index.html updated  ({os.path.getsize(HTML_PATH):,} bytes)")
print(f"Latest data date: {latest}")
