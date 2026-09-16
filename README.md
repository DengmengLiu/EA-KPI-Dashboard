# EA RISE Impact KPI Dashboard

A self-contained HTML dashboard hosted on GitHub Pages with automatic refresh from Excel.

---

## 🚀 First-Time Setup (do this once)

### Step 1 – Run the setup script locally

Make sure Python 3.8+ is installed. Then open a terminal in this folder and run:

```bash
python scripts/setup.py  "path/to/your/EA_RISE_Impact_KPI_Dashboard.html"
```

This will:
- Extract the data from your HTML → `data/data.json`
- Add refresh markers to `index.html`
- Copy your Excel (if found) to `data/latest.xlsx`
- Build **`ea-kpi-dashboard.zip`** ready to upload

---

### Step 2 – Create the GitHub repository

1. Go to **github.com** → click **New repository**
2. Name: `ea-kpi-dashboard`  |  Visibility: **Public**
3. Leave all other options as default → **Create repository**

---

### Step 3 – Upload the zip contents

1. Unzip `ea-kpi-dashboard.zip`
2. In your new GitHub repo, click **Add file → Upload files**
3. Drag **all the files and folders** from the unzipped folder into GitHub
4. Click **Commit changes**

---

### Step 4 – Enable GitHub Pages

1. Go to your repo → **Settings → Pages**
2. Under *Source*, select **Deploy from a branch**
3. Branch: `main` | Folder: `/ (root)`
4. Click **Save**

Your dashboard is now live at:
```
https://YOUR-USERNAME.github.io/ea-kpi-dashboard/
```
(Takes ~1 minute to deploy after the first push)

---

## 🔄 Refreshing the Dashboard (ongoing)

Every time you want to update the dashboard with fresh data:

1. Refresh your Excel file locally (from AFO or manually)
2. Go to your GitHub repo → `data/` folder
3. Click **Add file → Upload files**
4. Upload your new Excel as **`latest.xlsx`**
5. Click **Commit changes**

GitHub Actions automatically:
- Converts the Excel → `data/data.json`
- Injects the new data into `index.html`
- Commits the updated files

The dashboard refreshes within ~60–90 seconds. ✅

---

## 📁 Repository Structure

```
ea-kpi-dashboard/
├── index.html                    ← The dashboard (auto-updated by GitHub Actions)
├── data/
│   ├── data.json                 ← Auto-generated from Excel (do not edit)
│   └── latest.xlsx               ← Upload your fresh Excel here to trigger refresh
├── scripts/
│   ├── excel_to_json.py          ← Converts Excel columns → data.json
│   ├── inject_data.py            ← Splices data.json into index.html
│   └── setup.py                  ← One-time local setup helper
├── .github/
│   └── workflows/
│       └── refresh.yml           ← GitHub Actions pipeline
└── README.md
```

---

## 🛠 Adjusting the Column Mapping

If your Excel layout changes, edit `scripts/excel_to_json.py`.

The `MEASURES` list at the top defines which Excel column each KPI reads from:

```python
{"id": "compliance_check", "name": "Compliance Check", "bc": None, "col": 6},
#                                                                         ^ this is the 1-indexed Excel column
```

- **`col`** — column containing the KPI value (1 = yes, 0 = no)
- **`bc`** — column containing the baseline count (or `None` to use total account count)

Also check the constants at the top of the file:
```python
C_DATE   = 0   # column A (0-indexed)
C_REGION = 1   # column B
C_ACCID  = 2   # column C
C_NAME   = 3   # column D
C_EA     = 4   # column E
```

---

## ❓ Troubleshooting

| Problem | Fix |
|---|---|
| "No data rows found" | Check `C_DATE`, `C_REGION` in `excel_to_json.py` match your Excel columns |
| Dashboard shows old data | Check Actions tab in GitHub — look for red ❌ and click to see error |
| Excel not triggering refresh | Make sure you upload the file as exactly `data/latest.xlsx` |
| GitHub Actions fails with "Access denied" | Go to Settings → Actions → General → set Workflow permissions to **Read and write** |
