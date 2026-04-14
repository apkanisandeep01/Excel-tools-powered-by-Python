# 📊 Excel Toolkit — Streamlit App

A powerful, non-tech-friendly Excel processing tool with 7 operations.

## Features
| # | Tool | What it does |
|---|------|-------------|
| 1 | 🔗 Merge Files → One Sheet | Stack multiple files row-by-row into one sheet |
| 2 | ✂️ Split by Column | Split file into separate files by column values |
| 3 | 🗑️ Delete Columns | Remove unwanted columns and download cleaned file |
| 4 | 📚 Merge Files → One Workbook | Each file becomes a sheet in one workbook |
| 5 | 📤 Split Workbook → Files | Each sheet becomes its own file (ZIP download) |
| 6 | ➕ Append All Sheets → One Sheet | Stack all sheets from multiple workbooks |
| 7 | 🔀 Pandas-Style Join / Merge | Database-style LEFT / INNER / OUTER join |

## 🚀 Deploy on Streamlit Cloud (Free)

### Step 1 – Push to GitHub
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/excel-toolkit.git
git push -u origin main
```

### Step 2 – Deploy
1. Go to https://share.streamlit.io
2. Click **New app**
3. Select your GitHub repo
4. Set **Main file path** to `app.py`
5. Click **Deploy**

That's it! Your app will be live at `https://YOUR_USERNAME-excel-toolkit-app-XXXXX.streamlit.app`

## Run Locally
```bash
pip install -r requirements.txt
streamlit run app.py
```
