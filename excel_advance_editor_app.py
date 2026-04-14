import streamlit as st
import pandas as pd
import io
import zipfile
import base64
from datetime import datetime
from utils import (
    read_excel_with_header,
    df_to_excel_bytes,
    df_to_csv_bytes,
    show_preview,
    show_metrics,
    download_buttons,
)

st.set_page_config(
    page_title="Python Powered Excel",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ═══════════════════════════════════════════════════════════════════════════
#  FULL BLACK DOPAMINE THEME
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

/* ── GLOBAL BLACK BASE ── */
*, *::before, *::after { box-sizing: border-box; }
html, body { background: #000000 !important; }
.stApp, .stApp > * { background: #000000 !important; font-family: 'Inter', sans-serif; }
[data-testid="stAppViewContainer"] { background: #000000 !important; }
[data-testid="stMain"], [data-testid="block-container"] { background: #000000 !important; }
section[data-testid="stSidebar"] ~ div { background: #000000 !important; }

/* ── SIDEBAR ── */
[data-testid="stSidebar"] {
    background: #0a0a0a !important;
    border-right: 1px solid #1e1e1e !important;
    padding: 0 !important;
}
[data-testid="stSidebar"] > div:first-child { padding: 20px 16px 20px 16px !important; }

.sidebar-brand {
    text-align: center;
    padding: 18px 10px 14px;
    margin-bottom: 6px;
}
.sidebar-brand .brand-icon {
    font-size: 40px; display: block; margin-bottom: 6px;
    filter: drop-shadow(0 0 12px rgba(139,92,246,0.7));
}
.sidebar-brand .brand-title {
    font-size: 30px; font-weight: 900; color: #ffffff;
    letter-spacing: -0.3px; line-height: 1;
}
.sidebar-brand .brand-sub {
    font-size: 18px; font-weight: 800; color: #555; margin-top: 4px;
    text-transform: uppercase; letter-spacing: 1.5px;
}

.sidebar-divider {
    border: none; border-top: 2px solid #1a1a1a; margin: 22px 0;
}

/* ── SIDEBAR RADIO (equal-size boxes) ── */
[data-testid="stSidebar"] .stRadio > label { display: none !important; }
[data-testid="stSidebar"] .stRadio > div {
    display: flex !important;
    flex-direction: column !important;
    gap: 6px !important;
}
[data-testid="stSidebar"] .stRadio > div > label {
    display: flex !important;
    align-items: center !important;
    width: 100% !important;
    min-height: 48px !important;
    height: 48px !important;
    background: #111111 !important;
    border: 1.5px solid #222222 !important;
    border-radius: 10px !important;
    padding: 0 14px !important;
    font-size: 13px !important;
    font-weight: 700 !important;
    color: #aaaaaa !important;
    cursor: pointer;
    transition: all 0.15s ease;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
[data-testid="stSidebar"] .stRadio > div > label:hover {
    background: #1a1a2e !important;
    border-color: #f5f7fa !important;
    color: #ffffff !important;
    transform: translateX(3px);
}
[data-testid="stSidebar"] .stRadio > div > label[data-checked="true"],
[data-testid="stSidebar"] .stRadio > div > label:has(input:checked) {
    background: linear-gradient(135deg, #1e0a4e 0%, #2d1280 100%) !important;
    border-color: #7c3aed !important;
    color: #ffffff !important;
    box-shadow: 0 0 0 1px #7c3aed, 0 4px 16px rgba(124,58,237,0.3) !important;
}
[data-testid="stSidebar"] input[type="radio"] { display: none !important; }

.sidebar-footer {
    margin-top: 14px;
    text-align: center;
    font-size: 22px;
    color: #f5f7fa;
    font-weight: 600;
    letter-spacing: 0.5px;
}

/* ── MAIN CONTENT AREA ── */
[data-testid="block-container"] {
    padding: 28px 40px !important;
    max-width: 1100px !important;
}

/* ── SECTION HEADER ── */
.sec-header {
    display: flex; align-items: center; gap: 18px;
    margin-bottom: 30px;
}
.sec-header-icon {
    font-size: 32px;
    filter: drop-shadow(0 0 10px rgba(139,92,246,0.8));
}
.sec-header-text { flex: 1; }
.sec-header-title {
    font-size: 50px; font-weight: 900; color: #ffffff;
    line-height: 1; letter-spacing: -0.5px;
}
.sec-header-sub { font-size: 20px; color: #f5f7fa; margin-top: 3px; font-weight: 800; }
.sec-header-line {
    height: 2px;
    background: linear-gradient(90deg, #7c3aed, #06b6d4, transparent);
    border-radius: 2px;
    margin-bottom: 24px;
    margin-top: -16px;
}

/* ── OP BOX ── */
.op-box {
    background: #f2f2f2;
    border: 2px solid #1e1e1e;
    border-left: 6px solid #7c3aed;
    border-bottom: 4px solid #7c3aed;
    border-radius: 12px;
    padding: 18px 22px;
    margin-bottom: 22px;
}
.op-box-title { font-size: 18px; font-weight: 800; color: #020303; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 5px; }
.op-box-desc { font-size: 14px; color: #020303; line-height: 1.6; }
.op-box-desc b { color: #020303; }

/* ── STEP CARD ── */
.step-card {
    background: #e4e6eb;
    border: 5px solid #1e1e1e;
    border-radius: 255px;
    padding: 18px 22px;
    margin-bottom: 14px;
}
.step-label {
    font-size: 11px; font-weight: 800; color: #e4e6eb;
    text-transform: uppercase; letter-spacing: 1.5px; margin-bottom: 8px;
}

/* ── METRIC CHIPS ── */
.metric-row { display: flex; gap: 10px; flex-wrap: wrap; margin: 16px 0; }
.metric-chip {
    background: #0d0d0d;
    border: 1px solid #1e1e1e;
    border-radius: 8px; padding: 10px 18px;
    font-size: 13px; font-weight: 800;
}
.metric-chip.rows  { color: #a78bfa; border-color: #2d1a5e; }
.metric-chip.cols  { color: #34d399; border-color: #0d3326; }
.metric-chip.size  { color: #60a5fa; border-color: #0d1f3c; }

/* ── MAP BOX (join key) ── */
.map-box {
    background: #0d0d0d;
    border: 1px solid #1e1e2e;
    border-top: 3px solid #7c3aed;
    border-radius: 10px; padding: 16px 18px; margin: 10px 0;
}
.map-box-title { font-size: 12px; font-weight: 800; color: #7c3aed; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 10px; }

/* ── FILE UPLOADER ── */
[data-testid="stFileUploader"] {
    background: #0f0f0f !important;
    border: 2px solid #e8f0fc !important;
    border-radius: 12px !important;
    transition: border-color 0.3s;
}
[data-testid="stFileUploader"]:hover { border-color: #7c3aed !important; }
[data-testid="stFileUploader"] * { color: #e8f0fc !important;font-size: 20px }
[data-testid="stFileUploader"] small { color: #e8f0fc !important; font-size: 16px !important; }

/* ── ALL LABELS ── */
label, .stSelectbox > label, .stMultiSelect > label,
.stNumberInput > label, .stFileUploader > label,
[data-testid="stWidgetLabel"] {
    color: #cccccc !important;
    font-weight: 800 !important;
    font-size: 13px !important;
    text-transform: uppercase !important;
    letter-spacing: 0.8px !important;
    margin-bottom: 6px !important;
}

/* ── SELECTBOX / MULTISELECT ── */
.stSelectbox [data-baseweb="select"] > div,
.stMultiSelect [data-baseweb="select"] > div {
    background: #111111 !important;
    border: 1.5px solid #2a2a2a !important;
    border-radius: 8px !important;
    min-height: 42px !important;
}
.stSelectbox [data-baseweb="select"] > div:hover,
.stMultiSelect [data-baseweb="select"] > div:hover {
    border-color: #7c3aed !important;
}
.stSelectbox [data-baseweb="select"] span,
.stSelectbox [data-baseweb="select"] div,
.stMultiSelect [data-baseweb="select"] span,
.stMultiSelect [data-baseweb="select"] div {
    color: #ffffff !important; font-weight: 600 !important; font-size: 14px !important;
    background: transparent !important;
}
/* Dropdown menu */
[data-baseweb="menu"] {
    background: #111111 !important;
    border: px solid #2a2a2a !important;
    border-radius: 10px !important;
}
[data-baseweb="menu"] li {
    color: #e6e9ed !important; font-weight: 600 !important;
    font-size: 13px !important;
    background: #111111 !important;
}
[data-baseweb="menu"] li:hover {
    background: #1e0a4e !important; color: #ffffff !important;
}
/* Multi-select tags */
[data-baseweb="tag"] {
    background: #2d1280 !important;
    border-radius: 6px !important;
}
[data-baseweb="tag"] span { color: #e2d9f3 !important; font-weight: 700 !important; }

/* ── NUMBER INPUT ── */
.stNumberInput input {
    background: #111111 !important; color: #ffffff !important;
    font-weight: 700 !important; font-size: 20px !important;
    border: 2.5px solid #e6e9ed !important; border-radius: 12px !important;
}
.stNumberInput input:focus { border-color: #7c3aed !important; }
.stNumberInput button {
    background: #1a1a1a !important; border-color: #2a2a2a !important;
    color: #ffffff !important;
}

/* ── PROCESS BUTTON ── */
.stButton > button {
    background: linear-gradient(135deg, #7c3aed 0%, #4f46e5 50%, #06b6d4 100%) !important;
    color: #ffffff !important;
    border-radius: 10px !important;
    font-weight: 900 !important;
    font-size: 15px !important;
    border: none !important;
    padding: 13px 32px !important;
    letter-spacing: 0.3px !important;
    width: 100% !important;
    transition: all 0.2s !important;
    box-shadow: 0 4px 20px rgba(124,58,237,0.4) !important;
    text-transform: uppercase !important;
}
.stButton > button:hover {
    box-shadow: 0 6px 28px rgba(124,58,237,0.65) !important;
    transform: translateY(-1px) !important;
}
.stButton > button:active { transform: translateY(0px) !important; }

/* ── DOWNLOAD BUTTONS ── */
.stDownloadButton > button {
    background: #111111 !important;
    border: 1.5px solid #2a2a2a !important;
    color: #e6e9ed !important;
    border-radius: 10px !important;
    font-weight: 800 !important;
    font-size: 18px !important;
    padding: 11px 24px !important;
    width: 100% !important;
    transition: all 0.15s !important;
}
.stDownloadButton > button:hover {
    background: #1a0a3e !important;
    border-color: #7c3aed !important;
    box-shadow: 0 0 16px rgba(124,58,237,0.3) !important;
}

/* ── DATAFRAME ── */
[data-testid="stDataFrame"] {
    border-radius: 10px; overflow: hidden;
    border: 1px solid #1e1e1e !important;
}
[data-testid="stDataFrame"] thead th {
    background: #0d0d0d !important; color: #a78bfa !important;
    font-weight: 900 !important; font-size: 14px !important;
    text-transform: uppercase; letter-spacing: 0.5px;
}
[data-testid="stDataFrame"] tbody td {
    color: #cccccc !important; font-weight: 500 !important;
    background: #080808 !important;
}
[data-testid="stDataFrame"] tbody tr:nth-child(even) td { background: #0d0d0d !important; }

/* ── EXPANDER ── */
.stExpander {
    background: #e6e9ed !important;
    border: 1.5px solid #1e1e1e !important;
    border-radius: 2px !important;
}
.stExpander summary {
    font-weight: 800 !important; color: #aaaaaa !important;
    font-size: 13px !important; background: #0d0d0d !important;
}
.stExpander summary:hover { color: #ffffff !important; }
.stExpander [data-testid="stExpanderDetails"] {
    background: #080808 !important;
    border-top: 1px solid #1a1a1a !important;
}

/* ── ALERTS ── */
.stAlert { border-radius: 10px !important; font-weight: 700 !important; font-size: 13px !important; }
[data-testid="stNotification"] { border-radius: 10px !important; }
/* Info */
div[data-testid="stAlertContainer"][data-baseweb="notification"][kind="info"] {
    background: #071825 !important; border-color: #0369a1 !important;
}
/* Success */
div[data-testid="stAlertContainer"][data-baseweb="notification"][kind="success"] {
    background: #071a0f !important; border-color: #15803d !important;
}
/* Warning */
div[data-testid="stAlertContainer"][data-baseweb="notification"][kind="warning"] {
    background: #1a1200 !important; border-color: #a16207 !important;
}
/* Error */
div[data-testid="stAlertContainer"][data-baseweb="notification"][kind="error"] {
    background: #1a0707 !important; border-color: #b91c1c !important;
}

/* ── SECTION SEPARATOR ── */
.sep { border: none; border-top: 1px solid #1a1a1a; margin: 20px 0; }

/* ── SUCCESS GLOW BOX ── */
.result-header {
    background: #071a0f;
    border: 1px solid #15803d;
    border-radius: 10px;
    padding: 14px 20px;
    margin: 16px 0;
    display: flex; align-items: center; gap: 12px;
}
.result-header-icon { font-size: 24px; }
.result-header-text { font-size: 15px; font-weight: 800; color: #4ade80; }
.result-header-sub  { font-size: 12px; color: #555; margin-top: 2px; }

/* Markdown text in dark */
.stMarkdown, .stMarkdown p, .stMarkdown li, .stMarkdown span {
    color: #888888 !important;
}
.stMarkdown h3, .stMarkdown h4 { color: #cccccc !important; font-weight: 800 !important; }

/* Caption */
.stCaption { color: #444 !important; font-size: 11px !important; }

/* Scrollbar */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #0a0a0a; }
::-webkit-scrollbar-thumb { background: #2a2a2a; border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: #7c3aed; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════
TOOLS = {
    "Merge Excels":         "merge_flat",
    "Split into Excels":    "split_col",
    "Delete Columns":       "delete_cols",
    "Merge Workbook":       "merge_workbook",
    "Split Workbook":       "split_workbook",
    "Merge All Sheets":     "append_sheets",
    "Join / Merge":         "pandas_merge",
}

with st.sidebar:
    st.markdown("""
    <div class="sidebar-brand">
      <span class="brand-icon">📑</span>
      <div class="brand-title">Python Powered Excel</div>
    </div>
    <hr class="sidebar-divider">
    """, unsafe_allow_html=True)


    choice = st.radio("tool", list(TOOLS.keys()), label_visibility="collapsed")
    tool = TOOLS[choice]

    st.markdown("""
    <hr class="sidebar-divider">
    <div class="sidebar-footer">UPLOAD · PREVIEW · PROCESS · DOWNLOAD</div>
    """, unsafe_allow_html=True)




# ═══════════════════════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════════════════════
TOOL_META = {
    "merge_flat":     ("", "Merge Files",      "Helps in merging the multiple excel files in one single file"),
    "split_col":      ("", "Split into Many Excel",              "Splits into multiple excels based on the input values"),
    "delete_cols":    ("", "Delete Columns",               "Remove the Columns by selecting from dropdown Download only necessary files"),
    "merge_workbook": ("", "Merge Files into Workbook",   "Each file becomes a sheet inside one Excel workbook"),
    "split_workbook": ("", "Split Workbook to Excels",       "Each sheet  gets exported as own file"),
    "append_sheets":  ("", "Append All Sheets to one Sheet",      "Stack all tabs from multiple workbooks into one sheet"),
    "pandas_merge":   ("", "SQL-Style Join / Merge",       "Match rows across two files by key columns (LEFT / INNER / OUTER)"),
}

def page_header(key):
    icon, title, sub = TOOL_META[key]
    st.markdown(f"""
    <div class="sec-header">
      <span class="sec-header-icon">{icon}</span>
      <div class="sec-header-text">
        <div class="sec-header-title">{title}</div>
        <div class="sec-header-sub">{sub}</div>
      </div>
    </div>
    <div class="sec-header-line"></div>
    """, unsafe_allow_html=True)

def op_box(desc):
    st.markdown(f"""
    <div class="op-box">
      <div class="op-box-title">ℹ️ How will it help</div>
      <div class="op-box-desc">{desc}</div>
    </div>""", unsafe_allow_html=True)

def step_wrap(label, content_fn):
    st.markdown(f'<div class="step-card"><div class="step-label">{label}</div>', unsafe_allow_html=True)
    result = content_fn()
    st.markdown('</div>', unsafe_allow_html=True)
    return result

def require_files(files, n=2, label="files"):
    if not files:
        st.info(f"👆  Upload {n} or more {label} to get started.")
        return False
    if len(files) < n:
        st.error(f"❌  You uploaded **{len(files)}** file — need **at least {n}**. Add more above.")
        return False
    return True

def sep():
    st.markdown('<hr class="sep">', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
#  TOOL 1 — Merge Files → One Sheet
# ═══════════════════════════════════════════════════════════════════════════
if tool == "merge_flat":
    page_header("merge_flat")
    op_box("Upload <b>two or more</b> Excel / CSV files. All rows are stacked into one master sheet — like copy-pasting from many files into one.")

    files = st.file_uploader("📂  Upload Excel / CSV files", type=["xlsx","xls","csv"],
                              accept_multiple_files=True, key="mf_upload")
    if not require_files(files, 2):
        st.stop()

    header_row = st.number_input("📌  Header row number  (1 = first row)",
                                  min_value=1, max_value=20, value=1, key="mf_hdr")
    sep()

    previews, all_ok = {}, True
    # for f in files:
    #     try:
    #         df_p = read_excel_with_header(f, header_row)
    #         previews[f.name] = df_p
    #         with st.expander(f" {f.name}   —   {len(df_p):,} rows  ×  {len(df_p.columns)} cols"):
    #             show_preview(df_p)
    #     except Exception as e:
    #         st.error(f"❌  Cannot read **{f.name}**: {e}")
    #         all_ok = False
    for f in files:
        try:
            df_p = read_excel_with_header(f, header_row)
            previews[f.name] = df_p

            # Header
            st.markdown(f"### {f.name}")
            st.caption(f"{len(df_p):,} rows × {len(df_p.columns)} columns")

            # Always show top 5
            st.dataframe(df_p.head(5), use_container_width=True)

        except Exception as e:
            st.error(f"❌ Cannot read **{f.name}**: {e}")
            all_ok = False
        

    if not all_ok:
        st.warning("⚠️  Fix the errors above, then try again.")
        st.stop()

    sep()
    if st.button("🔗  MERGE ALL FILES", key="mf_btn"):
        try:
            merged = pd.concat(list(previews.values()), ignore_index=True)
            st.success(f"✅  Merged {len(files)} files — **{len(merged):,} total rows**")
            show_metrics(merged)
            st.dataframe(merged, use_container_width=True)
            download_buttons(merged, "merge_all_files")
        except Exception as e:
            st.error(f"❌  Merge failed: {e}")


# ═══════════════════════════════════════════════════════════════════════════
#  TOOL 2 — Split by Column
# ═══════════════════════════════════════════════════════════════════════════
elif tool == "split_col":
    page_header("split_col")
    op_box("Pick any column (e.g. <b>District</b>, <b>State</b>, <b>Category</b>) and get a <b>separate file per unique value</b>. Download files one-by-one or all together as a ZIP.")

    file = st.file_uploader("📂  Upload one Excel / CSV file", type=["xlsx","xls","csv"], key="sc_upload")
    if not file:
        st.info("👆  Upload a file to get started.")
        st.stop()

    try:
        header_row = st.number_input("📌  Header row number", min_value=1, max_value=20, value=1, key="sc_hdr")
        df = read_excel_with_header(file, header_row)
    except Exception as e:
        st.error(f"❌  Cannot read file: {e}")
        st.stop()

    with st.expander("Preview of uploaded file (top 5 rows)", expanded=True):
        show_preview(df)
    sep()

    split_col = st.selectbox("Column to SPLIT BY", ["— choose a column —"] + list(df.columns), key="sc_col")
    if split_col == "— choose a column —":
        st.warning("⚠️  Select a column above to continue.")
        st.stop()

    unique_vals = df[split_col].dropna().unique()
    st.info(f"🗂️  **{len(unique_vals)} unique values** found in '{split_col}' → will create **{len(unique_vals)} files**")
    sep()

    if st.button("SPLIT NOW", key="sc_btn"):
        if len(unique_vals) == 0:
            st.error("❌  No data found in selected column.")
            st.stop()
        try:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w") as zf:
                for val in unique_vals:
                    grp = df[df[split_col] == val].reset_index(drop=True)
                    safe = str(val)
                    for ch in r'/\:*?[]': safe = safe.replace(ch, "_")
                    with st.expander(f" {split_col} = {val}   ({len(grp):,} rows)", expanded=True):
                        show_metrics(grp)
                        st.dataframe(grp.head(5), use_container_width=True)
                        download_buttons(grp, f"split_by_{split_col}_{safe}")
                    zf.writestr(f"{safe}.xlsx", df_to_excel_bytes(grp))
            zip_buf.seek(0)
            sep()
            ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            st.success(f"✅  Split complete! {len(unique_vals)} files ready.")
            st.download_button(" Download ALL Files as ZIP", zip_buf, f"split_files_{ts}.zip", "application/zip")
        except Exception as e:
            st.error(f"❌  Split failed: {e}")


# ═══════════════════════════════════════════════════════════════════════════
#  TOOL 3 — Delete Columns
# ═══════════════════════════════════════════════════════════════════════════
elif tool == "delete_cols":
    page_header("delete_cols")
    op_box("Select columns to remove — checkbox style. Preview the cleaned file and download instantly.")

    file = st.file_uploader("📂  Upload one Excel / CSV file", type=["xlsx","xls","csv"], key="dc_upload")
    if not file:
        st.info("👆  Upload a file to get started.")
        st.stop()

    try:
        header_row = st.number_input("📌  Header row number", min_value=1, max_value=20, value=1, key="dc_hdr")
        df = read_excel_with_header(file, header_row)
    except Exception as e:
        st.error(f"❌  Cannot read file: {e}")
        st.stop()

    with st.expander("Preview uploaded file (top 5 rows)", expanded=True):
        show_preview(df)
    sep()

    to_delete = st.multiselect("Select columns to DELETE", df.columns.tolist(), key="dc_cols")
    if not to_delete:
        st.warning("⚠️  Select at least one column to delete from the list above.")
        st.stop()

    st.info(f"Will remove **{len(to_delete)}** column(s): {', '.join(str(c) for c in to_delete)}")
    sep()

    if st.button("DELETE SELECTED COLUMNS", key="dc_btn"):
        try:
            result = df.drop(columns=to_delete)
            st.success(f"✅  Deleted {len(to_delete)} column(s). **{len(result.columns)} columns** remaining.")
            show_metrics(result)
            st.dataframe(result, use_container_width=True)
            download_buttons(result, "delete_columns")
        except Exception as e:
            st.error(f"❌  Delete failed: {e}")


# ═══════════════════════════════════════════════════════════════════════════
#  TOOL 4 — Merge Files → One Workbook
# ═══════════════════════════════════════════════════════════════════════════
elif tool == "merge_workbook":
    page_header("merge_workbook")
    op_box("Each uploaded file becomes a <b>separate sheet (tab)</b> inside one Excel workbook. Perfect for monthly or regional reports.")

    files = st.file_uploader("📂  Upload Excel / CSV files (2 or more)", type=["xlsx","xls","csv"],
                              accept_multiple_files=True, key="mw_upload")
    if not require_files(files, 2):
        st.stop()

    header_row = st.number_input("📌  Header row number", min_value=1, max_value=20, value=1, key="mw_hdr")
    sep()

    previews, all_ok = {}, True
    for f in files:
        try:
            df_p = read_excel_with_header(f, header_row)
            previews[f.name] = df_p
            with st.expander(f"{f.name}   —   {len(df_p):,} rows  ×  {len(df_p.columns)} cols",expanded=True):
                show_preview(df_p)
        except Exception as e:
            st.error(f"❌  Cannot read **{f.name}**: {e}")
            all_ok = False

    if not all_ok:
        st.warning("⚠️  Fix errors above before creating workbook.")
        st.stop()

    st.info(f"**{len(files)} sheets** will be created in the output workbook.")
    sep()

    if st.button("CREATE WORKBOOK", key="mw_btn"):
        try:
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                for fname, df in previews.items():
                    sname = fname[:31]
                    for ch in r'/\*[]:?': sname = sname.replace(ch, "_")
                    df.to_excel(writer, sheet_name=sname, index=False)
            out.seek(0)
            ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            st.success(f"✅  Workbook with **{len(files)} sheets** ready!")
            st.download_button("Download Workbook (.xlsx)", out, f"merge_workbook_{ts}.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"❌  Failed: {e}")


# ═══════════════════════════════════════════════════════════════════════════
#  TOOL 5 — Split Workbook → Files
# ═══════════════════════════════════════════════════════════════════════════
elif tool == "split_workbook":
    page_header("split_workbook")
    op_box("Upload a workbook with <b>multiple sheets</b>. Each sheet is saved as its own Excel file and bundled into a ZIP download.")

    file = st.file_uploader("📂  Upload a multi-sheet Excel file", type=["xlsx","xls"], key="sw_upload")
    if not file:
        st.info("👆  Upload a file to get started.")
        st.stop()

    try:
        header_row = st.number_input("📌  Header row number", min_value=1, max_value=20, value=1, key="sw_hdr")
        all_sheets = pd.read_excel(file, sheet_name=None, header=header_row - 1)
    except Exception as e:
        st.error(f"❌  Cannot read file: {e}")
        st.stop()

    sheet_names = list(all_sheets.keys())
    st.info(f"Found **{len(sheet_names)} sheets**: {', '.join(sheet_names)}")

    sheets_to_remove = st.multiselect("Remove sheets before exporting (optional)", sheet_names, key="sw_del")
    sheets_final = [s for s in sheet_names if s not in sheets_to_remove]

    if not sheets_final:
        st.error("❌  All sheets removed — keep at least one.")
        st.stop()

    st.success(f"Will export **{len(sheets_final)} sheet(s)**: {', '.join(sheets_final)}")

    for s in sheets_final:
        with st.expander(f"Sheet: {s}   ({len(all_sheets[s]):,} rows)"):
            show_preview(all_sheets[s])
    sep()

    if st.button("SPLIT WORKBOOK", key="sw_btn"):
        try:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w") as zf:
                for s in sheets_final:
                    zf.writestr(f"{s}.xlsx", df_to_excel_bytes(all_sheets[s]))
            zip_buf.seek(0)
            ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            st.success(f"✅  Split into {len(sheets_final)} files!")
            st.download_button("Download All Sheets as ZIP", zip_buf, f"split_workbook_{ts}.zip", "application/zip")
        except Exception as e:
            st.error(f"❌  Split failed: {e}")


# ═══════════════════════════════════════════════════════════════════════════
#  TOOL 6 — Append All Sheets → One Sheet
# ═══════════════════════════════════════════════════════════════════════════
elif tool == "append_sheets":
    page_header("append_sheets")
    op_box("Upload workbooks where every sheet has the <b>same columns</b>. All data is stacked into one master sheet — great for quarterly reports in separate tabs.")

    files = st.file_uploader("📂  Upload Excel files", type=["xlsx","xls"],
                              accept_multiple_files=True, key="as_upload")
    if not require_files(files, 1, "Excel files"):
        st.stop()

    header_row = st.number_input("📌  Header row number", min_value=1, max_value=20, value=1, key="as_hdr")
    sep()

    all_dfs, all_ok = [], True
    for f in files:
        try:
            sheets = pd.read_excel(f, sheet_name=None, header=header_row - 1)
            st.info(f"📁  **{f.name}** — {len(sheets)} sheet(s): {', '.join(sheets.keys())}")
            for sname, df in sheets.items():
                with st.expander(f"{f.name} → {sname}   ({len(df):,} rows)", expanded=True):
                    show_preview(df)
                all_dfs.append(df)
        except Exception as e:
            st.error(f"❌  Cannot read **{f.name}**: {e}")
            all_ok = False

    if not all_ok:
        st.warning("⚠️  Fix errors above before appending.")
        st.stop()

    sep()
    if st.button("APPEND ALL SHEETS", key="as_btn"):
        try:
            result = pd.concat(all_dfs, ignore_index=True)
            st.success(f"✅  Appended {len(all_dfs)} sheets — **{len(result):,} total rows**!")
            show_metrics(result)
            st.dataframe(result, use_container_width=True)
            download_buttons(result, "append_all_sheets")
        except Exception as e:
            st.error(f"❌  Append failed: {e}")


# ═══════════════════════════════════════════════════════════════════════════
#  TOOL 7 — Pandas-style Join / Merge
# ═══════════════════════════════════════════════════════════════════════════
elif tool == "pandas_merge":
    page_header("pandas_merge")
    op_box("Join two files like a <b>database JOIN</b> — match rows by key columns even when column names are different across files. Choose Left, Inner, or Outer join.")

    files = st.file_uploader("📂  Upload exactly 2 Excel / CSV files", type=["xlsx","xls","csv"],
                              accept_multiple_files=True, key="pm_upload")

    if not files:
        st.info("👆  Upload 2 files to get started.")
        st.stop()
    if len(files) < 2:
        st.error(f"❌  Uploaded **{len(files)}** file — need **exactly 2**.")
        st.stop()
    if len(files) > 2:
        st.warning(f"⚠️  {len(files)} files uploaded — only the **first two** will be used.")
        files = files[:2]

    header_row = st.number_input("📌  Header row number", min_value=1, max_value=20, value=1, key="pm_hdr")
    sep()

    dfs, all_ok = {}, True
    for f in files:
        try:
            df = read_excel_with_header(f, header_row)
            dfs[f.name] = df
            with st.expander(f"{f.name}   ({len(df):,} rows  ×  {len(df.columns)} cols)", expanded=True):
                show_preview(df)
        except Exception as e:
            st.error(f"❌  Cannot read **{f.name}**: {e}")
            all_ok = False

    if not all_ok:
        st.warning("⚠️  Fix errors above before merging.")
        st.stop()

    file_names = list(dfs.keys())
    left_name, right_name = file_names[0], file_names[1]
    left_df,  right_df   = dfs[left_name], dfs[right_name]
    left_opts  = ["— choose —"] + list(left_df.columns)
    right_opts = ["— choose —"] + list(right_df.columns)

    sep()
    st.markdown("### Map Join Keys")
    st.caption("Column names can be different across files — just map them below.")

    n_keys = st.selectbox("How many key columns to join on?", [1, 2, 3], key="pm_nkeys")

    key_pairs, valid_keys = [], True
    for i in range(n_keys):
        st.markdown(f'<div class="map-box"><div class="map-box-title"> Key {i+1}</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.caption(f"FROM  →  {left_name}")
            lk = st.selectbox("left col", left_opts, key=f"pm_lk_{i}", label_visibility="collapsed")
        with c2:
            st.caption(f"FROM  →  {right_name}")
            rk = st.selectbox("right col", right_opts, key=f"pm_rk_{i}", label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)
        if lk == "— choose —" or rk == "— choose —":
            valid_keys = False
        else:
            key_pairs.append((lk, rk))

    if not valid_keys:
        st.warning("⚠️  Select all join key columns before proceeding.")
        st.stop()

    sep()
    HOW_OPTIONS = [
        "LEFT — Keep ALL rows from left file  (missing right = blank)",
        "INNER — Keep ONLY rows that match in BOTH files",
        "OUTER — Keep ALL rows from BOTH files  (fill missing with blank)",
    ]
    HOW_MAP = {
        HOW_OPTIONS[0]: "left",
        HOW_OPTIONS[1]: "inner",
        HOW_OPTIONS[2]: "outer",
    }
    how_choice = st.selectbox(" Join Type", HOW_OPTIONS, key="pm_how")
    sep()

    if st.button("RUN MERGE", key="pm_btn"):
        try:
            rename_map = {rk: lk for lk, rk in key_pairs if lk != rk}
            right_work = right_df.rename(columns=rename_map)
            join_cols  = [lk for lk, rk in key_pairs]
            result     = pd.merge(left_df, right_work, on=join_cols, how=HOW_MAP[how_choice])
            st.success(f"✅  Merge complete — **{len(result):,} rows** in result!")
            show_metrics(result)
            st.dataframe(result, use_container_width=True)
            download_buttons(result, "join_merge")
        except Exception as e:
            st.error(f"❌  Merge failed: {e}")
            st.info("💡  Tip: Make sure join key columns have matching values across both files.")
            
            
