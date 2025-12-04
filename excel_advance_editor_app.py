import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
from datetime import datetime

# =========================
# Page Config
# =========================
st.set_page_config(page_title="📊 Excel advance editor", layout="wide")
st.title("📊 Advance Excel tools powered by Python")

# =========================
# Helpers
# =========================
def safe_filename(value) -> str:
    """Sanitize values for use in filenames."""
    s = str(value)
    return "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in s).strip().replace(" ", "_")

def load_dataframe_with_sheet_picker(uploaded_file, header_row: int, key_prefix: str):
    """
    Load a CSV or Excel file into a DataFrame.
    If Excel has multiple sheets, show a sheet picker.
    header_row is 1-based in the UI; convert to 0-based for pandas.
    """
    try:
        name = uploaded_file.name.lower()
        if name.endswith(".csv"):
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, header=header_row - 1)
        else:
            # Inspect sheets
            uploaded_file.seek(0)
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names

            if len(sheet_names) > 1:
                sheet = st.selectbox(
                    f"Select sheet for **{uploaded_file.name}**",
                    sheet_names,
                    key=f"{key_prefix}_sheet_{uploaded_file.name}"
                )
            else:
                sheet = sheet_names[0]

            # Read the chosen sheet
            uploaded_file.seek(0)
            return pd.read_excel(uploaded_file, sheet_name=sheet, header=header_row - 1)
    except Exception as e:
        st.error(f"❌ Error reading {uploaded_file.name}: {e}")
        return None

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Return an Excel file (bytes) from a DataFrame."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# =========================
# Tabs
# =========================
tab1, tab2, tab3, tab4 = st.tabs([
    "📂 Combine Files",
    "✂️ Split by Column",
    "🗑 Drop Columns",
    "👁 View Selected Columns"
])

# ==========================================================
# Tab 1: Combine Files (min 2 files)
# ==========================================================
with tab1:
    uploaded_files = st.file_uploader(
        "Upload multiple Excel/CSV files to combine",
        type=["xlsx", "xls", "xlsm", "xlsb", "csv"],
        accept_multiple_files=True,
        key="combine_files"
    )

    if uploaded_files:
        if len(uploaded_files) < 2:
            st.warning("⚠️ Please upload at least two files to combine.")
        else:
            header_row = st.number_input(
                "Select header row (1 = first row)",
                min_value=1, value=1, key="combine_header"
            )

            dfs = []
            for i, file in enumerate(uploaded_files):
                df = load_dataframe_with_sheet_picker(file, header_row, key_prefix=f"combine_{i}")
                if df is not None:
                    dfs.append(df)
                    st.write(f"**Preview: {file.name}** (Rows: {df.shape[0]}, Columns: {df.shape[1]})")
                    st.dataframe(df.head())

            if dfs and len(dfs) >= 2 and st.button("🔗 Combine All"):
                combined_df = pd.concat(dfs, ignore_index=True)
                st.success(f"✅ Combined {len(dfs)} files/sheets into one dataframe.")
                st.dataframe(combined_df.head(10))
            
                # Create timestamp
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                filename = f"combined_data_{timestamp}.xlsx"
            
                # Provide download button with renamed file
                st.download_button(
                    "⬇️ Download Combined Excel",
                    to_excel_bytes(combined_df),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# ==========================================================
# Tab 2: Split by Column (with preview toggle + ZIP)
# ==========================================================
with tab2:
    uploaded_file = st.file_uploader(
        "Upload one Excel/CSV file to split by a column",
        type=["xlsx", "xls", "xlsm", "xlsb", "csv"],
        key="split_file"
    )

    if uploaded_file:
        header_row = st.number_input(
            "Select header row (1 = first row)",
            min_value=1, value=1, key="split_header"
        )
        enable_preview = st.checkbox("Enable Preview Before Processing", value=True, key="split_preview")

        df = load_dataframe_with_sheet_picker(uploaded_file, header_row, key_prefix="split")
        if df is not None and not df.empty:
            if enable_preview:
                st.markdown("### Preview of uploaded file")
                st.dataframe(df.head())

            split_column = st.selectbox("Select column to split by", df.columns, key="split_column")
            if split_column:
                groups = dict(tuple(df.groupby(split_column, dropna=True)))
                st.info(f"Found **{len(groups)}** group(s) in “{split_column}”.")

                # Individual group previews & downloads
                for name, group in groups.items():
                    if enable_preview:
                        st.markdown(f"#### Group: `{split_column} = {name}` (Rows: {len(group)})")
                        st.dataframe(group.head())

                    fname = f"{safe_filename(split_column)}_{safe_filename(name)}.xlsx"
                    st.download_button(
                        f"⬇️ Download {split_column} = {name}",
                        to_excel_bytes(group),
                        fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_split_{fname}"
                    )

                # ZIP all groups
                if st.button("⬇️ Download All Groups as ZIP"):
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zf:
                        for name, group in groups.items():
                            excel_bytes = to_excel_bytes(group)
                            fname = f"{safe_filename(split_column)}_{safe_filename(name)}.xlsx"
                            zf.writestr(fname, excel_bytes)
                    st.download_button(
                        "⬇️ Download ZIP",
                        zip_buffer.getvalue(),
                        "split_files.zip",
                        mime="application/zip",
                        key="dl_split_zip"
                    )

# ==========================================================
# Tab 3: Drop Columns (multi-file, preview toggle)
# ==========================================================
with tab3:
    uploaded_files_drop = st.file_uploader(
        "Upload one or more Excel/CSV files to drop columns",
        type=["xlsx", "xls", "xlsm", "xlsb", "csv"],
        accept_multiple_files=True,
        key="drop_files"
    )

    if uploaded_files_drop:
        header_row = st.number_input(
            "Select header row (1 = first row)",
            min_value=1, value=1, key="drop_header"
        )
        enable_preview_drop = st.checkbox("Enable Preview Before Dropping", value=True, key="drop_preview")

        first_df = load_dataframe_with_sheet_picker(uploaded_files_drop[0], header_row, key_prefix="drop_first")
        if first_df is None or first_df.empty:
            st.warning("The first file could not be read or is empty.")
        else:
            st.markdown("### First 5 rows of the first uploaded file")
            st.dataframe(first_df.head())  # Show first 5 rows after uploading

            columns_to_drop = st.multiselect("Select columns to drop", list(first_df.columns), key="drop_cols")

            if columns_to_drop:
                for i, file in enumerate(uploaded_files_drop):
                    df = load_dataframe_with_sheet_picker(file, header_row, key_prefix=f"drop_{i}")
                    if df is None:
                        continue

                    df_dropped = df.drop(columns=[c for c in columns_to_drop if c in df.columns], errors="ignore")
                    st.markdown(
                        f"### Processed File: `{file.name}` "
                        f"(Rows: {df_dropped.shape[0]}, Cols: {df_dropped.shape[1]})"
                    )
                    st.dataframe(df_dropped.head())  # Show first 5 rows after editing

                    out_name = f"{file.name.rsplit('.', 1)[0]}_dropped.xlsx"
                    st.download_button(
                        f"⬇️ Download {file.name} without Selected Columns",
                        to_excel_bytes(df_dropped),
                        out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_drop_{i}"
                    )

# ==========================================================
# Tab 4: View Selected Columns (fixed for full download)
# ==========================================================
with tab4:
    uploaded_file_view = st.file_uploader(
        "Upload one Excel/CSV file to view selected columns",
        type=["xlsx", "xls", "xlsm", "xlsb", "csv"],
        key="view_file"
    )

    if uploaded_file_view:
        header_row = st.number_input(
            "Select header row (1 = first row)",
            min_value=1, value=1,
            key="view_header"
        )
        df = load_dataframe_with_sheet_picker(uploaded_file_view, header_row, key_prefix="view")
        if df is not None and not df.empty:
            st.markdown("### First 5 rows of uploaded file")
            st.dataframe(df.head())  # Show first 5 rows after uploading

            selected_columns = st.multiselect("Select columns to view", list(df.columns), key="view_cols")
            num_rows = st.number_input("Number of rows to display", min_value=1, value=5, key="view_rows")

            # Preview limited rows
            if selected_columns:
                view_df = df[selected_columns].head(num_rows)
                full_df = df[selected_columns]   # Full data for download
            else:
                view_df = df.head(num_rows)
                full_df = df                   # Full data for download

            st.markdown("### Preview after selecting columns")
            st.dataframe(view_df)

            # Download full dataframe of selected columns
            st.download_button(
                "⬇️ Download Excel",
                to_excel_bytes(full_df),
                "view_selected_columns.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_view"
            )
