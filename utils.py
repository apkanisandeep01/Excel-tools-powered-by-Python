import io
from datetime import datetime
import pandas as pd
import streamlit as st


def read_excel_with_header(file, header_row: int) -> pd.DataFrame:
    """Read excel or csv with custom header row (1-indexed)."""
    name = getattr(file, "name", "")
    hdr = header_row - 1
    if name.endswith(".csv"):
        return pd.read_csv(file, header=hdr)
    return pd.read_excel(file, header=hdr)


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def show_preview(df: pd.DataFrame, n: int = 5):
    st.markdown(f"<small>Showing top {min(n, len(df))} of {len(df)} rows &nbsp;|&nbsp; {len(df.columns)} columns</small>", unsafe_allow_html=True)
    st.dataframe(df.head(n), use_container_width=True)


def show_metrics(df: pd.DataFrame):
    st.markdown(
        f"""<div class="metric-row">
            <div class="metric-chip">📏 {len(df):,} Rows</div>
            <div class="metric-chip">📐 {len(df.columns)} Columns</div>
            <div class="metric-chip">💾 ~{df.memory_usage(deep=True).sum() // 1024} KB</div>
        </div>""",
        unsafe_allow_html=True,
    )


def _make_safe_filename(stem: str) -> str:
    return "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in stem)


def _timestamped_name(stem: str, ext: str) -> str:
    safe_stem = _make_safe_filename(stem)
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return f"{safe_stem}_{ts}.{ext}"


def download_buttons(df: pd.DataFrame, stem: str):
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "📥 Download as Excel",
            df_to_excel_bytes(df),
            _timestamped_name(stem, "xlsx"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_excel_{stem}",
        )
    with c2:
        st.download_button(
            "📥 Download as CSV",
            df_to_csv_bytes(df),
            _timestamped_name(stem, "csv"),
            "text/csv",
            key=f"dl_csv_{stem}",
        )
