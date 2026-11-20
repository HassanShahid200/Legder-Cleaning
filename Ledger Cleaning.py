"""Streamlit web app to clean accounting exports with selectable, rearrangeable columns."""

import io
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

APP_TITLE = "Accounting Data Cleaner"
DEFAULT_COLUMNS = [
    "Acct Code",
    "Acct Name",
    "Posting Date",
    "Origin",
    "Doc No",
    "Remarks",
    "Offset Acct",
    "Offset Acct Name",
    "Deb./Cred. (LC)",
    "Balance Due (LC)",
]
REQUIRED_COLUMNS = {"Series", "Due Date", "Remarks", "Doc. No."}

st.set_page_config(page_title=APP_TITLE, layout="wide")


def read_input(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    extension = Path(file_name).suffix.lower()
    buffer = io.BytesIO(file_bytes)
    if extension == ".txt":
        return pd.read_csv(buffer, sep="\t", encoding="utf-16", low_memory=False)
    if extension == ".csv":
        return pd.read_csv(buffer, low_memory=False)
    if extension == ".xlsx":
        return pd.read_excel(buffer, engine="openpyxl")
    if extension == ".xls":
        return pd.read_excel(buffer, engine="xlrd")
    raise ValueError(f"Unsupported file format: {extension or 'unknown'}")


def transform_data(data: pd.DataFrame) -> pd.DataFrame:
    df = data.copy()
    df.columns = df.columns.astype(str).str.strip()
    missing = sorted(REQUIRED_COLUMNS - set(df.columns))
    if missing:
        raise ValueError("Missing required column(s): " + ", ".join(missing))

    # Only a missing Series value identifies an account header row.
    df["Acct Code"] = df["Due Date"].where(df["Series"].isna()).ffill()
    df["Acct Name"] = df["Remarks"].where(df["Series"].isna()).ffill()
    df = df.loc[~df["Series"].isna()].copy()

    df[["Origin", "Doc No"]] = df["Doc. No."].str.split(n=1, expand=True)
    ordered = [c for c in DEFAULT_COLUMNS if c in df.columns]
    ordered += [c for c in df.columns if c not in ordered]
    return df.reindex(columns=ordered)


@st.cache_data(show_spinner="Reading and cleaning file...")
def load_and_transform(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    raw = read_input(file_bytes, file_name)
    return transform_data(raw)


def format_output(data: pd.DataFrame) -> pd.DataFrame:
    df = data.copy()
    if "Posting Date" in df:
        df["Posting Date"] = pd.to_datetime(
            df["Posting Date"], format="%d.%m.%y", errors="coerce"
        )
    return df


def clean_excel_text(data: pd.DataFrame) -> pd.DataFrame:
    """Remove control characters that are invalid inside an Excel workbook."""
    df = data.copy()
    text_columns = df.select_dtypes(include=["object", "string"]).columns
    for column in text_columns:
        df[column] = df[column].map(
            lambda value: ILLEGAL_CHARACTERS_RE.sub("", value)
            if isinstance(value, str)
            else value
        )
    return df


def build_excel_bytes(data: pd.DataFrame) -> bytes:
    """Create and validate an .xlsx file in memory."""
    safe_data = clean_excel_text(data)
    buffer = io.BytesIO()
    safe_data.to_excel(buffer, index=False, engine="openpyxl")

    # Reopen to verify the workbook is valid before offering it for download.
    buffer.seek(0)
    workbook = load_workbook(buffer, read_only=True, data_only=False)
    workbook.close()
    return buffer.getvalue()


def build_csv_bytes(data: pd.DataFrame) -> bytes:
    return data.to_csv(index=False).encode("utf-8-sig")


def move_column(column_name: str, direction: int) -> None:
    cols = list(st.session_state["selected_cols"])
    if column_name not in cols:
        return
    idx = cols.index(column_name)
    new_idx = idx + direction
    if 0 <= new_idx < len(cols):
        cols[idx], cols[new_idx] = cols[new_idx], cols[idx]
        st.session_state["selected_cols"] = cols


def main() -> None:
    st.title(APP_TITLE)
    st.caption(
        "Upload a ledger export (.txt, .csv, .xlsx, .xls), choose and order your columns, "
        "and download the cleaned file."
    )

    uploaded_file = st.file_uploader(
        "Upload Input File",
        type=["txt", "csv", "xlsx", "xls"],
    )

    if uploaded_file is None:
        return

    try:
        prepared = load_and_transform(uploaded_file.getvalue(), uploaded_file.name)
    except Exception as error:
        st.error(f"Could not process the file: {error}")
        return

    all_columns = list(prepared.columns)
    file_id = f"{uploaded_file.name}_{uploaded_file.size}"

    # Reset default column selection whenever a new file is uploaded.
    if st.session_state.get("current_file_id") != file_id:
        st.session_state["current_file_id"] = file_id
        st.session_state["selected_cols"] = [
            c for c in DEFAULT_COLUMNS if c in all_columns
        ]

    st.subheader("1. Select and Arrange Columns")
    st.multiselect(
        "Output columns (exported in the order shown below)",
        options=all_columns,
        key="selected_cols",
    )

    selected_cols = st.session_state["selected_cols"]
    if not selected_cols:
        st.warning("Select at least one column to preview and export.")
        return

    with st.expander("Reorder selected columns (Move Up / Move Down)"):
        col_to_move = st.selectbox("Pick a column to move", options=selected_cols)
        btn_col1, btn_col2, _ = st.columns([1, 1, 6])
        btn_col1.button(
            "Move Up",
            on_click=move_column,
            args=(col_to_move, -1),
            use_container_width=True,
        )
        btn_col2.button(
            "Move Down",
            on_click=move_column,
            args=(col_to_move, 1),
            use_container_width=True,
        )

    output = format_output(prepared.loc[:, selected_cols])

    st.subheader("2. Preview Cleaned Data")
    m1, m2 = st.columns(2)
    m1.metric("Total Rows", f"{len(output):,}")
    m2.metric("Selected Columns", len(output.columns))
    st.dataframe(output.head(100), use_container_width=True)

    st.subheader("3. Export Cleaned File")
    export_format = st.radio(
        "Choose export format",
        options=["Excel (.xlsx)", "CSV (.csv)"],
        horizontal=True,
    )

    base_name = Path(uploaded_file.name).stem
    if export_format == "Excel (.xlsx)":
        file_data = build_excel_bytes(output)
        file_name = f"{base_name}_cleaned.xlsx"
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    else:
        file_data = build_csv_bytes(output)
        file_name = f"{base_name}_cleaned.csv"
        mime = "text/csv"

    st.download_button(
        label=f"Download {file_name}",
        data=file_data,
        file_name=file_name,
        mime=mime,
        type="primary",
    )


if __name__ == "__main__":
    main()
