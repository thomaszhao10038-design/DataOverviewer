import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.utils import get_column_letter

# -----------------------------------------------------
# PAGE CONFIG
# -----------------------------------------------------
st.set_page_config(
    page_title="EnergyAnalyser",
    layout="wide",
    initial_sidebar_state="expanded"
)


# -----------------------------------------------------
# HELPER: PROCESS SINGLE CSV
# -----------------------------------------------------
def process_sheet(df):
    """
    Extract Date, Time, PSum(W) from dataset.
    Preserves behaviour of Code 1 but cleaner.
    """
    df = df.copy()

    # Identify timestamp column
    timestamp_col = None
    for col in df.columns:
        if "time" in col.lower():
            timestamp_col = col
            break

    if timestamp_col is None:
        st.error("No timestamp column found.")
        return None

    df["Timestamp"] = pd.to_datetime(df[timestamp_col], errors="coerce")
    df = df.dropna(subset=["Timestamp"])

    df["Date"] = df["Timestamp"].dt.date
    df["Time"] = df["Timestamp"].dt.strftime("%H:%M:%S")

    # Identify PSum column
    psum_col = None
    for col in df.columns:
        if "psum" in col.lower():
            psum_col = col
            break

    if psum_col is None:
        st.error("No PSum(W) column found.")
        return None

    df["PSum (W)"] = df[psum_col]

    return df[["Date", "Time", "PSum (W)"]]


# -----------------------------------------------------
# MERGED EXCEL BUILDER (FULL CODE 2)
# -----------------------------------------------------
def build_output_excel(all_processed_data):
    """
    all_processed_data = {
        "filename1": {date1: df, date2: df},
        "filename2": {date3: df, date4: df}
    }

    Produces 1 Excel file containing all sheets.
    """
    output = BytesIO()
    wb = Workbook()

    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    # Formatting presets (from Code 2)
    header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
    bold_font = Font(bold=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # -----------------------------------------------------
    # CREATE SHEETS FOR EACH INPUT FILE + EACH DATE
    # -----------------------------------------------------
    for filename, daily_dict in all_processed_data.items():
        for date_label, df in daily_dict.items():

            sheet_name = f"{filename[:20]}_{date_label}"
            ws = wb.create_sheet(sheet_name)

            # ---- Write headers ----
            for col_idx, col_name in enumerate(df.columns, start=1):
                cell = ws.cell(row=1, column=col_idx, value=col_name)
                cell.font = bold_font
                cell.fill = header_fill
                cell.border = thin_border
                ws.column_dimensions[get_column_letter(col_idx)].width = 20

            # ---- Write data ----
            for r_idx, row in df.iterrows():
                for c_idx, value in enumerate(row, start=1):
                    ws.cell(row=r_idx + 2, column=c_idx, value=value)

            # ---- Add Chart (same as Code 2) ----
            if "PSum (W)" in df.columns:
                chart = LineChart()
                chart.title = f"Power Profile – {sheet_name}"
                chart.y_axis.title = "Power (W)"
                chart.x_axis.title = "Time"

                last_row = len(df) + 1

                cat_col = df.columns.get_loc("Time") + 1
                categories = Reference(ws, min_col=cat_col, min_row=2, max_row=last_row)

                psum_col = df.columns.get_loc("PSum (W)") + 1
                values = Reference(ws, min_col=psum_col, min_row=1, max_row=last_row)

                series = Series(values, title="PSum (W)")
                chart.series.append(series)
                chart.set_categories(categories)

                ws.add_chart(chart, "G2")

    wb.save(output)
    return output.getvalue()


# -----------------------------------------------------
# STREAMLIT UI (MULTIPLE CSV INPUT)
# -----------------------------------------------------
st.title("⚡ EnergyAnalyser — Multi-Day Electricity Dataset Processor")

uploaded_files = st.file_uploader("Upload one or more CSV files", type=["csv"], accept_multiple_files=True)

if uploaded_files:
    all_processed_data = {}

    for file in uploaded_files:
        st.subheader(f"📂 {file.name}")
        df_input = pd.read_csv(file)
        st.dataframe(df_input.head())

        processed = process_sheet(df_input)

        if processed is not None:
            daily_dict = {
                str(date): group.reset_index(drop=True)
                for date, group in processed.groupby("Date")
            }

            all_processed_data[file.name] = daily_dict

            # Preview grouped days
            for d, sub in daily_dict.items():
                with st.expander(f"📅 {file.name} — {d}"):
                    st.dataframe(sub.head())

    # -------------------------------------------------
    # Build combined Excel file
    # -------------------------------------------------
    excel_output = build_output_excel(all_processed_data)

    st.download_button(
        label="📥 Download Combined Excel (Charts + Formatting)",
        data=excel_output,
        file_name="EnergyAnalysis_Combined.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
