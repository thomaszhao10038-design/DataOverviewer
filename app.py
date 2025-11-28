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
# HELPER: PROCESS SINGLE SHEET
# -----------------------------------------------------
def process_sheet(df):
    """
    Extract 'Date', 'Time', 'PSum (W)' from uploaded dataset.
    """
    df = df.copy()

    # Find timestamp column
    timestamp_col = None
    for col in df.columns:
        if "time" in col.lower():
            timestamp_col = col
            break

    if timestamp_col is None:
        st.error("No timestamp column found.")
        return None

    # Convert to datetime
    df["Timestamp"] = pd.to_datetime(df[timestamp_col], errors="coerce")
    df = df.dropna(subset=["Timestamp"])

    # Extract date & time
    df["Date"] = df["Timestamp"].dt.date
    df["Time"] = df["Timestamp"].dt.strftime("%H:%M:%S")

    # Find PSum column
    psum_col = None
    for col in df.columns:
        if "psum" in col.lower():
            psum_col = col
            break

    if psum_col is None:
        st.error("No PSum column found.")
        return None

    df["PSum (W)"] = df[psum_col]

    # Final tidy dataset
    return df[["Date", "Time", "PSum (W)"]]


# -----------------------------------------------------
# EXCEL BUILDER WITH CHARTS (MERGED FROM CODE 2)
# -----------------------------------------------------
def build_output_excel(processed_dict):
    """
    Converts {sheet_name: dataframe} → full Excel workbook
    with formatting + line charts.
    """
    output = BytesIO()
    wb = Workbook()

    # remove default sheet
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    # Styling
    header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
    bold_font = Font(bold=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # -----------------------------------------------------
    # PROCESS EACH SHEET
    # -----------------------------------------------------
    for sheet_name, df in processed_dict.items():
        ws = wb.create_sheet(sheet_name)

        # Write headers
        for col_idx, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = bold_font
            cell.fill = header_fill
            cell.border = thin_border
            ws.column_dimensions[get_column_letter(col_idx)].width = 20

        # Write data
        for r_idx, row in df.iterrows():
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx + 2, column=c_idx, value=value)

        # -------------------------------------------------
        # Add Excel Line Chart
        # -------------------------------------------------
        if "PSum (W)" in df.columns:
            chart = LineChart()
            chart.title = f"Power Profile – {sheet_name}"
            chart.y_axis.title = "Power (W)"
            chart.x_axis.title = "Time"

            # category axis = time
            cat_col = df.columns.get_loc("Time") + 1
            last_row = len(df) + 1
            categories = Reference(ws, min_col=cat_col, min_row=2, max_row=last_row)

            # data series
            psum_col = df.columns.get_loc("PSum (W)") + 1
            values = Reference(ws, min_col=psum_col, min_row=1, max_row=last_row)

            series = Series(values, title="PSum (W)")
            chart.series.append(series)
            chart.set_categories(categories)

            ws.add_chart(chart, "G2")

    wb.save(output)
    return output.getvalue()


# -----------------------------------------------------
# STREAMLIT UI
# -----------------------------------------------------
st.title("⚡ EnergyAnalyser — Multi-Day Electricity Dataset Processor")

uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])


if uploaded_file:
    df_input = pd.read_csv(uploaded_file)

    st.subheader("Preview of Uploaded Data")
    st.dataframe(df_input.head())

    # Process dataset into dictionary of date → DataFrame
    df_processed = process_sheet(df_input)

    if df_processed is not None:
        # Split by day
        processed_dict = {
            str(date): group.reset_index(drop=True)
            for date, group in df_processed.groupby("Date")
        }

        st.subheader("Processed Dataset by Day")
        st.write(f"Detected **{len(processed_dict)}** separate day(s).")

        # Show preview of each day
        for k, v in processed_dict.items():
            with st.expander(f"📅 {k}"):
                st.dataframe(v.head())

        # Generate Excel
        excel_data = build_output_excel(processed_dict)

        st.download_button(
            label="📥 Download Full Excel with Charts",
            data=excel_data,
            file_name="EnergyAnalysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
