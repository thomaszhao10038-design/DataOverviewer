import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.utils import get_column_letter
import datetime

# --- Configuration & Constants ---
HEADER_ROW_INDEX = 2  # Data headers are on the 3rd row (index 2)
PSUM_RAW_NAME = 'PSum (W)'      # Name used after extraction from CSV
POWER_COL_OUT = 'PSumW'         # Name used after 10-min aggregation and in Excel

# --- Helper: Excel Column Letter to Index ---
def excel_col_to_index(col_str):
    """Convert Excel column letter (A, B, BI) to 0-based index."""
    col_str = col_str.upper().strip()
    index = 0
    for char in col_str:
        if 'A' <= char <= 'Z':
            index = index * 26 + (ord(char) - ord('A') + 1)
        else:
            raise ValueError(f"Invalid character in column string: {col_str}")
    return index - 1

# --- Stage 1: Process CSV Files (Extraction and Clean up) ---
def process_uploaded_files(uploaded_files, columns_config, header_index):
    """
    Reads CSVs, extracts Date, Time, and PSum columns based on user config,
    and consolidates. Handles header identification and numeric conversion.
    Returns a dictionary of raw dataframes with standardized column names.
    """
    processed_data = {}
    col_indices = list(columns_config.keys())

    if len(set(col_indices)) != 3:
        # This check is largely handled by the main app's input validation but kept for safety.
        st.error("Date, Time, and PSum must come from three unique columns.")
        return {}

    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        try:
            # Read CSV using no header initially
            # Use 'None' for header to get raw column indices.
            df_full = pd.read_csv(uploaded_file, header=None, encoding='ISO-8859-1', low_memory=False)

            # Assign header names from the target row (e.g., Row 3 / index 2)
            if header_index >= len(df_full):
                 st.error(f"File **{filename}**: Header index {header_index + 1} is out of bounds (file has {len(df_full)} rows).")
                 continue

            header_row = df_full.iloc[header_index].astype(str)
            df_full.columns = header_row

            # Start data from row index + 1
            df_full = df_full[header_index + 1:].reset_index(drop=True)

            # --- Extraction ---
            # Extract data using the 0-based column indices configured by the user.
            df_extracted = df_full.iloc[:, col_indices].copy()
            df_extracted.columns = list(columns_config.values()) # Assign standardized names

            # 1. PSum numeric conversion (handle potential commas, coerce errors)
            power_series = df_extracted[PSUM_RAW_NAME].astype(str).str.strip()
            # Replace comma decimal separator with dot
            power_series = power_series.str.replace(',', '.', regex=False)
            df_extracted[PSUM_RAW_NAME] = pd.to_numeric(power_series, errors='coerce')

            # 2. Date and Time formatting cleanup
            # Robust Date parsing: using dayfirst=True to handle common European formats (dd/mm/yyyy)
            df_extracted['Date'] = pd.to_datetime(df_extracted['Date'], errors='coerce', dayfirst=True).dt.date
            df_extracted['Time'] = pd.to_datetime(df_extracted['Time'], errors='coerce', format='%H:%M:%S').dt.time
            
            df_final = df_extracted[['Date', 'Time', PSUM_RAW_NAME]].copy()
            df_final.dropna(subset=['Date', 'Time', PSUM_RAW_NAME], inplace=True)

            # Create sheet name (safe)
            sheet_name = filename.replace('.csv','').replace('.','_').strip()[:31]
            processed_data[sheet_name] = df_final

        except Exception as e:
            st.error(f"Error processing file **{filename}**: {e}")
            continue

    return processed_data

# --- Stage 2 Helper: Process Single Sheet (10-min Resampling and Filtering) ---
def process_sheet(df_raw):
    """
    Processes a single DataFrame sheet from Stage 1: combines Date/Time,
    rounds timestamps to 10-minute intervals, filters out leading/trailing
    zero periods (non-active time), and pads the remaining data to a continuous
    10-min index for accurate charting.
    """
    df = df_raw.copy()
    
    # Combine Date (datetime.date) and Time (datetime.time) into a single Timestamp (datetime.datetime)
    df['Timestamp'] = pd.to_datetime(df['Date'].astype(str) + ' ' + df['Time'].astype(str), errors="coerce")
    
    # Rename for internal use in Stage 2
    df = df.rename(columns={PSUM_RAW_NAME: POWER_COL_OUT})

    # Drop rows where we failed to create a Timestamp or Power value
    df = df.dropna(subset=['Timestamp', POWER_COL_OUT])
    if df.empty:
        return pd.DataFrame()
    
    # --- CORE LOGIC: FILTER LEADING AND TRAILING ZEROS (ACTIVE PERIOD) ---
    # Find indices where power reading is non-zero
    non_zero_indices = df[df[POWER_COL_OUT].abs() != 0].index
    
    if non_zero_indices.empty:
        return pd.DataFrame()
        
    first_valid_idx = non_zero_indices.min()
    last_valid_idx = non_zero_indices.max()
    
    # Slice the DataFrame to keep data between the first and last active reading.
    df = df.loc[first_valid_idx:last_valid_idx].copy()
    # ----------------------------------------------------
    
    # Resample data to 10-minute intervals (Summing all readings within the interval)
    df_indexed = df.set_index('Timestamp')
    df_indexed.index = df_indexed.index.floor('10min')
    resampled_data = df_indexed[POWER_COL_OUT].groupby(level=0).sum()
    
    df_out = resampled_data.reset_index()
    df_out.columns = ['Rounded', POWER_COL_OUT]
    
    if df_out.empty or df_out['Rounded'].isna().all():
        return pd.DataFrame()
    
    # Get the original dates present in the processed data
    original_dates = set(df_out['Rounded'].dt.date)
    
    # Create a full 10-minute index from the start of the first day to the end of the last day
    min_dt = df_out['Rounded'].min().normalize() # Start of first day
    max_dt_exclusive = (df_out['Rounded'].max() + pd.Timedelta(minutes=10)).normalize() # Start of the day after the last reading
    
    full_time_index = pd.date_range(
        start=min_dt.to_pydatetime(),
        end=max_dt_exclusive.to_pydatetime(),
        freq='10min',
        inclusive='left'
    )
    
    # Reindex with the full index, filling missing slots with NaN (blank)
    df_indexed_for_reindex = df_out.set_index('Rounded')
    df_padded_series = df_indexed_for_reindex[POWER_COL_OUT].reindex(full_time_index)
    
    grouped = df_padded_series.reset_index().rename(columns={'index': 'Rounded'})
    grouped.columns = ['Rounded', POWER_COL_OUT]
    
    grouped[POWER_COL_OUT] = grouped[POWER_COL_OUT].astype(float)
    grouped["Date"] = grouped["Rounded"].dt.date
    grouped["Time"] = grouped["Rounded"].dt.strftime("%H:%M")
    
    # Filter back to only the dates originally present in the file
    grouped = grouped[grouped["Date"].isin(original_dates)]
    
    # Add kW column (absolute value)
    grouped['kW'] = grouped[POWER_COL_OUT].abs() / 1000
    
    return grouped

# --- Stage 2 Final: Build Excel (Excel Formatting and Charting) ---
def build_output_excel(sheets_dict):
    """Creates the final formatted Excel file with data, charts, and summary."""
    wb = Workbook()
    # Remove the default sheet created by openpyxl
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # Style definitions
    header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid') # Light Blue
    title_font = Font(bold=True, size=12, color='000080') # Dark Blue Bold
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    total_sheet_data = {}
    sheet_names_list = []

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(sheet_name)
        sheet_names_list.append(sheet_name)
        
        # Get unique dates and sort them
        dates = sorted(df["Date"].unique())
        
        col_start = 1
        max_row_used = 0
        daily_max_summary = []
        day_intervals = []

        # --- Data Layout (Daily Columns) ---
        for date in dates:
            day_data_full = df[df["Date"] == date].sort_values("Time")
            # Only use non-NaN data for summary statistics
            day_data_active = day_data_full.dropna(subset=[POWER_COL_OUT])

            n_rows = len(day_data_full)
            day_intervals.append(n_rows) # number of 10-min intervals for this day

            data_start_row = 3
            merge_start = data_start_row
            merge_end = merge_start + n_rows - 1

            date_str_full = date.strftime('%Y-%m-%d')
            date_str_short = date.strftime('%d-%b')

            # Row 1: Merge date header (Long Date)
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+3)
            date_title_cell = ws.cell(row=1, column=col_start, value=date_str_full)
            date_title_cell.alignment = Alignment(horizontal="center")
            date_title_cell.font = title_font
            date_title_cell.fill = header_fill

            # Row 2: Sub-headers
            ws.cell(row=2, column=col_start, value="Date (UTC Offset)")
            ws.cell(row=2, column=col_start+1, value="Time Stamp (10-min)")
            ws.cell(row=2, column=col_start+2, value="Active Power (W)")
            ws.cell(row=2, column=col_start+3, value="kW (Absolute)")
            
            for c in range(col_start, col_start + 4):
                 ws.cell(row=2, column=c).fill = header_fill
                 ws.cell(row=2, column=c).font = Font(bold=True)
                 ws.cell(row=2, column=c).alignment = Alignment(horizontal="center")

            # UTC Offset Column (Merged Date Cell)
            ws.merge_cells(start_row=merge_start, start_column=col_start, end_row=merge_end, end_column=col_start)
            date_cell = ws.cell(row=merge_start, column=col_start, value=date)
            date_cell.alignment = Alignment(horizontal="center", vertical="center")
            # Set the number format to ensure Excel interprets the numeric value as a date
            date_cell.number_format = 'YYYY-MM-DD'

            # Fill data (starts at row 3)
            for idx, r in enumerate(day_data_full.itertuples(), start=merge_start):
                time_cell = ws.cell(row=idx, column=col_start+1, value=r.Time)
                time_cell.number_format = numbers.FORMAT_DATE_TIME3 # H:MM
                
                ws.cell(row=idx, column=col_start+2, value=getattr(r, POWER_COL_OUT)).number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
                ws.cell(row=idx, column=col_start+3, value=r.kW).number_format = numbers.FORMAT_NUMBER_00

            # --- Summary Stats ---
            stats_row_start = merge_end + 1
            
            # Summary Calculations (using active data only)
            sum_w = day_data_active[POWER_COL_OUT].sum()
            mean_w = day_data_active[POWER_COL_OUT].mean()
            max_w = day_data_active[POWER_COL_OUT].max()
            sum_kw = day_data_active['kW'].sum()
            mean_kw = day_data_active['kW'].mean()
            max_kw = day_data_active['kW'].max()

            # Write Summary
            summary_labels = ["Total Sum", "Average", "Max Peak"]
            summary_values_W = [sum_w, mean_w, max_w]
            summary_values_kW = [sum_kw, mean_kw, max_kw]
            
            for i, label in enumerate(summary_labels):
                row = stats_row_start + i
                ws.cell(row=row, column=col_start+1, value=label).font = Font(bold=True)
                ws.cell(row=row, column=col_start+2, value=summary_values_W[i]).number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
                ws.cell(row=row, column=col_start+3, value=summary_values_kW[i]).number_format = numbers.FORMAT_NUMBER_00
                
            max_row_used = max(max_row_used, stats_row_start + len(summary_labels) - 1)
            daily_max_summary.append((date_str_short, max_kw))

            # Collect data for "Total" sheet
            if date not in total_sheet_data:
                total_sheet_data[date] = {}
            total_sheet_data[date][sheet_name] = max_kw

            # Move to the next column group (4 columns wide)
            col_start += 4
            
            # Set column widths for visibility
            for c_idx in range(1, col_start):
                 ws.column_dimensions[get_column_letter(c_idx)].width = 15

        # --- Add Line Chart for Individual Sheet ---
        if dates and day_intervals:
            chart = LineChart()
            chart.title = f"Daily 10-Minute Absolute Power Profile - {sheet_name}"
            chart.y_axis.title = "kW"
            chart.x_axis.title = "Time of Day"
            chart.height = 12.5
            chart.width = 23
            
            # Determine the maximum number of intervals across all days for consistent chart height
            max_rows = max(day_intervals) 
            first_time_col = 2
            
            # Categories are the Local Time Stamps from the first day's column group (col_start + 1)
            categories_ref = Reference(ws, min_col=first_time_col, min_row=3, max_row=2 + max_rows)

            col_start_chart_data = 1
            for i, n_rows in enumerate(day_intervals):
                # Data Ref is the 'kW' column (col_start + 3)
                data_ref = Reference(ws, min_col=col_start_chart_data+3, min_row=3, max_col=col_start_chart_data+3, max_row=2+n_rows)
                date_title_str = dates[i].strftime('%d-%b')
                s = Series(values=data_ref, title=date_title_str)
                chart.series.append(s)
                col_start_chart_data += 4

            chart.set_categories(categories_ref)
            ws.add_chart(chart, f'F{max_row_used+2}') # Insert chart below data columns, slightly to the right

        # --- Add Daily Max Summary Table for Individual Sheet ---
        if daily_max_summary:
            start_row = max_row_used + 5
            
            ws.cell(row=start_row, column=1, value="Daily Max Power (kW) Summary").font = title_font
            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=2)
            start_row += 1

            ws.cell(row=start_row, column=1, value="Day").fill = header_fill
            ws.cell(row=start_row, column=2, value="Max (kW)").fill = header_fill
            
            for d, (date_str, max_kw) in enumerate(daily_max_summary):
                row = start_row+1+d
                ws.cell(row=row, column=1, value=date_str).border = thin_border
                ws.cell(row=row, column=2, value=max_kw).number_format = numbers.FORMAT_NUMBER_00
                ws.cell(row=row, column=2).border = thin_border
                
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 15

    # -----------------------------
    # CREATE "TOTAL" SHEET
    # -----------------------------
    if total_sheet_data:
        ws_total = wb.create_sheet("Total")
        
        # Prepare Headers
        headers = ["Date"] + sheet_names_list + ["Total Load (kW)"]
        
        # Write Headers
        for col_idx, header_text in enumerate(headers, 1):
            cell = ws_total.cell(row=1, column=col_idx, value=header_text)
            cell.font = title_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border
            # Set width for all columns on the Total sheet
            ws_total.column_dimensions[get_column_letter(col_idx)].width = 20

        # Write Data
        sorted_dates = sorted(total_sheet_data.keys())
        
        for row_idx, date_obj in enumerate(sorted_dates, 2):
            # Date Column
            date_cell = ws_total.cell(row=row_idx, column=1, value=date_obj)
            date_cell.number_format = 'YYYY-MM-DD'
            date_cell.border = thin_border
            date_cell.alignment = Alignment(horizontal="center")
            
            row_total_load = 0
            
            # Sheet Max Load Columns
            for col_idx, sheet_name in enumerate(sheet_names_list, 2):
                val = total_sheet_data[date_obj].get(sheet_name, 0)
                if pd.isna(val): val = 0
                
                cell = ws_total.cell(row=row_idx, column=col_idx, value=val)
                cell.number_format = numbers.FORMAT_NUMBER_00
                cell.border = thin_border
                row_total_load += val
            
            # Total Load Column (Last column)
            total_col_index = len(sheet_names_list) + 2
            total_cell = ws_total.cell(row=row_idx, column=total_col_index, value=row_total_load)
            total_cell.number_format = numbers.FORMAT_NUMBER_00
            total_cell.border = thin_border
            total_cell.font = Font(bold=True)

        # Add Chart to Total Sheet
        if sorted_dates:
            chart_total = LineChart()
            chart_total.title = "Daily Max Power Overview"
            chart_total.y_axis.title = "Max Power (kW)"
            chart_total.x_axis.title = "Date"
            
            chart_total.height = 15
            chart_total.width = 30
            
            data_max_row = len(sorted_dates) + 1
            total_cols = len(sheet_names_list) + 2
            
            # Chart Data Reference: Cover all value columns (Sheet 1...Sheet N + Total Load)
            data_ref = Reference(ws_total, min_col=2, min_row=1, max_col=total_cols, max_row=data_max_row)
            chart_total.add_data(data_ref, titles_from_data=True)

            # Set smooth=False for straight lines between daily peaks
            for s in chart_total.series:
                s.smooth = False
            
            # Category Axis: Date Column (Col 1, data starts at row 2)
            cats_ref = Reference(ws_total, min_col=1, min_row=2, max_row=data_max_row)
            chart_total.set_categories(cats_ref)
            
            ws_total.add_chart(chart_total, "B" + str(data_max_row + 3))

    stream = BytesIO()
    # Ensure the default sheet is removed if it wasn't already (for safety)
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
        
    wb.save(stream)
    stream.seek(0)
    return stream


# --- Main Application Pipeline ---
def app():
    st.set_page_config(layout="wide", page_title="Energy Data Pipeline")
    
    st.title("⚡ Energy Data Pipeline: CSV Consolidation & 10-Min Analysis")
    st.markdown("""
    This application automates the analysis of raw CSV power data:
    1. **Extraction:** Extracts **Date**, **Time**, and **PSum (W)** columns using user-configured column letters.
    2. **Resampling:** Converts data to a **10-minute summation interval** (for total energy consumed in that period).
    3. **Filtering:** Automatically removes initial and final periods of zero readings to focus the report only on the **active usage time**.
    4. **Reporting:** Generates a professional Excel report with daily profiles, statistics, and a summary chart of maximum daily demand.
    """)
    st.write("---")

    # --- Sidebar: Column Configuration (Stage 1 Config) ---
    st.sidebar.header("⚙️ Raw CSV Column Configuration")
    st.sidebar.markdown("Specify the column letters for **Date**, **Time**, and **Total Active Power (PSum)**. The app assumes the header row is **Row 3** of your CSV.")

    # Use persistent session state for inputs
    if 'date_col_str' not in st.session_state:
        st.session_state.date_col_str = 'A'
    if 'time_col_str' not in st.session_state:
        st.session_state.time_col_str = 'B'
    if 'psum_col_str' not in st.session_state:
        st.session_state.psum_col_str = 'BI'

    date_col_str = st.sidebar.text_input("Date Column Letter:", value=st.session_state.date_col_str, key='date_col_input')
    time_col_str = st.sidebar.text_input("Time Column Letter:", value=st.session_state.time_col_str, key='time_col_input')
    ps_um_col_str = st.sidebar.text_input("PSum (W) Column Letter:", value=st.session_state.psum_col_str, key='psum_col_input',
                                          help="PSum (Total Active Power) column. Must be in Watts (W).")

    # Convert column letters to indices
    try:
        date_col_index = excel_col_to_index(date_col_str)
        time_col_index = excel_col_to_index(time_col_str)
        ps_um_col_index = excel_col_to_index(ps_um_col_str)

        # Check for unique columns
        if len({date_col_index, time_col_index, ps_um_col_index}) != 3:
            st.error("Configuration Error: Date, Time, and PSum must use three different column letters.")
            return

        COLUMNS_TO_EXTRACT = {
            date_col_index: 'Date',
            time_col_index: 'Time',
            ps_um_col_index: PSUM_RAW_NAME
        }

    except ValueError as e:
        st.error(f"Configuration Error: {e}")
        return

    # --- Main Area: File Upload ---
    uploaded_files = st.file_uploader(
        "Upload Raw CSV files (Max 10 per batch)", 
        type=['csv'], 
        accept_multiple_files=True
    )
    if uploaded_files and len(uploaded_files) > 10:
        st.warning(f"You uploaded {len(uploaded_files)} files. Only the first 10 will be processed.")
        uploaded_files = uploaded_files[:10]
    
    st.write("---")

    # --- Execution Button ---
    if uploaded_files:
        if st.button(f"🚀 Run Full Pipeline on {len(uploaded_files)} File(s)", type="primary"):
            
            # Use a spinner to show activity
            with st.spinner("Processing files... This may take a moment."):
                
                # 1. STAGE 1: Consolidation and Extraction
                st.info("Starting Stage 1: Consolidating and cleaning raw data...")
                processed_raw_data_dict = process_uploaded_files(
                    uploaded_files, 
                    COLUMNS_TO_EXTRACT, 
                    HEADER_ROW_INDEX
                )

                if not processed_raw_data_dict:
                    st.error("Stage 1 failed: No files were successfully processed. Check file structure or column letters.")
                    return

                st.success(f"Stage 1 Complete: Consolidated data from {len(processed_raw_data_dict)} file(s).")
                st.info("Starting Stage 2: 10-Minute Resampling, Zero-Filtering, and Analysis...")

                # 2. STAGE 2: Analysis (10-min resampling and filtering)
                final_processed_data_dict = {}
                
                # Use progress bar for Stage 2
                progress_bar = st.progress(0)
                
                for i, (sheet_name, df_raw) in enumerate(processed_raw_data_dict.items()):
                    # Call process_sheet (which now knows the fixed column names)
                    processed_df = process_sheet(df_raw)
                    
                    if not processed_df.empty:
                        final_processed_data_dict[sheet_name] = processed_df
                    else:
                        st.warning(f"File **{sheet_name}**: No usable data found after filtering (data might be entirely zero or contain too many errors). This file will be skipped in the final report.")
                    
                    progress_bar.progress((i + 1) / len(processed_raw_data_dict))
                
                progress_bar.empty()

                if not final_processed_data_dict:
                    st.error("Stage 2 failed: No usable data found for analysis across all files. Please check input data quality.")
                    return
                
                st.success(f"Stage 2 Complete: Analyzed and prepared data for {len(final_processed_data_dict)} sheet(s).")

                # 3. FINAL STEP: Generate Excel Output
                st.info("Generating final Excel report with charts and summaries...")
                
                try:
                    excel_data = build_output_excel(final_processed_data_dict)
                    
                    # Default filename generation logic
                    file_names_without_ext = [f.name.rsplit('.', 1)[0] for f in uploaded_files]
                    if len(file_names_without_ext) > 1:
                        first_name = file_names_without_ext[0][:17].strip() + ("..." if len(file_names_without_ext[0]) > 17 else "")
                        default_filename = f"{first_name}_and_{len(file_names_without_ext)-1}_More_Analyzed.xlsx"
                    elif file_names_without_ext:
                        default_filename = f"{file_names_without_ext[0]}_Analyzed.xlsx"
                    else:
                        default_filename = "EnergyAnalyser_Final_Report.xlsx"
                    
                    # Allow user to customize filename
                    custom_filename = st.text_input("Output Excel Filename:", value=default_filename)

                    # Download Button
                    st.balloons()
                    st.header("✅ Processing Complete")
                    st.download_button(
                        label="📥 Download Final Excel Report",
                        data=excel_data,
                        file_name=custom_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                except Exception as e:
                    st.exception(f"An unexpected error occurred during Excel report generation: {e}")
    else:
        st.info("Upload CSV files to begin the energy data pipeline.")

if __name__ == "__main__":
    app()
