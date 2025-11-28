import streamlit as st
import pandas as pd
from io import BytesIO
import time
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.utils import get_column_letter
import datetime
import re

# Streamlit page config (do this early)
st.set_page_config(layout="wide", page_title="Energy Data Pipeline")

# --- Configuration & Constants ---
HEADER_ROW_INDEX = 2
PSUM_RAW_NAME = 'PSum (W)'      # Name used after extraction from CSV
POWER_COL_OUT = 'PSumW'         # Name used after 10-min aggregation and in Excel

# --- Helper: Excel Column Letter to Index ---
def excel_col_to_index(col_str: str) -> int:
    """Convert Excel column letter (A, B, BI) to 0-based index."""
    col_str = str(col_str).upper().strip()
    if col_str == "":
        raise ValueError("Empty column string provided.")
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
    Reads CSVs, extracts Date, Time, and PSum columns, and consolidates.
    Returns a dictionary of raw dataframes with standardized column names.
    """
    processed_data = {}
    col_indices = list(columns_config.keys())
    
    # Ensure three unique indices
    if len(set(col_indices)) != 3:
        st.error("Date, Time, and PSum must come from three unique columns.")
        return {}

    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        try:
            # Read CSV into a full DataFrame without a header so we can pick the header row specified
            # Try multiple encodings if needed
            raw_bytes = uploaded_file.read()
            try:
                df_full = pd.read_csv(BytesIO(raw_bytes), header=None, encoding='utf-8', low_memory=False)
            except Exception:
                df_full = pd.read_csv(BytesIO(raw_bytes), header=None, encoding='ISO-8859-1', low_memory=False)

            # Defensive: ensure header_index exists in the file
            if header_index >= len(df_full):
                st.warning(f"File {filename}: header row index {header_index} is beyond file length ({len(df_full)} rows). Trying row 0 as header.")
                header_index_used = 0
            else:
                header_index_used = header_index

            # Assign header names from the target row (as strings)
            header_row = df_full.iloc[header_index_used].astype(str).fillna('').tolist()
            df_full.columns = header_row

            # Take data from header_index + 1 onward
            df_full = df_full.iloc[header_index_used + 1 :].reset_index(drop=True)

            # If the provided column indices exceed available columns, attempt to map by names
            max_index_in_file = df_full.shape[1] - 1
            if any(idx > max_index_in_file or idx < 0 for idx in col_indices):
                # Try to detect columns by name heuristics
                st.warning(f"File {filename}: one or more column indices out of range. Attempting to infer Date/Time/PSum columns by name.")
                lower_cols = [c.lower() for c in df_full.columns]
                try_map = {}
                for need_name in ['date', 'time', 'psum']:
                    found = None
                    for i, cn in enumerate(lower_cols):
                        if need_name in cn:
                            found = i
                            break
                    if found is not None:
                        try_map[found] = 'Date' if need_name == 'date' else ('Time' if need_name == 'time' else PSUM_RAW_NAME)
                if len(try_map) < 3:
                    st.error(f"Couldn't infer columns for file {filename}. Please check your column letters or file format.")
                    continue
                # override col_indices and columns_config
                col_indices = list(try_map.keys())
                columns_config = {k: try_map[k] for k in try_map}

            # Extract relevant columns using iloc and assign standardized names
            df_extracted = df_full.iloc[:, col_indices].copy()
            df_extracted.columns = list(columns_config.values())

            # Standardize column presence -> ensure Date, Time, PSum column exist
            # If combined datetime present in one of them, attempt splitting
            # 1. PSum numeric conversion (handle thousands separators, commas vs dots)
            if PSUM_RAW_NAME in df_extracted.columns:
                power_series = df_extracted[PSUM_RAW_NAME].astype(str).str.strip()
                # Replace common non-numeric characters (commas as decimal separators or thousand separators)
                # If both '.' and ',' present, prefer '.' as decimal and remove ','
                def normalize_number_text(txt):
                    if pd.isna(txt):
                        return ''
                    s = str(txt).strip()
                    # remove spaces
                    s = s.replace(" ", "")
                    # if both dot and comma -> assume comma thousands -> remove commas
                    if s.count('.') > 0 and s.count(',') > 0:
                        s = s.replace(',', '')
                    else:
                        # if many commas and no dot, assume comma decimal -> replace comma with dot
                        if s.count(',') == 1 and s.count('.') == 0:
                            s = s.replace(',', '.')
                        # else remove stray commas
                        s = s.replace(',', '')
                    # remove non-digit except dot and minus
                    s = re.sub(r'[^0-9\.\-]', '', s)
                    return s

                power_series = power_series.map(normalize_number_text)
                df_extracted[PSUM_RAW_NAME] = pd.to_numeric(power_series, errors='coerce')

            # 2. Date and Time formatting cleanup (standardize date strings)
            # If Date or Time look empty but there's a combined datetime column, try to parse it
            if ('Date' not in df_extracted.columns or df_extracted['Date'].isna().all()) and ('Time' not in df_extracted.columns or df_extracted['Time'].isna().all()):
                # try to find any column containing both date & time strings
                found_dt_col = None
                for col in df_full.columns:
                    sample = df_full[col].astype(str).head(10).str.cat(sep=' ')
                    if re.search(r'\d{1,2}[:/\\-]\d{1,2}[:/\\-]\d{2,4}', sample) and re.search(r'\d{1,2}:\d{2}', sample):
                        found_dt_col = col
                        break
                if found_dt_col:
                    parsed = pd.to_datetime(df_full[found_dt_col], errors='coerce', dayfirst=True)
                    df_extracted['Date'] = parsed.dt.strftime('%d/%m/%Y')
                    df_extracted['Time'] = parsed.dt.strftime('%H:%M:%S')

            # Force attempt to parse Date and Time separately if they exist
            if 'Date' in df_extracted.columns:
                df_extracted['Date'] = pd.to_datetime(df_extracted['Date'], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y')
            if 'Time' in df_extracted.columns:
                # Sometimes time might be like "12:30:00" or "12:30" or include AM/PM
                df_extracted['Time'] = pd.to_datetime(df_extracted['Time'], errors='coerce').dt.strftime('%H:%M:%S')

            # Final trimmed DataFrame for this sheet
            # Keep only Date, Time, and PSum
            keep_cols = []
            for c in ['Date', 'Time', PSUM_RAW_NAME]:
                if c in df_extracted.columns:
                    keep_cols.append(c)
            if len(keep_cols) < 3:
                st.warning(f"File {filename}: missing one of Date/Time/{PSUM_RAW_NAME} after parsing. Attempting best-effort extraction.")
                # attempt to fill missing columns with NaNs
                for c in ['Date', 'Time', PSUM_RAW_NAME]:
                    if c not in df_extracted.columns:
                        df_extracted[c] = pd.NA
                keep_cols = ['Date', 'Time', PSUM_RAW_NAME]

            df_final = df_extracted[['Date', 'Time', PSUM_RAW_NAME]].copy()

            # Sheet name safe
            sheet_name = filename.rsplit('.', 1)[0].replace('.', '_').strip()[:31]
            processed_data[sheet_name] = df_final

        except Exception as e:
            st.error(f"Error processing file **{filename}**: {e}")
            continue

    return processed_data

# --- Stage 2 Helper: Process Single Sheet (10-min Resampling and Filtering) ---
def process_sheet(df):
    """
    Processes a single DataFrame sheet from Stage 1: combines Date/Time, 
    rounds timestamps to 10-minute intervals, filters out leading/trailing 
    zero periods, and prepares data for Excel output.
    """
    
    # Check for mandatory columns from Stage 1
    if 'Date' not in df.columns or 'Time' not in df.columns or PSUM_RAW_NAME not in df.columns:
        return pd.DataFrame()

    # Combine Date and Time into a single Timestamp column; fallback if time missing
    df = df.copy()
    combined = (df['Date'].fillna('') + ' ' + df['Time'].fillna('')).str.strip()
    df['Timestamp'] = pd.to_datetime(combined, errors="coerce", dayfirst=True)
    if df['Timestamp'].isna().all():
        # try parse Date column alone as timestamp
        df['Timestamp'] = pd.to_datetime(df['Date'], errors='coerce', dayfirst=True)

    # Rename for internal use in Stage 2
    df = df.rename(columns={PSUM_RAW_NAME: POWER_COL_OUT})

    # clean numeric values in POWER_COL_OUT (already attempted in stage 1, but be defensive)
    if POWER_COL_OUT in df.columns:
        df[POWER_COL_OUT] = pd.to_numeric(df[POWER_COL_OUT], errors='coerce')

    df = df.dropna(subset=['Timestamp', POWER_COL_OUT])
    if df.empty:
        return pd.DataFrame()
    
    # --- CORE LOGIC: FILTER LEADING AND TRAILING ZEROS (ACTIVE PERIOD) ---
    try:
        non_zero_indices = df[df[POWER_COL_OUT].abs() != 0].index
    except Exception:
        non_zero_indices = df[df[POWER_COL_OUT] != 0].index
    
    if non_zero_indices.empty:
        return pd.DataFrame()
        
    first_valid_idx = non_zero_indices.min()
    last_valid_idx = non_zero_indices.max()
    
    # Slice the DataFrame to keep data between the first and last active reading.
    df = df.loc[first_valid_idx:last_valid_idx].copy()
    # ----------------------------------------------------
    
    # Resample data to 10-minute intervals
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
    min_dt = df_out['Rounded'].min().floor('D')
    max_dt_exclusive = df_out['Rounded'].max().ceil('D') 
    
    # inclusive='left' not supported on older pandas — use end = max_dt_exclusive - 10min if necessary
    full_time_index = pd.date_range(
        start=min_dt.to_pydatetime(),
        end=(max_dt_exclusive - pd.Timedelta(minutes=10)).to_pydatetime(),
        freq='10min'
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
    
    # Reorder columns for downstream Excel builder
    grouped = grouped[['Date', 'Time', POWER_COL_OUT, 'kW', 'Rounded']]
    return grouped

# --- Stage 2 Final: Build Excel (Excel Formatting and Charting) ---
def build_output_excel(sheets_dict):
    """Creates the final formatted Excel file with data, charts, and summary."""
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        try:
            wb.remove(wb['Sheet'])
        except Exception:
            pass

    header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
    title_font = Font(bold=True, size=12)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                 top=Side(style='thin'), bottom=Side(style='thin'))

    total_sheet_data = {}
    sheet_names_list = []

    for sheet_name, df in sheets_dict.items():
        # df expected: columns ['Date','Time',POWER_COL_OUT,'kW','Rounded']
        ws = wb.create_sheet(sheet_name[:31])
        sheet_names_list.append(sheet_name[:31])
        # If Date column is object (python date), convert to datetime.date
        if df.empty:
            continue

        # Ensure Date column is datetime.date objects
        if not pd.api.types.is_datetime64_any_dtype(df['Rounded']):
            df['Rounded'] = pd.to_datetime(df['Rounded'], errors='coerce')
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.date

        dates = sorted(df["Date"].dropna().unique())
        col_start = 1
        max_row_used = 0
        daily_max_summary = []
        day_intervals = []
        
        for date in dates:
            day_data_full = df[df["Date"] == date].sort_values("Rounded")
            day_data_active = day_data_full.dropna(subset=[POWER_COL_OUT])
            
            n_rows = len(day_data_full)
            day_intervals.append(n_rows)
            
            data_start_row = 3
            merge_start = data_start_row
            merge_end = merge_start + n_rows - 1

            # date is a date object; format strings
            date_str_full = date.strftime('%Y-%m-%d')
            date_str_short = date.strftime('%d-%b')

            # Row 1: Merge date header (Long Date)
            ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+3)
            ws.cell(row=1, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center")
            
            # Row 2: Sub-headers
            ws.cell(row=2, column=col_start, value="UTC Offset (minutes)")
            ws.cell(row=2, column=col_start+1, value="Local Time Stamp")
            ws.cell(row=2, column=col_start+2, value="Active Power (W)")
            ws.cell(row=2, column=col_start+3, value="kW")

            # Merge UTC column for the date range and set the date object
            if merge_end >= merge_start:
                ws.merge_cells(start_row=merge_start, start_column=col_start, end_row=merge_end, end_column=col_start)
                date_cell = ws.cell(row=merge_start, column=col_start, value=date)
                date_cell.alignment = Alignment(horizontal="center", vertical="center")
                # Excel date format
                date_cell.number_format = 'YYYY-MM-DD' 

            # Fill data (starts at row 3)
            for idx_offset, r in enumerate(day_data_full.itertuples(), start=merge_start):
                # r.Time may be like 'HH:MM'
                ws.cell(row=idx_offset, column=col_start+1, value=r.Time)
                # Active Power (W)
                ws.cell(row=idx_offset, column=col_start+2, value=getattr(r, POWER_COL_OUT))
                # kW
                ws.cell(row=idx_offset, column=col_start+3, value=getattr(r, 'kW'))

            # Summary stats
            stats_row_start = merge_end + 1
            sum_w = float(day_data_active[POWER_COL_OUT].sum()) if not day_data_active.empty else 0.0
            mean_w = float(day_data_active[POWER_COL_OUT].mean()) if not day_data_active.empty else 0.0
            max_w = float(day_data_active[POWER_COL_OUT].max()) if not day_data_active.empty else 0.0
            sum_kw = float(day_data_active['kW'].sum()) if not day_data_active.empty else 0.0
            mean_kw = float(day_data_active['kW'].mean()) if not day_data_active.empty else 0.0
            max_kw = float(day_data_active['kW'].max()) if not day_data_active.empty else 0.0

            ws.cell(row=stats_row_start, column=col_start+1, value="Total")
            ws.cell(row=stats_row_start, column=col_start+2, value=sum_w)
            ws.cell(row=stats_row_start, column=col_start+3, value=sum_kw)
            ws.cell(row=stats_row_start+1, column=col_start+1, value="Average")
            ws.cell(row=stats_row_start+1, column=col_start+2, value=mean_w)
            ws.cell(row=stats_row_start+1, column=col_start+3, value=mean_kw)
            ws.cell(row=stats_row_start+2, column=col_start+1, value="Max")
            ws.cell(row=stats_row_start+2, column=col_start+2, value=max_w)
            ws.cell(row=stats_row_start+2, column=col_start+3, value=max_kw)

            max_row_used = max(max_row_used, stats_row_start+2)
            daily_max_summary.append((date_str_short, max_kw)) 

            # Collect data for "Total" sheet
            if date not in total_sheet_data:
                total_sheet_data[date] = {}
            total_sheet_data[date][sheet_name] = max_kw

            col_start += 4

        # Add Line Chart for Individual Sheet
        if dates and any(day_intervals):
            chart = LineChart()
            chart.title = f"Daily 10-Minute Absolute Power Profile - {sheet_name}"
            chart.y_axis.title = "kW"
            chart.x_axis.title = "Time"
            chart.height = 12.5 
            chart.width = 23 

            max_rows = max(day_intervals) if day_intervals else 0
            first_time_col = 2
            categories_ref = Reference(ws, min_col=first_time_col, min_row=3, max_row=2 + max_rows)

            col_start = 1
            for i, n_rows in enumerate(day_intervals):
                if n_rows <= 0:
                    col_start += 4
                    continue
                data_ref = Reference(ws, min_col=col_start+3, min_row=3, max_col=col_start+3, max_row=2+n_rows)
                date_title_str = dates[i].strftime('%d-%b') if i < len(dates) else f"Day{i+1}"
                s = Series(values=data_ref, title=date_title_str)
                chart.series.append(s)
                col_start += 4

            chart.set_categories(categories_ref)
            try:
                ws.add_chart(chart, f'G{max_row_used+2}')
            except Exception:
                pass

        # Add Daily Max Summary Table for Individual Sheet
        if daily_max_summary:
            start_row = max_row_used + 5 
            
            ws.cell(row=start_row, column=1, value="Daily Max Power (kW) Summary").font = title_font
            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=2)
            start_row += 1

            ws.cell(row=start_row, column=1, value="Day").fill = header_fill
            ws.cell(row=start_row, column=2, value="Max (kW)").fill = header_fill

            for d, (date_str, max_kw) in enumerate(daily_max_summary):
                row = start_row+1+d
                ws.cell(row=row, column=1, value=date_str)
                vcell = ws.cell(row=row, column=2, value=max_kw)
                vcell.number_format = numbers.FORMAT_NUMBER_00

            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 15

    # -----------------------------
    # CREATE "TOTAL" SHEET
    # -----------------------------
    if total_sheet_data:
        ws_total = wb.create_sheet("Total")
        
        # Prepare Headers
        headers = ["Date"] + sheet_names_list + ["Total Load"]
        
        # Write Headers
        for col_idx, header_text in enumerate(headers, 1):
            cell = ws_total.cell(row=1, column=col_idx, value=header_text)
            cell.font = title_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border
            ws_total.column_dimensions[get_column_letter(col_idx)].width = 20

        # Write Data
        sorted_dates = sorted(total_sheet_data.keys())
        
        for row_idx, date_obj in enumerate(sorted_dates, 2):
            # Applying date format here too, just in case
            date_cell = ws_total.cell(row=row_idx, column=1, value=date_obj)
            date_cell.number_format = 'YYYY-MM-DD'
            date_cell.border = thin_border
            date_cell.alignment = Alignment(horizontal="center")
            
            row_total_load = 0.0
            
            # Sheet Columns
            for col_idx, sheet_name in enumerate(sheet_names_list, 2):
                val = total_sheet_data[date_obj].get(sheet_name, 0)
                if pd.isna(val): val = 0
                try:
                    val = float(val)
                except Exception:
                    val = 0.0
                
                cell = ws_total.cell(row=row_idx, column=col_idx, value=val)
                cell.number_format = numbers.FORMAT_NUMBER_00
                cell.border = thin_border
                row_total_load += val
            
            # Total Load Column
            total_cell = ws_total.cell(row=row_idx, column=len(sheet_names_list) + 2, value=row_total_load)
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
            
            data_ref = Reference(ws_total, min_col=2, min_row=1, max_col=total_cols, max_row=data_max_row)
            chart_total.add_data(data_ref, titles_from_data=True)

            for s in chart_total.series:
                s.smooth = False
            
            cats_ref = Reference(ws_total, min_col=1, min_row=2, max_row=data_max_row)
            chart_total.set_categories(cats_ref)
            
            try:
                ws_total.add_chart(chart_total, "B" + str(data_max_row + 3))
            except Exception:
                pass

    stream = BytesIO()
    try:
        if 'Sheet' in wb.sheetnames:
            try:
                wb.remove(wb['Sheet'])
            except Exception:
                pass
        wb.save(stream)
        stream.seek(0)
        return stream.getvalue()
    finally:
        try:
            stream.close()
        except Exception:
            pass

# --- Main Application Pipeline ---
def app():
    st.title("⚡ Energy Data Pipeline: CSV Consolidation & 10-Min Analysis")
    st.markdown("""
    This application performs a single, two-stage process:
    1. **Extraction (Stage 1):** Upload raw CSV files. The application extracts **Date**, **Time**, and **PSum (W)** based on the column letters you configure, using **Row 3** (index 2) as the header.
    2. **Analysis (Stage 2):** The extracted data is cleaned, resampled to **10-minute intervals**, filtered to the active period (zero readings at start/end are removed), and formatted into a comprehensive Excel report with charts.
    """)
    st.write("---")

    # --- Sidebar: Column Configuration (Stage 1 Config) ---
    st.sidebar.header("⚙️ Raw CSV Column Configuration")
    st.sidebar.markdown("Specify the column letters for extraction. Data reading starts from **Row 3**.")

    date_col_str = st.sidebar.text_input("Date Column Letter (Default: A)", value='A')
    time_col_str = st.sidebar.text_input("Time Column Letter (Default: B)", value='B')
    ps_um_col_str = st.sidebar.text_input("PSum Column Letter (Default: BI)", value='BI',
                                            help="PSum (Total Active Power) column.")

    # Convert column letters to indices
    try:
        date_col_index = excel_col_to_index(date_col_str)
        time_col_index = excel_col_to_index(time_col_str)
        ps_um_col_index = excel_col_to_index(ps_um_col_str)

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
        "Upload Raw CSV files (Max 10)", 
        type=['csv'], 
        accept_multiple_files=True
    )
    if uploaded_files and len(uploaded_files) > 10:
        st.warning(f"You uploaded {len(uploaded_files)} files. Only the first 10 will be processed.")
        uploaded_files = uploaded_files[:10]
    
    st.write("---")

    # --- Execution Button ---
    if uploaded_files:
        if st.button(f"🚀 Run Full Pipeline on {len(uploaded_files)} File(s)"):
            
            # Use a spinner to show activity
            with st.spinner("Processing files... This may take a moment."):
                
                st.info("Starting Stage 1: Consolidating and cleaning raw data...")
                
                # 1. STAGE 1: Consolidation
                processed_raw_data_dict = process_uploaded_files(
                    uploaded_files, 
                    COLUMNS_TO_EXTRACT, 
                    HEADER_ROW_INDEX
                )

                if not processed_raw_data_dict:
                    st.error("Stage 1 failed: No files were successfully processed. Check file structure or column letters.")
                    return

                st.success(f"Stage 1 Complete: Consolidated data from {len(processed_raw_data_dict)} file(s).")
                st.info("Starting Stage 2: 10-Minute Resampling and Analysis...")

                # 2. STAGE 2: Analysis (10-min resampling and filtering)
                final_processed_data_dict = {}
                
                # Use progress bar for Stage 2
                progress_bar = st.progress(0)
                total_files = len(processed_raw_data_dict)
                
                for i, (sheet_name, df_raw) in enumerate(processed_raw_data_dict.items()):
                    # Call process_sheet (which now knows the fixed column names)
                    processed_df = process_sheet(df_raw)
                    
                    if not processed_df.empty:
                        final_processed_data_dict[sheet_name] = processed_df
                    
                    percent = int(((i + 1) / total_files) * 100)
                    progress_bar.progress(percent)
                    time.sleep(0.05)  # slight pause so UI updates visibly for many small files
                
                progress_bar.empty()

                if not final_processed_data_dict:
                    st.error("Stage 2 failed: No usable data found after 10-minute resampling and zero-filtering. Data might be entirely zero or contain too many errors.")
                    return
                
                st.success(f"Stage 2 Complete: Analyzed and prepared data for {len(final_processed_data_dict)} sheet(s).")

                # 3. FINAL STEP: Generate Excel Output
                st.info("Generating final Excel report with charts and summaries...")
                
                try:
                    excel_bytes = build_output_excel(final_processed_data_dict)
                    
                    # Default filename generation
                    file_names_without_ext = [f.name.rsplit('.', 1)[0] for f in uploaded_files]
                    if len(file_names_without_ext) > 1:
                        first_name = file_names_without_ext[0][:17] + "..." if len(file_names_without_ext[0]) > 20 else file_names_without_ext[0]
                        default_filename = f"{first_name}_and_{len(file_names_without_ext)-1}_Analyzed.xlsx"
                    else:
                        default_filename = f"{file_names_without_ext[0]}_Analyzed.xlsx" if file_names_without_ext else "EnergyAnalyser_Final_Report.xlsx"
                    
                    custom_filename = st.text_input("Output Excel Filename:", value=default_filename)

                    # Download Button
                    st.balloons()
                    st.header("✅ Processing Complete")
                    st.download_button(
                        label="📥 Download Final Excel Report",
                        data=excel_bytes,
                        file_name=custom_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                except Exception as e:
                    st.exception(f"An unexpected error occurred during Excel generation: {e}")

if __name__ == "__main__":
    app()
