import streamlit as st
import pandas as pd
from io import BytesIO
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.utils import get_column_letter

# --- Configuration ---
# Constants for Code 1
PSUM_OUTPUT_NAME = 'PSum (W)' 
DATE_FORMAT_MAP = {
    "DD/MM/YYYY": "%d/%m/%Y %H:%M:%S",
    "YYYY-MM-DD": "%Y-%m-%d %H:%M:%S"
}

# Constants for Code 2
POWER_COL_OUT = 'PSumW'

# --- Helper Function for Excel Column Conversion (from Code 1) ---
def excel_col_to_index(col_str):
    """
    Converts an Excel column string (e.g., 'A', 'AA', 'BI') to a 0-based column index.
    Raises a ValueError if the string is invalid.
    """
    col_str = col_str.upper().strip()
    index = 0
    # A=1, B=2, ..., Z=26
    for char in col_str:
        if 'A' <= char <= 'Z':
            index = index * 26 + (ord(char) - ord('A') + 1)
        else:
            raise ValueError(f"Invalid character in column string: {col_str}")
    
    return index - 1

# -------------------------------------
# --- Functions from Code 2.py (Analysis) ---
# -------------------------------------

def process_sheet(df, date_col, time_col, psum_col):
    """
    Processes a single DataFrame sheet (derived from Code 1 output): cleans data, 
    rounds timestamps to 10-minute intervals, filters out leading/trailing zero periods, 
    and prepares data for Excel output.
    """
    # 1. Combine Date and Time columns into a single timestamp string/series
    # Date format is consistently 'DD/MM/YYYY' from Code 1 output
    combined_dt_series = df[date_col].astype(str) + ' ' + df[time_col].astype(str)
    
    # Convert the combined string to datetime objects, using the *known* format from Code 1
    df['Timestamp'] = pd.to_datetime(combined_dt_series, errors="coerce", format='%d/%m/%Y %H:%M:%S')
    timestamp_col = 'Timestamp'
    
    # 2. Clean and convert power column (handle potential commas, although Code 1 should output clean data)
    power_series = df[psum_col].astype(str).str.strip()
    # Replace comma decimal separator if it somehow persisted or was introduced
    power_series = power_series.str.replace(',', '.', regex=False) 
    df[psum_col] = pd.to_numeric(power_series, errors='coerce')
    
    # 3. Drop rows where we couldn't parse the timestamp or power value
    df = df.dropna(subset=[timestamp_col, psum_col])
    if df.empty:
        return pd.DataFrame()
    
    # --- CORE LOGIC: FILTER LEADING AND TRAILING ZEROS ---
    
    # Identify indices where the absolute power reading is non-zero
    # np.isclose is safer than != 0 for float comparisons
    non_zero_indices = df[~np.isclose(df[psum_col].abs(), 0.0)].index
    
    if non_zero_indices.empty:
        return pd.DataFrame() 
        
    # Get the index of the first and last non-zero reading
    first_valid_idx = non_zero_indices.min()
    last_valid_idx = non_zero_indices.max()
    
    # Slice the DataFrame to keep data between the first and last active reading.
    df = df.loc[first_valid_idx:last_valid_idx].copy()

    # ----------------------------------------------------
    
    # Resample data to 10-minute intervals
    df_indexed = df.set_index(timestamp_col)
    df_indexed.index = df_indexed.index.floor('10min')
    # Sum the power values within each 10-minute slot
    resampled_data = df_indexed[psum_col].groupby(level=0).sum()
    
    df_out = resampled_data.reset_index()
    df_out.columns = ['Rounded', POWER_COL_OUT] 
    
    if df_out.empty or df_out['Rounded'].isna().all():
        return pd.DataFrame()
    
    # Get the original dates present in the processed data
    original_dates = set(df_out['Rounded'].dt.date)
    
    # Create a full 10-minute index from the start of the first day to the end of the last day
    min_dt = df_out['Rounded'].min().floor('D')
    max_dt_exclusive = df_out['Rounded'].max().ceil('D') 
    if min_dt >= max_dt_exclusive:
        return pd.DataFrame()
    
    full_time_index = pd.date_range(
        start=min_dt.to_pydatetime(),
        end=max_dt_exclusive.to_pydatetime(),
        freq='10min',
        inclusive='left'
    )
    
    # Reindex with the full index, filling missing slots with NaN (blank) instead of 0.
    df_indexed_for_reindex = df_out.set_index('Rounded')
    df_padded_series = df_indexed_for_reindex[POWER_COL_OUT].reindex(full_time_index) 
    
    grouped = df_padded_series.reset_index().rename(columns={'index': 'Rounded'})
    grouped.columns = ['Rounded', POWER_COL_OUT]
    
    # Ensure the column is float type to correctly hold NaN values
    grouped[POWER_COL_OUT] = grouped[POWER_COL_OUT].astype(float) 

    grouped["Date"] = grouped["Rounded"].dt.date
    grouped["Time"] = grouped["Rounded"].dt.strftime("%H:%M") 
    
    # Filter back to only the dates originally present in the file
    grouped = grouped[grouped["Date"].isin(original_dates)]
    
    # Add kW column (absolute value)
    grouped['kW'] = grouped[POWER_COL_OUT].abs() / 1000
    
    return grouped

def prepare_chart_data(analysis_results):
    """
    Aggregates the daily maximum kW for all sheets into a single DataFrame
    suitable for Streamlit charting and returns the total_sheet_data dictionary.
    """
    total_sheet_data = {}
    
    for sheet_name, df in analysis_results.items():
        # Ensure 'Date' column is converted to datetime.date objects for consistent grouping
        df['Date'] = pd.to_datetime(df['Date']).dt.date
        
        # Get daily max kW for the sheet
        daily_maxes = df.groupby('Date')['kW'].max().reset_index()
        
        for _, row in daily_maxes.iterrows():
            date = row['Date']
            max_kw = row['kW']
            
            if date not in total_sheet_data:
                total_sheet_data[date] = {}
            total_sheet_data[date][sheet_name] = max_kw

    # Convert the dictionary structure to a DataFrame
    dates = sorted(total_sheet_data.keys())
    # Create an index with just the dates (not datetime objects)
    chart_df = pd.DataFrame(index=dates)
    sheet_names_list = sorted(list(set(sheet for date_data in total_sheet_data.values() for sheet in date_data.keys())))
    
    for date in dates:
        for sheet_name in sheet_names_list:
            # Fill the DataFrame with the max kW value, or 0 if missing for that day/sheet
            chart_df.loc[date, sheet_name] = total_sheet_data[date].get(sheet_name, 0)

    # Calculate Total Load column
    chart_df['Total Load (kW)'] = chart_df[sheet_names_list].sum(axis=1)
    chart_df.index.name = 'Date'
    
    return chart_df, total_sheet_data

def build_output_excel(sheets_dict, total_sheet_data):
    """
    Creates the final formatted Excel file with data, charts, and summary, 
    using the pre-calculated total_sheet_data for the 'Total' sheet.
    """
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
    title_font = Font(bold=True, size=12)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                              top=Side(style='thin'), bottom=Side(style='thin'))

    sheet_names_list = list(sheets_dict.keys())

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(sheet_name)
        # Ensure 'Date' column is converted to datetime.date objects for sorting
        df['Date'] = pd.to_datetime(df['Date']).dt.date
        dates = sorted(df["Date"].unique())
        col_start = 1
        max_row_used = 0
        daily_max_summary = []

        day_intervals = []
        
        for date in dates:
            day_data_full = df[df["Date"] == date].sort_values("Time")
            day_data_active = day_data_full.dropna(subset=[POWER_COL_OUT])
            
            n_rows = len(day_data_full) 
            day_intervals.append(n_rows)
            
            data_start_row = 3
            merge_start = data_start_row
            merge_end = merge_start + n_rows - 1

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

            # Merge UTC column (Starts at row 3 now)
            ws.merge_cells(start_row=merge_start, start_column=col_start, end_row=merge_end, end_column=col_start)
            ws.cell(row=merge_start, column=col_start, value=date_str_full).alignment = Alignment(horizontal="center", vertical="center")

            # Fill data (starts at row 3)
            for idx, r in enumerate(day_data_full.itertuples(), start=merge_start):
                ws.cell(row=idx, column=col_start+1, value=r.Time)
                # Use np.nan instead of value=None for openpyxl compatibility 
                ws.cell(row=idx, column=col_start+2, value=getattr(r, POWER_COL_OUT) if not pd.isna(getattr(r, POWER_COL_OUT)) else None) 
                ws.cell(row=idx, column=col_start+3, value=r.kW if not pd.isna(r.kW) else None)

            # Summary stats
            stats_row_start = merge_end + 1
            sum_w = day_data_active[POWER_COL_OUT].sum()
            mean_w = day_data_active[POWER_COL_OUT].mean()
            max_w = day_data_active[POWER_COL_OUT].max()
            sum_kw = day_data_active['kW'].sum()
            mean_kw = day_data_active['kW'].mean()
            max_kw = day_data_active['kW'].max()

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

            col_start += 4

        # Add Line Chart for Individual Sheet
        if dates:
            chart = LineChart()
            chart.title = f"Daily 10-Minute Absolute Power Profile - {sheet_name}"
            chart.y_axis.title = "kW"
            chart.x_axis.title = "Time"
            
            chart.height = 12.5 
            chart.width = 23    

            max_rows = max(day_intervals)
            first_time_col = 2
            categories_ref = Reference(ws, min_col=first_time_col, min_row=3, max_row=2 + max_rows)

            col_start = 1
            for i, n_rows in enumerate(day_intervals):
                data_ref = Reference(ws, min_col=col_start+3, min_row=3, max_col=col_start+3, max_row=2+n_rows)
                date_title_str = dates[i].strftime('%d-%b')
                s = Series(values=data_ref, title=date_title_str)
                chart.series.append(s)
                col_start += 4

            chart.set_categories(categories_ref)
            ws.add_chart(chart, f'G{max_row_used+2}')

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
                ws.cell(row=row, column=2, value=max_kw).number_format = numbers.FORMAT_NUMBER_00

            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 15

    # -----------------------------
    # CREATE "TOTAL" SHEET
    # -----------------------------
    if total_sheet_data:
        ws_total = wb.create_sheet("Total")
        
        sorted_dates = sorted(total_sheet_data.keys())
        # Re-derive sheet names list in case sheets_dict was empty for some reason
        all_sheet_names = sorted(list(set(sheet for date_data in total_sheet_data.values() for sheet in date_data.keys())))
        
        headers = ["Date"] + all_sheet_names + ["Total Load"]
        
        for col_idx, header_text in enumerate(headers, 1):
            cell = ws_total.cell(row=1, column=col_idx, value=header_text)
            cell.font = title_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border
            ws_total.column_dimensions[get_column_letter(col_idx)].width = 20

        for row_idx, date_obj in enumerate(sorted_dates, 2):
            date_cell = ws_total.cell(row=row_idx, column=1, value=date_obj.strftime('%Y-%m-%d'))
            date_cell.border = thin_border
            date_cell.alignment = Alignment(horizontal="center")
            
            row_total_load = 0
            
            for col_idx, sheet_name in enumerate(all_sheet_names, 2):
                val = total_sheet_data[date_obj].get(sheet_name, 0)
                if pd.isna(val): val = 0
                
                cell = ws_total.cell(row=row_idx, column=col_idx, value=val)
                cell.number_format = numbers.FORMAT_NUMBER_00
                cell.border = thin_border
                row_total_load += val
            
            total_cell = ws_total.cell(row=row_idx, column=len(all_sheet_names) + 2, value=row_total_load)
            total_cell.number_format = numbers.FORMAT_NUMBER_00
            total_cell.border = thin_border
            total_cell.font = Font(bold=True)

        # Add Chart to Total Sheet
        if sorted_dates:
            chart_total = LineChart()
            chart_total.title = "Daily Max Load Overview" 
            chart_total.y_axis.title = "Max Power (kW)"
            chart_total.x_axis.title = "Date"
            
            chart_total.height = 15
            chart_total.width = 30
            
            data_max_row = len(sorted_dates) + 1
            total_cols = len(all_sheet_names) + 2
            
            # Data from all sheets + Total Load column
            data_ref = Reference(ws_total, min_col=2, min_row=1, max_col=total_cols, max_row=data_max_row)
            chart_total.add_data(data_ref, titles_from_data=True)

            for s in chart_total.series:
                s.smooth = False
            
            cats_ref = Reference(ws_total, min_col=1, min_row=2, max_row=data_max_row)
            chart_total.set_categories(cats_ref)
            
            ws_total.add_chart(chart_total, "B" + str(data_max_row + 3))

    stream = BytesIO()
    if 'Sheet' in wb.sheetnames and len(wb.sheetnames) > len(sheets_dict) + (1 if total_sheet_data else 0):
        wb.remove(wb['Sheet'])
        
    wb.save(stream)
    stream.seek(0)
    return stream

# -------------------------------------
# --- Functions from Code 1.py (Consolidation) ---
# -------------------------------------

def process_uploaded_files(uploaded_files, file_configs):
    """
    Reads multiple CSV files, extracts configured columns, cleans PSum data, 
    and returns a dictionary of DataFrames based on individual file configurations.
    """
    processed_data = {}
    
    for i, uploaded_file in enumerate(uploaded_files):
        filename = uploaded_file.name
        config = file_configs[i] 
        
        try:
            date_col_index = excel_col_to_index(config['date_col_str'])
            time_col_index = excel_col_to_index(config['time_col_str'])
            ps_um_col_index = excel_col_to_index(config['psum_col_str'])
            
            columns_to_extract = {
                date_col_index: 'Date',
                time_col_index: 'Time',
                ps_um_col_index: PSUM_OUTPUT_NAME
            }
            col_indices = list(columns_to_extract.keys())
            
            if len(set(col_indices)) != 3:
                st.error(f"Error for file **{filename}**: Date, Time, and PSum must be extracted from three unique column indices. Check columns {config['date_col_str']}, {config['time_col_str']}, {config['psum_col_str']}.")
                continue
                
            header_index = int(config['start_row_num']) - 1 
            date_format_string = DATE_FORMAT_MAP.get(config['selected_date_format'])
            separator = config['delimiter_input']
            
            df_full = pd.read_csv(
                uploaded_file, 
                header=header_index, 
                encoding='ISO-8859-1', 
                low_memory=False,
                sep=separator
            )
            
            max_index = max(col_indices)
            if df_full.shape[1] < max_index + 1:
                st.error(f"File **{filename}** failed to read data correctly. It only has {df_full.shape[1]} columns. This usually means the **CSV Delimiter** ('{separator}') is incorrect for this file.")
                continue

            df_extracted = df_full.iloc[:, col_indices].copy()
            
            temp_cols = {
                k: v for k, v in columns_to_extract.items()
            }
            df_extracted.columns = temp_cols.values()
            
            if PSUM_OUTPUT_NAME in df_extracted.columns:
                df_extracted[PSUM_OUTPUT_NAME] = pd.to_numeric(
                    df_extracted[PSUM_OUTPUT_NAME].astype(str).str.replace(',', '.', regex=False), # Added comma replacement here for robustness
                    errors='coerce' 
                )

            combined_dt_str = df_extracted['Date'].astype(str) + ' ' + df_extracted['Time'].astype(str)

            datetime_series = pd.to_datetime(
                combined_dt_str, 
                errors='coerce',
                format=date_format_string 
            )
            
            valid_dates_count = datetime_series.count()
            if valid_dates_count == 0:
                st.warning(f"File **{filename}**: No valid dates could be parsed. Check the 'Date Format for Parsing' setting (**{config['selected_date_format']}**) and ensure the 'Date' and 'Time' columns contain valid data starting from Row {config['start_row_num']}.")
                continue

            # Output Date is consistently DD/MM/YYYY for the next stage
            df_final = pd.DataFrame({
                'Date': datetime_series.dt.strftime('%d/%m/%Y'), 
                'Time': datetime_series.dt.strftime('%H:%M:%S'),
                PSUM_OUTPUT_NAME: df_extracted[PSUM_OUTPUT_NAME] 
            })

            sheet_name = filename.replace('.csv', '').replace('.', '_').strip()[:31]
            
            processed_data[sheet_name] = df_final
            
        except ValueError as e:
            st.error(f"Configuration Error for file **{filename}**: Invalid column letter entered: {e}. Please use valid Excel column notation (e.g., A, C, AA).")
            continue
        except Exception as e:
            st.error(f"Error processing file **{filename}**. An unexpected error occurred. Error: {e}")
            continue
            
    return processed_data


@st.cache_data
def to_excel_consolidation(data_dict):
    """
    Writes consolidated DataFrames to an in-memory Excel file, 
    setting Date/Time columns to text format (Code 1's method).
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in data_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            text_format = workbook.add_format({'num_format': '@'})
            
            try:
                if 'Date' in df.columns:
                    date_col_index = df.columns.get_loc('Date')
                    worksheet.set_column(date_col_index, date_col_index, 12, text_format) 
                
                if 'Time' in df.columns:
                    time_col_index = df.columns.get_loc('Time')
                    worksheet.set_column(time_col_index, time_col_index, 10, text_format)
            except Exception as e:
                # print(f"Error applying explicit xlsxwriter formats: {e}") # Suppress console output in final app
                pass
            
    output.seek(0)
    return output.getvalue()


# -------------------------------------
# --- Main Streamlit Application Logic ---
# -------------------------------------

def run_analysis(consolidated_data):
    """Handles the execution of Stage 2 analysis and updates session state."""
    
    analysis_results = {}
    total_processed_days = 0
    
    with st.spinner("Analyzing data and generating report..."):
        for sheet_name, df_raw in consolidated_data.items():
            
            processed_df = process_sheet(df_raw, 'Date', 'Time', PSUM_OUTPUT_NAME)
            
            if not processed_df.empty:
                analysis_results[sheet_name] = processed_df
                total_processed_days += len(processed_df['Date'].unique())
            
        if analysis_results:
            chart_data_df, total_sheet_data = prepare_chart_data(analysis_results)
            output_stream = build_output_excel(analysis_results, total_sheet_data)

            # Store results in session state
            st.session_state['analysis_results'] = analysis_results
            st.session_state['chart_data_df'] = chart_data_df
            st.session_state['total_processed_days'] = total_processed_days
            st.session_state['analysis_excel_stream'] = output_stream
            st.session_state['analysis_status'] = "completed"
        else:
            st.session_state['analysis_status'] = "failed"
            st.session_state['analysis_results'] = None
            st.session_state['chart_data_df'] = None
            st.error("No data was suitable for 10-minute analysis. Please check the raw data for valid power readings.")


def display_analysis_results(chart_data_df, total_processed_days, analysis_excel_stream):
    """Displays the chart and download button in the main body."""
    
    st.header("📈 Daily Max Load Overview (Quick Analysis)")
    
    st.success(f"Analysis complete! Report generated with data for {total_processed_days} day(s) across {len(chart_data_df.columns) - 1} source(s).")
    st.balloons()
    
    # 1. Display Chart in Streamlit
    st.markdown("#### Total Load Trend")
    st.line_chart(chart_data_df, use_container_width=True)
    
    st.markdown("#### Underlying Data (Max Daily kW)")
    st.dataframe(chart_data_df)

    # 2. Download Button
    default_analysis_filename = "10Min_Power_Analysis_Report.xlsx"
    
    st.download_button(
        label="⬇️ Download Analysis Report (10-Min Intervals Excel)",
        data=analysis_excel_stream,
        file_name=default_analysis_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def app():
    st.set_page_config(layout="wide", page_title="EnergyAnalyser Pro")
    
    # --- Title Change Here ---
    st.title("⚡ DataAnalyser Pro for MSB")
    
    # -------------------------------------
    # --- Sidebar Content ---
    # -------------------------------------
    with st.sidebar:
        st.header("Quick Analysis")
        
        if 'consolidated_data' in st.session_state and st.session_state['consolidated_data']:
            st.info(f"Consolidated data for {len(st.session_state['consolidated_data'])} source(s) is ready.")
            
            # Button for Stage 2 (Analysis) - This is the main change
            if st.button("🚀 Run 10-Minute Power Analysis", key="sidebar_run_analysis"):
                run_analysis(st.session_state['consolidated_data'])
                # Rerun the app to show results in the main body
                st.rerun() 
        else:
            st.warning("Upload and consolidate files in the main body first.")

    
    # -------------------------------------
    # --- Main Content: Stage 1 Consolidation ---
    # -------------------------------------
    
    st.markdown("""
        ### **Stage 1: CSV Data Consolidation**
        Upload your raw energy data CSV files to extract **Date**, **Time**, and **PSum** and consolidate them into an intermediate dataset.
    """)

    uploaded_files = st.file_uploader(
        "Choose up to 10 CSV files", 
        type=["csv"], 
        accept_multiple_files=True
    )
    
    if uploaded_files and len(uploaded_files) > 10:
        st.warning(f"You have uploaded {len(uploaded_files)} files. Only the first 10 will be processed.")
        uploaded_files = uploaded_files[:10]
        
    file_configs = []
    
    if uploaded_files:
        st.header("Individual File Configuration")
        st.warning("Please verify the Delimiter, Start Row, and Column Letters for each file below.")
        
        for i, uploaded_file in enumerate(uploaded_files):
            with st.expander(f"⚙️ Settings for **{uploaded_file.name}**", expanded=i == 0):
                
                # Column Configuration
                st.subheader("Column Letters")
                date_col_str = st.text_input("Date Column Letter", value='A', key=f'date_col_str_{i}')
                time_col_str = st.text_input("Time Column Letter", value='B', key=f'time_col_str_{i}')
                ps_um_col_str = st.text_input("PSum Column Letter", value='BI', key=f'psum_col_str_{i}', help="PSum (Total Active Power) column letter in this file (e.g., 'BI').")

                # CSV File Settings
                st.subheader("CSV File Parsing")
                delimiter_input = st.text_input("CSV Delimiter (Separator)", value=',', key=f'delimiter_input_{i}', help="The character used to separate values (e.g., ',', ';', or '\\t').")
                start_row_num = st.number_input("Header Row Number", min_value=1, value=3, key=f'start_row_num_{i}', help="The row number that contains the column headers.")
                selected_date_format = st.selectbox("Date Format for Parsing", options=["DD/MM/YYYY", "YYYY-MM-DD"], index=0, key=f'selected_date_format_{i}')
                
                config = {
                    'date_col_str': date_col_str, 'time_col_str': time_col_str, 'psum_col_str': ps_um_col_str,
                    'delimiter_input': delimiter_input, 'start_row_num': start_row_num, 
                    'selected_date_format': selected_date_format,
                }
                file_configs.append(config)
                
        if st.button("🚀 Process & Consolidate Files", key="run_consolidation"):
            processed_data_dict = process_uploaded_files(uploaded_files, file_configs)
            
            if processed_data_dict:
                st.session_state['consolidated_data'] = processed_data_dict
                st.session_state['analysis_status'] = "pending" # Reset analysis status
                
                st.header("Consolidated Raw Data Output")
                
                first_sheet_name = next(iter(processed_data_dict))
                st.subheader(f"Preview of: {first_sheet_name}")
                st.dataframe(processed_data_dict[first_sheet_name].head())
                st.success(f"Successfully processed {len(processed_data_dict)} of {len(uploaded_files)} file(s).")

                # Generate default filename for raw data
                file_names_without_ext = [f.name.rsplit('.', 1)[0] for f in uploaded_files]
                default_filename = "EnergyAnalyser_Consolidated_Raw_Data.xlsx"
                if len(file_names_without_ext) > 1:
                    first_name = file_names_without_ext[0][:17] + "..." if len(file_names_without_ext[0]) > 20 else file_names_without_ext[0]
                    default_filename = f"{first_name}_and_{len(file_names_without_ext) - 1}_More_Consolidated.xlsx"
                elif file_names_without_ext:
                    default_filename = f"{file_names_without_ext[0]}_Consolidated.xlsx"

                custom_filename = st.text_input(
                    "Output Excel Filename (Raw):",
                    value=default_filename,
                    key="output_filename_input_raw",
                    help="Enter the name for the final Excel file with raw extracted data."
                )
                
                excel_data = to_excel_consolidation(processed_data_dict)
                
                st.download_button(
                    label="📥 Download Consolidated Data (Raw)",
                    data=excel_data,
                    file_name=custom_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.error("No data could be successfully processed in Stage 1.")
    
    
    # -------------------------------------
    # --- Main Content: Stage 2 Analysis Results Display ---
    # -------------------------------------
    if 'analysis_status' in st.session_state and st.session_state['analysis_status'] == "completed":
        st.markdown("---")
        display_analysis_results(
            st.session_state['chart_data_df'], 
            st.session_state['total_processed_days'],
            st.session_state['analysis_excel_stream']
        )
    elif 'analysis_status' in st.session_state and st.session_state['analysis_status'] == "failed":
        st.markdown("---")
        st.error("The Quick Analysis failed. Check raw data validity (Stage 1) and ensure power readings are non-zero.")


if __name__ == "__main__":
    # Initialize session state for data persistence between runs
    if 'consolidated_data' not in st.session_state:
        st.session_state['consolidated_data'] = None
    if 'analysis_status' not in st.session_state:
        st.session_state['analysis_status'] = "pending" # pending, completed, or failed
        
    app()
