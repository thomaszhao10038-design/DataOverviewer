import streamlit as st
import pandas as pd
from io import BytesIO
import numpy as np

# Imports for advanced Excel output (Code 2 functionality)
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.utils import get_column_letter

# --- Configuration for Streamlit Page ---
st.set_page_config(
    page_title="EnergyAnalyser Pro",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Constants for Data Processing ---
# Output column name for raw extraction (used in Code 1 logic)
PSUM_RAW_OUTPUT_NAME = 'PSum (W)' 
# Output column name for processed data (used in Code 2 logic)
PSUM_PROCESSED_OUTPUT_NAME = 'PSumW' 

# ----------------------------------------
# 1. HELPER FUNCTIONS (From Code 1)
# ----------------------------------------

def excel_col_to_index(col_str):
    """
    Converts an Excel column string (e.g., 'A', 'AA', 'BI') to a 0-based column index.
    Raises a ValueError if the string is invalid.
    """
    col_str = col_str.upper().strip()
    index = 0
    for char in col_str:
        if 'A' <= char <= 'Z':
            # Calculate the 1-based index 
            index = index * 26 + (ord(char) - ord('A') + 1)
        else:
            raise ValueError(f"Invalid character in column string: {col_str}")
    
    # Convert 1-based index to 0-based index for Pandas (A=0, B=1)
    return index - 1

def to_excel_raw(dataframes_dict):
    """
    Writes a dictionary of DataFrames to a single Excel file, 
    with each key as a sheet name (for raw output).
    Uses pandas to_excel for simplicity.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in dataframes_dict.items():
            # Only save the raw data columns we extracted
            df[['Date', 'Time', PSUM_RAW_OUTPUT_NAME]].to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()


# ----------------------------------------
# 2. DATA PROCESSING (From Code 2 - Advanced Cleaning)
# ----------------------------------------

def process_sheet(df_raw, date_col, time_col, psum_col):
    """
    Processes a single DataFrame sheet: combines Date/Time, cleans power data, 
    rounds timestamps to 10-minute intervals, calculates the mean power, 
    and filters out leading/trailing zero-periods.
    
    This function expects date_col, time_col, and psum_col to be the names 
    of the columns in df_raw, which will be 'Date', 'Time', and 'PSum (W)' 
    based on the raw extraction.
    """
    df = df_raw.copy()
    df.columns = df.columns.astype(str).str.strip()
    
    # 1. Combine Date and Time columns into a single timestamp string/series
    combined_dt_series = df[date_col].astype(str) + ' ' + df[time_col].astype(str)
    
    # Convert the combined string to datetime objects
    # dayfirst=True is kept for robust parsing of date formats like dd/mm/yyyy
    df['Timestamp'] = pd.to_datetime(combined_dt_series, errors="coerce", dayfirst=True)
    
    # 2. Clean and convert power column (handle commas as decimal separators for potential Euro format)
    power_series = df[psum_col].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    df[PSUM_PROCESSED_OUTPUT_NAME] = pd.to_numeric(power_series, errors='coerce')
    
    # Drop rows where Timestamp or Power is invalid/missing
    df.dropna(subset=['Timestamp', PSUM_PROCESSED_OUTPUT_NAME], inplace=True)
    if df.empty:
        return pd.DataFrame()

    # 3. Filter out leading/trailing zero-periods (finding the first and last non-zero reading)
    # Convert to numeric, errors='coerce' turns non-numeric to NaN
    power_values = pd.to_numeric(df[PSUM_PROCESSED_OUTPUT_NAME], errors='coerce').fillna(0)
    
    non_zero_indices = power_values[power_values != 0].index
    
    if non_zero_indices.empty:
        # Only zeros or invalid data
        return pd.DataFrame()
        
    first_non_zero_idx = non_zero_indices.min()
    last_non_zero_idx = non_zero_indices.max()
    
    # Keep only the data range from the first non-zero reading to the last non-zero reading
    df_filtered = df.loc[first_non_zero_idx:last_non_zero_idx].copy()
    
    # 4. Resample to 10-minute intervals and calculate the mean power
    df_filtered.set_index('Timestamp', inplace=True)
    # '10Min' groups data into 10-minute bins and 'mean()' calculates the average power in that bin
    df_resampled = df_filtered[PSUM_PROCESSED_OUTPUT_NAME].resample('10Min').mean().to_frame()
    
    # 5. Reset index and format output
    df_resampled.reset_index(inplace=True)
    df_resampled.rename(columns={'Timestamp': 'Timestamp (10-Min Avg)'}, inplace=True)
    
    # Add back separate Date and Time columns for Excel readability
    df_resampled['Date'] = df_resampled['Timestamp (10-Min Avg)'].dt.strftime('%Y-%m-%d')
    df_resampled['Time'] = df_resampled['Timestamp (10-Min Avg)'].dt.strftime('%H:%M:%S')
    
    # Reorder columns for final output
    return df_resampled[['Date', 'Time', PSUM_PROCESSED_OUTPUT_NAME, 'Timestamp (10-Min Avg)']]


# ----------------------------------------
# 3. ADVANCED EXCEL WRITING (From Code 2 - Styling and Charts)
# ----------------------------------------

def to_excel_with_charts(result_sheets):
    """
    Writes a dictionary of processed DataFrames to an Excel file with specific styling
    and generates a line chart for each sheet.
    """
    output = BytesIO()
    wb = Workbook()
    
    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    # Define styles
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    data_font = Font(color="000000")
    thin_border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))
    
    # Number format for PSum (W)
    num_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
    
    for sheet_name, df in result_sheets.items():
        ws = wb.create_sheet(title=sheet_name[:31]) # Truncate sheet name to max 31 chars
        
        # 1. Write Header
        headers = ['Date', 'Time', f'{PSUM_PROCESSED_OUTPUT_NAME} (W)']
        ws.append(headers)
        
        # Apply header styling
        for col_idx, cell in enumerate(ws[1]):
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # 2. Write Data
        for r_idx, row in df.iterrows():
            # Use only Date, Time, and Power value for the sheet
            ws.append([row['Date'], row['Time'], row[PSUM_PROCESSED_OUTPUT_NAME]])
            
        # 3. Apply Data Styling and Column Widths
        date_col_letter = get_column_letter(1)
        time_col_letter = get_column_letter(2)
        power_col_letter = get_column_letter(3)

        ws.column_dimensions[date_col_letter].width = 15
        ws.column_dimensions[time_col_letter].width = 15
        ws.column_dimensions[power_col_letter].width = 18

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=3):
            for cell in row:
                cell.font = data_font
                cell.border = thin_border
            # Format Power column (Column C)
            power_cell = row[2]
            power_cell.number_format = num_format
            
        # 4. Generate Chart (Line Chart of PSum over Time)
        chart = LineChart()
        chart.title = f"Power Consumption ({sheet_name})"
        chart.style = 10
        chart.x_axis.title = "Time"
        chart.y_axis.title = "Power (W)"
        
        # Data range (C2 to last row of Power data)
        data = Reference(ws, min_col=3, min_row=2, max_col=3, max_row=ws.max_row)
        # Category labels (B2 to last row of Time data)
        cats = Reference(ws, min_col=2, min_row=2, max_col=2, max_row=ws.max_row)

        series = Series(data, title=f"{PSUM_PROCESSED_OUTPUT_NAME} (W)")
        chart.series.append(series)

        chart.set_categories(cats)
        
        # Position the chart below the data
        chart_start_row = ws.max_row + 2 
        ws.add_chart(chart, f'A{chart_start_row}')


    # Save the workbook to the BytesIO stream
    wb.save(output)
    return output.getvalue()


# ----------------------------------------
# 4. STREAMLIT APPLICATION LOGIC
# ----------------------------------------

# --- App Title and Description ---
st.title("⚡ EnergyAnalyser Pro: Data Consolidation & 10-Min Processing")
st.markdown("""
    This tool allows you to upload multiple raw energy data CSV files, configure the Date, Time, and PSum columns for each, 
    and generate **two** consolidated Excel outputs:
    1.  **Raw Data:** Extracted Date, Time, and PSum (one sheet per file).
    2.  **Processed Data:** 10-minute average power data, cleaned, filtered for active periods, styled, and charted (one sheet per file).
""")

# --- File Uploader ---
uploaded_files = st.file_uploader(
    "Upload Raw Energy Data CSV Files (Max 10)",
    type=['csv'],
    accept_multiple_files=True
)

# --- State Management for File Configuration ---
if 'file_configs' not in st.session_state:
    st.session_state.file_configs = {}

if uploaded_files:
    
    # Clean up configs for files that were removed
    current_files = {file.name for file in uploaded_files}
    st.session_state.file_configs = {
        k: v for k, v in st.session_state.file_configs.items() if k in current_files
    }
    
    st.sidebar.header("File Settings")
    st.sidebar.markdown("Configure the column indices for each uploaded file.")

    for i, uploaded_file in enumerate(uploaded_files):
        file_name = uploaded_file.name
        
        # Initialize config if new file
        if file_name not in st.session_state.file_configs:
            st.session_state.file_configs[file_name] = {
                'header_row': 1,
                'data_start_row': 2,
                'date_col': 'A',
                'time_col': 'B',
                'psum_col': 'C'
            }

        with st.sidebar.expander(f"⚙️ Settings for: {file_name}", expanded=(i == 0)):
            config = st.session_state.file_configs[file_name]

            # Custom Keying to ensure independent widgets
            def config_key(key): return f"{file_name}_{key}"

            config['header_row'] = st.number_input(
                "Header Row Number:", 
                min_value=1, max_value=50, value=config['header_row'], 
                key=config_key('header_row'),
                help="The 1-based row number containing the column headers (e.g., 'Date', 'Time')."
            )
            config['data_start_row'] = st.number_input(
                "Data Start Row Number:", 
                min_value=2, max_value=50, value=config['data_start_row'], 
                key=config_key('data_start_row'),
                help="The 1-based row number where the actual data begins (should be > Header Row)."
            )
            
            st.markdown("---")
            st.markdown("**Column Letters (e.g., A, C, BI):**")
            
            config['date_col'] = st.text_input(
                "Date Column Letter:", 
                value=config['date_col'], 
                key=config_key('date_col')
            )
            config['time_col'] = st.text_input(
                "Time Column Letter:", 
                value=config['time_col'], 
                key=config_key('time_col')
            )
            config['psum_col'] = st.text_input(
                "PSum Column Letter:", 
                value=config['psum_col'], 
                key=config_key('psum_col'),
                help="The column containing the Active Power (PSum, P, Power, etc.)."
            )
            
            # Save config back (though Streamlit handles it via session_state updates from keys)
            st.session_state.file_configs[file_name] = config

    st.markdown("---")
    
    # --- Processing Button ---
    if st.button("🚀 Process and Generate Excel Files", type="primary"):
        
        processed_raw_data_dict = {}
        processed_converted_data_dict = {}
        successful_files = []
        
        status_bar = st.progress(0, text="Starting data processing...")

        for i, uploaded_file in enumerate(uploaded_files):
            file_name = uploaded_file.name
            config = st.session_state.file_configs.get(file_name)
            
            if not config:
                st.warning(f"Configuration missing for {file_name}. Skipping.")
                continue

            status_bar.progress((i + 1) / len(uploaded_files), text=f"Processing {file_name}...")
            st.subheader(f"Processing: {file_name}")

            try:
                # Calculate 0-based indices and skip rows
                header_row_index = config['header_row'] - 1
                row_skip = config['data_start_row'] - 1 
                
                if row_skip < header_row_index:
                    st.error(f"Data Start Row ({config['data_start_row']}) must be after or the same as Header Row ({config['header_row']}). Please adjust settings for {file_name}.")
                    continue
                
                # --- Read CSV with correct header row ---
                df = pd.read_csv(uploaded_file, header=header_row_index, low_memory=False)
                # Clean column names (strip whitespace)
                df.columns = df.columns.astype(str).str.strip()
                
                # --- Convert Excel letters to 0-based column indices ---
                date_col_index = excel_col_to_index(config['date_col'])
                time_col_index = excel_col_to_index(config['time_col'])
                psum_col_index = excel_col_to_index(config['psum_col'])
                
                # --- Validation ---
                max_cols = df.shape[1]
                if any(idx < 0 or idx >= max_cols for idx in [date_col_index, time_col_index, psum_col_index]):
                    st.error(f"One or more column letters for {file_name} refer to a column outside the CSV's bounds ({max_cols} columns found).")
                    continue
                
                # --- Extract RAW Data (Code 1 Logic) ---
                
                # Use iloc to select data starting from data_start_row
                df_raw_extracted = pd.DataFrame({
                    'Date': df.iloc[row_skip - header_row_index:, date_col_index].values,
                    'Time': df.iloc[row_skip - header_row_index:, time_col_index].values,
                    PSUM_RAW_OUTPUT_NAME: df.iloc[row_skip - header_row_index:, psum_col_index].values
                })
                
                # Drop rows where all three key columns are NaN
                df_raw_extracted.dropna(how='all', inplace=True)

                if df_raw_extracted.empty:
                    st.warning(f"Raw extraction for {file_name} resulted in no usable data after removing empty rows.")
                    continue
                
                # Store for Raw Output
                sheet_name = file_name.replace('.csv', '').replace('.', '_')[:31] # Max 31 chars
                processed_raw_data_dict[sheet_name] = df_raw_extracted
                
                # --- Process CONVERTED Data (Code 2 Logic) ---
                
                # The raw extracted DataFrame is passed to process_sheet
                # The column names are fixed based on the DataFrame creation above
                processed_df = process_sheet(
                    df_raw_extracted, 
                    date_col='Date', 
                    time_col='Time', 
                    psum_col=PSUM_RAW_OUTPUT_NAME
                )

                if not processed_df.empty:
                    processed_converted_data_dict[sheet_name] = processed_df
                    st.success(f"✅ {file_name}: Extracted raw data and successfully generated 10-minute processed data.")
                    successful_files.append(file_name)
                else:
                    st.warning(f"⚠️ {file_name}: Raw data extracted, but 10-minute processing resulted in no usable data (might be all zeros or parsing issues).")


            except ValueError as e:
                st.error(f"Error processing {file_name} due to column letter conversion: {e}")
            except Exception as e:
                st.error(f"An unexpected error occurred while processing {file_name}: {e}")
                import traceback
                st.exception(e)

        status_bar.empty()
        st.markdown("---")

        if successful_files:
            st.balloons()
            st.success(f"Processing complete for {len(successful_files)} of {len(uploaded_files)} files. Ready for download!")

            col1, col2 = st.columns(2)

            # --- RAW DATA OUTPUT (Code 1 Output) ---
            with col1:
                st.subheader("1. Download Raw Extracted Data")
                st.markdown("This file contains the original Date, Time, and PSum values, one sheet per file.")
                
                file_names_without_ext = [f.replace('.csv', '') for f in successful_files]
                
                if len(file_names_without_ext) > 1:
                    first_name = file_names_without_ext[0]
                    default_filename_raw = f"{first_name}_and_{len(file_names_without_ext) - 1}_More_Raw.xlsx"
                elif file_names_without_ext:
                    default_filename_raw = f"{file_names_without_ext[0]}_Raw.xlsx"
                else:
                    default_filename_raw = "EnergyAnalyser_Raw_Data.xlsx"

                custom_filename_raw = st.text_input(
                    "Raw Excel Filename:",
                    value=default_filename_raw,
                    key="output_filename_input_raw",
                    help="Enter the name for the final Excel file with raw extracted data."
                )
                
                # Generate Excel file for raw data
                excel_data_raw = to_excel_raw(processed_raw_data_dict)
                
                # Download Button for raw data
                st.download_button(
                    label="📥 Download Raw Consolidated Data",
                    data=excel_data_raw,
                    file_name=custom_filename_raw,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Click to download the Excel file with one sheet per uploaded CSV file."
                )

            # --- PROCESSED DATA OUTPUT (Code 2 Output) ---
            with col2:
                st.subheader("2. Download 10-Min Processed Data")
                st.markdown("This file contains the 10-minute averaged data with styling and charts.")
                
                file_names_without_ext = [f.replace('.csv', '') for f in processed_converted_data_dict.keys()]
                
                if len(file_names_without_ext) > 1:
                    first_name = file_names_without_ext[0]
                    default_filename_proc = f"{first_name}_and_{len(file_names_without_ext) - 1}_Processed.xlsx"
                elif file_names_without_ext:
                    default_filename_proc = f"{file_names_without_ext[0]}_Processed.xlsx"
                else:
                    default_filename_proc = "EnergyAnalyser_Processed_Data.xlsx"

                custom_filename_proc = st.text_input(
                    "Processed Excel Filename:",
                    value=default_filename_proc,
                    key="output_filename_input_proc",
                    help="Enter the name for the final Excel file with 10-minute processed data."
                )

                # Generate Excel file for processed data with openpyxl styling/charts
                excel_data_proc = to_excel_with_charts(processed_converted_data_dict)
                
                # Download Button for processed data
                st.download_button(
                    label="📥 Download Processed & Charted Data",
                    data=excel_data_proc,
                    file_name=custom_filename_proc,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Click to download the Excel file with 10-minute aggregated data, charts, and styling."
                )
                
        else:
            st.error("No data could be successfully processed. Please review the error messages above and adjust the configurations in the file settings.")
else:
    st.sidebar.markdown("Upload files to configure settings.")
