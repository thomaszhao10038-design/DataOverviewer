import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.utils import get_column_letter

# --- Configuration & Constants ---

# Constants from Code 1
MAX_FILES = 10
PSUM_OUTPUT_NAME = "PSum (W)"
# Constants from Code 2
POWER_COL_OUT = 'PSumW'

# --- Helper Functions (From Code 1) ---

def excel_col_to_index(col_str):
    """Converts an Excel column string (e.g., 'A', 'AA', 'BI') to a 0-based column index."""
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

def process_csv(uploaded_file, header_row, date_col_str, time_col_str, psum_col_str):
    """
    Processes a single CSV file based on user-defined column indices and header row.
    Returns a DataFrame with only Date, Time, and PSum (W) columns.
    """
    try:
        # Calculate 0-based indices
        header_index = header_row - 1
        date_col_idx = excel_col_to_index(date_col_str)
        time_col_idx = excel_col_to_index(time_col_str)
        psum_col_idx = excel_col_to_index(psum_col_str)

        # Read CSV with specified header and usecols
        df = pd.read_csv(
            uploaded_file,
            header=header_index,
            usecols=[date_col_idx, time_col_idx, psum_col_idx],
            index_col=False,
            encoding='utf-8',
            errors='replace',
            low_memory=False
        )
        
        # Rename columns to standard names for step 2
        df.columns = ["Date", "Time", PSUM_OUTPUT_NAME]

        # Drop rows where PSum is NaN or empty after reading (common CSV cleanup)
        df.dropna(subset=[PSUM_OUTPUT_NAME], inplace=True)

        return df, None
    except ValueError as e:
        return None, f"Configuration Error: {e}"
    except IndexError:
        return None, "Index Error: One of the column letters is outside the bounds of the CSV data."
    except Exception as e:
        return None, f"An unexpected error occurred during CSV parsing: {e}"

# --- Helper Functions (From Code 2) ---

def pcm_to_wav(pcm_data, sample_rate):
    """Placeholder for the original function, but not used since we aren't calling TTS."""
    # This function is not required for the data processing logic but was in the original Code 2,
    # and since the context suggests it might be a remnant, I'll keep the required data processing
    # functions and omit the unnecessary utility ones to keep the focus.
    # The actual data processing functions are process_sheet, create_chart, and build_output_excel.
    pass

def process_sheet(df, date_col, time_col, psum_col):
    """
    Processes a single DataFrame sheet: combines datetime, rounds timestamps to 10-min intervals,
    filters leading/trailing zeros, and aggregates by 10-min intervals.
    """
    df.columns = df.columns.astype(str).str.strip()
    
    # 1. Combine Date and Time columns into a single timestamp series
    combined_dt_series = df[date_col].astype(str) + ' ' + df[time_col].astype(str)
    
    # Convert the combined string to datetime objects
    # dayfirst=True is kept for robust parsing of date formats like dd/mm/yyyy
    df['Timestamp'] = pd.to_datetime(combined_dt_series, errors="coerce", dayfirst=True)
    timestamp_col = 'Timestamp'
    
    # 2. Clean and convert power column (handle commas as decimal separators)
    # The PSum column is already named PSUM_OUTPUT_NAME from Step 1.
    power_series = df[psum_col].astype(str).str.replace(',', '', regex=False)
    # Convert to numeric, errors='coerce' turns non-numeric into NaN
    df[POWER_COL_OUT] = pd.to_numeric(power_series, errors='coerce')
    
    # Drop rows where PSum is NaN or Timestamp is NaT
    df.dropna(subset=[POWER_COL_OUT, timestamp_col], inplace=True)
    
    if df.empty:
        return pd.DataFrame()

    # 3. Time rounding and setting index
    df.set_index(timestamp_col, inplace=True)
    # Round timestamp index to the nearest 10 minutes
    # Apply a small offset (1 sec) before floor to ensure 10:10:00.001 -> 10:10:00 (instead of 10:00:00 if using floor directly)
    df.index = (df.index + pd.Timedelta(seconds=1)).floor('10min')

    # 4. Filter leading/trailing zero periods
    first_nonzero_idx = df[df[POWER_COL_OUT] > 0].index.min()
    last_nonzero_idx = df[df[POWER_COL_OUT] > 0].index.max()

    if first_nonzero_idx is pd.NaT or last_nonzero_idx is pd.NaT:
        # All readings are zero or missing after cleanup
        return pd.DataFrame()

    # Resample and take the mean of PSum for each 10-minute bucket.
    # The mean is used to aggregate any multiple readings within the 10-min window.
    processed_df = df.loc[first_nonzero_idx:last_nonzero_idx, POWER_COL_OUT].resample('10min').mean().fillna(0).to_frame()

    # Calculate Daily Energy (Wh to kWh)
    # 10-minute average power (W) * (10/60) hours = Wh
    # Wh / 1000 = kWh
    processed_df['Daily_Energy_kWh'] = (processed_df[POWER_COL_OUT] * (10/60) / 1000)
    
    # Extract Date and Time back into columns for Excel formatting
    processed_df['Date'] = processed_df.index.strftime('%Y-%m-%d')
    processed_df['Time'] = processed_df.index.strftime('%H:%M:%S')

    # Reorder columns
    processed_df = processed_df[['Date', 'Time', POWER_COL_OUT, 'Daily_Energy_kWh']]

    return processed_df

def create_chart(ws, df, sheet_name):
    """Creates a line chart for PSumW and adds it to the openpyxl worksheet."""
    
    # Get the row count for data references
    max_row = len(df) + 1
    
    # Create a Line Chart
    chart = LineChart()
    chart.title = f"10-Minute Average Power ({sheet_name})"
    chart.style = 10
    chart.y_axis.title = 'Average PSum (W)'
    chart.x_axis.title = 'Time Interval'
    
    # Data Reference (PSumW column is column 3 (C) in 1-based indexing)
    data = Reference(ws, min_col=3, min_row=2, max_col=3, max_row=max_row)
    
    # Categories Reference (Date/Time combined, Col 1 & 2)
    # We will use the 'Time' column as categories (Col 2 or B)
    cats = Reference(ws, min_col=2, min_row=2, max_col=2, max_row=max_row)
    
    chart.set_categories(cats)
    series = Series(data, title=POWER_COL_OUT)
    chart.append(series)
    
    # Set the chart size and position (starting at E2)
    ws.add_chart(chart, "E2")


def build_output_excel(result_sheets):
    """
    Takes a dictionary of processed DataFrames and generates a multi-sheet Excel file
    with charts using openpyxl.
    """
    output = BytesIO()
    wb = Workbook()
    
    # Define styles
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    thin_border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))

    # Remove the default sheet created by Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
        
    for sheet_name, df in result_sheets.items():
        # Clean sheet name for Excel (max 31 chars, no invalid chars)
        safe_sheet_name = sheet_name[:31].replace('[', '(').replace(']', ')')
        
        ws = wb.create_sheet(title=safe_sheet_name)
        
        # Write headers
        headers = df.columns.tolist()
        ws.append(headers)
        
        # Apply header style
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Write data rows and apply data styles/formatting
        for r_idx, row in enumerate(df.itertuples(index=False), 2):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.border = thin_border
                
                # Apply number formatting for power and energy columns
                if c_idx == 3: # PSumW (Column C)
                    cell.number_format = '0.00'
                elif c_idx == 4: # Daily_Energy_kWh (Column D)
                    cell.number_format = '0.0000'
        
        # Auto-adjust column widths
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter # Get the column name
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            # Cap width at a reasonable max to prevent overly wide columns
            ws.column_dimensions[column].width = min(adjusted_width, 40)

        # Create and add the chart for this sheet
        if not df.empty and len(df) > 1:
            create_chart(ws, df, safe_sheet_name)

    # Save the workbook to the BytesIO stream
    wb.save(output)
    output.seek(0)
    return output

# --- Streamlit Application Main Logic ---

def app():
    st.title("⚡ Consolidated Energy Data Processor")
    st.markdown("""
        Upload your raw energy data CSV files to perform a two-stage process:
        1.  **Extraction & Consolidation**: Extracts Date, Time, and PSum data based on your column settings.
        2.  **Time-Series Analysis**: Rounds data to 10-minute intervals, filters non-operating periods, and calculates daily energy (kWh).
        
        The final output is a single Excel file with analysis, aggregation, and a power trend chart for each uploaded file.
    """)

    # --- Step 1: File Upload and Configuration ---
    
    st.header("1. Upload Raw CSV Files")
    uploaded_files = st.file_uploader(
        "Choose CSV Files (Max 10)",
        type=['csv'],
        accept_multiple_files=True,
        key='csv_uploader'
    )
    
    if not uploaded_files:
        st.info("Please upload your raw data CSV file(s) to begin.")
        return

    if len(uploaded_files) > MAX_FILES:
        st.warning(f"Maximum of {MAX_FILES} files allowed. Only the first {MAX_FILES} will be processed.")
        uploaded_files = uploaded_files[:MAX_FILES]

    st.sidebar.header("Data Extraction Settings")
    st.sidebar.markdown("Define the location of key data in your raw CSV files.")
    
    # Default settings for initial convenience
    default_header_row = 2
    default_date_col = 'A'
    default_time_col = 'B'
    default_psum_col = 'V' # Re-introduced as default, not hardcoded

    # Use a dictionary to store configuration for each file
    file_configs = {}
    
    # --- Configuration Section in Sidebar ---
    
    st.sidebar.markdown("---")
    st.sidebar.markdown("**Global Defaults (Apply to all files):**")
    
    # Global Inputs
    global_header = st.sidebar.number_input("Header Row Number:", min_value=1, value=default_header_row, key="g_header")
    global_date = st.sidebar.text_input("Date Column Letter:", value=default_date_col, max_chars=3, key="g_date").upper()
    global_time = st.sidebar.text_input("Time Column Letter:", value=default_time_col, max_chars=3, key="g_time").upper()
    # PSum Column Letter (W) is re-introduced
    global_psum = st.sidebar.text_input("PSum Column Letter (W):", value=default_psum_col, max_chars=3, key="g_psum").upper()
    
    st.sidebar.markdown("---")
    st.sidebar.markdown("**Individual File Overrides (Optional):**")
    
    # Loop to allow individual overrides (as in Code 1)
    for i, file in enumerate(uploaded_files):
        with st.sidebar.expander(f"⚙️ {file.name}"):
            file_configs[file.name] = {
                'header_row': st.number_input("Header Row:", min_value=1, value=global_header, key=f"h_{i}"),
                'date_col': st.text_input("Date Col:", value=global_date, key=f"d_{i}").upper(),
                'time_col': st.text_input("Time Col:", value=global_time, key=f"t_{i}").upper(),
                # Now uses individual or global input for PSum
                'psum_col': st.text_input("PSum Col:", value=global_psum, key=f"p_{i}").upper(), 
            }

    # --- Step 2: In-Memory Processing ---

    st.header("2. Process Data")
    
    if st.button("Start Combined Processing"):
        
        # Dictionary to hold the intermediate, consolidated data (PSum only)
        intermediate_data_dict = {}
        # Dictionary to hold the final, aggregated and analyzed data
        final_data_dict = {}
        
        st.subheader("Extraction Results (Code 1 Logic)")
        
        # --- Stage 1: Extraction (Code 1 Logic) ---
        processing_bar_stage1 = st.progress(0, text="Starting raw data extraction...")
        
        for i, file in enumerate(uploaded_files):
            config = file_configs[file.name]
            
            # Process the CSV using the defined configuration
            df_intermediate, error = process_csv(
                file, 
                config['header_row'], 
                config['date_col'], 
                config['time_col'], 
                config['psum_col']
            )

            progress = (i + 1) / len(uploaded_files)
            processing_bar_stage1.progress(progress, text=f"Extracting {file.name}...")

            if error:
                st.error(f"Error processing **{file.name}**: {error}")
                continue
            
            if not df_intermediate.empty:
                intermediate_data_dict[file.name] = df_intermediate
                st.success(f"Extracted data from **{file.name}** successfully.")
            else:
                st.warning(f"**{file.name}** extracted successfully, but no usable data rows were found.")

        processing_bar_stage1.empty()
        st.success("Stage 1: Raw data extraction complete.")
        
        if not intermediate_data_dict:
            st.error("No data was successfully extracted for analysis. Please check your configurations.")
            return

        st.markdown("---")
        st.subheader("Analysis and Aggregation Results (Code 2 Logic)")

        # --- Stage 2: Analysis and Aggregation (Code 2 Logic) ---
        processing_bar_stage2 = st.progress(0, text="Starting time-series analysis...")

        for i, (sheet_name, df_intermediate) in enumerate(intermediate_data_dict.items()):
            
            # The columns are guaranteed to be "Date", "Time", and "PSum (W)" from Stage 1
            date_col = "Date"
            time_col = "Time"
            psum_col = PSUM_OUTPUT_NAME # "PSum (W)"

            # Run the time-series processing (10-minute aggregation, filtering)
            processed_df = process_sheet(df_intermediate.copy(), date_col, time_col, psum_col)

            progress = (i + 1) / len(intermediate_data_dict)
            processing_bar_stage2.progress(progress, text=f"Analyzing {sheet_name}...")

            if not processed_df.empty:
                final_data_dict[sheet_name] = processed_df
                st.success(f"Analysis for **{sheet_name}** complete. Resulting data has {len(processed_df)} time intervals.")
            else:
                st.warning(f"**{sheet_name}** had no usable data after time-series filtering (all readings were zero/missing).")

        processing_bar_stage2.empty()
        st.success("Stage 2: Time-series analysis complete.")

        # --- Step 3: Download Final Output ---
        
        if final_data_dict:
            st.balloons()
            st.header("3. Download Final Results")
            st.info("Generating final Excel file with 10-minute aggregation, daily energy calculations, and charts...")
            
            # Generate the final Excel file stream
            output_stream = build_output_excel(final_data_dict)
            
            # Determine a consolidated filename
            file_names_without_ext = [f.name.rsplit('.', 1)[0] for f in uploaded_files]
            if len(file_names_without_ext) > 1:
                first_name = file_names_without_ext[0].split('_')[0] if '_' in file_names_without_ext[0] else file_names_without_ext[0]
                final_filename = f"{first_name}_{len(file_names_without_ext) - 1}_More_Analyzed.xlsx"
            elif file_names_without_ext:
                final_filename = f"{file_names_without_ext[0]}_Analyzed.xlsx"
            else:
                final_filename = "EnergyAnalyser_Analyzed_Output.xlsx"

            st.download_button(
                label="📥 Download Final Analyzed Excel (Multi-Sheet with Charts)",
                data=output_stream,
                file_name=final_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="final_download"
            )
        else:
            st.error("No data could be processed in either stage. Please review your CSV files and column configurations.")

if __name__ == "__main__":
    app()
