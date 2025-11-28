import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.utils import get_column_letter

# --- Configuration & Constants ---
MAX_FILES = 10
PSUM_OUTPUT_NAME = "PSum (W)" # Standard name after raw extraction (Code 1 output)
POWER_COL_OUT = 'PSumW'      # Column name used internally in the final aggregated DataFrame (Code 2 logic)

st.set_page_config(
    page_title="Consolidated Energy Analyzer",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Helper Functions (Stage 1: Extraction - From Code 1) ---

def excel_col_to_index(col_str):
    """Converts an Excel column string (e.g., 'A', 'AA') to a 0-based column index."""
    col_str = col_str.upper().strip()
    index = 0
    for char in col_str:
        if 'A' <= char <= 'Z':
            # Calculate the 1-based index (A=1, B=2, ...)
            index = index * 26 + (ord(char) - ord('A') + 1)
        else:
            raise ValueError(f"Invalid character in column string: {col_str}")
    
    # Convert 1-based index to 0-based index for Pandas (A=0, B=1)
    return index - 1

def process_csv(uploaded_file, header_row, date_col_str, time_col_str, psum_col_str):
    """
    Processes a single CSV file, extracts data based on column letters, 
    and returns a DataFrame with standard column names (Date, Time, PSum (W)).
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
        
        # Rename columns to standard names
        df.columns = ["Date", "Time", PSUM_OUTPUT_NAME]

        # Drop rows where PSum is NaN or empty after reading
        df.dropna(subset=[PSUM_OUTPUT_NAME], inplace=True)

        return df, None
    except ValueError as e:
        return None, f"Configuration Error: {e}"
    except IndexError:
        return None, "Index Error: One of the column letters is outside the bounds of the CSV data."
    except Exception as e:
        # Catch pandas read errors, encoding issues, etc.
        return None, f"An unexpected error occurred during CSV parsing: {type(e).__name__}: {e}"

# --- Helper Functions (Stage 2: Analysis & Output - From Code 2) ---

def process_sheet(df, date_col, time_col, psum_col):
    """
    Processes a single DataFrame sheet: combines datetime, rounds timestamps to 10-min intervals,
    filters leading/trailing zeros, and aggregates by 10-min intervals.
    
    Note: This function assumes the input DataFrame columns are already standardized
    to 'Date', 'Time', and PSUM_OUTPUT_NAME ('PSum (W)') from Stage 1.
    """
    # 1. Combine Date and Time columns into a single timestamp series
    combined_dt_series = df[date_col].astype(str) + ' ' + df[time_col].astype(str)
    
    # Convert the combined string to datetime objects
    df['Timestamp'] = pd.to_datetime(combined_dt_series, errors="coerce", dayfirst=True)
    timestamp_col = 'Timestamp'
    
    # 2. Clean and convert power column (handle commas as decimal separators)
    # The psum_col here is PSUM_OUTPUT_NAME ("PSum (W)")
    power_series = df[psum_col].astype(str).str.replace(',', '', regex=False)
    df[POWER_COL_OUT] = pd.to_numeric(power_series, errors='coerce')
    
    # Drop rows where PSum is NaN or Timestamp is NaT
    df.dropna(subset=[POWER_COL_OUT, timestamp_col], inplace=True)
    
    if df.empty:
        return pd.DataFrame()

    # 3. Time rounding and setting index
    df.set_index(timestamp_col, inplace=True)
    # Round timestamp index to the nearest 10 minutes
    df.index = (df.index + pd.Timedelta(seconds=1)).floor('10min')

    # 4. Filter leading/trailing zero periods
    first_nonzero_idx = df[df[POWER_COL_OUT] > 0].index.min()
    last_nonzero_idx = df[df[POWER_COL_OUT] > 0].index.max()

    if first_nonzero_idx is pd.NaT or last_nonzero_idx is pd.NaT:
        return pd.DataFrame()

    # Resample and aggregate: mean PSum for each 10-minute bucket.
    processed_df = df.loc[first_nonzero_idx:last_nonzero_idx, POWER_COL_OUT].resample('10min').mean().fillna(0).to_frame()

    # 5. Calculate Daily Energy
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
    
    max_row = len(df) + 1
    
    chart = LineChart()
    chart.title = f"10-Minute Average Power ({sheet_name})"
    chart.style = 10
    chart.y_axis.title = 'Average PSum (W)'
    chart.x_axis.title = 'Time Interval'
    
    # Data Reference (PSumW column is C/3)
    data = Reference(ws, min_col=3, min_row=2, max_col=3, max_row=max_row)
    # Categories Reference (Time column is B/2)
    cats = Reference(ws, min_col=2, min_row=2, max_col=2, max_row=max_row)
    
    chart.set_categories(cats)
    series = Series(data, title=POWER_COL_OUT)
    chart.append(series)
    
    # Place chart starting at E2
    chart.width = 15
    chart.height = 10
    ws.add_chart(chart, "E2")


def build_output_excel(result_sheets):
    """
    Generates a multi-sheet Excel file with aggregated data and charts using openpyxl.
    """
    output = BytesIO()
    wb = Workbook()
    
    # Define styles
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Remove the default sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
        
    for sheet_name, df in result_sheets.items():
        # Clean sheet name for Excel
        safe_sheet_name = sheet_name[:31].replace('[', '(').replace(']', ')')
        
        ws = wb.create_sheet(title=safe_sheet_name)
        
        # Write headers
        ws.append(df.columns.tolist())
        
        # Apply header style
        for col_idx, _ in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Write data rows and apply data styles/formatting
        for r_idx, row in enumerate(df.itertuples(index=False), 2):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.border = thin_border
                
                # Apply number formatting
                if c_idx == 3: # PSumW (Column C)
                    cell.number_format = '0.00'
                elif c_idx == 4: # Daily_Energy_kWh (Column D)
                    cell.number_format = '0.0000'
        
        # Auto-adjust column widths
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            ws.column_dimensions[column].width = min(max_length + 2, 40)

        # Create and add the chart
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
        This application combines raw data extraction from CSVs and time-series aggregation/analysis into a single flow.

        1.  **Extraction**: Raw CSV data is read and key columns are extracted based on your settings.
        2.  **Analysis**: The extracted data is aggregated to 10-minute intervals, filtered for active operation, and energy (kWh) is calculated.
        3.  **Output**: A final multi-sheet Excel file with the analyzed data and a chart for each file is generated.
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
    
    # Default settings
    default_header_row = 2
    default_date_col = 'A'
    default_time_col = 'B'
    default_psum_col = 'V' 

    # --- Global Configuration Section in Sidebar ---
    st.sidebar.markdown("---")
    st.sidebar.markdown("**Global Defaults (Apply to all files):**")
    
    global_header = st.sidebar.number_input("Header Row Number:", min_value=1, value=default_header_row, key="g_header", help="The row number where column names are located.")
    global_date = st.sidebar.text_input("Date Column Letter:", value=default_date_col, max_chars=3, key="g_date").upper()
    global_time = st.sidebar.text_input("Time Column Letter:", value=default_time_col, max_chars=3, key="g_time").upper()
    global_psum = st.sidebar.text_input("PSum Column Letter (W):", value=default_psum_col, max_chars=3, key="g_psum").upper()
    
    st.sidebar.markdown("---")
    st.sidebar.markdown("**Individual File Overrides (Optional):**")
    
    # Loop to allow individual overrides (as in Code 1)
    file_configs = {}
    for i, file in enumerate(uploaded_files):
        with st.sidebar.expander(f"⚙️ {file.name}"):
            file_configs[file.name] = {
                'header_row': st.number_input("Header Row:", min_value=1, value=global_header, key=f"h_{i}"),
                'date_col': st.text_input("Date Col:", value=global_date, key=f"d_{i}").upper(),
                'time_col': st.text_input("Time Col:", value=global_time, key=f"t_{i}").upper(),
                'psum_col': st.text_input("PSum Col:", value=global_psum, key=f"p_{i}").upper(), 
            }

    # --- Step 2: In-Memory Processing and Analysis ---

    st.header("2. Process and Analyze Data")
    
    if st.button("Start Combined Extraction & Analysis"):
        
        intermediate_data_dict = {} # Data after Stage 1 (Extraction)
        final_data_dict = {}        # Data after Stage 2 (Analysis)
        
        st.subheader("Stage 1: Raw Data Extraction")
        processing_bar_stage1 = st.progress(0, text="Starting raw data extraction...")
        
        # --- Stage 1: Extraction (Code 1 Logic) ---
        for i, file in enumerate(uploaded_files):
            config = file_configs[file.name]
            
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
                st.success(f"Extracted **{file.name}** successfully.")
            else:
                st.warning(f"**{file.name}** extracted successfully, but no usable data rows were found.")

        processing_bar_stage1.empty()
        st.success("Stage 1: Raw data extraction complete.")
        
        if not intermediate_data_dict:
            st.error("No data was successfully extracted for analysis. Please check your configurations.")
            return

        st.markdown("---")
        st.subheader("Stage 2: Time-Series Analysis and Aggregation")

        # --- Stage 2: Analysis and Aggregation (Code 2 Logic) ---
        processing_bar_stage2 = st.progress(0, text="Starting time-series analysis...")

        for i, (sheet_name, df_intermediate) in enumerate(intermediate_data_dict.items()):
            
            # The columns are standardized by Stage 1:
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
                final_filename = f"{first_name}_and_{len(file_names_without_ext) - 1}_More_Analyzed.xlsx"
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
            st.error("No data could be processed across all stages. Please review your configurations and data quality.")

if __name__ == "__main__":
    app()
