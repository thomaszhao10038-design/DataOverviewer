import streamlit as st
import pandas as pd
from io import BytesIO
import re

# --- Configuration for Streamlit Page ---
st.set_page_config(
    page_title="EnergyAnalyser",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Constants for Data Processing ---
PSUM_OUTPUT_NAME = "PSum (W)"

# --- Helper Function for Excel Column Conversion ---
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
            # Calculate the 1-based index
            index = index * 26 + (ord(char) - ord('A') + 1)
        else:
            raise ValueError(f"Invalid character in column string: {col_str}")
    
    # Convert 1-based index to 0-based index for Pandas (A=0, B=1)
    return index - 1

# --- Helper Function to Create Excel in Memory (for Download) ---
def to_excel(data_dict):
    """
    Converts a dictionary of pandas DataFrames into an Excel file in memory (BytesIO).
    Each key in the dictionary becomes a sheet name.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in data_dict.items():
            # Ensure sheet name is valid (max 31 chars, no special chars)
            safe_sheet_name = sheet_name[:31]
            # Replace invalid characters with underscore
            safe_sheet_name = re.sub(r'[\\/:*?\[\]]', '_', safe_sheet_name)
            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
    
    return output.getvalue()

# --- Core Data Processing Function (Extract Raw Data) ---
def process_file(uploaded_file, date_col_idx, time_col_idx, psum_col_idx):
    """
    Reads a CSV file, selects the specified columns by index, and returns a clean DataFrame.
    """
    try:
        # Read the CSV file. Use a robust separator detection or assume common ones.
        df = pd.read_csv(uploaded_file, encoding='latin1', on_bad_lines='skip')
    except Exception as e:
        st.error(f"Error reading {uploaded_file.name}: {e}")
        return None

    # Determine the maximum valid column index
    max_idx = len(df.columns) - 1
    
    # Check if indices are within bounds
    if any(idx > max_idx for idx in [date_col_idx, time_col_idx, psum_col_idx]):
        st.error(f"Configuration Error in {uploaded_file.name}: One or more column indices (Date/Time/PSum) exceed the actual number of columns ({len(df.columns)}). Please check your column letters.")
        return None

    try:
        # Get column names using 0-based indices
        date_col_name = df.columns[date_col_idx]
        time_col_name = df.columns[time_col_idx]
        psum_col_name = df.columns[psum_col_idx]

        # Select and rename columns
        processed_df = df[[date_col_name, time_col_name, psum_col_name]].copy()
        processed_df.columns = ["Date", "Time", PSUM_OUTPUT_NAME]

        # Simple cleaning for PSum: convert to numeric, coercing errors to NaN
        processed_df[PSUM_OUTPUT_NAME] = pd.to_numeric(processed_df[PSUM_OUTPUT_NAME], errors='coerce')
        
        # Drop rows where PSum is NaN after conversion
        processed_df = processed_df.dropna(subset=[PSUM_OUTPUT_NAME])

        return processed_df
    
    except IndexError:
        st.error(f"Index Error processing {uploaded_file.name}. Ensure all column letters are valid and within the file's column count.")
        return None
    except Exception as e:
        st.error(f"General error processing columns for {uploaded_file.name}: {e}")
        return None

# --------------------------------------------------------------------------
# --- STREAMLIT UI LAYOUT ---
# --------------------------------------------------------------------------

st.title("⚡ EnergyAnalyser: Data Consolidation")
st.markdown("""
    Upload your raw energy data CSV files (up to 10) to extract **Date**, **Time**, and **PSum** and consolidate them into a single Excel file.
""")

# --- File Uploader ---
uploaded_files = st.file_uploader(
    "Upload CSV Files (Max 10)",
    type=["csv"],
    accept_multiple_files=True
)

if uploaded_files:
    if len(uploaded_files) > 10:
        st.warning("Only the first 10 uploaded files will be processed.")
        uploaded_files = uploaded_files[:10]

    # --- Sidebar Configuration for Columns ---
    st.sidebar.header("Column Configuration")
    st.sidebar.markdown(
        "Enter the Excel column letter (e.g., `A`, `B`, `C`, etc.) for the data in your uploaded files."
    )
    
    # Use st.form for grouping inputs and preventing excessive reruns
    with st.sidebar.form("column_config_form"):
        # Default values are common in energy loggers
        date_col_str = st.text_input(
            "Date Column Letter:", "A", max_chars=3, help="e.g., A"
        )
        time_col_str = st.text_input(
            "Time Column Letter:", "B", max_chars=3, help="e.g., B"
        )
        psum_col_str = st.text_input(
            "PSum (Power) Column Letter:", "C", max_chars=3, help="e.g., C"
        )
        
        submitted = st.form_submit_button("Process Files")

else:
    st.sidebar.markdown("Upload files to configure settings.")
    # Exit script early if no files are uploaded
    submitted = False
    
# --- Main Logic Execution ---
if uploaded_files and submitted:
    
    # 1. Convert column letters to 0-based indices
    try:
        date_col_idx = excel_col_to_index(date_col_str)
        time_col_idx = excel_col_to_index(time_col_str)
        psum_col_idx = excel_col_to_index(psum_col_str)
    except ValueError as e:
        st.error(f"Invalid column configuration: {e}")
        st.stop()
        
    st.info(f"Processing files using 0-based column indices: Date={date_col_idx}, Time={time_col_idx}, PSum={psum_col_idx}")

    processed_data_dict = {}
    file_names_without_ext = []
    
    progress_bar = st.progress(0)
    total_files = len(uploaded_files)
    
    # 2. Process each uploaded file
    for i, file in enumerate(uploaded_files):
        st.markdown(f"**Processing File {i+1}/{total_files}:** `{file.name}`")
        
        # Check if file has data and process it
        processed_df = process_file(file, date_col_idx, time_col_idx, psum_col_idx)
        
        if processed_df is not None and not processed_df.empty:
            sheet_name = file.name.rsplit('.', 1)[0]
            processed_data_dict[sheet_name] = processed_df
            file_names_without_ext.append(sheet_name)
            st.success(f"Successfully extracted {len(processed_df)} records.")
        else:
            st.warning(f"Skipped file `{file.name}` due to errors or lack of valid data.")
            
        progress_bar.progress((i + 1) / total_files)

    st.markdown("---")
    
    # 3. Handle Output Generation
    if processed_data_dict:
        st.balloons()
        st.success(f"Consolidation complete! {len(processed_data_dict)} file(s) processed successfully.")

        # --- Dynamic Filename Generation ---
        if file_names_without_ext:
            if len(file_names_without_ext) == 1:
                default_filename = f"{file_names_without_ext[0]}_Consolidated.xlsx"
            else:
                first_name = file_names_without_ext[0]
                default_filename = f"{first_name}_and_{len(file_names_without_ext) - 1}_More_Consolidated.xlsx"
        else:
             default_filename = "EnergyAnalyser_Consolidated_Data.xlsx"


        custom_filename = st.text_input(
            "Output Excel Filename:",
            value=default_filename,
            key="output_filename_input_raw",
            help="Enter the name for the final Excel file with raw extracted data."
        )
        
        # Generate Excel file for raw data
        excel_data = to_excel(processed_data_dict)
        
        # Download Button for raw data
        st.download_button(
            label="📥 Download Consolidated Data (Date, Time, PSum)",
            data=excel_data,
            file_name=custom_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Click to download the Excel file with one sheet per uploaded CSV file."
        )

    else:
        st.error("No data could be successfully processed. Please review the error messages above and adjust the configurations in the file settings.")
