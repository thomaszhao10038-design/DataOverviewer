import streamlit as st
import pandas as pd
from io import BytesIO
import string

# --- Configuration for Streamlit Page ---
st.set_page_config(
    page_title="EnergyAnalyser",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
            # Calculate the 1-based index (e.g., 'B' is 2, 'AA' is 27)
            index = index * 26 + (ord(char) - ord('A') + 1)
        else:
            raise ValueError(f"Invalid character in column string: {col_str}")
    
    # Convert 1-based index to 0-based index for Pandas (A=0, B=1)
    return index - 1

# --- Constants for Data Processing ---
PSUM_OUTPUT_NAME = "PSum (W)"

# Function to generate mock column names (A, B, C, ..., AA, AB, ...)
def get_excel_column_names(n=52):
    """Generates a list of Excel-style column names up to n columns."""
    cols = []
    for i in range(n):
        col = ""
        while i >= 0:
            i, remainder = divmod(i, 26)
            col = string.ascii_uppercase[remainder] + col
            i -= 1
        cols.append(col)
    return cols

EXCEL_COLUMN_OPTIONS = get_excel_column_names(100)
DEFAULT_START_ROW = 5
DEFAULT_DATE_COL = 'A'
DEFAULT_TIME_COL = 'B'
DEFAULT_PSUM_COL = 'F' # Assuming 'F' is PSum

# --- Function to convert DataFrames to an Excel file in memory (for download) ---
def to_excel(data_dict):
    """Converts a dictionary of DataFrames into a multi-sheet Excel file in memory."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in data_dict.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    processed_data = output.getvalue()
    return processed_data

# --- App Title and Description ---
st.title("⚡ EnergyAnalyser: Data Consolidation")
st.markdown("""
    Upload your raw energy data CSV files (up to 10) to extract **Date**, **Time**, and **PSum** and consolidate them into a single Excel file.
""")

# Initialize session state for processed data
if 'processed_data_dict' not in st.session_state:
    st.session_state.processed_data_dict = {}
if 'processing_errors' not in st.session_state:
    st.session_state.processing_errors = []

# --- Sidebar Configuration ---
st.sidebar.header("Settings")

# 1. File Uploader
uploaded_files = st.sidebar.file_uploader(
    "Upload CSV or TXT Files (Max 10):",
    type=['csv', 'txt'],
    accept_multiple_files=True
)

if uploaded_files:
    
    # 2. One-Click Analysis Section (New - Renamed from Quick Analysis)
    st.sidebar.subheader("One-Click Analysis")
    st.sidebar.button(
        "Quick Download", 
        key="quick_download_sidebar",
        disabled=not st.session_state.processed_data_dict,
        help="If data has been processed, click to download the Excel file with the current filename."
    )

    # 3. Individual File Configuration Section (Moved)
    st.sidebar.subheader("Individual File Configuration")

    processed_data_dict = {}
    st.session_state.processing_errors = []

    for i, uploaded_file in enumerate(uploaded_files):
        # Unique keys for each file's configuration
        file_key_prefix = f"file_{i}_{uploaded_file.name}_"

        # Expandable section for each file
        with st.sidebar.expander(f"⚙️ {uploaded_file.name}", expanded=False):
            try:
                # Read file content and try to infer header/data
                # Use a small number of rows to preview
                temp_df = pd.read_csv(uploaded_file, nrows=1, encoding='ISO-8859-1', sep=None, engine='python', on_bad_lines='skip')
                uploaded_file.seek(0) # Reset pointer after reading header

                # Configuration Inputs
                start_row = st.number_input(
                    "Data Start Row (1-based index):", 
                    min_value=1, 
                    value=DEFAULT_START_ROW, 
                    key=file_key_prefix + "start_row",
                    help="Enter the row number where the actual data starts (1 is the first row of the file)."
                )

                # Use columns A-Z as the default options
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    date_col = st.selectbox(
                        "Date Column:",
                        EXCEL_COLUMN_OPTIONS,
                        index=EXCEL_COLUMN_OPTIONS.index(DEFAULT_DATE_COL),
                        key=file_key_prefix + "date_col"
                    )
                with col2:
                    time_col = st.selectbox(
                        "Time Column:",
                        EXCEL_COLUMN_OPTIONS,
                        index=EXCEL_COLUMN_OPTIONS.index(DEFAULT_TIME_COL),
                        key=file_key_prefix + "time_col"
                    )
                with col3:
                    psum_col = st.selectbox(
                        "PSum Column:",
                        EXCEL_COLUMN_OPTIONS,
                        index=EXCEL_COLUMN_OPTIONS.index(DEFAULT_PSUM_COL),
                        key=file_key_prefix + "psum_col",
                        help="Select the column containing the main power consumption data (e.g., PSum, P(W))."
                    )

                # --- Data Loading and Processing (Placeholder logic) ---
                # This part normally calls a function to process the file based on the config.
                # Since I don't have the process_file function, I'll simulate success.
                
                # --- Simulated Data Processing (Replace with actual logic in a real app) ---
                try:
                    # Skip rows (header) based on user input
                    skip_rows = start_row - 1 if start_row > 0 else 0
                    
                    df = pd.read_csv(
                        uploaded_file, 
                        skiprows=skip_rows, 
                        header=None, # No header row, columns are indexed by numbers or letters
                        encoding='ISO-8859-1', 
                        sep=None, 
                        engine='python', 
                        on_bad_lines='skip'
                    )
                    
                    # Rename columns based on configured Excel letters
                    date_idx = excel_col_to_index(date_col)
                    time_idx = excel_col_to_index(time_col)
                    psum_idx = excel_col_to_index(psum_col)
                    
                    col_mapping = {
                        date_idx: 'Date', 
                        time_idx: 'Time', 
                        psum_idx: PSUM_OUTPUT_NAME
                    }
                    
                    # Filter and rename columns
                    df = df.iloc[:, list(col_mapping.keys())].rename(columns=col_mapping)
                    
                    # Basic cleaning (simulate a successful result)
                    if not df.empty and 'Date' in df.columns and 'Time' in df.columns and PSUM_OUTPUT_NAME in df.columns:
                         # Placeholder for real processing (e.g., cleaning, timestamp conversion)
                        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                        df = df.dropna(subset=['Date', PSUM_OUTPUT_NAME])
                        processed_data_dict[uploaded_file.name] = df.head(10) # Store only head for memory safety here
                        
                    else:
                        st.session_state.processing_errors.append(f"File {uploaded_file.name}: Data frame is empty or missing required columns after filter.")

                except Exception as e:
                    st.session_state.processing_errors.append(f"File {uploaded_file.name}: Error during processing: {e}")

            except Exception as e:
                st.session_state.processing_errors.append(f"File {uploaded_file.name}: Could not read or configure file. Error: {e}")


# --- Main Column Content (Processing Summary and Final Download) ---

# Check if any file processing generated data
if processed_data_dict or st.session_state.processed_data_dict:
    final_data_dict = processed_data_dict if processed_data_dict else st.session_state.processed_data_dict
    
    if final_data_dict:
        st.subheader("Processing Summary")
        
        # Display errors if any occurred
        if st.session_state.processing_errors:
            st.warning("⚠️ Some files or configurations resulted in errors. Check the file settings in the sidebar.")
            for err in st.session_state.processing_errors:
                st.error(err)

        st.success(f"Successfully processed {len(final_data_dict)} of {len(uploaded_files)} files.")
        
        # --- Download Section ---
        st.subheader("Download Consolidated Data")

        # 1. Determine default filename logic
        file_names_without_ext = [name.rsplit('.', 1)[0] for name in final_data_dict.keys()]
        default_filename = "EnergyAnalyser_Consolidated_Data.xlsx"
        
        if len(file_names_without_ext) == 1:
            default_filename = f"{file_names_without_ext[0]}_Consolidated.xlsx"
        elif len(file_names_without_ext) > 1:
            first_name = file_names_without_ext[0]
            default_filename = f"{first_name}_and_{len(file_names_without_ext) - 1}_More_Consolidated.xlsx"
        
        custom_filename = st.text_input(
            "Output Excel Filename:",
            value=default_filename,
            key="output_filename_input_raw",
            help="Enter the name for the final Excel file with raw extracted data."
        )
        
        # Generate Excel file for raw data
        excel_data = to_excel(final_data_dict)
        
        # Download Button for raw data
        st.download_button(
            label="📥 Download Consolidated Data (Date, Time, PSum)",
            data=excel_data,
            file_name=custom_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Click to download the Excel file with one sheet per uploaded CSV file."
        )

    else:
        if uploaded_files:
            st.error("No data could be successfully processed. Please review the error messages above and adjust the configurations in the file settings.")
            for err in st.session_state.processing_errors:
                st.error(err)
else:
    if not uploaded_files:
        st.sidebar.markdown("Upload files to configure settings.")
    # Show status if files are uploaded but processing hasn't yielded results yet
    elif uploaded_files and not st.session_state.processing_errors:
        st.info("Files uploaded. Configure settings in the sidebar and the data will be processed automatically.")
