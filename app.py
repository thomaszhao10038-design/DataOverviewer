import streamlit as st
import pandas as pd
from io import BytesIO

# --- Configuration ---
# Assuming input sheets have a 'aDate' column and a 'Power (kW)' column.
DATE_COLUMN = 'Date'
POWER_COLUMN = 'Power (kW)'
REQUIRED_COLUMNS = [DATE_COLUMN, POWER_COLUMN]

st.set_page_config(
    page_title="Multi-Sheet Load Aggregator",
    layout="wide",
    initial_sidebar_state="expanded"
)

def process_excel_data(uploaded_file):
    """Reads all sheets, calculates daily max power for each, and generates the 'Total' summary."""
    try:
        # Read ALL sheets from the Excel file into a dictionary of DataFrames
        # sheet_name=None returns a dict {sheet_name: DataFrame}
        all_sheets_data = pd.read_excel(uploaded_file, sheet_name=None)

    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return None

    summary_data = {}

    for sheet_name, df in all_sheets_data.items():
        # Check for required columns
        if not all(col in df.columns for col in REQUIRED_COLUMNS):
            st.warning(f"Skipping sheet '{sheet_name}': Missing required columns ({', '.join(REQUIRED_COLUMNS)}).")
            continue

        # Ensure 'Date' column is in datetime format and set as index
        try:
            df[DATE_COLUMN] = pd.to_datetime(df[DATE_COLUMN])
        except Exception as e:
            st.error(f"Error converting '{DATE_COLUMN}' in sheet '{sheet_name}' to datetime: {e}")
            continue

        # Create a day-only key for grouping
        df['Date_Only'] = df[DATE_COLUMN].dt.date

        # Calculate the daily maximum power for the current sheet
        # Group by the date and find the max of the power column
        daily_max_power = df.groupby('Date_Only')[POWER_COLUMN].max().reset_index()

        # Rename the max power column to the sheet name for the summary table
        daily_max_power.rename(columns={'Date_Only': 'date', POWER_COLUMN: sheet_name}, inplace=True)

        # Store the result
        summary_data[sheet_name] = daily_max_power.set_index('date')


    if not summary_data:
        st.error("No valid sheets found to process. Please ensure sheets contain 'Date' and 'Power (kW)' columns.")
        return None

    # --- Step 3: Combine all daily max power data ---
    
    # Start with the first sheet's data
    first_sheet_name = list(summary_data.keys())[0]
    total_df = summary_data[first_sheet_name].copy()
    
    # Merge the rest of the sheets one by one
    for i, (sheet_name, df_data) in enumerate(summary_data.items()):
        if i == 0:
            continue # Skip the first one as it's the base
        
        # Merge on the index (which is 'date')
        total_df = pd.merge(total_df, df_data, left_index=True, right_index=True, how='outer')

    # --- Step 4: Calculate the Total Load ---
    
    # Select only the sheet columns (by excluding the index)
    sheet_columns = [col for col in total_df.columns if col != 'date']
    
    # Calculate the 'total load' by summing across the row (axis=1)
    total_df['total load'] = total_df[sheet_columns].sum(axis=1)

    # Reset the index to make 'date' a column again, as requested by the output structure
    total_df.reset_index(inplace=True)
    total_df.rename(columns={'index': 'date'}, inplace=True)
    
    # Sort by date for clean presentation
    total_df.sort_values(by='date', inplace=True)

    return total_df


# --- Streamlit UI ---
st.title("💡 Daily Max Load Aggregator")
st.markdown("""
Upload an Excel file with multiple sheets, where each sheet contains time-series load data (with 'Date' and 'Power (kW)' columns). 
The app will generate a summary sheet named 'Total' showing the daily maximum power for each sheet and the aggregated total load.
""")

uploaded_file = st.file_uploader(
    "Upload your Multi-Sheet Excel File (.xlsx or .xls)",
    type=['xlsx', 'xls']
)

if uploaded_file is not None:
    st.subheader("Processing Data...")

    # Process the file
    result_df = process_excel_data(uploaded_file)

    if result_df is not None:
        st.success("✅ Data processing complete! The 'Total' summary sheet is ready.")
        st.subheader("Generated 'Total' Summary Table")
        
        # Display the result
        st.dataframe(result_df, use_container_width=True)

        # --- Download Button ---
        @st.cache_data
        def convert_df_to_excel(df):
            """Converts the final DataFrame to an Excel file in memory."""
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Total', index=False)
            return output.getvalue()

        # Generate the Excel data
        excel_data = convert_df_to_excel(result_df)
        
        st.download_button(
            label="⬇️ Download 'Total' Sheet as Excel",
            data=excel_data,
            file_name="Total_Max_Load_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Download the resulting summary table as a new Excel file named 'Total'."
        )

        st.markdown("---")
        st.subheader("Raw Data Display (First 50 Rows of the Result)")
        st.table(result_df.head(50))
