import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, numbers
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.utils import get_column_letter

# --- Configuration ---
POWER_COL_OUT = 'PSumW'

# -----------------------------
# PROCESS SINGLE SHEET
# -----------------------------
def process_sheet(df, date_col, time_col, psum_col):
    """
    Processes a single DataFrame sheet: cleans data, rounds timestamps to 10-minute intervals,
    filters out leading/trailing zero periods, and prepares data for Excel output.
    
    This version combines separate Date and Time columns into a single timestamp index,
    resamples to 10-minute intervals (summing power), and removes inactive periods
    before the first non-zero reading and after the last non-zero reading.
    """
    df.columns = df.columns.astype(str).str.strip()
    
    # 1. Combine Date and Time columns into a single timestamp string/series
    combined_dt_series = df[date_col].astype(str) + ' ' + df[time_col].astype(str)
    
    # Convert the combined string to datetime objects
    df['Timestamp'] = pd.to_datetime(combined_dt_series, errors="coerce", dayfirst=True)
    timestamp_col = 'Timestamp'
    
    # 2. Clean and convert power column (handle commas as decimal separators)
    power_series = df[psum_col].astype(str).str.strip()
    # Replace comma decimal separator with dot, then convert to numeric
    power_series = power_series.str.replace(',', '.', regex=False)
    df[psum_col] = pd.to_numeric(power_series, errors='coerce')
    
    # 3. Drop rows where we couldn't parse the timestamp or power value
    df = df.dropna(subset=[timestamp_col, psum_col])
    if df.empty:
        return pd.DataFrame()
    
    # --- CORE LOGIC: FILTER LEADING AND TRAILING ZEROS (Inactive Period) ---
    
    # Identify indices where the absolute power reading is non-zero
    non_zero_indices = df[df[psum_col].abs() != 0].index
    
    if non_zero_indices.empty:
        # If all valid readings are zero, return an empty DataFrame (no usable period)
        return pd.DataFrame() 
        
    # Get the index of the first and last non-zero reading
    first_valid_idx = non_zero_indices.min()
    last_valid_idx = non_zero_indices.max()
    
    # Slice the DataFrame to keep data only between the first and last active reading.
    df = df.loc[first_valid_idx:last_valid_idx].copy()

    # ----------------------------------------------------------------------
    
    # 4. Resample data to 10-minute intervals
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
    # Use ceil('D') to include the last day's readings up to 23:50
    max_dt_exclusive = df_out['Rounded'].max().ceil('D') 
    if min_dt >= max_dt_exclusive:
        return pd.DataFrame()
    
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
    
    # Ensure the column is float type to correctly hold NaN values
    grouped[POWER_COL_OUT] = grouped[POWER_COL_OUT].astype(float) 

    grouped["Date"] = grouped["Rounded"].dt.date
    grouped["Time"] = grouped["Rounded"].dt.strftime("%H:%M") 
    
    # Filter back to only the dates originally present in the file
    grouped = grouped[grouped["Date"].isin(original_dates)]
    
    # Add kW column (absolute value). Since NaN * 1000 = NaN, this works fine.
    grouped['kW'] = grouped[POWER_COL_OUT].abs() / 1000
    
    return grouped

# -----------------------------
# BUILD EXCEL
# -----------------------------
def build_output_excel(sheets_dict):
    """Creates the final formatted Excel file with data, charts, and summary."""
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    header_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
    title_font = Font(bold=True, size=12)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))

    # Data structure for the "Total" sheet
    # Format: { date_obj: { sheet_name: max_kw, ... }, ... }
    total_sheet_data = {}
    sheet_names_list = []

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(sheet_name)
        sheet_names_list.append(sheet_name)
        dates = sorted(df["Date"].unique())
        col_start = 1
        max_row_used = 0
        daily_max_summary = []

        day_intervals = []
        
        # Structure:
        # Row 1: Merged Date Header (Full Date)
        # Row 2: Sub-headers (Time, W, kW)
        # Row 3: Start of data (Time, W, kW)

        for date in dates:
            # Get all data for the day (including NaNs for missing periods)
            day_data_full = df[df["Date"] == date].sort_values("Time")
            
            # Data used for calculations (excluding the new NaNs from outside the active period)
            day_data_active = day_data_full.dropna(subset=[POWER_COL_OUT])
            
            n_rows = len(day_data_full) # Use full count for row structure
            day_intervals.append(n_rows)
            
            data_start_row = 3 # Data starts at Row 3 
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
                ws.cell(row=idx, column=col_start+2, value=getattr(r, POWER_COL_OUT)) 
                ws.cell(row=idx, column=col_start+3, value=r.kW)

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

            # Collect data for "Total" sheet
            if date not in total_sheet_data:
                total_sheet_data[date] = {}
            total_sheet_data[date][sheet_name] = max_kw

            col_start += 4

        # Add Line Chart for Individual Sheet
        if dates:
            chart = LineChart()
            chart.title = f"Daily 10-Minute Absolute Power Profile - {sheet_name}"
            chart.y_axis.title = "kW"
            chart.x_axis.title = "Time"
            
            # Set Chart Size
            chart.height = 12.5 
            chart.width = 23    

            max_rows = max(day_intervals) if day_intervals else 0
            if max_rows > 0:
                first_time_col = 2
                # Categories ref: starts at row 3, ends at 2+max_rows
                categories_ref = Reference(ws, min_col=first_time_col, min_row=3, max_row=2 + max_rows)

                col_start = 1
                for i, n_rows in enumerate(day_intervals):
                    # Data ref: starts at row 3, ends at 2+n_rows
                    data_ref = Reference(ws, min_col=col_start+3, min_row=3, max_col=col_start+3, max_row=2+n_rows)
                    
                    # Get series name from dates list using index
                    date_title_str = dates[i].strftime('%d-%b')
                    
                    # Use Series object directly
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
            # Date Column
            date_cell = ws_total.cell(row=row_idx, column=1, value=date_obj.strftime('%Y-%m-%d'))
            date_cell.border = thin_border
            date_cell.alignment = Alignment(horizontal="center")
            
            row_total_load = 0
            
            # Sheet Columns
            for col_idx, sheet_name in enumerate(sheet_names_list, 2):
                val = total_sheet_data[date_obj].get(sheet_name, 0)
                if pd.isna(val): val = 0
                
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
            # Set the title to "Overview" as requested
            chart_total.title = "Overview" 
            chart_total.y_axis.title = "Max Power (kW)"
            chart_total.x_axis.title = "Date"
            
            # Set Chart Size
            chart_total.height = 15
            chart_total.width = 30
            
            data_max_row = len(sorted_dates) + 1
            total_cols = len(sheet_names_list) + 2
            
            # Chart Data Reference: Cover all value columns (Sheet 1...Sheet N + Total Load)
            data_ref = Reference(ws_total, min_col=2, min_row=1, max_col=total_cols, max_row=data_max_row)
            chart_total.add_data(data_ref, titles_from_data=True)

            # Iterate through all generated series and set smooth=False for straight lines
            for s in chart_total.series:
                s.smooth = False
            
            # Category Axis: Date Column (Col 1)
            cats_ref = Reference(ws_total, min_col=1, min_row=2, max_row=data_max_row)
            chart_total.set_categories(cats_ref)
            
            ws_total.add_chart(chart_total, "B" + str(data_max_row + 3))

    stream = BytesIO()
    if 'Sheet' in wb.sheetnames and len(wb.sheetnames) > len(sheets_dict) + (1 if total_sheet_data else 0):
        wb.remove(wb['Sheet'])
        
    wb.save(stream)
    stream.seek(0)
    return stream

# -----------------------------
# STREAMLIT APP
# -----------------------------
def app():
    st.set_page_config(layout="wide", page_title="Electricity Data Converter")
    st.title("📊 Excel 10-Minute Electricity Data Converter")
    st.markdown("""
        Upload an **Excel file (.xlsx)** with time-series data (e.g., the output from the consolidation tool, Code 1). Each sheet is processed to calculate total absolute power (W) in 10-minute intervals. 
        
        **Input Format Expected:** Separate columns for **Date**, **Time**, and **PSum (W)**.
        
        **Processing Logic:**
        1. Resamples data to 10-minute intervals, summing the Power (W).
        2. Filters out inactive periods: Leading and trailing zero values (data outside the first and last non-zero reading) are removed and appear blank in the output.
        
        The output Excel file includes:
        - **Individual Sheet Analysis:** A **line chart** and a **Max Power Summary table** for each day.
        - **Total Summary Sheet:** A comparative table and graph of daily max power across all sheets, **including the Total Load series**.
    """)

    uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"])

    if uploaded:
        xls = pd.ExcelFile(uploaded)
        result_sheets = {}
        st.write("---")

        for sheet_name in xls.sheet_names:
            st.markdown(f"**Processing sheet: `{sheet_name}`**")
            try:
                # Read the sheet, assuming the date/time/psum columns are already cleaned by Code 1
                df = pd.read_excel(uploaded, sheet_name=sheet_name)
            except Exception as e:
                st.error(f"Error reading sheet '{sheet_name}': {e}")
                continue

            df.columns = df.columns.astype(str).str.strip()
            
            # --- COLUMN DETECTION (Case-insensitive check for expected columns) ---
            date_col = next((c for c in df.columns if c in ["Date","DATE","date"]), None)
            time_col = next((c for c in df.columns if c in ["Time","TIME","time"]), None)
            
            if not date_col or not time_col:
                st.error(f"No valid Date and/or Time column in sheet '{sheet_name}' (expected: Date, Time).")
                continue

            psum_col = next((c for c in df.columns if c in ["PSum (W)","Psum (W)","PSum","P (W)","Power","PSUMW"]), None)
            if not psum_col:
                st.error(f"No valid PSum column in sheet '{sheet_name}' (expected: PSum (W) or similar).")
                continue

            # --- PROCESS DATA ---
            processed = process_sheet(df, date_col, time_col, psum_col)
            
            if not processed.empty:
                result_sheets[sheet_name] = processed
                st.success(f"Sheet '{sheet_name}' processed successfully. Found {len(processed)} 10-minute intervals.")
                # Show a small preview
                st.dataframe(processed.head(), use_container_width=True)
            else:
                st.warning(f"Sheet '{sheet_name}' resulted in no data after processing (check for all zero or invalid readings within the sheet's active period).")

        st.write("---")

        # --- Download Section ---
        if result_sheets:
            st.subheader("✅ All Sheets Processed")

            # Default filename generation
            default_filename = uploaded.name.replace(".xlsx", "_10min_Analysis.xlsx").replace(".csv", "_10min_Analysis.xlsx")
            
            # Allow user to customize output filename
            custom_filename = st.text_input(
                "Output Excel Filename:",
                value=default_filename,
                key="output_filename_analysis",
                help="Enter the name for the final Excel file with 10-minute interval data, charts, and summaries."
            )

            with st.spinner("Building the final formatted Excel file... This may take a moment for large files..."):
                try:
                    excel_data_output = build_output_excel(result_sheets)
                except Exception as e:
                    st.error(f"An error occurred while building the final Excel file: {e}")
                    return
            
            # Download Button
            st.download_button(
                label="📥 Download 10-Minute Analysis Excel",
                data=excel_data_output,
                file_name=custom_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Click to download the Excel file containing 10-minute summarized data, charts, and total summaries."
            )
            st.balloons()
        else:
            st.error("No sheets could be successfully processed. Please review the sheet processing warnings above.")

if __name__ == '__main__':
    # Ensure the app runs correctly when executed
    app()
