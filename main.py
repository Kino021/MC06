import pandas as pd
import streamlit as st
import math
from io import BytesIO

# Set up the page configuration
st.set_page_config(layout="wide", page_title="MC06 MONITORING", page_icon="📊", initial_sidebar_state="expanded")

# Title of the app
st.title('MC06 MONITORING')

# Data loading function with file upload support
@st.cache_data
def load_data(uploaded_file, remark_column='Remark'):
    df = pd.read_excel(uploaded_file)
    # Check if the specified remark column exists
    if remark_column not in df.columns:
        st.error(f"Column '{remark_column}' not found in the uploaded file. Available columns: {list(df.columns)}")
        return None
    # Filter out rows where the remark column contains "broken promise" (case-insensitive)
    df = df[~df[remark_column].astype(str).str.contains("broken promise", case=False, na=False)]
    # Explicitly parse the 'Date' column with the expected format
    try:
        df['Date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y', errors='coerce')
    except ValueError:
        # If the format doesn't match, try a different common format or let pandas infer
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    if df['Date'].isna().all():
        st.error("Could not parse the 'Date' column. Please ensure the dates are in a recognizable format (e.g., DD-MM-YYYY).")
        return None
    return df

# Function to create a single Excel file with all summaries (Overall Summary first, no collector summaries)
def create_combined_excel_file(summary_dfs, overall_summary_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        main_header_format = workbook.add_format({
            'bg_color': '#000080', 'font_color': '#FFFFFF', 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 14
        })
        header_format = workbook.add_format({
            'bg_color': '#FF0000', 'font_color': '#FFFFFF', 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        cell_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        date_format = workbook.add_format({'num_format': 'mmm dd, yyyy', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        date_range_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        time_format = workbook.add_format({'num_format': 'hh:mm:ss', 'border': 1, 'align': 'center', 'valign': 'vcenter'})

        # Process the overall summary sheet (per client) first
        overall_summary_df.to_excel(writer, sheet_name="Overall_Summary", index=False, startrow=2, header=False)
        worksheet = writer.sheets["Overall_Summary"]
        worksheet.merge_range('A1:U1', "Overall Summary per Client", main_header_format)
        for col_idx, col in enumerate(overall_summary_df.columns):
            worksheet.write(1, col_idx, col, header_format)
        for row_idx in range(len(overall_summary_df)):
            for col_idx, value in enumerate(overall_summary_df.iloc[row_idx]):
                if col_idx == 0:  # 'Date Range' column
                    worksheet.write(row_idx + 2, col_idx, value, date_range_format)
                elif col_idx in [11, 12, 13, 14]:  # Talk Time columns
                    worksheet.write(row_idx + 2, col_idx, value, time_format)
                else:
                    worksheet.write(row_idx + 2, col_idx, value, cell_format)
        for col_idx, col in enumerate(overall_summary_df.columns):
            if col_idx == 0:
                max_length = max(overall_summary_df[col].astype(str).map(len).max(), len(str(col)))
            else:
                max_length = max(overall_summary_df[col].astype(str).map(len).max(), len(str(col)))
            worksheet.set_column(col_idx, col_idx, max_length + 2)

        # Process each client's summary sheet
        for client, summary_df in summary_dfs.items():
            summary_df.to_excel(writer, sheet_name=f"Summary_{client[:31]}", index=False, startrow=2, header=False)
            worksheet = writer.sheets[f"Summary_{client[:31]}"]
            worksheet.merge_range('A1:T1', f"Daily Summary for {client}", main_header_format)
            for col_idx, col in enumerate(summary_df.columns):
                worksheet.write(1, col_idx, col, header_format)
            for row_idx in range(len(summary_df)):
                for col_idx, value in enumerate(summary_df.iloc[row_idx]):
                    if col_idx == 0:  # 'Day' column
                        worksheet.write_datetime(row_idx + 2, col_idx, value, date_format)
                    elif col_idx in [10, 11, 12, 13]:  # Talk Time columns
                        worksheet.write(row_idx + 2, col_idx, value, time_format)
                    else:
                        worksheet.write(row_idx + 2, col_idx, value, cell_format)
            for col_idx, col in enumerate(summary_df.columns):
                if col_idx == 0:
                    max_length = max(summary_df[col].astype(str).map(lambda x: len('MMM DD, YYYY')).max(), len(str(col)))
                else:
                    max_length = max(summary_df[col].astype(str).map(len).max(), len(str(col)))
                worksheet.set_column(col_idx, col_idx, max_length + 2)

    return output.getvalue()

# Function to create Excel file for client summaries only (Overall Summary first)
def create_client_summary_excel(summary_dfs, overall_summary_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        main_header_format = workbook.add_format({
            'bg_color': '#000080', 'font_color': '#FFFFFF', 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 14
        })
        header_format = workbook.add_format({
            'bg_color': '#FF0000', 'font_color': '#FFFFFF', 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        cell_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        date_format = workbook.add_format({'num_format': 'mmm dd, yyyy', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        date_range_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        time_format = workbook.add_format({'num_format': 'hh:mm:ss', 'border': 1, 'align': 'center', 'valign': 'vcenter'})

        # Process the overall summary sheet (per client) first
        overall_summary_df.to_excel(writer, sheet_name
