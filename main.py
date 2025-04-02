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
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    # Filter out rows where 'Remark' contains "broken promise" (case-insensitive)
    df = df[~df['Remark'].astype(str).str.contains("broken promise", case=False, na=False)]
    return df

# Function to create a single Excel file with multiple sheets, auto-fit columns, borders, middle alignment, red headers, and custom date formats
def create_combined_excel_file(summary_dfs, overall_summary_df, sheet_prefix, main_header_text):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        main_header_format = workbook.add_format({
            'bg_color': '#000080',  # Navy blue background
            'font_color': '#FFFFFF',  # White text
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 14
        })
        header_format = workbook.add_format({
            'bg_color': '#FF0000',  # Red background
            'font_color': '#FFFFFF',  # White text
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        cell_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        date_format = workbook.add_format({
            'num_format': 'mmm dd, yyyy',  # e.g., Mar 25, 2025
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        date_range_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        time_format = workbook.add_format({
            'num_format': 'hh:mm:ss',  # e.g., 01:23:45
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        # Process each summary sheet
        for key, summary_df in summary_dfs.items():
            sheet_name = f"{sheet_prefix}_{key[:31]}"
            summary_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2, header=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.merge_range('A1:V1', f"{main_header_text} {key}", main_header_format)
            for col_idx, col in enumerate(summary_df.columns):
                worksheet.write(1, col_idx, col, header_format)
            for row_idx in range(len(summary_df)):
                for col_idx, value in enumerate(summary_df.iloc[row_idx]):
                    if col_idx == 0:  # 'Day' or 'Date Range' column
                        worksheet.write(row_idx + 2, col_idx, value, date_format if sheet_prefix == "Summary" else date_range_format)
                    elif col_idx in [10, 12, 13, 14]:  # Talk Time columns
                        worksheet.write(row_idx + 2, col_idx, value, time_format)
                    else:
                        worksheet.write(row_idx + 2, col_idx, value, cell_format)
            for col_idx, col in enumerate(summary_df.columns):
                max_length = max(summary_df[col].astype(str).map(len).max(), len(str(col)))
                worksheet.set_column(col_idx, col_idx, max_length + 2)

        # Process the overall summary sheet
        overall_summary_df.to_excel(writer, sheet_name=f"Overall_{sheet_prefix}", index=False, startrow=2, header=False)
        worksheet = writer.sheets[f"Overall_{sheet_prefix}"]
        worksheet.merge_range('A1:W1', f"Overall {main_header_text}", main_header_format)
        for col_idx, col in enumerate(overall_summary_df.columns):
            worksheet.write(1, col_idx, col, header_format)
        for row_idx in range(len(overall_summary_df)):
            for col_idx, value in enumerate(overall_summary_df.iloc[row_idx]):
                if col_idx == 0:  # 'Date Range' column
                    worksheet.write(row_idx + 2, col_idx, value, date_range_format)
                elif col_idx in [12, 14, 15, 16]:  # Talk Time columns
                    worksheet.write(row_idx + 2, col_idx, value, time_format)
                else:
                    worksheet.write(row_idx + 2, col_idx, value, cell_format)
        for col_idx, col in enumerate(overall_summary_df.columns):
            max_length = max(overall_summary_df[col].astype(str).map(len).max(), len(str(col)))
            worksheet.set_column(col_idx, col_idx, max_length + 2)

    return output.getvalue()

# File uploader for Excel file
uploaded_file = st.sidebar.file_uploader("Upload Daily Remark File", type="xlsx")

# Define columns
col1, col2 = st.columns(2)

if uploaded_file is not None:
    df = load_data(uploaded_file)

    # Ensure 'Time' column is in datetime format
    df['Time'] = pd.to_datetime(df['Time'], errors='coerce').dt.time

    # Ensure 'Date' column is in datetime format
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

    # Ensure 'Talk Time Duration' and 'Call Duration' are numeric
    df['Talk Time Duration'] = pd.to_numeric(df['Talk Time Duration'], errors='coerce').fillna(0)
    df['Call Duration'] = pd.to_numeric(df['Call Duration'], errors='coerce').fillna(0)

    # Define Positive Skip conditions
    positive_skip_keywords = [
        "BRGY SKIPTRACE_POS - LEAVE MESSAGE CALL SMS",
        "BRGY SKIPTRACE_POS - LEAVE MESSAGE FACEBOOK",
        "POS VIA DIGITAL SKIP - OTHER SOCMED PLATFORMS",
        "POSITIVE VIA DIGITAL SKIP - FACEBOOK",
        "POSITIVE VIA DIGITAL SKIP - GOOGLE SEARCH",
        "POSITIVE VIA DIGITAL SKIP - INSTAGRAM",
        "POSITIVE VIA DIGITAL SKIP - LINKEDIN",
        "POSITIVE VIA DIGITAL SKIP - OTHER SOCMED",
        "POSITIVE VIA DIGITAL SKIP - OTHER SOCMED PLATFORMS",
        "POSITIVE VIA DIGITAL SKIP - VIBER",
        "POS VIA SOCMED - GOOGLE SEARCH",
        "POS VIA SOCMED - LINKEDIN",
        "POS VIA SOCMED - OTHER SOCMED PLATFORMS",
        "POS VIA SOCMED - FACEBOOK",
        "POS VIA SOCMED - VIBER",
        "POS VIA SOCMED - INSTAGRAM",
        "POS VIA DIGITAL SKIP - OTHER SOCMED PLATFORMS",
        "LS VIA SOCMED - T5 BROKEN PTP SPLIT AND OTP",
        "LS VIA SOCMED - T6 NO RESPONSE (SMS & EMAIL)",
        "LS VIA SOCMED - T7 PROMO OFFER LETTER",
        "LS VIA SOCMED - T9 RESTRUCTURING",
        "LS VIA SOCMED - T1 NOTIFICATION",
        "LS VIA SOCMED - T12 THIRD PARTY TEMPLATE",
        "LS VIA SOCMED - T8 AMNESTY PROMO TEMPLATE",
        "LS VIA SOCMED - T4 BROKEN PTP EPA",
        "LS VIA SOCMED - T6 NO RESPONSE SMS AND EMAIL",
        "LS VIA SOCMED - OTHERS",
        "LS VIA SOCMED - T10 PRE TERMINATION OFFER",
    ]

    # Define Negative Skip status conditions
    negative_skip_status = [
        "BRGY SKIP TRACING_NEGATIVE - CLIENT UNKNOWN",
        "BRGY SKIP TRACING_NEGATIVE - MOVED OUT",
        "BRGY SKIP TRACING_NEGATIVE - UNCONTACTED",
        "NEG VIA DIGITAL SKIP - OTHER SOCMED PLATFORMS",
        "NEGATIVE VIA DIGITAL SKIP - FACEBOOK",
        "NEGATIVE VIA DIGITAL SKIP - GOOGLE SEARCH",
        "NEGATIVE VIA DIGITAL SKIP - INSTAGRAM",
        "NEGATIVE VIA DIGITAL SKIP - LINKEDIN",
        "NEGATIVE VIA DIGITAL SKIP - OTHER SOCMED",
        "NEGATIVE VIA DIGITAL SKIP - OTHER SOCMED PLATFORMS",
        "NEGATIVE VIA DIGITAL SKIP - VIBER",
        "NEG VIA SOCMED - OTHER SOCMED PLATFORMS",
        "NEG VIA SOCMED - FACEBOOK",
        "NEG VIA SOCMED - VIBER",
        "NEG VIA SOCMED - GOOGLE SEARCH",
        "NEG VIA SOCMED - LINKEDIN",
        "NEG VIA SOCMED - INSTAGRAM",
    ]

    # Define RPC Skip status conditions
    rpc_skip_status = [
        "RPC_POS SKIP WITH REPLY - OTHER SOCMED",
        "RPC_POSITIVE SKIP WITH REPLY - FACEBOOK",
        "RPC_POSITIVE SKIP WITH REPLY - GOOGLE SEARCH",
        "RPC_POSITIVE SKIP WITH REPLY - INSTAGRAM",
        "RPC_POSITIVE SKIP WITH REPLY - LINKEDIN",
        "RPC_POSITIVE SKIP WITH REPLY - OTHER SOCMED PLATFORMS",
        "RPC_POSITIVE SKIP WITH REPLY - VIBER",
        "RPC_REPLY FROM SOCMED - VIBER",
        "RPC_REPLY FROM SOCMED - LINKEDIN",
        "RPC_POS SKIP WITH REPLY - OTHER SOCMED",
        "RPC_POSITIVE SKIP WITH REPLY - FACEBOOK",
        "RPC_POSITIVE SKIP WITH REPLY - VIBER",
        "RPC_REPLY FROM SOCMED - FACEBOOK",
        "RPC_REPLY FROM SOCMED - OTHER SOCMED PLAN",
    ]

    # Dictionaries to store summary DataFrames
    client_summary_dfs = {}
    collector_summary_dfs = {}

    with col1:
        st.write("## Summary Table by Day (Per Client)")
        min_date = df['Date'].min().date()
        max_date = df['Date'].max().date()
        start_date, end_date = st.date_input("Select date range", [min_date, max_date], min_value=min_date, max_value=max_date)
        filtered_df = df[(df['Date'].dt.date >= start_date) & (df['Date'].dt.date <= end_date)]

        for client, client_group in filtered_df.groupby('Client'):
            with st.container():
                st.subheader(f"Client: {client}")
                summary_table = []
                for date, date_group in client_group.groupby(client_group['Date'].dt.date):
                    valid_group = date_group[(date_group['Call Duration'].notna()) & 
                                            (date_group['Call Duration'] > 0) & 
                                            (date_group['Remark By'].str.lower() != "system")]
                    total_agents = valid_group['Remark By'].nunique()
                    collectors = ', '.join(valid_group['Remark By'].unique())  # List of unique collectors for the day
                    total_connected = date_group[date_group['Call Status'] == 'CONNECTED']['Account No.'].count()
                    total_talk_time_seconds = date_group['Talk Time Duration'].sum()
                    hours, remainder = divmod(int(total_talk_time_seconds), 3600)
                    minutes, seconds = divmod(remainder, 60)
                    formatted_talk_time = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                    talk_time_ave_seconds = total_talk_time_seconds / total_agents if total_agents > 0 else 0
                    ave_hours, ave_remainder = divmod(int(talk_time_ave_seconds), 3600)
                    ave_minutes, ave_seconds = divmod(ave_remainder, 60)
                    talk_time_ave_str = f"{ave_hours:02d}:{ave_minutes:02d}:{ave_seconds:02d}"
                    positive_skip_count = sum(date_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))
                    negative_skip_count = date_group[date_group['Status'].isin(negative_skip_status)].shape[0]
                    rpc_skip_count = date_group[date_group['Status'].isin(rpc_skip_status)].shape[0]
                    total_skip = positive_skip_count + negative_skip_count + rpc_skip_count
                    positive_skip_connected = date_group[(date_group['Call Status'] == 'CONNECTED') & 
                                                        (date_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))]['Account No.'].count()
                    negative_skip_connected = date_group[(date_group['Call Status'] == 'CONNECTED') & 
                                                        (date_group['Status'].isin(negative_skip_status))]['Account No.'].count()
                    rpc_skip_connected = date_group[(date_group['Call Status'] == 'CONNECTED') & 
                                                  (date_group['Status'].isin(rpc_skip_status))]['Account No.'].count()
                    positive_skip_talk_time_seconds = date_group[(date_group['Call Status'] == 'CONNECTED') & 
                                                                (date_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))]['Talk Time Duration'].sum()
                    negative_skip_talk_time_seconds = date_group[(date_group['Call Status'] == 'CONNECTED') & 
                                                                (date_group['Status'].isin(negative_skip_status))]['Talk Time Duration'].sum()
                    rpc_skip_talk_time_seconds = date_group[(date_group['Call Status'] == 'CONNECTED') & 
                                                          (date_group['Status'].isin(rpc_skip_status))]['Talk Time Duration'].sum()
                    pos_hours, pos_remainder = divmod(int(positive_skip_talk_time_seconds), 3600)
                    pos_minutes, pos_seconds = divmod(pos_remainder, 60)
                    positive_skip_talk_time = f"{pos_hours:02d}:{pos_minutes:02d}:{pos_seconds:02d}"
                    neg_hours, neg_remainder = divmod(int(negative_skip_talk_time_seconds), 3600)
                    neg_minutes, neg_seconds = divmod(neg_remainder, 60)
                    negative_skip_talk_time = f"{neg_hours:02d}:{neg_minutes:02d}:{neg_seconds:02d}"
                    rpc_hours, rpc_remainder = divmod(int(rpc_skip_talk_time_seconds), 3600)
                    rpc_minutes, rpc_seconds = divmod(rpc_remainder, 60)
                    rpc_skip_talk_time = f"{rpc_hours:02d}:{rpc_minutes:02d}:{rpc_seconds:02d}"
                    positive_skip_ave = round(positive_skip_count / total_agents, 2) if total_agents > 0 else 0
                    negative_skip_ave = round(negative_skip_count / total_agents, 2) if total_agents > 0 else 0
                    rpc_skip_ave = round(rpc_skip_count / total_agents, 2) if total_agents > 0 else 0
                    total_skip_ave = round(total_skip / total_agents, 2) if total_agents > 0 else 0
                    connected_ave = round(total_connected / total_agents, 2) if total_agents > 0 else 0
                    summary_table.append([
                        date, collectors, total_agents, total_connected, positive_skip_count, negative_skip_count, rpc_skip_count, total_skip,
                        positive_skip_connected, negative_skip_connected, rpc_skip_connected, 
                        positive_skip_talk_time, negative_skip_talk_time, rpc_skip_talk_time,
                        formatted_talk_time, positive_skip_ave, negative_skip_ave, rpc_skip_ave, total_skip_ave, connected_ave, talk_time_ave_str
                    ])
                summary_df = pd.DataFrame(summary_table, columns=[
                    'Day', 'Collectors Names', 'Collectors Count', 'Total Connected', 'Positive Skip', 'Negative Skip', 'RPC Skip', 'Total Skip',
                    'Positive Skip Connected', 'Negative Skip Connected', 'RPC Skip Connected', 
                    'Positive Skip Talk Time', 'Negative Skip Talk Time', 'RPC Skip Talk Time',
                    'Talk Time (HH:MM:SS)', 'Positive Skip Ave', 'Negative Skip Ave', 'RPC Skip Ave', 'Total Skip Ave', 'Connected Ave', 'Talk Time Ave'
                ])
                st.dataframe(summary_df)
                client_summary_dfs[client] = summary_df

        st.write("## Summary Table by Day (Per Collector)")
        for collector, collector_group in filtered_df.groupby('Remark By'):
            if collector.lower() == "system":
                continue
            with st.container():
                client = collector_group['Client'].iloc[0]  # Assuming each collector is tied to one client
                st.subheader(f"Agent: {collector} (Client: {client})")
                summary_table = []
                for date, date_group in collector_group.groupby(collector_group['Date'].dt.date):
                    valid_group = date_group[(date_group['Call Duration'].notna()) & 
                                            (date_group['Call Duration'] > 0)]
                    total_connected = date_group[date_group['Call Status'] == 'CONNECTED']['Account No.'].count()
                    total_talk_time_seconds = date_group['Talk Time Duration'].sum()
                    hours, remainder = divmod(int(total_talk_time_seconds), 3600)
                    minutes, seconds = divmod(remainder, 60)
                    formatted_talk_time = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                    positive_skip_count = sum(date_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))
                    negative_skip_count = date_group[date_group['Status'].isin(negative_skip_status)].shape[0]
                    rpc_skip_count = date_group[date_group['Status'].isin(rpc_skip_status)].shape[0]
                    total_skip = positive_skip_count + negative_skip_count + rpc_skip_count
                    positive_skip_connected = date_group[(date_group['Call Status'] == 'CONNECTED') & 
                                                        (date_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))]['Account No.'].count()
                    negative_skip_connected = date_group[(date_group['Call Status'] == 'CONNECTED') & 
                                                        (date_group['Status'].isin(negative_skip_status))]['Account No.'].count()
                    rpc_skip_connected = date_group[(date_group['Call Status'] == 'CONNECTED') & 
                                                  (date_group['Status'].isin(rpc_skip_status))]['Account No.'].count()
                    positive_skip_talk_time_seconds = date_group[(date_group['Call Status'] == 'CONNECTED') & 
                                                                (date_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))]['Talk Time Duration'].sum()
                    negative_skip_talk_time_seconds = date_group[(date_group['Call Status'] == 'CONNECTED') & 
                                                                (date_group['Status'].isin(negative_skip_status))]['Talk Time Duration'].sum()
                    rpc_skip_talk_time_seconds = date_group[(date_group['Call Status'] == 'CONNECTED') & 
                                                          (date_group['Status'].isin(rpc_skip_status))]['Talk Time Duration'].sum()
                    pos_hours, pos_remainder = divmod(int(positive_skip_talk_time_seconds), 3600)
                    pos_minutes, pos_seconds = divmod(pos_remainder, 60)
                    positive_skip_talk_time = f"{pos_hours:02d}:{pos_minutes:02d}:{pos_seconds:02d}"
                    neg_hours, neg_remainder = divmod(int(negative_skip_talk_time_seconds), 3600)
                    neg_minutes, neg_seconds = divmod(neg_remainder, 60)
                    negative_skip_talk_time = f"{neg_hours:02d}:{neg_minutes:02d}:{neg_seconds:02d}"
                    rpc_hours, rpc_remainder = divmod(int(rpc_skip_talk_time_seconds), 3600)
                    rpc_minutes, rpc_seconds = divmod(rpc_remainder, 60)
                    rpc_skip_talk_time = f"{rpc_hours:02d}:{rpc_minutes:02d}:{rpc_seconds:02d}"
                    manual_calls = date_group[date_group['Remark Type'] == 'outgoing'].shape[0]
                    manual_accounts = date_group[date_group['Remark Type'] == 'outgoing']['Account No.'].nunique()
                    summary_table.append([
                        date, client, total_connected, positive_skip_count, negative_skip_count, rpc_skip_count, total_skip,
                        positive_skip_connected, negative_skip_connected, rpc_skip_connected, 
                        positive_skip_talk_time, negative_skip_talk_time, rpc_skip_talk_time,
                        formatted_talk_time, manual_calls, manual_accounts
                    ])
                summary_df = pd.DataFrame(summary_table, columns=[
                    'Day', 'Client', 'Total Connected', 'Positive Skip', 'Negative Skip', 'RPC Skip', 'Total Skip',
                    'Positive Skip Connected', 'Negative Skip Connected', 'RPC Skip Connected', 
                    'Positive Skip Talk Time', 'Negative Skip Talk Time', 'RPC Skip Talk Time',
                    'Talk Time (HH:MM:SS)', 'Manual Calls', 'Manual Accounts'
                ])
                st.dataframe(summary_df)
                collector_summary_dfs[collector] = summary_df

    with col2:
        st.write("## Overall Summary per Client")
        with st.container():
            date_range_str = f"{start_date.strftime('%b %d %Y').upper()} - {end_date.strftime('%b %d %Y').upper()}"
            valid_df = filtered_df[(filtered_df['Call Duration'].notna()) & 
                                  (filtered_df['Call Duration'] > 0) & 
                                  (filtered_df['Remark By'].str.lower() != "system")]
            avg_collectors_per_client = valid_df.groupby(['Client', valid_df['Date'].dt.date])['Remark By'].nunique().groupby('Client').mean().apply(lambda x: math.ceil(x) if x % 1 >= 0.5 else round(x))

            overall_client_summary = []
            for client, client_group in filtered_df.groupby('Client'):
                total_agents = avg_collectors_per_client.get(client, 0)
                total_connected = client_group[client_group['Call Status'] == 'CONNECTED']['Account No.'].count()
                total_talk_time_seconds = client_group['Talk Time Duration'].sum()
                hours, remainder = divmod(int(total_talk_time_seconds), 3600)
                minutes, seconds = divmod(remainder, 60)
                formatted_talk_time = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                positive_skip_count = sum(client_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))
                negative_skip_count = client_group[client_group['Status'].isin(negative_skip_status)].shape[0]
                rpc_skip_count = client_group[client_group['Status'].isin(rpc_skip_status)].shape[0]
                total_skip = positive_skip_count + negative_skip_count + rpc_skip_count
                positive_skip_connected = client_group[(client_group['Call Status'] == 'CONNECTED') & 
                                                      (client_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))]['Account No.'].count()
                negative_skip_connected = client_group[(client_group['Call Status'] == 'CONNECTED') & 
                                                      (client_group['Status'].isin(negative_skip_status))]['Account No.'].count()
                rpc_skip_connected = client_group[(client_group['Call Status'] == 'CONNECTED') & 
                                                (client_group['Status'].isin(rpc_skip_status))]['Account No.'].count()
                positive_skip_talk_time_seconds = client_group[(client_group['Call Status'] == 'CONNECTED') & 
                                                              (client_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))]['Talk Time Duration'].sum()
                negative_skip_talk_time_seconds = client_group[(client_group['Call Status'] == 'CONNECTED') & 
                                                              (client_group['Status'].isin(negative_skip_status))]['Talk Time Duration'].sum()
                rpc_skip_talk_time_seconds = client_group[(client_group['Call Status'] == 'CONNECTED') & 
                                                        (client_group['Status'].isin(rpc_skip_status))]['Talk Time Duration'].sum()
                pos_hours, pos_remainder = divmod(int(positive_skip_talk_time_seconds), 3600)
                pos_minutes, pos_seconds = divmod(pos_remainder, 60)
                positive_skip_talk_time = f"{pos_hours:02d}:{pos_minutes:02d}:{pos_seconds:02d}"
                neg_hours, neg_remainder = divmod(int(negative_skip_talk_time_seconds), 3600)
                neg_minutes, neg_seconds = divmod(neg_remainder, 60)
                negative_skip_talk_time = f"{neg_hours:02d}:{neg_minutes:02d}:{neg_seconds:02d}"
                rpc_hours, rpc_remainder = divmod(int(rpc_skip_talk_time_seconds), 3600)
                rpc_minutes, rpc_seconds = divmod(rpc_remainder, 60)
                rpc_skip_talk_time = f"{rpc_hours:02d}:{rpc_minutes:02d}:{rpc_seconds:02d}"
                daily_data = client_group.groupby(client_group['Date'].dt.date).agg({
                    'Remark By': lambda x: x[(client_group['Call Duration'].notna()) & 
                                            (client_group['Call Duration'] > 0) & 
                                            (client_group['Remark By'].str.lower() != "system")].nunique(),
                    'Account No.': lambda x: x[client_group['Call Status'] == 'CONNECTED'].count(),
                    'Status': [
                        lambda x: sum(x.astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False)),
                        lambda x: x.isin(negative_skip_status).sum(),
                        lambda x: x.isin(rpc_skip_status).sum(),
                        lambda x: x[(client_group['Call Status'] == 'CONNECTED') & 
                                   (x.astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))].count(),
                        lambda x: x[(client_group['Call Status'] == 'CONNECTED') & 
                                   (x.isin(negative_skip_status))].count(),
                        lambda x: x[(client_group['Call Status'] == 'CONNECTED') & 
                                   (x.isin(rpc_skip_status))].count()
                    ],
                    'Talk Time Duration': [
                        'sum',
                        lambda x: x[(client_group['Call Status'] == 'CONNECTED') & 
                                   (client_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))].sum(),
                        lambda x: x[(client_group['Call Status'] == 'CONNECTED') & 
                                   (client_group['Status'].isin(negative_skip_status))].sum(),
                        lambda x: x[(client_group['Call Status'] == 'CONNECTED') & 
                                   (client_group['Status'].isin(rpc_skip_status))].sum()
                    ]
                })
                daily_data.columns = ['Collectors', 'Total Connected', 
                                     'Positive Skip', 'Negative Skip', 'RPC Skip',
                                     'Positive Skip Connected', 'Negative Skip Connected', 'RPC Skip Connected',
                                     'Talk Time', 'Positive Skip Talk Time Seconds', 'Negative Skip Talk Time Seconds', 'RPC Skip Talk Time Seconds']
                daily_data['Total Skip'] = daily_data['Positive Skip'] + daily_data['Negative Skip'] + daily_data['RPC Skip']
                daily_data['Positive Skip Ave'] = daily_data['Positive Skip'] / daily_data['Collectors']
                daily_data['Negative Skip Ave'] = daily_data['Negative Skip'] / daily_data['Collectors']
                daily_data['RPC Skip Ave'] = daily_data['RPC Skip'] / daily_data['Collectors']
                daily_data['Total Skip Ave'] = daily_data['Total Skip'] / daily_data['Collectors']
                daily_data['Connected Ave'] = daily_data['Total Connected'] / daily_data['Collectors']
                daily_data['Talk Time Ave Seconds'] = daily_data['Talk Time'] / daily_data['Collectors']
                positive_skip_ave = round(daily_data['Positive Skip Ave'].mean(), 2) if not daily_data.empty else 0
                negative_skip_ave = round(daily_data['Negative Skip Ave'].mean(), 2) if not daily_data.empty else 0
                rpc_skip_ave = round(daily_data['RPC Skip Ave'].mean(), 2) if not daily_data.empty else 0
                total_skip_ave = round(daily_data['Total Skip Ave'].mean(), 2) if not daily_data.empty else 0
                connected_ave = round(daily_data['Connected Ave'].mean(), 2) if not daily_data.empty else 0
                talk_time_ave_seconds = daily_data['Talk Time Ave Seconds'].mean() if not daily_data.empty else 0
                ave_hours, ave_remainder = divmod(int(talk_time_ave_seconds), 3600)
                ave_minutes, ave_seconds = divmod(ave_remainder, 60)
                talk_time_ave_str = f"{ave_hours:02d}:{ave_minutes:02d}:{ave_seconds:02d}"
                overall_client_summary.append([
                    date_range_str, client, total_agents, total_connected, positive_skip_count, negative_skip_count, rpc_skip_count, total_skip,
                    positive_skip_connected, negative_skip_connected, rpc_skip_connected,
                    positive_skip_talk_time, negative_skip_talk_time, rpc_skip_talk_time,
                    positive_skip_ave, negative_skip_ave, rpc_skip_ave, total_skip_ave, formatted_talk_time, connected_ave, talk_time_ave_str
                ])
            overall_client_summary_df = pd.DataFrame(overall_client_summary, columns=[
                'Date Range', 'Client', 'Collectors', 'Total Connected', 'Positive Skip', 'Negative Skip', 'RPC Skip', 'Total Skip',
                'Positive Skip Connected', 'Negative Skip Connected', 'RPC Skip Connected', 
                'Positive Skip Talk Time', 'Negative Skip Talk Time', 'RPC Skip Talk Time',
                'Positive Skip Ave', 'Negative Skip Ave', 'RPC Skip Ave', 'Total Skip Ave', 'Talk Time (HH:MM:SS)', 'Connected Ave', 'Talk Time Ave'
            ])
            st.dataframe(overall_client_summary_df)

            # Generate Excel file for per-client data
            client_excel_data = create_combined_excel_file(client_summary_dfs, overall_client_summary_df, "Summary", "Daily Summary for")
            st.download_button(
                label="Download Per Client Results",
                data=client_excel_data,
                file_name="MC06_Monitoring_Per_Client_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.write("## Overall Summary per Collector")
        with st.container():
            overall_collector_summary = []
            for collector, collector_group in filtered_df.groupby('Remark By'):
                if collector.lower() == "system":
                    continue
                client = collector_group['Client'].iloc[0]  # Assuming each collector is tied to one client
                total_connected = collector_group[collector_group['Call Status'] == 'CONNECTED']['Account No.'].count()
                total_talk_time_seconds = collector_group['Talk Time Duration'].sum()
                hours, remainder = divmod(int(total_talk_time_seconds), 3600)
                minutes, seconds = divmod(remainder, 60)
                formatted_talk_time = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                positive_skip_count = sum(collector_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))
                negative_skip_count = collector_group[collector_group['Status'].isin(negative_skip_status)].shape[0]
                rpc_skip_count = collector_group[collector_group['Status'].isin(rpc_skip_status)].shape[0]
                total_skip = positive_skip_count + negative_skip_count + rpc_skip_count
                positive_skip_connected = collector_group[(collector_group['Call Status'] == 'CONNECTED') & 
                                                        (collector_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))]['Account No.'].count()
                negative_skip_connected = collector_group[(collector_group['Call Status'] == 'CONNECTED') & 
                                                        (collector_group['Status'].isin(negative_skip_status))]['Account No.'].count()
                rpc_skip_connected = collector_group[(collector_group['Call Status'] == 'CONNECTED') & 
                                                  (collector_group['Status'].isin(rpc_skip_status))]['Account No.'].count()
                positive_skip_talk_time_seconds = collector_group[(collector_group['Call Status'] == 'CONNECTED') & 
                                                                (collector_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))]['Talk Time Duration'].sum()
                negative_skip_talk_time_seconds = collector_group[(collector_group['Call Status'] == 'CONNECTED') & 
                                                                (collector_group['Status'].isin(negative_skip_status))]['Talk Time Duration'].sum()
                rpc_skip_talk_time_seconds = collector_group[(collector_group['Call Status'] == 'CONNECTED') & 
                                                          (collector_group['Status'].isin(rpc_skip_status))]['Talk Time Duration'].sum()
                pos_hours, pos_remainder = divmod(int(positive_skip_talk_time_seconds), 3600)
                pos_minutes, pos_seconds = divmod(pos_remainder, 60)
                positive_skip_talk_time = f"{pos_hours:02d}:{pos_minutes:02d}:{pos_seconds:02d}"
                neg_hours, neg_remainder = divmod(int(negative_skip_talk_time_seconds), 3600)
                neg_minutes, neg_seconds = divmod(neg_remainder, 60)
                negative_skip_talk_time = f"{neg_hours:02d}:{neg_minutes:02d}:{neg_seconds:02d}"
                rpc_hours, rpc_remainder = divmod(int(rpc_skip_talk_time_seconds), 3600)
                rpc_minutes, rpc_seconds = divmod(rpc_remainder, 60)
                rpc_skip_talk_time = f"{rpc_hours:02d}:{rpc_minutes:02d}:{rpc_seconds:02d}"
                manual_calls = collector_group[collector_group['Remark Type'] == 'outgoing'].shape[0]
                manual_accounts = collector_group[collector_group['Remark Type'] == 'outgoing']['Account No.'].nunique()
                overall_collector_summary.append([
                    date_range_str, collector, client, total_connected, positive_skip_count, negative_skip_count, rpc_skip_count, total_skip,
                    positive_skip_connected, negative_skip_connected, rpc_skip_connected,
                    positive_skip_talk_time, negative_skip_talk_time, rpc_skip_talk_time,
                    formatted_talk_time, manual_calls, manual_accounts
                ])
            overall_collector_summary_df = pd.DataFrame(overall_collector_summary, columns=[
                'Date Range', 'Collector', 'Client', 'Total Connected', 'Positive Skip', 'Negative Skip', 'RPC Skip', 'Total Skip',
                'Positive Skip Connected', 'Negative Skip Connected', 'RPC Skip Connected', 
                'Positive Skip Talk Time', 'Negative Skip Talk Time', 'RPC Skip Talk Time',
                'Talk Time (HH:MM:SS)', 'Manual Calls', 'Manual Accounts'
            ])
            st.dataframe(overall_collector_summary_df)

            # Generate Excel file for per-collector data
            collector_excel_data = create_combined_excel_file(collector_summary_dfs, overall_collector_summary_df, "Collector_Summary", "Daily Summary for Collector")
            st.download_button(
                label="Download Per Collector Results",
                data=collector_excel_data,
                file_name="MC06_Monitoring_Per_Collector_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
