import pandas as pd
import streamlit as st

# Set up the page configuration
st.set_page_config(layout="wide", page_title="MC06 MONITORING", page_icon="📊", initial_sidebar_state="expanded")

# Title of the app
st.title('MC06 MONITORING')

# Data loading function with file upload support
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    return df

# File uploader for Excel file
uploaded_file = st.sidebar.file_uploader("Upload Daily Remark File", type="xlsx")

if uploaded_file is not None:
    df = load_data(uploaded_file)

    # Ensure 'Time' column is in datetime format
    df['Time'] = pd.to_datetime(df['Time'], errors='coerce').dt.time

    # Ensure 'Date' column is in datetime format
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

    # Ensure 'Talk Time Duration' is numeric
    df['Talk Time Duration'] = pd.to_numeric(df['Talk Time Duration'], errors='coerce').fillna(0)

    # Define Positive Skip conditions (count if Status CONTAINS these keywords)
    positive_skip_keywords = [
        "BRGY SKIPTRACE_POS - LEAVE MESSAGE FACEBOOK",
        "POS VIA DIGITAL SKIP - OTHER SOCMED PLATFORMS",
        "POSITIVE VIA DIGITAL SKIP - FACEBOOK",
        "POSITIVE VIA DIGITAL SKIP - VIBER",
        "RPC_POS SKIP WITH REPLY - OTHER SOCMED",
        "RPC_POSITIVE SKIP WITH REPLY - FACEBOOK",
        "RPC_POSITIVE SKIP WITH REPLY - VIBER"
    ]

    # Define Negative Skip status conditions
    negative_skip_status = [
        "NEGATIVE VIA DIGITAL SKIP - FACEBOOK",
        "NEGATIVE VIA DIGITAL SKIP - VIBER",
        "NEG VIA DIGITAL SKIP - OTHER SOCMED PLATFORMS",
        "BRGY SKIP TRACING_NEGATIVE - CLIENT UNKNOWN",
        "BRGY SKIP TRACING_NEGATIVE - MOVED OUT"
    ]

    # Create Streamlit columns layout
    col1, col2 = st.columns(2)

    with col1:
        st.write("## Summary Table by Day")

        # Add date filter
        min_date = df['Date'].min().date()
        max_date = df['Date'].max().date()
        start_date, end_date = st.date_input("Select date range", [min_date, max_date], min_value=min_date, max_value=max_date)

        # Filter data based on date range
        filtered_df = df[(df['Date'].dt.date >= start_date) & (df['Date'].dt.date <= end_date)]

        # Initialize an empty summary table
        summary_table = []

        # Group by 'Date' and 'Client'
        for (date, client), date_group in filtered_df.groupby([filtered_df['Date'].dt.date, 'Client']):
            total_agents = date_group['Remark By'].nunique()
            total_connected = date_group[date_group['Call Status'] == 'CONNECTED']['Account No.'].count()

            # Sum Talk Time Duration in seconds
            total_talk_time_seconds = date_group['Talk Time Duration'].sum()

            # Convert total talk time to HH:MM:SS format
            formatted_talk_time = str(pd.to_timedelta(total_talk_time_seconds, unit='s'))

            # Calculate average talk time per agent
            talk_time_ave_seconds = total_talk_time_seconds / total_agents if total_agents > 0 else 0

            # Convert average talk time to HH:MM:SS format
            talk_time_ave_str = str(pd.to_timedelta(round(talk_time_ave_seconds), unit='s'))

            # Count Positive Skip
            positive_skip_count = sum(date_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))

            # Count Negative Skip
            negative_skip_count = date_group[date_group['Status'].isin(negative_skip_status)].shape[0]

            # Calculate Total Skip
            total_skip = positive_skip_count + negative_skip_count

            # Calculate connected average per agent
            connected_ave = round(total_connected / total_agents, 2) if total_agents > 0 else 0

            # Append results to the summary table
            summary_table.append([
                date, client, total_agents, total_connected, positive_skip_count, negative_skip_count, total_skip, formatted_talk_time, connected_ave, talk_time_ave_str
            ])

        # Convert to DataFrame and display
        summary_df = pd.DataFrame(summary_table, columns=[
            'Day', 'Client', 'Total Agents', 'Total Connected', 'Positive Skip', 'Negative Skip', 'Total Skip', 'Talk Time (HH:MM:SS)', 'Connected Ave', 'Talk Time Ave'
        ])
        st.write(summary_df)

    with col2:
        st.write("## Overall Summary per Date")

        overall_summary = []

        for date, date_group in filtered_df.groupby(filtered_df['Date'].dt.date):
            total_agents = date_group['Remark By'].nunique()
            total_connected = date_group[date_group['Call Status'] == 'CONNECTED']['Account No.'].count()

            # Sum Talk Time Duration in seconds
            total_talk_time_seconds = date_group['Talk Time Duration'].sum()

            # Convert total talk time to HH:MM:SS format
            formatted_talk_time = str(pd.to_timedelta(total_talk_time_seconds, unit='s'))

            # Calculate average talk time per agent
            talk_time_ave_seconds = total_talk_time_seconds / total_agents if total_agents > 0 else 0

            # Convert average talk time to HH:MM:SS format
            talk_time_ave_str = str(pd.to_timedelta(round(talk_time_ave_seconds), unit='s'))

            # Count Positive Skip
            positive_skip_count = sum(date_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))

            # Count Negative Skip
            negative_skip_count = date_group[date_group['Status'].isin(negative_skip_status)].shape[0]

            # Calculate Total Skip
            total_skip = positive_skip_count + negative_skip_count

            # Calculate connected average per agent
            connected_ave = round(total_connected / total_agents, 2) if total_agents > 0 else 0

            # Append results to the overall summary
            overall_summary.append([
                date, total_agents, total_connected, positive_skip_count, negative_skip_count, total_skip, formatted_talk_time, connected_ave, talk_time_ave_str
            ])

        # Convert to DataFrame and display
        overall_summary_df = pd.DataFrame(overall_summary, columns=[
            'Day', 'Total Agents', 'Total Connected', 'Positive Skip', 'Negative Skip', 'Total Skip', 'Talk Time (HH:MM:SS)', 'Connected Ave', 'Talk Time Ave'
        ])
        st.write(overall_summary_df)
