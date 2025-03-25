import pandas as pd
import streamlit as st
import math

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
        "RPC_POS SKIP WITH REPLY - OTHER SOCMED",
        "RPC_POSITIVE SKIP WITH REPLY - FACEBOOK",
        "RPC_POSITIVE SKIP WITH REPLY - GOOGLE SEARCH",
        "RPC_POSITIVE SKIP WITH REPLY - INSTAGRAM",
        "RPC_POSITIVE SKIP WITH REPLY - LINKEDIN",
        "RPC_POSITIVE SKIP WITH REPLY - OTHER SOCMED PLATFORMS",
        "RPC_POSITIVE SKIP WITH REPLY - VIBER",
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

        # Group by 'Client' first, then by 'Date' within each client
        for client, client_group in filtered_df.groupby('Client'):
            with st.container():
                st.subheader(f"Client: {client}")
                
                # Initialize an empty summary table for this client
                summary_table = []

                # Group by 'Date' within this client's data
                for date, date_group in client_group.groupby(client_group['Date'].dt.date):
                    total_agents = date_group['Remark By'].nunique()
                    total_connected = date_group[date_group['Call Status'] == 'CONNECTED']['Account No.'].count()

                    # Sum Talk Time Duration in seconds
                    total_talk_time_seconds = date_group['Talk Time Duration'].sum()

                    # Convert total talk time to HH:MM:SS format without days
                    hours, remainder = divmod(int(total_talk_time_seconds), 3600)
                    minutes, seconds = divmod(remainder, 60)
                    formatted_talk_time = f"{hours:02d}:{minutes:02d}:{seconds:02d}"

                    # Calculate average talk time per agent
                    talk_time_ave_seconds = total_talk_time_seconds / total_agents if total_agents > 0 else 0

                    # Convert average talk time to HH:MM:SS format without days
                    ave_hours, ave_remainder = divmod(int(talk_time_ave_seconds), 3600)
                    ave_minutes, ave_seconds = divmod(ave_remainder, 60)
                    talk_time_ave_str = f"{ave_hours:02d}:{ave_minutes:02d}:{ave_seconds:02d}"

                    # Count Positive Skip
                    positive_skip_count = sum(date_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))

                    # Count Negative Skip
                    negative_skip_count = date_group[date_group['Status'].isin(negative_skip_status)].shape[0]

                    # Calculate Total Skip
                    total_skip = positive_skip_count + negative_skip_count

                    # Calculate Positive Skip Average (Positive Skip / Collectors)
                    positive_skip_ave = round(positive_skip_count / total_agents, 2) if total_agents > 0 else 0

                    # Calculate Negative Skip Average (Negative Skip / Collectors)
                    negative_skip_ave = round(negative_skip_count / total_agents, 2) if total_agents > 0 else 0

                    # Calculate Total Skip Average (Total Skip / Collectors)
                    total_skip_ave = round(total_skip / total_agents, 2) if total_agents > 0 else 0

                    # Calculate connected average per agent
                    connected_ave = round(total_connected / total_agents, 2) if total_agents > 0 else 0

                    # Append results to the summary table for this client with Talk Time after Total Skip
                    summary_table.append([
                        date, total_agents, total_connected, positive_skip_count, negative_skip_count, total_skip,
                        formatted_talk_time, positive_skip_ave, negative_skip_ave, total_skip_ave, connected_ave, talk_time_ave_str
                    ])

                # Convert to DataFrame and display for this client with Talk Time after Total Skip
                summary_df = pd.DataFrame(summary_table, columns=[
                    'Day', 'Collectors', 'Total Connected', 'Positive Skip', 'Negative Skip', 'Total Skip',
                    'Talk Time (HH:MM:SS)', 'Positive Skip Ave', 'Negative Skip Ave', 'Total Skip Ave', 'Connected Ave', 'Talk Time Ave'
                ])
                st.write(summary_df)

    with col2:
        st.write("## Overall Summary per Client")

        # Format the date range string
        date_range_str = f"{start_date.strftime('%b %d %Y').upper()} - {end_date.strftime('%b %d %Y').upper()}"

        # Calculate average collectors per client from filtered_df and round up if .5 or greater
        avg_collectors_per_client = filtered_df.groupby(['Client', filtered_df['Date'].dt.date])['Remark By'].nunique().groupby('Client').mean().apply(lambda x: math.ceil(x) if x % 1 >= 0.5 else round(x))

        # Group by 'Client' and create separate containers
        for client, client_group in filtered_df.groupby('Client'):
            with st.container():
                st.subheader(f"Client: {client}")
                
                overall_summary = []

                # Use the rounded average collectors
                total_agents = avg_collectors_per_client[client]
                total_connected = client_group[client_group['Call Status'] == 'CONNECTED']['Account No.'].count()

                # Sum Talk Time Duration in seconds
                total_talk_time_seconds = client_group['Talk Time Duration'].sum()

                # Convert total talk time to HH:MM:SS format without days
                hours, remainder = divmod(int(total_talk_time_seconds), 3600)
                minutes, seconds = divmod(remainder, 60)
                formatted_talk_time = f"{hours:02d}:{minutes:02d}:{seconds:02d}"

                # Calculate average talk time per agent using the rounded average collectors
                talk_time_ave_seconds = total_talk_time_seconds / total_agents if total_agents > 0 else 0

                # Convert average talk time to HH:MM:SS format without days
                ave_hours, ave_remainder = divmod(int(talk_time_ave_seconds), 3600)
                ave_minutes, ave_seconds = divmod(ave_remainder, 60)
                talk_time_ave_str = f"{ave_hours:02d}:{ave_minutes:02d}:{ave_seconds:02d}"

                # Count Positive Skip
                positive_skip_count = sum(client_group['Status'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))

                # Count Negative Skip
                negative_skip_count = client_group[client_group['Status'].isin(negative_skip_status)].shape[0]

                # Calculate Total Skip
                total_skip = positive_skip_count + negative_skip_count

                # Calculate Positive Skip Average (Positive Skip / Collectors)
                positive_skip_ave = round(positive_skip_count / total_agents, 2) if total_agents > 0 else 0

                # Calculate Negative Skip Average (Negative Skip / Collectors)
                negative_skip_ave = round(negative_skip_count / total_agents, 2) if total_agents > 0 else 0

                # Calculate Total Skip Average (Total Skip / Collectors)
                total_skip_ave = round(total_skip / total_agents, 2) if total_agents > 0 else 0

                # Calculate connected average per agent using the rounded average collectors
                connected_ave = round(total_connected / total_agents, 2) if total_agents > 0 else 0

                # Append results to the overall summary with new averages after Total Skip and before Talk Time
                overall_summary.append([
                    date_range_str, total_agents, total_connected, positive_skip_count, negative_skip_count, total_skip,
                    positive_skip_ave, negative_skip_ave, total_skip_ave, formatted_talk_time, connected_ave, talk_time_ave_str
                ])

                # Convert to DataFrame and display for this client with new averages
                overall_summary_df = pd.DataFrame(overall_summary, columns=[
                    'Date Range', 'Collectors', 'Total Connected', 'Positive Skip', 'Negative Skip', 'Total Skip',
                    'Positive Skip Ave', 'Negative Skip Ave', 'Total Skip Ave', 'Talk Time (HH:MM:SS)', 'Connected Ave', 'Talk Time Ave'
                ])
                st.write(overall_summary_df)
