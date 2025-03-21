import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(layout="wide", page_title="Daily Remark Summary", page_icon="📊", initial_sidebar_state="expanded")

# Apply dark mode
st.markdown(
    """
    <style>
    .reportview-container {
        background: #2E2E2E;
        color: white;
    }
    .sidebar .sidebar-content {
        background: #2E2E2E;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.title('Daily Remark Summary')

@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)

    # Check if 'Date' column exists, and if not, print all columns
    if 'Date' not in df.columns:
        st.error("The 'Date' column was not found in the file. Please check the column names.")
        return None

    # Convert 'Date' to datetime if it isn't already
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

    # Exclude rows where 'Debtor' contains 'DEFAULT_LEAD_'
    df = df[~df['Debtor'].str.contains("DEFAULT_LEAD_", case=False, na=False)]

    # Exclude rows where STATUS contains 'BP' (Broken Promise) or 'ABORT'
    df = df[~df['Status'].str.contains('ABORT', na=False)]

    # Exclude rows where STATUS contains 'NEW'
    df = df[~df['Status'].str.contains('NEW', na=False)]

    # Exclude rows where REMARK contains certain keywords or phrases
    excluded_remarks = [
        "Broken Promise",
        "New files imported", 
        "Updates when case reassign to another collector", 
        "NDF IN ICS", 
        "FOR PULL OUT (END OF HANDLING PERIOD)", 
        "END OF HANDLING PERIOD"
    ]
    df = df[~df['Remark'].str.contains('|'.join(excluded_remarks), case=False, na=False)]

    # Exclude rows where "CALL STATUS" contains "OTHERS"
    df = df[~df['Call Status'].str.contains('OTHERS', case=False, na=False)]

    # Exclude rows where the date is a Sunday (weekday() == 6)
    df = df[df['Date'].dt.weekday != 6]  # 6 corresponds to Sunday

    return df

uploaded_file = st.sidebar.file_uploader("Upload Daily Remark File", type="xlsx")

if uploaded_file:
    df = load_data(uploaded_file)
    if df is not None:
        # Initialize an empty DataFrame for the summary table
        summary_columns = ['Day', 'Total Collectors', 'Total Connected', 'Total Talk Time']
        total_summary = pd.DataFrame(columns=summary_columns)

        # Define exclude_users if necessary (or remove if not applicable)
        exclude_users = []  # Add users to exclude if needed (e.g., system users)

        # Function to convert seconds to HH:MM:SS format
        def seconds_to_hms(seconds):
            hours = seconds // 3600
            minutes = (seconds % 3600) // 60
            seconds = seconds % 60
            return f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}"

        # Group by 'Date' and 'Client' (Campaign), instead of 'Collector'
        for date, client, group in df[~df['Remark By'].str.upper().isin(['SYSTEM'])].groupby([df['Date'].dt.date, 'Client']):

            # Calculate the metrics
            total_collectors = group['Remark By'].nunique()  # Count distinct collectors
            total_connected = group[group['Call Status'] == 'CONNECTED']['Account No.'].count()

            # Calculate the total talk time (in seconds), ensure it's numeric
            total_talk_time = pd.to_numeric(group['Talk Time Duration'], errors='coerce').sum()  # Sum of talk time in seconds
            total_talk_time_hms = seconds_to_hms(total_talk_time)  # Convert to HH:MM:SS format

            # Add the row to the summary
            total_summary = pd.concat([total_summary, pd.DataFrame([{
                'Day': date,
                'Total Collectors': total_collectors,
                'Total Connected': total_connected,
                'Total Talk Time': total_talk_time_hms  # Use the HH:MM:SS format here
            }])], ignore_index=True)

        # Add total summary row
        total_summary_row = {
            'Day': 'Total',  # Label for the summary row
            'Total Collectors': total_summary['Total Collectors'].sum(),
            'Total Connected': total_summary['Total Connected'].sum(),
            'Total Talk Time': seconds_to_hms(pd.to_numeric(total_summary['Total Talk Time'].apply(
                lambda x: sum([int(i.split(':')[0])*3600 + int(i.split(':')[1])*60 + int(i.split(':')[2]) for i in x.split() if isinstance(i, str)])
            ).sum()))  # Convert total talk time to HH:MM:SS format
        }

        # Add the total summary to the DataFrame
        total_summary = pd.concat([total_summary, pd.DataFrame([total_summary_row])], ignore_index=True)

        # Show the DataFrame
        st.dataframe(total_summary)
