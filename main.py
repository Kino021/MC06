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

def summary_table(df):
    # Group by Date
    summary_df = df.groupby('Date').agg(
        total_agents=('Remark By', lambda x: x.nunique()),  # Count unique agents in the 'Remark By' column
        total_talk_time=('Talk Time Duration', 'sum'),
        total_connected_calls=('Call Status', lambda x: x.str.contains("CONNECTED", case=False, na=False).sum())
    ).reset_index()

    # Calculate additional columns
    summary_df['Talktime AVE'] = summary_df['total_talk_time'] / summary_df['total_agents']
    summary_df['Connected AVE'] = summary_df['total_connected_calls'] / summary_df['total_agents']

    # Handle cases where total_agents might be zero to avoid division by zero errors
    summary_df['Talktime AVE'].replace([float('inf'), -float('inf')], 0, inplace=True)
    summary_df['Connected AVE'].replace([float('inf'), -float('inf')], 0, inplace=True)

    return summary_df

# Upload file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    df = load_data(uploaded_file)

    if df is not None:
        # Generate the summary table
        summary_df = summary_table(df)

        # Display the summary table
        st.subheader("Summary Table")
        st.write(summary_df)
