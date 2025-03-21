import pandas as pd
import streamlit as st

# Set up the page configuration
st.set_page_config(layout="wide", page_title="PRODUCTIVITY", page_icon="📊", initial_sidebar_state="expanded")

# Apply dark mode styling
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

# Title of the app
st.title('Daily Remark Summary')

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

    # Filter out specific users based on 'Remark By'
    exclude_users = ['FGPANGANIBAN', 'KPILUSTRISIMO', 'BLRUIZ', 'MMMEJIA', 'SAHERNANDEZ', 'GPRAMOS',
                     'JGCELIZ', 'SPMADRID', 'RRCARLIT', 'MEBEJER',
                     'SEMIJARES', 'GMCARIAN', 'RRRECTO', 'EASORIANO', 'EUGALERA','JATERRADO','LMLABRADOR']
    df = df[~df['Remark By'].isin(exclude_users)]

    # Create the columns layout
    col1, col2 = st.columns(2)

    with col1:
        st.write("## Summary Table by Day")

        # Add date filter
        min_date = df['Date'].min().date()
        max_date = df['Date'].max().date()
        start_date, end_date = st.date_input("Select date range", [min_date, max_date], min_value=min_date, max_value=max_date)

        filtered_df = df[(df['Date'].dt.date >= start_date) & (df['Date'].dt.date <= end_date)]

        # Initialize an empty DataFrame for the summary table
        summary_table = pd.DataFrame(columns=[ 
            'Day', 'Total Agents', 'Total Connected', 'Talk Time (HH:MM:SS)'
        ])

        # Group by 'Date'
        for date, date_group in filtered_df.groupby(filtered_df['Date'].dt.date):
            # Calculate metrics
            total_agents = date_group['Remark By'].nunique()  # Count unique agents for the day
            total_connected = date_group[date_group['Call Status'] == 'CONNECTED']['Account No.'].count()

            # Calculate total talk time in minutes
            total_talk_time = date_group['Talk Time Duration'].sum() / 60  # Convert from seconds to minutes

            # Round the total talk time to nearest second and convert to HH:MM:SS format
            rounded_talk_time = round(total_talk_time * 60)  # Round to nearest second
            talk_time_str = str(pd.to_timedelta(rounded_talk_time, unit='s'))  # Convert to Timedelta and then to string
            formatted_talk_time = talk_time_str.split()[2]  # Extract the time part from the string (HH:MM:SS)

            # Add the row to the summary table
            summary_table = pd.concat([summary_table, pd.DataFrame([{
                'Day': date,
                'Total Agents': total_agents,
                'Total Connected': total_connected,
                'Talk Time (HH:MM:SS)': formatted_talk_time,  # Add formatted talk time
            }])], ignore_index=True)

        # Calculate and append totals for the summary table
        total_agents = summary_table['Total Agents'].sum()  # Total Agents count across all days
        total_connected = summary_table['Total Connected'].sum()  # Total Connected count across all days

        # Calculate the total talk time for the total row
        total_talk_time_minutes = summary_table['Talk Time (HH:MM:SS)'].apply(
            lambda x: pd.to_timedelta(x).total_seconds() / 60).sum()  # Sum the talk time in minutes

        # Round the total talk time to the nearest second before converting to HH:MM:SS
        rounded_total_talk_time_minutes = round(total_talk_time_minutes)

        # Format the total talk time as HH:MM:SS
        rounded_total_talk_time_seconds = round(rounded_total_talk_time_minutes * 60)  # Round to nearest second
        total_talk_time_str = str(pd.to_timedelta(rounded_total_talk_time_seconds, unit='s')).split()[2]

        total_row = pd.DataFrame([{
            'Day': 'Total',
            'Total Agents': total_agents,
            'Total Connected': total_connected,
            'Talk Time (HH:MM:SS)': total_talk_time_str,  # Add formatted total talk time
        }])

        summary_table = pd.concat([summary_table, total_row], ignore_index=True)

        # Reorder columns to ensure the desired order
        column_order = ['Day', 'Total Agents', 'Total Connected', 'Talk Time (HH:MM:SS)']
        summary_table = summary_table[column_order]

        st.write(summary_table)
