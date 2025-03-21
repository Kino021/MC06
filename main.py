import pandas as pd
import streamlit as st

# Set up the page configuration
st.set_page_config(layout="wide", page_title="MC06 MONITORING", page_icon="📊", initial_sidebar_state="expanded")

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
            'Day', 'Client', 'Total Agents', 'Total Connected', 'Talk Time (HH:MM:SS)', 'Connected Ave', 'Talk Time Ave'
        ])

        # Group by 'Date' and 'Client'
        for (date, client), date_group in filtered_df.groupby([filtered_df['Date'].dt.date, 'Client']):
            # Filter rows where 'Call Duration' has a value (non-zero)
            valid_agents = date_group[date_group['Call Duration'].notna() & (date_group['Call Duration'] > 0)]

            # Calculate metrics
            total_agents = valid_agents['Remark By'].nunique()  # Count unique agents for the day where Call Duration > 0
            total_connected = date_group[date_group['Call Status'] == 'CONNECTED']['Account No.'].count()

            # Calculate total talk time in minutes
            total_talk_time = date_group['Talk Time Duration'].sum() / 60  # Convert from seconds to minutes

            # Round the total talk time to nearest second and convert to HH:MM:SS format
            rounded_talk_time = round(total_talk_time * 60)  # Round to nearest second
            talk_time_str = str(pd.to_timedelta(rounded_talk_time, unit='s'))  # Convert to Timedelta and then to string
            formatted_talk_time = talk_time_str.split()[2]  # Extract the time part from the string (HH:MM:SS)

            # Calculate "Connected Ave" and "Talk Time Ave"
            connected_ave = total_connected / total_agents if total_agents > 0 else 0
            talk_time_ave = total_talk_time / total_agents if total_agents > 0 else 0

            # Convert Talk Time Ave to HH:MM:SS format
            rounded_talk_time_ave = round(talk_time_ave * 60)  # Round to nearest second
            talk_time_ave_str = str(pd.to_timedelta(rounded_talk_time_ave, unit='s')).split()[2]

            # Add the row to the summary table
            summary_table = pd.concat([summary_table, pd.DataFrame([{
                'Day': date,
                'Client': client,
                'Total Agents': total_agents,
                'Total Connected': total_connected,
                'Talk Time (HH:MM:SS)': formatted_talk_time,  # Add formatted talk time
                'Connected Ave': round(connected_ave, 2),  # Round to 2 decimal places
                'Talk Time Ave': talk_time_ave_str  # Add formatted talk time average
            }])], ignore_index=True)

        # Calculate and append averages for the summary table
        # Avoid including the "Talk Time Ave" in the mean calculation (it should be treated as string)
        summary_table['Talk Time (Seconds)'] = summary_table['Talk Time (HH:MM:SS)'].apply(
            lambda x: pd.to_timedelta(x).total_seconds() / 60)  # Convert to minutes for the calculation

        # Calculate average for "Total Connected" and "Talk Time (in minutes)"
        total_connected_ave = summary_table['Total Connected'].mean()  # Average of Total Connected
        total_talk_time_ave_minutes = summary_table['Talk Time (Seconds)'].mean()  # Average talk time in minutes

        # Format the total talk time average as HH:MM:SS
        rounded_total_talk_time_ave_minutes = round(total_talk_time_ave_minutes)
        rounded_total_talk_time_ave_seconds = round(rounded_total_talk_time_ave_minutes * 60)  # Round to nearest second
        total_talk_time_ave_str = str(pd.to_timedelta(rounded_total_talk_time_ave_seconds, unit='s')).split()[2]

        # Averages for "Connected Ave"
        connected_ave_total = summary_table['Connected Ave'].mean()  # Calculate average of "Connected Ave"

        # Calculate average for "Talk Time Ave" (per-agent average across all days)
        talk_time_ave_total = summary_table['Talk Time Ave'].apply(
            lambda x: pd.to_timedelta(x).total_seconds() / 60).mean()  # Average of Talk Time Ave in minutes

        # Convert "Talk Time Ave" for total row back to HH:MM:SS format
        rounded_talk_time_ave_total = round(talk_time_ave_total * 60)
        total_talk_time_ave_str = str(pd.to_timedelta(rounded_talk_time_ave_total, unit='s')).split()[2]

        # Create a total row with averages
        total_row = pd.DataFrame([{
            'Day': 'Total',
            'Client': '',  # Leave Client blank for the total row
            'Total Agents': '',  # Leave Total Agents blank
            'Total Connected': round(total_connected_ave, 2),  # Use average for Total Connected
            'Talk Time (HH:MM:SS)': str(pd.to_timedelta(rounded_total_talk_time_ave_seconds, unit='s')).split()[2],  # Add average talk time
            'Connected Ave': round(connected_ave_total, 2),  # Use average for Connected Ave
            'Talk Time Ave': total_talk_time_ave_str  # Add average Talk Time Ave per agent
        }])

        # Add the total row to the summary table
        summary_table = pd.concat([summary_table, total_row], ignore_index=True)

        # Reorder columns to ensure the desired order
        column_order = ['Day', 'Client', 'Total Agents', 'Total Connected', 'Talk Time (HH:MM:SS)', 'Connected Ave', 'Talk Time Ave']
        summary_table = summary_table[column_order]

        # Display the updated summary table
        st.write(summary_table)

    # Now let's create the overall summary table per date
    with col2:
        st.write("## Overall Summary per Date")

        # Create a summary table grouped by 'Date'
        overall_summary = pd.DataFrame(columns=[ 
            'Day', 'Total Agents', 'Total Connected', 'Talk Time (HH:MM:SS)', 'Connected Ave', 'Talk Time Ave'
        ])

        # Group by 'Date'
        for date, date_group in filtered_df.groupby(filtered_df['Date'].dt.date):
            # Filter rows where 'Call Duration' has a value (non-zero)
            valid_agents = date_group[date_group['Call Duration'].notna() & (date_group['Call Duration'] > 0)]

            # Calculate metrics
            total_agents = valid_agents['Remark By'].nunique()  # Count unique agents for the day where Call Duration > 0
            total_connected = date_group[date_group['Call Status'] == 'CONNECTED']['Account No.'].count()

            # Calculate total talk time in minutes
            total_talk_time = date_group['Talk Time Duration'].sum() / 60  # Convert from seconds to minutes

            # Round the total talk time to nearest second and convert to HH:MM:SS format
            rounded_talk_time = round(total_talk_time * 60)  # Round to nearest second
            talk_time_str = str(pd.to_timedelta(rounded_talk_time, unit='s'))  # Convert to Timedelta and then to string
            formatted_talk_time = talk_time_str.split()[2]  # Extract the time part from the string (HH:MM:SS)

            # Calculate "Connected Ave" and "Talk Time Ave"
            connected_ave = total_connected / total_agents if total_agents > 0 else 0
            talk_time_ave = total_talk_time / total_agents if total_agents > 0 else 0

            # Convert Talk Time Ave to HH:MM:SS format
            rounded_talk_time_ave = round(talk_time_ave * 60)  # Round to nearest second
            talk_time_ave_str = str(pd.to_timedelta(rounded_talk_time_ave, unit='s')).split()[2]

            # Add the row to the overall summary table
            overall_summary = pd.concat([overall_summary, pd.DataFrame([{
                'Day': date,
                'Total Agents': total_agents,
                'Total Connected': total_connected,
                'Talk Time (HH:MM:SS)': formatted_talk_time,  # Add formatted talk time
                'Connected Ave': round(connected_ave, 2),  # Round to 2 decimal places
                'Talk Time Ave': talk_time_ave_str  # Add formatted talk time average
            }])], ignore_index=True)

        # Display the overall summary table per date
        st.write(overall_summary)
