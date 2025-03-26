import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from pandas import ExcelWriter

st.set_page_config(layout="wide", page_title="Daily Remark Summary", page_icon="📊", initial_sidebar_state="expanded")

st.title('Daily Remark Summary')

@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip().str.upper()
    df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
    df = df[df['DATE'].dt.weekday != 6]  # Exclude Sundays
    return df

uploaded_file = st.sidebar.file_uploader("Upload Daily Remark File", type="xlsx")

def to_excel(df):
    output = BytesIO()
    with ExcelWriter(output, engine='xlsxwriter', date_format='yyyy-mm-dd') as writer:
        workbook = writer.book
        
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#FFFF00',
        })
        center_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        header_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': 'red',
            'font_color': 'white',
            'bold': True
        })
        comma_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'num_format': '#,##0'
        })
        percent_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'num_format': '0.00%'
        })
        date_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'num_format': 'yyyy-mm-dd'
        })
        time_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'num_format': 'hh:mm:ss'
        })
        
        # Convert percentage strings back to floats for Excel
        df_for_excel = df.copy()
        for col in ['PENETRATION RATE (%)', 'CONNECTED RATE (%)', 'PTP RATE', 'CALL DROP RATIO #']:
            df_for_excel[col] = df_for_excel[col].str.rstrip('%').astype(float)
        
        df_for_excel.to_excel(writer, sheet_name="Summary", index=False, startrow=1)
        worksheet = writer.sheets["Summary"]
        
        worksheet.merge_range('A1:' + chr(65 + len(df.columns) - 1) + '1', "Daily Remark Summary", title_format)
        
        for col_num, col_name in enumerate(df_for_excel.columns):
            worksheet.write(1, col_num, col_name, header_format)
        
        for row_num in range(2, len(df_for_excel) + 2):
            for col_num, col_name in enumerate(df_for_excel.columns):
                value = df_for_excel.iloc[row_num - 2, col_num]
                if col_name == 'DATE':
                    if isinstance(value, (pd.Timestamp, datetime.date)):
                        worksheet.write_datetime(row_num, col_num, value, date_format)
                    else:
                        worksheet.write(row_num, col_num, value, date_format)
                elif col_name in ['TOTAL PTP AMOUNT', 'TOTAL BALANCE']:
                    worksheet.write(row_num, col_num, value, comma_format)
                elif col_name in ['PENETRATION RATE (%)', 'CONNECTED RATE (%)', 'PTP RATE', 'CALL DROP RATIO #']:
                    worksheet.write(row_num, col_num, value / 100, percent_format)
                elif col_name in ['TOTAL TALK TIME', 'TALK TIME AVE', 'POSITIVE SKIP TALK TIME', 'NEGATIVE SKIP TALK TIME']:
                    worksheet.write(row_num, col_num, value, time_format)
                else:
                    worksheet.write(row_num, col_num, value, center_format)
        
        for col_num, col_name in enumerate(df_for_excel.columns):
            max_len = max(
                df_for_excel[col_name].astype(str).str.len().max(),
                len(col_name)
            ) + 2
            worksheet.set_column(col_num, col_num, max_len)

    return output.getvalue()

if uploaded_file is not None:
    df = load_data(uploaded_file)
    df = df[~df['DEBTOR'].str.contains("DEFAULT_LEAD_", case=False, na=False)]
    df = df[~df['STATUS'].str.contains('ABORT', na=False)]
    
    excluded_remarks = [
        "Broken Promise", "New files imported", "Updates when case reassign to another collector", 
        "NDF IN ICS", "FOR PULL OUT (END OF HANDLING PERIOD)", "END OF HANDLING PERIOD", "New Assignment -",
    ]
    df = df[~df['REMARK'].str.contains('|'.join(excluded_remarks), case=False, na=False)]
    df = df[~df['CALL STATUS'].str.contains('OTHERS', case=False, na=False)]
    
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
        "RPC_POS SKIP WITH REPLY - OTHER SOCMED",
        "RPC_POSITIVE SKIP WITH REPLY - FACEBOOK",
        "RPC_POSITIVE SKIP WITH REPLY - GOOGLE SEARCH",
        "RPC_POSITIVE SKIP WITH REPLY - INSTAGRAM",
        "RPC_POSITIVE SKIP WITH REPLY - LINKEDIN",
        "RPC_POSITIVE SKIP WITH REPLY - OTHER SOCMED PLATFORMS",
        "RPC_POSITIVE SKIP WITH REPLY - VIBER",
        "POS VIA SOCMED - GOOGLE SEARCH",
        "POS VIA SOCMED - LINKEDIN",
        "POS VIA SOCMED - OTHER SOCMED PLATFORMS",
        "POS VIA SOCMED - FACEBOOK",
        "POS VIA SOCMED - VIBER",
        "RPC_REPLY FROM SOCMED - VIBER",
        "POS VIA SOCMED - INSTAGRAM",
        "RPC_REPLY FROM SOCMED - LINKEDIN",
        "RPC_POS SKIP WITH REPLY - OTHER SOCMED",
        "RPC_POSITIVE SKIP WITH REPLY - FACEBOOK",
        "POS VIA DIGITAL SKIP - OTHER SOCMED PLATFORMS",
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
        "NEG VIA SOCMED - OTHER SOCMED PLATFORMS",
        "LS VIA SOCMED - T5 BROKEN PTP SPLIT AND OTP",
        "LS VIA SOCMED - T6 NO RESPONSE (SMS & EMAIL)",
        "RPC_REPLY FROM SOCMED - FACEBOOK",
        "RPC_REPLY FROM SOCMED - OTHER SOCMED PLAN",
        "LS VIA SOCMED - T7 PROMO OFFER LETTER",
        "LS VIA SOCMED - T9 RESTRUCTURING",
        "LS VIA SOCMED - T1 NOTIFICATION",
        "NEG VIA SOCMED - FACEBOOK",
        "LS VIA SOCMED - OTHERS",
        "NEG VIA SOCMED - VIBER",
        "NEG VIA SOCMED - GOOGLE SEARCH",
        "LS VIA SOCMED - T10 PRE TERMINATION OFFER",
        "NEG VIA SOCMED - LINKEDIN",
        "LS VIA SOCMED - T12 THIRD PARTY TEMPLATE",
        "LS VIA SOCMED - T8 AMNESTY PROMO TEMPLATE",
        "LS VIA SOCMED - T4 BROKEN PTP EPA",
        "LETTER RECEIVED - THRU OTHER SOCMED",
        "LS VIA SOCMED - T6 NO RESPONSE SMS AND EMAIL",
        "NEG VIA SOCMED - INSTAGRAM",
    ]

    def format_seconds_to_hms(seconds):
        seconds = int(seconds)
        hours = seconds // 3600
        minutes = (seconds % 3600) // 60
        secs = seconds % 60
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"

    def calculate_summary(df):
        summary_columns = [
            'DATE', 'CLIENT', 'COLLECTORS', 'ACCOUNTS', 'TOTAL DIALED', 'PENETRATION RATE (%)', 'CONNECTED #', 
            'CONNECTED RATE (%)', 'CONNECTED ACC', 'TOTAL TALK TIME', 'TALK TIME AVE', 'PTP ACC', 'PTP RATE', 
            'TOTAL PTP AMOUNT', 'TOTAL BALANCE', 'CALL DROP #', 'SYSTEM DROP', 'CALL DROP RATIO #',
            'POSITIVE SKIP CONNECTED', 'NEGATIVE SKIP CONNECTED', 'POSITIVE SKIP TALK TIME', 'NEGATIVE SKIP TALK TIME'
        ]
        
        summary_table = pd.DataFrame(columns=summary_columns)
        
        df['DATE'] = df['DATE'].dt.date  

        for (date, client), group in df.groupby(['DATE', 'CLIENT']):
            accounts = group['ACCOUNT NO.'].nunique()
            total_dialed = group['ACCOUNT NO.'].count()
            connected = group[group['CALL STATUS'] == 'CONNECTED']['ACCOUNT NO.'].nunique()
            penetration_rate = (total_dialed / accounts * 100) if accounts != 0 else 0
            penetration_rate_formatted = f"{penetration_rate:.2f}%"
            connected_acc = group[group['CALL STATUS'] == 'CONNECTED']['ACCOUNT NO.'].count()
            connected_rate = (connected_acc / total_dialed * 100) if total_dialed != 0 else 0
            connected_rate_formatted = f"{connected_rate:.2f}%"
            ptp_acc = group[(group['STATUS'].str.contains('PTP', na=False)) & (group['PTP AMOUNT'] != 0)]['ACCOUNT NO.'].nunique()
            ptp_rate = (ptp_acc / connected * 100) if connected != 0 else 0
            ptp_rate_formatted = f"{ptp_rate:.2f}%"
            total_ptp_amount = group[(group['STATUS'].str.contains('PTP', na=False)) & (group['PTP AMOUNT'] != 0)]['PTP AMOUNT'].sum()
            total_balance = group[(group['PTP AMOUNT'] != 0)]['BALANCE'].sum()
            system_drop = group[(group['STATUS'].str.contains('DROPPED', na=False)) & (group['REMARK BY'] == 'SYSTEM')]['ACCOUNT NO.'].count()
            call_drop_count = group[(group['STATUS'].str.contains('NEGATIVE CALLOUTS - DROP CALL', na=False)) & 
                                  (~group['REMARK BY'].str.upper().isin(['SYSTEM']))]['ACCOUNT NO.'].count()
            call_drop_ratio = (call_drop_count / connected_acc * 100) if connected_acc != 0 else 0
            call_drop_ratio_formatted = f"{call_drop_ratio:.2f}%"

            collectors = group[group['CALL DURATION'].notna()]['REMARK BY'].nunique()
            total_talk_seconds = group['TALK TIME DURATION'].sum()
            total_talk_time = format_seconds_to_hms(total_talk_seconds)
            talk_time_ave = format_seconds_to_hms(total_talk_seconds / collectors) if collectors != 0 else "00:00:00"

            # New columns: Positive and Negative Skip metrics
            positive_skip_connected = group[(group['CALL STATUS'] == 'CONNECTED') & 
                                           (group['STATUS'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))]['ACCOUNT NO.'].count()
            negative_skip_connected = group[(group['CALL STATUS'] == 'CONNECTED') & 
                                           (group['STATUS'].isin(negative_skip_status))]['ACCOUNT NO.'].count()
            positive_skip_talk_time_seconds = group[(group['CALL STATUS'] == 'CONNECTED') & 
                                                   (group['STATUS'].astype(str).str.contains('|'.join(positive_skip_keywords), case=False, na=False))]['TALK TIME DURATION'].sum()
            negative_skip_talk_time_seconds = group[(group['CALL STATUS'] == 'CONNECTED') & 
                                                   (group['STATUS'].isin(negative_skip_status))]['TALK TIME DURATION'].sum()
            positive_skip_talk_time = format_seconds_to_hms(positive_skip_talk_time_seconds)
            negative_skip_talk_time = format_seconds_to_hms(negative_skip_talk_time_seconds)

            summary_data = {
                'DATE': date,
                'CLIENT': client,
                'COLLECTORS': collectors,
                'ACCOUNTS': accounts,
                'TOTAL DIALED': total_dialed,
                'PENETRATION RATE (%)': penetration_rate_formatted,
                'CONNECTED #': connected,
                'CONNECTED RATE (%)': connected_rate_formatted,
                'CONNECTED ACC': connected_acc,
                'TOTAL TALK TIME': total_talk_time,
                'TALK TIME AVE': talk_time_ave,
                'PTP ACC': ptp_acc,
                'PTP RATE': ptp_rate_formatted,
                'TOTAL PTP AMOUNT': total_ptp_amount,
                'TOTAL BALANCE': total_balance,
                'CALL DROP #': call_drop_count,
                'SYSTEM DROP': system_drop,
                'CALL DROP RATIO #': call_drop_ratio_formatted,
                'POSITIVE SKIP CONNECTED': positive_skip_connected,
                'NEGATIVE SKIP CONNECTED': negative_skip_connected,
                'POSITIVE SKIP TALK TIME': positive_skip_talk_time,
                'NEGATIVE SKIP TALK TIME': negative_skip_talk_time,
            }
            
            summary_table = pd.concat([summary_table, pd.DataFrame([summary_data])], ignore_index=True)
        
        return summary_table.sort_values(by=['DATE'])

    summary = calculate_summary(df)

    st.write("## Daily Remark Summary Table")
    st.write(summary)

    st.download_button(
        label="Download Summary as Excel",
        data=to_excel(summary),
        file_name=f"Daily_Remark_Summary_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
