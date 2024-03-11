import streamlit as st
import numpy as np
import pandas as pd
from datetime import datetime, date, timedelta
import os
from Asset_Cashflows_Thaniie import run_Asset_Cashflows

st.title('Asset Cashflows')
st.caption('Application to calculate asset AoM between current position and previous position')

# Valuation date default values
#dt_val_date_curr = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
#t_val_date_prev = dt_val_date_curr.replace(day=1) - datetime.timedelta(days=1)

# Correcting the folder paths
#str_pals_path_curr = 'C:/Users/E135863/AIA Group Ltd/FSR Internship - General/02. Structural Risk/Asset Analysis of Movement App/PALS/TH/IFRS9/'
#str_pals_path_prev = str_pals_path_curr # Repeated for emphasis; adjust if needed
#str_sch_path_curr = 'C:/Users/E135863/AIA Group Ltd/FSR Internship - General/02. Structural Risk/Asset Analysis of Movement App/PALS/SCH/'
#str_sch_path_prev = str_sch_path_curr # Repeated for emphasis; adjust if needed

# CSV file encoding
str_csv_encoding = 'cp874'

# Interface for uploading files and specifying paths
st.subheader('Input parameters', divider='violet')
val_date = st.date_input('Provide valuation date (*month end date only)', date.today().replace(day=1) -  timedelta(days=1)).strftime('%Y%m%d')
#val_date_curr = st.date_input("Provide valuation date for current position (*month end date only)", dt_val_date_curr)
#val_date_prev = st.date_input("Provide valuation date for previous position (*month end date only)", dt_val_date_prev)

csv_encoding = st.text_input("Provide encoding type of .csv file for current position", str_csv_encoding)
input_file_path = st.text_input("Enter the input file path", value='C:/Python_Intern/Asset_Movement_Excel/Asset_Movement_Jan2024_v11.xlsm', help='Path to the Excel file with input parameters. EX: C:/Users/E135863/AIA Group Ltd/FSR Internship - General/02. Structural Risk/Asset Analysis of Movement App/Asset_Movement_Jan2024_v11.xlsm ')
output_path_cashflow = st.text_input("Enter output file path", value='C:/Users/E135863/AIA Group Ltd/FSR Internship - General/02. Structural Risk/Asset Analysis of Movement App/Output', help='Directory where the output Excel file will be saved. EX: C:/Users/E135863/AIA Group Ltd/FSR Internship - General/02. Structural Risk/Asset Analysis of Movement App/Output')

test_pals_file_curr = st.file_uploader("Choose a file for PALS asset data for current position", help='Sample format: PALS_X1_FAM_BI_TH20231130.csv')
test_pals_file_prev = st.file_uploader("Choose a file for PALS asset data for previous position", help='Sample format: PALS_X1_FAM_BI_TH20231130.csv')

df_pals_file_curr = None
df_pals_file_prev = None

if test_pals_file_curr is not None:
    df_pals_file_curr = pd.read_csv(test_pals_file_curr, sep=';', encoding=str_csv_encoding)

if test_pals_file_prev is not None:
    df_pals_file_prev = pd.read_csv(test_pals_file_prev, sep=';', encoding=str_csv_encoding)

if df_pals_file_curr is not None:
    st.write(df_pals_file_curr)

if df_pals_file_prev is not None:
    st.write(df_pals_file_prev)

if st.button('Run Asset Cashflows'):
    if input_file_path and output_path_cashflow:
        try:
            result_message = run_Asset_Cashflows(input_file_path, output_path_cashflow, val_date)
            st.success(f"✅ Asset Cashflows ran successfully")
            st.success(f"✅ Output File Location: {result_message}")
            if os.path.exists(result_message):
                display_result = pd.read_excel(result_message)
                st.write(display_result)
            else:
                st.error('Asset Cashflows output file not found', icon="🚨")
        except Exception as e:
            st.error(f"🚨 An error occured during Asset Cashflows calculation: {e}")
    else:
        st.warning("🔥 Please specify both the input file path and the output file directory.")
