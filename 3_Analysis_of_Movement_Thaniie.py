
import streamlit as st
import sys
import ctypes
import numpy as np
import pandas as pd
from datetime import datetime, date, timedelta
import os
from Asset_Calculation_Thaniie import run_Asset_Calculation


st.title('Asset Modeling App')
st.caption('Python-based GUI application to calculate asset AoM between current position and previous position including KRD/convexity breakdown by-fund and derivatives programme')

#Valuation date default values
#dt_val_date_curr = dt.date.today().replace(day=1) - dt.timedelta(days=1)
#dt_val_date_prev = dt_val_date_curr.replace(day=1) - dt.timedelta(days=1)

#Folder paths (start and end with additional '\', e.g., '\\A\\' is '\A\')
#str_pals_path_curr = '\\\Goeqxpwfis010\pals\TH\IFRS9\\'
#str_pals_path_prev = '\\\Goeqxpwfis010\pals\TH\IFRS9\\'
#str_sch_path_curr = '\\\Goeqxpwfis010\pals\SCH\\'
#str_sch_path_prev = '\\\Goeqxpwfis010\pals\SCH\\'

#File names
#str_pals_file_curr  = 'PALS_X1_FAM_BI_TH'
#str_pals_file_prev  = 'PALS_X1_FAM_BI_TH'
#str_sch_red_file_curr  = 'PALS_RED_SCH_BI'
#str_sch_red_file_prev = 'PALS_RED_SCH_BI'
#str_sch_step_file_curr  = 'PALS_STEP_SCH_BI'
#str_sch_step_file_prev = 'PALS_STEP_SCH_BI'
#str_sch_sink_file_curr  = 'PALS_SINK_SCH_BI'
#str_sch_sink_file_prev = 'PALS_SINK_SCH_BI'

#.csv file encoding
str_csv_encoding = 'cp874'



#Text string for user interface for PALS and Schedules files
#ui_str_val_date_curr = 'Provide valuation date for current position (*month end date only)'
#ui_str_val_date_prev = 'Provide valuation date for previous position (*month end date only)'

#ui_str_pals_path_curr = 'Provide folder path of PALS asset data for current position'
#ui_str_pals_path_prev = 'Provide folder path of PALS asset data for previous position'
#ui_str_sch_path_curr = 'Provide folder path of redemption/step-Up/sinking schedules for current position'
#ui_str_sch_path_prev = 'Provide folder path of redemption/step-Up/sinking schedules for previous position'

#ui_str_pals_file_curr = 'Provide file name prefix of PALS asset data for current position'
#ui_str_pals_file_prev = 'Provide file name prefix of PALS asset data for previous position'
#ui_str_red_file_curr = 'Provide file name prefix of redemption schedule for current position'
#ui_str_red_file_prev = 'Provide file name prefix of redemption schedule for previous position'
#ui_str_step_file_curr = 'Provide file name prefix of step-up schedule for current position'
#ui_str_step_file_prev = 'Provide file name prefix of step-up schedule for previous position'
#ui_str_sink_file_curr = 'Provide file name prefix of sinking schedule for current position'
#ui_str_sink_file_prev = 'Provide file name prefix of sinking schedule for previous position'

ui_str_csv_encoding = 'Provide encoding type of .csv file for current position'



col1, col2 = st.columns([2,2], gap="small")

with col1:
    st.subheader('Input parameters', divider='blue')
    val_date = st.date_input('Provide valuation date (*month end date only)', date.today().replace(day=1) -  timedelta(days=1)).strftime('%Y%m%d')

    #Date selectors for valuation dates
    #val_date_curr = st.date_input(ui_str_val_date_curr , dt_val_date_curr).strftime('%Y%m%d')
    #val_date_prev = st.date_input(ui_str_val_date_prev, dt_val_date_prev).strftime('%Y%m%d')
    
    #Text selectors for .csv file encoding type
    csv_encoding = st.text_input(ui_str_csv_encoding, str_csv_encoding)
    test_pals_file_prev = st.file_uploader("Choose a file for PALS asset data for previous position", type = ['xlsx','xls','xlsm'], help='Sample format: \\\Goeqxpwfis010\pals\TH\IFRS9\PALS_X1_FAM_BI_TH20231130.csv')
    #df_pals_file_prev = pd.read_csv(test_pals_file_prev, sep = ';', encoding = str_csv_encoding)

    
with col2:
    st.subheader('PALS & schedules folders', divider='blue')
    file_path = st.text_input("Enter the path of the Excel file", 'C:\Python_Intern\Asset_Movement_Excel\Asset_Movement_Jan2024_v11.xlsm')
    output_path_user = st.text_input("Enter output file path", 'C:/Users/E135863/AIA Group Ltd/FSR Internship - General/02. Structural Risk/Asset Analysis of Movement App/Output')
    test_pals_file_curr = st.file_uploader("Choose a file for PALS asset data for current position", type = ['xlsx','xls','xlsm'],  help='Sample format: \\\Goeqxpwfis010\pals\TH\IFRS9\PALS_X1_FAM_BI_TH20231130.csv')
    #test_pals_file_prev = st.file_uploader("Choose a file for PALS asset data for previous position", help='Sample format: \\\Goeqxpwfis010\pals\TH\IFRS9\PALS_X1_FAM_BI_TH20231130.csv')
    #df_pals_file_curr = pd.read_csv(test_pals_file_curr, sep = ';', encoding = str_csv_encoding)
    #df_pals_file_prev = pd.read_csv(test_pals_file_prev, sep = ';', encoding = str_csv_encoding)
    
#     #Text selectors for .csv files folders
#     pals_path_curr = st.text_input(ui_str_pals_path_curr, str_pals_path_curr)
#     pals_path_prev = st.text_input(ui_str_pals_path_prev, str_pals_path_prev)
#     sch_path_curr = st.text_input(ui_str_sch_path_curr, str_sch_path_curr)
#     sch_path_prev = st.text_input(ui_str_sch_path_prev, str_sch_path_prev)

# with col3:
#     st.subheader('PALS & Schedules Files', divider='red')
#     #Text selectors for .csv files names
#     pals_file_curr = st.text_input(ui_str_pals_file_curr, str_pals_file_curr)
#     pals_file__prev = st.text_input(ui_str_pals_file_prev, str_pals_file_prev)
#     sch_red_file_curr = st.text_input(ui_str_red_file_curr, str_sch_red_file_curr)
#     sch_red_file_prev = st.text_input(ui_str_red_file_prev, str_sch_red_file_prev)
#     sch_step_file_curr = st.text_input(ui_str_step_file_curr, str_sch_step_file_curr)
#     sch_step_file_prev = st.text_input(ui_str_step_file_prev, str_sch_step_file_prev)
#     sch_sink_file_curr = st.text_input(ui_str_sink_file_curr, str_sch_sink_file_curr)
#     sch_sink_file_prev = st.text_input(ui_str_sink_file_prev , str_sch_sink_file_prev)
 
 
df_pals_file_curr = None      
#st.write(df_pals_file_curr)
#st.write(df_pals_file_prev)

# def file_selector(folder_path='.'):
#     filenames = os.listdir(folder_path)
#     selected_filename = st.selectbox('Select a file', filenames)
#     return os.path.join(folder_path, selected_filename)

# filename = file_selector()
# st.write('You selected `%s`' % filename)
if st.button('Run Asset Calculation'):
    if file_path and output_path_user:
        try:
            message = run_Asset_Calculation(file_path, output_path_user, val_date )
            st.success(f"✅ Asset Calculation ran successfully.")
            st.success(f"✅ Output File Location: {message}")
            #output_path_result = run_calibration(file_path, output_path_user)
            if os.path.exists(message):
                df_pals_file_curr = pd.read_excel(message)
                st.write(df_pals_file_curr)
            else:
                st.error('Asset Calculation output file not found', icon="🚨")
        except Exception as e:
            st.error(f"🚨 An error occured during Asset Calculation: {e}")
    else:
        st.warning("🔥 Please specify both the input file path and the output file directory.")
