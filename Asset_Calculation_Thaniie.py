# -*- coding: utf-8 -*-
"""
Created on Tue Feb 27 13:57:55 2024

The script to generate PALS .xlsx file containing asset data to be used for Market Risk Dashboard by TH Financal & Strategic Risk team

@Input: The script requires the following inputs (either from manual input by user or from VBA shell macro)

    *file_para = path to spreadsheet containing all inputs to be used in the script.
    *tab_para = spreadsheet tab name link to spreadsheet per file_para containing parameters to be used in the script.

@Output: The script will generate the PALS .xlsx file in specified path per tab_para with following specification
    
    *Input preparation: import valuation dates, PALS, discount curves, FX rate and other parameters for current/previous positions
    *Input transformation: define data classifier function, data classification, aggregate security then join the current/previous positions
    *Discount curve input transformation: Convert annual compounding discount curves to monthly compounding and assign dates
    *Function preparation: define cashflows projection function and spread goal-seeking function
    *Calculation loop: loop through each security and calculate duration, KRD, convexity, Z-spread and market values for movement analysis
    *Export output: export output table in .xlsx format to specified path

@author: Thaniie

"""       
#%% Import library and parameters from command prompt
import sys
import ctypes
import numpy as np
import pandas as pd
from datetime import datetime, date
import time
import os
import openpyxl
    

def run_Asset_Calculation(input_Path, output_directory, date_time_str):
        
    #Input parameters from VBA
    start_time = time.time()
    file_para = input_Path
    excel = pd.ExcelFile(file_para)
    #file_para = r'C:\Python_Intern\Asset_Movement_Excel\Asset_Movement_Jan2024_v11.xlsm'
    tab_para = 'Python_Para'
    

    # file_para = r'D:\ERM\01 Structural Risk\01 Market Risk Dashboard\082023\Asset_Movement_Aug2023_v9.3_EXE.xlsm'
    # tab_para = 'Python_Para'

    #%% Input preparation

    #Input parameters from Excel parameter file
    Parameters = pd.read_excel(excel, sheet_name = tab_para, engine='openpyxl')
    Value_Date = datetime.strptime(date_time_str, "%Y%m%d")
    rfr_file_name = "Asset_Movement_" + Value_Date.strftime("%b %Y") + ".xlsx"
    output_path = os.path.join(output_directory, rfr_file_name)

    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    #Valuation dates parameters
    val_year = Parameters.loc[Parameters.iloc[:,0] == 'Year'].iloc[0,1]
    val_month = Parameters.loc[Parameters.iloc[:,0] == 'Month'].iloc[0,1]
    val_day = Parameters.loc[Parameters.iloc[:,0] == 'Date'].iloc[0,1]
    val_year_prev = Parameters.loc[Parameters.iloc[:,0] == 'Year_Prev'].iloc[0,1]
    val_month_prev = Parameters.loc[Parameters.iloc[:,0] == 'Month_Prev'].iloc[0,1]
    val_day_prev = Parameters.loc[Parameters.iloc[:,0] == 'Date_Prev'].iloc[0,1]

    #Input and output paths parameters
    file_asset = Parameters.loc[Parameters.iloc[:,0] == 'Input Full Path: Asset'].iloc[0,1]
    file_redem = Parameters.loc[Parameters.iloc[:,0] == 'Input Full Path: Redemption Schedule'].iloc[0,1]
    file_step = Parameters.loc[Parameters.iloc[:,0] == 'Input Full Path: Step Schedule'].iloc[0,1]
    file_sink = Parameters.loc[Parameters.iloc[:,0] == 'Input Full Path: Sinking Schedule'].iloc[0,1]
    file_asset_prev = Parameters.loc[Parameters.iloc[:,0] == 'Input Full Path: Asset_Prev'].iloc[0,1]
    file_redem_prev = Parameters.loc[Parameters.iloc[:,0] == 'Input Full Path: Redemption Schedule_Prev'].iloc[0,1]
    file_step_prev = Parameters.loc[Parameters.iloc[:,0] == 'Input Full Path: Step Schedule_Prev'].iloc[0,1]
    file_sink_prev = Parameters.loc[Parameters.iloc[:,0] == 'Input Full Path: Sinking Schedule_Prev'].iloc[0,1]
    #file_output = Parameters.loc[Parameters.iloc[:,0] == 'Output Full Path'].iloc[0,1]

    #Input and output paths parameters
    tab_fx = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: FX Rate'].iloc[0,1]
    tab_curve = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Discount Curve'].iloc[0,1]
    tab_curve_dur = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Discount Curve_Dur'].iloc[0,1]
    tab_curve_prev = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Discount Curve_Prev'].iloc[0,1]
    tab_eff_date = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Effective Date'].iloc[0,1]
    tab_krd = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: KRD Factor'].iloc[0,1]
    tab_field_col = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Field Column'].iloc[0,1]
    tab_product = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Product Type'].iloc[0,1]
    Tab_COR = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: COR Type'].iloc[0,1]
    Tab_Rating = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Rating Type'].iloc[0,1]
    Tab_Result_Col = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Result Column'].iloc[0,1]
    tab_krd_Result_Col = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: KRD Result Column'].iloc[0,1]
    Tab_IRS = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: IRS Exchange Notional'].iloc[0,1]
    Tab_FAIR2 = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: FAIR2'].iloc[0,1]
    Tab_Eff_Date = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Effective Date'].iloc[0,1]


    #Valuation dates, effective duration shock, FX rate tables and MV indicator setting
    Val_Date_Prev = date(val_year_prev,val_month_prev,val_day_prev)
    Val_Date = date(val_year,val_month,val_day)
    Eff_Dur_Shock = Parameters.loc[Parameters.iloc[:,0] == 'Effective Duration Shock'].iloc[0,1]
    FX_Rate = pd.read_excel(file_para,sheet_name = tab_fx, names = ['Date','Index', 'Currency','Currency_Full','FX/USD','USD/FX'])
    FX_Rate['Date'] = pd.to_datetime(FX_Rate['Date']).dt.date
    MV_Indi = Parameters.loc[Parameters.iloc[:,0] == 'MV Indicator'].iloc[0,1]

    #.csv input encoding and redemption/sinking/stepped-up coupon flags setting
    Input_Encoding = Parameters.loc[Parameters.iloc[:,0] == 'Input Encoding'].iloc[0,1]
    Dur_Redemp_Flag = Parameters.loc[Parameters.iloc[:,0] == 'Duration: Redemption Flag'].iloc[0,1]
    Dur_Step_Flag = Parameters.loc[Parameters.iloc[:,0] == 'Duration: Step Flag'].iloc[0,1]
    Dur_Sink_Flag = Parameters.loc[Parameters.iloc[:,0] == 'Duration: Sinking Flag'].iloc[0,1]
    Spread_Redemp_Flag = Parameters.loc[Parameters.iloc[:,0] == 'Spread: Redemption Flag'].iloc[0,1]
    Spread_Step_Flag = Parameters.loc[Parameters.iloc[:,0] == 'Spread: Step Flag'].iloc[0,1]
    Spread_Sink_Flag = Parameters.loc[Parameters.iloc[:,0] == 'Spread: Sinking Flag'].iloc[0,1]

    #Row and column filter parameters including COR, credit rating, and FAIR2 fund
    Filter = list(pd.read_excel(file_para,sheet_name = tab_field_col)[tab_field_col])
    Asset_Type = pd.read_excel(file_para,sheet_name = tab_product)
    Product_Type = tuple(list(pd.read_excel(file_para,sheet_name = tab_product)[tab_product]))
    COR_Type = pd.read_excel(file_para,sheet_name = Tab_COR)
    Rating_Type = pd.read_excel(file_para,sheet_name = Tab_Rating)
    IRS_Ex_Notional = tuple(list(pd.read_excel(file_para,sheet_name = Tab_IRS)[Tab_IRS]))
    FAIR2 = pd.read_excel(file_para,sheet_name = Tab_FAIR2)

    #Result column name
    Result = tuple(list(pd.read_excel(file_para,sheet_name = Tab_Result_Col)[Tab_Result_Col]))
    KRD_Result = tuple(list(pd.read_excel(file_para,sheet_name = tab_krd_Result_Col)[tab_krd_Result_Col]))

    #Import asset data, schedule files, discount curves and asset effective date
    try:
        PALS = pd.read_csv(file_asset, sep = ';', encoding = Input_Encoding)
        PALS_Prev = pd.read_csv(file_asset_prev, sep = ';', encoding = Input_Encoding)
        REDEM = pd.read_csv(file_redem, sep = ';', encoding = Input_Encoding)
        REDEM_Prev = pd.read_csv(file_redem_prev, sep = ';', encoding = Input_Encoding)
        STEP = pd.read_csv(file_step, sep = ';', encoding = Input_Encoding)
        STEP_Prev = pd.read_csv(file_step_prev, sep = ';', encoding = Input_Encoding)
        SINK = pd.read_csv(file_sink, sep = ';', encoding = Input_Encoding)
        SINK_Prev = pd.read_csv(file_sink_prev, sep = ';', encoding = Input_Encoding)

    except FileNotFoundError as File_Input_Error:
        ctypes.windll.user32.MessageBoxW(0, "Asset/Schedule File: "+repr(File_Input_Error), repr(FileNotFoundError.__qualname__), 0)
        sys.exit(1)

    DISC = pd.read_excel(file_para,sheet_name = tab_curve)
    DISC_Dur = pd.read_excel(file_para,sheet_name = tab_curve_dur)
    DISC_Prev = pd.read_excel(file_para,sheet_name = tab_curve_prev)
    PARA_EFF_DATE = pd.read_excel(file_para,sheet_name = Tab_Eff_Date)
    PARA_KRD = pd.read_excel(file_para,sheet_name = tab_krd)

    #FX rate table transformation
    FX_Table = FX_Rate[FX_Rate['Date'] == Val_Date]
    FX_Table['Currency'] = FX_Table['Currency'].str.strip()
    FX_Table_Prev = FX_Rate[FX_Rate['Date'] == Val_Date_Prev]
    FX_Table_Prev['Currency'] = FX_Table_Prev['Currency'].str.strip()
    
    #%% Input transformation

    #Function to classify asset into inforce, buy, sell and maturity
    def Asset_Category (row):
        
        if row['True'] == 'left_only':
        
            return 'Buy'
    
        elif row['True'] == 'right_only':
        
            if row['Redemp_Year_Prev'] <= val_year and row['Redemp_Month_Prev'] <= val_month and (row['Redemp_Year_Prev'] != 0 or row['Redemp_Month_Prev'] != 0):
            
                return 'Maturity'
        
            else:
            
                return 'Sell'
    
        else:
        
            return 'Inforce'
    
    #Function to use clean MV if dirty MV is N/A
    def MV_Con_Prev (row):
    
        if MV_Indi == 0:
        
            if row['Dirty_MV_RC_Prev'] != 0:
            
                return row['Dirty_MV_RC_Prev']
        
            else:
            
                return row['MV_RC_Prev']
    
        else:
        
            return row['MV_RC_Prev']
    
    def MV_Con (row):
    
        if MV_Indi == 0:
        
            if row['Dirty_MV_RC'] != 0:
            
                return row['Dirty_MV_RC']
        
            else:
            
                return row['MV_RC']
    
        else:
        
            return row['MV_RC']

    #Function to classify fund, asset type, credit type, COR type, rating type, and swap type
    def Credit_Con (row):
    
        if row['Credit_Type_EC'] == 'GOV' or row['Credit_Type_EC_Prev'] == 'GOV':
        
            return 'GOV'
    
        else:
        
            return 'CORP'

    def Swap_Con (row):
    
        if row['Effective_Year'] > val_year_prev or (row['Effective_Year'] == val_year_prev and row['Effective_Month'] > val_month_prev):
        
            return 'Forward'
    
        else:
        
            return 'Spot'

    #Function to split NL fund
    def FAIR2_Con (row):
    
        if row['Company_Code'] == 'TH02':
        
            return 'NL'
    
        else:
        
            return row[Tab_FAIR2]

    #Asset data transformation 
    Fund = FAIR2[FAIR2['FAIR2'] != 'ILP']
    Fund = pd.concat([Fund['FAIR2'].drop_duplicates().reset_index(drop=True), pd.Series(['NL'])], ignore_index=True)

    #Remove unused product type from the data and construct new Data_Control column
    PALS = pd.DataFrame(PALS, columns = Filter)
    PALS = PALS[PALS['Product_Type'].isin(Product_Type)].copy()
    PALS = PALS.merge(FAIR2, how='left', on = FAIR2.columns[0])
    PALS[Tab_FAIR2] = PALS.apply (lambda row: FAIR2_Con(row), axis=1)
    PALS['Data_Control'] = PALS['Security_ID'] +'_'+ PALS['Currency'] +'_'+ PALS['Pay_Receive'] + '_' + PALS['Trade_Date'].astype(str)

    PALS_Prev = pd.DataFrame(PALS_Prev, columns = Filter)
    PALS_Prev = PALS_Prev[PALS_Prev['Product_Type'].isin(Product_Type)].copy()
    PALS_Prev = PALS_Prev.merge(FAIR2, how='left', on = FAIR2.columns[0])
    PALS_Prev[Tab_FAIR2] = PALS_Prev.apply (lambda row: FAIR2_Con(row), axis=1)
    PALS_Prev['Data_Control'] = PALS_Prev['Security_ID'] +'_'+ PALS_Prev['Currency'] +'_'+ PALS_Prev['Pay_Receive'] + '_' + PALS_Prev['Trade_Date'].astype(str)

    #Calculation loop to calculation par value of each asset under each fund
    for i in range(len(Fund)):
    
        PALS[Fund[i]] = PALS[PALS['FAIR2'] == Fund[i]].groupby('Data_Control').Par_Value.transform('sum')
        PALS[Fund[i]] = PALS.groupby('Data_Control')[Fund[i]].transform('sum') / PALS.groupby('Data_Control')[Fund[i]].transform('count')
    
        PALS_Prev[Fund[i]] = PALS_Prev[PALS_Prev['FAIR2'] == Fund[i]].groupby('Data_Control').Par_Value.transform('sum')
        PALS_Prev[Fund[i]] = PALS_Prev.groupby('Data_Control')[Fund[i]].transform('sum') / PALS_Prev.groupby('Data_Control')[Fund[i]].transform('count')


    #Aggregate par, dirty MV, accrued interest for each asset ID and calculate proportion to assign to each fund
    PALS['Par_Value'] = PALS.groupby('Data_Control').Par_Value.transform('sum')
    PALS[Fund] = PALS[Fund].div(PALS['Par_Value'], axis=0)
    PALS['Dirty_MV_RC'] = PALS.groupby('Data_Control').Dirty_MV_RC.transform('sum')
    PALS['MV_RC'] = PALS.groupby('Data_Control').MV_RC.transform('sum')
    PALS['Accrued_Interest_RC'] = PALS.groupby('Data_Control').Accrued_Interest_RC.transform('sum')
    PALS = PALS.drop_duplicates(subset='Data_Control').reset_index(drop=True).fillna(0)

    PALS_Prev['Par_Value'] = PALS_Prev.groupby('Data_Control').Par_Value.transform('sum')
    PALS_Prev[Fund] = PALS_Prev[Fund].div(PALS_Prev['Par_Value'], axis=0)
    PALS_Prev['Dirty_MV_RC'] = PALS_Prev.groupby('Data_Control').Dirty_MV_RC.transform('sum')
    PALS_Prev['MV_RC'] = PALS_Prev.groupby('Data_Control').MV_RC.transform('sum')
    PALS_Prev['Accrued_Interest_RC'] = PALS_Prev.groupby('Data_Control').Accrued_Interest_RC.transform('sum')
    PALS_Prev = PALS_Prev.drop_duplicates(subset='Data_Control').reset_index(drop=True).fillna(0)

    #Merge current and previous month data together based on new Data_Control and reassign the derivatives effective date from Excel table
    Asset_Data = pd.merge(PALS, PALS_Prev, how='outer', on='Data_Control', suffixes = ("","_Prev"), indicator="True")
    Asset_Data = Asset_Data.merge(PARA_EFF_DATE[['Security_ID','Effective_Year']], how='left', on='Security_ID').fillna({'Effective_Year':1900})
    Asset_Data = Asset_Data.merge(PARA_EFF_DATE[['Security_ID','Effective_Month']], how='left', on='Security_ID').fillna({'Effective_Month':1})
    Asset_Data = Asset_Data.drop_duplicates(subset='Data_Control').reset_index(drop=True)

    #Create new columns to store the corresponding FX rates under both current and previous valuation positions
    Asset_Data = Asset_Data.merge(FX_Table[['Currency','FX/USD']], how='left', on='Currency', copy=False)
    Asset_Data = Asset_Data.merge(FX_Table_Prev[['Currency','FX/USD']], how='left', left_on='Currency_Prev', right_on='Currency', copy=False, suffixes = ("","_Prev"))

    #Create new columns for asset category, MV in THB and par value in THB
    Asset_Data['Asset_Category'] = Asset_Data.apply (lambda row: Asset_Category(row), axis=1)
    Asset_Data['Dirty_MV_THB_Prev'] = Asset_Data.apply (lambda row: MV_Con_Prev(row), axis=1) * FX_Table_Prev.loc[FX_Table_Prev.loc[:,'Currency'] == 'THB']['FX/USD'].iloc[0]
    Asset_Data['Dirty_MV_THB'] = Asset_Data.apply (lambda row: MV_Con(row), axis=1) * FX_Table.loc[FX_Table.loc[:,'Currency'] == 'THB']['FX/USD'].iloc[0]
    Asset_Data['Par_Value_THB_Prev'] = Asset_Data['Par_Value_Prev'] * FX_Table_Prev.loc[FX_Table_Prev.loc[:,'Currency'] == 'THB']['FX/USD'].iloc[0] / Asset_Data['FX/USD_Prev']
    Asset_Data['Par_Value_THB'] = Asset_Data['Par_Value'] * FX_Table.loc[FX_Table.loc[:,'Currency'] == 'THB']['FX/USD'].iloc[0] / Asset_Data['FX/USD']
    Asset_Data['Accrued_Interest_THB_Prev'] = Asset_Data['Accrued_Interest_RC_Prev'] * FX_Table_Prev.loc[FX_Table_Prev.loc[:,'Currency'] == 'THB']['FX/USD'].iloc[0]
    Asset_Data['Accrued_Interest_THB'] = Asset_Data['Accrued_Interest_RC'] * FX_Table.loc[FX_Table.loc[:,'Currency'] == 'THB']['FX/USD'].iloc[0]

    #Create FAIR2 columns and asset type columns
    Asset_Data = Asset_Data.merge(Asset_Type, how='left', on = Asset_Type.columns[0])
    Asset_Data = Asset_Data.merge(Asset_Type, how='left', left_on = Asset_Type.columns[0] + '_Prev', right_on = Asset_Type.columns[0], suffixes=("", "_Prev"))

    #Create credit type, COR, Rating, Swap_Type and drop column duplicates
    Asset_Data['Credit_Type'] = Asset_Data.apply (lambda row: Credit_Con(row), axis=1)
    Asset_Data = Asset_Data.merge(COR_Type, how='left', on = COR_Type.columns[0]).fillna({COR_Type.columns[1]:'N/A'})
    Asset_Data = Asset_Data.merge(Rating_Type, how='left', on = Rating_Type.columns[0]).fillna({Rating_Type.columns[1]:'N/A'})
    Asset_Data['Swap_Type'] = Asset_Data.apply (lambda row: Swap_Con(row), axis=1)

    #Placeholder columns for calculation results, drop duplicate rows, drop ILP data and reset dataframe index 
    Asset_Data = Asset_Data.loc[:,~Asset_Data.columns.duplicated()]
    Asset_Data = Asset_Data.reindex(columns=[*Asset_Data.columns, *Result], fill_value=0)
    Asset_Data = Asset_Data.drop_duplicates()
    Asset_Data = Asset_Data[Asset_Data.FAIR2 != 'ILP']
    Asset_Data = Asset_Data[Asset_Data.FAIR2_Prev != 'ILP']
    Asset_Data.reset_index(drop=True, inplace=True)

    #Schedule files transformation
    REDEM = REDEM[REDEM['INDEX'].isin(Asset_Data['Security_ID'])]
    REDEM_Prev = REDEM_Prev[REDEM_Prev['INDEX'].isin(Asset_Data['Security_ID'])]
    STEP = STEP[STEP['INDEX'].isin(Asset_Data['Security_ID'])]
    STEP_Prev = STEP_Prev[STEP_Prev['INDEX'].isin(Asset_Data['Security_ID'])]
    SINK = SINK[SINK['INDEX'].isin(Asset_Data['Security_ID'])]
    SINK_Prev = SINK_Prev[SINK_Prev['INDEX'].isin(Asset_Data['Security_ID'])]

    #%% Discount curve input transformation

    #Cut down discount curve dataframes to less than tenor 100 years
    DISC = DISC.loc[DISC['Tenor']<=100]
    DISC_Dur = DISC_Dur.loc[DISC_Dur['Tenor']<=100]
    DISC_Prev = DISC_Prev.loc[DISC_Prev['Tenor']<=100]

    #Generate dataframe with dimension 12 months * 100 years = 1200 months then assign dates to each index, starting from index 0 = valuation date and convert to numpy array
    DATE_Prev = pd.DataFrame(np.concatenate([([i]) for i in range(0,1201)]),index=range(1201),columns=['Index'])
    DATE_Prev  = DATE_Prev.assign(Year=pd.date_range(Val_Date_Prev,periods = 1201, freq='M').year)
    DATE_Prev  = DATE_Prev.assign(Month=pd.date_range(Val_Date_Prev,periods = 1201, freq='M').month)
    Year_Prev = pd.date_range(Val_Date_Prev,periods = 1201, freq='M').year.to_numpy()
    Month_Prev = pd.date_range(Val_Date_Prev,periods = 1201, freq='M').month.to_numpy()

    DATE = pd.DataFrame(np.concatenate([([i]) for i in range(0,1201)]),index=range(1201),columns=['Index'])
    DATE = DATE.assign(Year=pd.date_range(Val_Date,periods = 1201, freq='M').year)
    DATE = DATE.assign(Month=pd.date_range(Val_Date,periods = 1201, freq='M').month)
    Year = pd.date_range(Val_Date,periods = 1201, freq='M').year.to_numpy()
    Month = pd.date_range(Val_Date,periods = 1201, freq='M').month.to_numpy()

    #Calculation loop - discount RFR curve preparation for curve used in movement analysis and curve used in duration calculation
    for i in range(1,len(DISC.columns)):
    
        Currency = DISC.columns[i][0:3]
    
        if Currency == 'Unn':
        
            continue
    
        else:
        
            #Convert annual compounding forward rates to monthly compounding for curve used in movement analysis and curve used in duration calculation
            DISC[Currency+'_Fwds_M'] = (1 + DISC[Currency+'_Fwds']) ** (1/12) - 1
            DISC[Currency+'_Fwds_M_Dur'] = (1 + DISC_Dur[Currency+'_Fwds']) ** (1/12) - 1
        
            #For the first loop, generate zero-dataframe with index = 1200 months as starting point and merge with monthly compounding forward rate dataframe
            if i == 1:
            
                Disc_M = pd.merge(pd.DataFrame(np.append(0,np.concatenate([([i]*12) for i in range(1,101)])),columns=['Tenor']), 
                                  DISC[['Tenor',Currency+'_Fwds_M']],how='left',on='Tenor').fillna(0)
                Disc_M.insert(1,'Month',Disc_M.index)
            
            else:
                 
                Disc_M = pd.merge(Disc_M, DISC[['Tenor',Currency+'_Fwds_M']],how='left',on='Tenor').fillna(0)
        
            #Merge dataframe for discount curve used in duration calculation and assign 1 as initial values for discount factor
            Disc_M = pd.merge(Disc_M, DISC[['Tenor',Currency+'_Fwds_M_Dur']],how='left',on='Tenor').fillna(0)
            Disc_M[Currency + '_Disc_RFR'] = 1
            Disc_M[Currency + '_Disc_RFR_Dur'] = 1
        
            #Loop to calculate discount factor for 1200 months for curve used in movement analysis and curve used in duration calculation
            for j in range(1,len(Disc_M)):
                   
                Disc_M.loc[j,Currency +'_Disc_RFR'] = Disc_M.loc[j-1,Currency +'_Disc_RFR']/(1+Disc_M.loc[j,Currency+'_Fwds_M'])
                Disc_M.loc[j,Currency +'_Disc_RFR_Dur'] = Disc_M.loc[j-1,Currency +'_Disc_RFR_Dur']/(1+Disc_M.loc[j,Currency+'_Fwds_M_Dur'])

    #Calculation loop - discount RFR curve preparation for curve used in movement analysis at previous position date
    for i in range(1,len(DISC_Prev.columns)):
    
        Currency = DISC_Prev.columns[i][0:3]
    
        if Currency == 'Unn':
        
            continue
    
        else:
        
            DISC[Currency+'_Fwds_M_Prev'] = (1 + DISC_Prev[Currency+'_Fwds']) ** (1/12) - 1
            Disc_M = pd.merge(Disc_M, DISC[['Tenor',Currency+'_Fwds_M_Prev']],how='left',on='Tenor').fillna(0)
            Disc_M[Currency + '_Disc_RFR_Prev'] = 1
            
            for j in range(1,len(Disc_M)):
                              
                Disc_M.loc[j,Currency +'_Disc_RFR_Prev'] = Disc_M.loc[j-1,Currency +'_Disc_RFR_Prev']/(1+Disc_M.loc[j,Currency+'_Fwds_M_Prev'])
                                
    Disc_M.to_numpy()

    #%% Function preparation
    #Calculation function - cashflows
    def Cashflows_Calc(Year, Month, Val_Date, Sec_ID, Type, Redemp_Year, Redemp_Month, Option_Redemp_Month, Eff_Year, Eff_Month, Redemption, Par, 
                       Coupon, Coupon_Freq, Float_Spread):
             
        #Return 0 if date is earlier than effective date or is equal to valuation date
        if Year < Eff_Year or (Year <= Eff_Year and Month < Eff_Month) or (Year == Val_Date.year and Month == Val_Date.month):
        
            return 0
    
        #Return 0 (ECCS IRS) or negative of par (CCS/ECCS BS) at effective date
        elif Year == Eff_Year and Month == Eff_Month:
        
            if Type == 'IRS':
            
                return 0
            
            else:
        
                return Par * -1
    
        #Return coupon (ECCS IRS) or coupon + redemption amount (others) at redemption date
        #Redemption amount and par adjusted for partial early redemption
        elif Year == Redemp_Year and Month == Option_Redemp_Month:
        
            if Type == 'IRS' and Sec_ID not in IRS_Ex_Notional: #for mortgage IRS which pay principle at redemption date
            
                return Par * (Coupon + Float_Spread) / Coupon_Freq
       
            else:
            
                return Par * (Coupon + Float_Spread) / Coupon_Freq + Redemption
    
            #Return coupon if year is earlier than redemption year and  month equal redemption month or if year is redemption year but month is earlier than redemption month
            #Redemption date adjusted for embedded option assumption
        elif (((Month - Redemp_Month) % (12 / Coupon_Freq) == 0 and Year < Redemp_Year) or 
              ((Month - Redemp_Month) % (12 / Coupon_Freq) == 0 and (Year == Redemp_Year and Month < Option_Redemp_Month))):
        
            return Par * (Coupon + Float_Spread) / Coupon_Freq
    
        else:
       
            return 0
   
    Cashflows_Calc = np.frompyfunc(Cashflows_Calc,Cashflows_Calc.__code__.co_argcount,1)

    #Calculation function - Lookup on redemption schedule and assign new redemption date
    def Redemption_Calc(Year, Month, Val_Date, Sec_ID, Type, Redemp_Year, Redemp_Month, Option_Redemp_Month, Eff_Year,Eff_Month, Redemption, Par, 
                        Coupon, Coupon_Freq, Float_Spread, Date_DF, Redemp_DF, Embbeded_Index, Discount_Curve, Redemp_Flag):
    
        #Check if redemption date is beyond 100 years
        try:
        
            Redemp_Indi = Date_DF['Index'].loc[(Date_DF['Year'] == Redemp_Year)&(Date_DF['Month'] == Redemp_Month)].tolist()[0]
    
        except IndexError:
        
            Redemp_Indi = 1200

        #Lookup on redemption schedule and assign new redemption amount and redemption year & month [subject to redemption flag setting] 
        if Redemp_DF.loc[Redemp_DF['INDEX'] == Sec_ID].empty or Redemp_Flag == 0:
        
            Redemption_Final = Par
            Redemp_Year_Final = Redemp_Year
            Option_Redemp_Month_Final = Redemp_Month
            Redemp_Indi_Final = Redemp_Indi
        
        else:
        
            #Create array of option exercise dates and redemption date for PV calculation loop    
            Redemp_Indi_List = np.append(np.where(np.append(Redemp_DF.loc[Redemp_DF['INDEX'] == Sec_ID].iloc[: , 1:].T,0) != 0),Redemp_Indi)
        
            #Loop to calculate PV for a given option exercise dates/redemption date
            for j in range(len(Redemp_Indi_List)):
            
                #Prepare redemption amount and redemption year & month for redemption date index j
                if Redemp_DF.loc[Redemp_DF['INDEX'] == Sec_ID].iloc[: , 1:].T.iloc[Redemp_Indi_List[j]].tolist()[0] == 0:
                
                    Redemption_Loop = Par
                
                else:
                
                    Redemption_Loop = Par * (Redemp_DF.loc[Redemp_DF['INDEX'] == Sec_ID].iloc[: , 1:].T.iloc[Redemp_Indi_List[j]].tolist()[0] 
                                             - (Coupon + Float_Spread)/Coupon_Freq)

                Redemp_Year_Loop = Year[Redemp_Indi_List[j]]
                Option_Redemp_Month_Loop = Month[Redemp_Indi_List[j]]
            
                #Calculate cashflows and PV for redemption date index j and store it in array
                Cashflows_Loop = Cashflows_Calc(Year, Month, Val_Date, Sec_ID, Type, Redemp_Year_Loop, Redemp_Month, Option_Redemp_Month_Loop,
                                                Eff_Year, Eff_Month, Redemption_Loop, Par, Coupon, Coupon_Freq, Float_Spread)
                PV_Loop = sum(Cashflows_Loop * Discount_Curve)
            
                if j == 0:
                    
                    PV_Loop_List = PV_Loop
                    
                else:
                    
                    PV_Loop_List = np.append(PV_Loop_List, PV_Loop)
                
            #Check whether it's callable (1) or puttable (2) and assign the final redemption date index
            if Embbeded_Index == 1:
            
                Redemp_Indi_Final = Redemp_Indi_List[np.where(PV_Loop_List==min(PV_Loop_List))][0]
            
            else:
                
                Redemp_Indi_Final = Redemp_Indi_List[np.where(PV_Loop_List==max(PV_Loop_List))][0]
            
            #Check whether the option is exercised or not and assign the final redemption amount
            if Redemp_Indi_Final == Date_DF['Index'].loc[(Date_DF['Year'] == Redemp_Year)&(Date_DF['Month'] == Redemp_Month)].tolist()[0] or Redemp_Indi == 1200:
            
                Redemption_Final = Par
            
            else:
                
                Redemption_Final = Par * (Redemp_DF.loc[Redemp_DF['INDEX'] == Sec_ID].iloc[: , 1:].T.iloc[Redemp_Indi_Final].tolist()[0] - (Coupon + Float_Spread)/Coupon_Freq)
        
            #Assign final redemption year and month
            Redemp_Year_Final = Year[Redemp_Indi_Final]
            Option_Redemp_Month_Final = Month[Redemp_Indi_Final]
        
        return Redemption_Final, Redemp_Year_Final, Option_Redemp_Month_Final, Redemp_Indi_Final

    #Calculation function - Lookup on sinking fund schedule to adjust par amount (for coupon calculation) and new cashflows (from partial early redemption) [subject to sinking flag setting] 
    def Sink_Calc(Sink_DF, Sec_ID, Par, Redemption, Redemp_Indi, Sink_Flag):
    
        if Sink_DF.loc[Sink_DF['INDEX'] == Sec_ID].empty or Sink_Flag == 0:
        
            Sink_CF = 0
            Par_Unit = 1
        
        else:
        
            #Calculate cashflows from early partial redemption
            Par_Unit = np.append(0,Sink_DF.loc[Sink_DF['INDEX'] == Sec_ID].iloc[: , 1:].T)
            Sink_CF = Par_Unit * Par
            Sink_CF[Redemp_Indi:] = 0
        
            #Calculate outstanding par value at each time point for coupon cashflows calculation
            for j in range(len(Par_Unit)):
            
                if j == 0:
                
                    Par_Unit[j] = 1
                
                else:
                
                    Par_Unit[j] = Par_Unit[j-1] - Par_Unit[j]
        
            Par_Unit = np.append(0,Par_Unit[:-1]) 
       
        return Par_Unit * Redemption, Par_Unit * Par, Sink_CF

    #Calculation function - Lookup on stepped up coupon schedule and assign new coupon rate [subject to stepped-up coupon flag setting] 
    def Step_Coupon_Calc(Step_DF, Sec_ID, Coupon, Step_Flag):
    
        if Step_DF.loc[Step_DF['INDEX'] == Sec_ID].empty or Step_Flag == 0:
        
            return Coupon
        
        else:
        
            return np.append(0,Step_DF.loc[Step_DF['INDEX'] == Sec_ID].iloc[: , 1:].T) / 100
    
    #Calculation function - goal seek spread
    #Set up objective function to goal seek to 0
    def PV_Spread_Calc(Spread, CF, MV, Disc_Curve, Month):
    
        return sum(CF * ((Disc_Curve ** (-12/Month) + Spread) ** (-Month/12))) - MV

    #Import reequired library and set up spread calculation
    from scipy.optimize import root
    def Spread_Calc(CF, MV, Disc_Curve, Month, x0):
    
        #Check if root finding is sucess or not, if not adjust the initial guess (x0) by +/- 1 bps and resolve for root
        if root(PV_Spread_Calc, x0, args=(CF, MV, Disc_Curve, Month)).success is False:
        
            return root(PV_Spread_Calc, x0+0.0001, args=(CF, MV, Disc_Curve, Month))
    
            if root(PV_Spread_Calc, x0+0.0001, args=(CF, MV, Disc_Curve, Month)).success is False:
            
                return root(PV_Spread_Calc, x0-0.0001, args=(CF, MV, Disc_Curve, Month))
        
            else:
            
            #Return 0 if root is impossible to be found after +/- 1 bps adjustment to x0
                return 0
        else:
        
            return root(PV_Spread_Calc, x0, args=(CF, MV, Disc_Curve, Month))

    #%% Calculation loop

    for i in range(len(Asset_Data)):
 
        Asset_Input = Asset_Data.iloc[i]
    
        #Calculate MV movement, cashflows, spread, duration and convexity for inforce assets
        if Asset_Input['Asset_Category'] == 'Inforce' and Asset_Input['Dirty_MV_THB'] != 0:
        
            #Assign discount curves based on currency and project cashflows as of current and previous valuation dates for spread and movement calculation
            Disc_RFR_Prev = Disc_M[Asset_Input['Currency_Prev']+'_Disc_RFR_Prev']
            Disc_RFR = Disc_M[Asset_Input['Currency']+'_Disc_RFR']
            Disc_RFR_Dur = Disc_M[Asset_Input['Currency']+'_Disc_RFR_Dur']
        
            #Redemption, sinking and stepped-up coupon schedules calculation
            ##For previous valuation date
            Redemp_Output_Prev = Redemption_Calc(Year_Prev, Month_Prev, Val_Date_Prev, Asset_Input['Security_ID'], Asset_Input['Product_Type'], 
                                                 Asset_Input['Redemp_Year'], Asset_Input['Redemp_Month'], Asset_Input['Redemp_Month'],
                                                 Asset_Input['Effective_Year'], Asset_Input['Effective_Month'], Asset_Input['Par_Value_THB_Prev'],
                                                 Asset_Input['Par_Value_THB_Prev'], Asset_Input['Coupon_Rate_Prev'], Asset_Input['Coupon_Freq'],
                                                 Asset_Input['Floating_Spread'], DATE_Prev, REDEM_Prev, Asset_Input['Embedded_Opt_Index'], 
                                                 Disc_RFR_Prev, Spread_Redemp_Flag)
            Sink_Output_Prev = Sink_Calc(SINK_Prev, Asset_Input['Security_ID'], Asset_Input['Par_Value_THB_Prev'], Redemp_Output_Prev[0], 
                                         Redemp_Output_Prev[3], Spread_Sink_Flag)
            Step_Output_Prev = Step_Coupon_Calc(STEP_Prev, Asset_Input['Security_ID'], Asset_Input['Coupon_Rate_Prev'], Spread_Step_Flag)
        
            ##For current valuation date
            Redemp_Output = Redemption_Calc(Year, Month, Val_Date, Asset_Input['Security_ID'], Asset_Input['Product_Type'], Asset_Input['Redemp_Year'], 
                                            Asset_Input['Redemp_Month'], Asset_Input['Redemp_Month'], Asset_Input['Effective_Year'],
                                            Asset_Input['Effective_Month'], Asset_Input['Par_Value_THB'], Asset_Input['Par_Value_THB'], 
                                            Asset_Input['Coupon_Rate'], Asset_Input['Coupon_Freq'], Asset_Input['Floating_Spread'], 
                                            DATE, REDEM,Asset_Input['Embedded_Opt_Index'],Disc_RFR,Spread_Redemp_Flag)
            Sink_Output = Sink_Calc(SINK, Asset_Input['Security_ID'], Asset_Input['Par_Value_THB'], Redemp_Output[0], 
                                    Redemp_Output[3], Spread_Sink_Flag)
            Step_Output = Step_Coupon_Calc(STEP, Asset_Input['Security_ID'], Asset_Input['Coupon_Rate'], Spread_Step_Flag)
                
            ##For current valuation date duration and convexity calculation     
            Redemp_Output_Dur = Redemption_Calc(Year, Month, Val_Date, Asset_Input['Security_ID'], Asset_Input['Product_Type'], Asset_Input['Redemp_Year'], 
                                                Asset_Input['Redemp_Month'], Asset_Input['Redemp_Month'], Asset_Input['Effective_Year'], 
                                                Asset_Input['Effective_Month'], Asset_Input['Par_Value_THB'], Asset_Input['Par_Value_THB'], 
                                                Asset_Input['Coupon_Rate'], Asset_Input['Coupon_Freq'], Asset_Input['Floating_Spread'], 
                                                DATE, REDEM, Asset_Input['Embedded_Opt_Index'], Disc_RFR, Dur_Redemp_Flag)
            Sink_Output_Dur = Sink_Calc(SINK, Asset_Input['Security_ID'], Asset_Input['Par_Value_THB'], Redemp_Output_Dur[0], 
                                        Redemp_Output_Dur[3], Dur_Sink_Flag)
            Step_Output_Dur = Step_Coupon_Calc(STEP, Asset_Input['Security_ID'], Asset_Input['Coupon_Rate'], Dur_Step_Flag)
        
            #Goal seek spread from a given cashflows and discount RFR curve
            ##For previous valuation date
            Cashflows_Prev = Cashflows_Calc(Year_Prev, Month_Prev, Val_Date_Prev, Asset_Input['Security_ID'], Asset_Input['Product_Type'], Redemp_Output_Prev[1], 
                                            Asset_Input['Redemp_Month'], Redemp_Output_Prev[2], Asset_Input['Effective_Year'], Asset_Input['Effective_Month'], 
                                            Sink_Output_Prev[0], Sink_Output_Prev[1], Step_Output_Prev, Asset_Input['Coupon_Freq'], Asset_Input['Floating_Spread']) + Sink_Output_Prev[2]
            Spread_Prev = Spread_Calc(Cashflows_Prev, Asset_Input['Dirty_MV_THB_Prev'], Disc_RFR_Prev, Disc_M['Month'], 0).x.item()
        
            ##For current valuation date
            Cashflows = Cashflows_Calc(Year, Month, Val_Date, Asset_Input['Security_ID'], Asset_Input['Product_Type'], Redemp_Output[1], 
                                       Asset_Input['Redemp_Month'], Redemp_Output[2], Asset_Input['Effective_Year'], Asset_Input['Effective_Month'], 
                                       Sink_Output[0], Sink_Output[1], Step_Output, Asset_Input['Coupon_Freq'], Asset_Input['Floating_Spread']) + Sink_Output[2]
            Spread = Spread_Calc(Cashflows, Asset_Input['Dirty_MV_THB'], Disc_RFR, Disc_M['Month'], 0).x.item()
        
            ##For current valuation date duration and convexity calculation
            Cashflows_Dur = Cashflows_Calc(Year, Month, Val_Date, Asset_Input['Security_ID'], Asset_Input['Product_Type'], Redemp_Output_Dur[1], 
                                           Asset_Input['Redemp_Month'], Redemp_Output_Dur[2], Asset_Input['Effective_Year'], Asset_Input['Effective_Month'], 
                                           Sink_Output_Dur[0], Sink_Output_Dur[1], Step_Output_Dur, Asset_Input['Coupon_Freq'], Asset_Input['Floating_Spread']) + Sink_Output_Dur[2]
            Spread_Dur = Spread_Calc(Cashflows_Dur, Asset_Input['Dirty_MV_THB'], Disc_RFR_Dur, Disc_M['Month'], 0).x.item()
        
            ##For current valuation date rollforward calculation
            Cashflows_Rollforward = Cashflows_Calc(Year, Month, Val_Date, Asset_Input['Security_ID'], Asset_Input['Product_Type'], Redemp_Output[1], 
                                                   Asset_Input['Redemp_Month'], Redemp_Output[2], Asset_Input['Effective_Year'], Asset_Input['Effective_Month'], Sink_Output[0], 
                                                   Sink_Output[1], Step_Output_Prev, Asset_Input['Coupon_Freq'], Asset_Input['Floating_Spread']) + Sink_Output[2]
        
            #Store result to asset dataframe given that spread goal seek is not converged    
            if  Spread_Calc(Cashflows_Prev,Asset_Input['Dirty_MV_THB_Prev'],Disc_RFR_Prev,Disc_M['Month'],0).success is False or Asset_Input['Redemp_Year'] == 9999:
    
                Spread_Prev = Spread = Spread_Dur = 0
                Asset_Data.iloc[i,Asset_Data.columns.get_loc('Calculated_MV_Rollforward')] = Asset_Input['Dirty_MV_THB_Prev']
                Asset_Data.iloc[i,Asset_Data.columns.get_loc('Calculated_MV_Interest_Rate')] = Asset_Input['Dirty_MV_THB_Prev']
                Asset_Data.iloc[i,Asset_Data.columns.get_loc('Calculated_MV_FX')] = Asset_Input['Dirty_MV_THB_Prev']
                Asset_Data.iloc[i,Asset_Data.columns.get_loc('Effective_Duration')] = 0
                Asset_Data.iloc[i,Asset_Data.columns.get_loc('Effective_Convexity')] = 0
            
                for k in range(len(KRD_Result)):
                    Asset_Data.iloc[i,Asset_Data.columns.get_loc(KRD_Result[k])] = 0
                
            else:
            
                #Discount curves and PV calculation
                Disc_Base_Prev = (Disc_RFR_Prev ** (-12/Disc_M['Month']) + Spread_Prev) ** (-Disc_M['Month']/12)
                Disc_Base = (Disc_RFR ** (-12/Disc_M['Month']) + Spread_Prev) ** (-Disc_M['Month']/12)
                Disc_Base_Dur = (Disc_RFR_Dur ** (-12/Disc_M['Month']) + Spread_Dur) ** (-Disc_M['Month']/12)
                Disc_Up = (Disc_Base_Dur ** (-12/Disc_M['Month']) + Eff_Dur_Shock) ** (-Disc_M['Month']/12)
                Disc_Dn = (Disc_Base_Dur ** (-12/Disc_M['Month']) - Eff_Dur_Shock) ** (-Disc_M['Month']/12)
            
                PV_Up = sum(Cashflows_Dur * Disc_Up)
                PV_Dn = sum(Cashflows_Dur * Disc_Dn)
            
                #Calculate by-step MV for rollforward step (with FX impact) and FX step then store result to dataframe
                Asset_Data.iloc[i,Asset_Data.columns.get_loc('Calculated_MV_Rollforward')] = sum(Cashflows_Rollforward * Disc_Base_Prev)
                Asset_Data.iloc[i,Asset_Data.columns.get_loc('Calculated_MV_FX')] = sum(Cashflows * Disc_Base)
            
                #Store duration and convexity in the data frame
                Asset_Data.iloc[i,Asset_Data.columns.get_loc('Effective_Duration')] = (PV_Dn - PV_Up)/(2 * Asset_Input['Dirty_MV_THB'] * Eff_Dur_Shock)
                Asset_Data.iloc[i,Asset_Data.columns.get_loc('Effective_Convexity')] = (PV_Dn + PV_Up - 2 * Asset_Input['Dirty_MV_THB'])/(Asset_Input['Dirty_MV_THB'] * (Eff_Dur_Shock ** 2))      
            
                for k in range(len(KRD_Result)):
                    Asset_Data.iloc[i,Asset_Data.columns.get_loc(KRD_Result[k])] = sum(Cashflows_Dur * (Disc_Dn - Disc_Up) * PARA_KRD[KRD_Result[k]] / 
                                                                                       (2 * Asset_Input['Dirty_MV_THB'] * Eff_Dur_Shock))
        
            #Store spread and cashflows (based on movement calculation) in data frame
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Spread_Prev')] = Spread_Prev
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Spread')] = Spread
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Cashflows')] = sum(Cashflows_Prev[:val_month - val_month_prev % 12 + 1])    
    
        #Calculate spread, duration and convexity for new assets
        elif Asset_Input['Asset_Category'] == 'Buy' and Asset_Input['Dirty_MV_THB'] != 0:
        
            #Assign discount curves based on currency valuation date and project cashflows for spread calculation
            Disc_RFR = Disc_M[Asset_Input['Currency']+'_Disc_RFR']
            Disc_RFR_Dur = Disc_M[Asset_Input['Currency']+'_Disc_RFR_Dur']
        
        
            #Redemption, sinking and stepped-up coupon schedules calculation
            ##For current valuation date
            Redemp_Output = Redemption_Calc(Year,Month, Val_Date, Asset_Input['Security_ID'], Asset_Input['Product_Type'], Asset_Input['Redemp_Year'], 
                                            Asset_Input['Redemp_Month'], Asset_Input['Redemp_Month'], Asset_Input['Effective_Year'],
                                            Asset_Input['Effective_Month'], Asset_Input['Par_Value_THB'], Asset_Input['Par_Value_THB'], 
                                            Asset_Input['Coupon_Rate'], Asset_Input['Coupon_Freq'], Asset_Input['Floating_Spread'], 
                                            DATE, REDEM, Asset_Input['Embedded_Opt_Index'], Disc_RFR, Spread_Redemp_Flag)
            Sink_Output = Sink_Calc(SINK, Asset_Input['Security_ID'], Asset_Input['Par_Value_THB'], Redemp_Output[0], Redemp_Output[3], Spread_Sink_Flag)
            Step_Output = Step_Coupon_Calc(STEP, Asset_Input['Security_ID'], Asset_Input['Coupon_Rate'], Spread_Step_Flag)
        
            ##For current valuation date duration and convexity calculation
            Redemp_Output_Dur = Redemption_Calc(Year, Month, Val_Date, Asset_Input['Security_ID'], Asset_Input['Product_Type'], Asset_Input['Redemp_Year'], 
                                                Asset_Input['Redemp_Month'], Asset_Input['Redemp_Month'], Asset_Input['Effective_Year'], 
                                                Asset_Input['Effective_Month'], Asset_Input['Par_Value_THB'], Asset_Input['Par_Value_THB'], 
                                                Asset_Input['Coupon_Rate'], Asset_Input['Coupon_Freq'], Asset_Input['Floating_Spread'], 
                                                DATE, REDEM, Asset_Input['Embedded_Opt_Index'], Disc_RFR, Dur_Redemp_Flag)
            Sink_Output_Dur = Sink_Calc(SINK, Asset_Input['Security_ID'], Asset_Input['Par_Value_THB'], Redemp_Output_Dur[0], Redemp_Output_Dur[3], Dur_Sink_Flag)
            Step_Output_Dur = Step_Coupon_Calc(STEP, Asset_Input['Security_ID'], Asset_Input['Coupon_Rate'], Dur_Step_Flag)
        
            #Goal seek spread from a given cashflows and discount RFR curve
            ##For current valuation date
            Cashflows = Cashflows_Calc(Year, Month, Val_Date, Asset_Input['Security_ID'], Asset_Input['Product_Type'], Redemp_Output[1], 
                                       Asset_Input['Redemp_Month'], Redemp_Output[2], Asset_Input['Effective_Year'], Asset_Input['Effective_Month'], 
                                       Sink_Output[0], Sink_Output[1], Step_Output, Asset_Input['Coupon_Freq'], Asset_Input['Floating_Spread']) + Sink_Output[2]
            Spread = Spread_Calc(Cashflows, Asset_Input['Dirty_MV_THB'], Disc_RFR, Disc_M['Month'], 0).x.item()
        
            ##For current valuation date duration and convexity calculation
            Cashflows_Dur = Cashflows_Calc(Year, Month, Val_Date, Asset_Input['Security_ID'], Asset_Input['Product_Type'], Redemp_Output_Dur[1], 
                                           Asset_Input['Redemp_Month'], Redemp_Output_Dur[2], Asset_Input['Effective_Year'], Asset_Input['Effective_Month'], 
                                           Sink_Output_Dur[0], Sink_Output_Dur[1], Step_Output_Dur, Asset_Input['Coupon_Freq'], Asset_Input['Floating_Spread']) + Sink_Output_Dur[2]
            Spread_Dur = Spread_Calc(Cashflows_Dur, Asset_Input['Dirty_MV_THB'], Disc_RFR_Dur, Disc_M['Month'], 0).x.item()
                
            #Store result to asset dataframe given that spread goal seek is not converged    
            if  Spread_Calc(Cashflows, Asset_Input['Dirty_MV_THB'], Disc_RFR, Disc_M['Month'], 0).success is False or Asset_Input['Redemp_Year'] == 9999:
            
                Spread = Spread_Dur = 0
                Asset_Data.iloc[i,Asset_Data.columns.get_loc('Effective_Duration')] = 0
                Asset_Data.iloc[i,Asset_Data.columns.get_loc('Effective_Convexity')] = 0
            
                for k in range(len(KRD_Result)):
                    Asset_Data.iloc[i,Asset_Data.columns.get_loc(KRD_Result[k])] = 0
                
            else:
            
                #Discount curves and PV calculation
                Disc_Base_Dur = (Disc_RFR_Dur ** (-12/Disc_M['Month']) + Spread_Dur) ** (-Disc_M['Month']/12)
                Disc_Up = (Disc_Base_Dur ** (-12/Disc_M['Month']) + Eff_Dur_Shock) ** (-Disc_M['Month']/12)
                Disc_Dn = (Disc_Base_Dur ** (-12/Disc_M['Month']) - Eff_Dur_Shock) ** (-Disc_M['Month']/12)
        
                PV_Up = sum(Cashflows_Dur * Disc_Up)
                PV_Dn = sum(Cashflows_Dur * Disc_Dn)
            
                #Store duration and convexity in the data frame
                Asset_Data.iloc[i,Asset_Data.columns.get_loc('Effective_Duration')] = (PV_Dn - PV_Up)/(2 * Asset_Input['Dirty_MV_THB'] * Eff_Dur_Shock)
                Asset_Data.iloc[i,Asset_Data.columns.get_loc('Effective_Convexity')] = (PV_Dn + PV_Up - 2 * Asset_Input['Dirty_MV_THB'])/(Asset_Input['Dirty_MV_THB'] * (Eff_Dur_Shock ** 2))      
            
                for k in range(len(KRD_Result)):
                    Asset_Data.iloc[i,Asset_Data.columns.get_loc(KRD_Result[k])] = sum(Cashflows_Dur * (Disc_Dn - Disc_Up) * PARA_KRD[KRD_Result[k]] / 
                                                                                       (2 * Asset_Input['Dirty_MV_THB'] * Eff_Dur_Shock))
            #Store spread and cashflows in data frame
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Spread')] = Spread
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Cashflows')] = 0
               
        #Store maturity cashflows
        elif Asset_Input['Asset_Category'] == 'Maturity':
        
            if Asset_Input['Product_Type_Prev'] == 'IRS' and Asset_Input['Security_ID_Prev'] not in IRS_Ex_Notional:
            
                Asset_Data.iloc[i,Asset_Data.columns.get_loc('Cashflows')] = Asset_Input['Par_Value_THB_Prev'] * (Asset_Input['Coupon_Rate_Prev'] + Asset_Input['Floating_Spread_Prev'])/Asset_Input['Coupon_Freq_Prev']
       
            else:
            
                Asset_Data.iloc[i,Asset_Data.columns.get_loc('Cashflows')] = Asset_Input['Par_Value_THB_Prev'] * (1 + ((Asset_Input['Coupon_Rate_Prev'] + Asset_Input['Floating_Spread_Prev'])/Asset_Input['Coupon_Freq_Prev']))
            
        else:
        
            #Store zero results for sell/maturity assets
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Spread_Prev')] = 0
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Spread')] = 0
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Calculated_MV_Rollforward')] = 0
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Calculated_MV_Interest_Rate')] = 0
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Calculated_MV_FX')] = 0
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Cashflows')] = 0
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Effective_Duration')] = 0
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Effective_Convexity')] = 0
            for k in range(len(KRD_Result)):
                    Asset_Data.iloc[i,Asset_Data.columns.get_loc(KRD_Result[k])] = 0

    #Calculate accrued interest impact then store MV for rollforward and interest rate steps by removing FX impact

    if MV_Indi == 0:

        Accrued_IR = Asset_Data['Accrued_Interest_THB'] - (Asset_Data['Accrued_Interest_THB_Prev'] * Asset_Data['Par_Value'] 
                                                           / Asset_Data['Par_Value_Prev'] * (FX_Table.loc[FX_Table.loc[:,'Currency'] == 'THB']['FX/USD'].iloc[0] 
                                                                                             / Asset_Data['FX/USD']) / (FX_Table_Prev.loc[FX_Table_Prev.loc[:,'Currency'] == 'THB']['FX/USD'].iloc[0] 
                                                                                                                        / Asset_Data['FX/USD_Prev']))
    else:
    
        Accrued_IR = 0
    
    Asset_Data['Calculated_MV_FX'] = Asset_Data['Calculated_MV_FX'] + Accrued_IR 

    Asset_Data['Calculated_MV_Interest_Rate'] = Asset_Data['Calculated_MV_FX'] * (FX_Table_Prev.loc[FX_Table_Prev.loc[:,'Currency'] == 'THB']['FX/USD'].iloc[0]
                                                                                  / Asset_Data['FX/USD_Prev']) / (FX_Table.loc[FX_Table.loc[:,'Currency'] == 'THB']['FX/USD'].iloc[0]
                                                                                                                  / Asset_Data['FX/USD'])
                                                                                                                                      
    Asset_Data['Calculated_MV_Rollforward'] = (Asset_Data['Calculated_MV_Rollforward'] + Accrued_IR) * (FX_Table_Prev.loc[FX_Table_Prev.loc[:,'Currency'] == 'THB']['FX/USD'].iloc[0]
                                                                                                        / Asset_Data['FX/USD_Prev']) / (FX_Table.loc[FX_Table.loc[:,'Currency'] == 'THB']['FX/USD'].iloc[0]
                                                                                                                                        / Asset_Data['FX/USD'])
                                                 
    #Calculate and store MV with adjustment for partial selling
    Asset_Data['Calculated_MV_Partial_Sell'] = Asset_Data['Dirty_MV_THB_Prev'] * Asset_Data['Par_Value']/Asset_Data['Par_Value_Prev'] 

    #Sum by contract ID for Swap movement under TH RBC balance sheet
    Asset_Data['Dirty_MV_THB_SecurityID_Prev'] = Asset_Data.groupby('Security_ID_Prev').Dirty_MV_THB_Prev.transform('sum')
    Asset_Data['Dirty_MV_THB_SecurityID'] = Asset_Data.groupby('Security_ID').Dirty_MV_THB.transform('sum')
    Asset_Data['Calculated_MV_Partial_Sell_SecurityID'] = Asset_Data.groupby('Security_ID').Calculated_MV_Partial_Sell.transform('sum')
    Asset_Data['Calculated_MV_Rollforward_SecurityID'] = Asset_Data.groupby('Security_ID').Calculated_MV_Rollforward.transform('sum')
    Asset_Data['Calculated_MV_Interest_Rate_SecurityID'] = Asset_Data.groupby('Security_ID').Calculated_MV_Interest_Rate.transform('sum')
    Asset_Data['Calculated_MV_FX_SecurityID'] = Asset_Data.groupby('Security_ID').Calculated_MV_FX.transform('sum')
    Asset_Data['Calculated_MV_Interest_Rate_THB'] = Asset_Data['Calculated_MV_Rollforward']
    Asset_Data.loc[Asset_Data['Currency'] == 'THB', 'Calculated_MV_Interest_Rate_THB'] = Asset_Data['Calculated_MV_Interest_Rate']                                         
    Asset_Data['Calculated_MV_Interest_Rate_THB_SecurityID'] = Asset_Data.groupby('Security_ID').Calculated_MV_Interest_Rate_THB.transform('sum')
        
    #%% Export output

    Asset_Data[['Par_Value','Par_Value_Prev','MV_RC','MV_RC_Prev','Dirty_MV_RC','Dirty_MV_RC_Prev']] = Asset_Data[['Par_Value','Par_Value_Prev','MV_RC','MV_RC_Prev','Dirty_MV_RC','Dirty_MV_RC_Prev']].fillna(0)
    Asset_Data.drop(['Company_Code','Company_Code_Prev','True', 'FAIR2','FAIR2_Prev','FX/USD','FX/USD_Prev'], inplace=True, axis=1)
    Asset_Data.index = np.arange(1, len(Asset_Data)+1)
    Asset_Data.to_excel(output_path, header=True, engine = 'openpyxl')
    print("--- %s seconds ---" % (time.time() - start_time))
    print('Asset Model Calculation ran sucessfully.')
    return output_path
