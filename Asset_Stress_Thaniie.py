# -*- coding: utf-8 -*-
"""
Created on Tue Feb 27 16:20:00 2024

@author: Thaniie
"""
#%% Import library and parameters from command prompt
import sys
import time
import numpy as np
import pandas as pd
from datetime import datetime, date
import openpyxl
import os

def run_Asset_Stress_Test(input_Path, output_directory, date_time_str):

    start_time = time.time()
    #Input parameters from VBA
    File_Para = input_Path
    excel = pd.ExcelFile(File_Para)
    #File_Para = r'C:\Python_Intern\Asset_Movement_Excel\Asset_Movement_Jan2024_v11.xlsm'
    Tab_Para = 'Python_Para_Stress'

    #%% Input preparation

    #Input parameters from Excel parameter file
    Parameters = pd.read_excel(excel, sheet_name = Tab_Para, engine='openpyxl')
    Value_Date = datetime.strptime(date_time_str, "%Y%m%d")
    rfr_file_name = "Asset_Movement_" + Value_Date.strftime("%b %Y") + ".xlsx"
    output_path = os.path.join(output_directory, rfr_file_name)
    
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
   
    
    #Valuation dates parameters
    Val_Year = Parameters.loc[Parameters.iloc[:,0] == 'Year'].iloc[0,1]
    Val_Month = Parameters.loc[Parameters.iloc[:,0] == 'Month'].iloc[0,1]
    Val_Day = Parameters.loc[Parameters.iloc[:,0] == 'Date'].iloc[0,1]

    #Input and output paths parameters
    File_Cashflows = Parameters.loc[Parameters.iloc[:,0] == 'Input Full Path: Cashflows'].iloc[0,1]
    #File_Output = Parameters.loc[Parameters.iloc[:,0] == 'Output Full Path: Stress Test'].iloc[0,1]

    #Input tab parameters
    Tab_Spread = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Spread Data'].iloc[0,1]
    Tab_Curve = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Discount Curve Stress'].iloc[0,1]
    Tab_Stress_IR = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Interest Rate Stress'].iloc[0,1]
    Tab_Stress_CS = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Credit Spread Stress'].iloc[0,1]
    Tab_KRD = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: KRD Factor'].iloc[0,1]
    Tab_KRD_Result_Col = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: KRD Result Column'].iloc[0,1]
    Tab_Result_Col = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Stress Result Column'].iloc[0,1]

    #Effective duration shock and interest rate stress type parameters
    Val_Date = date(Val_Year,Val_Month,Val_Day)
    Eff_Dur_Shock = Parameters.loc[Parameters.iloc[:,0] == 'Effective Duration Shock'].iloc[0,1]

    #Interest rate stress parameters
    IR_Shock_Indi = Parameters.loc[Parameters.iloc[:,0] == 'Interest Rate Stress Indicator'].iloc[0,1]
    IR_Shock_Sce = Parameters.loc[Parameters.iloc[:,0] == 'Interest Rate Stress Scenario'].iloc[0,1]
    IR_Shock_Type = Parameters.loc[Parameters.iloc[:,0] == 'Interest Rate Stress Type'].iloc[0,1]

    #Credit spread stress parameters
    CS_Shock_Indi = Parameters.loc[Parameters.iloc[:,0] == 'Credit Spread Stress Indicator'].iloc[0,1]
    CS_Shock_Sce = Parameters.loc[Parameters.iloc[:,0] == 'Credit Spread Stress Scenario'].iloc[0,1]
    CS_Shock_TTM_Min = Parameters.loc[Parameters.iloc[:,0] == 'Credit Spread Stress Min TTM'].iloc[0,1]
    CS_Shock_TTM_Max = Parameters.loc[Parameters.iloc[:,0] == 'Credit Spread Stress Max TTM'].iloc[0,1]
    CS_Shock_TTM_Multi = Parameters.loc[Parameters.iloc[:,0] == 'Credit Spread Stress Multiple TTM'].iloc[0,1]

    #Result column name
    Base_Result = list(pd.read_excel(File_Para,sheet_name = Tab_Result_Col)[Tab_Result_Col])
    Shock_Result =  tuple([col + '_Shocked' for col in Base_Result])
    Base_KRD_Result = list(pd.read_excel(File_Para,sheet_name = Tab_KRD_Result_Col)[Tab_KRD_Result_Col])
    KRD_Result = tuple([col + '_Shocked' for col in Base_KRD_Result])

    #Import asset data, cashflows, discount curve, KRD factor and IR & CS shocks
    Spread_Data = pd.read_excel(File_Para, sheet_name = Tab_Spread)
    Cashflows_Data = pd.read_excel(File_Cashflows).loc[:,'Data_Control':]
    Cashflows_Data.drop('Dirty_MV_THB', axis=1, inplace=True)
    DISC = pd.read_excel(File_Para,sheet_name = Tab_Curve)
    Para_IR_Shock = pd.read_excel(File_Para,sheet_name = Tab_Stress_IR)
    Para_CS_Shock = pd.read_excel(File_Para,sheet_name = Tab_Stress_CS)
    PARA_KRD = pd.read_excel(File_Para,sheet_name = Tab_KRD)

    #%% Input transformation

    #Function to classify credit spread shock
    def CS_Shock_Index (row):
        
        if row[Para_CS_Shock.columns[1]] == 'Local':
        
            return row[Para_CS_Shock.columns[0]] +'_'+ row[Para_CS_Shock.columns[1]] +'_'+ row[Para_CS_Shock.columns[2]]
    
        elif row[Para_CS_Shock.columns[1]] == 'US & Developed':
                    
            return row[Para_CS_Shock.columns[0]] +'_'+ row[Para_CS_Shock.columns[1]] +'_'+ row[Para_CS_Shock.columns[2]] +'_'+ row[Para_CS_Shock.columns[3]] +'_'+ row[Para_CS_Shock.columns[4]]
        
        else:
        
            return row[Para_CS_Shock.columns[0]] +'_'+ row[Para_CS_Shock.columns[1]] +'_'+ row[Para_CS_Shock.columns[2]] +'_'+ row[Para_CS_Shock.columns[3]]

    #Transform baseline asset spread data
    Spread_Data = Spread_Data.loc[:,~Spread_Data.columns.str.endswith('_Prev')]
    Spread_Data = Spread_Data.drop(Spread_Data.columns[0], axis=1)
    Spread_Data = Spread_Data.dropna()

    #Merge the cashflows and spread data as one tabl, calculate TTM of each asset and map the credit spread shock index to each asset
    Asset_Data = pd.merge(Cashflows_Data, Spread_Data[['Data_Control','Spread'] + Base_Result], how = 'inner', on = 'Data_Control')
    Asset_Data['TTM'] = round((Asset_Data['Redemp_Year'] - Val_Year + (Asset_Data['Redemp_Month'] - Val_Month) / 12) / CS_Shock_TTM_Multi) * CS_Shock_TTM_Multi
    Asset_Data['TTM'] = Asset_Data['TTM'].clip(CS_Shock_TTM_Min, CS_Shock_TTM_Max).astype(str)
    Asset_Data['CS_Shock_Index'] = Asset_Data.apply(lambda row: CS_Shock_Index(row), axis=1)

    #Map credit spread shock size and calculate shocked credit spread for each asset
    if CS_Shock_Indi == 'Y':
    
        Asset_Data = pd.merge(Asset_Data, Para_CS_Shock[['CS_Shock_Index', CS_Shock_Sce]], how = 'left', on = 'CS_Shock_Index').fillna(0)
        Asset_Data['Spread_Shocked'] = Asset_Data['Spread'] + Asset_Data[CS_Shock_Sce] 

    else:
    
        Asset_Data['Spread_Shocked'] = Asset_Data['Spread']

    Asset_Data = Asset_Data.reindex(columns=[*Asset_Data.columns, *Shock_Result], fill_value=0)

    #%% Discount curve input transformation

    #Discount curve, interest rate shock curve, currency list and and date array generation
    DISC = DISC.loc[DISC['Tenor']<=100]
    IR_Shock = pd.DataFrame(np.arange(1,101), index=range(0, 100), columns = ['Tenor'])
    Currency_List = list(set(Para_IR_Shock['Currency']))
    Currency_Fwd = [curr + '_Fwds' for curr in Currency_List]
    DISC = DISC[['Tenor'] + Currency_Fwd]

    if IR_Shock_Indi == 'Y':
    
        #Calculation loop - interest rate shock interpolation
        for i in range(len(Currency_List)):
            
            Currency = Currency_List[i]
            IR_Shock[Currency] = np.interp(IR_Shock['Tenor'], Para_IR_Shock[Para_IR_Shock.Currency == Currency]['Maturity'], 
                                                 Para_IR_Shock[Para_IR_Shock.Currency == Currency][IR_Shock_Sce])
    
        #Discount curve generation to handle interest rate shock via zero or par curves
        if IR_Shock_Type == 'Zeros' or  IR_Shock_Type == 'Pars':
        
            IR_Base_Fwd = DISC[Currency_Fwd]
            IR_Base_Fwd.columns = IR_Base_Fwd.columns.str.removesuffix('_Fwds')
        
            #Zcb curve generation for all currencies
            IR_Base_Zcb = pd.DataFrame(0, index=range(0, 100), columns = Currency_List)
        
            for i in range(len(Currency_List)):
        
                for j in range(len(IR_Base_Zcb)):
                    
                    if j == 0:
                    
                        IR_Base_Zcb.loc[0,Currency_List[i]] = 1/(1+IR_Base_Fwd.loc[0,Currency_List[i]])
                    
                    else:
                    
                        IR_Base_Zcb.loc[j,Currency_List[i]] = IR_Base_Zcb.loc[j-1,Currency_List[i]]/(1+IR_Base_Fwd.loc[j,Currency_List[i]])
               
            #Shock translation in zero curve basis to zcb curve basis
            if IR_Shock_Type == 'Zeros':
            
                IR_Shocked_Zero =  IR_Base_Zcb.pow(-1/DISC['Tenor'], axis=0)-1 + IR_Shock[Currency_List]        
                IR_Shocked_Zero = IR_Shocked_Zero.clip(0, None)
                IR_Shocked_Zcb = (1+IR_Shocked_Zero).pow(-DISC['Tenor'], axis=0)
    
            #Shock translation in par curve basis to zcb curve basis
            elif IR_Shock_Type == 'Pars':
    
                IR_Shocked_Par = (1-IR_Base_Zcb) / IR_Base_Zcb.cumsum() + IR_Shock[Currency_List]
                IR_Shocked_Par = IR_Shocked_Par.clip(0, None)
                IR_Shocked_Zcb = pd.DataFrame(0, index=range(0, 100), columns=Currency_List)
            
                for i in range(len(Currency_List)):
                
                    for j in range(len(IR_Shocked_Par)):
                    
                        if j == 0:
                        
                            IR_Shocked_Zcb.loc[0,Currency_List[i]] = 1/(1+IR_Shocked_Par.loc[0,Currency_List[i]])
                    
                        else:
                        
                            IR_Shocked_Zcb.loc[j,Currency_List[i]] = (1-IR_Shocked_Zcb.loc[:j-1,Currency_List[i]].cumsum()[j-1]*IR_Shocked_Par.loc[j,Currency_List[i]])/(1+IR_Shocked_Par.loc[j,Currency_List[i]])
        
            #Shock translation from zcb curve basis to forward curve basis for monthly discount RFR curve generation
            IR_Shocked_Fwd = IR_Shocked_Zcb.shift(1) / IR_Shocked_Zcb - 1
            IR_Shocked_Fwd.loc[0,:] = IR_Shocked_Zcb.loc[0,:]**-1 - 1
            IR_Shocked_Fwd = IR_Shocked_Fwd.clip(0, None)
            IR_Shock_Fwd = IR_Shocked_Fwd - IR_Base_Fwd
    
        else:
        
            IR_Shock_Fwd = IR_Shock
    
    else:
        
        IR_Shock_Fwd = pd.DataFrame(0, index=range(0, 100), columns=Currency_List)
        
    #Calculation loop - monthly discount RFR curve preparation    
    for i in range(1,len(DISC.columns)):
    
        urrency = DISC.columns[i][0:3]
    
        try:
        
            DISC[Currency+'_Fwds_M'] = (1 + DISC[Currency+'_Fwds']+IR_Shock_Fwd[Currency]) ** (1/12) - 1
        
        except KeyError:
          
            DISC[Currency+'_Fwds_M'] = (1 + DISC[Currency+'_Fwds']) ** (1/12) - 1
                             
        if i == 1:
        
            Disc_M = pd.merge(pd.DataFrame(np.append(0,np.concatenate([([i]*12) for i in range(1,101)])),columns=['Tenor']), 
                              DISC[['Tenor',Currency+'_Fwds_M']],how='left',on='Tenor').fillna(0)
            Disc_M.insert(1,'Month',Disc_M.index)
        
        else:
             
            Disc_M = pd.merge(Disc_M, DISC[['Tenor',Currency+'_Fwds_M']],how='left',on='Tenor').fillna(0)
                    
        Disc_M[Currency + '_Disc_RFR'] = 1
            
        for j in range(1,len(Disc_M)):
               
            Disc_M.loc[j,Currency +'_Disc_RFR'] = Disc_M.loc[j-1,Currency +'_Disc_RFR']/(1+Disc_M.loc[j,Currency+'_Fwds_M'])
                                                         
    Disc_M.to_numpy()

    #%% Calculation loop

    #Calculation loop - PV, duration, convexity
    for i in range(len(Asset_Data)):
    
        Asset_Input = Asset_Data.iloc[i]
    
        if  Asset_Input['Effective_Duration'] != 0:
        
            Cashflows_Input = Asset_Data.loc[i,0:1200].T
            Disc_RFR = Disc_M[Asset_Input['Currency']+'_Disc_RFR']
           
            #Discount curve and PV calculation
            Disc_Base = (Disc_RFR ** (-12/Disc_M['Month']) + Asset_Input['Spread_Shocked']) ** (-Disc_M['Month']/12)
            Disc_Up = (Disc_Base ** (-12/Disc_M['Month']) + Eff_Dur_Shock) ** (-Disc_M['Month']/12)
            Disc_Dn = (Disc_Base ** (-12/Disc_M['Month']) - Eff_Dur_Shock) ** (-Disc_M['Month']/12)
            PV_Base = sum(Cashflows_Input * Disc_Base)
            PV_Up = sum(Cashflows_Input * Disc_Up)
            PV_Dn = sum(Cashflows_Input * Disc_Dn)
        
            #Store result to asset dataframe
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Dirty_MV_THB_Shocked')] = PV_Base
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Effective_Duration_Shocked')] = (PV_Dn - PV_Up)/(2 * PV_Base * Eff_Dur_Shock)
            Asset_Data.iloc[i,Asset_Data.columns.get_loc('Effective_Convexity_Shocked')] = (PV_Dn + PV_Up - 2 * PV_Base)/(PV_Base * (Eff_Dur_Shock ** 2))
            for k in range(len(KRD_Result)):
                Asset_Data.iloc[i,Asset_Data.columns.get_loc(KRD_Result[k])] = sum(Cashflows_Input * (Disc_Dn - Disc_Up) * PARA_KRD[Base_KRD_Result[k]] / 
                                                                                   (2 * PV_Base * Eff_Dur_Shock))     
            
        else:
               
            continue

    Asset_Data['Dirty_MV_THB_Shocked'] = np.where(Asset_Data['Dirty_MV_THB_Shocked'] == 0, Asset_Data['Dirty_MV_THB'], Asset_Data['Dirty_MV_THB_Shocked'])

    #%% Export output

    if CS_Shock_Indi == 'Y':
    
        Asset_Data.drop(CS_Shock_Sce, axis=1, inplace=True)

    Asset_Data.drop(range(0,1201), axis=1, inplace=True)
    Asset_Data.index = np.arange(1, len(Asset_Data)+1)
    Asset_Data.to_excel(output_path, header=True)
    print("--- %s seconds ---" % (time.time() - start_time))
    print('Asset model stress testing ran sucessfully.')
    return output_path

