# -*- coding: utf-8 -*-
"""
Created on Mon Feb 13 9:00:00 2024

@author: Thaniie
"""       
#%% Import library and parameters from command prompt
import sys
import ctypes
import time
import numpy as np
import pandas as pd
from datetime import datetime, date
import os
import openpyxl

def run_Asset_Cashflows(input_Path, output_directory, date_time_str):


    start_time = time.time()        
    #Input parameters from VBA
    #File_Para = r'C:\Python_Intern\Asset_Movement_Excel\Asset_Movement_Jan2024_v11.xlsm'
    File_Para = input_Path
    excel = pd.ExcelFile(File_Para)
    Tab_Para = 'Python_Para_CF'
    Parameters = pd.read_excel(excel, sheet_name = Tab_Para, engine='openpyxl')
    Value_Date = datetime.strptime(date_time_str, "%Y%m%d")
    rfr_file_name = "Asset_Cashflows_" + Value_Date.strftime("%b %Y") + ".xlsx"

    

    #Valuation dates parameters
    Val_Year = Parameters.loc[Parameters.iloc[:,0] == 'Year'].iloc[0,1]
    Val_Month = Parameters.loc[Parameters.iloc[:,0] == 'Month'].iloc[0,1]
    Val_Day = Parameters.loc[Parameters.iloc[:,0] == 'Date'].iloc[0,1]

    #Input and output paths parameters
    File_Asset = Parameters.loc[Parameters.iloc[:,0] == 'Input Full Path: Asset'].iloc[0,1]
    File_Redem = Parameters.loc[Parameters.iloc[:,0] == 'Input Full Path: Redemption Schedule'].iloc[0,1]
    File_Step = Parameters.loc[Parameters.iloc[:,0] == 'Input Full Path: Step Schedule'].iloc[0,1]
    File_Sink = Parameters.loc[Parameters.iloc[:,0] == 'Input Full Path: Sinking Schedule'].iloc[0,1]
    #File_Output = Parameters.loc[Parameters.iloc[:,0] == 'Output Full Path'].iloc[0,1]
    output_path = os.path.join(output_directory, rfr_file_name)
    #output_path = output_directory

    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
        
    #Input and output paths parameters
    Tab_FX = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: FX Rate'].iloc[0,1]
    Tab_Curve = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Discount Curve'].iloc[0,1]
    Tab_Eff_Date = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Effective Date'].iloc[0,1]
    Tab_Field_Col = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Field Column'].iloc[0,1]
    Tab_Product = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Product Type'].iloc[0,1]
    Tab_COR = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: COR Type'].iloc[0,1]
    Tab_Rating = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: Rating Type'].iloc[0,1]
    Tab_IRS = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: IRS Exchange Notional'].iloc[0,1]
    Tab_FAIR2 = Parameters.loc[Parameters.iloc[:,0] == 'Input Tab: FAIR2'].iloc[0,1]


    #Valuation dates, effective duration shock, FX rate tables and MV indicator setting
    Val_Date = date(Val_Year,Val_Month,Val_Day)
    FX_Rate = pd.read_excel(File_Para,sheet_name = Tab_FX, names = ['Date','Index', 'Currency','Currency_Full','FX/USD','USD/FX'])
    FX_Rate['Date'] = pd.to_datetime(FX_Rate['Date']).dt.date

    #.csv input encoding and redemption/sinking/stepped-up coupon flags setting
    Input_Encoding = Parameters.loc[Parameters.iloc[:,0] == 'Input Encoding'].iloc[0,1]
    CF_Redemp_Flag = Parameters.loc[Parameters.iloc[:,0] == 'Cashflows: Redemption Flag'].iloc[0,1]
    CF_Step_Flag = Parameters.loc[Parameters.iloc[:,0] == 'Cashflows: Step Flag'].iloc[0,1]
    CF_Sink_Flag = Parameters.loc[Parameters.iloc[:,0] == 'Cashflows: Sinking Flag'].iloc[0,1]

    #Row and column filter parameters including COR, credit rating, and FAIR2 fund
    Filter = list(pd.read_excel(File_Para,sheet_name = Tab_Field_Col)[Tab_Field_Col])
    Asset_Type = pd.read_excel(File_Para,sheet_name = Tab_Product)
    Product_Type = tuple(list(pd.read_excel(File_Para,sheet_name = Tab_Product)[Tab_Product]))
    COR_Type = pd.read_excel(File_Para,sheet_name = Tab_COR)
    Rating_Type = pd.read_excel(File_Para,sheet_name = Tab_Rating)
    IRS_Ex_Notional = tuple(list(pd.read_excel(File_Para,sheet_name = Tab_IRS)[Tab_IRS]))
    FAIR2 = pd.read_excel(File_Para,sheet_name = Tab_FAIR2)

    #Import asset data, schedule files, discount curves and asset effective date
    try:
        PALS = pd.read_csv(File_Asset, sep = ';', encoding = Input_Encoding)
        REDEM = pd.read_csv(File_Redem, sep = ';', encoding = Input_Encoding)
        STEP = pd.read_csv(File_Step, sep = ';', encoding = Input_Encoding)
        SINK = pd.read_csv(File_Sink, sep = ';', encoding = Input_Encoding)
 
    except FileNotFoundError as File_Input_Error:
        ctypes.windll.user32.MessageBoxW(0, "Asset/Schedule File: "+repr(File_Input_Error), repr(FileNotFoundError.__qualname__), 0)
        sys.exit(1)

    DISC = pd.read_excel(File_Para,sheet_name = Tab_Curve)
    PARA_EFF_DATE = pd.read_excel(File_Para,sheet_name = Tab_Eff_Date)

    #FX rate table transformation
    FX_Table = FX_Rate[FX_Rate['Date'] == Val_Date]
    FX_Table['Currency'] = FX_Table['Currency'].str.strip()
    
   

    #Function to classifiy fund, asset type, credit type, COR type, rating type, and swap type
    def Credit_Con (row):
    
        if row['Credit_Type_EC'] == 'GOV' == 'GOV':
        
            return 'GOV'
    
        else:
        
            return 'CORP'

    def Swap_Con (row):
    
        if row['Effective_Year'] > Val_Year or (row['Effective_Year'] == Val_Year and row['Effective_Month'] > Val_Month):
        
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

    PALS = pd.DataFrame(PALS, columns = Filter)
    PALS = PALS[PALS['Product_Type'].isin(Product_Type)].copy()
    PALS = PALS.merge(FAIR2, how='left', on = FAIR2.columns[0])
    PALS[Tab_FAIR2] = PALS.apply (lambda row: FAIR2_Con(row), axis=1)
    PALS['Data_Control'] = PALS['Security_ID'] +'_'+ PALS['Currency'] +'_'+ PALS['Pay_Receive'] + '_' + PALS['Trade_Date'].astype(str)

    for i in range(len(Fund)):
    
        PALS[Fund[i]] = PALS[PALS['FAIR2'] == Fund[i]].groupby('Data_Control').Par_Value.transform('sum')
        PALS[Fund[i]] = PALS.groupby('Data_Control')[Fund[i]].transform('sum') / PALS.groupby('Data_Control')[Fund[i]].transform('count')
   
    PALS['Par_Value'] = PALS.groupby('Data_Control').Par_Value.transform('sum')
    PALS[Fund] = PALS[Fund].div(PALS['Par_Value'], axis=0)
    PALS['Dirty_MV_RC'] = PALS.groupby('Data_Control').Dirty_MV_RC.transform('sum')
    PALS['Accrued_Interest_RC'] = PALS.groupby('Data_Control').Accrued_Interest_RC.transform('sum')
    PALS = PALS.drop_duplicates(subset='Data_Control').reset_index(drop=True).fillna(0)

    Asset_Data = PALS.merge(PARA_EFF_DATE[['Security_ID','Effective_Year']], how='left', on='Security_ID').fillna({'Effective_Year':1900})
    Asset_Data = Asset_Data.merge(PARA_EFF_DATE[['Security_ID','Effective_Month']], how='left', on='Security_ID').fillna({'Effective_Month':1})
    Asset_Data = Asset_Data.drop_duplicates(subset='Data_Control').reset_index(drop=True)
    Asset_Data = Asset_Data.merge(FX_Table[['Currency','FX/USD']], how='left', on='Currency', copy=False)

    #Create new columns for asset category, MV in THB and par value in THB
    Asset_Data['Dirty_MV_THB'] = Asset_Data['Dirty_MV_RC'] * FX_Table.loc[FX_Table.loc[:,'Currency'] == 'THB']['FX/USD'].iloc[0]
    Asset_Data['Par_Value_THB'] = Asset_Data['Par_Value'] * FX_Table.loc[FX_Table.loc[:,'Currency'] == 'THB']['FX/USD'].iloc[0] / Asset_Data['FX/USD']
    Asset_Data['Accrued_Interest_THB'] = Asset_Data['Accrued_Interest_RC'] * FX_Table.loc[FX_Table.loc[:,'Currency'] == 'THB']['FX/USD'].iloc[0]

    #Create FAIR2 columns and asset type columns
    Asset_Data = Asset_Data.merge(Asset_Type, how='left', on = Asset_Type.columns[0])

    #Create credit type, COR, Rating, Swap_Type and drop column duplicates
    Asset_Data['Credit_Type'] = Asset_Data.apply (lambda row: Credit_Con(row), axis=1)
    Asset_Data = Asset_Data.merge(COR_Type, how='left', on = COR_Type.columns[0]).fillna({COR_Type.columns[1]:'N/A'})
    Asset_Data = Asset_Data.merge(Rating_Type, how='left', on = Rating_Type.columns[0]).fillna({Rating_Type.columns[1]:'N/A'})
    Asset_Data['Swap_Type'] = Asset_Data.apply (lambda row: Swap_Con(row), axis=1)

    #Placeholder columns for calculation results, drop duplicate rows, drop ILP data and reset dataframe index 
    Asset_Data = Asset_Data.loc[:,~Asset_Data.columns.duplicated()]
    Asset_Data = Asset_Data.drop_duplicates()
    Asset_Data = Asset_Data[Asset_Data.FAIR2 != 'ILP']
    Asset_Data.reset_index(drop=True, inplace=True)
    Asset_Data = pd.concat([Asset_Data, pd.DataFrame(0, index=range(len(Asset_Data)), columns=range(0,1201))], axis=1)

    #Schedule files transformation
    REDEM = REDEM[REDEM['INDEX'].isin(Asset_Data['Security_ID'])]
    STEP = STEP[STEP['INDEX'].isin(Asset_Data['Security_ID'])]
    SINK = SINK[SINK['INDEX'].isin(Asset_Data['Security_ID'])]


#Discount curve and date array generation
    DISC = DISC.loc[DISC['Tenor']<=100]

    DATE = pd.DataFrame(np.concatenate([([i]) for i in range(0,1201)]),index=range(1201),columns=['Index'])
    DATE = DATE.assign(Year=pd.date_range(Val_Date,periods = 1201, freq='M').year)
    DATE = DATE.assign(Month=pd.date_range(Val_Date,periods = 1201, freq='M').month)
    Year = pd.date_range(Val_Date,periods = 1201, freq='M').year.to_numpy()
    Month = pd.date_range(Val_Date,periods = 1201, freq='M').month.to_numpy()

    #Calculation loop - discount RFR curve preparation
    for i in range(1,len(DISC.columns)):
    
        Currency = DISC.columns[i][0:3]
    
        if Currency == 'Unn':
        
            continue
    
        else:
                    
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


    for i in range(len(Asset_Data)):
 
        Asset_Input = Asset_Data.iloc[i]
            
        #Assign discount curves based on currency and project cashflows as of current valuation date
        Disc_RFR = Disc_M[Asset_Input['Currency']+'_Disc_RFR']
        Redemp_Output = Redemption_Calc(Year, Month, Val_Date, Asset_Input['Security_ID'], Asset_Input['Product_Type'], Asset_Input['Redemp_Year'], 
                                        Asset_Input['Redemp_Month'], Asset_Input['Redemp_Month'], Asset_Input['Effective_Year'],
                                        Asset_Input['Effective_Month'], Asset_Input['Par_Value_THB'], Asset_Input['Par_Value_THB'], 
                                        Asset_Input['Coupon_Rate'], Asset_Input['Coupon_Freq'], Asset_Input['Floating_Spread'], 
                                        DATE, REDEM,Asset_Input['Embedded_Opt_Index'], Disc_RFR, CF_Redemp_Flag)
        Sink_Output = Sink_Calc(SINK, Asset_Input['Security_ID'], Asset_Input['Par_Value_THB'], Redemp_Output[0], 
                            Redemp_Output[3], CF_Sink_Flag)
        Step_Output = Step_Coupon_Calc(STEP, Asset_Input['Security_ID'], Asset_Input['Coupon_Rate'], CF_Step_Flag)
    
        Cashflows = Cashflows_Calc(Year, Month, Val_Date, Asset_Input['Security_ID'], Asset_Input['Product_Type'], Redemp_Output[1], 
                                   Asset_Input['Redemp_Month'], Redemp_Output[2], Asset_Input['Effective_Year'], Asset_Input['Effective_Month'], 
                                   Sink_Output[0], Sink_Output[1], Step_Output, Asset_Input['Coupon_Freq'], Asset_Input['Floating_Spread']) + Sink_Output[2]
    
    #Calculate and store calculated cashflows
        Asset_Data.loc[i,0:] = Cashflows
        
    

    Asset_Data.drop(['Company_Code', 'FAIR2','FX/USD'], inplace=True, axis=1)
    Asset_Data.index = np.arange(1, len(Asset_Data)+1)
    Asset_Data.to_excel(output_path, header=True, engine='openpyxl')
    print("--- %s seconds ---" % (time.time() - start_time))
    print("Asset model cashflows ran sucessfully.")
    return output_path#"Asset model cashflows ran sucessfully."
