import numpy as np
import pandas as pd
from scipy.stats import linregress
from scipy.stats.mstats import gmean
from datetime import datetime
import sys
import gc
import openpyxl
import os
from openpyxl import worksheet

def run_calibration(input_path, output_directory, date_time_str):


    filepath = input_path
    excel = pd.ExcelFile(filepath)
    tab_para = 'Python_Para_Curve'
    parameters = pd.read_excel(excel, sheet_name=tab_para)
    filename = parameters.loc[parameters.iloc[:,0] == 'Output Full Path'].iloc[0,1]
    Val_Date = datetime.strptime(date_time_str, "%Y%m%d")
    rfr_file_name = "RFR_" + Val_Date.strftime("%b %Y") + ".xlsx"
    
    tab_input = parameters.loc[parameters.iloc[:,0] == 'Input Tab: Curve Calibration'].iloc[0,1]
    tab_template = parameters.loc[parameters.iloc[:,0] == 'Input Tab: Curve Template'].iloc[0,1]
    curve_parameters = pd.read_excel(excel, sheet_name=tab_input, engine='openpyxl', index_col=0)
    curve_temp = pd.read_excel(excel, sheet_name=tab_template, engine='openpyxl', index_col=0)
    #output_path = os.path.join(output_directory, filename)
    output_path = os.path.join(output_directory, rfr_file_name)
    

    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
        

    #z = np.arange(1,10)

    def hh(z):
        return (z + np.exp(-z)) / 2

    def Hmat(u, v):
        return hh(u + v) - hh(np.abs(u-v))


    
        
    def SmithWilsonBruteForce(instrument, data_in, nrofcoup, CRA, UFRac, alfamin, Tau, T2):
    
        data = data_in
        nrofrates = int(np.sum(data[:,0])) #sum first coulumn of data to determine number of liquid input rates
    
        u = np.zeros(nrofrates) # row vector containg liquid maturities
        r = np.zeros(nrofrates) # row vector containg liquid rates
    
        j = 0
        for i in range(1, data.shape[0] + 1):
            if data[i-1, 0] == 1: # Indicator = 1 <-> liquid maturity/rate 
                j = j+1
                u[j-1] = data[i-1,1] # store liquid maturity in u-vector 
                r[j-1] = data[i-1,2] - CRA / 10000 # store liquid rate including cra in r-vector
        umax = np.max(u) # maximum liquid maturity
    
        # '''Note prices of all instruments are set to 1 by construction
        # Hence 1: if Instrument = Zero then for Zero i there is only one pay-off of (1+r(i))^u(i) at time u(i)
        # Hence 2: if Instrumnent = Swap or Bond then for Swap/Bond i there are pay-offs of r[i]/nrofcoup at time 1/nrofcoup, 2/nrofcoup, ... u[i] - 1/nrofcoup plus a final pay-off of 1+r[i]/nrofcoup at time u[i]'''
          
        lnUFR = np.log(1 + UFRac)
        if instrument == "Zero":
            nrofcoup = 1
        Q = np.zeros((int(nrofrates),int(umax * nrofcoup)))
        if instrument == 'Zero': #nrofcoup = 1 by definition
            for i in range(1, nrofrates+1):
                Q[i-1][int(u[i-1])-1] = np.exp(-lnUFR * u[i-1]) * ((1+r[i-1]) ** u[i-1])
        elif (instrument == 'Swap') | (instrument == 'Bond'):
            for i in range(1, nrofrates+1):
                for j in range(1, int(u[i-1]) * nrofcoup - 1 + 1):
                    Q[i-1][j-1] = np.exp(-lnUFR * j / nrofcoup) * r[i-1] / nrofcoup
                j = int(u[i-1]) * nrofcoup
                Q[i-1][j-1] = np.exp(-lnUFR * j / nrofcoup) * (1 + r[i-1] / nrofcoup)
    
        precision = 6 # number decimal for optimal alfa
    
        Tau = Tau / 10000 # As Inout is in basispoint it has to be transformed to number format
        galfa_output =  Galfa(alfamin, Q, nrofrates, umax, nrofcoup, T2, Tau)
        # This function calculates 2 outputs:
        # Output(1): g(alfa)-tau where g(alfa) is according to 165 of the specs
        # Output(2): gamma = Qb accordimg to 148 of the specs
        if galfa_output[0] <= 0:
            alfa = alfamin
            gamma = galfa_output[1]
        else: # scanning for the optimal alfa is based on the scan-procedure taken from
            stepsize = 0.1
            alfa = alfamin + stepsize
            for alfa in np.arange(alfamin + stepsize, 20 + 0.001, stepsize):
                if Galfa(alfa, Q, nrofrates, umax, nrofcoup, T2, Tau)[0] <= 0:
                    break
            for i in range(1,precision - 1 + 1):
                alfascanoutput = AlfaScan(alfa, stepsize, Q, nrofrates, umax, nrofcoup, T2, Tau) 
                alfa = alfascanoutput[0]
                stepsize = stepsize / 10
            gamma = alfascanoutput[1]
    
        # Optimal alfa and corresponding gamma have been determined 
        # Now the SW-present value function according to 154 of the specs can be calculated :p(v) = exp
    
        h = np.zeros((122, int(umax * nrofcoup)))
        g = np.zeros((122, int(umax * nrofcoup)))
    
        for i in range(122):
            for j in range(1, int(umax * nrofcoup) + 1):
                h[i][j-1] = Hmat(alfa * i, alfa * (j) / nrofcoup)
                if (j / nrofcoup) > i:
                    g[i][j-1] = alfa * (1 - np.exp(-alfa * (j) / nrofcoup) * np.cosh(alfa * i))
                else:
                    g[i][j-1] = alfa * np.exp(-alfa * i) * np.sinh(alfa * (j) / nrofcoup)
    
        tempdiscount = np.zeros(122)
        tempintensity = np.zeros(122)
    
        discount = np.zeros(122)
        fwintensity = np.zeros(122)
        yldintensity = np.zeros(122)
        forwardac = np.zeros(122)
        zeroac = np.zeros(122)
        forwardcc = np.zeros(122)
        zerocc = np.zeros(122)
    
        temptempdiscount = np.transpose(np.matmul(h, gamma)).squeeze()
        temptempintensity = np.transpose(np.matmul(g, gamma)).squeeze()
        for i in range(122):
            tempdiscount[i] = temptempdiscount[i]  
            tempintensity[i] = temptempintensity[i]
        
        temp = 0
        for i in range(1,int(umax * nrofcoup) + 1):
            temp = temp + (1 - np.exp(-alfa * i / nrofcoup)) * gamma[i-1,0]
    
        yldintensity[0] = lnUFR - alfa * temp
        fwintensity[0] = yldintensity[0]
        discount[0] = 1
        yldintensity[1] = lnUFR - np.log(1+tempdiscount[1])
        fwintensity[1] = lnUFR - tempintensity[1] / (1 + tempdiscount[1])
        discount[1] = np.exp(-lnUFR) * (1 + tempdiscount[1])
    
        zeroac[1] = 1 / discount[1] - 1 
        forwardac[1] = zeroac[1]
        for i in range(2, 120 + 1):
            yldintensity[i] = lnUFR - np.log(1 + tempdiscount[i]) / i
            fwintensity[i] = lnUFR - tempintensity[i] / (1 + tempdiscount[i])
            discount[i] = np.exp(-lnUFR * i) * (1 + tempdiscount[i])
            zeroac[i] = (1 / discount[i]) ** (1 / i) - 1
            forwardac[i] = discount[i - 1] / discount[i] - 1

        yldintensity[121] = 0
        fwintensity[121] = 0
        zeroac[121] = 0
        forwardac[121] = 0
        discount[121] = alfa

        for i in range(1, 120 + 1):
            forwardcc[i] = np.log(1 + forwardac[i])
            zerocc[i] = np.log(1 + zeroac[i])

        output = [0,0,0,0,0,0]
        output[0] = discount
        output[1] = yldintensity
        output[2] = zeroac
        output[3] = fwintensity
        output[4] = forwardcc
        output[5] = forwardac

        SmithWilsonBruteForce = output
        
        return SmithWilsonBruteForce
    
        

    def AlfaScan(lastalfa, stepsize, Q, mm, umax, nrofcoup, T2, Tau):
        for alfa in np.arange(lastalfa + stepsize / 10 - stepsize, lastalfa + 0.00000001, stepsize / 10):
            galfa_output = Galfa(alfa, Q, mm, umax, nrofcoup, T2, Tau)
            if galfa_output[0] <= 0:
                break
            
        output = [0,0]
    
        output[0] = alfa
        output[1] = galfa_output[1]
        AlfaScan = output
    
        return AlfaScan

    def Galfa(alfa, Q, mm, umax, nrofcoup, T2, Tau):

        h = np.zeros([int(umax*nrofcoup), int(umax*nrofcoup)])
        for i in range(1, int(umax * nrofcoup) + 1):
            for j in range(1, int(umax * nrofcoup) + 1):
                h[i-1,j-1] = Hmat(alfa * i / nrofcoup, alfa * j / nrofcoup)
            
        temp1 = np.zeros([mm, 1])
        for i in range(1, mm+1):
            temp1[i-1, 0] = 1 - np.sum(Q[i-1])
    
        b = np.matmul(np.linalg.inv(np.matmul(np.matmul(Q, h), np.transpose(Q))), temp1)

        gamma = np.matmul(np.transpose(Q), b)
        temp2 = 0
        temp3 = 0
        for i in range(1, int(umax * nrofcoup) + 1):
            temp2 = temp2 + gamma[i-1, 0] * i / nrofcoup
            temp3 = temp3 + gamma[i-1, 0] * np.sinh(alfa * i / nrofcoup)
    
        kappa = (1 + alfa * temp2) / temp3
    
        output = [0,0]
        output[0] = alfa / np.abs(1 - kappa * np.exp(T2 * alfa)) - Tau
        output[1] = gamma
        Galfa = output

        return Galfa

 

    calibration_para = curve_parameters.loc[:'Convergence Bandwidth']
    calibration_para_col = list(calibration_para.columns)
       
    Fwds = curve_temp.add_suffix('_Fwds')
    Zeros = curve_temp.add_suffix('_Zeros')
    ZCBs = curve_temp.add_suffix('_ZCBs')
    Pars = curve_temp.add_suffix('_Pars')

    for column in calibration_para_col:
    
        #Set parameters for each loop
        Currency = column
        LOP_year = curve_parameters [column]['LOP Year']
        instrument = curve_parameters [column]['Asset Type']
        nrofcoup = 0 if curve_parameters [column]['Asset Type'] == 'Zero' else curve_parameters[column]['Coupon Freq']  
        CRA = 0
        UFRac = curve_parameters [column]['UFR']
        alfamin = curve_parameters [column]['Min Alpha']
        Tau = curve_parameters [column]['Convergence Bandwidth']
        T2 = curve_parameters [column]['Convergence Year']
    
        #Create the data for each currency
        data = pd.DataFrame()
        data[column] = curve_parameters.loc[1:][column].copy()
        data['Time'] = data.index
        data['Maturity'] = np.where(data[column] != 0 ,  1,0) 
        data = data[[ 'Maturity','Time', column]].to_numpy()
    
        #First output of SmithWilson(In numpy array)
        output_step_1 = SmithWilsonBruteForce(instrument, data, nrofcoup, CRA, UFRac, alfamin, Tau, T2)
    
        #Change output_step_1 to be in dataframe format
        columns =["Discount factor","Spot intensityspot rate cc","Spot rate ac","Forward intensity","Forward rate cc", "Forward rate ac"]
        j = 0
        output_df_1 = pd.DataFrame()
        for col in columns:
            output_df_1[col] = output_step_1[j]
            j += 1
        output_df_1['Time'] = output_df_1.index
    
        #Take output step one to calculate alpha, slope, and intercept to calibrate forward and zero rate
        y = output_step_1[-1][1:LOP_year+1]
        x = np.arange(1, len(y)+1)
        optimal_alpha = output_step_1[0][-1]
        fitted_slope, fitted_intercept, _, _, _ = linregress(x, y)
    
        translate = pd.DataFrame(np.concatenate([np.zeros(1),np.ones(30), np.zeros(20)])\
                            ,columns = ['Liquid (1=Y/0=N)'])
        translate['Maturity'] = np.concatenate([np.zeros(1),np.arange(1,51)])
        translate['Time'] = translate.index
        translate = translate.drop(0)
        translate['Forward rate ac (LR)'] = (translate['Time'] <= LOP_year)* (fitted_intercept + fitted_slope * translate['Time'])
        translate['Spot rate ac (LR)'] = (translate['Time'] <= LOP_year) * translate['Forward rate ac (LR)'].expanding().apply(lambda x: gmean(1+x)-1)
        translate['Input rates at calculation date'] = translate['Spot rate ac (LR)']
    
        #Get the second output to calculate SmithWilson in 2 second time 
        new_input = translate[['Liquid (1=Y/0=N)','Maturity', 'Input rates at calculation date']]
        new_input = new_input.to_numpy()
        # Second output of SmithWilson(In numpy array)
        output_step_2 = SmithWilsonBruteForce("Zero", new_input, 0, 0, UFRac, alfamin, Tau, T2)
    
        # Change output_step_2 to be in dataframe format
        columns =["Discount factor","Spot intensityspot rate cc","Spot rate ac","Forward intensity","Forward rate cc", "Forward rate ac"]
        j = 0
        output_df_2 = pd.DataFrame()
        for col in columns:
            output_df_2[col] = output_step_2[j]
            j += 1
        output_df_2['Time'] = output_df_2.index
    
        #Get the yield curve output in each columns (Fwds, Zeros, ZCBs, and Pars)
        output = pd.DataFrame()
        output['Forward rate ac 1'] = output_df_1['Forward rate ac'].drop(0)
        output['Forward rate ac 2'] = output_df_2['Forward rate ac'].drop(0)
        output['Time'] = output.index
        output = output[:-1]
        output['Fwds'] = np.where(output['Time'] <= LOP_year, output['Forward rate ac 1'],output['Forward rate ac 2'])
    
        output = output[['Time','Fwds']]
        output['Zeros'] = 0
        for i in range(1, len(output)+1):
            if i == 1:
                output.loc[i, 'Zeros'] = output.loc[i, 'Fwds']
            else:
                output.loc[i, 'Zeros'] = ((1+output.loc[i-1, 'Zeros'])**(output.loc[i-1,'Time'])*(1+output.loc[i,'Fwds']))**(1/output.loc[i,'Time'])-1
        output['ZCBs'] = 1/(1+output['Zeros'])**output['Time']
        output['Pars'] = (1-output['ZCBs']) / output['ZCBs'].cumsum()
    
        #Assign each columns to be in each dataframe set
        Fwds[f'{column}_Fwds'] = output['Fwds']
        Zeros[f'{column}_Zeros'] = output['Zeros']
        ZCBs[f'{column}_ZCBs'] = output['ZCBs']
        Pars[f'{column}_Pars'] = output['Pars']

    gc.collect()

    with pd.ExcelWriter(output_path) as writer:
        Fwds.to_excel(writer, 'Fwds')
        Zeros.to_excel(writer, 'Zeros')
        ZCBs.to_excel(writer, 'ZCBs')
        Pars.to_excel(writer, 'Pars')

    writer.save()
    return output_path#"RFR curve calibration ran successfully."

  
