# %% import package

import numpy as np
import pandas as pd
from sklearn.metrics import r2_score
from lmfit import Parameters, minimize
import matplotlib.pyplot as plt

# %% LS regression

fit_params = Parameters()
fit_params.add('Q1', value = 10, min = 0.01, max = 1000)
fit_params.add('b1', value = 10, min = 0.01, max = 1000)
fit_params.add('Q2', value = 10, min = 0.01, max = 1000)
fit_params.add('b2', value = 10, min = 0.01, max = 1000)

def Dual_Langmuir(params, x, data = None):
    Q1 = params['Q1'].value
    b1 = params['b1'].value
    Q2 = params['Q2'].value
    b2 = params['b2'].value
    
    model = Q1 * x / (1 + b1 * x) + Q2 * x / (1 + b2 * x)
    if data is None:
        return model
    return model - data

# %% import pressure data

file = '3_output_Isotherm_pressure.xlsx'
Databook = pd.ExcelFile(file)
Original_Pressure = Databook.parse('Sheet1', header = 0, index_col = None)

Pressure = Original_Pressure.iloc[:, 4:]
Pressure = np.asarray(Pressure)

del file, Databook

# %% import loading data

file = '3_output_Isotherm_loading.xlsx'
Databook = pd.ExcelFile(file)
Original_Load = Databook.parse('Sheet1', header = 0, index_col = None)

Load = Original_Load.iloc[:, 4:]
Load = np.asarray(Load)

del file, Databook

# %% import already generated 4_inter_Isotherm_pressure.xlsx

Pressure_inter = pd.read_excel('4_inter_Isotherm_pressure.xlsx', header = None, index_col = None)
Loading_inter = pd.read_excel('4_inter_Isotherm_loading.xlsx', header = None, index_col = None)

# %% fit
# run with one of the following four coefficients one-by-one
coeff = 10
# coeff = 100
# coeff = 1000
# coeff = 10000

for i in range(0, len(Pressure)):
    if pd.isnull(Pressure_inter.iloc[i, 4]) == True:
        X = Pressure[i, :]
        Y = Load[i, :]
        X = X[~np.isnan(X)]
        Y = Y[~np.isnan(Y)]       
        X = X * coeff       
        out = minimize(Dual_Langmuir, fit_params, args = (X, ), kws =  {'data': Y})
        Y_predict = Dual_Langmuir(out.params, X)    
        R2 = r2_score(Y, Y_predict)

        if R2 >= 0.95:
            Pressure_inter.iloc[i, 4] = 0.1
            Pressure_inter.iloc[i, 5] = 1
            Pressure_inter.iloc[i, 6] = 35
            Loading_inter.iloc[i, 4] = Dual_Langmuir(out.params, 0.1 * coeff)
            Loading_inter.iloc[i, 5] = Dual_Langmuir(out.params, 1 * coeff)
            Loading_inter.iloc[i, 6] = Dual_Langmuir(out.params, 35 * coeff)
            
# %% save to excel

Writer = pd.ExcelWriter('4_inter_Isotherm_pressure.xlsx', engine = 'xlsxwriter')
Pressure_inter.to_excel(Writer, 'Sheet1', header = None, index = None)
Writer.save()

WWriter = pd.ExcelWriter('4_inter_Isotherm_loading.xlsx', engine = 'xlsxwriter')
Loading_inter.to_excel(WWriter, 'Sheet1', header = None, index = None)
WWriter.save()










        
        