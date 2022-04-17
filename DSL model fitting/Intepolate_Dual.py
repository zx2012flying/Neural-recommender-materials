# %% import package

import numpy as np
import pandas as pd
from sklearn.metrics import r2_score
from lmfit import Parameters, minimize

# %% LS regression

fit_params = Parameters()
fit_params.add('Q1', value = 10, min = 0.01, max = 10)
fit_params.add('b1', value = 10, min = 0.01, max = 1000)
fit_params.add('Q2', value = 10, min = 0.01, max = 10)
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

# %% fit R2 = 0.99      8606
       # R2 = 0.98      9411
       # R2 = 0.97      9855
       # R2 = 0.95      10388
       # R2 = 0.9       11078

Pressure_0_1 = []
Load_0_1 = []
Pressure_1 = []
Load_1 = []
Pressure_35 = []
Load_35 = []

s = 0

for i in range(0, len(Pressure)):
    X = Pressure[i, :]
    Y = Load[i, :]
    X = X[~np.isnan(X)]
    Y = Y[~np.isnan(Y)]
    
    out = minimize(Dual_Langmuir, fit_params, args = (X, ), kws =  {'data': Y})
    Y_predict = Dual_Langmuir(out.params, X)    
    R2 = r2_score(Y, Y_predict)

    if R2 >= 0.95:
        Pressure_0_1.append(0.1)
        Load_0_1.append(Dual_Langmuir(out.params, 0.1))
        Pressure_1.append(1)
        Load_1.append(Dual_Langmuir(out.params, 1))
        Pressure_35.append(35)
        Load_35.append(Dual_Langmuir(out.params, 35))
    else:
        Pressure_0_1.append(np.nan)
        Load_0_1.append(np.nan)
        Pressure_1.append(np.nan)
        Load_1.append(np.nan)
        Pressure_35.append(np.nan)
        Load_35.append(np.nan)

del X, Y, Y_predict

# %% add to oroginal

Add_pressure = Original_Pressure.iloc[:, 0:7]
Add_load = Original_Load.iloc[:, 0:7]

Add_pressure.iloc[:, 4] = Pressure_0_1
Add_pressure.iloc[:, 5] = Pressure_1
Add_pressure.iloc[:, 6] = Pressure_35

Add_load.iloc[:, 4] = Load_0_1
Add_load.iloc[:, 5] = Load_1
Add_load.iloc[:, 6] = Load_35

# %% save to excel

Writer = pd.ExcelWriter('4_inter_Isotherm_pressure.xlsx', engine = 'xlsxwriter')
Add_pressure.to_excel(Writer, 'Sheet1', header = None, index = None)
Writer.save()

WWriter = pd.ExcelWriter('4_inter_Isotherm_loading.xlsx', engine = 'xlsxwriter')
Add_load.to_excel(WWriter, 'Sheet1', header = None, index = None)
WWriter.save()










        
        
