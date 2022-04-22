# import package

import pandas as pd
import tensorflow as tf
import numpy as np

# %% import model 

model = tf.keras.models.load_model('Final_Neural_Recommender_System.h5')

# %% import matrix

Matrix = pd.read_excel('90-1261 Incomplete Matrix.xlsx', sheet_name = 'Sheet1', header = 0, index_col = 0)
Matrix = Matrix.fillna(0)

Matrix_data = np.asarray(Matrix)
M_shape = np.shape(Matrix_data)

for i in range(0, M_shape[0]):
    for j in range(0, M_shape[1]):
        if Matrix_data[i, j] == 0:
            Adsorbent_id = tf.constant([[i]])
            Adsorbate_id = tf.constant([[j]])
            Matrix.iloc[i, j] = model([Adsorbent_id, Adsorbate_id]).numpy()[0][0]
            
Writer = pd.ExcelWriter('Complete_matrix_35bar.xlsx', engine = 'xlsxwriter')
Matrix.to_excel(Writer, 'Sheet1', header = True, index = True)
Writer.save()