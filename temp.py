import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import glob
import os
import csv


all_files = glob.glob('C:/Users/Administrator/Downloads/test/TDR*.csv')
all_data_frames = []
i = 0
for file in all_files:
    if i==0:
        data_frame = pd.read_csv(file, index_col=None)
        all_data_frames.append(data_frame)
    else:
        data_frame = pd.read_csv(file, index_col=None)
        all_data_frames.append(data_frame.iloc[:, 1:]) 
    i=i+1   
    
data_frame_concat = pd.concat(all_data_frames, axis=1, ignore_index=False)
data_frame_concat.to_csv("C:/Users/Administrator/Downloads/test/TDR_result1.csv", index=False)


#print(i)
k=data_frame_concat.iloc[:222, 1:]
lencol=len(k.columns)
min_value=k.min()
max_value=k.max()
k_matrix=k.as_matrix()

for i in range(0, lencol):
    if min_value[i]

print(min_value[0])

#for i in range(0,lencol):
#   index.append(np.where(k_matrix[:, [i]]==min_value[i]))





    
#min_row_index=np.argmin(k[0])
#print(data_frame_concat['Time [ns]']

#index=np.where(k == min_value)
#print(len(k.columns))
#print(data.head(12))
#plt.figure(figsize=(8,4))
#plt.plot(data.time, data.result, label='전월')
#plt.grid()
#plt.legend()
#plt.show()

#all_data.to_excel("C:/Users/Administrator/Downloads/test/result.xlsx", header=True, index=True)