import pandas as pd
import sys
import numpy as np
import glob


all_data=pd.DataFrame()
for f in glob.glob('C:/Users/kim20/Downloads/sample*.xlsx'):
    df=pd.read_excel(f)
    all_data=all_data.append(df, ignore_index=True)
    
print(all_data.shape)
print(all_data.head())
all_data.to_excel("C:/Users/kim20/Downloads/result.xlsx", header=True, index=True)