import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
#matplotlib inline

data=pd.read_excel('C:/Users/kim20/Downloads/sample.xlsx')

print(data.head(12))
plt.figure(figsize=(8,4))
plt.plot(data.time, data.result, label='전월')
plt.grid()
plt.legend()
plt.show()