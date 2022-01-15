import pandas as pd
import sys

#엑셀 이름 입력

excel_names = ['SwID_Info.xls', 'SysID_Info.xls']
excels = [pd.ExcelFile(name) for name in excel_names]
frames = [x.parse(x.sheet_names[0], header=None,index_col=None) for x in excels]
frames[1:] = [df[1:] for df in frames[1:]]
combined = pd.concat(frames)

#파일저장

combined.to_excel("합친엑셀.xls", header=False, index=False)