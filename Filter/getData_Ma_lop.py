import pandas as pd
import os as os
import numpy as np
from openpyxl import load_workbook

folder = 'tkb.xlsx'
target = open("mhp.txt","r")
ls = []

for s in target:
    ls.append(s)

for i in range(0,len(ls)):
    ls[i] = ls[i].replace('\n','')
    ls[i] = ls[i].replace('\t','')

print(ls)
data = pd.read_excel(folder,sheet_name = 'Sheet1')

temp = data[data['Mã lớp'] == float(ls[0])]

for i in range(1,len(ls)-1):
    temp = temp.append(data[data['Mã lớp'] == float(ls[i])])
    print(temp)
    
temp.to_excel(r'result.xlsx','Sheet1')
