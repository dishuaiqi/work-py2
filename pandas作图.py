from docx import Document
import openpyxl
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt

太和二场=pd.read_excel(r'D:\Users\Administrator\Desktop\太和二场.xlsx',sheet_name=1)
plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
plt.rcParams['axes.unicode_minus']=False #用来正常显示负号
太和二场.sort_values(by='S/P值',inplace=True,ascending=False)
print(太和二场.columns)
太和二场.plot.barh(x='原始编号',y=['S/P值','平均值'])
plt.tight_layout()
plt.show()

