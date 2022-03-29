import os
import pandas as pd

# 将文件读取出来放一个列表里面

pwd = r'D:\Users\Administrator\Desktop\新建文件夹 (5)'  # 获取文件目录

# print(len(os.listdir(pwd)))
# # 新建列表，存放文件名
file_list = []
#
# 新建列表存放每个文件数据(依次读取多个相同结构的Excel文件并创建DataFrame)
dfs = []

for root, dirs, files in os.walk(pwd):  # 第一个为起始路径，第二个为起始路径下的文件夹，第三个是起始路径下的文件。
    for file in files:
        file_path = os.path.join(root, file)
        file_list.append(file_path)  # 使用os.path.join(dirpath, name)得到全路径
        df = pd.read_excel(file_path,sheet_name='城关')  # 将excel转换成DataFrame
        dfs.append(df)

# 将多个DataFrame合并为一个
df = pd.concat(dfs)

# 写入excel文件，不包含索引数据
df.to_excel(r'D:\Users\Administrator\Desktop\1.xlsx', index=False)
