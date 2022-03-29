import pymysql
import pandas as pd
from sqlalchemy import create_engine
import os
# 病原更新=pd.read_excel(r'D:\Users\Administrator\Desktop\病原更新.xlsx')
# 抗体=pd.read_excel(r'D:\Users\Administrator\Desktop\抗体.xlsx')

# 联系人=pd.read_excel(r'D:\Users\Administrator\Desktop\联系人.xlsx')
# 打开数据库连接

db = pymysql.connect(user='root',password='666666',host='localhost',database='实验室',charset='utf8mb4')

sqla = "CREATE DATABASE IF NOT EXISTS test character set utf8;"
cursor = db.cursor()
cursor.execute(sqla)

conn = create_engine('mysql+pymysql://root:666666@localhost:3306/实验室?charset=utf8')
sql1='SELECT * FROM `病原`; '

sql_2 ='SET @@global.sql_mode= "NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION";'
cursor.execute(sql_2)
# 病原更新.to_sql('病原更新', conn, schema='实验室', if_exists='replace')
# 抗体.to_sql('抗体', conn, schema='实验室', if_exists='replace')
# 联系人.to_sql('联系人', conn, schema='实验室', if_exists='replace')
# sq更新='INSERT INTO `抗体`(`index`,`日期`,`猪群种类`,`原始序号`,`检测类型`,`OD值`,`S/N值`,`结果判定`,`猪群来源`,`免疫记录`,`免疫剂量`,`疫苗厂家`,`试剂盒`) SELECT `index`,`日期`,`猪群种类`,`原始序号`,`检测类型`,`OD值`,`S/N值`,`结果判定`,`猪群来源`,`免疫记录`,`免疫剂量`,`疫苗厂家`,`试剂盒` FROM `抗体更新`;'
a=pd.read_sql(sql1,conn)
# cursor = db.cursor()
# cursor.execute(sq更新)
#
a.to_excel(r'D:\Users\Administrator\Desktop\病原.xlsx')
print('完成')
