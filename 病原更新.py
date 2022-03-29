import pymysql
import pandas as pd
from sqlalchemy import create_engine
import os
病原更新=pd.read_excel(r'D:\Users\Administrator\Desktop\病原更新.xlsx')



db = pymysql.connect(user='root',password='666666',host='localhost',database='实验室',charset='utf8mb4')

sqla = "CREATE DATABASE IF NOT EXISTS 实验室 character set utf8;"
cursor = db.cursor()
cursor.execute(sqla)

conn = create_engine('mysql+pymysql://root:666666@localhost:3306/实验室?charset=utf8')


sql_2 ='SET @@global.sql_mode= "NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION";'
cursor.execute(sql_2)
病原更新.to_sql('病原更新', conn, schema='实验室', if_exists='replace')

print('完成')

