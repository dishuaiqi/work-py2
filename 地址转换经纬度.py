# _*_ coding:utf-8 _*_
import requests
import json
import codecs
import re
from openpyxl import Workbook
wb = Workbook()
sheet = wb.active
sheet.title = "qiang"
def get_location(address, i):
    print(i)
    url = "http://restapi.amap.com/v3/geocode/geo"
    data = {
        'key': '36e40d76f8a91bc9de0b56eeca4b650b', #在高德地图开发者平台申请的key，需要替换为自己的key
        'address': address
    }
    r = requests.post(url, data=data).json()
    sheet["A{0}".format(i)].value = address.strip('\n')
    print(r)
    if r['status'] == '1':
        if len(r['geocodes']) > 0:
            GPS = r['geocodes'][0]['location']
            sheet["B{0}".format(i)].value = '[' + GPS +']'
        else:
            sheet["B{0}".format(i)].value = '[]'
    else:
       sheet["B{0}".format(i)].value = '未找到'
#将地址信息替换为自己的文件，一行代表一个地址，根据需要也可以自定义分隔符
f = codecs.open(r"地址信息.txt", "r", "utf-8")
i = 0
while True:
    line = f.readline()
    i = i + 1
    if not line:
        f.close()
        wb.save(r"保存文件1.xlsx")
        break
    get_location(line, i)
