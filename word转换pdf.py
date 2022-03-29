
from win32com.client import gencache
from win32com.client import constants
import os
import time

文件夹=r'D:\Users\Administrator\Desktop\新建文件夹 (4)'#放word 文件夹
文件夹2=r'D:\Users\Administrator\Desktop\新建文件夹 (5)'
a=os.listdir(文件夹)
pdf_name=[]
for i in a:
    a1=os.path.splitext(i)
    pdf_name.append(a1[0]+'.pdf')
print(pdf_name)

s=a[0]
print(a[0])
文件=[]
for i in a:
    a1=os.path.join(文件夹,i)
    文件.append(a1)
# print(文件)
# name=[]
# for i in range(len(a)):
#     ming=a[i][0:len(a[1])-5]+'.pdf'
#     name.append(ming)
# print(name)
# print(文件)
名字=[]
for i in pdf_name:
    a1 = os.path.join(文件夹2, i)
    名字.append(a1)
print(名字)
def word_pdf(wordpath,pdfpath):
    word = gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(wordpath)
    doc.ExportAsFixedFormat(pdfpath,constants.wdExportFormatPDF)
    word.Quit()


# print(name[2])
for i in range(len(文件)):
    wordpath=文件[i]
    pdfpath=名字[i]
    print(wordpath)
    word_pdf(wordpath,pdfpath)
    time.sleep(3)






























