from pdf2docx import Converter
import os
import time
# pdf_file = r'D:\Users\Administrator\Desktop\核酸保护液试验.pdf'
# docx_file = r'D:\Users\Administrator\Desktop\核酸保护液试验.docx'
def pdf_word(pdf_file,word_file):

    # convert pdf to docx
    cv = Converter(pdf_file)
    cv.convert(word_file, start=0, end=None)
    cv.close()


文件夹1=r'D:\Users\Administrator\Desktop\新建文件夹 (4)'#放word 文件夹

文件夹2=r'D:\Users\Administrator\Desktop\新建文件夹 (5)'#放pdf 文件夹

a=os.listdir(文件夹2)
print(a)
word_name=[]
for i in a:
    a1=os.path.splitext(i)
    word_name.append(a1[0]+'.docx')
print(word_name)
pdf_file1=[]
for i in range(len(a)):
    s=os.path.join(文件夹2,a[i])
    pdf_file1.append(s)
print(pdf_file1)
word_file1=[]
for i in range(len(word_name)):
     s=os.path.join(文件夹1,word_name[i])
     word_file1.append(s)
print(word_file1)
for i in range(len(a)):
    pdf_file=pdf_file1[i]
    word_file=word_file1[i]
    print(pdf_file)
    pdf_word(pdf_file,word_file)
