import os
from PyPDF2 import PdfFileMerger


target_path = r'D:\Users\Administrator\Desktop\新建文件夹 (5)'

pdf_lst = [f for f in os.listdir(target_path) if f.endswith('.pdf')]


tmp = []
for item in pdf_lst:
    tmp.append(item[:-4])

res = [int(i) for i in tmp if isinstance(i, str)]

ss = sorted(res)
ll = []
for j in ss:
    sj = str(j) + ".pdf"
    ll.append(sj)

pdf_lst = [os.path.join(target_path, filename) for filename in ll]

file_merger = PdfFileMerger()
for pdf in pdf_lst:
    file_merger.append(pdf)     # 合并pdf文件

file_merger.write(r"D:\Users\Administrator\Desktop\目录.pdf")
#此程序合成时只能用1、2、3等数字命名。