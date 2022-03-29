
import pytesseract
from PIL import Image
from docx import Document
import fitz
import os



#  打开PDF文件，生成一个对象
doc = fitz.open(r'D:\Users\Administrator\Desktop\新建文件夹 (2)\031.pdf')

# for pg in range(doc.pageCount):
#     page = doc[pg]
#     rotate = int(0)
#     # 每个尺寸的缩放系数为2，这将为我们生成分辨率提高四倍的图像。
#     zoom_x = 2.0
#     zoom_y = 2.0
#     trans = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
#     pm = page.getPixmap(matrix=trans, alpha=False)
#     pm.writePNG('%s.png' % pg)





pytesseract.pytesseract.tesseract_cmd=r'C:\Program Files\Tesseract-OCR\tesseract'

# 使用pytesseract对英文进行识别，lang参数可省略
# code = pytesseract.image_to_string(Image.open('2.png'))
# print(code)
# 使用pytesseract对中文（含英文，但识别率降低）进行识别
# code = pytesseract.image_to_string(Image.open(r'3.png'))
# print(code)
image=Image.open(r'0.png')
zifuweizhi=r'C:\Program Files\Tesseract-OCR\tessdata'
content=pytesseract.image_to_string(image,lang='chi_sim')
print(content)
with open(r'D:\Users\Administrator\Desktop\图片39.txt','w',encoding='utf-8') as fp:
    fp.write(content)

# Doc=Document(r'D:\Users\Administrator\Desktop\1.docx')
# Doc.add_paragraph(content)
# Doc.save(r'D:\Users\Administrator\Desktop\图片1.docx')