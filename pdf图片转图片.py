import fitz
import os


#  打开PDF文件，生成一个对象
doc = fitz.open(r'D:\Users\Administrator\Desktop\1.pdf')

for pg in range(doc.pageCount):
    page = doc[pg]
    rotate = int(0)
    # 每个尺寸的缩放系数为2，这将为我们生成分辨率提高四倍的图像。
    zoom_x = 2.0
    zoom_y = 2.0
    trans = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
    pm = page.getPixmap(matrix=trans, alpha=False)
    pm.writePNG('%s.png' % pg)





