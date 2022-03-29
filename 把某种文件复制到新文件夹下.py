# #!/usr/bin/env python3
# # -*- coding: utf-8 -*-
# # @Author: yudengwu(余登武）
# # @Date  : 2020/10/4
# #@email:1344732766@qq.com
import os
import shutil
#
print('输入格式：D:/老男孩视频/第02部分-Python之基础讲解(09-28)/day09-Python安装与初识')   # 复制过来的路径中间是用 \ 隔开的，可能会出现转译的效果，可以改为 /

path = r'D:\Users\Administrator\Desktop\新建文件夹\WeChat Files\wxid_oxph17qejeir31\FileStorage\File'       # 想要查找的文件夹路径
new_path = r'D:/测试'            # 想要保存的文件夹路径，需新建一个文件夹

for root, dirs, files in os.walk(path):
    for i in range(len(files)):
        print(files[i])     # 打印出文件夹中所有的文件名，包括本文件夹中子文件夹中的文件
        # 匹配想要复制转移的文件类型
        if files[i][-4:] == 'xlsx':
            file_path = root + '/' + files[i]
            new_file_path = new_path + '/' + files[i]
            shutil.copy(file_path, new_file_path)

















import os
import re

import time
from datetime import datetime
#首先定义规则，
key='单'
new_path = r'D:/测试1'
n_path=r'D:/测试2'
for filename in os.listdir(new_path):
    if key in filename:
        print('1')
        old_file=os.path.join(new_path,filename)
        new_file_path =os.path.join(n_path,filename)
        shutil.copy(old_file, new_file_path)








