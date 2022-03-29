import os
import sys
import os.path
import time
from shutil import Error
from shutil import copystat
from shutil import copy2

path_str = r"D:\Users\Administrator\Desktop\新建文件夹\WeChat Files\wxid_oxph17qejeir31\FileStorage\File\2021-05";


def copy_file(src_file, dst_dir):
    if os.path.isdir(dst_dir):
        pass;
    else:
        os.makedirs(dst_dir);
    print(src_file);
    print(dst_dir);
    copy2(src_file, dst_dir)


def walk_file(file_path):
    for root, dirs, files in os.walk(file_path, topdown=False):
        for name in files:
            com_name = os.path.join(root, name);
            t = os.stat(com_name);
            copy_path_str = path_str + r"\year" + str(time.localtime(t.st_mtime).tm_year) + r"\month" + str(
                time.localtime(t.st_mtime).tm_mon);
            print(copy_path_str);
            copy_file(com_name, copy_path_str);
        for name in dirs:
            walk_file(name);


walk_file(path_str);
