import os,shutil
import time
import datetime
allFileNum = 0
myfile=[]
mydir=[]
this_folder=input("原始路径：").replace("\\",'/')
those_folder=input("目标路径：").replace("\\",'/')
def printPath(level, path):
    global allFileNum
    '''''
    打印一个目录下的所有文件夹和文件
    '''
    # 所有文件夹，第一个字段是次目录的级别
    dirList = []
    # 所有文件
    fileList = []
    # 返回一个列表，其中包含在目录条目的名称(google翻译)
    files = os.listdir(path)
    # 先添加目录级别
    dirList.append(str(level))
    for f in files:
        if (os.path.isdir(path + '/' + f)):
            # 排除隐藏文件夹。因为隐藏文件夹过多
            if (f[0] == '.'):
                pass
            else:
                # 添加非隐藏文件夹
                dirList.append(f)
                mydir.append(path + '/' + f)
        if (os.path.isfile(path + '/' + f)):
            # 添加文件
            fileList.append(f)
            myfile.append(path + '/' + f)
    # 当一个标志使用，文件夹列表第一个级别不打印
    i_dl = 0
    for dl in dirList:
        if (i_dl == 0):
            i_dl = i_dl + 1
        else:
            # print("得到的文件夹",'-' * (int(dirList[0])), dl)
            # 打印目录下的所有文件夹和文件，目录级别+1
            printPath((int(dirList[0]) + 1), path + '/' + dl)
    for fl in fileList:
        # print("得到的文件路径",'-' * (int(dirList[0])), fl)
        # 随便计算一下有多少个文件
        allFileNum = allFileNum + 1

#文件处理
def judge_file(oldPath,location):
    def TimeStampToTime(timestamp):
        timeStruct = time.localtime(timestamp)
        return time.strftime('%Y-%m-%d %H:%M:%S', timeStruct)
    def move_file(new_dir):
        old_file_name = oldPath.split("/")[-1]
        # 将文件移动到新文件夹
        shutil.move(oldPath, new_dir + '/' + old_file_name)  # a->b
        print("-"*50+"已完成：%.2f" % ((location + 1) / allFileNum*100))
    # ImageDate = datetime.datetime.strftime(time.ctime(os.stat(imgPath).st_mtime), "%Y-%m-%d %H:%M:%S")
    a = os.stat(oldPath).st_mtime
    #得到文件创建时间，判断文件夹是否存在
    creat_time=TimeStampToTime(a)[:-9]
    print(creat_time) #str 2021-01-10 10:05:31
    new_dir="%s/%s"%(those_folder,creat_time)
    #不存在文件夹则创建文件夹
    if not os.path.exists(new_dir):
        os.makedirs(new_dir)
        move_file(new_dir)
    #如果存在就判断是否是重复文件
    else :
        if oldPath.split("/")[-1] in os.listdir(new_dir):
            print("重复文件，略过")
            pass
        else:
            move_file(new_dir)
def do_all():
    for i in myfile:
        judge_file(i,myfile.index(i))
printPath(1, this_folder)
do_all()
input()
print("完成")
