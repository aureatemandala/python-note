#-*- coding=utf-8 -*-

import os
import re       #正则表达式
import glob     #文件名搜索
import fnmatch  #匹配文件名
import time     #日期与时间
import datetime #日期与时间
from tempfile import TemporaryFile      #创建临时文件
from tempfile import TemporaryDirectory #创建临时文件夹
import shutil   #复制、移动
import zipfile  #压缩包


#获取当前python程序运行路径
print(os.getcwd())

#让python自动处理路径连接
print(os.path.join('first','second','last'))

#列出当前程序文件夹下的所有文件和文件夹并判断是否为文件夹
for file in os.listdir():
    print(file , os.path.isdir(file))

for file in os.scandir(os.getcwd()):
    print(file.name, file.path, '这是个文件', file.is_file())

#遍历文件夹
for dirpath, dirnames, files in os.walk(os.getcwd()):
    print(f'发现文件夹：{dirpath}')
    print(files)


#搜索文件名称
print('abc.txt'.startswith('abc'))
print('abc.txt'.endswith('txt'))

print(glob.glob('*正*Y*'))
# '**'表示任意层文件或文件夹 
# recursive=True 表示递归搜索
print(glob.glob('**/*.xlsx', recursive=True))



#匹配文件名
print(fnmatch.fnmatch('this_is_a_test_str','th*str'))
print(fnmatch.fnmatch('this_is_a_test_str','th*[_,a-z,0-9]str'))

#获取文件信息
for file in os.scandir(os.getcwd()):
    print(file.name, file.path, '这是一个文件夹', file.is_dir(), time.ctime(file.stat().st_mtime))
# file.stat()方法显示文件信息
# st_size 文件的体积大小，以字节为单位，除以1024为KB，以此类推
# st_atime 文件的最近访问时间
# st_mtime 文件的最近修改时间
# st_ctime Windows下，表示文件创建时间
# st_birthtime Linux与mac下，表示文件创建时间

#更直观的时间显示
that_time = datetime.datetime.fromtimestamp(6546168543)
print(that_time)
print(that_time.hour, that_time.minute, that_time.second)

#单独查询指定文件
file = os.stat('文件管理.py')
print(datetime.datetime.fromtimestamp(file.st_ctime))


#读取文件内容 不存在则创建
# r,w,a 读，写，添加
file = open('files/flow.txt', 'r', encoding='utf-8')
text = file.readlines()
print(text)
file.close()
#建议使用的写法 with ... as ...
with open('files/flow.txt', 'w', encoding='utf-8') as file:
    text = '多年以后，面对行刑队，\n奥雷里亚诺·布恩迪亚上校回想起小时父亲带他去看冰块的那个下午。'
    file.write(text)


#创建临时文件存储数据
with TemporaryFile('w+') as file:
    print(file.name)
    file.write('一如一本书籍包含了一个世界，这个残破美丽的世界也必被一本书所包含。')
    file.seek(0)
    text = file.readlines()
    print(text)

#创建临时文件夹
with TemporaryDirectory() as temp_dir:
    print(f'已创建临时文件夹，目录为：{temp_dir}')

#判断文件夹是否存在，并创建文件夹
if not os.path.exists('files/testdir'):
    os.mkdir('files/testdir')
#创建多层文件夹
#os.makedirs('files/testdir/secondtestdir/lasttestdir')

#复制文件
shutil.copy('./files/flow.txt', './files/testdir/flow.txt.bak')

#复制文件夹
shutil.copytree('./files/testdir', './files/second/testdir bak')

#删除文件
os.remove('./files/testdir/flow.txt.bak')

#删除文件夹
shutil.rmtree('./files/second/testdir bak')

#移动文件或文件夹
#shutil.move('要移动的文件或文件夹', '要移动到的位置')
#第二个参数，写文件夹位置，则移动到该文件夹下
#第二个参数，写文件路径，移动到这个路径并重命名
shutil.move('./files/flow.txt', './files/testdir/movetest.txt')
shutil.move('./files/testdir/movetest.txt', './files/flow.txt')
shutil.move('./files/testdir', './files/second/')
shutil.move('./files/second/testdir', './files/')

#重命名文件或文件夹  可以用来移动
os.rename('./files/test.xlsx', 'test.xls')
os.rename('test.xls', './files/test.xlsx')


#读取zip压缩包 中文编码调整
with zipfile.ZipFile('./files/ziptest.zip', 'r') as zipobj:
    for filename in zipobj.namelist():
        print(filename.encode('cp437').decode('gbk'))
        info = zipobj.getinfo(filename)
        print('源文件大小：', info.file_size,'压缩后的大小：', info.compress_size)

#将压缩包内单个文件解压出来 并解决中文问题
with zipfile.ZipFile('./files/ziptest.zip', 'r') as zipobj:
    zipobj.extract('flow.txt','./files/testdir')
    zh_name = '这是一个中文测试文件.txt'
    zh_name_cp = zh_name.encode('gbk').decode('cp437')
    zipobj.extract(zh_name_cp,'./files/testdir')
    os.rename('./files/testdir/'+zh_name_cp, './files/testdir/'+zh_name)

#完全解压压缩包 
#.extractall(path='解压到此位置', pwd=b'密码')
with zipfile.ZipFile('./files/zipwithpasswd.zip', 'r') as zipobj:
    zipobj.extractall(path='./files/second', pwd=b'123456')


#创建压缩包
# w 创建一个压缩包
# a 添加文件至压缩包
file_list = ['./files/testdir', './files/flow.txt', './files/test.xlsx', './files/这是一个中文测试文件.txt']
with zipfile.ZipFile('./files/create.zip', 'w') as zipobj:
    for file in file_list:
        zipobj.write(file)