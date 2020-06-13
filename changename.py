#-*- encoding=utf-8 -*-

import os

rootpath = 'C:\\Users\\inuba\\Videos\\资料'

for dirpath, dirnames, files in os.walk(rootpath):
    for fileobj in files:
        if len(fileobj) > 50:
            new_file_name = fileobj[57:]
            os.rename(os.path.join(dirpath,fileobj), os.path.join(dirpath,new_file_name))
        else:
            print('pass')