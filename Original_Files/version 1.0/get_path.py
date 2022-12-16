# 获取选择文件路径
import tkinter as tk
from tkinter import filedialog
import json
import pymsgbox as mb
# srcfile 需要复制、移动的文件   
# dstpath 目的地址
import os
import shutil
from glob import glob



#Info1
mb.alert('请选择pdf文件路径！','提示')

###input
root = tk.Tk()
root.withdraw()

# 获取文件夹路径
input_path = filedialog.askopenfilename()


  # 使用askdirectory函数选择文件夹
print('\n输入的文件：', input_path)


#json变量
control={"input_file":str(input_path),}
json.dump(control,open('config.json','w'),indent=4)

#Info1
mb.alert('文件选择完成！')






