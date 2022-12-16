#!/usr/bin/env python
# -*- coding: UTF-8 -*-
'''
@File    ：pdftoword_api_version.py
@IDE     ：PyCharm
@Author  ：akkk
@Date    ：2022/12/5 12:16
@ver 2.0 
@Description：pdf转word
@GitHub  ：github.com/tonyhuang0711
@ver.2 written on 16/12/2022
@finished at 17:09 UTC+8:00
'''

# 获取选择文件路径
import os
import requests
import time
from builtins import print
import tkinter as tk
from tkinter import filedialog
import json
import pymsgbox as mb
import win32api
import win32con
# srcfile 需要复制、移动的文件
# dstpath 目的地址
from glob import glob
import sys


def get_system():
    import os
    print(os.name)
    if os.name == 'nt':
        sys = 1
        return sys
    else:
        sys = 0
        return sys


sys = get_system()

# Info1
if sys == 1:
    import win32api
    import win32con
    win32api.MessageBox(0, "请选择pdf文件！", "选择文件", win32con.MB_OK)
elif sys == 0:
    mb.alert('请选择pdf文件路径!', '提示')

# input
root = tk.Tk()
root.withdraw()

# 获取文件夹路径
input_path = filedialog.askopenfilename()

# 使用askdirectory函数选择文件夹
print('\n输入的文件:', input_path)


def check_path():
    import os
    if not os.path.exists(input_path):
        print('文件不存在！')
        return False
    else:
        return True


check_path = check_path()


def rechoose():
    if check_path == True:
        print('读取成功！')
    else:
        if sys == 1:
            import win32api
            import win32con
            win32api.MessageBox(0, "文件不存在！是否重新选择！", "提示", win32con.MB_YESNO)
            if win32api.MessageBox(0, "文件不存在！是否重新选择！", "提示", win32con.MB_YESNO) == win32con.IDYES:
                rechoose_path = filedialog.askopenfilename()
                return rechoose_path
            elif win32api.MessageBox(0, "文件不存在！是否重新选择！", "提示", win32con.MB_YESNO) == win32con.IDNO:
                exit()
            else:
                print('错误！')
                exit()
        elif sys == 0:
            mb.confirm('文件不存在！', '提示', ['重新选择', '退出'])
        if mb.confirm == '重新选择':
            rechoose_path = filedialog.askopenfilename()
            return rechoose_path
        elif mb.confirm == '退出':
            exit()
        else:
            print('错误！')
            exit()


rechoose_path = rechoose()


def file():
    if check_path == True:
        return input_path
    elif check_path == False:
        return rechoose_path
    else:
        print('错误！')
        exit()


file = file()
print('文件路径:', file)

# Info1
mb.alert('文件选择完成！')


# 创建output dir
# 导入os模块
# 获取当前工作目录
cwd = os.getcwd()
# 判断是否存在文件夹/output
isExists = os.path.exists(cwd+'/output')
# 判断结果
if not isExists:
    # 如果不存在则创建目录
    # 创建目录操作函数
    os.makedirs(cwd+'/output')
    print(cwd+'/output 创建成功')
else:
    # 如果目录存在则不创建，并提示目录已存在
    print(cwd+'/output 目录已存在')

input_path = file
print(input_path)


class PDF2Word():
    def __init__(self):
        self.machineid = 'ccc052ee5200088b92342303c4ea9399'
        self.token = ''
        self.guid = ''
        self.keytag = ''

    def produceToken(self):
        url = 'https://app.xunjiepdf.com/api/producetoken'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64; rv:76.0) Gecko/20100101 Firefox/76.0',
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'X-Requested-With': 'XMLHttpRequest',
            'Origin': 'https://app.xunjiepdf.com',
            'Connection': 'keep-alive',
            'Referer': 'https://app.xunjiepdf.com/pdf2word/', }
        data = {'machineid': self.machineid}
        res = requests.post(url, headers=headers, data=data)
        res_json = res.json()
        if res_json['code'] == 10000:
            self.token = res_json['token']
            self.guid = res_json['guid']
            print('成功获取token')
            return True
        else:
            return False

    def uploadPDF(self, filepath):
        filename = filepath.split('/')[-1]
        files = {'file': open(filepath, 'rb')}
        url = 'https://app.xunjiepdf.com/api/Upload'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64; rv:76.0) Gecko/20100101 Firefox/76.0',
            'Accept': '*/*',
            'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
            'Content-Type': 'application/pdf',
            'Origin': 'https://app.xunjiepdf.com',
            'Connection': 'keep-alive',
            'Referer': 'https://app.xunjiepdf.com/pdf2word/', }
        params = (
            ('tasktype', 'pdf2word'),
            ('phonenumber', ''),
            ('loginkey', ''),
            ('machineid', self.machineid),
            ('token', self.token),
            ('limitsize', '2048'),
            ('pdfname', filename),
            ('queuekey', self.guid),
            ('uploadtime', ''),
            ('filecount', '1'),
            ('fileindex', '1'),
            ('pagerange', 'all'),
            ('picturequality', ''),
            ('outputfileextension', 'docx'),
            ('picturerotate', '0,undefined'),
            ('filesequence', '0,undefined'),
            ('filepwd', ''),
            ('iconsize', ''),
            ('picturetoonepdf', ''),
            ('isshare', '0'),
            ('softname', 'pdfonlineconverter'),
            ('softversion', 'V5.0'),
            ('validpagescount', '20'),
            ('limituse', '1'),
            ('filespwdlist', ''),
            ('fileCountwater', '1'),
            ('languagefrom', ''),
            ('languageto', ''),
            ('cadverchose', ''),
            ('pictureforecolor', ''),
            ('picturebackcolor', ''),
            ('id', 'WU_FILE_1'),
            ('name', filename),
            ('type', 'application/pdf'),
            ('lastModifiedDate', ''),
            ('size', ''),)
        res = requests.post(url, headers=headers, params=params, files=files)
        res_json = res.json()
        if res_json['message'] == '上传成功':
            self.keytag = res_json['keytag']
            print('成功上传PDF')
            return True
        else:
            return False

    def progress(self):
        url = 'https://app.xunjiepdf.com/api/Progress'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64; rv:76.0) Gecko/20100101 Firefox/76.0',
            'Accept': 'text/plain, */*; q=0.01',
            'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'X-Requested-With': 'XMLHttpRequest',
            'Origin': 'https://app.xunjiepdf.com',
            'Connection': 'keep-alive',
            'Referer': 'https://app.xunjiepdf.com/pdf2word/', }
        data = {
            'tasktag': self.keytag,
            'phonenumber': '',
            'loginkey': '',
            'limituse': '1'}
        res = requests.post(url, headers=headers, data=data)
        res_json = res.json()
        if res_json['message'] == '处理成功':
            print('PDF处理完成')
            return True
        else:
            print('PDF处理中(请不要中途退出！！)')
            return False

    def downloadWord(self, output):
        url = 'https://app.xunjiepdf.com/download/fileid/%s' % self.keytag
        res = requests.get(url)
        with open(output, 'wb') as f:
            f.write(res.content)
            print('PDF下载成功("%s")' % output)

    def convertPDF(self, filepath, outpath):
        filename = filepath.split('/')[-1]
        filename = filename.split('.')[0] + '.docx'
        self.produceToken()
        self.uploadPDF(filepath)
        while True:
            res = self.progress()
            if res == True:
                break
            time.sleep(1)
        self.downloadWord(outpath + filename)


if __name__ == '__main__':
    pdf2word = PDF2Word()
    pdf2word.convertPDF(file, "output/")

# Info1
if sys == 1:
    win32api.MessageBox(0, '转换完成！！', '提示', win32con.MB_OK)
    # 获取运行目录
    import os
    path = os.getcwd()
    # 打开运行目录下output文件夹
    os.startfile(path + "\output")
    sys.exit()

else:
    mb.alert('转换完成！！', '提示')
    sys.exit()
