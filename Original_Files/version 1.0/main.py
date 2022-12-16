#!/usr/bin/env python
# -*- coding: UTF-8 -*-
'''
@Project ：Python-crawler-set 
@File    ：pdftoword_api_version.py
@IDE     ：PyCharm
@Author  ：akkk
@Date    ：2022/12/5 12:16
'''
from builtins import print
import sys
import time
import requests
import json
import os
import pymsgbox as mb

#创建output dir
# 导入os模块
import os



# 读取json
from get_path import input_path


def read_json():
    with open("config.json", "r", encoding="utf-8") as f:
        data = json.load(f)
        json_str = str.replace("'", '"')
        # json.loads() ,要求json串格式中必须的双引号！！转换为字典
        from get_path import input_path
        json_dict = json.loads(input_path)
        return json_dict
    return data



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
            print('PDF处理中')
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
    pdf2word.convertPDF(input_path, "output/")

#Info1
mb.alert('转换完成！！','提示')
sys.exit()