#!/usr/bin/python
# -*-coding:utf-8-*-

import datetime
import json
import os

import keyboard
import prettytable as pt
import pyperclip
import requests
from PIL import Image, ImageGrab


def WelCome():
    '''打印程序签名'''
    tb = pt.PrettyTable()
    tb.field_names = ['名称', 'Sina图片外链生成工具']
    tb.add_row(['作者', 'wwcheng'])
    tb.add_row(['微信公众号', '15Sceconds'])
    tb.add_row(['GitHub项目地址', 'https://github.com/nigo81/tools-for-auditor'])
    tb.add_row(['运行环境', 'Python3'])
    tb.add_row(['', '截图后 使用组合键 Ctrl+Shift+A 来上传'])
    tb.add_row(['使用方法', '上传成功后 外链将直接拷贝到剪贴板中'])
    tb.add_row(['', '截图和外链文本将会保存在TEMP文件夹中'])
    print(tb)


def clipboard_save(pic_path):
    '''从将剪切板中的图片保存至本地'''
    pic = True
    im = ImageGrab.grabclipboard()
    if im != None:
        im.save(pic_path)
        print('2. 正在上传剪切板中的图片')
    else:
        print('2. 剪切板上未发现图片')
        pic = False
    return pic


def Sina_upload(pic_path):
    '''将保存的图片上传至Sina服务器，并返回图片外链'''
    try:
        url = 'https://api.top15.cn/picbed/picApi.php?type=multipart'
        files = {'file': ('imageGrab.png', open(pic_path, 'rb'))}
        data = requests.post(url, files=files)
        data = json.loads(data.text)
        img_url = data['url']  # https方式
        # img_url = r'<img src="%s"/>'%img_url  # html方式
        # img_url = r'![img](%s)'%img_url       # markdown方式
        print("3. 生成的图片外链：{}".format(img_url))
        return img_url
    except Exception as err:
        print("3. 图片上传失败：{}".format(err))
        return None


def save_to_clipboard(img_url, txt_path):
    '''将图片外链储存至剪切板，同时保存在imgurl.txt中'''
    if img_url != None:
        pyperclip.copy(img_url)
        pyperclip.paste()
        f = open(txt_path, 'a+')
        f.write(img_url+'\n')
        f.close
        print('4. 外链已经保存并拷贝到剪贴板中')


def temp_path():
    '''创建储存图片和外链文本的本地文件夹'''
    now_path = os.path.abspath(os.curdir)
    if not os.path.exists(now_path+r'\temp'):
        os.mkdir(now_path+r'\temp')
    time_str = datetime.datetime.strftime(
        datetime.datetime.now(), '%Y%m%d%H%M%S')
    pic_path = now_path+r'\temp\imageGrab_%s.png' % time_str
    txt_path = now_path+r'\temp\imageUrl.txt'
    path = [pic_path, txt_path]
    return path


def main():
    '''按键监控与主体逻辑'''
    print('\n1. 截图后 按组合键 Ctrl+Shift+A 开始上传')
    if keyboard.wait(hotkey='ctrl+shift+a') == None:
        if clipboard_save(temp_path()[0]):
            img_url = Sina_upload(temp_path()[0])
            save_to_clipboard(img_url, temp_path()[1])


if __name__ == "__main__":
    WelCome()
    while True:
        main()
