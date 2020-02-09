作者 / wwcheng
公众号 / 15Seconds

一、前言

Q1: 图床是什么？

图床就是专门用来存放图片，同时允许你把图片对外连接的网上空间。

举个例子来说，就是你在浏览器中输入图片外链，即可看到链接对应的图片。

Q2: 把图片转换成外链的形式有什么好处？（暂时对我而言）

1.在Markdown中插入图片

如果你经常使用Markdown写作的话，会发现添加图片比较繁琐，不是语法繁琐，而是你在Markdown中插入的本地图片上传到网上之后无法显示。一般的解决办法就是，先把图片上传至某某相册或者服务器上，再复制图片外链插入Markdown中，本工具可以直接把截图转换成外链，省去上传的麻烦，大大提高Markdown写作效率。


2.在Anki中插入图片

Anki是一款卡片记忆软件，可以通过卡片正面的提示和问题来记忆背面记录的内容，可以通过上传本地附件的形式在卡片背面插入图片，但是这样会对软件的同步功能造成很大负担，因为Anki服务器布置在国外，国内可以使用但速度慢，通过在卡片背面上传外链的方式可以有效解决这一问题。


3.······

Q3: 把图片转成外链万无一失吗？

本工具使用的是新浪的公开服务器，而不是自己搭建的图床服务器，可能有部分限制，网上有出现外链失效的情况，虽然我目前还没出现过，但是还是要做好备份工作。

同时需要考虑隐私和法律问题，最好上传不涉及隐私、不违反法律的图片。

有兴趣的可以参照源码自己购买服务器搭建图床。

二、工具使用方法

截图后，同时按下Ctrl+Shift+A，会将图片外链保存至电脑剪切板，按下Ctrl+V即可看见外链

如何截图？微信截图快捷键 Alt+A;  QQ截图快捷键 Ctrl+Alt+A;  Snipaste截图快捷键 F1


三、“硬核”区域

1.导入依赖环境
~~~
import datetime
import json
import os

import keyboard
import prettytable as pt
import pyperclip
import requests
from PIL import Image, ImageGrab
~~~
2.利用PrettyTable美化程序的欢迎界面
~~~
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
~~~
3.将剪切板中的图片保存至本地
~~~
def clipboard_save(pic_path):
    '''将剪切板中的图片保存至本地'''
    pic = True
    im = ImageGrab.grabclipboard()
    if im != None:
        im.save(pic_path)
        print('2. 正在上传剪切板中的图片')
    else:
        print('2. 剪切板上未发现图片')
        pic = False
    return pic
~~~
4.将保存的图片上传至Sina服务器，并返回图片外链

源码中提供了3中钟输出方式，可以直接输出为Markdown格式或者HTML <img>标签。
~~~
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
~~~
5.通过pyperclip将图片外链储存至剪切板，同时保存在imgurl.txt中
~~~
def save_to_clipboard(img_url, txt_path):
    '''将图片外链储存至剪切板，同时保存在imgurl.txt中'''
    if img_url != None:
        pyperclip.copy(img_url)
        pyperclip.paste()
        f = open(txt_path, 'a+')
        f.write(img_url+'\n')
        f.close
        print('4. 外链已经保存并拷贝到剪贴板中')
~~~
6.创建储存图片和外链文本的本地文件夹，用于备份图片
~~~
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
~~~
7.按键监控与主程序入口，最终封装为可执行文件
~~~
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
~~~
本公众号工具及源码会在Github上正式发布和后续更新。

Github项目地址：https://github.com/nigo81/tools-for-auditor

微信公众号：15Seconds

订阅公众号之后，在后台回复“tools”可获取15Seconds公众号所有工具！
