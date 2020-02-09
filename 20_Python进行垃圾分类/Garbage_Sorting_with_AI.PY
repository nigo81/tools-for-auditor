# --coding:utf-8--
import base64
import json
import os
import re
import requests
import win32ui
from urllib import parse
from bs4 import BeautifulSoup


def image_load():
    dlg = win32ui.CreateFileDialog(1)
    dlg.SetOFNInitialDir(os.path.abspath(os.curdir))
    dlg.DoModal()
    image_path = dlg.GetPathName()
    f = open(image_path, 'rb')
    image = base64.b64encode(f.read())
    f.close
    return image


def baiduai_query(image):
    request_url = "https://aip.baidubce.com/rest/2.0/image-classify/v2/advanced_general"
    access_token = '24.40a71124704aa4cd793357f9fc2782c0.2592000.1565068326.282335-16728698'
    request_url = request_url + "?access_token=" + access_token
    data = parse.urlencode({"image": image})
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    request = requests.post(request_url, data=data, headers=headers)
    r = json.loads(request.text)
    key1 = r['result'][0]['keyword']
    key2 = r['result'][1]['keyword']
    keyword = [key1, key2]
    return keyword


def lajifenlei_query(query_word):
    request = requests.get("https://lajifenleiapp.com/sk/"+query_word)
    soup = BeautifulSoup(request.text, 'lxml')
    respose = soup.find_all('div', attrs={'class': 'row'})
    pattern = re.compile(r".*?属于.*?", re.S)
    if re.match(pattern, respose[2].text):
        print(respose[2].text.strip()+'\n')
        print(respose[4].text.strip())
        print(respose[5].text.strip())
        print(respose[6].text.strip().replace('\n\n', ''))
    else:
        result = 'https://lajifenleiapp.com/ 网站暂未收录\“%s\”的相关垃圾分类信息' % query_word
        print(result)


def main():
    image = image_load()
    print('正在识别图片···')
    query_word = baiduai_query(image)
    print('AI识别结果1'.center(40, '-'))
    lajifenlei_query(query_word[0])
    print('AI识别结果2'.center(40, '-'))
    lajifenlei_query(query_word[1])
    print('识别结束'.center(40, '-'))


if __name__ == '__main__':
    str = input('\n'+"输入1开始识别图片并分类，输入0结束程序：")
    while str == '1':
        main()
        str = input('\n'+"输入1开始识别图片并分类，输入0结束程序：")
