# “你是什么垃圾？”Python垃圾分类帮你告别灵魂拷问

作  者 | 王文铖

公众号 | 15Seconds

小龙虾的黄是什么垃圾？小龙虾的壳是什么垃圾？小龙虾的肉是什么垃圾？粽叶是什么垃圾？粽肉是什么垃圾？你是什么垃圾？（小龙虾都要哭了。。。）

如果真的是按照这个分类的话，那真是吃小龙虾一时爽，打扫战场愁断肠啊！

也许这些追问显得有些较真和搞笑，但也从另一方面显示出近期人们对于垃圾分类的关注。

现在是人工智能（AI）的时代，今天的内容，是讨论如何利用人工智能对垃圾进行分类，从而让大家~~有吃小龙虾、喝奶茶的资格~~不再为垃圾丢哪里而烦恼。

前段时间接触了百度AI，今天的内容依然是借助百度AI开放平台，用Python对垃圾图片进行物体识别，再通过相关的垃圾分类查询网站进行分类，最终实现Python简易版的垃圾分类。

一、程序效果图



二、准备工作

导入网络爬虫常用的模块
~~~
import base64
import json
import os
import re
import requests
import win32ui
from urllib import parse
from bs4 import BeautifulSoup
~~~

三、导入图片
首先需要对图片进行64位编码（将一副图片数据编码成一串字符串，使用该字符串代替图像地址）。

导入图片函数如下：
~~~
def image_load():
    dlg = win32ui.CreateFileDialog(1)
    dlg.SetOFNInitialDir(os.path.abspath(os.curdir))
    dlg.DoModal()
    image_path = dlg.GetPathName()
    f = open(image_path, 'rb')
    image = base64.b64encode(f.read())
    f.close
    return image
~~~

四、物体识别

图像物体识别是人工智能的典型案例。输入一张图片，就能得到这张图片中的物体名称或者场景名称。

我这里采用的是百度AI的通用物体和场景识别API来实现垃圾图像的识别。借助百度较为成熟开放的技术，省去我们搜集图像样本，建立模型的时间。

物体识别函数如下：
~~~
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
~~~

五、垃圾分类

借助网站（https://lajifenleiapp.com/） 对图像识别结果进行爬取，最终返回分类结果。

~~~
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
~~~

六、程序入口

最后，添加循环和程序入口，最终封装为可执行文件（exe）。
~~~
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
~~~

结语

上面这些都是人工智能与实际生活相结合的一些案例，我相信在将来肯定会实现，但是实现的过程确实充满艰辛，这对物体识别的精确度要求非常高，之前目前的水平还达不到。图片内的物品数量、图片自身的素质等问题短期内都难以解决。

虽然前路不易，但是我对人工智能的未来却更为期待。

本公众号工具会在Github上正式发布和后续更新。

Github项目地址：https://github.com/nigo81/tools-for-auditor

微信公众号：15Seconds