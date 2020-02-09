import os
import random
import sys
import time
from urllib import request

import aircv as ac
import pandas as pd
import requests
from fake_useragent import UserAgent
from lxml import etree
from PIL import Image


class Qiyekuaizhao(object):

    def __init__(self):
        self.base_headers = {
            'Host': 'www.tianyancha.com',
            'Connection': 'keep-alive',
            'Cache-Control': 'max-age=0',
            'Upgrade-Insecure-Requests': '1',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-User': '?1',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
            'Sec-Fetch-Site': 'none',
            'Referer': 'https://www.tianyancha.com/',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'zh-CN,zh;q=0.9',
        }
        self.now_path = os.path.abspath(os.curdir)
        now_time = str(time.strftime('%Y%m%d-%H%M%S',time.localtime(time.time())))
        self.output_file_path = self.now_path + r'\Output\{}'.format(now_time) 
        os.makedirs(self.output_file_path)
        self.company_code = ''
        self.search_key = ''
        self.num = ''

    def cookies_load(self):
        '''内部函数：加载cookies，切换时间戳'''
        filepath = self.now_path + r'\Input\cookies.txt'
        cookies = (open(filepath, 'r').read()).strip()
        if len(cookies) < 50:
            print('请先在cookies.txt中填写cookies信息')
            return None
        else:
            cookies = cookies[0:-10] + str(int(time.time()))  # 10位数时间戳
            # cookies = cookies[0:-10]+str(int(time.time()*1000))#13位数时间戳
            return cookies

    def get_company_code(self):
        try:
            search_key_urlencode = request.quote(self.search_key)
            headers = {
                'User-Agent': UserAgent().random,
                'Cookie': self.cookies_load(),
            }
            headers.update(self.base_headers)
            search_url = 'https://www.tianyancha.com/search?key={}'.format(search_key_urlencode)
            r = requests.get(search_url, headers=headers)
            html = etree.HTML(r.text)
            self.company_code = html.xpath("//div[@class='search-result-single   ']/@data-id")[0]
        except Exception as e:
            pass

    def download_full_report(self):
        try:
            snapshot_url = 'https://www.tianyancha.com/snapshot/{}'.format(self.company_code)
            headers = {
                'User-Agent': UserAgent().random,
                'Cookie': self.cookies_load(),
            }
            headers.update(self.base_headers)
            r = requests.get(snapshot_url, headers=headers)
            html = etree.HTML(r.text)
            snapshot_img_url = html.xpath("//div[@class='snapshot']/div/div/a/@href")[0]
            # company_name = html.xpath("//a[@class='link-hover-click']/text()")[1]

            snapshot_img_name = self.num + '_' + self.company_code + '_full_report.png'
            snapshot_img_path = self.output_file_path + '\\' + snapshot_img_name
            request.urlretrieve(snapshot_img_url, snapshot_img_path)
            return snapshot_img_path
        except Exception as e:
            return None

    def crop_basic_info(self,snapshot_img_path):
        try:
            imgsrc = snapshot_img_path
            imgobj=self.now_path + r'\Input\fengzhi.png'
            imsrc = ac.imread(imgsrc)
            imobj = ac.imread(imgobj)
            match_result = ac.find_template(imsrc, imobj, 0.5)  
            if match_result is not None:
                match_result['shape']=(imsrc.shape[1],imsrc.shape[0])#0为高，1为宽
            y = match_result['rectangle'][0][1]

            img = Image.open(imgsrc)
            cropped = img.crop((0, 1575, img.size[0], y))  # (left, upper, right, lower)
            basic_info_name = self.num + '_'+ self.company_code + '_basic_info.png'
            basic_info_path = self.output_file_path + '\\' + basic_info_name
            cropped.save(basic_info_path)

            pic_code='<table><img src=\''+basic_info_path+'\'width=\'300\' height=\'400\'>'
            pic_code_path = self.output_file_path + r'\0_basicinfo_piccode.txt' 
            with open(pic_code_path,'a') as f2:
                f2.write(pic_code+'\n')

        except Exception as e:
            print('\n文件路径中不能存在中文字符！\n',e)

    def main(self, num, search_key):
        self.num, self.search_key = str(num), search_key
        self.get_company_code()
        snapshot_img_path = self.download_full_report()

        log_path = self.output_file_path + r'\log.txt'
        f1 = open(log_path,'a')
        if snapshot_img_path:
            self.crop_basic_info(snapshot_img_path)
            f1.write(self.num + '_' + self.search_key + ' 的企业信用报告已成功下载\n')
        else:
            f1.write(self.num + '_' + self.search_key + ' 的企业信用报告下载失败\n')

        sleep_time = round(random.uniform(0.5, 3), 2)
        time.sleep(sleep_time)
        if num != 1 and num % 20 == 0:
            print('为减小网站压力，每获取20个信息，休息30秒再继续')
            time.sleep(30)
        else:
            pass

def load_company():
    '''加载公司列表'''
    try:
        now_path = os.path.abspath(os.curdir)
        input_xls = now_path+r'\Input\company.xlsx'
        d = pd.read_excel(input_xls)
        df_li = d.values.tolist()
        company_list = [i[1] for i in df_li]
        if len(company_list) == 0:
            print('company.xlsx中未加载到任何公司名称')
        else:
            print('已成功加载{}个待查询公司名称'.format(len(company_list)))
            return company_list
    except:
        print('加载公司列表失败！')
    
def progress_bar(i,start_time,total=100,width=50):
    percent=width / total
    p=int(i*percent)
    t = str(round(time.time() - start_time, 2))
    bl= "["+str(int(p * 100 / width)) + "%"+"]"+"["+ str(i) +"/" + str(total) +"]"+"(time:"+ t +"s)"
    a ="下载进度: [" + "#" * p + "_" * (width-p) + "]" + bl
    sys.stdout.write("\r%s" %a)
    sys.stdout.flush()
    time.sleep(0.1)

if __name__ == "__main__":

    print('企业信用报告（图片版）批量下载工具','数据来源 / 天眼查 + 国家企业信用信息公示系统','制作者 / 王文铖','微信公众号 / 15Seconds','建议一次查询在20条以内，过于频繁会对IP有限制',sep = '\n',end = '\n\n')
    start_time = time.time()
    Q = Qiyekuaizhao()
    search_key_list = load_company()
    for num, search_key in enumerate(search_key_list, 1):
        progress_bar(num, start_time, len(search_key_list))
        Q.main(num, search_key)

    time.sleep(2)
