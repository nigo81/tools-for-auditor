# -*- coding: utf-8 -*-
# Author: wwcheng
import csv
import logging
import logging.handlers
import os
import random
import time
from urllib import request

import numpy as np
import pandas as pd
import prettytable as pt
import requests
import xlwt
from fake_useragent import UserAgent
from lxml import etree


class QiChaMao(object):
    def __init__(self):
        self.start_time = time.time()
        self.timestr = str(int(time.time() * 1000))
        self.nowpath = os.path.abspath(os.curdir)
        self.search_key = ''
        self.company_url = ''
        self.company_id = ''
        self.company_name = ''
        self.basic_info = []
        self.cyxx_info = ''
        self.fzjg_info = ''
        self.share_info = ''
        self.dwtz_info = ''
        self.company_list = []
        self.base_headers = {
            'Host': 'www.qichamao.com',
            'Connection': 'keep-alive',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'zh-CN,zh;q=0.9',
        }
        self.__logging()

    def __cookies_load(self):
        '''内部函数：加载cookies，切换时间戳'''
        filepath = self.nowpath+r'\Input\cookies.txt'
        cookies = (open(filepath, 'r').read()).strip()
        if len(cookies) < 50:
            self.logger.info('请先在cookies.txt中填写cookies信息')
            return None
        else:
            cookies = cookies[0:-10]+str(int(time.time() * 1000))
            return cookies

    def __logging(self):
        '''内部函数；用于记录日志'''
        logging.basicConfig(filename=self.nowpath+r"\Output\error.log",
                            filemode="w",
                            format="%(asctime)s %(name)s:%(levelname)s:%(message)s",
                            datefmt="%d-%M-%Y %H:%M:%S",
                            level=logging.INFO
                            )
        self.logger = logging.getLogger("15Scends")
        handler = logging.StreamHandler()
        handler.setLevel(logging.INFO)
        formatter = logging.Formatter(
            "%(asctime)s %(name)s %(levelname)s %(message)s")
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)

    def __get_html(self):
        '''内部函数：获取页面源码'''
        headers2 = {'Referer': self.search_url,
                    'Cookie': self.__cookies_load(),
                    'Upgrade-Insecure-Requests': '1',
                    'User-Agent': UserAgent().random, }
        headers2.update(self.base_headers)
        self.info_url = 'https://www.qichamao.com{}'.format(self.company_url)
        self.logger.info(self.info_url)
        r = requests.get(self.info_url, headers=headers2)
        text = r.text
        self.html = etree.HTML(text)

    def WelCome(self):
        '''打印程序描述 '''
        tb = pt.PrettyTable()
        tb.field_names = ['名称', '企业工商信息批量查询']
        tb.add_row(['作者', 'wwcheng'])
        tb.add_row(['微信公众号', '15Sceconds'])
        tb.add_row(
            ['GitHub项目地址', 'https://github.com/nigo81/tools-for-auditor'])
        tb.add_row(['数据来源', 'www.qichamao.com'])
        tb.add_row(['更新时间', '2019-08-07 13:09'])
        print(tb)

    def load_company(self):
        '''加载公司列表'''
        try:
            input_xls = self.nowpath+r'\Input\company_input.xlsx'
            d = pd.read_excel(input_xls)
            df_li = d.values.tolist()
            self.company_list = [i[1] for i in df_li]
            self.logger.info('部分数据预览如下：')
            self.logger.info(self.company_list[:3])
            if len(self.company_list) == 0:
                self.logger.info('company_input.xlsx中未加载到任何公司名称')
            else:
                self.logger.info('已成功加载{}个待查询公司名称'.format(
                    len(self.company_list)))
        except:
            self.logger.info('加载公司列表失败！')
            self.logger.exception("正在记录发生的错误")

    def get_companyid(self):
        ''''获取公司对应的链接'''
        search_key = request.quote(self.search_key)
        self.search_url = 'https://www.qichamao.com/search/all/{}?o=0&area=0&p=1'.format(
            search_key)
        headers1 = {'Cache-Control': 'max-age=0',
                    'Cookie': self.__cookies_load(),
                    'Upgrade-Insecure-Requests': '1',
                    'User-Agent': UserAgent().random, }
        headers1.update(self.base_headers)
        r = requests.get(self.search_url, headers=headers1)
        text = r.text
        html = etree.HTML(text)
        self.company_url = html.xpath('//a[@class="listsec_tit"]/@href')[0]
        self.company_id = (self.company_url.split('/')[-1]).split('.')[0]

    def get_basic(self):
        '''获取公司基本信息'''
        self.__get_html()
        self.company_name = self.html.xpath(
            '//div[@class="arthd_tit"]/h1/text()')[0]
        if self.search_key == self.company_name:
            match = '一致'
        else:
            match = '名称不一致'
        info1 = ['无']*2
        info2 = ['无']*14
        try:
            info1 = self.html.xpath(
                '//div[@class="arthd_info"]/span/text()')[:2]
            info1 = [(i.split('：'))[1] for i in info1]
            if len(info1)==1:
                info1=info1+['无']
            info2 = self.html.xpath(
                '//section[@class="pb-d2"]/ul[1]/li/span[@class="info"]/text()')
            info2 = [(i.replace('\n', '')).replace('--', '无') for i in info2]
            try:
                faren=self.html.xpath(
                    '//section[@class="pb-d2"]/ul[1]/li[6]/span[@class="info"]/a/text()')
            except:
                self.logger.exception("正在记录发生的错误")
                faren=['--']
            info2 = info2[:4]+faren+info2[-10:]  # 剔除名称
        except:
            self.logger.exception("正在记录发生的错误")
        self.basic_info = [self.search_key,
                           self.company_name, match]+info1+info2

    def get_cyxx(self):
        '''获取成员信息'''
        url = 'https://www.qichamao.com/orgcompany/SearchItemCYXX'
        headers3 = {
            'Accept': '*/*',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'Content-Length': '53',
            'Origin': 'https://www.qichamao.com',
            'Referer': self.info_url,
            'X-Requested-With': 'XMLHttpRequest',
            'Cookie': self.__cookies_load(),
            'User-Agent': UserAgent().random,
        }
        headers3.update(self.base_headers)
        try:
            hdoc_area = self.html.xpath('//input[@id="hdoc_area"]/@value')
            data = {
                'orgCode': self.company_id,
                'oc_area': hdoc_area,
            }
            r = requests.post(url, headers=headers3, data=data)
            cyxx = r.json()
            dataList = cyxx['dataList']
            data = [dataList[i]['om_name']+'|'+dataList[i]['om_position']
                    for i in range(len(dataList))]
            self.cyxx_info = ('+').join(data)
            if len(self.cyxx_info) == 0:
                self.cyxx_info = '无'
        except:
            self.logger.exception("正在记录发生的错误")

    def get_fzjg(self):
        '''获取分支机构信息'''
        try:
            names = self.html.xpath(
                '//ul[@class="list-table_td branch-art"]/li[2]/span[2]/a/text()')
            self.fzjg_info = ('+').join(names)
            if len(self.fzjg_info) == 0:
                self.fzjg_info = '无'
        except:
            self.logger.exception("正在记录发生的错误")

    def get_share(self):
        '''获取股东信息'''
        try:
            names = self.html.xpath(
                '//ul[@class="list-gd"]/li/div/div[1]/em/a/text()')
            gdlxs = self.html.xpath(
                '//ul[@class="list-gd"]/li/div/div[1]/span/text()')
            cgbls = self.html.xpath(
                '//ul[@class="list-gd"]/li/div/div[2]/p[1]/span/text()')
            shareinfo = [names[i]+'|'+gdlxs[i]+'|'+cgbls[i]
                         for i in range(len(names))]
            self.share_info = ('+').join(shareinfo)
            if len(self.share_info) == 0:
                self.share_info = '无'
        except:
            self.share_info = '无'
            self.logger.exception("正在记录发生的错误")

    def get_dwtz(self):
        '''获取对外投资信息'''
        try:
            names = self.html.xpath(
                '//ul[@class="list-table_td dwtz-art"]/li[2]/span[2]/a/text()')
            farens = self.html.xpath(
                '//ul[@class="list-table_td dwtz-art"]/li[3]/span[2]/a/text()')
            zczbs = self.html.xpath(
                '//ul[@class="list-table_td dwtz-art"]/li[4]/span[2]/text()')
            czbls = self.html.xpath(
                '//ul[@class="list-table_td dwtz-art"]/li[5]/span[2]/text()')
            czbls = [(czbl.replace('\n', '')).strip() for czbl in czbls]
            dwtzinfo = [names[i]+'|'+farens[i]+'|'+zczbs[i]+'|'+czbls[i]
                        for i in range(len(names))]
            self.dwtz_info = ('+').join(dwtzinfo)
            if len(self.dwtz_info) == 0:
                self.dwtz_info = '无'
        except:
            self.dwtz_info = '无'
            self.logger.exception("正在记录发生的错误")

    def Output_csv(self, content=''):
        '''输出文件'''
        self.output_path = self.nowpath + \
            r'\Output\company_output_{}.csv'.format(self.timestr)
        with open(self.output_path, 'a', newline='', encoding='utf-8-sig')as f:
            writer = csv.writer(f)
            writer.writerow(content)

    def csv_to_xlsx(self):
        '''转换为xlsx'''
        csv = pd.read_csv(self.output_path, encoding='utf-8-sig')
        csv.to_excel(self.output_path.replace(
            '.csv', '.xlsx'), sheet_name='data')

    def single_query(self, search_key):
        '''单次查询'''
        try:
            self.search_key = search_key
            self.get_companyid()
            self.get_basic()
            self.get_cyxx()
            self.get_fzjg()
            self.get_share()
            self.get_dwtz()
            result = self.basic_info + \
                [self.cyxx_info, self.fzjg_info,
                    self.share_info, self.dwtz_info,self.info_url]
            self.Output_csv(result)
            sleep_time = round(random.uniform(0.2, 0.6), 2)
            time.sleep(sleep_time)
        except Exception:
            self.logger.info('获取{}的信息时出错'.format(self.search_key))
            self.logger.exception("正在记录发生的错误")

    def main(self):
        '''主体循环'''
        self.WelCome()
        self.load_company()
        title = ['公司名称', '搜索结果', '名称是否一致', '企业电话', '企业邮箱', '统一社会信用代码',
                 '纳税人识别号', '注册号', '机构代码','法定代表人', '企业类型', '经营状态', '注册资本',
                 '成立日期', '登记机关', '经营期限', '所属地区', '核准日期', '企业地址', '经营范围',
                 '成员信息', '分支机构', '股东信息', '对外投资','链接']
        self.Output_csv(title)
        for i in range(len(self.company_list)):
            search_key = self.company_list[i]
            self.logger.info('当前进度：{}/{} {}'.format(i+1,
                                                    len(self.company_list), search_key))
            self.single_query(search_key)
            if i+1 != 1 and (i+1) % 30 == 0:
                self.csv_to_xlsx()
                self.logger.info('为减小网站压力，每获取{}个信息，休息30秒再继续'.format(str(i+1)))
                time.sleep(30)
            else:
                pass
        self.csv_to_xlsx()
        stamp = time.time() - self.start_time
        self.logger.info('查询总用时：'+str(stamp))

    def run(self):
        if self.__cookies_load():
            self.main()
            time.sleep(4)


if __name__ == "__main__":
    QiChaMao = QiChaMao()
    QiChaMao.run()
