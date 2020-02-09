import csv
import datetime
import json
import os
import random
import time
from urllib.parse import quote
from urllib.request import urlretrieve

import pandas as pd
import prettytable as pt
import requests
from selenium import webdriver


class Credit_China(object):

    def __init__(self):
        self.query_time = datetime.datetime.strftime(
            datetime.datetime.now(), '%Y-%m-%d %H-%M-%S')
        self.nowpath = os.path.abspath(os.curdir)
        self.__WelCome()
        self.__ceart_enviroment()
        self.__open_browser()
        self.maxjl = ''
        self.myAK = ''

    def __read_txt(self,input_path):
        with open(input_path, 'r') as f:
            content = f.read()
        f.close()
        return content

    def __WelCome(self):
        # 打印程序描述
        tb = pt.PrettyTable()
        tb.field_names = ['名称', '信用中国查询工具1.0']
        tb.add_row(['作者', '王文铖'])
        tb.add_row(['微信公众号', '15Seconds'])
        tb.add_row(['使用方法', '关注微信公众号'])
        tb.add_row(
            ['GitHub项目地址', 'https://github.com/nigo81/tools-for-auditor'])
        tb.add_row(['数据来源', 'https://www.creditchina.gov.cn/'])
        tb.add_row(['更新时间', '2019-11-07'])
        print(tb)
        print('\n')

    def __ceart_enviroment(self):
        # 创建本次输出的文件环境
        result_path = self.nowpath+'\\Output\\' + self.query_time
        img_path = result_path+r'\img'
        pdf_path = result_path+r'\pdf'
        xls_path = result_path+r'\xls'
        self.mypath = [result_path, img_path, pdf_path,xls_path]
        [os.makedirs(path) for path in self.mypath if not os.path.exists(path)]

    def __open_browser(self):
        # 打开浏览器
        if 'PROGRAMFILES(X86)' in os.environ:
            executable_path = self.nowpath+r'\Driver\geckodriver-win64.exe'
        else:
            executable_path = self.nowpath+r'\Driver\geckodriver-win32.exe'
        firefox_options = webdriver.FirefoxOptions()
        firefox_options.add_argument('--headless')

        self.browser = webdriver.Firefox(
            executable_path=executable_path,
            firefox_options=firefox_options,
            service_log_path=self.mypath[0]+r'\geckodriver.log'
        )
        self.browser.set_window_size(1400, 1000)

    def __write_to_csv(self, output_path,value_list):
        with open(output_path, 'a', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(value_list)

    def __csv_to_xlsx(self,output_path):
        '''转换为xlsx'''
        csv = pd.read_csv(output_path, encoding='utf-8-sig')
        csv.to_excel(output_path.replace('.csv', '.xlsx'), sheet_name='15Seconds')

    def load_input(self):
        try:
            input_path=self.nowpath+r'\Input\baidu_api_AK.txt'
            self.myAK = self.__read_txt(input_path)
            input_path=self.nowpath+r'\Input\预警距离阈值(公里)设置.txt'
            self.maxjl = int (self.__read_txt(input_path))
        except:
            self.myAK = 'Nwpg67DilX6ljmFb2QLd78nkKO7nop12'
            self.maxjl = 5
        try:
            input_path=self.nowpath+r'\Input\company_input.xlsx'
            df = pd.read_excel(input_path).dropna()
            company_list=df['公司名称'].tolist()
            company_list = [i.replace('）',')').replace('（','(') for i in company_list]
            companypos_list=df['企业提供地址'].tolist()
            if len(company_list) == 0:
                print('company_input.xlsx 中未发现公司名称')   
            return company_list,companypos_list
        except:
            print('公司名称读取错误，检查是否存在 company_input.xlsx 工作簿') 

    def download_img_pdf(self, company):
        company = company.replace(' ', '')
        self.xyDetail_url = r'https://www.creditchina.gov.cn/xinyongxinxixiangqing/xyDetail.html?searchState=1&entityType=1&keyword={}'.format(
            quote(company))
        self.pdf_url = r'https://public.creditchina.gov.cn/credit-check/pdf/download?companyName={}&entityType=1&uuid=&tyshxydm='.format(
            quote(company))
        try:
            self.browser.get(self.xyDetail_url)
            time.sleep(round(random.uniform(2, 3), 2))
            screenshot_img_path = self.mypath[1] + r'\{} 企业信息.png'.format(company)
            self.browser.save_screenshot(screenshot_img_path)
            credit_pdf_path = self.mypath[2] + r'\{} 信用报告.pdf'.format(company)
            urlretrieve(url=self.pdf_url, filename=credit_pdf_path)
        except Exception as e:
            print(e)
        
    def get_details(self, company):
        Details = 'https://public.creditchina.gov.cn/private-api/getTyshxydmDetailsContent?keyword={}&scenes=defaultscenario&entityType=1&searchState=1&uuid=&tyshxydm='.format(
            quote(company))
        TypeCount = 'https://public.creditchina.gov.cn/private-api/searchDateTypeCount?entityType=1&searchState=1&keyword={}'.format(
            quote(company))
        value = ['（空）']*8
        try:
            a = requests.get(Details)
            r_a = a.json()
            keys_a = ['jgmc', 'status', 'tyshxydm', 'name',
                    'enttype', 'esdate', 'regorg', 'dom']
            data = r_a['data']['headEntity']
            data_key = [k for k in data.keys()]
            for i in range(3):
                if keys_a[i] in data_key:
                    value[i] = data[keys_a[i]] 

            data = r_a['data']['data']['entity']
            data_key = [k for k in data.keys()]
            for i in range(3,8):
                if keys_a[i] in data_key:
                    value[i] = data[keys_a[i]] 
            time.sleep(2)
            # b = requests.get(TypeCount)
            # r_b = b.json()
            # keys_b = ['守信激励', '行政处罚', '行政许可']
            # [value.append(r_b['data'][i]) for i in keys_b]
            
        except Exception as e:
            print(e)
        finally:
            value[2] = str(value[2]) + '\t'
            value = [company] + value
            return value

    def cxzb(self,address):
        url = 'http://api.map.baidu.com/geocoding/v3/?address={}&output=json&ak={}&callback=showLocation'.format(address,self.myAK)
        r = requests.get(url)
        data = (r.text).replace('showLocation&&showLocation(','')
        data = json.loads(data[:-1])
        lng,lat = data['result']['location']['lng'],data['result']['location']['lat']
        zb = str(lat) + ',' + str(lng)
        return zb

    def cxjl(self,origins,destinations):
        url = 'http://api.map.baidu.com/routematrix/v2/driving?output=json&origins={}&destinations={}&ak={}'.format(origins,destinations,self.myAK)
        r = requests.get(url)
        data = json.loads(r.text)
        jl_m = data['result'][0]['distance']['value']
        jl_gl = str(round(jl_m/1000,2))
        if jl_gl <= self.maxjl:
            jieguo = ''
        else:
            jieguo = '警告'
        jl_gl = jl_gl + '公里'
        return jl_gl,jl_m,jieguo

    def main(self):
        title = ['搜索词','公司名称', '状态', '统一社会信用代码', '法定代表人', '企业类型',
                 '成立日期', '登记机关', '注册地址', '企业提供地址', '距离（公里）', '距离（数值）','是否超出阈值']
        output_path_1 = self.mypath[3] + r'\Basic_Info.csv'
        self.__write_to_csv(output_path_1,title[:9])
        output_path_2 = self.mypath[3] + r'\Address_Check.csv'
        self.__write_to_csv(output_path_2,title[:2]+title[-5:])
        company_list,companypos_list = self.load_input()
        for i in range(len(company_list)):
            print('{}/{} 正在下载 {} 的信用信息···'.format(i+1,len(company_list),company_list[i]))
            c.download_img_pdf(company_list[i])
            value = c.get_details(company_list[i])
            try:
                pos_one,pos_two=value[8],companypos_list[i]
                origins,destinations = self.cxzb(pos_one),self.cxzb(pos_two)
                jl_gl,jl_m,jieguo = self.cxjl(origins,destinations)
                value = value + [companypos_list[i],jl_gl,jl_m,jieguo]
            except:
                pass
            self.__write_to_csv(output_path_1,value[:9])
            self.__write_to_csv(output_path_2,value[:2]+value[-5:])
            print('{}/{} {} 的信用信息已经下载完成'.format(i+1,len(company_list),company_list[i]), '\n')
        self.__csv_to_xlsx(output_path_1)
        self.__csv_to_xlsx(output_path_2)

    def __del__(self):
        self.browser.quit()


if __name__ == "__main__":
    c = Credit_China()
    c.main()
    del c
