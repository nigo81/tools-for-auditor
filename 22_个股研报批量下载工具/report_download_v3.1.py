import csv
import datetime
import json
import logging
import logging.handlers
import os
import random
import re
import time
import urllib
from urllib.parse import quote
from urllib.request import urlretrieve

import execjs
import pandas as pd
import prettytable as pt
import requests

Tools_Info=['研报批量下载工具','东方财富Choice数据','2020/02/09','3.0','']
User_Agent_List = [
            # Opera
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36 OPR/26.0.1656.60",
            "Opera/8.0 (Windows NT 5.1; U; en)",
            "Mozilla/5.0 (Windows NT 5.1; U; en; rv:1.8.1) Gecko/20061208 Firefox/2.0.0 Opera 9.50",
            "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; en) Opera 9.50",
            # Firefox
            "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:34.0) Gecko/20100101 Firefox/34.0",
            "Mozilla/5.0 (X11; U; Linux x86_64; zh-CN; rv:1.9.2.10) Gecko/20100922 Ubuntu/10.10 (maverick) Firefox/3.6.10",
            # Safari
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/534.57.2 (KHTML, like Gecko) Version/5.1.7 Safari/534.57.2",
            # chrome
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.71 Safari/537.36",
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11",
            "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/534.16 (KHTML, like Gecko) Chrome/10.0.648.133 Safari/534.16",
            # 360
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/30.0.1599.101 Safari/537.36",
            "Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; rv:11.0) like Gecko",
            # 淘宝浏览器
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.11 (KHTML, like Gecko) Chrome/20.0.1132.11 TaoBrowser/2.0 Safari/536.11",
            # 猎豹浏览器
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/21.0.1180.71 Safari/537.1 LBBROWSER",
            "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E; LBBROWSER)",
            "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; QQDownload 732; .NET4.0C; .NET4.0E; LBBROWSER)",
            # QQ浏览器
            "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E; QQBrowser/7.0.3698.400)",
            "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; QQDownload 732; .NET4.0C; .NET4.0E)",
            # sogou浏览器
            "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.84 Safari/535.11 SE 2.X MetaSr 1.0",
            "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; Trident/4.0; SV1; QQDownload 732; .NET4.0C; .NET4.0E; SE 2.X MetaSr 1.0)",
            # maxthon浏览器
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Maxthon/4.4.3.4000 Chrome/30.0.1599.101 Safari/537.36",
            # UC浏览器
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/38.0.2125.122 UBrowser/4.0.3214.0 Safari/537.36",
        ]

class Template(object):

    def __init__(self):
        self.now_time = datetime.datetime.strftime(datetime.datetime.now(), '%Y%m%d_%H-%M-%S')
        self.now_path = os.path.abspath(os.curdir)
        self.__ceart_folder()
        self.__logging()
        self.__print_tool_info()

    def __print_tool_info(self):
        # 打印程序信息
        tb = pt.PrettyTable()
        tb.field_names = ['名称', str(Tools_Info[0])]
        tb.add_row(['作者', '王文铖'])
        tb.add_row(['微信公众号', '15Seconds'])
        tb.add_row(['数据来源', Tools_Info[1]])
        tb.add_row(['更新时间', Tools_Info[2]])
        tb.add_row(['版本号', Tools_Info[3]])
        tb.add_row(['使用方法', '关注微信公众号 15Seconds,查看历史文章'])
        tb.add_row(['GitHub项目地址', 'https://github.com/nigo81/tools-for-auditor'])
        print(tb)

    def __ceart_folder(self):
        # 创建本次输出的文件环境
        result_path = self.now_path+'\\Output\\' + self.now_time
        # img_path = result_path+r'\img'
        # pdf_path = result_path+r'\pdf'
        # xls_path = result_path+r'\xls'
        # self.path_list = [result_path, img_path, pdf_path, xls_path]
        self.path_list = [result_path]
        [os.makedirs(path) for path in self.path_list if not os.path.exists(path)]

    def __logging(self):
        # 用于记录日志  logger_error用于记录error，logger用于print
        log_path=self.path_list[0]+r"\log.log"
        logging.basicConfig(filename = log_path,
                            filemode = "w",
                            format = "%(asctime)s %(name)s:%(levelname)s:%(message)s",
                            datefmt = "%d-%M-%Y %H:%M:%S",
                            level=logging.INFO)
        handler = logging.StreamHandler()
        handler.setLevel(logging.INFO)
        formatter = logging.Formatter("%(asctime)s %(name)s %(levelname)s %(message)s")
        handler.setFormatter(formatter)
        self.logger_error = logging.getLogger(" Error_log ")
        self.logger = logging.getLogger('['+ str(Tools_Info[0])+' '+str(Tools_Info[3])+']')
        self.logger.addHandler(handler)

    def __write_to_csv(self, output_path, value=[]):
        with open(output_path, 'a', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(value)

    def __csv_to_xlsx(self, output_path):
        # 转换为xlsx
        csv = pd.read_csv(output_path, encoding='utf-8-sig')
        csv.to_excel(output_path.replace('.csv', '.xlsx'), sheet_name='15Seconds')

    def __read_txt(self, input_path):
        # 加载cookies或者api_key
        with open(input_path, 'r') as f:
            content = f.read()
        f.close()
        return content

    def __read_xlsx(self, input_path):
        try:
            file_name = input_path.split('\\')[-1]
            df = pd.read_excel(input_path,sheet_name=int(self.function_select)-1).fillna('')
            if self.function_select =='1':
                code_list = df['股票代码'].tolist()
                code_list = [str(c).zfill(6) for c in code_list]
            elif self.function_select =='2':
                code_list = df['行业代码'].tolist()
                code_list = [str(c).zfill(3) for c in code_list]
            elif self.function_select =='3':    
                code_list = df['机构代码'].tolist()
                code_list = [str(c).zfill(8) for c in code_list]
            start_time_list = df['查询起始日期'].apply(lambda x:x.strftime('%Y-%m-%d')).tolist()
            end_time_list = df['查询结束日期'].apply(lambda x:x.strftime('%Y-%m-%d')).tolist()
            if len(code_list) == 0:
                self.logger.info('{} 中未发现公司名称'.format(file_name))
            return code_list,start_time_list,end_time_list
        except:
            self.logger.info('INPUT文件读取错误，检查 {} 工作簿是否存在'.format(file_name))

    def checkNameValid(self,name=None):
        reg = re.compile(r'[\\/:*?"<>|\r\n]+')
        valid_name = reg.findall(name)
        if valid_name:
            for nv in valid_name:
                name = name.replace(nv, "_")
        return name

    def set_function(self):
        fuction_select = input('>>请选择需要执行的功能?\n    1-个股研报下载\n    2-行业研报下载\n    3-机构研报下载\n\n>>在下方输入功能对应的数字，按Enter确认，默认功能1\n')
        fuction_select = fuction_select.replace(' ','')
        if fuction_select == '2' or fuction_select == '3':
            func = fuction_select
        else:
            func ='1'
        self.logger.info('开始执行功能 {}'.format(func))
        return func

    def get_one_page(self,code,pageNo=1,beginTime='',endTime=''):
        if self.function_select =='1':
            base_url = 'http://reportapi.eastmoney.com/report/list?'
        elif self.function_select =='2':
            base_url = 'http://reportapi.eastmoney.com/report/list?'
        elif self.function_select =='3':    
            base_url = 'http://reportapi.eastmoney.com/report/dg?'

        # JS逆向1
        cb = execjs.eval("'datatable' + Math.floor(Math.random() * 10000000 + 1)")
        timestr = str(round(time.time() * 1000))
        data = {
            'cb': cb,
            'pageNo': pageNo,
            'pageSize': '50',
            'industryCode': '*',
            'industry': '*',
            'rating': '*',
            'ratingchange': '*',
            'beginTime': beginTime,
            'endTime': endTime,
            'fields': '',
            'qType': '0',
            '_': timestr,
        }
        if self.function_select =='1':
            code_dic = {'code': code}
            data.update(code_dic)
            # JS逆向2
            if code[0]==0 or code[0]==2 or code[0]==3:
                codeid = '0.'+code
            else:
                codeid = '1.'+code
            headers = {
                'Referer': 'http://data.eastmoney.com/report/singlestock.jshtml?quoteid={}&stockcode={}'.format(codeid, code),
                'User-Agent': random.choice(User_Agent_List),
            }
        elif self.function_select =='2':
            code_dic = {'industryCode': code,
                        'orgCode':'',
                        'rcode':'',
                        'qType': '1'
                        }
            data.update(code_dic)
            headers = {
                'Referer': 'http://data.eastmoney.com/report/industry.jshtml',
                'User-Agent': random.choice(User_Agent_List),
            }
        elif self.function_select =='3':
            data = {
                'cb': cb,
                'pageNo': pageNo,
                'pageSize': '50',
                'author': '*',
                'orgCode': code,
                'beginTime': beginTime,
                'endTime': endTime,
                'qType': '',
                '_': timestr,
                }
            headers = {
                'Referer': 'http://data.eastmoney.com/report/orgpublish.jshtml?orgcode={}'.format(code),
                'User-Agent': random.choice(User_Agent_List),
            }
        query_string_parameters = urllib.parse.urlencode(data)
        url = base_url + query_string_parameters
        # print(url)
        r = requests.get(url, headers=headers)
        r = r.text[(len(cb)+1):-1]
        result = json.loads(r)

        return result

    def download_reports(self,keys):
        # 用代码和时间戳创建文件夹
        code,beginTime,endTime=keys[0],keys[1],keys[2]
        file_path = self.path_list[0]+'\\'+ self.checkNameValid(('_').join([code,beginTime,endTime]))
        if not os.path.exists(file_path):
            os.makedirs(file_path)
        # 判断总页数和总个数
        result = self.get_one_page(code,1,beginTime,endTime)
        hits = result['hits']
        self.logger.info('[{}]成功获取{}个研报'.format(code,hits))
        TotalPage = result['TotalPage']
        # 循环下载每页报告
        if hits > 0 :
            data_list=result['data']
            for page in range(1,TotalPage+1):
                if page * 50 < hits:
                    num = 50
                else:
                    num = hits - (page-1)*50
                if page >1:
                    result = self.get_one_page(code,page,beginTime,endTime)
                    data_list = data_list + result['data']
                for i in range(num):
                    try:
                        d = result['data'][i]
                        title, infoCode, orgSName, publishDate, researcher = d['title'], d['infoCode'], d['orgSName'], d['publishDate'][:10], d['researcher']
                        dl_url = 'http://pdf.dfcfw.com/pdf/H3_{}_1.pdf'.format(infoCode)
                        pdf_name_0 = ('_').join([publishDate, orgSName, title, researcher]) + '.pdf'
                        pdf_name = file_path+'\\'+ self.checkNameValid(pdf_name_0)
                        urllib.request.urlretrieve(dl_url, pdf_name)
                        self.logger.info('[{}]当前下载进度：{}/{}'.format(code,i+(page-1)*50+1,hits))
                        name_0 = ('_').join([publishDate, orgSName, title]) + '.pdf'
                        self.logger.info('当前研报文件名：{}'.format(name_0))
                    except:
                        pass

            try:
                pf = pd.DataFrame(data_list)
                pf.fillna('', inplace=True)
                pf = pf.ix[:, ::-1]
                file_path=file_path+'\\'+'{}_reports_data.xlsx'.format(code)
                pf.to_excel(file_path, encoding='utf-8', index=False)
                self.logger.info('[{}]研报汇总表格数据已生成'.format(code))
            except Exception as e:
                self.logger.info(e)
        else:
            pass

    def main(self):
        self.function_select = self.set_function()
        input_path = self.now_path+r'\Input\report_code_input.xlsx'
        code_list,start_time_list,end_time_list = self.__read_xlsx(input_path)
        # print(code_list,start_time_list,end_time_list)
        for i in range(len(code_list)):
            keys=[code_list[i],start_time_list[i],end_time_list[i]]
            keys = [str(k).replace(' ','') for k in keys]
            self.download_reports(keys)
            self.logger.info('[{}]研报已下载完成\n'.format(code_list[i]))

    def __del__(self):
        pass

if __name__ == "__main__":
    code_example = Template()
    code_example.main()
    del code_example
