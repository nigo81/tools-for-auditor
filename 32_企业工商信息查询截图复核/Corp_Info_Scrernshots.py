import csv
import datetime
import json
import logging
import logging.handlers
import os
import random
import re
import time
from urllib.parse import quote
from urllib.request import urlretrieve

import pandas as pd
import prettytable as pt
import requests
from lxml import etree
from PIL import Image
from selenium import webdriver

# 全局常量
Tools_Info=['企业工商信息截图工具','','2020-01-09','v1.0','']
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
Now_Time = datetime.datetime.strftime(datetime.datetime.now(), '%Y%m%d_%H-%M-%S')
Now_Path = os.path.abspath(os.curdir)

class Corp_Info_Scrernshots(object):

    def __init__(self):
        self.__ceart_folder()
        self.__logging()
        self.__print_tool_info()
        self.search_url = ''
        self.company_url = ''
        self.company_pos = ''
        # self.html
        self.base_headers = {
            'Host': 'www.qichamao.com',
            'Connection': 'keep-alive',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'zh-CN,zh;q=0.9',
        }

    def __print_tool_info(self):
        # 打印程序信息
        tb = pt.PrettyTable()
        tb.field_names = ['名称', str(Tools_Info[0])]
        tb.add_row(['作者', '王文铖'])
        tb.add_row(['微信公众号', '15Seconds'])
        # tb.add_row(['数据来源', Tools_Info[1]])
        tb.add_row(['更新时间', Tools_Info[2]])
        tb.add_row(['版本号', Tools_Info[3]])
        tb.add_row(['使用方法', '关注微信公众号 15Seconds,查看历史文章'])
        tb.add_row(['GitHub项目地址', 'https://github.com/nigo81/tools-for-auditor'])
        print(tb)

    def __ceart_folder(self):
        # 创建本次输出的文件环境
        self.result_path = Now_Path+'\\Output\\' + Now_Time
        os.makedirs(self.result_path)

    def __logging(self):
        # 用于记录日志  logger_error用于记录error，logger用于print
        log_path=self.result_path+r"\log.log"
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

    def __cookies_load(self):
        # 加载cookies或者api_key
        filepath = Now_Path+r'\config\cookies.txt'
        cookies = (open(filepath, 'r').read()).strip()
        if len(cookies) < 50:
            self.logger.info('请先在cookies.txt中填写cookies信息')
            return None
        else:
            cookies = cookies[0:-10]+str(int(time.time() * 1000))
            return cookies

    def authorization(self):
        # 获取授权
        self.logger.info('尝试获取15Seconds微信公众号的授权')
        with open(Now_Path + r'\config\authorization.config','r') as f:
            authorization_url = f.read()
        response = requests.get(authorization_url)
        response_json = response.json()
        authorization_info = response_json['clientVars']['padTitle']
        if authorization_info == '15Seconds微信公众号授权成功':
            return True
        else:
            self.logger.info(authorization_info)
            return False

    def read_xlsx(self):
        # 读取公司列表
        try:
            input_xls = Now_Path+r'\Input\company_input.xlsx'
            d = pd.read_excel(input_xls).fillna('0')#########################
            df = d.values.tolist()
            df_li =[i for i in df if i!=['0','0','0']]
            company_list = [str(i[1]) for i in df_li]
            try:
                companypos_list = [i[2] for i in df_li]
            except:
                companypos_list = ['0']*(len(company_list))
            self.logger.info('部分数据预览如下：')
            self.logger.info(company_list[:5])
            if len(company_list) == 0:
                self.logger.info('company_input.xlsx中未加载到任何公司名称')
            else:
                self.logger.info('已成功加载{}个待查询公司名称'.format(len(company_list)))
            return company_list,companypos_list
        except:
            self.logger.info('加载公司列表失败！')
            self.logger.exception("正在记录发生的错误")

    def get_company_url(self):
        # 获取公司对应的链接
        self.search_url = 'https://www.qichamao.com/search/all/{}?o=0&area=0&p=1'.format(quote(self.company_name))
        headers1 = {'Cache-Control': 'max-age=0',
                    'Cookie': self.__cookies_load(),
                    'Upgrade-Insecure-Requests': '1',
                    'User-Agent': random.choice(User_Agent_List), }
        headers1.update(self.base_headers)
        try:
            r = requests.get(self.search_url, headers = headers1)
            text = r.text
            html = etree.HTML(text)
            company_url_0 = html.xpath('//a[@class="listsec_tit"]/@href')[0]
            self.company_url = 'https://www.qichamao.com{}'.format(company_url_0)
            self.logger.info('{} 的网页链接获取成功'.format(self.company_name))
            self.company_id = (self.company_url.split('/')[-1]).split('.')[0]
        except Exception as e:
            self.company_url = ''
            self.logger.info('{} 的网页链接获取失败'.format(self.company_name))
            self.logger.info(e)

    def get_page_source(self):
        # 获取页面源码
        headers2 = {'Referer': self.search_url,
                    'Cookie': self.__cookies_load(),
                    'Upgrade-Insecure-Requests': '1',
                    'User-Agent': random.choice(User_Agent_List), }
        headers2.update(self.base_headers)
        try:
            r = requests.get(self.company_url, headers=headers2)
            text = r.text
            self.html = etree.HTML(text)
        except:
            self.html = ''
            self.logger.info('{} 的页面源码获取失败，已记录在error_log.txt中'.format(self.company_name))
            error_log = self.result_path + '\\error_log.txt'
            c = '{} 的页面源码获取失败'.format(self.company_name)
            with open(error_log, "a") as f:
                f.write(c +'\n')

    def get_basic(self):
        # 获取公司基本信息
        self.logger.info('{} 的基本工商信息正在获取'.format(self.company_name))
        try:
            company_name_qcm = self.html.xpath('//div[@class="t"]/h1/text()')[0]
            if self.company_name == company_name_qcm:
                match = '一致'
            else:
                match = '名称不一致'
            info = self.html.xpath('//div[@class="qd-table-body li-half f14"]/ul[1]/li/div')
            info = [i.xpath('string(.)').strip() for i in info]
            info = [str(i)+'\t' for i in info]
            basic_info = [self.company_name,company_name_qcm, match] + info
        except:
            self.logger.exception("正在记录发生的错误")
            basic_info = [self.company_name] +['-']*16  
        return basic_info

    def get_cyxx(self):
        # 获取成员信息
        url = 'https://www.qichamao.com/orgcompany/SearchItemCYXX'
        headers3 = {
            'Accept': '*/*',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'Content-Length': '53',
            'Origin': 'https://www.qichamao.com',
            'Referer': self.company_url,
            'X-Requested-With': 'XMLHttpRequest',
            'Cookie': self.__cookies_load(),
            'User-Agent': random.choice(User_Agent_List),
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
                    for i in range(len(dataList))][::-1]
            cyxx_info = ('+').join(data)
            if len(cyxx_info) == 0:
                cyxx_info = '无'
        except:
            cyxx_info = '无'
            self.logger.exception("正在记录发生的错误")
        return cyxx_info

    def get_fzjg(self):
        # 获取分支机构信息
        try:
            names = self.html.xpath('//div[@id="M_fzjg"]/div[2]/div/div/ul[2]/li[2]/span/a/text()')
            fzjg_info = ('+').join(names)
            if len(fzjg_info) == 0:
                fzjg_info = '无'
        except:
            fzjg_info = '无'
            self.logger.exception("正在记录发生的错误")
        return fzjg_info

    def get_share(self):
        # 获取股东信息
        try:
            names = self.html.xpath('//div[@id="M_gdxx"]/div[2]/div/div/ul/li[2]/span/a/text()')
            rj = self.html.xpath('//div[@id="M_gdxx"]/div[2]/div/div/ul/li[3]/span[2]/text()')
            bl = self.html.xpath('//div[@id="M_gdxx"]/div[2]/div/div/ul/li[4]/span[2]/text()')
            sj= self.html.xpath('//div[@id="M_gdxx"]/div[2]/div/div/ul/li[5]/span[2]/text()')
            if len(names) != 0:
                shareinfo = [names[i]+'|'+rj[i]+'|'+bl[i]+'|'+sj[i]
                             for i in range(len(names))]
                share_info = ('+').join(shareinfo)
            else:
                share_info = '无'
        except:
            share_info = '无'
            self.logger.exception("正在记录发生的错误")
        return share_info

    def get_dwtz(self):
        # 获取对外投资信息
        try:
            names = self.html.xpath('//div[@id="M_dwtz"]/div[2]/div/div/ul/li[2]/span/a/text()')
            farens = self.html.xpath('//div[@id="M_dwtz"]/div[2]/div/div/ul/li[3]/span[2]/a/text()')
            zczbs = self.html.xpath('//div[@id="M_dwtz"]/div[2]/div/div/ul/li[4]/span[2]/text()')
            czbls = self.html.xpath('//div[@id="M_dwtz"]/div[2]/div/div/ul/li[5]/span[2]/text()')
            czbls = [(czbl.replace('\n', '')).strip() for czbl in czbls]
            dwtzinfo = [names[i]+'|'+farens[i]+'|'+zczbs[i]+'|'+czbls[i]for i in range(len(names))]
            dwtz_info = ('+').join(dwtzinfo)
            if len(dwtz_info) == 0:
                dwtz_info = '无'
        except:
            dwtz_info = '无'
            self.logger.exception("正在记录发生的错误")
        return dwtz_info

    def Output_csv(self, content=''):
        # 输出文件
        output_path = self.result_path + r'\company_info.csv'
        with open(output_path, 'a', newline='', encoding='utf-8-sig')as f:
            writer = csv.writer(f)
            writer.writerow(content)

    def csv_to_xlsx(self):
        # 转换为xlsx
        output_path = self.result_path + r'\company_info.csv'
        csv = pd.read_csv(output_path, encoding='utf-8-sig',error_bad_lines=False)
        csv.to_excel(output_path.replace('.csv', '.xlsx'), sheet_name='15Seconds')

    def set_browser_parameters(self):
        # 设置浏览器参数
        if 'PROGRAMFILES(X86)' in os.environ:
            executable_path = Now_Path+r'\Driver\geckodriver-win64.exe'
        else:
            executable_path = Now_Path+r'\Driver\geckodriver-win32.exe'
        firefox_options = webdriver.FirefoxOptions()
        firefox_options.add_argument('--headless')
        self.browser = webdriver.Firefox(
            executable_path = executable_path,
            firefox_options = firefox_options,
            service_log_path = self.result_path+r'\geckodriver.log'
        )

    def get_screenshot(self,num):
        # 开始截图
        self.browser.set_window_size(1000,1000)
        self.browser.get(self.company_url)
        # target = self.browser.find_element_by_xpath('//div[@id="M_jbxx1"]')
        # self.browser.execute_script("arguments[0].scrollIntoView();", target)
        js = 'window.scrollBy(0,320)'
        self.browser.execute_script(js)
        time.sleep(0.5)
        screenshot_path = self.result_path + '\\{}_{}.png'.format(num,self.company_name)
        self.browser.save_screenshot(screenshot_path)
        self.logger.info('{} 的工商截图已获取'.format(self.company_name))
        return screenshot_path

    def crop_png(self,screenshot_path):
        # 裁剪截图
        im = Image.open(screenshot_path)
        im1 = im.crop((0, 120, im.size[0]-140, im.size[1]))
        im1.save(screenshot_path)
        self.logger.info('{} 的工商截图已裁剪完成'.format(self.company_name))
    
    def get_myAK_maxjl(self):
        try:
            with open(Now_Path + r'\config\baidumap_myAK.config','r') as f:
                myAK = f.read()
            with open(Now_Path + r'\config\预警距离阈值(公里)设置.txt','r') as f:
                maxjl = f.read()
        except:
            myAK = 'Nwpg67DilX6ljmFb2QLd78nkKO7nop12'
            maxjl = 5
        return myAK,maxjl

    def cxzb(self,address,myAK):
        # 查询坐标
        url = 'http://api.map.baidu.com/geocoding/v3/?address={}&output=json&ak={}&callback=showLocation'.format(address,myAK)
        r = requests.get(url)
        data = (r.text).replace('showLocation&&showLocation(','')
        data = json.loads(data[:-1])
        lng,lat = data['result']['location']['lng'],data['result']['location']['lat']
        zb = str(lat) + ',' + str(lng)
        return zb

    def cxjl(self,pos1,pos2,myAK,maxjl):
        # 查询距离
        url = 'http://api.map.baidu.com/routematrix/v2/driving?output=json&origins={}&destinations={}&ak={}'.format(pos1,pos2,myAK)
        r = requests.get(url)
        data = json.loads(r.text)
        jl_m = data['result'][0]['distance']['value']
        jl_gl = round(jl_m/1000, 2)
        if float(jl_gl) <= float(maxjl):
            jieguo = '未超出'
        else:
            jieguo = '超出'
        return jl_gl, jl_m, jieguo

    def single_search(self,i):
        try:
            juli_info = ['-']*4
            result = [self.company_name] +['-']*26 
            self.get_company_url()
            self.get_page_source()
            if self.html != '':
                basic_info = self.get_basic()
                cyxx_info = self.get_cyxx()
                fzjg_info = self.get_fzjg()
                share_info = self.get_share()
                dwtz_info = self.get_dwtz()
                self.logger.info('{} 的基本工商信息获取成功'.format(self.company_name))
                result = basic_info + [cyxx_info,fzjg_info,share_info,dwtz_info]
                # 截图模块
                if self.function_id =='2' or self.function_id =='4':
                    screenshot_path = self.get_screenshot(i+1) 
                    self.crop_png(screenshot_path)
                # 测距模块
                if (self.function_id =='3' or self.function_id =='4'):
                    if self.company_pos !='0':
                        myAK,maxjl = self.get_myAK_maxjl() 
                        address1 = basic_info[17] 
                        pos1 = self.cxzb(address1,myAK)
                        address2 = self.company_pos
                        pos2 = self.cxzb(address2,myAK)
                        jl_gl, jl_m, jieguo = self.cxjl(pos1, pos2,myAK,maxjl)
                        juli_info = [address2, str(jl_gl)+'公里', jl_m, jieguo]
                        self.logger.info('{} 的地址复核成功'.format(self.company_name))
                    else:
                        self.logger.info('{} 的地址未填，无法复核'.format(self.company_name))
                result =result + juli_info
            self.Output_csv(result)
            # 写入
            
        except Exception as e:
            self.logger.info('{} 的基本工商信息获取失败,继续下一个'.format(self.company_name))
            self.logger.exception(e)
            error_log = self.result_path + '\\error_log.txt'
            with open(error_log, "a") as f:
                f.write(self.company_name+'\n')

    def main(self):
        # 主程序
        start_time = time.time()
        if self.authorization():
            self.logger.info('15Seconds微信公众号授权成功')
            id = input('''  \n请选择需要执行功能对应数字，默认第1个功能 \n  1、工商信息基础表格\n  2、工商信息基础表格+截图\n  3、工商信息基础表格+地址核对\n  4、工商信息基础表格+截图+地址核对\n\n ''')
            self.function_id = id.strip()
            self.logger.info('开始执行功能{}'.format(id))
            company_list,companypos_list= self.read_xlsx()
            self.set_browser_parameters()  # 截图模块-加载浏览器
            title = ['公司名称', '搜索结果', '名称是否一致',  '法定代表人', '纳税人识别号', '名称', '机构代码',
                 '注册号', '注册资本', '统一社会信用代码', '登记机关', '经营状态','成立日期', '企业类型', 
                 '经营期限', '所属地区', '核准日期', '注册地址', '经营范围',
                 '成员信息', '分支机构', '股东信息', '对外投资'] + ['企业提供地址', '距离（公里）', '距离（数值）', '是否超出阈值']
            self.Output_csv(title)
            for i in range(len(company_list)):
                self.company_name = company_list[i]
                if self.function_id =='3' or self.function_id =='4':
                    self.company_pos = companypos_list[i] # 测距模块
                self.logger.info('当前进度：{}/{} {}'.format(i+1,len(company_list),self.company_name))
                self.single_search(i) # 单次查询
                t = round(random.uniform(2, 7), 2)
                self.logger.info('---为友好可持续的获取数据，等待 {} 秒---'.format(t))
                time.sleep(t)  
            self.csv_to_xlsx()
        else:
            self.logger.info('15Seconds微信公众号授权失败')
        stamp = time.time() - start_time
        self.logger.info('查询总用时：{} 秒'.format(str(stamp)))

    def __del__(self):
        try:
            self.browser.quit()
        except:
            pass


if __name__ == "__main__":
    code_example = Corp_Info_Scrernshots()
    code_example.main()
    del code_example
