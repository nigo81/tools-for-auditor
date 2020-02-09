import requests
import time
import json
import random
import csv
import prettytable as pt
import os


def WelCome():
    # 打印程序描述
    tb = pt.PrettyTable()
    tb.field_names = ['名称', '专利批量查询工具1.0']
    tb.add_row(['作者', '王文铖'])
    tb.add_row(['微信公众号', '15Seconds'])
    tb.add_row(['', ''])
    tb.add_row(['', '关注微信公众号'])
    tb.add_row(['使用方法', '可查看工具详细使用方法'])
    tb.add_row(['', '还能获取更多实用工具'])
    tb.add_row(['', ''])
    tb.add_row(['GitHub项目地址', 'https://github.com/nigo81/tools-for-auditor'])
    tb.add_row(['', ''])
    tb.add_row(['更新时间', '2019-11-01'])
    print(tb)
    print('\n')

def get_query_result(searchword):
    url = 'https://www.tiikong.com/patent/queryresult/getList.do?page=1&pageSize=10&search=(检索关键字:{})&searchType=IntelligentSearch'.format(
        searchword)
    with open('cookies.txt', 'r') as f:
        cookies = f.read()
    f.close()
    user_agent_list = [
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
    headers = {'User-Agent': random.choice(user_agent_list),
               'Host': 'www.tiikong.com',
               'Connection': 'keep-alive',
               'Cache-Control': 'max-age=0',
               'Upgrade-Insecure-Requests': '1',
               'Sec-Fetch-User': '?1',
               'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
               'Sec-Fetch-Mode': 'navigate',
               'Sec-Fetch-Site': 'none',
               'Accept-Encoding': 'gzip, deflate, br',
               'Accept-Language': 'zh-CN,zh;q=0.9',
               'Cookie': cookies,
               }
    response = requests.get(url, headers=headers)
    text = json.loads(response.text)
    return text

def get_patent_detail(text):
    data = text['data'][0]
    data_key = [k for k in data.keys()]

    key_list = ['patentNo', 'pubNo', 'lastLegalStatus', 'issueDate', 'appDate', 'pubDate', 'appCountry', 'patentHoldersuniqueName', 'patentHolderscountry',
                'inventorsuniqueName', 'abstracttextCN', 'abstracttextEN']

    value_list = []
    for i in key_list:
        if i in data_key:
            value = data[i]
        else:
            value = ''
        value_list.append(value)
    for i in range(0,2):
        value_list[i] = value_list[i]+'\t'#将长数字加上制表符
    value_list[1] = value_list[1].replace("<span class='highLightKeyword'>","").replace('</span>','')
    for i in range(2,5):
        value_list[i] = value_list[i][:10]#处理时间
    for i in range(7,10):
        value_list[i] = ','.join(value_list[i]) #将列表转化位字符串
    value_list = value_list[:10] + [value_list[10]+value_list[11]]#将中英文摘要合并
    return value_list

def write_to_csv(output_path,value_list):
    with open(output_path, 'a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(value_list)

def read_searchword_list(input_path):
    with open(input_path, 'r') as f:
        searchword_list = f.readlines()
        searchword_list = [a.replace(' ', '').replace(
            '\n', '') for a in searchword_list if a.replace(' ', '').replace('\n', '') != '']
    f.close()
    return searchword_list

if __name__ == "__main__":
    nowpath = os.path.abspath(os.curdir)
    nowtime = time.strftime('%Y%m%d%H%M%S',time.localtime(time.time()))
    output_path = nowpath+r'\output-{}.csv'.format(nowtime)
    input_path = nowpath+r'\input.txt'

    WelCome()
    data_name = ['搜索关键词','申请号', '公开号', '最新法律状态', '专利授权日期', '申请日期', '公告日期', '申请人国家',
                 '专利权人', '专利权人国家', '发明人',  '摘要']
    write_to_csv(output_path,data_name)
    searchword_list = read_searchword_list(input_path)
    for searchword in searchword_list:
        try:
            text = get_query_result(searchword)
            value_list = get_patent_detail(text)
            value_list = [searchword + '\t'] + value_list
            time.sleep(0.5)
            write_to_csv(output_path,value_list)
            [print(str(data_name[i])+':'+str(value_list[i]))
            for i in range(len(data_name))]
            print('\n')
        except:
            print('查询 {} 的信息时出错！尝试下一个'.format(searchword))



