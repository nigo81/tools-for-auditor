import json
import os
import re
import sys
import time
import urllib

import execjs
import pandas as pd
import requests
from fake_useragent import UserAgent


def checkNameValid(name=None):
    reg = re.compile(r'[\\/:*?"<>|\r\n]+')
    valid_name = reg.findall(name)
    if valid_name:
        for nv in valid_name:
            name = name.replace(nv, "_")
    return name


def progress_bar(i,total=100,width=50):
    percent=width / total
    p=int(i*percent)
    bl= "["+str(int(p * 100 / width)) + "%"+"]"+"["+ str(i) +"/" + str(total) +"]"
    a ="下载进度: [" + "#" * p + "_" * (width-p) + "]" + bl
    sys.stdout.write("\r%s" %a)
    sys.stdout.flush()
    time.sleep(0.1)


def get_one_page(code,pageNo=1,beginTime='',endTime=''):
    base_url = 'http://reportapi.eastmoney.com/report/list?'
    # JS逆向1
    cd = execjs.eval("'datatable' + Math.floor(Math.random() * 10000000 + 1)")
    timestr = str(round(time.time() * 1000))
    data = {
        'cb': cd,
        'pageNo': pageNo,
        'pageSize': '50',
        'code': code,
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
    query_string_parameters = urllib.parse.urlencode(data)
    url = base_url + query_string_parameters
    # JS逆向2
    if code[0]==0 or code[0]==2 or code[0]==3:
        codeid = '0.'+code
    else:
        codeid = '1.'+code
    headers = {
        'Referer': 'http://data.eastmoney.com/report/singlestock.jshtml?quoteid={}&stockcode={}'.format(codeid, code),
        'User-Agent': UserAgent().random,
    }
    r = requests.get(url, headers=headers)
    r = r.text[(len(cd)+1):-1]
    result = json.loads(r)
    return(result)


def download_reports(keys):
    # 用股票代码和时间戳创建文件夹
    code,beginTime,endTime=keys[0],keys[1],keys[2]
    timestr = str(round(time.time() * 1000))
    file_path = os.path.abspath(os.curdir)+'\\'+ checkNameValid(('_').join([code,timestr,beginTime,endTime]))
    if not os.path.exists(file_path):
        os.makedirs(file_path)
    # 判断总页数和总个数
    result = get_one_page(code,1,beginTime,endTime)
    hits = result['hits']
    print('成功获取{}个研报'.format(hits))
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
                result = get_one_page(code,page,beginTime,endTime)
                data_list = data_list + result['data']
            for i in range(num):
                progress_bar(i+(page-1)*50+1,hits,50)
                try:
                    d = result['data'][i]
                    title, infoCode, orgSName, publishDate, researcher = d['title'], d['infoCode'], d['orgSName'], d['publishDate'][:10], d['researcher']
                    dl_url = 'http://pdf.dfcfw.com/pdf/H3_{}_1.pdf'.format(infoCode)
                    pdf_name = ('_').join([publishDate, orgSName, title, researcher]) + '.pdf'
                    pdf_name = file_path+'\\'+ checkNameValid(pdf_name)
                    urllib.request.urlretrieve(dl_url, pdf_name)
                except Exception as e:
                    pass
        try:
            pf = pd.DataFrame(data_list)
            pf.fillna('', inplace=True)
            pf = pf.ix[:, ::-1]
            file_path=file_path+'\\'+r'01_reports_data.xlsx'
            pf.to_excel(file_path, encoding='utf-8', index=False)
            print('\n研报汇总表格数据已生成')
        except Exception as e:
            pass
    else:
        pass


if __name__ == "__main__":
    print('\n个股研报批量下载工具\n微信公众号：15Seconds\n作者：王文铖\n')
    while 1:
        code=input('请输入对应股票代码（6位数字，必填，按Enter确认）：')
        beginTime=input('请输入查询起始日期（格式：yyyy-mm-dd,选填，按Enter确认）：')
        endTime=input('请输入查询结束日期（格式：yyyy-mm-dd,选填，按Enter确认）：')
        keys=[i.replace(' ','') for i in [code,beginTime,endTime]]
        download_reports(keys)
        print('研报已下载完成\n')
