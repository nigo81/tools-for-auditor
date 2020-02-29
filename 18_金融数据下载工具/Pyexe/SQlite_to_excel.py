#------------------------------------------------------------------------------
# -*- coding: utf-8 -*-        
# Name: SQlite_to_excel.py
# Email:w-wc@foxmail.com
# Author: wwcheng
# Last Modified: 2019-06-28 13:07
# Description: 
#------------------------------------------------------------------------------
import sqlite3
import pandas as pd
import datetime
import os
import sys


def config_read(filepath):
    f = open(filepath, 'r')
    config=f.read()
    config=config.strip()
    return config

def sqlite_query(SQLPath,table_name):
    conn = sqlite3.connect(SQLPath)
    sql_cmd='select * from %s;'%table_name
    re=pd.read_sql(sql_cmd, conn)
    conn.close()
    return re

def tushare_to_excel(excelPath,sqldata):
    writer = pd.ExcelWriter(excelPath)
    sqldata.to_excel(writer, 'Sheet1')
    writer.save()

def get_time_string():
    time_str = datetime.datetime.strftime(datetime.datetime.now(),'%Y%m%d%H%M%S')
    return time_str

def errorlog_write(filename,content):
    with open(filename, 'w', encoding='GB2312') as f:
        f.write(content)
        f.close()

def get_domanic_path():
    '''返回当前程序所在目录的父目录的相对路径'''
    path=os.path.dirname(os.path.realpath(sys.executable))
    path = path.split('\\')
    thispypath = "\\".join(path[:-1])
    return thispypath

def main():
    try:
        thispypath = get_domanic_path()
        download_config=config_read(thispypath+'\Config\download_config.ini')
        table_name=download_config+'_wwcheng'
        data=sqlite_query(thispypath+'\SQLiteSpy\TushareToExcel.db',table_name)
        print(data)
        excel_name=download_config+'_'+get_time_string()
        excelPath=thispypath+"\Output\%s.xlsx"%(excel_name)
        tushare_to_excel(excelPath,data)
        if os.path.exists(excelPath):
            errorlog_write(thispypath+'\Config\ErrorReport.ini', "%s 已下载完成，请打开output文件夹查看"%excel_name)
        else:
            errorlog_write(thispypath+'\Config\ErrorReport.ini',"出现错误未能下载成功，请检查您输入的参数格式是否有误！")
    except Exception as e:
        print('error:'+str(e))
        errorlog_write(thispypath+'\Config\ErrorReport.ini',"出现错误未能获取数据，请检查您输入的参数格式是否有误！" +str(e))

if __name__ == '__main__':
    main()