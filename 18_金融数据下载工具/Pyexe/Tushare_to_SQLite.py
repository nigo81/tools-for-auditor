#------------------------------------------------------------------------------
# -*- coding: utf-8 -*-        
# Name: Tushare_to_SQLite_2.0.py
# Email:w-wc@foxmail.com
# Author: wwcheng
# Last Modified: 2019-06-29 11:03
# Description: 
#------------------------------------------------------------------------------
import tushare as ts
import sqlite3
import pandas as pd
import configparser
import sys

def config_read(config_path,sdk=True):
    config = configparser.ConfigParser()
    config.read(config_path, encoding="utf-8")
    Parameter = ''
    for i in config.options('Parameter'):
        inp = config.get('Parameter', i)
        Parameter = Parameter + ',' + inp
    Parameter = Parameter[1:]
    Token = config.get('Token', 'token')
    code = 'ts.pro_api(%s).query(%s)' % (Token, Parameter)
    Table_name=str(config.get('Parameter', 'sdk'))[1:-1]
    if sdk==False:
        return Table_name
    else:
        return code

def tushare_to_sqlite(SQLPath,table_name,data):
    conn = sqlite3.connect(SQLPath)
    cur = conn.cursor()
    try:
        cur.execute('DROP TABLE %s;' %table_name)
    except sqlite3.OperationalError:
        print('数据库中不存在表“%s”'%table_name)
    conn.commit()
    data.to_sql(table_name, con=conn, if_exists='append',index=False)
    sql_cmd='select * from %s;'%table_name
    re=pd.read_sql(sql_cmd, conn)
    conn.close()

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
        config_path = thispypath+'\Config\post_config.ini'
        tushare_code=config_read(config_path)
        print(tushare_code)
        data = eval(tushare_code)
        print(data)
        table_name = config_read(config_path,False)+'_wwcheng'
        tushare_to_sqlite(thispypath+'\SQLiteSpy\TushareToExcel.db',table_name,data)
        errorlog_write(thispypath+'\Config\ErrorReport.ini', "数据已保存至数据库中的表 %s，可以进行预览"%table_name)
    except Exception as e:
        errorlog_write(thispypath+'\Config\ErrorReport.ini',"出现错误未能获取数据，请检查您输入的参数格式是否有误！" +str(e))

if __name__ == '__main__':
    main()