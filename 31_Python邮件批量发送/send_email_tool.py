import datetime
import email.mime.multipart
import email.mime.text
import os
import random
import smtplib
import time
from email.header import Header  # 处理邮件主题
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr, parseaddr  # 用于构造特定格式的收发邮件地址

import numpy as np
import pandas as pd
import prettytable as pt

now_time = datetime.datetime.strftime(datetime.datetime.now(), '%Y%m%d-%H%M%S')
now_path = os.path.abspath(os.curdir)

def print_tool_info():
    tb = pt.PrettyTable()
    tb.field_names = ['名称', '邮件批量发送工具']
    tb.add_row(['作者', '王文铖'])
    tb.add_row(['微信公众号', '15Seconds'])
    tb.add_row(['更新时间', '2019-12-16'])
    tb.add_row(['版本号', 'v1.0'])
    print(tb)

def __format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))

def load_info():
    input_path = now_path+r'\邮箱配置表.xlsx'
    df = pd.read_excel(input_path).fillna('')
    data = np.array(df)
    print('邮箱配置加载成功')
    return data

def send_email(email_config):
    msg = email.mime.multipart.MIMEMultipart()
    msg['from'] = __format_addr('{}<{}>'.format(email_config[4],email_config[0]))
    msg['to'] = __format_addr('{}<{}>'.format(email_config[5],email_config[2]))
    msg['subject'] = email_config[3]
    if email_config[7]=='plain':
        txt = email.mime.text.MIMEText(email_config[6], email_config[7], 'utf-8')
    elif email_config[7]=='html':
        f1 = open(email_config[6], 'r') 
        content = f1.read() 
        f1.close() 
        txt = email.mime.text.MIMEText(content, email_config[7], 'utf-8')  # plain or html
    msg.attach(txt)

    path_list=email_config[8:14]
    path_list=[p for p in path_list if p !='']

    for path in path_list:
        part = MIMEApplication(open(path, 'rb').read())
        part.add_header('Content-Disposition', 'attachment',filename=path.split('\\')[-1])
        msg.attach(part)

    smtp.login(email_config[0], email_config[1])
    smtp.sendmail(email_config[0], email_config[2], str(msg))

if __name__ == "__main__":
    
    print_tool_info()
    email_data = load_info()
    print('正在尝试连接邮箱服务器')
    smtp = smtplib.SMTP()
    f2 = open(now_path+r'\config.config', 'r') 
    config = f2.readlines() 
    config = [c.replace('\n','') for c in config]
    f2.close() 
    smtp.connect(config[0],config[1])

    for i in range(len(email_data)):
        try:
            email_config = email_data[i]
            send_email(email_config)
            a = round(random.uniform(1, 2), 2)
            time.sleep(a)
            print('邮件发送进度：{}/{}  {}<{}>已发送成功'.format(i+1,len(email_data),email_config[5],email_config[2]))
            content = '【发送正常】 '+ str(email_config[5]) +' + '+ str(email_config[2]) 
            log_path = now_path+r'\log\log-{}.txt'.format(now_time)
            f3 = open(log_path,'a') 
            f3.write(content+ '\n')
            f3.close
        except Exception as err:
            content = '【出现错误】 '+ str(email_config[5]) +' + '+ str(email_config[2]) +' + '+ str(err)
            print(content)
            log_path = now_path+r'\log\log-{}.txt'.format(now_time)
            f3 = open(log_path,'a') 
            f3.write(content+ '\n')
            f3.close
    smtp.quit()
    print('发送完成，关闭邮箱服务器')
    time.sleep(2)
