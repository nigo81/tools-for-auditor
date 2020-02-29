# -*- coding: utf-8 -*-
import datetime
import os
import re
import time

import jieba.analyse
import numpy as np
import pandas as pd
import tushare as ts
from PIL import Image
from wordcloud import WordCloud

# 基础变量设置-可自行修改
# tushare_token获取链接：https://tushare.pro/register?reg=278708
tushare_token = 'd3276c7ba97ac3c839dcedd6b1c74d828b6286793326929673bd1c21'
day_start = r'20200219'  # 起始日期
day_end = r'20200226'  # 结束日期

# 获取当期时间和当前路径
now_time = datetime.datetime.strftime(datetime.datetime.now(), '%Y%m%d_%H%M%S')
now_path = os.path.abspath(os.curdir)

# 计算时间段包含的日期
def calculate_days(day_start,day_end):
    day_start = datetime.datetime.strptime(day_start,'%Y%m%d')
    day_end = datetime.datetime.strptime(day_end,'%Y%m%d')
    day_list = []
    while day_start <= day_end:
        day = day_start.strftime('%Y%m%d')
        day_list.append(day)
        day_start += datetime.timedelta(days=1)  # 下一个日期
    return day_list

# 通过tushare获取cctv_news
def get_news(day_list):
    pro = ts.pro_api(tushare_token)
    df_list = []
    for d in range(len(day_list)):  # 开始遍历每一个日期
        print('{}/{} now is getting the news of {}...'.format(d+1,len(day_list),day_list[d]))
        df = pro.cctv_news(date = day_list[d])  # CCTV新闻联播
        df_list.append(df)
        # cctvnews_file = now_path + '\\' + r'cctvnews_{}.csv'.format(day_list[d])
        # df.to_csv(cctvnews_file,encoding='utf-8-sig') # 把每天的cctv_news保存到CSV
        time.sleep(0.5)
    return df_list

# 合并所选时间段cctv_news的内容
def get_content(df_list):
    df_all = pd.concat(df_list)  # 连接函数concat
    df_all = df_all.reset_index()  # 重置索引
    series_content = df_all['content']  # 提取表格的content列
    content_string = ''
    for content in series_content:  # 遍历Series中的每一行congtent（每一天的content）
        content = str(content)  # 转为字符串
        content = re.sub("[\s+\.\!\/_,$%^*(+\"\']+|[+——！，。？、~@#￥%……&*（）]+", "",content)  # 去除标点符号
        content_string = content_string + content  # 汇总
    return content_string

# 使用jieba对cctv_news进行关键字词提取
def split_word(content_string):
    result=jieba.analyse.textrank(content_string,topK=300,withWeight=True)
    stopwords=['中国','全面','表示','会议','单位','企业','方式','国家']  # 停用词列表
    keywords = dict()
    for i in result:
        if i[0] in stopwords:
            pass
        else:
            keywords[i[0]]=i[1]
    return keywords

# 根据提取关键字词绘制词云
def creat_wordcloud(keywords):
    font = r'C:\Windows\Fonts\simfang.ttf'
    image= Image.open(now_path+r'\map_mask.png')
    mask = np.array(image)
    wordcloud_cctvnews = WordCloud(collocations=False,
                                   font_path=font,
                                   width=2400,
                                   height=2400,
                                   margin=2,
                                   background_color='white',
                                   mask = mask).generate_from_frequencies(keywords)
    # wordcloud_cctvnews = WordCloud(font_path=font).generate(keywords)                           
    wordcloud_cctvnews.to_file(now_path + r'\\cctvnews_{}.jpg'.format(now_time))
    print('Cctvnews wordcloud has been creat successfully.')

if __name__ == "__main__":
    day_list = calculate_days(day_start,day_end)
    df_list = get_news(day_list)
    content_string = get_content(df_list)
    # with open("test.txt","w") as f:
    #     f.write(content_string)
    # with open("test.txt","r") as f:
    #     content_string = f.read()
    keywords = split_word(content_string)
    creat_wordcloud(keywords)
