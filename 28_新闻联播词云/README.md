## 程序猿用什么姿势看新闻联播

作者 \ 王文铖<br/>微信公众号 \ 15Seconds

---
公众号（15Seconds）第一篇文章*Tushare_to_Excel*就已经介绍了Tushare金融大数据社区，Tushare_to_Excel可以实现Tushare金融数据与Excel的对接，无需编程基础。本篇文章利用Tushare新闻联播接口，学习一下词云的绘制。

## 新闻联播
![Tushare](https://user-gold-cdn.xitu.io/2020/2/27/170854494de48e11?w=223&h=64&f=png&s=4410)

[Tushare_CCTV_news](https://waditu.com/document/2?doc_id=154 "Tushare_CCTV_news"):<br/>https://waditu.com/document/2?doc_id=154

![cctvnews_demo](http://tushare.org/img/cctv_news.png)

## 姿势讲解

选取的数据：2020年2月19日-2020年2月26日的新闻联播文字稿

1. 安装python工具（*查看安装教程*）
2. 通过tushare获取cctv_news
3. 使用jieba分词进行关键字词提取
4. 根据提取关键字词绘制词云

## 词云展示
![cctvnews_wordclound_nomask](https://user-gold-cdn.xitu.io/2020/2/27/17085350c9d622c2?w=2400&h=2400&f=jpeg&s=749416)

![cctvnews_wordclound_mapmask](https://user-gold-cdn.xitu.io/2020/2/27/1708535a08acd2a6?w=5000&h=4393&f=jpeg&s=1190429)

## 背景样式

![map_mask.png](https://user-gold-cdn.xitu.io/2020/2/29/17090b6ee80d9fe8?w=5000&h=4393&f=png&s=8961214)

## 工具声明

本公众号软件未经作者许可，不得用于商业用途，仅做编程学习交流之用

<img style="width:4em;display: block;" src="https://image2.135editor.com/cache/remote/aHR0cHM6Ly9tbWJpei5xbG9nby5jbi9tbWJpel9naWYvN1FSVHZrSzJxQzVBZTdUTDByaWJrdEtXRGRpYjRMdjQ3eHNvZzRHNzJoUDRUaWNlRWZpYlo4S2liU0RCZzJTUk52RHpPYThhYllCTGRRUU5SWTN0a1pNSXJmZy8wP3d4X2ZtdD1naWY=">

本公众号工具及源码会在[Github](https://github.com/nigo81/tools-for-auditor "Github项目地址")上正式发布和后续更新。<br/>订阅公众号之后，在后台回复“**tools**”可获取 15Seconds 公众号所有工具，**转载或者提问**可以后台留言或者在后台加我**个人微信**交流！
