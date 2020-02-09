
## 企业工商信息批量查询工具

作者 / wwcheng

公众号 / 15Seconds

一、引言

在审计过程中，经常需要查询一些企业的工商登记信息，如关联方、主要客户与供应商、重要往来单位等，据此做出职业判断或获取审计证据。

现在查询工商信息的网站也很多，天眼查、企查查、百度信用、启信宝……而且这些网站也搭建了相关的数据中心，提供很多方便的收费API(大概0.1/次左右)，可以直接调用，通过这几天测试，最终选择企查猫进行爬虫，虽然麻烦一些但是速度不错(约30个/分钟)，反爬较弱，最重要的是不收费。

工具最终得出的结果如下：

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5lfixd80uj317p0l3jxs.jpg)

二、简要流程

1 获取Cookies

2 填入待查的公司名称

3 运行程序

4 得到企业的工商信息

三、学习内容

1 列表生成式

info = [(i.split('：'))[1] for i in info1]

2 切片

info = info2[:4]+info2[-10:]

3 XPATH

 html.xpath('//section[@class="pb-d2"]/ul[1]/li/span[@class="info"]/text()')

四、详细使用步骤

0 下载程序

微云下载链接：https://share.weiyun.com/560P7Zd

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5lnpxsapaj30hv08xjrz.jpg)

1 获取Cookies

Cookies最典型的应用是判定注册用户是否已经登录网站，用户可能会得到提示，是否在下一次进入此网站时保留用户信息以便简化登录手续，所以我们需要获取Cookies保持登录状态

工具里面提供的Cookies是具有时效性的，具体多久失效没有测试，用了一天没问题，用户最好学会获取自己的Cookeis，这样才一劳永逸。

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5lf3rjjcjj30pw0i276g.jpg)

搜索企查猫，注册登录后，右键单击页面，选择检查-Network，刷新当前页面后，找到自己的Cookeis，如下图（图片看不清的话，下载链接中有原图）：

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5lfb7pexsj31h30sm49l.jpg)

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5lfesljzlj31gw0smgx1.jpg)

2 填入待查的公司名称

在company_input.csv中输入需要查询的公司名称，最好输入完整名称，保存好工作表。

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5lf2wjs4hj30lj07ngm4.jpg)

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5lf1mog8mj306003oa9z.jpg)

3 运行程序

双击exe文件开始运行程序，每次查询随机等待零点几秒，每查询30个会暂停一分钟，毕竟爬取也不要太过分

4 得到企业的工商信息

最终会得到company_output_时间戳.csv的文件，第一列为查询的公司名称，第二列为实际匹配到的结果，第三列名称不一致的，原因可能是公司名称填错、股改等

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5lf4n8952j30m50723yy.jpg)

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5lf09uzckj30q002k0t0.jpg)

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5lf12wc7zj30lr02r0sw.jpg)
