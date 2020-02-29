**Tushare_to_Excel 的“十八般武艺”**

------

> 作者：王文铖 / 微信公众号：15Seconds

通过Excel轻松查询并下载全品类的金融数据，不仅包含股票、指数、基金、期货、期权、债券、外汇等行业大数据，而且包括了数字货币行情等区块链数据，甚至还有一些特色数据，如新闻联播文字稿等。具体包含的数据接口多达107种（如下图）。



![1.png](https://user-gold-cdn.xitu.io/2020/2/26/170811d7f1ea7f30?w=1195&h=639&f=png&s=146770)



**Tushare_to_Excel 的数据来源**

------

 

本工具的数据来源是Tushare，Tushare是一个金融大数据开放社区，免费提供各类金融数据和区块链数据，用户可以通过http、Python、Matlab、R语言方式来免费地获取数据。而Tushare的数据是来自于新浪财经、腾讯财经、上交所和深交所等，并能做到及时更新。

> Tushare 官网：https://tushare.pro/   
>
> Tushare 微信公众号：挖地兔

![0.png](https://user-gold-cdn.xitu.io/2020/2/26/170811cd40e0a1cd?w=1478&h=540&f=png&s=1019488)


**Tushare_to_Excel 的制作初衷**

------



1. 能站着就不要“爬”。提供一个不需要编程基础也能轻松获得金融数据的方案，对于没有时间、精力、兴趣去学习编程技能（数据库、爬虫等）的人来说，只通过Excel这一基本工具就能得到最终结果便是最理想的。（但是仍然鼓励大家学习编程！）

2. 从我个人角度，是为了通过这一个小工具的制作来熟悉SQL、Python、VBA等，因此我选择了Tushare的Python SDK和“乞丐版”数据库SQLite来完成这个工具。



**Tushare_to_Excel 的使用方法**

------



**1. 下载文件后，运行“Tushare_to_Excel.xlsm”**

提示：打开excel之后弹出的安全警告一定要选择“启用内容”或者“启用宏”

![3.png](https://user-gold-cdn.xitu.io/2020/2/26/170811dc3ef47c1e?w=599&h=136&f=png&s=28480)

![2.png](https://user-gold-cdn.xitu.io/2020/2/26/170811daefcefdac?w=855&h=722&f=png&s=199718)



**2. 注册Tushare账号，获取Token值**

注册地址：官网 或者 我的个人推荐地址（https://tushare.pro/register?reg=278708）

通过我的个人推荐地址注册账号，我会获得积分，积分与可使用接口数量和调用频次有关，希望大家支持！

注册成功之后，进入个人主页—接口TOKEN，复制token至excel表格G1单元格。

![4.png](https://user-gold-cdn.xitu.io/2020/2/26/170811ddb7a00840?w=948&h=302&f=png&s=35087)

 

**3. 双击三级目录具体项目，进入对应的数据接口**

![5.png](https://user-gold-cdn.xitu.io/2020/2/26/170811dedfe39e6c?w=637&h=436&f=png&s=80063)

**4. 将对应的输入参数和输出参数填在E列中，注意格式要求**

![6.png](https://user-gold-cdn.xitu.io/2020/2/26/170811e34064cea2?w=1343&h=734&f=png&s=329921)

**5. 依次单击“查询数据”，“预览数据”，“下载数据”按钮**

点击查询数据之后，等待几秒钟，会弹出提示，如果提示表格已经保存好了，就可以预览数据，再选择下载数据；如果提示出现错误，则需要检查参数的填写是否符合规范，更详细的参数文档和示例文件可以进入官网查看。

![7.png](https://user-gold-cdn.xitu.io/2020/2/26/170811e41a57e9d8?w=833&h=369&f=png&s=91722)

![8.png](https://user-gold-cdn.xitu.io/2020/2/26/170811e53d24472f?w=785&h=563&f=png&s=211487)

![9.png](https://user-gold-cdn.xitu.io/2020/2/26/170811e6a3f0ac09?w=678&h=440&f=png&s=74120)

PS：工具中包含了自己制作的程序，杀毒软件可能会提示发现可疑程序，要允许程序运行。

**6. 打开本工作簿路径下output文件夹，查看下载成功的数据表**

![10.png](https://user-gold-cdn.xitu.io/2020/2/26/170811e788d1f31c?w=793&h=440&f=png&s=107082)



