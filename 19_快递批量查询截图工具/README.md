
## 快递批量查询截图工具，可查所有快递+输出长截图

作者 / 王文铖

公众号 / 15Seconds

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g60mx2dfo8j31go0o748r.jpg)

一 引言

本文在上一版本基础上做出大量改进，在原有顺丰速运查询的基础上再加入快递100网站，也就是说这两个网站能查到的快递单号，本工具都可以查询并截长图，这两个网站也基本覆盖了大家的查询需求了。

同时，对截图做了深度优化。查询顺丰快递（通过顺丰速运官网）将会输出三张截图，包括物流节点高清长截图、运单详情页和电子存根页；查询其他快递（通过快递100官网）时，只输出物流节点高清长截图。当然，查询时不需要特地将顺丰和非顺丰的分开，工具内嵌了自动检索单号所属物流公司的模块。

二 工具截图效果预览

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g61d2cvwyrj30jq06xq3y.jpg)

三 如何使用工具

1 下载工具文件

微云下载链接：https://share.weiyun.com/5RBczT7  

如果链接失效，关注微信公众号15Seconds获取最新的工具

下载解压后，如果系统杀毒软家报木马，是因为文件中包含自己制作的程序，如果是从本公众下载的文件，可以忽略报错提示，其他途径获取的工具，需要自己识别，工具内容如下：

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g61d7cyyjmj30h30blaaa.jpg)

2 安装浏览器

打开 Driver 文件夹，根据自己的电脑系统版本安装对应的Firefox浏览器，这个浏览器是整个程序运行必需的依赖环境，通过模拟手工操作浏览器，达到查询快递和截图的目的，效果甚至比人工截图更好。

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g60mfk8vi1j30mo05q74p.jpg)

3 输入快递单号

在 Input 文件夹中的 bill_number.xlsx 工作簿中输入待查询的快递单号，工作簿的名称和格式不能改变，否则程序无法运行。

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g60mg1my64j30kf03k0sr.jpg)

4 运行程序

双击 Express_Tracking.exe 开始运行程序，同时打开手机微信准备扫码登陆，扫码后程序才会运行。可以使用非对应的微信账号扫码，但是只有当扫码的微信对应的手机号与快递单号发件人或收件人手机号一致时，才可以获取运单详情页和电子存根页。

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g60mmj16b1j30ks08bjry.jpg)

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g61dq3ma78j30hq0dawg3.jpg)

5 输出物流信息和截图

程序运行结束之后，可以进入 Output 文件夹查看输入结果， img、log、temp、xls文件夹中分别保存了截图、日志、临时文件、物流节点。

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g60mgndfdej30kn05dmxc.jpg)

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g60mlh1ixjj30l704dglq.jpg)
