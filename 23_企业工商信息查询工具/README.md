
## 企业工商信息批量查询工具V2

作者 / wwcheng

公众号 / 15Seconds

点此进入企业工商信息批量查询工具V1

一、更新内容

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5rb3baxnlj30i3066q3a.jpg)

1 增加6个字段

在第一版基础之上增加法人、成员信息、分支机构、股东信息、对外投资和链接这6个字段信息，部分字段信息包含多行，可用“+”号分列

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5rb9ccc5xj311z028wer.jpg)

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5rba4pck9j30wf02ggly.jpg)

2 程序框架优化

Input文件夹用于输入，Output文件夹用于输出结果，Readme文件夹包含cookies的查看方式以及工具使用说明文档

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5rax9tqb8j30mw06f0sz.jpg)

3 将csv换成xlsx格式

将需要查询的公司名称填入 \Input\company_input.xlsx 文件中

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5rb0jsmv2j30lj04aglr.jpg)

查询结束后将会把结果输出为csv和xlsx两种格式，同时将返回日志文件error.log（用于记录运行结果和出现的BUG）

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5rb1a4m84j30lh04rq38.jpg)

4 其他优化内容

如，解决查询已注销公司时输出的信息错列的BUG，优化提示语句，优化程序运行速度（但是新加入6个字段，其实速度上与上一版本差不多）等

二、结果预览

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5rbw8xqajj31cj0l0ah9.jpg)
