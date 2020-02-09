## [Python爬虫练习] 批量下载个股研报（附工具）

![mjW5w9.png](https://s2.ax1x.com/2019/08/30/mjW5w9.png)

作者 / 王文铖

公众号 / 15Seconds

程序运行示例GIF：

![mjWron.gif](https://s2.ax1x.com/2019/08/30/mjWron.gif)

### 参数设置：

![mjWRQU.png](https://s2.ax1x.com/2019/08/30/mjWRQU.png)

参数1：股票代码（6位数字，必填）

参数2：查询起始日期（格式：yyyy-mm-dd,选填）

参数3：查询结束日期（格式：yyyy-mm-dd,选填）

### 运行结果：

程序下载速度很快，基本只需几秒钟，结束后将会输出 研报PDF+Excel数据（比官网展示的数据还要详细）

![mjWfL4.png](https://s2.ax1x.com/2019/08/30/mjWfL4.png)

![mjq0IS.png](https://s2.ax1x.com/2019/08/30/mjq0IS.png)

### 程序框架：

主体的程序其实就是获取单页的JSON数据，然后解析JSON来下载PDF。在抓取单页数据前，有两个地方需要对JS进行逆向解密，从而获取正确的请求头内容。

![mjWcWV.png](https://s2.ax1x.com/2019/08/30/mjWcWV.png)

### 新体验—进度条

第一次在程序中加入进度条，感觉效果还不错，有“download”的感觉。下面提供的进度条函数，可以自己根据需要修改，可以改变它的图例样式。

```
import sys
import time

'''下载进度条函数'''
def progress_bar(name,i,total=100,width=50):
    # 参数：进度条名称，当前进度，进度总数，进度条宽度
    p=int(i* width / total)
    percent= "["+str(int(p * 100 / width)) + "%"+"]"+"["+ str(i) +"/" + str(total) +"]"
    bar_str = name + " [" + "#" * p + "_" * (width-p) + "]" + percent
    sys.stdout.write("\r%s" %bar_str)
    sys.stdout.flush()
    time.sleep(0.1)

name, total, width ='demo1:', 10, 10
for i in range(1,total+1):
    progress_bar(name ,i,total,width)
```

进度条操作示例gif：

![mjWyiq.gif](https://s2.ax1x.com/2019/08/30/mjWyiq.gif)