## 如何高效批量对工作簿进行创建、合并拆分、转换格式？

本文提供一个批量操作工作簿方案，通过VBA制作的Excel工具实现，利用这一工具可以做到批量对Excel工作簿进行创建、多种方式合并、拆分重组、重命名、转换格式、插入目录链接等。

当然，目前也有很多Excel插件可以实现上述部分功能，但是插件会对Excel加载速度会有影响，而且插件中大部分功能其实用不上，对于追求效率或者不想安装插件的话，可以尝试本文提出的解决方案，同时，如果你正在学习VBA，也可以借助这个工具练习VBA中工作簿的操作。

详细使用说明可以查看下面的视频（视频虽长，录制不易）

使用这一工具之前，先了解一下两个概念：工作簿（workbook）与工作表（worksheet）

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g6by68shltj310e0jawgt.jpg)

上图中有4个工作簿，每个工作簿有4个工作表，这个文件夹中共4个工作簿，16个工作表。

1 合并所有工作簿内所有工作表至总表

使用情景：把一个文件夹内所有的工作簿内的工作表内容汇总至同一个工作表中，可用于把示例文件中16个工作表汇总为1个工作表

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g6bwzyf0dsj30xp0o0aq2.jpg)

2 合并所有工作簿内特定名称的工作表至总表

使用情景：把一个文件夹内所有的工作簿内指定名称的工作表汇总至同一个工作表中，可用于把示例文件中16个工作表中名称为“2017年-3月”的汇总为1个工作表

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g6bx0hrjdkj30y20o0h1t.jpg)

3 合并所有工作簿内表名包含关键字的工作表至总表

使用情景：把一个文件夹内所有的工作簿内包含关键字的工作表汇总至同一个工作表中，可用于把示例文件中16个工作表中包含关键字“3月”的汇总为1个工作表

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g6bx11i6mvj30y20o0qj4.jpg)

4 合并所有工作簿内特定位置工作表至总表

使用情景：把一个文件夹内所有的工作簿内特定位置的工作表汇总至同一个工作表中，可用于把示例文件中每个工作簿中第3个表的汇总为1个工作表

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g6bx1hfrqtj30y20o0wuo.jpg)

5 保留格式合并所有工作簿内至同一个工作簿

使用情景：把一个文件夹内所有的工作簿在保留格式的前提下汇总至同一个工作簿内中，可用于把示例文件中4个工作簿保留格式汇总至1个工作簿（可以用于快速对单个主体的所有科目底稿调整格式，批量打印）

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g6bx21f8gbj30y20o04en.jpg)


6 XLS批量转换为XLSX

使用情景：将一个文件夹内XLS格式的工作簿批量转换为XLSX格式

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g6bx4qpfbfj30ux0nmmzs.jpg)

7 批量更换工作簿文件名

使用情景：批量更换一个文件夹内所有工作簿的文件名

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g6bx5jufuzj30uv0mnjte.jpg)

8 批量创建工作簿或工作表

使用情景：在一个文件夹内批量新建多个工作簿，或者在一个工作簿内创建多个工作表

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g6bx68jlm1j30uq0nhdhs.jpg)

9 批量拆分重组工作簿

使用情景：将一个文件夹内所有的16个工作表拆分为单独的工作簿，再对这16个工作簿归类至对应文件夹内（具体效果可以查看视频）

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g6bx71kkp7j30uo0nkjtc.jpg)

10 插入目录链接和返回目录链接

使用情景：为一个包含多个工作表的工作簿插入目录链接和返回目录链接

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g6bx79s0qnj30ux0nmtao.jpg)

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g6bywx156pj30uv0ia0u5.jpg)
