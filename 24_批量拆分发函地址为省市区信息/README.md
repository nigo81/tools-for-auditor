![UTOOLS1565525767440.png](http://yanxuan.nosdn.127.net/fd331abaa32580834135c46c348b2f34.png)

很多时候审计师在发函过程中使用顺丰。

但是要求填写省、市、区信息。

如何快速、批量拆分出省、市、区信息就是我们的需求。当然还要面对本身就缺失这些信息的情况。

![UTOOLS1565525961785.png](http://yanxuan.nosdn.127.net/fb7f2454abe3ce6f011deff524ed6257.png)

在github上，已经有人造出了轮子，我们直接用起来。

## 安装依赖库

这里我们需要用到cpca及pandas。

在本机已经安装好python3的情况下，用ctrl+r调用cmd命令行。

```
pip install cpca

pip install pandas
```
## 填写地址

打开input.xlsx文件，填写需要拆分的地址后关闭。

![UTOOLS1565526249780.png](http://yanxuan.nosdn.127.net/8712a681fae483035853efed8d4b35b3.png)

## 运行代码

在命令行中position.py文件路径下运行代码。或者用vscode等IDE中运行。
```python
import cpca
import pandas as pd
df=pd.read_excel('input.xlsx','Sheet1',index_col=None,na_values=['NA'])
location_str=list(df['地址'])
if location_str:
    df = cpca.transform(location_str)
    df.to_excel('out_put.xlsx',sheet_name='Sheet1')
    print(df.head())
else:
    print('请正确填写地址')
```
## 生成结果
![UTOOLS1565526478146.png](http://yanxuan.nosdn.127.net/10c68b1ec94f2c3cde371b79635bf895.png)


![UTOOLS1565526514661.png](http://yanxuan.nosdn.127.net/0d8ae91bd0dcc49cac43904c87d28c62.png)

output.xlsx文件中存储着拆分结果。

## 结语

由于用pyinstaller打包成exe后，不能成功运行。所以就只有这个python文件。

如果需要使用需要安装python3环境。