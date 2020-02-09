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