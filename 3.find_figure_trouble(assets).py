#!/usr/bin/env python
# coding: utf-8

# In[1]:


# 將財產分配 load 程式
import pandas as pd
import numpy as np
import time
from datetime import date
import os
import datetime
import re
import configparser as cp

def parse_btsid(value):
    pattern = 'R\d{4,6}|\d{4,7}[uUlLnNgG]'
    m = re.search(pattern, value)
    if m and m.group(0):
        return m.group(0)
    else:
        return '沒填寫'
    
def find_earlier_form(now1):
    path= './data_property_form'
    date_list = os.listdir(path)
    
    i=0    
    for i in range(30):
        date_name = 'form({}).xls'.format(now1)
        if date_name in date_list:
            break 
        else:
            now1 = now1 - datetime.timedelta(days=1)
    return(date_name)  

def find_earlier_Only_one():
    list1=['','','']
    path= './data_Only_One'
    date_list = os.listdir(path)
    n=0
    now1= date.today() 
    for i in range(30):   
        date_name_3G = 'Only_One_3G({}).xlsm'.format(now1).replace('-','.')
        date_name_4G = 'Only_One({}).xlsm'.format(now1).replace('-','.')
        date_name_5G = 'Only_One_5G({}).xlsm'.format(now1).replace('-','.')
        if date_name_3G in date_list:
            if list1[0] =='':
                list1[0] = date_name_3G
                n=n+1
        if date_name_4G in date_list:
            if list1[1] =='':
                list1[1] = date_name_4G
                n=n+1
        if date_name_5G in date_list:
            if list1[2] =='':
                list1[2] = date_name_5G
                n=n+1
        if n==3:
            break
        else:
            now1 = now1 - datetime.timedelta(days=1)
    return(list1)
#-------------------讀取config.ini------------------------------#  
filename = '.\config.ini'

# 配置文件读入
inifile = cp.ConfigParser()
inifile.read(filename, 'UTF-8')

# 读取 figure 部分 
address = inifile.get("figure", "address")
address = re.sub('[\r\n\t]', '', address)
address = address.split(',')

bts_device_name = inifile.get("figure", "bts_device_name")
bts_device_name = re.sub('[\r\n\t]', '', bts_device_name)
bts_device_name = bts_device_name.split(',')

bts_device_ant_name = inifile.get("figure", "bts_device_ant_name")
bts_device_ant_name = re.sub('[\r\n\t]', '', bts_device_ant_name)
bts_device_ant_name = bts_device_ant_name.split(',')

#----------------讀取固定檔案  名稱:  xls_df -----------------------#
select_day  = date.today() 
used_day = find_earlier_form(select_day) # 使用 def 
file1_path = "./data_property_form/{}".format(used_day)
xls_df = pd.read_excel(file1_path)

# ------------將基地台中文名稱取得， Dataframe 名稱: btsname_df------#
# all_files =['Only_One_3G(2021.05.13).xlsm','Only_One(2021.05.13).xlsm','Only_One_5G(2021.05.13).xlsm']
all_files = find_earlier_Only_one() # 使用 def
n = 0
for filename in all_files:
    begin_df = pd.read_excel('./data_Only_One/'+filename,sheet_name = 1 ) # only_one 增加了匯入 sheet
    list_colname = list(begin_df.head())
    begin_df.rename(columns = {list_colname[0]:'編號',list_colname[1]:'基地台名稱'},inplace = True)
    begin_df = begin_df[['編號','基地台名稱']]
    if n == 0 :
        btsname_df = begin_df
        n = n+1
    else:
        btsname_df = btsname_df.append(begin_df)
        
#----------------寫入檔案-------------------------#   
# 讀取資料庫 & 寫入 trouble 分析檔
today = date.today() 
writer =pd.ExcelWriter('./results(analysis)/bts_figure({}).xlsx'.format(today))
writer1 = pd.ExcelWriter('./results(analysis)/bts_trouble_shooting({}).xlsx'.format(today))

# 將「異動者」未指配項目列出，置於 sheet「異動者未指配」
values = {'異動者':'未指配','裝設地點':'未指配'}
xls_df = xls_df.fillna(value=values)
xls_df['異動者'] =  xls_df['異動者'].str.strip()
xls_df[xls_df['異動者']=='未指配'].to_excel(writer1, sheet_name = '異動者未指配',index=False)

# 將備料「放置地點」不對列出，置於 sheet「備料地點不對」
prepare_feed_loc = address

# 找出「備料地點不對」 
xls_df['裝設地點'] = xls_df['裝設地點'].str.strip()
xls_df['裝設地點'] = xls_df['裝設地點'].str.replace(',\(嘉義','嘉義',regex=True).str.replace('樓\)','樓',regex=True).str.replace('F\)','F',regex=True)
device_df = xls_df[xls_df['設備狀態'] =='備援/備用']
temp_df = device_df[~device_df['裝設地點'].isin(prepare_feed_loc)].copy()
temp_df = temp_df.drop(temp_df.loc[temp_df['裝設地點'].str.contains(pat = '\(R00000\)')].index) # 去除 repeater，HB 備料 (R00000)
temp_df.to_excel(writer1, sheet_name = '備料地點不對',index=False)

# 「使用中」料「放置地點」不對列出，置於 sheet「使用中地點未指配」
xls_df['設備狀態'] = xls_df['設備狀態'].str.strip()
state_df= xls_df[xls_df['設備狀態']=='使用中']
state_df[state_df['裝設地點'] =='未指配'].to_excel(writer1, sheet_name = '使用中地點未指配',index=False)

# 將「設備名稱」裡「微型基地台設備(行通)」各別列出置於 sheet「微型基地台」
xls_df['設備名稱'] = xls_df['設備名稱'].str.strip()
xls_df[xls_df['設備名稱']=='微型基地台設備(行通)'].to_excel(writer, sheet_name = '微型基地台B',index=False)
# 將 xls_df 去除「微型基地台設備(行通)」，得新的資料庫 「xls_df」
xls_df = xls_df.drop(xls_df.loc[xls_df['設備名稱']=='微型基地台設備(行通)'].index)


# 將「設備名稱」裡「增波器(Booster)(行通)」各別列出置於 sheet「增波器」
temp_df = xls_df[xls_df['設備名稱']=='增波器(Booster)(行通)']
titles = temp_df['裝設地點'].map(parse_btsid)
temp_df.insert(0,'編號',titles)
temp_df.to_excel(writer, sheet_name = '增波器B',index=False)
xls_df = xls_df.drop(xls_df.loc[xls_df['設備名稱']=='增波器(Booster)(行通)'].index)


# 將「設備名稱」裡「轉發器(Repeater)(行通)」各別列出置於 sheet「轉發器」，!!! 但天線尚未計算在內(Repeater 有天線設備) !!!
temp_df = xls_df[xls_df['設備名稱']=='轉發器(Repeater)(行通)']
titles = temp_df['裝設地點'].map(parse_btsid)
temp_df.insert(0,'編號',titles)
xls_df = xls_df.drop(xls_df.loc[xls_df['設備名稱']=='轉發器(Repeater)(行通)'].index)

# 以基地台財產 bts_df為基礎，置於 sheet「各別基地台財產」，從中找出基地台編號放置於 「基地台編號」的 column中
bts_df = xls_df[xls_df['設備名稱'].isin(bts_device_name)] 
titles = bts_df['裝設地點'].map(parse_btsid)
bts_df.insert(0,'編號',titles)

# 將 (Repeater) 天線資料拿出來 ，「轉發器」與 「轉發器天線」合併 
temp1_df = bts_df[bts_df['編號'].str.contains(pat = 'R')]
temp_df.append(temp1_df).to_excel(writer, sheet_name = '轉發器B',index=False)
bts_df = bts_df.drop(bts_df.loc[bts_df['編號'].str.contains(pat = 'R')].index)

# 將基地台中文名稱加入 bts_df 中，且置於第二欄
bts_df = pd.merge(bts_df,btsname_df,  how="left")
cols = bts_df.columns.tolist()
cols.insert(1, cols.pop(cols.index('基地台名稱'))) 
bts_df = bts_df[cols]  

# Delete these row indexes from dataFrame
indexNames = bts_df[ bts_df['廠牌'].isnull()].index
bts_df.drop(indexNames , inplace=True)
bts_df.to_excel(writer, sheet_name = '基地台',index=False)


# 將「基地台設備」資料另存於 sheet 「基地台設備」
unant_df = bts_df[~bts_df['設備名稱'].isin(bts_device_ant_name)] 
unant_df.to_excel(writer, sheet_name = '基地台設備',index=False)

# 將「基地台天線」資料另存於 sheet 「基地台天線」
ant_df = bts_df[bts_df['設備名稱'].isin(bts_device_ant_name)] 
ant_df.to_excel(writer, sheet_name = '基地台天線',index=False)

#################  trouble shooting   ##################

# 找出「財產名稱」與「基地台編號」不符合，列出 「trouble_shooting1.xlxs」 Sheet「財產名稱_編號不符」
# temp_df = bts_df[bts_df['財產名稱'].isin(['4G行動寬頻系統共構設備','4G行動寬頻基地台','4G系統行動寬頻介接設備','4G系統基地台充電設備','4G系統蓄電池'])]
# temp_df = temp_df[temp_df['編號'].str.contains(pat = '[UN]',regex = True, case = False)]
# temp1_df = bts_df[bts_df['財產名稱'].isin(['5G/4G/3G室外涵蓋天線','5G系統蓄電池','5G基地台射頻模組','5G基地台基頻模組','5G基地台彙集設備'])]
# temp1_df = temp1_df[temp1_df['編號'].str.contains(pat = '[UL]',regex = True, case = False)]
# temp_df.append(temp1_df).to_excel(writer1, sheet_name = '財產名稱_編號不符',index=False)

# 將各「設備名稱」歸類於設備狀態:[使用中]，而基地台名稱為:'',列出 「trouble_shooting1.xlxs」 Sheet「設備狀態_不匹配」
temp_df = bts_df[(bts_df['設備狀態']=='使用中') & (bts_df['基地台名稱'].isnull())]
temp_df = temp_df.drop(temp_df.loc[temp_df['使用單位']=='嘉義品改股'].index)
temp_df.to_excel(writer1, sheet_name = '使用中_編號不符',index=False)

# 將基地台停用設備不等於 0 列出 
temp1_df = bts_df[(bts_df['設備狀態']=='停用') & (bts_df['本月設備淨值']!=0)]
temp1_df.to_excel(writer1, sheet_name = '停用設備淨值不為0',index=False)



# 關閉寫入檔案
writer.save()
writer1.save()



# In[ ]:





# In[ ]:




