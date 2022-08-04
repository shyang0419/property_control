#!/usr/bin/env python
# coding: utf-8

# In[23]:


# 將財產分配 load 程式

# ANT-assets天線採用 => 財產名稱:'3G行動電話系統天線','２Ｇ及３Ｇ行動電話室外涵蓋型共用天線系統','4G/3G/2G行動通信室內涵蓋天線',
# ANT-assets                    '4G/3G/2G行動通信室外涵蓋天線','5G/4G/3G室外涵蓋天線','２Ｇ及３Ｇ行動電話室內涵蓋型共用天線系統'
# ANT-assets排除 =>設備名稱包含:'共構'

import pandas as pd
import numpy as np
import time
from datetime import date
import datetime
import os
import re
import configparser as cp

def parse_btsid(value):
    pattern = 'R\d{4,6}|\d{4,7}[uUlLnNgG]'
    m = re.search(pattern, value)
    if m and m.group(0):
        return m.group(0)
    else:
        return '沒填寫'

def choice_char(value):
    pattern = '([1-9])[0-9][0-9]' 
    m = re.search(pattern, str(value))
    if m and m.group(1):
        return m.group(1)
    
def find_earlier_eNodebList():
    list1=['','']
    path= './data_eNodebList'
    date_list = os.listdir(path)
    n=0
    now1= date.today() 
    for i in range(30):   
        date_name_4G = 'eNodebList_4G({}).xlsx'.format(now1)
        date_name_5G = 'eNodebList_5G({}).xlsx'.format(now1)
        if date_name_4G in date_list:
            if list1[0] =='':
                list1[0] = date_name_4G
                n=n+1
        if date_name_5G in date_list:
            if list1[1] =='':
                list1[1] = date_name_5G
                n=n+1
        if n==2:
            break
        else:
            now1 = now1 - datetime.timedelta(days=1)
    return(list1)

def find_last_record():
    list1=['']
    path= './results(analysis)'
    date_list = os.listdir(path)
    now= date.today() 
    now1 = now - datetime.timedelta(days=1)
    for i in range(30):   
        antenna_df = 'bts_antenna_switch({}).xlsx'.format(now1)
        if antenna_df in date_list:
            if list1[0] =='':
                list1[0] = antenna_df
            break
        else:
            now1 = now1 - datetime.timedelta(days=1)
    return(list1)

def find_earlier_in_stock():
    list1=['']
    path= './data_in_stock'
    date_list = os.listdir(path)
    now= date.today() 
    for i in range(120):   
        pcb_df = '天線庫存({}).xlsx'.format(now)
        if pcb_df in date_list:
            if list1[0] =='':
                list1[0] = pcb_df
            break
        else:
            now = now - datetime.timedelta(days=1)
    return(list1)
#-------------------讀取config.ini------------------------------#  
filename = '.\config.ini'

# 配置文件读入
inifile = cp.ConfigParser()
inifile.read(filename, 'UTF-8')

# 读取 antenna 部分 
property_name = inifile.get("antenna", "property_name")
property_name = re.sub('[\r\n\t]', '', property_name)
property_name = property_name.split(',') 

except_smallcell = inifile.get("antenna", "except_smallcell")
except_smallcell = re.sub('[\r\n\t]', '', except_smallcell)
except_smallcell = except_smallcell.split(',') 

 # --------讀取 assets 基地台資料庫 (assets_df) ----------- 
today = date.today() 
file1_name = 'bts_figure({}).xlsx'.format(today)
file1_path = "./results(analysis)/{}".format(file1_name)
assets_df = pd.read_excel(file1_path, sheet_name = '基地台天線',dtype='str')

# -------讀取 nmoss 基地台資料庫 (nmoss_df) -------------
nmoss_files = find_earlier_eNodebList() # 使用 def 
# ---------寫入資料庫  ---------------------------------
today = date.today() 
writer =pd.ExcelWriter('./results(analysis)/bts_antenna_switch({}).xlsx'.format(today))   

#---------調整 assets_df 格式 -----------

assets_df = assets_df.drop(assets_df.loc[assets_df['設備名稱'].str.contains("共構")].index)
assets_df = assets_df[assets_df['財產名稱'].isin(property_name)]
assets_df['廠牌'] = assets_df['廠牌'].str.capitalize()
assets_df = assets_df[(~assets_df['編號'].str.contains('U')) & (assets_df['數量'] == "1")]
assets_df.sort_values(by = ['編號'],inplace = True)
assets_df.to_excel(writer,sheet_name ='assets',index=False)

assets_inuse_df = assets_df[(assets_df['設備狀態']=='使用中') & (assets_df['數量'] == "1")]  
assets_inuse_df = assets_inuse_df[assets_inuse_df['編號'].str.contains('N|L')]  #
assets_spare_df = assets_df[(assets_df['設備狀態']=='備援/備用') & (assets_df['數量']== "1")]
assets_stop_df = assets_df[(assets_df['設備狀態']=='停用') & (assets_df['數量']== "1")]
assets_loc_df = assets_df[(assets_df['設備狀態']=='佔位置') & (assets_df['數量']== "1")]
assets_loss_df = assets_df[(assets_df['設備狀態']=='已遺失') & (assets_df['數量']== "1")]

# assets_inuse_df.to_excel(writer,sheet_name ='assets(使用中)',index=False)
# assets_spare_df.to_excel(writer,sheet_name ='assets(備用)',index=False)
# assets_stop_df.to_excel(writer,sheet_name ='assets(停用)',index=False)
# assets_loc_df.to_excel(writer,sheet_name ='assets(佔位)',index=False)
# assets_loss_df.to_excel(writer,sheet_name ='assets(遺失)',index=False)

#---------調整 assets_df 格式 -----------

assets_data_df = assets_inuse_df[['編號','基地台名稱','廠牌','型式/號','財產編號','異動者']]
temp_df = assets_inuse_df['廠牌'] + '_'+ assets_inuse_df['型式/號'].astype(str)
assets_data_df = assets_data_df.drop(['廠牌','型式/號'],axis =1)
assets_data_df.insert(2,'天線型號(assets)',temp_df)
# assets_data_df.to_excel(writer,sheet_name ='assets',index=False)  ###

#---------調整 nmoss_df 格式 ------------
small_bts_exception =except_smallcell # read from config.ini
n = 0
for filename in nmoss_files:
    begin_df = pd.read_excel('./data_eNodebList/'+ filename,sheet_name = 0,dtype='str')
    list_colname = list(begin_df.head())
    if n==0:
        begin_df.rename(columns = {list_colname[2]:'編號'},inplace = True)
        temp4G_df = begin_df
        n = n + 1
    else:
        begin_df.rename(columns = {list_colname[1]:'編號'},inplace = True)
        temp5G_df =begin_df
        
temp4G_df = temp4G_df[['編號','扇區編號(sectorno)','基地台名稱(BName)','天線廠牌1(AntennaBrand1)','天線型號1(AntennaType1)','天線廠牌2(AntennaBrand2)','天線型號2(AntennaType2)','天線廠牌3(AntennaBrand3)','天線型號3(AntennaType3)']] 

#--調整 1支天線 tri-sector (splitter) 2 port [Commscope V360QS-C3-3XR]--
index1 = temp4G_df.loc[temp4G_df['天線型號1(AntennaType1)']=='V360QS-C3-3XR'].index
temp4G_df.loc[index1,'天線廠牌2(AntennaBrand2)':'天線型號3(AntennaType3)'] = np.NaN

#---------特例調整 (612299L: 用 Andrew_3X-V65A-3XR 天線，但卻於外部作 splitter)--
index5 = temp4G_df.loc[(temp4G_df['編號']=='612299L') & (temp4G_df['天線型號1(AntennaType1)']=='3X-V65A-3XR')].index
temp4G_df.loc[index5,'天線廠牌2(AntennaBrand2)':'天線型號3(AntennaType3)'] = np.NaN

#---------特例調整 bi-direction (612231L， 616810L)避免line 174 drop_duplicate 產生錯誤 -------#
index6 = temp4G_df.loc[(temp4G_df['編號']=='612231L') & (temp4G_df['天線型號1(AntennaType1)']=='80010664')].index
temp4G_df.loc[index6,'天線廠牌2(AntennaBrand2)':'天線型號3(AntennaType3)'] = np.NaN
index7 = temp4G_df.loc[(temp4G_df['編號']=='616810L') & (temp4G_df['天線型號1(AntennaType1)']=='80010864')].index
temp4G_df.loc[index7,'天線廠牌2(AntennaBrand2)':'天線型號3(AntennaType3)'] = np.NaN

#----調整 tri-sector天線 (3方向但只有一支天線 6 port)[BROADRADIO_LLLOX306R-D],[Andrew_3X-V65A-3XR],[COMMSCOPE_NNNOX310R]--
except_trisector_df = temp4G_df[temp4G_df['天線型號1(AntennaType1)'].isin(['LLLOX306R-D','3X-V65A-3XR','NNNOX310R'])]                       
except_trisector_df = except_trisector_df.drop_duplicates(subset=['編號','天線型號1(AntennaType1)'], keep='first') # except_trisector_df 後面必須加回去
index2 = temp4G_df.loc[temp4G_df['天線型號1(AntennaType1)'].isin(['LLLOX306R-D','3X-V65A-3XR','NNNOX310R'])].index
temp4G_df.drop(index2,axis = 0,inplace = True)
temp4G_df = pd.concat([temp4G_df,except_trisector_df])
temp4G_df = temp4G_df.reset_index(drop = True)

#----------調整 small cell 用 2支天線--------------------------------------------------
smallbts_with_twoant = temp4G_df[temp4G_df['編號'].isin(small_bts_exception)]
temp4G_df = temp4G_df.drop(temp4G_df.loc[temp4G_df['編號'].isin(small_bts_exception)].index)
#-----------------------------------------------------------------------------
temp4G_df = temp4G_df.drop_duplicates()
#-------2022.5.14再次將被包含之天線去除 Triple > birection >sector --------------#
temp4G_df.sort_values(by = ['編號','扇區編號(sectorno)','天線型號1(AntennaType1)','天線型號2(AntennaType2)','天線型號3(AntennaType3)'],inplace = True)
temp4G_df = temp4G_df.drop_duplicates(subset=['編號','扇區編號(sectorno)','天線型號1(AntennaType1)'], keep='first')
#-----------------------2022.5.14 修改至此--------------------------------#
temp4G_df = pd.concat([smallbts_with_twoant, temp4G_df], ignore_index=True)
temp4G_df = temp4G_df.reset_index(drop = True)
temp4G_df = temp4G_df.drop(['扇區編號(sectorno)'], axis=1)

#  5G 開始
temp5G_df = temp5G_df[['編號','細胞編號(CellID)','基地台名稱(BName)','天線廠牌1(AntennaBrand1)','天線型號1(AntennaType1)','天線廠牌2(AntennaBrand2)','天線型號2(AntennaType2)','天線廠牌3(AntennaBrand3)','天線型號3(AntennaType3)']]   

temp_df['check'] = temp5G_df['細胞編號(CellID)'].map(choice_char)
temp5G_df = temp5G_df.drop(['細胞編號(CellID)'], axis =1)
temp5G_df['check'] = temp_df['check']
temp5G_df = temp5G_df.drop_duplicates()
temp5G_df = temp5G_df.drop(['check'], axis =1)
# temp5G_df.to_excel(writer,sheet_name ='temp5G',index=False)  #

nmoss_orig_df = pd.concat([temp4G_df, temp5G_df])
nmoss_orig_df = nmoss_orig_df.reset_index(drop = True)

nmoss_ante1_df = nmoss_orig_df[['編號','基地台名稱(BName)','天線廠牌1(AntennaBrand1)','天線型號1(AntennaType1)']] 
nmoss_ante1_df = nmoss_ante1_df.dropna(axis=0)
nmoss_ante1_df['天線型號(nmoss)'] = nmoss_ante1_df['天線廠牌1(AntennaBrand1)'].str.capitalize() + '_' + nmoss_ante1_df['天線型號1(AntennaType1)'].astype(str)
nmoss_ante1_df = nmoss_ante1_df.drop(['天線廠牌1(AntennaBrand1)','天線型號1(AntennaType1)'],axis =1)

nmoss_ante2_df = nmoss_orig_df[['編號','基地台名稱(BName)','天線廠牌2(AntennaBrand2)','天線型號2(AntennaType2)']] 
nmoss_ante2_df = nmoss_ante2_df.dropna(axis=0)
nmoss_ante2_df['天線型號(nmoss)'] = nmoss_ante2_df['天線廠牌2(AntennaBrand2)'].str.capitalize() + '_' + nmoss_ante2_df['天線型號2(AntennaType2)'].astype(str)
nmoss_ante2_df = nmoss_ante2_df.drop(['天線廠牌2(AntennaBrand2)','天線型號2(AntennaType2)'],axis =1)

nmoss_ante3_df = nmoss_orig_df[['編號','基地台名稱(BName)','天線廠牌3(AntennaBrand3)','天線型號3(AntennaType3)']] 
nmoss_ante3_df = nmoss_ante3_df.dropna(axis=0)
nmoss_ante3_df['天線型號(nmoss)'] = nmoss_ante3_df['天線廠牌3(AntennaBrand3)'].str.capitalize() + '_' + nmoss_ante3_df['天線型號3(AntennaType3)'].astype(str)
nmoss_ante3_df = nmoss_ante3_df.drop(['天線廠牌3(AntennaBrand3)','天線型號3(AntennaType3)'],axis =1)


nmoss_data_df = pd.concat([nmoss_ante1_df, nmoss_ante2_df, nmoss_ante3_df])
nmoss_data_df = nmoss_data_df.reset_index(drop = True)
#
nmoss_data_df.sort_values(by = ['編號','天線型號(nmoss)'],inplace = True)
nmoss_data_df = nmoss_data_df.reset_index(drop = True)
nmoss_data_df = nmoss_data_df.drop(nmoss_data_df.loc[nmoss_data_df['天線型號(nmoss)'].str.contains(pat ='Nokia_')].index)
nmoss_data_df.to_excel(writer,sheet_name ='nmoss',index=False)  ##
nmoss_data_df['天線型號(nmoss)'] = nmoss_data_df['天線型號(nmoss)'].str.replace(')','',regex=True).str.replace('(','',regex=True).str.replace('UT45-N2','UT45',regex=True).str.replace('不可申請證照','',regex=True).str.replace('1710-2690','',regex=True)
#nmoss_data_df['天線型號(nmoss)'] = nmoss_data_df['天線型號(nmoss)'].str.replace('Gammanu','Gamma nu')

#---------合併 (assets) (nmoss) 兩資料庫內容 ------------
full_df = pd.concat([assets_data_df, nmoss_data_df])
full_df = full_df.reset_index(drop = True)
full_df = full_df[['編號','基地台名稱(BName)','天線型號(nmoss)','天線型號(assets)','財產編號','異動者']]
full_df.sort_values(by = ['編號','天線型號(nmoss)','天線型號(assets)'],inplace = True)
full_df = full_df.reset_index(drop = True)
# full_df.to_excel(writer,sheet_name ='combined',index=False)

# --------比較 (assets_df) (netact_df)兩資料庫內容 ----------- 
bts_id = list(full_df['編號'].unique())
for i in bts_id:
    nmoss_tmp_df = full_df[(full_df['編號']== i)&(full_df['天線型號(assets)'].isnull())]
    check1 = list(nmoss_tmp_df['天線型號(nmoss)'].str.replace('Gammanu','Gamma nu',regex=True).str.replace('Andrew_DBXLH-6565A-VTM','Commscope_DBXLH-6565A-VTM',regex=True).str.replace('Andrew_3X-V65A-3XR','Commscope_3X-V65A-3XR',regex=True).
                 str.replace('Argus_NOX310R','Commscope_NOX310R',regex=True).str.replace('LLPX202F0','LPX202F',regex=True).str.replace('Andrew_HBX-6516DS-VTM','Commscope_HBX-6516DS-VTM',regex=True).str.replace('Argus_NNNOX310R','Commscope_NNNOX310R',regex=True).str.replace('Andrew_DBXDH-6565B-VTM','Commscope_DBXDH-6565B-VTM',regex=True))
    check_1 = [s.upper() for s in check1]
    check_1.sort() 
    
    assets_tmp_df = full_df[(full_df['編號']== i)&(full_df['天線型號(nmoss)'].isnull())]
    check2 = list(assets_tmp_df['天線型號(assets)'].str.replace('Andrew_DBXLH-6565A-VTM','Commscope_DBXLH-6565A-VTM',regex=True).str.replace('Andrew_3X-V65A-3XR','CommScope_3X-V65A-3XR',regex=True).
                 str.replace('Argus_NOX310R','Commscope_NOX310R',regex=True).str.replace('LLPX202F0','LPX202F',regex=True).str.replace('Andrew_HBX-6516DS-VTM','Commscope_HBX-6516DS-VTM',regex=True).str.replace('Argus_NNNOX310R','Commscope_NNNOX310R',regex=True).str.replace('Andrew_DBXDH-6565B-VTM','Commscope_DBXDH-6565B-VTM',regex=True))
     
    check_2 =[s.upper() for s in check2] 
    check_2.sort()  
    
    if check_1 != check_2 :
        index1 = full_df.loc[full_df['編號']==i].index 
        full_df.loc[index1 ,'check'] = 'X'
        
        del_a = []
        a = check_2
        b = check_1
        for j in a:
            if j in b:
                del_a.append(j)
                b.remove(j)
        for j in del_a:
            a.remove(j)           
    
        full_df.loc[index1 ,'缺料'] = str(b)
        full_df.loc[index1 ,'多餘'] = str(a)
    else:
        index2 = full_df.loc[full_df['編號']==i].index   
        full_df.loc[index2 ,'check'] = 'O'
        
full_df.to_excel(writer,sheet_name ='天線調整',index = False)
worksheet = writer.sheets['天線調整']
worksheet.set_column("B:E",26)

#---缺料統計表-1--缺料統計表-----------------------#
lack_df= full_df[['編號','基地台名稱(BName)','異動者','check','缺料','多餘']]
lack1_df = lack_df.loc[lack_df.check=='X'].copy()
right=lack1_df.drop_duplicates(subset=['編號','缺料','多餘'],keep="last")
right=right[['編號','異動者','check','缺料','多餘']]
left=lack1_df.drop_duplicates(subset=['編號','缺料','多餘'],keep="first")
left=left[['編號','基地台名稱(BName)']]
result = pd.merge(left, right, how='outer', on=['編號'])
result['異動者']=result['異動者'].str.strip()
result['異動者']= result['異動者'].replace(to_replace =np.nan,value ='__無__',regex=True)
result.to_excel(writer,sheet_name ='天線調整-1',index = False)
worksheet = writer.sheets['天線調整-1']
worksheet.set_column("A:A",10)
worksheet.set_column("B:B",30)
worksheet.set_column("C:C",10)
worksheet.set_column("E:F",45)
#------缺料統計表-2---- 製作「設備淨值」分析--2022/7/29----------------#
df1 = result.loc[:,['異動者','缺料']]
df1['缺料量']=0
df1['缺料'] = df1['缺料'].str.split(',')
df1 = df1.explode('缺料')
df1['缺料']=df1['缺料'].str.replace(pat=']',repl='',regex=True).str.replace(pat='[',repl='',regex=True).str.replace(pat="'",repl="",regex=True)
df1['缺料']=df1['缺料'].str.strip()
analysis_table = df1.groupby(['異動者','缺料'],as_index=False).agg("count") 
analysis1_table = pd.DataFrame(analysis_table)
analysis1_table = analysis1_table.replace('',np.nan, regex=True)
analysis1_table.dropna(axis=0, how='any', inplace=True)
analysis1_table.rename(columns={'缺料': '天線型號'}, inplace=True)

df2 = result.loc[:,['異動者','多餘']]
df2['多餘量']=0
df2['多餘'] = df2['多餘'].str.split(',')
df2 = df2.explode('多餘')
df2['多餘']=df2['多餘'].str.replace(pat=']',repl='',regex=True).str.replace(pat='[',repl='',regex=True).str.replace(pat="'",repl="",regex=True)
df2['多餘']=df2['多餘'].str.strip()
analysis_table = df2.groupby(['異動者','多餘'],as_index=False).agg("count") 
analysis2_table = pd.DataFrame(analysis_table)
analysis2_table = analysis2_table.replace('',np.nan, regex=True)
analysis2_table.dropna(axis=0, how='any', inplace=True)
analysis2_table.rename(columns={'多餘': '天線型號'}, inplace=True)

analysis_table = pd.merge(analysis1_table,analysis2_table,how='outer')
analysis_table.sort_values(by=['異動者'],inplace=True)

#-----缺料統計表-3   -- 將固定資產 --備用，停用、佔位 加入----------
spare_df = assets_spare_df[['異動者','廠牌','型式/號']].copy()
spare_df['型式/號'] = spare_df['廠牌'].str.upper() +'_'+ spare_df['型式/號'].str.upper()
spare_df.rename(columns={'型式/號': '天線型號'}, inplace=True)
spare_df=spare_df.drop(labels='廠牌',axis=1) 
spare_df['備用']=0
spare_table = spare_df.groupby(['異動者','天線型號'],as_index=False).agg("count") 
spare2_table = pd.DataFrame(spare_table)
analysis_table = pd.merge(analysis_table,spare2_table,how='outer')

stop_df = assets_stop_df[['異動者','廠牌','型式/號']].copy()
stop_df['型式/號'] = stop_df['廠牌'].str.upper() +'_'+ stop_df['型式/號'].str.upper()
stop_df.rename(columns={'型式/號': '天線型號'}, inplace=True)
stop_df=stop_df.drop(labels='廠牌',axis=1) 
stop_df['停用']=0
stop_table = stop_df.groupby(['異動者','天線型號'],as_index=False).agg("count") 
stop2_table = pd.DataFrame(stop_table)
analysis_table = pd.merge(analysis_table,stop2_table,how='outer')

loc_df = assets_loc_df[['異動者','廠牌','型式/號']].copy()
loc_df['型式/號'] = loc_df['廠牌'].str.upper() +'_'+ loc_df['型式/號'].str.upper()
loc_df.rename(columns={'型式/號': '天線型號'}, inplace=True)
loc_df=loc_df.drop(labels='廠牌',axis=1) 
loc_df['佔位']=0
loc_table = loc_df.groupby(['異動者','天線型號'],as_index=False).agg("count") 
loc2_table = pd.DataFrame(loc_table)
analysis_table = pd.merge(analysis_table,loc2_table,how='outer')

loss_df = assets_loss_df[['異動者','廠牌','型式/號']].copy()
loss_df['型式/號'] = loss_df['廠牌'].str.upper() +'_'+ loss_df['型式/號'].str.upper()
loss_df.rename(columns={'型式/號': '天線型號'}, inplace=True)
loss_df=loss_df.drop(labels='廠牌',axis=1) 
loss_df['遺失']=0
loss_table = loss_df.groupby(['異動者','天線型號'],as_index=False).agg("count") 
loss2_table = pd.DataFrame(loss_table)
analysis_table = pd.merge(analysis_table,loss2_table,how='outer')

analysis_table.sort_values(by=['異動者','天線型號'],inplace=True)
analysis_table.to_excel(writer,sheet_name ='天線調整-2',index=False)
worksheet = writer.sheets['天線調整-2']
worksheet.set_column("A:A",10)
worksheet.set_column("B:B",36)
worksheet.set_column("C:D",10)
#---------------------------------------------------------------


# 成績值比較 先選以前值
last_record =find_last_record() # 使用 def Macro
last_record_df = pd.read_excel('./results(analysis)/'+ last_record[0],sheet_name = '完成數',dtype= {'日期': str})
#last_record_df['日期'] = last_record_df['日期'].dt.strftime('%Y-%m-%d')

grade_df = full_df.loc[full_df.check=='O'].copy()
grade_df['異動者'].fillna(method='bfill',inplace=True)
grade_df.drop(['基地台名稱(BName)','天線型號(nmoss)','天線型號(assets)','財產編號','check'],axis =1,inplace=True)
grade_df = grade_df.drop_duplicates()
grade1_df = pd.DataFrame(grade_df['異動者'].value_counts()).T
grade1_df['日期']=str(today)
grade1_df.insert(0, '日期', grade1_df.pop('日期'))

result = pd.concat([last_record_df, grade1_df],ignore_index= True)
result = result.reset_index(drop = True)

result.tail(15).to_excel(writer,sheet_name ='完成數',index = False)
worksheet = writer.sheets['完成數']
worksheet.set_column("A:A",12)

#============建立統計表(begin)===================#

assets_inuse_df['廠牌'] = assets_inuse_df['廠牌'].str.capitalize()
assets_inuse_df['廠牌'] = assets_inuse_df['廠牌'].str.replace('Argus','Commscope',regex=True)
assets_inuse_df['型式/號'] = assets_inuse_df['型式/號'].astype(str).str.upper()
assets_inuse_df['assets_使用中'] = assets_inuse_df['廠牌']+'_'+assets_inuse_df['型式/號'].astype(str)
assets_inuse_df['assets_使用中'] = assets_inuse_df['assets_使用中'].str.replace('Andrew_DBXDH-6565B-VTM','Commscope_DBXDH-6565B-VTM').str.replace('Andrew_3X-V65A-3XR','Commscope_3X-V65A-3XR').str.replace('Andrew_DBXLH-6565A-VTM','Commscope_DBXLH-6565A-VTM')
assets_inuse_count = assets_inuse_df['assets_使用中'].value_counts()
assets_inuse_stic = pd.DataFrame(assets_inuse_count)
assets_inuse_stic.index.name="天線型式"

assets_spare1_df = assets_spare_df.copy()
assets_spare1_df['廠牌'] = assets_spare1_df['廠牌'].str.capitalize()
assets_spare1_df['廠牌'] = assets_spare1_df['廠牌'].str.replace('Argus','Commscope')
assets_spare1_df['型式/號'] = assets_spare1_df['型式/號'].astype(str).str.upper()
assets_spare1_df['assets_備用'] = assets_spare1_df['廠牌']+'_'+assets_spare1_df['型式/號'].astype(str)
assets_spare1_df['assets_備用'] = assets_spare1_df['assets_備用'].str.replace('Andrew_DBXDH-6565B-VTM','Commscope_DBXDH-6565B-VTM').str.replace('Andrew_3X-V65A-3XR','Commscope_3X-V65A-3XR').str.replace('Andrew_DBXLH-6565A-VTM','Commscope_DBXLH-6565A-VTM')
assets_spare_count = assets_spare1_df['assets_備用'].value_counts()
assets_spare_stic = pd.DataFrame(assets_spare_count)
assets_spare_stic.index.name="天線型式"

assets_spare_north =  assets_spare1_df[assets_spare1_df['使用單位']=='北嘉義基維股']
assets_spare_north_cnt = assets_spare_north['assets_備用'].value_counts()
assets_spare_norstic = pd.DataFrame(assets_spare_north_cnt)
assets_spare_norstic.rename(columns={'assets_備用':'北基備用'},inplace=True)
assets_spare_norstic.index.name="天線型式"

assets_spare_south =  assets_spare1_df[assets_spare1_df['使用單位']=='南嘉義基維股']
assets_spare_south_cnt = assets_spare_south['assets_備用'].value_counts()
assets_spare_soustic = pd.DataFrame(assets_spare_south_cnt)
assets_spare_soustic.rename(columns={'assets_備用':'南基備用'},inplace=True)
assets_spare_soustic.index.name="天線型式"

assets_spare_qual =  assets_spare1_df[assets_spare1_df['使用單位']=='嘉義品改股']
assets_spare_qual_cnt = assets_spare_qual['assets_備用'].value_counts()
assets_spare_qualstic = pd.DataFrame(assets_spare_qual_cnt)
assets_spare_qualstic.rename(columns={'assets_備用':'品改備用'},inplace=True)
assets_spare_qualstic.index.name="天線型式"

assets_stop1_df = assets_stop_df.copy()
assets_stop1_df['廠牌'] = assets_stop1_df['廠牌'].str.capitalize()
assets_stop1_df['廠牌'] = assets_stop1_df['廠牌'].str.replace('Argus','Commscope')
assets_stop1_df['型式/號'] = assets_stop1_df['型式/號'].astype(str).str.upper()
assets_stop1_df['assets_停用'] = assets_stop1_df['廠牌']+'_'+assets_stop1_df['型式/號'].astype(str)
assets_stop1_df['assets_停用'] = assets_stop1_df['assets_停用'].str.replace('Andrew_DBXDH-6565B-VTM','Commscope_DBXDH-6565B-VTM').str.replace('Andrew_3X-V65A-3XR','Commscope_3X-V65A-3XR').str.replace('Andrew_DBXLH-6565A-VTM','Commscope_DBXLH-6565A-VTM')
assets_stop_count  = assets_stop1_df['assets_停用'].value_counts()
assets_stop_stic = pd.DataFrame(assets_stop_count)
assets_stop_stic.index.name="天線型式"

assets_stop_north =  assets_stop1_df[assets_stop1_df['使用單位']=='北嘉義基維股']
assets_stop_north_cnt = assets_stop_north['assets_停用'].value_counts()
assets_stop_norstic = pd.DataFrame(assets_stop_north_cnt)
assets_stop_norstic.rename(columns={'assets_停用':'北基停用'},inplace=True)
assets_stop_norstic.index.name="天線型式"

assets_stop_south =  assets_stop1_df[assets_stop1_df['使用單位']=='南嘉義基維股']
assets_stop_south_cnt = assets_stop_south['assets_停用'].value_counts()
assets_stop_soustic = pd.DataFrame(assets_stop_south_cnt)
assets_stop_soustic.rename(columns={'assets_停用':'南基停用'},inplace=True)
assets_stop_soustic.index.name="天線型式"

nmoss_data_df['天線型號(nmoss)'] = nmoss_data_df['天線型號(nmoss)'].str.replace('Andrew_DBXLH-6565A-VTM','Commscope_DBXLH-6565A-VTM').str.replace('Gammanu','Gamma nu')
nmoss_data_df['天線型號(nmoss)'] = nmoss_data_df['天線型號(nmoss)'].str.replace('LLPX202F0','LPX202F').str.replace('Andrew_DBXDH-6565B-VTM','Commscope_DBXDH-6565B-VTM').str.replace('Andrew_3X-V65A-3XR','Commscope_3X-V65A-3XR').str.replace('AARC','Aarc').str.replace('CommScope','Commscope').str.replace('COMMSCOPE','Commscope').str.replace('BROADRADIO','Broadradio')
nmoss_data_df['天線型號(nmoss)'] = nmoss_data_df['天線型號(nmoss)'].str.replace('Commscope_HBX-6516DS-VTM','Andrew_HBX-6516DS-VTM')

nmoss_data_df_count = nmoss_data_df['天線型號(nmoss)'].value_counts()
nmoss_data_df_stic = pd.DataFrame(nmoss_data_df_count)
nmoss_data_df_stic.rename(columns={'天線型號(nmoss)':'nmoss_使用中'},inplace=True)
nmoss_data_df_stic.index.name="天線型式"


statistics_df = assets_inuse_stic.join(assets_spare_stic,how='outer')
statistics_df = statistics_df.join(assets_spare_norstic,how='outer')
statistics_df = statistics_df.join(assets_spare_soustic,how='outer')
statistics_df = statistics_df.join(assets_spare_qualstic,how='outer')
statistics_df = statistics_df.join(assets_stop_stic,how='outer')
statistics_df = statistics_df.join(assets_stop_norstic,how='outer')
statistics_df = statistics_df.join(assets_stop_soustic,how='outer')
statistics_df = statistics_df.join(nmoss_data_df_stic,how='outer')
statistics_df = statistics_df.fillna(0)
statistics_df[['assets_使用中','assets_備用','北基備用','南基備用','品改備用','assets_停用','北基停用','南基停用','nmoss_使用中']] = statistics_df[['assets_使用中','assets_備用','北基備用','南基備用','品改備用','assets_停用','北基停用','南基停用','nmoss_使用中']].astype(int)


#============讀取電路板庫存量(in stock)================#
file1_name = find_earlier_in_stock() # 使用 def 
file1_path = "./data_in_stock/{}".format(file1_name[0])
in_stock_df = pd.read_excel(file1_path, sheet_name = '統計')
in_stock_df.fillna(value=0, inplace=True)
in_stock_df['庫存_嘉義']= in_stock_df['庫存_北基']+in_stock_df['庫存_南基']+in_stock_df['庫存_品改']
in_stock_df.set_index("天線型式" , inplace=True)

statistics_df = statistics_df.join(in_stock_df,how='outer')
statistics_df['財編缺額'] = statistics_df['nmoss_使用中'] + statistics_df['庫存_嘉義'] - statistics_df['assets_使用中']- statistics_df['assets_備用'] - statistics_df['assets_停用'] 
statistics_df = statistics_df[['assets_使用中','assets_備用','assets_停用','nmoss_使用中','庫存_嘉義','財編缺額',                               '北基備用','北基停用','庫存_北基','南基備用','南基停用','庫存_南基','品改備用','庫存_品改']]
statistics_df.fillna(value=0, inplace=True)

statistics_df = statistics_df.reset_index()
stat_style = statistics_df.style.applymap(lambda x: 'background-color:#ADD8E6', subset=["北基備用"])     .applymap(lambda x: 'background-color:#ADD8E6', subset=["庫存_北基"])     .applymap(lambda x: 'background-color:#ADD8E6', subset=["北基停用"])     .applymap(lambda x: 'background-color:#FFFF74', subset=["南基停用"])     .applymap(lambda x: 'background-color:#FFFF74', subset=["南基備用"])     .applymap(lambda x: 'background-color:#FFFF74', subset=["庫存_南基"]) 

stat_style.to_excel(writer,sheet_name ='統計表',index = False)
worksheet = writer.sheets['統計表']
worksheet.set_column("A:A",34)
worksheet.set_column("B:G",13)
worksheet.set_column("H:O",11)

#----------------統計表 大於 > 0  ---------------------#
# 1.assets_使用中 < nmoss_使用中 ，assets_備用&assets_停用 !=0  提出
# 2.nmoss_使用中 == nmoss_使用中 
# 備用 < 庫存， 停用 !=0 提出
# 備用 > 庫存  提出
# 3.assets_使用中 > nmoss_使用中 提出

A = statistics_df['財編缺額'] > 0
stat_great_0_df = statistics_df[A]

B = stat_great_0_df['assets_使用中'] < stat_great_0_df['nmoss_使用中']
C = stat_great_0_df['assets_備用'] != 0
D = stat_great_0_df['assets_停用'] != 0
E = B&(C|D)

F = stat_great_0_df['assets_使用中'] == stat_great_0_df['nmoss_使用中']
G = stat_great_0_df['南基備用'] < stat_great_0_df['庫存_南基']
H = stat_great_0_df['南基停用'] != 0 
I = stat_great_0_df['北基備用'] < stat_great_0_df['庫存_北基']
J = stat_great_0_df['北基停用'] != 0
K = stat_great_0_df['南基備用']> stat_great_0_df['庫存_南基']
L = stat_great_0_df['北基備用']> stat_great_0_df['庫存_北基']
M = F&((G&H)|(I&J)|K|L)

N = stat_great_0_df['assets_使用中'] > stat_great_0_df['nmoss_使用中']

stat_great_1_df = stat_great_0_df[E|M|N]

great_style = stat_great_1_df.style.applymap(lambda x: 'background-color:#ADD8E6', subset=["assets_備用"])       .applymap(lambda x: 'background-color:#ADD8E6', subset=["assets_停用"])       .applymap(lambda x: 'background-color:#FFFF74', subset=["assets_使用中"])       .applymap(lambda x: 'background-color:#FFFF74', subset=["nmoss_使用中"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["北基備用"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["南基備用"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_北基"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_南基"])

great_style.to_excel(writer,sheet_name ='缺額>0修正',index = False)
worksheet = writer.sheets['缺額>0修正']
worksheet.set_column("A:A",34)
worksheet.set_column("B:C",15)
worksheet.set_column("D:F",11)
worksheet.set_column("G:G",13)
worksheet.set_column("H:I",11)
worksheet.set_column("J:J",15)
worksheet.set_column("K:O",11)

#----------------統計表 小於 <=0  ----------------------#
# 1.assets_使用中 < nmoss_使用中  提出
# 2.assets_使用中 == nmoss_使用中 
# 備用 < 庫存， 停用 !=0 提出
# 備用 > 庫存  提出
# 3.assets_使用中 > nmoss_使用中 提出


A = statistics_df['財編缺額'] <= 0

stat_less_0_df = statistics_df[A]

B = stat_less_0_df['assets_使用中'] < stat_less_0_df['nmoss_使用中']

C = stat_less_0_df['assets_使用中'] == stat_less_0_df['nmoss_使用中']
D = stat_less_0_df['南基備用'] < stat_less_0_df['庫存_南基']
E = stat_less_0_df['南基停用'] != 0 
F = stat_less_0_df['南基備用'] >  stat_less_0_df['庫存_南基']
G = stat_less_0_df['北基備用'] < stat_less_0_df['庫存_北基']
H = stat_less_0_df['北基停用'] != 0 
I = stat_less_0_df['北基備用'] >  stat_less_0_df['庫存_北基']
J = C&((D&E)|F|(G&H)|I)

K = stat_less_0_df['assets_使用中'] > stat_less_0_df['nmoss_使用中']

stat_less_1_df = stat_less_0_df[B|J|K]

less_style = stat_less_1_df.style.applymap(lambda x: 'background-color:#ADD8E6', subset=["assets_使用中"])     .applymap(lambda x: 'background-color:#ADD8E6', subset=["nmoss_使用中"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["北基備用"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["南基備用"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_北基"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_南基"]) 

less_style.to_excel(writer,sheet_name ='缺額<=0修正',index = False)
worksheet = writer.sheets['缺額<=0修正']
worksheet.set_column("A:A",34)
worksheet.set_column("B:C",15)
worksheet.set_column("D:G",11)
worksheet.set_column("H:I",11)
worksheet.set_column("J:J",15)
worksheet.set_column("K:O",11)

#===============建立統計表(end)========================#
writer.save()




# In[ ]:





# In[ ]:




