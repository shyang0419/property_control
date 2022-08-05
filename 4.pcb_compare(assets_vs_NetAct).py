#!/usr/bin/env python
# coding: utf-8

# In[17]:


# 將財產分配 load 程式
# PCB-assets天線採用 => 財產名稱:'4G行動寬頻基地台','5G基地台射頻模組','5G基地台基頻模組'

import pandas as pd
import numpy as np
import time
from datetime import date
import datetime
import re
import os
import warnings
import configparser as cp
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def parse_btsid(value):
    pattern = 'R\d{4,6}|\d{4,7}[uUlLnNgG]'
    m = re.search(pattern, value)
    if m and m.group(0):
        return m.group(0)
    else:
        return '沒填寫'

def detect_star(value):
    pattern = '([A-Za-z]{4,6})\*2|[A-Za-z0-9,;\[\]+]{4,24}'  # 還必須考慮,;[+]+等
    m = re.search(pattern, str(value))
    if m and m.group(1):
        return m.group(1) + ' ' + m.group(1)
    elif m and m.group(0):
        return m.group(0) 
    
def find_earlier_HW_Link():
    list1=['','']
    path= './data_HW_Link'
    date_list = os.listdir(path)
    n=0
    now1= date.today() 
    for i in range(30):   
        date_name_4G = '4G HW_Link_{}E.xlsm'.format(now1).replace('-','')
        date_name_5G = '5G HW_Link_{}E.xlsm'.format(now1).replace('-','')
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

def assets_duplicate(data):
    total_len=len(data)
    del_index=[]
    for i in range(total_len):
        if len(data[i])>1:
            res = data[i]
            del_index.append(res)
            output=[res[i:i + 1] for i in range(0, len(res), 1)]
            for j in range(len(output)):
                data.append(output[j])     

    for i in del_index:
        data.remove(i)
    data.sort()  
    return data

def find_last_record():
    list1=['']
    path= './results(analysis)'
    date_list = os.listdir(path)
    now= date.today() 
    now1 = now - datetime.timedelta(days=1)
    for i in range(30):   
        antenna_df = 'bts_circuitboard_switch({}).xlsx'.format(now1)
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
        pcb_df = '電路板庫存({}).xlsx'.format(now)
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

# 读取 circuitboard 部分 
property_name = inifile.get("circuitboard", "property_name")
property_name = re.sub('[\r\n\t]', '', property_name)
property_name = property_name.split(',')

not_check_pcb = inifile.get("circuitboard", "not_check_pcb")
not_check_pcb = re.sub('[\r\n\t]', '', not_check_pcb)
not_check_pcb = not_check_pcb.split(',')

# --------讀取 assets 基地台資料庫 (assets_df) ----------- 
today = date.today() 
file1_name = 'bts_figure({}).xlsx'.format(today)
file1_path = "./results(analysis)/{}".format(file1_name)
assets_df = pd.read_excel(file1_path, sheet_name = '基地台設備')

# ---------寫入資料庫  ---------------------------------
today = date.today() 
writer =pd.ExcelWriter('./results(analysis)/bts_circuitboard_switch({}).xlsx'.format(today))   

#---------調整 assets_df 格式 -----------

# ----將 3G「FRGY」電路板放入 4G 5G 財產中--------------------------------#
assets_FRGY_df = assets_df[(assets_df['財產名稱'].isin(['3G行動電話收發訊系統'])) & 
                             (assets_df['型式/號'].str.contains("FRGY"))].copy()
assets_FRGY_df['型式/號'] = assets_FRGY_df['型式/號'].str.replace("Flexi-RRH","",regex=True).str.replace('(','',regex=True).str.replace(')','',regex=True)
assets_FRGY_to_spare_df  = assets_FRGY_df[~assets_FRGY_df['編號'].str.contains("L")] # assets_FRGY_to_spare_df :3G 的 FRGY
assets_FRGY_to_spare_df.to_excel(writer,sheet_name ='可用3G(FRGY)',index=False) 

index1 = assets_FRGY_df.loc[~assets_FRGY_df['編號'].str.contains("L")].index   
assets_FRGY_df.drop(index1, inplace = True)  # assets_FRGY_df :3G FRGY 用於 4G

#------------將 3G「FXEB」電路板放入 4G 5G 財產中--------------------------#
assets_FXEB_df = assets_df[(assets_df['財產名稱'].isin(['3G行動電話收發訊系統'])) & 
                             (assets_df['型式/號'].str.contains("FXEB"))].copy()

#------------將 3G「FXDB」電路板放入 4G 5G 財產中--------------------------#
assets_FXDB_df = assets_df[(assets_df['財產名稱'].isin(['3G行動電話收發訊系統'])) & 
                             (assets_df['型式/號'].str.contains("FXDB"))].copy()
assets_FXDB_df['型式/號'] = assets_FXDB_df['型式/號'].str.replace("Flexi-RRH","",regex=True).str.replace('(','',regex=True).str.replace(')','',regex=True)
assets_FXDB_to_spare_df  = assets_FXDB_df[~assets_FXDB_df['編號'].str.contains("L")] # assets_FXDB_to_spare_df :3G 的 FXDB
assets_FXDB_to_spare_df.to_excel(writer,sheet_name ='可用3G(FXDB)',index=False)

index1 = assets_FXDB_df.loc[~assets_FXDB_df['編號'].str.contains("L")].index   
assets_FXDB_df.drop(index1, inplace = True)  # assets_FXDB_df :3G FXDB 用於 4G

#--------------將 4G 設備用於3G 去除掉----------------------------------------#
assets_df = assets_df[assets_df['財產名稱'].isin(property_name)]

for_3G_assets_df = assets_df[assets_df['編號'].str.contains("U")]
for_3G_assets_df.to_excel(writer,sheet_name ='4G(assets)用於3G',index=False) 

assets_df = assets_df.drop(assets_df.loc[assets_df['編號'].str.contains("U")].index)
assets_df = assets_df.reset_index(drop = True)                                                        

#---------------將3G 的財產 FRGY、FXEB、FXDB 併入 4/5G 系統-----------------------#
assets_df = pd.concat([assets_df, assets_FRGY_df,assets_FXEB_df,assets_FXDB_df])
#-------------------------------------------------------------------------------#

assets_df = assets_df.reset_index(drop = True)
assets_df['型式/號'] = assets_df['型式/號'].str.replace("/","",regex=True).str.replace("AirScale","",regex=True).str.replace('(','',regex=True).str.replace(')','',regex=True).str.replace('_','',regex=True).str.replace(';',',',regex=True).str.replace('AWHQCASiR','AWHQC',regex=True)

assets_df = assets_df[assets_df['數量']== 1]
assets_df.sort_values(by = ['編號'],inplace = True)
assets_df.to_excel(writer,sheet_name ='assets',index=False)

assets_inuse_df = assets_df[(assets_df['設備狀態']=='使用中') & (assets_df['數量']== 1)]
assets_spare_df = assets_df[(assets_df['設備狀態']=='備援/備用') & (assets_df['數量']== 1)]
assets_stop_df = assets_df[(assets_df['設備狀態']=='停用') & (assets_df['數量']== 1)]
assets_loc_df = assets_df[(assets_df['設備狀態']=='佔位置') & (assets_df['數量']== 1)]

assets_inuse_df = assets_inuse_df.sort_values(by=['編號'])
assets_inuse_df.rename(columns={'基地台名稱':'基地台名(assets)','型式/號':'型式/號(assets)'},inplace =True)
assets_inuse_df.sort_values(by = ['編號','基地台名(assets)','型式/號(assets)'],inplace = True)
                         
# assets_inuse_df.to_excel(writer,sheet_name ='assets(使用中)',index=False)
# assets_spare_df.to_excel(writer,sheet_name ='assets(備用)',index=False)
# assets_stop_df.to_excel(writer,sheet_name ='assets(停用)',index=False)
# assets_loc_df.to_excel(writer,sheet_name ='assets(佔位)',index=False)

#---------調整 netact_df 格式 ------------

# =====讀取 NetAct 基地台資料庫 (netact_df)=========#
n = 0
all_files = find_earlier_HW_Link() # 使用 def 
for filename in all_files:
    begin_df = pd.read_excel('./data_HW_Link/'+filename,sheet_name = 'Main',header = 5 )
    list_colname = list(begin_df.head())
    begin_df.rename(columns = {list_colname[0]:'編號',list_colname[1]:'基地台名稱',list_colname[4]:'硬體元件',list_colname[5]:'Serial Number'},inplace = True)
    begin_df = begin_df[['編號','基地台名稱','硬體元件','Serial Number']]
    begin_df['編號'] = begin_df['編號'].astype(str)
    begin_df['Serial Number'] = begin_df['Serial Number'].astype(str)
    begin_df['L_column'] = 'L'
    begin_df['N_column'] = 'N'
    if n == 0 :
        begin_df['編號'] = begin_df[['編號','L_column']].apply(''.join, axis=1)
        netact_df = begin_df
        n = n+1
    else:
        begin_df['編號'] = begin_df[['編號','N_column']].apply(''.join, axis=1)
        netact_df = netact_df.append(begin_df)
        
values = {'硬體元件': 'NULL','Serial Number': 'NULL'}
netact_df.fillna(value=values,inplace = True)

netact_df = netact_df.drop(['L_column','N_column'],axis = 1)  
netact_df = netact_df.reset_index(drop=True)

netact_df = netact_df.drop(netact_df.loc[netact_df['硬體元件'].str.contains(pat ='-1')].index)
netact_df = netact_df.drop(netact_df.loc[netact_df['硬體元件'].str.contains(pat ='473764A.102')].index)
netact_df = netact_df.drop(netact_df.loc[netact_df['硬體元件'].str.contains(pat ='APHA')].index)

netact_df = netact_df.drop(netact_df.loc[~netact_df['Serial Number'].str.contains(pat = '[A-Z]',regex= True)].index)

netact_df = netact_df[~netact_df['硬體元件'].isin(['FR2EB','FR2HB','FWEA_FREA','FWHN_FRHN'])]
netact_df['硬體元件'] = netact_df['硬體元件'].str.replace('ASIB AirScale Common','ASIB',regex=True).str.replace('S4-90M-R1-V3','',regex=True).str.replace('(','',regex=True).str.replace(')','',regex=True).str.replace('S4-90M-R1-V2','',regex=True)
netact_df['硬體元件'] = netact_df['硬體元件'].str.replace('RV4S4-65A-R6','',regex=True).str.strip()


netact_df = netact_df.dropna(subset =['基地台名稱'])
netact_df = netact_df.reset_index(drop = True)
netact_df.rename(columns={'基地台名稱':'基地台名(NetAct)','硬體元件':'型式/號(NetAct)'},inplace =True)
netact_df.sort_values(by=['編號','型式/號(NetAct)'],inplace = True)
netact_df = netact_df.reset_index(drop = True)
netact_df.to_excel(writer,sheet_name ='NetAct',index=False)   


#---------合併 (assets_df) (netact_df) 兩資料庫內容 ------------
full_df = pd.concat([netact_df, assets_inuse_df])
full_df = full_df.reset_index(drop = True)
full_df = full_df[['編號','基地台名(NetAct)','型式/號(NetAct)','型式/號(assets)','財產編號','異動者']]
full_df.sort_values(by = ['編號','型式/號(NetAct)','型式/號(assets)'],inplace = True)
full_df = full_df.reset_index(drop = True)


# --------比較 (assets_df) (netact_df)兩資料庫內容 ----------- 
full_df['型式/號(assets)'] = full_df['型式/號(assets)'].map(detect_star)
full_df['型式/號(assets)'] = full_df['型式/號(assets)'].str.replace(',',' ',regex=True).str.replace(';',' ',regex=True).str.replace('[','',regex=True).str.replace(']','',regex=True).str.replace('+',' ',regex=True)
# full_df.to_excel(writer,sheet_name ='基地台設備調整')

bts_id = list(full_df['編號'].unique())
for i in bts_id:
    netact_tmp_df = full_df[(full_df['編號']== i)&(full_df['型式/號(assets)'].isnull())]
    check1 = sorted(list(netact_tmp_df['型式/號(NetAct)'].str.split(' ')))
    assets_tmp_df = full_df[(full_df['編號']== i)&(full_df['型式/號(NetAct)'].isnull())]
    check2 = sorted(list(assets_tmp_df['型式/號(assets)'].str.split(' ')))    
    check2 = assets_duplicate(check2)
    
    # --------去除掉 FBBA & FBBC 電路板等的比較----
#     check1 = [x for x in check1 if x not in (['FBBA'],['FBBC'],['36CHDWDM'],['1830VWM'],['7250IXR'],\
#                                              ['7250IXR10PORT_10G'],['CWDM_SFP'],['EX3400-24T'],['18CHCOMPACTCWDM'])]
#     check2 = [x for x in check2 if x not in (['FBBA'],['FBBC'],['36CHDWDM'],['1830VWM'],['7250IXR'],\
#                                              ['7250IXR10PORT_10G'],['CWDM_SFP'],['EX3400-24T'],['18CHCOMPACTCWDM'])]  
    check1 = [x for x in check1 if x not in(eval(','.join(not_check_pcb)))]
    check2 = [x for x in check2 if x not in(eval(','.join(not_check_pcb)))]  



    # ------------------------------------------ 
    
    if check1 != check2:
        index1 = full_df.loc[full_df['編號']==i].index   
        full_df.loc[index1 ,'check'] = 'X'

        del_a = []
        a = check2
        b = check1
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
        
full_df.to_excel(writer,sheet_name ='設備調整',index=False)
worksheet = writer.sheets['設備調整']
worksheet.set_column("A:A",9)
worksheet.set_column("B:E",18)

#-----缺料統計表-1 ---------------------2022/7/29-------------#
lack_df= full_df[['編號','基地台名(NetAct)','異動者','check','缺料','多餘']]
lack1_df = lack_df.loc[lack_df.check=='X'].copy()
right=lack1_df.drop_duplicates(subset=['編號','缺料','多餘'],keep="last")
right=right[['編號','異動者','check','缺料','多餘']]
left=lack1_df.drop_duplicates(subset=['編號','缺料','多餘'],keep="first")
left=left[['編號','基地台名(NetAct)']]
result = pd.merge(left, right, how='outer', on=['編號'])
result['異動者']= result['異動者'].replace(to_replace =np.nan,value ='__無__',regex=True)
result.to_excel(writer,sheet_name ='設備調整-1',index = False)
worksheet = writer.sheets['設備調整-1']
worksheet.set_column("A:A",9)
worksheet.set_column("B:B",30)
worksheet.set_column("C:D",10)
worksheet.set_column("E:F",30)
#-----缺料統計表-2   - 製作「設備淨值」分析--2022/7/29----------------#
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
analysis1_table.rename(columns={'缺料': '型式/號'}, inplace=True)

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
analysis2_table.rename(columns={'多餘': '型式/號'}, inplace=True)

analysis_table = pd.merge(analysis1_table,analysis2_table,how='outer')

#-----缺料統計表-3   -- 將固定資產 --備用，停用、佔位 加入----------
spare_df = assets_spare_df[['異動者','型式/號']].copy()
spare_df['備用']=0
spare_table = spare_df.groupby(['異動者','型式/號'],as_index=False).agg("count") 
spare2_table = pd.DataFrame(spare_table)
analysis_table = pd.merge(analysis_table,spare2_table,how='outer')

stop_df  = assets_stop_df[['異動者','型式/號']].copy()
stop_df['停用']=0
stop_table = stop_df.groupby(['異動者','型式/號'],as_index=False).agg("count") 
stop2_table = pd.DataFrame(stop_table)
analysis_table = pd.merge(analysis_table,stop2_table,how='outer')

loc_df   = assets_loc_df[['異動者','型式/號']].copy()
loc_df['佔位']=0
loc_table = loc_df.groupby(['異動者','型式/號'],as_index=False).agg("count") 
loc2_table = pd.DataFrame(loc_table)
analysis_table = pd.merge(analysis_table,loc2_table,how='outer')

analysis_table.sort_values(by=['異動者','型式/號'],inplace=True)
analysis_table.to_excel(writer,sheet_name ='設備調整-2',index=False)
worksheet = writer.sheets['設備調整-2']
worksheet.set_column("A:A",9)
worksheet.set_column("B:B",15)
worksheet.set_column("C:D",10)
#----------------------------------------------------------------



# 成績值比較 先選以前值
last_record =find_last_record() # 使用 def Macro
last_record_df = pd.read_excel('./results(analysis)/'+ last_record[0],sheet_name = '完成數',dtype= {'日期': str})

grade_df = full_df.loc[full_df.check=='O'].copy()
grade_df['異動者'].fillna(method='bfill',inplace=True)
grade_df.drop(['基地台名(NetAct)','型式/號(NetAct)','型式/號(assets)','財產編號','check'],axis =1,inplace=True)
grade_df = grade_df.drop_duplicates()
grade1_df = pd.DataFrame(grade_df['異動者'].value_counts()).T
grade1_df['日期']=str(today)
grade1_df.insert(0, '日期', grade1_df.pop('日期'))

result = pd.concat([last_record_df, grade1_df],ignore_index= True)

result.tail(15).to_excel(writer,sheet_name ='完成數',index = False)
worksheet = writer.sheets['完成數']
worksheet.set_column("A:C",10)

#============建立統計表(begin)===================#
assets_inuse_count = assets_inuse_df['型式/號(assets)'].value_counts()
assets_inuse_stic = pd.DataFrame(assets_inuse_count)
assets_inuse_stic.rename(columns={'型式/號(assets)':'assets_使用中'},inplace=True)
assets_inuse_stic.index.name="電路板"

assets_spare_count = assets_spare_df['型式/號'].value_counts()
assets_spare_stic = pd.DataFrame(assets_spare_count)
assets_spare_stic.rename(columns={'型式/號':'assets_備用'},inplace=True)
assets_spare_stic.index.name="電路板"

assets_spare_north =  assets_spare_df[assets_spare_df['使用單位']=='北嘉義基維股']
assets_spare_north_cnt = assets_spare_north['型式/號'].value_counts()
assets_spare_norstic = pd.DataFrame(assets_spare_north_cnt)
assets_spare_norstic.rename(columns={'型式/號':'北基備用'},inplace=True)
assets_spare_norstic.index.name="電路板"

assets_spare_south =  assets_spare_df[assets_spare_df['使用單位']=='南嘉義基維股']
assets_spare_south_cnt = assets_spare_south['型式/號'].value_counts()
assets_spare_soustic = pd.DataFrame(assets_spare_south_cnt)
assets_spare_soustic.rename(columns={'型式/號':'南基備用'},inplace=True)
assets_spare_soustic.index.name="電路板"

assets_stop1_df = assets_stop_df.copy()
# assets_stop1_df['型式/號']= assets_stop1_df['型式/號'].str.replace('AirScale','',regex=True).str.replace('(','',regex=True).str.replace(')','',regex=True)
assets_stop_count  = assets_stop1_df['型式/號'].value_counts()
assets_stop_stic = pd.DataFrame(assets_stop_count)
assets_stop_stic.rename(columns={'型式/號':'assets_停用'},inplace=True)
assets_stop_stic.index.name="電路板"

assets_stop_north =  assets_stop1_df[assets_stop1_df['使用單位']=='北嘉義基維股']
assets_stop_north_cnt = assets_stop_north['型式/號'].value_counts()
assets_stop_norstic = pd.DataFrame(assets_stop_north_cnt)
assets_stop_norstic.rename(columns={'型式/號':'北基停用'},inplace=True)
assets_stop_norstic.index.name="電路板"

assets_stop_south =  assets_stop1_df[assets_stop1_df['使用單位']=='南嘉義基維股']
assets_stop_south_cnt = assets_stop_south['型式/號'].value_counts()
assets_stop_soustic = pd.DataFrame(assets_stop_south_cnt)
assets_stop_soustic.rename(columns={'型式/號':'南基停用'},inplace=True)
assets_stop_soustic.index.name="電路板"
assets_stop_soustic

netact_df_count = netact_df['型式/號(NetAct)'].value_counts()
netact_df_stic = pd.DataFrame(netact_df_count)
netact_df_stic.rename(columns={'型式/號(NetAct)':'NetAct_使用中'},inplace=True)
netact_df_stic.index.name="電路板"

statistics_df = assets_inuse_stic.join(assets_spare_stic,how='outer')
statistics_df = statistics_df.join(assets_spare_norstic,how='outer')
statistics_df = statistics_df.join(assets_spare_soustic,how='outer')
statistics_df = statistics_df.join(assets_stop_stic,how='outer')
statistics_df = statistics_df.join(assets_stop_norstic,how='outer')
statistics_df = statistics_df.join(assets_stop_soustic,how='outer')
statistics_df = statistics_df.join(netact_df_stic,how='outer')
statistics_df = statistics_df.fillna(0)
statistics_df[['assets_使用中','assets_備用','北基備用','南基備用','assets_停用','北基停用','南基停用','NetAct_使用中']]= statistics_df[['assets_使用中','assets_備用','北基備用','南基備用','assets_停用','北基停用','南基停用','NetAct_使用中']].astype(int)

#============讀取電路板庫存量(in stock)================#
file1_name = find_earlier_in_stock() # 使用 def 
file1_path = "./data_in_stock/{}".format(file1_name[0])
in_stock_df = pd.read_excel(file1_path, sheet_name = '統計')
in_stock_df.fillna(value=0, inplace=True)
in_stock_df['庫存_嘉義']= in_stock_df['庫存_北基']+in_stock_df['庫存_南基']
in_stock_df.set_index("電路板" , inplace=True)

statistics_df = statistics_df.join(in_stock_df,how='outer')
statistics_df['財編缺額'] = statistics_df['NetAct_使用中'] + statistics_df['庫存_嘉義'] - statistics_df['assets_使用中']- statistics_df['assets_備用'] - statistics_df['assets_停用'] 
statistics_df = statistics_df[['assets_使用中','assets_備用','assets_停用','NetAct_使用中','庫存_嘉義','財編缺額',                               '北基備用','北基停用','庫存_北基','南基備用','南基停用','庫存_南基']]
statistics_df.fillna(value=0, inplace=True)


statistics_df = statistics_df.reset_index()
stat_style = statistics_df.style.applymap(lambda x: 'background-color:#ADD8E6', subset=["北基備用"])     .applymap(lambda x: 'background-color:#ADD8E6', subset=["庫存_北基"])     .applymap(lambda x: 'background-color:#ADD8E6', subset=["北基停用"])     .applymap(lambda x: 'background-color:#FFFF74', subset=["南基停用"])     .applymap(lambda x: 'background-color:#FFFF74', subset=["南基備用"])     .applymap(lambda x: 'background-color:#FFFF74', subset=["庫存_南基"]) 
stat_style.to_excel(writer,sheet_name ='統計表',index = False)
worksheet = writer.sheets['統計表']
worksheet.set_column("A:A",10)
worksheet.set_column("B:B",15)
worksheet.set_column("C:D",12)
worksheet.set_column("E:F",14)
worksheet.set_column("G:M",10)

#----------------統計表 大於 > 0  ---------------------#
# 1.assets_使用中 < Netact_使用中 ，assets_備用 &assets_停用 !=0  提出
# 2.nmoss_使用中 == Netact_使用中 
# 備用 < 庫存， 停用 !=0 提出
# 備用 > 庫存  提出
# 3.assets_使用中 > Netact_使用中 提出

A = statistics_df['財編缺額'] > 0
stat_great_0_df = statistics_df[A]

B = stat_great_0_df['assets_使用中'] < stat_great_0_df['NetAct_使用中']
C = stat_great_0_df['assets_備用'] != 0
D = stat_great_0_df['assets_停用'] != 0
E = B&(C|D)

F = stat_great_0_df['assets_使用中'] == stat_great_0_df['NetAct_使用中']
G = stat_great_0_df['南基備用'] < stat_great_0_df['庫存_南基']
H = stat_great_0_df['南基停用'] != 0 
I = stat_great_0_df['北基備用'] < stat_great_0_df['庫存_北基']
J = stat_great_0_df['北基停用'] != 0
K = stat_great_0_df['南基備用']> stat_great_0_df['庫存_南基']
L = stat_great_0_df['北基備用']> stat_great_0_df['庫存_北基']
M = F&((G&H)|(I&J)|K|L)

N = stat_great_0_df['assets_使用中'] > stat_great_0_df['NetAct_使用中']

stat_great_1_df = stat_great_0_df[E|M|N]

great_style = stat_great_1_df.style.applymap(lambda x: 'background-color:#ADD8E6', subset=["assets_備用"])       .applymap(lambda x: 'background-color:#ADD8E6', subset=["assets_停用"])       .applymap(lambda x: 'background-color:#FFFF74', subset=["assets_使用中"])       .applymap(lambda x: 'background-color:#FFFF74', subset=["NetAct_使用中"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["北基備用"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["南基備用"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_北基"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_南基"])

great_style.to_excel(writer,sheet_name ='缺額>0修正',index = False)
worksheet = writer.sheets['缺額>0修正']
worksheet.set_column("A:A",10)
worksheet.set_column("B:B",15)
worksheet.set_column("C:H",12)
worksheet.set_column("I:I",15)
worksheet.set_column("J:M",12)

#----------------統計表 小於 < = 0  ---------------------#
# 1.assets_使用中 < NetAct_使用中  提出
# 2.assets_使用中 == NetAct_使用中 
# 備用 < 庫存， 停用 !=0 提出
# 備用 > 庫存  提出
# 3.assets_使用中 > NetAct_使用中 提出

A = statistics_df['財編缺額'] <= 0

stat_less_0_df = statistics_df[A]

B = stat_less_0_df['assets_使用中'] < stat_less_0_df['NetAct_使用中']

C = stat_less_0_df['assets_使用中'] == stat_less_0_df['NetAct_使用中']
D = stat_less_0_df['南基備用'] < stat_less_0_df['庫存_南基']
E = stat_less_0_df['南基停用'] != 0 
F = stat_less_0_df['南基備用'] >  stat_less_0_df['庫存_南基']
G = stat_less_0_df['北基備用'] < stat_less_0_df['庫存_北基']
H = stat_less_0_df['北基停用'] != 0 
I = stat_less_0_df['北基備用'] >  stat_less_0_df['庫存_北基']
J = C&((D&E)|F|(G&H)|I)

K = stat_less_0_df['assets_使用中'] > stat_less_0_df['NetAct_使用中']

stat_less_1_df = stat_less_0_df[B|J|K]

less_style = stat_less_1_df.style.applymap(lambda x: 'background-color:#ADD8E6', subset=["assets_使用中"])     .applymap(lambda x: 'background-color:#ADD8E6', subset=["NetAct_使用中"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["北基備用"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["南基備用"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_北基"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_南基"]) 

less_style.to_excel(writer,sheet_name ='缺額<=0修正',index = False)
worksheet = writer.sheets['缺額<=0修正']
worksheet.set_column("A:A",10)
worksheet.set_column("B:B",15)
worksheet.set_column("C:H",12)
worksheet.set_column("I:I",15)
worksheet.set_column("J:M",12)

#===============建立統計表(end)========================#

# 關閉寫入檔案
writer.save()


# In[ ]:





# In[ ]:




