#!/usr/bin/env python
# coding: utf-8

# In[4]:


# 本程式包括 3G ,4G ,5G
# 將程式load以下設備程式
#本程式與 pcb     比較: 多 => '5G基地台彙集設備'(此項qty所獨有)
#本程式與 antenna 比較: 少 => '4G/3G/2G行動通信室內涵蓋天線'

# FAMIS 財產名稱:'3G行動電話系統天線','２Ｇ及３Ｇ行動電話室外涵蓋型共用天線系統','4G/3G/2G行動通信室外涵蓋天線','5G/4G/3G室外涵蓋天線'
#               '4G行動寬頻基地台','5G基地台射頻模組','5G基地台基頻模組','5G基地台彙集設備'
# FAMIS row 203 排除 famis_df['規範'].str.contains("共構")                                            

# ASSETS財產名稱: 3G行動電話系統天線','２Ｇ及３Ｇ行動電話室外涵蓋型共用天線系統','4G/3G/2G行動通信室外涵蓋天線','5G/4G/3G室外涵蓋天線'
#               '4G行動寬頻基地台',,'5G基地台射頻模組','5G基地台基頻模組','5G基地台彙集設備'
# ASSETS row 135 排除 assets_df['設備名稱'].str.contains("共構")

import pandas as pd
import numpy as np
import time
from datetime import date
import os
import datetime
import re
import configparser as cp

def find_last_record():
    list1=['']
    path= './results(analysis)'
    date_list = os.listdir(path)
    now= date.today() 
    now1 = now - datetime.timedelta(days=1)
    for i in range(30):   
        famis_vs_assets = 'famis_vs_assets({}).xlsx'.format(now1)
        if famis_vs_assets in date_list:
            if list1[0] =='':
                list1[0] = famis_vs_assets
            break
        else:
            now1 = now1 - datetime.timedelta(days=1)
    return(list1)

def lookfor_num(value):
    pattern = '\*([1-9])'
    m = re.search(pattern, value)
    if m and m.group(1):
        #return int(m.group(1))
        return 1
    else:
        return 1


def parse_label(value):
    pattern = '廠牌--(Gamma NU|Gamma Nu|\w{2,13})'
    m = re.search(pattern, value)        
    if m and m.group(1):
        return m.group(1)
    else:
        return'沒填寫'
    
def style_only_4Gdevice(value):
    find_pattern = re.compile(r'[FA]{1}[A-KM-Z]{1}[2A-Z]{2,4}|Air[Ss]cale|Small Cell|Flexi Zone BTS|L1800 Micro RRH|LAA Micro RRH')
    match_result = find_pattern.findall(value)
    if match_result:
        return match_result
    else:
        return'無'


def style_except_4Gdevice(value):
    pattern = '型式--(.{1,32})'
    m = re.search(pattern, value)
    if m and m.group(1):
        return m.group(1)
    else:
        return'沒填寫'

def remove_virtual(value):
    list1= ['Small Cell','Flexi Zone BTS','L1800 Micro RRH','LAA Micro RRH','AirScale']
    for x in list1:
        if len(value)>1  and x in value:
            value.remove(x)
    return value



def trans_list(value):
    b=[]
    b.append(value)
    return b
    
def delete_string(value):
    nPos=value.find('】')
    value = value.replace(value[nPos:],'')
    return value
    
    
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

def find_earlier_basefamis(now1):
    path= './data_basefamis'
    date_list = os.listdir(path)
    
    i=0    
    for i in range(60):
        date_name = '嘉義中心財產({}).xls'.format(now1)
        if date_name in date_list:
            break 
        else:
            now1 = now1 - datetime.timedelta(days=1)
    return(date_name) 

def list_element_sep(value):
    value =sorted(str(value[0]).split(','))
    return value
#-------------------讀取config.ini------------------------------#  
filename = '.\config.ini'

# 配置文件读入
inifile = cp.ConfigParser()
inifile.read(filename, 'UTF-8')

# 读取 famis 部分 
property_name = inifile.get("famis", "property_name")
property_name = re.sub('[\r\n\t]', '', property_name)
property_name = property_name.split(',')

#----------------寫入檔案--------------------------    
select_day  = date.today() 
writer =pd.ExcelWriter('./results(analysis)/famis_vs_assets({}).xlsx'.format(select_day))

#----------------讀取 Assets form 固定資產檔案--------------------------
select_day  = date.today() 
used_day = find_earlier_form(select_day) # 使用 def 
file1_path = "./data_property_form/{}".format(used_day)
assets_df = pd.read_excel(file1_path,usecols=['設備名稱','廠牌','型式/號','裝設地點','財產編號','財產名稱','使用單位','數量','異動者'],dtype={'型式/號':str})
## assets_df = assets_df[assets_df['設備名稱'].isin(['全向型天線(行通)','拋物面/平板型天線(行通)','室內涵蓋天線(行通)',
##                                            '指向型天線(行通)','基地台設備/AAS(行通)','蓄電池組(電力)','基站擴充軟硬體/施工費(行通)'])]

# assets_df = assets_df.drop(assets_df.loc[assets_df['設備名稱'].str.contains("共構") & (assets_df['財產名稱']=='3G行動電話系統天線')].index)
assets_df['設備名稱'].fillna(value='未填寫', inplace=True)
assets_df = assets_df.drop(assets_df.loc[assets_df['設備名稱'].str.contains("共構")].index)
#------2022.4.8 增加3G行動電話收發訊系統 FXEB, FXDB 進入-----------#    
assets_FXEB_df = assets_df[(assets_df['財產名稱'].isin(['3G行動電話收發訊系統'])) & 
                             (assets_df['型式/號'].str.contains("FXEB"))].copy() # 電路板只用於 4G 

assets_FXDB_df = assets_df[(assets_df['財產名稱'].isin(['3G行動電話收發訊系統'])) & 
                              (assets_df['型式/號'].str.contains("FXDB"))].copy() # 電路板 3G/4G皆使用
assets_FXDB_df = assets_FXDB_df[assets_FXDB_df['裝設地點'].str.contains("L", case=True)]

assets_FRGY_df = assets_df[(assets_df['財產名稱'].isin(['3G行動電話收發訊系統'])) & 
                              (assets_df['型式/號'].str.contains("FRGY"))].copy() # 電路板 3G/4G皆使用
assets_FRGY_df = assets_FRGY_df[assets_FRGY_df['裝設地點'].str.contains("L", case=True)]

# assets_df = assets_df[assets_df['財產名稱'].isin(['3G行動電話系統天線','２Ｇ及３Ｇ行動電話室外涵蓋型共用天線系統',\
#     '4G/3G/2G行動通信室外涵蓋天線','4G行動寬頻基地台','5G/4G/3G室外涵蓋天線','5G基地台射頻模組','5G基地台基頻模組',\
#     '5G基地台彙集設備','4G行動寬頻系統共構設備','4G/3G/2G行動通信室內涵蓋天線','２Ｇ及３Ｇ行動電話室內涵蓋型共用天線系統'])]

assets_df = assets_df[assets_df['財產名稱'].isin(property_name)]

assets_df = pd.concat([assets_df,assets_FXEB_df,assets_FXDB_df,assets_FRGY_df])
# #-----------------------------------------------------------------#

#assets_df.drop(assets_df[(assets_df['數量']==0) | (assets_df['使用單位']=='嘉義品改股')].index,axis =0,inplace =True)
assets_df.drop(assets_df[(assets_df['數量']==0)].index,axis =0,inplace =True)
assets_df['使用單位'] = assets_df['使用單位'].str.replace('嘉義品改股','品改').str.replace('北嘉義基維股','北基').str.replace('南嘉義基維股','南基')

assets_df['型式/號'] = assets_df['型式/號'].str.replace(' ','',regex=True).str.replace('(','',regex=True).str.replace(')','',regex=True).str.replace(';',',',regex=True).str.replace('//','',regex=True)
assets_df['型式/號'] = assets_df['型式/號'].str.replace('7720','7720.00',regex=True).str.replace('+',',',regex=True).str.replace(']','',regex=True).str.replace('[','',regex=True).str.replace('/','',regex=True).str.replace('AirScale','',regex=True).str.replace('Flexi','',regex=True)
assets_df['型式/號'] = assets_df['型式/號'].str.replace('RFSI','I',regex=True) # HB 廠牌
assets_df['廠牌'] = assets_df['廠牌'].str.replace('OptoConn奈德國際','Optoconn')

assets_df.drop(columns= ['設備名稱'],inplace =True)
assets_df.rename(columns = {'數量':'數量(assets)','型式/號':'型式/號(assets)','廠牌':'廠牌(assets)'},inplace = True)


assets_df['型式/號(assets)']= assets_df['型式/號(assets)'].map(trans_list)
assets_df['型式/號(assets)']= assets_df['型式/號(assets)'].map(list_element_sep)
assets_df = assets_df[['財產編號','廠牌(assets)','型式/號(assets)','數量(assets)','財產名稱','使用單位','異動者']]
assets_df.to_excel(writer, sheet_name = 'assets_all',index=False)
assets_df.drop(columns= ['財產名稱','使用單位'],inplace =True)

#----------------讀取  famis 固定資產檔案---------------------
select_day  = date.today() 
used_day = find_earlier_basefamis(select_day) # 使用 def 
file2_path = "./data_basefamis/{}".format(used_day)
famis_df = pd.read_excel(file2_path, header=4, usecols=['財產編號＋列帳年月','使用單位','主從財產別','財產名稱','規範'],dtype={'規範':str})
famis_df.drop(famis_df[(famis_df['主從財產別']==2)].index,axis =0,inplace =True)



#------2022.4.8 增加3G行動電話收發訊系統 FXEB, FXDB, FRGY 進入-----------#
famis_FXEB_df = famis_df[(famis_df['財產名稱'].isin(['3G行動電話收發訊系統'])) & 
                          (famis_df['規範'].str.contains("FXEB"))].copy() # 電路板只用於 4G. 財產在 3/4G

famis_FXDB_df = famis_df[(famis_df['財產名稱'].isin(['3G行動電話收發訊系統'])) & 
                             (famis_df['規範'].str.contains("FXDB"))].copy()
famis_FXDB_df = famis_FXDB_df[famis_FXDB_df['規範'].str.contains("\d{6}L", regex=True)]  # 電路板用於 3/4G .財產在 3/4G

famis_FRGY_df = famis_df[(famis_df['財產名稱'].isin(['3G行動電話收發訊系統'])) & 
                             (famis_df['規範'].str.contains("FRGY"))].copy()
famis_FRGY_df = famis_FRGY_df[famis_FRGY_df['規範'].str.contains("\d{6}L", regex=True)]   # 電路板用於 3/4G.財產在 3G

famis_df = famis_df.drop(famis_df.loc[famis_df['規範'].str.contains("共構")].index)

# famis_df = famis_df[famis_df['財產名稱'].isin(['3G行動電話系統天線','２Ｇ及３Ｇ行動電話室外涵蓋型共用天線系統',\
#            '4G/3G/2G行動通信室外涵蓋天線','4G行動寬頻基地台','5G/4G/3G室外涵蓋天線','5G基地台射頻模組','5G基地台基頻模組',\
#            '5G基地台彙集設備','4G行動寬頻系統共構設備','4G/3G/2G行動通信室內涵蓋天線','２Ｇ及３Ｇ行動電話室內涵蓋型共用天線系統'])]

famis_df = famis_df[famis_df['財產名稱'].isin(property_name)]

famis_df = pd.concat([famis_df,famis_FXEB_df,famis_FXDB_df,famis_FRGY_df])
# famis_df = pd.concat([famis_df,famis_FXEB_df])

#-----------------------------------------------------------------#
famis_df['使用單位'] = famis_df['使用單位'].str.replace('5953-D31J02','').str.replace('5953-D31J03','').str.replace('5953-D31J04','')
famis_df['使用單位'] = famis_df['使用單位'].str.replace('(','',regex=True).str.replace(')','',regex=True)
famis_df['使用單位'] = famis_df['使用單位'].str.strip()

famis_df['使用單位'] = famis_df['使用單位'].str.replace('嘉義中心二股品改','品改').str.replace('嘉義營運中心北嘉義基維','北基').str.replace('嘉義營運中心南嘉義基維','南基')
famis_df.rename(columns={'財產編號＋列帳年月': '財產編號'}, inplace=True)
famis_df['財產編號'] = famis_df['財產編號'].map(lambda x:x.replace(x[-6:],''))
famis_df.drop(columns= ['主從財產別'],inplace =True)
#famis_df.to_excel(writer, sheet_name = 'Famis資產',index=False)

#---------(4G 行動寬頻基地台 使用: 「廠牌」，品名) ， (other device，天線使用:  「廠牌」、型式) -----
famis_df['廠牌(Famis)'] = famis_df['規範'].map(parse_label) # for 「廠牌」
famis_df['廠牌(Famis)'] = famis_df['廠牌(Famis)'].str.replace('NSN','Nokia')
famis_df['廠牌(Famis)'] = famis_df['廠牌(Famis)'].str.replace('OptoConn奈德國際','Optoconn')
#---------(分離 4G 行動寬頻基地台 與其他 device，天線 -----------
exact4G_famis = famis_df[famis_df['財產名稱']=='4G行動寬頻基地台'].copy()
famis_df.drop(famis_df[famis_df['財產名稱']=='4G行動寬頻基地台'].index,inplace = True)
famis_df['型式/號(Famis)'] = famis_df['規範'].map(style_except_4Gdevice)
famis_df['型式/號(Famis)'] = famis_df['型式/號(Famis)'].map(delete_string)
famis_df['數量(Famis)'] = famis_df['規範'].map(lookfor_num)
famis_df['型式/號(Famis)']= famis_df['型式/號(Famis)'].str.replace('/','',regex=True).str.replace('AirScale','',regex=True).str.replace('x1','',regex=True)
famis_df['型式/號(Famis)']= famis_df['型式/號(Famis)'].str.replace(' ','',regex=True).str.replace('\*1','',regex=True).str.replace('(','',regex=True).str.replace(')','',regex=True)
famis_df['型式/號(Famis)']= famis_df['型式/號(Famis)'].str.replace('RFSI','I',regex=True) # HB 廠牌
famis_df['型式/號(Famis)']= famis_df['型式/號(Famis)'].map(trans_list)
famis_df.to_excel(writer,sheet_name ='no_4Gdevice(Famis)',index=False)

exact4G_famis['型式/號(Famis)'] = exact4G_famis['規範'].map(style_only_4Gdevice)
exact4G_famis['型式/號(Famis)'] = exact4G_famis['型式/號(Famis)'].map(lambda x: sorted(list(set(x))))
exact4G_famis['型式/號(Famis)'] = exact4G_famis['型式/號(Famis)'].map(remove_virtual)

exact4G_famis['數量(Famis)'] = exact4G_famis['規範'].map(lookfor_num)
exact4G_famis.to_excel(writer,sheet_name ='4Gdevice(Famis)',index=False)

famis_df = pd.concat([famis_df,exact4G_famis])
famis_df = famis_df.reset_index(drop = True)
famis_df.to_excel(writer,sheet_name ='Famis_all',index=False)

#famis_df.to_excel('famis.xlsx',index=False)
#exact4G_famis.to_excel('exact4Gdevice.xlsx',index=False

# --------------將固定資產 assets 與 Famis資料合併 以 join outer 方式進行---------------

famis_df['型式/號(Famis)'] = famis_df['型式/號(Famis)'].map(lambda x: x[0] if len(x)==1 else x[0]+','+x[1])
assets_df['型式/號(assets)'] = assets_df['型式/號(assets)'].map(lambda x: x[0] if len(x)==1 else x[0]+','+x[1])
famis_df['型式/號(Famis)'] = famis_df['型式/號(Famis)'].str.replace('AWHQCASiR','AWHQC',regex=True)
assets_df['型式/號(assets)'] = assets_df['型式/號(assets)'].str.replace('AWHQCASiR','AWHQC',regex=True)

famis_df = famis_df.set_index("財產編號")
assets_df = assets_df.set_index("財產編號")

both_df = famis_df.join(assets_df,how ='outer')
both_df = both_df.reset_index()

#------------------比較 Famis 與 assets 資料-----------
#both_df = both_df[~both_df['財產編號'].str.contains('-001')]

both_df['廠牌(Famis)'] = both_df['廠牌(Famis)'].str.title()
both_df['廠牌(assets)'] = both_df['廠牌(assets)'].str.title()



both_df['type(Famis)'] = both_df['型式/號(Famis)']
# both_df['型式/號(assets)'] = both_df['型式/號(assets)'].str.replace('7720','7720.00') # Powerwave中馳 固定資產匯出會刪掉.00 所以補上
both_df['type(assets)'] = both_df['型式/號(assets)']

both_df['type(Famis)'] = both_df['type(Famis)'].str.replace('Small Cell','FW2EHB').str.replace('L1800 Micro RRH','AHEJ').str.replace('LAA Micro RRH','AZRB')
both_df['type(Famis)'] = both_df['type(Famis)'].str.replace('Airscale','ASIA').str.replace('AirScale','ASIA').str.replace('Flexi Zone BTS','FWHN')



index1 = both_df.loc[both_df['廠牌(Famis)']!=both_df['廠牌(assets)']].index
both_df.loc[index1,'廠牌Check']='X'
index1 = both_df.loc[both_df['廠牌(Famis)']==both_df['廠牌(assets)']].index
both_df.loc[index1,'廠牌Check']='O'

index1 = both_df.loc[both_df['type(Famis)'].str.upper()!=both_df['type(assets)'].str.upper()].index
both_df.loc[index1,'型式Check']='X'
index1 = both_df.loc[both_df['type(Famis)'].str.upper()==both_df['type(assets)'].str.upper()].index
both_df.loc[index1,'型式Check']='O'

index1 = both_df.loc[both_df['數量(Famis)']!=both_df['數量(assets)']].index
both_df.loc[index1,'數量Check']='X'
index1 = both_df.loc[both_df['數量(Famis)']==both_df['數量(assets)']].index
both_df.loc[index1,'數量Check']='O'



both_df = both_df[['財產編號','使用單位','財產名稱','規範','廠牌(Famis)','廠牌(assets)','廠牌Check','型式/號(Famis)','型式/號(assets)','型式Check','數量(Famis)','數量(assets)','數量Check','異動者']]
both_df.to_excel(writer,sheet_name ='比對',index=False)
worksheet = writer.sheets['比對']
worksheet.set_column("A:A",20)
worksheet.set_column("B:B",13)
worksheet.set_column("C:M",15)
#-------------------未完成數-------------------------
# 成績值比較 先選以前值
last_record =find_last_record() # 使用 def Macro
last_record_df = pd.read_excel('./results(analysis)/'+ last_record[0],sheet_name = '未完成數',dtype= {'日期': str})

both_df['異動者'].fillna('None', inplace=True)
names = sorted(list(both_df['異動者'].unique()))

brand_df = both_df.groupby(['異動者','廠牌Check']).size()
type_df = both_df.groupby(['異動者','型式Check']).size()
qty_df = both_df.groupby(['異動者','數量Check']).size()

dict_brand =dict(brand_df)
dict_type =dict(type_df)
dict_qty =dict(qty_df)

brand_list=[]
type_list =[]
qty_list =[]

for i in names:
    brand_list.append(dict_brand.get((i,'X'),0))
    type_list.append(dict_type.get((i,'X'),0))
    qty_list.append(dict_qty.get((i,'X'),0))

today = str(date.today())     
grade_df = pd.DataFrame([brand_list,type_list,qty_list],columns=names)
check = ['廠牌', '型式','數量']
grade_df.insert(0,"check",check, True)
grade_df.insert(0,"日期",today, True)

result = pd.concat([last_record_df, grade_df],ignore_index= True)
result = result.reset_index(drop = True)

result.tail(21).to_excel(writer,sheet_name ='未完成數',index=False)
worksheet = writer.sheets['未完成數']
worksheet.set_column("A:A",10)

#============建立統計表(begin)============5/20 line 351 'index':'assets_設備'=======#
assets_df = assets_df.reset_index() 
assets_df['assets_設備'] = assets_df['廠牌(assets)'].str.capitalize() + '_'+ assets_df['型式/號(assets)'].str.upper()
assets_reduce_df = assets_df[['assets_設備','數量(assets)']]
assets_stic_df = assets_reduce_df.groupby('assets_設備')

famis_df = famis_df.reset_index()     
famis_df['Famis_設備'] = famis_df['廠牌(Famis)'].str.capitalize() + '_'+famis_df['型式/號(Famis)'].str.upper()
famis_reduce_df = famis_df[['Famis_設備','數量(Famis)']]
famis_stic_df = famis_reduce_df.groupby('Famis_設備')

statistics_df = assets_stic_df.sum().join(famis_stic_df.sum(),how='outer')
statistics_df.fillna(0,inplace=True)
statistics_df = statistics_df.reset_index()
statistics_df.rename(columns = {'index':'assets_設備'},inplace = True)
statistics_df['缺額']= statistics_df['數量(assets)'] - statistics_df['數量(Famis)']
statistics_df.to_excel(writer,sheet_name ='統計表',index=False)
worksheet = writer.sheets['統計表']
worksheet.set_column("A:A",34)
worksheet.set_column("B:C",15)
worksheet.set_column("D:D",12)


#============建立統計差異表===================#
# 成績值比較 先選以前值
last_record =find_last_record() # 使用 def Macro
last_record_df = pd.read_excel('./results(analysis)/'+ last_record[0],sheet_name = '統計表') #,dtype= {'日期': str}
last_record_df.rename(columns = {'數量(assets)':'數量(assets_last)','數量(Famis)':'數量(Famis_last)','缺額':'缺額(last)'},inplace = True)

last_record_df = last_record_df.set_index("assets_設備")
statistics_df = statistics_df.set_index("assets_設備")

both_df = last_record_df.join(statistics_df,how ='outer')
both_df = both_df.reset_index()

index1 = both_df.loc[(both_df['數量(assets_last)']!=both_df['數量(assets)'])|(both_df['數量(Famis_last)']!=both_df['數量(Famis)'])].index
both_df.loc[index1,'無差異']='           X'
index1 = both_df.loc[(both_df['數量(assets_last)']==both_df['數量(assets)'])&(both_df['數量(Famis_last)']==both_df['數量(Famis)'])].index
both_df.loc[index1,'無差異']='           O'



both_df.to_excel(writer,sheet_name ='統計表差異',index=False)

worksheet = writer.sheets['統計表差異']
worksheet.set_column("A:A",34)
worksheet.set_column("B:C",16)
worksheet.set_column("D:F",12)
worksheet.set_column("G:G",6)
worksheet.set_column("H:H",8)

writer.save()




# In[ ]:





# In[ ]:




