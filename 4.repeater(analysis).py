#!/usr/bin/env python
# coding: utf-8

# In[1]:


# 分析 repeater 天線，其殘值 

import pandas as pd
import numpy as np
import time
from datetime import date
import datetime
import os
import re

def check0(value):
    if value==0 :
        return '殘值=0'
    else:
        return '殘值>0'
    
def change_na_column1(value):
    list0 = value
    j=''
    n=-1
    for i in list0:
        n=n+1
        if i==j:
            list0[n]=''
        else:
            j=i  
    return list0

def change_na_column2(value1,value2):
    compare1=change_na_column1(value1)
    list0 = value2
    j=''
    n=-1
    for i in list0:
        n=n+1
        if (i==j) & (compare1[n]==''):
            list0[n]=''
        else:
            j=i  
    return list0

def change_na_column3(value1,value2,value3):
    compare1=change_na_column1(value1)
    compare2=change_na_column2(value1,value2)
    list0 = value3
    j=''
    n=-1
    for i in list0:
        n=n+1
        if (i==j) & (compare2[n]==''):
            list0[n]=''
        else:
            j=i  
    return list0

 # --------讀取 assets 基地台資料庫 (assets_df) ----------- 
today = date.today() 
file1_name = 'bts_figure({}).xlsx'.format(today)
file1_path = "./results(analysis)/{}".format(file1_name)
assets_df = pd.read_excel(file1_path, sheet_name = '轉發器B')

# ---------寫入資料庫  ---------------------------------
today = date.today() 
writer =pd.ExcelWriter('./results(analysis)/repeater_analysis({}).xlsx'.format(today))   
#---------區分 設備狀態  -----------
assets_df.drop(assets_df[assets_df['數量']==0].index,axis=0,inplace=True)
assets_residual_df = assets_df[assets_df['本月設備淨值']==0] 
assets_residual_df.to_excel(writer,sheet_name ='殘值0',index=False)

assets_repeater_df = assets_df[assets_df['財產名稱'].isin(['4G行動寬頻轉發器','5G/4G轉發器'])]
assets_repeater_ant_df = assets_df[~assets_df['財產名稱'].isin(['4G行動寬頻轉發器','5G/4G轉發器'])]

assets_repeater_df.to_excel(writer,sheet_name ='主體',index=False)
assets_repeater_ant_df.to_excel(writer,sheet_name ='天線',index=False) 

assets_repeater_use_df = assets_repeater_df[assets_repeater_df['設備狀態']=='使用中']
assets_repeater_use_df.to_excel(writer,sheet_name ='使用中',index=False)

assets_repeater_spare_df = assets_repeater_df[assets_repeater_df['設備狀態']=='備援/備用']
assets_repeater_spare_df.to_excel(writer,sheet_name ='備用',index=False)

assets_repeater_spare_df = assets_repeater_df[assets_repeater_df['設備狀態']=='停用']
assets_repeater_spare_df.to_excel(writer,sheet_name ='停用',index=False)

assets_repeater_spare_df = assets_repeater_df[assets_repeater_df['設備狀態']=='佔位置']
assets_repeater_spare_df.to_excel(writer,sheet_name ='佔位置',index=False)


#---------------1. 製作「repeater設備淨值」分析------------------#
analysis_df = assets_repeater_df.loc[:,['設備狀態','本月設備淨值','使用單位','編號']]
analysis_df.rename(columns = {'編號':'數量'},inplace = True)
analysis_df['本月設備淨值']=analysis_df['本月設備淨值'].map(check0)
analysis_table = analysis_df.groupby(['使用單位','設備狀態','本月設備淨值'],as_index=False).agg("count") 
analysis1_table = pd.DataFrame(analysis_table)
#-------將 table以第一個為主，其餘填NA-----
analysis1_table['使用單位'] =change_na_column1(analysis1_table['使用單位'].copy())          
analysis1_table['設備狀態'] =change_na_column2(analysis1_table['使用單位'].copy(),analysis1_table['設備狀態'].copy())
#----------------------------------------------------------------
analysis1_table.to_excel(writer,sheet_name ='主體淨值表',index=False)
worksheet = writer.sheets['主體淨值表']
worksheet.set_column("A:C",15)
worksheet.set_column("D:D",8)

#----------------2.製作「主體種類」分析-------------------------------#
analysis_df = assets_repeater_df.loc[:,['設備狀態','本月設備淨值','使用單位','編號','廠牌','型式/號']]
analysis_df.rename(columns = {'編號':'數量'},inplace = True)
analysis_df['本月設備淨值']=analysis_df['本月設備淨值'].map(check0)
analysis_table = analysis_df.groupby(['使用單位','設備狀態','本月設備淨值','廠牌','型式/號',],as_index=False).agg("count") 
analysis1_table = pd.DataFrame(analysis_table)
#-------將 table重複處，以第一個為主，其餘填NA-----
analysis1_table['使用單位'] = change_na_column1(analysis1_table['使用單位'].copy())              
analysis1_table['設備狀態'] = change_na_column2(analysis1_table['使用單位'].copy(),analysis1_table['設備狀態'].copy())               
analysis1_table['本月設備淨值'] = change_na_column3(analysis1_table['使用單位'].copy(),analysis1_table['設備狀態'].copy(),                                             analysis1_table['本月設備淨值'].copy())
#-----------------------------------------------------
analysis1_table.to_excel(writer,sheet_name ='主體種類表',index=False)
worksheet = writer.sheets['主體種類表']
worksheet.set_column("A:D",15)
worksheet.set_column("E:E",20)
worksheet.set_column("F:F",8)

#---------------3. 製作天線淨值 分析------------------#
analysis_df = assets_repeater_ant_df.loc[:,['設備狀態','本月設備淨值','使用單位','編號']]
analysis_df.rename(columns = {'編號':'數量'},inplace = True)
analysis_df['本月設備淨值']=analysis_df['本月設備淨值'].map(check0)
analysis_table = analysis_df.groupby(['使用單位','設備狀態','本月設備淨值'],as_index=False).agg("count") 
analysis1_table = pd.DataFrame(analysis_table)
#-------將 table以第一個為主，其餘填NA-----
analysis1_table['使用單位'] =change_na_column1(analysis1_table['使用單位'].copy())          
analysis1_table['設備狀態'] =change_na_column2(analysis1_table['使用單位'].copy(),analysis1_table['設備狀態'].copy())
#----------------------------------------------------------------
analysis1_table.to_excel(writer,sheet_name ='天線淨值表',index=False)
worksheet = writer.sheets['天線淨值表']
worksheet.set_column("A:C",15)
worksheet.set_column("D:D",8)

#----------------4.製作「天線種類」分析-------------------------------#
analysis_df = assets_repeater_ant_df.loc[:,['設備狀態','本月設備淨值','使用單位','編號','廠牌','型式/號']]
analysis_df.rename(columns = {'編號':'數量'},inplace = True)
analysis_df['本月設備淨值']=analysis_df['本月設備淨值'].map(check0)
analysis_table = analysis_df.groupby(['使用單位','設備狀態','本月設備淨值','廠牌','型式/號',],as_index=False).agg("count") 
analysis1_table = pd.DataFrame(analysis_table)
#-------將 table重複處，以第一個為主，其餘填NA-----
analysis1_table['使用單位'] = change_na_column1(analysis1_table['使用單位'].copy())              
analysis1_table['設備狀態'] = change_na_column2(analysis1_table['使用單位'].copy(),analysis1_table['設備狀態'].copy())               
analysis1_table['本月設備淨值'] = change_na_column3(analysis1_table['使用單位'].copy(),analysis1_table['設備狀態'].copy(),                                             analysis1_table['本月設備淨值'].copy())
analysis1_table.to_excel(writer,sheet_name ='天線種類表',index=False)
worksheet = writer.sheets['天線種類表']
worksheet.set_column("A:D",15)
worksheet.set_column("E:E",20)
worksheet.set_column("F:F",8)

writer.save()




# In[ ]:





# In[ ]:




