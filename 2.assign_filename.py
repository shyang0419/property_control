#!/usr/bin/env python
# coding: utf-8

# In[48]:


# 選擇 下載目錄的檔案 更名、移動目錄
import os
import time
from datetime import date
import datetime
import pandas as pd
import numpy as np
import smtplib
import shutil

#---------------讀取資料檔---------------------
today = date.today() 
files_list = os.listdir(r"C:\Users\lanjy\Downloads")
path = "C:/Users/lanjy/Downloads"
fname=['form.xls','eNodebList_4G.xlsx','eNodebList_5G.xlsx']

for i in fname:
    if i in files_list: 
        if i=='form.xls':
            new_file = i[0:i.index('.')]+'('+str(today)+')'+'.xls'
            os.rename(os.path.join(path, i), os.path.join(path, new_file))
            shutil.move(r'C:/Users/lanjy/Downloads/{}'.format(new_file),r'C:/Users/lanjy/exercise_1/data_property_form/{}'.format(new_file))
            
        else:
            new_file = i[0:i.index('.')]+'('+str(today)+')'+'.xlsx'
            os.rename(os.path.join(path, i), os.path.join(path, new_file))
            shutil.move(r'C:/Users/lanjy/Downloads/{}'.format(new_file),r'C:/Users/lanjy/exercise_1/data_eNodebList/{}'.format(new_file))
   

   


# In[ ]:





# In[ ]:




