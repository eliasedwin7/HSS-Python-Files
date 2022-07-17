# -*- coding: utf-8 -*-
"""
Created on Fri Jul  1 09:58:21 2022

@author: mails
"""


import pandas as pd
import numpy as np
import re


path='D:/2022-2023/Inv.HSS-000-Lulu International Convention Centre-Trichur(066)-VFD.xls'

df =pd.read_excel(path,header=None)

def FindBillItems(df):
    start_index=df.index[df.iloc[:,0]=='Sl.No.']
    end_index=df.index[df.iloc[:,0]=='Bank Details:']

    items=df.iloc[start_index[0]:end_index[0],:]

    items=items.dropna(axis=1, how='all')
    items.reset_index(drop=True, inplace=True)

    items=items.fillna('')



    column_names=(items.iloc[0,:].str.cat(items.iloc[1,:],)).tolist()

    items=items.drop(items.index[[0,1,2]])



    header_items=items.columns
    for i in range (0,len(items.columns)):
        items.rename(columns={header_items[i]:column_names[i]},inplace=True)
    
    items.dropna(how='all',axis=1,inplace=True)

    bool_item=items['Sl.No.'].str.match(r'\d').astype(bool)

    items=items[bool_item]
    items=items.replace('',0)

    items.info()
    
    return items
