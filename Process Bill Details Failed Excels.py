# -*- coding: utf-8 -*-
"""
Created on Mon Jul  4 10:58:39 2022

@author: mails
"""
import pandas as pd
import numpy as np
import re

path='Inv.HSS-024-Courtyard by Marriot(024)-Service Charge.xls'

#j='Inv.HSS-000-Lulu International Convention Centre-Trichur(066)-VFD.xls'
#i=j

df =pd.read_excel(i,header=None)

start_index=df.index[df.iloc[:,0]=='Sl.No.']
end_index=df.index[df.iloc[:,0]=='Bank Details:']

items=df.iloc[start_index[0]:end_index[0],:]

items=items.dropna(axis=1, how='all')
items.reset_index(drop=True, inplace=True)

items=items.fillna('')



column_names=(items.iloc[0,:].str.cat(items.iloc[1,:],)).tolist()
#items.iloc[:,0]

start=items.index[items.iloc[:,0]=='Sl.No.']
end=items.index[(items.iloc[:,0]=='1')+1]
items=items.drop(items.index[(list(range(start[0],end[0])))])



header_items=items.columns
for i in range (0,len(items.columns)):
    items.rename(columns={header_items[i]:column_names[i]},inplace=True)

items.dropna(how='all',axis=1,inplace=True)

bool_item=items['Sl.No.'].str.match(r'\d').astype(bool)

items=items[bool_item]

items=items.replace('',0)
items.columns==''
items=items.loc[:,items.columns!='']

bill_t=ProcessBillDetails(df)
invoice=FindInvoiceDetails(bill_t)
transportation=FindTransportationDetails(bill_t)
receiever=FindReceiver(bill_t)
consignee=FindConsignee(bill_t)
BillItems=FindBillItems(df)

Invoice_No =invoice['Invoice No '].values
Invoice_No=Invoice_No[0]                             
transportation['Invoice No ']=Invoice_No
receiever['Invoice No ']=Invoice_No
consignee['Invoice No ']=Invoice_No
BillItems['Invoice No ']=Invoice_No
BillItems['File Name']=i
