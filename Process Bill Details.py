# -*- coding: utf-8 -*-
"""
Created on Fri Jul  1 12:11:28 2022

@author: mails
"""


import pandas as pd
import numpy as np
import re
import os



    
def ProcessBillDetails(df):    

    start_index=df.index[df.iloc[:,0]=='GSTIN: 32BBZPP6232J1Z4']
    end_index=df.index[df.iloc[:,0]=='State Code        :']

    bill=df.iloc[start_index[0]:end_index[0],:]

    bill=bill.dropna(axis=1, how='all')

    bill_t=bill.transpose()
    bill_t=bill_t.dropna(axis=1, how='all')
    bill_t.reset_index(drop=True, inplace=True)


    bill_t.drop([7,9,10,11],axis=1,inplace=True)
    
    return bill_t


#Process Transportation

# Invoice No.                       :		
# Invoice Date                        :		



def RemoveSpecialCharColumnNames(new_header):
    for i in range(0,len(new_header)):
        cleanString = re.sub('\W+',' ', new_header[i] )
        new_header[i]=cleanString
    return new_header

def FindInvoiceDetails(bill_t):
    InvoiceDet=bill_t.iloc[0:3,:5]
    InvoiceDet.dropna(axis=1, how='all',inplace=True)
    InvoiceDet.dropna(how='all',inplace=True)

    #remove special char from column names
    new_header= InvoiceDet.iloc[0].tolist()
    new_header=RemoveSpecialCharColumnNames(new_header)
    invoice=InvoiceDet[1:]
    invoice.columns=new_header
    
    return invoice
 
# Transportation Mode
# Vehicle No.                            
# Date of Supply                        
# Place of Supply                       
# Payment Terms    
 
def FindTransportationDetails(bill_t):             
    TransportDT =bill_t.iloc[-3:,:5]
    TransportDT.dropna( how='all',inplace=True)    
    new_header= TransportDT.iloc[0].tolist()    
    new_header=RemoveSpecialCharColumnNames(new_header)
    Transport=TransportDT[1:]
    Transport.columns=new_header
    return Transport
                   
   
# Details of Receiver - Billed to:
def FindReceiver(bill_t):
    ReceieverDT =bill_t.iloc[0:2,-9:]
    ReceieverDT.reset_index(drop=True, inplace=True)
    ReceieverDT.iloc[0,0]='Name'
    col=ReceieverDT.loc[:, ReceieverDT.columns != 19]
    ReceieverDT['concat'] = pd.Series(col.fillna('').values.tolist()).str.join(',')
    ReceieverDT['concat'][0]='Complete Address'
    
    ReceieverDT=ReceieverDT.fillna('')
    ReceieverDT.dropna( how='all',inplace=True)
    
    new_header= ReceieverDT.iloc[0].tolist()    
    new_header=RemoveSpecialCharColumnNames(new_header)
    Receiever=ReceieverDT[1:]
    Receiever.columns=new_header
    
    Receiever=Receiever.replace('',np.nan)
    Receiever.dropna(how='all',axis=1,inplace=True)
    
    #Receiever.columns
    return Receiever[['Name','Complete Address']]
    
# Details of Consignee - Shipped to:
   
    
def FindConsignee(bill_t):
    ConsigneeDT =bill_t.iloc[-3:,-9:]
    ConsigneeDT.reset_index(drop=True, inplace=True)
    ConsigneeDT.dropna( how='all',inplace=True)
    #col=ConsigneeDT.columns
    ConsigneeDT.iloc[0,0]='Name'
    col=ConsigneeDT.loc[:, ConsigneeDT.columns != 19]
    ConsigneeDT['concat'] = pd.Series(col.fillna('').values.tolist()).str.join(',')
    ConsigneeDT['concat'][0]='Complete Address'
    ConsigneeDT=ConsigneeDT.fillna('')
    

    new_header= ConsigneeDT.iloc[0].tolist()    
    new_header=RemoveSpecialCharColumnNames(new_header)
    Consignee=ConsigneeDT[1:]
    Consignee.columns=new_header
    
    Consignee=Consignee.replace('',np.nan)
    Consignee.dropna(how='all',axis=1,inplace=True)
    Consignee = Consignee.rename(columns={'Name': 'Consignee Name', 'Complete Address': 'Consignee Complete Address'})
    return Consignee[['Consignee Name','Consignee Complete Address']]

def FindBillItems(df):
    start_index=df.index[df.iloc[:,0]=='Sl.No.']
    end_index=df.index[df.iloc[:,0]=='Bank Details:']

    items=df.iloc[start_index[0]:end_index[0],:]

    items=items.dropna(axis=1, how='all')
    items.reset_index(drop=True, inplace=True)

    items=items.fillna('')



    column_names=(items.iloc[0,:].str.cat(items.iloc[1,:],)).tolist()

    #items=items.drop(items.index[[0,1,2]])
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
    items=items.loc[:,items.columns!='']
    #items.info()
    
    return items







#def main():

    
   # Get the list of all files and directories
Dir = "D:/2022-2023/"
path =[]

for x in os.listdir(Dir):
     if x.endswith(".xls"):
        full_path=Dir+x
        path.append(full_path)
           #print(x)
BillHeader_Final =[]
BillItems_Final=[]
BillHeader_columns=''
BillItems_columns=''
    #path='D:/2022-2023/Inv.HSS-48-Fortkochi Hotels(Eighth Bastion)-Fort Kochi(47)-Tank, contactor, OLR Relay & SC.xls'
for i in path:
        df =pd.read_excel(i,header=None)
    
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
        #BillItems['File Name']=i
        transportation['File Name']=i
    


        from functools import reduce
        dfs = [invoice, receiever,consignee,transportation]
        BillHeader = reduce(lambda left,right: pd.merge(left,right,on='Invoice No '), dfs) 
        BillHeader_columns=BillHeader.columns.to_list()
        BillHeadData=BillHeader#.to_numpy()
        BillHeader_Final.append(BillHeadData)
        
        BillItems_Final.append(BillItems)
        BillItems_columns=BillItems.columns
    #Final_data=pd.merge(BillItems,BillHeader, on='Invoice No ')
    
#final_data=pd.concat(BillItems_Final,axis=1, ignore_index=False)    
#final_data=pd.DataFrame.from_dict(map(dict,BillItems_Final))


    
#final_inv=pd.concat(BillHeader_Final,axis=1, ignore_index=False)    
#final_inv=pd.DataFrame.from_dict(map(dict,BillHeader_Final))
BillH_Final=np.array(BillHeader_Final,dtype=object)
#final_inv=pd.DataFrame(BillH_Final,columns=BillHeader_columns)
final_inv=pd.DataFrame(np.concatenate(BillH_Final),columns=BillHeader_columns)

FinalItemList=[]
ItemHeaderColumn=''
#FinalItemList[1].shape[1]
for i in BillItems_Final:
    if(i.shape[1]==17) and (i.shape[0]!=0): #change
        FinalItemList.append(i.to_numpy())
        ItemHeaderColumn=i.columns
        
BillI_Final=np.array(FinalItemList,dtype=object)
final_Items=pd.DataFrame(np.concatenate(BillI_Final),columns=ItemHeaderColumn)



final_Items.to_excel('C:/Users/mails/Downloads/ItemList.xlsx',index=False)
final_inv.to_excel('C:/Users/mails/Downloads/InvoiceList.xlsx',index=False)

#if  __name__=="__main__" :
   # main()
final_Items.sum()
