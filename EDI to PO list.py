#!/usr/bin/env python
# coding: utf-8

import os
import re
import xlwt

## function EDItoExcel
def EDItoExcel(Filename): 

    v=open(Filename,"r")

    pos=v.read()+"AMT*GV*3524.36~SE*105*20693~ST*850*20694~~GE*42*850000887~IEA*1*850000887~"
    pos=pos.replace("=","")
    pos=pos.replace("\n","")
    ##print line
    pattern_1=r'BEG\*00\*SA\*[0-9]+?\*[\w\W]*?~ST\*[0-9]+?\*[0-9]+?'
    pattern_test=r'BEG\*00\*SA\*([0-9]+)\*'
    po=re.compile(pattern_1)
    po_list=re.findall(po,pos)

    test=re.compile(pattern_test)
    
    po_list2=re.findall(test,pos) ##for checking

    pattern_2=r'PO1\*([0-9]+?)\*([0-9]+?)\*EA\*(.*?)\*LE\*IN\*([0-9]{9})\*UP.*?\*VN\*(.+?)\*' ##Line#,Quantity,price,item#,Sku#
    pattern_3=r'BEG\*00\*SA\*([0-9]+?)\*'
    pattern_4=r'925485US00\*2127643132\*([0-9]{8})'
    re_line=re.compile(pattern_2)
    re_po=re.compile(pattern_3)
    date=re.findall(re.compile(pattern_4),pos)


    po_line=[]
    for n in po_list:
        ponum=tuple(re.findall(re_po,n))
        line_sku=re.findall(re_line,n)
        for m in line_sku:
            po_line.append((ponum+m))

    length=len(po_line)

    workbook=xlwt.Workbook(encoding='ascii')
    worksheet=workbook.add_sheet('po_line')
    worksheet.write(0,0,label="PO#")
    worksheet.write(0,1,label="Line#")
    worksheet.write(0,2,label="Units")
    worksheet.write(0,3,label="Price")
    worksheet.write(0,4,label="Item#")
    worksheet.write(0,5,label="Sku#")
    worksheet.write(0,8,label="PO#2")

    for n in range(1,length+1):  
        for i in range(6):
            if i==5:
                worksheet.write(n,i,label=po_line[n-1][i])
            elif i==3:
                worksheet.write(n,i,label=float('0'+po_line[n-1][i]))
            else: worksheet.write(n,i,label=int('0'+po_line[n-1][i]))

    for n in range(1,len(po_list2)+1):
        worksheet.write(n,8,label=int(po_list2[n-1])) ## for checking

    filename=str(len(po_list2))+"PO-"+date[0]+".xls"
    workbook.save(filename)
         
    v.close()

    return;

def getTxt(FileList):
    txtList=[]
    length=len(FileList)
    for i in range(0,length):
        if FileList[i].endswith(".txt"):
            txtList.append(FileList[i])

    return txtList;
        
        

FileList=os.listdir(os.getcwd())
txtList=getTxt(FileList)

for i in range(0,len(txtList)):
    EDItoExcel(txtList[i])




