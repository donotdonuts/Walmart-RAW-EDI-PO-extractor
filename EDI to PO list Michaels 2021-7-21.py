#!/usr/bin/env python
# coding: utf-8
##summary for manual order
##updated 2019-8-23

import os
import re
import xlwt

## function EDItoExcel
def EDItoExcel(Filename): 

    v=open(Filename,"r")

    raw=v.read()
    raw=raw.replace("=","")
    raw=raw.replace("\n","")
    ##print line
    pattern_header=r'BEG\*00\*SA\*([0-9]+?)\*[\w\W]*?DTM\*010\*([0-9]{8})~DTM\*001\*([0-9]{8})'
    ## po number, shipping date and cancel date
    po=re.compile(pattern_header)
    po_header=re.findall(pattern_header,raw)


    pattern_line=r'PO1\*\*([0-9]+?)\*EA\*(.*?)\*\*IN\*([0-9]{6})\*[\w\W]*?VN\*(.*?)~PID' ##quantity, price, item#,vendor number
    re_line=re.compile(pattern_line)
    line_info = re.findall(re_line,raw)

    lines_count = len(line_info)

    workbook=xlwt.Workbook(encoding='ascii')
    worksheet=workbook.add_sheet('PO_info')
    worksheet.write(0,0,label="PO#")
    worksheet.write(0,1,label="shipping date")
    worksheet.write(0,2,label="cancel date")
    worksheet.write(0,3,label="item#")
    worksheet.write(0,4,label="vendor#")
    worksheet.write(0,5,label="price")
    worksheet.write(0,6,label="quantity")

    for row in range(1,lines_count+1):
        worksheet.write(row,0,label=po_header[0][0])
        worksheet.write(row,1,label=po_header[0][1])
        worksheet.write(row,2,label=po_header[0][2])
        worksheet.write(row,3,label=line_info[row-1][2])
        worksheet.write(row,4,label=line_info[row-1][3])
        worksheet.write(row,5,label=line_info[row-1][1])
        worksheet.write(row,6,label=line_info[row-1][0])


    filename=po_header[0][0]+".xls"
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




