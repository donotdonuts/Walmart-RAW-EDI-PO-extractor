#!/usr/bin/env python
# coding: utf-8
##used for manual keyin SSO DSDC type
##updated 2019-8-

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
   

    pattern_2=r'PO1\*([0-9]+?)\*([0-9]+?)\*EA\*(.*?)\*LE\*IN\*([0-9]{9})\*UP.*?\*VN\*(.+?)\*' ##[7]Line#,[8]Quantity,[9]price,[10]item#,[11]Sku#
    pattern_3=r'BEG\*00\*SA\*([0-9]+?)\*\*([0-9]{8})[\w\W]*?REF\*PD\*([\w\W]*?)~REF\*IA\*([0-9]{9})[\w\W]*?DTM\*038\*([0-9]{8})~DTM\*037\*([0-9]{8})~DTM\*063\*([0-9]{8})'
    ## [0]PO number, [1]PO date [2]event code [3]vendor sub number [4]Not later than, [5]No earlier than, [6]MABD
    pattern_4=r'925485US00\*2127643132\*([0-9]{8})'

    pattern_st_lines=r'PO1\*.*?~AMT'
    pattern_sts=r'~N1\*SN\*\*UL\*([0-9]{13})'#[12] store gln
    
    
    re_line=re.compile(pattern_2)
    re_po=re.compile(pattern_3)
    date=re.findall(re.compile(pattern_4),pos)
    re_st_lines=re.compile(pattern_st_lines)
    re_sts=re.compile(pattern_sts)


    po_line=[]
    for singlePo in po_list:
        ponum=re.findall(re_po,singlePo)
        line_sku=re.findall(re_line,singlePo)
        st_lines=re.findall(re_st_lines,singlePo)
        Store_GLN=re.findall(re_sts,singlePo)
        for i in range(len(line_sku)):
            poInfo=ponum[0]
            lineInfo=line_sku[i]
            temp=poInfo+lineInfo+tuple(Store_GLN)
            po_line.append(temp)

                

    global index
    index=1
    num_Po=len(po_list2)
    write_to_Excel(po_line,num_Po,Filename)
                    
    v.close()

    return;   
         


def getTxt(FileList):
    txtList=[]
    length=len(FileList)
    for i in range(0,length):
        if FileList[i].endswith(".txt"):
            txtList.append(FileList[i])

    return txtList;

def write_to_Excel(po_line,num_Po,Filename):
    length=len(po_line)
    print length
    if length>=65535:
        write_to_Excel(po_line[65535:],num_Po,Filename)
        length=length-len(po_line[65535:])
        global index
        index=index+1
        
        print index
        
        

    workbook=xlwt.Workbook(encoding='ascii')
    worksheet=workbook.add_sheet('po_line')
    worksheet.write(0,0,label="PO#")
    worksheet.write(0,1,label="PO date")
    worksheet.write(0,2,label="Event Code")
    worksheet.write(0,3,label="9digit#")
    worksheet.write(0,4,label="No Later Than")
    worksheet.write(0,5,label="No Earlier than")
    worksheet.write(0,6,label="MABD")
    worksheet.write(0,7,label="Line#")
    worksheet.write(0,8,label="Units")
    worksheet.write(0,9,label="Price")
    worksheet.write(0,10,label="Item#")
    worksheet.write(0,11,label="Sku#")
    worksheet.write(0,12,label="store gln")
    print "execute"
    


    for n in range(1,length+1):  
        for i in range(13):
            if i==11 or i==2:
                worksheet.write(n,i,label=po_line[n-1][i])
            elif i==9:
                worksheet.write(n,i,label=float('0'+po_line[n-1][i]))
            else: worksheet.write(n,i,label=int('0'+po_line[n-1][i]))

    filename=str(Filename+" Part "+str(index)+".xls")
    workbook.save(filename)
    return;

        

FileList=os.listdir(os.getcwd())
txtList=getTxt(FileList)

for i in range(0,len(txtList)):
    print "turning"+txtList[i]
    EDItoExcel(txtList[i])




