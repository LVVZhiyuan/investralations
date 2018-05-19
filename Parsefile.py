# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import os
import win32com
from win32com.client import Dispatch, constants
from docx import Document
from docx.shared import Inches
import re
import xlwt
import Parsedocx
import ParseDoc
import Parsepdf

def write_excel(df):  
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok=True) 

    rowidx = 0
    for j in range(0,len(df)):
        
        if len(df[j]["psslist"]) == 0:
            sheet1.write(rowidx,0,  df[j]["file"])
            sheet1.write(rowidx,1,  df[j]["code"])
            sheet1.write(rowidx,2,  df[j]["name"])
            sheet1.write(rowidx,3,  df[j]["tme"])
            sheet1.write(rowidx,4,  df[j]["pre_ty"])
            sheet1.write(rowidx,5,  df[j]["ty"])
            sheet1.write(rowidx,6,  df[j]["master"])
            sheet1.write(rowidx,7,  df[j]["content_len"])
            sheet1.write(rowidx,8,  df[j]["asks"])
            sheet1.write(rowidx,9,  df[j]["pss"])
            rowidx = rowidx+1
        else:
            for name in df[j]["psslist"]:
                sheet1.write(rowidx,0,  df[j]["file"])
                sheet1.write(rowidx,1,  df[j]["code"])
                sheet1.write(rowidx,2,  df[j]["name"])
                sheet1.write(rowidx,3,  df[j]["tme"])
                sheet1.write(rowidx,4,  df[j]["pre_ty"])
                sheet1.write(rowidx,5,  df[j]["ty"])
                sheet1.write(rowidx,6,  df[j]["master"])
                sheet1.write(rowidx,7,  df[j]["content_len"])
                sheet1.write(rowidx,8,  df[j]["asks"])
                sheet1.write(rowidx,9,  df[j]["pss"])
                ss = re.split('%',name)
                sheet1.write(rowidx,10,  ss[0])
                sheet1.write(rowidx,11,  ss[1])
                rowidx = rowidx+1

    f.save('E:\\wjh\\error\\2012\\2012.xlsx')
  
PATH = "D:\\wjh\\2012\\"
def parsefile(files,outs,errorfiles): 
    pros = 0
    for doc in files:
        if (os.path.splitext(doc)[1] == '.DOCX' or os.path.splitext(doc)[1] == '.docx') and str(doc).find("投资者关系活动记录表") !=-1:
            try:
                res = Parsedocx.parse_docx(PATH,doc)
                outs.append(res)
            except Exception as e:
                print(e)
                shutil.copyfile(PATH+doc,'E:\\wjh\\error\\2012\\'+doc)
                errorfiles.append(doc)
        elif (os.path.splitext(doc)[1] == '.DOC' or os.path.splitext(doc)[1] == '.doc') and str(doc).find("投资者关系活动记录表") !=-1:
            try:
                res = ParseDoc.parse_doc(PATH,doc)
                outs.append(res)
            except Exception as e:
                print(e)
                shutil.copyfile(PATH+doc,'E:\\wjh\\error\\2012\\'+doc)
                errorfiles.append(doc)
        elif (os.path.splitext(doc)[1] == '.PDF' or os.path.splitext(doc)[1] == '.pdf') and str(doc).find("投资者关系活动记录表") !=-1:
            try:
                res = Parsepdf.parse_pdf(PATH,doc)
                outs.append(res)
            except Exception as e:
                print(e)
                shutil.copyfile(PATH+doc,'E:\\wjh\\error\\2012\\'+doc)
                errorfiles.append(doc)
        pros = pros+1
        print(str(pros)+'/'+str(len(files)))
    
import multiprocessing
import shutil
    
if __name__ == "__main__":
    
    outs1 = []
    errorfiles1 = []
    outs2 = []
    errorfiles2 = []
    outs3 = []
    errorfiles3 = []
    outs4 = []
    errorfiles4 = []
    doc_files = os.listdir(PATH)
    doc_files = doc_files
    file1 = doc_files[0:int(len(doc_files)/4)]
    file2 = doc_files[int(len(doc_files)/4):int(2*len(doc_files)/4)]
    file3 = doc_files[int(2*len(doc_files)/4):int(3*len(doc_files)/4)]
    file4 = doc_files[int(3*len(doc_files)/4):int(len(doc_files))]
    
    #print(file1)
    outs1 = multiprocessing.Manager().list()
    outs2 = multiprocessing.Manager().list()
    outs3 = multiprocessing.Manager().list()
    outs4 = multiprocessing.Manager().list()
    
    '''outs1 = multiprocessing.Queue();
    outs2 = multiprocessing.Queue();
    outs3 = multiprocessing.Queue();
    outs4 = multiprocessing.Queue();'''
    
    p1 = multiprocessing.Process(target=parsefile,args=(file1,outs1,errorfiles1))
    p2 = multiprocessing.Process(target=parsefile,args=(file2,outs2,errorfiles2))
    p3 = multiprocessing.Process(target=parsefile,args=(file3,outs3,errorfiles3))
    p4 = multiprocessing.Process(target=parsefile,args=(file4,outs4,errorfiles4))
    
    p1.start()
    p2.start()
    p3.start()
    p4.start()
    
    p1.join()
    p2.join()
    p3.join()
    p4.join()
    
    #p11 = outs1.get()ile
    
    '''outs = []
    while outs1.qsize() !=0:
        outs.append(outs1.get())
    while outs2.qsize() !=0:
        outs.append(outs2.get())
    while outs3.qsize() !=0:
        outs.append(outs3.get())
    while outs4.qsize() !=0:
        outs.append(outs4.get())'''
    outs1.extend(outs2)
    outs1.extend(outs3)
    outs1.extend(outs4)
        
    write_excel(outs1)