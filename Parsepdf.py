# -*- coding: utf-8 -*-
"""
Created on Sun May 13 23:40:19 2018

@author: admin
"""
#df = tabula.read_pdf(r'D:\wjh\2013-08-27-千红制药：2013年8月26日投资者关系活动记录表.PDF', encoding='gbk', pages='all')
#print(df)
#for indexs in df.index:
#    print(df.loc[indexs].values[1].strip())

import pdfplumber
import re

def findty(desc):
    res = ''
    for i in desc.split():
        if i.find("■") !=-1 or i.find("√") !=-1  or i.find("☑") !=-1:
            res = res+" "+i[1:len(i)]
    return res

def findname(names):
    ss = re.split('[：\s；。——、）（]',names)
    companys = []
    persons = []
    curcompany ="空公司"
    for i in ss:
        if len(i)>3:
            companys.append(i)
            curcompany = i
        if 1<len(i)<=3:
            persons.append(curcompany+"%"+i)
    return persons

def parse_pdf(path,file):
    pdf = pdfplumber.open(path+file)
    p0 = pdf.pages[0]
    
    ss = re.split('[：\s]',p0.extract_text())
    ss = [x for x in ss if len(x)>0]
    code = ss[1]
    name = ss[3]
    table = p0.extract_table()
    ty_head = table[0][0]
    pre_ty = table[0][1]
    pss_head = table[1][0]
    pss = table[1][1]
    tme_head = table[2][0]
    tme = table[2][1]
    master = table[4][1]
    content = table[5][1]
    #print(content)
    content_len = str(len(content))
    
    askcontent=''
    for i in pdf.pages:
        askcontent = askcontent+str(i.extract_text())
    asks = max(askcontent.count('？'),askcontent.count('答：'))
    asks = max(asks,askcontent.count('问：'))
    
    ty = findty(pre_ty)
    
    psslist = findname(pss)
    
    pdf.close()
    
    return {"file":file,"code":code,"name":name,"tme_head":tme_head,"tme":tme,
            "pss_head":pss_head,"pss":pss,"ty_head":ty_head,"ty":ty,"pre_ty":pre_ty,'psslist':psslist,'master':master,'content_len':content_len,'asks':asks}
    
if __name__ == '__main__':  
    path = u'D:\\wjh\\2015\\'
    file = u'2015-01-09-华润三九：2015年1月8日投资者关系活动记录表.PDF'
    extractdic = parse_pdf(path,file)
    for i in extractdic:
        print(i,':',extractdic[i],'\n')
    