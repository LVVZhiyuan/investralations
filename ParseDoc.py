# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import win32com
from win32com.client import Dispatch, constants
import re

column0 = [u'文件名称',u'股票代码',u'股票名称',u'时间头',u'时间',u'参与单位与人员头',u'参与单位与人员头',u'活动类别头',u'活动类别']   

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

w = win32com.client.Dispatch('Word.Application')

def parse_doc(path,file):
    try:
        doc = w.Documents.Open( FileName = path+file )
        t = doc.Tables[0]
        p1 = doc.Paragraphs[0]
        ss = re.split('[：\s]',str(p1))
        if str(p1).count('证券代码') > 0:        
            ss = [x for x in ss if len(x)>0]
            code = ss[1]
            name = ss[-1]
        else:
            p1 = doc.Paragraphs[1]
            ss = re.split('[：\s]',str(p1))
            ss = [x for x in ss if len(x)>0]
            code = ss[1]
            name = ss[-1]
        ty_head = t.Rows[0].Cells[0].Range.Text
        pre_ty = t.Rows[0].Cells[1].Range.Text
        pss_head = t.Rows[1].Cells[0].Range.Text
        pss = t.Rows[1].Cells[1].Range.Text
        tme_head = t.Rows[2].Cells[0].Range.Text 
        tme = t.Rows[2].Cells[1].Range.Text
        ty = findty(pre_ty)
        psslist = findname(pss)
        
        master = t.Rows[4].Cells[1].Range.Text
        content = t.Rows[5].Cells[1].Range.Text
        content_len = str(len(content))
    
        asks = max(content.count('？'),content.count('答：'))
        asks = max(asks,content.count('问：'))
        
        doc.Close()
        return {"file":file,"code":code,"name":name,"tme_head":tme_head,"tme":tme,
                "pss_head":pss_head,"pss":pss,"ty_head":ty_head,"ty":ty,"pre_ty":pre_ty,'psslist':psslist,'master':master,'content_len':content_len,'asks':asks}
        #return {"file":file,"code":code,"name":name,"tme_head":tme_head,"tme":tme,"pss_head":pss_head,"pss":pss,"ty_head":ty_head,"ty":ty,"pre_ty":pre_ty,'psslist':psslist}
    except Exception as e:
        doc.Close()
        raise Exception('readdoc error')
if __name__ == "__main__":
    path = u'D:\\wjh\\2015\\'
    file = u'2015-01-07-许继电气：2015年1月6日投资者关系活动记录表.DOC'
    extractdic = parse_doc(path,file)
    for i in extractdic:
        print(i,':',extractdic[i],'\n')
    
    