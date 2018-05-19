# -*- coding: utf-8 -*-
"""
Created on Fri May 11 22:42:52 2018
@author: Administrator
"""
import urllib.request
import re
import time
from urllib.error import URLError, HTTPError

header = {
   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.96 Safari/537.36'
}
def getfilename(response):  
    res=[]
    patternname = re.compile('''20.*www.cninfo.com.cn" target=.*>''')
    resultname = patternname.findall(response.read().decode('utf-8'))
    
    regex = re.compile('''\d{4}-\d{2}-\d{2}''')
    regexend = re.compile('''title=".*">''')
    regextype = re.compile('\w*\?')
    for substr in resultname:
        reresult = re.findall(regex, substr)
        reresultend = re.findall(regexend, substr)
        reresultype = re.findall(regextype, substr)
        #print(substr)
        item = reresult[0]+'-'+reresultend[0][7:len(reresultend)-3]+'.'+reresultype[0][0:len(reresultype[0])-1]
        item1 = item.replace('/','')
        item2 = item1.replace('\\','')
        item3 = item2.replace('*','')
        res.append(item3)
    return res

def getfilelink(response):
    patternlike = re.compile("http.*www.cninfo.com.cn")
    resultlike = patternlike.findall(response.read().decode('utf-8'))
    return(resultlike)
    
def getfile(url):    
    request = urllib.request.Request(url, headers=header)
    r = urllib.request.urlopen(request)
    links = getfilelink(r)
    time.sleep(2)
    r = urllib.request.urlopen(request)
    names = getfilename(r)
    
    print("len(links)："+str(len(links)))
    print("len(names)："+str(len(names)))
    if len(links) != len(names):
        print("解析网页错误")
    name=0
    for i in links:
        time.sleep(1)
        request = urllib.request.Request(i, headers=header)
        try:
            file = urllib.request.urlopen(request)
        except HTTPError as e:
            print('The (www.python.org)server could not fulfill the request.')
            print('Error code: ', e.code)
            time.sleep(137)
            file = urllib.request.urlopen(request)
        with open('d:/wjh/'+names[name],'wb') as f:
            f.write(file.read())
        f.close()
        name=name+1
for i in range(1,517):
    url =('http://irm.cninfo.com.cn/ircs/interaction/irmInformationList.do?pageNo='
    +str(i))+'&stkcode=&beginDate=2012-01-01&endDate=2012-12-03&keyStr=&irmType='
    getfile(url)

    #r = urllib.request.urlopen(url)
   # names = getfilename(r)
   # for name in names:
    #    print(name)