# https-github.com-new
import requests
import bs4
import re
import json
import xlrd
from queue import Queue
from threading import Thread
from openpyxl import load_workbook
class MyThread(Thread):
    def __init__(self,company_list,cookie_dict,headers,session,put_queue):
        Thread.__init__(self)
        self.company_list=company_list
        self.cookie_dict=cookie_dict
        self.headers=headers
        self.session=session
        self.put_queue=put_queue
    def run(self):
        company(self.company_list,self.cookie_dict,self.headers,self.session,self.put_queue)
class SaveThread(Thread):
    def __init__(self,get_queue):
        Thread.__init__(self)
        self.get_queue=get_queue
    def run(self):
        while True:
            save(self.get_queue)
def save(get_queue):
    data_dict=get_queue.get(block=True,timeout=None)
    wb=load_workbook(u'aa.xlsx')
    ws=wb['2011及之前上市的高新技术企业']
    ws.append([ data_dict['company'],
                data_dict['name'],
                data_dict['inview'],
                data_dict['APO'],
                data_dict['APD'],
                data_dict['PN'],
                data_dict['PD'],
                data_dict['ipc'],])
    wb.save(u'aa.xlsx')
    print("保存成功")
def company(company_list,cookie_dict,headers,session,put_queue):
    for keyword in company_list:
        word=keyword
        print(word)
        literatureSF="复合申请人与发明人=("+word+")"
        executableSearchExp="VDB:(IBI='"+keyword+"')"
        data = {'searchCondition.searchExp': keyword,
                'searchCondition.dbId': 'VDB',
                'resultPagination.limit': '100',
                'searchCondition.searchType': 'Sino_foreign',
                'wee.bizlog.modulelevel': '0200101',
                }
        url = 'http://pss-system.cnipa.gov.cn/sipopublicsearch/patentsearch/executeSmartSearch1207-executeSmartSearch.shtml'
        response=session.post(url,data=data,headers=headers,cookies=cookie_dict).text
        keydata= json.loads(response)
        counter=keydata['searchResultDTO']['pagination']['totalCount']
        pagin=int(counter/12)
        if float(pagin)<counter/12:
            pagin+=1
        print(pagin)
        start=0
        data_list = []
        num=0
        for i in range(0,pagin):
            list={'resultPagination.limit':"12",
            'resultPagination.sumLimit':"10",
            'resultPagination.start':start,
            'resultPagination.totalCount':counter,
            'searchCondition.searchType':'Sino_foreign',
            "searchCondition.originalLanguage":' ',
            "searchCondition.extendInfo['MODE']":'MODE_SMART',
            'searchCondition.extendInfo["STRATEGY"]':'',
            'searchCondition.searchExp':word,
            'searchCondition.executableSearchExp':executableSearchExp,
            'searchCondition.dbId':'',
            'searchCondition.literatureSF':literatureSF,
            'searchCondition.targetLanguage':'',
            'searchCondition.resultMode':'SEARCH_MODE',
            'searchCondition.strategy':'',
            'searchCondition.searchKeywords':word,}
            URLS='http://pss-system.cnipa.gov.cn/sipopublicsearch/patentsearch/showSearchResult-startWa.shtml'
            response=session.post(url,data=data,headers=headers,cookies=cookie_dict).text
            text=session.post(URLS,data=list,headers=headers,cookies=cookie_dict).text
            aa=json.loads(text)

            for i in aa['searchResultDTO']['searchResultRecord']:
                data_dict={}
                data_dict['company']=keyword
                data_dict['name']=i['fieldMap']['TIVIEW']
                data_dict['inview']=i['fieldMap']['INVIEW']
                data_dict['APO']=i['fieldMap']['AP']
                data_dict['APD']=i['fieldMap']['APD']
                data_dict['PN']=i['fieldMap']['PN']
                data_dict['PD']=i['fieldMap']['PD']
                data_dict['ipc']=i['fieldMap']['ICST']
                data_list.append(data_dict.copy())
                print(data_dict)
                put_queue.put(data_dict.copy(),block=True,timeout=None)
                num+=1
                print(num)
            start+=12
def main():
    cookie='IS_LOGIN=true; WEE_SID=Pj33UWwuwiIdQPYKzOf5jgMUtOe98VMFiNtSYdnYP9qEEJxlyVyI!-527884025!-1457108056!1571812371502; JSESSIONID=Pj33UWwuwiIdQPYKzOf5jgMUtOe98VMFiNtSYdnYP9qEEJxlyVyI!-527884025!-1457108056'
    cookie_dict = {i.split("=")[0]:i.split("=")[-1] for i in cookie.split(";")}
    headers={ 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
    'Accept-Encoding': 'gzip, deflate',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'Cache-Control': 'max-age=0',
    'Cookie': 'IS_LOGIN=true; WEE_SID=Pj33UWwuwiIdQPYKzOf5jgMUtOe98VMFiNtSYdnYP9qEEJxlyVyI!-527884025!-1457108056!1571812371502; JSESSIONID=Pj33UWwuwiIdQPYKzOf5jgMUtOe98VMFiNtSYdnYP9qEEJxlyVyI!-527884025!-1457108056',
    'Host': 'pss-system.cnipa.gov.cn',
    'Proxy-Connection': 'keep-alive',
    'Referer': 'http://pss-system.cnipa.gov.cn/sipopublicsearch/patentsearch/tableSearch-showTableSearchIndex.shtml',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36'}
    print(cookie_dict)
    data={'searchCondition.searchExp':'中国石油',
      'searchCondition.dbId':'VDB',
      'resultPagination.limit':'100',
      'searchCondition.searchType':'Sino_foreign',
      'wee.bizlog.modulelevel':'0200101',
      }
    session=requests.session()
    company_name=[]
    xlsx=xlrd.open_workbook(u'更正版-高新技术企业名称.xlsx')
    table=xlsx.sheet_by_name('2011及之前上市的高新技术企业')
    put_queue=Queue()
    get_queue=Queue()
    for i in range(table.nrows):
        if i==0:
            continue
        company_name.append(table.cell(i,1).value)
    num=int(len(company_name)/9)
    a=0
    b=0
    savetd=SaveThread(put_queue)
    savetd.start()
    print(company_name)
    for i in range(9):
        a=b
        b+=num
        thread=MyThread(company_name[a:b],cookie_dict,headers,session,put_queue)
        thread.start()
if __name__ =='__main__':
    main()
