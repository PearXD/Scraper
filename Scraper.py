# -*- coding: utf-8 -*-
import requests
requests.urllib3.disable_warnings()

import time
import re
import os

#JSON解析
import json

#xpath解析
from lxml import etree
from lxml import html
#EXCEL
import openpyxl
import logging
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(message)s",
                    datefmt = '%m-%d %H:%M:%S'
                    )


import datetime


s参与者编号='参与者编号'
s中央结算系统参与者名称='中央结算系统参与者名称'
s持股量='持股量'
s占已发行股份='占已发行股份'

class 披露易(object):
    def __init__(self,*arg):
        self.cnheaders={
            "host":"sc.hkexnews.hk",
            "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
            "accept-encoding": "gzip, deflate",
            "accept-language": "zh-CN,zh;q=0.9,en;q=0.8,zh-TW;q=0.7,en-US;q=0.6,zh-HK;q=0.5",
            "cache-control": "max-age=0",
            "content-length": "321",
            "content-type": "application/x-www-form-urlencoded",
            "origin": "https://sc.hkexnews.hk",
            "sec-ch-ua-mobile": "?0",
            "sec-fetch-dest": "document",
            "sec-fetch-mode": "navigate",
            "sec-fetch-site": "same-origin",
            "sec-fetch-user": "?1",
            "upgrade-insecure-requests": "1",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.93 Safari/537.36"
        }
        
        self.enheaders={
            "host":"www.hkexnews.hk",
            "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
            "accept-encoding": "gzip, deflate",
            "accept-language": "zh-CN,zh;q=0.9,en;q=0.8,zh-TW;q=0.7,en-US;q=0.6,zh-HK;q=0.5",
            "cache-control": "max-age=0",
            "content-length": "321",
            "content-type": "application/x-www-form-urlencoded",
            "origin": "https://ww.hkexnews.hk",
            "sec-ch-ua-mobile": "?0",
            "sec-fetch-dest": "document",
            "sec-fetch-mode": "navigate",
            "sec-fetch-site": "same-origin",
            "sec-fetch-user": "?1",
            "upgrade-insecure-requests": "1",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.93 Safari/537.36"
        }

        # 代理IP
        self.proxies={'https':'127.0.0.1:8888'}
        self.proxies={}
        
        self.session =requests.session()
        self.session.verify=False


    def search(self,code,dt):
        if not code or not dt :return 
        code=str(code).strip()
        while len(code)<5:
            code='0'+code
            pass

        sdt=dt.strftime('%Y/%m/%d')
        url='https://sc.hkexnews.hk/TuniS/www.hkexnews.hk/sdw/search/searchsdw_c.aspx'
        #url='https://www.hkexnews.hk/sdw/search/searchsdw.aspx'
        data={
                "__EVENTTARGET": "btnSearch",
                "__EVENTARGUMENT": "",
                "__VIEWSTATE": "/wEPDwULLTIwNTMyMzMwMThkZHNjXATvSlyVIlPSDhuziMEZMG94",
                "__VIEWSTATEGENERATOR": "",
                "today": datetime.datetime.now().strftime('%Y%m%d'),
                "sortBy": "shareholding",
                "sortDirection": "desc",
                "alertMsg": "",
                "txtShareholdingDate": sdt,
                "txtStockCode": code,
                "txtStockName": "",
                "txtParticipantID": "",
                "txtParticipantName": "",
                "txtSelPartID": ""
            }

        shtml=''

        ec=1
        while ec<11:
            status=None
            if 1:
                rs=self.session.post(url,headers=self.cnheaders,timeout=30,proxies=self.proxies,data=data)
                if rs.status_code==200:
                    shtml=rs.text
                    break
                else:
                    status=rs.status_code
            # except:pass
            logging.warning(f'获取[搜索]源码失败 重试{ec}/10 状态{status}')
            ec+=1
            time.sleep(3)

        if shtml:
            itemlist=[]
            wc=etree.HTML(shtml)
            trlist=wc.xpath('//tbody/tr')
            for tr in trlist:
                line={}
                line['code']=code
                line['date']=sdt

                tdlist=tr.xpath('td')
                for td in tdlist:
                    k=td.xpath('string(div[1])').strip()
                    v=td.xpath('string(div[2])').strip()
                    if not k or not v:continue

                    if k:k=k.replace(':','').split('(')[0].split('（')[0].split('/')[0].strip()
                    if k==s持股量 or k==s占已发行股份:
                        v=float(v.replace(',','').replace('%',''))
                        pass

                    line[k]=v
                
                itemlist.append(line)
                #print(line)
                #return 

                pass
            
            # if itemlist:print(itemlist[-1])
            count=len(itemlist)
            logging.info(f'日期{sdt} {code} 搜索到 {count} 条记录')
            return itemlist


        logging.info(f'日期{sdt} {code} 没有搜索结果')


myapp=披露易()
def calcu(list1,list2):
    ''' 计算差值 '''
    rlist=[]
    for l in list1+list2:
        line={}
        for k in l:
            if type(l[k])!=type(''):continue
            line[k]=l[k]
            pass
        if not line in rlist:rlist.append(line)

    for l in rlist:
        f1={}
        for l1 in list1:
            if l.get(s参与者编号,'')+l.get(s中央结算系统参与者名称,'') == l1.get(s参与者编号,'')+l1.get(s中央结算系统参与者名称,''):
                f1=l1
                break

        f2={}
        for l1 in list2:
            if l.get(s参与者编号,'')+l.get(s中央结算系统参与者名称,'') == l1.get(s参与者编号,'')+l1.get(s中央结算系统参与者名称,''):
                f2=l1
                break

        l['list']=[f1,f2]
        l[s持股量]=f1.get(s持股量,0)-f2.get(s持股量,0)
        l[s占已发行股份]=f1.get(s占已发行股份,0)-f2.get(s占已发行股份,0)
        pass
    return rlist


def get_twodays_data(code):
    ''' 获取相邻两天数据差 '''

    nowdt=datetime.datetime.now()
    itemlist1=[];itemlist2=[]
    for i in range(0,10):
        itemlist0=myapp.search(code,nowdt)
        if itemlist0:
            if not itemlist1:itemlist1=itemlist0
            else:itemlist2=itemlist0

        nowdt=nowdt+datetime.timedelta(days=-1)
        if itemlist1 and itemlist2:break
   
    clist=calcu(itemlist1,itemlist2)
    return clist




def main(codelist=[]):
    ''' 执行获取 '''
    wb=openpyxl.Workbook()
    sheet=wb.worksheets[0]

    sheet.column_dimensions['C'].width=40   # C列列宽
    sheet.column_dimensions['F'].width=5   # C列列宽
    sheet.column_dimensions['J'].width=5   # C列列宽
    # 表头
    sheet.append(['股票代码','参与者编号','中央结算系统参与者名称','持股量（变化）','占已发行股份（变化）',None,'持股量','占已发行股份','日期',None,'上一日持股量','上一日占已发行股份','日期'])
    
    # 线程操作
    from concurrent.futures import ThreadPoolExecutor
    pool = ThreadPoolExecutor(max_workers=10)    

    ret=[]
    for code in codelist:    
        # 多线程提交 非阻塞
        ret.append( pool.submit(get_twodays_data,code))
        time.sleep(1)

    # ret=[pool.submit(get_twodays_data,code) for code in codelist]
    result=[task.result() for task in ret]     # 取结果 阻塞
    
    # 输出
    if result:
        for clist in result:
            if not clist :continue
            for item in clist:
                baseline=[item.get('code'),item.get(s参与者编号),item.get(s中央结算系统参与者名称),item.get(s持股量),item.get(s占已发行股份)]
                line1=[None,item['list'][0].get(s持股量),item['list'][0].get(s占已发行股份),item['list'][0].get('date')]
                line2=[None,item['list'][1].get(s持股量),item['list'][1].get(s占已发行股份),item['list'][1].get('date')]

                sheet.append(baseline+line1+line2)

    # 保存
    sdt=datetime.datetime.now().strftime('%Y%m%d %H%M%S')
    fpath=f'截止{sdt}数据.xlsx'
    wb.save(fpath)
    logging.info(f'文件已保存到 {fpath}')
    print('')


if __name__ == '__main__':   
    # 自己编辑这个code集合
    scode='522, 3309, 291, 1600, 884, 2616, 1995, 3360, 3606, 1310, 1691, 1888, 823, 960, 1999, 1896, 1336, 1833, 1658, 9909, 1787, 2313, 1516, 1686, 826, 856, 520, 3868, 868, 968, 2020, 1675, 2588, 3998, 384, 2128, 1589, 2319, 1610, 586, 2688, 451, 3900, 1112, 1137, 754, 1970, 817, 9922, 148, 2331, 973, 1268, 425, 316, 1308, 1918, 669, 6110, 168, 3933, 2500, 881'
    # scode='522, 3309'
    codelist=scode.split(',')

    main(codelist)
    logging.info('本次采集完毕 30秒后退出')
    time.sleep(30)

        

    
    
      











    
