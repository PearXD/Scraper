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

    nowdt=datetime.datetime.now() + datetime.timedelta(days=-1)
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
    scode='522, 3309, 291, 1600, 884, 2616, 1995, 3360, 3606, 1310, 1691, 1888, 823, 960, 1999, 1896, 1336, 1833, 1658, 9909, 1787, 2313, 1516, 1686, 826, 856, 520, 3868, 868, 968, 2020, 1675, 2588, 3998, 384, 2128, 1589, 2319, 1610, 586, 2688, 451, 3900, 1112, 1137, 754, 1970, 817, 9922, 148, 2331, 973, 1268, 425, 316, 1308, 1918, 669, 6110, 168, 3933, 2500, 881, 700, 9988, 1398, 3690, 3968, 939, 2318, 1288, 1299, 857, 3988, 941, 2628, 9618, 1658, 1024, 1211, 1810, 386, 2269, 388, 9999, 2359, 9888, 2020, 1088, 9633, 3328, 2333, 883, 1919, 6030, 16, 6618, 9626, 2202, 1876, 2601, 3908, 11, 2313, 2899, 6690, 2388, 981, 1339, 708, 1928, 998, 960, 66, 27, 2382, 6098, 175, 6969, 6066, 914, 2618, 669, 267, 6160, 1, 1109, 291, 241, 3, 1988, 6818, 2331, 688, 728, 1772, 3692, 2007, 2196, 1766, 3759, 2057, 2, 1113, 6862, 3347, 788, 2319, 9961, 6099, 12, 3606, 1929, 2338, 763, 2688, 2611, 823, 6185, 6886, 6186, 6837, 2328, 1913, 1816, 168, 390, 1336, 1177, 881, 384, 1093, 3333, 1179, 1972, 1997, 968, 2238, 1801, 868, 6806, 1918, 762, 3993, 853, 753, 1038, 1776, 1055, 9688, 285, 1186, 1209, 1193, 9698, 1800, 916, 992, 1898, 1833, 17, 6, 288, 9901, 2600, 9992, 2016, 6881, 316, 268, 1787, 6666, 656, 83, 1378, 772, 4, 1157, 6865, 670, 101, 1066, 322, 6178, 1516, 1877, 1308, 6823, 754, 19, 6110, 358, 1821, 1691, 1999, 3323, 270, 1171, 909, 3958, 1099, 2727, 2018, 813, 6699, 2638, 1618, 1548, 151, 902, 1128, 2039, 6198, 6078, 2883, 3969, 2518, 6993, 2607, 3888, 135, 3380, 3799, 9995, 489, 1044, 6060, 2066, 1359, 2208, 586, 873, 6127, 2128, 136, 177, 1347, 6808, 1888, 874, 3898, 1579, 1951, 3618, 1313, 9668, 3998, 867, 9926, 336, 884, 3319, 1268, 3800, 1585, 1513, 347, 2689, 696, 880, 148, 991, 9666, 966, 2588, 247, 2282, 1528, 836, 1030, 3808, 293, 9922, 23, 780, 2013, 636, 144, 522, 1882, 839, 2866, 2158, 973, 425, 2880, 2799, 3383, 354, 598, 338, 2869'
    # scode='522, 3309'
    codelist=scode.split(',')

    main(codelist)
    logging.info('本次采集完毕 30秒后退出')
    time.sleep(30)

        

    
    
      











    
