# -*- codeing = utf-8 -*-
# @Author : BBBBBigSand
# @File : pc.py
# @Software : PyCharm
import requests
import time
import re
import pandas as pd
import json
import xlwt

def get_url(num):
    page_url = []
    urlfirst = 'https://club.jd.com/comment/productPageComments.action?callback=&productId=100019125569&score=0&sortType=5&page='
    urllast = '&pageSize=10&isShadowSku=0&fold=1'
    for i in range(1,num,1):
        url = urlfirst+str(i)+urllastA
        page_url.append(url)
    return page_url

def get_content(url_lists):
    headers = {
        'cookie': '__jdu=1657471358907950237233; shshshfpa=5e60e9fc-e80d-eccb-b526-d0821fbd6c21-1664368924; areaId=22; shshshfpb=wHziNQ_UBwNblrVCx5F1a1A; ipLoc-djd=22-1930-50948-57092; jwotest_product=99; pinId=F2UtrazlsSb3guDaJcl14Q; pin=jd_TiRXtMoQncML; unick=jd_TiRXtMoQncML; _tp=oMSExK3tAEevA6w%2Bd0L5fA%3D%3D; _pst=jd_TiRXtMoQncML; PCSYCityID=CN_510000_510100_0; user-key=09c0bc21-1e26-4476-90bf-39c3d22a7221; unpl=JF8EALFnNSttXUNdBxIASEcSH1pSW1UOQ0cDaGZVAQ5ZTFcBEwJOGhN7XlVdXhRKHx9sbxRUXVNOXQ4YCysSEXteU11bD00VB2xXVgQFDQ8WUUtBSUt-SVxRWFULSBMCa2IFZG1bS2QFGjIbFRRNWFJeXwxCHwJpbwFVXlpNVwcZMisVIHttVlpZC0oSM25XBGQfDBdUBhIDHhBdS1pQWFgOSxUHZm8EUlVcSlcHHQEZECBKbVc; __jdv=76161171|baidu-search|t_262767352_baidusearch|cpc|304792250541_0_498383bd3e674979a160adb062591d92|1664423311172; mt_xid=V2_52007VwMVVltdVFodTRBUBGELFlNeWlRZHUspVQdiBEdbXllOUk0cG0AAZwAbTlULBl4DQBBcVzcCFQFeUFFYL0oYXwZ7AhpOXlBDWh9CHFUOZQIiUm1YYlMcSBxYB2UKFFVtXFZZHQ%3D%3D; shshshfp=3bd7d7963280ea6e9ba1e07c190050fb; joyya=1664423313.1664423315.28.1m7k1ch; __jdc=122270672; ip_cityCode=1715; JSESSIONID=A4636DDEDB2940709A61A661E7284766.s1; jsavif=0; __jda=122270672.1657471358907950237233.1657471359.1664693822.1664705193.13; token=d51c7d6f7e9313047e0f846e004cab50,2,924836; __tk=afbd0179f4dca49d6df5bcf0bad8607b,2,924836; __jdb=122270672.2.1657471358907950237233|13.1664705193; shshshsID=b6bdfac42339ddcb80873cae4beead66_2_1664705226683; 3AB9D23F7A4B3C9B=XPUDZWSZAFHBHZIEKBTS4MOVFJLICWXMQQV6ISETNTHYNXT7MSJREE5IE5HC5B544IE6EOCBIL3B2INDOQMP4PNSI4',
        'referer': 'https://item.jd.com/100019125569.html',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36 Edg/105.0.1343.53',
    }

    name = []
    useid = []
    comment = []
    auctionSku = []
    prosize = []
    crtime = []

    for i in range(0, len(url_lists)):
        print('正在爬取第{}页评论'.format(str(i + 1)))
        data = requests.get(url_lists[i], headers=headers).text
        time.sleep(1)
        jdata = json.loads(data)
        for ecomment in jdata['comments']:
            j = 0
            name.append(ecomment['nickname'])
            useid.append(ecomment['id'])
            comment.append(ecomment['content'])
            auctionSku.append(ecomment['productColor'])
            prosize.append(ecomment['productSize'])
            crtime.append(ecomment['creationTime'])

    workbook = xlwt.Workbook(encoding="utf-8", style_compression=0)
    worksheet = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
    col=("序号","用户名","ID","评论内容","评论时间","型号","规格")

    for i in range(0,7):
        worksheet.write(0,i,col[i])

    for i in range(0,601):
        worksheet.write(i + 1, 0, i)
        worksheet.write(i + 1, 1, name[i])
        worksheet.write(i + 1, 2, useid[i])
        worksheet.write(i + 1, 3, comment[i])
        worksheet.write(i + 1, 4, crtime[i])
        worksheet.write(i + 1, 5, auctionSku[i])
        worksheet.write(i + 1, 6, prosize[i])

    workbook.save('联想者.xls')


if __name__ == '__main__':
    num = 65
    url_list = get_url(num)
    get_content(url_list)