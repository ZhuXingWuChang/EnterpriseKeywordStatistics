"""
搜索网站 ：百度
搜索公司：卫星石化、洪汇新材、德美化工（三选一）
过滤关键词：污染、破坏、排污、偷排、废
水、超标、泄露、爆炸、死亡、事故、安全、违规、烟尘、溢
油、漏油、溃坝、损失、瓦斯、致癌、毒、黑名单、毁林、违法、
调查、废气、废渣、黑榜、恶、脏、整顿、整改、污水、黑烟、
霉素、臭味、噪音、胡排、乱倒、铬渣、矿渣、锰渣、毒气、
泥浆、血铅、废尘、黑粉、放射性、有害、超标、环境违法、
破坏环境
目的：爬出一家公司在搜狗中所有的媒体报告，统计其数量，然后用过滤关键词进行筛选（不知道这一步可不可以用Python完成）

https://www.sogou.com/web?query=
"""

import requests
import re
import time
import random
import jieba
from lxml import etree
from collections import OrderedDict
from pyexcel_xlsx import get_data
from pyexcel_xlsx import save_data

if __name__ == '__main__':
    row_1_data = [u"污染", u"破坏", u"排污", u"偷排", u"废水", u"超标", u"泄露", u"爆炸", u"死亡", u"事故", u"安全", u"违规"
        , u"烟尘", u"溢油", u"漏油", u"溃坝", u"损失", u"瓦斯", u"致癌", u"毒", u"黑名单", u"毁林", u"违法"
        , u"调查", u"废弃", u"废渣", u"黑榜", u"恶", u"脏", u"整顿", u"整改", u"污水", u"黑烟", u"霉素", u"臭味"
        , u"噪音", u"胡排", u"乱倒", u"铬渣", u"矿渣", u"锰渣", u"毒气", u"泥浆", u"血铅", u"废尘", u"黑粉", u"放射性"
        , u"有害", u"超标", u"环境违法", u"破坏环境"]  # 每一行的数据
    row_1_data.insert(0, u"关键词")
    ent_kw_cnt = []
    for i in row_1_data:
        if i == '关键词':
            ent_kw_cnt.append('卫星石化')
        else:
            ent_kw_cnt.append(0)
    data = OrderedDict()
    row_2_data = ent_kw_cnt
    sheet_1 = [row_1_data, row_2_data]
    data.update({u"Sheet1": sheet_1})
    save_data("./关键词频.xlsx", data)

    sheet = dict(get_data(r"./关键词频.xlsx", encodings="utf-8"))['Sheet1']
    keyword = sheet[0]
    ent_kw_cnt = sheet[1]
    enterprise = sheet[1][0]
    del keyword[0]
    del ent_kw_cnt[0]

    headers = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTM'
                      'L, like Gecko) Chrome/90.0.4430.72 Safari/537.36 Edg/90.0.818.42',
        'Connection': 'close',
    }

    loop_cnt = 0
    for kw in keyword:
        kw_cnt = 0
        for count in range(0, 121, 10):
            params = {
                'wd': enterprise + ' ' + kw,
                'pn': count,
                'oq': enterprise + ' ' + kw,
                'ie': 'utf-8',
                'fenlei': '256',
                'rsv_idx': '1',
            }
            url = 'https://www.baidu.com/s'
            sleep_num1 = random.randint(90, 120)
            time.sleep(sleep_num1)
            try:
                print("正在获取请求......")
                response = requests.get(url=url, headers=headers, params=params)
                print("获取请求成功！")
                response.encoding = 'utf-8'
                page_text = response.text
                r_ex = '<div class="c-abstract">(.*?)</div>'
                content_list = re.findall(r_ex, page_text, re.S)
                print('content_list:', end=' ')
                print(content_list)
                for content in content_list:
                    content_fenci = jieba.lcut(content, cut_all=True)
                    print('content_fenci:', end=' ')
                    print(content_fenci)
                    for fenci in content_fenci:
                        if fenci == kw:
                            kw_cnt = kw_cnt + 1
                    ent_kw_cnt[loop_cnt] = kw_cnt
                    print(kw + "出现频次:", end=' ')
                    print(kw_cnt)
            except requests.exceptions.ConnectionError:
                print("------这一条爬取失败,休息3分钟-----")
                time.sleep(180)
                continue
        loop_cnt = loop_cnt + 1

    if ent_kw_cnt[0] != "卫星石化":
        ent_kw_cnt.insert(0, u"卫星石化")
    row_2_data = ent_kw_cnt
    sheet_1[1] = row_2_data
    data.update({u"Sheet1": sheet_1})
    save_data("./关键词频.xlsx", data)

    print('爬取完毕')
