# -*- coding:utf-8  -*-
# @Time     : 2022/8/4 22:38
# @Author   : BGLB
# @Software : PyCharm
import json
import os
import time
import traceback
from contextlib import closing
from functools import reduce

import requests
from openpyxl import load_workbook

header = {
    'authority': 'tgc.tmall.com',
    'method': 'POST',
    'scheme': 'https',
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'bx-v': '2.2.2',
    'cache-control': 'no-cache',
    'content-type': 'application/json;charset=UTF-8',
    'cookie': 'cna=xXW2GqIv9RMCAbfvnGIq63vN; xlly_s=1; sgcookie=E100Qj86ID%2Fk0TrRWldMLBRIThghQTZGrW5na7sraTtAUpWD60g6oElStEvTvbC4iQIYjlctaQVdq9mffqVWdfGCud2BzCvMQYNnpd6vBQE2V%2Fs%3D; t=070b3f95a641ba0ca93d5bf37b525788; csg=94bb2a3c; _tb_token_=fef7e593eb6e3; cookie2=1aa13d513e4c871dfb77f7ab5001eef9; SCMLOCALE=zh-cn; _nk_=scm09620215; cookie17=UUpgR1XIK6lQS2vrBQ%3D%3D; SCMSESSID=1aa13d513e4c871dfb77f7ab5001eef9@HAVANA; SCMBIZTYPE=176000; X-XSRF-TOKEN=c39a6889-0197-406b-85c1-23acd3e78714; XSRF-TOKEN=846c7565-c106-46bb-a6e0-51411ad3ba5e; l=eB_izLpnL70Cr9-ABOfwhurza77OMIRfguPzaNbMiOCPOefWREFNW6xiTKLXCnGVnst6R3Wrj_IwBPTEGyznh3v4Gd3hJvSzqdTh.; tfstk=cQDGBNifhfPs-nnr0Aw6nVbtp5kcZC64rxkILhiqGCNUk-kFir5FaRdpxPYW7s1..; isg=BKmpjBp67zOSR9OHa2XjKRzauFUDdp2o7NeytUueWBDZEsgkk8creXBE1LYkijXg',
    'origin': 'https://tgc.tmall.com',
    'pragma': 'no-cache',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 ' \
                  'Safari/537.36',
    'x-xsrf-token': '846c7565-c106-46bb-a6e0-51411ad3ba5e',
}


def get_excel():
    """

    :return:
    """
    url = 'https://tgc.tmall.com/api/v1/orderNew/tradeOrderDownload.htm?spm=a26364.20167850.0.0.59664840JWIIAA&_input_charset=utf-8&taskId=40332923&c2mNavigatorShellPage=2&c2mNavigatorPageOpener=1'
    a = request_download_big_file({'url': url, 'headers': header}, './data.csv')
    print(a)


def request_download_big_file(request_kwargs, local_path, ):
    try:
        if os.path.exists(local_path):
            os.remove(local_path)
        local_path_tmp = local_path+'.tmp'
        param = {'url': '', 'stream': True, 'headers': ''}
        param.update(request_kwargs)
        with closing(requests.get(**param)) as r:
            # r = requests.get(url=file_url, verify=False, stream=True)
            if r.status_code == 200:
                with open(local_path_tmp, "wb") as f:
                    # f.write(r.content)
                    for chunk in r.iter_content(chunk_size=4096):
                        if chunk:
                            f.write(chunk)
            else:
                return False
        os.rename(local_path_tmp, local_path)
    except Exception:
        print(traceback.format_exc())
        return False
    return True


def get_qOsi(mainOrder):
    url = 'https://tgc.tmall.com/ds/api/v1/o/qOsi'
    param = {"mainOrderId": mainOrder, "infoKeys": ["buyerNick", "fullName", "mobilephone", "fullAddress"]}
    rep = requests.post(url, json=param, headers=header)
    result = rep.json()

    if result.get('success'):
        return True, result.get('data')
    return False, result.get('errorMessage')


def get_order():
    url = 'https://tgc.tmall.com/api/v1/orderNew/getTradeOrders.htm'
    order_list = []
    page = 1
    pageSize = 10
    param = {
        'pageNo': page,
        'pageSize': pageSize,
        'sourceTradeId': '',
        'status': 'PAID'
    }
    rep = requests.get(url, params=param, headers=header)
    total = rep.json().get('paginator', {}).get('total')
    if total > pageSize:
        while True:
            rep = requests.get(url, params=param, headers=header)
            if rep.json().get('success'):
                order = rep.json().get('data')
                for item in order:
                    print(item)
                    order_list.extend(item.get('detailOrders'))

            if len(order_list) >= total:
                break
            param['pageNo'] = page+1
    else:
        order = rep.json().get('data')
        for item in order:
            order_list.append(item.get('detailOrders'))
    run_function = lambda x, y: x if y in x else x+[y]
    order_list = reduce(run_function, [[], ]+order_list)

    return order_list


def read_excel(path):
    """

    :return:
    """
    res = []
    try:
        wb = load_workbook(path)
        sheet = wb.active
        n = 1
        for row in sheet.iter_rows(values_only=True, min_row=2):
            n += 1
            mainOrder = row[1]
            if mainOrder:
                flag, data = get_qOsi(mainOrder)
                print(f'{mainOrder}: {data}')
                if flag and data:
                    sheet.cell(row=n, column=3, value=data.get('buyerNick'))
                    sheet.cell(row=n, column=12, value=data.get('fullName'))
                    sheet.cell(row=n, column=13, value=data.get('fullAddress'))
                    sheet.cell(row=n, column=14, value=data.get('mobilephone'))
                # time.sleep(.1)
        wb.save('订单数据.xlsx')
    except Exception:
        print(traceback.format_exc())
        return False, traceback.format_exc()
    return True, res


def save_data(data):
    """

    :return:
    """
    with open('./data.json', encoding='utf8', mode='w') as f:
        json.dump(data, f, ensure_ascii=False)


if __name__ == '__main__':
    a = get_order()
    save_data(a)
    print(len(a))
