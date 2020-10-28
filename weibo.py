#!/usr/bin/env python
# -*- coding: UTF-8 -*-

import requests
import json
import re
import html
from lxml import etree
import sys
from pyquery import PyQuery as pq
import xlsxwriter

def get_leaderboard(category):
    url = "https://weibo.com/a/aj/transform/loadingmoreunlogin?ajwvr=6"
    count = 0
    page = 0
    total_nick_name_list = []
    total_user_url_list = []
    while count < 100:
        page = page + 1
        params = {
            'category': category,
            'page': page
        }
        res = requests.get(url, params=params)
        res_json = html.unescape(res.text)
        res_dict = json.loads(res_json, strict=False)
        res_html = etree.HTML(res_dict["data"])
        # list nickname
        nick_name = res_html.xpath("//div[@class=\"subinfo_box clearfix\"]/a[2]/span/text()")
        # list user url
        user_url = res_html.xpath("//div[@class=\"subinfo_box clearfix\"]/a[1]/@href")

        new_url_list = []
        # url处理，统一去除//weibo.com
        for item in user_url:
            item = item.replace('//weibo.com', '')
            item = r"https://weibo.com" + item
            new_url_list.append(item)

        if len(total_nick_name_list) + len(nick_name) < 100:
            total_nick_name_list.extend(nick_name)
            total_user_url_list.extend(new_url_list)
            count += len(nick_name)
        else:
            for index, item in enumerate(nick_name):
                if count < 100:
                    total_nick_name_list.append(item)
                    total_user_url_list.append(new_url_list[index])
                    count += 1
                else:
                    break
    return total_nick_name_list, total_user_url_list

#获取uid
def get_user_uid(total_nick_name_list, total_user_url_list):
    user_uid_list = []
    for index, item in enumerate(total_user_url_list):
        pattern = re.compile(r"(?<=u/)\d+")
        match = pattern.findall(item)
        if match:
            user_uid_list.append(match[0])
        else:
            nick_name = total_nick_name_list[index]
            uid = find_uid(nick_name)
            if uid:
                user_uid_list.append(uid)
            else:
                print("ERROR!")
    return user_uid_list


def find_uid(nick_name):
    search_url = "https://s.weibo.com/weibo/"
    headers = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.111 Safari/537.36',
        'cookie': 'SINAGLOBAL=7633674053407.922.1602316719977; SUB=_2AkMoxIqBf8NxqwJRmPkczGLmaIt2wwrEieKemHtaJRMxHRl-yT9jqmcjtRB6A0SkbrIim2oOjW3dyOhg-CZUVfUrJFnI; SUBP=0033WrSXqPxfM72-Ws9jqgMF55529P9D9WhOUcszPzoeeHiF2M4.QLmD; login_sid_t=89f0a3f910e9e3dfe5f95fc8dfd039d3; cross_origin_proto=SSL; wb_view_log=2560*14401; _s_tentry=weibo.com; Apache=8683473062600.699.1603798480325; ULV=1603798480328:3:3:1:8683473062600.699.1603798480325:1602766491099; UOR=,,github.com',
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9'
    }
    try:
        find_url = search_url + nick_name
        res = requests.get(url=find_url, headers=headers)
        res_html = etree.HTML(res.text)
        uid = res_html.xpath("//div[@class=\"card card-user-b s-pg16 s-brt1\"]/div[@class=\"info\"]/div/a[3]/@uid")
        return uid[0]
    except:
        return None


def get_json(params, headers):
    """获取网页中json数据"""
    url = 'https://m.weibo.cn/api/container/getIndex?'
    try:
        res = requests.get(url, params=params, headers=headers)
        if res.status_code == 200:
            return res.json()
    except requests.ConnectionError as e:
        return('Error', e.args)

def standardize_info(weibo):
    """标准化信息，去除乱码"""
    for k, v in weibo.items():
        if 'bool' not in str(type(v)) and 'int' not in str(
                type(v)) and 'list' not in str(
                    type(v)) and 'long' not in str(type(v)):
            weibo[k] = v.replace(u'\u200b', '').encode(
                sys.stdout.encoding, 'ignore').decode(sys.stdout.encoding)
    return weibo


def get_user_info(uid):
    """获取用户信息"""
    params = {'containerid': '100505' + uid}
    headers = {
        'host': 'm.weibo.cn',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.111 Safari/537.36',
        'cookie': 'SINAGLOBAL=7633674053407.922.1602316719977; SUB=_2AkMoxIqBf8NxqwJRmPkczGLmaIt2wwrEieKemHtaJRMxHRl-yT9jqmcjtRB6A0SkbrIim2oOjW3dyOhg-CZUVfUrJFnI; SUBP=0033WrSXqPxfM72-Ws9jqgMF55529P9D9WhOUcszPzoeeHiF2M4.QLmD; login_sid_t=89f0a3f910e9e3dfe5f95fc8dfd039d3; cross_origin_proto=SSL; wb_view_log=2560*14401; _s_tentry=weibo.com; Apache=8683473062600.699.1603798480325; ULV=1603798480328:3:3:1:8683473062600.699.1603798480325:1602766491099; UOR=,,github.com',
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'referer': 'https://m.weibo.com/u/' + uid
    }
    js = get_json(params, headers)
    if js['ok']:
        info = js['data']['userInfo']
        user_info = {}
        user_info['id'] = uid
        user_info['screen_name'] = info.get('screen_name', '')
        user_info['gender'] = info.get('gender', '')
        user_info['statuses_count'] = info.get('statuses_count', 0)
        user_info['followers_count'] = info.get('followers_count', 0)
        user_info['follow_count'] = info.get('follow_count', 0)
        user_info['description'] = info.get('description', '')
        user = standardize_info(user_info)
        return user


def get_page(uid, page_num):
    params = {
        'type': 'uid',
        'value': uid,
        'containerid': '107603' + uid,
        'page': page_num
    }
    headers = {
        'host': 'm.weibo.cn',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.111 Safari/537.36',
        'cookie': 'SINAGLOBAL=7633674053407.922.1602316719977; SUB=_2AkMoxIqBf8NxqwJRmPkczGLmaIt2wwrEieKemHtaJRMxHRl-yT9jqmcjtRB6A0SkbrIim2oOjW3dyOhg-CZUVfUrJFnI; SUBP=0033WrSXqPxfM72-Ws9jqgMF55529P9D9WhOUcszPzoeeHiF2M4.QLmD; login_sid_t=89f0a3f910e9e3dfe5f95fc8dfd039d3; cross_origin_proto=SSL; wb_view_log=2560*14401; _s_tentry=weibo.com; Apache=8683473062600.699.1603798480325; ULV=1603798480328:3:3:1:8683473062600.699.1603798480325:1602766491099; UOR=,,github.com',
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'referer': 'https://m.weibo.com/u/' + uid
    }
    return get_json(params, headers)

def parse_page(json):
    if json:
        items = json.get('data').get('cards')
        weibo_list = []
        for item in items:
            if item.__contains__("mblog"):
                item = item.get('mblog')
                weibo = {}
                weibo['id'] = item.get('id')
                weibo['text'] = pq(item.get('text')).text()
                weibo['attitudes'] = item.get('attitudes_count')
                weibo['comments'] = item.get('comments_count')
                weibo['reposts'] = item.get('reposts_count')
                weibo_list.append(weibo)
            else:
                continue
        return weibo_list
    else:
        return None

def get_comment(weibo_id):
    url = 'https://m.weibo.cn/comments/hotflow?'
    headers = {
        'host': 'm.weibo.cn',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.111 Safari/537.36',
        'cookie': 'SINAGLOBAL=7633674053407.922.1602316719977; SUB=_2AkMoxIqBf8NxqwJRmPkczGLmaIt2wwrEieKemHtaJRMxHRl-yT9jqmcjtRB6A0SkbrIim2oOjW3dyOhg-CZUVfUrJFnI; SUBP=0033WrSXqPxfM72-Ws9jqgMF55529P9D9WhOUcszPzoeeHiF2M4.QLmD; login_sid_t=89f0a3f910e9e3dfe5f95fc8dfd039d3; cross_origin_proto=SSL; wb_view_log=2560*14401; _s_tentry=weibo.com; Apache=8683473062600.699.1603798480325; ULV=1603798480328:3:3:1:8683473062600.699.1603798480325:1602766491099; UOR=,,github.com',
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    }
    params = {'id': weibo_id,
              'mid': weibo_id,
              'max_id_type': 0}
    try:
        res = requests.get(url, params=params, headers=headers)
        if res.status_code == 200:
            content = res.json()
            if content.__contains__("data"):
                if content['data'].__contains__("data"):
                    comment_str = ""
                    for index, item in enumerate(content['data']['data']):
                        text = item['text']
                        comment = re.sub('<[^<]+?>', '', text).replace('\n', '').strip()
                        comment_str += str(index) + comment + '\n'
                    return comment_str
            else:
                print("该微博暂时无评论！")
                return ""
    except requests.ConnectionError as e:
        return ""

# 生成excel文件
def generate_user_info_excel(user_info_list, file_code):
    workbook = xlsxwriter.Workbook('user_info_'+file_code+'.xlsx')
    worksheet = workbook.add_worksheet()

    # 用符号标记位置，例如：A列1行
    worksheet.write('A1', '用户id')
    worksheet.write('B1', '昵称')
    worksheet.write('C1', '性别')
    worksheet.write('D1', '微博数')
    worksheet.write('E1', '粉丝数')
    worksheet.write('F1', '关注数')
    worksheet.write('G1', '微博描述')
    row = 1
    col = 0
    for item in user_info_list:
        # 使用write_string方法，指定数据格式写入数据
        worksheet.write_string(row, col, str(item['id']))
        worksheet.write_string(row, col + 1, str(item['screen_name']))
        worksheet.write_string(row, col + 2, str(item['gender']))
        worksheet.write_string(row, col + 3, str(item['statuses_count']))
        worksheet.write_string(row, col + 4, str(item['followers_count']))
        worksheet.write_string(row, col + 5, str(item['follow_count']))
        worksheet.write_string(row, col + 6, str(item['description']))
        row += 1
    workbook.close()

def generate_user_weibo_excel(user_weibo_list, uid):
    workbook = xlsxwriter.Workbook('user_weibo_'+uid+'.xlsx')
    worksheet = workbook.add_worksheet()

    # 用符号标记位置，例如：A列1行
    worksheet.write('A1', '用户id')
    worksheet.write('B1', '微博id')
    worksheet.write('C1', '微博内容')
    worksheet.write('D1', '点赞数')
    worksheet.write('E1', '评论数')
    worksheet.write('F1', '转发数')
    worksheet.write('G1', '评论汇总')
    row = 1
    col = 0
    for item in user_weibo_list:
        # 使用write_string方法，指定数据格式写入数据
        worksheet.write_string(row, col, str(item['uid']))
        worksheet.write_string(row, col + 1, str(item['id']))
        worksheet.write_string(row, col + 2, str(item['text']))
        worksheet.write_string(row, col + 3, str(item['attitudes']))
        worksheet.write_string(row, col + 4, str(item['comments']))
        worksheet.write_string(row, col + 5, str(item['reposts']))
        worksheet.write_string(row, col + 6, str(item['comments_word']))
        row += 1
    workbook.close()

#获取榜单前100数据
def get_top_data(category, total_user_uid):
    user_info_list = []
    for uid in total_user_uid:
        user = get_user_info(uid)
        user_info_list.append(user)
        print("log user:"+str(user))
    generate_user_info_excel(user_info_list, category)

def get_weibo_and_comment(uid):
    page_num = 0
    sum = 0
    weibo_list = []
    while True:
        page = get_page(uid, page_num)
        #print(page)
        results = parse_page(page)
        #print(len(results))
        if len(results) > 0 and sum < 20:
            page_num += 1
            for result in results:
                result["uid"] = uid
                result["comments_word"] = ""
                if result.__contains__("id"):
                    result["comments_word"] = get_comment(result['id'])
                weibo_list.append(result)
                sum += 1
                if sum == 20:
                    break
        else:
            break
    if len(weibo_list) > 0:
        generate_user_weibo_excel(weibo_list, uid)

if __name__ == "__main__":

    # 99991 日榜 99992周榜 9999月榜
    directory = ["99991", "99992", "99993"]
    for category in directory:
        total_nick_name_list, total_user_url_list = get_leaderboard(category=category)
        total_user_uid = get_user_uid(total_nick_name_list, total_user_url_list)
        get_top_data(category, total_user_uid)
        for uid in total_user_uid:
            get_weibo_and_comment(uid)
