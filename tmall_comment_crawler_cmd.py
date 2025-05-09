#!/usr/bin/env python
# -*- coding: utf-8 -*-

import hashlib
import json
import random
import re
import time

import pandas as pd
import requests


class TmallCommentCrawler:
    def __init__(self):
        self.headers = {
            'Accept-Encoding': 'gzip, deflate, br',
            'Cache-Control': 'no-cache',
            'Connection': 'keep-alive',
            'Cookie': 'lid=tb358500429; wk_cookie2=1f0b7a77b5db43fd137d6f01281dd37d; wk_unb=VyyQ6i7zQ9P42g==; xlly_s=1; cna=nKs0H6yfViMCAXWbBi0ADJo7; dnk=tb358500429; uc1=existShop=false&cookie16=WqG3DMC9UpAPBHGz5QBErFxlCA==&cookie15=W5iHLLyFOGW7aA==&cookie21=URm48syIYn73&pas=0&cookie14=UoYaje3TqSHagg==; uc3=vt3=F8dD2EXY8uO57DhnvFo=&lg2=URm48syIIVrSKA==&id2=VyyQ6i7zQ9P42g==&nk2=F5RGM4iqC+NCac0=; tracknick=tb358500429; _l_g_=Ug==; uc4=nk4=0@FY4NBmpKNZ/A4kp8c1bvfSQxmqXuTQ==&id4=0@VXtfFdPVcAuwvjXKHFRCdw5uTaKB; unb=4019587167; lgc=tb358500429; cookie1=Vy67CEguTnSJl/VzmGz5paJ62i9fFGT9HvjV9icKaQ8=; login=true; cookie17=VyyQ6i7zQ9P42g==; cookie2=2948f234937fd2688799a08198c2f966; _nk_=tb358500429; sgcookie=E100+OKU4J4+vHhBNLiuG+BSCvX2DTBr0ADJ3b4XKMWrVROm2AJeJ2vxwz/2WJ51rbvqwdSDctiTOUlGbj1Zu7V1DXMYZIsUZx6Ek+paiYqcQa0=; cancelledSubSites=empty; sg=972; t=ffa7190fdbb2626d89558148ebee6b74; csg=3c2a9fec; sn=; _tb_token_=581303797776e; mtop_partitioned_detect=1; _m_h5_tk=4967d33ce878be64ff2bf85271fdfd05_1746456131231; _m_h5_tk_enc=b6d3ca6fc035c62a0ee5be79c0018685; isg=BK6u9yeUSoEFaLD6n_GC3FZD_wRwr3Kprmv70dh3A7FNu04VQDwCuVS5dydXZmrB; bxuab=0; tfstk=gUjqhK2sKoEVN8OvogtZ8WJfJntvuhPQ3GO6IOXMhIA06-pMbs1kSc6bkRywa1dXlnbXIc5WN1i1hG7-b9BOc-K_HnBvXhVQOh-NHtKO1elogirkEdWknV00P3qGC5cuOkZCnrvvfTVCfc-ndQJMjKYijT2yIplDjFvcE8JMCj0im12zUdvZnqxMjuXkdp-Ms1xgU3vvZKxMo1XohsCA6X9HoRgconTx7b8XttArjDkp3Ec11Viq049BlBYVDcY54K82KtjvVma2EgsyRUM_J3XdPsvyYugDxT7GTwfTR4KVIZf9zsF-zQ_lDOAMPAmCU9Wc-3j_IvLPZd8HjUkijT-vZFsDxkcyeaXfSiCqQcvAcMTw9UyiXFtlAebhgAedEnvG6eI_9mdcI9IdRHrSoHXPS67N4Th9ETJdXwMr7EvJUBNzaukxY1nNJGDEWVLk-LRQHF0tWEvJUBwzUV39rLvyO-LG.',
            'Host': 'h5api.m.tmall.com',
            'accept': '*/*',
            'accept-language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
            'referer': 'https://detail.tmall.com/',
            'sec-ch-ua': '"Chromium";v="136", "Microsoft Edge";v="136", "Not.A/Brand";v="99"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'script',
            'sec-fetch-mode': 'no-cors',
            'sec-fetch-site': 'same-site',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0'
        }
        self.base_url = 'https://h5api.m.tmall.com/h5/mtop.taobao.rate.detaillist.get/6.0/'
        self.last_error = ""  # 存储最后一次错误信息
        
        # 从Cookie中提取token进行签名计算
        self._extract_token_from_cookie()
        
    def _extract_token_from_cookie(self):
        """从Cookie中提取token用于签名计算"""
        cookie_str = self.headers['Cookie']
        self.token = ''
        
        # 尝试提取_m_h5_tk
        m = re.search(r'_m_h5_tk=([^_]+)_', cookie_str)
        if m:
            self.token = m.group(1)
            print(f"已从Cookie中提取token: {self.token}")
        else:
            print("警告: 无法从Cookie中提取token，签名可能无效")
        
    def get_comments(self, item_id, page_num=5, order_type=""):
        """
        获取商品评论
        :param item_id: 商品ID
        :param page_num: 页数，每页20条评论，默认抓取5页共100条
        :param order_type: 排序方式，为空表示默认排序，"feedbackdate"表示按时间排序
        :return: 评论数据列表
        """
        all_comments = []
        
        for page in range(1, page_num + 1):
            print(f"正在爬取第 {page} 页评论...")
            
            # 构建API请求参数
            timestamp = int(time.time() * 1000)
            
            data = {
                "showTrueCount": False,
                "auctionNumId": str(item_id),
                "pageNo": page,
                "pageSize": 20,
                "rateType": "",
                "searchImpr": "-8",
                "orderType": order_type,
                "expression": "",
                "rateSrc": "pc_rate_list"
            }
            
            # 使用正确的方式生成签名
            data_str = json.dumps(data)
            sign = self._generate_sign(timestamp, data_str)
            
            # 设置最后一次错误为空
            self.last_error = ""
            
            params = {
                'jsv': '2.7.4',
                'appKey': '12574478',
                't': timestamp,
                'sign': sign,
                'api': 'mtop.taobao.rate.detaillist.get',
                'v': '6.0',
                'isSec': 0,
                'ecode': 1,
                'timeout': 20000,
                'type': 'jsonp',
                'dataType': 'jsonp',
                'jsonpIncPrefix': 'pcdetail',
                'callback': f'mtopjsonppcdetail{random.randint(10, 99)}',
                'data': data_str
            }
            
            try:
                response = requests.get(self.base_url, params=params, headers=self.headers)
                response.raise_for_status()
                
                # 解析JSONP响应
                try:
                    json_str = re.search(r'mtopjsonppcdetail\d+\((.*)\)', response.text).group(1)
                    result = json.loads(json_str)
                    
                    # 检查API调用是否成功
                    if "SUCCESS" in result.get('ret', [''])[0]:
                        # 提取评论数据
                        if 'data' in result and 'rateList' in result['data']:
                            comments = result['data']['rateList']
                            all_comments.extend(comments)
                            print(f"成功获取第 {page} 页的 {len(comments)} 条评论")
                        else:
                            print(f"第 {page} 页没有找到评论数据")
                    else:
                        error_msg = f"API调用失败: {result.get('ret')}"
                        print(error_msg)
                        print(f"响应内容: {response.text[:200]}...")
                        self.last_error = error_msg
                        
                        # 如果是鉴权问题，尝试更新Cookie
                        if "FAIL_SYS_TOKEN_EMPTY" in response.text or "FAIL_SYS_ILLEGAL_ACCESS" in response.text:
                            auth_error = "鉴权失败，请更新Cookie和token"
                            print(auth_error)
                            self.last_error = auth_error
                        
                except Exception as e:
                    error_msg = f"解析第 {page} 页响应时出错: {e}"
                    print(error_msg)
                    print(f"响应内容: {response.text[:200]}...")
                    self.last_error = error_msg
                
                # 防止请求过快
                time.sleep(random.uniform(1, 2))
                
            except Exception as e:
                error_msg = f"爬取第 {page} 页评论时出错: {e}"
                print(error_msg)
                self.last_error = error_msg
                continue
                
        return all_comments
    
    def _generate_sign(self, timestamp, data_str):
        """
        根据天猫的签名算法生成正确的sign
        签名公式: md5(token + "&" + timestamp + "&" + appKey + "&" + data)
        """
        if not self.token:
            # 如果没有token，返回一个随机签名（将无法正常工作）
            return hashlib.md5(str(random.random()).encode('utf-8')).hexdigest()
            
        # 正确的签名计算
        sign_str = f"{self.token}&{timestamp}&12574478&{data_str}"
        return hashlib.md5(sign_str.encode('utf-8')).hexdigest()
    
    def save_to_excel(self, comments, output_file=None, filter_empty_comments=False):
        """
        将评论数据保存到Excel文件
        :param comments: 评论数据列表
        :param output_file: 输出文件名，若为None则自动生成
        :param filter_empty_comments: 是否过滤掉空评价（"此用户没有填写评价。"）
        """
        # 提取所有可能的字段
        data = []
        
        # 如果未指定输出文件名，则根据爬取内容自动生成
        if output_file is None and comments:
            # 获取商品ID和商品标题
            item_id = comments[0].get('auctionNumId', '')
            item_title = comments[0].get('auctionTitle', '')
            
            # 处理标题，移除特殊字符，限制长度
            if item_title:
                # 移除文件名中不允许的字符
                item_title = re.sub(r'[\\/:*?"<>|]', '', item_title)
                # 限制标题长度，避免文件名过长
                if len(item_title) > 30:
                    item_title = item_title[:30] + '...'
            
            # 生成文件名：商品ID_商品标题_评论数量_日期.xlsx
            current_date = time.strftime("%Y%m%d_%H%M%S", time.localtime())
            output_file = f"{item_id}_{item_title}_{len(comments)}条评论_{current_date}.xlsx"
        
        # 如果仍然没有文件名（如评论为空），使用默认文件名
        if not output_file:
            output_file = f"天猫商品评论_{time.strftime('%Y%m%d%H%M%S', time.localtime())}.xlsx"
        
        # 记录过滤掉的空评价数量
        filtered_count = 0
        
        for comment in comments:
            try:
                # 获取评论内容
                feedback = comment.get('feedback', '')
                
                # 如果启用了空评价过滤，且评论内容为"此用户没有填写评价。"，则跳过
                if filter_empty_comments and feedback == "此用户没有填写评价。":
                    filtered_count += 1
                    continue
                
                # 创建一个全面的评论数据字典
                item = {
                    # 基本信息
                    '用户昵称': comment.get('userNick', ''),
                    '评论内容': feedback,
                    '评论时间': comment.get('createTime', ''),
                    '评论时间间隔': comment.get('createTimeInterval', ''),
                    '评价日期': comment.get('feedbackDate', ''),
                    '评论ID': comment.get('id', ''),
                    '商品ID': comment.get('auctionNumId', ''),
                    '商品标题': comment.get('auctionTitle', ''),
                    'SKUID': comment.get('skuId', ''),
                    
                    # 商品规格
                    '商品规格': ', '.join([f"{k}: {v}" for k, v in comment.get('skuMap', {}).items()]) if comment.get('skuMap') else '',
                    '规格字符串': comment.get('skuValueStr', ''),
                    
                    # 评价相关
                    '评价类型': "好评" if comment.get('rateType') == "1" else ("中评" if comment.get('rateType') == "0" else "差评"),
                    '是否匿名': "是" if comment.get('annoy') == "1" else "否",
                    '是否置顶': "是" if comment.get('topRate') == "1" else "否",
                    '是否有详情': "是" if comment.get('hasDetail') == "1" else "否",
                    '是否复购': "是" if comment.get('repeatBusiness') == "1" else "否",
                    '是否金牌用户': "是" if comment.get('goldUser') == "1" else "否",
                    '是否黑名单用户': "是" if comment.get('formalBlackUser') == "true" else "否",
                    '是否复制': "是" if comment.get('copy') == "true" else "否",
                    '是否本人': "是" if comment.get('own') == "true" else "否",
                    '结构标签结束大小': comment.get('structTagEndSize', ''),
                    
                    # 互动信息
                    '点赞数': comment.get('interactInfo', {}).get('likeCount', 0),
                    '评论数': comment.get('interactInfo', {}).get('commentCount', 0),
                    '阅读数': comment.get('interactInfo', {}).get('readCount', 0),
                    '是否已点赞': "是" if comment.get('interactInfo', {}).get('alreadyLike') == "true" else "否",
                    '可否评论': "是" if comment.get('interactInfo', {}).get('enableComment') == "true" else "否",
                    '可否点赞': "是" if comment.get('interactInfo', {}).get('enableLike') == "true" else "否",
                    '可否分享': "是" if comment.get('interactInfo', {}).get('enableShare') == "true" else "否",
                    
                    # 商家回复
                    '商家回复': comment.get('reply', ''),
                    
                    # 用户信息
                    '用户ID': comment.get('userId', ''),
                    '用户信用等级': comment.get('creditLevel', ''),
                    '用户星级': comment.get('userStar', ''),
                    '用户头像URL': comment.get('headPicUrl', ''),
                    '用户头像框URL': comment.get('headFrameUrl', ''),
                    '用户主页URL': comment.get('userIndexURL', ''),
                    '用户标记': comment.get('userMark', ''),
                    '减少用户昵称': comment.get('reduceUserNick', ''),
                    
                    # 分享信息
                    '分享URL': comment.get('share', {}).get('shareURL', ''),
                    '详情URL': comment.get('share', {}).get('detailUrl', ''),
                    '详情分享URL': comment.get('share', {}).get('detailShareUrl', ''),
                    '支持分享': "是" if comment.get('share', {}).get('shareSupport') == "true" else "否",
                    
                    # 添加购物车URL
                    '添加购物车URL': comment.get('addCartUrl', ''),
                    
                    # 权限信息
                    '允许评论': "是" if comment.get('allowComment') == "true" else "否",
                    '允许互动': "是" if comment.get('allowInteract') == "true" else "否",
                    '允许笔记': "是" if comment.get('allowNote') == "true" else "否",
                    '允许举报评论': "是" if comment.get('allowReportReview') == "true" else "否",
                    '允许举报用户': "是" if comment.get('allowReportUser') == "true" else "否",
                    '允许屏蔽评论': "是" if comment.get('allowShieldReview') == "true" else "否",
                    '允许屏蔽用户': "是" if comment.get('allowShieldUser') == "true" else "否",
                    
                    # 额外信息
                    '用户等级': comment.get('extraInfoMap', {}).get('userGrade', ''),
                    '举报URL': comment.get('extraInfoMap', {}).get('report_url', ''),
                }
                
                # 用户标签列表
                if 'userTagList' in comment and comment['userTagList']:
                    for i, tag in enumerate(comment['userTagList']):
                        item[f'用户标签_{i+1}_代码'] = tag.get('tagCode', '')
                        item[f'用户标签_{i+1}_描述'] = tag.get('tagDesc', '')
                        item[f'用户标签_{i+1}_图标'] = tag.get('tagIconPic', '')
                
                data.append(item)
            except Exception as e:
                print(f"处理评论数据时出错: {e}")
                continue
        
        # 创建DataFrame并保存
        if data:
            df = pd.DataFrame(data)
            df.to_excel(output_file, index=False)
            print(f"评论数据已保存到 {output_file}")
            
            # 显示过滤信息
            if filter_empty_comments and filtered_count > 0:
                print(f"已过滤 {filtered_count} 条空评价")
        else:
            print("没有评论数据可以保存")
            if filter_empty_comments and filtered_count > 0:
                print(f"所有 {filtered_count} 条评论均为空评价，已全部过滤")

def main():
    # 使用示例
    crawler = TmallCommentCrawler()
    
    # 请输入商品ID
    item_id = input("请输入商品ID (例如: 714871191114): ")
    
    # 请用户输入要爬取的页数
    try:
        page_num = int(input("请输入要爬取的页数 (每页20条评论，默认5页): ") or "5")
        if page_num <= 0:
            print("页数必须大于0，已设置为默认值5页")
            page_num = 5
    except ValueError:
        print("输入无效，已设置为默认值5页")
        page_num = 5
    
    # 询问用户是否过滤空评价
    filter_empty = input("是否过滤空评价 (\"此用户没有填写评价。\") (y/n, 默认n): ").lower() == 'y'
    
    print(f"即将爬取{page_num}页，共{page_num*20}条评论...")
    if filter_empty:
        print("已启用空评价过滤")
    
    # 获取评论
    comments = crawler.get_comments(item_id, page_num)
    
    # 保存到Excel，使用自动生成的文件名
    crawler.save_to_excel(comments, filter_empty_comments=filter_empty)
    
    print(f"共获取 {len(comments)} 条评论")

if __name__ == "__main__":
    main() 