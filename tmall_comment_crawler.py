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
            'Cookie': '你的cookie',
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
        
    def get_comments(self, item_id, page_num=5):
        """
        获取商品评论
        :param item_id: 商品ID
        :param page_num: 页数，每页20条评论，默认抓取5页共100条
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
                "orderType": "feedbackdate",
                "expression": "",
                "rateSrc": "pc_rate_list"
            }
            
            # 使用正确的方式生成签名
            data_str = json.dumps(data)
            sign = self._generate_sign(timestamp, data_str)
            
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
                        print(f"API调用失败: {result.get('ret')}")
                        print(f"响应内容: {response.text[:200]}...")
                        
                        # 如果是鉴权问题，尝试更新Cookie
                        if "FAIL_SYS_TOKEN_EMPTY" in response.text or "FAIL_SYS_ILLEGAL_ACCESS" in response.text:
                            print("鉴权失败，请更新Cookie和token")
                        
                except Exception as e:
                    print(f"解析第 {page} 页响应时出错: {e}")
                    print(f"响应内容: {response.text[:200]}...")
                
                # 防止请求过快
                time.sleep(random.uniform(1, 2))
                
            except Exception as e:
                print(f"爬取第 {page} 页评论时出错: {e}")
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
    
    def save_to_excel(self, comments, output_file=None):
        """
        将评论数据保存到Excel文件
        :param comments: 评论数据列表
        :param output_file: 输出文件名，若为None则自动生成
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
            current_date = time.strftime("%Y%m%d", time.localtime())
            output_file = f"{item_id}_{item_title}_{len(comments)}条评论_{current_date}.xlsx"
        
        # 如果仍然没有文件名（如评论为空），使用默认文件名
        if not output_file:
            output_file = f"天猫商品评论_{time.strftime('%Y%m%d%H%M%S', time.localtime())}.xlsx"
        
        for comment in comments:
            try:
                # 创建一个全面的评论数据字典
                item = {
                    # 基本信息
                    '用户昵称': comment.get('userNick', ''),
                    '评论内容': comment.get('feedback', ''),
                    '评论时间': comment.get('createTime', ''),
                    '评论时间间隔': comment.get('createTimeInterval', ''),
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
            try:
                # 尝试保存为Excel
                df.to_excel(output_file, index=False)
                print(f"评论数据已保存到 {output_file}")
            except ModuleNotFoundError:
                # 如果缺少openpyxl库，保存为CSV作为备选
                csv_file = output_file.replace('.xlsx', '.csv')
                df.to_csv(csv_file, index=False, encoding='utf-8-sig')
                print(f"缺少openpyxl库，无法保存为Excel。数据已保存到CSV文件: {csv_file}")
                print("提示: 可以通过命令 'pip install openpyxl' 安装Excel支持")
        else:
            print("没有评论数据可以保存")

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
    
    print(f"即将爬取{page_num}页，共{page_num*20}条评论...")
    
    # 获取评论
    comments = crawler.get_comments(item_id, page_num)
    
    # 保存到Excel，使用自动生成的文件名
    crawler.save_to_excel(comments)
    
    print(f"共获取 {len(comments)} 条评论")

if __name__ == "__main__":
    main() 