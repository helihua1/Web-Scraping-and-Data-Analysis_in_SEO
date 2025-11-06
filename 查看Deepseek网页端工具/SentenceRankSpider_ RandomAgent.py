import pandas as pd
import requests
import time
import random
from bs4 import BeautifulSoup

"""
在百度搜索中查询指定关键词，查找目标URL的排名
:param query: 搜索关键词
:param target_url: 目标URL
:return: 排名 (1-100) 或 "100+" 或 "未找到"
"""


# 随机 User-Agent 列表
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:118.0) Gecko/20100101 Firefox/118.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.1 Safari/605.1.15"
]

# 百度搜索参数
BAIDU_URL = "https://www.baidu.com/s"

def get_baidu_rank(session, query, target_url):
    max_page = 2  # 最多查前 20 条

    for page in range(max_page):
        headers = {
            "User-Agent": random.choice(USER_AGENTS),
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Referer": "https://www.baidu.com/"
        }
        params = {
            "wd": query,  # 不要提前quote，让 requests 自动编码
            "pn": page * 10
        }

        try:
            # params 会被 自动附加到 URL 的查询字符串部分
            response = session.get(BAIDU_URL, params=params, headers=headers, timeout=3)
            response.encoding = "utf-8"
            html = response.text

            # 检测是否被反爬（验证码页面）
            if "百度安全验证" in html or "verify.baidu.com" in html:
                print("⚠检测到验证码页面，暂停 60 秒...")
                time.sleep(60)
                continue

            soup = BeautifulSoup(html, "html.parser")
            content_left = soup.find('div', id='content_left')
            if not content_left:
                continue

            results = content_left.find_all('div', class_='result c-container xpath-log new-pmd')
            for idx, result in enumerate(results, start=1 + page * 10):
                mu_url = result.get('mu')
                print(mu_url)
                if mu_url and target_url in mu_url:
                    rank = result.get('id')
                    print('id=' + rank)
                    return rank

            time.sleep(random.uniform(1.5, 5))  # 随机延时
            print("延时")
        except Exception as e:
            print(f"请求错误: {e}")
            time.sleep(5)

    return "未找到"

def main():
    file_path = r'C:\Users\zhang\Desktop\数据筛选软件\每日任务\查标题在百度中的排名\需要爬.xlsx'

    try:
        df = pd.read_excel(file_path, header=None, usecols=[0, 1, 2])
        print(f"成功读取Excel，共{len(df)}行数据")
    except Exception as e:
        print(f"读取Excel失败: {e}")
        return

    if len(df.columns) < 3:
        print("Excel 至少需要三列数据")
        return

    session = requests.Session()  # 维持会话和 Cookie

    for i, row in df.iterrows():
        url = str(row[1]).strip()
        query = str(row[2]).strip()

        print(f"\n处理第{i+1}行: 关键词='{query}' | 目标URL: {url}")

        rank = get_baidu_rank(session, query, url)
        df.at[i, '排名'] = rank
        print(f"排名结果: {rank}")

        if (i + 1) % 5 == 0:
            df.to_excel(file_path, index=False)
            print(f"已保存进度到第{i + 1}行")

    df.to_excel(file_path, index=False)
    print("✅ 所有数据处理完成！")

if __name__ == "__main__":
    main()
