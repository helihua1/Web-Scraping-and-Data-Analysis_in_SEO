import re
import pandas as pd
import requests
from Tools.scripts.verify_ensurepip_wheels import print_notice
from bs4 import BeautifulSoup
import pandas as pd

# 输入的关键词和基础网址
keyword_list = ["193491.html",'188856.html','73160.html','187304.html','94540.html']
base_url = "https://health.fynews.net"
for keyword in keyword_list:
    # 打开并读取日志文件
    with open(r"D:\sort\ai爬虫统计\自行搜索后\health.fynews.net.log", "r") as file:
        log_lines = file.readlines()

    # 用来存放匹配的URL
    urls = []

    # 遍历每一行日志
    for line in log_lines:
        if keyword in line:  # 如果行中包含关键词
            urls.append(line)
            # # 使用正则表达式提取 GET 请求后的 URL
            # match = re.search(r'GET (\S+)', line)
            # if match:
            #     url_path = match.group(1)  # 获取匹配的路径部分
            #     full_url = base_url + url_path  # 拼接完整URL
            #     urls.append(full_url)

    #
    # urls = [url for url in urls if not url.endswith('.ico')]

    print(keyword)
    print('个数为:',len(urls))
    # 用来存放爬取到的数据
    titles = []
    keywords = []

    # # 遍历每个 URL 进行请求并提取数据
    # for url in urls:
    #     try:
    #         # 发起请求
    #         response = requests.get(url, timeout=3)
    #         response.raise_for_status()  # 如果响应状态不是 200，会抛出异常
    #         # 尝试自动推断字符编码
    #         response.encoding = response.apparent_encoding
    #         # 解析 HTML
    #         soup = BeautifulSoup(response.text, 'html.parser')
    #
    #         # 提取 title 和 keywords
    #         title = soup.find('title').get_text() if soup.find('title')\
    #             else '无'
    #         print(keyword,title)
    #
    #         meta_keywords = soup.find('meta', {'name': 'keywords'})
    #         keywords_content = meta_keywords['content'] if meta_keywords else '无'
    #
    #         # 将数据添加到列表
    #         titles.append(title)
    #         keywords.append(keywords_content)
    #
    #     except requests.RequestException as e:
    #         # 如果请求失败，则记录错误
    #         titles.append('请求失败')
    #         keywords.append('请求失败')

    # 将结果保存到 Excel 文件
    df = pd.DataFrame({
        "URL": urls,
        # "Title": titles,
        # "Keywords": keywords
    })
    df.to_excel(rf"D:\sort\ai爬虫统计\自行搜索后\{keyword}统计.xlsx", index=False)

    print("网页数据提取并保存至 extracted_data.xlsx 完成！")
