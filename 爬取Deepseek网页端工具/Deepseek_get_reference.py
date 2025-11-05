import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
# import pyperclip  # 添加这个导入
import os
import re
import openpyxl

from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
'''
deepseek开启联网搜索，查看引用了哪些网站
'''




'''
处理excel文档输出成列表
'''
def process_file(path):
    df = pd.read_excel(path, header=None, usecols=[0])
    sentences = df[0].tolist()
    sentences = [str(sentence) for sentence in sentences]
    return sentences






if __name__ == "__main__":
    file_path = r"D:\sort\selenium文档\deepseek查看引用网站\deepseek查看引用网站.xlsx"
    output_file_path = r"D:\sort\selenium文档\deepseek查看引用网站\deepseek查看引用网站_查找结果.xlsx"

    # sandbox 是 Chrome 的安全机制，默认启用。
    # 在某些 Linux 环境（特别是 root 权限运行或容器/虚拟机环境中），启用 sandbox 会导致 Chrome 无法正常启动，所以加上 --no-sandbox 来禁用它。
    # 在 Windows 下一般没影响。
    q1 = Options()
    q1.add_argument('--no-sandbox')
    q1.add_experimental_option('detach', True)
    # 设置窗口位置为屏幕右下角
    q1.add_argument('--window-position=1500,900')  # 调整坐标值以适应您的屏幕
    q1.add_argument('--window-size=800,600')  # 设置窗口大小

    # 添加用户数据目录，保持登录状态
    user_data_dir = os.path.join(os.getcwd(), 'chrome_user_data')
    print(user_data_dir)
    q1.add_argument(f'--user-data-dir={user_data_dir}')
    df = pd.DataFrame(columns=['cite_index', 'source', 'href', 'title', 'snippet', 'date'])
    results = []
    keywords = process_file(file_path)
    print('关键词：',keywords)
    for keyword in keywords:
        if keyword == "" or keyword == "nan" or keyword == "None":
            continue
        # text = f"根据百度百科和百度健康医典和药品网，根据药物名称生成的关于药物的内容介绍【{keyword}】【别名】，【是否为激素药物】，【处方药or非处方药】，【剂型和规格】，【价格范围】，【适用疾病】，【用法用量】，【不良反应】，【禁忌】，【注意事项】，【特殊人群用药】，【是否能用于治疗银屑病】，生成400字以内的介绍，生成结果不要带任何表情emoji，不要包含表格，不要出现其他网站链接，不要出现序列号，一个药品内容介绍一个段落；"
        # 仪器
        # text = f"根据百度百科和百度健康医典，结合【仪器：vd3治疗仪】生成关于仪器的以下内容【别名】，【工作原理】，【价格收费】，【治疗周期】，【不良反应】，【禁忌】，【注意事项】，【是否用于治疗银屑病】，生成结果不要带任何表情emoji，不要包含表格，生成400字以内的介绍，一个仪器内容介绍一个段落；"
        text = keyword
        print(keyword)
        # max_retries = 3  # 最大重试次数
        max_retries = 1
        retry_count = 0
        success = False

        while retry_count < max_retries and not success:
            try:
                a1 = webdriver.Chrome(service=Service('chromedriver.exe'), options=q1)
                a1.get("https://chat.deepseek.com/")
                a1.implicitly_wait(300)

                # 点击发送消息框
                a1.find_element(By.XPATH,
                "/html/body/div[1]/div/div/div[2]/div[3]/div/div/div[2]/div[2]/div/div/div[1]"
                                ).click()

                # 输入内容
                a1.find_element(By.XPATH,
                '/html/body/div[1]/div/div/div[2]/div[3]/div/div/div[2]/div[2]/div/div/div[1]/textarea'
                                ).send_keys(text)

                # 点击发送
                a1.find_element(By.XPATH,
                '/html/body/div[1]/div/div/div[2]/div[3]/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/div[1]'
                                ).click()



                # # 等待引用网站列表中有内容
                # first = '/html/body/div[1]/div/div/div[2]/div[3]/div[1]/div[2]/div/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]/span'
                #
                # WebDriverWait(a1, 20).until(
                #     EC.presence_of_element_located((By.XPATH, first))
                # )

                time.sleep(10)
                # 点击引用网站标签，弹出引用网站列表
                a1.find_element(By.XPATH,
                '/html/body/div[1]/div/div/div[2]/div[3]/div[1]/div[2]/div/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]'
                                ).click()


                time.sleep(2)

                html = a1.page_source  # 或者直接读取保存的html文件
                soup = BeautifulSoup(html, 'html.parser')

                result = {}

                # 找到目标 div
                target_div = soup.find('div', class_='dc433409')
                if target_div:
                    # 遍历所有 <a> 标签
                    links = target_div.find_all('a', href=True)
                    for idx, a in enumerate(links, start=1):
                        href = a['href']

                        # 提取各个 span 内容
                        source_span = a.find('span', class_='d2eca804')
                        source = source_span.get_text(strip=True) if source_span else ""

                        date_span = a.find('span', class_='caa1ee14')
                        date = date_span.get_text(strip=True) if date_span else ""

                        index_span = a.find('span', class_='ds-markdown-cite')
                        cite_index = index_span.get_text(strip=True) if index_span else ""

                        # 找标题（class名中有 search-view-card__title）
                        title_tag = a.find('div', class_='search-view-card__title')
                        title = title_tag.get_text(strip=True) if title_tag else ""

                        # 找摘要（class名中有 search-view-card__snippet）
                        snippet_tag = a.find('div', class_='search-view-card__snippet')
                        snippet = snippet_tag.get_text(strip=True) if snippet_tag else ""

                        # 放入结果字典
                        result[0] = [keyword,'','','','','']
                        result[idx] = [cite_index,source, title,href, snippet,date]

                    for lst in result.values():
                        # 将 list 转成 DataFrame，然后用 concat 添加
                        df = pd.concat([df, pd.DataFrame([lst], columns=df.columns)], ignore_index=True)

                    success = True
                # 输出结果
                for k, v in result.items():
                    print(k, v)
                print('执行完毕')
                a1.close()

            except Exception as e:
                retry_count += 1
                print(f"第{retry_count}次尝试失败：{e}")
                if a1:
                    try:
                        a1.close()
                    except:
                        pass

                if retry_count < max_retries:
                    print(f"正在重试第{retry_count + 1}次...")
                    time.sleep(5)  # 等待5秒后重试
                else:
                    print(f"关键词 '{keyword}' 处理失败，已重试{max_retries}次")



    # 保存结果到新的Excel文件
    df.to_excel(output_file_path, index=False, header=False)
    print(f"处理完成！结果已保存至：{output_file_path}")