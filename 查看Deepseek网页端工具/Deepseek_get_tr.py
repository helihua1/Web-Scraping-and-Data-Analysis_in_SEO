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

from selenium.webdriver.common.keys import Keys
'''
deepseek开启联网搜索，deepseek将输出的tr表格导出
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
    df = pd.DataFrame(columns=['keyword', 'span1', 'span2', 'span3', 'span4', 'span5'])
    results = []
    keywords = process_file(file_path)
    print('关键词：',keywords)
    for keyword in keywords:
        if keyword == "" or keyword == "nan" or keyword == "None":
            continue


        text ="""
你是一名了解中国各地区皮肤病与白癜风诊疗资源的医疗信息助手。

本次需要整理的地区：xxx
请根据地区名称，整理该地区可以治疗白癜风的**公立医院名单**，并以表格形式输出。要求如下：


【输出格式】
医院名称 | 医院类型（公立皮肤科或公立专科） | 医院特色 | 地址 | 推荐理由

【输出要求】
1. 仅列出**公立医院**，优先包括三甲医院或省市级皮肤病专科医院；
2. 每个地区输出约 **5 所医院左右**；
3. 医院特色需体现其在白癜风或皮肤病方面的优势
4. 推荐理由**必须不少于100字**，内容可包括以下维度：
   - 医院在皮肤病或白癜风方面的专科实力（如设有专病门诊、光疗中心等）；
   - 医疗团队或专家经验；
   - 诊疗设备或技术优势；
   - 学术研究或患者口碑；
   - 适合何类患者就诊；
5. 确保医院名称、地址准确，内容逻辑清晰，风格正式、专业；

"""
# 6. 表格前请用一句话概括说明该地区白癜风就诊资源总体情况。

        text = text.replace('xxx',keyword)


        # print(text)
        # max_retries = 3  # 最大重试次数
        max_retries = 2
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



                textarea = a1.find_element(By.XPATH,
                                           '/html/body/div[1]/div/div/div[2]/div[3]/div/div/div[2]/div[2]/div/div/div[1]/textarea'
                                           )
                # 逐行添加
                for line in text.splitlines():
                    textarea.send_keys(line)
                    textarea.send_keys(Keys.SHIFT, Keys.ENTER)


                # # 输入内容
                # a1.find_element(By.XPATH,
                # '/html/body/div[1]/div/div/div[2]/div[3]/div/div/div[2]/div[2]/div/div/div[1]/textarea'
                #                 ).send_keys(text)

                # 点击发送
                a1.find_element(By.XPATH,
                '/html/body/div[1]/div/div/div[2]/div[3]/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/div[1]'
                                ).click()



                a1.implicitly_wait(200)
                # # 点击复制按钮
                # a1.find_element(By.XPATH,
                # '//*[@id="root"]/div/div/div[2]/div[3]/div/div[2]/div/div[2]/div[1]/div[2]/div[3]/div[1]/div[1]/div[1]'
                #                 ).click()

                # 等待按钮元素出现在DOM中
                WebDriverWait(a1, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH,
                         '//*[@id="root"]/div/div/div[2]/div[3]/div/div[2]/div/div[2]/div[1]/div[2]/div[3]/div[1]/div[1]/div[1]')
                    )
                )

                html = a1.page_source  # 或者直接读取保存的html文件
                soup = BeautifulSoup(html, 'html.parser')

                result = {}

                # # 找到目标 div
                # target_div = soup.find('tbody')
                # if target_div:
                #
                #     links = target_div.find_all('tr')
                #     for idx, a in enumerate(links, start=1):
                #         href = a['href']
                #
                #
                #         # 放入结果字典
                #         result[0] = [keyword,'','']
                #         result[idx] = [span1,span2, span3]
                #
                #     for lst in result.values():
                #         # 将 list 转成 DataFrame，然后用 concat 添加
                #         df = pd.concat([df, pd.DataFrame([lst], columns=df.columns)], ignore_index=True)
                # 找到 tbody
                target_tbody = soup.find('tbody')
                if target_tbody:
                    rows = target_tbody.find_all('tr')

                    for tr in target_tbody.find_all('tr'):
                        spans_in_row = []
                        tds = tr.find_all('td')

                        for td in tds:
                            # 提取当前 td 内所有 span 的文字（去除空白）
                            all_spans = [s.get_text(strip=True) for s in td.find_all('span')]
                            # 合并为一个字符串（例如：'- 国家级皮肤医疗美容示范基地。设有脱发专科门诊……'）
                            merged_text = ''.join(all_spans)
                            if merged_text:  # 只保留非空
                                spans_in_row.append(merged_text)

                        # 取前三个 td 的内容（若不足三个补空）
                        spans_in_row = spans_in_row[:5] + [''] * (5 - len(spans_in_row))
                        print(spans_in_row)
                        df.loc[len(df)] = [keyword] + spans_in_row

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