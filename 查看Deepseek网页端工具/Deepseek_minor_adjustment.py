import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pyperclip  # 添加这个导入
import os
import re
import openpyxl
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
'''
deepseek开启联网搜索，deepseek将输出的微调后的结果输出到控制台
'''

def process_text(text):
    return text

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
    df = pd.DataFrame(columns=['keyword', 'span1', 'span2', 'span3'])
    results = []
    keywords = process_file(file_path)


    # 每次10条
    for i in range(0, len(keywords), 10):
        keyword = keywords[i:i + 10]
        # 拼接
        keyword = "\n".join(keyword)

        if keyword == "" or keyword == "nan" or keyword == "None":
            continue

        #提示词
        prompt = """
You are a professional text rewriting model.
Your task is to slightly rewrite the given sentence according to the following strict rules:

1. The meaning of the original sentence must stay the same.
2. Keep all key feature words in the sentence (do not remove or change them).
   Example: in "Nanjing vitiligo hospital ranking", keep words like "Nanjing", "vitiligo", "hospital".
3. Select 2–5 other words or parts of the sentence to modify using one or more of these operations:
   - Replace with a close synonym
   - Add a short word or phrase
   - Remove a minor word
4. The total sentence length must not change by more than ±30%.
5. Output only the rewritten version, with no explanations."""
        text = keyword   +  prompt

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
                # print(text)
                # # 输入内容
                # a1.find_element(By.XPATH,
                # '/html/body/div[1]/div/div/div[2]/div[3]/div/div/div[2]/div[2]/div/div/div[1]/textarea'
                #                 ).send_keys(text)

                textarea = a1.find_element(By.XPATH,
                                           '/html/body/div[1]/div/div/div[2]/div[3]/div/div/div[2]/div[2]/div/div/div[1]/textarea'
                                           )

                # 逐行添加
                for line in text.splitlines():
                    textarea.send_keys(line)
                    textarea.send_keys(Keys.SHIFT, Keys.ENTER)


                    # 点击发送
                a1.find_element(By.XPATH,
                '/html/body/div[1]/div/div/div[2]/div[3]/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/div[1]'
                                ).click()

                a1.implicitly_wait(300)
                # 点击复制按钮
                a1.find_element(By.XPATH,
                '//*[@id="root"]/div/div/div[2]/div[3]/div/div[2]/div/div[2]/div[1]/div[2]/div[3]/div[1]/div[1]/div[1]'
                                ).click()
                # 读取剪切板
                try:
                    clipboard_content = pyperclip.paste()

                    text = process_text(clipboard_content)
                    print(text)
                    results.append([keyword, text])
                    success = True  # 标记成功
                except Exception as e:
                    print(f"读取剪切板失败：{e}")
                    success = True  # 即使剪切板失败也标记为成功，避免无限重试

                time.sleep(3)
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
