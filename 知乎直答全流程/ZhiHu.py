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
'''
生成提示词并输入知乎问答获取结果
'''
def process_text(text):
    # 1. 如果行中有 ###，行尾加冒号
    text = re.sub(r"(.*###.*?)(\n|$)", r"\1：\2", text)

    # 2. 去掉 * -- 空格 # 和换行符
    text = re.sub(r'(\*|\s|#|--)', "", text)
    text = text.replace("\n", "")  # 拼成一行

    # 3. 去掉数字编号，如 1. 2. 3. 9.
    text = re.sub(r"\d+\.", "", text)

    return text

def process_file(path):
    df = pd.read_excel(path, header=None, usecols=[0])
    sentences = df[0].tolist()
    sentences = [str(sentence) for sentence in sentences]
    return sentences

if __name__ == "__main__":
    file_path = r"D:\sort\selenium文档\zbw 生成提示词.xlsx"
    output_file_path = r"D:\sort\selenium文档\查找知乎直答_查找结果.xlsx"

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
    
    results = []
    keywords = process_file(file_path)
    for keyword in keywords:
        if keyword == "" or keyword == "nan" or keyword == "None":
            continue
        text = f"根据百度百科和百度健康医典和药品网，根据药物名称生成的关于药物的内容介绍【{keyword}】【别名】，【是否为激素药物】，【处方药or非处方药】，【剂型和规格】，【价格范围】，【适用疾病】，【用法用量】，【不良反应】，【禁忌】，【注意事项】，【特殊人群用药】，【是否能用于治疗银屑病】，生成400字以内的介绍，生成结果不要带任何表情emoji，不要包含表格，不要出现其他网站链接，不要出现序列号，一个药品内容介绍一个段落；"
        # 仪器
        # text = f"根据百度百科和百度健康医典，结合【仪器：vd3治疗仪】生成关于仪器的以下内容【别名】，【工作原理】，【价格收费】，【治疗周期】，【不良反应】，【禁忌】，【注意事项】，【是否用于治疗银屑病】，生成结果不要带任何表情emoji，不要包含表格，生成400字以内的介绍，一个仪器内容介绍一个段落；"
        
        max_retries = 3  # 最大重试次数
        retry_count = 0
        success = False
        
        while retry_count < max_retries and not success:
            try:
                a1 = webdriver.Chrome(service=Service('chromedriver.exe'), options=q1)
                a1.get("https://zhida.zhihu.com/")
                a1.implicitly_wait(10)

                # a1.find_element(By.XPATH,'//*[@id="fullScreen"]/div[1]/div/div/div[2]/div/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div/div').click()
                a1.find_element(By.XPATH,'//*[@id="fullScreen"]/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div/div').click()

                # a1.find_element(By.XPATH,'//*[@id="fullScreen"]/div[1]/div/div/div[2]/div/div/div/div[1]/div[2]/div/div/div[1]/div[2]/div/div/div/div/div').send_keys(text)
                a1.find_element(By.XPATH,'//*[@id="fullScreen"]/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div/div/div[1]/div/div/div').send_keys(text)

                # a1.find_element(By.XPATH,'//*[@id="fullScreen"]/div[1]/div/div/div[2]/div/div/div/div[2]/div[2]/div[3]').click()
                a1.find_element(By.XPATH,'//*[@id="fullScreen"]/div[1]/div/div/div[2]/div/div/div[2]/div[2]/div[3]').click()

                a1.implicitly_wait(300)
                a1.find_element(By.XPATH, '//*[@id="fullScreen"]/div[1]/div/div[2]/div[1]/div[1]/div/div[1]/div/div[3]/div/div/div/div[4]/div/div[1]/div[1]/div').click()
                # 读取剪切板
                try:
                    clipboard_content = pyperclip.paste()

                    text = process_text(clipboard_content)
                    print(text)
                    results.append([keyword,text])
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

    agg_attr1 = [item[0] for item in results]
    agg_attr2 = [item[1] for item in results]

    result_data = {
        0: agg_attr1,
        1: agg_attr2
    }

    result_df = pd.DataFrame(result_data)

    # 保存结果到新的Excel文件
    result_df.to_excel(output_file_path, index=False, header=False)
    print(f"处理完成！结果已保存至：{output_file_path}")