import os
from itertools import product
from openpyxl import Workbook
'''
同目录下有一个 模版.txt 文件，内容有如：
【地区】治疗【病种】医院排名
等等

遍历读取每一行，空行跳过。作为一个list。
同目录下有一个 替换数据 文件夹，有多个txt文件
【地区】.txt   
【病种】.txt   

分别读取每个txt的文件名，遍历读取里面每一行，如
【地区】.txt   中是：
济南
石家庄
长春

【病种】.txt   中是：
白癜风
银屑病

根据 替换数据 文件夹中的 'txt文件名'作为识别占位符，和‘txt内容’作为替换的内容 ，替换 模版.txt 中的 内容。

济南治疗白癜风医院排名
石家庄治疗白癜风医院排名
长春治疗白癜风医院排名
济南治疗银屑病医院排名
石家庄治疗银屑病医院排名
长春治疗银屑病医院排名

所有替换完毕后，输出为  为excel

'''
# 1. 读取模版文件
with open("模版.txt", "r", encoding="utf-8") as f:
    templates = [line.strip() for line in f if line.strip()]  # 去掉空行

# 2. 读取替换数据文件夹
replace_folder = "替换数据"
replace_dict = {}

for filename in os.listdir(replace_folder):
    if filename.endswith(".txt"):
        key = filename.replace(".txt", "")  # 占位符名，如 【地区】 或 【病种】
        path = os.path.join(replace_folder, filename)

        # open() 打开文件后会占用系统资源（文件句柄）。
        # 使用 with 可以确保在代码块执行完毕后自动关闭文件，无论期间是否出现异常。
        # 如果不使用 with，你必须手动调用 f.close()，否则可能导致文件一直被占用，尤其是大量文件时容易出错。
        with open(path, "r", encoding="utf-8") as f:
            replace_dict[key] = [line.strip() for line in f if line.strip()]

# 3. 多重替换，生成笛卡尔积
results = []

for template in templates:
    # 找出模版中所有占位符
    placeholders = [ph for ph in replace_dict.keys() if ph in template]

    if not placeholders:
        results.append((template, template))  # 没有占位符，替换后与原模版相同
        continue

    # 对占位符进行排序：先【地区】, 再【病种】，其他按原顺序
    sorted_placeholders = []
    if '【地区】' in placeholders:
        sorted_placeholders.append('【地区】')
    if '【病种】' in placeholders:
        sorted_placeholders.append('【病种】')
    for ph in placeholders:
        if ph not in sorted_placeholders:
            sorted_placeholders.append(ph)


    # replace_dict[ph]是list ，所以结果是[[],[],...]
    lists_to_product = [replace_dict[ph] for ph in sorted_placeholders]

    # product(*lists_to_product) 会生成 所有可能组合（笛卡尔积）：
    # 比如，3个list：地区list，等级，疾病
    # ('济南', '三甲', '白癜风')
    # ('济南', '三甲', '银屑病')
    # ('济南', '二甲', '白癜风')
    # ('济南', '二甲', '银屑病')
    # ('石家庄', '三甲', '白癜风')
    # ('石家庄', '三甲', '银屑病')
    # ('石家庄', '二甲', '白癜风')
    # ('石家庄', '二甲', '银屑病')
    for combo in product(*lists_to_product):
        temp = template
        # zip() 是 Python 内置函数，用于 把多个可迭代对象按顺序“打包”成元组。当任意一个可迭代对象耗尽时，zip 就停止。
        # a = [1, 2, 3]
        # b = ['a', 'b', 'c']
        # for x, y in zip(a, b):
        #     print(x, y)
        # 1 a
        # 2 b
        # 3 c
        # for ph, val in ... → 循环解包元组，每次取一个占位符和对应的值。
        for ph, val in zip(placeholders, combo):
            temp = temp.replace(ph, val)
        # 保存替换后的句子，原模板，和每个占位符的值
        results.append([temp, template, *combo])

# 4. 输出到 Excel
wb = Workbook()#Workbook() 是 创建一个新的 Excel 工作簿（Workbook）。
ws = wb.active#wb.active 获取当前活跃的工作表（默认是第一个 Sheet）。

for row in results:
    ws.append(row)

wb.save("output-模版输出结果.xlsx")
print("生成完成，已保存为 output.xlsx")
