import os
from itertools import product
from openpyxl import Workbook

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
    # replace_dict[ph]是list ，所以结果是[[],[],...]
    lists_to_product = [replace_dict[ph] for ph in placeholders]

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
        results.append((temp, template))  # 保存替换后的句子和原模版

# 4. 输出到 Excel
wb = Workbook()
ws = wb.active
# ws.title = "结果"
# ws.append(["替换后句子", "原模版句子"])  # 表头

for replaced, original in results:
    ws.append([replaced, original])

wb.save("output-模版输出结果.xlsx")
print("生成完成，已保存为 output.xlsx")
