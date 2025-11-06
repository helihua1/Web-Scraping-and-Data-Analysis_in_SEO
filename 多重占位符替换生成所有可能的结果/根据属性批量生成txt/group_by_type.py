import pandas as pd
import os
'''
excel第一列是属性，第二列是具体的内容。
把相同属性的内容放在一个txt。中，每个txt命名为：属性+自定义。
'''
# ======== 用户配置区域 ========
excel_path = '各机构所在城市 (1).xlsx'  # Excel 文件路径
custom_suffix = '医院'   # 文件名后缀
output_dir = 'output_txt'+ custom_suffix # 输出文件夹

# ============================

# 读取Excel（自动识别第一行作为表头）
df = pd.read_excel(excel_path)

# 检查是否包含所需的两列
if df.shape[1] < 2:
    raise ValueError("Excel 至少需要两列：第一列为属性，第二列为内容")

# 取前两列
df = df.iloc[:, :2]
df.columns = ['属性', '内容']

# 创建输出文件夹
os.makedirs(output_dir, exist_ok=True)

# 按属性分组并输出txt
for attr, group in df.groupby('属性'):
    filename = f"【{attr}{custom_suffix}】.txt"
    filepath = os.path.join(output_dir, filename)

    # 写入每条内容，每行一个
    with open(filepath, 'w', encoding='utf-8') as f:
        for content in group['内容'].dropna():
            f.write(str(content).strip() + '\n')

    print(f"✅ 已生成：{filepath}")

print("\n🎉 所有文件生成完毕！")
