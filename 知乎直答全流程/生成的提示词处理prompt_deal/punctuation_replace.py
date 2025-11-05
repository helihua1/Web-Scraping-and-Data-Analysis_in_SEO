# -*- coding: utf-8 -*-
import os
import re


def full_to_half(text):
    """
    将字符串中的全角字符转换为半角字符。
    主要处理：字母、数字、空格、标点符号。
    """
    # 创建全角字符到半角字符的映射表
    # 范围包括：全角字母、数字、空格、常见标点
    full_width_chars = "ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ０１２３４５６７８９　！＂＃＄％＆＇（）＊＋，－．／：；＜＝＞？＠［＼］＾＿｀｛｜｝～"
    half_width_chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 !\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~"

    # 创建映射表
    mapping_table = str.maketrans(full_width_chars, half_width_chars)

    # 使用映射表进行转换
    return text.translate(mapping_table)


if __name__ == "__main__":
    # 定义替换规则
    replacements = {
        '①': '',
        '②': '',
        '③': '',
        '④': '',
        '⑤': '',
        '⑥': '',
        '（': '(',
        '）': ')',
        '”': '',
        '“': '',
        '’': '',
        '‘': '',
        '"':'',
        '、':',',
        "~":'-',
        '。': ';',
        '×':'x',
        '®':'',
        ',,':'',
        '()':'',
        # '/片': '1片',
        # '/次': '1次',
        # '/日': '1天',
        # '/支': '1支',
        # '/': '，',
        '≤': '小于等于',
        '≥': '大于等于',
        '】': '',
        '【': '',
        '，':',',
        '；':';',
        '：':':'

    }

    # 文件路径
    input_file = r'D:\sort\selenium文档\新建 文本文档.txt'   # 原始文件
    output_file = input_file # 替换后的文件
    pattern = re.compile(r'\b(\d+)\.(?!\d)')  # \b保证是单词边界，(?!\d)表示后面不是数字

    # 读取文件
    with open(input_file, 'r', encoding='utf-8') as f:
        content = f.read()


    # ”1.”“2.”等 如果右边不是数字就删掉，
    content = pattern.sub('', content)

    content = full_to_half(content)

    # 逐条替换
    for old, new in replacements.items():
        content = content.replace(old, new)

    # 去除 注： 的内容
    content = re.sub(r'注:.*?;', '', content)
    os.remove(input_file)
    content = re.sub(r'\(注:.*?\)', '', content)
    os.remove(input_file)
    print(f"文件已删除: {input_file}")

    # 写入新文件
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(content)


    print("替换完成，结果已保存到", output_file)
