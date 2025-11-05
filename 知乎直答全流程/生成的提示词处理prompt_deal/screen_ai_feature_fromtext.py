import re
'''
    todo
    选第一句（第一个;之前）
    #  有   未找到 字样
    #  有  并非 不属于 需要审查
 #  需要删除句子： 整理   介绍 根据    依据 药品介绍;
 
#  如下  删
'''


def clean_english(text):
    # 删除括号中含有超过4个字母英文单词的内容（含括号）
    text = re.sub(r'（[^（）]*[a-zA-Z]{5,}[^（）]*）', '', text)

    # 英文前10个中文字内如果出现 '别名',则删除''别名''到单词的内容，比如：阿普米斯特片别名是Apremilast是一种口服磷酸二酯酶4 删除后 ：阿普米斯特片是一种口服磷酸二酯酶4
    # 使用正则匹配"别名"到英文单词的模式，别名和英文单词之间可能有中文字符
    pattern = re.compile(r'别名[^a-zA-Z]{0,10}[a-zA-Z]{5,}')
    text = pattern.sub('', text)

    # 删除单独出现的超过4个字母英文单词
    text = re.sub(r'\b[a-zA-Z]{5,}\b', '', text)

    # # 清理多余空格
    # text = re.sub(r'\s+', ' ', text).strip()

    return text

def extract_from_txt(file_path):
    results = []
    pattern = re.compile(r"(?:根据百度|根据权威|基于|根据现有搜索|根据公开|以下是)(.*?)[，：:]")  # 匹配“根据xxx:”或“基于xxx:”

    pattern2 = re.compile(r'（字数：(.*?)）')



    # 字数 ，以下，以上
    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()

            # match = pattern.search(line)
            # if match:
            #     results.append(match.group(0).strip())  # 提取括号里的内容
            # match2 = pattern2.search(line)
            # if match2:
            #     results.append(match2.group(0).strip())  # 提取括号里的内容


            # 删除匹配到的部分
            cleaned = pattern.sub("", line)
            cleaned2 = pattern2.sub("", cleaned)

            results.append(cleaned2.strip())

    return results


if __name__ == "__main__":
    file_path = r"D:\sort\selenium文档\新建 文本文档.txt"  # 你的txt文件
    extracted = extract_from_txt(file_path)

    for line in extracted:
        line = clean_english(line)
        print( line)
