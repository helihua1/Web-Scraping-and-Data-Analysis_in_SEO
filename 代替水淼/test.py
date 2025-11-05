import re
import sys

import processor
import shutil
from pathlib import Path
import pandas as pd
import os
import openpyxl
"""
词组替换
"""
class URLManager:
    def __init__(self, file_path="picture_url.txt"):
        self.file_path = file_path
        self.used_urls = []  # 存储已使用的URL
        self.available_urls = []  # 存储可用的URL
        self._load_urls()
    
    def _load_urls(self):
        """加载URL文件"""
        if os.path.exists(self.file_path):
            with open(self.file_path, 'r', encoding='utf-8') as f:
                # 读取所有非空行且不包含"已使用："的行
                self.available_urls = [line.strip() for line in f if line.strip() and '已使用：' not in line]
        else:
            self.available_urls = []
    
    def get_next_url(self):
        """获取下一个可用的URL"""
        if not self.available_urls:
            return None
        
        # 获取第一个可用URL
        url = self.available_urls.pop(0)
        # 添加到已使用列表（带前缀）
        used_url = f"已使用：{url}"
        self.used_urls.append(used_url)
        
        return url
    
    def get_url_from_used(self):
        """从已使用的URL中获取一个URL（循环使用）"""
        if not self.used_urls:
            return None
        
        # 获取第一个已使用的URL，移除"已使用："前缀
        used_url = self.used_urls.pop(0)
        url = used_url.replace("已使用：", "")
        
        # 将URL重新添加到已使用列表的末尾，实现循环使用
        self.used_urls.append(used_url)
        
        return url
    
    def save_used_urls(self):
        """保存已使用的URL回文件"""
        # 合并已使用和未使用的URL
        all_urls = self.used_urls + self.available_urls
        
        with open(self.file_path, 'w', encoding='utf-8') as f:
            for url in all_urls:
                f.write(url + '\n')

def insert_picture_url(text: str, url_manager: URLManager) -> str:
    """
    在text的第三行后插入一行URL
    
    Args:
        text: 原始文本
        url_manager: URL管理器实例
    
    Returns:
        插入URL后的文本
    """
    lines = text.split('\n')
    
    # 获取下一个URL
    url = url_manager.get_next_url()
    if url is None:
        #         try:
        #     raise Exception("URL不足请补充!!!!!!")
        # except Exception as e:
        #     print(e)
        # input("按回车键退出...")
        # sys.exit(0)
        
        # 如果可用URL不足，从已使用的URL中取
        url = url_manager.get_url_from_used()
        if url is None:
            try:
                raise Exception("URL不足请补充!!!!!!")
            except Exception as e:
                print(e)
            input("按回车键退出...")
            sys.exit(0)

    
    # 在第三行后插入新行（索引从0开始，所以在索引3的位置插入）
    if len(lines) > 1:
        lines.insert(1, url)
    else:
        # 如果行数不足3行，添加到末尾
        lines.append(url)
    
    return '\n'.join(lines)
def wrap_unlabeled_lines(text):
    """
    每行如果没有标签包裹，就加p标签包裹
    """
    # 分割文本为行
    lines = text.split('\n')
    result_lines = []

    # 正则表达式匹配以<xx>开头和结尾的行
    tag_pattern = re.compile(r'^\s*<[^>]+>.*</[^>]+>\s*$')
    
    # 检测包含英文和字符的行
    # specific_tags_pattern = re.compile(r'^[a-zA-Z0-9\s\.,!?;:\'"\-\(\)\[\]{}@#$%^&*+=_<>\/\\|~`]+$', re.IGNORECASE)
    # 检测是否由'<'开始,'>'结束。
    specific_tags_pattern = re.compile(r'^<.*>$')
    for line in lines:
        # 如果是空行，保持原样
        if not line.strip():
            result_lines.append(line)
            continue

        # 检查是否已经被标签包裹或者  '<'开始,'>'结束
        if tag_pattern.match(line) or specific_tags_pattern.search(line):
            result_lines.append(line)
        else:
            # 用<p>标签包裹
            result_lines.append(f'<p>{line}</p>')

    # 重新组合为完整文本
    return '\n'.join(result_lines)

def insert_txt_name(text, filename):
    """
    在text的第一行前插入一行，内容是<h2>当前txt的名字</h2>
    
    Args:
        text: 原始文本
        filename: txt文件名（不包含路径）
    
    Returns:
        插入标题后的文本
    """
    # 获取文件名（不包含扩展名）
    # name_without_ext = Path(filename).stem
    name_without_ext = filename.rsplit('.', 1)[0]  # 简单去除扩展名
    
    # 创建标题行
    title_line = f"<h2>{name_without_ext}</h2>"
    
    # 在文本开头插入标题行
    return title_line + '\n' + text

def load_replacement_rules(excel_file):
    """从Excel文件加载替换规则"""
    try:
        # 使用openpyxl直接读取，保持原始格式
        workbook = openpyxl.load_workbook(excel_file, data_only=False)  # data_only=False保持公式和格式
        worksheet = workbook.active
        
        replacement_dict = {}
        
        # 遍历所有有数据的行
        for row in worksheet.iter_rows(min_row=1, values_only=False):
            if len(row) >= 2 and row[0].value is not None and row[1].value is not None:
                # 获取原始显示值
                old_word = str(row[0].value).strip()
                new_word = str(row[1].value).strip()
                
                # 如果单元格是百分比格式，获取格式化后的值
                if row[0].number_format and '%' in row[0].number_format:
                    old_word = f"{row[0].value * 100}%" if isinstance(row[0].value, (int, float)) else str(row[0].value)
                
                if row[1].number_format and '%' in row[1].number_format:
                    new_word = f"{row[1].value * 100}%" if isinstance(row[1].value, (int, float)) else str(row[1].value)
                
                if old_word and new_word:
                    replacement_dict[old_word] = new_word
        
        return replacement_dict
    except Exception as e:
        print(f"读取Excel文件出错: {e}")
        return {}

def replace_words_in_text(text, replacement_dict):
    """在文本中执行词组替换"""

    # # 打印替换规则
    # print(f"成功加载 {len(replacement_rules)} 条替换规则")
    #     # 显示替换规则
    # for old, new in list(replacement_rules.items())[:1000]:  # 只显示前1000条
    #     print(f"'{old}' -> '{new}'")
    
    result = text
    # 按照键的长度降序排列，避免短词覆盖长词的问题
    sorted_rules = sorted(replacement_dict.items(), key=lambda x: len(x[0]), reverse=True)
    
    for old_word, new_word in sorted_rules:
        # 使用正则表达式进行精确匹配，避免部分匹配
        pattern = re.escape(old_word)
        result = re.sub(pattern, new_word, result)
    
    return result
def process_txt_files(source_dir: str):
    """
    处理指定目录下所有txt文件
    
    Args:
        source_dir: 源目录路径
    """
    source_path = Path(source_dir)
    
    if not source_path.exists():
        print(f"错误：目录 '{source_dir}' 不存在")
        return
    
    if not source_path.is_dir():
        print(f"错误：'{source_dir}' 不是一个目录")
        return
    
    # 创建目标目录
    target_dir_name = f"{source_path.name}_处理结果"
    target_path = source_path.parent / target_dir_name
    
    # 如果目标目录已存在，先删除
    if target_path.exists():
        shutil.rmtree(target_path)
        print(f"已删除现有目标目录：{target_path}")
    
    # 创建目标目录
    target_path.mkdir(parents=True, exist_ok=True)
    print(f"创建目标目录：{target_path}")
    
    # 统计信息
    processed_count = 0
    error_count = 0
    
    # 遍历源目录下的所有txt文件
    for txt_file in source_path.rglob("*.txt"):
        try:
            print(f"正在处理：{txt_file}")
            
            # 读取原文件内容
            with open(txt_file, 'r', encoding='utf-8', errors='ignore') as f:
                original_content = f.read()
            
            # 使用main方法处理内容
            processed_content = main(original_content, txt_file.name)
            
            # 计算相对路径
            relative_path = txt_file.relative_to(source_path)
            
            # 创建目标文件路径
            target_file_path = target_path / relative_path
            
            # 确保目标文件的父目录存在
            target_file_path.parent.mkdir(parents=True, exist_ok=True)
            
            # 写入处理后的内容
            with open(target_file_path, 'w', encoding='utf-8') as f:
                f.write(processed_content)
            
            print(f"  -> 已保存到：{target_file_path}")
            processed_count += 1
            
        except Exception as e:
            print(f"  -> 处理失败：{e}")
            error_count += 1

    # 使用过的url标记
    url_manager.save_used_urls()
    # 输出统计信息
    print(f"\n处理完成！")
    print(f"成功处理：{processed_count} 个文件")
    print(f"处理失败：{error_count} 个文件")
    print(f"结果保存在：{target_path}")

def clean_style_tags(text: str) -> str:
    # 1.1 删除成对的 <style>...</style>，DOTALL (. 匹配换行)
    text = re.sub(r"<style>.*?</style>", "", text, flags=re.DOTALL | re.IGNORECASE)
    # 1.2 删除单独的 <style> 或 </style>
    text = re.sub(r"</?style>", "", text, flags=re.IGNORECASE)
    return text

def clean_tags(text: str) -> str:

    # 流量站使用的
    s = 'a|img|body|html|head|thead|tbody|meta|img|title|div|menu|section|nav|script|span|hr|center|mark|s|p|hr|address|caption|blockquote|i'
    keywords = s.split('|')

    # 构建正则，匹配<>中包含任意关键字的内容
    # [^>]*
    # 方括号 [] 表示字符集。
    # ^> 表示非 > 的任意字符。
    # * 表示匹配 0 次或多次。
    # 整体意思：匹配 < 后直到遇到 > 之前的任意字符。
    # 例子：<kkbody=123> 中 kkbody=123 就被 [^>]* 匹配到。
    # pattern = r"<[^>]*(" + "|".join(keywords) + r")[^>]*>"

    # 注意这里：标签名要单独捕获，而不是随便包含
    pattern = r"</?\s*(?:" + "|".join(keywords) + r")\b[^>]*>"
    # 替换匹配的部分为空
    result = re.sub(pattern, "", text, flags=re.IGNORECASE)
    return result

def clean_url(text: str) -> str:
    # # 规则1：删除 www. 开头的网址
    # #www\.：这里的 \. 是为了匹配字面上的点 .
    # # 在正则里 . 默认表示“任意字符”，如果你想匹配点本身，就需要 \.
    # # \w，表示字母、数字和下划线
    # pattern1 = r"www\.[\w\.\-\/\?\=\&]*"


    # 改为类似于baidu.com，baidu.cn 这种 ‘.’两侧是英文的，都要删除包括 xx.xx和 xxx.xxx.xxx
    # \b单词边界，确保匹配完整的域名,当前的正则表达式 pattern1 使用了 \b 单词边界，但 \b 只匹配英文单词的边界，不匹配中文字符的边界。删掉\b
    pattern1 = r'(?:[a-zA-Z0-9-]+\.)+[a-zA-Z]{2,}(?:\.[a-zA-Z]{2,})*'
    # 规则2：删除 英文:// 后跟英文数字和符号的URL
    pattern2 = r"[a-zA-Z]+://[\w\.\-\/\?\=\&]*"
    # # 优化版本（更清晰）
    # pattern2_optimized = r"[a-zA-Z]+://[\w./?=&%-]*"

    # 先删除规则2，再删除规则1（顺序可调）
    text_cleaned = re.sub(pattern2, "", text)
    text_cleaned = re.sub(pattern1, "", text_cleaned)
    return text_cleaned

def clean_line_by_keywords(text: str) -> str:
    # 关键字列表
    keywords = [
        "h2", "html", "XX", "此处可加入", "故意留白", "排版错误",
        "请自行补充", "故意留空", "虚构", "------", "<!-- ", "意识流分割", "意识流", "好的，没问题"
    ]

    # 构建正则，用 '|' 连接关键字，并用 re.escape 转义特殊字符
    pattern = "|".join(map(re.escape, keywords))

    # 按行过滤，re.search 匹配行中任意位置
    filtered_lines = [line for line in text.splitlines() if not re.search(pattern, line, flags=re.IGNORECASE)]

    # 合并回文本
    result = "\n".join(filtered_lines)
    return result

def clean_keywords(text: str) -> str:
    s = '?|&emsp;|*|#|最终，|最后，|第一，|然后，|其次，|另外，|还有，|温馨提示，|&ensp;|总之，|接下来，|&bull;|<p></p>|综上，|科研、|教学|、科研|世界级|横空出世后|教学|、科研|世界级|横空出世后|首屈一指|纳晶微晶祛白技术|、、|<\li>|<\ol>|<\p>|科研|<h3/>|国际|国内外|韩国|德国|日本|<p>. </p>|<p>li></p>|<h3></h3>|笔者|国际|单独一个|免费复查|分期付款|免费义诊|重点|免费|头个|首个|连锁|化学剥脱术|威特|意大利|英国|军队|国家|原装|首批|创建于2016年|教授|成立于2016年|统一写|成立于2016年|创办于2016年|自2016年成立以来|于2016年创建|创立于2016年|2016年创立|始建于2016年|创建于2016年|始建于2016年|建院于2016年|自2016年创建至今|2016年成立至今|于2016年成立|自2016年建院以来|医院不主动提及非国营、非私企、无分院等信息|非私企|及非国营|主动提及|分院|标签|创立于2016年|始建于2016年的|成立时间为2016年|创建于2016年|2016年建院以来|2016年建院至今|创于2016年|医院于2016年建成|其于2016年成立|医院成立于 2016 年|医院于2016年建成|自2016年创立以来|世界|首批|认证|无毒|无害|不含激素|没有毒|有毒|创建于2012年|成立于2012年|创办于2012年|自2012年成立以来|于2012年创建|创立于2012年|2012年创立|始建于2012年|建院于2012年|自2012年创建至今|2012年成立至今|于2012年成立|自2012年建院以来|始建于2012年的|成立时间为2012年|2012年建院以来|2012年建院至今|创于2012年|医院于2012年建成|其于2012年成立|医院成立于 2012 年|自2012年创立以来|所以，|`|<p></p>|虽然不是医保定点医院，但|<!--img>|再次注意，|图片：'
    keywords = s.split('|')

    # 构建正则表达式（转义特殊字符）
    pattern = '|'.join(map(re.escape, keywords))

    # 替换为空字符串
    clean_text = re.sub(pattern, '', text)
    return clean_text

def clean_between_Parentheses(text: str) -> str:
    # # 定义左右括号变量
    # a = '（此处'
    # a1 = '(此处'
    # a2 = '（文中'
    # a3 = '(文中'
    # b = '）'
    # b1 = ')'
    #
    # # 使用 f-string 构建正则表达式
    # pattern = f'{re.escape(a)}.*?{re.escape(b)}'
    #
    # # 替换为 ''
    # clean_text = re.sub(pattern, '', text)
    import re

    a_list = ['（此处', '(此处', '（文中', '(文中']
    b_list = ['）', ')']

    # 对特殊字符进行转义并构建选择模式
    a_escaped = [re.escape(a) for a in a_list]
    b_escaped = [re.escape(b) for b in b_list]

    a_pattern = '(?:' + '|'.join(a_escaped) + ')'
    b_pattern = '(?:' + '|'.join(b_escaped) + ')'

    pattern = f'{a_pattern}.*?{b_pattern}'
    clean_text = re.sub(pattern, '', text)
    return clean_text

def clean_email(text: str) -> str:
    # 邮箱匹配正则
    pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]*'

    # 替换为空字符串
    clean_text = re.sub(pattern, '', text)

    return clean_text

def  clean_empty(text: str) -> str:
    # 分割每行，去除首尾空格，并过滤空行
    clean_lines = [line.strip() for line in text.splitlines() if line.strip()]

    # 重新拼接成字符串
    clean_text = "\n".join(clean_lines)
    return clean_text



def main(text: str, filename: str = None):
    global replacement_rules
    global url_manager

    text = clean_style_tags(text)
    text = clean_tags(text)
    text = clean_line_by_keywords(text)
    text = clean_keywords(text)
    text = clean_between_Parentheses(text)
    text = clean_email(text)
    text = clean_url(text)
    text = clean_empty(text)
    text = replace_words_in_text(text, replacement_rules)


    text = wrap_unlabeled_lines(text)
    text = insert_picture_url(text, url_manager)
    # 如果有文件名，在开头插入标题
    if filename:
        text = insert_txt_name(text, filename)
    
    return text
def main_cli():
    """命令行交互界面"""
    print("=== TXT文件批量处理工具 ===")
    print("功能：处理指定目录下所有txt文件，并保存到新目录")
    print()
    
    while True:
        source_dir = input("请输入要处理的目录路径（输入 'quit' 退出）：").strip()
        
        if source_dir.lower() == 'quit':
            print("退出程序")
            break
        
        if not source_dir:
            print("请输入有效的目录路径")
            continue
        
        # 处理路径中的引号
        source_dir = source_dir.strip('"\'')
        
        process_txt_files(source_dir)
        print("\n" + "="*50 + "\n")

if __name__ == "__main__":
    sample = """
这是正常的一行
<h2>标题</h2>
请自行补充一些内容
<h3>这一行是有效的1</h3>
故意留白111
11111好的，没问题11111111111
dfsdff???dfssf?dsfhshdf 

(此处)
()(此处有狗）

请联系我：abc123@example.com 或者 xyz456@domain.cn，谢谢！

   这是 

  
百度baidu.com1111
     

    """
    excel_file = "词组替换.xlsx"
    if os.path.exists(excel_file):
        print("正在加载替换规则...")
        replacement_rules = load_replacement_rules(excel_file)
        # print(f"成功加载 {len(replacement_rules)} 条替换规则")
        #     # 显示替换规则
        # for old, new in list(replacement_rules.items())[:1000]:  # 只显示前1000条
        #     print(f"'{old}' -> '{new}'")

    # # 使用read()一次性读取整个文件
    # with open('picture_url.txt', 'r') as f:
    #     content = f.read()
    #     print(content)

    # 初始化URL管理器
    url_manager = URLManager()

    main_cli()
    # text = main(sample)
    # print(text)
    # text = main(sample)
    # print(text)

