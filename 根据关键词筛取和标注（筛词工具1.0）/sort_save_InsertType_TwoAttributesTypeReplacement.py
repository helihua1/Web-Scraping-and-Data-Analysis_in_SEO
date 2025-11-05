import re

import pandas as pd
import os

'''
PRIORITY_FILE_PATH 表格格式 :疾病	栏目	素材类型	关键词
SORTED_FILE_PATH 表格格式 :句子
句子中包含PRIORITY_FILE_PATH表格中的关键词，视为比对成功，将句子保存到OUTPUT_FILE_PATH表格的后几列中。未匹配的句子保存到OUTPUT_FILE_PATH表格的第1列。
OUTPUT_FILE_PATH 表格格式 :未匹配的句子	|关键词属性1	关键词属性2	关键词属性3	匹配到的关键词	匹配到的句子

相比sort_save_InsertType.py增加功能：
1，比对时英文统一小写，防止大小写影响比对结果
2，比对句子中是否包含白癜风银屑病相关关键词，加入比对成功的结果
'''


def read_config():
    """读取配置文件中的路径设置"""
    config = {}
    try:
        with open('config.txt', 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    config[key.strip()] = value.strip()
        return config
    except FileNotFoundError:
        print("警告：未找到config.txt配置文件，使用默认路径")
        return {
            'SORTED_FILE_PATH': r"D:\sort\A.xlsx",
            'PRIORITY_FILE_PATH': r"D:\sort\确定可以保存的词.xlsx",
            'OUTPUT_FILE_PATH': r"D:\sort\sort_save_InsertType_TwoAttributesTypeReplacement.xlsx"
        }


def process_excel_files():
    # 从配置文件读取路径设置
    config = read_config()

    # 文件路径设置
    # 先去字典 config 里找键 SORTED_FILE_PATH 的值
    # 如果找到了，就用配置文件里指定的路径
    # 如果没找到（比如 config.txt 没写这个键），就用后面的默认值：r"D:\sort\A.xlsx"
    sorted_file_path = config.get('SORTED_FILE_PATH', r"D:\sort\A.xlsx")
    priority_file_path = config.get('PRIORITY_FILE_PATH', r"D:\sort\确定可以保存的词.xlsx")
    output_file_path = config.get('OUTPUT_FILE_PATH', r"D:\sort\sort_save_InsertType_TwoAttributesTypeReplacement.xlsx")

    print(f"使用配置文件中的路径：")
    print(f"排序文件：{sorted_file_path}")
    print(f"优先词文件：{priority_file_path}")
    print(f"输出文件：{output_file_path}")
    print("-" * 50)

    try:
        print("开始读取excel")
        # 读取sorted.xlsx中的句子（第1列）
        sorted_df = pd.read_excel(sorted_file_path, header=None, usecols=[0])


        sentences = sorted_df[0].tolist()  # 获取第1列所有句子

        # 读取优先聚合.xlsx中的词（第4列和第8列）及其前3列属性
        # 第4列（索引3）的前3列属性为索引0、1、2；第8列（索引7）的前3列属性为索引4、5、6
        priority_df = pd.read_excel(priority_file_path, header=None, usecols=[0, 1, 2, 3, 4, 5, 6, 7])

        # 存储第4列关键词及其属性（前3列）
        priority1_words = []
        priority1_attrs = {}  # 键:关键词, 值:属性列表[attr0, attr1, attr2]
        # 存储第8列关键词及其属性（前3列）
        priority2_words = []
        priority2_attrs = {}  # 键:关键词, 值:属性列表[attr4, attr5, attr6]
        print("开始检索")
        # 提取关键词及其对应的属性
        for idx, row in priority_df.iterrows():
            # 处理第4列（索引3）的关键词及其属性
            word1 = row[3]
            if not pd.isna(word1):
                priority1_words.append(word1)
                # 获取前3列属性（索引0、1、2）
                priority1_attrs[word1] = [row[0], row[1], row[2]]

            # 处理第8列（索引7）的关键词及其属性
            word2 = row[7]
            if not pd.isna(word2):
                priority2_words.append(word2)
                # 获取前3列属性（索引4、5、6）
                priority2_attrs[word2] = [row[4], row[5], row[6]]

        # 用于存储匹配到的句子及对应的关键词和属性
        matched_dict = {}  # 键:关键词, 值:该关键词匹配到的句子列表
        remaining_sentences = []

        # 遍历所有句子，检查是否包含优先聚合词并记录匹配的关键词
        for sentence in sentences:
            if pd.isna(sentence):  # 不对空值进行比对
                remaining_sentences.append(sentence)
                continue

            matched_word = None
            #  先检查第4列优先级高的词
            for word in priority1_words:
                if str(word).lower() in str(sentence).lower():
                    matched_word = word
                    break

            # 如果没有匹配到第4列的词，检查第8列的词
            if matched_word is None:
                for word in priority2_words:
                    if str(word).lower() in str(sentence).lower():
                        matched_word = word
                        break

            # 处理匹配结果
            if matched_word is not None:
                # 将句子添加到对应关键词的列表中
                if matched_word not in matched_dict:
                    matched_dict[matched_word] = []
                matched_dict[matched_word].append(sentence)
            else:
                remaining_sentences.append(sentence)

        # 按关键词优先级顺序聚合句子及相关信息（属性、关键词、句子）
        aggregated_data = []  # 每个元素: [attr1, attr2, attr3, 关键词, 句子]
        # 先处理第4列的关键词（按原顺序）
        for word in priority1_words:
            if word in matched_dict:
                # 获取该关键词对应的属性
                attrs = priority1_attrs[word]
                # 为每个匹配的句子添加属性、关键词和句子
                for sentence in matched_dict[word]:
                    aggregated_data.append([attrs[0], attrs[1], attrs[2], word, sentence])
                del matched_dict[word]  # 避免重复处理

        # 再处理第8列的关键词（按原顺序）
        for word in priority2_words:
            if word in matched_dict:
                # 获取该关键词对应的属性
                attrs = priority2_attrs[word]
                # 为每个匹配的句子添加属性、关键词和句子
                for sentence in matched_dict[word]:
                    aggregated_data.append([attrs[0], attrs[1], attrs[2], word, sentence])
                del matched_dict[word]



        BaiDianFeng = []
        print("开始检索白癜风银屑病相关关键词")
        # 增加功能  
        # 遍历aggregated_data，如果baidianfeng或yinxiebing相关词汇在如果item[4](句子)中，则item[0]改为白癜风或银屑病。
        # 遍历remaining_sentences，如果baidianfeng或yinxiebing相关词汇在如果句子中，则加入aggregated_data:[白癜风或银屑病,'','','',sentence]，否则还留在remaining_sentences中。
        baidianfeng = ["白面风","白癜风","白碘风","白殿风","白斑风","白颠风","白癫风","白点癫风","白屑风","白瘕风","白颠疯","白电风","白痶风","白疒癜风","白癞癜风","白电疯","百点风","百癜疯","白巅风","白垫风","白瘨风","白店风","白巅峰","白殿凤","白典风","白点风","白广癜风","白殿疯","白殿","自殿风","白癜凤","白癫","白癫疯","白疯癫","白癲风","白驳风","白癞风","白点癫疯","白点癫","白瘢风","白癜","白颠","白班癫疯","白癣风","白壂风","百瘕风","白淀风","白癜疯","白班风","白斑疯","白癬风","白班疯","白巅疯","白蚀病","白斑病","白风病","白麻风","白皮风","白斑症","白风癣","白块风","白蚀","白驳","皮肤白斑","白癞","白秃风","白蚀风","色素脱失性白斑病","白班病","色素脱失症","色素脱落症","色素脱失","白头风","白蚀症","白驳症","白电凤病","白佃风","白滇风","白掂风","白癫病","白颠病","白点疯","白癫症","白斑性癫风","白点颠风","白点癜风","白点颠疯","白点癜疯","色素减退斑","晕痣","白点1癫风","白瘾风","白疯颠","白癖风","白点一癫风","白癍风","白风点癫","白点癞风","白点癜","白点颠","白点巅峰","白淀疯","白壂疯","白爹风","白跌风","白瘹风","白跌疯","白叠风","白丁风","白蝶疯","白蝶风","白丁疯","白鼎疯","白顶风","白定疯","白腚风","白风癫","自癜风","白臀风","白痴风","白疒癫疯","白病癫风","白病癫","白颤风","白厂颠风","白颤疯","白瘨疯","白巅病","白巅凤","白典疯","白定风","白鼎风","白风癜","白片风","白天风","白腆风","白广癫疯","白瘕疯","白疯疯病","白痹风","白色癫疯","白厩风","白脸风","晕痔","白颠风","白癫疯","白廒风","白瘢疯","白瘢凤","白瘢冈","白癍病","白癍疯","白壁风","白臂风","白边风","白编风","白鞭风","白鞭疯","白扁风","白变风","白变疯","白变凤","白便风","白遍风","白飙风","白瘭风","白鬓风","炎症色素减退白斑","炎症性白斑","炎症后色素减退白斑","炎症后白斑","炎症色素","色素减退","白㿄风","白搬风","白陛风","白陛疯","白廦风","白避风","白璧风","白边疯","白边峰","白边锋","白边凤","白便疯","白遍疯","白瘪风","白疒殿风","白疒疯","白病风","白病疯","白波风","白玻疯","白博疯","白簸风","白廍风","白藏风","白痴癜风","白痴疯疯","白痴疯","白疵风","白瘯风","白戴风","白戴疯","白瘅风","白得风","白倒风","白癖疯","白登风","白登疯","白滴风","白底疯","白弟风","白帝癫风","白帝癜风","白缔疯","白嗲风","白掂疯","白滇疯","白滇峰","白瘨","白癲病","白碘疯","白典凤","白电癫风","白电癫","白电凤","白甸疯","白店枫","白店疯","白店凤","白垫疯","白叠疯","白碟风","白顶疯","白嵿疯","白订疯","白段风","白额风","白风颠","白风殿","白疯病","白疯巅","白疯淀","白疯殿","白疯癜","白疯风","白峰巅","白广癫风","白广癫冈","白广殿风","白广癜","白户癫疯","白痪风","白癀风","白毁风","白屐风","白箕风","白见风","白见疯","白建风","白贱风","白剑风","白健风","白厩疯","白就风","白廄风","白厥风","白连风","白利风","白连疯","白连凤","白瘤风","白癃风","白瘰风","白屡疯","白履风","白瘼风","白瘼疯","白内风","白劈风","白譬风","白偏风","白偏疯","白偏锋","白篇风","白翩风","白谝风","白片疯","白颇风","白颇疯","白瘦风","白厮风","白瘶风","白瘫风","白天癜风","白填风","白填疯","白阗疯","白痶疯","白鲜风","白鲜疯","白廯风","白痫风","白痫疯","白显风","白线风","白消风","白疫癜风","白殷风","白瘀风","白展风","白瘵风","白展殿风","白珍风","白真殿风","百瘢风","百边风","百扁风","百痴风","百颠风","百颠疯","百巅风","百巅峰","百癫风","百癫疯","百癫凤","百典风","百典疯","百点疯","百电风","百店风","百垫风","百淀风","百淀疯","百殿风","百殿疯","百殿凤","百癜风","百丁风","百丁疯","百顶风","百顶疯","百定风","百定疯","自颠风","自瘢风","自变风","自颠疯","自巅风","自巅疯","自巅峰","自癫风","自癫疯","自点风","自殿疯","自电风","自店风","自点疯","自淀风","自癜疯","自丁风","自瘕风","癜风","百巅疯","白带白癜风","白带癜风","白带疯","痞白","痞白症","龙舐","白顶峰","白瘕病","白捡风","白尖锋","白经风","白连峰","白莲风","白莲疯","白散风","白天癫风","白尉风","百变风","白厢风","白旋风","白丹病","白帝风","白底风","白堤风","白的风","白地风","白奌风","白电峰","白淀病","白皮病","斑驳病","白厨风","白癣症","白虎风","白痁疯","白巅","白颊疯","白驳病","白廯","白面疯"]
        yinxiebing = ["银屑病","牛皮皮癣","牛皮肤癣","牛皮癣","牛批癣","牛脾癣","牛疲癣","牛匹癣","牛辟癣","牛皮鲜","牛皮显","牛皮线","牛皮仙","牛皮轩","牛批藓","牛皮癖","牛皮藓","牛皮痟","牛皮廨","白疕","牛疲藓","牛痞藓","牛癖鲜","牛屁藓","牛藓病","牛屑癣","牛癣图","牛皮广癣","牛皮廯","银屑癣","银消病","银痟病","银削病","银宵病","银血病","牛藓","牛癣","牛银","牛皮病癣","副银病","白银屑","负银屑","副银屑","牛鼻癣","牛痹癣","牛血癣","银光屑","银鳞病","银皮病","银皮屑","银皮癣","银皮症","银屏癣","银钱癣","银悄病","银翘病","银翘藓","银鞘病","银翘癣","银锡病","银消癣","银消症","银绡癣","银硝藓","银销癣","银霄病","银肖病","银屑斑","银屑点","银屑风","银屑疾","银屑甲","银屑鲜","银屑廯","银屑藓","银屑屑","银屑血","银屑炎","银屑症","银癣病","银雪癣","银血癣","银血症","长银屑","滴状银屑","点滴银病","点滴银屑","点状银屑","副银屑症","牛皮斑癣","牛皮恶癣","牛皮康癣","牛皮皮藓","牛皮头癣","牛皮顽癣","牛皮小癣","牛皮屑癣","牛皮选癣","牛皮血癣","牛皮一癣","牛皮银屑","牛皮之癣","牛钱癣病","银病","银肤癣","银屑牛癣","银屑皮癣","牛皮有癣","牛皮初期癣","银皮肤癣","银皮肤病","银宵湿","银币癣","银销病","银肖疾","银血屑","银癣","牛皮癬","牛皮屑","牛皮病","牛皮消","牛斑癣","牛股藓","牛肉癣","银宵癣","银皮藓","银俏病","银头皮屑","银藓","银硝病","银线病","银癣苪","银削","银雪病","银元癣","白银癣","牛反癣","牛皮选","牛皮先","牛皮血","牛钱藓","牛屁癣","俞银皮","银鳞屑","银泻病","牛皮好癣","牛皮荃","牛皮想","牛皮炎","牛气癣","牛有癣","银前癣","牛皮蘚","午皮癣","牛皮瘢","牛皮m癣","牛皮痒","牜皮癣","银宵疯","牛皮顽屑","牛皮皮肤癣","银消痛","银嘱病","牛皮虫","牛皮的癣","牛皮个癣","银綃病","银肖","银消","皮廯","皮藓","皮屑","皮癣","屁藓","屁癣","生藓","生癣","湿藓","湿疹","手廯","手藓","手癣","体藓","体屑","体癣","头鲜","头廯","头癣","腿藓","腿癣","癣疹","长廯","长癣","掌脓包","掌脓疱","掌拓病","掌跖病","掌跖脓","掌跖症","扁平苔藓","扁平苔癣","红藓","红癣","红皮病","红皮型","副银屑病","银屑病甲"]
        # 功能1：处理aggregated_data中的数据
        for item in aggregated_data:
            sentence = item[4]  # 获取句子
            # 添加 NaN 检查
            if pd.isna(sentence):
                continue
            
            # 检查是否包含白癜风关键词
            if any(keyword in sentence for keyword in baidianfeng):
                item[0] = "白癜风"  # 将第一项改为"白癜风"
            # 检查是否包含银屑病关键词
            elif any(keyword in sentence for keyword in yinxiebing):
                item[0] = "银屑病"  # 将第一项改为"银屑病"

        # 功能2：处理remaining_sentences中的数据
        new_remaining_sentences = []
        for sentence in remaining_sentences:
            # 添加 NaN 检查
            if pd.isna(sentence):
                new_remaining_sentences.append(sentence)
                continue
            
            # 检查是否包含白癜风关键词
            if any(keyword in sentence for keyword in baidianfeng):
                # 添加到aggregated_data，前三个属性为空
                aggregated_data.append(["白癜风", "", "", "", sentence])
            # 检查是否包含银屑病关键词
            elif any(keyword in sentence for keyword in yinxiebing):
                # 添加到aggregated_data，前三个属性为空
                aggregated_data.append(["银屑病", "", "", "", sentence])
            else:
                # 不包含任何疾病关键词，保留在remaining_sentences中
                new_remaining_sentences.append(sentence)
        print("开始制作excel")
        # 准备构建结果DataFrame的数据
        max_length = max(len(aggregated_data), len(remaining_sentences))
        # 拆分聚合数据到各列
        agg_attr1 = [item[0] for item in aggregated_data] + [None] * (max_length - len(aggregated_data))
        agg_attr2 = [item[1] for item in aggregated_data] + [None] * (max_length - len(aggregated_data))
        agg_attr3 = [item[2] for item in aggregated_data] + [None] * (max_length - len(aggregated_data))
        agg_keywords = [item[3] for item in aggregated_data] + [None] * (max_length - len(aggregated_data))
        agg_sentences = [item[4] for item in aggregated_data] + [None] * (max_length - len(aggregated_data))

        # 创建新的DataFrame，列分布：
        # 0: 未匹配的句子
        # 3-5: 关键词的3列属性
        # 6: 匹配到的关键词
        # 7: 匹配到的句子（原需求中的第7列）

        result_data = {
            0: new_remaining_sentences + [None] * (max_length - len(new_remaining_sentences)),
            1: [None] * max_length,  # 预留列1
            2: [None] * max_length,  # 预留列2
            3: agg_attr1,  # 关键词属性1
            4: agg_attr2,  # 关键词属性2
            5: agg_attr3,  # 关键词属性3
            6: agg_keywords,  # 匹配到的关键词
            7: agg_sentences  # 匹配到的句子
        }

        result_df = pd.DataFrame(result_data)

        # 保存结果到新的Excel文件
        result_df.to_excel(output_file_path, index=False, header=False)
        print(f"处理完成！结果已保存至：{output_file_path}")

    except Exception as e:
        print(f"处理过程中发生错误：{str(e)}")


if __name__ == "__main__":
    print("程序开始执行。。。")
    process_excel_files()
    input("按回车键退出...")