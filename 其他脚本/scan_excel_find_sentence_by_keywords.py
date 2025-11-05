import os
import pandas as pd


"""
遍历folder目录下的多个excel，如果有包含keyword的行，打印输出
"""
def find_keyword_in_excels(folder_path, keyword):
    # 遍历目录
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(('.xlsx', '.xls')):
                file_path = os.path.join(root, file)
                try:
                    # 读取所有sheet
                    xls = pd.ExcelFile(file_path)
                    for sheet_name in xls.sheet_names:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)  # 全部读为字符串，避免数值丢失
                        # 查找包含关键词的行
                        mask = df.apply(lambda row: row.astype(str).str.contains(keyword, na=False).any(), axis=1)
                        matched_rows = df[mask]
                        if not matched_rows.empty:
                            # print(f"\n文件: {file_path}, 工作表: {sheet_name}")
                            # print(matched_rows.fillna('').to_string(index=False, header=False))
                            output_str = matched_rows.fillna('').to_string(index=False, header=False)

                            # 打印结果
                            print(output_str)

                            # # 如果包含“有效”，再额外提示
                            # if '有效' in output_str:
                            #     print(kw +'检测到有效对话！！！！！！！！！！！！！！！!!!!!!!!!!!!!!!!!!!!！！！！！！！！！！！！！！！！！！！！')


                except Exception as e:
                    print(f"读取 {file_path} 出错: {e}")

if __name__ == "__main__":
    folder = r"C:\Users\zhang\Desktop\数据筛选软件\每日任务\轨迹沟通记录"  # 修改为你的目录
    keywords  = [
        'm.lzbbb.com'
        , 'm.fzjfh.com'
        , 'm.yc-kw.com'
        , 'm.hdqhkq.com'
        , 'm.fjbdfyjs.com'
        , 'm.fjsbdf120.com'
        , 'm.rdbtz.com'
        , 'm.confirm4task.com'
        , 'm.zzbdf1.com'
        , 'm.shtpfb.com'
        , 'm.ytjsyhjhs.com'
        , 'm.zzxjwz.com'
    ]
    for kw in keywords:
          # 修改为你要查找的关键词
        print(kw)
        find_keyword_in_excels(folder, kw)
        print('--------------------------------------')
    print('结束')