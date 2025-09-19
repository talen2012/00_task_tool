import pandas as pd

# @Time    : 2025/09/09/10:43
# @Author  : talen
# @File    : summary_category_name.py

"""
记录一个一级分类，及其二级分类、三级分类，每个占据一个单元格
其中三级分类记录时要明确其二级分类，如组织部:智慧干训；
"""

def generate_category_name_sets(file_path):
    """
    读取指定Excel文件的“生态基础表明细-能力”sheet
    生成一级分类的集合、
    二级分类的字典，键为一级分类，值为该一级分类下的二级分类集合
    三级分类的字典，键为(一级分类, 二级分类)元组，值为该二级分类下的三级分类集合
    """
    # Open the Excel file and read the "生态基础表明细-能力" sheet
    try:
        df = pd.read_excel(file_path, sheet_name="生态基础表明细-能力", header=0) # header=0表示表头在第一行（索引从0开始）
    except Exception as e:
        print(f"Error reading file: {e}")
        return

    # 初始化分类集合和字典
    first_level_set = set()
    second_level_dict = {}
    third_level_dict = {}

    # 获得一级分类、二级分类、三级分类
    category_1_2_3_df = df.loc[:, ["一级分类（行业）","二级分类（子行业）","三级分类（领域）", "处理类别（保留、删减、新增）"]] #

    # 去除空行
    category_1_2_3_df = category_1_2_3_df.dropna(how="all")

    for _, row in category_1_2_3_df.iterrows():
        first_level = row.iloc[0]
        second_level = row.iloc[1]
        third_level = row.iloc[2]
        handle_method = row.iloc[3]
        if handle_method == "删减":
            continue
        if pd.isna(first_level) or first_level == "/":
            continue
        first_level_set.add(first_level)

        if pd.isna(second_level) or second_level == "/":
            continue
        # setdefault,如果键不存在，则添加键并设置默认值，否则返回现有值
        second_level_dict.setdefault(first_level, set()).add(second_level)

        if pd.isna(third_level) or third_level == "/":
            continue
        third_level_dict.setdefault((first_level, second_level), set()).add(third_level)

    print(f"一级分类共 {len(first_level_set)} 个")
    for key, value in second_level_dict.items():
        print(f"一级分类 {key} 有二级分类 {len(value)} 个")
    for key, value in third_level_dict.items():
        print(f"一级分类 {key[0]} 二级分类 {key[1]} 下有三级分类 {len(value)} 个")
    print("")

    return first_level_set, second_level_dict, third_level_dict

def compare_summary_from_excel(file_path, first_level_set, second_level_dict, third_level_dict):
    """
    文件中第一个sheet是已经汇总好的一二三级名称
    读取后对比从明细表里生成的内容、输出差异
    """
    df = pd.read_excel(file_path, sheet_name="生态基础表（生态能力图谱+集团行业+标包）",header=0) # 这个表里虽然表头在第一行，但占了两行
    category_1_2_3_summary_df = df.loc[:, ["一级分类","二级分类","三级分类"]]
    category_1_2_3_summary_df = category_1_2_3_summary_df.dropna(how="all")
    first_level_set_in_summary = set()
    second_level_dict_in_summary = {}
    third_level_dict_in_summary = {}

    # iterrows()返回(index, Series), 支持用列名访问，但速度慢、开销大，不适用大数据量遍历
    # itertuples返回一个namedtuple, 用.Index，.列名属性（列名中有空格、特殊字符会被下划线替换），列名是中文带括号还挺麻烦的
    for row in category_1_2_3_summary_df.itertuples(): # row的数据类型是Series
        first_level_text = row.一级分类 # 格式类似 "政务行业、政法公安行业"
        second_level_text= row.二级分类 # 格式类似 "组织部、工会"
        third_level_text = row.三级分类 # 格式类似 "组织部:智慧干训、智慧党建；"
        # 处理一级分类
        first_level_set_in_summary.add(first_level_text)
        # 处理二级分类
        second_level_dict_in_summary.setdefault(first_level_text,set()).update(str(second_level_text).split("、")) # update用于集合的合并
        # 处理三级分类
        # 先用；将文本分割为 二级分类：三级分类、三级分类 注意末尾也可能有；
        # split返回一个列表，跳过空元素、提醒格式不对
        for item in str(third_level_text).split("；"):
            if not item or item == "/":
                continue
            if "：" not in item:
                print(f"三级分类 {item} 书写格式有误，缺少\"：\"！")
                continue
            if len(item.split("：")) != 2:
                print(print(f"三级分类 {item} 书写格式有误，太多\"：\"！"))
                continue
            second_part, third_part = item.split("：")
            if not second_part or not third_part:
                print(f"三级分类 {item} 书写格式有误，缺少二级或三级！")
                continue
            third_level_dict_in_summary.setdefault((first_level_text, second_part), set()).update(str(third_part).split("、"))
    # 求两种方式获得各级分类名称汇总的差集、如有则输出
    # 一级分类
    diff_first_from_generated_to_summary = first_level_set - first_level_set_in_summary
    diff_first_from_summary_to_generated = first_level_set_in_summary - first_level_set
    if diff_first_from_generated_to_summary:
        print(f"\n明细表生成的分类中，多出的一级分类有：{diff_first_from_generated_to_summary}")
    if diff_first_from_summary_to_generated:
        print(f"\n首页已汇总的分类中，多出的一级分类有：{diff_first_from_summary_to_generated}")
    # 二级分类
    for key, _ in second_level_dict.items():
        print(f"\n明细表生成的一级分类 {key} 中：")
        if key not in second_level_dict_in_summary:
            print(f"\n明细表生成的一级分类 {key} 中：")
            print("该分类在首页汇总表中不存在！")
            continue
        
        diff_second_from_generated_to_summary = second_level_dict[key] - second_level_dict_in_summary[key]
        diff_second_from_summary_to_generated = second_level_dict_in_summary[key] - second_level_dict[key]
        if diff_second_from_generated_to_summary:
            print(f"\n明细表生成的一级分类 {key} 中：")
            print(f"明细表生成的分类，多出的二级分类有：{diff_second_from_generated_to_summary}")
        if diff_second_from_summary_to_generated:
            print(f"\n明细表生成的一级分类 {key} 中：")
            print(f"首页已汇总的分类，多出的二级分类有：{diff_second_from_summary_to_generated}")
     # 三级分类
    for key, _ in third_level_dict.items():
        if key not in third_level_dict_in_summary:
            print(f"\n明细表生成的一级分类 {key[0]} 二级分类 {key[1]} 中：")
            print("该分类在首页汇总表中不存在")
            continue
        diff_third_from_generated_to_summary = third_level_dict[key] - third_level_dict_in_summary[key]
        if diff_third_from_generated_to_summary:
            print(f"\n明细表生成的一级分类 {key[0]} 二级分类 {key[1]} 中：")
            print(f"明细表生成的分类，多出来的三级分类有：{diff_third_from_generated_to_summary}")
        diff_third_from_summary_to_generated = third_level_dict_in_summary[key] - third_level_dict[key]
        if diff_third_from_summary_to_generated:
            print(f"\n明细表生成的一级分类 {key[0]} 二级分类 {key[1]} 中：")
            print(f"首页已汇总的分类，多出来的三级分类有：{diff_third_from_summary_to_generated}")

    return 

if __name__ == "__main__":
    file_path = "03_category_name\\合作生态能力图谱-0910_test.xlsx"
    first_level_set, second_level_dict, third_level_dict = generate_category_name_sets(file_path)
    compare_summary_from_excel(file_path, first_level_set, second_level_dict, third_level_dict)