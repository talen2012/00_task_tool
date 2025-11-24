import pandas as pd

# @Time    : 2025/09/09/10:43
# @Author  : talen
# @File    : summary_category_name.py

"""
对中标情况进行分析，仅24年中标，仅25年中标，24年和25年均中标
"""

def generate_company_name_set(data_frame):
    """
    生成公司名称集合
    """
    # 筛选出“市”列内容为“省管”、“西安”、“西咸”的行
    print("筛选省管、西安、西咸的中标公司...")
    filtered_df = data_frame[data_frame["市"].isin(["省管", "西安", "西咸"])] # isin方法用于筛选多值
    # 从filtered_df中提取“中标公司”列
    company_series = filtered_df["中标公司"].dropna() # 去除空值
    # 使用集合去重
    company_set = set(company_series)
    print(f"中标公司共 {len(company_set)} 家")
    return company_set


if __name__ == "__main__":
    bid_info_2024_file = ".\\04_win_bid_company_analysis\\24年中标截止12月底.xlsx"
    bid_info_2025_file = ".\\04_win_bid_company_analysis\\25年中标-截止8月.xlsx"
    bid_info_chanshu_2025_file = ".\\04_win_bid_company_analysis\\陕西招投标数据-数说123-1016.xlsx"

    result_file = ".\\04_win_bid_company_analysis\\中标情况分析_24年25年.xlsx"
    # 读取Excel文件
    try:
        df_2024 = pd.read_excel(bid_info_2024_file, sheet_name="全量中标")
    except Exception as e:
        print(f"读取 {bid_info_2024_file} 时发生错误: {e}")

    try:
        df_2025 = pd.read_excel(bid_info_2025_file, sheet_name="全量中标清单-用于统计中标份额")
    except Exception as e:
        print(f"读取 {bid_info_2025_file} 时发生错误: {e}")

    try:
        df_2025_chanshu = pd.read_excel(bid_info_chanshu_2025_file, sheet_name="中标项目")
    except Exception as e:
        print(f"读取 {bid_info_chanshu_2025_file} 时发生错误: {e}")

    # 生成公司名称集合
    print("2024年中标公司名称集合:")
    company_name_set_2024 = generate_company_name_set(df_2024)
    print("2025年中标公司名称集合:")
    company_name_set_2025_bid = generate_company_name_set(df_2025)
    company_name_set_2025_chanshu = generate_company_name_set(df_2025_chanshu)
    # 合并2025年两个数据源的公司名称集合
    company_name_set_2025 = company_name_set_2025_bid.union(company_name_set_2025_chanshu)
    print(f"2025年中标公司名称集合（合并后）共 {len(company_name_set_2025)} 家")

    # 计算仅24年中标的公司
    only_2024 = company_name_set_2024 - company_name_set_2025
    print(f"仅24年中标公司数量: {len(only_2024)}")
    
    # 计算仅25年中标的公司
    only_2025 = company_name_set_2025 - company_name_set_2024
    print(f"仅25年中标公司数量: {len(only_2025)}")

    # 计算24年和25年均中标的公司
    both_24_and_25 = company_name_set_2024.intersection(company_name_set_2025)
    print(f"24年和25年均中标公司数量: {len(both_24_and_25)}")

    # 保存结果到Excel文件，第一列公司名称，第二列中标年份
    result_rows = []
    for company in only_2024:
        result_rows.append({"公司名称": company, "中标年份": "仅2024"})
    for company in only_2025:
        result_rows.append({"公司名称": company, "中标年份": "仅2025"})
    for company in both_24_and_25:
        result_rows.append({"公司名称": company, "中标年份": "2024和2025"})

    result_df = pd.DataFrame(result_rows, columns=["公司名称", "中标年份"])
    result_df.to_excel(result_file, index=False)
    print(f"结果已保存到 {result_file}")