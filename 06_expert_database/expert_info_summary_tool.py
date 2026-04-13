# @Time    : 2026/01/14/15:44
# @Author  : talen
# @File    : expert_info_summary_tool.py

import pandas as pd

def summarize_expert_info(df: pd.DataFrame) -> pd.DataFrame:
    """
    从文件“西安专家.xls”中提取专家职称水平、专家等级、专业特长等信息，生成专家信息汇总字典。
    """
    expert_info_map = {}

    # 数据预处理
    # 使用电话作为索引，
    df = df.dropna(subset=["电话"])  # 去除“电话”列为空的行
    
    for _, row in df.iterrows():
        phone_num  = str(row["电话"])
        expert_level = str(row["专家等级"]) if pd.notna(row["专家等级"]) else ""

        if phone_num not in expert_info_map:
            expert_info_map[phone_num] = {"姓名": "/", "地市": "/", "部门": "/", "专家等级": "/", "产数认证": "/", "职称水平": "/", "技术骨干": "/", "学历": "/", "专业特长": "/", "来源": "/"}

        expert_info_map[phone_num]["姓名"] = str(row["姓名"]) if pd.notna(row["姓名"]) else "/"
        # 地市
        expert_info_map[phone_num]["地市"] = str(row["地市"]) if pd.notna(row["地市"]) else "/"
        # 部门
        expert_info_map[phone_num]["部门"] = str(row["部门"]) if pd.notna(row["部门"]) else "/"
        # 姓名

        # 专家等级
        if "四级" in expert_level:
            expert_info_map[phone_num]["专家等级"] = "四级"
        elif "三级" in expert_level:
            expert_info_map[phone_num]["专家等级"] = "三级"
        elif "二级" in expert_level:
            expert_info_map[phone_num]["专家等级"] = "二级"
        elif "一级" in expert_level:
            expert_info_map[phone_num]["专家等级"] = "一级"

        # 产数认证
        if "L4" in expert_level:
            expert_info_map[phone_num]["产数认证"] = "L4"
        elif "L3" in expert_level:
            expert_info_map[phone_num]["产数认证"] = "L3"
        elif "L2" in expert_level:
            expert_info_map[phone_num]["产数认证"] = "L2"
        elif "L1" in expert_level:
            expert_info_map[phone_num]["产数认证"] = "L1"

        # 职称水平
        if "获得资格认证" in expert_level:
            expert_info_map[phone_num]["职称水平"] = "获得职称"

        # 高校毕业生
        if "高校毕业生" in expert_level:
            expert_info_map[phone_num]["学历"] = "高校毕业生"

        # 技术骨干
        if "技术骨干" in expert_level:
            expert_info_map[phone_num]["技术骨干"] = "技术骨干"

        # 专业特长
        if expert_info_map[phone_num]["专业特长"] == "/" :
            expert_info_map[phone_num]["专业特长"] = str(row["专业特长"]) if pd.notna(row["专业特长"]) else "/"
        else:
            expert_info_map[phone_num]["专业特长"] += "/" + str(row["专业特长"]) if pd.notna(row["专业特长"]) else ""

        # 来源
        expert_info_map[phone_num]["来源"] = "西安专家.xls"

    # 将字典转换为DataFrame
    summary_df = pd.DataFrame.from_dict(expert_info_map, orient='index') # orient='index'表示字典的键作为行索引
    summary_df.index.name = '电话'
    summary_df = summary_df.reset_index() # 将索引转换为列

    return summary_df

def filter_chan_shu_dui_wu(df: pd.DataFrame) -> pd.DataFrame:
    """
    从文件“人力视图产数队伍明细”中筛选L2以上人员，及专家信息
    """
    # 将“产数工程师级别”、“研发工程师级别”、“云网工程师级别”列拼接成一列
    df["认证"] = df["产数工程师级别"].fillna("") + "/" + df["研发工程师级别"].fillna("") + "/" + df["云网工程师级别"].fillna("")+ "/" + df["专家级别"].fillna("")
    # 筛选包含L2、L3、L4或专家的行
    filtered_df = df[df["认证"].str.contains("L2|L3|L4|专家", na=False)] # str属性返回字符串方法的集合, na=False表示忽略NaN值
    filtered_df = filtered_df.drop(columns=["认证"])  # 删除辅助列
    return filtered_df

if __name__ == "__main__":
    input_file = ".\\06_expert_database\\西安专家.xlsx"
    output_file = ".\\06_expert_database\\专家信息汇总.xlsx"

    # # 读取Excel文件
    # try:
    #     df_expert = pd.read_excel(input_file, sheet_name="导出专家")
    # except Exception as e:
    #     print(f"读取 {input_file} 时发生错误: {e}")
    #     exit(1)

    # # 提取专家信息汇总
    # summary_df = summarize_expert_info(df_expert)

    # # 保存结果到Excel文件
    # try:
    #     summary_df.to_excel(output_file, index=False)
    #     print(f"专家信息汇总已保存到 {output_file}")
    # except Exception as e:
    #     print(f"保存到 {output_file} 时发生错误: {e}")

    # 筛选产数队伍中的L2及以上人员
    try:
        df_chanshu = pd.read_excel(".\\06_expert_database\\人力视图产数队伍明细-西安.xlsx", sheet_name="Sheet1", skiprows=1)
    except Exception as e:
        print(f"读取 人人力视图产数队伍明细-西安.xlsx 时发生错误: {e}")
        exit(1)

    df_filtered_chanshu = filter_chan_shu_dui_wu(df_chanshu) 
    df_filtered_chanshu.to_excel(".\\06_expert_database\\产数队伍L2及以上人员.xlsx", index=False) # index=False表示不保存行索引
    print(f"产数队伍L2及以上人员已保存到 .\\06_expert_database\\产数队伍L2及以上人员.xlsx")