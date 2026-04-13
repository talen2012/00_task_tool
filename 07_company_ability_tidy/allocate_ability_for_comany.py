import pandas as pd

# @Time      : 2026/02/02/10:20
# @Author    : talen
# @File      : allocate_ability_for_company.py

def stat_company_ability(ability_file, company_file):
    '''
    四级能力分类表格每一行后半部分都附有相应的公司名称
    提取公司表格中的公司推荐行业、生态名称、生态来源、能力方案（解决方案及产品等）、企业简介（简要描述行业地位或市场规模等）
    给每个公司按四级能力分类表格分配能力，不同能力分行，公司信息复制

    '''
    try:
        df_ability = pd.read_excel(ability_file, sheet_name="能力类型视图清单 行业能力0202 (2)", header=1)
        df_company = pd.read_excel(company_file, sheet_name="行业上报")
    except Exception as e:
        print(f"错误：读取Excel文件失败 - {str(e)}")
        return
    
    df_company["上报时间"] = pd.to_datetime(df_company["上报时间"], errors='coerce').dt.strftime('%y/%m/%d')

    # 公司能力字典，键为公司名称，值为一个列表，列表元素是该公司的各项能力
    companys_ability_map = {}
    # 四级能力分类表格，从“公司信息”列开始为公司名称
    company_start_col = df_ability.columns.get_loc("公司信息")
    for idx, row in df_ability.iterrows():
        for col in df_ability.columns[company_start_col:]:
            company_name = str(row[col]).strip()
            if not pd.isna(company_name) and company_name not in ["", "无", "-", "/", "未知"]:
                ability_info = {
                    "序号": row["序号"],
                    "一级能力": row["一级能力"],
                    "二级能力": row["二级能力"],
                    "三级能力": row["三级能力"],
                    "四级能力": row["四级能力"],
                    "技术要求": row["技术要求"]
                }
                if company_name not in companys_ability_map:
                    companys_ability_map[company_name] = []
                companys_ability_map[company_name].append(ability_info)

    # 公司信息字典，键为“推荐行业-生态名称”，值为公司信息
    # 因为同一个公司可能有多个推荐行业，所以采用这种“组合键”
    companys_info_map = {}
    for _, row in df_company.iterrows():
        key = f"{row['推荐行业']}-{row['生态名称']}"
        if key not in companys_info_map:
            companys_info_map[key] = {
                "推荐行业": row["推荐行业"],
                "生态名称": str(row["生态名称"]).strip(),
                "生态来源": row["生态来源"],
                "能力方案": row["能力方案（解决方案及产品等）"],
                "企业简介": row["企业简介（简要描述行业地位或市场规模等）"],
                "上报时间": row["上报时间"],
                "备注": row["备注"]
            }

    # 生成汇总表格
    # 把公司信息和公司能力结合起来，左半部分是公司信息，右半部分是公司能力
    # 由于一个公司有多项能力，所以需要多行表示
    output_rows = []
    for _, company_info in companys_info_map.items():
        # 查找公司能力
        company_name = company_info["生态名称"]
        if pd.notna(company_name):
            if company_name in companys_ability_map.keys():
                for ability in companys_ability_map[company_name]:
                    output_row = company_info.copy()
                    output_row.update(ability) # 合并两个字典，为什么不直接+，因为字典不支持+
                    output_rows.append(output_row)
            else:
                output_row = company_info.copy()
                output_rows.append(output_row)
                

    # 生成DataFrame并保存为Excel
    df_output = pd.DataFrame(output_rows)
    output_file = ".\\07_company_ability_tidy\\公司能力分配结果.xlsx"
    try:
        df_output.to_excel(output_file, index=False)
        print(f"公司能力分配结果已保存到 {output_file}")
    except Exception as e:
        print(f"保存Excel文件时出错：{e}")

if __name__ == "__main__":
    ability_file = ".\\07_company_ability_tidy\\能力类型视图清单表 0202v2.xlsx"
    company_file = ".\\07_company_ability_tidy\\生态合作伙伴推荐表汇总--全量公司.xlsx"
    stat_company_ability(ability_file, company_file)



