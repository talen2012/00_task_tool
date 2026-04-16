import tkinter as tk
import pandas as pd
import openpyxl
import os
import re

# @Time     : 2026/04/13/16:00
# @Author   : talen
# @File     : generate_biz_report_weekly.py

#TODO: 界面设计

class BizReportApp:
    def __init__(self):
        # 变量
        self.map_required_cols = {"清单底表", "通报模板"}
        # 目前百万以上商机清单，预计签约时间只到月，格式是"202604", 需要用户在"是否目标时间段"列手动标记是否在统计的目标时间段内，否则无法统计到段时间维度
        self.above_million_required_cols = {"集团商机编码", "单元", "行业标识", "预计签约月", "是否目标时间段","预估签约金额(万元)"}
        self.below_million_required_cols = {"集团商机编码", "单元", "行业标识", "预计签约日期", "商机预测金额(万元)"}
        self.contrat_status_required_cols = {"单元（新）", "行业（新）", "签约日期", "省口径签约额", "商机编码"}
        self.unit_industry_map_filepath = None
        self.report_template_filepath = None
        self.biz_above_million_filepath = "09_business_opportunity_weekly_report\\0410_百万以上.xlsx"
        self.biz_below_million_filepath = "09_business_opportunity_weekly_report\\0411_百万以下.xlsx"
        self.contract_status_filepath = "09_business_opportunity_weekly_report\\0410_签约情况.xlsx"
        self.start_date = None
        self.end_date = None
        self.strict_mode = None # 签约情况中的合同，必须有对应商机编码，并且该商机编码存在于商机清单中，否则不计入统计
        self.loose_mode = None # 不考虑合同与商机编码的对应关系，直接统计签约情况中的合同数据

        # 名称映射表及通报模板默认文件路径，如果存在
        default_unit_industry_map_file_path = os.path.join(os.path.dirname(__file__), "reference", "单元_行业名称映射表.xlsx")
        default_report_template_file_path = os.path.join(os.path.dirname(__file__), "reference", "通报模板.xlsx")
        if os.path.exists(default_unit_industry_map_file_path):
            self.unit_industry_map_filepath = default_unit_industry_map_file_path
        if os.path.exists(default_report_template_file_path):
            self.report_template_filepath = default_report_template_file_path

    def _open_source_tabel(self):
        """
        使用pandas打开底表文件, 并检查是否存在指定列
        @param biz_above_million_file: 100万以上商机清单文件路径
        @param biz_below_million_file: 100万以下商机清单文件路径
        @param contract_status_file: 合同实际签约情况文件路径
        @param unit_industry_map_file: 单元_行业名称映射表
        """
        # 1. 读取单元/行业名称映射文件
        print("正在加载单元/行业名称映射表...")
        try:
            df_unit_map = pd.read_excel(self.unit_industry_map_filepath, sheet_name="单元", header=1)
            df_industry_map = pd.read_excel(self.unit_industry_map_filepath, sheet_name="行业", header=1)
            # 如果"清单底表"和"通报模板"两列不存在，就抛出Key error
            unit_missing_cols = self.map_required_cols - set(df_unit_map.columns)
            industry_missing_cols = self.map_required_cols - set(df_industry_map.columns)
            if unit_missing_cols:
                raise KeyError(f"Sheet【单元】缺少列: {list(unit_missing_cols)}")
            if industry_missing_cols:
                raise KeyError(f"Sheet【行业】缺少列: {list(industry_missing_cols)}")
        except KeyError as ve:
            print(f"错误:单元/行业名称映射文件格式不正确\n{str(ve)}")
            return None
        except Exception as e:
            print(f"错误:单元/行业名称映射文件格式不正确\n{str(e)}")
            return None
        print("成功加载单元/行业名称映射表！")
        
        # 2. 读取百万以上商机清单
        print("正在读取百万以上商机清单...")
        try:
            df_biz_above_million = pd.read_excel(self.biz_above_million_filepath, sheet_name="百万以上")
            above_million_missing_cols = self.above_million_required_cols - set(df_biz_above_million.columns)
            if above_million_missing_cols:
                raise KeyError(f"Sheet【百万以上】缺少列: {list(above_million_missing_cols)}")
        except KeyError as ve:
            print(f"错误:百万以上商机清单文件格式不正确\n{str(ve)}")
            return None
        except Exception as e:
            print(f"错误:百万以上商机清单文件格式不正确\n{str(e)}")
            return None
        print("成功读取百万以上商机清单！")

        # 3. 读取百万以下商机清单
        print("正在读取百万以下商机清单...")
        try:
            df_biz_below_million = pd.read_excel(self.biz_below_million_filepath, sheet_name="百万以下")
            
            below_million_missing_cols = self.below_million_required_cols - set(df_biz_below_million.columns)
            if below_million_missing_cols:
                raise KeyError(f"Sheet【百万以下】缺少列: {list(below_million_missing_cols)}") 
        except KeyError as ve:
            print(f"错误:百万以下商机清单文件格式不正确\n{str(ve)}")
            return None
        except Exception as e:
            print(f"错误:百万以下商机清单文件格式不正确\n{str(e)}")
            return None
        print("成功读取百万以下商机清单！")

        # 4. 读取商机签约情况
        try:
            df_contrat_status = pd.read_excel(self.contract_status_filepath, sheet_name="签约清单")
            
            contrat_status_missing_cols = self.contrat_status_required_cols - set(df_contrat_status.columns)
            if contrat_status_missing_cols:
                raise KeyError(f"Sheet【签约清单】缺少列: {list(contrat_status_missing_cols)}") 
        except KeyError as ve:
            print(f"错误:商机签约情况文件格式不正确\n{str(ve)}")
            return None
        except Exception as e:
            print(f"错误:商机签约情况文件格式不正确\n{str(e)}")
            return None
        
        # 5. 成功返回5个DataFrame
        return df_biz_above_million, df_biz_below_million, df_contrat_status, df_unit_map, df_industry_map

    def _summarize_by_unit_and_industry(self, df_biz_above_million, df_biz_below_million, df_contract_status, df_unit_map, df_industry_map): 
        """
        根据100万以上/以下商机清单，及合同签约情况生成每周商机进展报告
        @param df_biz_above_million: 100万以上商机清单
        @param df_biz_below_million: 100万以下商机清单
        @param df_contract_status: 合同实际签约情况
        @param df_unit_map: 单元名称映射表
        @param df_industry_map: 行业名称映射表
        """
        # 1. 数据预处理
        # 1.1 筛选必须列
        df_biz_above_million_required = df_biz_above_million[list(self.above_million_required_cols)].copy() # 注意：要新建一个DataFrame副本，避免SettingWithCopyWarning 
        df_biz_below_million_required = df_biz_below_million[list(self.below_million_required_cols)].copy()
        df_contract_status_required = df_contract_status[list(self.contrat_status_required_cols)].copy()
        # 1.2 构建单元名称映射表、行业名称映射表，可能存在多对一映射关系
        ## 注意不要用dict，因为清单底表列存在"高新","南高新"这种相互包含的关键字
        ## 并且本来就是要逐个遍历，所以即使使用dict也不会提高效率
        ## 因此使用list保留顺序, 并根据关键字长度排个序，长的关键字在前边，避免"南高新"被"高新"给匹配走
        unit_map_list = df_unit_map[["清单底表", "通报模板"]].values.tolist() #values是将DataFrame转换成二维numpy数组，tolist()再将其转换成列表
        unit_map_list.sort(key=lambda x: len(x[0]), reverse=True)
        industry_map_list = df_industry_map[["清单底表", "通报模板"]].values.tolist()
        industry_map_list.sort(key=lambda x: len(x[0]), reverse=True)
        # 1.3 根据单元_行业名称映射表，将底表的单元、行业两列映射成与模板中一致
        # 注意：非绝对匹配，只要底表名称包含"清单底表"列的某个键，就可以映射到对应的"通报模板"名称，不成功会标记为"unmatched"
        print("正在映射单元和行业名称...")
        df_biz_above_million_required["单元"] = df_biz_above_million_required["单元"].apply(lambda x: next((v for k, v in unit_map_list if k in str(x)), "unmatched")) # next()函数用于返回第一个满足条件的元素，如果没有满足条件的元素，则返回默认值x（即原值）
        df_biz_below_million_required["单元"] = df_biz_below_million_required["单元"].apply(lambda x: next((v for k, v in unit_map_list if k in str(x)), "unmatched"))
        df_contract_status_required["单元（新）"] = df_contract_status_required["单元（新）"].apply(lambda x: next((v for k, v in unit_map_list if k in str(x)), "unmatched"))

        df_biz_above_million_required["行业标识"] = df_biz_above_million_required["行业标识"].apply(lambda x: next((v for k, v in industry_map_list if k in str(x)), "unmatched"))
        df_biz_below_million_required["行业标识"] = df_biz_below_million_required["行业标识"].apply(lambda x: next((v for k, v in industry_map_list if k in str(x)), "unmatched"))
        df_contract_status_required["行业（新）"] = df_contract_status_required["行业（新）"].apply(lambda x: next((v for k, v in industry_map_list if k in str(x)), "unmatched"))
        print("成功映射单元和行业名称！")

        # 1.4 保存预处理之后的数据到文件，留待后续参考
        def save_prehandled_data(origin_filepath, sheet_name, df):
            file_name = os.path.basename(origin_filepath)
            new_filepath = re.sub(
                r'(\.xlsx|\.xls)$', 
                f"_预处理_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}\\1",
                file_name,
                flags=re.IGNORECASE)
            prehandled_output_dir = os.path.join(os.path.dirname(__file__), "预处理后的数据")
            os.makedirs(prehandled_output_dir, exist_ok=True)
            try:
                df.to_excel(
                    os.path.join(prehandled_output_dir, new_filepath),
                    sheet_name = sheet_name,
                    index = False
                )
            except Exception as e:
                print(f"错误:保存预处理数据失败\n{str(e)}")
            print(f"预处理后的数据文件已保存至: \n{prehandled_output_dir}")
        

        save_prehandled_data(self.biz_above_million_filepath, "百万以上", df_biz_above_million_required)
        save_prehandled_data(self.biz_below_million_filepath, "百万以下", df_biz_below_million_required)
        save_prehandled_data(self.contract_status_filepath, "签约清单", df_contract_status_required)


        # 根据用户输入的结束日期确定统计时间段、统计当月、后续三月
        # current_month = pd.to_datetime(self.end_date).month
        # month_after_1_months = (pd.to_datetime(self.end_date) + pd.DateOffset(months=1)).month
        # month_after_2_months = (pd.to_datetime(self.end_date) + pd.DateOffset(months=2)).month
        # month_after_3_months = (pd.to_datetime(self.end_date) + pd.DateOffset(months=3)).month
        # TODO: 根据单元_行业名称映射表，将底表的单元、行业两列映射成与模板中一致
        # TODO: 统计与保存
        # TODO: 保存签约清单中商机编码为空、商机编码不在商机清单的合同编号，留待进一步处理
        

    def _save_to_report_file(self, df_unit, df_industry):
        """
        将统计结果进报告模板文件里并保存
        @param df_unit 按单元维度统计的商机计划与签约信息
        @param df_industry 按行业维度统计的商机计划与签约信息
        @report_file_path 最终要填写的报告模板文件地址
        @param end_date 报告结束日期
        """
        # 根据用户输入的起止日期确定统计时间段、统计当月、后续三月
        period_start = pd.to_datetime(self.start_date)
        period_end = pd.to_datetime(self.end_date)
        month = pd.to_datetime(self.end_date).strftime("%-m月") # %-表示去掉前导0
        month_after_1_months = period_end + pd.DateOffset(months=1).strftime("%-m月")
        month_after_2_months = period_end + pd.DateOffset(months=2).strftime("%-m月")
        month_after_3_months = period_end + pd.DateOffset(months=3).strftime("%-m月")

    def run_biz_analysis_workflow(self):
        # 1. 读取文件并检查Sheet名、列名
        df_set = self._open_source_tabel()
        if not df_set:
            print("错误", "数据源文件格式不正确！")
            return
        self._summarize_by_unit_and_industry(*df_set)
        # TODO：根据用户选择，进行宽口径、严口径或两种口径的统计 

if __name__ == "__main__":
    app = BizReportApp()
    app.run_biz_analysis_workflow()