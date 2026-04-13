import tkinter as tk
import pandas as pd
import openpyxl

# @Time     : 2026/04/13/16:00
# @Author   : talen
# @File     : generate_biz_report_weekly.py

def open_source_tabel(biz_above_million_file, df_biz_below_million_file, df_contract_status_file, df_unit_industry_map_file):
    """
    使用pandas打开底表文件，并检查是否存在指定列
    @param biz_above_million_file: 100万以上商机清单文件路径
    @param biz_below_million_file: 100万以下商机清单文件路径
    @param contract_status_file: 合同实际签约情况文件路径
    @param nit_industry_map_file: 单元_行业名称映射表
    """

def generate_weekly_report(df_biz_above_million, df_biz_below_million, df_contract_status, df_unit_industry_map, start_date, end_date): 
    """
    根据100万以上/以下商机清单，及合同签约情况生成每周商机进展报告
    @param df_biz_above_million: 100万以上商机清单
    @param df_biz_below_million: 100万以下商机清单
    @param df_contract_status: 合同实际签约情况
    @param df_unit_industry_map: 单元_行业名称映射表
    @param start_date: 报告起始日期
    @param end_date: 报告结束日期
    """
    # TODO: 根据单元_行业名称映射表，将底表的单元、行业两列映射成与模板中一致
    # TODO：