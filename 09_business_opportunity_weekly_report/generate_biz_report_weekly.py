import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk # ttk是tkinter的一个子模块，提供了更现代化的组件，如Combobox、Treeview等
from tkcalendar import DateEntry # 日历选择框
import pandas as pd
import openpyxl
import os
import re
import yaml
from typing import Dict
from datetime import datetime
import platform

# @Time     : 2026/04/13/16:00
# @Author   : talen
# @File     : generate_biz_report_weekly.py
class BizReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("商机报告生成工具")
        self.root.geometry("800x600")
        # ============================= 用户输入变量 =============================
        self.biz_above_million_filepath = tk.StringVar()
        self.biz_below_million_filepath = tk.StringVar()
        self.contract_status_filepath = tk.StringVar()
        self.biz_code_filepath = tk.StringVar()
        self.status_msg = tk.StringVar(value="等待用户输入...")

        self.start_date = tk.StringVar()
        self.end_date = tk.StringVar()
        # 供统计月下拉选择
        self.stat_year = tk.StringVar(value=pd.Timestamp.now().strftime("%Y"))
        self.stat_month = tk.StringVar(value=pd.Timestamp.now().strftime("%m"))
        # 统计口径，每种对应一个统计结果文件，允许同时选择
        self.years_for_select = [str(y) for y in range(2020, 2051)]
        self.months_for_select = [f"{m:02d}" for m in range(1, 13)]

        self.strict_mode = tk.BooleanVar() # 签约情况中的合同，必须有对应商机编码，并且该商机编码存在于商机清单中，否则不计入统计
        self.loose_mode = tk.BooleanVar(value=True) # 不考虑合同与商机编码的对应关系，直接统计签约情况中的合同数据

        self.generate_btn = None # 初始化生成报告按钮，后续动态控制状态

        # ============ reference目录下的文件路径是写死的，不要移动 ==================
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_filepath = os.path.join(self.base_dir, "reference", "config.yaml")
        self.unit_industry_map_filepath = os.path.join(self.base_dir, "reference", "单元_行业名称映射表.xlsx")
        self._report_template_filepath = os.path.join(self.base_dir, "reference", "通报模板.xlsx")

        # =============================== 创建UI ==================================
        self._create_ui()

        # ============== 从配置文件读取目标文件名、sheet名及必须列 ===================
        # 百万以上
        self.above_million_filename_keyword = None
        self.above_million_sheet = None
        self.above_million_required_cols = None
        # 百万以下
        self.below_million_filename_keyword = None
        self.below_million_sheet = None
        self.below_million_required_cols = None
        # 签约情况
        self.contract_status_filename_keyword = None
        self.contract_status_sheet = None
        self.contract_status_required_cols = None
        # 全量商机编码清单
        self.biz_code_filename_keyword = None
        self.biz_code_sheet = None
        self.biz_code_required_cols = None

        # 配置文件加载及状态
        self.config_ok = False
        self._load_config()
        
        # =============== 自动查找程序目录下可能的商机清单及签约情况文件 ==============
        self._auto_find_files()

        # ===================== 绑定输入变化监听事件（实时校验输入）==================
        self._bind_input_trace()
        
        # ========================= 初始检查用户输入完整度 ==========================
        self._check_input_completeness()

    def _auto_find_files(self):
        if not self.config_ok:
            return
        for file_name in os.listdir(self.base_dir):  # listdir仅第一层，不递归，返回字符串列表，os.walk可遍历
            # 只匹配excel文件
            if file_name.startswith(("~$")): continue # 跳过隐藏文件
            if file_name.lower().endswith(('.xlsx', '.xls')):
                if self.above_million_filename_keyword in file_name:
                    self.biz_above_million_filepath.set(os.path.join(self.base_dir, file_name))
                    self._log(f"自动匹配到百万以上商机文件：{file_name}")
                if self.below_million_filename_keyword in file_name:
                    self.biz_below_million_filepath.set(os.path.join(self.base_dir, file_name))
                    self._log(f"自动匹配到百万以下商机文件：{file_name}")
                if self.contract_status_filename_keyword in file_name:
                    self.contract_status_filepath.set(os.path.join(self.base_dir, file_name))
                    self._log(f"自动匹配到签约情况文件：{file_name}")
                if self.biz_code_filename_keyword in file_name:
                    self.biz_code_filepath.set(os.path.join(self.base_dir, file_name))
                    self._log(f"自动匹配到商机编码清单文件：{file_name}")

    def _create_ui(self):
        # ============================= 文件选择区域 =============================
        frame_files = tk.LabelFrame(self.root, text="文件选择", padx=10, pady=10) # 框架内部的内边距
        frame_files.pack(fill="x", padx=10, pady=5) # 框架和父窗口/其它组件的边距
        # 1. 百万以上商机清单
        tk.Label(frame_files, text="百万以上商机清单:").grid(row=0, column=0, sticky='w')
        tk.Entry(frame_files, textvariable=self.biz_above_million_filepath, width=80).grid(row=0, column=1, padx=5)
        tk.Button(frame_files, text="浏览...", command=self._select_biz_above_million_file).grid(row=0, column=2)
        # 2. 百万以下商机清单
        tk.Label(frame_files, text='百万以下商机清单:').grid(row=1, column=0, sticky='w')
        tk.Entry(frame_files, textvariable=self.biz_below_million_filepath, width=80).grid(row=1, column=1, padx=5)
        tk.Button(frame_files, text='浏览...', command=self._select_biz_below_million_file).grid(row=1, column=2)
        # 3. 签约情况表
        tk.Label(frame_files, text='商机签约情况表:').grid(row=2, column=0, sticky='w')
        tk.Entry(frame_files, textvariable=self.contract_status_filepath, width=80).grid(row=2, column=1, padx=5)
        tk.Button(frame_files, text='浏览...', command=self._select_contract_status_file).grid(row=2, column=2)
        # 4.商机编码清单
        tk.Label(frame_files, text='商机编码全量清单:').grid(row=3, column=0, sticky='w')
        tk.Entry(frame_files, textvariable=self.biz_code_filepath, width=80).grid(row=3, column=1, padx=5)
        tk.Button(frame_files, text='浏览...', command=self._select_biz_code_file).grid(row=3, column=2)
     
        # ================================= 设置 =================================
        frame_settings = tk.LabelFrame(self.root, text='设置', padx=5, pady=5)
        frame_settings.pack(fill='x', padx=10, pady=5)
        # 1. 日期选择
        frame_settings_date = tk.LabelFrame(frame_settings, text='日期选择', padx=10, pady=10)
        frame_settings_date.pack(side='left', padx=10, pady=5,anchor='n')
        tk.Label(frame_settings_date, text="开始日期:").grid(row=0, column=0, sticky='w')
        DateEntry(frame_settings_date, textvariable=self.start_date, date_pattern="yyyy-mm-dd", width=14).grid(row=0, column=1, padx=5)
        tk.Label(frame_settings_date, text='结束日期:').grid(row=0, column=2, padx=5)
        DateEntry(frame_settings_date, textvariable=self.end_date, date_pattern="yyyy-mm-dd", width=14).grid(row=0, column=3, padx=5)
        tk.Label(frame_settings_date, text="统计月:").grid(row=0, column=4, padx=5)
        ttk.Combobox(frame_settings_date, textvariable=self.stat_year, values=self.years_for_select, width=6).grid(row=0, column=5)
        ttk.Combobox(frame_settings_date, textvariable=self.stat_month, values=self.months_for_select, width=3).grid(row=0, column=6)

        # 2. 统计口径
        frame_settings_caliber = tk.LabelFrame(frame_settings, text="统计口径", padx=10, pady=10)
        frame_settings_caliber.pack(side='left', padx=10, pady=5,anchor='n')
        ttk.Checkbutton(frame_settings_caliber, text="宽口径", variable=self.loose_mode).grid(row=0, column=0, padx=5)
        ttk.Checkbutton(frame_settings_caliber, text='严口径', variable=self.strict_mode).grid(row=0, column=1, padx=5)

        # =============================== 操作区域 ===============================
        frame_actions = tk.LabelFrame(self.root, padx=10, pady=10)
        frame_actions.pack(fill='x', side='top', padx=10, pady=5)
        # 初始化生成报告按钮（默认置灰）
        self.generate_btn = tk.Button(
            frame_actions,
            text="生成报告",
            command=self.run_biz_analysis_workflow,
            bg="#cccccc",
            fg="#666666",
            font=('Microsoft Yahei', 10, 'bold'),
            state=tk.DISABLED
            )
        self.generate_btn.pack(side='left', padx=5)
        
        tk.Label(frame_actions, textvariable=self.status_msg, fg='blue').pack(side='left', padx=50)

        # =============================== 日志区域 ===============================
        frame_log = tk.LabelFrame(self.root, text='运行日志', padx=10, pady=10)
        frame_log.pack(fill='both', expand=True, padx=10, pady=5)
        self.log_text = scrolledtext.ScrolledText(frame_log, height=20)
        self.log_text.pack(fill="both", expand=True)

    def _bind_input_trace(self):
        """绑定所有输入变量的变化监听（实时校验）"""
        # 监听文件路径的变化
        self.biz_above_million_filepath.trace_add("write", lambda *args: self._check_input_completeness())
        self.biz_below_million_filepath.trace_add("write", lambda *args: self._check_input_completeness())
        self.contract_status_filepath.trace_add("write", lambda *args: self._check_input_completeness())
        self.biz_code_filepath.trace_add("write", lambda *args: self._check_input_completeness())
        # 监听日期、统计月变化
        self.start_date.trace_add("write", lambda *args: self._check_input_completeness())
        self.end_date.trace_add("write", lambda *args: self._check_input_completeness())
        self.stat_year.trace_add("write", lambda *args: self._check_input_completeness())
        self.stat_month.trace_add("write", lambda *args: self._check_input_completeness())
        # 监听统计口径变化
        self.strict_mode.trace_add("write", lambda *args: self._check_input_completeness())
        self.loose_mode.trace_add("write", lambda *args: self._check_input_completeness())

    def _check_input_completeness(self):
        """分类校验，失败弹窗+直接返回，更新状态和按钮样式"""
        def wait_input():
            self.generate_btn.config(bg="#cccccc", fg="#666666", state=tk.DISABLED)
            self.status_msg.set("等待用户输入...")
            self._log("等待用户输入...")
        # ============================= 文件路径校验 =============================
        # 1. 路径非空
        required_filepath_inputs = [
            self.biz_above_million_filepath.get().strip(),
            self.biz_below_million_filepath.get().strip(),
            self.contract_status_filepath.get().strip(),
            self.biz_code_filepath.get().strip()
        ]
        if not all(required_filepath_inputs):
            wait_input()
            return
        # 2. 路径有效
        file_var_map = [
            (self.biz_above_million_filepath, "百万以上商机文件"),
            (self.biz_below_million_filepath, "百万以下商机文件"),
            (self.contract_status_filepath, "签约情况文件"),
            (self.biz_code_filepath, "商机编码清单文件")
        ]
        for var, name in file_var_map:
            path = var.get().strip()
            if not os.path.exists(path):
                self.generate_btn.config(bg="#cccccc", fg="#666666", state=tk.DISABLED)
                messagebox.showerror("路径无效", f"【{name}】不存在或为空！")
                self.status_msg.set("文件路径有误")
                self._log(f"错误: 【{name}】不存在或为空！")
                return
                
        # =============================== 日期校验 ===============================
        # 1. 日期非空且有效
        start = self.start_date.get().strip()
        end = self.end_date.get().strip()
        if start and end:
            try:
                start = pd.to_datetime(start)
                end = pd.to_datetime(end)
                if start > end:
                    self.generate_btn.config(bg="#cccccc", fg="#666666", state=tk.DISABLED)
                    messagebox.showerror("日期无效", "开始日期不能晚于结束日期！")
                    self._log("错误: 开始日期不能晚于结束日期！")
                    self.status_msg.set("日期输入有误")
                    return
            except:
                self.generate_btn.config(bg="#cccccc", fg="#666666", state=tk.DISABLED)
                messagebox.showerror("日期无效", "日期格式不正确")
                self._log("错误: 日期格式有误！")
                self.status_msg.set("日期输入有误")
                return
        else:
            wait_input()
            return

        # =============================== 月份校验 ===============================
        year_month_ok = self.stat_year.get().strip() and self.stat_month.get().strip()
        
        # =============================== 月份校验 ===============================
        mode_ok = self.loose_mode.get() or self.strict_mode.get()
        # 还有配置文件是否加载OK
        if year_month_ok and mode_ok and self.config_ok:
            self.status_msg.set("就绪")
            self._log("所有输入已填写且校验通过")
            # 更新按钮样式（可点击）
            self.generate_btn.config(
                bg="#4CAF50",
                fg="white",
                state=tk.NORMAL  # 启用按钮
            )
        else:
            wait_input()

    def _log(self, message):
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.see(tk.END)
        self.root.update_idletasks() # 实时刷新日志

    def _select_biz_above_million_file(self):
        filename = filedialog.askopenfilename(
            initialdir=self.base_dir,
            filetypes=[('Excel file', '*.xlsx;*.xls')]
            )
        if filename:
            self.biz_above_million_filepath.set(filename)

    def _select_biz_below_million_file(self):
        filename = filedialog.askopenfilename(
            initialdir=self.base_dir,
            filetypes=[('Excel file', '*.xlsx;*.xls')]
        )
        if filename:
            self.biz_below_million_filepath.set(filename)
    
    def _select_contract_status_file(self):
        filename = filedialog.askopenfilename(
            initialdir=self.base_dir,
            filetypes=[('Excel file', '*.xlsx;*.xls')]
        )
        if filename:
            self.contract_status_filepath.set(filename)

    def _select_biz_code_file(self):
        filename = filedialog.askopenfilename(
            initialdir=self.base_dir,
            filetypes=[('Excel file', '*.xlsx;*.xls')]
        )
        if filename:
            self.biz_code_filepath.set(filename)

    def _load_config(self): 
        """
        加载配置文件
        """ 
        self._log("正在加载配置文件...")
        # ============================= 加载config文件 =============================
        self.config: Dict = None
        config_load_success = True
        try:
            # 1. 检查文件是否存在
            if not os.path.exists(self.config_filepath):
                raise FileNotFoundError(f"配置文件不存在！请检查路径：{self.config_filepath}")
            # 2. 读取并解析yaml文件
            with open(self.config_filepath, 'r', encoding='utf-8') as f:
                self.config = yaml.safe_load(f)
            # 3. 校验yaml是否为空
            if not self.config:
                raise ValueError(f"配置文件内容为空, 请检查：{self.config_filepath}")
        except PermissionError as pe:
            self._log(f"错误：加载配置失败！\n无权限读取配置文件: {self.config_filepath}\n{str(pe)}")
            config_load_success = False
        except ValueError as ve:
            self._log(f"错误: 加载配置失败！\n{str(ve)}")
            config_load_success = False
        except FileNotFoundError as fnfe:
            self._log(f"错误：加载配置失败！\n{str(fnfe)}")
            config_load_success = False
        except Exception as e:
            self._log(f"错误：加载配置失败！\n{str(e)}")
            config_load_success = False

        if config_load_success ==False:
            self.config_ok = False
            return

        # ===================== 读取配置（自动校验字段是否存在） =====================
        try:
            # 百万以上
            self.above_million_filename_keyword = self.config["above_million"]["workbook_name"]
            self.above_million_sheet = self.config["above_million"]["sheet_name"]
            self.above_million_required_cols = self.config["above_million"]["required_cols"]
            # 百万以下
            self.below_million_filename_keyword = self.config["below_million"]["workbook_name"]
            self.below_million_sheet = self.config["below_million"]["sheet_name"]
            self.below_million_required_cols = self.config["below_million"]["required_cols"]
            # 签约情况
            self.contract_status_filename_keyword = self.config["contract_status"]["workbook_name"]
            self.contract_status_sheet = self.config["contract_status"]["sheet_name"]
            self.contract_status_required_cols = self.config["contract_status"]["required_cols"]
            # 全量商机编码清单
            self.biz_code_filename_keyword = self.config['biz_code']['workbook_name']
            self.biz_code_sheet = self.config['biz_code']['sheet_name']
            self.biz_code_required_cols = self.config['biz_code']['required_cols']
            # 合同排除关键词
            # self.contrat_exclude_keywords = self.config["contract_exclude_keywords"]
        # 捕获：配置文件少写了字段（比如漏了 above_million）
        except KeyError as ke:
            self._log(f"错误: 加载配置失败！\n配置文件缺少关键字段: \n{str(ke)}")
            config_load_success = False
        except Exception as e:
            self._log(f"错误: 加载配置失败！\n{str(e)}")
            config_load_success = False
        
        if config_load_success == True:
            self._log("成功加载配置文件！")
            self.config_ok = True
        
    def _open_source_table(self):
        """使用pandas打开底表文件, 并检查是否存在指定列"""
        # 1. 读取单元/行业名称映射文件
        self._log("正在加载单元/行业名称映射表...")
        try:
            df_unit_map = pd.read_excel(
                self.unit_industry_map_filepath,
                sheet_name="单元", 
                header=1,
                usecols=['清单底表', '通报模板']
                )
            df_industry_map = pd.read_excel(
                self.unit_industry_map_filepath,
                sheet_name="行业",
                header=1,
                usecols=['清单底表', '通报模板']
                )
        except ValueError as ve:
            self._log(f"错误: 单元/行业名称映射文件缺少必需列！\n{str(ve)}")
            return None
        except Exception as e:
            self._log(f"错误: 单元/行业名称映射文件格式不正确！\n{str(e)}")
            return None
        self._log("成功加载单元/行业名称映射表！")
        
        # 2. 读取百万以上商机清单
        self._log("正在读取百万以上商机...")
        try:
            df_biz_above_million = pd.read_excel(
                self.biz_above_million_filepath.get().strip(),
                sheet_name=self.above_million_sheet,
                usecols=self.above_million_required_cols
                )
        except ValueError as ve: # usecols列缺失，捕获ValueError
            self._log(f"错误: 百万以上商机文件缺少必需列！\n{str(ve)}")
            return None
        except Exception as e:
            self._log(f"错误: 百万以上商机文件格式不正确！\n{str(e)}")
            return None
        self._log(f"成功读取百万以上商机{len(df_biz_above_million)} 个！")

        # 3. 读取百万以下商机清单
        self._log("正在读取百万以下商机...")
        try:
            df_biz_below_million = pd.read_excel(
                self.biz_below_million_filepath.get().strip(),
                sheet_name=self.below_million_sheet,
                usecols=self.below_million_required_cols)
        except ValueError as ve: # usecols列缺失，捕获ValueError
            self._log(f"错误: 百万以下商机文件缺少必需列！\n{str(ve)}")
            return None
        except Exception as e:
            self._log(f"错误: 百万以下商机文件格式不正确！\n{str(e)}")
            return None
        self._log(f"成功读取百万以下商机 {len(df_biz_below_million)} 个！")

        # 4. 读取商机签约情况
        self._log("正在读取签约商机清单...")
        try:
            df_contrat_status = pd.read_excel(
                self.contract_status_filepath.get().strip(),
                sheet_name=self.contract_status_sheet,
                usecols=self.contract_status_required_cols)
        except ValueError as ve: # usecols列缺失，捕获ValueError
            self._log(f"错误: 签约情况文件缺少必需列！\n{str(ve)}")
            return None
        except Exception as e:
            self._log(f"错误: 签约情况文件格式不正确！\n{str(e)}")
            return None
        self._log(f"成功读取签约商机 {len(df_contrat_status)} 个！")

        # 5. 读取全量商机清单及预计签约日期，用于严口径统计
        # # 由于商机系统导出的全量数据太多，读取时间会很长，只保留需要的两列，是一种适宜的方法
        self._log("正在读取全量商机编码清单及预计签约日期...")
        try:
            df_biz_code = pd.read_excel(
                self.biz_code_filepath.get().strip(),
                sheet_name=self.biz_code_sheet,
                usecols=self.biz_code_required_cols
                )
        except ValueError as ve: # usecols列缺失，捕获ValueError
            self._log(f"错误: 全量商机编码文件缺少必需列！\n{str(ve)}")
            return None
        except Exception as e:
            self._log(f"错误：全量商机编码文件格式不正确！\n{str(e)}")
            return None
        self._log(f"成功读取商机编码 {len(df_biz_code) }个！")
        
        # 5. 成功返回6个DataFrame
        return df_biz_above_million, df_biz_below_million, df_contrat_status, df_unit_map, df_industry_map, df_biz_code

    def _summarize_by_unit_and_industry(self, df_biz_above_million, df_biz_below_million, df_contract_status, df_unit_map, df_industry_map, df_biz_code): 
        """
        根据100万以上/以下商机清单，及合同签约情况生成每周商机进展报告
        @param df_biz_above_million: 100万以上商机清单
        @param df_biz_below_million: 100万以下商机清单
        @param df_contract_status: 合同实际签约情况
        @param df_unit_map: 单元名称映射表
        @param df_industry_map: 行业名称映射表
        @param df_biz_code: 全量商机编码及预计签约日期, 用于严口径校验
        """
        # 1. 数据预处理
        
        # # 1.1 删去签约情况表中，"比对"列内容包含"集成""陕数""省公司"的合同
        # contract_exclude_pattern = "|".join(self.contrat_exclude_keywords)
        # df_contract_status_required = df_contract_status_required[
        #     ~df_contract_status_required[self.above_million_required_cols[0]].str.contains(contract_exclude_pattern, na=False) #na视为不包含关键词
        #     ].copy()
        # self._log(f"已清理\"{self.above_million_required_cols[0]}\"列含 {self.contrat_exclude_keywords} 的签约数据，剩余 {len(df_contract_status_required)} 个！")
        
        # 1.1 根据配置文件确定必需列的名称
        # 百万以上
        above_million_bizcode_col_name, above_million_unit_col_name, above_million_industry_col_name, above_million_date_col_name, above_million_money_col_name, above_million_target_col_name = self.above_million_required_cols
        # 百万以下
        below_million_bizcode_col_name, below_million_unit_col_name, below_million_industry_col_name, below_million_date_col_name, below_million_money_col_name = self.below_million_required_cols
        # 签约情况
        contract_status_unit_col_name, contract_status_industry_col_name, contract_status_date_col_name, contract_status_money_col_name, contract_status_bizcode_col_name = self.contract_status_required_cols
        # 商机编码清单
        biz_code_bizcode_col_name, biz_code_date_col_name = self.biz_code_required_cols
        
        # 1.3 确保数据类型正确
        # # 日期列
        df_biz_above_million[above_million_date_col_name] = pd.to_datetime(
            df_biz_above_million[above_million_date_col_name],
            format="%Y%m", # 底表格式是202604
            errors="coerce" # 无法转换时填充为空，不报错
        )#.dt.strftime("%Y/%m/%d") # 自动补1号，例如20260401
        df_biz_below_million[below_million_date_col_name] = pd.to_datetime(
            df_biz_below_million[below_million_date_col_name],
            errors="coerce"
        ).dt.strftime("%Y/%m/%d")
        df_contract_status[contract_status_date_col_name] = pd.to_datetime(
            df_contract_status[contract_status_date_col_name],
            errors="coerce"
        ).dt.strftime("%Y/%m/%d")
        df_biz_code[biz_code_date_col_name] = pd.to_datetime(
            df_biz_code[biz_code_date_col_name],
            errors='coerce'
        )
        # # 数值列
        df_biz_above_million[above_million_money_col_name] = pd.to_numeric(
            df_biz_above_million[above_million_money_col_name],
            errors="coerce" # 无法转换时填充为空，不报错
        ).fillna(0).astype('float64')
        df_biz_below_million[below_million_money_col_name] = pd.to_numeric(
            df_biz_below_million[below_million_money_col_name],
            errors="coerce"
        ).fillna(0).astype('float64')
        df_contract_status[contract_status_money_col_name] = pd.to_numeric(
            df_contract_status[contract_status_money_col_name],
            errors="coerce"
        ).fillna(0).astype('float64')
        # # 将签约情况表中的"省口径签约额"列单位从"元"改为"万元"
        df_contract_status[contract_status_money_col_name] = df_contract_status[contract_status_money_col_name] / 10000
        
        # 1.4 构建单元名称映射表、行业名称映射表，可能存在多对一映射关系
        ## 注意不要用dict，因为清单底表列存在"高新","南高新"这种相互包含的关键字
        ## 并且本来就是要逐个遍历，所以即使使用dict也不会提高效率
        ## 因此使用list保留顺序, 并根据关键字长度排个序，长的关键字在前边，避免"南高新"被"高新"给匹配走
        unit_map_list = df_unit_map[["清单底表", "通报模板"]].values.tolist() #values是将DataFrame转换成二维numpy数组，tolist()再将其转换成列表
        unit_map_list.sort(key=lambda x: len(x[0]), reverse=True)
        industry_map_list = df_industry_map[["清单底表", "通报模板"]].values.tolist()
        industry_map_list.sort(key=lambda x: len(x[0]), reverse=True)

        # 1.5 根据单元_行业名称映射表，将底表的单元、行业两列映射成与模板中一致
        # 注意：非绝对匹配，只要底表名称包含"清单底表"列的某个键，就可以映射到对应的"通报模板"名称，不成功会标记为"unmatched"
        self._log("正在映射单元和行业名称...")
        df_biz_above_million[above_million_unit_col_name] = df_biz_above_million[above_million_unit_col_name].apply(lambda x: next((v for k, v in unit_map_list if k in str(x)), "unmatched")) # next()函数用于返回第一个满足条件的元素，如果没有满足条件的元素，则返回默认值x（即原值）
        df_biz_below_million[below_million_unit_col_name] = df_biz_below_million[below_million_unit_col_name].apply(lambda x: next((v for k, v in unit_map_list if k in str(x)), "unmatched"))
        df_contract_status[contract_status_unit_col_name] = df_contract_status[contract_status_unit_col_name].apply(lambda x: next((v for k, v in unit_map_list if k in str(x)), "unmatched"))

        df_biz_above_million[above_million_industry_col_name] = df_biz_above_million[above_million_industry_col_name].apply(lambda x: next((v for k, v in industry_map_list if k in str(x)), "unmatched"))
        df_biz_below_million[below_million_industry_col_name] = df_biz_below_million[below_million_industry_col_name].apply(lambda x: next((v for k, v in industry_map_list if k in str(x)), "unmatched"))
        df_contract_status[contract_status_industry_col_name] = df_contract_status[contract_status_industry_col_name].apply(lambda x: next((v for k, v in industry_map_list if k in str(x)), "unmatched"))
        self._log("成功映射单元和行业名称！")

        # 1.6 保存预处理之后的数据到文件，留待后续参考
        def save_prehandled_data(origin_filepath, sheet_name, df):
            file_name = os.path.basename(origin_filepath)
            new_filepath = re.sub(
                r'(\.xlsx|\.xls)$', 
                fr'_预处理_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}\1',
                file_name,
                flags=re.IGNORECASE)
            prehandled_output_dir = os.path.join(self.base_dir, "预处理后的数据")
            os.makedirs(prehandled_output_dir, exist_ok=True)
            abs_new_filepath = os.path.join(prehandled_output_dir, new_filepath)
            try:
                df.to_excel(
                    abs_new_filepath,
                    sheet_name = sheet_name,
                    index = False
                )
            except Exception as e:
                self._log(f"错误:保存预处理数据失败\n{str(e)}", level="error")
                return
            self._log(f"预处理后的数据文件已保存至: \n{abs_new_filepath}")
            
        save_prehandled_data(self.biz_above_million_filepath.get(), self.above_million_sheet, df_biz_above_million)
        save_prehandled_data(self.biz_below_million_filepath.get(), self.below_million_sheet, df_biz_below_million)
        save_prehandled_data(self.contract_status_filepath.get(), self.contract_status_sheet, df_contract_status)

        # 2. 根据用户输入的结束日期确定统计时间段、统计当月及后续三月
        # 2.1 商机量、商机金额总计
        df_agg_above_million_by_unit = df_biz_above_million.groupby("单元").agg(
            商机量=("单元", "size"),
            商机金额=("")
        )

        # 2.1 时间段统计
        # # 百万以上需要用户在"是否目标时间段"


        
        current_month = pd.to_datetime(self.end_date).month
        month_after_1_months = (pd.to_datetime(self.end_date) + pd.DateOffset(months=1)).month
        month_after_2_months = (pd.to_datetime(self.end_date) + pd.DateOffset(months=2)).month
        month_after_3_months = (pd.to_datetime(self.end_date) + pd.DateOffset(months=3)).month
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
        # 月份跨平台兼容
        month_format = "%#m月" if platform.system() =="Windows" else "%-m月"
        month = pd.to_datetime(self.end_date).strftime(month_format) 
        month_after_1_months = (period_end + pd.DateOffset(months=1)).strftime(month_format)
        month_after_2_months = (period_end + pd.DateOffset(months=2)).strftime(month_format)
        month_after_3_months = (period_end + pd.DateOffset(months=3)).strftime(month_format)

    def run_biz_analysis_workflow(self):
        # 1. 读取文件并检查Sheet名、列名
        df_set = self._open_source_table()
        if not df_set:
            return
        self._summarize_by_unit_and_industry(*df_set)
        # TODO：根据用户选择，进行宽口径、严口径或两种口径的统计 

if __name__ == "__main__":
    root = tk.Tk() # Tk()是创建一个Tkinter应用程序的主窗口对象，所有的组件都要放到这个主窗口上，
    app = BizReportApp(root)
    # mainloop()是Tkinter应用程序的事件循环，负责监听用户的操作并做出相应，必须在创建完所有组建后调用
    root.mainloop()