#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
多段数据电压电流曲线对比工具
支持多文件上传、日期标签、自动处理数据后生成交互式HTML图表
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
import re
from datetime import datetime, timedelta
from threading import Thread
import webbrowser

# 尝试导入中文支持
try:
    import locale
    locale.setlocale(locale.LC_ALL, 'zh_CN.UTF-8')
except:
    pass


class VoltageCurrentPlotter:
    def __init__(self, root):
        self.root = root
        self.root.title("电压电流对比图工具")
        self.root.geometry("1200x850")  # 增加高度以容纳新组件

        # 数据存储: [(file_path, date_label, color_index), ...]
        self.file_list = []

        # 上一轮文件列表备份（用于恢复功能）
        self.last_file_list = []

        # 历史记录存储
        self.history_files = []  # [(timestamp, file_path)]

        # 上次选择的日期（用于日期记忆）
        self.last_selected_date = datetime.now().strftime("%Y-%m-%d")

        # 日期段生成的日期列表
        self.generated_dates = []

        # 保存日期段的起始日期（用于重置后恢复）
        self.date_range_start = None

        # 标记是否从日期列表选择了日期
        self.selected_from_list = False

        # 预定义颜色方案 - 电压和电流各用一种颜色
        self.voltage_color = '#DC143C'  # 深红色
        self.current_color = '#1E90FF'  # 深蓝色

        # 电流单位配置
        self.current_units = ['A', 'mA', 'μA']  # 可用单位
        self.current_unit_factors = {'A': 1, 'mA': 1000, 'μA': 1000000}  # 转换因子
        self.selected_current_unit = tk.StringVar(value='mA')  # 默认单位改为mA

        # 自动添加文件开关（日期段模式下）
        self.auto_add_files = tk.BooleanVar(value=False)  # 默认关闭

        # 智能分离曲线开关
        self.smart_separate_curves = tk.BooleanVar(value=True)  # 默认开启

        # 创建历史曲线文件夹
        self.history_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), '历史曲线')
        if not os.path.exists(self.history_folder):
            os.makedirs(self.history_folder)

        self._load_history()
        self._setup_ui()

    def _load_history(self):
        """加载历史记录"""
        history_file = os.path.join(self.history_folder, 'chart_history.txt')
        if os.path.exists(history_file):
            try:
                with open(history_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if line:
                            parts = line.split('|')
                            if len(parts) >= 2:
                                timestamp = parts[0]
                                file_path = parts[1]
                                self.history_files.append((timestamp, file_path))
            except Exception as e:
                pass

    def _save_history(self, timestamp, file_path):
        """保存历史记录"""
        history_file = os.path.join(self.history_folder, 'chart_history.txt')
        try:
            with open(history_file, 'a', encoding='utf-8') as f:
                f.write(f"{timestamp}|{file_path}\n")
        except Exception as e:
            pass

    def _setup_ui(self):
        """设置GUI界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=0)  # 历史记录固定宽度
        main_frame.rowconfigure(2, weight=1)  # 文件列表
        main_frame.rowconfigure(3, weight=0)  # 日期列表（不自动扩展）

        # 标题
        title_label = ttk.Label(
            main_frame,
            text="电压电流对比图工具",
            font=("Microsoft YaHei", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=10)  # 跨两列

        # 文件上传区域
        upload_frame = ttk.LabelFrame(main_frame, text="文件上传", padding="10")
        upload_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        # 第一行：选择文件
        ttk.Button(upload_frame, text="📁 选择文件 (CSV/TXT)",
                  command=self._select_files).grid(row=0, column=0, padx=5, pady=3, sticky=tk.W)

        self.file_count_label = ttk.Label(upload_frame, text="未选择文件", foreground="gray")
        self.file_count_label.grid(row=0, column=1, padx=10, sticky=tk.W)

        # 第二行：日期选择
        ttk.Label(upload_frame, text="数据日期:").grid(row=1, column=0, padx=5, pady=3, sticky=tk.W)

        self.date_entry = ttk.Entry(upload_frame, width=15)
        self.date_entry.grid(row=1, column=1, padx=5, pady=3, sticky=tk.W)
        self.date_entry.insert(0, self.last_selected_date)

        # 日期选择按钮
        ttk.Button(upload_frame, text="📅 选择日期",
                  command=self._select_date).grid(row=1, column=2, padx=5, pady=3)

        # 前一天/后一天按钮
        ttk.Button(upload_frame, text="◀ 前一天",
                  command=self._previous_day, width=8).grid(row=1, column=3, padx=2, pady=3)
        ttk.Button(upload_frame, text="后一天 ▶",
                  command=self._next_day, width=8).grid(row=1, column=4, padx=2, pady=3)

        # 电流单位选择
        ttk.Label(upload_frame, text="电流单位:").grid(row=1, column=5, padx=5, pady=3, sticky=tk.W)
        current_unit_combo = ttk.Combobox(upload_frame, textvariable=self.selected_current_unit,
                                          values=self.current_units, width=5, state='readonly')
        current_unit_combo.grid(row=1, column=6, padx=5, pady=3)

        # 添加到列表按钮
        ttk.Button(upload_frame, text="➕ 添加到列表",
                  command=self._add_file_with_date).grid(row=1, column=7, padx=5, pady=3)

        # 第三行：日期段批量生成
        ttk.Label(upload_frame, text="日期段:").grid(row=2, column=0, padx=5, pady=3, sticky=tk.W)
        ttk.Button(upload_frame, text="📅 生成日期段",
                  command=self._generate_date_range).grid(row=2, column=1, padx=5, pady=3, sticky=tk.W)
        ttk.Button(upload_frame, text="🗑️ 清空日期段",
                  command=self._clear_date_range).grid(row=2, column=2, padx=5, pady=3, sticky=tk.W)
        # 自动上传文件开关
        ttk.Checkbutton(upload_frame, text="自动上传文件",
                       variable=self.auto_add_files).grid(row=2, column=3, padx=10, pady=3, sticky=tk.W)
        # 智能分离曲线开关
        ttk.Checkbutton(upload_frame, text="智能分离曲线",
                       variable=self.smart_separate_curves).grid(row=2, column=4, padx=10, pady=3, sticky=tk.W)

        # 已选文件列表
        list_frame = ttk.LabelFrame(main_frame, text="已上传文件列表", padding="10")
        list_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        # 创建树形视图
        columns = ("序号", "文件名", "日期标签", "路径")
        self.file_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=6)

        self.file_tree.heading("序号", text="序号")
        self.file_tree.heading("文件名", text="文件名")
        self.file_tree.heading("日期标签", text="日期标签")
        self.file_tree.heading("路径", text="路径")

        self.file_tree.column("序号", width=50, anchor=tk.CENTER)
        self.file_tree.column("文件名", width=200, anchor=tk.W)
        self.file_tree.column("日期标签", width=120, anchor=tk.CENTER)
        self.file_tree.column("路径", width=400, anchor=tk.W)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)

        self.file_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # 按钮区
        btn_frame = ttk.Frame(list_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=5)

        ttk.Button(btn_frame, text="删除选中", command=self._remove_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="清空列表", command=self._clear_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="恢复上一轮", command=self._restore_last_file_list).pack(side=tk.LEFT, padx=5)

        # 日期段列表区域（在文件列表下方）
        date_list_frame = ttk.LabelFrame(main_frame, text="生成的日期段列表（点击日期快速填入）", padding="10")
        date_list_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        date_list_frame.columnconfigure(0, weight=1)

        # 创建日期列表（使用Listbox显示生成的日期）
        date_list_container = ttk.Frame(date_list_frame)
        date_list_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.date_listbox = tk.Listbox(date_list_container, height=5, exportselection=False)
        self.date_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 日期列表滚动条
        date_scrollbar = ttk.Scrollbar(date_list_container, orient=tk.VERTICAL, command=self.date_listbox.yview)
        date_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.date_listbox.configure(yscrollcommand=date_scrollbar.set)

        # 绑定点击事件
        self.date_listbox.bind('<<ListboxSelect>>', self._on_date_select)

        # 日期列表提示
        date_hint = ttk.Label(date_list_frame, text="提示：生成日期段后，点击日期自动填入上方", foreground="gray", font=("Arial", 9))
        date_hint.grid(row=1, column=0, pady=3)

        # 历史记录区域（跨越row=2和row=3）
        history_frame = ttk.LabelFrame(main_frame, text="历史记录", padding="10")
        history_frame.grid(row=2, column=1, rowspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        history_frame.columnconfigure(0, weight=1)
        history_frame.rowconfigure(0, weight=1)

        # 创建历史记录树形视图
        hist_columns = ("时间", "文件名")
        self.history_tree = ttk.Treeview(history_frame, columns=hist_columns, show="headings", height=6)

        self.history_tree.heading("时间", text="生成时间")
        self.history_tree.heading("文件名", text="文件名")

        self.history_tree.column("时间", width=140, anchor=tk.CENTER)
        self.history_tree.column("文件名", width=180, anchor=tk.W)

        # 添加滚动条
        hist_scrollbar = ttk.Scrollbar(history_frame, orient=tk.VERTICAL, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=hist_scrollbar.set)

        self.history_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        hist_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # 历史记录按钮
        hist_btn_frame = ttk.Frame(history_frame)
        hist_btn_frame.grid(row=1, column=0, columnspan=2, pady=5)

        ttk.Button(hist_btn_frame, text="🔍 查看曲线", command=self._view_history).pack(side=tk.LEFT, padx=3)
        ttk.Button(hist_btn_frame, text="📂 打开文件夹", command=self._open_history_folder).pack(side=tk.LEFT, padx=3)
        ttk.Button(hist_btn_frame, text="🗑️ 清空历史", command=self._clear_history).pack(side=tk.LEFT, padx=3)

        # 双击查看
        self.history_tree.bind('<Double-1>', lambda e: self._view_history())

        # 加载历史记录到列表
        self._refresh_history_list()

        # 处理按钮
        process_frame = ttk.Frame(main_frame)
        process_frame.grid(row=4, column=0, columnspan=2, pady=10)

        ttk.Button(process_frame, text="开始处理并生成图表",
                  command=self._start_processing, width=25).pack(side=tk.LEFT, padx=10)
        ttk.Button(process_frame, text="重置",
                  command=self._reset_all, width=15).pack(side=tk.LEFT, padx=10)

        # 状态提示区域
        status_frame = ttk.LabelFrame(main_frame, text="状态信息", padding="10")
        status_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)

        self.status_text = ScrolledText(status_frame, height=8, wrap=tk.WORD)
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        self._log("欢迎使用电压电流对比图工具！")
        self._log("请选择CSV/TXT文件并添加日期标签后开始处理。")

    def _log(self, message, level="INFO"):
        """记录日志信息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        color_map = {
            "INFO": "black",
            "SUCCESS": "green",
            "WARNING": "orange",
            "ERROR": "red"
        }

        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n", level)
        self.status_text.tag_config(level, foreground=color_map.get(level, "black"))
        self.status_text.see(tk.END)

    def _select_files(self):
        """选择文件"""
        files = filedialog.askopenfilenames(
            title="选择数据文件",
            filetypes=[
                ("支持文件", "*.csv *.txt *.xlsx *.xls"),
                ("CSV文件", "*.csv"),
                ("文本文件", "*.txt"),
                ("Excel文件", "*.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
        )
        if files:
            self.selected_files = files
            self.file_count_label.config(text=f"已选择 {len(files)} 个文件", foreground="green")
            self._log(f"已选择 {len(files)} 个文件", "INFO")

            # 如果开启了自动上传且从日期列表选择了日期，自动添加到列表
            if self.auto_add_files.get() and self.selected_from_list:
                self._add_file_with_date()
                # 注意：不要在这里重置 selected_from_list
                # _add_file_with_date 函数内部会根据日期段情况来控制这个标志

    def _select_date(self):
        """选择日期"""
        # 创建日期选择对话框
        date_dialog = tk.Toplevel(self.root)
        date_dialog.title("选择日期")
        date_dialog.geometry("300x400")
        date_dialog.transient(self.root)
        date_dialog.grab_set()

        # 使用上次选择的日期作为默认值
        try:
            last_date = datetime.strptime(self.last_selected_date, "%Y-%m-%d")
            current_year = last_date.year
            current_month = last_date.month
            current_day = last_date.day
        except:
            now = datetime.now()
            current_year = now.year
            current_month = now.month
            current_day = now.day

        # 年份选择
        ttk.Label(date_dialog, text="年份:").pack(pady=5)
        year_var = tk.StringVar(value=str(current_year))
        year_combo = ttk.Combobox(date_dialog, textvariable=year_var,
                                   values=[str(y) for y in range(current_year - 5, current_year + 6)],
                                   width=10, state='readonly')
        year_combo.pack(pady=5)

        # 月份选择
        ttk.Label(date_dialog, text="月份:").pack(pady=5)
        month_var = tk.StringVar(value=str(current_month))
        month_combo = ttk.Combobox(date_dialog, textvariable=month_var,
                                    values=[str(m).zfill(2) for m in range(1, 13)],
                                    width=10, state='readonly')
        month_combo.pack(pady=5)

        # 日期选择
        ttk.Label(date_dialog, text="日期:").pack(pady=5)
        day_var = tk.StringVar(value=str(current_day))
        day_combo = ttk.Combobox(date_dialog, textvariable=day_var,
                                  values=[str(d).zfill(2) for d in range(1, 32)],
                                  width=10, state='readonly')
        day_combo.pack(pady=5)

        def on_ok():
            year = year_var.get()
            month = month_var.get().zfill(2)
            day = day_var.get().zfill(2)
            self.date_entry.delete(0, tk.END)
            self.date_entry.insert(0, f"{year}-{month}-{day}")
            date_dialog.destroy()

        ttk.Button(date_dialog, text="确定", command=on_ok).pack(pady=20)

        # 居中显示
        date_dialog.update_idletasks()
        x = (date_dialog.winfo_screenwidth() // 2) - (date_dialog.winfo_width() // 2)
        y = (date_dialog.winfo_screenheight() // 2) - (date_dialog.winfo_height() // 2)
        date_dialog.geometry(f"+{x}+{y}")

        self.root.wait_window(date_dialog)

    def _previous_day(self):
        """切换到前一天"""
        try:
            current_date = datetime.strptime(self.date_entry.get(), "%Y-%m-%d")
            prev_date = current_date - timedelta(days=1)
            self.date_entry.delete(0, tk.END)
            self.date_entry.insert(0, prev_date.strftime("%Y-%m-%d"))
            self.last_selected_date = prev_date.strftime("%Y-%m-%d")
        except:
            pass

    def _next_day(self):
        """切换到后一天"""
        try:
            current_date = datetime.strptime(self.date_entry.get(), "%Y-%m-%d")
            next_date = current_date + timedelta(days=1)
            self.date_entry.delete(0, tk.END)
            self.date_entry.insert(0, next_date.strftime("%Y-%m-%d"))
            self.last_selected_date = next_date.strftime("%Y-%m-%d")
        except:
            pass

    def _generate_date_range(self):
        """生成日期段"""
        # 创建日期段选择对话框
        range_dialog = tk.Toplevel(self.root)
        range_dialog.title("生成日期段")
        range_dialog.geometry("350x450")
        range_dialog.transient(self.root)
        range_dialog.grab_set()

        # 使用当前日期作为默认值
        try:
            current_date = datetime.strptime(self.date_entry.get(), "%Y-%m-%d")
            start_year = current_date.year
            start_month = current_date.month
            start_day = current_date.day
            end_year = current_date.year
            end_month = current_date.month
            end_day = current_date.day
        except:
            now = datetime.now()
            start_year = now.year
            start_month = now.month
            start_day = now.day
            end_year = now.year
            end_month = now.month
            end_day = now.day

        # 开始日期选择
        ttk.Label(range_dialog, text="开始日期", font=("Arial", 10, "bold")).pack(pady=5)

        start_frame = ttk.Frame(range_dialog)
        start_frame.pack(pady=5)

        ttk.Label(start_frame, text="年:").pack(side=tk.LEFT, padx=2)
        start_year_var = tk.StringVar(value=str(start_year))
        start_year_combo = ttk.Combobox(start_frame, textvariable=start_year_var,
                                        values=[str(y) for y in range(start_year - 5, start_year + 6)],
                                        width=6, state='readonly')
        start_year_combo.pack(side=tk.LEFT, padx=2)

        ttk.Label(start_frame, text="月:").pack(side=tk.LEFT, padx=2)
        start_month_var = tk.StringVar(value=str(start_month).zfill(2))
        start_month_combo = ttk.Combobox(start_frame, textvariable=start_month_var,
                                         values=[str(m).zfill(2) for m in range(1, 13)],
                                         width=4, state='readonly')
        start_month_combo.pack(side=tk.LEFT, padx=2)

        ttk.Label(start_frame, text="日:").pack(side=tk.LEFT, padx=2)
        start_day_var = tk.StringVar(value=str(start_day).zfill(2))
        start_day_combo = ttk.Combobox(start_frame, textvariable=start_day_var,
                                       values=[str(d).zfill(2) for d in range(1, 32)],
                                       width=4, state='readonly')
        start_day_combo.pack(side=tk.LEFT, padx=2)

        # 结束日期选择
        ttk.Label(range_dialog, text="结束日期", font=("Arial", 10, "bold")).pack(pady=5)

        end_frame = ttk.Frame(range_dialog)
        end_frame.pack(pady=5)

        ttk.Label(end_frame, text="年:").pack(side=tk.LEFT, padx=2)
        end_year_var = tk.StringVar(value=str(end_year))
        end_year_combo = ttk.Combobox(end_frame, textvariable=end_year_var,
                                      values=[str(y) for y in range(end_year - 5, end_year + 6)],
                                      width=6, state='readonly')
        end_year_combo.pack(side=tk.LEFT, padx=2)

        ttk.Label(end_frame, text="月:").pack(side=tk.LEFT, padx=2)
        end_month_var = tk.StringVar(value=str(end_month).zfill(2))
        end_month_combo = ttk.Combobox(end_frame, textvariable=end_month_var,
                                       values=[str(m).zfill(2) for m in range(1, 13)],
                                       width=4, state='readonly')
        end_month_combo.pack(side=tk.LEFT, padx=2)

        ttk.Label(end_frame, text="日:").pack(side=tk.LEFT, padx=2)
        end_day_var = tk.StringVar(value=str(end_day).zfill(2))
        end_day_combo = ttk.Combobox(end_frame, textvariable=end_day_var,
                                     values=[str(d).zfill(2) for d in range(1, 32)],
                                     width=4, state='readonly')
        end_day_combo.pack(side=tk.LEFT, padx=2)

        def on_generate():
            try:
                start_date = datetime(int(start_year_var.get()),
                                     int(start_month_var.get()),
                                     int(start_day_var.get()))
                end_date = datetime(int(end_year_var.get()),
                                   int(end_month_var.get()),
                                   int(end_day_var.get()))

                if start_date > end_date:
                    messagebox.showerror("错误", "开始日期不能晚于结束日期！")
                    return

                # 生成日期列表
                self.generated_dates = []
                current = start_date
                while current <= end_date:
                    self.generated_dates.append(current.strftime("%Y-%m-%d"))
                    current += timedelta(days=1)

                # 保存��期段起始日期（用于重置后恢复）
                self.date_range_start = self.generated_dates[0]

                # 更新Listbox显示
                self.date_listbox.delete(0, tk.END)
                for date_str in self.generated_dates:
                    self.date_listbox.insert(tk.END, date_str)

                # 自动选中第一个日期
                if self.generated_dates:
                    self.date_listbox.selection_set(0)
                    self.date_listbox.see(0)
                    # 触发选中事件，填入日期
                    self._on_date_select()

                # 自动开启"自动上传文件"功能
                self.auto_add_files.set(True)
                self._log(f"已生成日期段：{len(self.generated_dates)} 天（{self.generated_dates[0]} 至 {self.generated_dates[-1]}），已自动开启'自动上传文件'功能", "SUCCESS")
                range_dialog.destroy()

            except Exception as e:
                messagebox.showerror("错误", f"日期格式错误：{str(e)}")

        ttk.Button(range_dialog, text="生成日期段", command=on_generate).pack(pady=20)

        # 居中显示
        range_dialog.update_idletasks()
        x = (range_dialog.winfo_screenwidth() // 2) - (range_dialog.winfo_width() // 2)
        y = (range_dialog.winfo_screenheight() // 2) - (range_dialog.winfo_height() // 2)
        range_dialog.geometry(f"+{x}+{y}")

        self.root.wait_window(range_dialog)

    def _clear_date_range(self):
        """清空日期段列表"""
        self.generated_dates = []
        self.date_listbox.delete(0, tk.END)
        # 清除保存的起始日期
        self.date_range_start = None
        # 重置自动添加标志
        self.selected_from_list = False
        self._log("已清空日期段列表", "INFO")

    def _on_date_select(self, event=None):
        """日期列表点击事件"""
        selection = self.date_listbox.curselection()
        if selection:
            index = selection[0]
            selected_date = self.date_listbox.get(index)
            self.date_entry.delete(0, tk.END)
            self.date_entry.insert(0, selected_date)
            self.last_selected_date = selected_date
            # 只有在自动上传开关开启时，才标记为自动添加模式
            if self.auto_add_files.get():
                self.selected_from_list = True
            else:
                self.selected_from_list = False

    def _refresh_history_list(self):
        """刷新历史记录列表"""
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)

        for timestamp, file_path in self.history_files:
            file_name = os.path.basename(file_path)
            self.history_tree.insert("", tk.END, values=(timestamp, file_name))

    def _view_history(self):
        """查看选中的历史曲线"""
        selection = self.history_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一条历史记录！")
            return

        item = self.history_tree.item(selection[0])
        file_path = item['values'][1]

        if os.path.exists(file_path):
            webbrowser.open(f'file:///{file_path.replace(os.sep, "/")}')
            self._log(f"正在打开: {os.path.basename(file_path)}", "INFO")
        else:
            messagebox.showerror("错误", "文件不存在！")

    def _open_history_folder(self):
        """打开历史记录文件夹"""
        os.startfile(self.history_folder)

    def _clear_history(self):
        """清空历史记录"""
        if not self.history_files:
            messagebox.showinfo("提示", "历史记录为空！")
            return

        if messagebox.askyesno("确认", "确定要清空所有历史记录吗？\n这不会删除已生成的HTML文件。"):
            self.history_files.clear()
            history_file = os.path.join(self.history_folder, 'chart_history.txt')
            if os.path.exists(history_file):
                os.remove(history_file)
            self._refresh_history_list()
            self._log("历史记录已清空", "INFO")

    def _add_file_with_date(self):
        """添加文件到列表（带日期标签）"""
        if not hasattr(self, 'selected_files') or not self.selected_files:
            messagebox.showwarning("警告", "请先选择文件！")
            return

        date_str = self.date_entry.get().strip()
        if not self._validate_date(date_str):
            messagebox.showerror("错误", "日期格式错误！请使用 YYYY-MM-DD 格式\n示例: 2024-05-20")
            return

        last_inserted_item = None

        for file_path in self.selected_files:
            file_name = os.path.basename(file_path)

            # 检查是否已存在
            if any(f[0] == file_path for f in self.file_list):
                self._log(f"文件已存在: {file_name}", "WARNING")
                continue

            # 不再需要color_idx，电压电流统一颜色
            self.file_list.append((file_path, date_str, 0))

            # 更新树形视图
            last_inserted_item = self.file_tree.insert("", tk.END, values=(
                len(self.file_list),
                file_name,
                date_str,
                file_path
            ))

        # 自动滚动到最后添加的文件位置
        if last_inserted_item:
            self.file_tree.see(last_inserted_item)

        count = len(self.selected_files)
        self._log(f"成功添加 {count} 个文件，日期: {date_str}", "SUCCESS")
        self._log(f"当前共加载 {len(self.file_list)} 个文件", "INFO")

        # 记住本次选择的日期，方便下次使用
        self.last_selected_date = date_str

        # 如果日期段列表不为空，自动进入下一天
        if self.generated_dates:
            current_date_str = self.date_entry.get().strip()
            try:
                current_date = datetime.strptime(current_date_str, "%Y-%m-%d")
                # 查找当前日期在生成列表中的位置
                for i, gen_date_str in enumerate(self.generated_dates):
                    if gen_date_str == current_date_str:
                        # 移动到下一个日期
                        if i + 1 < len(self.generated_dates):
                            next_date = self.generated_dates[i + 1]
                            self.date_entry.delete(0, tk.END)
                            self.date_entry.insert(0, next_date)
                            self.last_selected_date = next_date
                            # 自动选中列表中的下一个日期
                            self.date_listbox.selection_clear(0, tk.END)
                            self.date_listbox.selection_set(i + 1)
                            self.date_listbox.see(i + 1)
                            self._log(f"已自动切换到: {next_date}", "INFO")
                            # 在日期段模式下，保持自动添加功能
                            self.selected_from_list = True
                        break
            except:
                pass

        # 只在非自动模式下删除selected_files（避免影响连续上传）
        if not self.selected_from_list:
            del self.selected_files

    def _validate_date(self, date_str):
        """验证日期格式"""
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
            return True
        except ValueError:
            return False

    def _remove_selected(self):
        """删除选中项"""
        selection = self.file_tree.selection()
        if not selection:
            return

        indices = sorted([int(self.file_tree.item(item)["values"][0]) - 1 for item in selection], reverse=True)

        for idx in indices:
            self.file_list.pop(idx)

        self._refresh_file_list()
        self._log(f"已删除 {len(indices)} 个文件", "INFO")

    def _clear_all(self):
        """清空所有文件"""
        self.file_list.clear()
        self._refresh_file_list()
        self._log("已清空所有文件", "WARNING")

    def _refresh_file_list(self):
        """刷新文件列表显示"""
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        for idx, (file_path, date_str, color_idx) in enumerate(self.file_list, 1):
            file_name = os.path.basename(file_path)
            self.file_tree.insert("", tk.END, values=(
                idx, file_name, date_str, file_path
            ))

    def _reset_all(self):
        """重置所有"""
        self._clear_all()
        self.status_text.delete(1.0, tk.END)
        self.progress['value'] = 0
        self._log("已重置，请重新选择文件", "INFO")

    def _restore_last_file_list(self):
        """恢复上一轮文件列表"""
        if not self.last_file_list:
            messagebox.showinfo("提示", "没有可恢复的文件列表！\n请先生成一次图表后再使用此功能。")
            return

        # 询问用户是否恢复
        if not messagebox.askyesno("确认", f"确定要恢复上一轮的文件列表吗？\n共 {len(self.last_file_list)} 个文件。\n\n注意：这将清空当前列表。"):
            return

        # 清空当前列表
        self.file_list.clear()

        # 复制上一轮文件列表
        self.file_list = self.last_file_list.copy()

        # 刷新显示
        self._refresh_file_list()
        self._log(f"已恢复上一轮文件列表（{len(self.file_list)} 个文件）", "SUCCESS")

        # 更新当前文件数量提示
        self.file_count_label.config(text=f"", foreground="gray")

    def _auto_reset_after_process(self):
        """处理完成后自动清空文件列表"""
        # 保存当前文件列表到备份
        if self.file_list:
            self.last_file_list = self.file_list.copy()
            self._log(f"已保存上一轮文件列表（{len(self.last_file_list)} 个文件），可通过'恢复上一轮'按钮恢复", "INFO")

        self._clear_all()
        self.progress['value'] = 0
        # 不清空状态日志，保留处理记录

        # 如果有日期段起始日期，恢复到起始日期并选中
        if self.date_range_start and self.generated_dates:
            self.date_entry.delete(0, tk.END)
            self.date_entry.insert(0, self.date_range_start)
            self.last_selected_date = self.date_range_start
            # 自动选中列表中的第一个日期
            self.date_listbox.selection_clear(0, tk.END)
            self.date_listbox.selection_set(0)
            self.date_listbox.see(0)
            self._log(f"已自动恢复到日期段起始日期：{self.date_range_start}", "INFO")

    def _parse_time_value(self, time_value, date_label):
        """解析时间值，支持多种格式"""
        # 如果是 datetime.time 对象（从Excel读取的时间）
        if hasattr(time_value, 'hour') and hasattr(time_value, 'minute') and hasattr(time_value, 'second'):
            try:
                base_date = datetime.strptime(date_label, "%Y-%m-%d")
                return base_date.replace(
                    hour=time_value.hour,
                    minute=time_value.minute,
                    second=time_value.second,
                    microsecond=getattr(time_value, 'microsecond', 0)
                )
            except:
                pass

        # 如果是数字（Excel序列值）
        if isinstance(time_value, (int, float)):
            try:
                # Excel基准日期: 1899-12-30
                base_date = datetime(1899, 12, 30)
                full_datetime = base_date + timedelta(days=float(time_value))
                return full_datetime
            except:
                pass

        # 如果是字符串
        if isinstance(time_value, str):
            time_value = time_value.strip()

            # 格式1: HH:MM:SS.mmm 或 HH:MM:SS
            match = re.match(r'^(\d{1,2}):(\d{2}):(\d{2})(\.\d+)?$', time_value)
            if match:
                h, m, s, ms = match.groups()
                ms = float(ms) if ms else 0
                try:
                    base_date = datetime.strptime(date_label, "%Y-%m-%d")
                    return base_date.replace(hour=int(h), minute=int(m),
                                           second=int(s), microsecond=int(ms*1000000))
                except:
                    pass

            # 格式2: YYYY-MM-DD HH:MM:SS
            try:
                return datetime.strptime(time_value, "%Y-%m-%d %H:%M:%S")
            except:
                try:
                    return datetime.strptime(time_value, "%Y-%m-%d %H:%M:%S.%f")
                except:
                    pass

            # 格式3: Excel日期格式
            try:
                return pd.to_datetime(time_value)
            except:
                pass

        return None

    def _load_and_process_file(self, file_info):
        """加载并处理单个文件"""
        file_path, date_label, color_idx = file_info

        try:
            # 根据文件扩展名选择读取方式
            file_ext = os.path.splitext(file_path)[1].lower()

            # 用于存储调试信息
            debug_info = ""

            if file_ext in ['.xlsx', '.xls']:
                # Excel文件 - 自动选择引擎
                try:
                    if file_ext == '.xlsx':
                        # 尝试使用 openpyxl 引擎
                        df = pd.read_excel(file_path, engine='openpyxl')
                    else:
                        # .xls 文件，尝试不同引擎
                        try:
                            df = pd.read_excel(file_path, engine='xlrd')
                        except:
                            # 如果 xlrd 不可用，尝试 openpyxl（支持 .xls 的某些情况）
                            df = pd.read_excel(file_path)

                    # 记录Excel文件信息用于调试
                    debug_info = f" [列名: {list(df.columns)}, 行数: {len(df)}]"
                    print(f"[调试] 读取Excel文件: {os.path.basename(file_path)}{debug_info}")
                except ImportError as e:
                    return None, f"缺少Excel读取库，请安装: pip install openpyxl xlrd"
                except Exception as e:
                    return None, f"无法读取Excel文件: {os.path.basename(file_path)} - {str(e)}"
            else:
                # CSV/TXT文件 - 尝试不同的编码
                for encoding in ['utf-8', 'gbk', 'gb2312', 'utf-8-sig']:
                    try:
                        df = pd.read_csv(file_path, encoding=encoding)
                        break
                    except:
                        continue
                else:
                    return None, f"无法读取文件编码: {os.path.basename(file_path)}"

            # 查找时间、电压、电流列
            time_col = None
            voltage_col = None
            current_col = None

            # 清理列名：去除前后空格和特殊字符
            clean_columns = {}
            for col in df.columns:
                clean_col = str(col).strip()
                clean_columns[clean_col] = col

            for clean_col, original_col in clean_columns.items():
                col_lower = clean_col.lower()
                # 匹配时间列
                if '时间' in clean_col or 'time' in col_lower:
                    time_col = original_col
                # 匹配电压列
                elif '电压' in clean_col or 'voltage' in col_lower:
                    voltage_col = original_col
                # 匹配电流列
                elif '电流' in clean_col or 'current' in col_lower:
                    current_col = original_col

            print(f"找到的列 - 时间: {time_col}, 电压: {voltage_col}, 电流: {current_col}")

            if time_col is None:
                return None, f"未找到时间列: {os.path.basename(file_path)}，文件列名: {list(df.columns)}"
            if voltage_col is None:
                return None, f"未找到电压列: {os.path.basename(file_path)}，文件列名: {list(df.columns)}"
            if current_col is None:
                return None, f"未找到电流列: {os.path.basename(file_path)}，文件列名: {list(df.columns)}"

            # 处理数据
            data_rows = []
            error_count = 0
            time_error = 0  # 时间解析错误
            value_error = 0  # 电压电流都为空

            # 输出前3行原始数据用于调试
            print(f"\n[调试] 文件 {os.path.basename(file_path)} 前3行原始数据:")
            for i in range(min(3, len(df))):
                sample_row = df.iloc[i]
                print(f"  第{i+1}行: 时间={sample_row[time_col]} (类型:{type(sample_row[time_col]).__name__}), "
                      f"电压={sample_row[voltage_col]}, 电流={sample_row[current_col]}")

            for idx, row in df.iterrows():
                try:
                    time_val = self._parse_time_value(row[time_col], date_label)
                    if time_val is None:
                        time_error += 1
                        # 只打印前3个错误示例
                        if time_error <= 3:
                            print(f"[调试] 第{idx+1}行时间解析失败: 原始值={row[time_col]}, 类型={type(row[time_col]).__name__}")
                        error_count += 1
                        continue

                    # 处理电压和电流值，空值或无效值设为None
                    try:
                        voltage_val = row[voltage_col]
                        if pd.isna(voltage_val) or voltage_val == '' or str(voltage_val).strip() == '':
                            voltage = None
                        else:
                            voltage = float(voltage_val)
                    except:
                        voltage = None

                    try:
                        current_val = row[current_col]
                        if pd.isna(current_val) or current_val == '' or str(current_val).strip() == '':
                            current = None
                        else:
                            current = float(current_val)
                    except:
                        current = None

                    # 如果电压和电流都为空，跳过该行
                    if voltage is None and current is None:
                        value_error += 1
                        # 只打印前3个错误示例
                        if value_error <= 3:
                            print(f"[调试] 第{idx+1}行电压电流都为空: 时间={time_val}, "
                                  f"电压原始值={row[voltage_col]}, 电流原始值={row[current_col]}")
                        error_count += 1
                        continue

                    data_rows.append({
                        'datetime': time_val,
                        'time_str': time_val.strftime("%H:%M:%S"),
                        'voltage': voltage,
                        'current': current
                    })
                except Exception as e:
                    if error_count < 3:
                        print(f"[调试] 第{idx+1}行处理异常: {str(e)}")
                    error_count += 1
                    continue

            print(f"[调试] 处理完成: 有效数据{len(data_rows)}行, "
                  f"时间解析错误{time_error}行, 电压电流为空{value_error}行")

            if not data_rows:
                return None, f"无有效数据: {os.path.basename(file_path)} (共{len(df)}行, 跳过{error_count}行)"

            df_processed = pd.DataFrame(data_rows)
            df_processed = df_processed.sort_values('datetime')

            return df_processed, f"成功: {os.path.basename(file_path)} ({len(data_rows)} 行, {error_count} 行错误){debug_info}"

        except Exception as e:
            return None, f"处理失败: {os.path.basename(file_path)} - {str(e)}"

    def _calculate_smart_ticks(self, axis_max):
        """计算智能刻度值，使曲线在视觉上分离"""
        # 计算合适的刻度间隔
        if axis_max <= 10:
            step = 2
        elif axis_max <= 20:
            step = 5
        elif axis_max <= 50:
            step = 10
        elif axis_max <= 100:
            step = 20
        elif axis_max <= 200:
            step = 50
        else:
            step = 100

        # 生成刻度值
        max_tick = int(axis_max / step) * step + step
        ticks = list(range(0, max_tick + 1, step))

        # 确保刻度不超过最大值太多
        if ticks and ticks[-1] > axis_max * 1.2:
            ticks = [t for t in ticks if t <= axis_max * 1.1]

        return ticks if ticks else None

    def _generate_plotly_chart(self, data_frames):
        """生成交互式HTML图表"""
        if not data_frames:
            return None

        # 获取选择的电流单位（仅用于显示）
        selected_unit = self.selected_current_unit.get()

        # 创建子图（双Y轴）
        fig = make_subplots(
            specs=[[{"secondary_y": True}]]
        )

        # 合并所有数据到两条曲线（电压和电流）
        # 需要在不同日期之间、同一天内的时间空隙处插入None以断开曲线
        all_voltage_data = []
        all_current_data = []
        all_datetimes = []
        all_date_labels = []  # 保存每个数据点对应的日期标签

        # 数据断开���值设置（秒）
        DATA_GAP_THRESHOLD = 3  # 数据间隙超过3秒，认为设备关机，断开曲线
        FILE_GAP_THRESHOLD = 60  # 不同文件之间超过1分钟，断开曲线（可根据实际需求调整）

        # 先按日期排序数据帧
        sorted_data_frames = sorted(data_frames, key=lambda x: x[2])  # 按date_label排序

        for idx, (df, _, date_label, _) in enumerate(sorted_data_frames):
            # 如果不是第一个文件，检查与前一个文件的时间间隙
            if idx > 0 and all_datetimes:
                prev_max_time = max(all_datetimes)
                current_min_time = df['datetime'].min()
                # 计算时间差（秒）
                time_diff_seconds = (current_min_time - prev_max_time).total_seconds()

                # 如果时间差超过阈值，插入None断开曲线
                if time_diff_seconds > FILE_GAP_THRESHOLD:
                    all_datetimes.append(current_min_time - timedelta(seconds=1))
                    all_voltage_data.append(None)
                    all_current_data.append(None)
                    all_date_labels.append('')

            # 处理当前文件内部的时间空隙
            df_sorted = df.sort_values('datetime').reset_index(drop=True)
            prev_datetime = None

            for _, row in df_sorted.iterrows():
                current_datetime = row['datetime']

                # 检查与前一个数据点的时间间隙
                if prev_datetime is not None:
                    time_diff_seconds = (current_datetime - prev_datetime).total_seconds()
                    # 如果时间差超过阈值（设备开关机），插入None断开曲线
                    if time_diff_seconds > DATA_GAP_THRESHOLD:
                        all_datetimes.append(current_datetime - timedelta(milliseconds=1))
                        all_voltage_data.append(None)
                        all_current_data.append(None)
                        all_date_labels.append('')

                all_datetimes.append(current_datetime)
                all_voltage_data.append(row['voltage'])
                all_current_data.append(row['current'])
                all_date_labels.append(date_label)

                prev_datetime = current_datetime

        if not all_datetimes:
            return None

        min_time = min(all_datetimes)
        max_time = max(all_datetimes)

        # 创建组合数据框并排序
        combined_df = pd.DataFrame({
            'datetime': all_datetimes,
            'voltage': all_voltage_data,
            'current': all_current_data,
            'date_label': all_date_labels
        })
        combined_df = combined_df.sort_values('datetime')

        # 只添加两条曲线：电压和电流
        # 电压曲线（左Y轴）- 统一红色
        fig.add_trace(
            go.Scatter(
                x=combined_df['datetime'],
                y=combined_df['voltage'],
                mode='lines',
                name="电压",
                line=dict(color=self.voltage_color, width=2),
                hovertemplate=(
                    '<b>电压</b><br>' +
                    '日期: %{text}<br>' +
                    '时间: %{x|%Y-%m-%d %H:%M:%S}<br>' +
                    '数值: %{y:.2f} V<br>' +
                    '<extra></extra>'
                ),
                text=combined_df['date_label']
            ),
            secondary_y=False
        )

        # 电流曲线（右Y轴）- 统一蓝色
        fig.add_trace(
            go.Scatter(
                x=combined_df['datetime'],
                y=combined_df['current'],
                mode='lines',
                name="电流",
                line=dict(color=self.current_color, width=2),
                hovertemplate=(
                    '<b>电流</b><br>' +
                    '日期: %{text}<br>' +
                    '时间: %{x|%Y-%m-%d %H:%M:%S}<br>' +
                    f'数值: %{{y:.2f}} {selected_unit}<br>' +
                    '<extra></extra>'
                ),
                text=combined_df['date_label']
            ),
            secondary_y=True
        )

        # 格式化时间范围字符串
        start_time_str = min_time.strftime('%Y-%m-%d %H:%M:%S')
        end_time_str = max_time.strftime('%Y-%m-%d %H:%M:%S')

        # 计算电压和电流的最大值，用于设置Y轴范围
        max_voltage = combined_df['voltage'].max()
        max_current = combined_df['current'].max()

        # 智能分离曲线逻辑
        if self.smart_separate_curves.get() and max_voltage > 0 and max_current > 0:
            # 计算电压和电流的比例
            ratio = max_voltage / max_current

            # 如果比例在0.1到10之间，说明数值范围接近，可能导致曲线重叠
            if 0.1 < ratio < 10:
                # 智能分离：让两条曲线在Y轴上错开显示
                # 策略：电压轴留10%空间，电流轴留更多空间（如50%或100%）
                # 这样电流曲线会显示在图表的下半部分，与电压曲线分离

                voltage_max = max_voltage * 1.1  # 电压轴留10%空间
                current_max = max_current * 1.8  # 电流轴留80%空间，使曲线下移

                # 设置自定义刻度，使视觉上更清晰
                # 电压刻度：接近实际数据范围
                voltage_ticks = self._calculate_smart_ticks(voltage_max)
                # 电流刻度：错开的刻度值
                current_ticks = self._calculate_smart_ticks(current_max)
            else:
                # 数值范围差异大，不需要分离
                voltage_max = max_voltage * 1.1 if max_voltage > 0 else 10
                current_max = max_current * 1.1 if max_current > 0 else 100
                voltage_ticks = None
                current_ticks = None
        else:
            # 未启用智能分离，使用原始逻辑
            voltage_max = max_voltage * 1.1 if max_voltage > 0 else 10
            current_max = max_current * 1.1 if max_current > 0 else 100
            voltage_ticks = None
            current_ticks = None

        # 设置X轴 - 严格从数据开始时间到结束时间，不添加额外边距
        fig.update_xaxes(
            title_text="",  # 去掉X轴标题
            tickformat="%H:%M:%S",
            showgrid=True
        )

        # 设置Y轴 - 从0开始
        if voltage_ticks:
            # 使用自定义刻度值
            fig.update_yaxes(
                title_text="电压 (V)",
                secondary_y=False,
                showgrid=True,
                gridcolor='lightgray',
                range=[0, voltage_max],
                tickmode='array',
                tickvals=voltage_ticks
            )
        else:
            # 自动刻度
            fig.update_yaxes(
                title_text="电压 (V)",
                secondary_y=False,
                showgrid=True,
                gridcolor='lightgray',
                range=[0, voltage_max]
            )

        if current_ticks:
            # 使用自定义刻度值
            fig.update_yaxes(
                title_text=f"电流 ({selected_unit})",
                secondary_y=True,
                showgrid=False,
                range=[0, current_max],
                tickmode='array',
                tickvals=current_ticks
            )
        else:
            # 自动刻度
            fig.update_yaxes(
                title_text=f"电流 ({selected_unit})",
                secondary_y=True,
                showgrid=False,
                range=[0, current_max]
            )

        # 更新布局 - 简洁干净
        fig.update_layout(
            title={
                'text': "电压电流对比图",
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 20, 'color': '#333'}
            },
            # 在X轴上方添加开始/结束时间标注
            annotations=[
                {
                    'x': min_time,
                    'y': 1.05,
                    'xref': 'x',
                    'yref': 'paper',
                    'text': f'▶ {start_time_str}',
                    'showarrow': False,
                    'font': {'size': 11, 'color': '#666'},
                    'xanchor': 'left'
                },
                {
                    'x': max_time,
                    'y': 1.05,
                    'xref': 'x',
                    'yref': 'paper',
                    'text': f'{end_time_str} ◀',
                    'showarrow': False,
                    'font': {'size': 11, 'color': '#666'},
                    'xanchor': 'right'
                }
            ],
            # 使用 'x unified' 模式同时显示电压和电流
            hovermode='x unified',
            template='plotly_white',
            height=650,
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="center",
                x=0.5,
                bgcolor='rgba(255,255,255,0.8)',
                bordercolor='#ccc',
                borderwidth=1
            ),
            margin=dict(l=60, r=60, t=100, b=60),  # 增加顶部margin给annotations
            plot_bgcolor='white',
            # 在layout中也设置X轴范围，确保大量数据时也能正确显示
        )

        # 生成HTML
        html_content = fig.to_html(
            include_plotlyjs=True,
            config={
                'displayModeBar': True,
                'displaylogo': False,
                'scrollZoom': True,  # 启用鼠标滚轮缩放
                'dragmode': 'pan',  # 默认拖拽平移模式，更直观
                'modeBarButtonsToRemove': ['lasso2d', 'select2d'],
                'locale': 'zh-CN',  # 设置中文区域
                'toImageButtonOptions': {
                    'format': 'png',
                    'filename': '电压电流对比图',
                    'height': 650,
                    'width': 1200,
                    'scale': 2
                }
            }
        )

        # 添加JavaScript用于动态更新annotations和下载时隐藏滑块
        js_script = """
<script>
(function() {
    // 等待DOM和Plotly完全加载
    function waitForChart() {
        var gd = document.querySelectorAll('.plotly-graph-div')[0];
        if (!gd) {
            setTimeout(waitForChart, 100);
            return;
        }

        // 保存原始状态
        var sliderVisible = true;
        var exportInProgress = false;

        // 监听下载事件，隐藏滑块
        gd.on('plotly_beforeexport', function(eventdata) {
            if (exportInProgress) return;
            exportInProgress = true;

            // 保存当前滑块状态
            sliderVisible = gd.layout.xaxis.rangeslider.visible !== false;

            // 临时隐藏滑块，保持时间轴
            Plotly.relayout(gd, {
                'xaxis.rangeslider.visible': false
            });

            // 等待一小段时间确保布局更新完成
            setTimeout(function() {
                // 触发重绘以确保滑块被隐藏
                Plotly.Plots.resize(gd);
            }, 100);
        });

        // 监听下载完成事件，恢复滑块
        gd.on('plotly_afterexport', function(eventdata) {
            // 恢复滑块显示
            if (sliderVisible) {
                Plotly.relayout(gd, {
                    'xaxis.rangeslider.visible': true
                });
            }

            exportInProgress = false;
        });

        // 格式化日期时间函数
        function formatDateTime(date) {
            var year = date.getFullYear();
            var month = String(date.getMonth() + 1).padStart(2, '0');
            var day = String(date.getDate()).padStart(2, '0');
            var hours = String(date.getHours()).padStart(2, '0');
            var minutes = String(date.getMinutes()).padStart(2, '0');
            var seconds = String(date.getSeconds()).padStart(2, '0');
            return year + '-' + month + '-' + day + ' ' + hours + ':' + minutes + ':' + seconds;
        }

        // 更新annotations中的日期显示
        function updateDateDisplay() {
            var xaxis = gd.layout.xaxis;
            if (xaxis && xaxis.range && xaxis.range[0] && xaxis.range[1]) {
                var startTime = new Date(xaxis.range[0]);
                var endTime = new Date(xaxis.range[1]);

                var startStr = '▶ ' + formatDateTime(startTime);
                var endStr = formatDateTime(endTime) + ' ◀';

                // 更新annotations
                var newAnnotations = gd.layout.annotations.map(function(ann, idx) {
                    if (idx === 0) {
                        return Object.assign({}, ann, {text: startStr, x: xaxis.range[0]});
                    } else if (idx === 1) {
                        return Object.assign({}, ann, {text: endStr, x: xaxis.range[1]});
                    }
                    return ann;
                });

                Plotly.relayout(gd, {
                    annotations: newAnnotations
                });
            }
        }

        // 监听图表重绘事件（包括滑块变化、缩放等）
        gd.on('plotly_relayout', function(eventdata) {
            var hasXAxisChange = eventdata && (
                'xaxis.range[0]' in eventdata ||
                'xaxis.range[1]' in eventdata ||
                'xaxis.range' in eventdata ||
                'xaxis.autorange' in eventdata
            );

            if (hasXAxisChange || eventdata === undefined) {
                updateDateDisplay();
            }
        });

        // 监听滚轮事件，实现实时更新
        gd.on('plotly_wheel', function(eventdata) {
            setTimeout(updateDateDisplay, 100);
        });

        // 初始化时也更新一次
        setTimeout(updateDateDisplay, 500);
    }

    // 开始等待图表加载
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', waitForChart);
    } else {
        waitForChart();
    }
})();
</script>
"""
        # 添加Plotly中文语言包
        locale_script = """
<script src="../plotly-locale-zh-cn.js"></script>
"""
        # 将JavaScript插入到HTML中
        html_content = html_content.replace('</body>', locale_script + js_script + '</body>')

        return html_content

    def _start_processing(self):
        """开始处理数据"""
        if not self.file_list:
            messagebox.showwarning("警告", "请先添加文件！")
            return

        # 在新线程中处理
        thread = Thread(target=self._process_data_thread)
        thread.daemon = True
        thread.start()

    def _process_data_thread(self):
        """数据处理线程"""
        try:
            total = len(self.file_list)
            self.progress['maximum'] = total

            data_frames = []

            for idx, file_info in enumerate(self.file_list):
                file_path = file_info[0]
                self.root.after(0, lambda p=idx+1: self.progress.configure(value=p))
                self.root.after(0, lambda f=file_path: self._log(f"正在处理: {os.path.basename(f)}", "INFO"))

                df, msg = self._load_and_process_file(file_info)

                self.root.after(0, lambda m=msg: self._log(m, "SUCCESS" if "成功" in m else "ERROR"))

                if df is not None:
                    data_frames.append((df, file_info[0], file_info[1], file_info[2]))

            if not data_frames:
                self.root.after(0, lambda: messagebox.showerror("错误", "没有有效的数据可处理！"))
                return

            self.root.after(0, lambda: self._log("正在生成图表...", "INFO"))

            html_content = self._generate_plotly_chart(data_frames)

            if html_content is None:
                self.root.after(0, lambda: messagebox.showerror("错误", "生成图表失败！"))
                return

            # 保存文件到历史文件夹
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            timestamp_display = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            output_path = os.path.join(self.history_folder, f"电压电流对比图_{timestamp}.html")

            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)

            # 保存到历史记录
            self.history_files.append((timestamp_display, output_path))
            self._save_history(timestamp_display, output_path)
            self.root.after(0, lambda: self._refresh_history_list())

            self.root.after(0, lambda p=output_path: self._log(f"图表已生成: {p}", "SUCCESS"))
            self.root.after(0, lambda p=output_path: self._log(f"正在打开浏览器...", "INFO"))

            # 打开浏览器
            self.root.after(0, lambda p=output_path: webbrowser.open(f'file:///{p.replace(os.sep, "/")}'))

            self.root.after(0, lambda: self.progress.configure(value=total))

            # 自动清空文件列表，方便下次处理
            self.root.after(500, self._auto_reset_after_process)

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"处理失败: {str(e)}"))
            self.root.after(0, lambda e=str(e): self._log(f"错误: {e}", "ERROR"))


def main():
    root = tk.Tk()
    app = VoltageCurrentPlotter(root)
    root.mainloop()


if __name__ == "__main__":
    main()