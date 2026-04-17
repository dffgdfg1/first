#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""电压电流曲线分析 - GUI界面"""

import os
import sys
import threading
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path

from analyze_voltage_current import (
    VOLTAGE_SEGMENTS, VOLTAGE_TOLERANCE, MIN_STABLE_POINTS,
    OUTLIER_STD_FACTOR, PROJECT_CONFIGS, VoltageCurrentAnalyzer, ChartScreenshot,
    generate_report,
)


class TextHandler(logging.Handler):
    """将日志重定向到tkinter Text控件"""

    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record) + "\n"
        self.text_widget.after(0, self._append, msg)

    def _append(self, msg):
        self.text_widget.configure(state="normal")
        self.text_widget.insert(tk.END, msg)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state="disabled")


class AnalyzeGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("电压电流曲线自动化分析")
        self.root.geometry("820x660")
        self.root.minsize(720, 560)
        self.running = False
        self.thread = None
        self._last_folder_list = []  # 保存上一轮文件夹列表
        self._build_ui()
        self._setup_logging()

    # ── UI构建 ──────────────────────────────────────────────
    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # === 路径选择区 ===
        path_frame = ttk.LabelFrame(self.root, text="路径设置", padding=8)
        path_frame.pack(fill=tk.X, **pad)

        ttk.Label(path_frame, text="HTML文件夹:").grid(row=0, column=0, sticky=tk.W)
        self.var_html_dir = tk.StringVar(value=str(Path(__file__).parent / "html"))
        ttk.Entry(path_frame, textvariable=self.var_html_dir, width=60).grid(row=0, column=1, sticky=tk.EW, padx=4)
        ttk.Button(path_frame, text="浏览…", command=self._browse_html).grid(row=0, column=2)

        ttk.Label(path_frame, text="输出文件夹:").grid(row=1, column=0, sticky=tk.W, pady=(4, 0))
        self.var_out_dir = tk.StringVar(value=str(Path(__file__).parent / "output"))
        ttk.Entry(path_frame, textvariable=self.var_out_dir, width=60).grid(row=1, column=1, sticky=tk.EW, padx=4, pady=(4, 0))
        ttk.Button(path_frame, text="浏览…", command=self._browse_output).grid(row=1, column=2, pady=(4, 0))
        path_frame.columnconfigure(1, weight=1)

        # === 参数配置区 ===
        param_frame = ttk.LabelFrame(self.root, text="分析参数", padding=8)
        param_frame.pack(fill=tk.X, **pad)

        self.var_tolerance = tk.DoubleVar(value=VOLTAGE_TOLERANCE)
        self.var_min_points = tk.IntVar(value=MIN_STABLE_POINTS)
        self.var_outlier = tk.DoubleVar(value=OUTLIER_STD_FACTOR)  # 默认3.5，更宽松
        self.var_screenshot = tk.BooleanVar(value=True)
        self.var_project = tk.StringVar(value='D2J')

        # 参数说明：电压容差(V) - 匹配目标电压时的允许偏差范围
        #          最小稳定点数 - 判定为有效电压段所需的最少数据点
        #          异常值倍数 - IQR异常值检测的标准差倍数
        params = [
            ("电压容差(V):", self.var_tolerance, 0, "匹配目标电压的允许偏差"),
            ("最小稳定点数:", self.var_min_points, 1, "有效段最少数据点"),
            ("异常值倍数:", self.var_outlier, 2, "IQR异常检测倍数"),
        ]
        for label, var, col, tooltip in params:
            lbl = ttk.Label(param_frame, text=label)
            lbl.grid(row=0, column=col * 2, sticky=tk.W, padx=(8 if col else 0, 2))
            entry = ttk.Entry(param_frame, textvariable=var, width=8)
            entry.grid(row=0, column=col * 2 + 1, padx=(0, 12))
            # 添加工具提示（鼠标悬停显示）
            self._create_tooltip(entry, tooltip)

        ttk.Checkbutton(param_frame, text="截取图表截图", variable=self.var_screenshot).grid(row=0, column=6, padx=8)

        # 批量分析输出路径选项
        self.var_output_to_source = tk.BooleanVar(value=True)
        ttk.Checkbutton(param_frame, text="输出到源文件夹", variable=self.var_output_to_source).grid(row=0, column=7, padx=8)

        # 项目选择
        ttk.Label(param_frame, text="项目类型:").grid(row=1, column=0, sticky=tk.W, padx=(0, 2), pady=(4, 0))
        project_combo = ttk.Combobox(param_frame, textvariable=self.var_project, values=['D2J', 'G2V'],
                                      state='readonly', width=10)
        project_combo.grid(row=1, column=1, sticky=tk.W, pady=(4, 0))
        ttk.Label(param_frame, text="D2J: 6.5V/9V/14V/16V/18V  |  G2V: 9V/14V/16V",
                  font=("", 8), foreground="gray").grid(row=1, column=2, columnspan=5, sticky=tk.W, padx=8, pady=(4, 0))

        # === 模式选择区 ===
        mode_frame = ttk.LabelFrame(self.root, text="分析模式", padding=8)
        mode_frame.pack(fill=tk.X, **pad)

        self.var_mode = tk.StringVar(value="normal")
        ttk.Radiobutton(mode_frame, text="普通模式", variable=self.var_mode, value="normal").pack(side=tk.LEFT, padx=8)
        ttk.Radiobutton(mode_frame, text="P03专用模式", variable=self.var_mode, value="p03").pack(side=tk.LEFT, padx=8)

        # P03输出层级控制
        ttk.Label(mode_frame, text="P03输出层级:", font=("", 9)).pack(side=tk.LEFT, padx=(16, 2))
        self.var_p03_level = tk.IntVar(value=4)
        ttk.Spinbox(mode_frame, from_=1, to=10, textvariable=self.var_p03_level, width=5).pack(side=tk.LEFT, padx=2)

        ttk.Label(mode_frame, text="提示: P03模式请依次上传Rt、Tmax、Tmin三个文件夹（每个包含1-6号样机）",
                  font=("", 8), foreground="gray").pack(side=tk.LEFT, padx=16)

        # === 文件夹列表区（支持直接选择多个HTML文件夹）===
        folder_list_frame = ttk.LabelFrame(self.root, text="批量分析模式（可选，留空则使用默认文件夹模式）", padding=8)
        folder_list_frame.pack(fill=tk.BOTH, expand=True, **pad)

        # 左侧：文件夹列表
        list_container = ttk.Frame(folder_list_frame)
        list_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.folder_listbox = tk.Listbox(list_container, yscrollcommand=scrollbar.set, height=6, selectmode=tk.EXTENDED)
        self.folder_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.folder_listbox.yview)

        # 右侧：操作按钮
        btn_container = ttk.Frame(folder_list_frame)
        btn_container.pack(side=tk.RIGHT, fill=tk.Y, padx=(8, 0))

        ttk.Button(btn_container, text="添加文件夹...", command=self._add_folders, width=12).pack(pady=2)
        ttk.Button(btn_container, text="移除选中", command=self._remove_folders, width=12).pack(pady=2)
        ttk.Button(btn_container, text="清空列表", command=self._clear_folders, width=12).pack(pady=2)
        ttk.Button(btn_container, text="恢复上一轮", command=self._restore_folders, width=12).pack(pady=2)
        ttk.Frame(btn_container, height=10).pack()  # 间隔
        ttk.Label(btn_container, text="提示：\n添加文件夹后将\n按顺序批量分析\n结果保存在output/1,2,3...",
                  font=("", 8), foreground="gray", justify=tk.LEFT).pack(pady=2)

        # === 控制按钮区 ===
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.X, **pad)

        self.btn_start = ttk.Button(btn_frame, text="▶ 开始分析", command=self._start)
        self.btn_start.pack(side=tk.LEFT, padx=4)
        self.btn_stop = ttk.Button(btn_frame, text="■ 停止", command=self._stop, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=4)
        self.btn_open_output = ttk.Button(btn_frame, text="打开输出目录", command=self._open_output)
        self.btn_open_output.pack(side=tk.RIGHT, padx=4)

        # === 进度条 ===
        prog_frame = ttk.Frame(self.root)
        prog_frame.pack(fill=tk.X, **pad)

        self.var_progress_text = tk.StringVar(value="就绪")
        ttk.Label(prog_frame, textvariable=self.var_progress_text).pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(prog_frame, mode="determinate", maximum=100)
        self.progress.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(8, 0))

        # === 日志输出区 ===
        log_frame = ttk.LabelFrame(self.root, text="运行日志", padding=4)
        log_frame.pack(fill=tk.BOTH, expand=True, **pad)

        self.log_text = scrolledtext.ScrolledText(log_frame, state="disabled", wrap=tk.WORD,
                                                   font=("Consolas", 9), height=16)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    # ── 工具提示 ──────────────────────────────────────────
    def _create_tooltip(self, widget, text):
        """为控件添加鼠标悬停提示"""
        def on_enter(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            label = tk.Label(tooltip, text=text, background="lightyellow", relief=tk.SOLID, borderwidth=1, font=("", 8))
            label.pack()
            widget._tooltip = tooltip

        def on_leave(event):
            if hasattr(widget, '_tooltip'):
                widget._tooltip.destroy()
                del widget._tooltip
            _ = event  # 使用event参数避免警告

        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

    # ── 日志 ──────────────────────────────────────────────
    def _setup_logging(self):
        handler = TextHandler(self.log_text)
        handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"))
        root_logger = logging.getLogger()
        root_logger.addHandler(handler)
        root_logger.setLevel(logging.INFO)

    # ── 路径浏览 ──────────────────────────────────────────
    def _browse_html(self):
        d = filedialog.askdirectory(title="选择HTML文件夹")
        if d:
            self.var_html_dir.set(d)

    def _browse_output(self):
        d = filedialog.askdirectory(title="选择输出文件夹")
        if d:
            self.var_out_dir.set(d)

    def _open_output(self):
        out = self.var_out_dir.get()
        if os.path.isdir(out):
            os.startfile(out)
        else:
            messagebox.showwarning("提示", "输出目录不存在")

    # ── 文件夹列表管理 ──────────────────────────────────────
    def _add_folders(self):
        """添加HTML文件夹到列表"""
        folder = filedialog.askdirectory(title="选择包含HTML文件的文件夹")
        if folder and folder not in self.folder_listbox.get(0, tk.END):
            self.folder_listbox.insert(tk.END, folder)
            # 自动选中新添加的项
            last_index = self.folder_listbox.size() - 1
            self.folder_listbox.selection_clear(0, tk.END)
            self.folder_listbox.selection_set(last_index)
            self.folder_listbox.activate(last_index)
            # 滚动到最新添加的项
            self.folder_listbox.see(last_index)

    def _remove_folders(self):
        """移除选中的文件夹"""
        selected = self.folder_listbox.curselection()
        for idx in reversed(selected):
            self.folder_listbox.delete(idx)

    def _clear_folders(self):
        """清空文件夹列表"""
        self.folder_listbox.delete(0, tk.END)

    def _restore_folders(self):
        """恢复上一轮的文件夹列表"""
        if not self._last_folder_list:
            messagebox.showinfo("提示", "没有上一轮的文件夹记录")
            return
        self.folder_listbox.delete(0, tk.END)
        for folder in self._last_folder_list:
            self.folder_listbox.insert(tk.END, folder)

    # ── 日志辅助 ──────────────────────────────────────────
    def _log(self, msg):
        logging.getLogger().info(msg)

    def _set_progress(self, value, text=None):
        self.progress["value"] = value
        if text:
            self.var_progress_text.set(text)

    # ── 启动 / 停止 ──────────────────────────────────────
    def _start(self):
        # 检查是否使用批量文件夹模式
        folder_list = list(self.folder_listbox.get(0, tk.END))

        if folder_list:
            # 批量文件夹模式
            output_dir = Path(self.var_out_dir.get())
            output_dir.mkdir(parents=True, exist_ok=True)
            self._start_batch_folder_mode(folder_list, output_dir)
        else:
            # 默认文件夹模式
            base_dir = Path(self.var_html_dir.get())
            output_dir = Path(self.var_out_dir.get())

            if not base_dir.exists():
                messagebox.showerror("错误", f"HTML文件夹不存在:\n{base_dir}")
                return

            output_dir.mkdir(parents=True, exist_ok=True)
            self._start_folder_mode(base_dir, output_dir)

    def _start_folder_mode(self, base_dir, output_dir):
        """默认文件夹模式启动（1-6号样机）"""
        # 写回参数到模块级变量
        import analyze_voltage_current as avc
        avc.VOLTAGE_TOLERANCE = self.var_tolerance.get()
        avc.MIN_STABLE_POINTS = self.var_min_points.get()
        avc.OUTLIER_STD_FACTOR = self.var_outlier.get()

        self.running = True
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self._set_progress(0, "分析中…")

        self.thread = threading.Thread(target=self._run_analysis, args=(base_dir, output_dir), daemon=True)
        self.thread.start()

    def _start_batch_folder_mode(self, folder_list, output_dir):
        """批量文件夹模式启动"""
        import analyze_voltage_current as avc
        avc.VOLTAGE_TOLERANCE = self.var_tolerance.get()
        avc.MIN_STABLE_POINTS = self.var_min_points.get()
        avc.OUTLIER_STD_FACTOR = self.var_outlier.get()

        self.running = True
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self._set_progress(0, "分析中…")

        # 根据模式选择不同的处理方法
        mode = self.var_mode.get()
        if mode == "p03":
            self.thread = threading.Thread(target=self._run_p03_analysis, args=(folder_list, output_dir), daemon=True)
        else:
            self.thread = threading.Thread(target=self._run_batch_folder_analysis, args=(folder_list, output_dir), daemon=True)
        self.thread.start()

    def _stop(self):
        self.running = False
        self._log("用户请求停止，将在当前样机完成后终止…")

    # ── 后台分析线程 ──────────────────────────────────────
    def _run_analysis(self, base_dir: Path, output_dir: Path):
        """默认文件夹模式分析（1-6号样机）"""
        try:
            self._do_analysis(base_dir, output_dir)
        except Exception as e:
            self._log(f"分析异常: {e}")
        finally:
            self.root.after(0, self._on_finished)

    def _run_batch_folder_analysis(self, folder_list, output_dir: Path):
        """批量文件夹模式分析"""
        try:
            self._do_batch_folder_analysis(folder_list, output_dir)
        except Exception as e:
            self._log(f"分析异常: {e}")
        finally:
            self.root.after(0, self._on_finished)

    def _do_batch_folder_analysis(self, folder_list, output_dir: Path):
        """处理用户选择的多个HTML文件夹（每个文件夹包含1-6号样机子文件夹）"""
        total_folders = len(folder_list)
        output_to_source = self.var_output_to_source.get()

        for folder_idx, html_folder in enumerate(folder_list, 1):
            if not self.running:
                self._log("已停止")
                break

            folder_path = Path(html_folder)
            folder_name = folder_path.name

            # 根据用户选择决定输出路径
            if output_to_source:
                # 输出到源文件夹下的output子目录
                sub_output_dir = folder_path / "output"
            else:
                # 输出到统一output目录下的编号子文件夹
                sub_output_dir = output_dir / str(folder_idx)

            sub_output_dir.mkdir(parents=True, exist_ok=True)

            self._log(f"\n{'='*60}")
            self._log(f"处理文件夹 [{folder_idx}/{total_folders}]: {folder_name}")
            self._log(f"输出目录: {sub_output_dir}")
            self._log(f"{'='*60}")

            # 使用原有的process_all逻辑分析1-6号样机
            self._analyze_folder_with_samples(folder_path, sub_output_dir, folder_idx, total_folders)

        self._log("\n" + "="*60)
        self._log("批量分析完成！")
        self._log("="*60)

    def _analyze_folder_with_samples(self, base_dir: Path, output_dir: Path, folder_idx, total_folders):
        """分析包含1-6号样机子文件夹的根目录"""
        all_results = {}
        screenshot = None
        do_screenshot = self.var_screenshot.get()

        work_screenshot_dir = output_dir / "工作模式截图"
        work_screenshot_dir.mkdir(parents=True, exist_ok=True)
        sleep_screenshot_dir = output_dir / "休眠电流截图"
        sleep_screenshot_dir.mkdir(parents=True, exist_ok=True)

        browser_ok = False
        if do_screenshot and folder_idx == 1:  # 只在第一个文件夹时初始化浏览器
            try:
                screenshot = ChartScreenshot()
                screenshot.init_browser()
                browser_ok = True
                self.screenshot_instance = screenshot
            except Exception:
                self._log("  浏览器初始化失败，跳过截图")
        elif do_screenshot and hasattr(self, 'screenshot_instance'):
            screenshot = self.screenshot_instance
            browser_ok = True

        total_samples = 6
        for i in range(1, total_samples + 1):
            if not self.running:
                self._log("已停止")
                break

            sample_folder = base_dir / str(i)

            # 更新进度
            overall_progress = ((folder_idx - 1) * 100 + (i / total_samples * 100)) / total_folders
            self.root.after(0, self._set_progress, int(overall_progress),
                          f"[文件夹{folder_idx}/{total_folders}] 处理 {i}号样机 ({i}/{total_samples})")

            if not sample_folder.exists():
                self._log(f"  ⚠ {i}号样机文件夹不存在: {sample_folder}")
                continue

            self._log(f"\n  处理 {i}号样机 ({i}/{total_samples})")

            # 获取该样机文件夹下的所有HTML文件
            html_files = sorted([f for f in sample_folder.iterdir() if f.suffix == '.html'])
            if len(html_files) == 0:
                self._log(f"    ⚠ {sample_folder} 中没有HTML文件，跳过")
                continue

            # 通过文件名末尾是否带-l区分：带-l的是休眠模式，不带-l的是工作模式
            work_file = None
            sleep_file = None
            for hf in html_files:
                # 检查文件名（不含扩展名）是否以-l结尾
                name_without_ext = hf.stem  # 例如 "D2J P03-l" 或 "D2J P03"
                if name_without_ext.endswith('-l'):
                    sleep_file = hf  # 文件名末尾带-l → 休眠模式
                else:
                    work_file = hf   # 文件名末尾不带-l → 工作模式

            # 即使没有找到文件，也继续处理（用于截图）
            if not work_file and not sleep_file:
                self._log(f"    ⚠ 未找到工作模式或休眠模式文件")
                continue

            all_results[i] = {}

            # 1. 分析工作模式（单位A）- 从工作模式HTML提取各电压段电流
            if work_file:
                self._log(f"    分析工作模式: {work_file.name}")
                try:
                    project_type = self.var_project.get()
                    analyzer_work = VoltageCurrentAnalyzer(str(work_file), outlier_factor=self.var_outlier.get(),
                                                          project_type=project_type)
                    all_results[i]['work'] = analyzer_work.analyze()
                    all_results[i]['unit_work'] = analyzer_work.current_unit
                except Exception as e:
                    self._log(f"      ⚠ 工作模式分析失败: {e}")

            # 2. 分析休眠电流（单位uA）- 从休眠模式HTML提取14V休眠电流
            if sleep_file:
                self._log(f"    分析休眠电流: {sleep_file.name}")
                try:
                    project_type = self.var_project.get()
                    analyzer_sleep = VoltageCurrentAnalyzer(str(sleep_file), outlier_factor=self.var_outlier.get(),
                                                           project_type=project_type)
                    sleep_result = analyzer_sleep.analyze_sleep()
                    if sleep_result:
                        all_results[i]['sleep'] = sleep_result
                        all_results[i]['unit_sleep'] = analyzer_sleep.current_unit
                        min_i, max_i, avg_i = sleep_result
                        self._log(f"      休眠14V: {min_i:.2f}-{max_i:.2f} (平均:{avg_i:.2f})uA")
                    else:
                        self._log(f"      ⚠ 未检测到休眠电流数据")
                except Exception as e:
                    self._log(f"      ⚠ 休眠电流分析失败: {e}")

            # 3. 截图 - 无论是否检测到电流数据，都要截图
            if browser_ok and screenshot:
                if work_file:
                    work_img = work_screenshot_dir / f"{i}号样机工作电压电流.png"
                    self._log(f"    截取工作模式图表: {work_file.name}")
                    screenshot.capture(str(work_file), str(work_img))
                if sleep_file:
                    sleep_img = sleep_screenshot_dir / f"{i}号样机休眠电流.png"
                    self._log(f"    截取休眠模式图表: {sleep_file.name}")
                    screenshot.capture(str(sleep_file), str(sleep_img))

        # 最后一个文件夹处理完后关闭浏览器
        if folder_idx == total_folders and hasattr(self, 'screenshot_instance'):
            self.screenshot_instance.close()
            del self.screenshot_instance

        # 生成汇总报告
        if all_results and self.running:
            self._log(f"\n  生成汇总报告...")
            project_type = self.var_project.get()
            voltage_segments = PROJECT_CONFIGS[project_type]['voltage_segments']
            generate_report(all_results, output_dir, voltage_segments)
            self._log(f"  ✓ 报告已保存到: {output_dir}")

    def _do_analysis(self, base_dir: Path, output_dir: Path):
        """默认文件夹模式分析"""
        all_results = {}
        screenshot = None
        do_screenshot = self.var_screenshot.get()

        work_screenshot_dir = output_dir / "工作模式截图"
        work_screenshot_dir.mkdir(parents=True, exist_ok=True)
        sleep_screenshot_dir = output_dir / "休眠电流截图"
        sleep_screenshot_dir.mkdir(parents=True, exist_ok=True)

        browser_ok = False
        if do_screenshot:
            try:
                screenshot = ChartScreenshot()
                screenshot.init_browser()
                browser_ok = True
            except Exception:
                self._log("浏览器初始化失败，跳过截图")

        total = 6
        for i in range(1, total + 1):
            if not self.running:
                self._log("已停止")
                break

            pct = int((i - 1) / total * 100)
            self.root.after(0, self._set_progress, pct, f"正在处理 {i}号样机 ({i}/{total})")

            folder = base_dir / str(i)
            if not folder.exists():
                self._log(f"文件夹不存在: {folder}")
                continue

            html_files = sorted([f for f in folder.iterdir() if f.suffix == '.html'])
            if len(html_files) == 0:
                self._log(f"{folder} 中没有HTML文件，跳过")
                continue

            # 通过文件名末尾是否带-l区分：带-l的是休眠模式，不带-l的是工作模式
            work_file = None
            sleep_file = None
            for hf in html_files:
                # 检查文件名（不含扩展名）是否以-l结尾
                name_without_ext = hf.stem  # 例如 "D2J P03-l" 或 "D2J P03"
                if name_without_ext.endswith('-l'):
                    sleep_file = hf  # 文件名末尾带-l → 休眠模式
                else:
                    work_file = hf   # 文件名末尾不带-l → 工作模式

            # 即使没有找到文件，也继续处理（用于截图）
            if not work_file and not sleep_file:
                self._log(f"  ⚠ 未找到工作模式或休眠模式文件")
                continue

            all_results[i] = {}

            # 1. 分析工作模式（单位A）
            if work_file:
                self._log(f"[{i}号样机] 分析工作模式: {work_file.name}")
                try:
                    project_type = self.var_project.get()
                    analyzer_work = VoltageCurrentAnalyzer(str(work_file), outlier_factor=self.var_outlier.get(),
                                                          project_type=project_type)
                    all_results[i]['work'] = analyzer_work.analyze()
                    all_results[i]['unit_work'] = analyzer_work.current_unit
                except Exception as e:
                    self._log(f"  ⚠ 工作模式分析失败: {e}")

            # 2. 分析休眠电流（单位uA）
            if sleep_file:
                self._log(f"[{i}号样机] 分析休眠电流: {sleep_file.name}")
                try:
                    project_type = self.var_project.get()
                    analyzer_sleep = VoltageCurrentAnalyzer(str(sleep_file), outlier_factor=self.var_outlier.get(),
                                                           project_type=project_type)
                    sleep_result = analyzer_sleep.analyze_sleep()
                    if sleep_result:
                        all_results[i]['sleep'] = sleep_result
                        all_results[i]['unit_sleep'] = analyzer_sleep.current_unit
                        min_i, max_i, avg_i = sleep_result
                        self._log(f"  休眠14V: {min_i:.2f}-{max_i:.2f} (平均:{avg_i:.2f})uA")
                    else:
                        self._log(f"  ⚠ 未检测到休眠电流数据")
                except Exception as e:
                    self._log(f"  ⚠ 休眠电流分析失败: {e}")

            # 3. 截图 - 无论是否检测到电流数据，都要截图
            if browser_ok:
                if work_file:
                    work_img = work_screenshot_dir / f"{i}号样机工作电压电流.png"
                    self._log(f"[{i}号样机] 截取工作模式图表: {work_file.name}")
                    screenshot.capture(str(work_file), str(work_img))
                if sleep_file:
                    sleep_img = sleep_screenshot_dir / f"{i}号样机休眠电流.png"
                    self._log(f"[{i}号样机] 截取休眠模式图表: {sleep_file.name}")
                    screenshot.capture(str(sleep_file), str(sleep_img))

        if screenshot and browser_ok:
            screenshot.close()

        # 生成报告
        if all_results and self.running:
            self._log("生成汇总报告…")
            project_type = self.var_project.get()
            voltage_segments = PROJECT_CONFIGS[project_type]['voltage_segments']
            generate_report(all_results, output_dir, voltage_segments)
            self._log(f"报告已保存到: {output_dir}")

    def _run_p03_analysis(self, folder_list, output_dir: Path):
        """P03专用模式分析"""
        self._log("===== 启动P03专用模式分析 =====")
        try:
            self._do_p03_analysis(folder_list)
        except Exception as e:
            self._log(f"分析异常: {e}")
        finally:
            self.root.after(0, self._on_finished)

    def _do_p03_analysis(self, folder_list):
        """处理P03模式的文件夹列表 - 用户依次上传Rt、Tmax、Tmin文件夹"""
        if len(folder_list) % 3 != 0:
            self._log("⚠ P03模式需要3的倍数个文件夹（Rt、Tmax、Tmin为一组）")
            return

        total_groups = len(folder_list) // 3

        for group_idx in range(total_groups):
            if not self.running:
                self._log("已停止")
                break

            # 每组3个文件夹：Rt、Tmax、Tmin
            rt_folder = Path(folder_list[group_idx * 3])
            tmax_folder = Path(folder_list[group_idx * 3 + 1])
            tmin_folder = Path(folder_list[group_idx * 3 + 2])

            # 输出到第一组第一个文件夹的上N级目录下（N由用户设置）
            output_level = self.var_p03_level.get()
            base_output_dir = rt_folder
            for _ in range(output_level):
                base_output_dir = base_output_dir.parent
            group_output_dir = base_output_dir / "output"
            group_output_dir.mkdir(exist_ok=True)

            self._log(f"\n{'='*60}")
            self._log(f"处理P03组 [{group_idx + 1}/{total_groups}]")
            self._log(f"  Rt: {rt_folder.name}")
            self._log(f"  Tmax: {tmax_folder.name}")
            self._log(f"  Tmin: {tmin_folder.name}")
            self._log(f"输出目录: {group_output_dir}")
            self._log(f"{'='*60}")

            # 处理三个温度条件，收集所有结果
            temp_folders = [
                ('Rt', rt_folder),
                ('Tmax', tmax_folder),
                ('Tmin', tmin_folder)
            ]

            all_temp_results = {}  # {temp_name: all_results}

            for temp_idx, (temp_name, temp_folder) in enumerate(temp_folders):
                if not self.running:
                    break
                self._log(f"\n  处理温度条件: {temp_name}")
                temp_output_dir = group_output_dir / temp_name
                temp_output_dir.mkdir(exist_ok=True)

                # 更新进度条
                progress = int(((group_idx * 3 + temp_idx) / (total_groups * 3)) * 100)
                self.root.after(0, self._set_progress, progress, f"P03组{group_idx+1} - {temp_name}")

                # 使用普通模式的分析逻辑（截图+数据分析，但不生成单独Excel）
                results = self._analyze_p03_folder(temp_folder, temp_output_dir, group_idx * 3 + temp_idx + 1, total_groups * 3)
                if results:
                    all_temp_results[temp_name] = results

            # 生成合并的Excel报告
            if all_temp_results and self.running:
                self._log(f"\n  生成P03汇总报告...")
                self._generate_p03_report(all_temp_results, group_output_dir)

        self._log("\n" + "="*60)
        self._log("P03批量分析完成！")
        self._log("="*60)

    def _analyze_p03_folder(self, base_dir: Path, output_dir: Path, folder_idx: int, total_folders: int) -> dict:
        """分析P03某个温度条件下的1-6号样机，返回结果但不生成Excel"""
        all_results = {}
        screenshot = None
        do_screenshot = self.var_screenshot.get()

        work_screenshot_dir = output_dir / "工作模式截图"
        work_screenshot_dir.mkdir(parents=True, exist_ok=True)
        sleep_screenshot_dir = output_dir / "休眠电流截图"
        sleep_screenshot_dir.mkdir(parents=True, exist_ok=True)

        browser_ok = False
        if do_screenshot and folder_idx == 1:
            try:
                screenshot = ChartScreenshot()
                screenshot.init_browser()
                browser_ok = True
                self.screenshot_instance = screenshot
            except Exception:
                self._log("    ⚠ 浏览器初始化失败，跳过截图")
        elif do_screenshot and hasattr(self, 'screenshot_instance'):
            screenshot = self.screenshot_instance
            browser_ok = True

        total = 6
        for i in range(1, total + 1):
            if not self.running:
                break

            folder = base_dir / str(i)
            if not folder.exists():
                continue

            self._log(f"    处理 {i}号样机")

            html_files = sorted([f for f in folder.iterdir() if f.suffix == '.html'])
            if not html_files:
                self._log(f"      ⚠ 无HTML文件，跳过")
                continue

            # 通过文件名末尾是否带-l区分：带-l的是休眠模式，不带-l的是工作模式
            work_file = None
            sleep_file = None
            for hf in html_files:
                # 检查文件名（不含扩展名）是否以-l结尾
                name_without_ext = hf.stem  # 例如 "D2J P03-l" 或 "D2J P03"
                if name_without_ext.endswith('-l'):
                    sleep_file = hf  # 文件名末尾带-l → 休眠模式
                else:
                    work_file = hf   # 文件名末尾不带-l → 工作模式

            all_results[i] = {}

            # 分析工作模式
            if work_file:
                self._log(f"      分析工作模式: {work_file.name}")
                try:
                    project_type = self.var_project.get()
                    analyzer_work = VoltageCurrentAnalyzer(str(work_file), outlier_factor=self.var_outlier.get(),
                                                          project_type=project_type)
                    all_results[i]['work'] = analyzer_work.analyze()
                    all_results[i]['unit_work'] = analyzer_work.current_unit
                except Exception as e:
                    self._log(f"      ⚠ 工作模式分析失败: {e}")

            # 分析休眠电流
            if sleep_file:
                self._log(f"      分析休眠电流: {sleep_file.name}")
                try:
                    project_type = self.var_project.get()
                    analyzer_sleep = VoltageCurrentAnalyzer(str(sleep_file), outlier_factor=self.var_outlier.get(),
                                                           project_type=project_type)
                    sleep_result = analyzer_sleep.analyze_sleep()
                    if sleep_result:
                        all_results[i]['sleep'] = sleep_result
                        all_results[i]['unit_sleep'] = analyzer_sleep.current_unit
                        min_i, max_i, avg_i = sleep_result
                        self._log(f"        休眠14V: {min_i:.2f}-{max_i:.2f} (平均:{avg_i:.2f})uA")
                except Exception as e:
                    self._log(f"      ⚠ 休眠电流分析失败: {e}")

            # 截图 - 无论是否检测到电流数据，都要截图
            if browser_ok and screenshot:
                if work_file:
                    work_img = work_screenshot_dir / f"{i}号样机工作电压电流.png"
                    self._log(f"      截取工作模式图表: {work_file.name}")
                    screenshot.capture(str(work_file), str(work_img))
                if sleep_file:
                    sleep_img = sleep_screenshot_dir / f"{i}号样机休眠电流.png"
                    self._log(f"      截取休眠模式图表: {sleep_file.name}")
                    screenshot.capture(str(sleep_file), str(sleep_img))

        # 最后一个文件夹处理完后关闭浏览器
        if folder_idx == total_folders and hasattr(self, 'screenshot_instance'):
            self.screenshot_instance.close()
            del self.screenshot_instance

        return all_results

    def _generate_p03_report(self, all_temp_results: dict, output_dir: Path):
        """生成P03合并Excel报告，将Rt/Tmax/Tmin数据整合到一个表"""
        import pandas as pd
        from openpyxl.styles import Font

        project_type = self.var_project.get()
        voltage_segments = PROJECT_CONFIGS[project_type]['voltage_segments']

        rows = []
        for temp_name in ['Rt', 'Tmax', 'Tmin']:
            if temp_name not in all_temp_results:
                continue
            results = all_temp_results[temp_name]
            for num in range(1, 7):
                row = {'温度条件': temp_name, '样机编号': f'{num}号样机'}
                for voltage in voltage_segments:
                    if num in results and 'work' in results[num]:
                        data = results[num]['work']
                        if voltage in data:
                            min_i, max_i, avg_i = data[voltage]
                            unit = results[num].get('unit_work', 'A')
                            if unit == 'uA':
                                row[f'{voltage}V'] = f"{min_i:.0f}-{max_i:.0f}{unit}"
                                row[f'{voltage}V平均'] = f"{avg_i:.0f}{unit}"
                            else:
                                row[f'{voltage}V'] = f"{min_i:.2f}-{max_i:.2f}{unit}"
                                row[f'{voltage}V平均'] = f"{avg_i:.2f}{unit}"
                        else:
                            row[f'{voltage}V'] = "N/A"
                            row[f'{voltage}V平均'] = "N/A"
                    else:
                        row[f'{voltage}V'] = "N/A"
                        row[f'{voltage}V平均'] = "N/A"

                # 休眠电流
                if num in results and 'sleep' in results[num]:
                    min_i, max_i, avg_i = results[num]['sleep']
                    unit = results[num].get('unit_sleep', 'A')
                    if unit == 'A':
                        row['休眠电流(14V)'] = f"{min_i*1e6:.2f}-{max_i*1e6:.2f}uA"
                        row['休眠电流(14V)平均'] = f"{avg_i*1e6:.2f}uA"
                    else:
                        row['休眠电流(14V)'] = f"{min_i:.2f}-{max_i:.2f}{unit}"
                        row['休眠电流(14V)平均'] = f"{avg_i:.2f}{unit}"
                else:
                    row['休眠电流(14V)'] = "N/A"
                    row['休眠电流(14V)平均'] = "N/A"

                rows.append(row)

        df = pd.DataFrame(rows)

        output_file = output_dir / "P03电压电流分析汇总.xlsx"
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='工作模式', index=False)
                worksheet = writer.sheets['工作模式']
                arial_font = Font(name='Arial', size=9)
                for row in worksheet.iter_rows():
                    for cell in row:
                        cell.font = arial_font
            self._log(f"  ✓ P03汇总报告已保存: {output_file}")
        except PermissionError:
            import datetime
            ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            alt_file = output_dir / f"P03电压电流分析汇总_{ts}.xlsx"
            with pd.ExcelWriter(alt_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='工作模式', index=False)
                worksheet = writer.sheets['工作模式']
                arial_font = Font(name='Arial', size=9)
                for row in worksheet.iter_rows():
                    for cell in row:
                        cell.font = arial_font
            self._log(f"  ⚠ 原文件被占用，已保存到: {alt_file}")

    def _on_finished(self):
        self.running = False
        self.btn_start.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self._set_progress(100, "完成")
        self._log("===== 分析流程结束 =====")
        # 保存当前列表并清空
        current_list = list(self.folder_listbox.get(0, tk.END))
        if current_list:
            self._last_folder_list = current_list
            self.folder_listbox.delete(0, tk.END)


# ── 入口 ──────────────────────────────────────────────────
def main():
    root = tk.Tk()
    AnalyzeGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
