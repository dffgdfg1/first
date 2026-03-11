#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""文件移动工具 - 根据关键词查找并移动文件"""

import os
import shutil
import threading
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path


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


class MoveFilesGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("文件移动工具 - 关键词查找移动")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        self.running = False
        self.thread = None
        self._build_ui()
        self._setup_logging()

    # ── UI构建 ──────────────────────────────────────────────
    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # === 说明区 ===
        info_frame = ttk.LabelFrame(self.root, text="功能说明", padding=8)
        info_frame.pack(fill=tk.X, **pad)

        info_text = "1. 添加一个或多个一级文件夹（父文件夹）\n" \
                   "2. 输入关键词（支持多个，用逗号分隔）\n" \
                   "3. 软件会在每个一级文件夹的二级文件夹中查找包含关键词的文件夹\n" \
                   "4. 将匹配的三级文件复制/移动到二级文件夹"
        ttk.Label(info_frame, text=info_text, foreground="blue", font=("", 9), justify=tk.LEFT).pack(anchor=tk.W)

        # === 路径选择区 ===
        path_frame = ttk.LabelFrame(self.root, text="一级文件夹列表（父文件夹）", padding=8)
        path_frame.pack(fill=tk.BOTH, expand=True, **pad)

        # 左侧：文件夹列表
        list_container = ttk.Frame(path_frame)
        list_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar_folders = ttk.Scrollbar(list_container)
        scrollbar_folders.pack(side=tk.RIGHT, fill=tk.Y)

        self.folder_listbox = tk.Listbox(list_container, yscrollcommand=scrollbar_folders.set,
                                         height=6, selectmode=tk.EXTENDED)
        self.folder_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_folders.config(command=self.folder_listbox.yview)

        # 右侧：操作按钮
        btn_container = ttk.Frame(path_frame)
        btn_container.pack(side=tk.RIGHT, fill=tk.Y, padx=(8, 0))

        ttk.Button(btn_container, text="添加文件夹...", command=self._add_folder, width=14).pack(pady=2)
        ttk.Button(btn_container, text="批量添加...", command=self._add_folders_batch, width=14).pack(pady=2)
        ttk.Button(btn_container, text="移除选中", command=self._remove_folders, width=14).pack(pady=2)
        ttk.Button(btn_container, text="清空列表", command=self._clear_folders, width=14).pack(pady=2)

        # === 关键词输入区 ===
        keyword_frame = ttk.LabelFrame(self.root, text="关键词设置", padding=8)
        keyword_frame.pack(fill=tk.X, **pad)

        ttk.Label(keyword_frame, text="关键词（多个用逗号分隔）:").pack(side=tk.LEFT, padx=(0, 4))
        self.var_keywords = tk.StringVar()
        ttk.Entry(keyword_frame, textvariable=self.var_keywords, width=60).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # === 选项配置区 ===
        option_frame = ttk.LabelFrame(self.root, text="操作选项", padding=8)
        option_frame.pack(fill=tk.X, **pad)

        self.var_operation = tk.StringVar(value="copy")
        self.var_overwrite = tk.BooleanVar(value=False)
        self.var_delete_empty = tk.BooleanVar(value=True)
        self.var_html_only = tk.BooleanVar(value=False)
        self.var_case_sensitive = tk.BooleanVar(value=False)

        ttk.Radiobutton(option_frame, text="复制文件", variable=self.var_operation,
                       value="copy").grid(row=0, column=0, sticky=tk.W, padx=8, pady=2)
        ttk.Radiobutton(option_frame, text="移动文件", variable=self.var_operation,
                       value="move").grid(row=0, column=1, sticky=tk.W, padx=8, pady=2)

        ttk.Checkbutton(option_frame, text="覆盖同名文件",
                       variable=self.var_overwrite).grid(row=0, column=2, sticky=tk.W, padx=8, pady=2)
        ttk.Checkbutton(option_frame, text="删除空文件夹",
                       variable=self.var_delete_empty).grid(row=0, column=3, sticky=tk.W, padx=8, pady=2)

        ttk.Checkbutton(option_frame, text="仅处理HTML文件",
                       variable=self.var_html_only).grid(row=1, column=0, sticky=tk.W, padx=8, pady=2)
        ttk.Checkbutton(option_frame, text="关键词区分大小写",
                       variable=self.var_case_sensitive).grid(row=1, column=1, sticky=tk.W, padx=8, pady=2)

        # === 预览区 ===
        preview_frame = ttk.LabelFrame(self.root, text="匹配预览（点击'扫描预览'查看将要处理的文件夹）", padding=8)
        preview_frame.pack(fill=tk.BOTH, expand=True, **pad)

        # 预览列表
        preview_container = ttk.Frame(preview_frame)
        preview_container.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(preview_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.preview_listbox = tk.Listbox(preview_container, yscrollcommand=scrollbar.set, height=8)
        self.preview_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.preview_listbox.yview)

        # 预览按钮
        preview_btn_frame = ttk.Frame(preview_frame)
        preview_btn_frame.pack(fill=tk.X, pady=(4, 0))

        ttk.Button(preview_btn_frame, text="🔍 扫描预览", command=self._scan_preview).pack(side=tk.LEFT, padx=2)
        self.lbl_preview_count = ttk.Label(preview_btn_frame, text="", foreground="green")
        self.lbl_preview_count.pack(side=tk.LEFT, padx=8)

        # === 控制按钮区 ===
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.X, **pad)

        self.btn_start = ttk.Button(btn_frame, text="▶ 开始处理", command=self._start)
        self.btn_start.pack(side=tk.LEFT, padx=4)
        self.btn_stop = ttk.Button(btn_frame, text="■ 停止", command=self._stop, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=4)

        ttk.Label(btn_frame, text="  ", font=("", 1)).pack(side=tk.LEFT)
        self.lbl_stats = ttk.Label(btn_frame, text="", foreground="green", font=("", 9))
        self.lbl_stats.pack(side=tk.LEFT, padx=8)

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
                                                   font=("Consolas", 9), height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    # ── 日志 ──────────────────────────────────────────────
    def _setup_logging(self):
        handler = TextHandler(self.log_text)
        handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                                              datefmt="%H:%M:%S"))
        root_logger = logging.getLogger()
        root_logger.addHandler(handler)
        root_logger.setLevel(logging.INFO)

    def _log(self, msg):
        logging.getLogger().info(msg)

    def _set_progress(self, value, text=None):
        self.progress["value"] = value
        if text:
            self.var_progress_text.set(text)

    # ── 路径浏览 ──────────────────────────────────────────
    def _add_folder(self):
        """添加单个一级文件夹"""
        folder = filedialog.askdirectory(title="选择一级文件夹（父文件夹）")
        if folder and folder not in self.folder_listbox.get(0, tk.END):
            self.folder_listbox.insert(tk.END, folder)
            self._log(f"已添加一级文件夹: {folder}")

    def _add_folders_batch(self):
        """批量添加多个一级文件夹"""
        parent_folder = filedialog.askdirectory(title="选择包含多个一级文件夹的父文件夹")
        if not parent_folder:
            return

        parent_path = Path(parent_folder)
        subfolders = [str(f) for f in parent_path.iterdir() if f.is_dir()]

        if not subfolders:
            messagebox.showinfo("提示", "该文件夹下没有子文件夹")
            return

        added_count = 0
        existing_folders = self.folder_listbox.get(0, tk.END)
        for folder in subfolders:
            if folder not in existing_folders:
                self.folder_listbox.insert(tk.END, folder)
                added_count += 1

        self._log(f"批量添加了 {added_count} 个一级文件夹")

    def _remove_folders(self):
        """移除选中的文件夹"""
        selected = self.folder_listbox.curselection()
        if not selected:
            messagebox.showinfo("提示", "请先选择要移除的文件夹")
            return
        for idx in reversed(selected):
            self.folder_listbox.delete(idx)

    def _clear_folders(self):
        """清空文件夹列表"""
        if self.folder_listbox.size() == 0:
            return
        if messagebox.askyesno("确认", "确定要清空所有文件夹吗？"):
            self.folder_listbox.delete(0, tk.END)

    # ── 扫描预览 ──────────────────────────────────────────
    def _scan_preview(self):
        """扫描并预览将要处理的文件夹"""
        folder_list = list(self.folder_listbox.get(0, tk.END))
        keywords_str = self.var_keywords.get().strip()

        if not folder_list:
            messagebox.showwarning("提示", "请先添加一级文件夹")
            return

        if not keywords_str:
            messagebox.showwarning("提示", "请输入关键词")
            return

        # 清空预览列表
        self.preview_listbox.delete(0, tk.END)

        # 解析关键词
        keywords = [k.strip() for k in keywords_str.split(',') if k.strip()]
        case_sensitive = self.var_case_sensitive.get()

        self._log("="*60)
        self._log(f"开始扫描 {len(folder_list)} 个一级文件夹")
        self._log(f"关键词: {keywords}")
        self._log("="*60)

        total_matched = 0

        # 遍历每个一级文件夹
        for parent_folder in folder_list:
            parent_path = Path(parent_folder)

            if not parent_path.exists():
                self._log(f"⚠ 文件夹不存在: {parent_folder}")
                continue

            self._log(f"\n扫描: {parent_path.name}")

            # 遍历二级文件夹
            try:
                for level2_folder in parent_path.iterdir():
                    if not level2_folder.is_dir():
                        continue

                    # 在二级文件夹中查找包含关键词的三级文件夹
                    for level3_folder in level2_folder.iterdir():
                        if not level3_folder.is_dir():
                            continue

                        folder_name = level3_folder.name
                        folder_name_check = folder_name if case_sensitive else folder_name.lower()

                        # 检查是否包含任一关键词
                        for keyword in keywords:
                            keyword_check = keyword if case_sensitive else keyword.lower()
                            if keyword_check in folder_name_check:
                                total_matched += 1
                                display_text = f"[{parent_path.name}] {level2_folder.name} ← {level3_folder.name} (匹配: {keyword})"
                                self.preview_listbox.insert(tk.END, display_text)
                                self._log(f"  匹配: {level2_folder.name}/{level3_folder.name}")
                                break

            except Exception as e:
                self._log(f"  扫描出错: {e}")

        self.lbl_preview_count.config(text=f"找到 {total_matched} 个匹配的文件夹")
        self._log(f"\n扫描完成，共找到 {total_matched} 个匹配的文件夹")

        if total_matched == 0:
            messagebox.showinfo("提示", "未找到匹配的文件夹")

    # ── 启动 / 停止 ──────────────────────────────────────
    def _start(self):
        folder_list = list(self.folder_listbox.get(0, tk.END))
        keywords_str = self.var_keywords.get().strip()

        if not folder_list:
            messagebox.showwarning("提示", "请先添加一级文件夹")
            return

        if not keywords_str:
            messagebox.showwarning("提示", "请输入关键词")
            return

        self.running = True
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self._set_progress(0, "处理中...")
        self.lbl_stats.config(text="")

        keywords = [k.strip() for k in keywords_str.split(',') if k.strip()]
        self.thread = threading.Thread(target=self._run_process, args=(folder_list, keywords), daemon=True)
        self.thread.start()

    def _stop(self):
        self.running = False
        self._log("用户请求停止，将在当前操作完成后终止...")

    # ── 后台处理线程 ──────────────────────────────────────
    def _run_process(self, folder_list, keywords):
        """后台线程执行文件处理"""
        try:
            self._do_process(folder_list, keywords)
        except Exception as e:
            self._log(f"处理异常: {e}")
            import traceback
            self._log(traceback.format_exc())
        finally:
            self.root.after(0, self._on_finished)

    def _do_process(self, folder_list, keywords):
        """执行文件处理操作"""
        operation = self.var_operation.get()
        overwrite = self.var_overwrite.get()
        delete_empty = self.var_delete_empty.get()
        html_only = self.var_html_only.get()
        case_sensitive = self.var_case_sensitive.get()

        total_processed = 0
        total_errors = 0

        self._log("="*60)
        self._log(f"开始处理 {len(folder_list)} 个一级文件夹")
        self._log(f"关键词: {keywords}")
        self._log(f"操作: {'复制' if operation == 'copy' else '移动'}")
        self._log(f"选项: 覆盖={overwrite}, 删除空文件夹={delete_empty}, 仅HTML={html_only}, 区分大小写={case_sensitive}")
        self._log("="*60)

        # 遍历每个一级文件夹
        for folder_idx, parent_folder in enumerate(folder_list, 1):
            if not self.running:
                self._log("已停止")
                break

            parent_path = Path(parent_folder)

            if not parent_path.exists():
                self._log(f"\n⚠ 文件夹不存在: {parent_folder}")
                continue

            self._log(f"\n{'='*60}")
            self._log(f"[{folder_idx}/{len(folder_list)}] 处理一级文件夹: {parent_path.name}")
            self._log(f"{'='*60}")

            matched_folders = []

            # 第一步：收集该一级文件夹下所有匹配的文件夹
            for level2_folder in parent_path.iterdir():
                if not level2_folder.is_dir():
                    continue

                for level3_folder in level2_folder.iterdir():
                    if not level3_folder.is_dir():
                        continue

                    folder_name = level3_folder.name
                    folder_name_check = folder_name if case_sensitive else folder_name.lower()

                    for keyword in keywords:
                        keyword_check = keyword if case_sensitive else keyword.lower()
                        if keyword_check in folder_name_check:
                            matched_folders.append((level2_folder, level3_folder))
                            break

            self._log(f"找到 {len(matched_folders)} 个匹配的文件夹")

            if len(matched_folders) == 0:
                self._log("没有找到匹配的文件夹，跳过")
                continue

            # 第二步：处理每个匹配的文件夹
            for idx, (level2_folder, level3_folder) in enumerate(matched_folders, 1):
                if not self.running:
                    self._log("已停止")
                    break

                self._log(f"\n  [{idx}/{len(matched_folders)}] 处理: {level2_folder.name}/{level3_folder.name}")

                # 更新进度
                overall_progress = ((folder_idx - 1) * 100 + (idx / len(matched_folders) * 100)) / len(folder_list)
                self.root.after(0, self._set_progress, int(overall_progress),
                              f"[{folder_idx}/{len(folder_list)}] 处理中: {level3_folder.name}")

                try:
                    processed, errors = self._process_folder(
                        level3_folder, level2_folder, operation, overwrite, html_only
                    )
                    total_processed += processed
                    total_errors += errors

                    self._log(f"    完成: {operation} {processed} 个文件, {errors} 个错误")

                    # 如果是移动操作且设置了删除空文件夹
                    if operation == "move" and delete_empty:
                        try:
                            remaining = list(level3_folder.iterdir())
                            if not remaining:
                                level3_folder.rmdir()
                                self._log(f"    已删除空文件夹: {level3_folder.name}")
                        except Exception as e:
                            self._log(f"    删除文件夹失败: {e}")

                except Exception as e:
                    self._log(f"    错误: {e}")
                    total_errors += 1

        self._log("\n" + "="*60)
        self._log(f"全部完成! 共处理 {total_processed} 个文件, {total_errors} 个错误")
        self._log("="*60)

        # 更新统计信息
        stats_text = f"处理: {total_processed} 个文件  |  错误: {total_errors} 个"
        self.root.after(0, lambda: self.lbl_stats.config(text=stats_text))

    def _process_folder(self, source_folder, target_folder, operation, overwrite, html_only):
        """处理单个文件夹：将三级文件复制/移动到二级文件夹"""
        processed_count = 0
        error_count = 0

        # 获取源文件夹中的文件
        if html_only:
            files = [f for f in source_folder.iterdir() if f.is_file() and f.suffix.lower() in ['.html', '.htm']]
        else:
            files = [f for f in source_folder.iterdir() if f.is_file()]

        self._log(f"  找到 {len(files)} 个文件")

        for file in files:
            if not self.running:
                break

            target_path = target_folder / file.name

            try:
                # 检查目标文件是否存在
                if target_path.exists():
                    if overwrite:
                        self._log(f"    覆盖: {file.name}")
                    else:
                        # 生成新文件名
                        base_name = file.stem
                        extension = file.suffix
                        counter = 1
                        while target_path.exists():
                            new_name = f"{base_name}_{counter}{extension}"
                            target_path = target_folder / new_name
                            counter += 1
                        self._log(f"    重命名: {file.name} → {target_path.name}")

                # 执行复制或移动
                if operation == "copy":
                    shutil.copy2(str(file), str(target_path))
                else:  # move
                    shutil.move(str(file), str(target_path))

                processed_count += 1

            except Exception as e:
                self._log(f"    失败 {file.name}: {e}")
                error_count += 1

        return processed_count, error_count

    def _on_finished(self):
        """处理完成后的清理工作"""
        self.running = False
        self.btn_start.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self._set_progress(100, "完成")


# ── 入口 ──────────────────────────────────────────────────
def main():
    root = tk.Tk()
    MoveFilesGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
