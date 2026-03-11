#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""图片批量缩放工具 - GUI界面"""

import os
import sys
import threading
from pathlib import Path
from typing import List, Tuple
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from PIL import Image, ImageTk
import datetime


class ImageResizer:
    """图片缩放处理类"""

    RESIZE_MODES = {
        "保持比例(适应)": "fit",
        "保持比例(填充)": "fill",
        "拉伸": "stretch",
        "裁剪中心": "crop"
    }

    def __init__(self, log_callback=None):
        self.log_callback = log_callback

    def log(self, msg):
        """输出日志"""
        if self.log_callback:
            self.log_callback(msg)

    def resize_image(self, input_path: str, output_path: str,
                    width: int, height: int, mode: str = "fit",
                    quality: int = 95, format: str = None) -> bool:
        """
        缩放单张图片

        Args:
            input_path: 输入图片路径
            output_path: 输出图片路径
            width: 目标宽度
            height: 目标高度
            mode: 缩放模式 (fit/fill/stretch/crop)
            quality: 输出质量 (1-100)
            format: 输出格式 (None表示保持原格式)

        Returns:
            是否成功
        """
        try:
            img = Image.open(input_path)
            original_format = img.format

            # 转换RGBA到RGB (如果输出格式不支持透明度)
            if img.mode == 'RGBA' and (format == 'JPEG' or (format is None and original_format == 'JPEG')):
                background = Image.new('RGB', img.size, (255, 255, 255))
                background.paste(img, mask=img.split()[3])
                img = background

            target_size = (width, height)

            if mode == "fit":
                # 保持比例，适应目标尺寸（可能有留白）
                img.thumbnail(target_size, Image.Resampling.LANCZOS)
                # 创建目标尺寸的画布，居中放置
                result = Image.new('RGB', target_size, (255, 255, 255))
                offset = ((target_size[0] - img.size[0]) // 2,
                         (target_size[1] - img.size[1]) // 2)
                result.paste(img, offset)
                img = result

            elif mode == "fill":
                # 保持比例，填充目标尺寸（可能裁剪）
                img_ratio = img.size[0] / img.size[1]
                target_ratio = width / height

                if img_ratio > target_ratio:
                    # 图片更宽，按高度缩放
                    new_height = height
                    new_width = int(height * img_ratio)
                else:
                    # 图片更高，按宽度缩放
                    new_width = width
                    new_height = int(width / img_ratio)

                img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

                # 裁剪到目标尺寸
                left = (new_width - width) // 2
                top = (new_height - height) // 2
                img = img.crop((left, top, left + width, top + height))

            elif mode == "stretch":
                # 直接拉伸到目标尺寸
                img = img.resize(target_size, Image.Resampling.LANCZOS)

            elif mode == "crop":
                # 从中心裁剪
                left = (img.size[0] - width) // 2
                top = (img.size[1] - height) // 2
                right = left + width
                bottom = top + height

                # 如果目标尺寸大于原图，先放大
                if width > img.size[0] or height > img.size[1]:
                    scale = max(width / img.size[0], height / img.size[1])
                    new_size = (int(img.size[0] * scale), int(img.size[1] * scale))
                    img = img.resize(new_size, Image.Resampling.LANCZOS)
                    left = (img.size[0] - width) // 2
                    top = (img.size[1] - height) // 2
                    right = left + width
                    bottom = top + height

                img = img.crop((left, top, right, bottom))

            # 保存图片
            save_format = format if format else original_format
            if save_format == 'JPEG':
                img.save(output_path, format=save_format, quality=quality, optimize=True)
            else:
                img.save(output_path, format=save_format, quality=quality)

            return True

        except Exception as e:
            self.log(f"  ✗ 处理失败: {e}")
            return False


class ImageResizerGUI:
    """图片缩放工具GUI"""

    def __init__(self, root):
        self.root = root
        self.root.title("图片批量缩放工具")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)

        self.running = False
        self.thread = None
        self.image_files = []
        self.preview_image = None

        self._build_ui()

    def _build_ui(self):
        """构建界面"""
        pad = {"padx": 8, "pady": 4}

        # === 图片选择区 ===
        file_frame = ttk.LabelFrame(self.root, text="图片文件", padding=8)
        file_frame.pack(fill=tk.BOTH, expand=True, **pad)

        # 左侧：文件列表
        list_container = ttk.Frame(file_frame)
        list_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_listbox = tk.Listbox(list_container, yscrollcommand=scrollbar.set,
                                       height=8, selectmode=tk.EXTENDED)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.file_listbox.bind('<<ListboxSelect>>', self._on_file_select)
        scrollbar.config(command=self.file_listbox.yview)

        # 右侧：操作按钮和预览
        right_container = ttk.Frame(file_frame)
        right_container.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(8, 0))

        btn_container = ttk.Frame(right_container)
        btn_container.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(btn_container, text="添加图片...", command=self._add_images, width=12).pack(pady=2)
        ttk.Button(btn_container, text="添加文件夹...", command=self._add_folder, width=12).pack(pady=2)
        ttk.Button(btn_container, text="移除选中", command=self._remove_images, width=12).pack(pady=2)
        ttk.Button(btn_container, text="清空列表", command=self._clear_images, width=12).pack(pady=2)

        # 预览区
        preview_frame = ttk.LabelFrame(right_container, text="预览", padding=4)
        preview_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(8, 0))

        self.preview_label = tk.Label(preview_frame, text="选择图片查看预览",
                                      bg="gray90", width=20, height=10)
        self.preview_label.pack(fill=tk.BOTH, expand=True)

        # === 尺寸设置区 ===
        size_frame = ttk.LabelFrame(self.root, text="目标尺寸", padding=8)
        size_frame.pack(fill=tk.X, **pad)

        # 第一行：宽度和高度
        row1 = ttk.Frame(size_frame)
        row1.pack(fill=tk.X, pady=2)

        ttk.Label(row1, text="宽度(px):").pack(side=tk.LEFT, padx=(0, 4))
        self.var_width = tk.IntVar(value=800)
        ttk.Entry(row1, textvariable=self.var_width, width=10).pack(side=tk.LEFT, padx=(0, 16))

        ttk.Label(row1, text="高度(px):").pack(side=tk.LEFT, padx=(0, 4))
        self.var_height = tk.IntVar(value=600)
        ttk.Entry(row1, textvariable=self.var_height, width=10).pack(side=tk.LEFT, padx=(0, 16))

        # 常用尺寸快捷按钮
        ttk.Label(row1, text="快捷:").pack(side=tk.LEFT, padx=(8, 4))
        presets = [
            ("1920×1080", 1920, 1080),
            ("1280×720", 1280, 720),
            ("800×600", 800, 600),
            ("640×480", 640, 480)
        ]
        for label, w, h in presets:
            ttk.Button(row1, text=label, width=10,
                      command=lambda w=w, h=h: self._set_size(w, h)).pack(side=tk.LEFT, padx=2)

        # === 缩放模式区 ===
        mode_frame = ttk.LabelFrame(self.root, text="缩放模式", padding=8)
        mode_frame.pack(fill=tk.X, **pad)

        self.var_mode = tk.StringVar(value="保持比例(适应)")
        mode_row = ttk.Frame(mode_frame)
        mode_row.pack(fill=tk.X)

        for mode_name in ImageResizer.RESIZE_MODES.keys():
            ttk.Radiobutton(mode_row, text=mode_name, variable=self.var_mode,
                          value=mode_name).pack(side=tk.LEFT, padx=8)

        # 模式说明
        mode_desc = {
            "保持比例(适应)": "保持原图比例，按长边或宽边中较小的比例缩放，完整显示图片，不足部分留白",
            "保持比例(填充)": "保持原图比例，按长边或宽边中较大的比例缩放，填满目标尺寸，超出部分裁剪",
            "拉伸": "直接拉伸到目标尺寸，不保持比例，可能变形",
            "裁剪中心": "从图片中心裁剪到目标尺寸，不缩放（如原图小于目标则先放大）"
        }
        self.mode_desc_label = ttk.Label(mode_frame, text=mode_desc["保持比例(适应)"],
                                        font=("", 8), foreground="gray")
        self.mode_desc_label.pack(pady=(4, 0))

        # 绑定模式变化事件
        self.var_mode.trace('w', lambda *args: self.mode_desc_label.config(
            text=mode_desc.get(self.var_mode.get(), "")))

        # === 输出设置区 ===
        output_frame = ttk.LabelFrame(self.root, text="输出设置", padding=8)
        output_frame.pack(fill=tk.X, **pad)

        row2 = ttk.Frame(output_frame)
        row2.pack(fill=tk.X, pady=2)

        ttk.Label(row2, text="输出目录:").pack(side=tk.LEFT, padx=(0, 4))
        self.var_output_dir = tk.StringVar(value=str(Path(__file__).parent / "resized_images"))
        ttk.Entry(row2, textvariable=self.var_output_dir, width=50).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(row2, text="浏览...", command=self._browse_output).pack(side=tk.LEFT)

        row3 = ttk.Frame(output_frame)
        row3.pack(fill=tk.X, pady=(4, 2))

        ttk.Label(row3, text="输出格式:").pack(side=tk.LEFT, padx=(0, 4))
        self.var_format = tk.StringVar(value="保持原格式")
        format_combo = ttk.Combobox(row3, textvariable=self.var_format, width=12,
                                   values=["保持原格式", "JPEG", "PNG", "WEBP"], state="readonly")
        format_combo.pack(side=tk.LEFT, padx=(0, 16))

        ttk.Label(row3, text="图片质量:").pack(side=tk.LEFT, padx=(0, 4))
        self.var_quality = tk.IntVar(value=95)
        quality_spin = ttk.Spinbox(row3, from_=1, to=100, textvariable=self.var_quality, width=8)
        quality_spin.pack(side=tk.LEFT, padx=(0, 16))

        self.var_keep_structure = tk.BooleanVar(value=False)
        ttk.Checkbutton(row3, text="保持文件夹结构", variable=self.var_keep_structure).pack(side=tk.LEFT, padx=8)

        self.var_overwrite = tk.BooleanVar(value=False)
        ttk.Checkbutton(row3, text="覆盖已存在文件", variable=self.var_overwrite).pack(side=tk.LEFT, padx=8)

        # === 控制按钮区 ===
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.X, **pad)

        self.btn_start = ttk.Button(btn_frame, text="▶ 开始处理", command=self._start)
        self.btn_start.pack(side=tk.LEFT, padx=4)
        self.btn_stop = ttk.Button(btn_frame, text="■ 停止", command=self._stop, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=4)
        self.btn_open_output = ttk.Button(btn_frame, text="打开输出目录", command=self._open_output)
        self.btn_open_output.pack(side=tk.RIGHT, padx=4)

        # 统计信息
        self.var_stats = tk.StringVar(value="已选择: 0 张图片")
        ttk.Label(btn_frame, textvariable=self.var_stats).pack(side=tk.LEFT, padx=16)

        # === 进度条 ===
        prog_frame = ttk.Frame(self.root)
        prog_frame.pack(fill=tk.X, **pad)

        self.var_progress_text = tk.StringVar(value="就绪")
        ttk.Label(prog_frame, textvariable=self.var_progress_text).pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(prog_frame, mode="determinate", maximum=100)
        self.progress.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(8, 0))

        # === 日志输出区 ===
        log_frame = ttk.LabelFrame(self.root, text="处理日志", padding=4)
        log_frame.pack(fill=tk.BOTH, expand=True, **pad)

        self.log_text = scrolledtext.ScrolledText(log_frame, state="disabled", wrap=tk.WORD,
                                                  font=("Consolas", 9), height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    # ── 辅助方法 ──────────────────────────────────────────
    def _set_size(self, width, height):
        """设置快捷尺寸"""
        self.var_width.set(width)
        self.var_height.set(height)

    def _browse_output(self):
        """浏览输出目录"""
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.var_output_dir.set(d)

    def _open_output(self):
        """打开输出目录"""
        out = self.var_output_dir.get()
        if os.path.isdir(out):
            os.startfile(out)
        else:
            messagebox.showwarning("提示", "输出目录不存在")

    def _add_images(self):
        """添加图片文件"""
        files = filedialog.askopenfilenames(
            title="选择图片文件",
            filetypes=[
                ("图片文件", "*.jpg *.jpeg *.png *.bmp *.gif *.webp *.tiff"),
                ("所有文件", "*.*")
            ]
        )
        for f in files:
            if f not in self.image_files:
                self.image_files.append(f)
                self.file_listbox.insert(tk.END, Path(f).name)
        self._update_stats()

    def _add_folder(self):
        """添加文件夹中的所有图片"""
        folder = filedialog.askdirectory(title="选择包含图片的文件夹")
        if not folder:
            return

        extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp', '.tiff'}
        folder_path = Path(folder)

        count = 0
        for file_path in folder_path.rglob('*'):
            if file_path.suffix.lower() in extensions:
                file_str = str(file_path)
                if file_str not in self.image_files:
                    self.image_files.append(file_str)
                    # 显示相对路径
                    try:
                        rel_path = file_path.relative_to(folder_path)
                        self.file_listbox.insert(tk.END, str(rel_path))
                    except:
                        self.file_listbox.insert(tk.END, file_path.name)
                    count += 1

        self._update_stats()
        self._log(f"从文件夹添加了 {count} 张图片")

    def _remove_images(self):
        """移除选中的图片"""
        selected = self.file_listbox.curselection()
        for idx in reversed(selected):
            del self.image_files[idx]
            self.file_listbox.delete(idx)
        self._update_stats()

    def _clear_images(self):
        """清空图片列表"""
        self.image_files.clear()
        self.file_listbox.delete(0, tk.END)
        self._update_stats()
        self.preview_label.config(image='', text="选择图片查看预览")
        self.preview_image = None

    def _update_stats(self):
        """更新统计信息"""
        count = len(self.image_files)
        self.var_stats.set(f"已选择: {count} 张图片")

    def _on_file_select(self, event=None):
        """文件选择事件 - 显示预览"""
        _ = event  # 使用event参数避免警告
        selection = self.file_listbox.curselection()
        if not selection:
            return

        idx = selection[0]
        if idx >= len(self.image_files):
            return

        file_path = self.image_files[idx]
        try:
            img = Image.open(file_path)
            # 缩放到预览区域
            img.thumbnail((200, 200), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self.preview_label.config(image=photo, text='')
            self.preview_image = photo  # 保持引用
        except Exception as e:
            self.preview_label.config(image='', text=f"预览失败\n{e}")
            self.preview_image = None

    def _log(self, msg):
        """输出日志"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        full_msg = f"[{timestamp}] {msg}\n"
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, full_msg)
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def _set_progress(self, value, text=None):
        """设置进度"""
        self.progress["value"] = value
        if text:
            self.var_progress_text.set(text)

    # ── 启动 / 停止 ──────────────────────────────────────
    def _start(self):
        """开始处理"""
        if not self.image_files:
            messagebox.showwarning("提示", "请先添加图片文件")
            return

        # 验证参数
        try:
            width = self.var_width.get()
            height = self.var_height.get()
            if width <= 0 or height <= 0:
                messagebox.showerror("错误", "宽度和高度必须大于0")
                return
        except:
            messagebox.showerror("错误", "请输入有效的宽度和高度")
            return

        # 创建输出目录
        output_dir = Path(self.var_output_dir.get())
        output_dir.mkdir(parents=True, exist_ok=True)

        self.running = True
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self._set_progress(0, "处理中...")

        self.thread = threading.Thread(target=self._run_processing, daemon=True)
        self.thread.start()

    def _stop(self):
        """停止处理"""
        self.running = False
        self._log("用户请求停止，将在当前图片完成后终止...")

    def _run_processing(self):
        """后台处理线程"""
        try:
            self._do_processing()
        except Exception as e:
            self._log(f"处理异常: {e}")
        finally:
            self.root.after(0, self._on_finished)

    def _do_processing(self):
        """执行图片处理"""
        width = self.var_width.get()
        height = self.var_height.get()
        mode = ImageResizer.RESIZE_MODES[self.var_mode.get()]
        quality = self.var_quality.get()
        output_dir = Path(self.var_output_dir.get())
        keep_structure = self.var_keep_structure.get()
        overwrite = self.var_overwrite.get()

        # 获取输出格式
        format_str = self.var_format.get()
        output_format = None if format_str == "保持原格式" else format_str

        # 创建处理器
        resizer = ImageResizer(log_callback=self._log)

        total = len(self.image_files)
        success_count = 0
        skip_count = 0
        fail_count = 0

        self._log(f"开始处理 {total} 张图片")
        self._log(f"目标尺寸: {width}×{height}")
        self._log(f"缩放模式: {self.var_mode.get()}")
        self._log(f"输出格式: {format_str}")
        self._log(f"图片质量: {quality}")
        self._log("-" * 60)

        for idx, input_path in enumerate(self.image_files, 1):
            if not self.running:
                self._log("处理已停止")
                break

            # 更新进度
            progress = int((idx - 1) / total * 100)
            self.root.after(0, self._set_progress, progress, f"处理中 ({idx}/{total})")

            input_file = Path(input_path)
            self._log(f"[{idx}/{total}] {input_file.name}")

            # 确定输出路径
            if keep_structure:
                # 保持文件夹结构（如果是从文件夹添加的）
                try:
                    # 尝试获取相对路径
                    rel_path = input_file.relative_to(input_file.parent.parent)
                    output_path = output_dir / rel_path.parent / input_file.name
                except:
                    output_path = output_dir / input_file.name
            else:
                output_path = output_dir / input_file.name

            # 如果指定了输出格式，修改扩展名
            if output_format:
                ext_map = {"JPEG": ".jpg", "PNG": ".png", "WEBP": ".webp"}
                output_path = output_path.with_suffix(ext_map.get(output_format, output_path.suffix))

            # 检查文件是否已存在
            if output_path.exists() and not overwrite:
                self._log(f"  ⊙ 文件已存在，跳过")
                skip_count += 1
                continue

            # 创建输出目录
            output_path.parent.mkdir(parents=True, exist_ok=True)

            # 处理图片
            try:
                success = resizer.resize_image(
                    input_path=str(input_file),
                    output_path=str(output_path),
                    width=width,
                    height=height,
                    mode=mode,
                    quality=quality,
                    format=output_format
                )

                if success:
                    # 获取文件大小
                    original_size = input_file.stat().st_size / 1024  # KB
                    new_size = output_path.stat().st_size / 1024  # KB
                    self._log(f"  ✓ 完成 ({original_size:.1f}KB → {new_size:.1f}KB)")
                    success_count += 1
                else:
                    fail_count += 1

            except Exception as e:
                self._log(f"  ✗ 失败: {e}")
                fail_count += 1

        # 完成统计
        self._log("-" * 60)
        self._log(f"处理完成！")
        self._log(f"  成功: {success_count} 张")
        if skip_count > 0:
            self._log(f"  跳过: {skip_count} 张")
        if fail_count > 0:
            self._log(f"  失败: {fail_count} 张")
        self._log(f"输出目录: {output_dir}")

    def _on_finished(self):
        """处理完成"""
        self.running = False
        self.btn_start.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self._set_progress(100, "完成")


# ── 入口 ──────────────────────────────────────────────────
def main():
    root = tk.Tk()
    ImageResizerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
