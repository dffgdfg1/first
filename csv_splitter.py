import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import pandas as pd
import os
from pathlib import Path
import threading
from datetime import datetime

class CSVSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV 文件拆分工具")
        self.root.geometry("700x650")
        
        self.EXCEL_MAX_ROWS = 1048575
        self.selected_file = None
        self.is_processing = False
        
        self.setup_ui()
    
    def setup_ui(self):
        title_label = tk.Label(
            self.root, 
            text="CSV 文件拆分工具", 
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=15)
        
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=10, padx=20, fill=tk.X)
        
        self.file_label = tk.Label(
            file_frame, 
            text="未选择文件", 
            bg="lightgray", 
            width=50, 
            anchor="w",
            padx=10
        )
        self.file_label.pack(side=tk.LEFT, padx=(0, 10))
        
        select_btn = tk.Button(
            file_frame, 
            text="选择 CSV", 
            command=self.select_file,
            padx=10
        )
        select_btn.pack(side=tk.LEFT)
        
        rows_frame = tk.Frame(self.root)
        rows_frame.pack(pady=10, padx=20, fill=tk.X)
        
        tk.Label(rows_frame, text="每个文件最大行数:").pack(side=tk.LEFT)
        self.rows_entry = tk.Entry(rows_frame, width=15)
        self.rows_entry.insert(0, str(self.EXCEL_MAX_ROWS))
        self.rows_entry.pack(side=tk.LEFT, padx=10)
        tk.Label(rows_frame, text="(不含表头)").pack(side=tk.LEFT)
        
        encoding_frame = tk.Frame(self.root)
        encoding_frame.pack(pady=10, padx=20, fill=tk.X)
        
        tk.Label(encoding_frame, text="文件编码:").pack(side=tk.LEFT)
        self.encoding_var = tk.StringVar(value="utf-8")
        
        encodings = ["utf-8", "gbk", "gb2312", "utf-8-sig", "latin1"]
        encoding_menu = ttk.Combobox(
            encoding_frame, 
            textvariable=self.encoding_var,
            values=encodings,
            width=12,
            state="readonly"
        )
        encoding_menu.pack(side=tk.LEFT, padx=10)
        
        clean_frame = tk.Frame(self.root)
        clean_frame.pack(pady=10, padx=20, fill=tk.X)
        
        self.auto_clean_var = tk.BooleanVar(value=True)
        clean_check = tk.Checkbutton(
            clean_frame,
            text="智能去除无关信息（自动检测表头位置）",
            variable=self.auto_clean_var
        )
        clean_check.pack(side=tk.LEFT)
        
        self.progress = ttk.Progressbar(
            self.root, 
            mode='indeterminate', 
            length=500
        )
        self.progress.pack(pady=15)
        
        self.status_label = tk.Label(
            self.root, 
            text="请选择要拆分的 CSV 文件", 
            fg="gray",
            wraplength=550
        )
        self.status_label.pack(pady=5)
        
        # 日志区域
        log_frame = tk.LabelFrame(self.root, text="处理日志", padx=10, pady=10)
        log_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=12,
            width=70,
            font=("Consolas", 9),
            bg="#f5f5f5",
            fg="#333333"
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        self.process_btn = tk.Button(
            self.root, 
            text="开始拆分", 
            command=self.start_processing,
            state=tk.DISABLED,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 12, "bold"),
            width=20,
            height=2
        )
        self.process_btn.pack(pady=15)
    
    def log(self, message, level="INFO"):
        """添加日志信息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # 根据级别设置颜色
        color_map = {
            "INFO": "#0066cc",
            "SUCCESS": "#00aa00",
            "WARNING": "#ff8800",
            "ERROR": "#cc0000"
        }
        
        log_entry = f"[{timestamp}] {message}\n"
        
        self.log_text.insert(tk.END, log_entry)
        
        # 为最后一行设置颜色
        last_line_start = self.log_text.index("end-2c linestart")
        last_line_end = self.log_text.index("end-1c")
        
        tag_name = f"{level}_{timestamp}"
        self.log_text.tag_add(tag_name, last_line_start, last_line_end)
        self.log_text.tag_config(tag_name, foreground=color_map.get(level, "#333333"))
        
        # 自动滚动到底部
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def select_file(self):
        filename = filedialog.askopenfilename(
            title="选择 CSV 文件",
            filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")]
        )
        
        if filename:
            self.selected_file = filename
            display_name = os.path.basename(filename)
            if len(display_name) > 40:
                display_name = display_name[:37] + "..."
            self.file_label.config(text=display_name)
            self.process_btn.config(state=tk.NORMAL)
            
            file_size = os.path.getsize(filename) / (1024 * 1024)
            self.update_status(f"文件已选择 ({file_size:.1f} MB)，点击开始拆分", "blue")
            self.root.after(0, self.log, f"已选择文件: {os.path.basename(filename)} ({file_size:.1f} MB)")
    
    def update_status(self, message, color="blue"):
        self.status_label.config(text=message, fg=color)
    
    def find_header_row(self, filepath, encoding):
        """智能检测表头所在行"""
        try:
            with open(filepath, 'r', encoding=encoding, errors='ignore') as f:
                lines = []
                for i, line in enumerate(f):
                    if i >= 50:
                        break
                    lines.append(line.strip())
                
                for i, line in enumerate(lines):
                    if not line:
                        continue
                    
                    comma_count = line.count(',')
                    
                    if comma_count >= 2:
                        fields = line.split(',')
                        has_text = any(any(c.isalpha() for c in field) for field in fields)
                        
                        if has_text:
                            return i
                
                return 0
                
        except Exception as e:
            self.root.after(0, self.log, f"检测表头时出现警告: {str(e)}", "WARNING")
            return 0
    
    def start_processing(self):
        if not self.selected_file or self.is_processing:
            return
        
        try:
            max_rows = int(self.rows_entry.get())
            if max_rows <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("错误", "请输入有效的行数")
            return
        
        # 清空日志
        self.log_text.delete(1.0, tk.END)
        
        self.is_processing = True
        self.process_btn.config(state=tk.DISABLED, text="处理中...")
        self.progress.start()
        
        self.root.after(0, self.log, "=" * 60)
        self.root.after(0, self.log, "开始处理 CSV 文件")
        self.root.after(0, self.log, "=" * 60)
        
        thread = threading.Thread(
            target=self.process_file, 
            args=(max_rows,),
            daemon=True
        )
        thread.start()
    
    def process_file(self, max_rows_per_file):
        try:
            input_file = Path(self.selected_file)
            output_dir = input_file.parent / f"{input_file.stem}_拆分"
            output_dir.mkdir(exist_ok=True)
            
            encoding = self.encoding_var.get()
            
            self.root.after(0, self.log, f"输出目录: {output_dir}")
            self.root.after(0, self.log, f"文件编码: {encoding}")
            self.root.after(0, self.log, f"每个文件最大行数: {max_rows_per_file:,}")
            
            # 智能检测表头
            header_row = 0
            if self.auto_clean_var.get():
                self.root.after(0, self.update_status, "正在检测表头位置...", "blue")
                self.root.after(0, self.log, "正在智能检测表头位置...")
                header_row = self.find_header_row(self.selected_file, encoding)
                if header_row > 0:
                    self.root.after(0, self.log, f"检测到表头在第 {header_row + 1} 行，将跳过前 {header_row} 行", "WARNING")
                else:
                    self.root.after(0, self.log, "表头在第 1 行，无需跳过")
            
            self.root.after(0, self.update_status, "正在读取 CSV 文件...", "blue")
            self.root.after(0, self.log, "开始读取 CSV 文件...")
            
            file_count = 0
            total_rows_processed = 0
            buffer = []
            
            chunk_size = 50000
            
            # 读取 CSV
            try:
                if header_row > 0:
                    reader = pd.read_csv(
                        self.selected_file,
                        encoding=encoding,
                        chunksize=chunk_size,
                        skiprows=list(range(header_row)),
                        on_bad_lines='skip',
                        low_memory=False
                    )
                else:
                    reader = pd.read_csv(
                        self.selected_file,
                        encoding=encoding,
                        chunksize=chunk_size,
                        on_bad_lines='skip',
                        low_memory=False
                    )
            except:
                self.root.after(0, self.log, "C 引擎读取失败，切换到 Python 引擎", "WARNING")
                if header_row > 0:
                    reader = pd.read_csv(
                        self.selected_file,
                        encoding=encoding,
                        chunksize=chunk_size,
                        skiprows=list(range(header_row)),
                        on_bad_lines='skip',
                        engine='python'
                    )
                else:
                    reader = pd.read_csv(
                        self.selected_file,
                        encoding=encoding,
                        chunksize=chunk_size,
                        on_bad_lines='skip',
                        engine='python'
                    )
            
            chunk_count = 0
            # 持续读取所有数据块
            for chunk in reader:
                chunk_count += 1
                
                # 清理数据
                original_len = len(chunk)
                chunk = chunk.dropna(how='all')
                cleaned_len = len(chunk)
                
                if original_len != cleaned_len:
                    self.root.after(0, self.log, f"数据块 {chunk_count}: 清理了 {original_len - cleaned_len} 个空行")
                
                if len(chunk) == 0:
                    continue
                
                buffer.append(chunk)
                total_rows_processed += len(chunk)
                
                self.root.after(0, self.log, f"读取数据块 {chunk_count}: {len(chunk):,} 行 (累计: {total_rows_processed:,} 行)")
                
                buffer_size = sum(len(df) for df in buffer)
                
                # 写入文件
                while buffer_size >= max_rows_per_file:
                    file_count += 1
                    
                    combined_df = pd.concat(buffer, ignore_index=True)
                    
                    to_save = combined_df.iloc[:max_rows_per_file]
                    remaining = combined_df.iloc[max_rows_per_file:]
                    
                    output_file = output_dir / f"{input_file.stem}_part{file_count}.xlsx"
                    
                    self.root.after(0, self.log, f"正在生成文件 {file_count}: {len(to_save):,} 行...")
                    to_save.to_excel(output_file, index=False, engine='openpyxl')
                    self.root.after(0, self.log, f"✓ 文件 {file_count} 已保存: {output_file.name}", "SUCCESS")
                    
                    if len(remaining) > 0:
                        buffer = [remaining]
                        buffer_size = len(remaining)
                    else:
                        buffer = []
                        buffer_size = 0
                    
                    self.root.after(
                        0,
                        self.update_status,
                        f"已生成 {file_count} 个文件，已读取 {total_rows_processed:,} 行",
                        "blue"
                    )
            
            # 保存最后剩余的数据
            if buffer:
                file_count += 1
                combined_df = pd.concat(buffer, ignore_index=True)
                output_file = output_dir / f"{input_file.stem}_part{file_count}.xlsx"
                
                self.root.after(0, self.log, f"正在生成最后一个文件 {file_count}: {len(combined_df):,} 行...")
                combined_df.to_excel(output_file, index=False, engine='openpyxl')
                self.root.after(0, self.log, f"✓ 文件 {file_count} 已保存: {output_file.name}", "SUCCESS")
            
            self.root.after(0, self.progress.stop)
            self.root.after(0, self.log, "=" * 60)
            self.root.after(0, self.log, f"处理完成！", "SUCCESS")
            self.root.after(0, self.log, f"总行数: {total_rows_processed:,}", "SUCCESS")
            self.root.after(0, self.log, f"生成文件数: {file_count}", "SUCCESS")
            self.root.after(0, self.log, f"保存位置: {output_dir}", "SUCCESS")
            self.root.after(0, self.log, "=" * 60)
            
            summary = f"拆分完成！\n\n"
            if header_row > 0:
                summary += f"已跳过前 {header_row} 行无关信息\n"
            summary += f"总行数: {total_rows_processed:,}\n"
            summary += f"生成文件数: {file_count}\n"
            summary += f"每个文件: 最多 {max_rows_per_file:,} 行\n\n"
            summary += f"保存位置:\n{output_dir}"
            
            self.root.after(
                0,
                lambda: messagebox.showinfo("完成", summary)
            )
            self.root.after(
                0,
                self.update_status,
                f"✓ 完成！{file_count} 个文件，共 {total_rows_processed:,} 行数据",
                "green"
            )
            
        except Exception as e:
            self.root.after(0, self.progress.stop)
            error_msg = str(e)
            
            self.root.after(0, self.log, "=" * 60, "ERROR")
            self.root.after(0, self.log, f"处理失败: {error_msg}", "ERROR")
            self.root.after(0, self.log, "=" * 60, "ERROR")
            
            if "codec" in error_msg.lower() or "decode" in error_msg.lower():
                error_msg += "\n\n建议：尝试更换编码格式（gbk 或 gb2312）"
                self.root.after(0, self.log, "建议：尝试更换编码格式", "WARNING")
            
            self.root.after(
                0,
                lambda: messagebox.showerror("错误", f"处理失败:\n{error_msg}")
            )
            self.root.after(
                0,
                self.update_status,
                f"✗ 处理失败",
                "red"
            )
        
        finally:
            self.is_processing = False
            self.root.after(0, self.process_btn.config, {"state": tk.NORMAL, "text": "开始拆分"})

if __name__ == "__main__":
    root = tk.Tk()
    app = CSVSplitterApp(root)
    root.mainloop()
