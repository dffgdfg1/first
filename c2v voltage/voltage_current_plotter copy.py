import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import os
import threading
from datetime import datetime

# 设置中文字体支持
plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False


class DataMergerApp:
    """数据合成与可视化应用"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("数据曲线合成工具")
        self.root.geometry("1200x800")
        
        self.data_files = []
        self.merged_data = None
        self.sampling_rate = 100
        
        self.setup_ui()
    
    def setup_ui(self):
        """设置用户界面"""
        # 顶部控制面板
        control_frame = tk.Frame(self.root, padx=10, pady=10)
        control_frame.pack(side=tk.TOP, fill=tk.X)
        
        # 第一行按钮
        button_frame1 = tk.Frame(control_frame)
        button_frame1.pack(fill=tk.X, pady=5)
        
        tk.Button(button_frame1, text="选择文件", command=self.select_files, 
                 bg="#4CAF50", fg="white", padx=20, pady=5).pack(side=tk.LEFT, padx=5)
        
        tk.Button(button_frame1, text="合成曲线", command=self.start_merge,
                 bg="#2196F3", fg="white", padx=20, pady=5).pack(side=tk.LEFT, padx=5)
        
        tk.Button(button_frame1, text="导出数据", command=self.export_data,
                 bg="#FF9800", fg="white", padx=20, pady=5).pack(side=tk.LEFT, padx=5)
        
        tk.Button(button_frame1, text="清除", command=self.clear_all,
                 bg="#f44336", fg="white", padx=20, pady=5).pack(side=tk.LEFT, padx=5)
        
        # 采样设置
        sampling_frame = tk.Frame(control_frame)
        sampling_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(sampling_frame, text="采样率（每N个点取1个）:").pack(side=tk.LEFT, padx=5)
        
        self.sampling_var = tk.StringVar(value="100")
        sampling_entry = tk.Entry(sampling_frame, textvariable=self.sampling_var, width=10)
        sampling_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Label(sampling_frame, text="(1=不采样，100=每100个点取1个)").pack(side=tk.LEFT, padx=5)
        
        # 进度条
        self.progress_frame = tk.Frame(control_frame)
        self.progress_frame.pack(fill=tk.X, pady=5)
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate')
        self.progress_bar.pack(fill=tk.X, padx=5)
        
        self.progress_label = tk.Label(self.progress_frame, text="")
        self.progress_label.pack(pady=2)
        
        # 文件列表框架
        list_frame = tk.Frame(self.root, padx=10)
        list_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=False)
        
        tk.Label(list_frame, text="已选择的文件:", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        
        # 创建带滚动条的列表框
        listbox_frame = tk.Frame(list_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        scrollbar = tk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(listbox_frame, height=5, yscrollcommand=scrollbar.set)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # 图表框架
        self.figure_frame = tk.Frame(self.root)
        self.figure_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 状态栏
        self.status_label = tk.Label(self.root, text="就绪", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)
    
    def select_files(self):
        """选择数据文件"""
        files = filedialog.askopenfilenames(
            title="选择数据文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("CSV文件", "*.csv"),
                ("文本文件", "*.txt"),
                ("所有文件", "*.*")
            ]
        )
        
        if files:
            self.data_files.extend(files)
            self.update_file_list()
            self.status_label.config(text=f"已选择 {len(self.data_files)} 个文件")
    
    def update_file_list(self):
        """更新文件列表显示"""
        self.file_listbox.delete(0, tk.END)
        for file in self.data_files:
            self.file_listbox.insert(tk.END, os.path.basename(file))
    
    def parse_time_with_microseconds(self, time_str):
        """解析包含微秒的时间字符串: 01/04/1970 00:37:48.437212648"""
        if pd.isna(time_str):
            return None
            
        try:
            if isinstance(time_str, datetime):
                return time_str
            
            time_str = str(time_str).strip()
            
            # 分离日期时间和微秒部分
            if '.' in time_str:
                datetime_part, microseconds_part = time_str.rsplit('.', 1)
                dt = datetime.strptime(datetime_part, "%m/%d/%Y %H:%M:%S")
                # 只取前6位微秒（Python datetime只支持6位）
                microseconds = int(microseconds_part[:6].ljust(6, '0'))
                return dt.replace(microsecond=microseconds)
            else:
                return datetime.strptime(time_str, "%m/%d/%Y %H:%M:%S")
        except Exception:
            try:
                return pd.to_datetime(time_str)
            except Exception:
                return None
    
    def read_excel_file_chunked(self, filepath, chunk_size=100000):
        """分块读取Excel文件，提取第1列（电流）和第14列（时间）"""
        try:
            all_chunks = []
            
            # Excel不支持chunksize，需要一次性读取后分块处理
            df = pd.read_excel(filepath, header=None)
            
            # 分块处理
            for i in range(0, len(df), chunk_size):
                chunk = df.iloc[i:i+chunk_size]
                chunk_data = self.process_chunk_columns(chunk)
                if not chunk_data.empty:
                    all_chunks.append(chunk_data)
            
            if all_chunks:
                return pd.concat(all_chunks, ignore_index=True)
            return pd.DataFrame()
        
        except Exception as e:
            self.update_status(f"读取Excel文件失败: {str(e)}")
            return None
    
    def read_text_file_chunked(self, filepath, chunk_size=100000):
        """分块读取文本文件，提取第1列（电流）和第14列（时间）"""
        try:
            all_chunks = []
            
            for chunk in pd.read_csv(filepath, header=None, chunksize=chunk_size, 
                                    on_bad_lines='skip', encoding='utf-8'):
                chunk_data = self.process_chunk_columns(chunk)
                if not chunk_data.empty:
                    all_chunks.append(chunk_data)
            
            if all_chunks:
                return pd.concat(all_chunks, ignore_index=True)
            return pd.DataFrame()
        
        except Exception as e:
            self.update_status(f"读取文本文件失败: {str(e)}")
            return None
    
    def process_chunk_columns(self, chunk):
        """处理数据块，提取第1列和第14列"""
        data = []
        
        # 检查是否有足够的列
        if chunk.shape[1] < 14:
            self.update_status(f"警告：数据列数不足14列，实际只有{chunk.shape[1]}列")
            return pd.DataFrame()
        
        for _, row in chunk.iterrows():
            try:
                # 第1列是电流（索引0）
                current = float(row[0])
                
                # 第14列是时间（索引13）
                timestamp = self.parse_time_with_microseconds(row[13])
                
                if timestamp:
                    data.append({'timestamp': timestamp, 'current': current})
            except (ValueError, TypeError, IndexError):
                continue
        
        return pd.DataFrame(data)
    
    def downsample_data(self, df, sampling_rate):
        """降采样数据"""
        if sampling_rate <= 1:
            return df
        
        # 按时间排序
        df = df.sort_values('timestamp').reset_index(drop=True)
        
        # 每N个点取1个
        sampled_df = df.iloc[::sampling_rate].copy()
        
        return sampled_df
    
    def read_data_file(self, filepath, file_index, total_files):
        """读取数据文件（自动识别格式）"""
        ext = os.path.splitext(filepath)[1].lower()
        filename = os.path.basename(filepath)
        
        self.update_progress(file_index, total_files, f"正在读取: {filename}")
        
        if ext in ['.xlsx', '.xls']:
            df = self.read_excel_file_chunked(filepath)
        elif ext in ['.csv', '.txt']:
            df = self.read_text_file_chunked(filepath)
        else:
            df = self.read_text_file_chunked(filepath)
        
        return df
    
    def start_merge(self):
        """启动合成任务（在后台线程中）"""
        if not self.data_files:
            messagebox.showwarning("警告", "请先选择数据文件")
            return
        
        try:
            self.sampling_rate = int(self.sampling_var.get())
            if self.sampling_rate < 1:
                self.sampling_rate = 1
        except ValueError:
            messagebox.showerror("错误", "采样率必须是正整数")
            return
        
        # 在后台线程中执行
        thread = threading.Thread(target=self.merge_curves, daemon=True)
        thread.start()
    
    def merge_curves(self):
        """合成多个文件的曲线"""
        self.update_status("正在合成曲线...")
        
        all_data = []
        total_files = len(self.data_files)
        
        for i, filepath in enumerate(self.data_files):
            df = self.read_data_file(filepath, i + 1, total_files)
            if df is not None and not df.empty:
                all_data.append(df)
        
        if not all_data:
            self.update_status("合成失败：没有成功读取任何数据")
            self.root.after(0, lambda: messagebox.showerror("错误", "没有成功读取任何数据"))
            return
        
        self.update_progress(0, 100, "正在合并数据...")
        
        # 合并所有数据
        merged = pd.concat(all_data, ignore_index=True)
        
        self.update_progress(30, 100, "正在排序...")
        
        # 按时间排序
        merged = merged.sort_values('timestamp').reset_index(drop=True)
        
        original_count = len(merged)
        
        self.update_progress(60, 100, "正在降采样...")
        
        # 降采样
        self.merged_data = self.downsample_data(merged, self.sampling_rate)
        
        sampled_count = len(self.merged_data)
        
        self.update_progress(90, 100, "正在绘制曲线...")
        
        # 绘制曲线
        self.root.after(0, self.plot_curve)
        
        self.update_progress(100, 100, "完成")
        
        status_msg = f"合成完成：原始 {original_count:,} 个点，采样后 {sampled_count:,} 个点"
        self.update_status(status_msg)
    
    def plot_curve(self):
        """绘制时间-电流曲线"""
        for widget in self.figure_frame.winfo_children():
            widget.destroy()
        
        fig, ax = plt.subplots(figsize=(10, 5))
        
        ax.plot(self.merged_data['timestamp'], self.merged_data['current'], 
               linewidth=1, color='#2196F3', label='电流', alpha=0.8)
        
        ax.set_xlabel('时间', fontsize=12)
        ax.set_ylabel('电流 (A)', fontsize=12)
        ax.set_title('时间-电流曲线', fontsize=14, fontweight='bold')
        ax.grid(True, alpha=0.3)
        ax.legend()
        
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, master=self.figure_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def export_data(self):
        """导出合成后的数据"""
        if self.merged_data is None or self.merged_data.empty:
            messagebox.showwarning("警告", "没有可导出的数据，请先合成曲线")
            return
        
        filepath = filedialog.asksaveasfilename(
            title="保存数据",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel文件", "*.xlsx"),
                ("CSV文件", "*.csv"),
                ("文本文件", "*.txt")
            ]
        )
        
        if filepath:
            try:
                export_data = self.merged_data.copy()
                # 保留完整的时间戳格式（包含微秒）
                export_data['timestamp'] = export_data['timestamp'].dt.strftime('%m/%d/%Y %H:%M:%S.%f')
                
                ext = os.path.splitext(filepath)[1].lower()
                if ext == '.xlsx':
                    export_data.to_excel(filepath, index=False, columns=['current', 'timestamp'])
                else:
                    export_data.to_csv(filepath, index=False, columns=['current', 'timestamp'])
                
                messagebox.showinfo("成功", f"数据已导出到:\n{filepath}")
                self.update_status("数据导出成功")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败:\n{str(e)}")
    
    def clear_all(self):
        """清除所有数据"""
        self.data_files = []
        self.merged_data = None
        self.file_listbox.delete(0, tk.END)
        
        for widget in self.figure_frame.winfo_children():
            widget.destroy()
        
        self.progress_bar['value'] = 0
        self.progress_label.config(text="")
        self.update_status("已清除所有数据")
    
    def update_progress(self, current, total, message=""):
        """更新进度条"""
        def update():
            progress = (current / total) * 100 if total > 0 else 0
            self.progress_bar['value'] = progress
            self.progress_label.config(text=message)
        
        self.root.after(0, update)
    
    def update_status(self, message):
        """更新状态栏"""
        def update():
            self.status_label.config(text=message)
        
        self.root.after(0, update)


def main():
    """主函数"""
    root = tk.Tk()
    app = DataMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
