import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib
from scipy import signal
import warnings
warnings.filterwarnings('ignore')

matplotlib.use('TkAgg')

class VoltageCurrentAnalyzerPro:
    def __init__(self, root):
        self.root = root
        self.root.title("通用电压电流对比分析器 Pro")
        self.root.geometry("1400x800")
        
        # 设置样式
        self.setup_styles()
        
        # 数据存储
        self.voltage_data = None
        self.current_data = None
        self.time_aligned = False
        
        # 创建菜单栏
        self.create_menu()
        
        # 创建主界面布局
        self.create_main_layout()
        
    def setup_styles(self):
        """设置界面样式"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置标签样式
        style.configure('Title.TLabel', font=('微软雅黑', 12, 'bold'))
        style.configure('Subtitle.TLabel', font=('微软雅黑', 10))
        
        # 配置按钮样式
        style.configure('Primary.TButton', font=('微软雅黑', 10))
        style.configure('Success.TButton', font=('微软雅黑', 10))
        
    def create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="新建", command=self.new_project)
        file_menu.add_command(label="打开配置", command=self.open_config)
        file_menu.add_command(label="保存配置", command=self.save_config)
        file_menu.add_separator()
        file_menu.add_command(label="导出PNG", command=lambda: self.export_chart('png'))
        file_menu.add_command(label="导出PDF", command=lambda: self.export_chart('pdf'))
        file_menu.add_command(label="导出CSV", command=self.export_csv)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        
        # 工具菜单
        tool_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="工具", menu=tool_menu)
        tool_menu.add_command(label="滤波设置", command=self.filter_settings)
        tool_menu.add_command(label="标定参数", command=self.calibration_settings)
        tool_menu.add_command(label="计算器", command=self.open_calculator)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="用户手册", command=self.show_user_manual)
        help_menu.add_command(label="关于", command=self.show_about)
        
    def create_main_layout(self):
        """创建主界面布局"""
        # 左侧面板 - 数据源配置
        left_frame = ttk.LabelFrame(self.root, text="数据源配置", padding=10)
        left_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        
        # 电压数据文件
        ttk.Label(left_frame, text="电压数据文件:", style='Subtitle.TLabel').grid(row=0, column=0, sticky="w")
        ttk.Button(left_frame, text="选择文件...", 
                  command=self.load_voltage_file).grid(row=1, column=0, pady=(0,5))
        self.voltage_file_label = ttk.Label(left_frame, text="未选择", foreground="gray")
        self.voltage_file_label.grid(row=2, column=0, sticky="w", pady=(0,20))
        
        # 电流数据文件
        ttk.Label(left_frame, text="电流数据文件:", style='Subtitle.TLabel').grid(row=3, column=0, sticky="w")
        ttk.Button(left_frame, text="选择文件...", 
                  command=self.load_current_file).grid(row=4, column=0, pady=(0,5))
        self.current_file_label = ttk.Label(left_frame, text="未选择", foreground="gray")
        self.current_file_label.grid(row=5, column=0, sticky="w")
        
        # 中央控制区 - 对比参数设置
        center_frame = ttk.LabelFrame(self.root, text="对比参数设置", padding=10)
        center_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        
        # 对比标题
        ttk.Label(center_frame, text="对比标题:", style='Subtitle.TLabel').grid(row=0, column=0, sticky="w")
        self.title_entry = ttk.Entry(center_frame, width=25)
        self.title_entry.insert(0, "电压-电流时序对比图")
        self.title_entry.grid(row=1, column=0, pady=(0,10))
        
        # 时间范围选项
        ttk.Label(center_frame, text="时间范围:", style='Subtitle.TLabel').grid(row=2, column=0, sticky="w", pady=(5,0))
        self.auto_align_var = tk.BooleanVar(value=True)
        self.auto_align_check = ttk.Checkbutton(center_frame, text="自动对齐时间轴", 
                                               variable=self.auto_align_var)
        self.auto_align_check.grid(row=3, column=0, sticky="w")
        
        self.sync_zoom_var = tk.BooleanVar(value=True)
        self.sync_zoom_check = ttk.Checkbutton(center_frame, text="同步缩放", 
                                              variable=self.sync_zoom_var)
        self.sync_zoom_check.grid(row=4, column=0, sticky="w")
        
        # 显示选项
        ttk.Label(center_frame, text="显示选项:", style='Subtitle.TLabel').grid(row=5, column=0, sticky="w", pady=(10,0))
        self.show_trend_var = tk.BooleanVar(value=False)
        self.show_trend_check = ttk.Checkbutton(center_frame, text="显示趋势线", 
                                               variable=self.show_trend_var)
        self.show_trend_check.grid(row=6, column=0, sticky="w")
        
        self.show_stats_var = tk.BooleanVar(value=False)
        self.show_stats_check = ttk.Checkbutton(center_frame, text="显示统计指标", 
                                               variable=self.show_stats_var)
        self.show_stats_check.grid(row=7, column=0, sticky="w")
        
        self.highlight_anomaly_var = tk.BooleanVar(value=True)
        self.highlight_anomaly_check = ttk.Checkbutton(center_frame, text="异常点高亮", 
                                                      variable=self.highlight_anomaly_var)
        self.highlight_anomaly_check.grid(row=8, column=0, sticky="w")
        
        # 控制按钮
        ttk.Button(center_frame, text="生成对比图表", 
                  command=self.generate_chart, style='Primary.TButton').grid(row=9, column=0, pady=(20,5))
        ttk.Button(center_frame, text="清空选择", 
                  command=self.clear_selection, style='Success.TButton').grid(row=10, column=0, pady=5)
        
        # 右侧面板 - 分析视图
        right_frame = ttk.LabelFrame(self.root, text="图表显示区域", padding=10)
        right_frame.grid(row=0, column=2, padx=5, pady=5, sticky="nsew")
        
        # 视图选项
        ttk.Label(right_frame, text="双Y轴对比图:", style='Subtitle.TLabel').grid(row=0, column=0, sticky="w")
        ttk.Label(right_frame, text="左轴: 电压(V)").grid(row=1, column=0, sticky="w")
        ttk.Label(right_frame, text="右轴: 电流(A)").grid(row=2, column=0, sticky="w", pady=(0,10))
        
        ttk.Label(right_frame, text="时间轴: 时间戳").grid(row=3, column=0, sticky="w", pady=(0,10))
        
        ttk.Label(right_frame, text="可选视图:", style='Subtitle.TLabel').grid(row=4, column=0, sticky="w")
        self.view_mode = tk.StringVar(value="overlay")
        ttk.Radiobutton(right_frame, text="叠加对比", variable=self.view_mode, 
                       value="overlay").grid(row=5, column=0, sticky="w")
        ttk.Radiobutton(right_frame, text="分屏显示", variable=self.view_mode, 
                       value="split").grid(row=6, column=0, sticky="w")
        
        # 图表显示区域
        self.fig, (self.ax1, self.ax2) = plt.subplots(2, 1, figsize=(8, 6))
        self.canvas = FigureCanvasTkAgg(self.fig, master=right_frame)
        self.canvas.get_tk_widget().grid(row=7, column=0, pady=10)
        
        # 底部面板 - 分析结果
        bottom_frame = ttk.LabelFrame(self.root, text="数据分析结果", padding=10)
        bottom_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        
        # 相关性分析
        correlation_frame = ttk.Frame(bottom_frame)
        correlation_frame.grid(row=0, column=0, padx=20, sticky="w")
        ttk.Label(correlation_frame, text="相关性分析:", style='Subtitle.TLabel').grid(row=0, column=0, sticky="w")
        self.correlation_label = ttk.Label(correlation_frame, text="相关系数: --")
        self.correlation_label.grid(row=1, column=0, sticky="w")
        self.phase_label = ttk.Label(correlation_frame, text="相位差: -- ms")
        self.phase_label.grid(row=2, column=0, sticky="w")
        
        # 统计信息
        stats_frame = ttk.Frame(bottom_frame)
        stats_frame.grid(row=0, column=1, padx=20, sticky="w")
        ttk.Label(stats_frame, text="统计信息:", style='Subtitle.TLabel').grid(row=0, column=0, sticky="w")
        self.voltage_range_label = ttk.Label(stats_frame, text="电压范围: -- V")
        self.voltage_range_label.grid(row=1, column=0, sticky="w")
        self.current_range_label = ttk.Label(stats_frame, text="电流范围: -- A")
        self.current_range_label.grid(row=2, column=0, sticky="w")
        self.power_label = ttk.Label(stats_frame, text="平均功率: -- W")
        self.power_label.grid(row=3, column=0, sticky="w")
        
        # 异常检测
        anomaly_frame = ttk.Frame(bottom_frame)
        anomaly_frame.grid(row=0, column=2, padx=20, sticky="w")
        ttk.Label(anomaly_frame, text="异常检测:", style='Subtitle.TLabel').grid(row=0, column=0, sticky="w")
        self.overvoltage_label = ttk.Label(anomaly_frame, text="过压点: 0")
        self.overvoltage_label.grid(row=1, column=0, sticky="w")
        self.overcurrent_label = ttk.Label(anomaly_frame, text="过流点: 0")
        self.overcurrent_label.grid(row=2, column=0, sticky="w")
        self.anomaly_power_label = ttk.Label(anomaly_frame, text="异常功耗点: 0")
        self.anomaly_power_label.grid(row=3, column=0, sticky="w")
        
        # 配置网格权重
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(2, weight=2)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=0)
        
    def load_voltage_file(self):
        """加载电压数据文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV文件", "*.csv"), ("Excel文件", "*.xlsx;*.xls"), ("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if file_path:
            self.voltage_data = self.load_data_file(file_path)
            if self.voltage_data is not None:
                self.voltage_file_label.config(text=file_path.split('/')[-1])
                
    def load_current_file(self):
        """加载电流数据文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV文件", "*.csv"), ("Excel文件", "*.xlsx;*.xls"), ("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if file_path:
                self.current_data = self.load_data_file(file_path)
                if self.current_data is not None:
                    self.current_file_label.config(text=file_path.split('/')[-1])
                    
    def load_data_file(self, file_path):
        """加载数据文件"""
        try:
            if file_path.endswith('.csv'):
                data = pd.read_csv(file_path)
            elif file_path.endswith(('.xlsx', '.xls')):
                data = pd.read_excel(file_path)
            elif file_path.endswith('.txt'):
                data = pd.read_csv(file_path, delimiter='\t')
            else:
                messagebox.showerror("错误", "不支持的文件格式")
                return None
                
            # 检查必要列
            if len(data.columns) < 2:
                messagebox.showerror("错误", "数据文件需要至少包含时间和数值两列")
                return None
                
            return data
        except Exception as e:
            messagebox.showerror("错误", f"加载文件失败: {str(e)}")
            return None
            
    def align_time_axis(self, voltage_df, current_df):
        """对齐时间轴"""
        if len(voltage_df) != len(current_df):
            # 如果长度不同，进行重采样
            min_length = min(len(voltage_df), len(current_df))
            return voltage_df.iloc[:min_length], current_df.iloc[:min_length]
        return voltage_df, current_df
        
    def calculate_statistics(self, voltage, current):
        """计算统计信息"""
        stats = {}
        
        # 基本统计
        stats['voltage_min'] = np.min(voltage)
        stats['voltage_max'] = np.max(voltage)
        stats['voltage_mean'] = np.mean(voltage)
        stats['voltage_std'] = np.std(voltage)
        
        stats['current_min'] = np.min(current)
        stats['current_max'] = np.max(current)
        stats['current_mean'] = np.mean(current)
        stats['current_std'] = np.std(current)
        
        # 相关性
        if len(voltage) == len(current):
            correlation = np.corrcoef(voltage, current)[0, 1]
            stats['correlation'] = correlation
            
            # 相位差（简化计算）
            cross_corr = signal.correlate(voltage - np.mean(voltage), 
                                         current - np.mean(current))
            lag = np.argmax(cross_corr) - (len(voltage) - 1)
            stats['phase_lag'] = lag
        else:
            stats['correlation'] = np.nan
            stats['phase_lag'] = 0
            
        # 功率计算
        if len(voltage) == len(current):
            power = voltage * current
            stats['power_min'] = np.min(power)
            stats['power_max'] = np.max(power)
            stats['power_mean'] = np.mean(power)
            stats['power'] = power
        else:
            stats['power'] = None
            
        return stats
        
    def detect_anomalies(self, voltage, current, stats):
        """检测异常点"""
        anomalies = {}
        
        # 过压检测（假设阈值为平均值的±3倍标准差）
        voltage_threshold = stats['voltage_mean'] + 3 * stats['voltage_std']
        under_voltage_threshold = stats['voltage_mean'] - 3 * stats['voltage_std']
        overvoltage_idx = np.where((voltage > voltage_threshold) | (voltage < under_voltage_threshold))[0]
        anomalies['overvoltage'] = overvoltage_idx
        
        # 过流检测
        current_threshold = stats['current_mean'] + 3 * stats['current_std']
        overcurrent_idx = np.where(current > current_threshold)[0]
        anomalies['overcurrent'] = overcurrent_idx
        
        # 异常功耗检测
        if stats['power'] is not None:
            power_threshold = stats['power_mean'] + 3 * np.std(stats['power'])
            anomaly_power_idx = np.where(stats['power'] > power_threshold)[0]
            anomalies['anomaly_power'] = anomaly_power_idx
            
        return anomalies
        
    def generate_chart(self):
        """生成对比图表"""
        if self.voltage_data is None or self.current_data is None:
            messagebox.showwarning("警告", "请先选择电压和电流数据文件")
            return
            
        try:
            # 清理图表
            self.ax1.clear()
            self.ax2.clear()
            
            # 获取数据
            voltage_df = self.voltage_data
            current_df = self.current_data
            
            # 对齐时间轴
            if self.auto_align_var.get():
                voltage_df, current_df = self.align_time_axis(voltage_df, current_df)
                self.time_aligned = True
                
            # 假设第一列是时间，第二列是数值
            voltage_time = voltage_df.iloc[:, 0].values
            voltage_values = voltage_df.iloc[:, 1].values
            
            current_time = current_df.iloc[:, 0].values
            current_values = current_df.iloc[:, 1].values
            
            # 计算统计信息
            stats = self.calculate_statistics(voltage_values, current_values)
            
            # 检测异常
            anomalies = self.detect_anomalies(voltage_values, current_values, stats)
            
            # 更新分析结果
            self.update_analysis_results(stats, anomalies)
            
            # 根据视图模式绘制图表
            if self.view_mode.get() == "overlay":
                self.plot_overlay(voltage_time, voltage_values, current_time, current_values, 
                                stats, anomalies)
            else:
                self.plot_split(voltage_time, voltage_values, current_time, current_values, 
                              stats, anomalies)
                
            self.canvas.draw()
            
        except Exception as e:
            messagebox.showerror("错误", f"生成图表时出错: {str(e)}")
            
    def plot_overlay(self, v_time, v_values, c_time, c_values, stats, anomalies):
        """绘制叠加对比图"""
        # 左Y轴 - 电压
        color1 = 'tab:blue'
        self.ax1.set_xlabel('时间')
        self.ax1.set_ylabel('电压 (V)', color=color1)
        line1 = self.ax1.plot(v_time, v_values, color=color1, label='电压', alpha=0.8)
        self.ax1.tick_params(axis='y', labelcolor=color1)
        
        # 高亮异常点
        if self.highlight_anomaly_var.get() and len(anomalies['overvoltage']) > 0:
            self.ax1.scatter(v_time[anomalies['overvoltage']], 
                           v_values[anomalies['overvoltage']], 
                           color='red', s=50, zorder=5, label='过压点')
            
        # 右Y轴 - 电流
        self.ax2 = self.ax1.twinx()
        color2 = 'tab:red'
        self.ax2.set_ylabel('电流 (A)', color=color2)
        line2 = self.ax2.plot(c_time, c_values, color=color2, label='电流', alpha=0.8)
        self.ax2.tick_params(axis='y', labelcolor=color2)
        
        # 高亮异常点
        if self.highlight_anomaly_var.get() and len(anomalies['overcurrent']) > 0:
            self.ax2.scatter(c_time[anomalies['overcurrent']], 
                           c_values[anomalies['overcurrent']], 
                           color='orange', s=50, zorder=5, label='过流点')
            
        # 添加标题和图例
        title = self.title_entry.get()
        self.ax1.set_title(title)
        
        # 组合图例
        lines1, labels1 = self.ax1.get_legend_handles_labels()
        lines2, labels2 = self.ax2.get_legend_handles_labels()
        self.ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper right')
        
        # 显示趋势线
        if self.show_trend_var.get():
            z1 = np.polyfit(range(len(v_values)), v_values, 1)
            p1 = np.poly1d(z1)
            self.ax1.plot(v_time, p1(range(len(v_values))), 'b--', alpha=0.5, label='电压趋势')
            
            z2 = np.polyfit(range(len(c_values)), c_values, 1)
            p2 = np.poly1d(z2)
            self.ax2.plot(c_time, p2(range(len(c_values))), 'r--', alpha=0.5, label='电流趋势')
            
    def plot_split(self, v_time, v_values, c_time, c_values, stats, anomalies):
        """绘制分屏显示图"""
        # 电压图
        self.ax1.set_title('电压波形')
        self.ax1.set_xlabel('时间')
        self.ax1.set_ylabel('电压 (V)')
        self.ax1.plot(v_time, v_values, 'b-', alpha=0.8, label='电压')
        
        # 高亮异常点
        if self.highlight_anomaly_var.get() and len(anomalies['overvoltage']) > 0:
            self.ax1.scatter(v_time[anomalies['overvoltage']], 
                           v_values[anomalies['overvoltage']], 
                           color='red', s=50, zorder=5, label='过压点')
            
        if self.show_trend_var.get():
            z1 = np.polyfit(range(len(v_values)), v_values, 1)
            p1 = np.poly1d(z1)
            self.ax1.plot(v_time, p1(range(len(v_values))), 'b--', alpha=0.5, label='趋势线')
            
        self.ax1.legend()
        self.ax1.grid(True, alpha=0.3)
        
        # 电流图
        self.ax2.set_title('电流波形')
        self.ax2.set_xlabel('时间')
        self.ax2.set_ylabel('电流 (A)')
        self.ax2.plot(c_time, c_values, 'r-', alpha=0.8, label='电流')
        
        # 高亮异常点
        if self.highlight_anomaly_var.get() and len(anomalies['overcurrent']) > 0:
            self.ax2.scatter(c_time[anomalies['overcurrent']], 
                           c_values[anomalies['overcurrent']], 
                           color='orange', s=50, zorder=5, label='过流点')
            
        if self.show_trend_var.get():
            z2 = np.polyfit(range(len(c_values)), c_values, 1)
            p2 = np.poly1d(z2)
            self.ax2.plot(c_time, p2(range(len(c_values))), 'r--', alpha=0.5, label='趋势线')
            
        self.ax2.legend()
        self.ax2.grid(True, alpha=0.3)
        
        # 调整布局
        self.fig.tight_layout()
        
    def update_analysis_results(self, stats, anomalies):
        """更新分析结果显示"""
        # 相关性分析
        correlation_text = f"相关系数: {stats.get('correlation', 0):.3f}"
        self.correlation_label.config(text=correlation_text)
        
        phase_lag = stats.get('phase_lag', 0)
        phase_text = f"相位差: {abs(phase_lag)} 个采样点"
        self.phase_label.config(text=phase_text)
        
        # 统计信息
        voltage_range_text = f"电压范围: {stats['voltage_min']:.2f}-{stats['voltage_max']:.2f} V"
        self.voltage_range_label.config(text=voltage_range_text)
        
        current_range_text = f"电流范围: {stats['current_min']:.3f}-{stats['current_max']:.3f} A"
        self.current_range_label.config(text=current_range_text)
        
        if stats.get('power_mean') is not None:
            power_text = f"平均功率: {stats['power_mean']:.1f} W"
            self.power_label.config(text=power_text)
            
        # 异常检测
        self.overvoltage_label.config(text=f"过压点: {len(anomalies.get('overvoltage', []))}")
        self.overcurrent_label.config(text=f"过流点: {len(anomalies.get('overcurrent', []))}")
        self.anomaly_power_label.config(text=f"异常功耗点: {len(anomalies.get('anomaly_power', []))}")
        
    def clear_selection(self):
        """清空选择"""
        self.voltage_data = None
        self.current_data = None
        self.voltage_file_label.config(text="未选择")
        self.current_file_label.config(text="未选择")
        self.ax1.clear()
        self.ax2.clear()
        self.canvas.draw()
        
        # 清空分析结果
        self.correlation_label.config(text="相关系数: --")
        self.phase_label.config(text="相位差: -- ms")
        self.voltage_range_label.config(text="电压范围: -- V")
        self.current_range_label.config(text="电流范围: -- A")
        self.power_label.config(text="平均功率: -- W")
        self.overvoltage_label.config(text="过压点: 0")
        self.overcurrent_label.config(text="过流点: 0")
        self.anomaly_power_label.config(text="异常功耗点: 0")
        
    def new_project(self):
        """新建项目"""
        self.clear_selection()
        
    def open_config(self):
        """打开配置"""
        messagebox.showinfo("信息", "打开配置功能")
        
    def save_config(self):
        """保存配置"""
        messagebox.showinfo("信息", "保存配置功能")
        
    def export_chart(self, format_type):
        """导出图表"""
        if format_type == 'png':
            file_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG图片", "*.png")]
            )
            if file_path:
                self.fig.savefig(file_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("成功", f"图表已导出到: {file_path}")
        elif format_type == 'pdf':
            file_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF文档", "*.pdf")]
            )
            if file_path:
                self.fig.savefig(file_path, bbox_inches='tight')
                messagebox.showinfo("成功", f"图表已导出到: {file_path}")
                
    def export_csv(self):
        """导出CSV"""
        if self.voltage_data is not None and self.current_data is not None:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV文件", "*.csv")]
            )
            if file_path:
                combined_data = pd.DataFrame({
                    '时间': self.voltage_data.iloc[:, 0],
                    '电压': self.voltage_data.iloc[:, 1],
                    '电流': self.current_data.iloc[:, 1]
                })
                combined_data.to_csv(file_path, index=False)
                messagebox.showinfo("成功", f"数据已导出到: {file_path}")
                
    def filter_settings(self):
        """滤波设置"""
        messagebox.showinfo("信息", "滤波设置功能")
        
    def calibration_settings(self):
        """标定参数"""
        messagebox.showinfo("信息", "标定参数功能")
        
    def open_calculator(self):
        """打开计算器"""
        import os
        os.system("calc.exe" if os.name == 'nt' else "gnome-calculator")
        
    def show_user_manual(self):
        """显示用户手册"""
        manual_text = """
        通用电压电流对比分析器 Pro 使用说明
        
        1. 数据加载：
           - 分别选择电压和电流数据文件
           - 支持格式：CSV、Excel、TXT
           - 数据应包含时间和数值两列
        
        2. 参数配置：
           - 设置对比标题
           - 选择时间轴对齐方式
           - 配置显示选项
        
        3. 分析功能：
           - 双Y轴对比显示
           - 相关性分析
           - 相位差计算
           - 异常点检测
           - 功率计算
        
        4. 导出功能：
           - 导出PNG图片
           - 导出PDF报告
           - 导出CSV数据
        """
        messagebox.showinfo("用户手册", manual_text)
        
    def show_about(self):
        """关于对话框"""
        about_text = """
        通用电压电流对比分析器 Pro
        版本: 1.0.0
        
        功能特点：
        1. 双Y轴电压电流对比分析
        2. 实时功率计算
        3. 相位差检测
        4. 异常点自动识别
        5. 多种导出格式支持
        
        作者: AI Assistant
        日期: 2026-01-22
        """
        messagebox.showinfo("关于", about_text)

def main():
    """主函数"""
    root = tk.Tk()
    app = VoltageCurrentAnalyzerPro(root)
    root.mainloop()

if __name__ == "__main__":
    main()