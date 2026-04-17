import tkinter as tk
from tkinter import ttk, messagebox
import random
import json
import os

class RandomNumberGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("随机数生成器")
        
        # 根据图片调整窗口大小
        self.root.geometry("850x750")  # 稍微加宽以容纳单位输入框
        self.root.configure(bg='#f0f0f0')
        
        # 加载配置
        self.config_file = "generator_config.json"
        self.load_config()
        
        # 设置样式
        self.setup_styles()
        
        # 创建界面
        self.create_widgets()
        
    def load_config(self):
        """加载配置文件"""
        default_config = {
            "mode": "range",  # 默认改为范围值模式以匹配图片
            "include_unit": True,
            "unit": "A",  # 新增：默认单位
            "auto_copy": False,
            "add_serial": True,
            "add_pass": True,
            "start_number": 33,
            "last_deviation": 0.50  # 根据图片默认值
        }
        
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            except:
                self.config = default_config
        else:
            self.config = default_config
            
    def save_config(self):
        """保存配置"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except:
            pass
        
    def setup_styles(self):
        """设置界面样式"""
        style = ttk.Style()
        style.theme_use('clam')
        
    def create_widgets(self):
        """创建界面组件"""
        # 创建主容器
        main_container = ttk.Frame(self.root, padding="10")
        main_container.pack(fill="both", expand=True)
        
        # 左侧区域
        left_panel = ttk.Frame(main_container, width=380)
        left_panel.pack(side="left", fill="both", padx=(0, 10))
        
        # 右侧区域
        right_panel = ttk.Frame(main_container)
        right_panel.pack(side="right", fill="both", expand=True)
        
        # 左侧面板内容
        title_label = ttk.Label(
            left_panel, 
            text="随机数生成器", 
            font=("微软雅黑", 16, "bold")
        )
        title_label.pack(pady=(0, 15))
        
        # 数据模式选择
        mode_frame = ttk.LabelFrame(left_panel, text="数据模式", padding=15)
        mode_frame.pack(fill="x", pady=10)
        
        mode_radio_frame = ttk.Frame(mode_frame)
        mode_radio_frame.pack(fill="x", pady=5)
        
        self.mode_var = tk.StringVar(value=self.config.get("mode", "range"))
        
        single_btn = ttk.Radiobutton(
            mode_radio_frame,
            text="单数值模式 (如: 11.03)",
            value="single",
            variable=self.mode_var,
            command=self.update_mode
        )
        single_btn.pack(anchor="w", pady=2)
        
        range_btn = ttk.Radiobutton(
            mode_radio_frame,
            text="范围值模式 (如: 0.48-0.87A)",
            value="range",
            variable=self.mode_var,
            command=self.update_mode
        )
        range_btn.pack(anchor="w", pady=2)
        
        # 数字输入框框架
        input_frame = ttk.LabelFrame(left_panel, text="数字输入", padding=15)
        input_frame.pack(fill="x", pady=10)
        
        input_grid = ttk.Frame(input_frame)
        input_grid.pack(fill="x")
        
        # 根据模式显示不同的输入框
        self.value_frame = ttk.Frame(input_grid)
        self.value_frame.pack(fill="x", pady=5)
        
        # 新增：单位输入框架
        unit_frame = ttk.Frame(input_frame)
        unit_frame.pack(fill="x", pady=(5, 0))
        
        ttk.Label(unit_frame, text="单位:", font=("微软雅黑", 10)).pack(side="left", padx=(0, 5))
        
        self.unit_var = tk.StringVar(value=self.config.get("unit", "A"))
        self.unit_entry = ttk.Entry(
            unit_frame,
            textvariable=self.unit_var,
            width=10,
            font=("微软雅黑", 10)
        )
        self.unit_entry.pack(side="left", padx=5)
        
        # 单位提示标签
        unit_hint_label = ttk.Label(
            unit_frame,
            text="(如: A, Ω, V, mA, μA)",
            font=("微软雅黑", 9),
            foreground="gray"
        )
        unit_hint_label.pack(side="left", padx=5)
        
        # 波动范围调节框架
        deviation_frame = ttk.LabelFrame(left_panel, text="波动范围 (±)", padding=15)
        deviation_frame.pack(fill="x", pady=10)
        
        value_frame = ttk.Frame(deviation_frame)
        value_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(value_frame, text="当前值:", font=("微软雅黑", 10)).pack(side="left", padx=5)
        self.deviation_label = ttk.Label(
            value_frame, 
            text=f"{self.config.get('last_deviation', 0.50):.2f}", 
            font=("微软雅黑", 10, "bold"),
            foreground="blue"
        )
        self.deviation_label.pack(side="left", padx=5)
        
        self.deviation_var = tk.DoubleVar(value=self.config.get("last_deviation", 0.50))
        self.deviation_scale = ttk.Scale(
            deviation_frame,
            from_=0.0,
            to=100.0,
            value=self.config.get("last_deviation", 0.50),
            variable=self.deviation_var,
            command=self.update_deviation_label,
            length=300
        )
        self.deviation_scale.pack(fill="x", padx=5, pady=5)
        
        # 快捷按钮框架 - 按照图片中的按钮设置
        quick_buttons_frame = ttk.LabelFrame(deviation_frame, text="快捷设置", padding=10)
        quick_buttons_frame.pack(fill="x", pady=(10, 0))
        
        # 按照图片中的快捷按钮设置
        quick_buttons = [
            ("±0.05", 0.05), ("±0.06", 0.06), ("±0.07", 0.07), ("±0.08", 0.08), ("±0.09", 0.09),
            ("±0.10", 0.10), ("±0.15", 0.15), ("±0.20", 0.20), ("±0.25", 0.25), ("±0.30", 0.30)
        ]
        
        # 第一行按钮
        row1_frame = ttk.Frame(quick_buttons_frame)
        row1_frame.pack(fill="x", pady=2)
        for i in range(5):
            text, value = quick_buttons[i]
            btn = ttk.Button(
                row1_frame,
                text=text,
                width=8,
                command=lambda v=value: self.set_preset_deviation(v)
            )
            btn.pack(side="left", padx=1, expand=True, fill="x")
        
        # 第二行按钮
        row2_frame = ttk.Frame(quick_buttons_frame)
        row2_frame.pack(fill="x", pady=2)
        for i in range(5, 10):
            text, value = quick_buttons[i]
            btn = ttk.Button(
                row2_frame,
                text=text,
                width=8,
                command=lambda v=value: self.set_preset_deviation(v)
            )
            btn.pack(side="left", padx=1, expand=True, fill="x")
        
        # 自定义输入波动范围
        custom_frame = ttk.Frame(deviation_frame)
        custom_frame.pack(fill="x", pady=(10, 0))
        
        ttk.Label(custom_frame, text="自定义:", font=("微软雅黑", 9)).pack(side="left", padx=(0, 5))
        
        self.custom_deviation_var = tk.StringVar()
        self.custom_deviation_entry = ttk.Entry(
            custom_frame,
            textvariable=self.custom_deviation_var,
            width=8,
            font=("微软雅黑", 9)
        )
        self.custom_deviation_entry.pack(side="left", padx=5)
        
        ttk.Label(custom_frame, text="(0.01-5.00)", font=("微软雅黑", 9), foreground="gray").pack(side="left", padx=(0, 5))
        
        set_custom_btn = ttk.Button(
            custom_frame,
            text="设置",
            command=self.set_custom_deviation,
            width=6
        )
        set_custom_btn.pack(side="left", padx=5)
        
        # 表格设置框架 - 只保留自动复制功能
        table_frame = ttk.LabelFrame(left_panel, text="表格设置", padding=15)
        table_frame.pack(fill="x", pady=10)
        
        # 只保留自动复制选项
        self.auto_copy_var = tk.BooleanVar(value=self.config.get("auto_copy", False))
        self.auto_copy_cb = ttk.Checkbutton(
            table_frame,
            text="生成后自动复制到剪贴板",
            variable=self.auto_copy_var
        )
        self.auto_copy_cb.pack(anchor="w", pady=5)
        
        # 功能按钮框架
        button_frame = ttk.Frame(left_panel)
        button_frame.pack(pady=10)
        
        self.generate_btn = ttk.Button(
            button_frame,
            text="生成数据",
            command=self.generate_numbers,
            width=12
        )
        self.generate_btn.pack(side="left", padx=5)
        
        self.generate_10_btn = ttk.Button(
            button_frame,
            text="生成10行",
            command=lambda: self.generate_numbers(10),
            width=12
        )
        self.generate_10_btn.pack(side="left", padx=5)
        
        self.clear_btn = ttk.Button(
            button_frame,
            text="清空",
            command=self.clear_all,
            width=8
        )
        self.clear_btn.pack(side="left", padx=5)
        
        # 右侧面板内容
        result_frame = ttk.LabelFrame(right_panel, text="生成结果", padding=15)
        result_frame.pack(fill="both", expand=True)
        
        # 控制栏
        control_frame = ttk.Frame(result_frame)
        control_frame.pack(fill="x", pady=(0, 10))
        
        # 生成行数设置
        group_frame = ttk.Frame(control_frame)
        group_frame.pack(side="left", fill="x", expand=True)
        
        ttk.Label(group_frame, text="生成行数:", font=("微软雅黑", 10)).pack(side="left", padx=(0, 5))
        
        self.group_var = tk.IntVar(value=6)
        self.group_spinbox = ttk.Spinbox(
            group_frame,
            from_=1,
            to=100,
            textvariable=self.group_var,
            width=8,
            font=("微软雅黑", 10)
        )
        self.group_spinbox.pack(side="left", padx=5)
        
        # 复制按钮
        copy_frame = ttk.Frame(control_frame)
        copy_frame.pack(side="right")
        
        self.copy_excel_btn = ttk.Button(
            copy_frame,
            text="复制到Excel",
            command=self.copy_to_excel_format,
            width=12
        )
        self.copy_excel_btn.pack(side="left", padx=2)
        
        self.copy_all_btn = ttk.Button(
            copy_frame,
            text="复制结果",
            command=self.copy_all_to_clipboard,
            width=10
        )
        self.copy_all_btn.pack(side="left", padx=2)
        
        # 创建Text控件用于显示结果
        text_frame = ttk.Frame(result_frame)
        text_frame.pack(fill="both", expand=True)
        
        self.result_text = tk.Text(
            text_frame,
            height=15,
            width=55,
            font=("Consolas", 11),
            bg='white',
            relief='sunken',
            borderwidth=1,
            wrap='none'
        )
        self.result_text.pack(side="left", fill="both", expand=True)
        
        # 添加滚动条
        v_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.result_text.yview)
        v_scrollbar.pack(side="right", fill="y")
        self.result_text.config(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(result_frame, orient="horizontal", command=self.result_text.xview)
        h_scrollbar.pack(fill="x", pady=(5, 0))
        self.result_text.config(xscrollcommand=h_scrollbar.set)
        
        self.result_text.config(state='disabled')
        
        # 使用说明标签 - 按图片中的文字
        instruction_text = """Excel中转粘贴法（推荐！）：
1. 点击"复制到Excel"按钮
2. 打开Excel表格
3. 点击第一个单元格
4. 按Ctrl+V粘贴
5. 数据会自动分配到不同列
6. 再从Excel复制到Word表格

或者直接在Word中：
1. 确保Word表格已创建
2. 在Word中使用"粘贴特殊"
3. 选择"无格式文本"或"只保留文本\""""
        
        instruction_label = ttk.Label(
            result_frame,
            text=instruction_text,
            font=("微软雅黑", 9),
            foreground="green",
            wraplength=400
        )
        instruction_label.pack(pady=(10, 5), anchor="w")
        
        # 状态标签
        self.status_label = ttk.Label(
            result_frame,
            text="",
            font=("微软雅黑", 9),
            foreground="blue"
        )
        self.status_label.pack(pady=(5, 0), anchor="w")
        
        # 绑定事件
        self.root.bind('<Return>', lambda event: self.generate_numbers())
        
        # 窗口关闭时保存配置
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 初始化
        self.update_mode()
        self.entry1.focus()
        
    def update_mode(self):
        """更新模式显示"""
        # 清除旧的输入框
        for widget in self.value_frame.winfo_children():
            widget.destroy()
        
        mode = self.mode_var.get()
        
        if mode == "single":
            # 单数值模式
            ttk.Label(self.value_frame, text="基准值:", font=("微软雅黑", 10)).pack(side="left", padx=(0, 5))
            self.entry1 = ttk.Entry(self.value_frame, width=12, font=("微软雅黑", 10))
            self.entry1.pack(side="left", padx=5)
            self.entry1.insert(0, "11.0")
            
            # 单数值模式也显示单位
            unit_label = ttk.Label(self.value_frame, text=self.unit_var.get(), font=("微软雅黑", 10))
            unit_label.pack(side="left", padx=(5, 0))
            
        else:
            # 范围值模式
            ttk.Label(self.value_frame, text="数值1:", font=("微软雅黑", 10)).pack(side="left", padx=(0, 5))
            self.entry1 = ttk.Entry(self.value_frame, width=10, font=("微软雅黑", 10))
            self.entry1.pack(side="left", padx=5)
            self.entry1.insert(0, "83.33")
            
            ttk.Label(self.value_frame, text="数值2:", font=("微软雅黑", 10)).pack(side="left", padx=(10, 5))
            self.entry2 = ttk.Entry(self.value_frame, width=10, font=("微软雅黑", 10))
            self.entry2.pack(side="left", padx=5)
            self.entry2.insert(0, "96.38")
            
            # 使用自定义单位
            unit_label = ttk.Label(self.value_frame, text=self.unit_var.get(), font=("微软雅黑", 10))
            unit_label.pack(side="left", padx=(5, 0))
        
    def update_deviation_label(self, value=None):
        """更新波动范围标签"""
        deviation = self.deviation_var.get()
        self.deviation_label.config(text=f"{deviation:.2f}")
        
    def set_preset_deviation(self, value):
        """设置预设波动范围"""
        self.deviation_var.set(value)
        self.update_deviation_label()
        
    def set_custom_deviation(self):
        """设置自定义波动范围"""
        try:
            value_str = self.custom_deviation_var.get().strip()
            if not value_str:
                return
                
            value = float(value_str)
            if 0.0 <= value <= 5.0:
                self.deviation_var.set(value)
                self.update_deviation_label()
                self.custom_deviation_var.set("")  # 清空输入框
            else:
                messagebox.showwarning("输入错误", "请输入0.00到5.00之间的数字")
        except ValueError:
            messagebox.showwarning("输入错误", "请输入有效的数字")
            
    def generate_numbers(self, num_groups=None):
        """生成随机数"""
        try:
            # 获取模式
            mode = self.mode_var.get()
            
            if num_groups is None:
                num_groups = self.group_var.get()
            
            # 获取自定义单位
            unit = self.unit_var.get().strip()
            if not unit:
                unit = "A"  # 默认值
            
            # 启用编辑，清空内容
            self.result_text.config(state='normal')
            self.result_text.delete(1.0, tk.END)
            
            if mode == "single":
                # 单数值模式
                str1 = self.entry1.get().strip()
                if not str1:
                    messagebox.showwarning("输入错误", "请输入基准值！")
                    return
                
                base_value = float(str1)
                deviation = self.deviation_var.get()
                
                results = []
                for i in range(num_groups):
                    # 生成随机数
                    random_value = random.uniform(max(0, base_value - deviation), base_value + deviation)
                    
                    # 只保留数值，去掉序号和PASS
                    value_str = f"{random_value:.2f}{unit}"
                    results.append(value_str)
                    self.result_text.insert(tk.END, f"{value_str}\n")
                    
                # 更新状态
                min_val = max(0, base_value - deviation)
                max_val = base_value + deviation
                self.status_label.config(
                    text=f"生成完成！数值范围: [{min_val:.3f}~{max_val:.3f}] | 共{num_groups}行"
                )
                
            else:
                # 范围值模式
                str1 = self.entry1.get().strip()
                str2 = self.entry2.get().strip()
                
                if not str1 or not str2:
                    messagebox.showwarning("输入错误", "请输入两个数值！")
                    return
                
                a = float(str1)
                b = float(str2)
                deviation = self.deviation_var.get()
                
                results = []
                for i in range(num_groups):
                    # 生成两个随机数
                    r1 = random.uniform(max(0, a - deviation), a + deviation)
                    r2 = random.uniform(max(0, b - deviation), b + deviation)
                    
                    # 保持相对顺序
                    if (a < b and r1 > r2) or (a > b and r1 < r2):
                        r1, r2 = r2, r1
                    
                    # 只保留数值范围，去掉序号和PASS
                    value_str = f"{r1:.2f}-{r2:.2f}{unit}"
                    results.append(value_str)
                    self.result_text.insert(tk.END, f"{value_str}\n")
                
                # 更新状态
                min1 = max(0, a - deviation)
                max1 = a + deviation
                min2 = max(0, b - deviation)
                max2 = b + deviation
                self.status_label.config(
                    text=f"生成完成！数值1范围: [{min1:.3f}~{max1:.3f}] | 数值2范围: [{min2:.3f}~{max2:.3f}] | 共{num_groups}行"
                )
            
            # 禁用编辑
            self.result_text.config(state='disabled')
            
            # 滚动到顶部
            self.result_text.see(1.0)
            
            # 自动复制
            if self.auto_copy_var.get() and results:
                self.copy_to_excel_format()
            
        except ValueError:
            messagebox.showerror("输入错误", "请输入有效的数字！")
        except Exception as e:
            messagebox.showerror("错误", f"发生错误: {str(e)}")
    
    def copy_to_excel_format(self):
        """复制为Excel可识别的格式"""
        try:
            text = self.result_text.get(1.0, tk.END).strip()
            if not text:
                messagebox.showinfo("提示", "没有可复制的内容")
                return
            
            # 使用制表符分隔的文本，Excel可以自动识别
            excel_text = text
            
            # 修复：使用正确的剪贴板方法
            self.root.clipboard_clear()
            self.root.clipboard_append(excel_text)  # 修复这里：使用clipboard_append而不是clipboard.append
            
            # 显示成功提示
            original_text = self.copy_excel_btn.cget("text")
            self.copy_excel_btn.config(text="已复制")
            self.root.after(1000, lambda: self.copy_excel_btn.config(text="复制到Excel"))
            
            # 显示详细的使用说明
            excel_guide = """✅ 已成功复制为Excel格式！

🎯 使用方法（Excel中转法）：
1. 打开Microsoft Excel
2. 点击任意单元格（建议从A1开始）
3. 按 Ctrl+V 粘贴
4. 数据会自动分配到不同列

🎯 然后复制到Word：
1. 在Excel中选中已粘贴的数据区域
2. 按 Ctrl+C 复制
3. 切换到Word表格
4. 点击第一个单元格
5. 按 Ctrl+V 粘贴

🎯 或者直接在Word中：
1. 确保Word表格已创建
2. 在表格的第一个单元格右键
3. 选择"粘贴特殊"
4. 选择"无格式文本"或"只保留文本\""""
            
            messagebox.showinfo("Excel中转粘贴指南", excel_guide)
            
        except Exception as e:
            messagebox.showerror("复制错误", f"复制失败: {str(e)}")
    
    def copy_all_to_clipboard(self, auto_mode=False):
        """复制所有结果到剪贴板"""
        try:
            text = self.result_text.get(1.0, tk.END).strip()
            if text:
                # 修复：使用正确的剪贴板方法
                self.root.clipboard_clear()
                self.root.clipboard_append(text)  # 修复这里：使用clipboard_append而不是clipboard.append
                
                if not auto_mode:
                    original_text = self.copy_all_btn.cget("text")
                    self.copy_all_btn.config(text="已复制")
                    self.root.after(1000, lambda: self.copy_all_btn.config(text="复制结果"))
                    
                return True
            else:
                if not auto_mode:
                    messagebox.showinfo("提示", "没有可复制的内容")
                return False
                
        except Exception as e:
            if not auto_mode:
                messagebox.showerror("复制错误", f"复制失败: {str(e)}")
            return False
    
    def clear_all(self):
        """清空所有输入和结果"""
        self.result_text.config(state='normal')
        self.result_text.delete(1.0, tk.END)
        self.result_text.config(state='disabled')
        self.status_label.config(text="")
        
        if hasattr(self, 'entry1'):
            self.entry1.delete(0, tk.END)
        if hasattr(self, 'entry2') and hasattr(self, 'mode_var') and self.mode_var.get() == "range":
            self.entry2.delete(0, tk.END)
        
        if hasattr(self, 'entry1'):
            self.entry1.focus()
    
    def on_closing(self):
        """窗口关闭时保存配置"""
        self.config["mode"] = self.mode_var.get()
        self.config["unit"] = self.unit_var.get()  # 保存单位设置
        self.config["auto_copy"] = self.auto_copy_var.get()
        self.config["last_deviation"] = self.deviation_var.get()
        self.save_config()
        self.root.destroy()

def main():
    root = tk.Tk()
    app = RandomNumberGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()