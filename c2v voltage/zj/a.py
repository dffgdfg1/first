import os
import tkinter as tk
from tkinter import messagebox, ttk
from pathlib import Path

class FolderGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("批量文件夹生成器 v3.0")
        self.root.geometry("580x500")  # 增加窗口高度以容纳新功能
        self.root.resizable(False, False)
        
        # 设置样式
        self.setup_styles()
        # 创建界面
        self.create_widgets()
    
    def setup_styles(self):
        # 创建样式
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置标签样式
        style.configure('Title.TLabel', font=('微软雅黑', 16, 'bold'), foreground='#2c3e50')
        style.configure('Normal.TLabel', font=('微软雅黑', 10), foreground='#34495e')
        
        # 配置按钮样式
        style.configure('Primary.TButton', font=('微软雅黑', 10, 'bold'), 
                       padding=10, background='#3498db', foreground='white')
        style.map('Primary.TButton',
                 background=[('active', '#2980b9')])
        
        style.configure('Success.TButton', font=('微软雅黑', 10, 'bold'), 
                       padding=10, background='#2ecc71', foreground='white')
        style.map('Success.TButton',
                 background=[('active', '#27ae60')])
        
        style.configure('Info.TButton', font=('微软雅黑', 10, 'bold'), 
                       padding=10, background='#9b59b6', foreground='white')
        style.map('Info.TButton',
                 background=[('active', '#8e44ad')])
        
        style.configure('Warning.TButton', font=('微软雅黑', 10, 'bold'), 
                       padding=10, background='#e67e22', foreground='white')
        style.map('Warning.TButton',
                 background=[('active', '#d35400')])
        
        # 配置单选按钮样式
        style.configure('TRadiobutton', font=('微软雅黑', 10))
    
    def create_widgets(self):
        # 标题
        title_label = ttk.Label(self.root, text="📁 批量文件夹生成器 v3.0", style='Title.TLabel')
        title_label.pack(pady=20)
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 起始序号
        ttk.Label(main_frame, text="起始序号:", style='Normal.TLabel').grid(
            row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.start_var = tk.StringVar(value="11")
        self.start_entry = ttk.Entry(main_frame, textvariable=self.start_var, 
                                    font=('微软雅黑', 10), width=20)
        self.start_entry.grid(row=0, column=1, padx=(10, 0), pady=(0, 5), sticky=tk.W)
        
        # 结束序号
        ttk.Label(main_frame, text="结束序号:", style='Normal.TLabel').grid(
            row=1, column=0, sticky=tk.W, pady=5)
        self.end_var = tk.StringVar(value="33")
        self.end_entry = ttk.Entry(main_frame, textvariable=self.end_var, 
                                  font=('微软雅黑', 10), width=20)
        self.end_entry.grid(row=1, column=1, padx=(10, 0), pady=5, sticky=tk.W)
        
        # 文件夹名称 - 修改这里：不再自动去除空格
        ttk.Label(main_frame, text="文件夹名称:", style='Normal.TLabel').grid(
            row=2, column=0, sticky=tk.W, pady=5)
        self.prefix_var = tk.StringVar(value="测试数据")
        # 注意：我们需要在获取值时处理空格，但不在 StringVar 初始化时去除
        self.prefix_entry = ttk.Entry(main_frame, textvariable=self.prefix_var, 
                                     font=('微软雅黑', 10), width=20)
        self.prefix_entry.grid(row=2, column=1, padx=(10, 0), pady=5, sticky=tk.W)
        
        # 序号前缀（新功能）- 修改这里：不再自动去除空格
        ttk.Label(main_frame, text="序号前缀:", style='Normal.TLabel').grid(
            row=3, column=0, sticky=tk.W, pady=5)
        self.prefix_before_num_var = tk.StringVar(value="")
        self.prefix_before_num_entry = ttk.Entry(main_frame, textvariable=self.prefix_before_num_var, 
                                               font=('微软雅黑', 10), width=20)
        self.prefix_before_num_entry.grid(row=3, column=1, padx=(10, 0), pady=5, sticky=tk.W)
        
        # 序号单位 - 修改这里：不再自动去除空格
        ttk.Label(main_frame, text="序号单位:", style='Normal.TLabel').grid(
            row=4, column=0, sticky=tk.W, pady=5)
        
        # 创建框架用于放置单位输入框和位置选项
        unit_position_frame = ttk.Frame(main_frame)
        unit_position_frame.grid(row=4, column=1, padx=(10, 0), pady=5, sticky=tk.W+tk.E)
        
        # 单位输入框
        self.unit_var = tk.StringVar(value="")
        self.unit_entry = ttk.Entry(unit_position_frame, textvariable=self.unit_var, 
                                   font=('微软雅黑', 10), width=8)
        self.unit_entry.pack(side=tk.LEFT)
        
        # 添加序号位置选项
        ttk.Label(unit_position_frame, text="  序号位置:", style='Normal.TLabel').pack(side=tk.LEFT, padx=(10, 5))
        
        # 序号位置选择
        self.position_var = tk.StringVar(value="front")  # 默认在前面
        position_frame = ttk.Frame(unit_position_frame)
        position_frame.pack(side=tk.LEFT)
        
        ttk.Radiobutton(position_frame, text="在前面", variable=self.position_var, 
                       value="front", command=self.update_info).pack(side=tk.LEFT)
        ttk.Radiobutton(position_frame, text="在后面", variable=self.position_var, 
                       value="back", command=self.update_info).pack(side=tk.LEFT, padx=(10, 0))
        
        # 序号位数选项
        ttk.Label(main_frame, text="序号位数:", style='Normal.TLabel').grid(
            row=5, column=0, sticky=tk.W, pady=5)
        
        digits_frame = ttk.Frame(main_frame)
        digits_frame.grid(row=5, column=1, padx=(10, 0), pady=5, sticky=tk.W)
        
        self.digits_var = tk.StringVar(value="auto")
        ttk.Radiobutton(digits_frame, text="自动", variable=self.digits_var, 
                       value="auto", command=self.update_info).pack(side=tk.LEFT)
        ttk.Radiobutton(digits_frame, text="固定:", variable=self.digits_var, 
                       value="fixed", command=self.update_info).pack(side=tk.LEFT, padx=(10, 0))
        
        self.digits_num_var = tk.StringVar(value="2")
        self.digits_entry = ttk.Entry(digits_frame, textvariable=self.digits_num_var, 
                                     font=('微软雅黑', 10), width=5, state='disabled')
        self.digits_entry.pack(side=tk.LEFT, padx=(5, 0))
        
        # 保存路径
        ttk.Label(main_frame, text="保存路径:", style='Normal.TLabel').grid(
            row=6, column=0, sticky=tk.W, pady=5)
        
        path_frame = ttk.Frame(main_frame)
        path_frame.grid(row=6, column=1, padx=(10, 0), pady=5, sticky=tk.W+tk.E)
        
        self.path_var = tk.StringVar(value=str(Path.home() / "Desktop" / "批量文件夹"))
        self.path_entry = ttk.Entry(path_frame, textvariable=self.path_var, 
                                   font=('微软雅黑', 10))
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        browse_btn = ttk.Button(path_frame, text="浏览", width=8,
                               command=self.browse_directory)
        browse_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        # 信息显示区域
        self.info_label = ttk.Label(main_frame, text="", style='Normal.TLabel',
                                   foreground='#2c3e50', wraplength=450)
        self.info_label.grid(row=7, column=0, columnspan=2, pady=(20, 10))
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=8, column=0, columnspan=2, pady=20)
        
        # 生成按钮
        generate_btn = ttk.Button(button_frame, text="🚀 生成文件夹", 
                                 style='Success.TButton', command=self.generate_folders)
        generate_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 重置按钮
        reset_btn = ttk.Button(button_frame, text="🔄 重置", 
                              style='Info.TButton', command=self.reset_to_example)
        reset_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 测试按钮
        test_btn = ttk.Button(button_frame, text="🧪 测试名称", 
                             style='Warning.TButton', command=self.test_folder_names)
        test_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 清空按钮
        clear_btn = ttk.Button(button_frame, text="🗑️ 清空", 
                              style='Primary.TButton', command=self.clear_fields)
        clear_btn.pack(side=tk.LEFT)
        
        # 绑定事件
        self.start_entry.bind('<KeyRelease>', self.update_info)
        self.end_entry.bind('<KeyRelease>', self.update_info)
        self.prefix_entry.bind('<KeyRelease>', self.update_info)
        self.prefix_before_num_entry.bind('<KeyRelease>', self.update_info)
        self.unit_entry.bind('<KeyRelease>', self.update_info)
        self.digits_entry.bind('<KeyRelease>', self.update_info)
        
        # 绑定位数模式改变事件
        self.digits_var.trace('w', self.on_digits_mode_changed)
        
        # 设置示例
        self.reset_to_example()
    
    def on_digits_mode_changed(self, *args):
        """当序号位数模式改变时启用/禁用位数输入框"""
        if self.digits_var.get() == "fixed":
            self.digits_entry.config(state='normal')
        else:
            self.digits_entry.config(state='disabled')
        self.update_info()
    
    def update_info(self, event=None):
        try:
            start = int(self.start_var.get())
            end = int(self.end_var.get())
            # 修改这里：直接从 Entry 小部件获取值，不通过 strip()
            prefix = self.prefix_entry.get()
            prefix_before = self.prefix_before_num_entry.get()  # 新功能：序号前缀
            unit = self.unit_entry.get()
            position = self.position_var.get()
            digits_mode = self.digits_var.get()
            
            if start > end:
                self.info_label.config(text="❌ 起始序号不能大于结束序号！", foreground='#e74c3c')
                return
            
            count = end - start + 1
            if count <= 0:
                self.info_label.config(text="❌ 请检查输入参数！", foreground='#e74c3c')
                return
            
            # 构建示例
            example1 = self.generate_folder_name(start, prefix, prefix_before, unit, position, digits_mode)
            example2 = self.generate_folder_name(end, prefix, prefix_before, unit, position, digits_mode)
            
            example = f"例如：{example1}"
            if count > 1:
                if count <= 5:
                    examples = []
                    for i in range(start, min(start+3, end+1)):
                        examples.append(self.generate_folder_name(i, prefix, prefix_before, unit, position, digits_mode))
                    if count > 3:
                        examples.append("...")
                        examples.append(example2)
                    example = f"例如：{', '.join(examples)}"
                else:
                    mid_example = self.generate_folder_name(start+1, prefix, prefix_before, unit, position, digits_mode)
                    example = f"例如：{example1}, {mid_example}, ..., {example2}"
            
            # 显示生成信息
            position_text = "前" if position == "front" else "后"
            
            prefix_before_text = f"，序号前缀：'{prefix_before}'" if prefix_before else ""
            unit_text = f"，单位：'{unit}'" if unit else ""
            
            if digits_mode == "fixed":
                try:
                    digits = int(self.digits_num_var.get())
                    if digits > 0:
                        digits_text = f"，固定{digits}位"
                    else:
                        digits_text = ""
                except ValueError:
                    digits_text = ""
            else:
                digits_text = ""
            
            # 显示空格提示
            space_note = ""
            if ' ' in prefix or ' ' in prefix_before or ' ' in unit:
                space_note = "\n⚠️ 注意：文件夹名称包含空格"
            
            self.info_label.config(
                text=f"📊 将生成 {count} 个文件夹\n序号在{position_text}面{prefix_before_text}{unit_text}{digits_text}\n{example}{space_note}",
                foreground='#2c3e50'
            )
            
        except ValueError:
            self.info_label.config(text="❌ 请输入有效的数字！", foreground='#e74c3c')
        except Exception as e:
            self.info_label.config(text=f"❌ 错误：{str(e)}", foreground='#e74c3c')
    
    def generate_folder_name(self, num, prefix, prefix_before, unit, position, digits_mode):
        """生成文件夹名称的辅助函数"""
        # 处理位数补零
        if digits_mode == "fixed":
            try:
                digits = int(self.digits_num_var.get())
                if digits > 0:
                    num_str = str(num).zfill(digits)
                else:
                    num_str = str(num)
            except ValueError:
                num_str = str(num)
        else:
            num_str = str(num)
        
        # 组合各部分
        prefix_part = f"{prefix_before}{num_str}" if prefix_before else num_str
        
        # 根据位置构建文件夹名
        if position == "front":
            if unit:
                return f"{prefix_part}{unit}{prefix}"
            else:
                return f"{prefix_part}{prefix}"
        else:  # 序号在后面
            if unit:
                return f"{prefix}{prefix_part}{unit}"
            else:
                return f"{prefix}{prefix_part}"
    
    def browse_directory(self):
        from tkinter import filedialog
        directory = filedialog.askdirectory(
            title="选择保存文件夹的目录",
            initialdir=self.path_var.get()
        )
        if directory:
            self.path_var.set(directory)
    
    def generate_folders(self):
        try:
            # 获取输入值 - 修改这里：直接从 Entry 小部件获取值
            start = int(self.start_var.get())
            end = int(self.end_var.get())
            prefix = self.prefix_entry.get()  # 直接获取，保留空格
            prefix_before = self.prefix_before_num_entry.get()  # 直接获取，保留空格
            unit = self.unit_entry.get()  # 直接获取，保留空格
            position = self.position_var.get()
            digits_mode = self.digits_var.get()
            save_path = self.path_var.get().strip()  # 路径仍然可以去除首尾空格
            
            # 验证输入 - 修改这里：允许纯空格或空字符串
            if not prefix and not prefix.strip():
                messagebox.showerror("错误", "文件夹名称不能为空！")
                return
            
            if start > end:
                messagebox.showerror("错误", "起始序号不能大于结束序号！")
                return
            
            if not save_path:
                messagebox.showerror("错误", "请选择保存路径！")
                return
            
            # 获取固定位数
            fixed_digits = None
            if digits_mode == "fixed":
                try:
                    fixed_digits = int(self.digits_num_var.get())
                    if fixed_digits <= 0:
                        fixed_digits = None
                except ValueError:
                    fixed_digits = None
            
            # 创建保存目录
            save_dir = Path(save_path)
            try:
                save_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建目录：{e}")
                return
            
            # 生成文件夹
            created_count = 0
            failed_folders = []
            
            for i in range(start, end + 1):
                folder_name = self.generate_folder_name(i, prefix, prefix_before, unit, position, digits_mode)
                folder_path = save_dir / folder_name
                
                try:
                    folder_path.mkdir(exist_ok=False)
                    created_count += 1
                    print(f"✓ 已创建: {folder_path}")
                except FileExistsError:
                    failed_folders.append(folder_name)
                except Exception as e:
                    failed_folders.append(f"{folder_name} (错误: {str(e)})")
            
            # 显示结果
            result_message = f"✅ 完成！\n\n成功创建: {created_count} 个文件夹\n保存位置: {save_path}"
            
            if failed_folders:
                if len(failed_folders) > 5:
                    result_message += f"\n\n失败/已存在: {len(failed_folders)} 个文件夹\n(前5个: {', '.join(failed_folders[:5])}...)"
                else:
                    result_message += f"\n\n失败/已存在的文件夹:\n{', '.join(failed_folders)}"
            
            messagebox.showinfo("生成完成", result_message)
            
            # 更新信息显示
            self.info_label.config(
                text=f"✅ 已生成 {created_count} 个文件夹到:\n{save_path}",
                foreground='#27ae60'
            )
            
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字！")
        except Exception as e:
            messagebox.showerror("错误", f"生成过程中出现错误：{str(e)}")
    
    def test_folder_names(self):
        """测试生成的文件夹名称"""
        try:
            start = int(self.start_var.get())
            end = int(self.end_var.get())
            # 修改这里：直接从 Entry 小部件获取值
            prefix = self.prefix_entry.get()
            prefix_before = self.prefix_before_num_entry.get()
            unit = self.unit_entry.get()
            position = self.position_var.get()
            digits_mode = self.digits_var.get()
            
            if start > end:
                messagebox.showerror("错误", "起始序号不能大于结束序号！")
                return
            
            count = end - start + 1
            if count <= 0:
                messagebox.showerror("错误", "请检查输入参数！")
                return
            
            # 显示前5个文件夹名称
            folder_names = []
            for i in range(start, min(start + 5, end + 1)):
                folder_name = self.generate_folder_name(i, prefix, prefix_before, unit, position, digits_mode)
                folder_names.append(folder_name)
            
            if count > 5:
                folder_names.append(f"...")
                last_name = self.generate_folder_name(end, prefix, prefix_before, unit, position, digits_mode)
                folder_names.append(last_name)
            
            names_text = "\n".join(folder_names)
            messagebox.showinfo("测试文件夹名称", 
                              f"📋 文件夹名称预览 (共{count}个):\n\n{names_text}")
            
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字！")
        except Exception as e:
            messagebox.showerror("错误", f"测试过程中出现错误：{str(e)}")
    
    def reset_to_example(self):
        """重置为示例值"""
        self.start_var.set("11")
        self.end_var.set("33")
        # 修改这里：通过 StringVar 的 set 方法设置初始值
        self.prefix_var.set("测试数据")
        self.prefix_before_num_var.set("")  # 新功能：序号前缀
        self.unit_var.set("")
        self.position_var.set("front")
        self.digits_var.set("auto")
        self.digits_num_var.set("2")
        self.path_var.set(str(Path.home() / "Desktop" / "批量文件夹"))
        self.on_digits_mode_changed()  # 更新位数输入框状态
        self.update_info()
    
    def clear_fields(self):
        """清空所有字段"""
        self.start_var.set("1")
        self.end_var.set("10")
        # 修改这里：通过 StringVar 的 set 方法清空
        self.prefix_var.set("文件夹")
        self.prefix_before_num_var.set("")  # 新功能：序号前缀
        self.unit_var.set("")
        self.position_var.set("front")
        self.digits_var.set("auto")
        self.digits_num_var.set("2")
        self.path_var.set(str(Path.home() / "Desktop" / "批量文件夹"))
        self.on_digits_mode_changed()  # 更新位数输入框状态
        self.update_info()

def main():
    root = tk.Tk()
    app = FolderGeneratorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()