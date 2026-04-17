import sys
import os
import pandas as pd
import numpy as np
from matplotlib.figure import Figure

# 手动设置matplotlib后端为PyQt6
import matplotlib
matplotlib.use('QtAgg')
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar

from PyQt6.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                             QWidget, QPushButton, QLabel, QFileDialog, QMessageBox,
                             QGroupBox, QSpinBox, QDoubleSpinBox, QProgressBar,
                             QCheckBox, QComboBox, QTextEdit, QDialog)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont

class DataProcessor(QThread):
    """数据处理线程"""
    progress_updated = pyqtSignal(int)
    data_loaded = pyqtSignal(pd.DataFrame, dict)  # 返回DataFrame和列信息字典
    error_occurred = pyqtSignal(str)
    
    def __init__(self, file1_path, file2_path, auto_align=True):
        super().__init__()
        self.file1_path = file1_path
        self.file2_path = file2_path
        self.auto_align = auto_align
    
    def detect_time_column(self, df):
        """智能识别时间列"""
        # 首先检查列名中是否包含时间关键词
        time_keywords = ['time', '时间', 'datetime', 'date', 'timestamp', '日期', '时刻']
        
        for col in df.columns:
            col_str = str(col).lower()
            for keyword in time_keywords:
                if keyword in col_str:
                    return col
        
        # 如果没有找到包含关键词的列，检查数据类型
        for col in df.columns:
            try:
                # 检查是否为日期时间类型
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    return col
                
                # 检查是否为时间格式的字符串
                sample = df[col].iloc[0] if len(df[col]) > 0 else None
                if isinstance(sample, (pd.Timestamp, np.datetime64)):
                    return col
                
                # 尝试将第一行转换为datetime
                try:
                    pd.to_datetime(df[col].head(10), errors='raise')
                    return col
                except:
                    pass
                    
            except:
                continue
        
        # 如果还找不到，使用第一列
        print(f"警告：未找到时间列，使用第一列 '{df.columns[0]}' 作为时间列")
        return df.columns[0]
    
    def convert_time_to_seconds(self, time_series):
        """将时间转换为秒数"""
        try:
            # 如果是datetime类型，转换为秒数
            if pd.api.types.is_datetime64_any_dtype(time_series):
                return (time_series - time_series.min()).dt.total_seconds()
            
            # 尝试转换为datetime
            try:
                datetime_series = pd.to_datetime(time_series, errors='coerce')
                if not datetime_series.isnull().all():
                    return (datetime_series - datetime_series.min()).dt.total_seconds()
            except:
                pass
            
            # 如果是数值类型，直接返回
            if pd.api.types.is_numeric_dtype(time_series):
                return time_series
            
            # 尝试转换为数值
            numeric_series = pd.to_numeric(time_series, errors='coerce')
            if not numeric_series.isnull().all():
                return numeric_series
            
            # 如果转换失败，使用索引
            print("警告：无法转换时间列，使用索引作为时间")
            return pd.Series(range(len(time_series)))
            
        except Exception as e:
            print(f"时间转换错误: {e}")
            return pd.Series(range(len(time_series)))
    
    def align_time_series(self, df1, df2):
        """对齐两个时间序列"""
        if 'Time' not in df1.columns or 'Time' not in df2.columns:
            raise ValueError("数据中缺少Time列")
            
        # 转换为秒数
        time1 = self.convert_time_to_seconds(df1['Time']).values
        time2 = self.convert_time_to_seconds(df2['Time']).values
        
        # 计算时间范围
        min_time = min(time1.min(), time2.min())
        max_time = max(time1.max(), time2.max())
        
        # 计算时间间隔
        if len(time1) > 1:
            interval1 = (time1.max() - time1.min()) / len(time1)
        else:
            interval1 = 1
            
        if len(time2) > 1:
            interval2 = (time2.max() - time2.min()) / len(time2)
        else:
            interval2 = 1
        
        # 选择较大的时间间隔作为基准
        target_interval = max(interval1, interval2)
        
        # 创建统一的时间轴
        aligned_time = np.arange(min_time, max_time + target_interval, target_interval)
        
        return aligned_time, target_interval
    
    def resample_data(self, df, target_time):
        """将数据重采样到目标时间轴"""
        if len(df) == 0 or 'Time' not in df.columns:
            return pd.DataFrame({'Time': target_time})
        
        resampled_data = pd.DataFrame({'Time': target_time})
        
        # 原始时间数据
        original_time = self.convert_time_to_seconds(df['Time']).values
        
        for col in df.columns:
            if col == 'Time':
                continue
                
            try:
                y = pd.to_numeric(df[col], errors='coerce').values
                mask = ~np.isnan(y)
                
                if np.sum(mask) > 1:  # 至少需要2个点才能插值
                    valid_time = original_time[mask]
                    valid_y = y[mask]
                    
                    # 使用线性插值
                    resampled_y = np.interp(target_time, valid_time, valid_y)
                    resampled_data[col] = resampled_y

                else:
                    resampled_data[col] = np.nan
                    
            except Exception as e:
                print(f"列 {col} 插值失败: {e}")
                resampled_data[col] = np.nan
        
        return resampled_data
    
    def identify_temperature_column(self, df):
        """识别温度列类型"""
        column_types = {}
        
        for col in df.columns:
            if col == 'Time':
                continue
                
            try:
                data = pd.to_numeric(df[col], errors='coerce').dropna()
                if len(data) < 2:
                    continue
                
                variance = data.var()
                if pd.isna(variance):
                    variance = 0
                    
                unique_ratio = len(data.unique()) / len(data) if len(data) > 0 else 0
                col_lower = str(col).lower()
                
                # 基于列名判断
                if any(pattern in col_lower for pattern in ['sv', 'sp', '设定', 'set', 'target']):
                    column_types[col] = 'Temperature_SV'
                elif any(pattern in col_lower for pattern in ['s1', '热电偶', '外部', 'ext']):
                    column_types[col] = 'Temperature_S1'
                elif any(pattern in col_lower for pattern in ['pv', '实际', '监控', 'actual', 'monitor']):
                    column_types[col] = 'Temperature_PV'
                else:
                    # 基于数据特征判断
                    if '温度' in col_lower or 'temp' in col_lower or 't ' in col_lower:
                        if unique_ratio < 0.3:  # 变化少
                            column_types[col] = 'Temperature_SV'
                        elif variance > 1:  # 波动大
                            column_types[col] = 'Temperature_PV'
                        else:  # 相对稳定
                            column_types[col] = 'Temperature_S1'
                        
            except Exception as e:
                print(f"识别列 {col} 类型时出错: {e}")
                continue
        
        return column_types
    
    def process_single_file(self, file_path):
        """处理单个文件"""
        try:
            print(f"正在处理文件: {os.path.basename(file_path)}")
            
            # 读取Excel文件
            try:
                if file_path.endswith('.xls'):
                    df = pd.read_excel(file_path, engine='xlrd')
                else:
                    df = pd.read_excel(file_path, engine='openpyxl')
            except Exception as e:
                print(f"使用默认引擎读取文件: {e}")
                df = pd.read_excel(file_path)
            
            print(f"文件列名: {list(df.columns)}")
            
            # 识别时间列
            time_col = self.detect_time_column(df)
            print(f"识别到时间列: {time_col}")
            
            # 识别温度列类型
            temp_types = self.identify_temperature_column(df)
            print(f"识别到的温度列: {temp_types}")
            
            # 只保留时间列和识别到的温度列
            cols_to_keep = [time_col] + list(temp_types.keys())
            print(f"保留的列: {cols_to_keep}")
            
            df_selected = df[cols_to_keep].copy()
            
            # 重命名列
            rename_dict = {time_col: 'Time'}
            for col, col_type in temp_types.items():
                rename_dict[col] = col_type
            
            df_selected = df_selected.rename(columns=rename_dict)
            
            # 确保时间列为数值
            print(f"转换时间列前数据类型: {df_selected['Time'].dtype}")
            df_selected['Time'] = self.convert_time_to_seconds(df_selected['Time'])
            print(f"转换时间列后数据类型: {df_selected['Time'].dtype}")
            
            # 确保温度列为数值
            for col in ['Temperature_S1', 'Temperature_SV', 'Temperature_PV']:
                if col in df_selected.columns:
                    df_selected[col] = pd.to_numeric(df_selected[col], errors='coerce')
                    print(f"列 {col}: {df_selected[col].notna().sum()} 个有效数据点")
            
            print(f"处理后列名: {list(df_selected.columns)}")
            return df_selected
            
        except Exception as e:
            print(f"处理文件 {os.path.basename(file_path)} 时详细错误:")
            import traceback
            traceback.print_exc()
            raise ValueError(f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}")
    
    def run(self):
        try:
            self.progress_updated.emit(10)
            
            print("开始处理文件...")
            
            # 处理两个文件
            df1 = self.process_single_file(self.file1_path)
            self.progress_updated.emit(30)
            
            df2 = self.process_single_file(self.file2_path)
            self.progress_updated.emit(50)
            
            print(f"文件1列名: {list(df1.columns)}")
            print(f"文件2列名: {list(df2.columns)}")
            print(f"文件1形状: {df1.shape}")
            print(f"文件2形状: {df2.shape}")
            
            # 对齐时间序列
            if self.auto_align:
                print("开始对齐时间序列...")
                aligned_time, interval = self.align_time_series(df1, df2)
                print(f"对齐后时间范围: {aligned_time[0]:.2f} 到 {aligned_time[-1]:.2f}, 间隔: {interval:.2f}")
                
                # 重采样数据
                df1_aligned = self.resample_data(df1, aligned_time)
                df2_aligned = self.resample_data(df2, aligned_time)
                
                print(f"对齐后文件1列名: {list(df1_aligned.columns)}")
                print(f"对齐后文件2列名: {list(df2_aligned.columns)}")
                
                # 合并数据前重命名列以避免冲突
                df1_aligned = df1_aligned.add_suffix('_1')
                df2_aligned = df2_aligned.add_suffix('_2')
                
                # 重命名时间列
                df1_aligned = df1_aligned.rename(columns={'Time_1': 'Time'})
                df2_aligned = df2_aligned.rename(columns={'Time_2': 'Time'})
                
                # 合并数据
                merged_df = pd.merge(df1_aligned, df2_aligned, on='Time', how='outer')
                print(f"合并后列名: {list(merged_df.columns)}")
                
            else:
                print("不使用时间对齐，直接合并...")
                # 重命名列以避免冲突
                df1 = df1.add_suffix('_1')
                df2 = df2.add_suffix('_2')
                df1 = df1.rename(columns={'Time_1': 'Time'})
                df2 = df2.rename(columns={'Time_2': 'Time'})
                
                merged_df = pd.concat([df1, df2], axis=1)
            
            self.progress_updated.emit(70)
            
            # 检查合并后的数据
            print(f"合并后数据形状: {merged_df.shape}")
            print(f"合并后列名: {list(merged_df.columns)}")
            
            # 合并相同类型的列
            for temp_type in ['Temperature_S1', 'Temperature_SV', 'Temperature_PV']:
                cols_1 = [col for col in merged_df.columns if col.startswith(f"{temp_type}_1")]
                cols_2 = [col for col in merged_df.columns if col.startswith(f"{temp_type}_2")]
                
                print(f"查找 {temp_type}: 文件1列={cols_1}, 文件2列={cols_2}")
                
                if cols_1 and cols_2:
                    # 合并两个文件的同类型数据
                    col1 = cols_1[0]
                    col2 = cols_2[0]
                    merged_df[temp_type] = merged_df[col1].combine_first(merged_df[col2])
                    # 删除原始列
                    merged_df = merged_df.drop(columns=[col1, col2], errors='ignore')
                elif cols_1:
                    # 只有文件1有该类型数据
                    merged_df = merged_df.rename(columns={cols_1[0]: temp_type})
                elif cols_2:
                    # 只有文件2有该类型数据
                    merged_df = merged_df.rename(columns={cols_2[0]: temp_type})
            
            # 确保有Time列
            if 'Time' not in merged_df.columns and len(merged_df) > 0:
                print("警告: 未找到Time列，使用索引作为时间")
                merged_df['Time'] = range(len(merged_df))
            
            # 按时间排序
            if 'Time' in merged_df.columns:
                merged_df = merged_df.sort_values('Time')
            
            # 移除完全为NaN的列
            merged_df = merged_df.dropna(axis=1, how='all')
            
            # 填充NaN值
            for col in merged_df.columns:
                if col != 'Time':
                    merged_df[col] = merged_df[col].interpolate(method='linear', limit_direction='both')
            
            print(f"最终数据形状: {merged_df.shape}")
            print(f"最终列名: {list(merged_df.columns)}")
            
            self.progress_updated.emit(90)
            
            # 准备列信息
            column_info = {}
            for col in ['Temperature_S1', 'Temperature_SV', 'Temperature_PV']:
                if col in merged_df.columns:
                    data = merged_df[col].dropna()
                    if len(data) > 0:
                        column_info[col] = {
                            'count': len(data),
                            'mean': float(data.mean()) if not pd.isna(data.mean()) else 0,
                            'std': float(data.std()) if not pd.isna(data.std()) else 0,
                            'min': float(data.min()) if not pd.isna(data.min()) else 0,
                            'max': float(data.max()) if not pd.isna(data.max()) else 0
                        }
                        print(f"{col}: {len(data)}个数据点, 均值: {data.mean():.2f}")
            
            self.progress_updated.emit(100)
            
            # 发送数据和列信息
            self.data_loaded.emit(merged_df, column_info)
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            error_msg = f"处理数据时出错: {str(e)}"
            print(error_msg)
            print(f"详细错误信息:\n{error_details}")
            self.error_occurred.emit(f"{error_msg}\n\n详细信息:\n{error_details}")

class InteractivePlotCanvas(FigureCanvas):
    """交互式matplotlib画布"""
    def __init__(self, parent=None, width=10, height=6, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        super().__init__(self.fig)
        self.setParent(parent)
        
        # 创建子图
        self.ax = self.fig.add_subplot(111)
        self.fig.tight_layout()
        
        # 初始化数据存储
        self.time_data = None
        self.plot_data = None
        self.plot_lines = {}  # 存储每条曲线的Line2D对象
        self.markers = {}  # 存储标记点
        self.marker_texts = {}  # 存储标记文本
        
        # 启用鼠标交互
        self.mpl_connect('motion_notify_event', self.on_mouse_move)
        self.mpl_connect('button_press_event', self.on_mouse_click)
        self.mpl_connect('scroll_event', self.on_mouse_scroll)
        
        # 数据点信息存储
        self.data_points = []  # 存储所有数据点信息
        
        # 当前选中的点
        self.selected_point = None
        self.info_text = None
        
    def set_data(self, merged_df):
        """设置数据并绘制初始图表"""
        # 清空画布
        self.ax.clear()
        
        # 获取时间数据
        if 'Time' in merged_df.columns and len(merged_df['Time'].dropna()) > 0:
            self.time_data = merged_df['Time'].dropna().values
        else:
            self.time_data = np.arange(len(merged_df))
        
        # 存储所有数据点信息
        self.data_points = []
        
        # 绘制三条曲线
        styles = {
            'Temperature_S1': {'color': 'r-', 'label': 'Temperature_S1 (外部热电偶)', 'marker': 'o', 'zorder': 5},
            'Temperature_SV': {'color': 'b-', 'label': 'Temperature_SV (设定值)', 'marker': 's', 'zorder': 4},
            'Temperature_PV': {'color': 'g-', 'label': 'Temperature_PV (温箱监控)', 'marker': '^', 'zorder': 3}
        }
        
        # 清空存储
        self.plot_lines = {}
        self.markers = {}
        self.marker_texts = {}
        
        for col_type, style in styles.items():
            if col_type in merged_df.columns:
                data = pd.to_numeric(merged_df[col_type], errors='coerce').dropna().values
                
                if len(data) > 0:
                    min_len = min(len(self.time_data), len(data))
                    if min_len > 0:
                        time_plot = self.time_data[:min_len]
                        data_plot = data[:min_len]
                        
                        # 绘制曲线
                        line, = self.ax.plot(time_plot, data_plot, 
                                           style['color'], 
                                           label=style['label'],
                                           linewidth=1.5, 
                                           alpha=0.8,
                                           marker='',  # 暂时不显示标记点
                                           markersize=6,
                                           markevery=1,
                                           zorder=style['zorder'])
                        
                        # 存储Line2D对象
                        self.plot_lines[col_type] = line
                        
                        # 收集所有数据点
                        for t, v in zip(time_plot, data_plot):
                            self.data_points.append({
                                'time': t,
                                'value': v,
                                'type': col_type,
                                'label': style['label']
                            })
        
        # 按时间排序数据点
        self.data_points.sort(key=lambda x: x['time'])
        
        # 设置图表属性
        self.setup_axes()
        
        # 创建信息文本框
        self.create_info_box()
        
        # 刷新画布
        self.draw()
    
    def setup_axes(self):
        """设置坐标轴属性"""
        # 智能设置X轴刻度
        if len(self.time_data) > 0:
            time_range = self.time_data[-1] - self.time_data[0]
            if time_range > 3600 * 24:  # 大于1天
                interval = 3600 * 6  # 6小时

                self.ax.set_xlabel('时间 (小时)', fontsize=12)
                ticks = np.arange(self.time_data[0], self.time_data[-1] + interval, interval)

                self.ax.set_xticks(ticks[ticks <= self.time_data[-1]])

                self.ax.set_xticklabels([f"{tick/3600:.1f}" for tick in ticks if tick <= self.time_data[-1]], rotation=45)

            elif time_range > 3600:  # 大于1小时

                interval = 3600  # 1小时

                self.ax.set_xlabel('时间 (小时)', fontsize=12)

                ticks = np.arange(self.time_data[0], self.time_data[-1] + interval, interval)

                self.ax.set_xticks(ticks[ticks <= self.time_data[-1]])

                self.ax.set_xticklabels([f"{tick/3600:.1f}" for tick in ticks if tick <= self.time_data[-1]], rotation=45)

            elif time_range > 60:  # 大于1分钟

                interval = 300  # 5分钟

                self.ax.set_xlabel('时间 (分钟)', fontsize=12)

                ticks = np.arange(self.time_data[0], self.time_data[-1] + interval, interval)

                self.ax.set_xticks(ticks[ticks <= self.time_data[-1]])

                self.ax.set_xticklabels([f"{tick/60:.0f}" for tick in ticks if tick <= self.time_data[-1]], rotation=45)

            else:  # 小于1分钟

                interval = 10  # 10秒

                self.ax.set_xlabel('时间 (秒)', fontsize=12)

                ticks = np.arange(self.time_data[0], self.time_data[-1] + interval, interval)

                self.ax.set_xticks(ticks[ticks <= self.time_data[-1]])

                self.ax.set_xticklabels([f"{tick:.0f}" for tick in ticks if tick <= self.time_data[-1]], rotation=45)
        
        self.ax.set_ylabel('温度 (°C)', fontsize=12)
        self.ax.set_title('温度数据合成曲线 (可交互)', fontsize=14, fontweight='bold')
        self.ax.legend(fontsize=10, loc='upper right')
        self.ax.grid(True, alpha=0.3, linestyle='--')
        
        # 自动调整坐标轴
        self.ax.relim()
        self.ax.autoscale_view()
    
    def create_info_box(self):
        """创建信息显示框"""
        if self.info_text is not None and self.info_text in self.ax.texts:
            self.info_text.remove()
        
        # 在图表右上角创建文本框
        self.info_text = self.ax.text(0.98, 0.98, '', 
                                     transform=self.ax.transAxes,
                                     fontsize=10,
                                     verticalalignment='top',
                                     horizontalalignment='right',
                                     bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.9),
                                     zorder=10)
        
        # 隐藏初始状态
        self.info_text.set_visible(False)
    
    def show_nearest_point_info(self, x, y):
        """显示最近数据点的信息"""
        if not self.data_points:
            return
            
        # 找到最近的数据点
        nearest_point = None
        min_distance = float('inf')
        
        for point in self.data_points:
            # 计算距离（欧几里得距离）
            distance = np.sqrt((point['time'] - x)**2 + (point['value'] - y)**2)
            if distance < min_distance:
                min_distance = distance
                nearest_point = point
        
        if nearest_point and min_distance < 0.05 * (self.ax.get_xlim()[1] - self.ax.get_xlim()[0]):
            # 显示最近点的信息
            info = f"时间: {nearest_point['time']:.1f}s\n"
            info += f"温度: {nearest_point['value']:.2f}°C\n"
            info += f"类型: {nearest_point['label']}\n"
            info += f"距离: {min_distance:.4f}"
            
            # 更新信息框
            self.info_text.set_text(info)
            self.info_text.set_visible(True)
            
            # 移除之前的标记点
            for key in self.markers:
                if self.markers[key] in self.ax.collections:
                    self.markers[key].remove()
                if self.marker_texts[key] in self.ax.texts:
                    self.marker_texts[key].remove()
            
            # 标记最近的点
            marker = self.ax.scatter([nearest_point['time']], [nearest_point['value']], 
                                   color='red', s=100, zorder=10,
                                   edgecolors='black', linewidth=2)
            
            # 添加标记文本
            marker_text = self.ax.text(nearest_point['time'], nearest_point['value'] + 1,
                                     f"{nearest_point['value']:.2f}°C",
                                     fontsize=9, fontweight='bold',
                                     bbox=dict(boxstyle='round', facecolor='yellow', alpha=0.8))
            
            # 存储标记
            self.markers['current'] = marker
            self.marker_texts['current'] = marker_text
            
            # 刷新画布
            self.draw()
    
    def clear_current_markers(self):
        """清除当前标记点"""
        for key in self.markers:
            if self.markers[key] in self.ax.collections:
                self.markers[key].remove()
        self.markers.clear()
        
        for key in self.marker_texts:
            if self.marker_texts[key] in self.ax.texts:
                self.marker_texts[key].remove()
        self.marker_texts.clear()
        
        # 隐藏信息框
        if self.info_text:
            self.info_text.set_visible(False)
        
        # 刷新画布
        self.draw()
    
    def on_mouse_move(self, event):
        """鼠标移动事件处理"""
        if event.inaxes == self.ax:
            x, y = event.xdata, event.ydata
            
            # 更新鼠标位置显示
            if hasattr(self, 'mouse_label'):
                self.mouse_label.setText(f"鼠标位置: ({x:.2f}, {y:.2f})")
            
            # 显示最近点的信息
            self.show_nearest_point_info(x, y)
        else:
            # 鼠标不在图表内时隐藏信息
            self.clear_current_markers()
    
    def on_mouse_click(self, event):
        """鼠标点击事件处理"""
        if event.inaxes == self.ax and event.button == 1:  # 左键点击
            x, y = event.xdata, event.ydata
            self.add_marker_at_position(x, y)
    
    def on_mouse_scroll(self, event):
        """鼠标滚轮事件处理"""
        if event.inaxes == self.ax:
            # 滚轮缩放功能
            cur_xlim = self.ax.get_xlim()
            cur_ylim = self.ax.get_ylim()
            
            xdata = event.xdata
            ydata = event.ydata
            
            scale_factor = 1.1 if event.button == 'up' else 0.9
            
            new_width = (cur_xlim[1] - cur_xlim[0]) * scale_factor
            new_height = (cur_ylim[1] - cur_ylim[0]) * scale_factor
            
            relx = (cur_xlim[1] - xdata) / (cur_xlim[1] - cur_xlim[0])
            rely = (cur_ylim[1] - ydata) / (cur_ylim[1] - cur_ylim[0])
            
            self.ax.set_xlim([xdata - new_width * (1 - relx), xdata + new_width * relx])
            self.ax.set_ylim([ydata - new_height * (1 - rely), ydata + new_height * rely])
            
            self.draw()
    
    def add_marker_at_position(self, x, y):
        """在指定位置添加标记点"""
        # 寻找最近的数据点
        nearest_point = None
        min_distance = float('inf')
        
        for point in self.data_points:
            distance = np.sqrt((point['time'] - x)**2 + (point['value'] - y)**2)
            if distance < min_distance:
                min_distance = distance
                nearest_point = point
        
        if nearest_point:
            # 创建标记
            marker_id = f"marker_{len(self.markers)}"
            
            # 移除旧的标记（如果有）
            if marker_id in self.markers and self.markers[marker_id] in self.ax.collections:
                self.markers[marker_id].remove()
            if marker_id in self.marker_texts and self.marker_texts[marker_id] in self.ax.texts:
                self.marker_texts[marker_id].remove()
            
            # 绘制标记点
            marker = self.ax.scatter([nearest_point['time']], [nearest_point['value']], 
                                   color='orange', s=80, zorder=9,
                                   edgecolors='darkred', linewidth=1.5,
                                   alpha=0.8)
            
            # 添加文本标签
            marker_text = self.ax.text(nearest_point['time'], nearest_point['value'] + 2,
                                     f"{nearest_point['label']}\n{nearest_point['value']:.2f}°C",
                                     fontsize=8, fontweight='bold',
                                     bbox=dict(boxstyle='round', facecolor='lightblue', alpha=0.8))
            
            # 存储标记
            self.markers[marker_id] = marker
            self.marker_texts[marker_id] = marker_text
            
            # 刷新画布
            self.draw()

class DataPlotterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.merged_data = None
        self.column_info = {}
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle('智能温度曲线合成工具 - 交互式')
        self.setGeometry(100, 100, 1600, 1000)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QHBoxLayout(central_widget)
        
        # 左侧控制面板
        left_panel = self.create_control_panel()
        main_layout.addWidget(left_panel, 1)
        
        # 右侧图表和交互面板
        right_panel = self.create_interactive_plot_panel()
        main_layout.addWidget(right_panel, 3)
        
    def create_control_panel(self):
        """创建左侧控制面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        
        # 标题
        title_label = QLabel('智能温度曲线合成工具\n(交互式)')
        title_label.setFont(QFont('Arial', 16, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        
        layout.addSpacing(20)
        
        # 文件上传组
        file_group = QGroupBox("文件上传")
        file_layout = QVBoxLayout(file_group)
        
        # 第一个文件
        self.file1_btn = QPushButton("选择第一个Excel文件\n(包含Temperature_S1)")
        self.file1_btn.clicked.connect(lambda: self.select_file(1))
        self.file1_label = QLabel("未选择文件")
        self.file1_label.setWordWrap(True)
        self.file1_label.setMaximumWidth(200)
        file_layout.addWidget(self.file1_btn)
        file_layout.addWidget(self.file1_label)
        
        file_layout.addSpacing(10)
        
        # 第二个文件
        self.file2_btn = QPushButton("选择第二个Excel文件\n(包含Temperature_SV/PV)")
        self.file2_btn.clicked.connect(lambda: self.select_file(2))
        self.file2_label = QLabel("未选择文件")
        self.file2_label.setWordWrap(True)
        self.file2_label.setMaximumWidth(200)
        file_layout.addWidget(self.file2_btn)
        file_layout.addWidget(self.file2_label)
        
        layout.addWidget(file_group)
        
        layout.addSpacing(20)
        
        # 绘图设置组
        plot_group = QGroupBox("交互式绘图设置")
        plot_layout = QVBoxLayout(plot_group)
        
        # 标记点设置
        marker_group = QGroupBox("标记点设置")
        marker_layout = QVBoxLayout(marker_group)
        
        marker_layout.addWidget(QLabel("标记点大小:"))
        marker_size_spinbox = QSpinBox()
        marker_size_spinbox.setRange(1, 200)
        marker_size_spinbox.setValue(80)
        marker_size_spinbox.setToolTip("标记点的大小")
        marker_layout.addWidget(marker_size_spinbox)
        
        marker_layout.addWidget(QLabel("标记点颜色:"))
        marker_color_combo = QComboBox()
        marker_color_combo.addItems(["红色", "蓝色", "绿色", "橙色", "紫色"])
        marker_layout.addWidget(marker_color_combo)
        
        plot_layout.addWidget(marker_group)
        
        # 缩放控制
        zoom_group = QGroupBox("缩放控制")
        zoom_layout = QVBoxLayout(zoom_group)
        
        zoom_layout.addWidget(QLabel("缩放因子:"))
        zoom_factor_spinbox = QDoubleSpinBox()
        zoom_factor_spinbox.setRange(0.1, 5.0)
        zoom_factor_spinbox.setValue(1.1)
        zoom_factor_spinbox.setSingleStep(0.1)
        zoom_layout.addWidget(zoom_factor_spinbox)
        
        plot_layout.addWidget(zoom_group)
        
        layout.addWidget(plot_group)
        
        layout.addSpacing(20)
        
        # 操作按钮组
        button_group = QGroupBox("操作")
        button_layout = QVBoxLayout(button_group)
        
        # 生成图表按钮
        self.plot_btn = QPushButton("生成交互式曲线")
        self.plot_btn.clicked.connect(self.plot_data)
        self.plot_btn.setEnabled(False)
        button_layout.addWidget(self.plot_btn)
        
        # 清除图表按钮

        self.clear_btn = QPushButton("清除图表")

        self.clear_btn.clicked.connect(self.clear_plot)

        button_layout.addWidget(self.clear_btn)
        
        # 清除标记按钮

        self.clear_markers_btn = QPushButton("清除标记点")

        self.clear_markers_btn.clicked.connect(self.clear_markers)

        button_layout.addWidget(self.clear_markers_btn)
        
        # 保存图片按钮

        self.save_btn = QPushButton("保存图片")

        self.save_btn.clicked.connect(self.save_plot)

        self.save_btn.setEnabled(False)

        button_layout.addWidget(self.save_btn)
        
        # 查看数据按钮

        self.view_data_btn = QPushButton("查看数据")

        self.view_data_btn.clicked.connect(self.view_data)

        self.view_data_btn.setEnabled(False)

        button_layout.addWidget(self.view_data_btn)
        
        # 导出数据按钮

        self.export_data_btn = QPushButton("导出数据")

        self.export_data_btn.clicked.connect(self.export_data)

        self.export_data_btn.setEnabled(False)

        button_layout.addWidget(self.export_data_btn)
        
        layout.addWidget(button_group)
        
        layout.addSpacing(20)
        
        # 实时鼠标位置显示

        self.mouse_label = QLabel("鼠标位置: (等待中...)")

        self.mouse_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(self.mouse_label)
        
        # 进度条

        self.progress_bar = QProgressBar()

        self.progress_bar.setVisible(False)

        layout.addWidget(self.progress_bar)
        
        # 状态标签

        self.status_label = QLabel("请选择Excel文件开始")

        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.status_label.setWordWrap(True)

        layout.addWidget(self.status_label)
        
        # 信息显示框

        info_group = QGroupBox("实时信息")

        info_layout = QVBoxLayout(info_group)

        
        self.info_display = QTextEdit()

        self.info_display.setReadOnly(True)

        self.info_display.setMaximumHeight(150)

        info_layout.addWidget(self.info_display)
        
        layout.addWidget(info_group)
        
        layout.addStretch()
        
        return panel
    
    def create_interactive_plot_panel(self):
        """创建右侧交互式图表面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        
        # 创建交互式matplotlib画布

        self.canvas = InteractivePlotCanvas(self, width=12, height=9)

        
        # 将画布的mouse_label连接到我们的标签

        self.canvas.mouse_label = self.mouse_label
        
        self.toolbar = NavigationToolbar(self.canvas, self)
        
        layout.addWidget(self.toolbar)

        layout.addWidget(self.canvas)
        
        # 添加交互控制按钮

        control_layout = QHBoxLayout()

        
        # 放大按钮

        zoom_in_btn = QPushButton("放大")

        zoom_in_btn.clicked.connect(self.zoom_in)

        control_layout.addWidget(zoom_in_btn)

        
        # 缩小按钮

        zoom_out_btn = QPushButton("缩小")

        zoom_out_btn.clicked.connect(self.zoom_out)

        control_layout.addWidget(zoom_out_btn)

        
        # 重置视图按钮

        reset_view_btn = QPushButton("重置视图")

        reset_view_btn.clicked.connect(self.reset_view)

        control_layout.addWidget(reset_view_btn)

        
        # 添加标记按钮

        add_marker_btn = QPushButton("添加标记")

        add_marker_btn.clicked.connect(self.add_marker)

        control_layout.addWidget(add_marker_btn)

        
        layout.addLayout(control_layout)
        
        return panel
    
    def select_file(self, file_num):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, f"选择第{file_num}个Excel文件", "", 
            "Excel files (*.xlsx *.xls);;All files (*)"
        )
        
        if file_path:
            if file_num == 1:
                self.file1_path = file_path
                self.file1_label.setText(os.path.basename(file_path))

            else:

                self.file2_path = file_path

                self.file2_label.setText(os.path.basename(file_path))

            
            # 如果两个文件都已选择，启用绘图按钮

            if hasattr(self, 'file1_path') and hasattr(self, 'file2_path'):

                self.plot_btn.setEnabled(True)

                self.status_label.setText("文件已选择完成，点击生成合成曲线")
    
    def plot_data(self):
        """绘制交互式曲线"""
        if not hasattr(self, 'file1_path') or not hasattr(self, 'file2_path'):
            QMessageBox.warning(self, "警告", "请先选择两个Excel文件！")
            return
        
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.plot_btn.setEnabled(False)
        
        # 显示状态信息

        self.status_label.setText("正在处理数据...")
        
        self.processor = DataProcessor(
            self.file1_path, 
            self.file2_path,
            auto_align=True
        )
        self.processor.progress_updated.connect(self.update_progress)
        self.processor.data_loaded.connect(self.on_data_loaded)
        self.processor.error_occurred.connect(self.on_error)
        self.processor.start()
    
    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)
        if value < 100:
            self.status_label.setText(f"正在处理数据... {value}%")
    
    def on_data_loaded(self, merged_df, column_info):
        """数据加载完成回调"""
        try:
            self.progress_bar.setVisible(False)
            self.merged_data = merged_df
            self.column_info = column_info
            
            # 设置交互式画布的数据
            self.canvas.set_data(merged_df)
            
            # 启用按钮
            self.save_btn.setEnabled(True)
            self.view_data_btn.setEnabled(True)
            self.export_data_btn.setEnabled(True)
            
            # 更新状态信息
            detection_text = "检测到的列:\n"
            for col_type, info in column_info.items():
                detection_text += f"{col_type}: {info['count']}点\n"
                detection_text += f"  均值: {info['mean']:.2f}°C\n"
                detection_text += f"  范围: [{info['min']:.2f}, {info['max']:.2f}]°C\n"
            
            time_data = self.canvas.time_data if self.canvas.time_data is not None else []
            if len(time_data) > 0:
                time_range = f"{time_data[0]:.0f} - {time_data[-1]:.0f}s"
                detection_text += f"\n时间范围: {time_range}"
                detection_text += f"\n数据点数: {len(time_data)}"
            
            self.status_label.setText(f"成功绘制交互式曲线")
            
            # 显示操作提示
            self.info_display.setPlainText(
                "操作说明:\n"
                "1. 鼠标悬停在图表上查看最近数据点信息\n"
                "2. 左键点击添加永久标记点\n"
                "3. 使用滚轮或工具栏按钮缩放\n"
                "4. 点击'清除标记点'按钮移除所有标记\n"
                "5. 点击'重置视图'按钮恢复原始视图"
            )
            
            self.plot_btn.setEnabled(True)
            
        except Exception as e:
            self.on_error(f"绘制图表时出错: {str(e)}")
    
    def on_error(self, error_msg):
        """错误处理"""
        self.progress_bar.setVisible(False)
        self.plot_btn.setEnabled(True)
        QMessageBox.critical(self, "错误", error_msg)
        self.status_label.setText("处理失败，请检查文件格式")
    
    def clear_plot(self):
        """清除图表"""
        if hasattr(self, 'canvas'):
            self.canvas.set_data(self.merged_data) if self.merged_data is not None else self.canvas.ax.clear()
            self.canvas.draw()
        self.save_btn.setEnabled(False)
        self.status_label.setText("图表已清除")
    
    def clear_markers(self):
        """清除所有标记点"""
        if hasattr(self, 'canvas'):
            self.canvas.clear_current_markers()
    
    def zoom_in(self):
        """放大视图"""
        if hasattr(self, 'canvas'):
            cur_xlim = self.canvas.ax.get_xlim()
            cur_ylim = self.canvas.ax.get_ylim()
            
            new_width = (cur_xlim[1] - cur_xlim[0]) * 0.9
            new_height = (cur_ylim[1] - cur_ylim[0]) * 0.9
            
            center_x = (cur_xlim[0] + cur_xlim[1]) / 2
            center_y = (cur_ylim[0] + cur_ylim[1]) / 2
            
            self.canvas.ax.set_xlim([center_x - new_width/2, center_x + new_width/2])
            self.canvas.ax.set_ylim([center_y - new_height/2, center_y + new_height/2])
            
            self.canvas.draw()
    
    def zoom_out(self):
        """缩小视图"""
        if hasattr(self, 'canvas'):
            cur_xlim = self.canvas.ax.get_xlim()
            cur_ylim = self.canvas.ax.get_ylim()
            
            new_width = (cur_xlim[1] - cur_xlim[0]) * 1.1

            new_height = (cur_ylim[1] - cur_ylim[0]) * 1.1

            
            center_x = (cur_xlim[0] + cur_xlim[1]) / 2

            center_y = (cur_ylim[0] + cur_ylim[1]) / 2

            
            self.canvas.ax.set_xlim([center_x - new_width/2, center_x + new_width/2])

            self.canvas.ax.set_ylim([center_y - new_height/2, center_y + new_height/2])

            
            self.canvas.draw()
    
    def reset_view(self):
        """重置视图"""
        if hasattr(self, 'canvas') and self.merged_data is not None:
            self.canvas.set_data(self.merged_data)
    
    def add_marker(self):
        """添加标记点"""
        if hasattr(self, 'canvas'):
            # 在视图中心添加标记点
            xlim = self.canvas.ax.get_xlim()
            ylim = self.canvas.ax.get_ylim()
            
            center_x = (xlim[0] + xlim[1]) / 2
            center_y = (ylim[0] + ylim[1]) / 2
            
            self.canvas.add_marker_at_position(center_x, center_y)
    
    def save_plot(self):
        """保存图表"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存图表", "interactive_temperature_plot.png",
            "PNG files (*.png);;JPG files (*.jpg);;PDF files (*.pdf);;All files (*)"
        )
        
        if file_path:
            try:
                self.canvas.fig.savefig(file_path, dpi=300, bbox_inches='tight')
                QMessageBox.information(self, "成功", f"图表已保存到: {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存失败: {str(e)}")
    
    def view_data(self):
        """查看数据"""
        if self.merged_data is not None and len(self.merged_data) > 0:
            dialog = QDialog(self)
            dialog.setWindowTitle("数据预览")
            dialog.setGeometry(200, 200, 1000, 700)
            
            layout = QVBoxLayout(dialog)
            
            # 数据信息汇总
            summary_label = QLabel("数据信息汇总")
            summary_label.setFont(QFont('Arial', 12, QFont.Weight.Bold))
            layout.addWidget(summary_label)
            
            info_text = QTextEdit()
            info_text.setReadOnly(True)
            
            info_content = f"数据形状: {self.merged_data.shape}\n"
            if 'Time' in self.merged_data.columns:
                info_content += f"时间范围: {self.merged_data['Time'].min():.1f}s - {self.merged_data['Time'].max():.1f}s\n"
            info_content += f"数据点数: {len(self.merged_data)}\n"
            info_content += f"列名: {list(self.merged_data.columns)}\n\n"
            
            # 显示所有数据（因为数据量不大）
            info_content += "全部数据:\n"
            info_content += self.merged_data.to_string()
            
            info_text.setText(info_content)
            layout.addWidget(info_text)
            
            dialog.exec()
    
    def export_data(self):
        """导出数据"""
        if self.merged_data is not None:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "导出数据", "temperature_data.csv",
                "CSV files (*.csv);;Excel files (*.xlsx);;All files (*)"
            )
            
            if file_path:
                try:
                    if file_path.endswith('.csv'):
                        self.merged_data.to_csv(file_path, index=False)
                    else:
                        self.merged_data.to_excel(file_path, index=False)
                    
                    QMessageBox.information(self, "成功", f"数据已导出到: {file_path}")
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")

def main():
    import matplotlib
    matplotlib.use('QtAgg')
    
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    window = DataPlotterApp()
    window.show()
    
    sys.exit(app.exec())

if __name__ == '__main__':
    main()