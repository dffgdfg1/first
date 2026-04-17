#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
电压电流曲线分析脚本
功能：自动分析HTML格式的电压电流数据，提取各电压段的电流范围，
      使用Selenium截取图表，生成汇总报告
"""

import os
import re
import sys
import json
import time
import logging
import numpy as np
import pandas as pd
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.webdriver import WebDriver as EdgeWebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image
import warnings
warnings.filterwarnings('ignore')

# ============ 配置参数 ============
VOLTAGE_SEGMENTS = [6.5, 9, 14, 16, 18]
VOLTAGE_TOLERANCE = 1.0       # 电压匹配容差
MIN_STABLE_POINTS = 30        # 最小稳定数据点数
OUTLIER_STD_FACTOR = 2.5      # 异常值标准差倍数
STABLE_WINDOW = 50            # 稳定性检测滑动窗口
STABLE_CV_THRESHOLD = 0.15    # 变异系数阈值（判断稳定）

# 日志配置
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
logger = logging.getLogger(__name__)

class HTMLDataExtractor:
    """从HTML文件中提取ECharts数据"""

    @staticmethod
    def extract(html_file: str) -> Tuple[List, List, str]:
        """
        提取电压和电流数据

        Returns:
            (voltage_data, current_data, current_unit)
        """
        with open(html_file, 'r', encoding='utf-8') as f:
            content = f.read()

        # 检测电流单位
        current_unit = 'A'
        if '"电流(uA)"' in content or '"电流(μA)"' in content:
            current_unit = 'uA'

        voltage_data = []
        current_data = []

        # 找到电压series的data数组
        v_data_text = HTMLDataExtractor._find_series_data(content, '"name": "电压"')
        if v_data_text:
            voltage_data = HTMLDataExtractor._parse_data_array(v_data_text)

        # 找到电流series的data数组
        i_data_text = HTMLDataExtractor._find_series_data(content, '"name": "电流"')
        if i_data_text:
            current_data = HTMLDataExtractor._parse_data_array(i_data_text)

        logger.info(f"提取 {len(voltage_data)} 个电压点, {len(current_data)} 个电流点, 单位: {current_unit}")
        return voltage_data, current_data, current_unit

    @staticmethod
    def _find_series_data(content: str, name_marker: str) -> Optional[str]:
        """通过name标记找到对应series的data数组文本，使用括号计数"""
        pos = content.find(name_marker)
        if pos == -1:
            # 尝试不带空格的格式
            name_marker_no_space = name_marker.replace(': ', ':')
            pos = content.find(name_marker_no_space)
            if pos == -1:
                return None

        # 从name位置向后找 "data": [
        data_key = '"data":'
        data_pos = content.find(data_key, pos)
        if data_pos == -1:
            return None

        # 找到 [ 的位置
        bracket_start = content.find('[', data_pos + len(data_key))
        if bracket_start == -1:
            return None

        # 用括号计数找到匹配的 ]
        depth = 0
        for i in range(bracket_start, len(content)):
            if content[i] == '[':
                depth += 1
            elif content[i] == ']':
                depth -= 1
                if depth == 0:
                    return content[bracket_start + 1:i]

        return None

    @staticmethod
    def _parse_data_array(data_text: str) -> List[Tuple[int, float]]:
        """解析 [[timestamp, value], ...] 格式的数据"""
        points = re.findall(r'\[\s*(\d+)\s*,\s*([-\d.eE+]+)\s*\]', data_text)
        return [(int(ts), float(val)) for ts, val in points]


class VoltageCurrentAnalyzer:
    """电压电流数据分析器"""

    def __init__(self, html_file: str, outlier_factor: float = None):
        self.html_file = html_file
        self.voltage_points = []  # [(timestamp, voltage), ...]
        self.current_points = []  # [(timestamp, current), ...]
        self.current_unit = 'A'
        # 使用传入的异常值倍数，如果没有传入则使用全局默认值
        self.outlier_factor = outlier_factor if outlier_factor is not None else OUTLIER_STD_FACTOR

    def analyze(self) -> Dict[float, Tuple[float, float]]:
        """执行分析，返回 {电压段: (最小电流, 最大电流)}"""
        try:
            self.voltage_points, self.current_points, self.current_unit = \
                HTMLDataExtractor.extract(self.html_file)
        except Exception as e:
            logger.error(f"数据提取失败: {e}")
            return {}

        if not self.voltage_points or not self.current_points:
            logger.warning(f"数据为空: {self.html_file}")
            return {}

        self._voltages = np.array([v for _, v in self.voltage_points])
        self._currents = np.array([v for _, v in self.current_points])

        # 检查数据长度是否一致
        if len(self._voltages) != len(self._currents):
            logger.warning(f"电压点数({len(self._voltages)})与电流点数({len(self._currents)})不一致，将只使用较短的长度")
            min_len = min(len(self._voltages), len(self._currents))
            self._voltages = self._voltages[:min_len]
            self._currents = self._currents[:min_len]

        results = {}
        for target_v in VOLTAGE_SEGMENTS:
            result = self._analyze_voltage_segment(self._voltages, self._currents, target_v)
            if result:
                results[target_v] = result
                logger.info(f"  {target_v}V: {result[0]:.4f}-{result[1]:.4f} (平均:{result[2]:.4f}) {self.current_unit}")

        return results

    def analyze_static(self) -> Optional[float]:
        """
        分析静态电流：计算整个数据的平均电流，保留两位小数
        返回 平均电流值 或 None
        """
        if not hasattr(self, '_currents'):
            try:
                self.voltage_points, self.current_points, self.current_unit = \
                    HTMLDataExtractor.extract(self.html_file)
                self._voltages = np.array([v for _, v in self.voltage_points])
                self._currents = np.array([v for _, v in self.current_points])

                # 检查数据长度是否一致
                if len(self._voltages) != len(self._currents):
                    logger.warning(f"电压点数({len(self._voltages)})与电流点数({len(self._currents)})不一致，将只使用较短的长度")
                    min_len = min(len(self._voltages), len(self._currents))
                    self._voltages = self._voltages[:min_len]
                    self._currents = self._currents[:min_len]
            except Exception:
                return None

        if len(self._currents) == 0:
            return None

        # 计算平均值并保留两位小数
        avg_current = float(np.mean(self._currents))
        result = round(avg_current, 2)
        logger.info(f"  静态电流: {result:.2f} {self.current_unit}")
        return result

    def analyze_sleep(self) -> Optional[Tuple[float, float, float]]:
        """
        分析休眠模式：提取14V段的休眠电流
        两种检测方式：
        1. 14V段持续约1分钟（400-800点）
        2. 数据最后一段14V的uA电流（新增）
        返回 (min_current, max_current, avg_current) 或 None
        """
        if not hasattr(self, '_voltages'):
            try:
                self.voltage_points, self.current_points, self.current_unit = \
                    HTMLDataExtractor.extract(self.html_file)
                self._voltages = np.array([v for _, v in self.voltage_points])
                self._currents = np.array([v for _, v in self.current_points])
            except Exception:
                return None

        voltages = self._voltages
        currents = self._currents

        # 检查数据长度是否一致
        max_valid_index = len(currents) - 1
        if len(voltages) != len(currents):
            logger.warning(f"电压点数({len(voltages)})与电流点数({len(currents)})不一致，将只使用前{len(currents)}个电压点")
            voltages = voltages[:len(currents)]

        # 找到14V段
        mask = np.abs(voltages - 14) <= VOLTAGE_TOLERANCE
        indices = np.where(mask)[0]
        if len(indices) == 0:
            return None

        groups = self._split_continuous(indices.tolist())

        # 检测所有包含密集低电流的14V段（不限制位置和点数范围）
        sleep_groups = []
        for group_idx, g in enumerate(groups):
            # 要求至少有50个点才认为是有效段
            if len(g) < 50:
                continue
            
            # 检查电流是否真的是休眠电流（uA单位时需要验证）
            if self.current_unit == 'uA':
                group_currents = currents[g]
                
                # 先过滤掉明显的高电流点（>10000uA），保留低电流部分
                low_current_mask = group_currents < 10000
                low_current_count = np.sum(low_current_mask)
                
                # 如果有足够多的低电流点（至少50个），说明可能是休眠电流
                if low_current_count >= 50:
                    low_currents = group_currents[low_current_mask]
                    avg_current = np.mean(low_currents)
                    max_current = np.max(low_currents)
                    
                    # 低电流部分的平均值应该<1000uA才是真正的休眠电流
                    if avg_current < 1000 and max_current < 5000:
                        sleep_groups.append(g)
                        logger.info(f"  检测到休眠电流段{group_idx+1} (共{len(g)}个点, 低电流点:{low_current_count}个, 平均:{avg_current:.0f}uA)")
                    else:
                        logger.debug(f"  跳过高电流段{group_idx+1} (共{len(g)}个点, 平均:{avg_current:.0f}uA, 最大:{max_current:.0f}uA)")
                else:
                    logger.debug(f"  跳过高电流段{group_idx+1} (共{len(g)}个点, 低电流点仅{low_current_count}个)")
            else:
                # 非uA单位，直接添加
                sleep_groups.append(g)

        if not sleep_groups:
            return None

        # 合并所有休眠组的稳定电流（只保留低电流部分）
        all_stable = []
        for g in sleep_groups:
            seg_c = currents[g]
            
            # 对于uA单位，只保留低电流部分（<10000uA）
            if self.current_unit == 'uA':
                low_mask = seg_c < 10000
                seg_c = seg_c[low_mask]
            
            # 去掉前后15%的过渡数据
            trim = int(len(seg_c) * 0.15)
            if trim > 0 and len(seg_c) > trim * 2:
                stable = seg_c[trim:-trim]
            else:
                stable = seg_c
            all_stable.extend(stable.tolist())

        if not all_stable:
            return None

        arr = np.array(all_stable)
        # 去除异常值
        filtered = self._remove_outliers(arr)
        if len(filtered) == 0:
            return None

        result = (float(np.min(filtered)), float(np.max(filtered)), float(np.mean(filtered)))
        logger.info(f"  休眠14V: {result[0]:.6f}-{result[1]:.6f} (平均:{result[2]:.6f}) {self.current_unit}")
        return result

    def _analyze_voltage_segment(self, voltages, currents, target_v):
        """分析单个电压段，返回 (min_current, max_current, avg_current) 或 None"""
        # 检查数据长度是否一致
        if len(voltages) != len(currents):
            logger.warning(f"电压点数({len(voltages)})与电流点数({len(currents)})不一致，将只使用前{len(currents)}个电压点")
            voltages = voltages[:len(currents)]

        # 找到匹配电压的索引
        mask = np.abs(voltages - target_v) <= VOLTAGE_TOLERANCE
        indices = np.where(mask)[0]
        if len(indices) == 0:
            return None

        # 分割为连续组
        groups = self._split_continuous(indices.tolist())

        # 过滤有效组（长度足够且平均电流足够高）
        valid_groups = []
        for g in groups:
            if len(g) < MIN_STABLE_POINTS:
                continue
            grp_currents = currents[g]
            avg_current = np.mean(grp_currents)
            max_current = np.max(grp_currents)

            # 过滤掉静态/空闲组（平均电流太低）
            if max_current < 0.01:
                continue

            # 14V特殊处理：只取持续时间较长的组（至少300个点，约30秒）
            if target_v == 14.0:
                # 要求至少300个点，不设上限（因为14V段长度可能变化）
                if len(g) < 300:
                    continue

            valid_groups.append((g, avg_current, max_current))

        if not valid_groups:
            return None

        # 按平均电流排序，取最高的那个组（最后一个上升阶段通常电流最高）
        valid_groups.sort(key=lambda x: x[1])
        best_group = valid_groups[-1][0]

        seg_currents = currents[best_group]

        # 14V特殊处理：智能检测稳定阶段
        if target_v == 14.0:
            # 使用滑动窗口检测稳定阶段
            stable_start = self._find_stable_region(seg_currents)

            if stable_start is not None and stable_start < len(seg_currents):
                # 从稳定点开始取数据
                seg_currents = seg_currents[stable_start:]

                # 检查后段是否有明显下降（可能是进入休眠）
                if len(seg_currents) > 100:
                    # 计算后20%的平均电流
                    tail_start = int(len(seg_currents) * 0.8)
                    tail_avg = np.mean(seg_currents[tail_start:])
                    overall_avg = np.mean(seg_currents)

                    # 如果后段平均电流明显低于整体平均（低于70%），去掉后段
                    if tail_avg < overall_avg * 0.7:
                        seg_currents = seg_currents[:tail_start]
            else:
                # 如果找不到稳定区域，使用旧逻辑：去掉前25%
                front_trim = int(len(seg_currents) * 0.25)
                if front_trim > 0:
                    seg_currents = seg_currents[front_trim:]

            # 过滤掉极小值（<0.20A，过滤掉明显的低电流过渡阶段）
            seg_currents = seg_currents[seg_currents >= 0.20]

            if len(seg_currents) == 0:
                return None

            # 去除异常值
            filtered = self._remove_outliers(seg_currents)
            if len(filtered) == 0:
                return None

            return (float(np.min(filtered)), float(np.max(filtered)), float(np.mean(filtered)))

        # 其他电压段：原有逻辑
        # 找到活跃区域：电流 > 最大值的5%
        max_c = np.max(seg_currents)
        if max_c <= 0:
            return None

        active_threshold = max_c * 0.05
        active_mask = seg_currents > active_threshold

        # 找到第一个和最后一个活跃点
        active_indices = np.where(active_mask)[0]
        if len(active_indices) < MIN_STABLE_POINTS:
            return None

        first_active = active_indices[0]
        last_active = active_indices[-1]

        # 跳过启动阶段：从第一个活跃点开始，跳过前30%
        active_range = last_active - first_active
        skip_start = first_active + int(active_range * 0.3)
        stable_currents = seg_currents[skip_start:last_active + 1]

        if len(stable_currents) < 10:
            return None

        # 去除异常值（IQR方法）
        filtered = self._remove_outliers(stable_currents)
        if len(filtered) == 0:
            return None

        return (float(np.min(filtered)), float(np.max(filtered)), float(np.mean(filtered)))

    def _find_stable_region(self, currents: np.ndarray, window_size: int = 50) -> int:
        """
        检测电流稳定区域的起始位置

        算法：
        1. 使用滑动窗口计算电流变化率（标准差）和平均电流
        2. 找到变化率小且平均电流高的窗口，即为稳定工作区域
        3. 返回稳定区域的起始索引

        参数：
            currents: 电流数组
            window_size: 滑动窗口大小（默认50个点，约5秒）

        返回：
            稳定区域的起始索引，如果找不到则返回None
        """
        if len(currents) < window_size * 2:
            return None

        # 计算每个窗口的标准差（变化率）和平均电流
        stds = []
        avgs = []
        for i in range(len(currents) - window_size):
            window = currents[i:i + window_size]
            stds.append(np.std(window))
            avgs.append(np.mean(window))

        stds = np.array(stds)
        avgs = np.array(avgs)

        # 找到平均电流最高的区域（排除前20%的上升阶段）
        start_search = int(len(avgs) * 0.2)
        if start_search >= len(avgs):
            return None

        # 找到平均电流最高的窗口
        max_avg_idx = start_search + np.argmax(avgs[start_search:])

        # 在高电流区域附近找稳定点
        # 搜索范围：最高电流窗口前后20%的范围
        search_range = int(len(avgs) * 0.2)
        search_start = max(start_search, max_avg_idx - search_range)
        search_end = min(len(avgs), max_avg_idx + search_range)

        # 在这个范围内找标准差最小的窗口
        if search_start >= search_end:
            return max_avg_idx

        min_std_idx = search_start + np.argmin(stds[search_start:search_end])

        # 确保选中的窗口平均电流足够高（至少是最大值的80%）
        if avgs[min_std_idx] < avgs[max_avg_idx] * 0.8:
            # 如果太低，直接返回最高电流窗口的位置
            return max_avg_idx

        return min_std_idx

    @staticmethod
    def _split_continuous(indices: List[int]) -> List[List[int]]:
        """将索引列表按连续性分组"""
        if not indices:
            return []
        groups = [[indices[0]]]
        for i in range(1, len(indices)):
            if indices[i] - indices[i-1] <= 2:
                groups[-1].append(indices[i])
            else:
                groups.append([indices[i]])
        return groups

    def _remove_outliers(self, data: np.ndarray) -> np.ndarray:
        """使用IQR方法去除异常值"""
        if len(data) < 10:
            return data
        q1 = np.percentile(data, 25)
        q3 = np.percentile(data, 75)
        iqr = q3 - q1
        lower = q1 - self.outlier_factor * iqr
        upper = q3 + self.outlier_factor * iqr
        return data[(data >= lower) & (data <= upper)]


class ChartScreenshot:
    """使用Selenium截取ECharts图表"""

    def __init__(self):
        self.driver = None

    def init_browser(self):
        """初始化无头浏览器"""
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--window-size=2560,1440')  # 提高窗口分辨率到2K
        options.add_argument('--force-device-scale-factor=2')  # 2倍缩放提高清晰度
        options.add_argument('--high-dpi-support=1')  # 启用高DPI支持
        options.add_argument('--disable-gpu')  # 禁用GPU加速，提高稳定性
        options.add_argument('--disable-software-rasterizer')  # 禁用软件光栅化
        try:
            # 直接使用 EdgeWebDriver 类而不是 webdriver.Edge()，避免 PyInstaller 打包问题
            self.driver = EdgeWebDriver(options=options)
            # 设置页面加载超时为600秒（10分钟）
            self.driver.set_page_load_timeout(600)
            # 设置脚本执行超时为600秒
            self.driver.set_script_timeout(600)
            # 设置Selenium命令执行超时为600秒（解决大文件加载超时问题）
            self.driver.command_executor._timeout = 600
            logger.info("浏览器初始化成功")
        except Exception as e:
            logger.error(f"浏览器初始化失败: {e}")
            logger.info("请确保已安装Edge浏览器和对应版本的EdgeDriver")
            raise

    def capture(self, html_file: str, output_path: str) -> bool:
        """
        截取HTML图表并保存为PNG，确保显示完整数据范围（0电压0电流可见）

        Args:
            html_file: HTML文件路径
            output_path: 输出PNG路径
        Returns:
            bool: 是否成功
        """
        if not self.driver:
            self.init_browser()

        try:
            file_url = Path(html_file).resolve().as_uri()
            self.driver.get(file_url)

            # 等待ECharts渲染完成（增加超时时间以处理大文件）
            time.sleep(5)
            WebDriverWait(self.driver, 180).until(
                lambda d: d.execute_script(
                    "return document.querySelector('canvas') !== null"
                )
            )
            time.sleep(3)

            # 页面缩放到70%，确保能看到底部时间轴
            self.driver.execute_script("document.body.style.zoom='0.7'")
            time.sleep(0.5)

            # 重置dataZoom并调整图表尺寸确保完整显示
            self.driver.execute_script("""
                var chart = echarts.getInstanceByDom(document.getElementById('chart-container'));
                if (!chart) return;

                // 重置dataZoom到显示全部数据
                chart.dispatchAction({
                    type: 'dataZoom',
                    start: 0,
                    end: 100
                });

                // 隐藏dataZoom滑块和toolbox
                chart.setOption({
                    dataZoom: [{show: false}],
                    toolbox: {show: false}
                });

                // 调整图表高度，确保完整显示
                var dom = chart.getDom();
                dom.style.height = '1200px';
                chart.resize();
            """)

            # 等待重绘完成
            time.sleep(1)

            # 截取整个可见页面
            self.driver.save_screenshot(output_path)

            # 打开图片并智能裁剪底部空白
            img = Image.open(output_path)
            img_array = np.array(img.convert('RGB'))

            height, width = img_array.shape[0], img_array.shape[1]
            crop_bottom = height

            # 方案：检测非白色/非背景色像素
            # 1. 先检测背景色（取底部10行的众数颜色）
            bottom_sample = img_array[-10:, :, :]
            # 计算平均背景色
            bg_color = np.mean(bottom_sample.reshape(-1, 3), axis=0)
            logger.info(f"检测到背景色: RGB({bg_color[0]:.0f}, {bg_color[1]:.0f}, {bg_color[2]:.0f})")

            # 2. 从底部向上扫描，找到第一行包含明显非背景色像素的位置
            found = False
            for y in range(height - 1, -1, -1):
                row = img_array[y, :, :]

                # 计算每个像素与背景色的欧氏距离
                diff = np.sqrt(np.sum((row - bg_color) ** 2, axis=1))

                # 统计明显不同于背景色的像素数量（距离>30）
                non_bg_pixels = np.sum(diff > 30)

                # 如果这一行有超过宽度5%的非背景像素，认为是内容行
                if non_bg_pixels > width * 0.05:
                    crop_bottom = min(y + 80, height)  # 保留80px边距
                    logger.info(f"检测到内容行: y={y}, 非背景像素={non_bg_pixels}, 裁剪到: {crop_bottom}px")
                    found = True
                    break

            # 裁剪图片
            if found and crop_bottom < height:
                img = img.crop((0, 0, img.width, crop_bottom))
                logger.info(f"裁剪底部空白: {height}px -> {crop_bottom}px (减少{height - crop_bottom}px)")
            else:
                logger.info(f"未检测到需要裁剪的底部空白")

            # 直接拉伸到3:2宽高比
            # 目标宽度为15.5cm，使用300 DPI（打印质量）
            # 15.5cm * 300 DPI / 2.54 = 1831px
            target_ratio = 3.0 / 2.0  # 3:2比例
            target_width_px = 1831
            target_height_px = int(target_width_px / target_ratio)  # 按3:2计算高度 = 1221px

            current_width = img.width
            current_height = img.height
            current_ratio = current_width / current_height

            logger.info(f"原始尺寸: {current_width}x{current_height}, 比例: {current_ratio:.2f}")
            logger.info(f"目标尺寸: {target_width_px}x{target_height_px}, 比例: {target_ratio:.2f}")

            # 使用LANCZOS高质量重采样，直接拉伸到目标尺寸
            resized = img.resize((target_width_px, target_height_px), Image.LANCZOS)
            logger.info(f"已拉伸到3:2比例 ({target_width_px}x{target_height_px})")

            # 设置DPI为300，确保打印质量
            resized.save(output_path, dpi=(300, 300), quality=95)

            logger.info(f"截图已保存: {output_path}")
            return True

        except Exception as e:
            logger.error(f"截图失败 {html_file}: {e}")
            return False

    def close(self):
        """关闭浏览器"""
        if self.driver:
            self.driver.quit()
            self.driver = None


def generate_report(all_results: Dict, output_dir: Path):
    """生成Excel汇总报告"""
    def build_table(mode_key: str) -> pd.DataFrame:
        rows = []
        # 按样机编号组织数据（行）
        for num in range(1, 7):
            row = {'样机编号': f'{num}号样机'}
            for voltage in VOLTAGE_SEGMENTS:
                if num in all_results and mode_key in all_results[num]:
                    data = all_results[num][mode_key]
                    if voltage in data:
                        min_i, max_i, avg_i = data[voltage]
                        unit = all_results[num].get('unit_' + mode_key, 'A')
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
            rows.append(row)
        return pd.DataFrame(rows)

    work_df = build_table('work')

    # 在工作模式表中追加休眠电流列（18V后面）
    sleep_col = []
    sleep_avg_col = []
    for num in range(1, 7):
        if num in all_results and 'sleep' in all_results[num]:
            min_i, max_i, avg_i = all_results[num]['sleep']
            unit = all_results[num].get('unit_sleep', 'A')
            if unit == 'A':
                sleep_col.append(f"{min_i*1e6:.2f}-{max_i*1e6:.2f}uA")
                sleep_avg_col.append(f"{avg_i*1e6:.2f}uA")
            else:
                sleep_col.append(f"{min_i:.2f}-{max_i:.2f}{unit}")
                sleep_avg_col.append(f"{avg_i:.2f}{unit}")
        else:
            sleep_col.append("N/A")
            sleep_avg_col.append("N/A")
    work_df['休眠电流(14V)'] = sleep_col
    work_df['休眠电流(14V)平均'] = sleep_avg_col

    output_file = output_dir / "电压电流分析汇总.xlsx"
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            work_df.to_excel(writer, sheet_name='工作模式', index=False)
            # 设置字体为Arial 9
            workbook = writer.book
            worksheet = writer.sheets['工作模式']
            from openpyxl.styles import Font
            arial_font = Font(name='Arial', size=9)
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.font = arial_font
        logger.info(f"汇总报告已保存: {output_file}")
    except PermissionError:
        # 文件被占用时，尝试用带时间戳的文件名
        import datetime
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        alt_file = output_dir / f"电压电流分析汇总_{ts}.xlsx"
        with pd.ExcelWriter(alt_file, engine='openpyxl') as writer:
            work_df.to_excel(writer, sheet_name='工作模式', index=False)
            # 设置字体为Arial 9
            workbook = writer.book
            worksheet = writer.sheets['工作模式']
            from openpyxl.styles import Font
            arial_font = Font(name='Arial', size=9)
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.font = arial_font
        logger.warning(f"原文件被占用，已保存到: {alt_file}")

    print("\n" + "=" * 80)
    print("工作模式数据汇总:")
    print("=" * 80)
    print(work_df.to_string(index=False))

    return output_file


def process_all(base_dir: Path, output_dir: Path):
    """批量处理所有文件夹"""
    all_results = {}
    screenshot = ChartScreenshot()

    # 创建分类截图目录
    work_screenshot_dir = output_dir / "工作模式截图"
    static_screenshot_dir = output_dir / "静态模式截图"
    work_screenshot_dir.mkdir(exist_ok=True)
    static_screenshot_dir.mkdir(exist_ok=True)

    try:
        screenshot.init_browser()
        browser_ok = True
    except Exception:
        browser_ok = False
        logger.warning("浏览器初始化失败，将跳过截图功能")

    total = 6
    for i in range(1, total + 1):
        folder = base_dir / str(i)
        if not folder.exists():
            logger.warning(f"文件夹不存在: {folder}")
            continue

        print(f"\n{'='*60}")
        print(f"处理 {i}号样机 ({i}/{total})")
        print(f"{'='*60}")

        html_files = sorted([f for f in folder.iterdir() if f.suffix == '.html'])
        if len(html_files) == 0:
            logger.warning(f"{folder} 中没有HTML文件，跳过")
            continue

        # 通过文件名末尾是否带-l区分：带-l的是休眠模式，不带-l的是工作模式
        work_file = None
        static_file = None
        for hf in html_files:
            # 检查文件名（不含扩展名）是否以-l结尾
            name_without_ext = hf.stem  # 例如 "D2J P03-l" 或 "D2J P03"
            if name_without_ext.endswith('-l'):
                static_file = hf  # 文件名末尾带-l → 休眠/静态模式
            else:
                work_file = hf   # 文件名末尾不带-l → 工作模式

        if not work_file and not static_file:
            logger.warning(f"{folder} 中未找到有效的HTML文件")
            continue

        all_results[i] = {}

        # 分析工作模式
        if work_file:
            print(f"\n分析工作模式: {work_file.name}")
            try:
                analyzer_w = VoltageCurrentAnalyzer(str(work_file))
                all_results[i]['work'] = analyzer_w.analyze()
                all_results[i]['unit_work'] = analyzer_w.current_unit
            except Exception as e:
                logger.error(f"工作模式分析失败: {e}")

        # 分析休眠模式（从工作模式文件中14V约1分钟的段）
        if work_file:
            print(f"\n分析休眠模式...")
            try:
                sleep_result = analyzer_w.analyze_sleep()
                if sleep_result:
                    all_results[i]['sleep'] = sleep_result
                    all_results[i]['unit_sleep'] = analyzer_w.current_unit
            except Exception as e:
                logger.error(f"休眠模式分析失败: {e}")

        # 分析静态模式
        if static_file:
            print(f"\n分析静态模式: {static_file.name}")
            try:
                analyzer_s = VoltageCurrentAnalyzer(str(static_file))
                all_results[i]['static'] = analyzer_s.analyze()
                all_results[i]['unit_static'] = analyzer_s.current_unit
            except Exception as e:
                logger.error(f"静态模式分析失败: {e}")

        # 截图 - 无论是否检测到电流数据，都要截图
        if browser_ok:
            if work_file:
                work_img = work_screenshot_dir / f"{i}号样机工作电压电流.png"
                logger.info(f"截取工作模式图表: {work_file.name}")
                screenshot.capture(str(work_file), str(work_img))

            if static_file:
                static_img = static_screenshot_dir / f"{i}号样机静态电流.png"
                logger.info(f"截取静态模式图表: {static_file.name}")
                screenshot.capture(str(static_file), str(static_img))

    if browser_ok:
        screenshot.close()

    # 生成汇总报告
    print("\n" + "=" * 60)
    print("生成汇总报告...")
    generate_report(all_results, output_dir)

    return all_results


def main():
    print("=" * 60)
    print("  电压电流曲线自动化分析脚本")
    print("=" * 60)

    base_dir = Path(__file__).parent / "html"
    output_dir = Path(__file__).parent / "output"
    output_dir.mkdir(exist_ok=True)

    if not base_dir.exists():
        logger.error(f"找不到html文件夹: {base_dir}")
        sys.exit(1)

    process_all(base_dir, output_dir)

    print("\n" + "=" * 60)
    print("分析完成！结果保存在 output 文件夹中")
    print("=" * 60)


if __name__ == "__main__":
    main()
