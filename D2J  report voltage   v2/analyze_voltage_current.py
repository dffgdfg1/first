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
import socket
import warnings
warnings.filterwarnings('ignore')

# ============ 配置参数 ============
# 项目配置
PROJECT_CONFIGS = {
    'D2J': {
        'voltage_segments': [6.5, 9, 14, 16, 18],
        'name': 'D2J项目',
        'min_14v_current': 0.20  # D2J项目14V最小电流阈值（A）
    },
    'G2V': {
        'voltage_segments': [9, 14, 16],
        'name': 'G2V项目',
        'min_14v_current': 0.01  # G2V项目14V最小电流阈值（A），允许更小的电流
    }
}

VOLTAGE_SEGMENTS = [6.5, 9, 14, 16, 18]  # 默认D2J
VOLTAGE_TOLERANCE = 1.0       # 电压匹配容差
MIN_STABLE_POINTS = 30        # 最小稳定数据点数
OUTLIER_STD_FACTOR = 3.5      # 异常值标准差倍数（更宽松，保留更多真实数据）
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

    def __init__(self, html_file: str, outlier_factor: float = None, project_type: str = 'D2J'):
        self.html_file = html_file
        self.voltage_points = []  # [(timestamp, voltage), ...]
        self.current_points = []  # [(timestamp, current), ...]
        self.current_unit = 'A'
        # 使用传入的异常值倍数，如果没有传入则使用全局默认值
        self.outlier_factor = outlier_factor if outlier_factor is not None else OUTLIER_STD_FACTOR
        # 项目类型和对应的电压段
        self.project_type = project_type
        project_config = PROJECT_CONFIGS.get(project_type, PROJECT_CONFIGS['D2J'])
        self.voltage_segments = project_config['voltage_segments']
        self.min_14v_current = project_config.get('min_14v_current', 0.20)

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

        # 按时间戳对齐电压和电流数据
        voltage_dict = {ts: v for ts, v in self.voltage_points}
        aligned_data = []

        for curr_ts, curr_val in self.current_points:
            # 查找最接近的电压时间戳
            if curr_ts in voltage_dict:
                volt_val = voltage_dict[curr_ts]
            else:
                # 找到最近的电压时间戳
                nearest_ts = min(voltage_dict.keys(), key=lambda ts: abs(ts - curr_ts))
                volt_val = voltage_dict[nearest_ts]
            aligned_data.append((volt_val, curr_val))

        self._voltages = np.array([v for v, c in aligned_data])
        self._currents = np.array([c for v, c in aligned_data])

        logger.info(f"对齐后数据点数: {len(self._voltages)}")

        results = {}
        for target_v in self.voltage_segments:
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

                # 按时间戳对齐电压和电流数据
                voltage_dict = {ts: v for ts, v in self.voltage_points}
                aligned_data = []

                for curr_ts, curr_val in self.current_points:
                    # 查找最接近的电压时间戳
                    if curr_ts in voltage_dict:
                        volt_val = voltage_dict[curr_ts]
                    else:
                        # 找到最近的电压时间戳
                        nearest_ts = min(voltage_dict.keys(), key=lambda ts: abs(ts - curr_ts))
                        volt_val = voltage_dict[nearest_ts]
                    aligned_data.append((volt_val, curr_val))

                self._voltages = np.array([v for v, c in aligned_data])
                self._currents = np.array([c for v, c in aligned_data])
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

                # 按时间戳对齐电压和电流数据
                voltage_dict = {ts: v for ts, v in self.voltage_points}
                aligned_data = []

                for curr_ts, curr_val in self.current_points:
                    # 查找最接近的电压时间戳
                    if curr_ts in voltage_dict:
                        volt_val = voltage_dict[curr_ts]
                    else:
                        # 找到最近的电压时间戳
                        nearest_ts = min(voltage_dict.keys(), key=lambda ts: abs(ts - curr_ts))
                        volt_val = voltage_dict[nearest_ts]
                    aligned_data.append((volt_val, curr_val))

                self._voltages = np.array([v for v, c in aligned_data])
                self._currents = np.array([c for v, c in aligned_data])
            except Exception:
                return None

        voltages = self._voltages
        currents = self._currents

        # ????????????????????RT-9V????9V?????14V?
        voltage_groups = []
        for target_v in self.voltage_segments:
            mask = np.abs(voltages - target_v) <= VOLTAGE_TOLERANCE
            indices = np.where(mask)[0]
            if len(indices) == 0:
                continue
            for g in self._split_continuous(indices.tolist()):
                voltage_groups.append((target_v, g))

        if not voltage_groups:
            return None

        sleep_groups = []
        for group_idx, (target_v, g) in enumerate(voltage_groups):
            if len(g) < 50:
                continue

            if self.current_unit == 'uA':
                group_currents = currents[g]

                low_current_mask = group_currents < 10000
                low_current_count = np.sum(low_current_mask)

                if low_current_count >= 50:
                    low_currents = group_currents[low_current_mask]
                    avg_current = np.mean(low_currents)
                    max_current = np.max(low_currents)

                    if avg_current < 1000 and max_current < 5000:
                        sleep_groups.append((target_v, g))
                        logger.info(f"  ???{target_v}V?????{group_idx+1} (?{len(g)}??, ????:{low_current_count}?, ??:{avg_current:.0f}uA)")
                    else:
                        logger.debug(f"  ??????{group_idx+1} (?{len(g)}??, ??:{avg_current:.0f}uA, ??:{max_current:.0f}uA)")
                else:
                    logger.debug(f"  ??????{group_idx+1} (?{len(g)}??, ?????{low_current_count}?)")
            else:
                group_currents = currents[g]
                nonzero_mask = group_currents > 0
                nonzero_count = np.sum(nonzero_mask)
                if nonzero_count >= 50:
                    sleep_groups.append((target_v, g))
                    logger.info(f"  ???{target_v}V?????{group_idx+1} (?{len(g)}??, ?????:{nonzero_count}?)")
                else:
                    logger.debug(f"  ???{group_idx+1} (?{len(g)}??, ??????{nonzero_count}?)")

        if not sleep_groups:
            return None

        all_stable = []
        sleep_voltage = None
        for target_v, g in sleep_groups:
            if sleep_voltage is None:
                sleep_voltage = target_v
            seg_c = currents[g]

            if self.current_unit == 'uA':
                low_mask = (seg_c > 0) & (seg_c < 10000)
                seg_c = seg_c[low_mask]
            else:
                seg_c = seg_c[seg_c > 0]

            trim = int(len(seg_c) * 0.05)
            if trim > 0 and len(seg_c) > trim * 2:
                stable = seg_c[trim:-trim]
            else:
                stable = seg_c
            all_stable.extend(stable.tolist())

        if not all_stable:
            return None

        arr = np.array(all_stable)
        filtered = self._remove_outliers(arr)
        if len(filtered) == 0:
            filtered = arr
        result = (float(np.min(filtered)), float(np.max(filtered)), float(np.mean(filtered)))
        voltage_label = f"{sleep_voltage}V" if sleep_voltage is not None else "??"
        logger.info(f"  ??{voltage_label}: {result[0]:.6f}-{result[1]:.6f} (??:{result[2]:.6f}) {self.current_unit}")
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

            # 过滤掉极小值（过滤掉明显的低电流过渡阶段）
            # 使用项目特定的阈值：D2J为0.20A，G2V为0.01A
            seg_currents = seg_currents[seg_currents >= self.min_14v_current]

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
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--disable-gpu')
        options.add_argument('--disable-software-rasterizer')
        try:
            # 在创建driver前设置socket级超时（适用于所有Selenium版本）
            socket.setdefaulttimeout(600)
            # 直接使用 EdgeWebDriver 类而不是 webdriver.Edge()，避免 PyInstaller 打包问题
            self.driver = EdgeWebDriver(options=options)
            # 设置页面加载超时为600秒（10分钟）
            self.driver.set_page_load_timeout(600)
            # 设置脚本执行超时为600秒
            self.driver.set_script_timeout(600)
            # 修改 command_executor 的 client_config timeout，覆盖 urllib3 请求的 read timeout
            # （socket.setdefaulttimeout 不能覆盖 urllib3 的显式超时，必须直接修改此值）
            self.driver.command_executor._client_config.timeout = 600
            logger.info("浏览器初始化成功")
        except Exception as e:
            logger.error(f"浏览器初始化失败: {e}")
            logger.info("请确保已安装Edge浏览器和对应版本的EdgeDriver")
            raise

    def capture(self, html_file: str, output_path: str, zoom_hours: float = None) -> bool:
        """
        截取HTML图表并保存为PNG，使用ECharts getDataURL直接导出画布，
        避免save_screenshot传输整个浏览器窗口导致的超时问题。

        Args:
            html_file: HTML文件路径
            output_path: 输出PNG路径
        Returns:
            bool: 是否成功
        """
        import base64
        if not self.driver:
            self.init_browser()

        try:
            file_url = Path(html_file).resolve().as_uri()
            logger.info(f"[截图] 加载页面: {file_url}")
            self.driver.get(file_url)
            logger.info(f"[截图] 页面加载完成，等待canvas元素...")

            # 等待canvas元素出现（最多3分钟）
            WebDriverWait(self.driver, 180).until(
                lambda d: d.execute_script("return document.querySelector('canvas') !== null")
            )
            logger.info(f"[截图] canvas元素已出现，等待ECharts实例初始化...")

            # 通过JS轮询等待ECharts实例初始化完成，最多等待60秒
            WebDriverWait(self.driver, 60).until(
                lambda d: d.execute_script(
                    "var c = echarts.getInstanceByDom(document.getElementById('chart-container'));"
                    "return c !== null && c !== undefined;"
                )
            )
            logger.info(f"[截图] ECharts实例就绪，准备图表...")

            # 准备图表：重置/设置dataZoom、隐藏控件、调整尺寸
            # zoom_hours有值时截取数据中间约N小时，用于生成局部截图
            self.driver.execute_script("""
                var chart = echarts.getInstanceByDom(document.getElementById('chart-container'));
                if (!chart) return;
                var zoomHours = arguments[0];
                var zoomStart = 0;
                var zoomEnd = 100;
                var option = chart.getOption();
                if (zoomHours && zoomHours > 0) {
                    var minTs = Infinity;
                    var maxTs = -Infinity;
                    var tsCount = 0;
                    (option.series || []).forEach(function(series) {
                        (series.data || []).forEach(function(point) {
                            if (Array.isArray(point) && point.length >= 2 && typeof point[0] === 'number') {
                                if (point[0] < minTs) minTs = point[0];
                                if (point[0] > maxTs) maxTs = point[0];
                                tsCount += 1;
                            }
                        });
                    });
                    if (tsCount > 1 && isFinite(minTs) && isFinite(maxTs)) {
                        var range = maxTs - minTs;
                        var zoomMs = zoomHours * 60 * 60 * 1000;
                        if (range > zoomMs) {
                            var center = minTs + range / 2;
                            var startTs = center - zoomMs / 2;
                            var endTs = center + zoomMs / 2;
                            zoomStart = Math.max(0, (startTs - minTs) / range * 100);
                            zoomEnd = Math.min(100, (endTs - minTs) / range * 100);
                        }
                    }
                }
                chart.setOption({
                    height: null,
                    dataZoom: [{
                        show: false,
                        xAxisIndex: [0],
                        filterMode: 'none',
                        start: zoomStart,
                        end: zoomEnd
                    }],
                    toolbox: {show: false},
                    grid: {bottom: 100},
                    xAxis: {
                        axisLabel: {
                            rotate: 0,
                            fontSize: 15,
                            lineHeight: 22,
                            formatter: '{yyyy}-{MM}-{dd}\\n{HH}:{mm}',
                            hideOverlap: true
                        }
                    }
                }, {replaceMerge: ['dataZoom']});
                chart.dispatchAction({ type: 'dataZoom', dataZoomIndex: 0, start: zoomStart, end: zoomEnd });
                var dom = chart.getDom();
                dom.style.width  = '1920px';
                dom.style.height = '1200px';
                chart.resize();
            """, zoom_hours)
            logger.info(f"[截图] 图表resize完成，等待重绘...")

            # 等待resize重绘完成（用JS检测，避免固定sleep）
            WebDriverWait(self.driver, 30).until(
                lambda d: d.execute_script(
                    "var c = echarts.getInstanceByDom(document.getElementById('chart-container'));"
                    "var w = c && c.getDom().offsetWidth;"
                    "return w >= 1900;"
                )
            )
            logger.info(f"[截图] 重绘完成，调用getDataURL导出...")

            # 用ECharts getDataURL直接导出画布为base64（pixelRatio=2保证清晰度）
            # 比save_screenshot快得多：无需传输整个浏览器窗口
            b64 = self.driver.execute_script("""
                var chart = echarts.getInstanceByDom(document.getElementById('chart-container'));
                if (!chart) return null;
                var url = chart.getDataURL({
                    type: 'png',
                    pixelRatio: 2,
                    backgroundColor: '#fff'
                });
                return url.replace('data:image/png;base64,', '');
            """)
            logger.info(f"[截图] getDataURL返回, 数据长度: {len(b64) if b64 else 0}")

            if not b64:
                logger.error(f"截图失败 {html_file}: getDataURL 返回空值")
                return False

            # base64解码为图片
            import io
            img_data = base64.b64decode(b64)
            img = Image.open(io.BytesIO(img_data))
            logger.info(f"[截图] 图片解码成功: {img.width}x{img.height}")

            # 直接拉伸到3:2宽高比（1831x1220，300DPI打印质量）
            target_width_px = 1831
            target_height_px = int(target_width_px / (3.0 / 2.0))  # = 1220px

            logger.info(f"原始尺寸: {img.width}x{img.height}, 比例: {img.width/img.height:.2f}")
            logger.info(f"目标尺寸: {target_width_px}x{target_height_px}, 比例: 1.50")

            resized = img.resize((target_width_px, target_height_px), Image.LANCZOS)
            logger.info(f"已拉伸到3:2比例 ({target_width_px}x{target_height_px})")

            resized.save(output_path, dpi=(300, 300))
            logger.info(f"截图已保存: {output_path}")
            return True

        except Exception as e:
            import traceback
            logger.error(f"截图失败 {html_file}: {type(e).__name__}: {e}")
            logger.info(f"截图异常详情:\n{traceback.format_exc()}")
            return False

    def close(self):
        """关闭浏览器"""
        if self.driver:
            self.driver.quit()
            self.driver = None


def generate_report(all_results: Dict, output_dir: Path, voltage_segments: List[float] = None):
    """生成Excel汇总报告"""
    if voltage_segments is None:
        voltage_segments = VOLTAGE_SEGMENTS

    def build_table(mode_key: str) -> pd.DataFrame:
        rows = []
        # 按样机编号组织数据（行）
        for num in range(1, 7):
            row = {'样机编号': f'{num}号样机'}
            for voltage in voltage_segments:
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
