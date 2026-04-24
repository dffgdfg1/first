#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
温湿度曲线对比工具
加载两个 Excel 数据源，智能对齐时间轴，绘制 S1 / PV / SV 三条曲线。
依赖: pip install pandas openpyxl matplotlib scipy
"""

import struct
import datetime
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.dates as mdates
import matplotlib.ticker as mticker
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure

# ── .740 文件解析 ─────────────────────────────────────────────────────────────

class Parser740:
    """解析 .740 二进制数据文件（每条记录 112 字节）"""
    RECORD_SIZE = 112

    def parse(self, filepath):
        with open(filepath, 'rb') as f:
            data = f.read()
        if len(data) % self.RECORD_SIZE != 0:
            raise ValueError(
                f"文件大小 {len(data)} 字节不是 {self.RECORD_SIZE} 的整数倍，"
                "可能不是有效的 .740 文件"
            )
        records = []
        for i in range(len(data) // self.RECORD_SIZE):
            rec = data[i * self.RECORD_SIZE:(i + 1) * self.RECORD_SIZE]
            if rec[0:1] != b'\x1a':
                raise ValueError(f"记录 {i} 头部标识不匹配: {rec[0:4].hex()}")
            ts = struct.unpack_from('<I', rec, 8)[0]
            channels = list(struct.unpack_from('<12f', rec, 16))
            records.append({'time': datetime.datetime.fromtimestamp(ts), 'channels': channels})
        return records


def load_740_as_dataframe(path: str) -> pd.DataFrame:
    """将 .740 文件解析为 DataFrame，列名为 通道1..通道12，时间为索引。"""
    parser = Parser740()
    records = parser.parse(path)
    rows = {f'通道{i+1}': [r['channels'][i] for r in records] for i in range(12)}
    df = pd.DataFrame(rows, index=pd.DatetimeIndex([r['time'] for r in records]))
    df.index.name = '时间'
    # 去掉全零的通道列
    df = df.loc[:, (df != 0).any()]
    return df.sort_index()


# ── 全局样式 ──────────────────────────────────────────────────────────────────
matplotlib.rcParams.update({
    "font.family":       ["Microsoft YaHei", "SimHei", "DejaVu Sans"],
    "axes.facecolor":    "#FFFFFF",
    "figure.facecolor":  "#FFFFFF",
    "axes.edgecolor":    "#CCCCCC",
    "axes.grid":         True,
    "grid.color":        "#E8E8E8",
    "grid.linewidth":    0.8,
    "axes.spines.top":   False,
    "axes.spines.right": False,
    "xtick.color":       "#666666",
    "ytick.color":       "#666666",
    "axes.labelcolor":   "#444444",
    "font.size":         10,
})

COLOR_S1  = "#4472C4"
COLOR_PV  = "#70AD47"
COLOR_SV  = "#FFC000"
COLOR_HUM = "#9DC3E6"


# ── 数据工具 ──────────────────────────────────────────────────────────────────

def load_datafile(path: str) -> pd.DataFrame:
    """读取 Excel 或 CSV，自动识别时间列并设为索引。
    CSV 使用 GBK 编码，'----' 替换为 NaN。
    兼容 pandas 3.x（移除了 infer_datetime_format）。"""
    p = Path(path)
    if p.suffix.lower() == ".csv":
        df = pd.read_csv(path, encoding="gbk", na_values=["----", "---", "--"])
    else:
        df = pd.read_excel(path, header=0)

    # 逐列尝试：跳过数值列，datetime64 直接用，字符串列尝试解析
    # 注意：pandas 3.x 字符串列 dtype 为 str，dtype==object 会是 False，
    #       故改用 is_numeric_dtype 来排除数值列
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            continue
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            return df.set_index(col).sort_index()
        try:
            parsed = pd.to_datetime(df[col])   # pandas 3.x: 无 infer_datetime_format
            if parsed.notna().mean() > 0.8:
                df[col] = parsed
                return df.set_index(col).sort_index()
        except Exception:
            continue

    # 3. 兜底：尝试把索引解析为时间
    try:
        df.index = pd.to_datetime(df.index)
    except Exception:
        pass
    return df.sort_index()


def rename_box_cols(df: pd.DataFrame) -> pd.DataFrame:
    """将温箱原始列名统一映射为标准名。
    支持：
      温度_PV / 温度_TPV / TPV  → Temperature_PV
      温度_SP / 温度_SV         → Temperature_SV
      湿度_PV                   → Humidity
    已是标准名的列直接跳过。"""
    STANDARD = {"Temperature_PV", "Temperature_SV", "Humidity", "Temperature_S1"}
    mapping = {}
    for c in df.columns:
        if c in STANDARD:
            continue
        cs = str(c)
        cu = cs.upper()
        if cu.strip() == "D91":
            mapping[c] = "Temperature_PV"
        elif cu.strip() == "D92":
            mapping[c] = "Temperature_SV"
        elif "温度" in cs:
            if any(k in cu for k in ["_TPV", "_PV"]) or cu.endswith("PV"):
                mapping[c] = "Temperature_PV"
            elif any(k in cu for k in ["_SP", "_SV"]) or cu.endswith(("SP", "SV")):
                mapping[c] = "Temperature_SV"
        elif cu.strip() in ("TPV", "温度_TPV"):
            mapping[c] = "Temperature_PV"
        elif "湿度" in cs and "PV" in cu:
            mapping[c] = "Humidity"
    return df.rename(columns=mapping)


def rename_sensor_cols(df: pd.DataFrame) -> pd.DataFrame:
    """将外接传感器的通道列名映射为 Temperature_S1（通道1）。"""
    mapping = {}
    for c in df.columns:
        cl = str(c)
        if "通道1" == cl.strip():
            mapping[c] = "Temperature_S1"
    return df.rename(columns=mapping)


def numeric_cols(df: pd.DataFrame) -> list:
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]


def find_col(columns: list, keywords: list):
    """按关键字（不区分大小写）在列名中匹配，找到第一个就返回。"""
    for kw in keywords:
        for c in columns:
            if kw.lower() in str(c).lower():
                return c
    return None


def best_offset_seconds(ref: pd.Series, target: pd.Series, max_min: int = 60) -> float:
    """
    互相关法估算 target 相对于 ref 的最佳时间偏移（秒）。
    返回正值 → target 需要向后移；负值 → 需要向前移。
    """
    start = max(ref.index.min(), target.index.min())
    end   = min(ref.index.max(), target.index.max())
    if start >= end:
        return 0.0
    idx = pd.date_range(start, end, freq="1min")
    tol = pd.Timedelta("2min")
    a = ref.reindex(idx, method="nearest", tolerance=tol).ffill().dropna()
    b = target.reindex(idx, method="nearest", tolerance=tol).ffill().dropna()
    common = a.index.intersection(b.index)
    if len(common) < 10:
        return 0.0
    va = (a[common] - a[common].mean()).to_numpy()
    vb = (b[common] - b[common].mean()).to_numpy()
    corr = np.correlate(va, vb, mode="full")
    lags  = np.arange(-(len(va) - 1), len(vb))          # 单位：分钟
    mask  = (lags >= -max_min) & (lags <= max_min)
    best  = lags[mask][np.argmax(corr[mask])]
    return float(best) * 60.0


def smooth(s: pd.Series, win: int) -> pd.Series:
    if win <= 1:
        return s
    return s.rolling(win, center=True, min_periods=1).mean()


def fit_to_ref(ref: pd.Series, target: pd.Series,
               alpha: float, smooth_win: int) -> pd.Series:
    """
    将 target 曲线向 ref 曲线靠拢，但保留真实物理差异。

    算法：加权混合
      1. 对 ref 线性插值到 target 的时间点
      2. 对 target 自身做平滑（去锯齿）
      3. result = alpha * ref_插值 + (1-alpha) * target_平滑

    alpha=0  → 只平滑自身，不靠拢
    alpha=1  → 完全重合到参考曲线
    建议范围：0.2 ~ 0.6
    """
    from scipy.signal import savgol_filter

    # 找共同时间范围
    t_start = max(ref.index.min(), target.index.min())
    t_end   = min(ref.index.max(), target.index.max())
    if t_start >= t_end:
        return target

    target_c = target[(target.index >= t_start) & (target.index <= t_end)].dropna()
    ref_c    = ref[(ref.index >= t_start) & (ref.index <= t_end)].dropna()
    if len(target_c) < 4 or len(ref_c) < 2:
        return target

    # 1. 把 ref 插值到 target 的时间点（线性插值，稳定可靠）
    ref_interp = ref_c.reindex(ref_c.index.union(target_c.index)) \
                      .interpolate(method="time") \
                      .reindex(target_c.index)

    # 2. 平滑 target 自身
    win = max(5, smooth_win | 1)
    tgt_vals = target_c.values.astype(float)
    if len(tgt_vals) > win:
        tgt_smooth = savgol_filter(tgt_vals, window_length=win, polyorder=3)
    else:
        tgt_smooth = tgt_vals

    # 3. 加权混合
    ref_vals = ref_interp.values.astype(float)
    result   = alpha * ref_vals + (1.0 - alpha) * tgt_smooth

    return pd.Series(result, index=target_c.index)


# ── 主应用 ────────────────────────────────────────────────────────────────────

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("温湿度曲线对比工具")
        self.state("zoomed")
        self.configure(bg="#F5F5F5")

        self.df_s1:  pd.DataFrame = None
        self.df_box: pd.DataFrame = None
        self._first_dir: str = ""   # 首个加载文件的目录

        # 控件变量
        self.path_s1   = tk.StringVar()
        self.path_box  = tk.StringVar()
        self.col_s1    = tk.StringVar()
        self.col_pv    = tk.StringVar()
        self.col_sv    = tk.StringVar()
        self.col_hum   = tk.StringVar(value="无")

        self.show_s1   = tk.BooleanVar(value=True)
        self.show_pv   = tk.BooleanVar(value=True)
        self.show_sv   = tk.BooleanVar(value=True)
        self.show_hum  = tk.BooleanVar(value=False)
        self.show_dev  = tk.BooleanVar(value=False)

        self.offset_min  = tk.DoubleVar(value=0.0)
        self.shift_a_min = tk.DoubleVar(value=0.0)   # 数据源 A 绝对平移
        self.shift_b_min = tk.DoubleVar(value=0.0)   # 数据源 B 绝对平移
        self.smooth_win  = tk.IntVar(value=1)
        self.dev_thresh  = tk.DoubleVar(value=5.0)
        self.status_var  = tk.StringVar(value="就绪 — 请加载两个 Excel 文件")

        # 参考曲线拟合
        self.ref_curve   = tk.StringVar(value="无")   # 无/S1/PV/SV
        self.spline_s    = tk.DoubleVar(value=1.0)    # 样条平滑因子
        self.resid_win   = tk.IntVar(value=9)          # 残差平滑窗口

        # Y 轴最小跨度（°C），防止曲线接近时坐标轴刻度过密
        self.y_min_span  = tk.DoubleVar(value=5.0)

        self._build_ui()

    # ── UI 搭建 ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        pane = tk.PanedWindow(self, orient="horizontal", sashwidth=4,
                              bg="#DDDDDD", sashrelief="flat")
        pane.pack(fill="both", expand=True, padx=4, pady=(4, 0))

        # 左侧：可滚动容器
        left_outer = tk.Frame(pane, bg="#F0F0F0", width=275)
        pane.add(left_outer, minsize=250, stretch="never")

        left_canvas = tk.Canvas(left_outer, bg="#F0F0F0",
                                highlightthickness=0, width=258)
        left_sb = ttk.Scrollbar(left_outer, orient="vertical",
                                command=left_canvas.yview)
        left_canvas.configure(yscrollcommand=left_sb.set)
        left_sb.pack(side="right", fill="y")
        left_canvas.pack(side="left", fill="both", expand=True)

        left = tk.Frame(left_canvas, bg="#F0F0F0")
        _win = left_canvas.create_window((0, 0), window=left, anchor="nw")

        left.bind("<Configure>",
                  lambda _: left_canvas.configure(
                      scrollregion=left_canvas.bbox("all")))
        left_canvas.bind("<Configure>",
                         lambda e: left_canvas.itemconfig(_win, width=e.width))

        # 鼠标在左侧面板任意子控件上都能滚动
        def _scroll(e):
            left_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        left_canvas.bind_all("<MouseWheel>", _scroll)

        right = tk.Frame(pane, bg="#FFFFFF")
        pane.add(right, minsize=600, stretch="always")

        self._build_controls(left)
        self._build_chart(right)

        # 右侧图表区取消全局滚轮绑定（还给 matplotlib 缩放用）
        right.bind("<Enter>", lambda _: left_canvas.unbind_all("<MouseWheel>"))
        right.bind("<Leave>", lambda _: left_canvas.bind_all("<MouseWheel>", _scroll))

        tk.Label(self, textvariable=self.status_var, anchor="w", padx=8,
                 bg="#E0E0E0", fg="#555555").pack(side="bottom", fill="x")

    def _sec(self, parent, title: str):
        f = tk.LabelFrame(parent, text=title, bg="#F0F0F0",
                          fg="#333333", font=("Microsoft YaHei", 9, "bold"),
                          bd=1, relief="groove", padx=6, pady=4)
        f.pack(fill="x", padx=6, pady=4)
        return f

    def _build_controls(self, parent):
        # ── 数据源 A ─────────────────────────────────────────────────
        sa = self._sec(parent, "数据源 A（外接传感器）")
        sa.columnconfigure(0, weight=1)
        tk.Entry(sa, textvariable=self.path_s1, width=20,
                 font=("Consolas", 8)).grid(row=0, column=0, sticky="ew")
        _btn(sa, "浏览", self._load_s1, COLOR_S1).grid(row=0, column=1, padx=(4, 0))
        tk.Label(sa, text="S1 温度列:", bg="#F0F0F0",
                 font=("Microsoft YaHei", 9)).grid(row=1, column=0, sticky="w", pady=2)
        self._cb_s1 = ttk.Combobox(sa, textvariable=self.col_s1,
                                    state="readonly", width=22)
        self._cb_s1.grid(row=2, column=0, columnspan=2, sticky="ew")

        # ── 数据源 B ─────────────────────────────────────────────────
        sb = self._sec(parent, "数据源 B（温箱监控）")
        sb.columnconfigure(0, weight=1)
        tk.Entry(sb, textvariable=self.path_box, width=20,
                 font=("Consolas", 8)).grid(row=0, column=0, sticky="ew")
        _btn(sb, "浏览", self._load_box, COLOR_S1).grid(row=0, column=1, padx=(4, 0))
        specs = [
            ("PV 实际值列:", self.col_pv, "_cb_pv"),
            ("SV 设定值列:", self.col_sv, "_cb_sv"),
            ("湿度列:",      self.col_hum, "_cb_hum"),
        ]
        for row_i, (lbl, var, attr) in enumerate(specs, start=1):
            tk.Label(sb, text=lbl, bg="#F0F0F0",
                     font=("Microsoft YaHei", 9)).grid(
                         row=row_i * 2 - 1, column=0, sticky="w", pady=2)
            cb = ttk.Combobox(sb, textvariable=var, state="readonly", width=22)
            cb.grid(row=row_i * 2, column=0, columnspan=2, sticky="ew")
            setattr(self, attr, cb)

        for var in (self.col_s1, self.col_pv, self.col_sv, self.col_hum):
            var.trace_add("write", lambda *_: self._redraw())

        # ── 时间对齐 ─────────────────────────────────────────────────
        sc = self._sec(parent, "时间对齐（S1 时间偏移）")
        tk.Label(sc, text="偏移量（分钟）：", bg="#F0F0F0",
                 font=("Microsoft YaHei", 9)).pack(anchor="w")
        self._sld = tk.Scale(sc, from_=-120, to=120, resolution=0.5,
                              orient="horizontal", variable=self.offset_min,
                              command=lambda _: self._redraw(),
                              bg="#F0F0F0", highlightthickness=0, length=230,
                              troughcolor="#DDDDDD", activebackground=COLOR_S1)
        self._sld.pack(fill="x")
        row = tk.Frame(sc, bg="#F0F0F0"); row.pack(fill="x", pady=2)
        tk.Entry(row, textvariable=self.offset_min, width=8).pack(side="left")
        tk.Label(row, text="分", bg="#F0F0F0").pack(side="left", padx=2)
        tk.Button(row, text="重置 0", bg="#F0F0F0", relief="flat",
                  command=self._reset_offset).pack(side="left", padx=6)
        tk.Button(sc, text="⚡ 自动对齐", bg="#70AD47", fg="white", relief="flat",
                  font=("Microsoft YaHei", 9, "bold"),
                  command=self._auto_align).pack(fill="x", pady=(4, 0))

        # ── 时间戳平移 ────────────────────────────────────────────────
        st = self._sec(parent, "时间戳平移（绝对平移）")
        tk.Label(st, text="对任意一方数据前移（-）或后移（+）。",
                 bg="#F0F0F0", fg="#666666",
                 font=("Microsoft YaHei", 8), wraplength=220,
                 justify="left").pack(anchor="w", pady=(0, 4))

        for lbl_text, var, color in [
            ("数据源 A（外接传感器）偏移（分钟）：", self.shift_a_min, COLOR_S1),
            ("数据源 B（温箱监控）偏移（分钟）：",   self.shift_b_min, COLOR_PV),
        ]:
            tk.Label(st, text=lbl_text, bg="#F0F0F0",
                     fg=color, font=("Microsoft YaHei", 8)).pack(anchor="w")
            row_s = tk.Frame(st, bg="#F0F0F0"); row_s.pack(fill="x", pady=(0, 4))
            tk.Scale(row_s, from_=-1440, to=1440, resolution=1,
                     orient="horizontal", variable=var,
                     command=lambda _: self._redraw(),
                     bg="#F0F0F0", highlightthickness=0, length=180,
                     troughcolor="#DDDDDD",
                     activebackground=color).pack(side="left", fill="x", expand=True)
            tk.Entry(row_s, textvariable=var, width=6).pack(side="left", padx=(4, 0))
        row_rst = tk.Frame(st, bg="#F0F0F0"); row_rst.pack(fill="x", pady=(2, 0))
        tk.Button(row_rst, text="重置 A", bg="#F0F0F0", relief="flat",
                  command=lambda: (self.shift_a_min.set(0.0), self._redraw())
                  ).pack(side="left", padx=(0, 4))
        tk.Button(row_rst, text="重置 B", bg="#F0F0F0", relief="flat",
                  command=lambda: (self.shift_b_min.set(0.0), self._redraw())
                  ).pack(side="left")

        # ── 平滑 ─────────────────────────────────────────────────────
        sd = self._sec(parent, "数据平滑（移动平均）")
        rw = tk.Frame(sd, bg="#F0F0F0"); rw.pack(fill="x")
        tk.Label(rw, text="窗口大小：", bg="#F0F0F0").pack(side="left")
        tk.Spinbox(rw, from_=1, to=120, textvariable=self.smooth_win,
                   width=5, command=self._redraw).pack(side="left", padx=4)
        tk.Label(rw, text="点", bg="#F0F0F0").pack(side="left")

        # ── 参考曲线拟合 ─────────────────────────────────────────────
        sf = self._sec(parent, "参考曲线拟合")
        tk.Label(sf, text="选一条曲线为基准，其余两条向其靠拢。",
                 bg="#F0F0F0", fg="#666666",
                 font=("Microsoft YaHei", 8), wraplength=220,
                 justify="left").pack(anchor="w", pady=(0, 4))

        row_ref = tk.Frame(sf, bg="#F0F0F0"); row_ref.pack(fill="x")
        tk.Label(row_ref, text="参考曲线：", bg="#F0F0F0").pack(side="left")
        self._cb_ref = ttk.Combobox(row_ref, textvariable=self.ref_curve,
                                     values=["无", "S1", "PV", "SV"],
                                     state="readonly", width=6)
        self._cb_ref.pack(side="left", padx=4)
        self.ref_curve.trace_add("write", lambda *_: self._redraw())

        tk.Label(sf, text="靠拢强度 α（0=只平滑，1=完全重合）：",
                 bg="#F0F0F0", font=("Microsoft YaHei", 8)).pack(anchor="w", pady=(6, 0))
        self._sld_s = tk.Scale(sf, from_=0.0, to=1.0, resolution=0.01,
                                orient="horizontal", variable=self.spline_s,
                                command=lambda _: self._redraw(),
                                bg="#F0F0F0", highlightthickness=0, length=220,
                                troughcolor="#DDDDDD", activebackground=COLOR_PV)
        self._sld_s.pack(fill="x")

        row_rw = tk.Frame(sf, bg="#F0F0F0"); row_rw.pack(fill="x", pady=2)
        tk.Label(row_rw, text="平滑窗口：", bg="#F0F0F0").pack(side="left")
        tk.Spinbox(row_rw, from_=5, to=51, increment=2,
                   textvariable=self.resid_win,
                   width=5, command=self._redraw).pack(side="left", padx=4)
        tk.Label(row_rw, text="点（奇数）", bg="#F0F0F0").pack(side="left")

        # ── 显示选项 ─────────────────────────────────────────────────
        se = self._sec(parent, "显示选项")
        for text, var, col in [
            ("Temperature_S1（外接）", self.show_s1,  COLOR_S1),
            ("Temperature_PV（实际）", self.show_pv,  COLOR_PV),
            ("Temperature_SV（设定）", self.show_sv,  COLOR_SV),
            ("湿度曲线",               self.show_hum, COLOR_HUM),
        ]:
            tk.Checkbutton(se, text=text, variable=var, bg="#F0F0F0",
                           fg=col, activeforeground=col, selectcolor="#F0F0F0",
                           font=("Microsoft YaHei", 9, "bold"),
                           command=self._redraw).pack(anchor="w")

        dr = tk.Frame(se, bg="#F0F0F0"); dr.pack(anchor="w", pady=2)
        tk.Checkbutton(dr, text="标注 S1/PV 偏差 >", variable=self.show_dev,
                       bg="#F0F0F0", selectcolor="#F0F0F0",
                       command=self._redraw).pack(side="left")
        tk.Entry(dr, textvariable=self.dev_thresh, width=5).pack(side="left")
        tk.Label(dr, text="°C", bg="#F0F0F0").pack(side="left")

        ry = tk.Frame(se, bg="#F0F0F0"); ry.pack(anchor="w", pady=2)
        tk.Label(ry, text="Y 轴最小范围：", bg="#F0F0F0",
                 font=("Microsoft YaHei", 9)).pack(side="left")
        tk.Entry(ry, textvariable=self.y_min_span, width=5).pack(side="left")
        tk.Label(ry, text="°C", bg="#F0F0F0").pack(side="left")

        # ── 导出 ─────────────────────────────────────────────────────
        sf = self._sec(parent, "导出")
        _btn(sf, "保存图表 PNG", self._export_png, COLOR_S1, full=True)
        _btn(sf, "导出合并数据 Excel", self._export_excel, "#70AD47", full=True)

        # 重置参数按钮
        tk.Button(parent, text="↺  重置参数", command=self._reset_params,
                  bg="#A0A0A0", fg="white", relief="flat",
                  font=("Microsoft YaHei", 9),
                  activebackground="#888888").pack(
                      fill="x", padx=6, pady=(0, 2))

        # 更新按钮
        tk.Button(parent, text="▶  更新图表", command=self._redraw,
                  bg="#FFC000", fg="white", relief="flat",
                  font=("Microsoft YaHei", 10, "bold"),
                  activebackground="#E6AC00", height=2).pack(
                      fill="x", padx=6, pady=(2, 8))

    def _build_chart(self, parent):
        self.fig = Figure(figsize=(12, 6), dpi=100, facecolor="#FFFFFF")
        self.ax1 = self.fig.add_subplot(111)
        self.ax2 = self.ax1.twinx()
        self.fig.subplots_adjust(top=0.92, bottom=0.13, left=0.07, right=0.93)

        self.canvas = FigureCanvasTkAgg(self.fig, master=parent)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        tb_frame = tk.Frame(parent, bg="#FFFFFF")
        tb_frame.pack(fill="x")
        self.toolbar = NavigationToolbar2Tk(self.canvas, tb_frame)
        self.toolbar.update()

    # ── 文件加载 ──────────────────────────────────────────────────────────────

    def _load_s1(self):
        p = filedialog.askopenfilename(
            title="选择外接传感器文件（.740 / Excel / CSV）",
            filetypes=[("支持的文件", "*.740 *.xlsx *.xls *.csv"), ("所有文件", "*.*")])
        if p:
            self.path_s1.set(p)
            threading.Thread(target=self._do_load_s1, args=(p,), daemon=True).start()

    def _do_load_s1(self, path):
        try:
            if Path(path).suffix.lower() == ".740":
                df = load_740_as_dataframe(path)
                df = rename_sensor_cols(df)   # 通道1 → Temperature_S1
            else:
                df = load_datafile(path)
                df = rename_sensor_cols(df)   # 通道1 → Temperature_S1
            self.df_s1 = df
            if not self._first_dir:
                self._first_dir = str(Path(path).parent)
            nc = numeric_cols(df)
            cols = ["无"] + nc
            guess = (find_col(nc, ["Temperature_S1", "s1", "通道1"])
                     or (nc[0] if nc else "无"))
            self.after(0, lambda: _set_cb(self._cb_s1, cols, guess))
            self.after(0, lambda: self.status_var.set(
                f"A 已加载：{Path(path).name}  |  {len(df)} 行，{len(nc)} 个数值列"))
            self.after(0, self._redraw)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("加载失败", str(e)))

    def _load_box(self):
        p = filedialog.askopenfilename(
            title="选择温箱监控文件（Excel / CSV）",
            filetypes=[("支持的文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")])
        if p:
            self.path_box.set(p)
            threading.Thread(target=self._do_load_box, args=(p,), daemon=True).start()

    def _do_load_box(self, path):
        try:
            df = load_datafile(path)
            df = rename_box_cols(df)   # 温度_PV→Temperature_PV, 温度_SP→Temperature_SV, 湿度_PV→Humidity
            self.df_box = df
            if not self._first_dir:
                self._first_dir = str(Path(path).parent)
            nc   = numeric_cols(df)
            opts = ["无"] + nc
            gpv  = find_col(nc, ["Temperature_PV"]) or (nc[0] if nc else "无")
            gsv  = find_col(nc, ["Temperature_SV"]) or (nc[1] if len(nc) > 1 else "无")
            ghum = find_col(nc, ["Humidity"]) or "无"
            self.after(0, lambda: _set_cb(self._cb_pv,  opts, gpv))
            self.after(0, lambda: _set_cb(self._cb_sv,  opts, gsv))
            self.after(0, lambda: _set_cb(self._cb_hum, opts, ghum))
            self.after(0, lambda: self.status_var.set(
                f"B 已加载：{Path(path).name}  |  {len(df)} 行，{len(nc)} 个数值列"))
            self.after(0, self._redraw)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("加载失败", str(e)))

    # ── 时间对齐 ──────────────────────────────────────────────────────────────

    def _reset_offset(self):
        self.offset_min.set(0.0)
        self._redraw()

    def _reset_params(self):
        self.offset_min.set(0.0)
        self.smooth_win.set(1)
        self.dev_thresh.set(5.0)
        self.y_min_span.set(5.0)
        self.ref_curve.set("无")
        self.spline_s.set(1.0)
        self.resid_win.set(9)
        self.show_s1.set(True)
        self.show_pv.set(True)
        self.show_sv.set(True)
        self.show_hum.set(False)
        self.show_dev.set(False)
        self._redraw()

    def _auto_align(self):
        if self.df_s1 is None or self.df_box is None:
            messagebox.showwarning("提示", "请先加载两个数据文件。")
            return
        if self.col_s1.get() in ("", "无") or self.col_pv.get() in ("", "无"):
            messagebox.showwarning("提示", "请先选择 S1 列和 PV 列。")
            return
        self.status_var.set("正在计算最优时间偏移…")
        threading.Thread(target=self._do_align, daemon=True).start()

    def _do_align(self):
        try:
            ref    = self.df_box[self.col_pv.get()].dropna()
            target = self.df_s1[self.col_s1.get()].dropna()
            off_s  = best_offset_seconds(ref, target, max_min=60)
            off_m  = round(max(-120, min(120, off_s / 60)), 2)
            self.after(0, lambda: self.offset_min.set(off_m))
            self.after(0, self._redraw)
            self.after(0, lambda: self.status_var.set(
                f"自动对齐完成，S1 偏移 {off_m:+.1f} 分钟"))
        except Exception as e:
            self.after(0, lambda: self.status_var.set(f"自动对齐失败：{e}"))

    # ── 数据提取 ──────────────────────────────────────────────────────────────

    def _series(self, df, col: str, offset_sec: float = 0.0):
        if df is None or col in ("", "无") or col not in df.columns:
            return None
        s = df[col].dropna()
        if not isinstance(s.index, pd.DatetimeIndex):
            return None
        if offset_sec:
            s = s.copy()
            s.index = s.index + pd.Timedelta(seconds=offset_sec)
        return smooth(s, self.smooth_win.get())

    # ── 绘图 ──────────────────────────────────────────────────────────────────

    def _redraw(self):
        self.ax1.cla()
        self.ax2.cla()
        off = self.offset_min.get() * 60
        lines, labels = [], []

        def plot_line(s, color, label):
            if s is None or s.empty:
                return
            ln, = self.ax1.plot(s.index, s.values, color=color, lw=1.5,
                                zorder=3, label=label)
            lines.append(ln); labels.append(label)

        # 先提取原始三条曲线
        # 数据源 A：相对偏移（对齐用）+ 绝对平移
        # 数据源 B：仅绝对平移
        off_a = off + self.shift_a_min.get() * 60
        off_b = self.shift_b_min.get() * 60
        s_s1 = self._series(self.df_s1, self.col_s1.get(), off_a)
        s_pv = self._series(self.df_box, self.col_pv.get(), off_b)
        s_sv = self._series(self.df_box, self.col_sv.get(), off_b)

        # 参考曲线拟合
        ref_key = self.ref_curve.get()
        if ref_key != "无":
            ref_map = {"S1": s_s1, "PV": s_pv, "SV": s_sv}
            ref_s   = ref_map.get(ref_key)
            sp_s    = self.spline_s.get()
            rw      = self.resid_win.get()
            if ref_s is not None and not ref_s.empty:
                try:
                    if ref_key != "S1" and s_s1 is not None:
                        s_s1 = fit_to_ref(ref_s, s_s1, sp_s, rw)
                    if ref_key != "PV" and s_pv is not None:
                        s_pv = fit_to_ref(ref_s, s_pv, sp_s, rw)
                    if ref_key != "SV" and s_sv is not None:
                        s_sv = fit_to_ref(ref_s, s_sv, sp_s, rw)
                except Exception as ex:
                    self.status_var.set(f"拟合失败：{ex}")

        if self.show_s1.get():
            plot_line(s_s1, COLOR_S1, "Temperature_S1")
        if self.show_pv.get():
            plot_line(s_pv, COLOR_PV, "Temperature_PV")
        if self.show_sv.get():
            plot_line(s_sv, COLOR_SV, "Temperature_SV")

        # 湿度右轴
        if self.show_hum.get():
            sh = self._series(self.df_box, self.col_hum.get(), off_b)
            if sh is not None and not sh.empty:
                ln, = self.ax2.plot(sh.index, sh.values,
                                    color=COLOR_HUM, lw=1.2, ls="--",
                                    zorder=2, label="湿度")
                self.ax2.set_ylabel("湿度 (%)", color=COLOR_HUM)
                self.ax2.tick_params(axis="y", labelcolor=COLOR_HUM)
                self.ax2.set_ylim(0, 110)
                self.ax2.spines["right"].set_visible(True)
                self.ax2.spines["right"].set_color("#CCCCCC")
                lines.append(ln); labels.append("湿度")
        else:
            self.ax2.set_yticks([])
            self.ax2.spines["right"].set_visible(False)

        # 偏差标注
        if self.show_dev.get():
            msg = self._draw_deviation(off)
            if msg:
                self.status_var.set(msg)

        # 格式
        self.ax1.set_ylabel("温度 (°C)")
        self.ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d %H:%M"))

        # Y 轴最小跨度：防止曲线接近时坐标轴放大失真
        if lines:
            ymin, ymax = self.ax1.get_ylim()
            span = self.y_min_span.get()
            if (ymax - ymin) < span:
                mid = (ymax + ymin) / 2
                self.ax1.set_ylim(mid - span / 2, mid + span / 2)

        # X 轴刻度：按时间跨度自适应间隔，并强制标注首尾时间点
        if lines:
            xmin, xmax = self.ax1.get_xlim()
            t_start = mdates.num2date(xmin).replace(tzinfo=None)
            t_end   = mdates.num2date(xmax).replace(tzinfo=None)
            span_h  = (t_end - t_start).total_seconds() / 3600
            if span_h > 7 * 24:
                freq = "24h"
            elif span_h > 3 * 24:
                freq = "12h"
            elif span_h > 24:
                freq = "6h"
            elif span_h > 12:
                freq = "3h"
            else:
                freq = "1h"
            reg_ticks = pd.date_range(
                start=pd.Timestamp(t_start).ceil(freq),
                end=pd.Timestamp(t_end).floor(freq),
                freq=freq)
            self.ax1.xaxis.set_major_locator(
                mticker.FixedLocator([mdates.date2num(t) for t in reg_ticks]))

        self.fig.autofmt_xdate(rotation=30, ha="right")
        self.ax1.set_title("温湿度数据对比",
                           loc="right",
                           fontsize=13, fontweight="bold",
                           color="#333333", pad=10)

        if lines:
            self.ax1.legend(lines, labels,
                            loc="upper left", ncol=len(lines),
                            bbox_to_anchor=(0.0, 1.06),
                            frameon=False, fontsize=9)

        self.canvas.draw_idle()

    def _draw_deviation(self, offset_sec: float):
        s1 = self._series(self.df_s1, self.col_s1.get(), offset_sec)
        pv = self._series(self.df_box, self.col_pv.get())
        if s1 is None or pv is None:
            return None
        thresh = self.dev_thresh.get()
        merged = pd.DataFrame({"s1": s1, "pv": pv}).dropna()
        if merged.empty:
            return None
        exceed = (merged["s1"] - merged["pv"]).abs() > thresh
        if not exceed.any():
            return None
        # 画连续超限区段
        in_seg, t0 = False, None
        for t, v in exceed.items():
            if v and not in_seg:
                in_seg, t0 = True, t
            elif not v and in_seg:
                in_seg = False
                self.ax1.axvspan(t0, t, alpha=0.12, color="red", zorder=1)
        if in_seg:
            self.ax1.axvspan(t0, exceed.index[-1], alpha=0.12, color="red", zorder=1)
        n = int(exceed.sum())
        return f"偏差超限（>{thresh}°C）：{n}/{len(exceed)} 个点"

    # ── 导出 ──────────────────────────────────────────────────────────────────

    def _export_png(self):
        p = filedialog.asksaveasfilename(
            title="保存图表", defaultextension=".png",
            initialdir=self._first_dir or None,
            filetypes=[("PNG 图片", "*.png"), ("SVG 矢量", "*.svg")])
        if not p:
            return
        # Windows 有时不自动加后缀
        if not (p.endswith(".png") or p.endswith(".svg")):
            p += ".png"
        try:
            self.fig.savefig(p, dpi=150, bbox_inches="tight")
            self.status_var.set(f"图表已保存：{Path(p).name}")
        except Exception as e:
            messagebox.showerror("保存失败", str(e))

    def _export_excel(self):
        if self.df_s1 is None and self.df_box is None:
            messagebox.showwarning("提示", "没有可导出的数据。")
            return
        p = filedialog.asksaveasfilename(
            title="导出数据", defaultextension=".xlsx",
            initialdir=self._first_dir or None,
            filetypes=[("Excel 文件", "*.xlsx")])
        if not p:
            return
        off = self.offset_min.get() * 60
        parts = {}
        for label, df, col, o in [
            ("Temperature_S1", self.df_s1, self.col_s1.get(), off),
            ("Temperature_PV", self.df_box, self.col_pv.get(), 0),
            ("Temperature_SV", self.df_box, self.col_sv.get(), 0),
        ]:
            s = self._series(df, col, o)
            if s is not None:
                parts[label] = s
        if not parts:
            messagebox.showwarning("提示", "所选列无有效数据。")
            return
        out = pd.DataFrame(parts)
        out.index.name = "时间"
        out.to_excel(p)
        self.status_var.set(f"数据已导出：{Path(p).name}  |  {len(out)} 行")


# ── 辅助函数 ──────────────────────────────────────────────────────────────────

def _btn(parent, text, cmd, bg, full=False):
    b = tk.Button(parent, text=text, command=cmd, bg=bg, fg="white",
                  relief="flat", activebackground=bg,
                  font=("Microsoft YaHei", 9))
    if full:
        b.pack(fill="x", pady=2)
    return b


def _set_cb(cb: ttk.Combobox, values: list, current: str):
    cb["values"] = values
    cb.set(current if current in values else (values[0] if values else ""))


# ── 入口 ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
