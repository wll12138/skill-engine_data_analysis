"""
engine_analysis.py — 发动机台架性能数据分析工具集

分析增压器匹配对比、扭矩/BSFC/增压压力、高原能力评估、数据可视化。
所有函数接受 DataFrame 和列名参数，带自动列名检测和错误处理。

用法示例：
    from engine_analysis import *
    df = load_excel("data.xlsx")
    df_a, df_b = split_groups(df, n_points=9)
    results = compare_turbochargers(df_a, df_b)
    plot_comparison(results)
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path
from typing import Optional, Dict, Tuple, List

# ────────────────────────────────────────────────────────────
# 常量
# ────────────────────────────────────────────────────────────

# 标准大气压 (kPa)
ATM_STANDARD = 101.325

# 各海拔大气压 (基于 ISO 2533 标准大气)
ALTITUDE_PRESSURE = {
    0:     101.325,
    1000:   89.9,
    2000:   79.5,
    3000:   70.1,
    4000:   61.6,
}

# 默认增压器转速限制 (rpm) — 需用户根据供应商确认
TURBO_SPEED_LIMIT_DEFAULT = 250000

# 默认转速点
DEFAULT_RPM_POINTS = [1000, 1500, 2000, 2500, 3000, 3500, 4000, 4500, 5000]

# 列名关键词映射 (检测顺序 = 优先匹配顺序)
COLUMN_PATTERNS = {
    "rpm":     ["DynoSpeed_Avg", "DynoSpeed", "转速", "rpm", "RPM", "SPEED", "EngineSpeed", "Epm_nEng"],
    "torque":  ["DynoTorque_Avg", "DynoTorque", "修正扭矩", "CorrTorque", "CorrTorqueEWG", "扭矩", "Torque", "TORQUE"],
    "bsfc":    ["BSFC", "燃油消耗率", "FuelCOSP", "FB_RATE"],
    "turbo_speed": ["TURBOSPEED_Avg", "TURBOSPEED", "增压器转速",
                    "Trbch_N", "TurboSpeed", "涡轮转速"],
    "boost":   ["BSTC_pActBoostPress_Avg", "BSTC_pActBoostPress",
                "增压压力", "Boost", "BOOST", "VBOOST", "P3", "P_Intake"],
    "egt":     ["EXHT_tMnfdTemp_Avg", "EXHT_tMnfdTemp", "排气温度",
                "EGT", "egt", "exhaust", "T_EXH", "MEANTEXH"],
    "backpressure": ["ExhP_pUpFstCat_Avg", "ExhP_pUpFstCat",
                     "背压", "Back", "back", "FT_TACT", "P_EXH"],
    "wg":      ["EWGC_rActlPos_Avg", "EWGC_rActlPos", "WG开度",
                "WG", "wg", "wastegate"],
    "airflow": ["AirFlow_Avg", "AirFlow", "进气流量", "AFS_dm", "air", "流量"],
    "power":   ["BrakePower_Avg", "BrakePower", "修正功率",
                "CorrBrkPwr", "CorrBrkPwrEWG", "功率", "Power", "POWER"],
    "intake_temp": ["AirTemp_Avg", "AirTemp", "进气温度",
                    "进气歧管温度", "T_Intake", "T_AIR", "T_AIR_IN", "T_ACS"],
    # 燃烧特性相关
    "cov":       ["IMEP1CO_Avg", "IMEP1CO", "COV", "cov",
                  "IMEPCOV", "CoV", "循环变动"],
    "ai50":      ["AI501_Avg", "AI501", "AI50", "CA50", "MFB50",
                  "A50", "a50", "燃烧相位"],
    "spark_act": ["SPK_dgActSpkAdv_Avg", "SPK_dgActSpkAdvAvg_Avg",
                  "SPK_dgActSpkAdv", "点火角", "点火提前角",
                  "SparkAdv", "SPK_dgMainSpkAdv"],
    "spark_mbt": ["SPK_dgMBTSpkAdv_Avg", "SPK_dgMBTSpkAdvAvg_Avg",
                  "SPK_dgMBTSpkAdv", "MBTSpkAdv"],
    "spark_delta": ["SPK_dgDltFromMBT_Avg", "SPK_dgDltFromMBT",
                    "DltFromMBT", "dltFromMBT", "退角", "点火退角"],
    "knock":     ["knockWndStrAng_Avg", "knockWndStrAng", "Knock",
                  "KNK", "knock", "爆震", "knockWnd"],
    "vvt":       ["VVT", "vvt", "VCT", "Cam", "cam", "进气门", "排气门",
                  "CamPhs", "CamPos"],
    "fuel_flow": ["Fuel_FuelConsume_Avg", "Fuel_FuelConsume",
                  "FuelConsume", "FuelMassFlow", "燃油消耗量",
                  "FuelFlow", "油耗量"],
    "imep":      ["IMEP1_Avg", "IMEPH1_Avg", "IMEPL1_Avg",
                  "Pmi", "平均有效压力"],
}

# ────────────────────────────────────────────────────────────
# 1. 数据加载与清洗
# ────────────────────────────────────────────────────────────

def load_excel(filepath: str, sheet_name: Optional[str] = None,
               skiprows: int = 0) -> pd.DataFrame:
    """读取 Excel 台架数据，支持自动探测 sheet 名。
    如果是 .csv 文件，自动转调 load_csv()。

    Args:
        filepath: Excel 文件路径
        sheet_name: sheet 名称，为 None 时打印所有可用 sheet
        skiprows: 跳过前 N 行 (如单位行)

    Returns:
        DataFrame

    用法:
        df = load_excel("data.xlsx", sheet_name="Sheet3", skiprows=1)
    """
    p = Path(filepath)
    if not p.exists():
        raise FileNotFoundError(f"文件不存在: {filepath}")

    ext = p.suffix.lower()
    if ext == '.csv':
        print("检测到 CSV 文件，自动使用 load_csv()")
        return load_csv(filepath, encoding='gbk', header_rows=0)

    xl = pd.ExcelFile(p)
    print(f"可用 sheet: {xl.sheet_names}")

    if sheet_name is None and len(xl.sheet_names) == 1:
        sheet_name = xl.sheet_names[0]
        print(f"使用唯一 sheet: {sheet_name}")
    elif sheet_name is None:
        sheet_name = xl.sheet_names[0]
        print(f"默认使用第一个 sheet: {sheet_name}")

    df = pd.read_excel(xl, sheet_name=sheet_name, skiprows=skiprows)
    print(f"读取完成: {df.shape[0]} 行 x {df.shape[1]} 列")
    return df


def load_csv(filepath: str, encoding: str = 'gbk', header_rows: int = 5,
             skip_time_cols: int = 3) -> pd.DataFrame:
    """读取台架 CSV 数据 (复杂多行表头, GBK 编码)。

    典型 ETAS INCA 输出的 CSV 有 5~8 行元数据/表头信息：
      rows 0-4: 元数据 (Logger description, Log period...)
      row 5:    真正的列名 (Time, DynoSpeed_Avg, ...)
      row 6:    单位行 (ms, rev/min, Nm, ...)
      row 7:    数据类型行 (Raw, Average, ...)
      row 8+:   实际数据

    Args:
        filepath: CSV 文件路径
        encoding: 文件编码 (默认 gbk)
        header_rows: 跳过的元数据行数 (默认 5)
        skip_time_cols: 跳过前 N 列时间戳列 (默认 3)

    Returns:
        清洗后的 DataFrame
    """
    p = Path(filepath)
    if not p.exists():
        raise FileNotFoundError(f"文件不存在: {filepath}")

    print(f"CSV: encoding={encoding}, header_rows={header_rows}")

    # 先读列名行
    col_df = pd.read_csv(p, encoding=encoding, header=None, nrows=1,
                         skiprows=header_rows)
    col_names = col_df.iloc[0].tolist()

    # 用列名读数据
    df = pd.read_csv(p, encoding=encoding, header=header_rows)
    df.columns = col_names[:len(df.columns)]
    df = clean_columns(df)

    # 跳过时间戳列
    if skip_time_cols > 0:
        time_cols = df.columns[:skip_time_cols].tolist()
        df = df.drop(columns=time_cols)
        print(f"跳过时间戳列: {time_cols}")

    print(f"读取完成: {df.shape[0]} 行 x {df.shape[1]} 列")
    return df


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """清理列名: 去除换行符、首尾空格、统一命名。"""
    df = df.copy()
    df.columns = (
        df.columns
        .str.replace('\n', '', regex=False)
        .str.replace('\r', '', regex=False)
        .str.strip()
    )
    return df


def ensure_numeric(df: pd.DataFrame,
                   exclude_cols: Optional[List[str]] = None) -> pd.DataFrame:
    """将 DataFrame 中的 object 列转为数值类型 (无法转换的置为 NaN)。"""
    df = df.copy()
    ex = set(exclude_cols or [])
    for col in df.columns:
        if col in ex:
            continue
        if df[col].dtype == 'object':
            df[col] = pd.to_numeric(df[col], errors='coerce')
    return df


def inspect_data(df: pd.DataFrame, head_n: int = 20) -> None:
    """打印数据概览: 列名、前若干行、数据类型。"""
    print(f"📐 列名 ({len(df.columns)}): {df.columns.tolist()}")
    print(f"\n📄 前 {head_n} 行:")
    print(df.head(head_n).to_string())
    print(f"\n🔢 数据类型:")
    print(df.dtypes)


# ────────────────────────────────────────────────────────────
# 2. 列名检测
# ────────────────────────────────────────────────────────────

def detect_column(df: pd.DataFrame, key: str,
                  case_sensitive: bool = False) -> Optional[str]:
    """在 DataFrame 中根据关键词模式检测目标列名 (模糊匹配)。

    Args:
        df: 目标 DataFrame
        key: COLUMN_PATTERNS 中的键名 (如 'rpm', 'torque', 'bsfc')
        case_sensitive: 是否大小写敏感

    Returns:
        匹配到的列名，或 None

    示例:
        rpm_col = detect_column(df, "rpm")  # 自动找到转速列
    """
    if key not in COLUMN_PATTERNS:
        raise ValueError(f"未知列类型 '{key}'，可用: {list(COLUMN_PATTERNS.keys())}")

    patterns = COLUMN_PATTERNS[key]
    columns = df.columns.tolist()

    if not case_sensitive:
        cols_lower = {c: c.lower() for c in columns}
        patterns = [p.lower() for p in patterns]
    else:
        cols_lower = {c: c for c in columns}

    # 优先完全匹配
    for col in columns:
        col_check = col if case_sensitive else col.lower()
        if col_check in patterns:
            return col

    # 再部分匹配 (包含关键词)
    for col in columns:
        col_check = col if case_sensitive else col.lower()
        for pat in patterns:
            if pat in col_check or col_check in pat:
                return col

    # 最后子串匹配 (separator 分割后匹配)
    for col in columns:
        col_check = col if case_sensitive else col.lower()
        parts = set(col_check.replace('_', ' ').replace('.', ' ').split())
        for pat in patterns:
            pat_parts = set(pat.replace('_', ' ').replace('.', ' ').split())
            if pat_parts & parts:
                return col
            # 单项匹配 (长度≥3 避免单字母误配, 如 'p' in '(rpm)')
            for pp in pat_parts:
                if len(pp) < 3:
                    continue
                for cp in parts:
                    if len(cp) < 3:
                        continue
                    if pp in cp or cp in pp:
                        return col

    return None


def detect_all_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """检测 DataFrame 中所有已知类型的列。

    Returns:
        {类型键: 列名} 字典

    示例:
        cols = detect_all_columns(df)
        # cols = {'rpm': '转速', 'torque': '扭矩(Bm)', 'bsfc': None, ...}
    """
    result = {}
    for key in COLUMN_PATTERNS:
        col = detect_column(df, key)
        if col is not None:
            result[key] = col
    return result


# ────────────────────────────────────────────────────────────
# 3. 数据分隔 (A/B 两组上下排列的情况)
# ────────────────────────────────────────────────────────────

def split_groups(df: pd.DataFrame, n_points: Optional[int] = None,
                 rpm_col: Optional[str] = None) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """将上下排列的两组数据按行分为 A、B 两组。

    常见场景: Excel 中前 N 行是增压器 A，后 N 行是增压器 B。

    Args:
        df: 原始 DataFrame
        n_points: 每组行数。为 None 时尝试按 DEFAULT_RPM_POINTS 推断
        rpm_col: 转速列名，用于辅助推断

    Returns:
        (df_a, df_b)

    用法:
        df_a, df_b = split_groups(df, n_points=9)
        df_a, df_b = split_groups(df, rpm_col="转速")  # 自动推断
    """
    if n_points is not None:
        m = len(df) // 2
        df_a = df.iloc[:n_points].copy().reset_index(drop=True)
        df_b = df.iloc[n_points:n_points * 2].copy().reset_index(drop=True)
        print(f"→ A 组: {len(df_a)} 行, B 组: {len(df_b)} 行")
        return df_a, df_b

    # 尝试自动推断 — 假设数据行数是标准转速点的 2 倍
    total = len(df)
    if total % 2 == 0:
        half = total // 2
        if rpm_col is not None and rpm_col in df.columns:
            # 看换行位置是否是标准 rpm 点的结束
            test = df[rpm_col].values[:half]
            # 如果前半部分比较连续/完整
            pass
        df_a = df.iloc[:half].copy().reset_index(drop=True)
        df_b = df.iloc[half:].copy().reset_index(drop=True)
        print(f"→ 自动推断: A 组 {len(df_a)} 行, B 组 {len(df_b)} 行")
        return df_a, df_b

    raise ValueError(
        f"无法自动推断分组行数 (数据总行数 {total} 不是偶数)。"
        f"请指定 n_points 参数。"
    )


# ────────────────────────────────────────────────────────────
# 4. 增压器对比分析
# ────────────────────────────────────────────────────────────

def _safe_float(arr):
    """安全转为 float ndarray，non-numeric 置 NaN。"""
    return pd.to_numeric(pd.Series(arr), errors='coerce').values.astype(float)


def compare_turbochargers(df_a: pd.DataFrame, df_b: pd.DataFrame,
                          name_a: str = "A", name_b: str = "B",
                          rpm_col: Optional[str] = None,
                          torque_col: Optional[str] = None,
                          turbo_speed_limit: int = TURBO_SPEED_LIMIT_DEFAULT
                          ) -> Dict:
    """双增压器全面对比分析。

    Args:
        df_a: 增压器 A 的数据
        df_b: 增压器 B 的数据
        name_a: 增压器 A 名称
        name_b: 增压器 B 名称
        rpm_col: 转速列名；为 None 时自动检测
        torque_col: 扭矩列名；为 None 时自动检测
        turbo_speed_limit: 增压器转速限制值 (rpm)，默认 250000

    Returns:
        Dict 包含:
          - summary: 关键指标汇总
          - scores: 加权评分
          - details: 各维度详细数据
          - col_map: 列名映射
          - names: (name_a, name_b)

    用法:
        results = compare_turbochargers(df_a, df_b, "方案A", "方案B")
        print(results["summary"])
        print(results["scores"])
    """
    # --- 列名检测 ---
    rpm_col = rpm_col or detect_column(df_a, "rpm")
    torque_col = torque_col or detect_column(df_a, "torque")

    if rpm_col is None:
        raise ValueError("无法检测到转速列，请通过 rpm_col= 指定")
    if torque_col is None:
        raise ValueError("无法检测到扭矩列，请通过 torque_col= 指定")

    col_map = detect_all_columns(df_a)
    col_map["rpm"] = rpm_col
    col_map["torque"] = torque_col

    # --- 安全提取数值 ---
    rpm = _safe_float(df_a[rpm_col].values)
    torque_a = _safe_float(df_a[torque_col].values)
    torque_b = _safe_float(df_b[torque_col].values)

    print(f"📊 列名映射: {col_map}")

    # ── 1. 扭矩分析 ──
    low_end = rpm <= 1500
    tq_low_a = np.nanmean(torque_a[low_end])
    tq_low_b = np.nanmean(torque_b[low_end])
    tq_peak_a = np.nanmax(torque_a)
    tq_peak_b = np.nanmax(torque_b)
    tq_avg_a = np.nanmean(torque_a)
    tq_avg_b = np.nanmean(torque_b)
    tq_peak_rpm_a = rpm[np.nanargmax(torque_a)] if not np.all(np.isnan(torque_a)) else 0
    tq_peak_rpm_b = rpm[np.nanargmax(torque_b)] if not np.all(np.isnan(torque_b)) else 0

    # ── 2. BSFC ──
    bsfc_a = bsfc_b = bsfc_main_a = bsfc_main_b = None
    if "bsfc" in col_map:
        bsfc_a = _safe_float(df_a[col_map["bsfc"]].values)
        bsfc_b = _safe_float(df_b[col_map["bsfc"]].values)
        # 排除 1000rpm (异常偏高)
        bsfc_main = rpm > 1000
        bsfc_main_a = np.nanmean(bsfc_a[bsfc_main])
        bsfc_main_b = np.nanmean(bsfc_b[bsfc_main])

    # ── 3. 增压器转速 ──
    speed_a = speed_b = speed_margin_a = speed_margin_b = None
    if "turbo_speed" in col_map:
        speed_a = _safe_float(df_a[col_map["turbo_speed"]].values)
        speed_b = _safe_float(df_b[col_map["turbo_speed"]].values)
        speed_margin_a = turbo_speed_limit - np.nanmax(speed_a)
        speed_margin_b = turbo_speed_limit - np.nanmax(speed_b)

    # ── 4. 排气温度 ──
    egt_a = egt_b = egt_max_a = egt_max_b = egt_avg_a = egt_avg_b = None
    if "egt" in col_map:
        egt_a = _safe_float(df_a[col_map["egt"]].values)
        egt_b = _safe_float(df_b[col_map["egt"]].values)
        egt_max_a = np.nanmax(egt_a)
        egt_max_b = np.nanmax(egt_b)
        egt_avg_a = np.nanmean(egt_a)
        egt_avg_b = np.nanmean(egt_b)

    # ── 5. 增压压力 ──
    boost_a = boost_b = None
    if "boost" in col_map:
        boost_a = _safe_float(df_a[col_map["boost"]].values)
        boost_b = _safe_float(df_b[col_map["boost"]].values)

    # ── 6. WG 开度 ──
    wg_a = wg_b = wg_high_a = wg_high_b = None
    if "wg" in col_map:
        wg_a = _safe_float(df_a[col_map["wg"]].values)
        wg_b = _safe_float(df_b[col_map["wg"]].values)
        high_rpm = rpm >= 3000
        wg_high_a = np.nanmean(wg_a[high_rpm])
        wg_high_b = np.nanmean(wg_b[high_rpm])

    # ── 7. 背压 ──
    bp_a = bp_b = None
    if "backpressure" in col_map:
        bp_a = _safe_float(df_a[col_map["backpressure"]].values)
        bp_b = _safe_float(df_b[col_map["backpressure"]].values)

    # ── 8. 功率 ──
    power_a = power_b = None
    if "power" in col_map:
        power_a = _safe_float(df_a[col_map["power"]].values)
        power_b = _safe_float(df_b[col_map["power"]].values)

    # ── 评分 ──
    scores = {name_a: 0.0, name_b: 0.0}
    weights = {}

    # 低速扭矩 (权重 2.0)
    weights["低速扭矩"] = (2.0, name_a if tq_low_a > tq_low_b else name_b)
    if tq_low_a > tq_low_b:
        scores[name_a] += 2.0
    else:
        scores[name_b] += 2.0

    # BSFC — 越低越好 (权重 2.0)
    if bsfc_main_a is not None:
        if bsfc_main_a < bsfc_main_b:
            scores[name_a] += 2.0
            weights["燃油经济性"] = (2.0, name_a)
        elif bsfc_main_b < bsfc_main_a:
            scores[name_b] += 2.0
            weights["燃油经济性"] = (2.0, name_b)
        else:
            weights["燃油经济性"] = (2.0, "持平")

    # 涡轮转速 — 高速段 (≥4000rpm) 越低越好 (余量大) (权重 1.5)
    if speed_a is not None:
        high_end = rpm >= 4000
        spd_high_a = np.nanmean(speed_a[high_end])
        spd_high_b = np.nanmean(speed_b[high_end])
        if spd_high_a < spd_high_b:
            scores[name_a] += 1.5
            weights["涡轮转速余量"] = (1.5, name_a)
        elif spd_high_b < spd_high_a:
            scores[name_b] += 1.5
            weights["涡轮转速余量"] = (1.5, name_b)
        else:
            weights["涡轮转速余量"] = (1.5, "持平")

    # 排气温度 — 越低越好 (权重 1.0)
    if egt_avg_a is not None:
        if egt_avg_a < egt_avg_b:
            scores[name_a] += 1.0
            weights["排气温度"] = (1.0, name_a)
        elif egt_avg_b < egt_avg_a:
            scores[name_b] += 1.0
            weights["排气温度"] = (1.0, name_b)
        else:
            weights["排气温度"] = (1.0, "持平")

    # WG 开度 — 高转速段越小越好 (权重 1.5)
    if wg_high_a is not None:
        if wg_high_a < wg_high_b:
            scores[name_a] += 1.5
            weights["WG效率"] = (1.5, name_a)
        elif wg_high_b < wg_high_a:
            scores[name_b] += 1.5
            weights["WG效率"] = (1.5, name_b)
        else:
            weights["WG效率"] = (1.5, "持平")

    # 峰值扭矩 (权重 1.0)
    if tq_peak_a > tq_peak_b:
        scores[name_a] += 1.0
        weights["峰值扭矩"] = (1.0, name_a)
    elif tq_peak_b > tq_peak_a:
        scores[name_b] += 1.0
        weights["峰值扭矩"] = (1.0, name_b)
    else:
        weights["峰值扭矩"] = (1.0, "持平")

    winner = name_a if scores[name_a] > scores[name_b] else name_b

    # ── 打包结果 ──
    return {
        "names": (name_a, name_b),
        "winner": winner,
        "scores": scores,
        "weights": weights,
        "col_map": col_map,
        "rpm": rpm,
        "summary": {
            "torque_low":      (round(tq_low_a, 1), round(tq_low_b, 1)),
            "torque_peak":     (round(tq_peak_a, 1), round(tq_peak_b, 1)),
            "torque_peak_rpm": (int(tq_peak_rpm_a), int(tq_peak_rpm_b)),
            "torque_avg":      (round(tq_avg_a, 1), round(tq_avg_b, 1)),
            "bsfc_avg":        (round(bsfc_main_a, 1) if bsfc_main_a else None,
                                round(bsfc_main_b, 1) if bsfc_main_b else None),
            "speed_margin":    (round(speed_margin_a) if speed_margin_a else None,
                                round(speed_margin_b) if speed_margin_b else None),
            "egt_max":         (round(egt_max_a) if egt_max_a else None,
                                round(egt_max_b) if egt_max_b else None),
            "egt_avg":         (round(egt_avg_a, 1) if egt_avg_a else None,
                                round(egt_avg_b, 1) if egt_avg_b else None),
            "wg_high_avg":     (round(wg_high_a, 1) if wg_high_a else None,
                                round(wg_high_b, 1) if wg_high_b else None),
            "turbo_speed_limit": turbo_speed_limit,
        },
        "raw": {
            "torque_a": torque_a, "torque_b": torque_b,
            "bsfc_a": bsfc_a, "bsfc_b": bsfc_b,
            "speed_a": speed_a, "speed_b": speed_b,
            "egt_a": egt_a, "egt_b": egt_b,
            "boost_a": boost_a, "boost_b": boost_b,
            "wg_a": wg_a, "wg_b": wg_b,
            "bp_a": bp_a, "bp_b": bp_b,
            "power_a": power_a, "power_b": power_b,
        },
    }


# ────────────────────────────────────────────────────────────
# 5. 高原能力评估
# ────────────────────────────────────────────────────────────

def calc_altitude_pressure(altitude_m: float) -> float:
    """计算给定海拔的标准大气压 (kPa)。

    使用 ISO 2533 标准大气模型:
        P = 101.325 × (1 - 0.0065 × h / 288.15) ^ 5.255
    """
    return ATM_STANDARD * (1 - 0.0065 * altitude_m / 288.15) ** 5.255


def assess_high_altitude(df_a: pd.DataFrame, df_b: pd.DataFrame,
                         name_a: str = "A", name_b: str = "B",
                         altitude_m: float = 3000,
                         turbo_speed_limit: int = TURBO_SPEED_LIMIT_DEFAULT,
                         rpm_col: Optional[str] = None
                         ) -> Dict:
    """评估两款增压器在指定海拔的高原性能。

    关键公式:
        Speed_alt = Speed_0 × √(P0 / P_alt)
    其中 P0 = 101.325 kPa, P_alt 为对应海拔的大气压。

    Args:
        df_a / df_b: 两组增压器数据
        name_a / name_b: 名称
        altitude_m: 海拔高度 (米)
        turbo_speed_limit: 增压器转速限制值
        rpm_col: 转速列名

    Returns:
        Dict 包含高原评估结果

    用法:
        ha = assess_high_altitude(df_a, df_b, "方案A", "方案B",
                                  altitude_m=3000)
    """
    rpm_col = rpm_col or detect_column(df_a, "rpm")
    if rpm_col is None:
        raise ValueError("无法检测到转速列")

    speed_col = detect_column(df_a, "turbo_speed")
    if speed_col is None:
        return {"error": "无法检测增压器转速列，无法评估高原能力"}

    rpm = _safe_float(df_a[rpm_col].values)
    speed_a = _safe_float(df_a[speed_col].values)
    speed_b = _safe_float(df_b[speed_col].values)

    P_alt = calc_altitude_pressure(altitude_m)
    ratio = np.sqrt(ATM_STANDARD / P_alt)

    speed_alt_a = speed_a * ratio
    speed_alt_b = speed_b * ratio

    max_a = np.nanmax(speed_alt_a)
    max_b = np.nanmax(speed_alt_b)
    margin_a = turbo_speed_limit - max_a
    margin_b = turbo_speed_limit - max_b

    def safety_label(margin):
        if margin > 30000:
            return "✅ 安全"
        elif margin > 15000:
            return "⚠️ 可接受"
        else:
            return "❌ 高风险"

    # 增压比分析 (如数据存在)
    boost_col = detect_column(df_a, "boost")
    pr_a = pr_b = None
    if boost_col is not None:
        boost_a = _safe_float(df_a[boost_col].values)
        boost_b = _safe_float(df_b[boost_col].values)
        pr_a = boost_a / P_alt
        pr_b = boost_b / P_alt

    return {
        "altitude_m": altitude_m,
        "P_alt_kPa": round(P_alt, 2),
        "P0_kPa": ATM_STANDARD,
        "speed_multiplier": round(ratio, 4),
        turbo_speed_limit: turbo_speed_limit,
        name_a: {
            "max_speed_alt": round(max_a),
            "margin": round(margin_a),
            "safety": safety_label(margin_a),
            "pressure_ratio_range": (
                (round(float(pr_a.min()), 3), round(float(pr_a.max()), 3))
                if pr_a is not None else None
            ),
        },
        name_b: {
            "max_speed_alt": round(max_b),
            "margin": round(margin_b),
            "safety": safety_label(margin_b),
            "pressure_ratio_range": (
                (round(float(pr_b.min()), 3), round(float(pr_b.max()), 3))
                if pr_b is not None else None
            ),
        },
    }


# ────────────────────────────────────────────────────────────
# 6. 数据可视化
# ────────────────────────────────────────────────────────────

def _setup_chinese_font():
    """设置 matplotlib 中文字体。macOS 优先用 PingFang。"""
    for font in ["PingFang SC", "Heiti TC", "Arial Unicode MS",
                  "Noto Sans CJK SC", "SimHei"]:
        try:
            plt.rcParams['font.family'] = [font, 'sans-serif']
            plt.rcParams['axes.unicode_minus'] = False
            fig, ax = plt.subplots()
            ax.set_title("测试")
            plt.close(fig)
            return
        except Exception:
            continue
    # fallback
    plt.rcParams['font.family'] = ['sans-serif']
    plt.rcParams['axes.unicode_minus'] = False


def plot_comparison(results: Dict, save_path: Optional[str] = None,
                    figsize: Tuple[int, int] = (16, 10)):
    """生成增压器对比图: 扭矩 / BSFC / 增压器转速 / EGT / 增压压力 / WG 开度。

    Args:
        results: compare_turbochargers() 返回值
        save_path: 图片保存路径，为 None 则显示
        figsize: 画布大小

    用法:
        plot_comparison(results, "/tmp/comparison.png")
    """
    _setup_chinese_font()

    name_a, name_b = results["names"]
    rpm = results["rpm"]
    r = results["raw"]
    sm = results["summary"]

    # 确定子图数量
    subplots = [1]  # 扭矩总是有
    labels = {1: "扭矩"}
    idx = 2
    for key, label in [("bsfc_a", "BSFC"), ("speed_a", "增压器转速"),
                       ("egt_a", "排气温度"), ("boost_a", "增压压力"),
                       ("wg_a", "WG开度")]:
        if r.get(key) is not None:
            subplots.append(idx)
            labels[idx] = label
            idx += 1

    n_plots = len(subplots)
    n_rows = (n_plots + 2) // 3  # 最多3列
    n_cols = min(n_plots, 3)

    fig, axes = plt.subplots(n_rows, n_cols, figsize=figsize)
    axes_flat = np.atleast_1d(axes).ravel()

    # 隐藏多余子图
    for i in range(n_plots, len(axes_flat)):
        axes_flat[i].set_visible(False)

    def _plot_ax(ax, y_a, y_b, title, ylabel="", ylim=None,
                 show_limit=None):
        if y_a is None or y_b is None:
            return
        ax.plot(rpm, y_a, 'o-', label=name_a, linewidth=1.5)
        ax.plot(rpm, y_b, 's-', label=name_b, linewidth=1.5)
        ax.set_title(title)
        ax.set_xlabel("转速 (rpm)")
        ax.set_ylabel(ylabel)
        if ylim:
            ax.set_ylim(ylim)
        if show_limit:
            ax.axhline(y=show_limit[0], color='r', linestyle='--',
                       alpha=0.5, label=show_limit[1])
        ax.legend(fontsize=9)
        ax.grid(True, alpha=0.3)
        ax.ticklabel_format(style='plain', axis='y')

    # 1. 扭矩
    _plot_ax(axes_flat[0], r["torque_a"], r["torque_b"],
             f"扭矩对比 (Nm)", "Nm")

    # 2. BSFC
    if r["bsfc_a"] is not None:
        _plot_ax(axes_flat[1], r["bsfc_a"], r["bsfc_b"],
                 "BSFC 对比 (g/kWh)", "g/kWh")

    # 3. 增压器转速
    if r["speed_a"] is not None:
        _plot_ax(axes_flat[2], r["speed_a"], r["speed_b"],
                 "增压器转速 (rpm)", "rpm",
                 show_limit=(sm.get("turbo_speed_limit", 250000),
                             f"限制 {sm.get('turbo_speed_limit', 250000)} rpm"))

    # 4. EGT
    if r["egt_a"] is not None:
        _plot_ax(axes_flat[3], r["egt_a"], r["egt_b"],
                 "排气温度 (°C)", "°C")

    # 5. 增压压力
    if r["boost_a"] is not None:
        _plot_ax(axes_flat[4], r["boost_a"], r["boost_b"],
                 "增压压力 (kPa)", "kPa")

    # 6. WG 开度
    if r["wg_a"] is not None:
        _plot_ax(axes_flat[5], r["wg_a"], r["wg_b"],
                 "WG开度 (%)", "%")

    plt.tight_layout()
    if save_path:
        Path(save_path).parent.mkdir(parents=True, exist_ok=True)
        plt.savefig(save_path, dpi=150, bbox_inches='tight')
        print(f"📈 图表已保存: {save_path}")
    else:
        plt.show()
    plt.close()


# ────────────────────────────────────────────────────────────
# 7. 报告生成
# ────────────────────────────────────────────────────────────

def _fmt(val_a, val_b, unit="", better="higher"):
    """格式化两个值的对比行，返回 (显示字符串, 胜者)"""
    if val_a is None or val_b is None:
        return ("-", "-", "-")
    result = f"{val_a} vs {val_b}"
    if unit:
        result = f"{val_a}{unit} vs {val_b}{unit}"
    if better == "higher":
        win = name_a if val_a > val_b else name_b
    else:
        win = name_a if val_a < val_b else name_b
    return result, win


def generate_text_report(results: Dict,
                         altitude_results: Optional[Dict] = None) -> str:
    """生成文本格式分析报告。

    Args:
        results: compare_turbochargers() 返回值
        altitude_results: assess_high_altitude() 返回值 (可选)

    Returns:
        格式化的报告字符串

    用法:
        report = generate_text_report(results, altitude_results)
        print(report)
    """
    name_a, name_b = results["names"]
    sm = results["summary"]
    weights = results["weights"]
    scores = results["scores"]
    winner = results["winner"]

    lines = []
    _w = lines.append

    _w(f"## 🔧 增压器对比分析报告")
    _w(f"")
    _w(f"### 📊 基本数据")
    _w(f"")
    _w(f"| 指标 | {name_a} | {name_b} | 优劣 |")
    _w(f"|------|----------|----------|------|")

    rows = [
        ("峰值扭矩",
         f"{sm['torque_peak'][0]} Nm @ {sm['torque_peak_rpm'][0]}rpm",
         f"{sm['torque_peak'][1]} Nm @ {sm['torque_peak_rpm'][1]}rpm"),
        ("平均扭矩",
         f"{sm['torque_avg'][0]} Nm",
         f"{sm['torque_avg'][1]} Nm"),
    ]

    if sm["bsfc_avg"][0] is not None and sm["bsfc_avg"][1] is not None:
        winner_bsfc = (name_a if sm["bsfc_avg"][0] < sm["bsfc_avg"][1]
                       else name_b)
        rows.append(("平均BSFC",
                     f"{sm['bsfc_avg'][0]} g/kWh",
                     f"{sm['bsfc_avg'][1]} g/kWh",
                     f"⭐ {winner_bsfc} 省油"))

    if sm["speed_margin"][0] is not None and sm["speed_margin"][1] is not None:
        winner_margin = (name_a if sm["speed_margin"][0] > sm["speed_margin"][1]
                         else name_b)
        rows.append(("转速余量",
                     f"{sm['speed_margin'][0]:,} rpm",
                     f"{sm['speed_margin'][1]:,} rpm",
                     f"⭐ {winner_margin} 安全"))

    if sm["egt_max"][0] is not None and sm["egt_max"][1] is not None:
        winner_egt = (name_a if sm["egt_max"][0] < sm["egt_max"][1]
                      else name_b)
        rows.append(("最高排气温度",
                     f"{sm['egt_max'][0]}°C",
                     f"{sm['egt_max'][1]}°C",
                     f"⭐ {winner_egt} 热管理好"))

    if sm["wg_high_avg"][0] is not None and sm["wg_high_avg"][1] is not None:
        winner_wg = (name_a if sm["wg_high_avg"][0] < sm["wg_high_avg"][1]
                     else name_b)
        rows.append(("WG高转速均开度",
                     f"{sm['wg_high_avg'][0]}%",
                     f"{sm['wg_high_avg'][1]}%",
                     f"⭐ {winner_wg} 效率高"))

    for row in rows:
        _w(f"| {' | '.join(str(x) for x in row)} |")

    _w(f"")
    _w(f"### 🏆 综合评分")
    _w(f"")
    _w(f"| 维度 (权重) | {name_a} | {name_b} | 胜者 |")
    _w(f"|------------|----------|----------|------|")
    for dim, (w, w_winner) in weights.items():
        a_mark = "✅" if w_winner == name_a else ""
        b_mark = "✅" if w_winner == name_b else ""
        _w(f"| {dim} (×{w}) | {a_mark} | {b_mark} | {w_winner} |")

    _w(f"| **总分** | **{scores[name_a]}** | **{scores[name_b]}** | **🏆 {winner}** |")

    _w(f"")
    _w(f"### 📈 关键曲线解读")
    _w(f"")
    # 自动生成解读骨架
    tq_low_a, tq_low_b = sm["torque_low"]
    _w(f"1. **低速段 (1000-1500rpm)** — "
       f"平均扭矩 {tq_low_a} Nm vs {tq_low_b} Nm, "
       f"{'👍' if tq_low_a > tq_low_b else '👍'}{name_a if tq_low_a > tq_low_b else name_b} 起步响应更好")
    _w(f"2. **中速段 (2000-3000rpm)** — "
       f"峰值扭矩 {sm['torque_peak'][0]} Nm @ {sm['torque_peak_rpm'][0]}rpm "
       f"vs {sm['torque_peak'][1]} Nm @ {sm['torque_peak_rpm'][1]}rpm")
    _w(f"3. **高速段 (3500-5000rpm)** — "
       f"关注增压器转速余量和 WG 开度效率")

    # 高原
    if altitude_results and "error" not in altitude_results:
        _w(f"")
        _w(f"### 🏔️ 高原能力评估 ({altitude_results['altitude_m']}m 海拔)")
        _w(f"")
        _w(f"| 指标 | {name_a} | {name_b} |")
        _w(f"|------|----------|----------|")
        d_a = altitude_results[name_a]
        d_b = altitude_results[name_b]
        _w(f"| 预计最高增压器转速 | {d_a['max_speed_alt']:,} rpm | {d_b['max_speed_alt']:,} rpm |")
        _w(f"| 安全余量 | {d_a['margin']:,} rpm | {d_b['margin']:,} rpm |")
        _w(f"| 安全性 | {d_a['safety']} | {d_b['safety']} |")
        if d_a.get("pressure_ratio_range") and d_b.get("pressure_ratio_range"):
            _w(f"| 增压比 (全速段) | {d_a['pressure_ratio_range'][0]}~{d_a['pressure_ratio_range'][1]} | "
               f"{d_b['pressure_ratio_range'][0]}~{d_b['pressure_ratio_range'][1]} |")
        _w(f"")
        _w(f"> 关键公式: Speed_alt = Speed_0 × √(P0 / P_alt)")
        _w(f"> P_alt = {altitude_results['P_alt_kPa']} kPa, "
           f"转速放大系数 = {altitude_results['speed_multiplier']}×")

    _w(f"")
    _w(f"### 💡 综合建议")
    _w(f"")
    _w(f"- **推荐方案:** {winner}")
    _w(f"- **核心理由:** 综合评分 {scores[winner]} 分 (vs {scores[name_a if winner==name_b else name_b]} 分)")
    _w(f"- **风险提示:** ")
    _w(f"  - 增压器转速限制值 ({sm.get('turbo_speed_limit', 250000)} rpm) 需供应商确认")
    if altitude_results and "error" not in altitude_results:
        if "❌" in altitude_results[name_a]["safety"] or \
           "❌" in altitude_results[name_b]["safety"]:
            _w(f"  - 高原可能超速，需确认 ECU 降扭策略")

    return "\n".join(lines)


# ────────────────────────────────────────────────────────────
# 8. 一站式分析工作流
# ────────────────────────────────────────────────────────────

def full_analysis(filepath: str,
                  name_a: str = "A", name_b: str = "B",
                  n_points: Optional[int] = None,
                  sheet_name: Optional[str] = None,
                  skiprows: int = 0,
                  turbo_speed_limit: int = TURBO_SPEED_LIMIT_DEFAULT,
                  altitude_m: Optional[float] = 3000,
                  save_plot: Optional[str] = None) -> Dict:
    """完整分析流程: 加载 → 分隔 → 对比 → 高原评估 → 可视化 → 报告。

    Args:
        filepath: Excel 文件路径
        name_a, name_b: 增压器名称
        n_points: 每组数据行数 (见 split_groups)
        sheet_name: Excel sheet 名
        skiprows: 跳过的起始行数
        turbo_speed_limit: 增压器转速限制
        altitude_m: 评估海拔 (米)，设为 None 则跳过
        save_plot: 图表保存路径，设为 None 则不保存

    Returns:
        {"results": ..., "altitude": ..., "report": "..."}

    用法 (快捷方式):
        out = full_analysis("260601-增压器对比数据.xlsx",
                            "方案A", "方案B", n_points=9)
        print(out["report"])
    """
    print(f"{'='*60}")
    print(f"🔧 增压器对比分析 — {name_a} vs {name_b}")
    print(f"{'='*60}\n")

    # Step 1: 加载数据
    df = load_excel(filepath, sheet_name=sheet_name, skiprows=skiprows)
    df = clean_columns(df)
    df = ensure_numeric(df)

    # Step 2: 分隔
    df_a, df_b = split_groups(df, n_points=n_points)

    # Step 3: 对比
    results = compare_turbochargers(
        df_a, df_b, name_a=name_a, name_b=name_b,
        turbo_speed_limit=turbo_speed_limit,
    )

    # Step 4: 高原评估
    altitude_results = None
    if altitude_m is not None:
        try:
            altitude_results = assess_high_altitude(
                df_a, df_b, name_a=name_a, name_b=name_b,
                altitude_m=altitude_m,
                turbo_speed_limit=turbo_speed_limit,
            )
        except Exception as e:
            print(f"⚠️ 高原评估跳过: {e}")

    # Step 5: 可视化
    if save_plot:
        try:
            plot_comparison(results, save_path=save_plot)
        except Exception as e:
            print(f"⚠️ 图表生成跳过: {e}")

    # Step 6: 报告
    report = generate_text_report(results, altitude_results)

    # 打印结果
    print(report)

    return {
        "results": results,
        "altitude": altitude_results,
        "report": report,
    }


# ────────────────────────────────────────────────────────────
# 辅助工具
# ────────────────────────────────────────────────────────────

def calculate_pressure_ratio(boost_kpa: np.ndarray,
                             altitude_m: float = 0) -> np.ndarray:
    """计算增压比 (PR = 增压压力 / 对应海拔的大气压)。

    Args:
        boost_kpa: 增压压力数组 (kPa)
        altitude_m: 海拔高度 (米)，默认海平面

    Returns:
        增压比数组

    用法:
        pr = calculate_pressure_ratio(boost_values, altitude_m=3000)
    """
    P_alt = calc_altitude_pressure(altitude_m)
    return boost_kpa / P_alt


def estimate_turbo_speed_at_altitude(speed_sea_level: np.ndarray,
                                     altitude_m: float) -> np.ndarray:
    """推算高原增压器转速。

    Speed_alt = Speed_0 × √(P0 / P_alt)

    Args:
        speed_sea_level: 海平面增压器转速数组 (rpm)
        altitude_m: 目标海拔 (米)

    Returns:
        高原增压器转速数组 (rpm)
    """
    P_alt = calc_altitude_pressure(altitude_m)
    ratio = np.sqrt(ATM_STANDARD / P_alt)
    return speed_sea_level * ratio


def wg_efficiency_assessment(wg_opening: np.ndarray, rpm: np.ndarray,
                             threshold_low: float = 10.0,
                             threshold_high: float = 20.0) -> str:
    """评估 WG 开度效率。

    - < 10%:   ✅ 匹配优秀，废气能量利用率高
    - 10~20%:  ⚡ 匹配良好
    - > 20%:   ⚠️ 匹配效率低，需要放掉大量废气

    Args:
        wg_opening: WG 开度数组 (%)
        rpm: 转速数组
        threshold_low: 低阈值
        threshold_high: 高阈值

    Returns:
        评估结论字符串
    """
    avg = np.nanmean(wg_opening)
    max_val = np.nanmax(wg_opening)

    if max_val < threshold_low:
        return f"✅ 匹配优秀 (平均 {avg:.1f}%, 最大 {max_val:.1f}%)"
    elif avg < threshold_high:
        return f"⚡ 匹配良好 (平均 {avg:.1f}%)"
    else:
        high_rpm_avg = np.nanmean(
            wg_opening[rpm >= 3000] if any(rpm >= 3000) else wg_opening
        )
        return f"⚠️ 匹配效率低 (平均 {avg:.1f}%, 高转速均 {high_rpm_avg:.1f}%)"


def print_data_structure(df: pd.DataFrame) -> None:
    """打印数据结构的友好界面，方便快速理解 Excel 布局。"""
    print(f"📋 行数: {df.shape[0]}, 列数: {df.shape[1]}")
    print(f"\n📐 列名:")
    for i, col in enumerate(df.columns):
        print(f"  [{i:2d}] {col}")
    print(f"\n📄 前 5 行 (缩略):")
    print(df.head(5).to_string(max_colwidth=20))
    print(f"\n📄 后 5 行 (缩略):")
    print(df.tail(5).to_string(max_colwidth=20))


# ────────────────────────────────────────────────────────────
# 9. 单发动机万有特性分析 (非 A/B 对比)
# ────────────────────────────────────────────────────────────

def assess_high_altitude_single(turbo_speeds: np.ndarray,
                                 rpm_values: np.ndarray,
                                 altitude_m: float = 3000,
                                 turbo_speed_limit: int = TURBO_SPEED_LIMIT_DEFAULT
                                 ) -> Dict:
    """评估单台增压器在指定海拔的高原性能。

    Speed_alt = Speed_0 × √(P0 / P_alt)
    其中 P0 = 101.325 kPa, P_alt 为对应海拔的大气压。

    Args:
        turbo_speeds: 海平面增压器转速数组 (rpm)
        rpm_values: 对应转速数组
        altitude_m: 海拔高度 (米)
        turbo_speed_limit: 增压器转速限制值

    Returns:
        Dict 包含高原评估结果
    """
    P_alt = calc_altitude_pressure(altitude_m)
    ratio = np.sqrt(ATM_STANDARD / P_alt)
    speed_alt = np.array(turbo_speeds) * ratio

    max_speed = np.nanmax(speed_alt)
    margin = turbo_speed_limit - max_speed

    def safety_label(m):
        if m > 30000: return "安全"
        elif m > 15000: return "可接受"
        else: return "高风险"

    return {
        "altitude_m": altitude_m,
        "P_alt_kPa": round(P_alt, 2),
        "speed_multiplier": round(ratio, 4),
        "turbo_speed_limit": turbo_speed_limit,
        "max_speed_alt": round(max_speed),
        "margin": round(margin),
        "safety": safety_label(margin),
    }


def _single_engine_performance_core(
    df: pd.DataFrame, rpm_col: str, torque_col: str,
    col_map: Dict,
    turbo_speed_limit: int = TURBO_SPEED_LIMIT_DEFAULT,
    altitude_m: Optional[float] = 3000,
    save_plot: Optional[str] = None,
    standard_engine: Optional[str] = None,
) -> Dict:
    """单发动机性能分析核心 — 接收已加载的数据，不负责 I/O。

    single_engine_analysis() 和 single_engine_full_analysis() 共享此函数，
    消除 full_analysis 中重复加载同一条数据的冗余 I/O。
    """
    # 提取数据
    data = {}
    data['rpm'] = _safe_float(df[rpm_col].values)
    data['torque'] = _safe_float(df[torque_col].values)
    mask = (data['rpm'] > 0) & (data['torque'] > 0)

    for key in ['bsfc', 'turbo_speed', 'boost', 'egt', 'wg', 'power', 'airflow']:
        if key in col_map:
            data[key] = _safe_float(df[col_map[key]].values)
        else:
            data[key] = None

    power_loaded = data.get('power')
    if power_loaded is None or (isinstance(power_loaded, np.ndarray) and np.all(np.isnan(power_loaded))):
        data['power'] = (
            (data['torque'] * data['rpm'] / 9549) if (data['torque'] is not None and data['rpm'] is not None) else None
        )

    rpm = data['rpm'][mask]
    torque = data['torque'][mask]

    # 关键指标
    max_tq_idx = np.nanargmax(torque)
    max_tq = torque[max_tq_idx]
    max_tq_rpm = rpm[max_tq_idx]

    power_vals = data['power'][mask] if data['power'] is not None else (
        torque * rpm / 9549
    )
    max_pwr_idx = np.nanargmax(power_vals)
    max_pwr = power_vals[max_pwr_idx]
    max_pwr_rpm = rpm[max_pwr_idx]

    summary = {
        "max_torque": (round(max_tq, 1), int(max_tq_rpm)),
        "max_power": (round(max_pwr, 1), int(max_pwr_rpm)),
        "max_power_hp": round(max_pwr * 1.341, 1),
        "data_points": int(mask.sum()),
        "rpm_range": (int(rpm.min()), int(rpm.max())),
    }

    if data['bsfc'] is not None:
        bsfc = data['bsfc'][mask]
        min_bsfc = np.nanmin(bsfc)
        min_bsfc_idx = np.nanargmin(bsfc)
        summary["min_bsfc"] = round(min_bsfc, 1)
        summary["min_bsfc_at"] = f"{round(bsfc[min_bsfc_idx],1)} g/kWh @ {int(rpm[min_bsfc_idx])} rpm, {round(torque[min_bsfc_idx],1)} Nm"
    if data['boost'] is not None:
        boost_vals = data['boost'][mask]
        summary["max_boost"] = round(np.nanmax(boost_vals), 1)
    if data['egt'] is not None:
        egt_vals = data['egt'][mask]
        summary["max_egt"] = round(np.nanmax(egt_vals))
    if data['turbo_speed'] is not None:
        ts_vals = data['turbo_speed'][mask]
        summary["max_turbo_speed"] = round(np.nanmax(ts_vals))
        summary["turbo_speed_limit"] = turbo_speed_limit

    altitude_results = None
    if altitude_m is not None and data['turbo_speed'] is not None:
        ts_filtered = data['turbo_speed'][mask]
        rpm_filtered = data['rpm'][mask]
        altitude_results = assess_high_altitude_single(
            ts_filtered, rpm_filtered,
            altitude_m=altitude_m, turbo_speed_limit=turbo_speed_limit,
        )

    # 构建 report
    report_parts = [f"发动机万有特性分析报告"]
    report_parts.append(f"")
    report_parts.append(f"最大扭矩: {summary['max_torque'][0]} Nm @ {summary['max_torque'][1]} rpm")
    report_parts.append(f"最大功率: {summary['max_power'][0]} kW ({summary['max_power_hp']} hp) @ {summary['max_power'][1]} rpm")
    if 'min_bsfc' in summary:
        report_parts.append(f"最低BSFC: {summary['min_bsfc']} g/kWh")
        report_parts.append(f"经济区点: {summary['min_bsfc_at']}")
    if 'max_boost' in summary:
        report_parts.append(f"最高增压压力: {summary['max_boost']} kPa")
    if 'max_egt' in summary:
        report_parts.append(f"最高排气温度: {summary['max_egt']} C")
    if 'max_turbo_speed' in summary:
        report_parts.append(f"最高增压器转速: {summary['max_turbo_speed']} rpm  (限制 {turbo_speed_limit})")

    if altitude_results and 'error' not in altitude_results:
        report_parts.append(f"")
        report_parts.append(f"高原评估 ({altitude_m}m):")
        report_parts.append(f"  放大系数: {altitude_results['speed_multiplier']}x")
        report_parts.append(f"  预估最高转速: {altitude_results['max_speed_alt']} rpm")
        report_parts.append(f"  安全余量: {altitude_results['margin']} rpm ({altitude_results['safety']})")

    report = "\n".join(report_parts)

    # 标准发动机对比
    standard_comparison = None
    if standard_engine == "B15HE":
        rpm_arr = data['rpm']
        torque_arr = data['torque']
        power_arr = data['power']
        bsfc_arr = data.get('bsfc')
        standard_comparison = compare_with_b15he_standard(
            rpm_arr, torque_arr, power_arr, bsfc_arr,
        )
        if standard_comparison and "report" in standard_comparison:
            report = _append_standard_comparison_to_report(report, standard_comparison)
            print(standard_comparison.get("report", ""))

    if save_plot:
        _plot_single_engine(data, col_map, save_path=save_plot)

    print(report)
    ret = {"summary": summary, "altitude": altitude_results, "report": report}
    if standard_comparison is not None:
        ret["standard_comparison"] = standard_comparison
    return ret


def single_engine_analysis(filepath: str,
                           encoding: str = 'gbk',
                           header_rows: int = 5,
                           skip_time_cols: int = 3,
                           turbo_speed_limit: int = TURBO_SPEED_LIMIT_DEFAULT,
                           altitude_m: Optional[float] = 3000,
                           save_plot: Optional[str] = None,
                           standard_engine: Optional[str] = None
                           ) -> Dict:
    """单发动机万有特性数据分析 (非 A/B 对比场景)。

    分析维度：扭矩特性 / BSFC 经济区 / 增压器工作线 /
    WG 开度分布 / 增压压力 / 排气温度 / 高原能力。

    Args:
        filepath: CSV 或 Excel 文件路径
        encoding: CSV 文件编码 (默认 gbk)
        header_rows: CSV 跳过表头行数
        skip_time_cols: CSV 跳过时间戳列数
        turbo_speed_limit: 增压器转速限制
        altitude_m: 评估海拔 (米)，None 则跳过
        save_plot: 图表保存路径，None 则不保存
        standard_engine: 标准发动机对标类型，如 "B15HE" (None=不对比)

    Returns:
        {"summary": ..., "altitude": ..., "report": "...",
         "standard_comparison": ...}  # 当 standard_engine 设置时额外返回
    """
    ext = Path(filepath).suffix.lower()

    if ext == '.csv':
        df = load_csv(filepath, encoding=encoding,
                      header_rows=header_rows, skip_time_cols=skip_time_cols)
    else:
        df = load_excel(filepath)
    df = ensure_numeric(df)

    # 自动检测关键信号列
    rpm_col = detect_column(df, "rpm")
    torque_col = detect_column(df, "torque")

    if rpm_col is None or torque_col is None:
        print(f"无法自动检测列名，可用列: {df.columns.tolist()}")
        # 尝试匹配 ETAS INCA 命名
        for c in df.columns:
            cl = c.lower()
            if 'speed' in cl and 'dyno' in cl and rpm_col is None:
                rpm_col = c
            if 'torque' in cl and 'dyno' in cl and torque_col is None:
                torque_col = c

    print(f"转速列: {rpm_col}, 扭矩列: {torque_col}")

    col_map = detect_all_columns(df)

    return _single_engine_performance_core(
        df, rpm_col, torque_col, col_map,
        turbo_speed_limit=turbo_speed_limit,
        altitude_m=altitude_m,
        save_plot=save_plot,
        standard_engine=standard_engine,
    )


def _plot_single_engine(data: Dict, col_map: Dict, save_path: str,
                        figsize=(18, 10)):
    """单发动机数据可视化 (内部函数)。"""
    _setup_chinese_font()
    try:
        import matplotlib.pyplot as plt
    except ImportError:
        print("matplotlib 未安装，跳过图表")
        return

    rpm = data['rpm']
    torque = data['torque']
    mask = (rpm > 0) & (torque > 0)
    rpm_m, tq_m = rpm[mask], torque[mask]

    fig, axes = plt.subplots(2, 4, figsize=figsize)

    # 1) Torque-RPM scatter colored by BSFC
    ax = axes[0, 0]
    if data['bsfc'] is not None:
        sc = ax.scatter(rpm_m, tq_m, c=data['bsfc'][mask], cmap='viridis', s=30, alpha=0.7)
        plt.colorbar(sc, ax=ax, label='BSFC (g/kWh)')
    else:
        ax.scatter(rpm_m, tq_m, s=30, alpha=0.7)
    ax.set_xlabel('RPM'); ax.set_ylabel('Torque (Nm)')
    ax.set_title('Torque-RPM (colored by BSFC)'); ax.grid(True, alpha=0.3)

    # 2) Power
    ax = axes[0, 1]
    pwr = data['power'] if data['power'] is not None else tq_m * rpm_m / 9549
    ax.scatter(rpm_m, pwr[mask] if isinstance(pwr, np.ndarray) else pwr, s=20, alpha=0.6)
    ax.set_xlabel('RPM'); ax.set_ylabel('Power (kW)')
    ax.set_title('Power'); ax.grid(True, alpha=0.3)

    # 3) BSFC distribution map
    ax = axes[0, 2]
    if data['bsfc'] is not None:
        sc = ax.scatter(rpm_m, tq_m, c=data['bsfc'][mask], cmap='RdYlGn_r', s=40, alpha=0.8)
        plt.colorbar(sc, ax=ax, label='BSFC (g/kWh)')
    ax.set_xlabel('RPM'); ax.set_ylabel('Torque (Nm)')
    ax.set_title('BSFC Distribution'); ax.grid(True, alpha=0.3)

    # 4) Boost pressure map
    ax = axes[0, 3]
    if data['boost'] is not None:
        sc = ax.scatter(rpm_m, tq_m, c=data['boost'][mask], cmap='Blues', s=40, alpha=0.8)
        plt.colorbar(sc, ax=ax, label='Boost (kPa)')
    ax.set_xlabel('RPM'); ax.set_ylabel('Torque (Nm)')
    ax.set_title('Boost Pressure'); ax.grid(True, alpha=0.3)

    # 5) WG opening
    ax = axes[1, 0]
    if data['wg'] is not None:
        sc = ax.scatter(rpm_m, data['wg'][mask], c=tq_m, cmap='plasma', s=30, alpha=0.7)
        plt.colorbar(sc, ax=ax, label='Torque (Nm)')
        ax.axhline(y=10, color='g', linestyle='--', alpha=0.5, label='Good <10%')
        ax.axhline(y=20, color='r', linestyle='--', alpha=0.5, label='Warn >20%')
        ax.legend(fontsize=8)
    ax.set_xlabel('RPM'); ax.set_ylabel('WG Open (%)')
    ax.set_title('WG Opening'); ax.grid(True, alpha=0.3)

    # 6) EGT
    ax = axes[1, 1]
    if data['egt'] is not None:
        sc = ax.scatter(rpm_m, data['egt'][mask], c=tq_m, cmap='hot', s=30, alpha=0.7)
        plt.colorbar(sc, ax=ax, label='Torque (Nm)')
    ax.set_xlabel('RPM'); ax.set_ylabel('EGT (C)')
    ax.set_title('Exhaust Temp'); ax.grid(True, alpha=0.3)

    # 7) Turbo speed
    ax = axes[1, 2]
    if data['turbo_speed'] is not None:
        sc = ax.scatter(rpm_m, data['turbo_speed'][mask], c=tq_m, cmap='viridis', s=30, alpha=0.7)
        plt.colorbar(sc, ax=ax, label='Torque (Nm)')
        limit = TURBO_SPEED_LIMIT_DEFAULT
        ax.axhline(y=limit, color='r', linestyle='--', alpha=0.5, label=f'Limit {limit}')
        ax.legend(fontsize=8)
    ax.set_xlabel('RPM'); ax.set_ylabel('Turbo Speed (rpm)')
    ax.set_title('Turbocharger Speed'); ax.grid(True, alpha=0.3)

    # 8) Airflow
    ax = axes[1, 3]
    if data['airflow'] is not None:
        sc = ax.scatter(rpm_m, data['airflow'][mask], c=tq_m, cmap='plasma', s=30, alpha=0.7)
        plt.colorbar(sc, ax=ax, label='Torque (Nm)')
    ax.set_xlabel('RPM'); ax.set_ylabel('Air Flow (kg/h)')
    ax.set_title('Intake Air Flow'); ax.grid(True, alpha=0.3)

    plt.tight_layout()
    Path(save_path).parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    print(f"Chart saved: {save_path}")
    plt.close()


# ────────────────────────────────────────────────────────────
# 10. 单发动机燃烧特性分析 (功率/油耗/COV/AI50/点火角/爆震/VVT)
# ────────────────────────────────────────────────────────────

def _safe_extract(df: pd.DataFrame, col_map: Dict, key: str,
                  mask: Optional[np.ndarray] = None) -> Optional[np.ndarray]:
    """安全提取列数据，返回 float ndarray 或 None。"""
    if key not in col_map:
        return None
    vals = _safe_float(df[col_map[key]].values)
    if mask is not None:
        vals = vals.copy()
        vals[~mask] = np.nan
    return vals


def single_engine_combustion_analysis(
    df: pd.DataFrame,
    rpm_col: str, torque_col: str,
    col_map: Dict,
    turbo_speed_limit: int = TURBO_SPEED_LIMIT_DEFAULT,
    altitude_m: Optional[float] = 3000,
    save_plot: Optional[str] = None
) -> Dict:
    """单发动机燃烧特性全面分析。

    分析维度：
      - 功率 / BSFC 经济区 / 燃油消耗量
      - COV (循环变动系数) — 燃烧稳定性
      - AI50 (CA50) — 燃烧相位
      - 点火角 (实际 / MBT / 退角)
      - 爆震窗口
      - VVT (可变气门正时)
      - IMEP (平均有效压力)
      - 增压器高原评估

    Args:
        df: 清洗后的 DataFrame
        rpm_col: 转速列名
        torque_col: 扭矩列名
        col_map: detect_all_columns() 返回的列名映射
        turbo_speed_limit: 增压器转速限制
        altitude_m: 评估海拔 (米)，None 则跳过
        save_plot: 图表保存路径，None 则不保存

    Returns:
        Dict 包含 summary, combustion, altitude, report
    """
    # 有效点筛选
    rpm = _safe_float(df[rpm_col].values)
    torque = _safe_float(df[torque_col].values)
    power_raw = _safe_extract(df, col_map, "power")
    mask = (rpm > 0) & (torque > 0)

    rpm_m = rpm[mask]
    torque_m = torque[mask]
    power_m = power_raw[mask] if power_raw is not None else (torque_m * rpm_m / 9549)

    # 按转速分组合并
    rpm_group = np.round(rpm_m / 50) * 50
    group_data = {"rpm": rpm_m, "torque": torque_m, "power": power_m}

    # 提取各燃烧信号
    for key in ['bsfc', 'cov', 'ai50', 'spark_act', 'spark_mbt',
                'spark_delta', 'knock', 'knock', 'vvt', 'fuel_flow',
                'imep', 'boost', 'egt', 'wg', 'turbo_speed', 'airflow']:
        vals = _safe_extract(df, col_map, key)
        group_data[key] = vals[mask] if vals is not None else None

    # ── 关键统计 ──
    summary = {}
    summary["data_points"] = int(mask.sum())
    summary["rpm_range"] = (int(rpm_m.min()), int(rpm_m.max()))

    # 功率
    max_pwr = np.nanmax(power_m)
    summary["max_power"] = (round(max_pwr, 1), int(rpm_m[np.nanargmax(power_m)]))
    summary["max_power_hp"] = round(max_pwr * 1.341, 1)

    # 扭矩
    max_tq = np.nanmax(torque_m)
    summary["max_torque"] = (round(max_tq, 1), int(rpm_m[np.nanargmax(torque_m)]))

    # BSFC
    bsfc = group_data['bsfc']
    if bsfc is not None:
        valid = bsfc > 0
        if valid.any():
            summary["min_bsfc"] = round(np.nanmin(bsfc[valid]), 1)
            best_idx = np.nanargmin(bsfc[valid])
            summary["bsfc_economy_zone"] = (
                f"{summary['min_bsfc']} g/kWh @ "
                f"{int(rpm_m[valid][best_idx])} rpm, "
                f"{round(torque_m[valid][best_idx], 1)} Nm"
            )
            # BSFC < 240 的经济区占比
            economy = (bsfc[valid] < 240) & (bsfc[valid] > 0)
            summary["bsfc_below_240_ratio"] = f"{economy.sum()}/{valid.sum()} ({economy.sum()/valid.sum()*100:.0f}%)"

    # COV
    cov = group_data['cov']
    if cov is not None:
        cov_valid = cov[(cov > 0) & (cov < 100)]  # 过滤异常值
        if len(cov_valid) > 0:
            summary["cov"] = {
                "min": round(float(cov_valid.min()), 2),
                "max": round(float(cov_valid.max()), 2),
                "mean": round(float(cov_valid.mean()), 2),
                "above_5pct": int((cov_valid > 5).sum()),
                "above_3pct_lowload": int(
                    ((cov_valid > 3) & (torque_m[(cov > 0) & (cov < 100)] < 30)).sum()
                ),
            }

    # AI50
    ai50 = group_data['ai50']
    if ai50 is not None:
        cov_mask = pd.Series(True, index=ai50.index) if cov is None else ((cov > 0) & (cov < 100))
        ai50_valid = ai50[(ai50 > -10) & (ai50 < 60) & cov_mask]
        if len(ai50_valid) > 0:
            optimal = ((ai50_valid >= 6) & (ai50_valid <= 12)).sum()
            late = (ai50_valid > 15).sum()
            early = (ai50_valid < 5).sum()
            summary["ai50"] = {
                "min": round(float(ai50_valid.min()), 1),
                "max": round(float(ai50_valid.max()), 1),
                "mean": round(float(ai50_valid.mean()), 1),
                "optimal_6_12_ratio": f"{optimal}/{len(ai50_valid)} ({optimal/len(ai50_valid)*100:.0f}%)",
                "late_gt_15": int(late),
                "early_lt_5": int(early),
            }

    # 点火角
    spark = group_data['spark_act']
    if spark is not None:
        spark_valid = spark[(spark > -50) & (spark < 100)]
        if len(spark_valid) > 0:
            summary["spark"] = {
                "min": round(float(spark_valid.min()), 1),
                "max": round(float(spark_valid.max()), 1),
                "mean": round(float(spark_valid.mean()), 1),
            }
        # 大负荷点火角
        full_load = torque_m > 150
        if full_load.any() and spark_valid.any():
            fl_spark = spark[full_load]
            summary["spark"]["full_load_mean"] = round(float(np.nanmean(fl_spark)), 1)

    # 点火退角
    delta = group_data['spark_delta']
    if delta is not None:
        delta_valid = delta[(delta >= 0) & (delta < 50)]
        if len(delta_valid) > 0:
            summary["spark_delta"] = {
                "min": round(float(delta_valid.min()), 1),
                "max": round(float(delta_valid.max()), 1),
                "mean": round(float(delta_valid.mean()), 1),
                "gt_5_retarded": int((delta_valid > 5).sum()),
            }

    # 增压器 / 高原
    if group_data['turbo_speed'] is not None:
        ts = group_data['turbo_speed']
        summary["max_turbo_speed"] = round(float(np.nanmax(ts)))
        summary["turbo_speed_limit"] = turbo_speed_limit

    altitude_results = None
    if altitude_m is not None and group_data['turbo_speed'] is not None:
        ts_valid = group_data['turbo_speed'][group_data['turbo_speed'] > 0]
        rpm_valid = group_data['rpm'][group_data['turbo_speed'] > 0]
        if len(ts_valid) > 0:
            altitude_results = assess_high_altitude_single(
                ts_valid, rpm_valid,
                altitude_m=altitude_m, turbo_speed_limit=turbo_speed_limit,
            )

    # ── 报告 ──
    report = _build_combustion_report(summary, altitude_results, altitude_m)
    print(report)

    if save_plot:
        _plot_combustion_analysis(group_data, save_path=save_plot)

    return {
        "summary": summary,
        "altitude": altitude_results,
        "report": report,
        "group_data": group_data,
    }


def _build_combustion_report(summary: Dict,
                              altitude_results: Optional[Dict] = None,
                              altitude_m: Optional[float] = None) -> str:
    """生成燃烧特性分析报告。"""
    lines = []
    _w = lines.append

    _w("## 🔥 发动机燃烧特性分析报告")
    _w("")

    # 基本性能
    _w("### 📊 基本性能")
    _w(f"有效数据点: {summary.get('data_points', '-')}")
    _w(f"转速范围: {summary.get('rpm_range', ('-','-'))[0]} - {summary.get('rpm_range', ('-','-'))[1]} rpm")
    _w("")

    _w("| 指标 | 数值 |")
    _w("|------|------|")
    _w(f"| 最大扭矩 | {summary.get('max_torque', ('-',''))[0]} Nm @ {summary.get('max_torque', ('','-'))[1]} rpm |")
    _w(f"| 最大功率 | {summary.get('max_power', ('-',''))[0]} kW ({summary.get('max_power_hp', '-')} hp) @ {summary.get('max_power', ('','-'))[1]} rpm |")
    if 'min_bsfc' in summary:
        _w(f"| 最低BSFC | {summary['min_bsfc']} g/kWh |")
    if 'bsfc_economy_zone' in summary:
        _w(f"| 经济区 | {summary['bsfc_economy_zone']} |")
    if 'bsfc_below_240_ratio' in summary:
        _w(f"| BSFC<240占比 | {summary['bsfc_below_240_ratio']} |")
    if 'max_turbo_speed' in summary:
        _w(f"| 最高增压器转速 | {summary['max_turbo_speed']:,} rpm (限制 {summary.get('turbo_speed_limit','-')}) |")

    # COV
    if 'cov' in summary:
        c = summary['cov']
        _w("")
        _w("### 🔄 COV 循环变动系数")
        _w(f"| 指标 | 数值 |")
        _w(f"|------|------|")
        _w(f"| 范围 | {c['min']}% – {c['max']}% |")
        _w(f"| 均值 | {c['mean']}% |")
        _w(f"| COV > 5% (不稳定) | {c['above_5pct']} 点 |")
        _w(f"| 低负荷 COV>3% | {c['above_3pct_lowload']} 点 |")
        if c['above_5pct'] > 0:
            _w("> ⚠️ 存在燃烧不稳定点，需关注")
        else:
            _w("> ✅ 全工况燃烧稳定")

    # AI50
    if 'ai50' in summary:
        a = summary['ai50']
        _w("")
        _w("### 🔥 AI50 (CA50) 燃烧相位")
        _w(f"| 指标 | 数值 |")
        _w(f"|------|------|")
        _w(f"| 范围 | {a['min']}° – {a['max']}° CA ATDC |")
        _w(f"| 均值 | {a['mean']}° CA ATDC |")
        _w(f"| 最佳区间 6-12° | {a['optimal_6_12_ratio']} |")
        _w(f"| 燃烧过迟 >15° | {a['late_gt_15']} 点 |")
        _w(f"| 燃烧过早 <5° | {a['early_lt_5']} 点 |")
        if a['late_gt_15'] > len([1 for _ in range(len(str(a['late_gt_15'])))]) * 10:
            _w("> ⚠️ 大量工况燃烧相位偏晚，影响热效率")

    # 点火角
    if 'spark' in summary:
        s = summary['spark']
        _w("")
        _w("### ⚡ 点火角")
        _w(f"| 指标 | 数值 |")
        _w(f"|------|------|")
        _w(f"| 实际点火角范围 | {s['min']}° – {s['max']}° BTDC |")
        _w(f"| 实际点火角均值 | {s['mean']}° BTDC |")
        if 'full_load_mean' in s:
            _w(f"| 大负荷(>150Nm)均值 | {s['full_load_mean']}° BTDC |")

    if 'spark_delta' in summary:
        d = summary['spark_delta']
        _w("")
        _w("### ⚡ 点火退角 (MBT差值)")
        _w(f"| 指标 | 数值 |")
        _w(f"|------|------|")
        _w(f"| 范围 | {d['min']}° – {d['max']}° |")
        _w(f"| 均值 | {d['mean']}° |")
        _w(f"| 退角 >5° (受爆震限制) | {d['gt_5_retarded']} 点 |")
        if d['gt_5_retarded'] > 10:
            _w("> 🔥 大量点存在明显退角，ECU受爆震限制标定偏保守")

    # 高原
    if altitude_results and 'error' not in altitude_results:
        _w("")
        _w(f"### 🏔️ 高原能力评估 ({altitude_m}m)")
        _w(f"| 指标 | 数值 |")
        _w(f"|------|------|")
        _w(f"| 预计最高增压器转速 | {altitude_results.get('max_speed_alt', '-'):,} rpm |")
        _w(f"| 安全余量 | {altitude_results.get('margin', '-')} rpm |")
        _w(f"| 安全性 | {altitude_results.get('safety', '-')} |")
        _w(f"| 转速放大系数 | {altitude_results.get('speed_multiplier', '-')}x |")

    _w("")
    _w("### 💡 综合建议")
    if 'bsfc' in summary:
        _w(f"- 最佳经济区: 控制 AI50 在 6-12°CA ATDC, BSFC 可低于 240 g/kWh")
    if 'spark_delta' in summary and summary['spark_delta'].get('gt_5_retarded', 0) > 10:
        _w(f"- 点火退角过大点较多 ({summary['spark_delta']['gt_5_retarded']}点)，建议检查爆震传感器标定")
    if altitude_results and 'error' not in altitude_results:
        if '高风险' in altitude_results.get('safety', ''):
            _w(f"- ⚠️ 高原存在超速风险，需确认 ECU 降扭策略")

    return "\n".join(lines)


def _plot_combustion_analysis(group_data: Dict, save_path: str,
                               figsize=(18, 16)):
    """燃烧特性可视化 (9子图)。"""
    try:
        import matplotlib.pyplot as plt
    except ImportError:
        print("matplotlib 未安装，跳过图表")
        return

    _setup_chinese_font()
    rpm = group_data['rpm']
    torque = group_data['torque']

    fig, axes = plt.subplots(3, 3, figsize=figsize)

    # 1. Power
    ax = axes[0, 0]
    if group_data['power'] is not None:
        ax.scatter(rpm, group_data['power'], c=torque, cmap='plasma', s=30, alpha=0.7)
    ax.set_xlabel('RPM'); ax.set_ylabel('Power (kW)')
    ax.set_title('Power'); ax.grid(True, alpha=0.3)

    # 2. BSFC map
    ax = axes[0, 1]
    if group_data['bsfc'] is not None:
        sc = ax.scatter(rpm, torque, c=group_data['bsfc'],
                        cmap='RdYlGn_r', s=40, alpha=0.8, vmin=220, vmax=400)
        plt.colorbar(sc, ax=ax, label='BSFC (g/kWh)')
        best = np.nanargmin(group_data['bsfc'])
        ax.plot(rpm[best], torque[best], 'r*', markersize=15,
                label=f"Best {group_data['bsfc'][best]:.0f} g/kWh")
        ax.legend(fontsize=8)
    ax.set_xlabel('RPM'); ax.set_ylabel('Torque (Nm)')
    ax.set_title('BSFC Map'); ax.grid(True, alpha=0.3)

    # 3. COV map
    ax = axes[0, 2]
    if group_data['cov'] is not None:
        sc = ax.scatter(rpm, torque, c=group_data['cov'],
                        cmap='RdYlGn_r', s=40, alpha=0.8, vmin=0, vmax=5)
        plt.colorbar(sc, ax=ax, label='COV (%)')
        ax.axhline(y=30, color='gray', linestyle=':', alpha=0.5)
    ax.set_xlabel('RPM'); ax.set_ylabel('Torque (Nm)')
    ax.set_title('COV - Combustion Stability'); ax.grid(True, alpha=0.3)

    # 4. AI50 map
    ax = axes[1, 0]
    if group_data['ai50'] is not None:
        sc = ax.scatter(rpm, torque, c=group_data['ai50'],
                        cmap='coolwarm', s=40, alpha=0.8, vmin=0, vmax=25)
        plt.colorbar(sc, ax=ax, label='AI50 (°CA ATDC)')
    ax.set_xlabel('RPM'); ax.set_ylabel('Torque (Nm)')
    ax.set_title('AI50 (CA50)'); ax.grid(True, alpha=0.3)

    # 5. Spark advance
    ax = axes[1, 1]
    if group_data['spark_act'] is not None:
        sc = ax.scatter(rpm, torque, c=group_data['spark_act'],
                        cmap='viridis', s=40, alpha=0.8)
        plt.colorbar(sc, ax=ax, label='Spark Adv (°BTDC)')
    ax.set_xlabel('RPM'); ax.set_ylabel('Torque (Nm)')
    ax.set_title('Spark Advance'); ax.grid(True, alpha=0.3)

    # 6. Spark delta from MBT
    ax = axes[1, 2]
    if group_data['spark_delta'] is not None:
        sc = ax.scatter(rpm, torque, c=group_data['spark_delta'],
                        cmap='Reds', s=40, alpha=0.8)
        plt.colorbar(sc, ax=ax, label='Delta (°CA)')
        ax.set_title('Spark Retard from MBT')
    elif group_data['spark_mbt'] is not None and group_data['spark_act'] is not None:
        delta = group_data['spark_mbt'] - group_data['spark_act']
        sc = ax.scatter(rpm, torque, c=delta, cmap='Reds', s=40, alpha=0.8)
        plt.colorbar(sc, ax=ax, label='Delta (°CA)')
        ax.set_title('Spark Retard (calc)')
    ax.set_xlabel('RPM'); ax.set_ylabel('Torque (Nm)')
    ax.grid(True, alpha=0.3)

    # 7. BSFC vs AI50
    ax = axes[2, 0]
    if group_data['ai50'] is not None and group_data['bsfc'] is not None:
        valid = (group_data['cov'] is None) | (group_data['cov'] < 5) if group_data['cov'] is not None else np.ones_like(rpm, dtype=bool)
        sc = ax.scatter(group_data['ai50'][valid], group_data['bsfc'][valid],
                        c=rpm[valid], cmap='viridis', s=30, alpha=0.7)
        plt.colorbar(sc, ax=ax, label='RPM')
    ax.set_xlabel('AI50 (°CA ATDC)'); ax.set_ylabel('BSFC (g/kWh)')
    ax.set_title('BSFC vs AI50'); ax.grid(True, alpha=0.3)

    # 8. Spark by load level
    ax = axes[2, 1]
    if group_data['spark_act'] is not None:
        load_bins = [(0, 30, 'Low <30Nm'), (30, 80, 'Mid 30-80'),
                     (80, 140, 'High 80-140'), (140, 999, 'Full >140Nm')]
        colors = ['blue', 'green', 'orange', 'red']
        for (lo, hi, label), cl in zip(load_bins, colors):
            idx = (torque >= lo) & (torque < hi)
            ax.scatter(rpm[idx], group_data['spark_act'][idx], s=15, alpha=0.6, c=cl, label=label)
        ax.legend(fontsize=8)
    ax.set_xlabel('RPM'); ax.set_ylabel('Spark Adv (°BTDC)')
    ax.set_title('Spark by Load Level'); ax.grid(True, alpha=0.3)

    # 9. COV vs IMEP
    ax = axes[2, 2]
    if group_data['imep'] is not None and group_data['cov'] is not None:
        sc = ax.scatter(group_data['imep'], group_data['cov'],
                        c=rpm, cmap='viridis', s=30, alpha=0.7)
        plt.colorbar(sc, ax=ax, label='RPM')
        ax.axhline(y=3, color='r', linestyle='--', alpha=0.5, label='Stability limit')
        ax.legend(fontsize=8)
    ax.set_xlabel('IMEP (bar)'); ax.set_ylabel('COV (%)')
    ax.set_title('COV vs IMEP'); ax.grid(True, alpha=0.3)

    plt.suptitle('Combustion Characteristics Analysis', fontsize=14)
    plt.tight_layout()
    Path(save_path).parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    print(f"Chart saved: {save_path}")
    plt.close()


# ────────────────────────────────────────────────────────────
# 11. 燃烧参数敏感性分析
# ────────────────────────────────────────────────────────────

# 可选依赖检测
try:
    from sklearn.ensemble import RandomForestRegressor
    from sklearn.inspection import permutation_importance
    from sklearn.model_selection import KFold, cross_val_score
    from sklearn.feature_selection import mutual_info_regression
    _HAS_SKLEARN = True
except ImportError:
    _HAS_SKLEARN = False


def analyze_combustion_sensitivity(
    filepath: str,
    encoding: str = 'gbk',
    header_rows: int = 5,
    skip_time_cols: int = 3,
) -> Dict:
    """燃烧参数敏感性分析 (ML 级别) — 多方法量化控制参数对 BSFC/COV 的影响。

    分析管线：
      1. 燃烧基线诊断 (调用 single_engine_combustion_analysis)
      2. Pearson + Spearman + 互信息 (线性 + 非线性 + 信息论)
      3. Random Forest 特征重要性 (impurity + permutation, 需 sklearn)
      4. 偏依赖分析 (特征对目标的非线性响应曲线)
      5. RPM 条件分析 (低/中/高转速分区)
      6. 负荷分区分析 (低/中/大/全负荷)
      7. 交叉验证稳定性 (KFold, 需 sklearn)

    Args:
        filepath: CSV 或 Excel 文件路径
        encoding: CSV 编码 (默认 gbk)
        header_rows: CSV 跳过表头行数
        skip_time_cols: CSV 跳过时间戳列数

    Returns:
        Dict 包含 report, feature_importance, combustion_baseline,
             sensitivity_matrix, rpm_analysis, load_zone_analysis, cv_stability
    """
    print(f"{'='*60}")
    print(" 燃烧参数敏感性分析 (ML)")
    print(f"{'='*60}")
    if _HAS_SKLEARN:
        print("  ✓ scikit-learn 可用 — Random Forest + 互信息 + 交叉验证")
    else:
        print("  ⚠️ scikit-learn 未安装 — 仅 Pearson/Spearman 相关分析")
        print("    安装: pip install scikit-learn")
    print()

    # ── 1. 加载数据 ──
    ext = Path(filepath).suffix.lower()
    if ext == '.csv':
        df = load_csv(filepath, encoding=encoding,
                      header_rows=header_rows, skip_time_cols=skip_time_cols)
    else:
        df = load_excel(filepath)
    df = ensure_numeric(df)

    # ── 2. 检测列名 ──
    rpm_col = detect_column(df, "rpm")
    torque_col = detect_column(df, "torque")
    if rpm_col is None or torque_col is None:
        raise ValueError(f"无法检测转速/扭矩列。可用列: {df.columns.tolist()}")
    print(f"转速列: {rpm_col}, 扭矩列: {torque_col}")

    col_map = detect_all_columns(df)
    col_map['rpm'] = rpm_col
    col_map['torque'] = torque_col
    combustion_cols = {k: v for k, v in col_map.items()
                       if v and k not in ('rpm', 'torque', 'power', 'airflow')}
    print(f"检测到的燃烧相关列: {combustion_cols}\n")

    # ── 3. 燃烧基线诊断 ──
    print("── 燃烧基线诊断 ──")
    try:
        combustion_baseline = single_engine_combustion_analysis(
            df, rpm_col, torque_col, col_map,
            turbo_speed_limit=TURBO_SPEED_LIMIT_DEFAULT,
            altitude_m=None, save_plot=None,
        )
        baseline_ok = True
    except Exception as e:
        print(f"  ⚠️ 基线诊断失败: {e}")
        combustion_baseline = {"summary": {}, "report": str(e)}
        baseline_ok = False

    # ── 4. 提取特征矩阵 ──
    rpm = _safe_float(df[rpm_col].values)
    torque = _safe_float(df[torque_col].values)
    mask = (rpm > 0) & (torque > 0)
    rpm_m = rpm[mask]
    torque_m = torque[mask]
    n_points = int(mask.sum())

    feature_config = {
        'ai50':        'AI50 (燃烧相位)',
        'spark_act':   '实际点火角',
        'spark_delta': '点火退角',
        'knock':       '爆震强度',
        'vvt':         'VVT 正时',
        'egt':         '排气温度',
        'boost':       '增压压力',
        'imep':        'IMEP',
    }
    features = {}
    for key in feature_config:
        vals = _safe_extract(df, col_map, key)
        features[key] = vals[mask] if vals is not None else None

    bsfc_raw = _safe_extract(df, col_map, "bsfc")
    cov_raw = _safe_extract(df, col_map, "cov")
    bsfc = bsfc_raw[mask] if bsfc_raw is not None else None
    cov = cov_raw[mask] if cov_raw is not None else None

    valid_bsfc = (bsfc > 0) & (bsfc < 1000) if bsfc is not None else np.zeros(n_points, dtype=bool)
    valid_cov = (cov > 0) & (cov < 100) if cov is not None else np.zeros(n_points, dtype=bool)
    feature_keys = [k for k in feature_config if features[k] is not None]

    print(f"有效数据点: {n_points}  (BSFC: {valid_bsfc.sum()}, COV: {valid_cov.sum()})")
    print(f"特征数: {len(feature_keys)} — {', '.join(feature_config[k] for k in feature_keys)}\n")

    if len(feature_keys) == 0:
        return {"report": "⚠️ 未检测到任何燃烧特性列。", "feature_importance": [],
                "sensitivity_matrix": {}, "combustion_baseline": combustion_baseline}

    # ── 5. Pearson + Spearman 相关性 ──
    corr_pearson = {"BSFC": {}, "COV": {}}
    corr_spearman = {"BSFC": {}, "COV": {}}
    from scipy.stats import spearmanr as _spearmanr_func

    for target_name, target_arr, valid_mask in [
        ("BSFC", bsfc, valid_bsfc), ("COV", cov, valid_cov),
    ]:
        if target_arr is None or valid_mask.sum() < 5:
            corr_pearson[target_name] = None
            corr_spearman[target_name] = None
            continue
        for fk in feature_keys:
            fv = features[fk]
            if fv is None: continue
            joint = valid_mask & ~np.isnan(fv) & (fv > -1e6) & (fv < 1e6)
            if joint.sum() < 5: continue
            # Pearson
            p = np.corrcoef(fv[joint], target_arr[joint])[0, 1]
            corr_pearson[target_name][fk] = round(float(p) if not np.isnan(p) else 0, 3)
            # Spearman
            try:
                s, _ = _spearmanr_func(fv[joint], target_arr[joint])
                corr_spearman[target_name][fk] = round(float(s) if not np.isnan(s) else 0, 3)
            except Exception:
                corr_spearman[target_name][fk] = 0

    # ── 6. 互信息 (scikit-learn) ──
    mi_scores = {"BSFC": {}, "COV": {}}
    if _HAS_SKLEARN:
        for target_name, target_arr, valid_mask in [
            ("BSFC", bsfc, valid_bsfc), ("COV", cov, valid_cov),
        ]:
            if target_arr is None or valid_mask.sum() < 10: continue
            joint_mask = valid_mask.copy()
            valid_fks = []
            for fk in feature_keys:
                joint_mask_fk = joint_mask & ~np.isnan(features[fk])
                if joint_mask_fk.sum() >= 10:
                    valid_fks.append(fk)
                joint_mask = joint_mask_fk
            if len(valid_fks) < 2 or joint_mask.sum() < 10: continue
            X_mi = np.column_stack([features[fk][joint_mask] for fk in valid_fks])
            y_mi = target_arr[joint_mask]
            try:
                mi = mutual_info_regression(X_mi, y_mi, random_state=42)
                for i, fk in enumerate(valid_fks):
                    mi_scores[target_name][fk] = round(float(mi[i]), 4)
            except Exception:
                pass

    # ── 7. Random Forest 特征重要性 ──
    rf_importance = []
    rf_permutation = []
    if _HAS_SKLEARN:
        for target_name, target_arr, valid_mask in [
            ("BSFC", bsfc, valid_bsfc), ("COV", cov, valid_cov),
        ]:
            if target_arr is None or valid_mask.sum() < 15: continue
            joint_mask = valid_mask.copy()
            for fk in feature_keys:
                joint_mask &= ~np.isnan(features[fk])
            if joint_mask.sum() < 15: continue
            X_rf = np.column_stack([features[fk][joint_mask] for fk in feature_keys])
            y_rf = target_arr[joint_mask]

            try:
                rf = RandomForestRegressor(n_estimators=100, max_depth=8,
                                           random_state=42, n_jobs=1)
                rf.fit(X_rf, y_rf)
                # Impurity importance
                total = rf.feature_importances_.sum()
                for i, fk in enumerate(feature_keys):
                    rf_importance.append({
                        "target": target_name,
                        "feature": fk,
                        "feature_cn": feature_config[fk],
                        "importance": round(float(rf.feature_importances_[i] / total) if total > 0 else 0, 4),
                        "method": "rf_impurity",
                    })
                # Permutation importance
                perm = permutation_importance(rf, X_rf, y_rf, n_repeats=5,
                                              random_state=42, n_jobs=1)
                for i, fk in enumerate(feature_keys):
                    rf_permutation.append({
                        "target": target_name,
                        "feature": fk,
                        "feature_cn": feature_config[fk],
                        "importance": round(float(perm.importances_mean[i]), 4),
                        "importance_std": round(float(perm.importances_std[i]), 4),
                        "method": "rf_permutation",
                    })
            except Exception as e:
                print(f"  ⚠️ RF 建模失败 ({target_name}): {e}")

    # ── 8. RPM 条件分析 ──
    rpm_bands = [
        ("低速 (<2000 rpm)", 0, 2000),
        ("中速 (2000-3500 rpm)", 2000, 3500),
        ("高速 (>3500 rpm)", 3500, 99999),
    ]
    rpm_analysis = {}
    for band_name, lo, hi in rpm_bands:
        band_idx = (rpm_m >= lo) & (rpm_m < hi)
        if band_idx.sum() < 5: continue
        band_corrs = {}
        for fk in feature_keys:
            fv = features[fk]
            if fv is None: continue
            band_fv = fv[band_idx]
            for target_name, target_arr, valid_mask in [
                ("BSFC", bsfc, valid_bsfc), ("COV", cov, valid_cov),
            ]:
                if target_arr is None: continue
                joint = valid_mask[band_idx] & ~np.isnan(band_fv) & (band_fv > -1e6) & (band_fv < 1e6)
                if joint.sum() < 5: continue
                c = np.corrcoef(band_fv[joint], target_arr[band_idx][joint])[0, 1]
                band_corrs.setdefault(target_name, {})[fk] = round(float(c) if not np.isnan(c) else 0, 3)
        if band_corrs:
            rpm_analysis[band_name] = {"n_points": int(band_idx.sum()), "correlations": band_corrs}

    # ── 9. 负荷分区分析 ──
    load_zones = [
        ("低负荷 (<50 Nm)", lambda t: t < 50),
        ("中负荷 (50-120 Nm)", lambda t: (t >= 50) & (t < 120)),
        ("大负荷 (120-200 Nm)", lambda t: (t >= 120) & (t < 200)),
        ("全负荷 (>=200 Nm)", lambda t: t >= 200),
    ]
    load_zone_analysis = {}
    for zone_name, zone_fn in load_zones:
        zone_idx = zone_fn(torque_m)
        if zone_idx.sum() < 5: continue
        zone_corrs = {}
        for fk in feature_keys:
            fv = features[fk]
            if fv is None: continue
            zone_fv = fv[zone_idx]
            for target_name, target_arr, valid_mask in [
                ("BSFC", bsfc, valid_bsfc), ("COV", cov, valid_cov),
            ]:
                if target_arr is None: continue
                joint = valid_mask[zone_idx] & ~np.isnan(zone_fv)
                if joint.sum() < 5: continue
                c = np.corrcoef(zone_fv[joint], target_arr[zone_idx][joint])[0, 1]
                zone_corrs.setdefault(target_name, {})[fk] = round(float(c) if not np.isnan(c) else 0, 3)
        if zone_corrs:
            load_zone_analysis[zone_name] = {"n_points": int(zone_idx.sum()), "correlations": zone_corrs}

    # ── 10. 交叉验证稳定性 ──
    cv_stability = []
    if _HAS_SKLEARN and rf_importance:
        for target_name, target_arr, valid_mask in [
            ("BSFC", bsfc, valid_bsfc), ("COV", cov, valid_cov),
        ]:
            if target_arr is None or valid_mask.sum() < 20: continue
            joint_mask = valid_mask.copy()
            for fk in feature_keys:
                joint_mask &= ~np.isnan(features[fk])
            if joint_mask.sum() < 20: continue
            X_cv = np.column_stack([features[fk][joint_mask] for fk in feature_keys])
            y_cv = target_arr[joint_mask]
            n_folds = min(5, joint_mask.sum() // 5)
            if n_folds < 2: continue
            try:
                rf_cv = RandomForestRegressor(n_estimators=100, max_depth=8,
                                              random_state=42, n_jobs=1)
                scores = cross_val_score(rf_cv, X_cv, y_cv, cv=n_folds,
                                         scoring='neg_mean_squared_error')
                cv_stability.append({
                    "target": target_name,
                    "cv_folds": n_folds,
                    "cv_mse_mean": round(float(-scores.mean()), 2),
                    "cv_mse_std": round(float(scores.std()), 2),
                    "cv_r2_mean": round(float(
                        cross_val_score(rf_cv, X_cv, y_cv, cv=n_folds,
                                        scoring='r2').mean()), 3),
                })
            except Exception:
                pass

    # ── 11. 合并特征重要性排名 ──
    feature_importance = _merge_feature_importance(
        corr_pearson, corr_spearman, mi_scores, rf_importance, rf_permutation,
        feature_config,
    )

    # ── 12. 偏依赖数据 ──
    partial_dependence = {}
    if _HAS_SKLEARN and feature_keys:
        for target_name, target_arr, valid_mask in [
            ("BSFC", bsfc, valid_bsfc), ("COV", cov, valid_cov),
        ]:
            if target_arr is None or valid_mask.sum() < 15: continue
            joint_mask = valid_mask.copy()
            for fk in feature_keys:
                joint_mask &= ~np.isnan(features[fk])
            if joint_mask.sum() < 15: continue
            X_pd = np.column_stack([features[fk][joint_mask] for fk in feature_keys])
            y_pd = target_arr[joint_mask]
            try:
                rf_pd = RandomForestRegressor(n_estimators=100, max_depth=8,
                                              random_state=42, n_jobs=1)
                rf_pd.fit(X_pd, y_pd)
                pd_data = {}
                for i, fk in enumerate(feature_keys):
                    fv = features[fk][joint_mask]
                    x_grid = np.linspace(np.nanmin(fv), np.nanmax(fv), 30)
                    X_grid = np.tile(X_pd.mean(axis=0), (30, 1))
                    X_grid[:, i] = x_grid
                    y_pred = rf_pd.predict(X_grid)
                    pd_data[fk] = {"x": x_grid.tolist(), "y": y_pred.tolist()}
                partial_dependence[target_name] = pd_data
            except Exception:
                pass

    # ── 13. 生成报告 ──
    report = _build_ml_sensitivity_report(
        feature_importance, corr_pearson, corr_spearman, mi_scores,
        rpm_analysis, load_zone_analysis, cv_stability,
        n_points, feature_keys, feature_config, baseline_ok, _HAS_SKLEARN,
    )

    return {
        "report": report,
        "feature_importance": feature_importance,
        "combustion_baseline": combustion_baseline,
        "sensitivity_matrix": {
            "pearson": corr_pearson, "spearman": corr_spearman,
            "mutual_info": mi_scores,
            "rf_impurity": rf_importance, "rf_permutation": rf_permutation,
        },
        "rpm_analysis": rpm_analysis,
        "load_zone_analysis": load_zone_analysis,
        "cv_stability": cv_stability,
        "partial_dependence": partial_dependence,
        "n_points": n_points,
        "detected_features": feature_keys,
        "has_sklearn": _HAS_SKLEARN,
    }


def _merge_feature_importance(pearson, spearman, mi, rf_imp, rf_perm, feature_config):
    """合并多种方法的特征重要性为统一排名。"""
    merged = {}
    for fk in feature_config:
        merged[fk] = {"feature": fk, "feature_cn": feature_config[fk],
                      "pearson_BSFC": None, "pearson_COV": None,
                      "spearman_BSFC": None, "spearman_COV": None,
                      "mi_BSFC": None, "mi_COV": None,
                      "rf_impurity_BSFC": None, "rf_impurity_COV": None,
                      "rf_permutation_BSFC": None, "rf_permutation_COV": None,
                      "aggregate_score": 0.0}

    for fk in feature_config:
        for target in ["BSFC", "COV"]:
            for corr_dict, prefix in [(pearson, "pearson"), (spearman, "spearman")]:
                if corr_dict.get(target) and fk in corr_dict[target]:
                    v = corr_dict[target][fk]
                    merged[fk][f"{prefix}_{target}"] = v
                    merged[fk]["aggregate_score"] += abs(v) * 0.15
            for mi_dict, prefix in [(mi, "mi")]:
                if mi_dict.get(target) and fk in mi_dict[target]:
                    merged[fk][f"{prefix}_{target}"] = mi_dict[target][fk]
                    merged[fk]["aggregate_score"] += mi_dict[target][fk] * 0.05
    for item in rf_imp:
        fk = item["feature"]
        merged[fk][f"rf_impurity_{item['target']}"] = item["importance"]
        merged[fk]["aggregate_score"] += item["importance"] * 0.35
    for item in rf_perm:
        fk = item["feature"]
        merged[fk][f"rf_permutation_{item['target']}"] = item["importance"]
        merged[fk]["aggregate_score"] += item["importance"] * 0.30

    result = sorted(merged.values(), key=lambda x: x["aggregate_score"], reverse=True)
    return result


def _build_ml_sensitivity_report(
    feature_importance, pearson, spearman, mi, rpm_analysis,
    load_zone_analysis, cv_stability, n_points, feature_keys, feature_config,
    baseline_ok, has_sklearn,
):
    """生成基于数据的敏感性分析结论，数据用表格展示。"""
    def _short(fk):
        m = {'ai50':'AI50','spark_act':'点火角','spark_delta':'退角',
             'knock':'爆震','vvt':'VVT','egt':'排温',
             'boost':'增压','imep':'IMEP'}
        return m.get(fk, fk)

    def _top(d, n=3):
        if not d: return []
        return sorted(d.items(), key=lambda x: abs(x[1]), reverse=True)[:n]

    def _tfmt(v):
        """格式化数值，加符号前缀。"""
        return f"{v:+.3f}"

    def _tbl(headers, rows):
        """生成 markdown 表格。"""
        lines = []
        sep = "| " + " | ".join(headers) + " |"
        lines.append(sep)
        lines.append("|" + "|".join(["------" for _ in headers]) + "|")
        for row in rows:
            lines.append("| " + " | ".join(str(c) for c in row) + " |")
        lines.append("")
        return "\n".join(lines)

    lines = []
    lines.append("# 燃烧参数敏感性分析结论")
    lines.append("")

    # ═══ 一、BSFC ═══
    lines.append("## 一、BSFC 主导因素")
    lines.append("")
    bsfc_items = _top(pearson.get("BSFC", {}), 8) if pearson.get("BSFC") else []

    if bsfc_items:
        rows = []
        for k, v in bsfc_items:
            flag = "▲▲" if abs(v) > 0.5 else ("▲" if abs(v) > 0.3 else "─")
            direction = "正" if v > 0 else "负"
            bar = "█" * min(int(abs(v) * 15), 15)
            rows.append([_short(k), flag, _tfmt(v), direction, bar])
        lines.append(_tbl(["参数", "强度", "r", "方向", "可视化"], rows))
        lines.append("")

    # 负荷分区
    if load_zone_analysis:
        lines.append("> 总体相关被负荷混淆，分负荷区后真实关系如下。注意大负荷区的符号与低负荷区完全相反。")
        lines.append("")
        headers = ["负荷区", "点数", "第1因子", "第2因子", "第3因子"]
        rows = []
        for zname in ["大负荷 (120-200 Nm)", "全负荷 (>=200 Nm)", "中负荷 (50-120 Nm)", "低负荷 (<50 Nm)"]:
            if zname not in load_zone_analysis: continue
            zdata = load_zone_analysis[zname]
            zc = zdata["correlations"].get("BSFC", {})
            top3 = _top(zc, 3)
            cells = [zname, str(zdata['n_points'])]
            for k, v in top3:
                cells.append(f"{_short(k)} {_tfmt(v)}")
            while len(cells) < 5:
                cells.append("-")
            rows.append(cells[:5])
        lines.append(_tbl(headers, rows))
        lines.append("")

    # 反转
    reversals = []
    for fk in feature_keys:
        signs = set()
        for zname, zdata in load_zone_analysis.items():
            v = zdata["correlations"].get("BSFC", {}).get(fk)
            if v is not None and abs(v) > 0.2:
                signs.add("+" if v > 0 else "-")
        if len(signs) > 1:
            reversals.append(_short(fk))
    if reversals:
        lines.append(f"> ⚠ **反转参数**: {', '.join(reversals)} — 不同负荷区影响方向相反，必须分区标定")
        lines.append("")

    # ═══ 二、COV ═══
    lines.append("## 二、COV 燃烧稳定性")
    lines.append("")
    cov_items = _top(pearson.get("COV", {}), 8) if pearson.get("COV") else []
    if cov_items:
        rows = []
        for k, v in cov_items:
            flag = "▲▲" if abs(v) > 0.5 else ("▲" if abs(v) > 0.3 else "─")
            rows.append([_short(k), flag, _tfmt(v)])
        lines.append(_tbl(["参数", "强度", "r"], rows))
        lines.append("")

        strong_cov = [(k,v) for k,v in cov_items if abs(v) > 0.5]
        if strong_cov:
            lines.append(f"> 燃烧稳定性瓶颈: {'、'.join(_short(k) for k,v in strong_cov[:3])}")
            lines.append("")

    # ═══ 三、RPM ═══
    if rpm_analysis:
        lines.append("## 三、转速分区")
        lines.append("")
        headers = ["转速段", "点数", "BSFC 主导", "COV 主导"]
        rows = []
        for band_name, data in rpm_analysis.items():
            bsfc_top = _top(data["correlations"].get("BSFC", {}), 2)
            cov_top = _top(data["correlations"].get("COV", {}), 2)
            bsfc_s = ", ".join(f"{_short(k)} {_tfmt(v)}" for k,v in bsfc_top if abs(v)>0.2)
            cov_s = ", ".join(f"{_short(k)} {_tfmt(v)}" for k,v in cov_top if abs(v)>0.2)
            rows.append([band_name, str(data['n_points']), bsfc_s or "-", cov_s or "-"])
        lines.append(_tbl(headers, rows))
        lines.append("")

    # ═══ 四、CV ═══
    if cv_stability:
        lines.append("## 四、分析可靠性")
        lines.append("")
        rows = []
        for cv in cv_stability:
            r2 = cv['cv_r2_mean']
            grade = "高" if r2 > 0.7 else ("中" if r2 > 0.4 else "低")
            rows.append([cv['target'], f"{r2:.3f}", f"{cv['cv_mse_mean']:.1f}±{cv['cv_mse_std']:.1f}", grade])
        lines.append(_tbl(["目标", "R²", "MSE", "可信度"], rows))
        lines.append("")

    # ═══ 五、建议 ═══
    lines.append("## 五、标定建议")
    lines.append("")
    suggestions = []

    if load_zone_analysis:
        big_load = load_zone_analysis.get("大负荷 (120-200 Nm)", {})
        bl_corr = big_load.get("correlations", {}).get("BSFC", {})
        bl_top = _top(bl_corr, 2)
        if bl_top:
            names = "、".join(_short(k) for k,v in bl_top if abs(v)>0.3)
            if names:
                suggestions.append(f"大负荷区 BSFC 优化重点: **{names}**")

        low_load = load_zone_analysis.get("低负荷 (<50 Nm)", {})
        ll_corr = low_load.get("correlations", {}).get("BSFC", {})
        ll_top = _top(ll_corr, 2)
        if ll_top:
            names = "、".join(_short(k) for k,v in ll_top if abs(v)>0.3)
            if names:
                suggestions.append(f"低负荷区 BSFC 优化重点: **{names}**（与大负荷区方向相反）")

    if cov_items:
        cov_strong = [(k,v) for k,v in cov_items if v > 0.5]
        if cov_strong:
            suggestions.append(f"燃烧稳定性瓶颈: **{'、'.join(_short(k) for k,v in cov_strong[:2])}**")

    if reversals:
        suggestions.append(f"必须分区标定: **{', '.join(reversals)}**")

    for i, s in enumerate(suggestions, 1):
        lines.append(f"{i}. {s}")
    lines.append("")

    return "\n".join(lines)


# ────────────────────────────────────────────────────────────
# 12. 一站式单发动机分析 (含燃烧特性)
# ────────────────────────────────────────────────────────────
def single_engine_full_analysis(
    filepath: str,
    encoding: str = 'gbk',
    header_rows: int = 5,
    skip_time_cols: int = 3,
    turbo_speed_limit: int = TURBO_SPEED_LIMIT_DEFAULT,
    altitude_m: Optional[float] = 3000,
    save_plot_performance: Optional[str] = None,
    save_plot_combustion: Optional[str] = None,
    standard_engine: Optional[str] = None,
) -> Dict:
    """一站式单发动机全分析: 性能 + 燃烧特性。

    自动检测并分析:
      - 性能: 扭矩/功率/BSFC/增压压力/WG开度/排温/涡轮转速
      - 燃烧: COV/AI50/点火角/点火退角/爆震/VVT/IMEP
      - 高原: 增压器高原能力评估
      - 标准对比: 当 standard_engine="B15HE" 时，对比 B15HE 标准数据

    Args:
        filepath: CSV 或 Excel 文件路径
        encoding: CSV 编码 (默认 gbk)
        header_rows: CSV 跳过表头行数
        skip_time_cols: CSV 跳过时间戳列数
        turbo_speed_limit: 增压器转速限制
        altitude_m: 评估海拔 (米)，None 则跳过
        save_plot_performance: 性能图表保存路径
        save_plot_combustion: 燃烧特性图表保存路径
        standard_engine: 标准发动机对标类型，如 "B15HE" (None=不对比)

    Returns:
        Dict 包含 performance, combustion, altitude, report
    """
    print(f"{'='*60}")
    print(" 发动机全分析 (性能 + 燃烧特性)")
    print(f"{'='*60}\n")

    # 加载
    ext = Path(filepath).suffix.lower()
    if ext == '.csv':
        df = load_csv(filepath, encoding=encoding,
                      header_rows=header_rows, skip_time_cols=skip_time_cols)
    else:
        df = load_excel(filepath)
    df = ensure_numeric(df)

    # 检测列
    rpm_col = detect_column(df, "rpm")
    torque_col = detect_column(df, "torque")
    if rpm_col is None or torque_col is None:
        raise ValueError(f"无法检测转速/扭矩列。可用列: {df.columns.tolist()}")
    print(f"转速列: {rpm_col}, 扭矩列: {torque_col}")

    col_map = detect_all_columns(df)
    col_map['rpm'] = rpm_col
    col_map['torque'] = torque_col
    print(f"检测到的列: {col_map}")

    # 燃烧特性分析 (包含性能指标)
    combustion_out = single_engine_combustion_analysis(
        df, rpm_col, torque_col, col_map,
        turbo_speed_limit=turbo_speed_limit,
        altitude_m=altitude_m,
        save_plot=save_plot_combustion,
    )

    # 性能分析 (复用已加载的数据，避免重复 I/O)
    performance_out = _single_engine_performance_core(
        df, rpm_col, torque_col, col_map,
        turbo_speed_limit=turbo_speed_limit,
        altitude_m=altitude_m,
        save_plot=save_plot_performance,
        standard_engine=standard_engine,
    )

    return {
        "performance": performance_out,
        "combustion": combustion_out,
        "report": combustion_out["report"],
    }


# ────────────────────────────────────────────────────────────
# B15HE 标准数据对比
# ────────────────────────────────────────────────────────────

# B15HE 标准数据文件路径 (SKILL assets 目录下的 Excel 文件)
_B15HE_STANDARD_PATH = Path(__file__).parent.parent / "assets" / "baseline_engine_database" / "260108_B15HE_BSFC_发动机标准数据_v1.0.xlsx"

# 外特性列索引映射 (0-based)
_B15HE_WOT_COL_MAP = {
    "rpm": 4,          # DynoSpeed_Avg
    "power": 5,        # BrakePower_Avg
    "torque": 6,       # DynoTorque_Avg
    "bsfc": 13,        # 油耗率 (BSFC_Avg)
    "fuel_flow": 12,   # 油耗量 (FuelMassFlowRate_Avg)
    "修正功率": 8,
    "修正扭矩": 9,
}


def load_b15he_standard(sheet: str = "外特性") -> Optional[pd.DataFrame]:
    """加载 B15HE 发动机标准数据。

    Args:
        sheet: "外特性" (WOT 全负荷曲线) 或 "B15HE万有数据" (万有特性)

    Returns:
        DataFrame 或 None (文件不存在时)
    """
    fp = _B15HE_STANDARD_PATH
    if not fp.exists():
        print(f"[WARNING] B15HE 标准数据文件不存在: {fp}")
        return None

    try:
        df = pd.read_excel(str(fp), sheet_name=sheet, header=0)

        # 清理列名 (去掉 \\XCP: 1 等后缀)
        df.columns = [str(c).split("\\")[0].strip() for c in df.columns]

        if sheet == "外特性":
            # 外特性: row 0 就是数据 (无单位行)
            return df
        elif sheet == "B15HE万有数据":
            # 万有数据: row 0 是单位行，需要跳过
            # 跳过第一行 (单位) 和后续的空行
            df = df.iloc[1:].copy()
            df = df.dropna(how="all")
            # 数值化关键列
            rpm_col = df.columns[4]  # DynoSpeed_Avg
            df[rpm_col] = pd.to_numeric(df[rpm_col], errors="coerce")
            df = df.dropna(subset=[rpm_col])
            return df
        else:
            raise ValueError(f"未知 sheet: {sheet}")
    except Exception as e:
        print(f"[WARNING] 加载 B15HE 标准数据失败 ({sheet}): {e}")
        return None


def _interp_at_rpm(std_rpm, std_val, target_rpm):
    """对标准数据进行线性插值，获取目标 RPM 点的标准值。"""
    if std_rpm is None or std_val is None:
        return None
    idx = np.searchsorted(std_rpm, target_rpm)
    if idx <= 0:
        return float(std_val[0])
    if idx >= len(std_rpm):
        return float(std_val[-1])
    if std_rpm[idx] == std_rpm[idx - 1]:
        return float(std_val[idx])
    ratio = (target_rpm - std_rpm[idx - 1]) / (std_rpm[idx] - std_rpm[idx - 1])
    return float(std_val[idx - 1] + ratio * (std_val[idx] - std_val[idx - 1]))


def compare_with_b15he_standard(
    test_rpm: np.ndarray,
    test_torque: np.ndarray,
    test_power: Optional[np.ndarray] = None,
    test_bsfc: Optional[np.ndarray] = None,
) -> Dict:
    """将测试数据的外特性 (WOT) 与 B15HE 标准对比。

    Args:
        test_rpm: 测试数据转速数组
        test_torque: 测试数据扭矩数组 (Nm)
        test_power: 测试数据功率数组 (kW), None 则自动计算
        test_bsfc: 测试数据 BSFC 数组 (g/kWh), None 则跳过

    Returns:
        dict: {
            "standard_rpm": [...],
            "standard_torque": [...],
            "standard_power": [...],
            "standard_bsfc": [...],
            "comparison_points": [{"rpm":..., "test_torque":..., "std_torque":..., "torque_diff":..., ...}, ...],
            "summary": {各项对比结论},
            "report": "文本报告"
        }
    """
    result = {"standard_engine": "B15HE"}
    std = load_b15he_standard("外特性")
    if std is None:
        result["report"] = "[B15HE 标准数据未加载，跳过对比]"
        return result

    # 提取标准数据
    std_rpm = pd.to_numeric(std.iloc[:, _B15HE_WOT_COL_MAP["rpm"]], errors="coerce").values
    std_torque = pd.to_numeric(std.iloc[:, _B15HE_WOT_COL_MAP["torque"]], errors="coerce").values
    std_power = pd.to_numeric(std.iloc[:, _B15HE_WOT_COL_MAP["power"]], errors="coerce").values
    std_bsfc = pd.to_numeric(std.iloc[:, _B15HE_WOT_COL_MAP["bsfc"]], errors="coerce").values

    valid_std = ~(np.isnan(std_rpm) | np.isnan(std_torque))
    std_rpm = std_rpm[valid_std]
    std_torque = std_torque[valid_std]
    std_power = std_power[valid_std]
    std_bsfc = std_bsfc[valid_std]

    result["standard_rpm"] = std_rpm.tolist()
    result["standard_torque"] = std_torque.tolist()
    result["standard_power"] = std_power.tolist()
    result["standard_bsfc"] = std_bsfc.tolist()

    # 排序标准数据 (按 RPM)
    sort_idx = np.argsort(std_rpm)
    std_rpm = std_rpm[sort_idx]
    std_torque = std_torque[sort_idx]
    std_power = std_power[sort_idx]
    std_bsfc = std_bsfc[sort_idx]

    # 筛选测试数据
    mask = (test_rpm > 0) & (test_torque > 0)
    tr = test_rpm[mask]
    tt = test_torque[mask]
    if test_power is not None:
        tp = test_power[mask]
    else:
        tp = tt * tr / 9549

    if test_bsfc is not None:
        tb = test_bsfc[mask]
    else:
        tb = None

    # 对每个测试 RPM 点，查找标准值 (插值)
    comparison_points = []
    for i in range(len(tr)):
        rpm_i = tr[i]
        std_tq = _interp_at_rpm(std_rpm, std_torque, rpm_i)
        std_pw = _interp_at_rpm(std_rpm, std_power, rpm_i)
        std_bsfc_i = _interp_at_rpm(std_rpm, std_bsfc, rpm_i)

        point = {
            "rpm": round(rpm_i, 0),
            "test_torque": round(tt[i], 1),
            "std_torque": round(std_tq, 1) if std_tq is not None else None,
            "torque_diff": round(tt[i] - std_tq, 1) if std_tq is not None else None,
            "torque_diff_pct": round((tt[i] - std_tq) / std_tq * 100, 1) if std_tq is not None and std_tq != 0 else None,
            "test_power": round(tp[i], 1),
            "std_power": round(std_pw, 1) if std_pw is not None else None,
            "power_diff": round(tp[i] - std_pw, 1) if std_pw is not None else None,
        }
        if tb is not None and std_bsfc_i is not None:
            point["test_bsfc"] = round(tb[i], 1)
            point["std_bsfc"] = round(std_bsfc_i, 1)
            point["bsfc_diff"] = round(tb[i] - std_bsfc_i, 1)
        comparison_points.append(point)

    result["comparison_points"] = comparison_points

    # ── 总结 ──
    summary = {}
    torque_diffs = [p["torque_diff"] for p in comparison_points if p["torque_diff"] is not None]
    power_diffs = [p["power_diff"] for p in comparison_points if p["power_diff"] is not None]

    if torque_diffs:
        summary["torque"] = {
            "mean_diff": round(np.mean(torque_diffs), 1),
            "max_gain": round(max(torque_diffs), 1),
            "max_loss": round(min(torque_diffs), 1),
        }
        avg_tq_diff_pct = np.mean([p["torque_diff_pct"] for p in comparison_points
                                    if p["torque_diff_pct"] is not None])
        summary["torque"]["avg_diff_pct"] = round(avg_tq_diff_pct, 1)

    if power_diffs:
        summary["power"] = {
            "mean_diff": round(np.mean(power_diffs), 1),
            "max_gain": round(max(power_diffs), 1),
            "max_loss": round(min(power_diffs), 1),
        }

    # BSFC 对比
    bsfc_diffs = [p.get("bsfc_diff") for p in comparison_points
                  if p.get("bsfc_diff") is not None]
    if bsfc_diffs:
        summary["bsfc"] = {
            "mean_diff": round(np.mean(bsfc_diffs), 1),
            "n_test_lower": sum(1 for d in bsfc_diffs if d < 0),  # 负值 = 测试油耗更低 (更好)
            "n_test_higher": sum(1 for d in bsfc_diffs if d > 0),  # 正值 = 测试油耗更高 (更差)
        }

    result["summary"] = summary

    # ── 生成报告 ──
    lines = [
        "=" * 60,
        " B15HE 标准数据对比",
        "=" * 60,
        "",
        "外特性 (WOT) 对比: 测试数据 vs B15HE 标准",
        "",
    ]
    header = f"{'RPM':>8} | {'测试扭矩':>8} | {'标准扭矩':>8} | {'差值':>8} | {'差%':>7} | {'测试功率':>8} | {'标准功率':>8} | {'功率差':>8}"
    if bsfc_diffs:
        header += f" | {'测试BSFC':>10} | {'标准BSFC':>10} | {'BSFC差':>9}"
    lines.append(header)

    lines.append("-" * (80 if not bsfc_diffs else 126))
    for p in comparison_points:
        line = f"{p['rpm']:>8.0f} | {p['test_torque']:>8.1f} | {p['std_torque'] or '-':>8} | {p['torque_diff'] or '-':>8} | {p['torque_diff_pct'] or '-':>7}"
        line += f" | {p['test_power']:>8.1f} | {p['std_power'] or '-':>8} | {p['power_diff'] or '-':>8}"
        if "test_bsfc" in p and p["test_bsfc"] is not None:
            line += f" | {p['test_bsfc']:>10.1f} | {p.get('std_bsfc', 0):>10.1f} | {p.get('bsfc_diff', 0):>+9.1f}"
        lines.append(line)

    lines.append("")
    lines.append("--- 综合结论 ---")
    if "torque" in summary:
        s = summary["torque"]
        lines.append(f"  扭矩: 平均 {s['mean_diff']:+.1f} Nm ({s['avg_diff_pct']:+.1f}%), "
                     f"最大增益 {s['max_gain']:+.1f} Nm, 最大损失 {s['max_loss']:+.1f} Nm")
    if "power" in summary:
        s = summary["power"]
        lines.append(f"  功率: 平均 {s['mean_diff']:+.1f} kW, "
                     f"最大增益 {s['max_gain']:+.1f} kW, 最大损失 {s['max_loss']:+.1f} kW")
    if "bsfc" in summary:
        s = summary["bsfc"]
        lines.append(f"  BSFC: 平均 {s['mean_diff']:+.1f} g/kWh "
                     f"(负值 = 测试油耗低于标准，更好)")
        if s["n_test_higher"] > len(bsfc_diffs) * 0.5:
            lines.append(f"  ⚠️ 注意: 大部分转速点测试油耗高于标准值 (需优化)")
        elif s["n_test_lower"] > len(bsfc_diffs) * 0.5:
            lines.append(f"  ✅ 优秀: 大部分转速点测试油耗低于标准值")

    lines.append("")

    result["report"] = "\n".join(lines)
    return result




# ────────────────────────────────────────────────────────────
# 通用标准数据对比框架
# ────────────────────────────────────────────────────────────

def load_standard_data(filepath: str, sheet_name: str = None,
                       col_map: Optional[Dict[str, int]] = None
                       ) -> Optional[pd.DataFrame]:
    """加载发动机标准数据文件（通用版本）。

    Args:
        filepath: 标准数据 Excel 路径
        sheet_name: Sheet 名，None 则自动检测
        col_map: 列索引映射 {rpm/torque/power/bsfc: index}，None 则自动检测

    Returns:
        DataFrame 或 None（加载失败时）

    用法:
        # 加载指定发动机标准数据
        std = load_standard_data("发动机标准数据.xlsx")
        std = load_standard_data("标准.xlsx", sheet_name="外特性",
                                 col_map={"rpm": 4, "torque": 9, "power": 8, "bsfc": 13})
    """
    p = Path(filepath)
    if not p.exists():
        print(f"[WARNING] 标准数据文件不存在: {filepath}")
        return None

    try:
        df = pd.read_excel(str(p), sheet_name=sheet_name, header=0)
        # 清理列名后缀
        df.columns = [str(c).split("\\")[0].strip() for c in df.columns]
        return df
    except Exception as e:
        print(f"[WARNING] 加载标准数据失败: {e}")
        return None


def compare_with_standard(
    test_rpm: np.ndarray,
    test_torque: np.ndarray,
    test_power: Optional[np.ndarray] = None,
    test_bsfc: Optional[np.ndarray] = None,
    standard_df: Optional[pd.DataFrame] = None,
    col_map: Optional[Dict[str, str]] = None,
    name: str = "测试",
    standard_name: str = "标准",
) -> Dict:
    """通用发动机外特性（WOT）标准数据对比。

    对比测试数据与标准数据的扭矩/功率/BSFC，输出各转速点的
    绝对差值、百分比差值，以及综合结论。

    Args:
        test_rpm: 测试数据转速数组
        test_torque: 测试数据扭矩数组 (Nm)
        test_power: 测试数据功率数组 (kW)，None 则跳过
        test_bsfc: 测试数据 BSFC 数组 (g/kWh)，None 则跳过
        standard_df: 标准数据 DataFrame
        col_map: 标准数据列名映射，如 {"rpm":"DynoSpeed_Avg","torque":"DynoTorque_Avg"}
                默认自动检测列名
        name: 测试方名称 (默认 "测试")
        standard_name: 标准方名称 (默认 "标准")

    Returns:
        dict: comparison_points, summary, report

    用法:
        std = load_standard_data("标准.xlsx")
        result = compare_with_standard(test_rpm, test_torque, test_power, test_bsfc, std)
        print(result["report"])
    """
    if standard_df is None:
        return {"report": f"[{standard_name} 标准数据未加载，跳过对比]"}

    # 检测标准数据列
    if col_map:
        std_rpm_col = col_map.get("rpm")
        std_torque_col = col_map.get("torque")
        std_power_col = col_map.get("power")
        std_bsfc_col = col_map.get("bsfc")
    else:
        std_rpm_col = detect_column(standard_df, "rpm")
        std_torque_col = detect_column(standard_df, "torque")
        std_power_col = detect_column(standard_df, "power")
        std_bsfc_col = detect_column(standard_df, "bsfc")

    if std_rpm_col is None or std_torque_col is None:
        return {"report": f"[错误: 无法从标准数据中检测转速/扭矩列，跳过对比]"}

    # 提取标准数据
    std_rpm = pd.to_numeric(standard_df[std_rpm_col], errors="coerce").values
    std_torque = pd.to_numeric(standard_df[std_torque_col], errors="coerce").values
    std_power = (pd.to_numeric(standard_df[std_power_col], errors="coerce").values
                 if std_power_col else None)
    std_bsfc = (pd.to_numeric(standard_df[std_bsfc_col], errors="coerce").values
                if std_bsfc_col else None)

    # 清理无效数据
    valid = ~(np.isnan(std_rpm) | np.isnan(std_torque))
    std_rpm = std_rpm[valid]
    std_torque = std_torque[valid]
    if std_power is not None:
        std_power = std_power[valid]
    if std_bsfc is not None:
        std_bsfc = std_bsfc[valid]

    # 按 RPM 排序
    idx = np.argsort(std_rpm)
    std_rpm = std_rpm[idx]; std_torque = std_torque[idx]
    if std_power is not None: std_power = std_power[idx]
    if std_bsfc is not None: std_bsfc = std_bsfc[idx]

    # 筛选测试数据
    mask = (test_rpm > 0) & (test_torque > 0)
    tr = test_rpm[mask]; tt = test_torque[mask]
    tp = (test_power[mask] if test_power is not None
          else tt * tr / 9549)    # 功率自动计算
    tb = test_bsfc[mask] if test_bsfc is not None else None

    # 逐点对比（插值）
    comparison_points = []
    for i in range(len(tr)):
        rpm_i = tr[i]
        std_tq = _interp_at_rpm(std_rpm, std_torque, rpm_i)
        std_pw = _interp_at_rpm(std_rpm, std_power, rpm_i) if std_power is not None else None
        std_bsfc_i = _interp_at_rpm(std_rpm, std_bsfc, rpm_i) if std_bsfc is not None else None

        point = {
            "rpm": round(rpm_i, 0),
            "test_torque": round(tt[i], 1),
            "std_torque": round(std_tq, 1) if std_tq else None,
            "torque_diff": round(tt[i] - std_tq, 1) if std_tq else None,
            "torque_diff_pct": round((tt[i] - std_tq) / std_tq * 100, 1)
                if std_tq and std_tq != 0 else None,
            "test_power": round(tp[i], 1),
            "std_power": round(std_pw, 1) if std_pw else None,
            "power_diff": round(tp[i] - std_pw, 1) if std_pw else None,
        }
        if tb is not None and std_bsfc_i is not None:
            point["test_bsfc"] = round(tb[i], 1)
            point["std_bsfc"] = round(std_bsfc_i, 1)
            point["bsfc_diff"] = round(tb[i] - std_bsfc_i, 1)
        comparison_points.append(point)

    # 汇总
    summary = {}
    torque_diffs = [p["torque_diff"] for p in comparison_points if p["torque_diff"] is not None]
    power_diffs = [p["power_diff"] for p in comparison_points if p["power_diff"] is not None]

    if torque_diffs:
        summary["torque"] = {
            "mean_diff": round(np.mean(torque_diffs), 1),
            "max_gain": round(max(torque_diffs), 1),
            "max_loss": round(min(torque_diffs), 1),
            "avg_diff_pct": round(np.mean([p["torque_diff_pct"] for p in comparison_points
                                            if p["torque_diff_pct"] is not None]), 1),
        }
    if power_diffs:
        summary["power"] = {
            "mean_diff": round(np.mean(power_diffs), 1),
            "max_gain": round(max(power_diffs), 1),
            "max_loss": round(min(power_diffs), 1),
        }

    bsfc_diffs = [p.get("bsfc_diff") for p in comparison_points if p.get("bsfc_diff") is not None]
    if bsfc_diffs:
        summary["bsfc"] = {
            "mean_diff": round(np.mean(bsfc_diffs), 1),
            "n_test_lower": sum(1 for d in bsfc_diffs if d < 0),
            "n_test_higher": sum(1 for d in bsfc_diffs if d > 0),
        }

    # 报告
    lines = [
        "=" * 60,
        f" {standard_name} 标准数据对比",
        "=" * 60, "",
        f"外特性 (WOT) 对比: {name} vs {standard_name}", "",
    ]
    has_header = True
    for p in comparison_points:
        if has_header:
            h = f"{'RPM':>8} | {'扭矩测试':>8} | {'扭矩标准':>8} | {'差值':>8} | {'差%':>7}"
            h += f" | {'功率测试':>8} | {'功率标准':>8} | {'功率差':>8}"
            if "test_bsfc" in p:
                h += f" | {'BSFC测试':>10} | {'BSFC标准':>10} | {'BSFC差':>9}"
            lines.append(h)
            lines.append("-" * (88 if "test_bsfc" not in p else 126))
            has_header = False

        line = (f"{p['rpm']:>8.0f} | {p['test_torque']:>8.1f} | "
                f"{p['std_torque'] or '-':>8} | {p['torque_diff'] or '-':>8} | "
                f"{p['torque_diff_pct'] or '-':>7} | {p['test_power']:>8.1f} | "
                f"{p['std_power'] or '-':>8} | {p['power_diff'] or '-':>8}")
        if "test_bsfc" in p and p["test_bsfc"] is not None:
            line += (f" | {p['test_bsfc']:>10.1f} | {p.get('std_bsfc', 0):>10.1f} | "
                     f"{p.get('bsfc_diff', 0):>+9.1f}")
        lines.append(line)

    lines.append("")
    lines.append("--- 综合结论 ---")
    if "torque" in summary:
        s = summary["torque"]
        lines.append(f"  扭矩: 平均 {s['mean_diff']:+.1f} Nm ({s['avg_diff_pct']:+.1f}%), "
                     f"最大增益 {s['max_gain']:+.1f} Nm, 最大损失 {s['max_loss']:+.1f} Nm")
    if "power" in summary:
        s = summary["power"]
        lines.append(f"  功率: 平均 {s['mean_diff']:+.1f} kW, "
                     f"最大增益 {s['max_gain']:+.1f} kW, 最大损失 {s['max_loss']:+.1f} kW")
    if "bsfc" in summary:
        s = summary["bsfc"]
        lines.append(f"  BSFC: 平均 {s['mean_diff']:+.1f} g/kWh "
                     f"(负值={name}油耗更低，更好)")
        verdict = "优秀" if s["n_test_lower"] > len(bsfc_diffs) * 0.5 else "需关注"
        lines.append(f"  {'✅' if verdict == '优秀' else '⚠️'} {verdict}: "
                     f"{s['n_test_lower']}/{len(bsfc_diffs)} 个点油耗低于标准")

    lines.append("")
    return {
        "standard_engine": standard_name,
        "comparison_points": comparison_points,
        "summary": summary,
        "report": "\n".join(lines),
    }


# ─── 修改 single_engine_analysis 以支持标准对比 ───

# (在函数调用后补充标准对比，通过外层包装函数实现)


def _append_standard_comparison_to_report(
    original_report: str,
    comparison: Dict,
) -> str:
    """将标准对比报告追加到原报告末尾。"""
    if "report" not in comparison or not comparison["report"]:
        return original_report
    if comparison["report"].startswith("[B15HE"):
        return original_report + "\n" + comparison["report"]
    return original_report + "\n\n" + comparison["report"]


# ────────────────────────────────────────────────────────────
# 辅助工具
# ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 4:
        print("用法: python engine_analysis.py <file.xlsx> <名称A> <名称B> [行数]")
        print("示例: python engine_analysis.py data.xlsx 方案A 方案B 9")
        sys.exit(1)

    fp = sys.argv[1]
    name_a = sys.argv[2]
    name_b = sys.argv[3]
    n_pts = int(sys.argv[4]) if len(sys.argv) > 4 else None

    full_analysis(fp, name_a, name_b, n_points=n_pts)
