---
name: engine-data-analysis
description: Analyze engine dyno/bench test data — combustion characteristics (COV/AI50/spark/knock/VVT/IMEP), performance parameters (torque/power/BSFC/boost/EGT), fuel consumption analysis, turbocharger matching comparison, high-altitude capability assessment, data visualization, and standard data benchmarking. Use when the user asks about engine performance data, combustion analysis, dyno data, turbocharger comparison, BSFC analysis, or any .xlsx/.csv engine test files.
---

# 发动机数据分析

分析发动机台架测试数据（dyno data）—— 燃烧特性（COV/AI50/点火角/爆震/VVT/IMEP）、性能参数（扭矩/功率/BSFC/增压压力/排气温度）、油耗分析、增压器匹配对比、高原能力评估、标准数据对标、数据可视化。

## 触发条件

当用户提及以下任何内容时，应加载本 skill：
- **发动机 / engine / 台架 / dyno / 测试数据**
- **燃烧 / 燃烧分析 / COV / 循环变动 / AI50 / CA50 / 燃烧相位**
- **点火角 / 点火提前角 / 点火退角 / MBT / 爆震 / Knock**
- **VVT / VCT / 可变气门正时 / IMEP / 平均有效压力**
- **增压器 / turbo / 涡轮 / 增压器匹配**
- **BSFC / 燃油消耗率 / 油耗 / 扭矩 / 功率**
- **增压压力 / 排气温度 / 背压 / WG 开度**
- **高原 / 高海拔 / 万有特性 / 对标 / 标准对比**
- **分析 .xlsx / .csv 格式的测试数据**

## 快速入口

核心分析逻辑在 `scripts/engine_analysis.py`，通过相对路径导入：

```python
from pathlib import Path
import sys
sys.path.insert(0, str(Path.home() / ".hermes/skills/data-science/engine-data-analysis/scripts"))
from engine_analysis import *
```

### 增压器 A/B 对比（一键分析）

```python
out = full_analysis(
    filepath="增压器对比数据.xlsx",
    name_a="方案A", name_b="方案B",
    n_points=9,               # 每组行数（不指定则自动推断）
    sheet_name="Sheet3",
    skiprows=0,               # 如果有单位行则设为 1
    turbo_speed_limit=250000, # 增压器转速限制，需用户确认
    altitude_m=3000,          # 高原评估海拔，设为 None 则跳过
    save_plot="/tmp/comparison.png",
)
print(out["report"])
```

### 单发动机全分析（非对比）

```python
out = single_engine_full_analysis(
    filepath="发动机万有数据.csv",
    encoding="gbk",            # CSV 编码
    header_rows=5,             # 跳过的元数据行数
    turbo_speed_limit=250000,
    altitude_m=3000,
    save_plot_performance="/tmp/performance.png",
    save_plot_combustion="/tmp/combustion.png",
)
print(out["report"])
```

### 机器学习分析（燃烧敏感性分析）

当数据包含燃烧相关信号并需要分析各参数之间的敏感性关系时：

```python
out = analyze_combustion_sensitivity(
    filepath="B15HTC万有数据.csv",
    encoding="gbk", header_rows=5,
    save_plots="/tmp/sensitivity/"
)
```

详见下方"燃烧特性分析"章节。

## 分析工作流

### Step 1: 理解数据结构

数据通常来自台架测试文件，常见格式：

| 格式 | 典型特征 | 读取方法 |
|------|---------|---------|
| `.xlsx` | 上下排列 A/B 两组值，含单位行 | `load_excel(fp, sheet, skiprows=1)` |
| `.csv` (GBK 编码) | 5~8 行表头信息，含时间戳列 | `load_csv(fp, encoding='gbk', header_rows=5)` |
| `.csv` (万有特性) | 单发动机多负荷点台架数据 | `load_csv(fp, encoding='gbk', header_rows=5)` |

先快速查看数据结构：
```python
df = load_excel("数据.xlsx")
print_data_structure(df)
```

### Step 2: 列名自动检测

数据集列名来源多样（中英文混排、含换行符），用 `detect_all_columns()` 自动匹配：

```python
cols = detect_all_columns(df)
# 返回: {'rpm': '转速', 'torque': '修正扭矩', 'bsfc': 'BSFC_Avg', 'cov': 'IMEP1CO_Avg', ...}
```

> ⚠️ **扭矩/功率优先使用修正值** — `detect_column()` 会优先匹配"修正扭矩"、"修正功率"（`CorrTorqueEWG`、`CorrBrkPwrEWG` 等）。如果数据中同时包含原始值和修正值，会优先选择修正值。
>
> 若需要手动指定：
> ```python
> torque_col = detect_column(df, "torque")  # 自动优先修正值
> torque_col = "DynoTorque_Avg"             # 手动指定原始值
> ```

### Step 3: 数据分隔（仅 A/B 对比）

A/B 两组数据上下排列时：

```python
df_a, df_b = split_groups(df, n_points=9)        # 明确指定行数
df_a, df_b = split_groups(df)                     # 或自动推断
```

### Step 4: 增压器对比分析

`compare_turbochargers()` 从 7 个维度加权评分：

| 维度 | 权重 | 说明 |
|------|------|------|
| 低速扭矩（1000-1500rpm） | ×2.0 | 起步响应 |
| BSFC 燃油经济性 | ×2.0 | 越低越好，自动排除 1000rpm 异常点 |
| 涡轮转速余量（≥4000rpm） | ×1.5 | 余量大则安全 |
| WG 开度（≥3000rpm） | ×1.5 | 越小废气利用率越高 |
| 排气温度 | ×1.0 | 越低越好 |
| 峰值扭矩 | ×1.0 | 越高越好 |

```python
results = compare_turbochargers(df_a, df_b, "方案A", "方案B")
print(results["scores"])    # {'方案A': 8.5, '方案B': 6.0}
print(results["winner"])    # '方案A'
```

### Step 5: 高原能力评估

关键公式：
```
Speed_alt = Speed_0 × √(P0 / P_alt)
P_alt = 101.325 × (1 - 0.0065 × h / 288.15) ^ 5.255
```

```python
ha = assess_high_altitude(df_a, df_b, "方案A", "方案B", altitude_m=3000)
print(ha["方案A"]["safety"])   # "✅ 安全" / "⚠️ 可接受" / "❌ 高风险"
```

### Step 6: 标准数据对比

如果有发动机标准数据文件，加载后进行对标分析，从 **性能（扭矩/功率/BSFC）**、**燃烧（COV/AI50/点火角）**、**油耗** 等方面对比：

```python
# 方法一：使用通用标准对比框架（推荐）
std = load_standard_data("发动机标准数据.xlsx", sheet_name="外特性")
result = compare_with_standard(
    test_rpm, test_torque, test_power, test_bsfc,
    standard_df=std, name="测试", standard_name="标准",
)
print(result["report"])

# 方法二：指定列映射（标准数据列名与测试数据不同时）
result = compare_with_standard(
    test_rpm, test_torque, test_power, test_bsfc,
    standard_df=std,
    col_map={"rpm": "DynoSpeed_Avg", "torque": "DynoTorque_Avg"},
)
```

### Step 7: 燃烧特性分析

当数据包含燃烧相关信号时，使用 `single_engine_combustion_analysis()` 或 `analyze_combustion_sensitivity()` 一站式分析：

```python
# 先检测所有列
col_map = detect_all_columns(df)
# col_map 包含: cov, ai50, spark_act, spark_mbt, spark_delta, knock, vvt, imep 等

# 燃烧特性分析
out = single_engine_combustion_analysis(
    df, rpm_col, torque_col, col_map,
    turbo_speed_limit=250000, altitude_m=3000,
    save_plot="/tmp/combustion.png",
)
print(out["report"])

# 燃烧敏感性分析（机器学习）
out = analyze_combustion_sensitivity(
    filepath="发动机数据.csv",
    encoding="gbk", header_rows=5,
    save_plots="/tmp/sensitivity/",
)
print(out["report"])
```

### Step 8: 可视化

```python
# A/B 对比图（2×3 子图）
plot_comparison(results, save_path="/tmp/comparison.png")

# 燃烧特性图（3×3 子图）
_plot_combustion_analysis(group_data, save_path="/tmp/combustion.png")
```

### Step 9: 报告生成

```python
# A/B 对比报告
report = generate_text_report(results, altitude_results)

# 燃烧特性报告
report = _build_combustion_report(summary, altitude)

# 打印即可
print(report)
```

## 标准数据对比

### A/B 对比场景中嵌入标准对比

在 A/B 增压器对比分析完成后，如果存在标准数据，可以额外增加标准对比分析：

```python
# 1. A/B 对比
results = compare_turbochargers(df_a, df_b, "方案A", "方案B")

# 2. 与标准数据对比（分别对 A 和 B 做标准对标）
std = load_standard_data("发动机标准数据.xlsx")

# 对比 A 组
result_a = compare_with_standard(rpm, torque_a, power_a, bsfc_a, std, name="方案A")
# 对比 B 组
result_b = compare_with_standard(rpm, torque_b, power_b, bsfc_b, std, name="方案B")

print(result_a["report"])
print(result_b["report"])
```

### 单发动机场景中的标准对比

在 `single_engine_analysis()` 或 `single_engine_full_analysis()` 中，通过 `standard_engine` 参数触发：

```python
# 如果标准数据文件名包含"标准"、"standard"、"对标"等关键词且与数据在同一目录，
# 可以通过 standard_engine 参数指定名称
out = single_engine_full_analysis(
    filepath="测试数据.csv",
    standard_engine="对标发动机",  # 仅作为对比报告中显示的标签
)
```

### 对比维度

| 维度 | 指标 | 说明 |
|------|------|------|
| 性能 | 扭矩 (Nm)、功率 (kW)、BSFC (g/kWh) | 外特性 WOT 逐点对比 |
| 燃烧 | COV、AI50、点火角 | 万有特性工况图对比（如有） |
| 油耗 | BSFC、燃油消耗量 | 经济区分布对比 |

### 标准数据文件约定

标准数据可以是 Excel 或 CSV 格式，需包含转速、扭矩列（推荐也包含功率、BSFC 列）。列名通过 `col_map` 参数映射，或使用 `detect_column()` 自动检测。

## 燃烧特性分析

### 自动检测的燃烧信号

| 信号类型 | COLUMN_PATTERNS 键名 | 说明 |
|---------|---------------------|------|
| COV | `cov` | 循环变动系数，<3% 稳定，>5% 不稳定 |
| AI50 (CA50) | `ai50` | 燃烧相位，最佳 6-12°CA ATDC |
| 实际点火角 | `spark_act` | 实际点火提前角 (°BTDC) |
| MBT 点火角 | `spark_mbt` | MBT 点火角 (°BTDC) |
| 点火退角 | `spark_delta` | 从 MBT 退角 (°CA)，>5° 可能受爆震限制 |
| 爆震 | `knock` | 爆震窗口 / 爆震强度 |
| VVT | `vvt` | 可变气门正时（进/排气） |
| 油耗量 | `fuel_flow` | 燃油消耗量 (kg/h) |
| IMEP | `imep` | 平均有效压力 (bar) |

### 燃烧特性可视化（3×3 子图）

| 位置 | 内容 | 说明 |
|------|------|------|
| (0,0) | 功率分布 | 全工况功率散点 |
| (0,1) | BSFC 经济区地图 | 标最低 BSFC 点 |
| (0,2) | COV 燃烧稳定性 | 含 <3% 稳定线参考 |
| (1,0) | AI50 (CA50) | 燃烧相位分布 |
| (1,1) | 点火角 | 实际点火提前角 |
| (1,2) | 点火退角 | 退角分布，诊断爆震限制 |
| (2,0) | BSFC vs AI50 | 燃烧相位对油耗影响 |
| (2,1) | 点火角按负荷 | 不同负荷点火策略 |
| (2,2) | COV vs IMEP | 负荷对稳定性影响 |

### 燃烧敏感性分析（机器学习）

`analyze_combustion_sensitivity()` 使用随机森林回归分析各参数对 BSFC 和 COV 的敏感度：

```python
from engine_analysis import analyze_combustion_sensitivity

out = analyze_combustion_sensitivity(
    filepath="发动机数据.csv",
    encoding="gbk", header_rows=5,
    save_plots="/tmp/sensitivity/",
)
# out["feature_importance"] — 各参数对 BSFC 的影响权重
# out["report"] — 敏感性分析报告
```

典型发现示例：
| 参数 | BSFC 权重 | COV 权重 | 结论 |
|------|-----------|---------|------|
| AI50 (CA50) | 0.35 | 0.12 | 显著影响 BSFC |
| 排气温度 | 0.20 | 0.05 | 反映燃烧品质 |
| 点火角 | 0.15 | 0.25 | 显著影响稳定性 |
| 增压压力 | 0.12 | 0.08 | 影响中等 |

## 关键公式与注意事项

### ✅ 增压比公式
```
增压比 = 增压压力(kPa) / 标准大气压(101.325 kPa)
```
> ❌ 不是 增压压力 / 排气背压

### ✅ 高原转速推算
```
Speed_alt = Speed_0 × √(P0 / P_alt)
```
> 使用前需用户确认增压器限制值和评估海拔

### ⚠️ 修正扭矩/修正功率
台架数据中可能同时存在原始值（如 `DynoTorque`）和修正值（如 `修正扭矩`、`CorrTorqueEWG`）。优先使用**修正值**，因为它已经过环境校正，更能反映发动机真实性能。

### ⚠️ BSFC 异常值
1000rpm 的 BSFC 值可能异常高（如 494 g/kWh vs 351 g/kWh），分析时自动排除 1000rpm 数据点。

### ⚠️ 各指标参考阈值

| 指标 | 优秀 | 良好 | 一般 | 差 |
|------|------|------|------|----|
| WG 开度 | < 10% | 10-20% | 20-30% | > 30% |
| COV 循环变动 | < 2% | 2-3% | 3-5% | > 5% |
| AI50 燃烧相位 | 6-12° | 12-15° | 15-20° | > 20° 或 < 5° |
| 点火退角 | < 3° | 3-5° | 5-8° | > 8° |
| 增压器转速余量 | > 50krpm | 30-50k | 15-30k | < 15k |
| 最低 BSFC | < 230 | 230-250 | 250-270 | > 270 g/kWh |
| 低速扭矩 (1000rpm) | > 100 Nm | 80-100 | 60-80 | < 60 |

## CSV 文件处理要点

台架 CSV 常见陷阱及处理方式：

```python
# 标准读取方式（跳过表头+时间列）
df = load_csv("数据.csv", encoding="gbk", header_rows=5, skip_time_cols=3)

# 若 GBK 报错，尝试 latin-1（数值不受影响）
df = load_csv("数据.csv", encoding="latin-1", header_rows=5)
```

**典型多行表头结构：**

| 行号 | 内容 |
|-----|------|
| 1~4 | 元数据（Logger description 等） |
| 5 | 列名（Time, DynoSpeed_Avg, ...） |
| 6 | 单位行（ms, rev/min, Nm） |
| 7 | 数据类型（Raw, Average） |
| 8+ | 实际数据 |

**其他要点：**
- **重名列** — ETAS INCA 输出的 CSV 可能有多列同名（如 3 个 Time），pandas 自动加 `.1` `.2` 后缀
- **1.#QNAN 值** — `ensure_numeric()` 自动转为 NaN
- **合并单元格** — 用 `df.ffill()` 处理

## 标准大气压参考

| 海拔 | 大气压(kPa) | 与海平面比值 |
|------|-------------|-------------|
| 海平面 0m | 101.325 | 1.000 |
| 1000m | 89.9 | 0.887 |
| 2000m | 79.5 | 0.785 |
| 3000m | 70.1 | 0.692 |
| 4000m | 61.6 | 0.608 |

## 列名检测参考

`engine_analysis.py` 内置的 `COLUMN_PATTERNS` 覆盖以下信号类型（按优先级排列）：

| 信号类型 | 键名 | 关键匹配词（第一优先级） |
|----------|------|------------------------|
| 转速 | `rpm` | 转速, rpm, SPEED, DynoSpeed |
| **扭矩（修正优先）** | **`torque`** | **修正扭矩, CorrTorque, CorrTorqueEWG > 扭矩, Torque** |
| **功率（修正优先）** | **`power`** | **修正功率, CorrBrkPwr, CorrBrkPwrEWG > 功率, Power** |
| BSFC | `bsfc` | BSFC, 燃油消耗率, FuelCOSP |
| 增压器转速 | `turbo_speed` | 增压器转速, TURBOSPEED, Trbch_N |
| 增压压力 | `boost` | 增压压力, Boost, BSTC_pActBoostPress, VBOOST, P3 |
| 排气温度 | `egt` | 排气温度, EGT, T_EXH, EXHT_tMnfdTemp |
| 背压 | `backpressure` | 背压, FT_TACT, P_EXH |
| WG 开度 | `wg` | WG开度, EWGC_rActlPos |
| COV | `cov` | COV, IMEPCOV, IMEP1CO |
| AI50 | `ai50` | AI50, CA50, MFB50 |
| 点火角 | `spark_act` | SPK_dgActSpkAdv, 点火角 |
| MBT | `spark_mbt` | SPK_dgMBTSpkAdv, MBT |
| 退角 | `spark_delta` | SPK_dgDltFromMBT, 退角 |
| 爆震 | `knock` | Knock, knockWnd |
| VVT | `vvt` | VVT, Cam, CamPhs |
| 油耗量 | `fuel_flow` | Fuel_FuelConsume, FuelMassFlow |
| IMEP | `imep` | IMEP, 平均有效压力 |
| 进气流量 | `airflow` | 进气流量, AirFlow, AFS_dm |
| 进气温度 | `intake_temp` | 进气温度, T_Intake, T_ACS |

> ETAS INCA 信号名通常带 `_Avg` 后缀（如 `DynoSpeed_Avg`、`IMEP1CO_Avg`），列名检测已覆盖 `_Avg` 变体。

## CLI 快速调试

```bash
python ~/.hermes/skills/data-science/engine-data-analysis/scripts/engine_analysis.py \
    "数据文件.xlsx" "方案A" "方案B" 9
```

## 常见问题

### Q: 列名检测不到怎么办？
```python
results = compare_turbochargers(df_a, df_b, rpm_col="我的转速列", torque_col="我的扭矩列")
```

### Q: CSV 读取乱码怎么办？
先试 `encoding='gbk'`，再试 `encoding='latin-1'`。数值不受编码影响。

### Q: 只想对比标准数据？
```python
std = load_standard_data("标准.xlsx")
result = compare_with_standard(test_rpm, test_torque, standard_df=std)
print(result["report"])
```

### Q: 如何计算增压比？
```python
pr = calculate_pressure_ratio(boost_values, altitude_m=3000)
```

### Q: 如何导出报告？
```python
report = generate_text_report(results, altitude_results)
print(report)          # 打印到控制台
with open("报告.md", "w") as f:
    f.write(report)    # 保存到文件
```

## 参考文件

| 文件 | 说明 |
|------|------|
| `scripts/engine_analysis.py` | 核心分析模块 |
| `references/etas_inca_signals.md` | ETAS INCA 信号命名规范 |
