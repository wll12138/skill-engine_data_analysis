# 发动机数据分析工作流

按用户目标选择入口，先确认数据结构和列名，再进入具体分析。

## 前置步骤：创建输出文件夹 + 环境配置

### 输出文件夹

每次分析第一步——在数据文件所在目录下创建时间戳文件夹。

```python
import os
from datetime import datetime

filepath = "<替换为用户的数据文件路径>"  # e.g. "D:/data/发动机数据.csv"
data_parent_dir = os.path.dirname(filepath)
out_dir = os.path.join(data_parent_dir, datetime.now().strftime("%Y%m%d-%H%M%S"))
os.makedirs(out_dir, exist_ok=True)
```

后续所有图表和报告均写入 `out_dir`，例如：
```python
save_plot = os.path.join(out_dir, "comparison.png")
```

### matplotlib 无头环境

在 sandbox、无 GUI 的服务器或 CI 环境中绘图时，**必须在导入 pyplot 前设置 `Agg` 后端**，否则 Tk 后端会报错 `_tkinter.TclError: Can't find a usable tk.tcl`。

```python
import matplotlib
matplotlib.use('Agg')
# 然后再 import engine_analysis 或其他绘图模块
```

`Agg` 是纯离屏渲染后端，不依赖任何 GUI 库，支持 PNG/PDF/SVG 输出。

## A/B 增压器或方案对比

适用于两组数据上下排列、每组相同转速点的 Excel 或 CSV。

```python
df = load_excel("增压器对比数据.xlsx", sheet_name="Sheet3", skiprows=0)
print_data_structure(df)
cols = detect_all_columns(df)

df_a, df_b = split_groups(df, n_points=9)
results = compare_turbochargers(df_a, df_b, "方案A", "方案B")
altitude = assess_high_altitude(df_a, df_b, "方案A", "方案B", altitude_m=3000)
report = generate_text_report(results, altitude)
print(report)
```

`compare_turbochargers()` 从低速扭矩、BSFC、涡轮转速余量、WG 开度、排气温度、峰值扭矩等维度加权评分。若用户没有确认涡轮转速限制，不要把风险判断说成定论。

## 单发动机综合分析

适用于万有特性、外特性或稳态点数据。

```python
out = single_engine_full_analysis(
    filepath="发动机万有数据.csv",
    encoding="gbk",
    header_rows=5,
    turbo_speed_limit=250000,
    altitude_m=3000,
    save_plot_performance=os.path.join(out_dir, "performance.png"),
    save_plot_combustion=os.path.join(out_dir, "combustion.png"),
)
print(out["report"])
```

如果数据包含 COV、AI50、点火角、退角、爆震、VVT 或 IMEP，综合分析应包含燃烧诊断；缺列时说明跳过的维度。

## 标准数据对标

标准数据可以是 Excel 或 CSV，至少应包含转速、扭矩列，推荐包含功率和 BSFC。

```python
std = load_standard_data("发动机标准数据.xlsx", sheet_name="外特性")
result = compare_with_standard(
    test_rpm,
    test_torque,
    test_power,
    test_bsfc,
    standard_df=std,
    name="测试发动机",
    standard_name="标准发动机",
)
print(result["report"])
```

标准数据列名与测试数据不一致时使用 `col_map`：

```python
result = compare_with_standard(
    test_rpm,
    test_torque,
    test_power,
    test_bsfc,
    standard_df=std,
    col_map={"rpm": "DynoSpeed_Avg", "torque": "DynoTorque_Avg"},
)
```

## B15HE 内置基准

内置基准文件位于：

```text
assets/baseline_engine_database/260108_B15HE_BSFC_发动机标准数据_v1.0.xlsx
```

单发动机分析中可用：

```python
out = single_engine_full_analysis(
    filepath="测试数据.csv",
    standard_engine="B15HE",
)
```

## 燃烧特性分析

先检测燃烧列：

```python
cols = detect_all_columns(df)
```

常见燃烧键名：

- `cov`: 循环变动系数
- `ai50`: AI50 / CA50 / MFB50
- `spark_act`: 实际点火角
- `spark_mbt`: MBT 点火角
- `spark_delta`: 从 MBT 退角
- `knock`: 爆震窗口或爆震强度
- `vvt`: 可变气门正时
- `fuel_flow`: 燃油消耗量
- `imep`: 平均有效压力

一站式分析：

```python
out = single_engine_combustion_analysis(
    df,
    rpm_col,
    torque_col,
    cols,
    turbo_speed_limit=250000,
    altitude_m=3000,
    save_plot=os.path.join(out_dir, "combustion.png"),
)
print(out["report"])
```

## 燃烧敏感性分析 

当用户要求参数敏感性、影响因子、机器学习或特征重要性时使用。

运行前自动补齐 sklearn 依赖：

```python
try:
    from sklearn.ensemble import RandomForestRegressor
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "scikit-learn", "-q"])
```

```python
out = analyze_combustion_sensitivity(
    filepath="发动机数据.csv",
    encoding="gbk",
    header_rows=5,
)
# 报告已由 _build_ml_sensitivity_report() 内置生成，直接保存
with open(os.path.join(out_dir, "敏感性报告.md"), "w", encoding="utf-8") as f:
    f.write(out["report"])
```

分析管线：

1. **燃烧基线诊断** — 调用 `single_engine_combustion_analysis()` 获得 COV/AI50/点火角等基线状态
2. **单变量相关分析** — Pearson + Spearman + 互信息，覆盖线性/单调/非线性关联
3. **Random Forest 特征重要性** — impurity + permutation，控制共线性后的独立贡献 (需 sklearn)
4. **RPM / 负荷分区分析** — 按转速段和负荷区分别计算相关，检测反转点
5. **偏依赖分析** — 特征对目标的非线性响应曲线 (需 sklearn)
6. **K-Fold 交叉验证** — 模型稳定性评估 (需 sklearn)
7. **生成报告** — markdown 表格 + 标定建议 (内置 `_build_ml_sensitivity_report`)

输出 markdown 表格报告（不生成图表）。报告风格：数据驱动分析结论——直接说「大负荷区退角与 BSFC 正相关 r=+0.74，低负荷区反转为负，必须分区标定」，不要解释名词和系数含义。

典型解读维度：

- AI50 对 BSFC 的影响
- 点火角和点火退角对 COV 的影响
- 排气温度、增压压力、VVT 对燃烧品质的间接影响
- 不同负荷 / 转速区间的敏感性差异
- 非线性关系检测 (Pearson ≠ Spearman)

### ⚠ Pearson ≠ RF 重要性

**RF 重要性衡量的是「控制其他变量后，该特征的独立预测贡献」。**
**Pearson 相关衡量的是「该特征与目标在全量数据中的线性共变」。**

当 IMEP 和扭矩已经解释 94% 的 BSFC 方差时，VVT、点火角等特征的 Pearson 相关性
很大程度是**与 IMEP/扭矩共线导致的虚假相关**，而非独立的因果关系。

| 特征 | Pearson r | RF Importance | 解读 |
|------|-----------|---------------|------|
| VVT | −0.574 ▲ | 0.1% ─ | Pearson 高是因为 VVT 与负荷共变 |
| 点火角 | +0.536 ▲ | 0.4% ─ | 同上 |
| IMEP | −0.433 ▲ | 50.5% ▲▲ | RF 揭示真实主导地位 |
| 扭矩 | −0.459 ▲ | 43.6% ▲▲ | 同上 |

> 不以 Pearson 排名直接做标定建议——必须用 RF 或分区分析确认独立贡献。
> Pearson 高 RF 低的特征，基本是搭了负荷的便车。

## 可视化

```python
plot_comparison(results, save_path=os.path.join(out_dir, "comparison.png"))
```

燃烧图通常包含功率分布、BSFC 地图、COV、AI50、点火角、退角、BSFC vs AI50、点火角按负荷、COV vs IMEP。保存图表后，在答复中给出路径。

## 常见问题

### 列名检测不到（系统性误匹配）

当多个列同时误匹配（如 turbo_speed→涡轮前压力、spark_act→温度偏移、spark_mbt→排温、imep→COV），最快方案是**手动构建 col_map 字典**，完全绕过 `detect_all_columns()`：

```python
# 1. 加载数据
df = load_csv("数据.csv", encoding="gbk", header_rows=5, skip_time_cols=3)
# 清理元数据行
df = df[pd.to_numeric(df['DynoSpeed_Avg'], errors='coerce').notna()].copy()
for col in df.columns:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# 2. 手动映射（按实际列名填写）
col_map = {
    'rpm': 'DynoSpeed_Avg',
    'torque': 'CorrTorqueEWG_Avg',     # 优先修正扭矩
    'power': 'CorrBrkPwrEWG_Avg',       # 优先修正功率
    'bsfc': 'BSFC_Avg',
    'turbo_speed': 'TURBOSPEED_Avg',
    'boost': 'BSTC_pActBoostPress_Avg',
    'egt': 'EXHT_tMnfdTemp_Avg',
    'cov': 'IMEP1CO_Avg',
    'ai50': 'AI501_Avg',
    'spark_act': 'SPK_dgActSpkAdvAvg_Avg',
    'spark_mbt': 'SPK_dgMBTSpkAdv_Avg',
    'spark_delta': 'SPK_dgDltFromMBT_Avg',
    'knock': 'KNOC_dgTotalRTDAvg_Avg',
    'vvt': 'VVT1_dgCurPosOffs_Avg',
    'imep': 'IMEP1_Avg',
}

# 3. 验证关键列存在
missing = [k for k, v in col_map.items() if v not in df.columns]
if missing:
    print(f"缺失列: {missing}")
    
# 4. 后续直接用 df[col_map['bsfc']] 引用
```

此方案适用于 `single_engine_full_analysis()` 内部检测失效时的手工分析管线。

### 已知误匹配案例

| 信号 | 正确列 | 曾经误匹配为 | 原因 |
|------|--------|-------------|------|
| turbo_speed | `TURBOSPEED_Avg` | `TurbineInPressG_Avg` | `turbine` ⊆ TurbineInPressG |
| spark_act | `SPK_dgActSpkAdv_Avg` | `ConvTempSpkOfst_Avg` | `SPK` ⊆ ConvTempSpkOfst |
| spark_mbt | `SPK_dgMBTSpkAdv_Avg` | `EXHT_tPortTempMBT_Avg` | `MBT` ⊆ PortTempMBT |
| imep | `IMEP1_Avg` | `IMEP1CO_Avg` | `IMEP` ⊆ IMEP1CO |
| wg | `EWGC_rActlPos_Avg` | `BCEW_rDesWGPos_Avg` | `WG` ⊆ DesWGPos |

### 快速判断是否误匹配

- 增压器转速显示 ~199 rpm → 误匹配为涡轮前压力
- IMEP 与 COV 检出同一列 → 误匹配
- 点火角数值异常（>50° 或 < -20°）→ 误匹配
- 某参数与目标全 1.0 相关 → 误匹配为同列

### 单个列名检测不到

手动传入列名：

```python
results = compare_turbochargers(
    df_a,
    df_b,
    rpm_col="我的转速列",
    torque_col="我的扭矩列",
)
```

### CSV 读取乱码

先试 `encoding="gbk"`，再试 `encoding="latin-1"`。数值列通常不受编码影响。

### 只想对比标准数据

```python
std = load_standard_data("标准.xlsx")
result = compare_with_standard(test_rpm, test_torque, standard_df=std)
print(result["report"])
```

### 导出报告

报告字符串已由分析函数内置生成。直接从 `out["report"]` / `result["report"]` 取字符串保存，**禁止自行重写**：

```python
# ✅ 正确做法：直接保存
report = out["report"]
with open(os.path.join(out_dir, "综合报告.md"), "w", encoding="utf-8") as f:
    f.write(report)

# ✗ 错误做法：读取 out["feature_importance"] 等中间数据后自行组织报告
```
