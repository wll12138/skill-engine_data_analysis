---
name: engine-data-analysis
description: Analyze engine dyno/bench test data — combustion characteristics (COV/AI50/spark/knock/VVT/IMEP), performance parameters (torque/power/BSFC/boost/EGT), fuel consumption analysis, turbocharger matching comparison, high-altitude capability assessment, data visualization, and standard data benchmarking. Use when the user asks about engine performance data, combustion analysis, dyno data, turbocharger comparison, BSFC analysis, or any .xlsx/.csv engine test files.
---

# 发动机数据分析

分析发动机台架测试数据（dyno data）：燃烧特性、性能参数、油耗、增压器匹配、高原能力、标准数据对标和可视化。

## 触发条件

当用户提及以下任何内容时，应加载本 skill：

- 发动机 / engine / 台架 / dyno / 测试数据
- 燃烧 / COV / AI50 / CA50 / 点火角 / MBT / 爆震 / VVT / IMEP
- 增压器 / turbo / 涡轮 / 增压器匹配 / WG / 排气温度 / 背压
- BSFC / 油耗 / 燃油消耗率 / 扭矩 / 功率 / 增压压力
- 高原 / 高海拔 / 万有特性 / 标准对比 / 对标
- `.xlsx` 或 `.csv` 格式的发动机测试数据

## 资源导航

- 核心分析模块：`scripts/engine_analysis.py`
- ETAS INCA 精确列名：`references/etas_inca_signals.md`
- 场景化工作流与示例：按需读取 `references/workflows.md`
- 阈值、公式、CSV 处理和列名参考：按需读取 `references/thresholds_and_formulas.md`
- B15HE 基准数据：`assets/baseline_engine_database/260108_B15HE_BSFC_发动机标准数据_v1.0.xlsx`
- 列名检测修复方案：`references/workflows.md`（「常见问题」章节）

只在需要对应细节时读取 reference 文件；优先从本文件获得入口和决策路径。

## 快速入口

通过相对路径导入核心分析逻辑：

```python
from pathlib import Path
import sys
sys.path.insert(0, str(Path.home() / "AppData/Local/hermes/skills/engineering/engine-data-analysis/scripts"))
from engine_analysis import *
```

> 以下示例假定 `out_dir` 已创建（代码见 `references/workflows.md` →「前置步骤」）。

### 先看数据结构

```python
df = load_excel("数据.xlsx")
print_data_structure(df)
cols = detect_all_columns(df)
print(cols)
```

`.csv` 常见为 GBK 编码、多行表头：

```python
df = load_csv("数据.csv", encoding="gbk", header_rows=5, skip_time_cols=3)
cols = detect_all_columns(df)
```

### 增压器 A/B 对比

```python
out = full_analysis(
    filepath="增压器对比数据.xlsx",
    name_a="方案A",
    name_b="方案B",
    n_points=9,               # 每组行数；不指定则自动推断
    sheet_name="Sheet3",
    skiprows=0,               # 如果有单位行可设为 1
    turbo_speed_limit=250000, # 需用户或供应商确认
    altitude_m=3000,          # 设为 None 则跳过高原评估
    save_plot=os.path.join(out_dir, "comparison.png"),
)
print(out["report"])
```

### 单发动机综合分析

```python
out = single_engine_full_analysis(
    filepath="发动机万有数据.csv",
    encoding="gbk",
    header_rows=5,
    turbo_speed_limit=250000,
    altitude_m=3000,
    save_plot_performance=os.path.join(out_dir, "performance.png"),
    save_plot_combustion=os.path.join(out_dir, "combustion.png"),
    standard_engine="B15HE",  # 不需要基准对标时设为 None
)
print(out["report"])
```

### 标准数据对比

```python
std = load_standard_data("发动机标准数据.xlsx", sheet_name="外特性")
result = compare_with_standard(
    test_rpm,
    test_torque,
    test_power,
    test_bsfc,
    standard_df=std,
    name="测试",
    standard_name="标准",
)
print(result["report"])
```

标准数据列名不一致时，传入 `col_map`：

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

## 分析工作流

0. 创建输出文件夹（详见「输出文件夹管理」）。
1. 识别文件类型和表头结构：Excel 先看 sheet，CSV 先确认编码、表头行、时间列。
2. 使用 `print_data_structure()` 和 `detect_all_columns()` 检查列名匹配结果。
3. 若数据为 A/B 上下排列，使用 `split_groups()` 分组后运行 `compare_turbochargers()` 或 `full_analysis()`。
4. 若是单发动机万有数据，运行 `single_engine_full_analysis()`，有燃烧列时会补充燃烧分析。
5. 涉及高原能力时，先确认 `turbo_speed_limit` 和目标 `altitude_m`。
6. 涉及标准对标时，优先使用 `load_standard_data()` + `compare_with_standard()`；B15HE 可用内置基准路径。
7. 生成图表时保存到输出文件夹，并在最终答复中给出图表路径和关键结论。

## 列名检测原则

`detect_column()` 和 `detect_all_columns()` 支持中英文混排、换行列名和 ETAS INCA `_Avg` 后缀。扭矩和功率优先匹配修正值，例如 `修正扭矩`、`CorrTorqueEWG`、`修正功率`、`CorrBrkPwrEWG`，再匹配原始值。

INCA 数据优先用精确全名（如 `TURBOSPEED_Avg`、`IMEP1_Avg`）做 Phase 1 匹配，中文/缩写作 fallback。禁止裸 `IMEP`/`SPK`/`MBT`/`turbine` 等跨信号边界的短模式。

如自动检测不符合用户意图，**首选方案是手动构建 `col_map` 字典**直接使用原始列名，完全绕过自动检测：

```python
col_map = {
    'rpm': 'DynoSpeed_Avg',
    'torque': 'CorrTorqueEWG_Avg',
    'power': 'CorrBrkPwrEWG_Avg',
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
```

后续分析直接通过 `df[col_map['bsfc']]` 引用，不再依赖检测结果。备选方案：手动传入单个列名，或运行时动态修正 `COLUMN_PATTERNS`（见 `references/workflows.md`「常见问题」章节）。

详细信号映射见 `references/etas_inca_signals.md` 和 `references/thresholds_and_formulas.md`。

## 场景选择

- A/B 增压器或方案对比：使用 `full_analysis()`。
- 单发动机性能 + 燃烧 + 高原综合分析：使用 `single_engine_full_analysis()`。
- 只做燃烧诊断：先检测 `cov`、`ai50`、`spark_act`、`spark_mbt`、`spark_delta`、`knock`、`vvt`、`imep`，再使用 `single_engine_combustion_analysis()`。
- 只做标准对标：使用 `load_standard_data()` 和 `compare_with_standard()`。
- 需要燃烧敏感性或机器学习分析：使用 `analyze_combustion_sensitivity()`。只输出 markdown 表格，不生成任何图表。详细见 `references/workflows.md`。

## 关键注意事项

- 列名自动检测可能误匹配：turbo_speed / spark_act / spark_mbt / imep 根因为 COLUMN_PATTERNS 中短模式跨越信号边界（`turbine`→涡轮压力，`SPK`→温度偏移，`MBT`→排温，`IMEP`→COV）。应以 `references/etas_inca_signals.md` 精确全名为首位模式。修复方案见 `references/workflows.md`「常见问题」章节。
- 高原转速评估前必须确认增压器转速限制值；默认值只是占位。
- BSFC 异常点，尤其 1000rpm，分析时应结合数据质量判断。
- 增压压力、排气背压、绝对压力和表压要确认单位和含义，不要混用。
- CSV 乱码优先尝试 encoding=gbk，失败再尝试 latin-1。
- 标准数据至少需要转速和扭矩列；功率和 BSFC 缺失时只输出可比维度。
- 输出要求：每次分析必须将所有报告写入同一个时间戳文件夹，禁止分散到多个文件夹。必须保存综合报告 .md，如有敏感性分析则保存敏感性报告 .md。图表为选配（save_plot 参数），不强制生成。答复时给出文件夹路径和文件清单。

## CLI 快速调试

```bash
python ~/AppData/Local/hermes/skills/engineering/engine-data-analysis/scripts/engine_analysis.py \
  "数据文件.xlsx" "方案A" "方案B" 9
```

## 输出文件夹管理

**每次分析必须在数据文件所在目录下创建 `YYYYMMDD-HHMMSS` 时间戳文件夹**，所有输出（报告 `.md`，选配图表 `.png`）写入该文件夹，答复时给出路径。禁止复用旧文件夹，禁止分析完成后补建。具体实现代码见 `references/workflows.md` →「前置步骤」。

## 输出要求

**⚠ 报告已由内置函数生成**：`generate_text_report()`、`_build_combustion_report()`、
`_build_ml_sensitivity_report()` 等函数已在返回 dict 的 `"report"` 键中生成完整 markdown。
**直接从该键取字符串保存为 .md，禁止自行重写、补充或重新排版。**

答复用户时优先给出：

- 数据读取和列名检测是否可靠（特别注意常见误匹配：涡轮转速、点火角、IMEP vs COV）
- 关键性能差异：扭矩、功率、BSFC、增压压力、排温、WG、涡轮转速
- 燃烧诊断：COV、AI50、点火退角、爆震风险
- 高原能力：目标海拔下的转速推算和安全余量
- 标准对标：领先/落后点、差值和风险
- 敏感性分析：BSFC/COV 主导参数、负荷/转速分区差异、反转点、标定优先级建议
- 图表或报告文件路径，所有输出存入同一个时间戳文件夹
