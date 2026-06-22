# engine-data-analysis

发动机台架数据综合分析工具集 — Hermes Agent Skill

## 功能

- **燃烧特性分析** — COV（循环变动系数）/ AI50（CA50 燃烧相位）/ 点火角（实际 / MBT / 退角）/ 爆震 / VVT / IMEP
- **增压器 A/B 对比** — 双增压器多维度加权评分（低速扭矩 / BSFC / 涡轮转速余量 / WG 效率 / 排温 / 峰值扭矩）
- **高原能力评估** — 根据 ISO 2533 标准大气模型推算高原增压器转速，评估安全余量
- **单发动机万有特性分析** — 全负荷 + 部分负荷稳态点分析（扭矩/功率/BSFC/增压压力/WG 开度/排温/涡轮转速）
- **标准数据对标** — 任意发动机标准数据的通用对比框架（支持自定义列映射和外特性/万有特性对标）
- **数据可视化** — 性能对比图（6 子图）、燃烧特性图（9 子图）、单机分析图（8 子图）
- **自动列名检测** — 20+ 种发动机信号自动模糊匹配（支持中文/英文/ETAS INCA 命名），扭矩/功率优先使用修正值
- **燃烧敏感性分析 (ML)** — Pearson + Spearman + 互信息 + Random Forest + 偏依赖 + RPM/负荷分区 + 交叉验证，输出 markdown 表格报告

## 快速使用

这个 skill 的核心不是单一脚本命令，而是按发动机台架数据分析场景选择入口：先理解数据结构和列名，再进行 A/B 对比、单发动机综合分析、高原评估或标准数据对标。

```python
from pathlib import Path
import sys
sys.path.insert(0, str(Path.home() / "AppData/Local/hermes/skills/engineering/engine-data-analysis/scripts"))
from engine_analysis import *

# 1. 先查看数据结构并自动识别信号列
df = load_excel("发动机台架数据.xlsx", sheet_name="Sheet1", skiprows=0)
print_data_structure(df)
cols = detect_all_columns(df)
print(cols)

# 2. 增压器/方案 A/B 对比：性能 + BSFC + WG + 排温 + 涡轮转速 + 高原能力
out = full_analysis(
    filepath="增压器对比数据.xlsx",
    name_a="方案A",
    name_b="方案B",
    n_points=9,
    turbo_speed_limit=250000,
    altitude_m=3000,
    save_plot="/tmp/turbo_comparison.png",
)
print(out["report"])

# 3. 单发动机综合分析：性能 + 燃烧 + 高原能力 + 可选标准对标
out = single_engine_full_analysis(
    filepath="发动机万有数据.csv",
    encoding="gbk", header_rows=5,
    turbo_speed_limit=250000,
    altitude_m=3000,
    save_plot_performance="/tmp/performance.png",
    save_plot_combustion="/tmp/combustion.png",
    standard_engine="B15HE",  # 不需要标准对标时设为 None
)
print(out["report"])

# 4. 通用标准数据对标：适合外特性/WOT 标准表
test_df = load_excel("测试外特性.xlsx")
test_cols = detect_all_columns(test_df)
std = load_standard_data("标准数据.xlsx")
result = compare_with_standard(
    test_rpm=test_df[test_cols["rpm"]].to_numpy(),
    test_torque=test_df[test_cols["torque"]].to_numpy(),
    test_power=test_df[test_cols["power"]].to_numpy() if test_cols["power"] else None,
    test_bsfc=test_df[test_cols["bsfc"]].to_numpy() if test_cols["bsfc"] else None,
    standard_df=std,
    name="测试发动机",
    standard_name="标准发动机",
)
print(result["report"])
```

### 5. 燃烧敏感性分析 (ML)

```python
sens = analyze_combustion_sensitivity(
    filepath="发动机万有数据.csv",
    encoding="gbk", header_rows=5,
)
print(sens["report"])
print(sens["feature_importance"])
```

分析管线：燃烧基线 → Pearson+Spearman+互信息 → Random Forest → RPM/负荷分区 → 偏依赖 → K-Fold 交叉验证 → 报告生成。无 sklearn 时自动回退。输出 markdown 表格报告。
```

## 支持的信号

| 类型 | 信号 | 检测优先级 |
|------|------|-----------|
| 性能 | rpm / torque(修正优先) / power(修正优先) / BSFC / boost / EGT / turbo_speed / WG / airflow / backpressure | 修正扭矩/功率优先于原始值 |
| 燃烧 | COV / AI50(CA50) / spark_act / spark_mbt / spark_delta / knock / VVT / fuel_flow / IMEP | 自动检测，有则分析无则跳过 |
| 高原 | 海平面转速推算，标准大气模型 | 需用户确认限制值 |

## 文件结构

```
engine-data-analysis/
├── SKILL.md                          # Skill 触发条件、入口和核心流程
├── README.md                         # 仓库说明
├── agents/
│   └── openai.yaml                   # UI 元数据 / 隐式调用配置
├── scripts/
│   └── engine_analysis.py            # 核心分析模块
├── assets/
│   └── baseline_engine_database/
│       └── 260108_B15HE_BSFC_发动机标准数据_v1.0.xlsx  # 对标基准数据库
├── references/
│   ├── etas_inca_signals.md          # ETAS INCA 信号命名规范
│   ├── workflows.md                  # 场景化分析工作流
│   └── thresholds_and_formulas.md    # 阈值、公式、CSV 与列名参考
```

## 维护与发布

GitHub 仓库：[johnhejunlin/skill-engine_data_analysis](https://github.com/johnhejunlin/skill-engine_data_analysis)

每次更新本 skill 后必须：
- 同步更新 `README.md`
- 将最新改动上传到 GitHub 仓库

## 更新日志

### 2026-06-08 — feat: implement analyze_combustion_sensitivity (ML)
- 实现 `analyze_combustion_sensitivity()` — 7 步 ML 管线（燃烧基线→Pearson+Spearman+互信息→RF→分区→偏依赖→CV→报告生成），sklearn 可选，无 sklearn 自动回退
- 新增 `_merge_feature_importance()` 5 方法融合排名、`_build_ml_sensitivity_report()` 数据驱动分析报告（markdown 表格）
- 修复 `_single_engine_performance_core` numpy `or` 崩溃、`max_power_rpm` KeyError、`_merge_feature_importance` fk 脱循环 bug
- 更新 `COLUMN_PATTERNS` 修复 ETAS INCA 常见误匹配（turbo_speed/spark_act/spark_mbt/imep/wg）
- SKILL.md 精简场景选择、输出要求，与 workflows.md 去重
- README.md 更新路径、功能列表、文件结构

### 2026-06-07 — refactor: optimize skill structure
- 精简 `SKILL.md`，保留触发条件、快速入口和核心工作流
- 新增 `agents/openai.yaml` UI 元数据
- 新增 `references/workflows.md` 和 `references/thresholds_and_formulas.md`
- 将 B15HE 基准数据库移动到 `assets/baseline_engine_database/`
- 同步更新 `_B15HE_STANDARD_PATH`

### 2026-06-07 — docs: add GitHub publish requirement
- 记录 skill 的 GitHub 仓库地址
- 明确每次更新 skill 后需同步更新 README 并上传到 GitHub

### 2026-06-07 — docs: clarify quick usage scope
- 将 README 定位从“性能数据分析”调整为“台架数据综合分析”
- 重写快速使用示例，按数据识别、A/B 对比、单发动机综合分析和标准对标组织
- 移除 README 中尚未在脚本实现的一键燃烧敏感性分析声明

### 2026-06-03 — docs: update file structure for baseline engine database
- 标准数据文件移入 `baseline engine database/` 子目录
- 同步更新 `_B15HE_STANDARD_PATH` 模块内文件路径
- README 文件结构更新

### 2026-06-02 — refactor: 通用化 + 修正值优先 + 标准对比框架
- 拓宽 skill 描述：燃烧 + 性能 + 油耗 + 增压器（作为子项）
- 扭矩/功率列名优先匹配修正值（修正扭矩 / CorrTorqueEWG / 修正功率 / CorrBrkPwrEWG）
- 替换所有具体品牌名（博马/奕森 → 方案A/方案B）
- 新增通用标准对比框架：`load_standard_data()` / `compare_with_standard()`
- 新增燃烧敏感性分析：`analyze_combustion_sensitivity()`
- SKILL.md 去除绝对路径（改用 `Path.home()`），重写内容结构

### 2026-06-02 — feat: add combustion analysis
- 新增燃烧特性分析：COV/AI50/点火角/点火退角/爆震/VVT/IMEP 信号检测
- 新增 `single_engine_combustion_analysis()` 燃烧一站式分析
- 新增 `_plot_combustion_analysis()` 9 子图燃烧可视化
- COLUMN_PATTERNS 新增 9 种信号类型

### 2026-06-02 — feat: add B15HE standard comparison
- 新增 B15HE 标准数据文件
- 新增 `compare_with_b15he_standard()` 外特性对比

### 2026-06-01 — initial: extract reusable module, slim down SKILL.md
- 将内联代码提取为 `scripts/engine_analysis.py` 结构化模块
- SKILL.md 精简 71%（35KB → 10KB）
- 增压器 7 维度加权评分 + 高原能力评估 + 自动列名检测
