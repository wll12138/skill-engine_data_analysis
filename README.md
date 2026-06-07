# engine-data-analysis

发动机台架性能数据分析工具集 — Hermes Agent Skill

## 功能

- **燃烧特性分析** — COV（循环变动系数）/ AI50（CA50 燃烧相位）/ 点火角（实际 / MBT / 退角）/ 爆震 / VVT / IMEP
- **增压器 A/B 对比** — 双增压器多维度加权评分（低速扭矩 / BSFC / 涡轮转速余量 / WG 效率 / 排温 / 峰值扭矩）
- **高原能力评估** — 根据 ISO 2533 标准大气模型推算高原增压器转速，评估安全余量
- **单发动机万有特性分析** — 全负荷 + 部分负荷稳态点分析（扭矩/功率/BSFC/增压压力/WG 开度/排温/涡轮转速）
- **标准数据对标** — 任意发动机标准数据的通用对比框架（支持自定义列映射和外特性/万有特性对标）
- **燃烧敏感性分析** — 基于随机森林回归，分析各参数对 BSFC 和 COV 的敏感度
- **数据可视化** — 性能对比图（6 子图）、燃烧特性图（9 子图）、单机分析图（8 子图）
- **自动列名检测** — 20+ 种发动机信号自动模糊匹配（支持中文/英文/ETAS INCA 命名），扭矩/功率优先使用修正值

## 快速使用

```python
from pathlib import Path
import sys
sys.path.insert(0, str(Path.home() / ".hermes/skills/data-science/engine-data-analysis/scripts"))
from engine_analysis import *

# 增压器 A/B 对比
out = full_analysis("对比数据.xlsx", "方案A", "方案B", n_points=9)
print(out["report"])

# 单发动机燃烧分析
out = single_engine_full_analysis(
    "发动机万有数据.csv",
    encoding="gbk", header_rows=5,
    save_plot_performance="/tmp/performance.png",
    save_plot_combustion="/tmp/combustion.png",
)
print(out["report"])

# 标准数据对比
std = load_standard_data("标准数据.xlsx")
result = compare_with_standard(rpm, torque, power, bsfc, standard_df=std)
print(result["report"])
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
├── SKILL.md                          # Hermes Agent 技能描述
├── README.md                         # 本文件
├── scripts/
│   └── engine_analysis.py            # 核心分析模块
├── references/
│   └── etas_inca_signals.md          # ETAS INCA 信号命名规范
└── baseline engine database/
    └── 260108_B15HE_BSFC_发动机标准数据_v1.0.xlsx  # 对标基准数据库
```

## 维护与发布

GitHub 仓库：[johnhejunlin/skill-engine_data_analysis](https://github.com/johnhejunlin/skill-engine_data_analysis)

每次更新本 skill 后必须：
- 同步更新 `README.md`
- 将最新改动上传到 GitHub 仓库

## 更新日志

### 2026-06-07 — docs: add GitHub publish requirement
- 记录 skill 的 GitHub 仓库地址
- 明确每次更新 skill 后需同步更新 README 并上传到 GitHub

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
