# ETAS INCA 台架信号命名参考

> 基于 B15HTC 万有数据分析实战提炼。包含通用命名规则、子系统前缀表、B15HTC 实测信号映射、常见坑和编码问题。
> 用于快速识别 CSV 数据中的燃烧/性能信号，减少列名搜索时间。

## 概述

ETAS INCA 是常见的发动机标定/数据采集工具。其导出的 CSV 文件有固定格式特征。

## CSV 文件结构

```
Row 0-4:    元数据 (Logger description, Log period, Statistics...)
Row 5:      真正的列名 (Time, DynoSpeed_Avg, DynoTorque_Avg, ...)
Row 6:      单位行 (ms, rev/min, Nm, kW, ...)
Row 7:      数据类型行 (Raw, Average, Average, ...)
Row 8+:     实际数据 (数值)
```

## 列名前缀规则

INCA 输出的信号名按子系统分类，以 `_` 分隔：

| 前缀 | 子系统 | 典型信号 |
|------|-------|---------|
| `Dyno` | 台架测功机 | `DynoSpeed_Avg`, `DynoTorque_Avg`, `BrakePower_Avg` |
| `BSFC` | 燃油消耗率 | `BSFC_Avg` |
| `BSTC` | 增压系统控制 | `BSTC_pActBoostPress_Avg`, `BSTC_pDsrdBoostPres_Avg` |
| `EWGC` | 废气门控制 | `EWGC_rActlPos_Avg`, `EWGC_rPosDsrd_Avg` |
| `EXHT` | 排气温度 | `EXHT_tMnfdTemp_Avg`, `EXHT_tMnfdTempFlt_Avg` |
| `TURBO` | 增压器 | `TURBOSPEED_Avg` |
| `SPK` | 点火系统 | `SPK_dgActSpkAdv_Avg`, `SPK_dgMBTSpkAdv_Avg`, `SPK_dgDltFromMBT_Avg` |
| `AI50` | 燃烧相位 | `AI501_Avg` (1缸), `AI502_Avg` (2缸) |
| `IMEP` | 平均有效压力 | `IMEP1_Avg`, `IMEP1CO_Avg` (COV), `IMEPH1_Avg` (高压) |
| `IMEP1CO` | 循环变动系数 | `IMEP1CO_Avg` — 缸1的COV |
| `Fuel` | 燃油系统 | `Fuel_FuelConsume_Avg`, `FuelMassFlowRate_Avg` |
| `knockWnd` | 爆震窗口 | `knockWndStrAng_Avg`, `knockWndWdtAng_Avg` |
| `AF` | 空燃比 | `AF_Avg`, `AFDirect_Avg` |
| `AirFlow` | 进气流量 | `AirFlow_Avg`, `CalcAirFlow_Avg` |
| `AirTemp` | 进气温度 | `AirTemp_Avg` |
| `EGT` / `T_EXH` | 排气温度(替代) | 部分项目用 `EXHT_tMnfdTemp_Avg` |
| `VVT` / `CamPhs` | 可变气门正时 | `CamPhs...`, `VVT...` (项目相关) |

## 后缀含义

| 后缀 | 含义 |
|------|------|
| `_Avg` | 统计平均值 (Statistics window 内) |
| `_Min` | 统计最小值 |
| `_Max` | 统计最大值 |
| `_Raw` | 原始请求值 (vs 实际值) |
| `_Dsrd` | 目标值 (desired) |
| `_Flt` | 滤波后值 (filtered) |

## 燃烧特性信号映射 (实战验证)

在 B15HTC 万有数据中验证过的信号映射：

```python
rpm          → DynoSpeed_Avg        # [3]
torque       → DynoTorque_Avg       # [4]
power        → BrakePower_Avg       # [5]
bsfc         → BSFC_Avg             # [18]
boost        → BSTC_pActBoostPress_Avg  # [19]
wg           → EWGC_rActlPos_Avg    # [77]
egt          → EXHT_tMnfdTemp_Avg   # [93]
turbo_speed  → TURBOSPEED_Avg       # [212]
cov          → IMEP1CO_Avg          # [6]
ai50         → AI501_Avg            # [227]
spark_act    → SPK_dgActSpkAdv_Avg  # [162]
spark_mbt    → SPK_dgMBTSpkAdv_Avg  # [168]
spark_delta  → SPK_dgDltFromMBT_Avg # [165]
knock        → knockWndStrAng_Avg   # [119]
fuel_flow    → Fuel_FuelConsume_Avg # [104]
imep         → IMEP1_Avg            # [229]
airflow      → AirFlow_Avg          # [9]
```

> 注意: 实际列索引可能因项目而异，以上索引基于 B15HTC 数据实测。

## 常见问题

### 多列同名
INCA 有时输出 3 列都叫 `Time` (Date/Time/ms)，pandas 自动加后缀 `.1` `.2`：
```python
# 前3列是时间戳，从第4列开始才是信号
df = raw.iloc[:, 3:]
```

### 编码问题
国内台架 CSV 通常用 GBK 编码：
```python
df = pd.read_csv("data.csv", encoding='gbk', header=5)
```
失败时尝试 `latin-1`（数值不受影响）。

### 数值格式
INCA 可能输出 `1.#QNAN` 等特殊 NaN 表示方式，需用 `pd.to_numeric(errors='coerce')` 处理：
```python
df = df.apply(pd.to_numeric, errors='coerce')
```
