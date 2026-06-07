# 阈值、公式与处理要点

## 增压器对比评分维度

| 维度 | 权重 | 说明 |
|------|------|------|
| 低速扭矩（1000-1500rpm） | x2.0 | 起步响应 |
| BSFC 燃油经济性 | x2.0 | 越低越好，需排查异常点 |
| 涡轮转速余量（>=4000rpm） | x1.5 | 余量大则安全 |
| WG 开度（>=3000rpm） | x1.5 | 越小废气利用率越高 |
| 排气温度 | x1.0 | 越低越好 |
| 峰值扭矩 | x1.0 | 越高越好 |

## 关键公式

### 增压比

```text
增压比 = 增压压力(kPa) / 标准大气压(101.325 kPa)
```

注意不要用排气背压作为分母。

### 高原转速推算

```text
Speed_alt = Speed_0 * sqrt(P0 / P_alt)
P_alt = 101.325 * (1 - 0.0065 * h / 288.15) ^ 5.255
```

使用前需要确认目标海拔和增压器转速限制。

## 标准大气压参考

| 海拔 | 大气压(kPa) | 与海平面比值 |
|------|-------------|-------------|
| 0m | 101.325 | 1.000 |
| 1000m | 89.9 | 0.887 |
| 2000m | 79.5 | 0.785 |
| 3000m | 70.1 | 0.692 |
| 4000m | 61.6 | 0.608 |

## 指标阈值

| 指标 | 优秀 | 良好 | 一般 | 差 |
|------|------|------|------|----|
| WG 开度 | < 10% | 10-20% | 20-30% | > 30% |
| COV 循环变动 | < 2% | 2-3% | 3-5% | > 5% |
| AI50 燃烧相位 | 6-12 deg | 12-15 deg | 15-20 deg | > 20 deg 或 < 5 deg |
| 点火退角 | < 3 deg | 3-5 deg | 5-8 deg | > 8 deg |
| 增压器转速余量 | > 50 krpm | 30-50 krpm | 15-30 krpm | < 15 krpm |
| 最低 BSFC | < 230 | 230-250 | 250-270 | > 270 g/kWh |
| 低速扭矩（1000rpm） | > 100 Nm | 80-100 | 60-80 | < 60 |

## CSV 文件处理

典型 ETAS INCA CSV：

| 行号 | 内容 |
|------|------|
| 0-4 | Logger description、Log period 等元数据 |
| 5 | 列名：Time、DynoSpeed_Avg 等 |
| 6 | 单位行：ms、rev/min、Nm 等 |
| 7 | 数据类型：Raw、Average 等 |
| 8+ | 实际数据 |

读取示例：

```python
df = load_csv("数据.csv", encoding="gbk", header_rows=5, skip_time_cols=3)
```

处理要点：

- 重名 `Time` 列会被 pandas 自动加 `.1`、`.2` 后缀。
- `1.#QNAN` 等特殊值用 `ensure_numeric()` 转为 NaN。
- 合并单元格或分组标签可用 `df.ffill()` 处理。
- 若 GBK 报错，尝试 `latin-1`；数值列通常不受影响。

## 列名检测参考

| 信号类型 | 键名 | 关键匹配词 |
|----------|------|------------|
| 转速 | `rpm` | 转速, rpm, SPEED, DynoSpeed |
| 扭矩（修正优先） | `torque` | 修正扭矩, CorrTorque, CorrTorqueEWG, 扭矩, Torque |
| 功率（修正优先） | `power` | 修正功率, CorrBrkPwr, CorrBrkPwrEWG, 功率, Power |
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

ETAS INCA 信号常带 `_Avg` 后缀，例如 `DynoSpeed_Avg`、`IMEP1CO_Avg`。
