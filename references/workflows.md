# 发动机数据分析工作流

按用户目标选择入口，先确认数据结构和列名，再进入具体分析。

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
    save_plot_performance="/tmp/performance.png",
    save_plot_combustion="/tmp/combustion.png",
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
    save_plot="/tmp/combustion.png",
)
print(out["report"])
```

## 燃烧敏感性分析

当用户要求参数敏感性、影响因子、机器学习或特征重要性时：

```python
out = analyze_combustion_sensitivity(
    filepath="发动机数据.csv",
    encoding="gbk",
    header_rows=5,
    save_plots="/tmp/sensitivity/",
)
print(out["report"])
print(out["feature_importance"])
```

典型解读维度：

- AI50 对 BSFC 的影响
- 点火角和点火退角对 COV 的影响
- 排气温度、增压压力、VVT 对燃烧品质的间接影响
- 不同负荷区间的敏感性差异

## 可视化

```python
plot_comparison(results, save_path="/tmp/comparison.png")
```

燃烧图通常包含功率分布、BSFC 地图、COV、AI50、点火角、退角、BSFC vs AI50、点火角按负荷、COV vs IMEP。保存图表后，在答复中给出路径。

## 常见问题

### 列名检测不到

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

```python
report = generate_text_report(results, altitude_results)
print(report)
with open("报告.md", "w", encoding="utf-8") as f:
    f.write(report)
```
