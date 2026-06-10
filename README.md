# Management-Statistics-No.4

## 1. 项目简介

本项目用于完成管理统计学课程中关于“广告战略与广告媒体对新产品周销量影响”的方差分析作业。题目要求根据表6-24和表6-25的数据，检验不同广告战略、广告媒体及其组合下的销售量均值是否存在显著差异。

项目使用 Python 完成数据读取、描述性统计、单因素方差分析、双因素方差分析、可视化图表生成，并最终输出 HTML 可视化报告和 Word 标准作业文档。

---

## 2. 研究问题

本项目主要回答以下四个问题：

1. 根据广告方案和广告媒体形式，给出显著性检验假设；
2. 检验三种不同广告战略下的销售量均值是否相同；
3. 检验三种广告策略和两种广告媒体组合下的销售量均值是否相同；
4. 对所使用的方差分析方法进行讨论。

---

## 3. 数据说明

数据文件位于：

```text
data/management_statistics_tables_6_24_6_25.xlsx
```

Excel 文件中包含两个工作表：

| 工作表名称      | 数据内容                   | 分析方法    |
| ---------- | ---------------------- | ------- |
| table_6_24 | 三种广告战略下20周周销量          | 单因素方差分析 |
| table_6_25 | 三种广告策略 × 两种广告媒体下10周周销量 | 双因素方差分析 |

### 表6-24字段

| 字段       | 含义                 |
| -------- | ------------------ |
| week     | 星期                 |
| strategy | 广告战略，包括便利性、高质量、低价格 |
| sales    | 周销量                |

### 表6-25字段

| 字段       | 含义                 |
| -------- | ------------------ |
| week     | 星期                 |
| strategy | 广告策略，包括便利性、高质量、低价格 |
| media    | 广告媒体，包括短视频平台、社交网站  |
| sales    | 周销量                |

---

## 4. 项目结构

```text
管理统计第四节作业/
│
├─ data/
│   └─ management_statistics_tables_6_24_6_25.xlsx
│
├─ outputs/
│   ├─ results.json
│   ├─ table_6_24_clean.csv
│   ├─ table_6_25_clean.csv
│   ├─ desc_table_6_24.csv
│   ├─ desc_table_6_25.csv
│   ├─ anova_oneway_table_6_24.csv
│   ├─ anova_twoway_table_6_25.csv
│   ├─ fig_1_strategy_mean.png
│   ├─ fig_2_strategy_boxplot.png
│   ├─ fig_3_combination_mean.png
│   └─ fig_4_interaction.png
│
├─ reports/
│   ├─ 广告战略方差分析报告.html
│   └─ 管理统计学_广告战略方差分析作业.docx
│
├─ 01_analysis.py
├─ 02_generate_html.py
├─ 03_generate_word.py
├─ requirements.txt
└─ README.md
```

---

## 5. 环境配置

建议使用独立虚拟环境运行本项目。

### 创建虚拟环境

```powershell
python -m venv .venv
```

### 激活虚拟环境

```powershell
.venv\Scripts\activate
```

### 安装依赖

```powershell
python -m pip install -r requirements.txt
```

如果没有 `requirements.txt`，可以手动安装：

```powershell
python -m pip install pandas numpy scipy statsmodels matplotlib openpyxl python-docx
```

---

## 6. 运行方法

请按照以下顺序运行代码。

### 第一步：运行统计分析与图表生成

```powershell
python 01_analysis.py
```

如果 Excel 文件路径不同，可以使用：

```powershell
python 01_analysis.py --input "data/management_statistics_tables_6_24_6_25.xlsx"
```

该脚本会完成：

* 数据读取与校验；
* 描述性统计；
* 单因素方差分析；
* 双因素方差分析；
* 方差齐性检验；
* 残差正态性检验；
* 图表生成；
* 结果保存到 `outputs/results.json`。

### 第二步：生成 HTML 报告

```powershell
python 02_generate_html.py
```

输出文件：

```text
reports/广告战略方差分析报告.html
```

### 第三步：生成 Word 作业文档

```powershell
python 03_generate_word.py
```

输出文件：

```text
reports/管理统计学_广告战略方差分析作业.docx
```

---

## 7. 主要分析方法

### 7.1 单因素方差分析

用于检验表6-24中三种广告战略下的平均周销量是否相同。

原假设：

```text
H0：μ便利性 = μ高质量 = μ低价格
```

备择假设：

```text
H1：三种广告战略下的平均周销量不全相同
```

### 7.2 双因素方差分析

用于检验表6-25中广告策略、广告媒体及二者交互作用是否影响平均周销量。

检验内容包括：

1. 广告策略主效应；
2. 广告媒体主效应；
3. 广告策略 × 广告媒体交互作用。

---

## 8. 核心结果摘要

### 表6-24：单因素方差分析

单因素方差分析结果为：

```text
F = 0.7215
p = 0.4904
```

由于 p > 0.05，因此在5%显著性水平下不能拒绝原假设。结论是：三种广告战略下的平均周销量没有显著差异。

### 表6-25：双因素方差分析

双因素方差分析结果为：

```text
广告策略主效应 p = 0.6866
广告媒体主效应 p = 0.0530
策略×媒体交互作用 p = 0.1184
```

三者均大于0.05，因此在5%显著性水平下，不能认为广告策略、广告媒体及二者交互作用对平均周销量具有显著影响。

---

## 9. 输出说明

### outputs 文件夹

`outputs` 文件夹保存中间结果和图表，供 HTML 和 Word 报告复用。

主要文件包括：

| 文件                          | 含义            |
| --------------------------- | ------------- |
| results.json                | 汇总后的统计结果和图表路径 |
| desc_table_6_24.csv         | 表6-24描述性统计    |
| desc_table_6_25.csv         | 表6-25描述性统计    |
| anova_oneway_table_6_24.csv | 单因素方差分析表      |
| anova_twoway_table_6_25.csv | 双因素方差分析表      |
| fig_1_strategy_mean.png     | 三种广告战略均值柱状图   |
| fig_2_strategy_boxplot.png  | 三种广告战略销量箱线图   |
| fig_3_combination_mean.png  | 策略×媒体组合均值柱状图  |
| fig_4_interaction.png       | 策略×媒体组合对比图    |

### reports 文件夹

`reports` 文件夹保存最终报告。

| 文件                    | 用途        |
| --------------------- | --------- |
| 广告战略方差分析报告.html       | 可视化展示版本   |
| 管理统计学_广告战略方差分析作业.docx | 正式提交或阅读版本 |

---

## 10. 方法讨论

本项目采用方差分析方法是合理的，因为研究目标是比较多个广告方案下的平均销量是否存在差异。表6-24只涉及广告战略一个因素，因此使用单因素方差分析；表6-25同时涉及广告策略和广告媒体两个因素，因此使用双因素方差分析。

需要注意的是，方差分析依赖样本独立、误差近似正态、方差齐性等前提。虽然题目假定不同城市除广告战略外其他条件完全相同，但现实中城市消费水平、人口结构、竞争环境、渠道条件和时间趋势都可能影响销量。因此，本项目结果应理解为基于样本数据的统计判断，不能仅凭样本均值直接决定最终广告方案。

---

## 11. 注意事项

1. 运行 `02_generate_html.py` 和 `03_generate_word.py` 前，应先运行 `01_analysis.py`；
2. 如果修改了 Excel 数据，需要重新运行全部脚本；
3. 如果图片没有更新，可以删除 `outputs` 文件夹中的旧图片后重新运行；
4. 如果图表中文乱码，请检查系统是否安装了微软雅黑、黑体或其他中文字体；
5. 如果 Word 中图片没有变化，通常是因为没有重新运行 `01_analysis.py` 或旧图片仍被引用。

---

## 12. 项目结论

根据现有样本数据，三种广告战略下的平均周销量没有显著差异；广告策略、广告媒体及二者交互作用在5%显著性水平下也均不显著。管理上可以参考样本均值，但不应仅凭样本均值直接判断某一广告战略或媒体组合显著优于其他方案。
