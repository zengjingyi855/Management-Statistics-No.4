# -*- coding: utf-8 -*-
"""
01_analysis.py
管理统计学作业：广告战略与广告媒体对周销量的影响分析

功能：
1. 读取表6-24和表6-25数据；
2. 检查数据结构与样本量；
3. 完成描述性统计、单因素方差分析、双因素方差分析；
4. 完成方差齐性检验和残差正态性检验；
5. 输出 CSV / JSON 结果和 PNG 可视化图表，供 HTML 与 Word 报告复用。

运行方式：
    python 01_analysis.py

如果 Excel 路径不同：
    python 01_analysis.py --input "data/management_statistics_tables_6_24_6_25.xlsx"
"""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
from typing import Any, Dict, List

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import scipy.stats as stats
import statsmodels.api as sm
from matplotlib import font_manager
from statsmodels.formula.api import ols


PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT = PROJECT_ROOT / "data" / "management_statistics_tables_6_24_6_25.xlsx"
OUTPUT_DIR = PROJECT_ROOT / "outputs"

STRATEGY_ORDER = ["便利性", "高质量", "低价格"]
MEDIA_ORDER = ["短视频平台", "社交网站"]
ALPHA = 0.05

# 图表配色：清爽亮蓝风，亮但不杂
# 图表配色：清爽亮蓝商务风
# 注意：这里保留 light_bg、red、purple 等键，是为了兼容后面所有画图函数，避免 KeyError。
COLORS = {
    "navy": "#0F2742",        # 深蓝灰：标题、坐标轴、文字
    "blue": "#2D6CDF",        # 主蓝
    "sky": "#5AA9E6",         # 浅蓝
    "cyan": "#00A6D6",        # 青蓝
    "green": "#10B981",       # 青绿，备用
    "orange": "#F59E0B",      # 橙色，对比色
    "red": "#EF476F",         # 箱线图均值点使用
    "purple": "#7C3AED",      # 备用
    "light_blue": "#F3F8FF",  # 浅蓝背景
    "light_bg": "#F3F8FF",    # 兼容画图函数里的背景色名称
    "grid": "#D9E6F2",        # 网格线
    "gray": "#64748B",        # 中性灰
}

# 表6-24：三种广告战略用同一蓝色系
STRATEGY_COLORS = [
    "#9CC9F5",
    "#5AA9E6",
    "#2D6CDF",
]

# 表6-25：按媒体分色，不要六种颜色乱飞
COMBINATION_COLORS = {
    "短视频平台": "#5AA9E6",
    "社交网站": "#F59E0B",
}

# 图4：两种广告媒体
MEDIA_COLORS = {
    "短视频平台": "#2D6CDF",
    "社交网站": "#F59E0B",
}


# -----------------------------------------------------------------------------
# 基础工具函数
# -----------------------------------------------------------------------------

def ensure_dirs() -> None:
    """创建输出目录。"""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def setup_chinese_font() -> None:
    """尽量设置中文字体，避免图表中文乱码。"""
    candidate_fonts = [
        "Microsoft YaHei", "SimHei", "PingFang SC", "Noto Sans CJK SC",
        "Source Han Sans SC", "WenQuanYi Micro Hei", "Arial Unicode MS",
        "AR PL UMing CN", "AR PL KaitiM GB"
    ]
    available_fonts = {f.name for f in font_manager.fontManager.ttflist}
    for font in candidate_fonts:
        if font in available_fonts:
            plt.rcParams["font.sans-serif"] = [font]
            break
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["figure.dpi"] = 130
    plt.rcParams["savefig.dpi"] = 180

def style_axis(ax) -> None:
    """统一设置图表风格。"""
    ax.set_facecolor("white")
    ax.grid(axis="y", color=COLORS["grid"], alpha=0.8, linewidth=0.8)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(COLORS["grid"])
    ax.spines["bottom"].set_color(COLORS["grid"])

    ax.tick_params(axis="both", colors=COLORS["navy"], labelsize=9)
    ax.xaxis.label.set_color(COLORS["navy"])
    ax.yaxis.label.set_color(COLORS["navy"])
    ax.title.set_color(COLORS["navy"])


def round_value(x: Any, ndigits: int = 4) -> Any:
    """把 numpy/pandas 数值转换为普通 Python 类型，便于写入 JSON。"""
    if pd.isna(x):
        return None
    if isinstance(x, (np.integer,)):
        return int(x)
    if isinstance(x, (np.floating, float)):
        if math.isfinite(float(x)):
            return round(float(x), ndigits)
        return None
    return x


def df_to_records(df: pd.DataFrame, ndigits: int = 4) -> List[Dict[str, Any]]:
    """DataFrame 转换为适合 JSON 保存的 records。"""
    records: List[Dict[str, Any]] = []
    for record in df.to_dict(orient="records"):
        records.append({str(k): round_value(v, ndigits) for k, v in record.items()})
    return records


def save_csv(df: pd.DataFrame, filename: str) -> Path:
    """保存 CSV 文件。"""
    path = OUTPUT_DIR / filename
    df.to_csv(path, index=False, encoding="utf-8-sig")
    return path


def pvalue_decision(p_value: float | None, alpha: float = ALPHA) -> str:
    """根据 p 值给出拒绝/不拒绝原假设的判断。"""
    if p_value is None or pd.isna(p_value):
        return "无法判断"
    return "拒绝原假设" if p_value < alpha else "不能拒绝原假设"


# -----------------------------------------------------------------------------
# 数据读取与校验
# -----------------------------------------------------------------------------

def read_data(excel_path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    """读取两个工作表。"""
    if not excel_path.exists():
        raise FileNotFoundError(f"找不到数据文件：{excel_path}")

    df_624 = pd.read_excel(excel_path, sheet_name="table_6_24")
    df_625 = pd.read_excel(excel_path, sheet_name="table_6_25")

    df_624 = clean_table_624(df_624)
    df_625 = clean_table_625(df_625)
    validate_data(df_624, df_625)

    return df_624, df_625


def clean_table_624(df: pd.DataFrame) -> pd.DataFrame:
    """清洗表6-24。"""
    expected_cols = ["week", "strategy", "sales"]
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    if list(df.columns) != expected_cols:
        raise ValueError(f"表6-24字段应为 {expected_cols}，实际为 {list(df.columns)}")

    df["week"] = pd.to_numeric(df["week"], errors="raise").astype(int)
    df["strategy"] = df["strategy"].astype(str).str.strip()
    df["sales"] = pd.to_numeric(df["sales"], errors="raise")
    df["strategy"] = pd.Categorical(df["strategy"], categories=STRATEGY_ORDER, ordered=True)
    return df.sort_values(["strategy", "week"]).reset_index(drop=True)


def clean_table_625(df: pd.DataFrame) -> pd.DataFrame:
    """清洗表6-25。"""
    expected_cols = ["week", "strategy", "media", "sales"]
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    if list(df.columns) != expected_cols:
        raise ValueError(f"表6-25字段应为 {expected_cols}，实际为 {list(df.columns)}")

    df["week"] = pd.to_numeric(df["week"], errors="raise").astype(int)
    df["strategy"] = df["strategy"].astype(str).str.strip()
    df["media"] = df["media"].astype(str).str.strip()
    df["sales"] = pd.to_numeric(df["sales"], errors="raise")
    df["strategy"] = pd.Categorical(df["strategy"], categories=STRATEGY_ORDER, ordered=True)
    df["media"] = pd.Categorical(df["media"], categories=MEDIA_ORDER, ordered=True)
    return df.sort_values(["strategy", "media", "week"]).reset_index(drop=True)


def validate_data(df_624: pd.DataFrame, df_625: pd.DataFrame) -> None:
    """检查两张表是否满足题目设计。"""
    errors: List[str] = []

    if len(df_624) != 60:
        errors.append(f"表6-24应有60行，实际{len(df_624)}行。")
    counts_624 = df_624.groupby("strategy", observed=False)["sales"].size().reindex(STRATEGY_ORDER)
    if not (counts_624 == 20).all():
        errors.append(f"表6-24每种策略应各20行，实际为：{counts_624.to_dict()}")
    if set(df_624["week"].unique()) != set(range(1, 21)):
        errors.append("表6-24的week应为1到20。")

    if len(df_625) != 60:
        errors.append(f"表6-25应有60行，实际{len(df_625)}行。")
    counts_625 = (
        df_625.groupby(["strategy", "media"], observed=False)["sales"]
        .size()
        .reset_index(name="n")
    )
    bad_625 = counts_625[counts_625["n"] != 10]
    if not bad_625.empty:
        errors.append("表6-25每个策略×媒体组合应各10行，异常组合为：\n" + bad_625.to_string(index=False))
    if set(df_625["week"].unique()) != set(range(1, 11)):
        errors.append("表6-25的week应为1到10。")

    if errors:
        raise ValueError("数据校验失败：\n" + "\n".join(errors))


# -----------------------------------------------------------------------------
# 统计分析
# -----------------------------------------------------------------------------

def descriptive_table_624(df: pd.DataFrame) -> pd.DataFrame:
    """表6-24描述性统计。"""
    desc = (
        df.groupby("strategy", observed=False)["sales"]
        .agg(n="count", mean="mean", std="std", median="median", min="min", max="max")
        .reindex(STRATEGY_ORDER)
        .reset_index()
    )
    desc["se"] = desc["std"] / np.sqrt(desc["n"])
    desc["ci95"] = stats.t.ppf(0.975, desc["n"] - 1) * desc["se"]
    return desc.round(4)


def descriptive_table_625(df: pd.DataFrame) -> pd.DataFrame:
    """表6-25描述性统计。"""
    desc = (
        df.groupby(["strategy", "media"], observed=False)["sales"]
        .agg(n="count", mean="mean", std="std", median="median", min="min", max="max")
        .reset_index()
    )
    desc["se"] = desc["std"] / np.sqrt(desc["n"])
    desc["ci95"] = stats.t.ppf(0.975, desc["n"] - 1) * desc["se"]
    return desc.round(4)


def one_way_anova(df: pd.DataFrame) -> tuple[pd.DataFrame, Dict[str, Any], Any]:
    """单因素方差分析：检验三种广告战略下平均销量是否相同。"""
    groups = [df.loc[df["strategy"] == s, "sales"].values for s in STRATEGY_ORDER]
    f_stat, p_value = stats.f_oneway(*groups)

    model = ols("sales ~ C(strategy)", data=df).fit()
    residuals = model.resid

    grand_mean = df["sales"].mean()
    group_stats = df.groupby("strategy", observed=False)["sales"].agg(["count", "mean"]).reindex(STRATEGY_ORDER)
    ss_between = float(sum(group_stats["count"] * (group_stats["mean"] - grand_mean) ** 2))
    ss_within = float(sum(((g - g.mean()) ** 2).sum() for g in groups))
    ss_total = float(((df["sales"] - grand_mean) ** 2).sum())
    df_between = len(STRATEGY_ORDER) - 1
    df_within = len(df) - len(STRATEGY_ORDER)
    df_total = len(df) - 1
    ms_between = ss_between / df_between
    ms_within = ss_within / df_within

    anova_table = pd.DataFrame({
        "来源": ["组间：广告战略", "组内：误差", "总计"],
        "平方和SS": [ss_between, ss_within, ss_total],
        "自由度df": [df_between, df_within, df_total],
        "均方MS": [ms_between, ms_within, np.nan],
        "F值": [f_stat, np.nan, np.nan],
        "p值": [p_value, np.nan, np.nan],
    }).round(4)

    levene_stat, levene_p = stats.levene(*groups, center="median")
    shapiro_stat, shapiro_p = stats.shapiro(residuals)

    summary = {
        "f_stat": round_value(f_stat),
        "p_value": round_value(p_value),
        "alpha": ALPHA,
        "decision": pvalue_decision(float(p_value)),
        "levene_stat": round_value(levene_stat),
        "levene_p": round_value(levene_p),
        "shapiro_stat": round_value(shapiro_stat),
        "shapiro_p": round_value(shapiro_p),
        "conclusion": (
            "在5%显著性水平下，三种广告战略下的平均周销量存在显著差异。"
            if p_value < ALPHA
            else "在5%显著性水平下，不能认为三种广告战略下的平均周销量存在显著差异。"
        ),
    }
    return anova_table, summary, model


def two_way_anova(df: pd.DataFrame) -> tuple[pd.DataFrame, Dict[str, Any], Any]:
    """双因素方差分析：检验广告策略、广告媒体及交互作用。"""
    model = ols("sales ~ C(strategy) * C(media)", data=df).fit()
    raw_table = sm.stats.anova_lm(model, typ=2)

    name_map = {
        "C(strategy)": "广告策略",
        "C(media)": "广告媒体",
        "C(strategy):C(media)": "策略×媒体交互作用",
        "Residual": "误差",
    }
    anova_table = raw_table.reset_index().rename(columns={
        "index": "来源",
        "sum_sq": "平方和SS",
        "df": "自由度df",
        "F": "F值",
        "PR(>F)": "p值",
    })
    anova_table["来源"] = anova_table["来源"].map(name_map).fillna(anova_table["来源"])
    anova_table["均方MS"] = anova_table["平方和SS"] / anova_table["自由度df"]
    anova_table = anova_table[["来源", "平方和SS", "自由度df", "均方MS", "F值", "p值"]].round(4)

    # 残差检验
    residuals = model.resid
    shapiro_stat, shapiro_p = stats.shapiro(residuals)
    groups = [
        g["sales"].values
        for _, g in df.groupby(["strategy", "media"], observed=False)
    ]
    levene_stat, levene_p = stats.levene(*groups, center="median")

    def get_p(source: str) -> float | None:
        row = anova_table[anova_table["来源"] == source]
        if row.empty:
            return None
        return float(row.iloc[0]["p值"])

    strategy_p = get_p("广告策略")
    media_p = get_p("广告媒体")
    interaction_p = get_p("策略×媒体交互作用")

    summary = {
        "alpha": ALPHA,
        "strategy_p": round_value(strategy_p),
        "media_p": round_value(media_p),
        "interaction_p": round_value(interaction_p),
        "strategy_decision": pvalue_decision(strategy_p),
        "media_decision": pvalue_decision(media_p),
        "interaction_decision": pvalue_decision(interaction_p),
        "levene_stat": round_value(levene_stat),
        "levene_p": round_value(levene_p),
        "shapiro_stat": round_value(shapiro_stat),
        "shapiro_p": round_value(shapiro_p),
        "conclusion": build_two_way_conclusion(strategy_p, media_p, interaction_p),
    }
    return anova_table, summary, model


def build_two_way_conclusion(strategy_p: float | None, media_p: float | None, interaction_p: float | None) -> str:
    """生成双因素方差分析结论。"""
    parts = []
    if strategy_p is not None:
        parts.append("广告策略主效应显著" if strategy_p < ALPHA else "广告策略主效应不显著")
    if media_p is not None:
        parts.append("广告媒体主效应显著" if media_p < ALPHA else "广告媒体主效应不显著")
    if interaction_p is not None:
        parts.append("策略与媒体的交互作用显著" if interaction_p < ALPHA else "策略与媒体的交互作用不显著")
    return "在5%显著性水平下，" + "、".join(parts) + "。"


# -----------------------------------------------------------------------------
# 可视化
# -----------------------------------------------------------------------------

def plot_strategy_mean(desc_624: pd.DataFrame) -> Path:
    """图1：表6-24三种广告战略平均销量柱状图。"""
    path = OUTPUT_DIR / "fig_1_strategy_mean.png"

    fig, ax = plt.subplots(figsize=(7.2, 4.8))
    fig.patch.set_facecolor(COLORS["light_bg"])

    bars = ax.bar(
        desc_624["strategy"].astype(str),
        desc_624["mean"],
        yerr=desc_624["ci95"],
        capsize=5,
        color=STRATEGY_COLORS,
        edgecolor="white",
        linewidth=1.2,
    )

    ax.set_title("表6-24：三种广告战略的平均周销量", fontsize=13, fontweight="bold", pad=14)
    ax.set_xlabel("广告战略")
    ax.set_ylabel("平均周销量")
    style_axis(ax)

    for bar, value in zip(bars, desc_624["mean"]):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            value + 6,
            f"{value:.2f}",
            ha="center",
            va="bottom",
            fontsize=9,
            color=COLORS["navy"],
            fontweight="bold",
        )

    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return path


def plot_strategy_boxplot(df_624: pd.DataFrame) -> Path:
    """图2：表6-24三种广告战略销量分布箱线图。"""
    path = OUTPUT_DIR / "fig_2_strategy_boxplot.png"

    fig, ax = plt.subplots(figsize=(7.2, 4.8))
    fig.patch.set_facecolor(COLORS["light_bg"])

    data = [df_624.loc[df_624["strategy"] == s, "sales"].values for s in STRATEGY_ORDER]
    box = ax.boxplot(
        data,
        tick_labels=STRATEGY_ORDER,
        showmeans=True,
        patch_artist=True,
        meanprops={
            "marker": "D",
            "markerfacecolor": COLORS["red"],
            "markeredgecolor": "white",
            "markersize": 6,
        },
        medianprops={
            "color": COLORS["navy"],
            "linewidth": 1.5,
        },
        whiskerprops={
            "color": COLORS["gray"],
            "linewidth": 1.2,
        },
        capprops={
            "color": COLORS["gray"],
            "linewidth": 1.2,
        },
    )

    for patch, color in zip(box["boxes"], STRATEGY_COLORS):
        patch.set_facecolor(color)
        patch.set_alpha(0.72)
        patch.set_edgecolor("white")
        patch.set_linewidth(1.2)

    ax.set_title("表6-24：三种广告战略的周销量分布", fontsize=13, fontweight="bold", pad=14)
    ax.set_xlabel("广告战略")
    ax.set_ylabel("周销量")
    style_axis(ax)

    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return path


def plot_combination_mean(desc_625: pd.DataFrame) -> Path:
    """图3：表6-25六种策略×媒体组合平均销量柱状图。"""
    path = OUTPUT_DIR / "fig_3_combination_mean.png"
    plot_df = desc_625.copy()
    plot_df["组合"] = plot_df["strategy"].astype(str) + "\n" + plot_df["media"].astype(str)

    # 按广告媒体分色：短视频平台统一蓝色，社交网站统一橙色
    # 不能把 COMBINATION_COLORS 这个字典直接传给 ax.bar 的 color 参数。
    bar_colors = [
        COMBINATION_COLORS[str(media)]
        for media in plot_df["media"]
    ]

    fig, ax = plt.subplots(figsize=(9.5, 5.2))
    fig.patch.set_facecolor(COLORS["light_bg"])

    bars = ax.bar(
        plot_df["组合"],
        plot_df["mean"],
        yerr=plot_df["ci95"],
        capsize=4,
        color=bar_colors,
        edgecolor="white",
        linewidth=1.2,
    )

    ax.set_title("表6-25：广告策略×广告媒体组合的平均周销量", fontsize=13, fontweight="bold", pad=14)
    ax.set_xlabel("策略×媒体组合")
    ax.set_ylabel("平均周销量")
    ax.tick_params(axis="x", labelsize=8)
    style_axis(ax)

    for bar, value in zip(bars, plot_df["mean"]):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            value + 6,
            f"{value:.2f}",
            ha="center",
            va="bottom",
            fontsize=8,
            color=COLORS["navy"],
            fontweight="bold",
        )

    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return path

def plot_interaction(desc_625: pd.DataFrame) -> Path:
    """图4：广告策略与广告媒体组合均值对比图。

    使用分组柱状图展示三种广告策略在两种广告媒体下的平均周销量。
    不在图内写长结论，避免遮挡数据；结论放在报告正文中说明。
    """
    path = OUTPUT_DIR / "fig_4_interaction.png"

    plot_df = desc_625.copy()

    mean_pivot = (
        plot_df.pivot(index="strategy", columns="media", values="mean")
        .reindex(STRATEGY_ORDER)
    )
    ci_pivot = (
        plot_df.pivot(index="strategy", columns="media", values="ci95")
        .reindex(STRATEGY_ORDER)
    )

    x = np.arange(len(STRATEGY_ORDER))
    width = 0.34

    bg_color = COLORS.get("light_bg", COLORS.get("light_blue", "#F3F8FF"))
    navy = COLORS.get("navy", "#0F2742")
    grid = COLORS.get("grid", "#D9E6F2")
    gray = COLORS.get("gray", "#64748B")

    fig, ax = plt.subplots(figsize=(8.8, 5.6))
    fig.patch.set_facecolor(bg_color)

    error_style = {
        "ecolor": gray,
        "elinewidth": 1.3,
        "capsize": 4,
        "capthick": 1.3,
    }

    bars_video = ax.bar(
        x - width / 2,
        mean_pivot["短视频平台"],
        width,
        yerr=ci_pivot["短视频平台"],
        error_kw=error_style,
        label="短视频平台",
        color=MEDIA_COLORS["短视频平台"],
        edgecolor="white",
        linewidth=1.2,
    )

    bars_social = ax.bar(
        x + width / 2,
        mean_pivot["社交网站"],
        width,
        yerr=ci_pivot["社交网站"],
        error_kw=error_style,
        label="社交网站",
        color=MEDIA_COLORS["社交网站"],
        edgecolor="white",
        linewidth=1.2,
    )

    ax.set_title(
        "表6-25：广告策略与广告媒体组合的平均周销量",
        fontsize=13,
        fontweight="bold",
        pad=16,
        color=navy,
    )
    ax.set_xlabel("广告策略")
    ax.set_ylabel("平均周销量")
    ax.set_xticks(x)
    ax.set_xticklabels(STRATEGY_ORDER)

    # 设置纵轴范围，给数值标签和误差线留空间
    y_max = max(
        (mean_pivot["短视频平台"] + ci_pivot["短视频平台"]).max(),
        (mean_pivot["社交网站"] + ci_pivot["社交网站"]).max(),
    )
    y_min = min(
        (mean_pivot["短视频平台"] - ci_pivot["短视频平台"]).min(),
        (mean_pivot["社交网站"] - ci_pivot["社交网站"]).min(),
    )
    ax.set_ylim(y_min - 35, y_max + 55)

    style_axis(ax)

    # 添加数值标签
    for bars in [bars_video, bars_social]:
        for bar in bars:
            value = bar.get_height()
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                value + 8,
                f"{value:.1f}",
                ha="center",
                va="bottom",
                fontsize=9,
                color=navy,
                fontweight="bold",
            )

    # 图例放到底部，避免遮挡柱形和标签
    legend = ax.legend(
        title="广告媒体",
        loc="upper center",
        bbox_to_anchor=(0.5, -0.12),
        ncol=2,
        frameon=True,
    )
    legend.get_frame().set_facecolor("white")
    legend.get_frame().set_edgecolor(grid)

    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)

    return path


def generate_plots(df_624: pd.DataFrame, df_625: pd.DataFrame, desc_624: pd.DataFrame, desc_625: pd.DataFrame) -> Dict[str, str]:
    """生成全部图表。"""
    paths = {
        "fig_1_strategy_mean": plot_strategy_mean(desc_624),
        "fig_2_strategy_boxplot": plot_strategy_boxplot(df_624),
        "fig_3_combination_mean": plot_combination_mean(desc_625),
        "fig_4_interaction": plot_interaction(desc_625),
    }
    return {k: str(v.relative_to(PROJECT_ROOT)) for k, v in paths.items()}


# -----------------------------------------------------------------------------
# 主流程
# -----------------------------------------------------------------------------

def run_analysis(excel_path: Path) -> Dict[str, Any]:
    """执行完整分析，并返回结果字典。"""
    ensure_dirs()
    setup_chinese_font()

    df_624, df_625 = read_data(excel_path)

    # 保存清洗后的原始数据，方便复核。
    save_csv(df_624, "table_6_24_clean.csv")
    save_csv(df_625, "table_6_25_clean.csv")

    desc_624 = descriptive_table_624(df_624)
    desc_625 = descriptive_table_625(df_625)
    anova_624, summary_624, _ = one_way_anova(df_624)
    anova_625, summary_625, _ = two_way_anova(df_625)

    save_csv(desc_624, "desc_table_6_24.csv")
    save_csv(desc_625, "desc_table_6_25.csv")
    save_csv(anova_624, "anova_oneway_table_6_24.csv")
    save_csv(anova_625, "anova_twoway_table_6_25.csv")

    figure_paths = generate_plots(df_624, df_625, desc_624, desc_625)

    results: Dict[str, Any] = {
        "project": "管理统计学：广告战略方差分析",
        "alpha": ALPHA,
        "input_file": str(excel_path.relative_to(PROJECT_ROOT)) if excel_path.is_relative_to(PROJECT_ROOT) else str(excel_path),
        "data_check": {
            "table_6_24_rows": int(len(df_624)),
            "table_6_25_rows": int(len(df_625)),
            "table_6_24_strategy_counts": {
                str(k): int(v) for k, v in df_624.groupby("strategy", observed=False).size().reindex(STRATEGY_ORDER).items()
            },
            "table_6_25_combination_counts": df_to_records(
                df_625.groupby(["strategy", "media"], observed=False).size().reset_index(name="n")
            ),
        },
        "desc_624": df_to_records(desc_624),
        "desc_625": df_to_records(desc_625),
        "anova_624": df_to_records(anova_624),
        "anova_625": df_to_records(anova_625),
        "summary_624": summary_624,
        "summary_625": summary_625,
        "figures": figure_paths,
        "interpretation": build_interpretation(summary_624, summary_625),
    }

    results_path = OUTPUT_DIR / "results.json"
    with open(results_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    return results


def build_interpretation(summary_624: Dict[str, Any], summary_625: Dict[str, Any]) -> Dict[str, str]:
    """生成报告中会复用的文字解释。"""
    media_p = summary_625.get("media_p")
    media_extra = ""
    if media_p is not None and ALPHA <= float(media_p) < 0.10:
        media_extra = "不过，广告媒体的p值接近0.05，说明社交网站可能具有更高销量表现，但在5%水平下证据仍不足。"

    return {
        "oneway": summary_624["conclusion"] + "因此，不能仅凭样本均值直接断定某一种广告战略显著优于其他战略。",
        "twoway": summary_625["conclusion"] + media_extra,
        "management": "从管理决策角度看，样本均值可以作为参考，但正式决策应以显著性检验为依据。本题数据下，低价格策略和社交网站组合在描述性均值上表现较好，但统计证据并不足以支持其必然优越。建议企业扩大样本范围、延长观察期，并进一步控制城市差异和时间趋势。",
        "method_discussion": "方差分析要求观测值相互独立、各组误差近似正态、组间方差齐性。本案例虽然假定城市条件相同，但现实中城市消费水平、人口结构、竞争状况和周销量时间相关性都可能影响检验结果。因此，结论应被理解为试验样本下的统计判断，而不是绝对的市场规律。",
    }


def print_console_summary(results: Dict[str, Any]) -> None:
    """在控制台打印关键结果。"""
    print("\n========== 数据校验 ==========")
    print(f"表6-24行数：{results['data_check']['table_6_24_rows']}")
    print(f"表6-25行数：{results['data_check']['table_6_25_rows']}")

    print("\n========== 单因素方差分析：表6-24 ==========")
    print(pd.DataFrame(results["anova_624"]).to_string(index=False))
    print("结论：", results["summary_624"]["conclusion"])

    print("\n========== 双因素方差分析：表6-25 ==========")
    print(pd.DataFrame(results["anova_625"]).to_string(index=False))
    print("结论：", results["summary_625"]["conclusion"])

    print("\n输出目录：", OUTPUT_DIR)
    print("结果文件：", OUTPUT_DIR / "results.json")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="广告战略与广告媒体方差分析")
    parser.add_argument("--input", type=str, default=str(DEFAULT_INPUT), help="Excel数据文件路径")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    excel_path = Path(args.input).resolve()
    results = run_analysis(excel_path)
    print_console_summary(results)


if __name__ == "__main__":
    main()
