# -*- coding: utf-8 -*-
"""
02_generate_html.py
根据 01_analysis.py 输出的 results.json 和图片，生成完整 HTML 报告。

运行前请先运行：
    python 01_analysis.py

然后运行：
    python 02_generate_html.py
"""

from __future__ import annotations

import base64
import html
import json
from pathlib import Path
from typing import Any, Dict, List

PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "outputs"
REPORT_DIR = PROJECT_ROOT / "reports"
RESULTS_PATH = OUTPUT_DIR / "results.json"
HTML_PATH = REPORT_DIR / "广告战略方差分析报告.html"


def load_results() -> Dict[str, Any]:
    if not RESULTS_PATH.exists():
        raise FileNotFoundError("找不到 outputs/results.json。请先运行 python 01_analysis.py")
    with open(RESULTS_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def fmt(value: Any) -> str:
    """HTML表格中的数值格式化。"""
    if value is None:
        return ""
    if isinstance(value, float):
        return f"{value:.4f}"
    return html.escape(str(value))


def num(value: Any, ndigits: int = 4) -> str:
    """正文中的数值格式化。"""
    if value is None:
        return ""
    try:
        return f"{float(value):.{ndigits}f}"
    except (TypeError, ValueError):
        return str(value)


def table_html(records: List[Dict[str, Any]], caption: str) -> str:
    """records生成HTML表格。"""
    if not records:
        return ""
    columns = list(records[0].keys())
    thead = "".join(f"<th>{html.escape(str(c))}</th>" for c in columns)
    rows = []
    for record in records:
        tds = "".join(f"<td>{fmt(record.get(c))}</td>" for c in columns)
        rows.append(f"<tr>{tds}</tr>")
    tbody = "\n".join(rows)
    return f"""
    <figure class="table-figure">
      <figcaption>{html.escape(caption)}</figcaption>
      <div class="table-wrap">
        <table>
          <thead><tr>{thead}</tr></thead>
          <tbody>{tbody}</tbody>
        </table>
      </div>
    </figure>
    """


def image_to_base64(relative_path: str) -> str:
    """将图片转成base64，保证HTML尽量自包含。"""
    path = PROJECT_ROOT / relative_path
    if not path.exists():
        raise FileNotFoundError(f"找不到图片：{path}")
    return base64.b64encode(path.read_bytes()).decode("ascii")


def image_html(relative_path: str, caption: str) -> str:
    img64 = image_to_base64(relative_path)
    return f"""
    <figure class="image-figure">
      <img src="data:image/png;base64,{img64}" alt="{html.escape(caption)}" />
      <figcaption>{html.escape(caption)}</figcaption>
    </figure>
    """


def build_html(results: Dict[str, Any]) -> str:
    s624 = results["summary_624"]
    s625 = results["summary_625"]
    interp = results["interpretation"]
    figs = results["figures"]

    f624 = num(s624.get("f_stat"))
    p624 = num(s624.get("p_value"))
    p_strategy = num(s625.get("strategy_p"))
    p_media = num(s625.get("media_p"))
    p_interaction = num(s625.get("interaction_p"))
    alpha = results.get("alpha", 0.05)

    html_text = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>广告战略方差分析报告</title>
  <style>
    :root {{
      --bg: #f3f8ff;
      --card: #ffffff;
      --text: #0f2742;
      --muted: #64748b;
      --line: #d9e6f2;
      --blue: #2d6cdf;
      --sky: #5aa9e6;
      --orange: #f59e0b;
      --soft-blue: #eaf4ff;
      --soft-orange: #fff7e6;
      --soft-green: #f0fdf4;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", Arial, sans-serif;
      background: linear-gradient(180deg, #eaf4ff 0%, #f7fbff 42%, #ffffff 100%);
      color: var(--text);
      line-height: 1.72;
    }}
    header {{
      padding: 46px 24px 34px;
      background: linear-gradient(135deg, #0f2742 0%, #2d6cdf 62%, #5aa9e6 100%);
      color: white;
      text-align: center;
    }}
    header h1 {{ margin: 0 0 12px; font-size: 32px; }}
    header p {{ margin: 0; opacity: .92; }}
    main {{ max-width: 1100px; margin: 0 auto; padding: 28px 18px 60px; }}
    section {{
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 26px;
      margin-bottom: 22px;
      box-shadow: 0 10px 28px rgba(15, 39, 66, .06);
    }}
    h2 {{
      margin-top: 0;
      border-left: 6px solid var(--blue);
      padding-left: 12px;
      font-size: 22px;
    }}
    h3 {{ margin-top: 22px; font-size: 18px; }}
    .answer-grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; }}
    .answer-card {{
      border: 1px solid var(--line);
      border-radius: 16px;
      padding: 18px;
      background: #ffffff;
      box-shadow: 0 6px 18px rgba(45, 108, 223, .06);
    }}
    .answer-card h3 {{ margin: 0 0 10px; color: var(--blue); font-size: 17px; }}
    .direct-answer {{
      margin-top: 12px;
      padding: 12px 14px;
      border-radius: 12px;
      background: var(--soft-blue);
      border-left: 5px solid var(--blue);
      font-weight: 700;
    }}
    .formula {{
      display: inline-block;
      background: #f8fafc;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 2px 8px;
      margin: 2px 0;
      font-family: "Times New Roman", serif;
    }}
    .kpi-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin: 18px 0; }}
    .kpi {{ background: var(--soft-blue); border-radius: 14px; padding: 16px; }}
    .kpi .label {{ color: var(--muted); font-size: 13px; }}
    .kpi .value {{ font-size: 22px; font-weight: 700; color: var(--blue); }}
    .note {{ background: var(--soft-orange); border: 1px solid #f8d98b; padding: 14px 16px; border-radius: 12px; }}
    .conclusion {{ background: var(--soft-green); border: 1px solid #bbf7d0; padding: 14px 16px; border-radius: 12px; }}
    .table-wrap {{ overflow-x: auto; }}
    table {{ border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 14px; }}
    th, td {{ border: 1px solid var(--line); padding: 8px 10px; text-align: center; }}
    th {{ background: #eaf4ff; font-weight: 700; color: var(--text); }}
    figcaption {{ color: var(--muted); font-size: 14px; margin: 8px 0; text-align: center; }}
    .image-figure img {{ display: block; max-width: 100%; margin: 12px auto; border-radius: 10px; border: 1px solid var(--line); }}
    code {{ background: #f3f4f6; padding: 2px 6px; border-radius: 5px; }}
    footer {{ text-align: center; color: var(--muted); margin-top: 26px; font-size: 13px; }}
    @media (max-width: 760px) {{
      .answer-grid, .kpi-grid {{ grid-template-columns: 1fr; }}
      header h1 {{ font-size: 26px; }}
      section {{ padding: 20px; }}
    }}
  </style>
</head>
<body>
<header>
  <h1>管理统计学作业：广告战略方差分析</h1>
  <p>基于表6-24与表6-25，对广告战略、广告媒体及其组合效果进行显著性检验</p>
</header>
<main>
  <section>
    <h2>一、讨论题直接作答</h2>
    <div class="answer-grid">
      <div class="answer-card">
        <h3>讨论题1：显著性检验假设</h3>
        <p>针对表6-24，检验三种广告战略下平均周销量是否相同。</p>
        <p><span class="formula">H<sub>0</sub>: μ<sub>便利性</sub> = μ<sub>高质量</sub> = μ<sub>低价格</sub></span></p>
        <p><span class="formula">H<sub>1</sub>: 三种广告战略均值不全相同</span></p>
        <p>针对表6-25，应分别检验广告策略主效应、广告媒体主效应和策略×媒体交互作用。</p>
        <div class="direct-answer">答案：本题应设置单因素方差分析假设和双因素方差分析假设。</div>
      </div>
      <div class="answer-card">
        <h3>讨论题2：三种广告战略均值是否相同</h3>
        <p>单因素方差分析结果为：F = {f624}，p = {p624}。</p>
        <p>由于 p = {p624} &gt; {alpha}，不能拒绝原假设。</p>
        <div class="direct-answer">答案：三种广告战略下的销售量均值没有显著差异。</div>
      </div>
      <div class="answer-card">
        <h3>讨论题3：策略与媒体组合均值是否相同</h3>
        <p>双因素方差分析结果为：广告策略 p = {p_strategy}，广告媒体 p = {p_media}，策略×媒体交互作用 p = {p_interaction}。</p>
        <p>三者均大于 {alpha}，因此不能拒绝相关原假设。</p>
        <div class="direct-answer">答案：在5%显著性水平下，广告策略、广告媒体及其交互作用均不显著，不能认为各组合销售量均值存在显著差异。</div>
      </div>
      <div class="answer-card">
        <h3>讨论题4：检验方法讨论</h3>
        <p>表6-24只有一个因素，适合使用单因素方差分析；表6-25包含广告策略和广告媒体两个因素，适合使用双因素方差分析。</p>
        <p>但方差分析依赖独立性、近似正态性和方差齐性。现实中城市差异和周销量时间相关性可能影响检验可靠性。</p>
        <div class="direct-answer">答案：方差分析方法适用，但结论应结合样本量、城市控制和时间相关性谨慎解释。</div>
      </div>
    </div>
  </section>

  <section>
    <h2>二、问题背景与研究目的</h2>
    <p>某厂家开发了一种新产品，营销经理需要在“便利性”“高质量”“低价格”三种广告战略之间进行选择。进一步地，题目还引入了两种广告媒体：短视频平台和社交网站。因此，本报告分别采用单因素方差分析和双因素方差分析，检验不同广告方案下的平均周销量是否存在显著差异。</p>
    <div class="kpi-grid">
      <div class="kpi"><div class="label">表6-24样本量</div><div class="value">{results['data_check']['table_6_24_rows']}</div></div>
      <div class="kpi"><div class="label">表6-25样本量</div><div class="value">{results['data_check']['table_6_25_rows']}</div></div>
      <div class="kpi"><div class="label">显著性水平</div><div class="value">α = {results['alpha']}</div></div>
    </div>
  </section>

  <section>
    <h2>三、检验假设展开</h2>
    <h3>1. 三种广告战略的单因素方差分析</h3>
    <p>原假设 H<sub>0</sub>：三种广告战略下的平均周销量相同，即 μ<sub>便利性</sub> = μ<sub>高质量</sub> = μ<sub>低价格</sub>。</p>
    <p>备择假设 H<sub>1</sub>：三种广告战略下的平均周销量不全相同。</p>
    <h3>2. 广告策略与广告媒体的双因素方差分析</h3>
    <p>广告策略主效应：检验三种广告策略的平均销量是否相同。</p>
    <p>广告媒体主效应：检验短视频平台与社交网站的平均销量是否相同。</p>
    <p>交互作用：检验广告策略的效果是否会因广告媒体不同而发生显著变化。</p>
  </section>

  <section>
    <h2>四、表6-24：三种广告战略的单因素方差分析</h2>
    {table_html(results['desc_624'], '表1  表6-24描述性统计')}
    {image_html(figs['fig_1_strategy_mean'], '图1  三种广告战略的平均周销量及95%置信区间')}
    {image_html(figs['fig_2_strategy_boxplot'], '图2  三种广告战略的周销量分布箱线图')}
    {table_html(results['anova_624'], '表2  单因素方差分析表')}
    <div class="conclusion"><strong>直接结论：</strong>F = {f624}，p = {p624}。在 α = {alpha} 下，p值大于0.05，因此三种广告战略均值不存在显著差异。</div>
    <p class="note">方差齐性检验 Levene p = {num(s624.get('levene_p'))}；残差正态性检验 Shapiro-Wilk p = {num(s624.get('shapiro_p'))}。这些检验用于辅助判断方差分析前提是否基本满足。</p>
  </section>

  <section>
    <h2>五、表6-25：广告策略×广告媒体的双因素方差分析</h2>
    {table_html(results['desc_625'], '表3  表6-25各组合描述性统计')}
    {image_html(figs['fig_3_combination_mean'], '图3  广告策略×广告媒体组合的平均周销量')}
    {image_html(figs['fig_4_interaction'], '图4  广告策略与广告媒体组合的平均周销量对比')}
    {table_html(results['anova_625'], '表4  双因素方差分析表')}
    <div class="conclusion"><strong>直接结论：</strong>广告策略 p = {p_strategy}，广告媒体 p = {p_media}，策略×媒体交互作用 p = {p_interaction}。三者均大于0.05，因此广告策略、广告媒体和二者交互作用均不显著。</div>
    <p class="note">广告媒体的p值接近0.05，说明社交网站可能有更高销量表现，但在5%显著性水平下证据仍不足。方差齐性检验 Levene p = {num(s625.get('levene_p'))}；残差正态性检验 Shapiro-Wilk p = {num(s625.get('shapiro_p'))}。</p>
  </section>

  <section>
    <h2>六、管理解释与方法讨论</h2>
    <h3>1. 管理解释</h3>
    <p>{html.escape(interp['management'])}</p>
    <h3>2. 方法讨论</h3>
    <p>{html.escape(interp['method_discussion'])}</p>
  </section>

  <section>
    <h2>七、最终结论</h2>
    <p>综合单因素和双因素方差分析结果，在5%显著性水平下，本题数据尚不足以证明三种广告战略、两种广告媒体以及二者交互作用会导致平均周销量的显著差异。描述性统计可以提示可能的市场方向，但正式决策仍需更多样本和更严格的试验控制。</p>
  </section>

  <footer>本报告由 Python 统计分析脚本自动生成。</footer>
</main>
</body>
</html>"""
    return html_text


def main() -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    results = load_results()
    html_text = build_html(results)
    HTML_PATH.write_text(html_text, encoding="utf-8")
    print(f"HTML报告已生成：{HTML_PATH}")


if __name__ == "__main__":
    main()
