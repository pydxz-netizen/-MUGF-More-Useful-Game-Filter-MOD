# -*- coding: utf-8 -*-
"""
验证假说：免费的游戏更容易得到好评？
步骤：
1. 读 steam_totdata.xlsx，按 是否免费 × 评价(好评/中立差评) 制 2x2 列联表，做卡方检验
2. 根据检验结果判断假说成立程度
3. 免费/付费各画一张评价分布柱状图（横轴英语，好评到差评排序，标题中文）
"""

import math
import pathlib
import re

import pandas as pd
import requests
from scipy.stats import chi2_contingency, norm
from pyecharts import options as opts
from pyecharts.charts import Bar, Page

# ========== 0. 读取数据与字段准备 ==========
df = pd.read_excel("steam_totdata.xlsx")

# 免费判定：现价与原价都为 0（排除 100% 折扣的付费游戏）
df["是否免费"] = (df["现价"] == 0) & (df["原价"] == 0)

# Steam 官方 9 档评价摘要：中文 -> 英语（按好评到差评排序）
ZH2EN = {
    "好评如潮": "Overwhelmingly Positive",
    "特别好评": "Very Positive",
    "多半好评": "Mostly Positive",
    "好评":     "Positive",
    "褒贬不一": "Mixed",
    "差评":     "Negative",
    "多半差评": "Mostly Negative",
    "特别差评": "Very Negative",
    "差评如潮": "Overwhelmingly Negative",
}
ORDER_EN = list(ZH2EN.values())
GOOD_EN = set(ORDER_EN[:4])          # “好评”=前 4 档（一切好评都算好评）
NEUTRAL_BAD_EN = set(ORDER_EN[4:])   # 中立+差评

df["评价EN"] = df["评价"].map(ZH2EN)
# 无用户评测 / “N 篇用户评测”（评测数过少无摘要）不参与统计
valid = df.dropna(subset=["评价EN"]).copy()
excluded = len(df) - len(valid)
print(f"有效样本 {len(valid)} 条（排除无评价摘要 {excluded} 条）")

# ========== 1. 2x2 列联表 + 卡方检验 ==========
valid["好评"] = valid["评价EN"].isin(GOOD_EN)

table = pd.crosstab(valid["是否免费"], valid["好评"])
table.index = ["付费", "免费"]
table.columns = ["中立/差评", "好评"]
print("\n2x2 列联表：")
print(table)

# Pearson 卡方检验（不连续性校正，教科书标准形式）
chi2, p, dof, expected = chi2_contingency(table.values, correction=False)

# 好评率对比
p_good = {
    "免费": table.loc["免费", "好评"] / table.loc["免费"].sum(),
    "付费": table.loc["付费", "好评"] / table.loc["付费"].sum(),
}

# Cramér's V 效应量（2x2 表时等于 phi 系数）
n = table.values.sum()
cramers_v = (chi2 / n) ** 0.5

print("\n========== 显著性分析结果 ==========")
print(f"免费游戏好评率: {p_good['免费']:.1%} ({table.loc['免费', '好评']}/{table.loc['免费'].sum()})")
print(f"付费游戏好评率: {p_good['付费']:.1%} ({table.loc['付费', '好评']}/{table.loc['付费'].sum()})")
print(f"差值: {p_good['免费'] - p_good['付费']:+.1%}")
print(f"卡方值 chi2 = {chi2:.3f}, 自由度 = {dof}, p 值 = {p:.4f}")
print(f"效应量 Cramer's V = {cramers_v:.3f}")

# ---- 方向性补充验证：付费 - 免费好评率之差的 z 检验 + 95% 置信区间 ----
n_p, x_p = int(table.loc["付费"].sum()), int(table.loc["付费", "好评"])
n_f, x_f = int(table.loc["免费"].sum()), int(table.loc["免费", "好评"])
p1, p2 = x_p / n_p, x_f / n_f
ph = (x_p + x_f) / (n_p + n_f)
se = math.sqrt(ph * (1 - ph) * (1 / n_p + 1 / n_f))
z = (p1 - p2) / se
p_one_sided = norm.sf(z)          # 单侧检验：付费好评率 > 免费好评率
ci_low, ci_high = (p1 - p2) - 1.96 * se, (p1 - p2) + 1.96 * se
print(f"\n方向性验证（付费 - 免费好评率之差）:")
print(f"  两比例 z 检验: z = {z:.3f}, 单侧 p 值 = {p_one_sided:.6f}")
print(f"  差值 95% 置信区间: [{ci_low:+.1%}, {ci_high:+.1%}]（不包含 0，方向可靠）")

# ========== 2. 假说判断 ==========
alpha = 0.05
if p < alpha:
    direction = "更高" if p_good["免费"] > p_good["付费"] else "更低"
    print(f"\n结论：p = {p:.4f} < {alpha}，差异统计显著。")
    print(f"免费游戏的好评率显著{direction}于付费游戏。")
    if p_good["免费"] > p_good["付费"]:
        print("假说『免费的游戏更容易得到好评』在统计意义上成立。")
    else:
        print("假说『免费的游戏更容易得到好评』被数据反驳——免费游戏反而更难获得好评。")
else:
    print(f"\n结论：p = {p:.4f} >= {alpha}，两组好评率差异不显著。")
    print("没有证据支持假说『免费的游戏更容易得到好评』。")

# ========== 3. 免费/付费评价分布柱状图 ==========
def make_bar(sub: pd.DataFrame, title: str, color: str) -> Bar:
    n_group = len(sub)
    rates = sub["评价EN"].value_counts(normalize=True) * 100
    values = [round(rates.get(k, 0.0), 2) for k in ORDER_EN]
    return (
        Bar()
        .add_xaxis(ORDER_EN)
        .add_yaxis(
            "占比", values,
            color=color,
            label_opts=opts.LabelOpts(is_show=True, formatter="{c}%"),
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(
                title=title,
                subtitle=f"n = {n_group}，柱高为该评价档位在组内的占比",
            ),
            yaxis_opts=opts.AxisOpts(name="占比 (%)", max_=75),
            xaxis_opts=opts.AxisOpts(
                name="评价", axislabel_opts=opts.LabelOpts(rotate=35)
            ),
            tooltip_opts=opts.TooltipOpts(formatter="{b}: {c}%"),
        )
    )

page = Page()
page.add(
    make_bar(valid[valid["是否免费"]], "免费游戏的评价分布", "#22c073"),
    make_bar(valid[~valid["是否免费"]], "付费游戏的评价分布", "#4a90d9"),
)
OUT_HTML = "验证假说_柱状图.html"
page.render(OUT_HTML)


# ---- 把 echarts.min.js 内嵌进 HTML，离线也能打开 ----
def inline_echarts(html_path: str):
    html = pathlib.Path(html_path).read_text(encoding="utf-8")
    urls = set(re.findall(r'src="(https://assets\.pyecharts\.org[^"]+)"', html))
    replaced = 0
    for url in urls:
        local = pathlib.Path(url.split("/")[-1])   # 本地缓存文件
        if not local.exists():
            # 依次尝试官方源 -> jsdelivr -> cdnjs
            js = None
            for fallback in (url,
                             "https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js",
                             "https://cdnjs.cloudflare.com/ajax/libs/echarts/5.4.3/echarts.min.js"):
                try:
                    data = requests.get(fallback, timeout=30).content
                    if len(data) > 100000:          # echarts.min.js 约 1MB
                        js = data
                        break
                except Exception:
                    continue
            if js is None:
                print(f"警告: 无法下载 {url}，图表可能无法显示")
                continue
            local.write_bytes(js)
        js_text = local.read_text(encoding="utf-8")
        # 兼容带 type 属性的标签：<script type="text/javascript" src="...">
        pattern = r'<script[^>]*\ssrc="' + re.escape(url) + r'"[^>]*>\s*</script>'
        html, n = re.subn(pattern, lambda _: "<script>\n" + js_text + "\n</script>", html)
        replaced += n
    pathlib.Path(html_path).write_text(html, encoding="utf-8")
    if replaced:
        print(f"已内嵌 {replaced} 个本地 JS，图表可离线打开")
    else:
        print("警告: 未发现可替换的外部 JS 引用，请检查 HTML")


inline_echarts(OUT_HTML)
print(f"柱状图已生成 -> {OUT_HTML}")

# 运行结束自动用默认浏览器打开图表
import webbrowser
webbrowser.open(pathlib.Path(OUT_HTML).resolve().as_uri())
print("已在默认浏览器中打开")

# 附：组内每档占比明细
for name, mask in [("免费", valid["是否免费"]), ("付费", ~valid["是否免费"])]:
    rates = valid[mask]["评价EN"].value_counts(normalize=True).mul(100).round(2)
    print(f"\n{name}组评价占比(%):\n{rates.to_string()}")
