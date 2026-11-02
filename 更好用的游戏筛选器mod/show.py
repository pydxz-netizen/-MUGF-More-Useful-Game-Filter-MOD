# -*- coding: utf-8 -*-
"""
性价比四象限气泡图
==================
x = 现价，y = 好评率%，气泡大小 = 评测数
四个象限：秋促捡漏区 / 大作区 / 低价观望区 / 高价避雷区
悬浮显示游戏名与数据，点击气泡跳转 Steam 商店页
数据源：steam_totdata.xlsx（只取付费且有评价摘要的游戏）
"""

import json
import math
import pathlib
import re

import pandas as pd
import requests
import webbrowser

from pyecharts import options as opts
from pyecharts.charts import Scatter
from pyecharts.commons.utils import JsCode
from pyecharts.globals import ThemeType
from pyecharts.options import MarkAreaItem, MarkLineItem, ScatterItem

OUT_HTML = "性价比四象限气泡图.html"

# ========== 1. 数据准备 ==========
df = pd.read_excel("steam_totdata.xlsx")
# 只取付费且好评率有效的游戏：免费游戏没有"价格性价比"，无摘要的没有纵坐标
d = df[(df["现价"] > 0) & df["好评率%"].notna()].copy()

x_med = float(d["现价"].median())          # 低价/高价分界：现价中位数
y_med = float(d["好评率%"].median())       # 好评/差评分界：好评率中位数
x_max = math.ceil(d["现价"].max() / 10) * 10
count_max = int(d["评测数"].max())

print(f"参与绘制游戏: {len(d)} 款 | 现价中位数 ¥{x_med:.0f} | 好评率中位数 {y_med:.0f}%")

# ========== 2. 生成气泡数据（value = [现价, 好评率, 评测数, 商店链接]） ==========
def bubble_size(count: int) -> float:
    """sqrt 缩放，防止大评测数游戏把其他气泡挤成针尖"""
    return round(8 + 32 * math.sqrt(max(count, 1) / count_max), 1)

rows, points = [], []
for _, row in d.iterrows():
    name = row["名称"]
    x = round(float(row["现价"]), 2)
    y = round(float(row["好评率%"]), 1)
    cnt = int(row["评测数"])
    size = bubble_size(cnt)
    rows.append([name, x, y, cnt, row["商店链接"], size])
    points.append(ScatterItem(
        name=name,
        value=[x, y, cnt, row["商店链接"]],
        symbol_size=size,
    ))

# 全量数据注入 JS（[名称, 现价, 好评率, 评测数, 链接, 气泡大小]）：
# 搜索时用 setOption 把 series 数据切换成"仅命中项"，清空则恢复全量
rows_json = json.dumps(rows, ensure_ascii=False)

# ========== 3. 四象限着色 + 分界线 ==========
def area(name, x0, x1, y0, y1, fill, font_color, pos):
    return MarkAreaItem(
        name=name, x=(x0, x1), y=(y0, y1),
        itemstyle_opts=opts.ItemStyleOpts(color=fill),
        label_opts=opts.LabelOpts(
            is_show=True, position=pos, font_size=15, font_weight="bold",
            color=font_color,
        ),
    )

chart = (
    Scatter(
        init_opts=opts.InitOpts(
            width="96%", height="760px",
            page_title="Steam 游戏性价比四象限",
            theme=ThemeType.LIGHT,
        )
    )
    .add_xaxis([0])   # x 轴仅用于展示，数据点自带横坐标
    .add_yaxis(
        "游戏",
        points,
        symbol="circle",
        itemstyle_opts=opts.ItemStyleOpts(color="#5B8FF9", opacity=0.55),
    )
    .set_series_opts(
        # 象限底色与名称
        markarea_opts=opts.MarkAreaOpts(
            is_silent=True,
            data=[
                area("秋促捡漏区", 0, x_med, y_med, 100, "rgba(34,192,115,0.10)", "#17834d", "insideTopLeft"),
                area("大作区", x_med, x_max, y_med, 100, "rgba(74,144,217,0.10)", "#2b6cb0", "insideTopRight"),
                area("低价观望区", 0, x_med, 0, y_med, "rgba(150,150,150,0.07)", "#8a8a8a", "insideBottomLeft"),
                area("高价避雷区", x_med, x_max, 0, y_med, "rgba(224,82,82,0.10)", "#b13a3a", "insideBottomRight"),
            ],
        ),
        # 两条中位数分界线
        markline_opts=opts.MarkLineOpts(
            is_silent=True, symbol=["none", "none"],
            linestyle_opts=opts.LineStyleOpts(color="#999", type_="dashed"),
            label_opts=opts.LabelOpts(font_size=11, color="#666"),
            data=[
                MarkLineItem(x=x_med, name=f"现价中位数 ¥{x_med:.0f}"),
                MarkLineItem(y=y_med, name=f"好评率中位数 {y_med:.0f}%"),
            ],
        ),
    )
    .set_global_opts(
        title_opts=opts.TitleOpts(
            title="Steam 游戏性价比四象限图",
            subtitle=(
                f"共 {len(d)} 款付费游戏（免费及无评价摘要者不参与）｜"
                f"分界线 = 全体中位数：现价 ¥{x_med:.0f}、好评率 {y_med:.0f}%｜"
                f"气泡大小 = 评测数，点击气泡可跳转商店页面"
            ),
        ),
        legend_opts=opts.LegendOpts(is_show=False),
        # 可缩放坐标系：横轴/纵轴各一条滑块选择范围和位置，滚轮可在画布内缩放
        datazoom_opts=[
            # 两个 inside 分别绑定 x / y 轴：滚轮仍同时缩放两轴适应窗口；
            # filter_mode="none" 纯视窗缩放不过滤数据点，彻底避免双轴联动
            opts.DataZoomOpts(type_="inside", xaxis_index=[0], filter_mode="none",
                              range_start=0, range_end=100),
            opts.DataZoomOpts(type_="inside", yaxis_index=[0], filter_mode="none",
                              range_start=0, range_end=100),
            opts.DataZoomOpts(type_="slider", xaxis_index=[0], filter_mode="none",
                              pos_bottom="8px", height=20, range_start=0, range_end=100),
            opts.DataZoomOpts(type_="slider", yaxis_index=[0], orient="vertical",
                              filter_mode="none",
                              pos_right="4px", width=20, range_start=0, range_end=100),
        ],
        xaxis_opts=opts.AxisOpts(
            type_="value",   # 必须显式数值轴，否则类目轴会让气泡横坐标错位
            name="现价 (元)", min_=0, max_=x_max,
            axislabel_opts=opts.LabelOpts(formatter="{value}"),
            splitline_opts=opts.SplitLineOpts(is_show=True,
                linestyle_opts=opts.LineStyleOpts(opacity=0.3)),
        ),
        yaxis_opts=opts.AxisOpts(
            name="好评率 (%)", min_=0, max_=100,
            splitline_opts=opts.SplitLineOpts(is_show=True,
                linestyle_opts=opts.LineStyleOpts(opacity=0.3)),
        ),
        tooltip_opts=opts.TooltipOpts(
            formatter=JsCode(
                """function (p) {
                    if (!p.value) return p.name;
                    return '<b>' + p.name + '</b><br/>现价: ¥' + p.value[0] +
                           '<br/>好评率: ' + p.value[1] + '%<br/>评测数: ' +
                           p.value[2] + '<br/>（点击气泡跳转商店页面）';
                }"""
            )
        ),
    )
    # 点击气泡跳转商店链接 + 顶部搜索框（子串过滤：只显示命中气泡）
    .add_js_funcs(
        """
        (function () {
            var dom = document.querySelector('[_echarts_instance_]');
            var chart = echarts.getInstanceByDom(dom);

            // 点击气泡 -> 新标签页打开商店链接
            chart.on('click', function (p) {
                if (p.value && p.value[3]) window.open(p.value[3], '_blank');
            });

            // ===== 搜索框：子串匹配，只显示命中气泡，其余隐藏 =====
            var wrap = document.createElement('div');
            wrap.style.cssText = 'width:96%;margin:6px auto 0;display:flex;gap:10px;align-items:center;';
            wrap.innerHTML =
                '<input id="game-search" type="text" placeholder="输入游戏名关键词（只显示匹配的游戏，如: stardew / 怪物猎人）" ' +
                'style="flex:1;height:34px;padding:0 12px;font-size:14px;border:1px solid #c9c9c9;border-radius:6px;outline:none;"/>' +
                '<span id="search-count" style="font-size:13px;color:#666;white-space:nowrap;"></span>';
            dom.parentNode.insertBefore(wrap, dom);

            var input = document.getElementById('game-search');
            var counter = document.getElementById('search-count');
            // 全量数据由 Python 注入：[名称, 现价, 好评率, 评测数, 链接, 气泡大小]
            var FULL = __ROWS_JSON__.map(function (r) {
                return {name: r[0], value: [r[1], r[2], r[3], r[4]], symbolSize: r[5]};
            });

            function doSearch() {
                var q = input.value.trim().toLowerCase();
                if (!q) {
                    chart.setOption({series: [{data: FULL}]});   // 恢复全部
                    counter.textContent = '';
                    return;
                }
                var hits = [];
                for (var i = 0; i < FULL.length; i++) {
                    if (String(FULL[i].name).toLowerCase().indexOf(q) >= 0) hits.push(FULL[i]);
                }
                // 只保留命中的气泡，其余隐藏
                chart.setOption({series: [{data: hits}]});
                counter.textContent = hits.length
                    ? '显示 ' + hits.length + ' 款'
                    : '无匹配游戏';
            }

            // 输入防抖后实时搜索
            var timer = null;
            input.addEventListener('input', function () {
                clearTimeout(timer);
                timer = setTimeout(doSearch, 300);
            });
        })();
        """.replace("__ROWS_JSON__", rows_json)
    )
)

chart.render(OUT_HTML)


# ========== 4. 内嵌 echarts.min.js，离线可打开 ==========
def inline_echarts(html_path: str):
    html = pathlib.Path(html_path).read_text(encoding="utf-8")
    urls = set(re.findall(r'src="(https://assets\.pyecharts\.org[^"]+)"', html))
    replaced = 0
    for url in urls:
        local = pathlib.Path(url.split("/")[-1])   # 本地缓存 echarts.min.js
        if not local.exists():
            js = None
            for fallback in (url,
                             "https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js",
                             "https://cdnjs.cloudflare.com/ajax/libs/echarts/5.4.3/echarts.min.js"):
                try:
                    data = requests.get(fallback, timeout=30).content
                    if len(data) > 100000:
                        js = data
                        break
                except Exception:
                    continue
            if js is None:
                print(f"警告: 无法下载 {url}")
                continue
            local.write_bytes(js)
        js_text = local.read_text(encoding="utf-8")
        pattern = r'<script[^>]*\ssrc="' + re.escape(url) + r'"[^>]*>\s*</script>'
        html, n = re.subn(pattern, lambda _: "<script>\n" + js_text + "\n</script>", html)
        replaced += n
    pathlib.Path(html_path).write_text(html, encoding="utf-8")
    if replaced:
        print(f"已内嵌 {replaced} 个本地 JS，可离线打开")
    else:
        print("警告: 未发现可替换的外部 JS 引用")


inline_echarts(OUT_HTML)
print(f"图表已生成 -> {OUT_HTML}")

# ========== 5. 运行完直接打开 ==========
webbrowser.open(pathlib.Path(OUT_HTML).resolve().as_uri())
print("已在默认浏览器中打开")

# 附：各象限代表游戏（写报告时可直接引用）
for name, mask in [
    ("秋促捡漏区(低价+好评)", (d["现价"] <= x_med) & (d["好评率%"] >= y_med)),
    ("高价避雷区(高价+差评)", (d["现价"] > x_med) & (d["好评率%"] < y_med)),
]:
    sub = d[mask].nlargest(5, "评测数")
    print(f"\n{name} 代表游戏（按评测数）:")
    for _, r in sub.iterrows():
        print(f"  {r['名称']} | ¥{r['现价']:.0f} | {r['好评率%']:.0f}% | {r['评测数']:,} 评测")
