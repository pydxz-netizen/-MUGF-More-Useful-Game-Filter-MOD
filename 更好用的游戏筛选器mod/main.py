# -*- coding: utf-8 -*-
"""
性价比四象限气泡图（含类型/评测数/发布日期筛选）
================================================
x = 现价，y = 好评率%；统一小气泡，颜色 = 评测数（右侧分段图例）
筛选器（合取）：名称搜索 ∧ 类型tag(析取) ∧ 评测数区间 ∧ 发布日期区间
其余：滚轮缩放 + x/y 独立滑块、点击气泡加入下方对比表格（表格+动态对比条形图）、HTML 内嵌 echarts 离线可用
数据源：steam_totdata.xlsx（有评价摘要即绘制，含免费游戏，免费现价记 0）
"""

import json
import math
import pathlib
import re
from collections import Counter
from datetime import datetime

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
# 只要求有评价摘要（好评率%）；免费游戏保留——否则"免费开玩"等 tag
# 会筛掉 88 款中的 81 款（免费游戏现价记 0 元，位于图表最左侧）
df = pd.read_excel("steam_totdata.xlsx")
# 现价为空（"暂不可购"等）无法落在价格轴上，直接排除，避免出现坐标 NaN / 表格里显示 "¥NaN"
d = df[df["好评率%"].notna() & df["现价"].notna()].copy()

x_med = float(d["现价"].median())          # 低价/高价分界：现价中位数
y_med = float(d["好评率%"].median())       # 好评/差评分界：好评率中位数
x_max = math.ceil(d["现价"].max() / 10) * 10

print(f"参与绘制游戏: {len(d)} 款 | 现价中位数 ¥{x_med:.0f} | 好评率中位数 {y_med:.0f}%")


# ========== 2. 字段解析与数据注入 ==========
def parse_release_ms(s):
    """'2025 年 12 月 10 日' -> 毫秒时间戳；无法解析返回 None"""
    if pd.isna(s):
        return None
    m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月(?:\s*(\d{1,2})\s*日)?", str(s))
    if not m:
        return None
    try:
        day = int(m.group(3)) if m.group(3) else 1
        return datetime(int(m.group(1)), int(m.group(2)), day).timestamp() * 1000
    except ValueError:
        return None


def split_tags(s):
    if pd.isna(s):
        return []
    return [t.strip() for t in str(s).split("、") if t.strip()]


# 类型 tag 并集（按出现次数降序），供下拉框展示
tag_freq = Counter(t for tags in d["类型"].map(split_tags) for t in tags)
tags_json = json.dumps(
    [{"name": t, "count": c} for t, c in tag_freq.most_common()],
    ensure_ascii=False,
)
print(f"类型 tag 并集: {len(tag_freq)} 个")

def fmt_date_text(s):
    """发售日期 -> 'YYYY-MM-DD'（无法解析时返回原文，空值返回空串）"""
    ms = parse_release_ms(s)
    if ms is not None:
        return datetime.fromtimestamp(ms / 1000).strftime("%Y-%m-%d")
    return "" if pd.isna(s) else str(s).strip()


# rows = [名称, 现价, 好评率, 评测数, 链接, 发售时间戳ms, 类型tag列表, 原价, 发布日期文本]
rows, points = [], []
for _, row in d.iterrows():
    name = row["名称"]
    x = round(float(row["现价"]), 2)
    y = round(float(row["好评率%"]), 1)
    cnt = int(row["评测数"])
    orig = row["原价"]
    orig_v = round(float(orig), 2) if pd.notna(orig) else None
    rows.append([name, x, y, cnt, row["商店链接"],
                 parse_release_ms(row["发售日期"]), split_tags(row["类型"]),
                 orig_v, fmt_date_text(row["发售日期"])])
    points.append(ScatterItem(
        name=name,
        value=[x, y, cnt, row["商店链接"]],
    ))

# 全量数据注入 JS：筛选时用 setOption 把 series 数据切换成"仅命中项"，清空恢复全量
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
        symbol_size=10,   # 统一小点，大小不再编码评测数（由颜色表达）
        label_opts=opts.LabelOpts(is_show=False),   # 1047 个点不显示文字标签，避免互相压字
        itemstyle_opts=opts.ItemStyleOpts(opacity=0.8),
    )
    .set_series_opts(
        # 象限底色与名称
        markarea_opts=opts.MarkAreaOpts(
            is_silent=True,
            data=[
                area("性价比区", -2, x_med, y_med, 101, "rgba(34,192,115,0.10)", "#17834d", "insideTopLeft"),
                area("大作区", x_med, x_max, y_med, 101, "rgba(74,144,217,0.10)", "#2b6cb0", "insideTopRight"),
                area("低价观望区", -2, x_med, 0, y_med, "rgba(150,150,150,0.07)", "#8a8a8a", "insideBottomLeft"),
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
                f"分界线 = 全体中位数：现价 ¥{x_med:.0f}、好评率 {y_med:.0f}%｜"
                f"颜色 = 评测数（右侧图例），点击气泡加入下方对比表格"
            ),
            title_textstyle_opts=opts.TextStyleOpts(
                font_size=25, font_weight="bold", color="#1f2937"),
            subtitle_textstyle_opts=opts.TextStyleOpts(
                font_size=13, color="#6b7280"),
        ),
        legend_opts=opts.LegendOpts(is_show=False),
        # 评测数 -> 颜色 分段图例（图表右侧，兼作点击筛选开关）
        visualmap_opts=opts.VisualMapOpts(
            type_="color",
            is_piecewise=True,
            dimension=2,   # value[2] = 评测数
            pieces=[
                {"min": 0, "max": 32, "label": "0~32 条评测", "color": "#A50F15"},
                {"min": 33, "max": 256, "label": "33~256", "color": "#F5C242"},
                {"min": 257, "max": 1024, "label": "257~1024", "color": "#2E8B57"},
                {"min": 1025, "label": "1024+ 条评测", "color": "#7D3AC1"},
            ],
            pos_right="6px", pos_top="middle", orient="vertical",
        ),
        # 可缩放坐标系：横轴/纵轴各一条滑块，滚轮画布内缩放；
        # filter_mode="none" 纯视窗缩放不过滤数据，双轴互不联动
        datazoom_opts=[
            opts.DataZoomOpts(type_="inside", xaxis_index=[0], filter_mode="none",
                              range_start=0, range_end=100),
            opts.DataZoomOpts(type_="inside", yaxis_index=[0], filter_mode="none",
                              range_start=0, range_end=100),
            opts.DataZoomOpts(type_="slider", xaxis_index=[0], filter_mode="none",
                              pos_bottom="8px", height=20, range_start=0, range_end=100),
            opts.DataZoomOpts(type_="slider", yaxis_index=[0], orient="vertical",
                              filter_mode="none",
                              pos_left="6px", width=20, range_start=0, range_end=100),
        ],
        xaxis_opts=opts.AxisOpts(
            type_="value",   # 必须显式数值轴，否则类目轴会让气泡横坐标错位
            name="现价 (元)", min_=-2, max_=x_max,   # 最低 -2：防止 x=0 的免费游戏贴边
            axislabel_opts=opts.LabelOpts(
                # 只隐藏 -2 这个留白端点，其余价格刻度（0~690）都正常显示
                formatter=JsCode(
                    """function (v) {
                        if (parseFloat(v) < 0) return '';
                        return v;
                    }"""
                )
            ),
            splitline_opts=opts.SplitLineOpts(is_show=True,
                linestyle_opts=opts.LineStyleOpts(opacity=0.3)),
        ),
        yaxis_opts=opts.AxisOpts(
            name="好评率 (%)", min_=0, max_=101,   # 上限 101：防止 100% 好评贴边
            axislabel_opts=opts.LabelOpts(
                formatter=JsCode(
                    """function (v) {
                        if (parseFloat(v) >= 101) return '';
                        return v;
                    }"""
                )
            ),
            splitline_opts=opts.SplitLineOpts(is_show=True,
                linestyle_opts=opts.LineStyleOpts(opacity=0.3)),
        ),
        tooltip_opts=opts.TooltipOpts(
            # 浮出框改到光标左下方
            position=JsCode(
                """function (point, params, dom, rect, size) {
                    var x = point[0] - size.contentSize[0] - 12;
                    if (x < 0) x = 0;
                    return [x, point[1] + 18];
                }"""
            ),
            formatter=JsCode(
                """function (p) {
                    if (!p.value) return p.name;
                    return '<b>' + p.name + '</b><br/>现价: ¥' + p.value[0] +
                           '<br/>好评率: ' + p.value[1] + '%<br/>评测数: ' +
                           p.value[2] + '<br/>（点击气泡加入下方对比表格）';
                }"""
            )
        ),
    )
    # 点击气泡加入对比表格 + 顶部搜索框 + 筛选栏（类型/评测数/发布日期）
    .add_js_funcs(
        """
        (function () {
            var dom = document.querySelector('[_echarts_instance_]');
            var chart = echarts.getInstanceByDom(dom);

            // 图表绘图区整体下移，给标题/副标题留出空间，避免遮挡轴名称
            // 注意：pyecharts 会在这段自定义 JS 之后才执行初次 setOption，
            // 因此必须延后到下一个事件循环再改样式，否则会被初始 option 覆盖
            setTimeout(function () {
            chart.setOption({
                grid: {top: 118},
                textStyle: {fontFamily: '"Segoe UI","Microsoft YaHei","PingFang SC",sans-serif'},
                legend: {textStyle: {color: '#475569'}},
                // 坐标轴刻度：自动隐藏重叠刻度，字号统一，避免数字挤成一团
                xAxis: {splitNumber: 8,
                        axisLabel: {hideOverlap: true, fontSize: 11, color: '#64748b'},
                        nameTextStyle: {fontSize: 13, color: '#475569', fontWeight: 'bold'}},
                yAxis: {splitNumber: 8,
                        axisLabel: {hideOverlap: true, fontSize: 11, color: '#64748b'},
                        nameTextStyle: {fontSize: 13, color: '#475569', fontWeight: 'bold'}},
                // 滑块内不再画密集的数据迷你条（上万点时像一片噪点）
                dataZoom: [{}, {}, {showDataShadow: false}, {showDataShadow: false}]
            });
            // 页面内容（搜索/筛选/对比区）插入后页面宽度可能变化，
            // 重算一次画布尺寸，避免画布比白色卡片窄一圈导致边框看起来没对齐
            chart.resize();
            }, 0);

            // 点击气泡 -> 加入下方对比表格（不再直接跳转商店）
            chart.on('click', function (p) {
                if (p.value && p.value[3]) addToCompare(p.name);
            });

            // ===== 图例筛选与数据同步：取消某段评测数后，真正移除该段点，避免幽灵点 =====
            // 要点：自己维护"用户勾选状态"作为权威值，不直接用 echarts 的 selected
            //（setOption 更新数据后 echarts 可能把 selected 重置为全选，导致筛掉的点 0.3s 后复活）
            var PIECES = [{min: 0, max: 32}, {min: 33, max: 256},
                          {min: 257, max: 1024}, {min: 1025}];
            var visualFlags = [true, true, true, true];   // 用户勾选状态（权威）
            var visualRanges = null;                      // null = 全部显示
            var lastApplyAt = 0;

            // selected 可能是数组 [false,true,...]，也可能是对象 {0:false}，统一成数组
            function normSelected(sel) {
                if (!sel || typeof sel === 'string') return null;   // 'all' / 空 => 全选
                var arr = [], i, v;
                if (Object.prototype.toString.call(sel) === '[object Array]') {
                    for (i = 0; i < PIECES.length; i++) arr.push(sel.length > i ? sel[i] !== false : true);
                    return arr;
                }
                for (i = 0; i < PIECES.length; i++) {
                    v = sel[i];
                    if (v === undefined) v = sel[String(i)];
                    arr.push(v !== false);   // 未记录视为选中
                }
                return arr;
            }
            function rangesFrom(flags) {
                var out = [], anyOff = false;
                for (var i = 0; i < PIECES.length; i++) {
                    if (!flags[i]) { anyOff = true; continue; }
                    out.push([PIECES[i].min == null ? -Infinity : PIECES[i].min,
                              PIECES[i].max == null ? Infinity : PIECES[i].max]);
                }
                return anyOff ? out : null;
            }
            function applyVisualFlags(flags) {
                if (!flags) return;
                var same = true;
                for (var i = 0; i < PIECES.length; i++) {
                    if (flags[i] !== visualFlags[i]) { same = false; break; }
                }
                if (same) return;
                visualFlags = flags;
                visualRanges = rangesFrom(flags);
                lastApplyAt = Date.now();
                applyAll();
            }
            function pushSelected() {   // 把勾选状态写回 echarts，保证图例显示与实际筛选一致
                try {
                    chart.dispatchAction({type: 'selectDataRange', visualMapIndex: 0,
                                          selected: visualFlags.slice()});
                } catch (e) {}
            }

            chart.on('datarangeselected', function (e) {   // visualMap 选区变化的官方事件
                applyVisualFlags(normSelected(e && e.selected));
            });
            chart.on('selectchanged', function () {         // 兼容：echarts5 通用选择事件
                var vm = (chart.getOption() || {}).visualMap;
                if (vm && vm[0]) applyVisualFlags(normSelected(vm[0].selected));
            });
            setInterval(function () {                       // 兜底：真实点击偶尔不发事件
                if (Date.now() - lastApplyAt < 600) return; // 刚改过数据，跳过（防重置误判）
                var vm = (chart.getOption() || {}).visualMap;
                if (!vm || !vm[0]) return;
                applyVisualFlags(normSelected(vm[0].selected));
            }, 400);

            // 全量数据：[名称, 现价, 好评率, 评测数, 链接, 发售时间戳ms, tags数组, 原价, 发布日期文本]
            var FULL = __ROWS_JSON__.map(function (r) {
                return {name: r[0], value: [r[1], r[2], r[3], r[4]], date: r[5], tags: r[6],
                        orig: r[7], dateText: r[8]};
            });
            var TAGS = __TAGS_JSON__;   // [{name, count}] 按次数降序
            var X_MED = __X_MED__, Y_MED = __Y_MED__;   // 未筛选全量数据的现价/好评率中位数
            var lastHits = FULL.slice();                // 当前筛选命中的游戏（批量导入用）
            window.__MED = {x: X_MED, y: Y_MED};        // 测试/调试钩子

            // ================= UI 构建 =================
            var css = document.createElement('style');
            css.textContent = [
                // ---- 全局版式 ----
                'body{margin:0;padding:18px 0 60px;background:linear-gradient(180deg,#f6f8fc 0%,#eef2f8 55%,#e8eefa 100%);',
                'font-family:"Segoe UI","Microsoft YaHei","PingFang SC",system-ui,sans-serif;color:#1f2937;}',
                '[_echarts_instance_]{margin:0 auto;background:#fff;border-radius:16px;box-shadow:0 10px 30px rgba(15,23,42,.08);}',
                // ---- 搜索栏 ----
                '.search-bar{display:flex;align-items:center;gap:10px;box-sizing:border-box;background:#fff;border:1px solid #e4e9f2;',
                'border-radius:14px;padding:7px 12px;box-shadow:0 6px 18px rgba(15,23,42,.06);transition:.2s;}',
                '.search-bar:focus-within{border-color:#9dbcf5;box-shadow:0 8px 22px rgba(59,130,246,.15);}',
                '.sb-input{flex:1;height:34px;border:none;outline:none;font-size:14px;background:transparent;',
                'padding:0 4px;color:#1f2937;font-family:inherit;}',
                '.sb-input::placeholder{color:#9aa4b2;}',
                '.sb-count{font-size:12px;color:#2563eb;background:#eef4ff;border:1px solid #dbe7ff;',
                'border-radius:999px;padding:2px 10px;white-space:nowrap;}',
                // ---- 筛选卡片 ----
                '.flt-card{background:#fff;border:1px solid #e4e9f2;border-radius:14px;padding:10px 12px;',
                'box-shadow:0 4px 14px rgba(15,23,42,.045);}',
                '.flt-label{font-size:12px;color:#64748b;margin-bottom:6px;font-weight:600;letter-spacing:.2px;}',
                '.flt-input{height:32px;padding:0 9px;font-size:13px;border:1px solid #d9e0ea;',
                'border-radius:9px;outline:none;box-sizing:border-box;width:100%;background:#fff;',
                'color:#1f2937;font-family:inherit;transition:.15s;}',
                '.flt-input:focus{border-color:#7aa7f2;box-shadow:0 0 0 3px rgba(59,130,246,.12);}',
                '.flt-sep{color:#b6bfcc;font-size:13px;}',
                // ---- 类型多选下拉 ----
                '#tag-panel{position:absolute;top:100%;left:0;right:0;max-height:250px;overflow-y:auto;',
                'background:#fff;border:1px solid #e4e9f2;border-radius:12px;z-index:999;display:none;',
                'box-shadow:0 14px 34px rgba(15,23,42,.16);padding:4px;}',
                '.tag-opt{padding:7px 10px;font-size:13px;cursor:pointer;display:flex;justify-content:space-between;',
                'border-radius:8px;transition:.12s;}',
                '.tag-opt:hover{background:#f1f5ff;}',
                '.tag-opt.on{color:#1d4ed8;font-weight:600;background:#eef4ff;}',
                '.tag-opt .cnt{color:#94a3b8;font-size:12px;font-weight:normal;}'
            ].join('');
            document.head.appendChild(css);

            // 搜索框（最上方）
            var searchWrap = document.createElement('div');
            searchWrap.className = 'search-bar';
            searchWrap.style.cssText = 'width:96%;margin:2px auto 0;';
            searchWrap.innerHTML =
                '<input id="game-search" class="sb-input" type="text" ' +
                'placeholder="输入游戏名关键词（只显示匹配的游戏，如: stardew / 怪物猎人）"/>' +
                '<span id="search-count" class="sb-count" style="display:none;"></span>';
            dom.parentNode.insertBefore(searchWrap, dom);

            // 筛选行（位于搜索框下方、图表标题上方）
            var row = document.createElement('div');
            row.style.cssText = 'width:96%;margin:10px auto 0;display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap;';
            row.innerHTML =
                '<div class="flt-card" style="flex:2.2;min-width:280px;position:relative;">' +
                  '<div class="flt-label">类型（析取 = 含任一类型，合取 = 同时含所有类型；Backspace 删除末位）</div>' +
                  '<div style="display:flex;gap:6px;">' +
                    '<select id="tag-mode" class="flt-input" style="flex:none;width:78px;">' +
                      '<option value="or" selected>析取</option>' +
                      '<option value="and">合取</option>' +
                    '</select>' +
                    '<input id="tag-input" class="flt-input" readonly placeholder="点击选择类型标签..." style="flex:1;">' +
                  '</div>' +
                  '<div id="tag-panel"></div>' +
                '</div>' +
                '<div class="flt-card" style="flex:1.1;min-width:190px;">' +
                  '<div class="flt-label">评测数范围（留空不限）</div>' +
                  '<div style="display:flex;gap:6px;align-items:center;">' +
                    '<input id="cnt-min" class="flt-input" type="number" min="0" placeholder="最少(0)">' +
                    '<span class="flt-sep">~</span>' +
                    '<input id="cnt-max" class="flt-input" type="number" min="0" placeholder="最多(∞)">' +
                  '</div>' +
                '</div>' +
                '<div class="flt-card" style="flex:1.3;min-width:220px;">' +
                  '<div class="flt-label">发布日期范围（年 或 年-月，留空不限）</div>' +
                  '<div style="display:flex;gap:6px;align-items:center;">' +
                    '<input id="date-min" class="flt-input" placeholder="最早 如 2020 或 2020-3">' +
                    '<span class="flt-sep">~</span>' +
                    '<input id="date-max" class="flt-input" placeholder="最晚 如 2025 或 2025-12">' +
                  '</div>' +
                '</div>';
            dom.parentNode.insertBefore(row, dom);

            // ================= 对比表格 + 动态对比条形图 =================
            var cmpCss = document.createElement('style');
            cmpCss.textContent = [
                '.cmp-card{background:#fff;border:1px solid #e4e9f2;border-radius:14px;',
                'padding:6px 12px;box-shadow:0 6px 18px rgba(15,23,42,.05);overflow:hidden;}',
                '.cmp-table{border-collapse:separate;border-spacing:0;width:100%;font-size:13px;}',
                '.cmp-table th{background:#f7f9fc;color:#475569;font-size:12px;font-weight:700;letter-spacing:.3px;',
                'padding:9px 10px;text-align:left;border-bottom:1px solid #e8edf5;white-space:nowrap;}',
                '.cmp-table td{padding:8px 10px;text-align:left;border-bottom:1px solid #f0f3f8;color:#334155;}',
                '.cmp-table tbody tr:hover td{background:#f8fbff;}',
                '.cmp-table a{color:#1d4ed8;text-decoration:none;font-weight:600;}',
                '.cmp-table a:hover{text-decoration:underline;}',
                '.cmp-del{width:22px;height:22px;border:1px solid #f1c2c2;background:#fff;color:#dc2626;',
                'border-radius:50%;cursor:pointer;line-height:1;font-size:14px;padding:0;transition:.15s;}',
                '.cmp-del:hover{background:#dc2626;border-color:#dc2626;color:#fff;}',
                '.cmp-btn{height:28px;padding:0 12px;border:1px solid #dbe3ef;background:#fff;color:#475569;',
                'border-radius:999px;cursor:pointer;font-size:12px;vertical-align:middle;transition:.15s;',
                'font-family:inherit;}',
                '.cmp-btn:hover{background:#f1f6ff;border-color:#bcd3f7;color:#1d4ed8;}',
                '.cmp-btn.on{background:#2563eb;border-color:#2563eb;color:#fff;font-weight:600;}',
                '.cmp-empty{color:#94a3b8;font-size:13px;padding:14px 4px;}',
                '.cmp-title{font-size:16px;font-weight:700;color:#1f2937;margin:16px 0 8px;',
                'display:flex;align-items:center;gap:10px;flex-wrap:wrap;}',
                '.cmp-title::before{content:"";width:4px;height:16px;border-radius:2px;background:#2563eb;}',
                '.cmp-sub{font-size:12px;color:#8b96a8;font-weight:400;}',
                '.cmp-title .cmp-btn{margin-left:0;}',
                '.cmp-import{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:16px 0 0;}',
                '#imp-tip{font-size:12px;color:#0f766e;}',
                '#cmp-note{font-size:12px;color:#b45309;font-weight:400;}',
                '.cmp-legend{display:flex;align-items:center;gap:16px;flex-wrap:wrap;',
                'font-size:12px;font-weight:700;margin:0 0 8px 4px;}',
                '.cmp-scale{margin:4px 6px 0;}',
                '.cmp-scale-track{padding:0 8px;}',
                '#cmp-scale-input{width:100%;height:22px;margin:0;background:transparent;',
                '-webkit-appearance:none;appearance:none;cursor:pointer;}',
                '#cmp-scale-input::-webkit-slider-runnable-track{height:6px;border-radius:3px;',
                'background:#e6ebf3;}',
                '#cmp-scale-input::-webkit-slider-thumb{-webkit-appearance:none;width:14px;height:14px;',
                'border-radius:50%;background:#2563eb;border:2px solid #fff;margin-top:-4px;',
                'box-shadow:0 1px 3px rgba(0,0,0,.25);}',
                '.cmp-scale-ticks{display:flex;justify-content:space-between;padding:0 8px;',
                'margin-top:1px;}',
                '.cmp-scale-ticks i{width:1px;height:6px;background:#94a3b8;}',
                '.cmp-scale-tip{font-size:12px;color:#64748b;margin-top:5px;text-align:right;}',
                '#cmp-chart{box-shadow:none!important;border-radius:12px!important;}'
            ].join('');
            document.head.appendChild(cmpCss);

            var cmp = [];          // 已加入对比的游戏（保持加入顺序）
            var sortKey = 'cnt';   // 对比图 y 轴排序依据：cnt=评测数 / rating=好评率 / price=现价
            var axisSkip = 0;      // 拖动条：已从右侧裁掉的"评测数最多"的游戏数量（0 = 全部完整显示）

            // 对比图配色：价格档决定色相，深色段 = 好评率占比，浅色段 = 非好评
            var PRICE_BANDS = [
                {max: 0,        name: '绿色：免费',           dark: '#1E7B44', light: '#A8DDB7'},
                {max: 40,       name: '黄色：40 元及以下',     dark: '#C79A16', light: '#F7E3A1'},
                {max: 120,      name: '红色：120 元及以下',    dark: '#C0392B', light: '#F3B7AE'},
                {max: Infinity, name: '紫色：超过 120 元',     dark: '#7D3AC1', light: '#D6C2EC'}
            ];
            function priceBand(p) {
                if (!p || p <= 0) return PRICE_BANDS[0];          // 免费
                for (var i = 1; i < PRICE_BANDS.length; i++) {
                    if (p <= PRICE_BANDS[i].max) return PRICE_BANDS[i];
                }
                return PRICE_BANDS[PRICE_BANDS.length - 1];
            }
            function priceLegendHtml() {
                return PRICE_BANDS.map(function (b) {
                    return '<span style="color:' + b.dark + '">' + b.name + '</span>';
                }).join('') +
                '<span class="cmp-sub">（柱长 = 评测数，深色段占该柱长度的比例 = 好评率，浅色段 = 非好评）</span>';
            }

            function sortedCmp() {   // 表格与图表共用同一排序
                var arr = cmp.slice();
                arr.sort(function (a, b) {
                    if (sortKey === 'cnt') return b.value[2] - a.value[2];       // 评测数降序
                    if (sortKey === 'rating') return b.value[1] - a.value[1];    // 好评率降序
                    return a.value[0] - b.value[0];                              // 现价升序（便宜的在上）
                });
                return arr;
            }

            var cmpWrap = document.createElement('div');
            cmpWrap.style.cssText = 'width:96%;margin:14px auto 0;';
            cmpWrap.innerHTML =
                '<div class="cmp-import">' +
                  '<button id="imp-all" class="cmp-btn">全部导入</button>' +
                  '<button id="imp-green" class="cmp-btn">绿色区域全部导入</button>' +
                  '<button id="imp-high" class="cmp-btn">高好评全部导入</button>' +
                  '<button id="imp-low" class="cmp-btn">低价格游戏全部导入</button>' +
                  '<button id="imp-clear" class="cmp-btn">清空全部</button>' +
                  '<span id="imp-tip"></span>' +
                '</div>' +
                '<div class="cmp-title">游戏对比表格' +
                '<span class="cmp-sub">（点击图上气泡加入；点击游戏名跳转商店；点击行首 − 删除该行）</span></div>' +
                '<div id="cmp-table-box" class="cmp-card"></div>' +
                '<div class="cmp-title" style="margin-top:18px">游戏对比图' +
                '<span class="cmp-sub">（柱长 = 评测数，深色段 = 好评率，柱色 = 价格档）</span>' +
                '<span id="cmp-note"></span>' +
                '<button id="sort-cnt" class="cmp-btn on">按评测数</button>' +
                '<button id="sort-rating" class="cmp-btn">按好评率</button>' +
                '<button id="sort-price" class="cmp-btn">按现价</button>' +
                '</div>' +
                '<div id="cmp-price-legend" class="cmp-legend">' + priceLegendHtml() + '</div>' +
                '<div id="cmp-chart" class="cmp-card" style="padding:6px"></div>' +
                '<div id="cmp-scale" class="cmp-scale"></div>';
            dom.parentNode.appendChild(cmpWrap);

            function buildTable() {
                var box = document.getElementById('cmp-table-box');
                if (!cmp.length) {
                    box.innerHTML = '<div class="cmp-empty">暂无游戏：点击上方气泡图中的圆点即可添加。</div>';
                    return;
                }
                var html = '<table class="cmp-table"><thead><tr><th></th><th>游戏名称</th><th>类型</th>' +
                    '<th>原价</th><th>现价</th><th>好评率</th><th>评测数</th><th>发布日期</th></tr></thead><tbody>';
                var arr = sortedCmp();   // 表格与对比图共用同一排序
                arr.forEach(function (g, i) {
                    html += '<tr><td><button class="cmp-del" data-i="' + i + '">−</button></td>' +
                        '<td><a href="' + g.value[3] + '" target="_blank">' + g.name + '</a></td>' +
                        '<td>' + (g.tags.join('、') || '—') + '</td>' +
                        '<td>' + (g.orig == null ? '—' : '¥' + g.orig) + '</td>' +
                        '<td>¥' + g.value[0] + '</td>' +
                        '<td>' + g.value[1] + '%</td>' +
                        '<td>' + g.value[2].toLocaleString() + '</td>' +
                        '<td>' + (g.dateText || '—') + '</td></tr>';
                });
                box.innerHTML = html + '</tbody></table>';
            }

            var cmpChart = null;
            function buildChart() {
                var box = document.getElementById('cmp-chart');
                if (cmpChart) { cmpChart.dispose(); cmpChart = null; }
                if (!cmp.length) {
                    document.getElementById('cmp-note').textContent = '';
                    box.innerHTML = '<div class="cmp-empty">暂无可视化：添加游戏后自动生成。</div>';
                    return;
                }
                box.innerHTML = '';
                // 大批量导入（可达上万款）时，条形图只渲染前 60 款，避免页面卡死；
                // 表格始终展示全部游戏
                var CHART_MAX = 60;
                var sortedAll = sortedCmp();   // 与表格共用同一排序
                var sorted = sortedAll.slice(0, CHART_MAX);
                document.getElementById('cmp-note').textContent =
                    sortedAll.length > CHART_MAX
                        ? '（共 ' + sortedAll.length + ' 款，图表仅显示前 ' + CHART_MAX + ' 款）'
                        : '';
                var byCount = sorted.slice().sort(function (a, b) { return b.value[2] - a.value[2]; });
                if (axisSkip > sorted.length - 1) axisSkip = sorted.length - 1;
                if (axisSkip < 0) axisSkip = 0;
                // 上限取"剩余游戏中评测数最大者"的真实评测数，使其柱子右端正好贴住图表右端
                var axisMax = Math.max(byCount[axisSkip] ? byCount[axisSkip].value[2] : 0, 1);

                box.style.height = (90 + sorted.length * 36) + 'px';
                cmpChart = echarts.init(box);
                cmpChart.setOption({
                    grid: {top: 30, left: 16, right: 60, bottom: 56, containLabel: true},
                    tooltip: {
                        position: function (point, params, dom2, rect, size) {
                            var x = point[0] - size.contentSize[0] - 12;
                            if (x < 0) x = 0;
                            return [x, point[1] + 18];
                        },
                        formatter: function (p) {
                            var g = sorted[p.dataIndex];
                            if (!g) return '';
                            return '<b>' + g.name + '</b><br/>现价: ¥' + g.value[0] +
                                   '<br/>好评率: ' + g.value[1] + '%<br/>评测数: ' + g.value[2];
                        }
                    },
                    xAxis: {type: 'value', name: '评测数（条）', nameLocation: 'middle', nameGap: 28,
                            min: 0, max: axisMax,   // 柱长直接表示评测数（上限随拖动条变化）
                            axisLabel: {   // 评测数跨度大，轴刻度用"万"简写
                                formatter: function (v) {
                                    return v >= 10000
                                        ? (v / 10000).toFixed(v % 10000 === 0 ? 0 : 1) + '万'
                                        : v;
                                }
                            },
                            splitLine: {show: true, lineStyle: {opacity: 0.3}}},
                    yAxis: {type: 'category',
                            data: sorted.map(function (g) { return g.name; }),
                            inverse: true,   // 排序第一位的显示在最上方
                            axisLabel: {width: 180, overflow: 'truncate'}},
                    series: [{
                        type: 'custom',   // custom 系列：一柱两段 + 明确的深浅分界线
                        renderItem: function (params, api) {
                            var idx = api.value(0);
                            var g = sorted[idx];
                            if (!g) return null;
                            var h = 16;                                   // 柱粗统一，不承载信息
                            var cnt = g.value[2];
                            var clipped = cnt > axisMax;                  // 评测数超出上限 -> 柱子被裁
                            var x0 = api.coord([0, idx])[0];
                            var xEnd = api.coord([Math.min(cnt, axisMax), idx])[0];
                            var totalW = Math.max(xEnd - x0, 1);          // 柱长 = 评测数（裁掉后贴右边界）
                            var band = priceBand(g.value[0]);
                            var top = api.coord([0, idx])[1] - h / 2;
                            var children = [
                                {type: 'rect',                            // 浅色段：非好评
                                 shape: {x: x0, y: top, width: totalW, height: h},
                                 style: {fill: band.light}}
                            ];
                            if (!clipped) {       // 只有在图内完整显示评测数的柱子才画好评率深色段
                                var pct = Math.max(0, Math.min(100, g.value[1]));
                                var darkW = totalW * pct / 100;
                                children.push({type: 'rect',                // 深色段：好评
                                               shape: {x: x0, y: top, width: darkW, height: h},
                                               style: {fill: band.dark}});
                                children.push({type: 'line',                // 深浅分界线
                                               shape: {x1: x0 + darkW, y1: top,
                                                       x2: x0 + darkW, y2: top + h},
                                               style: {stroke: '#111827', lineWidth: 1}});
                            }
                            return {type: 'group', children: children};
                        },
                        data: sorted.map(function (g, i) {
                            return {value: [i, g.value[2], g.value[1]],
                                    itemStyle: {color: priceBand(g.value[0]).dark}};
                        }),
                        encode: {x: 1, y: 0}
                    }]
                });
                // 点击柱子跳转商店（补充入口）
                cmpChart.on('click', function (p) {
                    var g = sorted[p.dataIndex];
                    if (g && g.value[3]) window.open(g.value[3], '_blank');
                });
                renderScale(sorted.length, axisSkip, axisMax);
            }

            // 对比图下方拖动条：刻度数（含两端）= 图中游戏数；右端 = 全部完整显示，
            // 每往左拖一格，就裁掉当前"评测数最多"的游戏，坐标轴上限换成次大者的评测数
            function renderScale(n, skip, axisMaxVal) {
                var box = document.getElementById('cmp-scale');
                if (!box) return;
                if (n < 2) { box.innerHTML = ''; return; }   // 只有 1 款时无需调节
                var tipText = '坐标轴上限：<b>' + axisMaxVal.toLocaleString() + '</b> 条' +
                    (skip ? '（已裁掉评测数最多的 ' + skip + ' 款，其柱子超出图表、仅显示浅色）'
                          : '（全部游戏完整显示评测数）');
                var exist = document.getElementById('cmp-scale-input');
                if (exist && parseInt(exist.getAttribute('max'), 10) === n - 1) {
                    // 刻度总数没变（只是重新渲染/换排序）：只更新位置与提示，避免打断拖动
                    exist.value = String(n - 1 - skip);
                    box.querySelector('.cmp-scale-tip').innerHTML = tipText;
                    return;
                }
                var ticks = '';
                for (var i = 0; i < n; i++) ticks += '<i></i>';
                box.innerHTML =
                    '<div class="cmp-scale-track">' +
                    '<input id="cmp-scale-input" type="range" min="0" max="' + (n - 1) +
                    '" step="1" value="' + (n - 1 - skip) + '">' +
                    '</div>' +
                    '<div class="cmp-scale-ticks">' + ticks + '</div>' +
                    '<div class="cmp-scale-tip">' + tipText + '</div>';
                document.getElementById('cmp-scale-input').oninput = function () {
                    axisSkip = (n - 1) - parseInt(this.value, 10);
                    buildChart();
                };
            }

            function refresh() { axisSkip = 0; buildTable(); buildChart(); }   // 游戏集合变化时拖动条回到右端
            function addToCompare(name) {
                var g = null;
                for (var i = 0; i < FULL.length; i++) {
                    if (FULL[i].name === name) { g = FULL[i]; break; }
                }
                // 按商店链接判断是否已在表中（同名不同链接是不同游戏）
                if (!g || cmp.some(function (x) { return x.value[3] === g.value[3]; })) return;
                cmp.push(g);
                refresh();
            }
            window.__addToCompare = addToCompare;   // 测试/调试钩子

            // ===== 批量导入：把"当前筛选结果"整批加入下方对比表格 =====
            // kind: all=全部 / green=绿色区域 / high=高好评 / low=低价
            function importFiltered(kind) {
                var has = {};
                cmp.forEach(function (g) { has[g.value[3]] = 1; });   // 按商店链接去重
                var added = 0;
                for (var i = 0; i < lastHits.length; i++) {
                    var g = lastHits[i];
                    if (kind === 'green' && !(g.value[0] < X_MED && g.value[1] > Y_MED)) continue;
                    if (kind === 'high' && !(g.value[1] > Y_MED)) continue;
                    if (kind === 'low' && !(g.value[0] < X_MED)) continue;
                    if (has[g.value[3]]) continue;
                    has[g.value[3]] = 1;
                    cmp.push(g);
                    added++;
                }
                refresh();
                document.getElementById('imp-tip').textContent = added
                    ? '已导入 ' + added + ' 款（表格共 ' + cmp.length + ' 款）'
                    : (lastHits.length ? '没有新增游戏（均已在表格中）' : '当前没有筛选出游戏');
            }
            document.getElementById('imp-all').addEventListener('click', function () { importFiltered('all'); });
            document.getElementById('imp-green').addEventListener('click', function () { importFiltered('green'); });
            document.getElementById('imp-high').addEventListener('click', function () { importFiltered('high'); });
            document.getElementById('imp-low').addEventListener('click', function () { importFiltered('low'); });
            document.getElementById('imp-clear').addEventListener('click', function () {
                cmp.length = 0;   // 清空当前表格与对比图
                refresh();
                document.getElementById('imp-tip').textContent = '已清空表格';
            });
            window.__importFiltered = importFiltered;   // 测试/调试钩子
            window.__cmp = cmp;                         // 测试/调试钩子

            // 行首 − 删除（事件委托，表格重建后依然有效）
            document.getElementById('cmp-table-box').addEventListener('click', function (e) {
                var btn = e.target.closest ? e.target.closest('.cmp-del') : null;
                if (!btn) return;
                // data-i 指向排序后数组的位置，用对象引用在原数组中定位删除
                var g = sortedCmp()[+btn.getAttribute('data-i')];
                var k = cmp.indexOf(g);
                if (k >= 0) { cmp.splice(k, 1); refresh(); }
            });

            // 对比图排序切换（表格与图表同步重排，按钮高亮当前排序）
            function setSort(key) {
                sortKey = key;
                // 排序只改变顺序、不改变集合，保留拖动条当前位置
                buildTable(); buildChart();
                document.getElementById('sort-cnt').classList.toggle('on', key === 'cnt');
                document.getElementById('sort-rating').classList.toggle('on', key === 'rating');
                document.getElementById('sort-price').classList.toggle('on', key === 'price');
            }
            document.getElementById('sort-cnt').addEventListener('click', function () { setSort('cnt'); });
            document.getElementById('sort-rating').addEventListener('click', function () { setSort('rating'); });
            document.getElementById('sort-price').addEventListener('click', function () { setSort('price'); });

            // 窗口尺寸变化时同步重算对比图画布，避免宽度与卡片边框不一致
            window.addEventListener('resize', function () { if (cmpChart) cmpChart.resize(); });

            refresh();   // 初始空状态提示

            var input = document.getElementById('game-search');
            var counter = document.getElementById('search-count');
            var tagInput = document.getElementById('tag-input');
            var tagPanel = document.getElementById('tag-panel');
            var tagMode = document.getElementById('tag-mode');
            var cntMin = document.getElementById('cnt-min');
            var cntMax = document.getElementById('cnt-max');
            var dateMin = document.getElementById('date-min');
            var dateMax = document.getElementById('date-max');

            // ---- 类型多选下拉 ----
            var selTags = [];
            TAGS.forEach(function (t) {
                var opt = document.createElement('div');
                opt.className = 'tag-opt';
                opt.innerHTML = '<span>' + t.name + '</span><span class="cnt">' + t.count + ' 款</span>';
                opt.onclick = function (e) {
                    e.stopPropagation();
                    var i = selTags.indexOf(t.name);
                    if (i >= 0) selTags.splice(i, 1); else selTags.push(t.name);
                    opt.classList.toggle('on', i < 0);
                    tagInput.value = selTags.join('、');
                    applyAll();
                };
                opt._tag = t.name;
                tagPanel.appendChild(opt);
            });
            tagInput.onclick = function (e) {
                e.stopPropagation();
                tagPanel.style.display = tagPanel.style.display === 'block' ? 'none' : 'block';
            };
            tagInput.addEventListener('keydown', function (e) {
                if (e.key === 'Backspace' && selTags.length) {
                    var last = selTags.pop();
                    tagInput.value = selTags.join('、');
                    for (var i = 0; i < tagPanel.children.length; i++) {
                        if (tagPanel.children[i]._tag === last) tagPanel.children[i].classList.remove('on');
                    }
                    applyAll();
                    e.preventDefault();
                }
            });
            document.addEventListener('click', function () { tagPanel.style.display = 'none'; });

            // ---- 日期输入解析：纯年份 -> 最早取1月/最晚取12月 ----
            function parseDateInput(s, isMin) {
                s = (s || '').trim();
                var m = s.match(/(\\d{4})(?:\\D+(\\d{1,2}))?(?:\\D+(\\d{1,2}))?/);
                if (!m) return isMin ? new Date(1900, 0, 1).getTime() : new Date().getTime();
                var y = +m[1], mo = m[2] ? +m[2] : null, dd = m[3] ? +m[3] : null;
                if (mo == null) return isMin ? new Date(y, 0, 1).getTime() : new Date(y, 11, 31, 23, 59, 59).getTime();
                if (dd == null) return isMin ? new Date(y, mo - 1, 1).getTime() : new Date(y, mo, 0, 23, 59, 59).getTime();
                return new Date(y, mo - 1, dd, 23, 59, 59).getTime();
            }

            // ---- 统一筛选（合取）：名称 ∧ 类型(析取) ∧ 评测数 ∧ 发布日期 ----
            function applyAll() {
                var q = input.value.trim().toLowerCase();
                var lo = cntMin.value === '' ? 0 : +cntMin.value;
                var hi = cntMax.value === '' ? Infinity : +cntMax.value;
                var dLo = parseDateInput(dateMin.value, true);
                var dHi = parseDateInput(dateMax.value, false);
                var dateFiltered = dateMin.value.trim() !== '' || dateMax.value.trim() !== '';

                var hits = [];
                for (var i = 0; i < FULL.length; i++) {
                    var g = FULL[i];
                    if (q && g.name.toLowerCase().indexOf(q) < 0) continue;
                    if (selTags.length) {
                        var ok = false;
                        if (tagMode.value === 'and') {
                            // 合取：必须同时包含所有选中标签
                            ok = true;
                            for (var j = 0; j < selTags.length; j++) {
                                if (g.tags.indexOf(selTags[j]) < 0) { ok = false; break; }
                            }
                        } else {
                            // 析取：包含任意一个选中标签即可
                            for (var j = 0; j < selTags.length; j++) {
                                if (g.tags.indexOf(selTags[j]) >= 0) { ok = true; break; }
                            }
                        }
                        if (!ok) continue;
                    }
                    if (g.value[2] < lo || g.value[2] > hi) continue;
                    if (visualRanges) {   // 图例已取消的评测数分段，点要真正移除（防幽灵点）
                        var inSeg = false;
                        for (var s = 0; s < visualRanges.length; s++) {
                            if (g.value[2] >= visualRanges[s][0] && g.value[2] <= visualRanges[s][1]) { inSeg = true; break; }
                        }
                        if (!inSeg) continue;
                    }
                    if (dateFiltered && (g.date == null || g.date < dLo || g.date > dHi)) continue;
                    hits.push(g);
                }
                chart.setOption({series: [{data: hits}]});
                lastHits = hits;   // 记录当前筛选结果，供"批量导入"按钮使用
                if (visualRanges) pushSelected();   // 重建数据后同步图例勾选状态
                if (hits.length === FULL.length) {
                    counter.textContent = '';
                    counter.style.display = 'none';
                } else {
                    counter.style.display = 'inline-block';
                    counter.textContent = hits.length ? '显示 ' + hits.length + ' 款' : '无匹配游戏';
                }
            }

            // 输入防抖
            var timer = null;
            function debounce() {
                clearTimeout(timer);
                timer = setTimeout(applyAll, 300);
            }
            input.addEventListener('input', debounce);
            tagMode.addEventListener('change', applyAll);
            cntMin.addEventListener('input', debounce);
            cntMax.addEventListener('input', debounce);
            dateMin.addEventListener('input', debounce);
            dateMax.addEventListener('input', debounce);
        })();
        """.replace("__ROWS_JSON__", rows_json).replace("__TAGS_JSON__", tags_json)
           .replace("__X_MED__", f"{x_med:g}").replace("__Y_MED__", f"{y_med:g}")
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
    ("性价比区(低价+好评)", (d["现价"] <= x_med) & (d["好评率%"] >= y_med)),
    ("高价避雷区(高价+差评)", (d["现价"] > x_med) & (d["好评率%"] < y_med)),
]:
    sub = d[mask].nlargest(5, "评测数")
    print(f"\n{name} 代表游戏（按评测数）:")
    for _, r in sub.iterrows():
        print(f"  {r['名称']} | ¥{r['现价']:.0f} | {r['好评率%']:.0f}% | {r['评测数']:,} 评测")
