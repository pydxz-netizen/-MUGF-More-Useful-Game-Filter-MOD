# -*- coding: utf-8 -*-
"""
Steam「日系角色扮演(JRPG)」类目爬虫（改良版）
=============================================
原脚本失败的两个原因：
1. SSL 证书验证失败 —— Steam 国内 CDN 返回的证书链不完整，Python 的 certifi
   验不过（浏览器能开是因为浏览器会自动补全证书链）。
   解决：用 truststore 让 Python 直接用 Windows 系统证书库。
2. 类目页是 React 动态渲染 —— HTML 源码里没有商品卡片，
   翻页也不是 URL 参数 offset=12/24/...，而是页面底部的「显示更多」按钮
   （每点一次加载 12 个，首屏自动加载的 12 个相当于 offset=0）。
   解决：用 Playwright 驱动本机 Edge 渲染页面，模拟滚动 + 点击按钮，
       按官方排序收集游戏 appid；详情再用稳定的 JSON 接口获取。

数据来源（全部为结构化 JSON，无需解析 HTML）：
- 类目列表:  Playwright 渲染 https://store.steampowered.com/category/rpg_jrpg
- 游戏详情:  https://store.steampowered.com/api/appdetails?appids=X
             （名称/类型/价格/发售日期）
- 用户评价:  https://store.steampowered.com/appreviews/X?json=1
             （好评率/评测数/评价摘要）

依赖：pip install requests pandas openpyxl playwright truststore
运行：py steamJRPG爬虫.py
"""

# ========== 第 1 步：导入必要的库 ==========
import random
import time

import pandas as pd                 # 数据整理 + 导出 Excel
import requests                     # 详情接口请求
import truststore
truststore.inject_into_ssl()        # 关键：修复 SSL 证书验证失败

from playwright.sync_api import sync_playwright   # 浏览器自动化

# ========== 第 2 步：全局配置 ==========
HUB_URL = "https://store.steampowered.com/category/rpg_jrpg?cc=CN&l=schinese" #目标位置
PER_PAGE = 12          # Steam 类目页每次「显示更多」加载 12 个（即 offset+12）
TOTAL_PAGES = 13       # 想要多少「页」（13 页 ≈ 156 款游戏）
OUT_XLSX = "steamJRPG信息收集.xlsx"  # 输出文件名
DETAIL_DELAY = (0.3, 0.6)   # 详情接口之间的随机延时（秒），防止被限流

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}

# 页面里统计 /app/ 链接的 JS（在浏览器里执行）
APPS_COUNT_JS = "document.querySelectorAll('a[href*=\"/app/\"]').length"
COUNT_JS = f"() => {APPS_COUNT_JS}"
IDS_JS = """() => {
    const out = [];
    for (const a of document.querySelectorAll('a[href*="/app/"]')) {
        const m = a.getAttribute('href').match(/\\/app\\/(\\d+)\\//);
        if (m && !out.includes(m[1])) out.push(m[1]);
    }
    return out;
}"""


# ========== 第 3 步：获取列表（浏览器渲染 + 模拟翻页） ==========
def collect_appids(pages: int = TOTAL_PAGES) -> list[str]:
    """打开类目页，通过「显示更多」按钮按每批 12 个收集游戏 appid（官方排序）"""
    target = pages * PER_PAGE
    with sync_playwright() as p:
        # channel="msedge" 直接用本机 Edge，无需额外下载浏览器
        browser = p.chromium.launch(channel="msedge", headless=True)
        page = browser.new_page(viewport={"width": 1400, "height": 900})
        page.goto(HUB_URL, timeout=60000, wait_until="domcontentloaded")
        page.wait_for_selector('a[href*="/app/"]', timeout=30000)
        print("首屏加载完成")

        stagnant = 0                      # 连续加载失败的次数
        while len(page.evaluate(IDS_JS)) < target and stagnant < 4:
            before = page.evaluate(COUNT_JS)
            # 等价于访问 offset=12/24/36... 的地址：点「显示更多」按钮
            try:
                page.locator("text=显示更多").first.click(timeout=3000)
            except Exception:
                pass
            # 兜底：把页面滚到底部，触发懒加载
            page.evaluate("() => window.scrollTo(0, document.body.scrollHeight)")
            try:
                page.wait_for_function(
                    f"(old) => {APPS_COUNT_JS} > old", arg=before, timeout=15000
                )
                stagnant = 0
                print(f"  已收集 {len(page.evaluate(IDS_JS))} 个游戏")
            except Exception:
                stagnant += 1
                print("  本批没有新内容，重试...")

        ids = page.evaluate(IDS_JS)
        browser.close()

    ids = ids[:target]
    print(f"列表收集完成：共 {len(ids)} 个 appid")
    return ids


# ========== 第 4 步：请求详情接口（纯 JSON，稳定） ==========
def fetch_json(url: str, params: dict | None = None, retries: int = 3):
    """GET 请求并返回 JSON，失败自动重试"""
    for attempt in range(1, retries + 1):
        try:
            resp = requests.get(url, params=params, headers=HEADERS, timeout=15)
            resp.raise_for_status()
            return resp.json()
        except Exception as e:
            print(f"  [重试 {attempt}/{retries}] {url} -> {e}")
            time.sleep(2)
    return None


def fetch_game_info(appid: str) -> dict | None:
    """合并 appdetails（详情）与 appreviews（评价）两个接口"""
    d = fetch_json("https://store.steampowered.com/api/appdetails",
                   params={"appids": appid, "l": "schinese", "cc": "CN"})
    if not d or not d.get(appid, {}).get("success"):
        return None
    info = d[appid]["data"]

    # 价格：price_overview 存在时含原价/现价/折扣；免费游戏没有该字段
    po = info.get("price_overview")
    if po:
        price_now = po.get("final_formatted", "")
        price_origin = po.get("initial_formatted") or price_now
        discount = po.get("discount_percent", 0)
    elif info.get("is_free"):
        price_now, price_origin, discount = "免费", "免费", 0
    else:
        price_now, price_origin, discount = "暂不可购", "暂不可购", 0

    # 评价：query_summary 里是中文摘要 + 好评数/总数
    rv = fetch_json(f"https://store.steampowered.com/appreviews/{appid}",
                    params={"json": 1, "num_per_page": 0,
                            "language": "schinese", "purchase_type": "all"})
    q = (rv or {}).get("query_summary", {})
    total, pos = q.get("total_reviews", 0), q.get("total_positive", 0)
    positive_pct = round(pos * 100 / total) if total else None

    return {
        "appid": appid,
        "名称": info.get("name", ""),
        "类型": "、".join(g["description"] for g in info.get("genres", [])),
        "评价": q.get("review_score_desc", "暂无评价"),
        "好评率%": positive_pct,
        "评测数": total,
        "原价": price_origin,
        "现价": price_now,
        "折扣%": discount,
        "发售日期": info.get("release_date", {}).get("date", ""),
        "商店链接": f"https://store.steampowered.com/app/{appid}",
    }


# ========== 第 5 步：处理数据 ==========
def process_data(rows: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    df = df.drop_duplicates(subset="appid")            # 按 appid 去重
    df.insert(0, "序号", range(1, len(df) + 1))
    return df.drop(columns=["appid"])


# ========== 第 6 步：导出 Excel ==========
def save_to_excel(df: pd.DataFrame, path: str = OUT_XLSX):
    df.to_excel(path, sheet_name="JRPG数据", index=False) #工作表名
    print(f"已保存 {len(df)} 条数据 -> {path}")


# ========== 主流程 ==========
def main():
    ids = collect_appids(TOTAL_PAGES)
    if not ids:
        print("未获取到任何游戏 ID，退出")
        return

    rows, failed = [], 0
    for i, appid in enumerate(ids, 1):
        info = fetch_game_info(appid)
        if info is None:
            failed += 1
            print(f"[{i}/{len(ids)}] {appid} 详情获取失败，跳过")
            continue
        rows.append(info)
        print(f"[{i}/{len(ids)}] {info['名称']} | {info['现价']} | {info['评价']}")
        time.sleep(random.uniform(*DETAIL_DELAY))      # 礼貌延时

    if not rows:
        print("未获取到任何数据")
        return

    print(f"完成：成功 {len(rows)} 条，失败 {failed} 条")
    df = process_data(rows)
    print(df.head(10))
    save_to_excel(df)


if __name__ == "__main__":
    main()
