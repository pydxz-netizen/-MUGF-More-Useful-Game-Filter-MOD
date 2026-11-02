# -*- coding: utf-8 -*-
"""
Steam 全站热门游戏扩充爬虫（与原有 8 份类目爬虫同源，均为 store.steampowered.com 官网数据）
=============================================================================================
思路：
- 列表页用搜索接口（服务端渲染的 JSON 片段，带 start/count 翻页，无需 Playwright）
  https://store.steampowered.com/search/results/?infinite=1&start=...&count=100&category1=998
  默认排序 = Steam 官方热度顺序，前 1 万名即全站最主流的游戏。
- 一次请求即可拿到：appid / 名称 / 发售日期 / 好评率% / 评测数 / 评价摘要 / 原价 / 现价 / 折扣% / 用户标签ID
  （卡片里 data-tooltip-html 就是「N 篇评测中有 M% 为好评」，与详情接口同源）
- 类型列沿用原数据格式（Steam 官方类型 genres），由标签 ID 映射到官方类型名：
  标签字典 https://store.steampowered.com/tagdata/populartags/schinese
- 断点续爬：每页结果缓存到 _steam_cache/，失败重试，中断后重跑自动跳过已抓页面

运行：
    py steam扩展爬虫.py          # 全量（默认抓 96 页 ≈ 9600 款，约 8~15 分钟）
    py steam扩展爬虫.py 3        # 只抓 3 页试跑
"""

import html as html_lib
import json
import pathlib
import random
import re
import sys
import time

import pandas as pd
import requests
import truststore
truststore.inject_into_ssl()   # 关键：修复 Steam 证书链不完整导致的 SSL 失败

# ========== 配置 ==========
OUT_XLSX = "steam扩展信息收集.xlsx"
CACHE_DIR = pathlib.Path("_steam_cache")
TARGET_PAGES = 100         # 100 页 × 100 = 10000 款（去重后与原有类目数据合并 ≈ 10100）
COUNT_PER_PAGE = 100       # 列表接口每页条数
MAX_RETRY = 6              # 单页最多重试次数
DELAY = (0.6, 1.2)         # 礼貌延时（秒）

# 原有数据使用的 20 个 Steam 官方类型名（类型列与旧数据保持一致）
GENRE_SET = {
    "动作", "冒险", "角色扮演", "策略", "模拟", "休闲", "独立", "体育",
    "免费开玩", "抢先体验", "大型多人在线", "竞速", "动画制作和建模",
    "设计和插画", "实用工具", "视频制作", "教育", "网络出版", "游戏开发", "照片编辑",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}
SESSION = requests.Session()
SESSION.headers.update(HEADERS)

LIST_URL = "https://store.steampowered.com/search/results/"
TAG_URL = "https://store.steampowered.com/tagdata/populartags/schinese"


# ========== 通用请求（超时/断线自动重试） ==========
def get_json(url: str, params: dict | None = None, tries: int = MAX_RETRY):
    for i in range(1, tries + 1):
        try:
            r = SESSION.get(url, params=params, timeout=30)
            if r.status_code == 200:
                return r.json()
            print(f"  [{i}/{tries}] HTTP {r.status_code}，重试")
        except Exception as e:
            print(f"  [{i}/{tries}] {type(e).__name__}，重试")
        time.sleep(min(2 * i, 10))
    return None


# ========== 1. 标签字典：标签ID -> 中文名（用于映射官方类型） ==========
def load_tag_map() -> dict[str, str]:
    """优先读本地缓存；失败时多次重试，仍失败则直接退出（否则类型列会全空）"""
    CACHE_DIR.mkdir(exist_ok=True)
    cache = CACHE_DIR / "tags_schinese.json"
    if cache.exists():
        tag_map = {str(k): v for k, v in json.loads(cache.read_text(encoding="utf-8")).items()}
        if tag_map:
            print(f"标签字典（本地缓存）：{len(tag_map)} 条")
            return tag_map

    tag_map = {}
    for lang in ("schinese", "english"):
        data = get_json(TAG_URL.replace("schinese", lang), tries=10)
        if not data:
            print(f"  标签字典 {lang} 获取失败")
            continue
        for t in data:
            tag_map.setdefault(str(t.get("tagid")), t.get("name", ""))
    if not tag_map:
        raise SystemExit("标签字典获取失败（网络问题），请重跑；已抓取的页面有缓存不会重复下载")
    cache.write_text(json.dumps(tag_map, ensure_ascii=False), encoding="utf-8")
    print(f"标签字典：{len(tag_map)} 条")
    return tag_map


# ========== 2. 列表页抓取（带缓存，可断点续爬） ==========
def fetch_list_page(page: int) -> str | None:
    CACHE_DIR.mkdir(exist_ok=True)
    cache = CACHE_DIR / f"page_{page:04d}.json"
    if cache.exists():
        return cache.read_text(encoding="utf-8")
    params = {"query": "", "start": page * COUNT_PER_PAGE, "count": COUNT_PER_PAGE,
              "category1": 998, "cc": "CN", "l": "schinese", "infinite": 1}
    j = get_json(LIST_URL, params)
    if not j or "results_html" not in j:
        return None
    cache.write_text(json.dumps(j, ensure_ascii=False), encoding="utf-8")
    return cache.read_text(encoding="utf-8")


# ========== 3. 卡片解析 ==========
def text_of(fragment: str) -> str:
    return html_lib.unescape(re.sub(r"<[^>]+>", " ", fragment)).strip()


def parse_card(card: str, tag_map: dict[str, str]) -> dict | None:
    m_appid = re.search(r'/app/(\d+)', card)
    m_name = re.search(r'<span class="title">(.*?)</span>', card, re.S)
    if not (m_appid and m_name):
        return None
    appid = m_appid.group(1)
    name = text_of(m_name.group(1))

    # 发售日期
    m_rel = re.search(r'class="search_released[^"]*"[^>]*>([\s\S]*?)</div>', card)
    released = text_of(m_rel.group(1)) if m_rel else ""

    # 评价：data-tooltip-html = "好评如潮<br>此游戏的 2,033 篇用户评测中有 99% 为好评。..."
    desc, pct, cnt = "", None, 0
    m_tip = re.search(r'data-tooltip-html="([^"]*)"', card)
    if m_tip:
        parts = [p.strip() for p in html_lib.unescape(m_tip.group(1)).split("<br>") if p.strip()]
        if parts:
            desc = parts[0]
        joined = " ".join(parts)
        m_num = re.search(r"([\d,]+)\s*篇.*?(\d+)%", joined)
        if m_num:
            cnt = int(m_num.group(1).replace(",", ""))
            pct = int(m_num.group(2))

    # 价格：免费游戏卡片里 discount_final_price 的文字就是「免费」
    m_final = re.search(r'class="discount_final_price[^"]*">([\s\S]*?)</div>', card)
    m_orig = re.search(r'class="discount_original_price">([\s\S]*?)</div>', card)
    m_disc = re.search(r'data-discount="(\d+)"', card)
    final_txt = text_of(m_final.group(1)) if m_final else ""
    if "免费" in final_txt:
        now, origin = "免费", "免费"
    elif "¥" in final_txt:
        now = "¥ " + final_txt.replace("¥", "").replace(",", "").strip()
        origin = now
        if m_orig:
            origin = "¥ " + text_of(m_orig.group(1)).replace("¥", "").replace(",", "").strip()
    else:
        now, origin = "暂不可购", "暂不可购"
    discount = int(m_disc.group(1)) if m_disc else 0

    # 标签 -> 官方类型（取交集，保证与旧数据同一套类型名）
    types = []
    m_tags = re.search(r'data-ds-tagids="\[([^\]]*)\]"', card)
    if m_tags:
        for tid in m_tags.group(1).split(","):
            nm = tag_map.get(tid.strip())
            if nm in GENRE_SET and nm not in types:
                types.append(nm)
    if now == "免费" and "免费开玩" not in types:
        types.append("免费开玩")

    return {
        "appid": appid,
        "名称": name,
        "类型": "、".join(types),
        "评价": desc or "暂无评价",
        "好评率%": pct,
        "评测数": cnt,
        "原价": origin,
        "现价": now,
        "折扣%": discount,
        "发售日期": released,
        "商店链接": f"https://store.steampowered.com/app/{appid}",
    }


# ========== 4. 主流程 ==========
def main():
    pages = int(sys.argv[1]) if len(sys.argv) > 1 else TARGET_PAGES
    tag_map = load_tag_map()

    rows, seen = [], set()
    for page in range(pages):
        raw = fetch_list_page(page)
        if not raw:
            print(f"第 {page + 1}/{pages} 页抓取失败，跳过")
            continue
        html_text = json.loads(raw)["results_html"]
        # 每个游戏是一整段 <a class="search_result_row">...</a>，内部没有嵌套 <a>
        card_blocks = re.findall(
            r'(<a href="https://store\.steampowered\.com/app/\d+[\s\S]*?</a>)', html_text
        )
        got = 0
        for block in card_blocks:
            info = parse_card(block, tag_map)
            if not info or info["appid"] in seen:
                continue
            seen.add(info["appid"])
            rows.append(info)
            got += 1
        print(f"第 {page + 1}/{pages} 页：解析 {got} 款（累计 {len(rows)}）")

        # 每 10 页落一次盘，避免中途失败全丢
        if (page + 1) % 10 == 0:
            save(rows)
        time.sleep(random.uniform(*DELAY))

    save(rows)

    df = pd.DataFrame(rows)
    print(f"\n共 {len(df)} 款 | appid 去重后 {df['appid'].nunique()} 款")
    print("有类型标注:", (df["类型"] != "").sum(), "| 有评价数据:", df["好评率%"].notna().sum())
    print("\n类型分布（前 12）:\n", df["类型"].str.split("、").explode().value_counts().head(12))
    print("\n价格示例:\n", df[["名称", "原价", "现价", "好评率%", "评测数"]].head(6).to_string())


def save(rows: list[dict]):
    if not rows:
        return
    df = pd.DataFrame(rows).drop_duplicates(subset="appid")
    df.insert(0, "序号", range(1, len(df) + 1))
    df.drop(columns=["appid"]).to_excel(OUT_XLSX, sheet_name="扩展数据", index=False)
    print(f"  -> 已写入 {len(df)} 条 {OUT_XLSX}")


if __name__ == "__main__":
    main()
