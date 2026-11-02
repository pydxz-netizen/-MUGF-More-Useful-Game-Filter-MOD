# -*- coding: utf-8 -*-
"""
爬虫基本框架：发送请求 → 解析响应 → 处理数据 → 导出 Excel
演示目标：https://quotes.toscrape.com（一个专门供爬虫练习的示例网站）
依赖：pip install requests beautifulsoup4 pandas openpyxl
"""

# ========== 第 1 步：导入必要的库 ==========
import time
import random

import requests                     # 发送 HTTP 请求
from bs4 import BeautifulSoup      # 解析 HTML
import pandas as pd                # 数据整理 + 导出 Excel

# ========== 第 2 步：全局配置 ==========
BASE_URL = "https://quotes.toscrape.com/page/{}/"   # 目标网址（带分页）
MAX_PAGE = 3                                        # 要爬取的页数
DELAY_RANGE = (1, 2)                                # 每次请求之间的随机间隔（秒），礼貌爬取

HEADERS = {
    # 伪装成浏览器，避免被网站拒绝
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}


# ========== 第 3 步：发送请求 ==========
def fetch_html(url: str, retries: int = 3) -> str | None:
    """请求页面，失败自动重试，成功返回 HTML 文本"""
    for attempt in range(1, retries + 1):
        try:
            resp = requests.get(url, headers=HEADERS, timeout=10)
            resp.raise_for_status()          # 状态码非 2xx 时抛出异常
            resp.encoding = resp.apparent_encoding
            return resp.text
        except requests.RequestException as e:
            print(f"[重试 {attempt}/{retries}] 请求失败: {url} -> {e}")
            time.sleep(2)
    return None


# ========== 第 4 步：解析响应 ==========
def parse_html(html: str) -> list[dict]:
    """从 HTML 中提取目标字段，返回列表[字典]结构"""
    soup = BeautifulSoup(html, "lxml")
    items = []
    for quote in soup.select("div.quote"):          # 每条名言所在节点
        items.append({
            "名言": quote.select_one("span.text").get_text(strip=True),
            "作者": quote.select_one("small.author").get_text(strip=True),
            "标签": "、".join(t.get_text(strip=True) for t in quote.select("a.tag")),
        })
    return items


# ========== 第 5 步：处理数据 ==========
def process_data(rows: list[dict]) -> pd.DataFrame:
    """清洗并整理成表格"""
    df = pd.DataFrame(rows)
    df = df.drop_duplicates()                       # 去重
    df = df.dropna()                                # 去掉空行
    df.insert(0, "序号", range(1, len(df) + 1))     # 加序号列
    return df


# ========== 第 6 步：导出 Excel ==========
def save_to_excel(df: pd.DataFrame, path: str = "爬取结果.xlsx"):
    """写入 Excel 文件（需要 openpyxl）"""
    df.to_excel(path, sheet_name="名言数据", index=False)
    print(f"已保存 {len(df)} 条数据 -> {path}")


# ========== 主流程 ==========
def main():
    all_rows = []
    for page in range(1, MAX_PAGE + 1):
        url = BASE_URL.format(page)
        print(f"正在爬取第 {page} 页: {url}")
        html = fetch_html(url)
        if html is None:
            print("该页请求失败，跳过")
            continue
        rows = parse_html(html)
        if not rows:                                # 没有内容说明已到最后一页
            print("没有更多内容，停止翻页")
            break
        all_rows.extend(rows)
        time.sleep(random.uniform(*DELAY_RANGE))    # 随机延时，降低请求频率

    if not all_rows:
        print("未获取到任何数据")
        return

    df = process_data(all_rows)
    print(df.head(10))                              # 预览前 10 行
    save_to_excel(df)


if __name__ == "__main__":
    main()
