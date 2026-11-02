# -*- coding: utf-8 -*-
"""
数据预处理：合并 8 份 steam*.xlsx → 按商店链接去重 → 价格清洗 → steam_totdata.xlsx
对应作业要求（数据预处理.md）：
1. 8 份 excel 合并，序号依次往下标
2. 去重：商店链接相同视为同一游戏（名字相同但链接不同的算不同游戏）
3. 清洗：原价/现价去掉 ¥ 转数字，免费 = 0.00
"""

import glob
import re

import pandas as pd

OUT_XLSX = "steam_totdata.xlsx"

# ========== 1. 合并 8 份文件 ==========
files = [f for f in sorted(glob.glob("steam*.xlsx")) if "totdata" not in f]
frames = []
for f in files:
    df = pd.read_excel(f)
    # 从文件名提取来源类目（如 JRPG / HORROR），后续可视化做分类对比必需
    df["类目"] = re.search(r"steam(.+?)信息收集", f).group(1)
    frames.append(df)

data = pd.concat(frames, ignore_index=True)
print(f"合并 {len(files)} 个文件，共 {len(data)} 行")

# ========== 2. 去重（按商店链接） ==========
before = len(data)
data = data.drop_duplicates(subset="商店链接", keep="first").reset_index(drop=True)
print(f"去重：{before} -> {len(data)} 行（删除 {before - len(data)} 条重复）")

# ========== 3. 价格清洗 ==========
def clean_price(v):
    """'¥ 268.20' -> 268.20；'免费' -> 0.00；'暂不可购'等异常 -> 空"""
    s = str(v).replace("¥", "").replace(",", "").strip()
    if s == "免费":
        return 0.00
    try:
        return float(s)
    except ValueError:
        return None

for col in ["原价", "现价"]:
    data[col] = data[col].map(clean_price)

for col in ["折扣%", "好评率%", "评测数"]:
    data[col] = pd.to_numeric(data[col], errors="coerce")

# ========== 4. 序号连续重编 + 导出 ==========
data["序号"] = range(1, len(data) + 1)
cols = ["序号", "类目", "名称", "类型", "评价", "好评率%", "评测数",
        "原价", "现价", "折扣%", "发售日期", "商店链接"]
data = data[cols]

data.to_excel(OUT_XLSX, sheet_name="Steam数据", index=False)
print(f"已保存 {len(data)} 条 -> {OUT_XLSX}")

# 快速自检
print("\n各文件行数:", {re.search(r'steam(.+?)信息', f).group(1): len(pd.read_excel(f)) for f in files})
print("各类目剩余条数:\n", data["类目"].value_counts())
print("\n现价统计:\n", data["现价"].describe().round(2))
print("\n现价为空(暂不可购等):", data["现价"].isna().sum(), "条")
print(data.head(8).to_string())
