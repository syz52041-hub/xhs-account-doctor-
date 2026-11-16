#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成演示数据：两份合成「笔记列表明细表.xlsx」（虚构数据，非真实账号）
用途：没有真实导出数据时，先跑这个生成演示文件，体验完整诊断流程。

    python make_demo_data.py

输出：./demo_data/笔记列表明细表_0930.xlsx / _1015.xlsx
之后即可运行诊断：
    python account_doctor.py ./demo_data/笔记列表明细表_0930.xlsx ./demo_data/笔记列表明细表_1015.xlsx
"""
import datetime as dt
import random
import os

import pandas as pd

random.seed(42)

TITLES = [
    ("干货", "新手做小红书必看的5个避坑干货"),
    ("干货", "涨粉慢？我总结了这套选题清单"),
    ("干货", "一条笔记的完整制作教程，从标题到封面"),
    ("干货", "新手博主常用方法：内容线拆解"),
    ("热点", "这瓜反转得离谱，全网都看错了"),
    ("热点", "年度名场面盘点，最后一个梗封神"),
    ("热点", "最近争议很大的事，说点别人不敢说的"),
    ("热点", "热搜上的事，聊聊我的看法"),
    ("日常", "今天出门的随手拍"),
    ("日常", "周末逛了趟集市"),
    ("干货", "被问爆的问题统一回答：关于定位"),
    ("热点", "评论区吵翻的瓜，我扒到了源头"),
    ("日常", "记录一个普通工作日"),
    ("干货", "拍摄角度攻略，手机也能拍大片"),
]


def gen_rows():
    rows = []
    t0 = dt.datetime(2026, 9, 1, 19, 30)
    for i, (tag, title) in enumerate(TITLES):
        exp = random.choice([0, 60, 180, 900, 3500, 12000, 28000])
        views = int(exp * random.uniform(0.08, 0.15)) if exp else 0
        likes = int(views * random.uniform(0.05, 0.12)) if views else 0
        comments = int(likes * random.uniform(0.05, 0.25)) if likes else 0
        collects = int(likes * random.uniform(0.1, 0.5)) if likes else 0
        shares = int(likes * 0.08) if likes else 0
        fans = int(likes * random.uniform(0, 0.06))
        ctr = random.uniform(0.04, 0.14) if exp else 0.0
        rows.append({
            "笔记标题": title,
            "首次发布时间": (t0 + dt.timedelta(days=i * 3)).strftime("%Y年%m月%d日%H时%M分%S秒"),
            "体裁": random.choice(["图文", "视频"]),
            "曝光": exp,
            "观看量": views,
            "封面点击率": round(ctr, 3),
            "点赞": likes,
            "评论": comments,
            "收藏": collects,
            "涨粉": fans,
            "分享": shares,
            "人均观看时长": 0.0,
            "弹幕": 0,
        })
    return rows


def main():
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "demo_data")
    os.makedirs(out_dir, exist_ok=True)

    df1 = pd.DataFrame(gen_rows())

    # 第二份导出：部分笔记数据增长 + 封面点击率回落（用于演示去重口径）
    df2 = df1.head(9).copy()
    for idx, r in df2.iterrows():
        r["曝光"] = int(r["曝光"] * 1.3) + 20
        r["观看量"] = int(r["观看量"] * 1.25)
        r["点赞"] = int(r["点赞"] * 1.2) + 3
        r["封面点击率"] = round(r["封面点击率"] * 0.75, 3)
    r_new = dict(df1.iloc[0])
    r_new["笔记标题"] = "新增：一篇迟到的复盘"
    r_new["涨粉"] = 12
    df2 = pd.concat([df2, pd.DataFrame([r_new])], ignore_index=True)

    p1 = os.path.join(out_dir, "笔记列表明细表_0930.xlsx")
    p2 = os.path.join(out_dir, "笔记列表明细表_1015.xlsx")
    df1.to_excel(p1, index=False)
    df2.to_excel(p2, index=False)
    print("演示数据已生成：")
    print("  ", p1)
    print("  ", p2)


if __name__ == "__main__":
    main()
