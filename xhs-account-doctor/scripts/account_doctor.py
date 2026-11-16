#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
小红书账号诊断器 (XHS Account Doctor)
=====================================
输入：小红书创作服务平台导出的「笔记列表明细表.xlsx」（支持多份合并）
输出：一份零外部依赖、手机友好的 HTML 诊断看板（数据不外泄、离线可看）

用法：
    python account_doctor.py                          # 自动扫描 ~/Downloads 里的笔记导出表
    python account_doctor.py "路径/文件.xlsx"          # 指定文件（可多个）
    python account_doctor.py --by-source a.xlsx b.xlsx # 按文件分账号/分平台对比
    python account_doctor.py --compare a.xlsx b.xlsx   # 自己 vs 对标 并排看板
    python account_doctor.py --config config.json      # 自定义内容线打标规则

依赖：pandas + openpyxl（pip install pandas openpyxl）

设计原则：
    - 零成本：数据来自官方后台导出，不买第三方数据服务、不爬接口
    - 零风险：手动导入，不碰登录态、不踩封号灰产
    - 不编造：所有结论都能追溯到后台真实数字；预期收益只做保守估算并明确标注
    - 本地化：产物是本地 HTML，不上传任何第三方
"""

import os
import sys
import re
import json
import html
import datetime as dt

import pandas as pd

# ----------------------------------------------------------------------------
# 配置
# ----------------------------------------------------------------------------
WORKDIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOADS = os.path.expanduser("~/Downloads")

# 内容线打标默认规则（按标题关键词）。任何人/账号都不同，建议用 --config 换成自己的。
DEFAULT_CONFIG = {
    "default_tag": "日常",
    "score_weights": {"fans": 5, "interactions": 1, "exposure": 0.01},
    "tag_rules": {
        "干货": ["干货", "教程", "攻略", "避坑", "总结", "清单", "方法", "新手"],
        "热点": ["爆料", "瓜", "争议", "反转", "离谱", "热搜", "名场面", "梗"],
    },
}

CFG = dict(DEFAULT_CONFIG)

TAG_PALETTE = [
    ("rgba(74,163,255,.18)", "#4aa3ff"),
    ("rgba(255,176,32,.18)", "#ffb020"),
    ("rgba(255,59,78,.16)", "#ff3b4e"),
    ("rgba(46,204,113,.18)", "#2ecc71"),
    ("rgba(178,102,255,.18)", "#b266ff"),
    ("rgba(0,206,209,.18)", "#00ced1"),
]


def load_config(path):
    """加载用户自定义配置（tag_rules / default_tag / score_weights）。"""
    global CFG
    if not path:
        return
    with open(path, "r", encoding="utf-8") as f:
        user = json.load(f)
    CFG = dict(DEFAULT_CONFIG)
    CFG.update({k: v for k, v in user.items() if v})


def tag_rules():
    return CFG["tag_rules"]


def default_tag():
    return CFG["default_tag"]


def tagname(t):
    return f"{t}线"


def all_tags():
    ts = list(CFG["tag_rules"].keys())
    if default_tag() not in ts:
        ts.append(default_tag())
    return ts


# 小红书后台导出列名 -> 内部字段
COL_MAP = {
    "笔记标题": "title",
    "首次发布时间": "publish_time",
    "体裁": "type",
    "曝光": "exposure",
    "观看量": "views",
    "封面点击率": "cover_ctr",
    "点赞": "likes",
    "评论": "comments",
    "收藏": "collects",
    "涨粉": "fans_gain",
    "分享": "shares",
    "人均观看时长": "avg_watch",
    "弹幕": "danmu",
}
NUM_COLS = ["exposure", "views", "cover_ctr", "likes", "comments",
            "collects", "fans_gain", "shares", "avg_watch", "danmu"]

# 综合得分公式（透明、可解释，权重可在 config.json 里改）
#   默认：涨粉权重最高（核心目标）；互动衡量内容质量；曝光衡量流量基数
def score_row(r):
    w = CFG["score_weights"]
    return r["fans_gain"] * w["fans"] + r["interactions"] + r["exposure"] * w["exposure"]


# ----------------------------------------------------------------------------
# 数据加载与清洗
# ----------------------------------------------------------------------------
def classify(title):
    t = str(title)
    for tag, kws in tag_rules().items():
        if any(k in t for k in kws):
            return tag
    return default_tag()


def parse_time(s):
    s = str(s)
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日(\d{1,2})时(\d{1,2})分(\d{1,2})秒", s)
    if m:
        return dt.datetime(*map(int, m.groups()))
    return None


def to_num(v):
    if pd.isna(v):
        return 0
    s = str(v).replace("%", "").replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0


def load_file(path):
    # 检测真实表头在第几行：第 1 行可能是导出占位标题
    raw = pd.read_excel(path, sheet_name=0, header=None, nrows=2)
    first = str(raw.iloc[0, 0])
    header = 0 if "笔记标题" in first else 1

    df = pd.read_excel(path, sheet_name=0, header=header)
    df = df.rename(columns=COL_MAP)
    df = df[[c for c in COL_MAP.values() if c in df.columns]].copy()

    for c in NUM_COLS:
        if c in df.columns:
            df[c] = df[c].apply(to_num)
        else:
            df[c] = 0

    df["publish_time"] = df["publish_time"].apply(parse_time)
    df["tag"] = df["title"].apply(classify)
    df["source"] = os.path.splitext(os.path.basename(path))[0]
    df["snapshot_ts"] = os.path.getmtime(path)   # 供 dedup 判定"哪份是最新快照"

    df["interactions"] = df["likes"] + df["comments"] + df["collects"] + df["shares"]
    df["interaction_rate"] = (df["interactions"] / df["views"].replace(0, pd.NA)).fillna(0)
    df["fan_eff"] = (df["fans_gain"] / df["exposure"].replace(0, pd.NA)).fillna(0)
    df["score"] = df.apply(score_row, axis=1)
    return df


def load_all(paths, by_source):
    frames = []
    for p in paths:
        df = load_file(p)
        frames.append(df)
    data = pd.concat(frames, ignore_index=True)
    return data, None


# ----------------------------------------------------------------------------
# 分析模块
# ----------------------------------------------------------------------------
def overview(data):
    return {
        "notes": len(data),
        "exposure": int(data["exposure"].sum()),
        "views": int(data["views"].sum()),
        "fans": int(data["fans_gain"].sum()),
        "interactions": int(data["interactions"].sum()),
        "ir": data["interaction_rate"].mean() if len(data) else 0,
        "avg_ctr": data["cover_ctr"].mean() if len(data) else 0,
        "date_min": data["publish_time"].min(),
        "date_max": data["publish_time"].max(),
        "sources": data["source"].nunique(),
    }


def module_redblack(data, n=5):
    d = data.sort_values("score", ascending=False).reset_index(drop=True)
    red = d.head(n)
    black = d.tail(n).sort_values("score")
    return red, black


def module_bomb(data):
    """爆款规律：按内容线分组均值 + 体裁分布 + 标题长度 vs 表现"""
    grp = data.groupby("tag").agg(
        n=("title", "count"),
        avg_exposure=("exposure", "mean"),
        avg_views=("views", "mean"),
        avg_interactions=("interactions", "mean"),
        avg_ir=("interaction_rate", "mean"),
        avg_fans=("fans_gain", "mean"),
        avg_fan_eff=("fan_eff", "mean"),
    ).reset_index().sort_values("avg_fans", ascending=False)

    type_dist = data.groupby("type").size().reset_index(name="n").sort_values("n", ascending=False)

    data = data.copy()
    data["title_len"] = data["title"].astype(str).str.len()
    len_corr = data["title_len"].corr(data["interactions"])
    return grp, type_dist, len_corr


def module_tag_compare(data):
    """内容线对比"""
    grp = data.groupby("tag").agg(
        n=("title", "count"),
        exposure=("exposure", "sum"),
        avg_exposure=("exposure", "mean"),
        avg_ir=("interaction_rate", "mean"),
        fans=("fans_gain", "sum"),
        avg_fan_eff=("fan_eff", "mean"),
    ).reset_index()
    return grp


def module_health(data):
    """流量健康诊断"""
    exp_mean = data["exposure"].mean()
    ir_mean = data["interaction_rate"].mean()

    zero_exp = data[data["exposure"] == 0]
    # 高曝光低互动：曝光高于均值但互动率低于均值
    leaked = data[(data["exposure"] > exp_mean) & (data["interaction_rate"] < ir_mean)]
    low_ctr = data[data["cover_ctr"] < data["cover_ctr"].mean()]

    return {
        "exp_mean": exp_mean,
        "ir_mean": ir_mean,
        "ctr_mean": data["cover_ctr"].mean(),
        "zero_exp_n": len(zero_exp),
        "leaked": leaked.sort_values("exposure", ascending=False).head(8),
        "low_ctr_n": len(low_ctr),
        "ctr_dist": data["cover_ctr"].describe(),
    }


# ----------------------------------------------------------------------------
# 分析层：漏斗诊断 / 归因 / 策略 / 可执行优化清单
# ----------------------------------------------------------------------------
def module_funnel(data):
    """曝光 -> 观看 -> 互动 -> 涨粉 的四段漏斗，定位最大漏点。"""
    e = data["exposure"].sum()
    v = data["views"].sum()
    i = data["interactions"].sum()
    f = data["fans_gain"].sum()
    steps = [
        ("曝光→观看", e, v),
        ("观看→互动", v, i),
        ("互动→涨粉", i, f),
        ("曝光→涨粉", e, f),
    ]
    rates = {}
    for name, a, b in steps:
        rates[name] = (b / a) if a else 0.0
    # 漏点判定：末端转化最低的那段
    leak = "互动→涨粉" if rates["互动→涨粉"] < 0.05 else "曝光→观看"
    zero_fan = int((data["fans_gain"] == 0).sum())
    return {
        "e": e, "v": v, "i": i, "f": f,
        "rates": rates,
        "leak": leak,
        "zero_fan": zero_fan,
        "zero_fan_ratio": zero_fan / max(len(data), 1),
    }


def _note_bucket(r, exp_mean):
    if r["exposure"] < 100:
        return "没流量（发布/推荐问题）"
    if r["fans_gain"] == 0:
        return "有流量没接住（缺关注引导）"
    return "有效涨粉"


def module_attr(data):
    """红黑榜归因：红榜共性 + 黑榜分桶。"""
    red = data.sort_values("score", ascending=False).head(5)
    black = data.sort_values("score").head(5)
    exp_mean = data["exposure"].mean()

    # 黑榜逐篇分桶
    buckets = {}
    for _, r in data.iterrows():
        b = _note_bucket(r, exp_mean)
        buckets.setdefault(b, []).append(r["title"])
    # 高质量但被埋没：曝光<100 且互动率进前 40%
    ir_thr = data["interaction_rate"].quantile(0.6)
    starved = data[(data["exposure"] < 100) & (data["interaction_rate"] >= ir_thr)]
    starved = starved.sort_values("interaction_rate", ascending=False)

    red_tags = red["tag"].value_counts().to_dict()
    return {
        "red_n": len(red),
        "red_tags": red_tags,
        "red_titles": list(red["title"].astype(str).str[:24]),
        "red_avg_exp": red["exposure"].mean(),
        "red_avg_ir": red["interaction_rate"].mean(),
        "black_titles": list(black["title"].astype(str).str[:24]),
        "buckets": {k: len(v) for k, v in buckets.items()},
        "starved_titles": list(starved["title"].astype(str).str[:24]),
        "starved_n": len(starved),
    }


def module_strategy(data):
    """内容线策略：谁做流量引擎、谁做质量引擎、谁聚焦。"""
    g = data.groupby("tag").agg(
        n=("title", "count"),
        exp=("exposure", "sum"),
        avg_exp=("exposure", "mean"),
        avg_ir=("interaction_rate", "mean"),
        fans=("fans_gain", "sum"),
        fan_eff=("fan_eff", "mean"),
    ).reset_index()
    g = g[g["n"] > 0]
    if g.empty:
        return []
    traffic_line = g.loc[g["exp"].idxmax(), "tag"]          # 流量引擎
    quality_line = g.loc[g["avg_ir"].idxmax(), "tag"]       # 质量引擎
    rows = []
    for _, r in g.iterrows():
        role = []
        if r["tag"] == traffic_line:
            role.append("流量引擎（保曝光）")
        if r["tag"] == quality_line:
            role.append("质量引擎（互动最高）")
        if r["n"] <= 2:
            role.append("样本太少，先加产观察")
        rows.append({
            "tag": r["tag"], "n": int(r["n"]), "exp": int(r["exp"]),
            "avg_ir": r["avg_ir"], "fans": int(r["fans"]),
            "fan_eff": r["fan_eff"], "role": "；".join(role) or "常规",
        })
    return rows


def build_actions(data, funnel, attr, strategy):
    """生成可执行优化清单：问题 -> 数据依据 -> 具体改法 -> 预期。"""
    e, f = funnel["e"], funnel["f"]
    cur_eff = (f / e) if e else 0
    # 预期用保守 3 倍提升估算，标注为参考
    tgt_eff = cur_eff * 3
    exp_fans = int(e * tgt_eff) if e else 0

    actions = []
    # 1. 关注引导（最高杠杆）
    actions.append({
        "pri": "P0",
        "problem": "关注引导缺失——漏斗末端断了",
        "evidence": f"总曝光 {e:,} 仅换 {f} 粉，曝光→涨粉率 {cur_eff*100:.3f}%；{funnel['zero_fan']}/{len(data)} 篇零涨粉",
        "fix": "① 每篇文末加 1 句关注钩子（例：『关注我，<你的核心更新内容>每天更』）；② 主页简介写清定位+关注理由；③ 置顶一条『新人必看』",
        "expected": f"若曝光→涨粉率提到当前 3 倍（{tgt_eff*100:.3f}%），同等流量月涨粉约 {exp_fans}（估算，非承诺）",
    })
    # 2. 高质量被埋没
    if attr["starved_n"] > 0:
        actions.append({
            "pri": "P1",
            "problem": f"高质量内容被埋没（{attr['starved_n']} 篇互动率高但曝光<100）",
            "evidence": "例：" + "；".join(attr["starved_titles"][:2]),
            "fix": "用高曝光内容线做标题钩子，正文/合集里嵌高质量内容交叉引流；或在流量笔记评论区/文末引导看你的深度内容",
            "expected": "把已有高质量内容从『无人看见』变『被流量线带出来』，互动转化复用",
        })
    # 3. 红榜共性复用
    red_tag = max(attr["red_tags"], key=attr["red_tags"].get) if attr["red_tags"] else "—"
    actions.append({
        "pri": "P1",
        "problem": f"红榜全是「{red_tag}」线——已有验证的流量打法没放大",
        "evidence": "红榜 TOP5：" + "；".join(attr["red_titles"][:3]),
        "fix": f"把「{red_tag}」线的标题套路做成模板，批量复用；其余线借这条线的流量结构",
        "expected": "复制已验证的曝光结构，降低单篇试错成本",
    })
    # 4. 黑榜零流量
    nb = attr["buckets"].get("没流量（发布/推荐问题）", 0)
    if nb > 0:
        actions.append({
            "pri": "P2",
            "problem": f"{nb} 篇几乎零曝光——没进推荐流",
            "evidence": "标题例：" + "；".join(attr["black_titles"][:2]),
            "fix": "检查发布时间（工作日晚 18–22 点、周末午间流量高）；首小时靠粉丝/社群撬动初始互动，触发推荐；标题前 15 字含搜索热词",
            "expected": "先把『发得出去』解决，再谈转化",
        })
    # 5. 内容线分工
    if strategy:
        qlines = [s for s in strategy if "质量引擎" in s["role"]]
        tlines = [s for s in strategy if "流量引擎" in s["role"]]
        if qlines and tlines:
            actions.append({
                "pri": "P2",
                "problem": "内容线没分工，质量线被流量线淹没",
                "evidence": f"质量引擎={qlines[0]['tag']}（互动最高但曝光少）；流量引擎={tlines[0]['tag']}（曝光高）",
                "fix": f"用「{tlines[0]['tag']}」做流量钩子吸新，用「{qlines[0]['tag']}」做粘性沉淀粉丝；两条线在标题/合集里互相导流",
                "expected": "流量与质量分开经营，既不缺曝光也不缺粘性",
            })
    return actions


def module_analysis(data):
    funnel = module_funnel(data)
    attr = module_attr(data)
    strategy = module_strategy(data)
    actions = build_actions(data, funnel, attr, strategy)
    return {"funnel": funnel, "attr": attr, "strategy": strategy, "actions": actions}


# ----------------------------------------------------------------------------
# HTML 渲染（纯 CSS + 内联条形，零外部依赖）
# ----------------------------------------------------------------------------
def tag_css():
    rules = ""
    for i, t in enumerate(all_tags()):
        bg, fg = TAG_PALETTE[i % len(TAG_PALETTE)]
        rules += f".tag.{t}{{background:{bg};color:{fg}}}\n"
    return rules


def build_css():
    return """
:root{--bg:#0f1115;--card:#1a1d24;--ink:#e8eaf0;--muted:#9aa0ad;--red:#ff3b4e;--green:#2ecc71;--amber:#ffb020;--blue:#4aa3ff;--line:#2a2e38;}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--ink);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;line-height:1.5;padding:16px;max-width:760px;margin:0 auto}
h1{font-size:22px;margin-bottom:4px}
h2{font-size:17px;margin:22px 0 10px;padding-left:10px;border-left:4px solid var(--red)}
.sub{color:var(--muted);font-size:13px;margin-bottom:6px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px;margin:10px 0}
.kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin:12px 0}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px;text-align:center}
.kpi .v{font-size:20px;font-weight:700}
.kpi .l{font-size:11px;color:var(--muted);margin-top:2px}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th,td{padding:7px 6px;text-align:left;border-bottom:1px solid var(--line)}
th{color:var(--muted);font-weight:600;font-size:11px}
.tag{display:inline-block;padding:1px 7px;border-radius:6px;font-size:11px;font-weight:600}
""" + tag_css() + """
.bar{height:8px;border-radius:5px;background:linear-gradient(90deg,var(--red),var(--amber));min-width:2px}
.barwrap{background:var(--line);border-radius:5px;height:8px;overflow:hidden;flex:1}
.row{display:flex;align-items:center;gap:8px;margin:5px 0;font-size:12.5px}
.row .name{width:64px;color:var(--muted);flex-shrink:0}
.badge{font-size:11px;padding:1px 6px;border-radius:5px}
.badge.good{background:rgba(46,204,113,.18);color:var(--green)}
.badge.bad{background:rgba(255,59,78,.18);color:var(--red)}
.note{font-size:12px;color:var(--muted);margin:6px 0}
.foot{font-size:11px;color:var(--muted);margin-top:26px;text-align:center;border-top:1px solid var(--line);padding-top:12px}
.pos{color:var(--red);font-weight:700}.neg{color:var(--green)}
"""


def esc(s):
    return html.escape(str(s))


def pct(x, d=1):
    return f"{x*100:.{d}f}%"


def kpi_card(v, l):
    return f'<div class="kpi"><div class="v">{v}</div><div class="l">{l}</div></div>'


def build_html(data, by_source):
    ov = overview(data)
    red, black = module_redblack(data)
    bomb_grp, type_dist, len_corr = module_bomb(data)
    tag_cmp = module_tag_compare(data)
    health = module_health(data)
    analysis = module_analysis(data)
    css = build_css()

    # 时间范围
    if ov["date_min"] and ov["date_max"]:
        span = f'{ov["date_min"]:%Y-%m-%d} ~ {ov["date_max"]:%Y-%m-%d}'
    else:
        span = "时间未知"

    # ---- 红黑榜行 ----
    def rbar(r):
        mx = max(data["score"].max(), 1)
        w = max(2, int(r["score"] / mx * 100))
        return f'<div class="barwrap"><div class="bar" style="width:{w}%"></div></div>'

    red_rows = ""
    for _, r in red.iterrows():
        red_rows += (f'<tr><td class="tag {esc(r["tag"])}">{esc(tagname(r["tag"]))}</td>'
                     f'<td>{esc(r["title"])[:22]}</td>'
                     f'<td>{int(r["exposure"])}</td><td>{int(r["fans_gain"])}</td>'
                     f'<td>{pct(r["interaction_rate"])}</td></tr>')
    black_rows = ""
    for _, r in black.iterrows():
        black_rows += (f'<tr><td class="tag {esc(r["tag"])}">{esc(tagname(r["tag"]))}</td>'
                       f'<td>{esc(r["title"])[:22]}</td>'
                       f'<td>{int(r["exposure"])}</td><td>{int(r["fans_gain"])}</td>'
                       f'<td>{pct(r["interaction_rate"])}</td></tr>')

    # ---- 内容线对比条形 ----
    max_exp = max(tag_cmp["avg_exposure"].max(), 1)
    max_ir = max(tag_cmp["avg_ir"].max(), 0.0001)
    max_fe = max(tag_cmp["avg_fan_eff"].max(), 0.0001)
    cmp_rows = ""
    for _, r in tag_cmp.iterrows():
        we = max(2, int(r["avg_exposure"] / max_exp * 100))
        wi = max(2, int(r["avg_ir"] / max_ir * 100))
        wf = max(2, int(r["avg_fan_eff"] / max_fe * 100))
        cmp_rows += f'''
        <div style="margin:10px 0">
          <div class="row"><span class="name">{esc(tagname(r["tag"]))}</span>
            <span style="width:90px;color:var(--muted)">曝光均值</span>
            <div class="barwrap"><div class="bar" style="width:{we}%"></div></div>
            <span style="width:54px;text-align:right">{int(r["avg_exposure"])}</span></div>
          <div class="row"><span class="name"></span>
            <span style="width:90px;color:var(--muted)">互动率</span>
            <div class="barwrap"><div class="bar" style="width:{wi}%"></div></div>
            <span style="width:54px;text-align:right">{pct(r["avg_ir"])}</span></div>
          <div class="row"><span class="name"></span>
            <span style="width:90px;color:var(--muted)">涨粉效率</span>
            <div class="barwrap"><div class="bar" style="width:{wf}%"></div></div>
            <span style="width:54px;text-align:right">{pct(r["avg_fan_eff"],2)}</span></div>
        </div>'''

    # ---- 内容线均值表 ----
    bomb_rows = ""
    for _, r in bomb_grp.iterrows():
        bomb_rows += (f'<tr><td class="tag {esc(r["tag"])}">{esc(tagname(r["tag"]))}</td>'
                      f'<td>{int(r["n"])}</td><td>{int(r["avg_exposure"])}</td>'
                      f'<td>{pct(r["avg_ir"])}</td><td>{int(r["avg_fans"])}</td>'
                      f'<td>{pct(r["avg_fan_eff"],2)}</td></tr>')

    type_rows = "".join(
        f'<tr><td>{esc(r["type"])}</td><td>{int(r["n"])}</td><td>{pct(r["n"]/len(data))}</td></tr>'
        for _, r in type_dist.iterrows())

    # ---- 流量诊断 ----
    leaked_rows = ""
    for _, r in health["leaked"].iterrows():
        leaked_rows += (f'<tr><td>{esc(r["title"])[:24]}</td>'
                        f'<td>{int(r["exposure"])}</td><td>{pct(r["interaction_rate"])}</td>'
                        f'<td class="neg">低互动</td></tr>')

    if pd.notna(len_corr):
        corr_txt = f"标题长度与互动的相关系数 {len_corr:.2f}（接近0表示长度不是关键，选题/封面更决定表现）"
    else:
        corr_txt = "标题长度与互动相关性不显著（样本或数值单一，暂无法判断）"

    bysrc = ""
    if by_source and ov["sources"] > 1:
        g = data.groupby("source").agg(
            n=("title", "count"), exp=("exposure", "sum"),
            fans=("fans_gain", "sum"), ir=("interaction_rate", "mean")).reset_index()
        bysrc_rows = "".join(
            f'<tr><td>{esc(r["source"])}</td><td>{int(r["n"])}</td><td>{int(r["exp"])}</td>'
            f'<td>{int(r["fans"])}</td><td>{pct(r["ir"])}</td></tr>' for _, r in g.iterrows())
        bysrc = f'''
        <h2>多账号 / 多平台对比</h2>
        <div class="card"><table>
          <tr><th>账号/平台(文件)</th><th>笔记</th><th>总曝光</th><th>总涨粉</th><th>互动率</th></tr>
          {bysrc_rows}
        </table>
        <div class="note">把不同平台的导出文件命名好（如 头条_账号名.xlsx）一起传入，即可横向对比。</div></div>'''

    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")

    # ---- 分析层：漏斗 / 归因 / 策略 / 优化清单 ----
    fn = analysis["funnel"]
    at = analysis["attr"]
    st = analysis["strategy"]
    rates = fn["rates"]

    funnel_rows = "".join(
        f'<div class="row"><span class="name">{k}</span>'
        f'<span style="width:80px;text-align:right;color:var(--ink);font-weight:600">{v*100:.2f}%</span>'
        f'<div class="barwrap"><div class="bar" style="width:{max(2,int(v*100))}%"></div></div></div>'
        for k, v in rates.items())

    # 漏点标红
    leak_txt = (f'最大漏点在 <span class="pos">【{fn["leak"]}】</span>：'
                f'你有流量（曝光 {fn["e"]:,}）但只换回 {fn["f"]} 个粉，'
                f'{fn["zero_fan"]}/{ov["notes"]} 篇笔记涨粉为 0。'
                f'问题不在"让人点进来"，在"让人留下来关注"。')

    bucket_rows = "".join(
        f'<div class="note">• {esc(k)}：{v} 篇</div>' for k, v in at["buckets"].items())

    strat_rows = ""
    for s in st:
        strat_rows += (f'<tr><td class="tag {esc(s["tag"])}">{esc(tagname(s["tag"]))}</td>'
                       f'<td>{s["n"]}</td><td>{int(s["exp"]):,}</td>'
                       f'<td>{pct(s["avg_ir"])}</td><td>{s["fans"]}</td>'
                       f'<td style="color:var(--amber)">{esc(s["role"])}</td></tr>')

    starved_txt = ("被埋没的高质量内容（互动率高但曝光<100）："
                   + "；".join(esc(t) for t in at["starved_titles"][:3])) if at["starved_n"] else "暂无明显高质量低曝光内容"

    action_cards = ""
    for a in analysis["actions"]:
        action_cards += f'''
        <div class="card" style="border-left:4px solid var(--amber)">
          <div class="badge bad" style="background:rgba(255,176,32,.18);color:var(--amber)">{esc(a["pri"])}</div>
          <div style="font-weight:700;margin:6px 0 4px">🔧 {esc(a["problem"])}</div>
          <div class="note">📊 数据依据：{esc(a["evidence"])}</div>
          <div class="note">✅ 具体改法：{esc(a["fix"])}</div>
          <div class="note" style="color:var(--muted)">📈 预期：{esc(a["expected"])}</div>
        </div>'''

    analysis_html = f'''
<h2>⑤ 诊断结论（不是摆数，是找病根）</h2>
<div class="card">
  <div style="font-weight:700;margin-bottom:6px">转化漏斗</div>
  {funnel_rows}
  <div class="note" style="margin-top:8px;color:var(--ink)">{leak_txt}</div>
</div>
<div class="card">
  <div style="font-weight:700;margin-bottom:6px">红黑榜归因</div>
  <div class="note">🔴 红榜 TOP5 共性：{esc("、".join(at["red_titles"][:3]))}（主要集中在 {esc("、".join(at["red_tags"].keys()))} 线）</div>
  <div class="note">⚫ 黑榜分桶：</div>
  {bucket_rows}
  <div class="note">💎 {esc(starved_txt)}</div>
</div>
<div class="card">
  <div style="font-weight:700;margin-bottom:6px">内容线分工建议</div>
  <table><tr><th>内容线</th><th>篇数</th><th>总曝光</th><th>互动率</th><th>涨粉</th><th>角色定位</th></tr>{strat_rows}</table>
  <div class="note">流量引擎保曝光、质量引擎做粘性，两条线在标题/合集里互相导流。</div>
</div>

<h2>⑥ 你的优化方案（照着做，不空谈）</h2>
{action_cards}
<div class="note" style="color:var(--muted)">⚠️ 预期数字为基于当前数据的保守估算（3 倍提升假设），非涨粉承诺。所有建议来自你后台真实数据，不编不猜。</div>
'''

    out = f'''<!DOCTYPE html><html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>账号诊断看板</title><style>{css}</style></head><body>
<h1>📊 账号诊断看板</h1>
<div class="sub">数据周期：{span} · 共 {ov["notes"]} 篇笔记 · 来源 {ov["sources"]} 份官方导出 · 生成 {now}</div>
<div class="sub" style="color:var(--amber)">⚠️ 数据来自你自己的小红书后台官方导出，零爬取、零第三方服务，全程本地不上传。</div>

<div class="kpis">
{kpi_card(f'{ov["exposure"]:,}', "总曝光")}
{kpi_card(f'{ov["views"]:,}', "总观看")}
{kpi_card(f'{ov["fans"]:,}', "总涨粉")}
{kpi_card(pct(ov["ir"]), "平均互动率")}
{kpi_card(f'{ov["avg_ctr"]:.2f}', "平均封面点击率")}
{kpi_card(f'{ov["interactions"]:,}', "总互动")}
</div>

<h2>① 单篇红黑榜</h2>
<div class="card"><div class="badge good">红榜 TOP5</div>
<table><tr><th>内容线</th><th>标题</th><th>曝光</th><th>涨粉</th><th>互动率</th></tr>{red_rows}</table></div>
<div class="card"><div class="badge bad">黑榜 BOTTOM5</div>
<table><tr><th>内容线</th><th>标题</th><th>曝光</th><th>涨粉</th><th>互动率</th></tr>{black_rows}</table>
<div class="note">综合得分 = 涨粉×权重 + 互动数 + 曝光×小权重（可在 config.json 里调），权重透明可解释。</div></div>

<h2>② 爆款规律</h2>
<div class="card">
<table><tr><th>内容线</th><th>篇数</th><th>曝光均值</th><th>互动率</th><th>涨粉均值</th><th>涨粉效率</th></tr>{bomb_rows}</table>
<div class="note">体裁分布：</div>
<table><tr><th>体裁</th><th>篇数</th><th>占比</th></tr>{type_rows}</table>
<div class="note">{corr_txt}</div>
</div>

<h2>③ 内容线对比</h2>
<div class="card">{cmp_rows}
<div class="note">条形越长代表该指标在各内容线里越高。谁在「涨粉效率」上领跑，下次就多压哪条线。</div></div>

<h2>④ 流量健康诊断</h2>
<div class="card">
<div class="row"><span class="name">曝光均值</span><span>{int(health["exp_mean"])}</span></div>
<div class="row"><span class="name">互动率均值</span><span>{pct(health["ir_mean"])}</span></div>
<div class="row"><span class="name">封面点击率</span><span>{health["ctr_mean"]:.2f}</span></div>
<div class="badge bad">零曝光笔记 {health["zero_exp_n"]} 篇</div>
<div class="badge bad">低点击率笔记 {health["low_ctr_n"]} 篇</div>
<div class="note">「高曝光低互动」清单（流量进了但没接住，问题在内容转化）：</div>
<table><tr><th>标题</th><th>曝光</th><th>互动率</th><th>诊断</th></tr>{leaked_rows or '<tr><td colspan=4>无明显漏桶</td></tr>'}</table>
</div>

{analysis_html}

{bysrc}

<div class="foot">小红书账号诊断器 · 数据 100% 来自你自己后台的官方导出，不外泄、不爬取、不上传</div>
</body></html>'''
    return out


COUNT_COLS = ["exposure", "views", "likes", "comments", "collects",
              "fans_gain", "shares", "avg_watch", "danmu"]


def dedup(data):
    """同一篇笔记会在不同日期导出中重复出现。三类字段三种取法（实战修正口径）：

    - **计数类**（曝光/观看/赞/评/藏/涨粉/转发/时长/弹幕）→ 取 `max`：这些只增不减。
    - **封面点击率** → 取**最新快照**的值。注意：不能对它取 `max`——
      点击率是**比率**、会随曝光增长而回落 ⇒ 旧快照的虚高值会被留下来。
    - **派生指标**（互动数/互动率/涨粉效率/score）→ 用合并后的计数**重算**，不取 max。

    依赖 `load_file` 写入的 `snapshot_ts`（文件 mtime）；缺列时退回旧行为。
    """
    if "snapshot_ts" not in data.columns:
        metric_cols = NUM_COLS + ["interactions", "interaction_rate", "fan_eff", "score"]
        out = data.groupby("title", as_index=False).agg({c: "max" for c in metric_cols})
        meta = data.drop_duplicates("title").set_index("title")
        for c in ["publish_time", "type", "tag", "source"]:
            out[c] = out["title"].map(meta[c])
        return out.sort_values("score", ascending=False).reset_index(drop=True)

    out = data.groupby("title", as_index=False).agg({c: "max" for c in COUNT_COLS})
    newest = (data.sort_values("snapshot_ts")
                  .drop_duplicates("title", keep="last")
                  .set_index("title"))
    for c in ["cover_ctr", "publish_time", "type", "tag", "source"]:
        out[c] = out["title"].map(newest[c])
    out["interactions"] = out["likes"] + out["comments"] + out["collects"] + out["shares"]
    out["interaction_rate"] = (out["interactions"] / out["views"].replace(0, pd.NA)).fillna(0)
    out["fan_eff"] = (out["fans_gain"] / out["exposure"].replace(0, pd.NA)).fillna(0)
    out["score"] = out.apply(score_row, axis=1)
    return out.sort_values("score", ascending=False).reset_index(drop=True)


def object_overview(data, name, top_tags):
    """把一个对象的明细数据聚合成账号级概览（用于并排对比）。"""
    ov = overview(data)
    tags = data["tag"].value_counts().to_dict()
    n = max(ov["notes"], 1)
    d = {
        "name": name,
        "notes": ov["notes"],
        "exposure": ov["exposure"],
        "views": ov["views"],
        "fans": ov["fans"],
        "ir": ov["ir"],
        "ctr": ov["avg_ctr"],
    }
    for t in top_tags:
        d[f"tag_{t}"] = tags.get(t, 0) / n
    return d


def build_compare_html(objects, top_tags):
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    cols = "".join(f"<th>{esc(o['name'])}</th>" for o in objects)

    def row(label, key, fmt, better="high"):
        vals = [o[key] for o in objects]
        best = max(vals) if better == "high" else min(vals)
        cells = ""
        for o in objects:
            v = o[key]
            cls = ' class="pos"' if (len(objects) > 1 and v == best) else ""
            cells += f"<td{cls}>{fmt(v)}</td>"
        return f"<tr><td>{label}</td>{cells}</tr>"

    rows = ""
    rows += row("笔记数", "notes", lambda x: f"{x:,}", "high")
    rows += row("总曝光", "exposure", lambda x: f"{x:,}", "high")
    rows += row("总观看", "views", lambda x: f"{x:,}", "high")
    rows += row("总涨粉", "fans", lambda x: f"{x:,}", "high")
    rows += row("平均互动率", "ir", lambda x: pct(x), "high")
    rows += row("平均封面点击率", "ctr", lambda x: f"{x:.2f}", "high")
    for t in top_tags:
        rows += row(f"{tagname(t)}占比", f"tag_{t}", lambda x: pct(x, 0), "high")

    diag = []
    if len(objects) >= 2:
        base = objects[0]
        for o in objects[1:]:
            if o["fans"] > base["fans"]:
                diag.append(f'{esc(o["name"])} 涨粉 {o["fans"]} &gt; {esc(base["name"])} {base["fans"]}，差距在转化端（关注引导）')
            if o["ir"] > base["ir"]:
                diag.append(f'{esc(o["name"])} 互动率 {pct(o["ir"])} &gt; {pct(base["ir"])}，内容质量更抓人')
            if o["ctr"] > base["ctr"]:
                diag.append(f'{esc(o["name"])} 封面点击率更高，第一道门（封面划不划得动）做得更好')
    diag_html = "".join(f'<div class="note">• {d}</div>' for d in diag) or \
        '<div class="note">仅单一对象，无可对比差距。传入多个文件即可并排。</div>'

    return f'''<!DOCTYPE html><html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>账号对比看板</title><style>{build_css()}</style></head><body>
<h1>⚔️ 自己 vs 对标 对比看板</h1>
<div class="sub">生成 {now} · 共 {len(objects)} 个对比对象 · 数据来自官方导出/真实拆解，零虚构</div>
<div class="card"><table><tr><th>指标</th>{cols}</tr>{rows}</table>
<div class="note">标红 = 该指标在对比对象中领先。封面点击率/互动率越高越好。</div></div>
<h2>差距诊断</h2><div class="card">{diag_html}</div>
<div class="foot">对比模式 · 把对标账号的公开指标填入同格式表格，与自己的后台导出一并传入即可并排对比。</div>
</body></html>'''


def find_downloads_files():
    files = []
    if os.path.isdir(DOWNLOADS):
        for n in sorted(os.listdir(DOWNLOADS)):
            if n.endswith(".xlsx") and "笔记" in n:
                files.append(os.path.join(DOWNLOADS, n))
    return files


def main():
    args = sys.argv[1:]
    by_source = False
    config_path = None
    out_dir = None
    paths = []
    i = 0
    while i < len(args):
        a = args[i]
        if a == "--by-source":
            by_source = True
        elif a == "--config":
            i += 1
            config_path = args[i] if i < len(args) else None
        elif a == "--out":
            i += 1
            out_dir = args[i] if i < len(args) else None
        elif a.startswith("--"):
            pass
        else:
            paths.append(a)
        i += 1

    load_config(config_path)

    out_dir = out_dir or os.path.join(WORKDIR, "outputs")
    os.makedirs(out_dir, exist_ok=True)

    if not paths:
        paths = find_downloads_files()
        if not paths:
            print("未找到笔记列表明细表.xlsx：请把文件放到 ~/Downloads，或显式传入路径。")
            sys.exit(1)

    if "--compare" in args:
        objs = []
        for p in paths:
            d = load_file(p)
            d = dedup(d)
            nm = os.path.splitext(os.path.basename(p))[0]
            objs.append(object_overview(d, nm, all_tags()))
        out = build_compare_html(objs, all_tags())
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
        outp = os.path.join(out_dir, f"账号对比看板_{stamp}.html")
        with open(outp, "w", encoding="utf-8") as f:
            f.write(out)
        print("对比看板已生成：", outp)
        return

    print(f"读取 {len(paths)} 个文件：")
    for p in paths:
        print("  -", p)

    data, _ = load_all(paths, by_source)
    print(f"合并后 {len(data)} 篇（去重前）")
    before = len(data)
    data = dedup(data)
    print(f"去重后 {len(data)} 篇唯一笔记（移除 {before - len(data)} 条重复快照）")
    print("内容线分布：", data["tag"].value_counts().to_dict())

    html_out = build_html(data, by_source)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    out_path = os.path.join(out_dir, f"账号诊断看板_{stamp}.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_out)
    print("看板已生成：", out_path)


if __name__ == "__main__":
    main()
