# -*- coding: utf-8 -*-
"""
印度尼西亚电子烟市场调研数据看板
结构参考 vape_ch_dashboard_v3.py（瑞士版），按印尼数据集独立适配。
数据文件：印尼热销产品榜单-20261030xlsx.xlsx
Sheet 1: popular sell (clear)   热销榜 120 条 / 5 个零售网站
Sheet 2: popular brand eliquid  Foom、RELX、OXVA 在售烟油库 164 条
运行：
pip install streamlit pandas plotly openpyxl
streamlit run id_dashboard_v1.py
"""
import os, glob, re, inspect
from collections import Counter
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="印尼电子烟市场调研数据看板", page_icon="🇮🇩", layout="wide", initial_sidebar_state="expanded")

# ------------------------------ 常量 ------------------------------
IDR_PER_USD = 16650                 # 用户调研概览使用的固定折算汇率，不是实时汇率
CNY_PER_IDR = 0.00037                # 用户指定：1 印尼盾 = 0.00037 人民币
CNY_PER_USD = 6.72                   # 用户指定：1 美元 = 6.72 人民币
TOPN = 20                           # 热销榜按前 20 名计算相对得分
FLAVOR_MIN_SUPPORT = 3             # 洞察默认只突出至少重复 3 次的口味/元素
FLAVOR_TAG_ALIASES = {
    "ice": "冰感", "mint": "薄荷", "menthol": "薄荷", "cool": "冰感",
    "cooling": "冰感", "koolada": "冰感", "strawberry": "草莓", "tea": "茶",
    "gummy": "软糖", "bubblegum": "泡泡糖", "soda": "汽水", "milk": "奶",
    "cream": "奶油", "custard": "卡仕达", "tobacco": "烟草", "mango": "芒果",
    "grape": "葡萄", "lychee": "荔枝", "watermelon": "西瓜", "peach": "桃",
    "apple": "苹果", "banana": "香蕉", "coffee": "咖啡", "caramel": "焦糖"
}

FLAVOR_CATEGORY_ALIASES = {
    "fruit": "水果", "fruits": "水果", "fruity": "水果", "dessert": "甜点",
    "sweet": "甜点", "drink": "饮料", "beverage": "饮料", "candy": "糖果",
    "mint": "薄荷", "menthol": "薄荷", "tobacco": "烟草"
}
DATA_DATE = "2026-10-30"
REGULATORY_CHECK_DATE = "2026-11-02"
SHEET_POP = "popular sell (clear)"
SHEET_BRAND = "popular brand eliquid"
FORM_EL = "瓶装烟油 E-liquid"
FORM_DP = "一次性电子烟 Disposable"
ICE_TAGS = {"ice", "mint", "menthol", "cool", "cooling", "koolada"}
T = "plotly_white"

COLORS = {
    "primary": "#166534", "secondary": "#0F766E", "gold": "#D6A21E", "fruit": "#F9735B",
    "drink": "#0891B2", "dessert": "#B45309", "candy": "#A855F7", "menthol": "#14B8A6",
    "tobacco": "#78350F", "other": "#94A3B8", "ice": "#38BDF8"
}
CAT_COLOR_MAP = {
    "水果": COLORS["fruit"], "饮料": COLORS["drink"], "甜点": COLORS["dessert"],
    "糖果": COLORS["candy"], "薄荷": COLORS["menthol"], "烟草": COLORS["tobacco"], "其他": COLORS["other"]
}
HEAT_SCALE = [[0, "#ECFDF5"], [0.55, "#86C9A4"], [1, COLORS["primary"]]]

# ------------------------------ 样式与布局 ------------------------------
st.markdown("""
<style>
.metric-card {
    background: linear-gradient(135deg, #166534 0%, #0F766E 100%);
    padding: 15px 17px; border-radius: 11px; color: white; text-align: center;
    margin: 3px 0; min-height: 96px; box-sizing: border-box; display: flex; flex-direction: column; justify-content: center;
}
.metric-card.alt { background: linear-gradient(135deg, #475569 0%, #1e293b 100%); }
.metric-card .value { font-size: 1.75rem; font-weight: 700; line-height: 1.2; }
.metric-card .label { font-size: 0.77rem; opacity: 0.92; margin-top: 5px; }
.price-note { color: #cbd5e1; font-size: 0.72rem; line-height: 1.25; margin-top: 3px; min-height: 1.1em; }
.element-selector-spacer { height: 68px; }

.section-title {
    font-size: 1.03rem; font-weight: 650; color: #166534;
    margin: 17px 0 9px; padding-left: 10px; border-left: 4px solid #0F766E;
}

.insight-box, .warning-box, .danger-box, .action-box, .summary-card {
    padding: 14px 18px; border-radius: 8px; margin: 10px 0;
    line-height: 1.7; font-size: 0.92rem; box-shadow: 0 2px 4px rgba(0,0,0,0.03);
}
.insight-box b, .summary-card b, .warning-box b, .danger-box b, .action-box b {
    display: block; margin-bottom: 6px; font-weight: 700; font-size: 0.95rem;
}
.insight-box, .summary-card { background: #ECFDF5; border-left: 4px solid #166534; }
.aligned-insight { min-height: 94px; }
.warning-box { background: #fff7ed; border-left: 4px solid #d97706; }
.danger-box { background: #FFF1F2; border-left: 4px solid #BE123C; }
.action-box { background: #f0fdf4; border-left: 4px solid #16a34a; }

.reg-card {
    background: #f8fafc; border-left: 4px solid #94a3b8;
    padding: 12px 16px; border-radius: 8px; margin: 8px 0; line-height: 1.65; font-size: 0.89rem;
}
.reg-card.restrict { border-left-color: #d97706; }
.reg-card.proposal { border-left-color: #991b1b; }
.reg-card.active { border-left-color: #16a34a; }
.reg-card b { display: block; margin-bottom: 4px; color: #0f172a; }

.tag-chip {
    display: inline-block; background: #CCFBF1; color: #115E59;
    padding: 2px 9px; border-radius: 999px; font-size: 0.75rem; margin: 2px;
}
</style>
""", unsafe_allow_html=True)

# ------------------------------ 通用函数 ------------------------------
_CHART_SEQ = 0
def next_key(prefix="chart"):
    global _CHART_SEQ
    _CHART_SEQ += 1
    return f"{prefix}{_CHART_SEQ}"

def params(fn):
    try: return inspect.signature(fn).parameters
    except (TypeError, ValueError): return {}

def width_kwargs(fn):
    p = params(fn); w = p.get("width")
    if w is not None and w.default == "stretch": return {"width": "stretch"}
    if "use_container_width" in p: return {"use_container_width": True}
    return {}

def show_chart(fig):
    kw = width_kwargs(st.plotly_chart)
    if "key" in params(st.plotly_chart): kw["key"] = next_key()
    st.plotly_chart(fig, **kw)

def show_df(df, **kwargs):
    kw = width_kwargs(st.dataframe)
    if "key" in params(st.dataframe): kw["key"] = next_key("df")
    st.dataframe(df, hide_index=True, **kw, **kwargs)

def metric(col, label, value, alt=False):
    cls = "metric-card alt" if alt else "metric-card"
    col.markdown(f'<div class="{cls}"><div class="value">{value}</div><div class="label">{label}</div></div>', unsafe_allow_html=True)

def price_metric(col, label, value, alt=False):
    """Render a price card with IDR as the main value and CNY as a muted note."""
    cls = "metric-card alt" if alt else "metric-card"
    if pd.notna(value):
        main_text = f"Rp {value:,.0f}"
        cny_text = f"约 ¥{value * CNY_PER_IDR:,.2f}"
    else:
        main_text = "—"
        cny_text = "人民币约价不可用"
    col.markdown(
        f'<div class="{cls}"><div class="value">{main_text}</div>'
        f'<div class="price-note">{cny_text}</div><div class="label">{label}</div></div>',
        unsafe_allow_html=True,
    )

def section(title): 
    st.markdown(f'<p class="section-title">{title}</p>', unsafe_allow_html=True)

def pct(n, d): return f"{n/d:.1%}" if d else "—"

def clean_text(x): 
    return re.sub(r"\s+", " ", str(x).replace("\xa0", " ").strip()) if pd.notna(x) else ""

def norm_brand(x):
    s = clean_text(x)
    m = {"foom": "Foom", "relx": "RELX", "oxva": "OXVA", "cmmg brew": "CMMG Brew", "dji(dunia juice indonesia)": "DJI (Dunia Juice Indonesia)"}
    return m.get(s.lower(), s)

def split_tags_raw(v):
    if pd.isna(v): return []
    return [clean_text(x) for x in re.split(r"[,，;/|、]+", str(v))
            if clean_text(x) and clean_text(x).lower() not in ("无", "nan", "none")]


def split_tags(v):
    # Lower-case normalized values are used only for matching/counting.
    return [x.lower() for x in split_tags_raw(v)]

def display_tag(tag):
    return FLAVOR_TAG_ALIASES.get(str(tag).lower().strip(), str(tag).lower().strip())

def display_category(cat):
    key = str(cat).lower().strip()
    return FLAVOR_CATEGORY_ALIASES.get(key, str(cat).strip())

def nums_in(v):
    if pd.isna(v): return []
    return [float(x) for x in re.findall(r"\d+(?:\.\d+)?", str(v))]

def parse_idr(v):
    if pd.isna(v): return None
    if isinstance(v, (int, float)): return float(v)
    s = clean_text(v).upper().replace("IDR", " ").replace("RP", " ").replace("RUPIAH", " ").replace("  ", " ")
    s = s.replace(",", "").replace(".", "")
    try: return float(s)
    except ValueError: return None

def top_n_ties(s, n=10, include_singletons=False):
    """Return the top values, retaining ties at the cutoff.

    When singleton values are excluded, filter them before applying the top-N
    cutoff so small filtered samples do not accidentally reintroduce count=1
    items.
    """
    s = s.dropna().sort_values(ascending=False, kind="mergesort")
    if not include_singletons:
        s = s[s > 1]
    if len(s) <= n:
        return s
    cut = s.iloc[n - 1]
    return s[s >= cut]

def top_counter(counter, n=10, include_singletons=False):
    return top_n_ties(pd.Series(dict(counter)), n, include_singletons=include_singletons)

def counts_ge2(values):
    """Return only values appearing at least twice, with stable descending order."""
    counts = values.dropna().value_counts()
    return counts[counts >= 2].sort_values(ascending=False, kind="mergesort")

def top_summary(series, n=3, unit="条"):
    v = top_n_ties(series.value_counts(), n, include_singletons=True)
    return "、".join(f"{k}（{int(x)}{unit}）" for k, x in v.items()) if len(v) else "暂无可用数据"

def tie_note(k, n):
    if k > n: st.caption(f"第 {n} 名出现并列，已一并显示并列项（共 {k} 项）。")

def vc_df(s, name, count="记录数"):
    v = s.fillna("未标注").value_counts()
    return pd.DataFrame({name: v.index.tolist(), count: v.values.tolist()})

def flavor_complexity(n):
    if n == 0: return "无标签"
    if n == 1: return "单一元素"
    if n == 2: return "双元素组合"
    return "三种及以上元素"

def format_idr(v):
    return f"Rp {v:,.0f}" if pd.notna(v) else "—"

def format_idr_cny(v):
    return f"Rp {v:,.0f}（约 ¥{v * CNY_PER_IDR:,.2f}）" if pd.notna(v) else "—"

def format_usd_cny(v):
    return f"US${v:,.2f}（约 ¥{v * CNY_PER_USD:,.2f}）" if pd.notna(v) else "—"

# ------------------------------ 数据文件 ------------------------------
_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_XLSX = os.path.join(_DIR, "印尼热销产品榜单-20261030xlsx.xlsx")

def resolve_excel():
    if os.path.exists(DEFAULT_XLSX): return DEFAULT_XLSX
    c = sorted(glob.glob(os.path.join(_DIR, "印尼*.xlsx")))
    return c[-1] if c else None

@st.cache_data
def load_data(path, mtime):
    pop = pd.read_excel(path, sheet_name=SHEET_POP)
    brand = pd.read_excel(path, sheet_name=SHEET_BRAND)
    for d in (pop, brand):
        d.columns = d.columns.astype(str).str.strip()
        for c in d.select_dtypes(include="object").columns: 
            d[c] = d[c].map(lambda x: clean_text(x) if pd.notna(x) else x)
            
    req_pop = ["网站名称", "产品类型", "排名", "品牌", "口味名称", "口味标签", "分类", "价格", "设备类别"]
    req_brand = ["网站名称", "品牌", "口味名称", "口味标签", "分类", "价格", "设备类别"]
    for c in req_pop:
        if c not in pop: raise ValueError(f"热销榜 Sheet 缺少字段：{c}")
    for c in req_brand:
        if c not in brand: raise ValueError(f"品牌烟油 Sheet 缺少字段：{c}")

    def enrich(d, ranked):
        d = d.copy()
        for c, default in [("网站名称", "未标注"), ("产品类型", "未标注"), ("品牌来源国", "未标注"), 
                           ("是否本地代理/本土品牌", "不确定"), ("含冰/薄荷", "不确定"), 
                           ("含烟草", "不确定"), ("分类", "其他")]:
            if c not in d: d[c] = default
            d[c] = d[c].fillna(default).astype(str).str.strip()
            
        d["品牌原始"] = d["品牌"].fillna("未标注").astype(str)
        d["品牌"] = d["品牌"].fillna("未标注").apply(norm_brand)
        d["口味名称"] = d["口味名称"].fillna("未命名").astype(str).str.strip()
        d["口味_key"] = d["口味名称"].str.lower().str.replace(r"\s+", " ", regex=True).str.strip()
        d["口味标签原文列表"] = d["口味标签"].apply(split_tags_raw) if "口味标签" in d else [[] for _ in range(len(d))]
        d["口味标签列表"] = d["口味标签"].apply(split_tags) if "口味标签" in d else [[] for _ in range(len(d))]
        d["口味标签展示"] = d["口味标签原文列表"].apply(lambda xs: list(xs))
        d["口味标签唯一列表"] = d["口味标签列表"].apply(lambda xs: list(dict.fromkeys(xs)))
        d["分类展示"] = d["分类"].apply(display_category)
        d["口味标签数"] = d["口味标签唯一列表"].apply(len)
        d["口味复杂度"] = d["口味标签数"].apply(flavor_complexity)
        d["价格_IDR"] = d["价格"].apply(parse_idr) if "价格" in d else None
        d["容量列表"] = d["容量ml"].apply(nums_in) if "容量ml" in d else [[] for _ in range(len(d))]
        d["容量ml_主规格"] = d["容量列表"].apply(lambda x: x[0] if x else None)
        d["容量ml_最大"] = d["容量列表"].apply(lambda x: max(x) if x else None)
        
        nic_col = next((c for c in d.columns if "尼古丁" in c and "mg" in c), None)
        if nic_col:
            d["尼古丁列表"] = d[nic_col].apply(nums_in)
        else:
            d["尼古丁列表"] = [[] for _ in range(len(d))]
            
        d["尼古丁最高"] = d["尼古丁列表"].apply(lambda x: max(x) if x else None)
        d["含冰_标签判断"] = d["口味标签列表"].apply(lambda x: any(t in ICE_TAGS or "ice" in t or "mint" in t for t in x))
        d["含烟草_标签判断"] = d["口味标签列表"].apply(lambda x: any("tobacco" in t for t in x))
        d["含冰_原始"] = d["含冰/薄荷"]
        d["含烟草_原始"] = d["含烟草"]
        
        d["数据校验提示"] = ""
        d.loc[(d["含冰/薄荷"] == "否") & d["含冰_标签判断"], "数据校验提示"] = "冰感标签与含冰标记冲突"
        d.loc[(d["含烟草"] == "否") & d["含烟草_标签判断"], "数据校验提示"] = "烟草标签与含烟草标记冲突"
        
        if ranked:
            d["排名"] = pd.to_numeric(d["排名"], errors="coerce")
            d["榜单得分"] = (TOPN + 1 - d["排名"]).clip(lower=1)
            d["产品_key"] = d["品牌"].astype(str) + " | " + d["口味_key"]
        else:
            d["排名"] = pd.NA; d["榜单得分"] = pd.NA
            
        for c in ["详情页", "榜单来源"]:
            if c not in d: d[c] = ""
            d[c] = d[c].fillna("").astype(str)
            
        d["单价_IDR每ml"] = d.apply(lambda r: r["价格_IDR"] / r["容量ml_主规格"] if pd.notna(r["价格_IDR"]) and pd.notna(r["容量ml_主规格"]) and r["容量ml_主规格"] > 0 else None, axis=1)
        d["价格_CNY"] = d["价格_IDR"] * CNY_PER_IDR
        return d

    return enrich(pop, True), enrich(brand, False)

excel_path = resolve_excel()
if not excel_path:
    st.error(f"未找到 Excel：请将「印尼热销产品榜单-20261030xlsx.xlsx」与脚本放在同一目录。当前目录：{_DIR}")
    st.stop()

try: 
    pop_all, brand_all = load_data(excel_path, os.path.getmtime(excel_path))
except Exception as e:
    st.error(f"Excel 读取失败：{e}"); st.stop()

# ------------------------------ 市场背景 ------------------------------
market = pd.DataFrame([
    [2018,197,130],[2019,330,200],[2020,352,230],[2021,447,300],
    [2022,667,380],[2023,860,450],[2024,1175,500],[2025,1357,525]
], columns=["年份", "市场规模_USD百万", "用户数_万人"])
market["同比"] = market["市场规模_USD百万"].pct_change() * 100
market["年份标签"] = market["年份"].astype(str) + market["年份"].apply(lambda x: "F" if x == 2025 else "")

channel = pd.DataFrame([
    [2018,35,60,5],[2019,35,60,5],[2020,50,45,5],[2021,40,55,5],
    [2022,40,55,5],[2023,45,45,10],[2024,40,45,15],[2025,37,45,18]
], columns=["年份", "线上", "专卖店", "传统零售"])

device = pd.DataFrame([
    [2022,18,40,30,9,3],[2023,20,30,35,10,5],[2024,22,15,40,15,8],[2025,20,10,42,16,12]
], columns=["年份", "基础套装", "高级套装", "开放式 Pod", "封闭式 Pod", "一次性"])

# ------------------------------ 侧边栏筛选 ------------------------------
with st.sidebar:
    st.markdown("## 🔍 筛选（热销榜数据）")
    devices = sorted(pop_all["设备类别"].dropna().unique().tolist())
    sel_devices = st.multiselect("产品形态", devices, default=devices)
    types = sorted(pop_all["产品类型"].dropna().unique().tolist())
    sel_types = st.multiselect("烟油/产品类型", types, default=types)
    brands = sorted(pop_all["品牌"].dropna().unique().tolist())
    sel_brands = st.multiselect("品牌", brands, default=brands)
    cats = sorted(pop_all["分类"].dropna().unique().tolist())
    sel_cats = st.multiselect("口味分类", cats, default=cats)
    sites = sorted(pop_all["网站名称"].dropna().unique().tolist())
    sel_sites = st.multiselect("数据来源网站", sites, default=sites)
    origins = sorted(pop_all["品牌来源国"].dropna().unique().tolist())
    sel_origins = st.multiselect("品牌来源国", origins, default=origins)
    
    st.markdown("---")
    st.caption(f"汇率参考：US$1 ≈ Rp {IDR_PER_USD:,}（固定调研口径）")
    st.caption("指定1Rp = 0.00037CNY      US$1 ≈6.72 CNY")
    st.caption(f"数据日期：{DATA_DATE} · 热销榜 {len(pop_all)} 条 · {pop_all['网站名称'].nunique()} 个网站")
    st.caption("左侧筛选仅联动热销榜分析；Foom / RELX / OXVA 在售口味库独立展示。")

fdf = pop_all[
    pop_all["设备类别"].isin(sel_devices) & pop_all["产品类型"].isin(sel_types) & 
    pop_all["品牌"].isin(sel_brands) & pop_all["分类"].isin(sel_cats) & 
    pop_all["网站名称"].isin(sel_sites) & pop_all["品牌来源国"].isin(sel_origins)
].copy()

if fdf.empty: 
    st.warning("当前筛选条件下无热销榜数据，请放宽筛选条件。")
    st.stop()

st.markdown("# 🇮🇩 印度尼西亚电子烟市场调研数据看板")
st.caption(f"热销榜筛选后 {len(fdf)} 条记录｜品牌 {fdf['品牌'].nunique()} 个｜口味 {fdf['口味_key'].nunique()} 个｜网站 {fdf['网站名称'].nunique()} 个；品牌烟油库 {len(brand_all)} 条（独立统计）")
st.markdown("---")

tabs = st.tabs(["📊 总览", "🍭 口味热榜", "🔖 口味元素", "🏷️ 品牌分析", "💉 产品规格", "💰 价格分析", "⭐ 品牌烟油库", "⚖️ 法规与税收", "📝 总结洞察"])

# ============================== TAB 0 总览 ==============================
with tabs[0]:
    st.markdown("## 📊 印尼电子烟市场概览")
    st.caption("本页将印尼宏观市场背景与本次 5 个网站热销榜样本合并呈现；两类数据口径不同，不能直接互相推算。")
    
    local_n = int((fdf["是否本地代理/本土品牌"] == "是").sum())
    ice_n = int((fdf["含冰/薄荷"] == "是").sum())
    el_n = int((fdf["设备类别"] == "瓶装烟油").sum())
    dp_n = int((fdf["设备类别"] == "一次性电子烟").sum())
    
    k = st.columns(6)
    for col, (lab, val, alt) in zip(k, [
        ("热销榜记录", len(fdf), False), ("覆盖品牌", fdf["品牌"].nunique(), False), 
        ("独立口味", fdf["口味_key"].nunique(), False), ("本土/本地代理标记", pct(local_n, len(fdf)), True), 
        ("含冰/薄荷标记", pct(ice_n, len(fdf)), True), ("覆盖网站", fdf["网站名称"].nunique(), False)
    ]): 
        metric(col, lab, val, alt)
        
    section("印尼市场规模与用户基础｜ECigIntelligence 截至 2025 年 5 月")
    m1, m2, m3, m4 = st.columns(4)
    metric(m1, "2025F 市场规模", f"US$ {market.iloc[-1]['市场规模_USD百万']:,.0f}M")
    metric(m2, "2025F 用户数", f"{market.iloc[-1]['用户数_万人']:,.0f} 万", True)
    metric(m3, "2024 市场规模", f"US$ {market.iloc[-2]['市场规模_USD百万']:,.0f}M")
    metric(m4, "2025F 市场规模同比", f"{market.iloc[-1]['同比']:.1f}%", True)
    
    fig = go.Figure()
    fig.add_trace(go.Bar(x=market["年份标签"], y=market["市场规模_USD百万"], name="市场规模（US$M）", yaxis="y"))
    fig.add_trace(go.Scatter(x=market["年份标签"], y=market["用户数_万人"], name="用户数（万人）", yaxis="y2", mode="lines+markers"))
    fig.update_layout(template=T, yaxis=dict(title="市场规模（US$M）"), yaxis2=dict(title="用户数（万人）", overlaying="y", side="right"), legend=dict(orientation="h"), height=390, margin=dict(t=20, b=20))
    show_chart(fig)
    
    st.markdown('<div class="warning-box"><b>增速放缓警示：</b>2018-2024 年 CAGR 约 34.7%，但 2025F 预测增速降至 15.5%。消费税收入 2025 年仅增长 7.38%，说明市场已从高速扩张进入成熟期。2025F 为预测值，不含 2026 年监管收紧影响。</div>', unsafe_allow_html=True)
    
    # 【补充】EC 列示的头部品牌
    section("ECigIntelligence 列示的头部品牌（调研视角）")
    st.caption("以下为 ECigIntelligence 调研资料列示的印尼头部品牌，排序未说明，仅供市场格局参考。")
    ec_brands = pd.DataFrame({
        "烟油品牌（Top e-liquid brands）": ["Foom", "Emkay Brewer", "Bequ", "The O Rama", "Upods", "EJM", "RELX", "YB"],
        "硬件品牌（Top hardware brands）": ["RELX", "OXVA", "Lost Vape", "Foom", "Voopoo", "Uwell", "Vaporesso", "Elf Bar"]
    })
    show_df(ec_brands)
    st.markdown('<div class="insight-box"><b>品牌格局交叉验证：</b>EC 列示的头部烟油品牌中，Foom、Emkay Brewer、RELX 在本次 5 个零售网站热销榜中同样高频出现，说明本土品牌（Foom、Emkay）与国际品牌（RELX）在印尼市场均具有较强影响力。硬件品牌中 RELX、OXVA、Foom 也出现在我们的品牌烟油库或热销榜中。EC 数据与本次零售榜单采集结果在头部品牌上具有较高一致性。</div>', unsafe_allow_html=True)

    left, right = st.columns(2)
    with left:
        section("销售渠道结构变化（ECigIntelligence 调研序列）")
        ch = channel.melt(id_vars="年份", var_name="渠道", value_name="占比")
        fig = px.bar(ch, x="年份", y="占比", color="渠道", barmode="stack", text="占比", template=T)
        fig.update_layout(yaxis_title="占比（%）", xaxis_title="", height=380)
        show_chart(fig)
        
    with right:
        section("设备类型结构变化（ECigIntelligence 调研序列）")
        dv = device.melt(id_vars="年份", var_name="设备类型", value_name="占比")
        fig = px.bar(dv, x="年份", y="占比", color="设备类型", barmode="stack", text="占比", template=T)
        fig.update_layout(yaxis_title="占比（%）", xaxis_title="", height=380)
        show_chart(fig)
        
    st.markdown('<div class="insight-box"><b>渠道与设备洞察：</b>专卖店仍是主要渠道（45%）；线上占比由 2020 年的 50% 回落至 37%，传统零售升至 18%。开放式 Pod 占比由 30% 升至 42%，一次性占比为 12%。</div>', unsafe_allow_html=True)
        
    section("热销榜样本：产品形态与烟油类型")
    left, right = st.columns(2)
    with left:
        d = vc_df(fdf["设备类别"], "产品形态")
        fig = px.pie(d, names="产品形态", values="记录数", hole=.42, color="产品形态", color_discrete_map={"瓶装烟油": COLORS["secondary"], "一次性电子烟": COLORS["primary"]}, template=T)
        fig.update_traces(textinfo="percent+label")
        fig.update_layout(height=400)
        show_chart(fig)
        
    with right:
        d = vc_df(fdf["产品类型"], "产品类型")
        fig = px.bar(d, x="产品类型", y="记录数", text="记录数", color="产品类型", template=T)
        fig.update_layout(showlegend=False, xaxis_title="", yaxis_title="榜单记录数", height=400)
        fig.update_traces(textposition="outside")
        show_chart(fig)
        
    st.markdown(f'<div class="insight-box"><b>样本结构：</b>瓶装烟油 {el_n} 条（{pct(el_n, len(fdf))}），一次性电子烟 {dp_n} 条（{pct(dp_n, len(fdf))}）。当前样本明显以瓶装烟油为主，Nic Salt / Freebase / Disposable Vape 的数量分别为：' + "、".join(f"{k} {v} 条" for k, v in fdf["产品类型"].value_counts().items()) + '。这与 ECigIntelligence 的设备结构数据（开放式 Pod 42%）基本吻合，但不代表全国市场。</div>', unsafe_allow_html=True)
        
    left, right = st.columns(2)
    with left:
        section("热销榜口味分类构成")
        d = vc_df(fdf["分类"], "口味分类")
        fig = px.pie(d, names="口味分类", values="记录数", hole=.4, color="口味分类", color_discrete_map=CAT_COLOR_MAP, template=T)
        fig.update_traces(textinfo="percent+label")
        fig.update_layout(height=400)
        show_chart(fig)
        
    with right:
        section("各网站样本量")
        d = vc_df(fdf["网站名称"], "网站").sort_values("记录数")
        fig = px.bar(d, x="记录数", y="网站", orientation="h", text="记录数", template=T, color_discrete_sequence=[COLORS["primary"]])
        fig.update_traces(textposition="outside")
        fig.update_layout(height=400)
        show_chart(fig)
        
    st.markdown('<div class="insight-box"><b>口味与网站洞察：</b>甜点、水果、饮料是主要分类；各网站样本量不均衡，频次统计会受到榜单长度影响。</div>', unsafe_allow_html=True)
    
# ============================== TAB 1 口味热榜 ==============================
with tabs[1]:
    st.markdown("## 🍭 热销口味榜单分析")
    left, right = st.columns(2)
    with left:
        section("口味名称出现频次≥2")
        fl = counts_ge2(fdf["口味_key"])
        names = fdf.drop_duplicates("口味_key").set_index("口味_key")["口味名称"].to_dict()
        d = pd.DataFrame({"口味": [names.get(x, x) for x in fl.index], "出现次数": fl.values}).sort_values("出现次数")
        if d.empty:
            st.info("当前筛选下没有口味名称达到频次≥2。")
            st.markdown('<div class="insight-box aligned-insight"><b>总榜结论：</b>没有口味名称达到频次≥2，无法判断重复出现的热门口味。</div>', unsafe_allow_html=True)
        else:
            fig = px.bar(d, x="出现次数", y="口味", orientation="h", text="出现次数", template=T, color_discrete_sequence=[COLORS["primary"]])
            fig.update_layout(height=max(420, 32 * len(d) + 120), xaxis_title="记录数", yaxis_title="")
            fig.update_traces(textposition="outside")
            show_chart(fig)
            leaders = fl[fl == fl.max()]
            leader_text = "、".join(f"{names.get(k, k)}（{int(v)}条）" for k, v in leaders.items())
            singleton_n = int((fdf["口味_key"].value_counts() == 1).sum())
            st.markdown(f'<div class="insight-box aligned-insight"><b>总榜结论：</b>共有 {len(fl)} 个口味名称达到频次≥2，最高频为 {leader_text}。另有 {singleton_n} 个口味仅出现 1 次，属于长尾记录。</div>', unsafe_allow_html=True)

    with right:
        section("不同产品类型的热门口味")
        type_options = sorted(fdf["产品类型"].unique())
        sel_type = st.selectbox("查看产品类型", type_options, key="flavor_type")
        sub = fdf[fdf["产品类型"] == sel_type]
        vc = counts_ge2(sub["口味_key"])
        d = pd.DataFrame({"口味": [names.get(x, x) for x in vc.index], "记录数": vc.values}).sort_values("记录数")
        if d.empty:
            st.info(f"{sel_type} 当前没有出现频次≥2的口味，样本重复度不足，无法判断热门口味。")
            st.markdown(f'<div class="insight-box aligned-insight"><b>{sel_type} 小结：</b>共 {len(sub)} 条记录，但没有口味名称达到频次≥2，不将单次出现口味判定为热门。</div>', unsafe_allow_html=True)
        else:
            fig = px.bar(d, x="记录数", y="口味", orientation="h", text="记录数", template=T, color_discrete_sequence=[COLORS["secondary"]])
            fig.update_layout(height=max(420, 32 * len(d) + 120), xaxis_title="记录数", yaxis_title="")
            fig.update_traces(textposition="outside")
            show_chart(fig)
            lead = vc[vc == vc.max()]
            lead_text = "、".join(f"{names.get(k, k)}（{int(v)}条）" for k, v in lead.items())
            st.markdown(f'<div class="insight-box aligned-insight"><b>{sel_type} 小结：</b>共 {len(sub)} 条记录；频次≥2的口味中，最高频为 {lead_text}。</div>', unsafe_allow_html=True)


    section("口味分类 × 产品类型")
    ct = pd.crosstab(fdf["分类"], fdf["产品类型"])
    fig = px.imshow(ct, text_auto=True, aspect="auto", color_continuous_scale=HEAT_SCALE, template=T)
    fig.update_layout(coloraxis_showscale=False, height=360, xaxis_title="产品类型", yaxis_title="口味分类")
    show_chart(fig)

    section("跨网站重复上榜的口味（同名口味出现于 ≥2 个网站）")
    dup = fdf.groupby("口味_key").agg(口味名称=("口味名称", "first"), 出现网站数=("网站名称", "nunique"), 记录数=("网站名称", "size"), 品牌数=("品牌", "nunique"), 平均排名=("排名", "mean")).reset_index()
    dup = dup[dup["出现网站数"] >= 2].sort_values(["出现网站数", "记录数"], ascending=False)
    if dup.empty:
        st.info("当前筛选下无跨站点重复口味。")
    else:
        dup["平均排名"] = dup["平均排名"].round(1)
        show_df(dup[["口味名称", "出现网站数", "记录数", "品牌数", "平均排名"]].head(50))
        st.markdown('<div class="insight-box"><b>跨站重复口味商业价值：</b>跨站点重复出现的口味名称具有更高的市场共识度，是重点关注的潜力口味。按口味名称（忽略大小写和多余空格）匹配；不同品牌的同名口味会合并，因此这里衡量的是口味名称跨站点覆盖，不等同于同一 SKU 重复上榜。</div>', unsafe_allow_html=True)

# ============================== TAB 2 口味元素 ==============================
with tabs[2]:
    st.markdown("## 🔖 口味元素与组合分析")
    tags = [t for row in fdf["口味标签列表"] for t in row]
    tag_display_map = {}
    for norm_row, raw_row in zip(fdf["口味标签列表"], fdf["口味标签原文列表"]):
        for norm, raw in zip(norm_row, raw_row):
            tag_display_map.setdefault(norm, raw)
    c1, c2, c3, c4 = st.columns(4)
    metric(c1, "标签总出现次数", len(tags))
    metric(c2, "不同标签数", len(set(tags)))
    metric(c3, "无标签记录", int((fdf["口味标签数"] == 0).sum()), True)
    metric(c4, "复合口味记录", int((fdf["口味标签数"] >= 2).sum()), True)

    brand_options = sorted(fdf["品牌"].dropna().unique())
    left, right = st.columns(2)
    with left:
        section("整体高频口味元素 Top 25（总榜）")
        # 与右侧品牌选择器占据相同的顶部空间，使两侧图表与结论框水平对齐。
        st.markdown('<div class="element-selector-spacer"></div>', unsafe_allow_html=True)
        tag_counts = Counter(tags)
        top = top_counter(tag_counts, 25, include_singletons=True)
        if len(top):
            d = pd.DataFrame({"元素": [tag_display_map.get(x, x) for x in top.index], "出现次数": top.values}).sort_values("出现次数")
            fig = px.bar(d, x="出现次数", y="元素", orientation="h", text="出现次数", template=T, color_discrete_sequence=[COLORS["primary"]])
            fig.update_layout(height=600, xaxis_title="出现次数", yaxis_title="")
            fig.update_traces(textposition="outside")
            show_chart(fig)
            top3 = top_counter(tag_counts, 3, include_singletons=True)
            lead = "、".join(f"{tag_display_map.get(k, k)}（{int(v)}次）" for k, v in top3.items())
            st.markdown(f'<div class="insight-box aligned-insight"><b>总榜结论：</b>前三高频元素为 {lead}。图表保留 Excel 中的原始英文标签，统计单位为榜单记录。</div>', unsafe_allow_html=True)
        else:
            st.info("没有可用口味标签。")
            st.markdown('<div class="insight-box aligned-insight"><b>总榜结论：</b>当前筛选下没有可用的口味标签数据。</div>', unsafe_allow_html=True)

    with right:
        section("按品牌查看高频口味元素")
        chosen_tag_brand = st.selectbox("选择品牌", brand_options, key="element_brand")
        brand_tags = [t for row in fdf.loc[fdf["品牌"] == chosen_tag_brand, "口味标签列表"] for t in row]
        btop = top_counter(Counter(brand_tags), 25, include_singletons=True)
        if len(btop):
            bd = pd.DataFrame({"元素": [tag_display_map.get(x, x) for x in btop.index], "出现次数": btop.values}).sort_values("出现次数")
            fig = px.bar(bd, x="出现次数", y="元素", orientation="h", text="出现次数", template=T, color_discrete_sequence=[COLORS["secondary"]])
            fig.update_layout(height=600, xaxis_title="出现次数", yaxis_title="")
            fig.update_traces(textposition="outside")
            show_chart(fig)
            st.markdown(f'<div class="insight-box aligned-insight"><b>{chosen_tag_brand} 小结：</b>{len(fdf[fdf["品牌"] == chosen_tag_brand])} 条热销榜记录；标签出现 {len(brand_tags)} 次。图表保留原始英文标签。</div>', unsafe_allow_html=True)
        else:
            st.info(f"{chosen_tag_brand} 暂无可用标签。")
            st.markdown(f'<div class="insight-box aligned-insight"><b>{chosen_tag_brand} 小结：</b>该品牌在当前筛选下暂无可用的口味标签记录。</div>', unsafe_allow_html=True)

    left, right = st.columns(2)
    with left:
        section("口味标签复杂度")
        d = vc_df(fdf["口味复杂度"], "标签复杂度")
        fig = px.pie(d, names="标签复杂度", values="记录数", hole=.4, template=T)
        fig.update_traces(textinfo="percent+label")
        fig.update_layout(height=400)
        show_chart(fig)
    with right:
        section("含冰 / 薄荷标记")
        d = vc_df(fdf["含冰/薄荷"], "含冰标记")
        fig = px.pie(d, names="含冰标记", values="记录数", hole=.4, template=T)
        fig.update_traces(textinfo="percent+label")
        fig.update_layout(height=400)
        show_chart(fig)

    total_n = len(fdf)
    ice_tag_n = int(sum(1 for row in fdf["口味标签唯一列表"] if any(t in ICE_TAGS or "ice" in t or "mint" in t for t in row)))
    double_n = int((fdf["口味标签数"] == 2).sum())
    multi3_n = int((fdf["口味标签数"] >= 3).sum())
    multi_n = double_n + multi3_n
    raw_ice_n = int((fdf["含冰/薄荷"] == "是").sum())
    st.markdown(f'<div class="insight-box"><b>口味组合与冰感元素结论：</b>多元素组合占比为 {pct(multi_n, total_n)}，其中双元素组合占比为 {pct(double_n, total_n)}，三种及以上元素占比为 {pct(multi3_n, total_n)}。冰感元素：标签侧识别到含冰/薄荷 {ice_tag_n} 条（{pct(ice_tag_n, total_n)}），原始字段明确标记“是”为 {raw_ice_n} 条（{pct(raw_ice_n, total_n)}）。</div>', unsafe_allow_html=True)

    section("各口味分类的高频元素")
    for cat in fdf["分类"].value_counts().index:
        sub = fdf[fdf["分类"] == cat]
        cat_counts = Counter([t for row in sub["口味标签列表"] for t in row])
        topcat = top_counter(cat_counts, 7, include_singletons=True)
        chips = " ".join(f'<span class="tag-chip">{t}（{n}）</span>' for t, n in topcat.items()) or "无标签"
        st.markdown(f'<div style="padding:7px 10px;margin:5px 0;background:#f8fafc;border-radius:6px"><b>{cat}</b>（{len(sub)}条） {chips}</div>', unsafe_allow_html=True)

    section("高频双元素组合（出现 ≥3 次，完整展示并列项）")
    combos = Counter()
    for row in fdf["口味标签唯一列表"]:
        u = sorted(row)
        for i in range(len(u)):
            for j in range(i + 1, len(u)):
                combos[(u[i], u[j])] += 1
    frequent = {k: v for k, v in combos.items() if v >= 3}
    if frequent:
        cc = sorted(frequent.items(), key=lambda x: (-x[1], x[0]))
        d = pd.DataFrame({"组合": [f"{a} + {b}" for (a, b), n in cc], "出现次数": [n for (a, b), n in cc]}).sort_values("出现次数")
        fig = px.bar(d, x="出现次数", y="组合", orientation="h", text="出现次数", template=T, color_discrete_sequence=[COLORS["secondary"]])
        fig.update_layout(height=max(360, 28 * len(d) + 100))
        fig.update_traces(textposition="outside")
        show_chart(fig)
        max_n = max(frequent.values())
        leaders = [f"{a} + {b}" for (a, b), v in frequent.items() if v == max_n]
        st.markdown(f'<div class="insight-box"><b>组合洞察：</b>最高频组合出现 {max_n} 次：' + "、".join(leaders) + f'。当前展示所有出现 3 次及以上的组合，共 {len(d)} 组。</div>', unsafe_allow_html=True)
    else:
        st.info("当前样本没有出现 3 次及以上的双元素组合。")

# ============================== TAB 3 品牌分析 ==============================
with tabs[3]:
    st.markdown("## 🏷️ 热销榜品牌分析")
    left, right = st.columns(2)
    with left:
        section("品牌上榜频次≥2")
        bc = counts_ge2(fdf["品牌"])
        if bc.empty:
            st.info("当前筛选下没有品牌达到上榜频次≥2。")
            st.markdown('<div class="insight-box aligned-insight"><b>频次小结：</b>当前筛选下没有品牌达到上榜频次≥2，无法判断重复上榜品牌。</div>', unsafe_allow_html=True)
        else:
            d = pd.DataFrame({"品牌": bc.index, "上榜记录数": bc.values}).sort_values("上榜记录数")
            fig = px.bar(d, x="上榜记录数", y="品牌", orientation="h", text="上榜记录数", template=T, color_discrete_sequence=[COLORS["primary"]])
            fig.update_layout(height=600, xaxis_title="上榜记录数", yaxis_title="")
            fig.update_traces(textposition="outside")
            show_chart(fig)
            top3 = top_n_ties(bc, 3, include_singletons=True)
            st.markdown(f'<div class="insight-box aligned-insight"><b>频次小结：</b>频次≥2的品牌中，最高频为：{"、".join(f"{b}（{int(n)}条）" for b, n in top3.items())}。单次上榜品牌不纳入本图；频次反映样本榜单占位，不是销量。</div>', unsafe_allow_html=True)

    with right:
        section("本土品牌 / 本地代理标记")
        d = vc_df(fdf["是否本地代理/本土品牌"], "本土/代理标记")
        fig = px.pie(d, names="本土/代理标记", values="记录数", hole=.4, template=T, color="本土/代理标记", color_discrete_map={"是": COLORS["primary"], "否": "#94a3b8", "不确定": COLORS["gold"]})
        fig.update_traces(textinfo="percent+label")
        fig.update_layout(height=600)
        show_chart(fig)
        local_yes = int((fdf["是否本地代理/本土品牌"] == "是").sum())
        st.markdown(f'<div class="insight-box aligned-insight"><b>本土化洞察：</b>标记为本土品牌/本地代理的记录占 {pct(local_yes, len(fdf))}（{local_yes}/{len(fdf)} 条）。这反映当前网站样本中的品牌标记结构，不能直接推导整个印尼市场的品牌份额；本土/本地代理品牌可作为销售线索优先核实。</div>', unsafe_allow_html=True)

    section("品牌来源国")
    d = vc_df(fdf["品牌来源国"], "来源国")
    fig = px.pie(d, names="来源国", values="记录数", hole=.4, template=T)
    fig.update_traces(textinfo="percent+label")
    fig.update_layout(height=400)
    show_chart(fig)
    country_top = top_n_ties(fdf["品牌来源国"].value_counts(), 1)
    if len(country_top):
        st.markdown('<div class="insight-box"><b>来源国洞察：</b>' + "、".join(f"{k} {int(v)} 条（{pct(v, len(fdf))}）" for k, v in country_top.items()) + f'；其余来源国合计 {len(fdf) - int(country_top.sum())} 条。该结构仅反映当前筛选样本。</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="insight-box"><b>来源国小结：</b>当前筛选下无品牌来源国数据。</div>', unsafe_allow_html=True)

    section("网站 × 品牌覆盖矩阵（Top 15）")
    topb = top_n_ties(fdf["品牌"].value_counts(), 15).index.tolist()
    cross = pd.crosstab(fdf["品牌"], fdf["网站名称"])
    cross = cross.loc[[b for b in topb if b in cross.index]]
    fig = px.imshow(cross, text_auto=True, aspect="auto", color_continuous_scale=HEAT_SCALE, template=T)
    fig.update_layout(coloraxis_showscale=False, height=max(360, min(720, 24 * len(cross))), xaxis_title="网站", yaxis_title="品牌")
    show_chart(fig)
    spread = fdf.groupby("品牌")["网站名称"].nunique().sort_values(ascending=False)
    spread_top = top_n_ties(spread, 1, include_singletons=True)
    st.markdown(f'<div class="insight-box"><b>渠道覆盖洞察：</b>覆盖网站最多的品牌（并列完整展示）：' + "、".join(f"{b}（{int(n)}个网站）" for b, n in spread_top.items()) + f'；仅在单一网站出现的品牌 {int((spread == 1).sum())} 个。多网站覆盖品牌具有更强的渠道渗透力。</div>', unsafe_allow_html=True)

    section("品牌综合热度榜 Top 20（榜单得分）")
    heat = fdf.groupby(["品牌", "品牌来源国"]).agg(上榜次数=("排名", "size"), 覆盖网站数=("网站名称", "nunique"), 平均排名=("排名", "mean"), 榜单得分=("榜单得分", "sum")).reset_index()
    heat["平均排名"] = heat["平均排名"].round(1)
    heat = heat.sort_values(["榜单得分", "上榜次数"], ascending=False)
    if len(heat) > 20: heat = heat[heat["榜单得分"] >= heat.iloc[19]["榜单得分"]]
    show_df(heat)
    tie_note(len(heat), 20)
    if not heat.empty:
        lead_score = heat["榜单得分"].max()
        lead_brands = heat[heat["榜单得分"] == lead_score]
        st.markdown(f'<div class="insight-box"><b>热度洞察：</b>综合得分领先品牌（并列完整展示）：' + "、".join(f"{r['品牌']}（{int(r['榜单得分'])}分）" for _, r in lead_brands.iterrows()) + '。榜单得分 = 各条记录（21 − 排名）之和；同时考虑名次和出现次数，仅作相对热度指标。</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="insight-box"><b>热度小结：</b>当前筛选下无品牌热度数据。</div>', unsafe_allow_html=True)

    section("选择品牌查看代表性口味")
    bsel = st.selectbox("选择热销榜品牌", sorted(fdf["品牌"].unique()), key="pop_brand_flavor")
    sub = fdf[fdf["品牌"] == bsel]
    # 这里展示品牌的代表性口味，允许仅出现 1 次的记录；“热门”口径只用于上方频次分析。
    fl = top_n_ties(sub["口味_key"].value_counts(), 15, include_singletons=True)
    if len(fl):
        names = sub.drop_duplicates("口味_key").set_index("口味_key")["口味名称"].to_dict()
        d = pd.DataFrame({"口味": [names.get(x, x) for x in fl.index], "记录数": fl.values}).sort_values("记录数")
        fig = px.bar(d, x="记录数", y="口味", orientation="h", text="记录数", template=T, color_discrete_sequence=[COLORS["secondary"]])
        fig.update_layout(height=max(400, 28 * len(d) + 100))
        fig.update_traces(textposition="outside")
        show_chart(fig)
    else:
        st.info("当前品牌没有可展示的口味记录。")

# ============================== TAB 4 产品规格 ==============================
with tabs[4]:
    st.markdown("## 💉 产品规格与尼古丁结构")
    st.caption("本页按采集到的规格字段统计，不对印尼法规未明确规定的尼古丁浓度自行设定合规上限。")
    left, right = st.columns(2)
    with left:
        section("尼古丁浓度档位分布（mg/mL）")
        vals = [x for row in fdf["尼古丁列表"] for x in row]
        if vals:
            d = pd.DataFrame(Counter(vals).items(), columns=["浓度mg/mL", "出现次数"]).sort_values("浓度mg/mL")
            d["浓度mg/mL"] = d["浓度mg/mL"].map(lambda x: f"{x:g}")
            fig = px.bar(d, x="浓度mg/mL", y="出现次数", text="出现次数", template=T, color_discrete_sequence=[COLORS["primary"]])
            fig.update_layout(xaxis_title="标注浓度（mg/mL）", yaxis_title="出现次数", height=400)
            fig.update_traces(textposition="outside")
            show_chart(fig)
        else: 
            st.info("无尼古丁浓度字段。")
            
    with right:
        section("Nic Salt / Freebase / Disposable 构成")
        d = vc_df(fdf["产品类型"], "产品类型")
        fig = px.pie(d, names="产品类型", values="记录数", hole=.4, template=T)
        fig.update_traces(textinfo="percent+label")
        fig.update_layout(height=400)
        show_chart(fig)
        
    st.markdown('<div class="insight-box"><b>尼古丁浓度洞察：</b>Nic Salt 以 30mg 为主，Freebase 以 3mg 为主，一次性电子烟样本主要为 20mg。法规要求检测，具体最高限值仍需以正式发布的适用文本为准。</div>', unsafe_allow_html=True)
        
    section("烟油容量规格分布")
    d = vc_df(fdf["容量ml_主规格"].dropna().map(lambda x: f"{x:g} mL"), "容量")
    if not d.empty:
        fig = px.bar(d, x="容量", y="记录数", text="记录数", template=T, color_discrete_sequence=[COLORS["secondary"]])
        fig.update_layout(xaxis_title="", yaxis_title="记录数", height=520, margin=dict(t=30, r=55, b=95, l=25))
        fig.update_xaxes(automargin=True, tickangle=0)
        fig.update_traces(cliponaxis=False, texttemplate="%{text}")
        fig.update_traces(textposition="outside")
        show_chart(fig)
        
    over20 = fdf[fdf["容量ml_主规格"].notna() & (fdf["容量ml_主规格"] > 20)]
    over20_counts = over20["容量ml_主规格"].value_counts().sort_index()
    over20_detail = "、".join(f"{x:g} mL {int(n)} 条" for x, n in over20_counts.items()) or "暂无"
    st.markdown(f'<div class="warning-box"><b>包装规格合规复核：</b>当前筛选中，容量标注大于 20 mL 的记录有 {len(over20)} 条（{pct(len(over20), len(fdf))}），主要为 {over20_detail}。依据 PP 28/2024 第 433 条，仍需结合产品属于开放式、封闭式或一次性系统，以及生产/进口日期、标签、许可、实际包装和过渡安排进行核验。线上商品页只能作为合规核验信号，不能单凭挂牌信息认定违法。商业落地前应由当地合规顾问确认适用条款和执行口径。</div>', unsafe_allow_html=True)

    section("产品类型 × 尼古丁浓度档位")
    nicrows = []
    for _, r in fdf.iterrows():
        for n in r["尼古丁列表"]: nicrows.append({"产品类型": r["产品类型"], "浓度": f"{n:g} mg"})
    if nicrows:
        nd = pd.DataFrame(nicrows)
        ct = pd.crosstab(nd["产品类型"], nd["浓度"])
        fig = px.imshow(ct, text_auto=True, aspect="auto", color_continuous_scale=HEAT_SCALE, template=T)
        fig.update_layout(coloraxis_showscale=False, height=340)
        show_chart(fig)

    section("容量 × 产品类型")
    cap = fdf.dropna(subset=["容量ml_主规格"]).groupby(["产品类型", "容量ml_主规格"]).size().reset_index(name="记录数")
    if not cap.empty:
        fig = px.bar(cap, x="容量ml_主规格", y="记录数", color="产品类型", barmode="group", text="记录数", template=T)
        fig.update_layout(xaxis_title="容量（mL）", yaxis_title="记录数", height=400)
        show_chart(fig)

# ============================== TAB 5 价格分析 ==============================
with tabs[5]:
    st.markdown("## 💰 印尼电子烟产品价格分析")
    pdf = fdf.dropna(subset=["价格_IDR"]).copy()
    if pdf.empty: 
        st.info("当前筛选没有可解析价格。")
    else:
        c = st.columns(5)
        metric(c[0], "价格样本数", len(pdf))
        price_metric(c[1], "均价", pdf["价格_IDR"].mean())
        price_metric(c[2], "中位价", pdf["价格_IDR"].median(), True)
        price_metric(c[3], "最低价", pdf["价格_IDR"].min())
        price_metric(c[4], "最高价", pdf["价格_IDR"].max())
        
        left, right = st.columns(2)
        with left:
            section("产品类型价格分布")
            fig = px.box(pdf, x="产品类型", y="价格_IDR", color="产品类型", points="all", template=T, hover_data=["品牌", "网站名称", "口味名称"])
            fig.update_layout(xaxis_title="", yaxis_title="标价（IDR；人民币折算见悬浮信息）", showlegend=False, height=450)
            show_chart(fig)
            
        with right:
            section("各品牌中位标价")
            bp = pdf.groupby("品牌")["价格_IDR"].agg(["median", "count"]).reset_index().sort_values("median")
            bp["标签"] = bp.apply(lambda r: f"Rp {r['median']:,.0f}（约 ¥{r['median'] * CNY_PER_IDR:,.2f}，n={int(r['count'])}）", axis=1)
            fig = px.bar(bp, x="median", y="品牌", orientation="h", text="标签", template=T, color_discrete_sequence=[COLORS["primary"]])
            fig.update_layout(xaxis_title="中位标价（Rp）", yaxis_title="", height=450)
            fig.update_xaxes(tickformat=",.0f")
            fig.update_traces(textposition="outside")
            show_chart(fig)
            
        price_by_type = pdf.groupby("产品类型")["价格_IDR"].agg(["median", "min", "max", "count"]).sort_values("median")
        price_txt = "；".join(f"{idx}：中位 {format_idr_cny(row['median'])}（范围 {format_idr_cny(row['min'])}–{format_idr_cny(row['max'])}，n={int(row['count'])}）" for idx, row in price_by_type.iterrows())
        st.markdown(f'<div class="insight-box"><b>价格带洞察：</b>当前筛选样本显示，{price_txt}。这是零售挂牌价的横截面比较，受网站、品牌、规格和促销状态影响，不能直接推导成本、毛利或统一建议零售价。</div>', unsafe_allow_html=True)
            
        section("本土 vs 国际品牌价格对比")
        pdf_copy = pdf.copy()
        pdf_copy["品牌类型"] = pdf_copy["品牌来源国"].apply(lambda x: "本土" if x == "印度尼西亚" else ("中国" if x == "中国" else "进口"))
        fig = px.box(pdf_copy, x="品牌类型", y="价格_IDR", color="品牌类型", points="all", template=T, hover_data=["品牌", "价格_IDR"])
        fig.update_layout(xaxis_title="品牌类型", yaxis_title="标价（IDR；人民币折算见悬浮信息）", height=400)
        show_chart(fig)
        
        section("ECigIntelligence 线上价格参考（用户调研概览）")
        price_ref = pd.DataFrame([
            ["基础套装 Basic kit", 9.46, 9.46 * CNY_PER_USD, 9.46 * IDR_PER_USD, "开放系统套装"],
            ["开放式 Pod 套装", 7.60, 7.60 * CNY_PER_USD, 7.60 * IDR_PER_USD, "开放系统套装"],
            ["高级套装 Advanced kit", 24.46, 24.46 * CNY_PER_USD, 24.46 * IDR_PER_USD, "开放系统套装"],
            ["30mL 自由基烟油（每 mL）", 0.10, 0.10 * CNY_PER_USD, 0.10 * IDR_PER_USD, "RTU"],
            ["30mL 盐烟油（每 mL）", 0.17, 0.17 * CNY_PER_USD, 0.17 * IDR_PER_USD, "RTU"],
            ["60mL 自由基烟油（每 mL）", 0.09, 0.09 * CNY_PER_USD, 0.09 * IDR_PER_USD, "RTU"],
            ["封闭式 Pod 套装", 7.46, 7.46 * CNY_PER_USD, 7.46 * IDR_PER_USD, "封闭系统产品"],
            ["预填充烟弹", 4.80, 4.80 * CNY_PER_USD, 4.80 * IDR_PER_USD, "封闭系统产品"],
            ["一次性", 7.45, 7.45 * CNY_PER_USD, 7.45 * IDR_PER_USD, "封闭系统产品"],
        ], columns=["品类", "线上价格参考_USD", "人民币折算", "印尼盾折算", "备注"])
        price_ref["线上价格参考"] = price_ref["线上价格参考_USD"].map(lambda x: f"US${x:,.2f}")
        price_ref["人民币折算"] = price_ref["人民币折算"].map(lambda x: f"¥{x:,.2f}")
        price_ref["印尼盾折算"] = price_ref["印尼盾折算"].map(lambda x: f"Rp {x:,.0f}")
        price_ref = price_ref[["品类", "线上价格参考", "人民币折算", "印尼盾折算", "备注"]]
        show_df(price_ref)
        st.caption(f"折算汇率按 Rp {IDR_PER_USD:,}/USD 固定值；此表为调研概览参考价，不与零售网站采集价直接混合比较。")
        

# ============================== TAB 6 品牌烟油库 ==============================
with tabs[6]:
    st.markdown("## ⭐ Foom / RELX / OXVA 品牌在售烟油分析")
    st.caption("本页使用独立 Sheet「popular brand eliquid」，共 164 条；它记录品牌站/官网采集到的在售口味，不是热销榜。")
    bdf = brand_all.copy()
    brands_lib = sorted(bdf["品牌"].dropna().unique().tolist())
    selected = st.multiselect("选择对比品牌", brands_lib, default=[b for b in ["Foom", "RELX", "OXVA"] if b in brands_lib], key="brand_lib_select")
    bdf = bdf[bdf["品牌"].isin(selected)]
    
    if bdf.empty: 
        st.info("请选择至少一个品牌。")
    else:
        c = st.columns(5)
        metric(c[0], "在售口味记录", len(bdf))
        metric(c[1], "品牌数", bdf["品牌"].nunique())
        metric(c[2], "饮料类记录", int((bdf["分类"] == "饮料").sum()), True)
        metric(c[3], "水果类记录", int((bdf["分类"] == "水果").sum()), True)
        metric(c[4], "甜点类记录", int((bdf["分类"] == "甜点").sum()), True)
        
        left, right = st.columns(2)
        with left:
            section("各品牌在售口味数量")
            d = vc_df(bdf["品牌"], "品牌", "在售记录数")
            fig = px.bar(d, x="品牌", y="在售记录数", text="在售记录数", template=T, color="品牌")
            fig.update_layout(showlegend=False, xaxis_title="", yaxis_title="记录数", height=400)
            fig.update_traces(textposition="outside")
            show_chart(fig)
            
        with right:
            section("品牌 × 口味分类结构")
            ct = pd.crosstab(bdf["品牌"], bdf["分类"])
            fig = px.bar(ct.reset_index().melt(id_vars="品牌", var_name="分类", value_name="记录数"), x="品牌", y="记录数", color="分类", barmode="stack", color_discrete_map=CAT_COLOR_MAP, template=T)
            fig.update_layout(xaxis_title="", yaxis_title="在售口味记录数", height=400)
            show_chart(fig)
            
        brand_lib_counts = bdf["品牌"].value_counts()
        brand_lib_txt = "、".join(f"{b} {int(n)}条" for b, n in brand_lib_counts.items())
        st.markdown(f'<div class="insight-box"><b>品牌烟油库洞察：</b>当前选定品牌的在售记录为 {brand_lib_txt}。记录数反映本次采集到的在售口味覆盖，不等同于销量或真实 SKU 总量。</div>', unsafe_allow_html=True)
            
        section("品牌口味分类占比对比（百分比）")
        ct = pd.crosstab(bdf["品牌"], bdf["分类"], normalize="index") * 100
        fig = px.imshow(ct.round(1), text_auto=".1f", aspect="auto", color_continuous_scale=HEAT_SCALE, template=T)
        fig.update_layout(coloraxis_showscale=False, height=300, xaxis_title="口味分类", yaxis_title="品牌")
        show_chart(fig)
        
        section("Foom 产品主题线索分析")
        if "Foom" in selected:
            foom_data = bdf[bdf["品牌"] == "Foom"].copy()
            if foom_data.empty:
                st.info("当前筛选没有 Foom 品牌烟油记录。")
            else:
                st.caption("以下分类根据本次采集到的 Foom 口味名称关键词归类，用于观察产品主题分布，不等同于官方完整系列、全部 SKU 或销量。")
                theme_rules = [
                    ("Ice Cream Series", [r"\bice\s*cream\b"]),
                    ("Barista Series", [r"\blatte\b", r"\bcoffee\b", r"\bcappuccino\b", r"\bfrappuccino\b"]),
                    ("Tea Series", [r"\btea\b", r"\boolong\b"]),
                    ("Bubblegum Series", [r"\bbubblegum\b"]),
                    ("Mint / Menthol Theme", [r"\bmenthol\b", r"\bmint\b"]),
                    ("Icy Theme", [r"\bicy\b", r"\biced\b", r"\bice\b"]),
                    ("Sour Series", [r"\bsour\b"]),
                    ("Drink / Beverage Theme", [r"\bmilkshake\b", r"\bmilk\b", r"\bsoda\b", r"\bwater\b", r"\bcola\b", r"\bsparkling\b", r"\bjuice\b", r"\bsmoothie\b"]),
                    ("Tropical / Fruit Theme", [r"\btropical\b", r"\bpassionfruit\b", r"\blychee\b", r"\bsoursop\b", r"\bstarfruit\b", r"\bpineapple\b", r"\bguava\b", r"\bmango\b", r"\bberry\b"]),
                    ("Asian / Local Theme", [r"\btaro\b", r"\bthai\b", r"\bvietnam\b", r"\bmung bean\b", r"\bmatcha\b"]),
                ]
                def classify_foom_theme(flavor_name):
                    text = str(flavor_name).strip().lower()
                    for theme, patterns in theme_rules:
                        if any(re.search(pattern, text) for pattern in patterns):
                            return theme
                    return "Other / Unclassified"
                foom_data["主题线索"] = foom_data["口味名称"].map(classify_foom_theme)
                series_df = foom_data["主题线索"].value_counts().rename_axis("主题线索").reset_index(name="记录数")
                series_df = series_df.sort_values(["记录数", "主题线索"], ascending=[False, True]).reset_index(drop=True)
                fig = px.bar(series_df, x="记录数", y="主题线索", orientation="h", text="记录数", template=T, color_discrete_sequence=[COLORS["secondary"]])
                # Plotly horizontal bars render the last category at the top;
                # reverse the descending data order so the largest bar is topmost.
                fig.update_yaxes(categoryorder="array", categoryarray=series_df["主题线索"].tolist()[::-1])
                fig.update_layout(xaxis_title="采集记录数", yaxis_title="", height=max(430, 32 * len(series_df) + 110))
                fig.update_traces(textposition="outside")
                show_chart(fig)
                top_theme = series_df.iloc[0]
                total_foom = len(foom_data)
                price_info = "价格字段不足"
                if "价格_IDR" in foom_data.columns:
                    foom_prices = foom_data["价格_IDR"].dropna().round(0).value_counts().sort_index()
                    if not foom_prices.empty:
                        price_info = "、".join(f"Rp {price:,.0f}（约 ¥{price * CNY_PER_IDR:,.2f}，{int(count)}条）" for price, count in foom_prices.items())
                st.markdown(f'<div class="insight-box"><b>Foom 产品结构结论：</b>本次采集到 Foom 在售口味记录 {total_foom} 条，数量最多的主题线索为「{top_theme["主题线索"]}」（{int(top_theme["记录数"])} 条，占 {top_theme["记录数"] / total_foom:.1%}）。价格记录显示：{price_info}。该结果反映采集页面中的口味命名和价格分布，不代表官方完整系列、全部 SKU 或销量。</div>', unsafe_allow_html=True)

        section("品牌烟油口味清单")
        detail_cols = [c for c in ["网站名称", "品牌", "口味名称", "口味标签", "分类", "含冰/薄荷", "含烟草", "产品类型", "品牌来源国", "详情页"] if c in bdf]
        detail = bdf[detail_cols].copy()
        if "价格_IDR" in bdf:
            detail["价格（Rp）"] = bdf["价格_IDR"].map(lambda x: f"Rp {x:,.0f}" if pd.notna(x) else "—")
            detail["价格（约人民币）"] = bdf["价格_CNY"].map(lambda x: f"¥{x:,.2f}" if pd.notna(x) else "—")
        show_df(detail)
        export_cols = [c for c in ["网站名称", "品牌", "口味名称", "口味标签", "分类", "含冰/薄荷", "含烟草", "产品类型", "品牌来源国", "价格_IDR", "价格_CNY", "详情页"] if c in bdf]
        export_df = bdf[export_cols].copy()
        st.download_button("下载所选品牌烟油口味清单 CSV", export_df.to_csv(index=False).encode("utf-8-sig"), "brand_eliquid_selected.csv", "text/csv", key="brand_eliquid_download")

# ============================== TAB 7 法规与税收 ==============================
with tabs[7]:
    st.markdown("## ⚖️ 印尼电子烟监管框架与消费税")
    st.caption("法规核验以政府法规数据库和主管部门公开信息为基础；动态事项、过渡安排与地方执行仍需落地前复核。")
    
    regs = [
        ("监管框架与适用范围｜已生效", "UU 17/2023《卫生法》及其实施条例 PP 28/2024 将电子烟纳入成瘾性物质监管；适用范围包括含尼古丁及其他相关液体、固体或气体形态产品，也涵盖合成尼古丁相关产品。BPOM 18/2025（2025-07-03 生效）规定烟草及电子烟产品的监督、实施与经营者报告要求。", "active"),
        ("电子烟禁令｜最新正式颁布", "印尼国家警察总长于 2026-10-30 签署并发布第 SE/9/X/2026 号通函，正式全面禁止警察系统内所有公务员及访客在工作环境内外使用电子烟，以防范依托咪酯等新型毒品滥用并维护机构形象，违者将面临纪律处罚。", "proposal"),
        ("最低销售年龄与销售场所｜已生效", "PP 28/2024 第 434 条：禁止向未满 21 周岁者及孕妇销售；禁止自动售货机销售；不得在出入口、经常通行区域陈列；学校及儿童游乐场所 200 米范围内禁止销售。", "restrict"),
        ("线上销售与社交媒体｜需区分渠道", "第 434 条禁止通过网站、电商平台及社交媒体销售；其中网站和电商平台在实施年龄验证时可适用例外，但该例外不适用于社交媒体。第 446 条另明确禁止在数字社交媒体上投放烟草及电子烟广告。", "restrict"),
        ("包装容量与产品规格｜已生效", "第 433 条：闭系统/一次性烟弹每个烟弹尼古丁液体不得超过 2 mL，且每包不超过 2 个烟弹；开放式/可补充系统液体仅允许 10 mL 或 20 mL 包装；固体电子烟每包 20 支。", "restrict"),
        ("健康警示与标签｜法定要求及过渡安排", "第 435–438 条要求包装符合设计和文字标准，并附图文健康警示；第 438 条规定警示图文覆盖包装正、反面各 50%。ECigIntelligence 对两年过渡期的解读指出，部分产品、检测、许可、标签和警示要求进入 2026-07-26 的实施节点。", "restrict"),
        ("尼古丁检测与最高限值｜法规已要求，数值须谨慎确认", "PP 28/2024 第 431 条要求生产、进口及流通电子烟的经营者具备许可、遵守最高尼古丁限值、对每个产品变体进行尼古丁检测，并向 BPOM 报告检测结果；检测须由认可实验室完成。截至本次调研，未找到已正式公布并生效的电子烟具体 mg/mL 数值。", "warning"),
        ("口味与添加剂｜不是简单的全面口味禁令", "PP 28/2024 第 432 条规定，生产烟草及电子烟产品不得使用未经科学证明对健康无害的添加材料；法规解释将 flavor、aroma、coloring 列为添加材料。合规重点是添加剂安全证明、检测与监督。", "restrict"),
        ("广告与促销｜已生效条款及 2026 年新规", "PP 28/2024 第 446 条禁止数字社交媒体广告；2026-09-14 生效的 Kepmenkes HK.01.07/MENKES/963/2026 进一步发布线上广告监督指南，明确将品牌植入、用户生成内容、产品测评/教程、网红内容和用户评价等纳入广告识别范围。", "restrict"),
        ("标准化/平装包装｜草案阶段", "印尼卫生部 2026-06-05 官方公告称正在起草烟草及电子烟包装和健康警示信息的部长条例，拟引入 plain packaging。该公告描述的是 RPMK 草案制定，不应标记为已正式生效法规。", "proposal"),
        ("全面禁止电子烟｜部门提议，不是禁令", "BNN 于 2026-09-09 官方发布消息，提出可研究通过总统条例全面禁止电子烟；BNN 同时说明仍需跨部门综合研究。该消息是部门提出的政策选项，不代表政府已作出全面禁售决定。", "proposal"),
    ]
    for title, detail, cls in regs: 
        st.markdown(f'<div class="reg-card {cls}"><b>{title}</b>{detail}</div>', unsafe_allow_html=True)
        
    section("电子烟消费税与最低零售价（PMK 96/2024，自 2025-01-01 起）")
    tax = pd.DataFrame([
        ["开放系统液体电子烟", "Rp 1,368/mL（约 ¥{:.2f}/mL）".format(1368 * CNY_PER_IDR), "Rp 636/mL（约 ¥{:.2f}/mL）".format(636 * CNY_PER_IDR)],
        ["封闭系统液体电子烟", "Rp 41,983/烟弹（约 ¥{:.2f}/烟弹）".format(41983 * CNY_PER_IDR), "Rp 6,776/mL（约 ¥{:.2f}/mL）".format(6776 * CNY_PER_IDR)],
        ["固体电子烟", "Rp 6,240/克（约 ¥{:.2f}/克）".format(6240 * CNY_PER_IDR), "Rp 3,074/克（约 ¥{:.2f}/克）".format(3074 * CNY_PER_IDR)]
    ], columns=["产品类别", "最低零售价 HJE", "消费税"])
    show_df(tax)
    st.caption("以上为 PMK 96/2024 附件列示、2025-01-01 起适用的最低零售价及消费税单位。HJE 与消费税计量单位不同，不能将封闭系统两列数字直接按同一单位比较。")
    
    st.markdown('<div class="warning-box"><b>税收口径提醒：</b>本次确认到 PMK 96/2024 的官方法规文本及其 2025 年适用税率。未将"2026 年税率是否调整"写成确定结论；正式报价、进口成本和税标核算前，应由印尼进口商/税务顾问再次确认现行税则、税标及地方税适用范围。</div>', unsafe_allow_html=True)
    
    section("关键法规时间线")
    timeline = pd.DataFrame([
        ["2024-07-26", "PP 28/2024 颁布", "已生效；部分条款设过渡期"],
        ["2025-01-01", "PMK 96/2024 税率及 HJE 标准开始适用", "已实施"],
        ["2025-07-03", "PerBPOM 18/2025 生效", "已实施"],
        ["2025-08-04", "Permenko PMK 2/2025 生效，建立尼古丁/焦油上限协调程序", "已实施；具体限值另行确定"],
        ["2026-03-10", "Kemenko PMK 就尼古丁/焦油上限研究征求意见", "研究/公众意见阶段"],
        ["2026-06-05", "卫生部宣布起草标准化包装部长条例", "草案阶段"],
        ["2026-07-26", "部分产品、检测、许可、标签和警示要求实施节点", "过渡期结束"],
        ["2026-09-09", "BNN 提出研究通过总统条例全面禁止电子烟", "部门提议"],
        ["2026-09-14", "Kepmenkes HK.01.07/MENKES/963/2026 发布线上广告监督指南", "已生效"],
    ], columns=["日期", "事项", "状态"])
    show_df(timeline)
    
    st.markdown('<div class="warning-box"><b>业务合规重点：</b>包装容量需按产品系统分类核验；线上销售和数字媒体广告需分别审查平台、年龄验证与内容规则；上市前准备产品许可、尼古丁检测、健康警示、实际包装和税标资料。当前样本只用于识别复核事项，不构成违法判断。</div>', unsafe_allow_html=True)

# ============================== TAB 8 总结洞察 ==============================
with tabs[8]:
    st.markdown("## 📝 印尼市场总结与业务参考")
    total = len(fdf)
    category_counts = fdf["分类展示"].value_counts()
    topbrands = top_n_ties(fdf["品牌"].value_counts(), 5)
    topfl = top_n_ties(fdf["口味_key"].value_counts(), 3)
    names = fdf.drop_duplicates("口味_key").set_index("口味_key")["口味名称"].to_dict()
    topfl_txt = "、".join(f"{names.get(k, k)}（{int(v)}条）" for k, v in topfl.items()) or "暂无可用数据"
    brand_txt = "、".join(f"{b}（{int(n)}条）" for b, n in topbrands.items()) or "暂无可用数据"
    local_n = int((fdf["是否本地代理/本土品牌"] == "是").sum())
    local_pct = local_n / total if total else 0
    cat_txt = "、".join(f"{k}（{int(v)}条，{v/total:.1%}）" for k, v in category_counts.items())
    product_mix = fdf["产品类型"].value_counts()
    product_mix_txt = "、".join(f"{k} {int(v)}条（{v/total:.1%}）" for k, v in product_mix.items())

    normalized_tag_rows = [list(dict.fromkeys(row)) for row in fdf["口味标签列表"]]
    normalized_tag_counts = Counter(t for row in normalized_tag_rows for t in row)
    normalized_tag_top = top_counter(normalized_tag_counts, 6, include_singletons=True)
    summary_tag_map = {}
    for norm_row, raw_row in zip(fdf["口味标签列表"], fdf["口味标签原文列表"]):
        for norm, raw in zip(norm_row, raw_row):
            summary_tag_map.setdefault(norm, raw)
    tag_txt = "、".join(f"{summary_tag_map.get(k, k)}（{int(v)}条）" for k, v in normalized_tag_top.items()) or "暂无可用标签"
    ice_tag_n = int(fdf["含冰_标签判断"].sum())
    ice_flag_n = int((fdf["含冰/薄荷"] == "是").sum())
    double_n = int((fdf["口味标签数"] == 2).sum())
    multi3_n = int((fdf["口味标签数"] >= 3).sum())
    multi_n = double_n + multi3_n
    combo_counts = Counter()
    for row in normalized_tag_rows:
        u = sorted(set(row))
        for i in range(len(u)):
            for j in range(i + 1, len(u)):
                combo_counts[(u[i], u[j])] += 1
    combo_series = pd.Series(combo_counts, dtype="int64") if combo_counts else pd.Series(dtype="int64")
    combo_top = top_counter(combo_series, 5, include_singletons=False)
    combo_short = "、".join(f"{summary_tag_map.get(a, a)} + {summary_tag_map.get(b, b)}（{int(v)}条）" for (a, b), v in combo_top.items()) or "暂无达到重复门槛的组合"

    st.markdown(f'<div class="summary-card"><b>市场背景</b>ECigIntelligence 资料显示，印尼电子烟市场仍有增长；该资料属于外部市场估计，不能直接转换为本次零售榜单的销量或产能目标。</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="summary-card"><b>产品结构</b>当前筛选样本共 {total} 条，主要产品类型为 {product_mix_txt}。该结构反映本次热门榜单样本，不代表整体市场份额。</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="summary-card"><b>口味偏好</b>高频分类为 {cat_txt}；重复频次≥2的口味包括 {topfl_txt}。高频元素包括 {tag_txt}。多元素组合占比 {pct(multi_n, total)}（双元素 {pct(double_n, total)}；三种及以上 {pct(multi3_n, total)}）。标签侧识别到含冰/薄荷 {ice_tag_n} 条，原始字段标记“是”为 {ice_flag_n} 条。组合示例：{combo_short}。</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="summary-card"><b>品牌与合作线索</b>热销样本中上榜记录较多的品牌包括 {brand_txt}。本土品牌/本地代理标记为“是”的记录占 {local_pct:.1%}（{local_n}/{total} 条）。该结果适合用于潜客优先级排序，不等同于品牌市场份额。</div>', unsafe_allow_html=True)
    st.markdown('<div class="danger-box"><b>合规前置</b>包装容量需按开放式、封闭式和一次性系统分类核验；线上销售和数字媒体广告需分别审查年龄验证、平台属性与内容规则；上市前准备产品许可、尼古丁检测、标签警示、实际包装和税标资料。当前样本只用于识别复核事项，不构成违法判断。</div>', unsafe_allow_html=True)
    st.markdown('<div class="action-box"><b>开发建议</b><ul><li>围绕水果、饮料/茶和甜点类口味，可将重点放在“凉味剂”上，测试无冰、轻冰和强冰版本。</li><li>优先验证高频元素组合，不把榜单频次直接等同于销量预测。</li><li>报价时根据产品类型、尼古丁类型、容量和渠道分别核算。</li><li>打样前确认配方、包装、检测、标签、税标和流通文件。</li></ul></div>', unsafe_allow_html=True)

    section("数据范围与边界说明")
    st.markdown(f'<div class="summary-card"><b>统一说明</b><ul><li><b>热销榜：</b>当前 Excel 包含 {len(pop_all)} 条记录，覆盖 {pop_all["网站名称"].nunique()} 个线上零售网站；记录数反映榜单采集结果，不代表销量或市场份额。</li><li><b>品牌烟油库：</b>共 {len(brand_all)} 条品牌在售口味记录，用于观察采集到的产品覆盖，不代表品牌完整 SKU。</li><li><b>口味元素：</b>根据 Excel 中已有标签统计；同一条记录中同时出现的标签用于组合分析。</li><li><b>市场背景：</b>市场规模、用户数、渠道和设备结构来自 ECigIntelligence 外部调研资料，不与本次零售榜单直接合并计算。</li><li><b>法规：</b>法规内容用于识别商业落地前的复核事项，最终应由印尼当地法律、税务和产品合规顾问确认。</li></ul><div>以上数据仅供业务分析和销售沟通参考，不构成销量预测、市场份额结论或合规承诺。</div></div>', unsafe_allow_html=True)

    st.markdown("---")
