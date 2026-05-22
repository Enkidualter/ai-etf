"""ETF Investment Assistant"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Generator, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

# ── Constants ─────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).resolve().parent
EXCEL_PATH    = BASE_DIR / "上证ETF(1).xlsx"
CACHE_DIR     = BASE_DIR / "data" / "cache"
PROFILE_CACHE = CACHE_DIR / "etf_sample_profile.csv"
PRICE_CACHE   = CACHE_DIR / "etf_sample_prices.csv"
SAMPLE_SIZE   = 10
CST           = "Asia/Shanghai"
SIDEBAR_W     = 252   # px

DEEPSEEK_BASE_URL = os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com").rstrip("/")
DEEPSEEK_MODEL    = os.environ.get("DEEPSEEK_MODEL",    "deepseek-chat")

AI_PROVIDERS: Dict[str, Dict[str, str]] = {
    "DeepSeek":        {"base_url": "https://api.deepseek.com",                               "default_model": "deepseek-chat"},
    "Gemini (Google)": {"base_url": "https://generativelanguage.googleapis.com/v1beta/openai", "default_model": "gemini-2.0-flash"},
    "OpenAI":          {"base_url": "https://api.openai.com/v1",                              "default_model": "gpt-4o-mini"},
    "Moonshot (Kimi)": {"base_url": "https://api.moonshot.cn/v1",                             "default_model": "moonshot-v1-8k"},
    "智谱 GLM":         {"base_url": "https://open.bigmodel.cn/api/paas/v4",                   "default_model": "glm-4-flash"},
    "硅基流动":         {"base_url": "https://api.siliconflow.cn/v1",                          "default_model": "Qwen/Qwen2.5-7B-Instruct"},
}

HELP_TEXT = {
    "宽基":    "宽基ETF通常跟踪沪深300、上证50、中证500这类大盘综合指数，买的是一篮子股票，适合长期配置的核心仓位。",
    "红利":    "红利ETF偏向高分红、现金流较稳定的公司，常被保守投资者关注，但也受利率、周期和行业集中影响。",
    "行业主题":"行业主题ETF集中投向某个行业或主题，例如消费、金融、半导体，机会集中，波动和回撤通常更大。",
    "回撤":    "回撤可理解为从阶段高点跌下多少。最大回撤-30%意味着从最高点买入会一度亏约30%。",
    "波动率":  "波动率衡量价格上下震荡的剧烈程度，更高的波动率更考验持有耐心和仓位控制。",
    "费率":    "费率是长期持有时持续付出的管理费、托管费，单年看不大，多年复利后会显著影响实际收益。",
    "规模":    "规模越大流动性和稳定性通常越好，规模很小的ETF需额外关注成交不活跃或清盘风险。",
    "流动性":  "流动性指买卖是否方便、价差是否大，流动性差时买入或卖出可能付出额外隐性成本。",
}


# ── CSS ───────────────────────────────────────────────────────────────────────
def build_css(dark: bool) -> str:
    if dark:
        v = dict(
            bg_page="#080d18", bg_sidebar="#0b1120", bg_card="#0e1628",
            bg_card2="#111c30", border_s="#192236", border_c="#1e2d45",
            text_h="#e8f0ff", text_p="#c0d0e8", text_sec="#8a9bb8",
            text_m="#4e6080", text_d="#2e3f58",
            accent="#4f8ef7", accent_btn="#2563eb",
            card_shadow="none",
            badge_bg="rgba(79,142,247,0.12)", badge_bdr="rgba(79,142,247,0.3)",
            chart_bg="#0e1628", chart_paper="#0e1628",
            chart_font="#4e6080", chart_grid="#192236",
            chart_line="#4f8ef7", chart_fill="rgba(79,142,247,0.08)",
        )
    else:
        v = dict(
            bg_page="#f5f6f8", bg_sidebar="#ffffff", bg_card="#ffffff",
            bg_card2="#f9fafb", border_s="#f0f1f3", border_c="#e4e6ea",
            text_h="#111827", text_p="#1f2937", text_sec="#4b5563",
            text_m="#9ca3af", text_d="#d1d5db",
            accent="#2563eb", accent_btn="#2563eb",
            card_shadow="0 1px 3px rgba(0,0,0,0.07),0 1px 2px rgba(0,0,0,0.04)",
            badge_bg="rgba(37,99,235,0.07)", badge_bdr="rgba(37,99,235,0.2)",
            chart_bg="#ffffff", chart_paper="#f5f6f8",
            chart_font="#9ca3af", chart_grid="#f0f1f3",
            chart_line="#2563eb", chart_fill="rgba(37,99,235,0.05)",
        )

    return f"""<style>
/* ── Always-open sidebar: hide all collapse/expand controls ── */
[data-testid="stSidebarCollapseButton"],
[data-testid="collapsedControl"],
button[title*="sidebar"],
button[aria-label*="sidebar"],
button[aria-label*="Sidebar"] {{
    display: none !important;
}}

/* ── Sidebar fixed width ── */
section[data-testid="stSidebar"] {{
    width: {SIDEBAR_W}px !important;
    min-width: {SIDEBAR_W}px !important;
    background: {v['bg_sidebar']} !important;
    border-right: 1px solid {v['border_s']} !important;
}}
section[data-testid="stSidebar"] > div:first-child {{
    width: {SIDEBAR_W}px !important;
    padding: 1.5rem 1rem !important;
}}

/* ── Hide chrome ── */
#MainMenu, footer, [data-testid="stToolbar"],
[data-testid="stDecoration"] {{
    display: none !important;
}}
header[data-testid="stHeader"] {{
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
    height: 0 !important;
    min-height: 0 !important;
}}

/* ── App background ── */
.stApp {{
    background: {v['bg_page']} !important;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", system-ui, sans-serif !important;
}}

/* ── Main content padding ── */
.main .block-container {{
    padding-top: 2rem !important;
    max-width: 100% !important;
}}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {{
    gap: 0;
    border-bottom: 1px solid {v['border_s']};
    background: transparent;
    padding: 0;
    margin-bottom: 0;
}}
.stTabs [data-baseweb="tab"] {{
    background: transparent !important;
    color: {v['text_m']};
    height: 40px;
    padding: 0 18px;
    font-size: 0.875rem;
    border-radius: 0;
    border-bottom: 2px solid transparent;
    margin-bottom: -1px;
    transition: color 0.15s ease;
    font-family: inherit !important;
}}
.stTabs [aria-selected="true"] {{
    color: {v['accent']} !important;
    border-bottom-color: {v['accent']} !important;
    background: transparent !important;
}}
.stTabs [data-baseweb="tab-panel"] {{ padding-top: 24px; }}

/* ── Buttons ── */
div[data-testid="stButton"] > button {{
    background: {v['bg_card']};
    color: {v['text_sec']};
    border: 1px solid {v['border_c']};
    border-radius: 8px;
    font-size: 0.875rem;
    font-family: inherit !important;
    transition: all 0.15s cubic-bezier(0.4,0,0.2,1);
}}
div[data-testid="stButton"] > button:hover {{
    border-color: {v['accent']};
    color: {v['accent']};
    transform: translateY(-1px);
}}
div[data-testid="stButton"] > button[kind="primary"] {{
    background: {v['accent_btn']};
    color: #fff;
    border: none;
    font-weight: 600;
    box-shadow: 0 1px 4px rgba(37,99,235,0.3);
}}
div[data-testid="stButton"] > button[kind="primary"]:hover {{
    background: #1d4ed8;
    color: #fff;
    transform: translateY(-1px);
    box-shadow: 0 4px 14px rgba(37,99,235,0.4);
}}

/* ── Metrics ── */
[data-testid="stMetricValue"] {{
    font-size: 1.4rem !important;
    font-weight: 700 !important;
    color: {v['text_h']} !important;
}}
[data-testid="stMetricLabel"] {{
    font-size: 0.7rem !important;
    color: {v['text_m']} !important;
    text-transform: uppercase;
    letter-spacing: 0.06em;
}}

/* ── Inputs ── */
.stSelectbox > div > div {{
    background: {v['bg_card']} !important;
    border-color: {v['border_c']} !important;
    border-radius: 8px !important;
    color: {v['text_p']} !important;
}}
.stTextInput > div > div > input,
[data-testid="stNumberInput"] input {{
    background: {v['bg_card']} !important;
    border-color: {v['border_c']} !important;
    border-radius: 8px !important;
    color: {v['text_p']} !important;
}}
.stSelectbox label, .stTextInput label,
[data-testid="stNumberInput"] label {{
    color: {v['text_sec']} !important;
    font-size: 0.82rem !important;
}}

/* ── Toggle ── */
[data-testid="stToggle"] span {{ background: {v['border_c']} !important; }}
[data-testid="stToggle"] input:checked + span {{
    background: {v['accent_btn']} !important;
}}

/* ── Expander ── */
details > summary {{
    background: {v['bg_card']} !important;
    border: 1px solid {v['border_c']} !important;
    border-radius: 8px !important;
    padding: 10px 16px !important;
    color: {v['text_sec']} !important;
    font-size: 0.875rem !important;
    cursor: pointer;
    transition: border-color 0.15s;
}}
details > summary:hover {{ border-color: {v['accent']} !important; }}
details[open] > summary {{ border-radius: 8px 8px 0 0 !important; }}
details > div {{
    background: {v['bg_card2']} !important;
    border: 1px solid {v['border_c']} !important;
    border-top: none !important;
    border-radius: 0 0 8px 8px !important;
    padding: 16px !important;
}}

/* ── Misc ── */
.stAlert {{ border-radius: 10px !important; }}
hr {{ border-color: {v['border_s']} !important; margin: 1rem 0 !important; }}
[data-testid="stDataFrame"] {{ border-radius: 10px; overflow: hidden; }}
h1, h2, h3, h4 {{ color: {v['text_h']} !important; font-family: inherit !important; }}
p, [data-testid="stMarkdownContainer"] p {{
    color: {v['text_sec']} !important;
    font-family: inherit !important;
}}
[data-testid="stCaption"] {{ color: {v['text_m']} !important; }}
.stMarkdown code {{ background: {v['bg_card2']} !important; border-radius: 4px; }}

/* ── Card ── */
.card {{
    background: {v['bg_card']};
    border: 1px solid {v['border_c']};
    border-radius: 12px;
    box-shadow: {v['card_shadow']};
    transition: transform 0.18s cubic-bezier(0.4,0,0.2,1),
                box-shadow 0.18s ease;
}}
.card:hover {{ transform: translateY(-2px); }}
.card--accent {{ border-color: {v['accent_btn']} !important; }}

/* ETF card internals */
.etf-name {{ font-size: 0.95rem; font-weight: 600; color: {v['text_h']}; margin-bottom: 4px; }}
.etf-meta {{ color: {v['text_m']}; font-size: 0.75rem; display: flex; align-items: center; gap: 5px; flex-wrap: wrap; }}
.etf-score {{
    padding: 3px 13px; border-radius: 20px;
    font-weight: 700; font-size: 0.9rem;
    white-space: nowrap; flex-shrink: 0; margin-left: 10px;
}}
.etf-grid {{
    display: grid; grid-template-columns: repeat(5,1fr);
    gap: 8px; border-top: 1px solid {v['border_s']};
    padding-top: 12px; margin-top: 12px;
}}
.etf-lbl {{ color: {v['text_d']}; font-size: 0.62rem; text-transform: uppercase; letter-spacing: .05em; margin-bottom: 3px; }}
.etf-val {{ color: {v['text_p']}; font-size: 0.82rem; font-weight: 500; }}

/* Match tag */
.mtag {{
    display: inline-flex; align-items: center;
    background: {v['badge_bg']}; color: {v['accent']};
    padding: 1px 7px; border-radius: 99px;
    font-size: 0.6rem; font-weight: 600;
    border: 1px solid {v['badge_bdr']};
}}

/* Domain card */
.domain-card {{ padding: 18px; height: 100%; }}

/* Section title */
.stitle {{ margin: 0 0 20px 0; }}
.stitle-h {{ margin: 0; font-size: 1.15rem; font-weight: 700; color: {v['text_h']}; letter-spacing: -0.01em; }}
.stitle-s {{ margin: 4px 0 0; font-size: 0.8rem; color: {v['text_m']}; }}

/* Sidebar labels */
.sb-label {{
    color: {v['text_m']}; font-size: 0.65rem;
    text-transform: uppercase; letter-spacing: .08em;
    margin-bottom: 6px; margin-top: 2px;
}}

/* Risk flag row */
.flag-row {{ display: flex; gap: 8px; align-items: flex-start; margin-bottom: 8px; }}
.flag-icon {{ color: #f59e0b; flex-shrink: 0; margin-top: 1px; }}
.flag-text {{ color: {v['text_sec']}; font-size: 0.84rem; line-height: 1.5; }}

/* Score bar */
.sbar-row {{ margin-bottom: 12px; }}
.sbar-hd {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px; }}
.sbar-lbl {{ color: {v['text_sec']}; font-size: 0.82rem; }}
.sbar-num {{ color: {v['text_p']}; font-size: 0.78rem; font-weight: 600; }}
.sbar-num span {{ color: {v['text_d']}; font-weight: 400; }}
.sbar-track {{ background: {v['border_s']}; border-radius: 99px; height: 5px; overflow: hidden; }}
.sbar-fill {{ height: 100%; border-radius: 99px; }}

/* Term card */
.term-card {{
    background: {v['bg_card']}; border: 1px solid {v['border_c']};
    border-radius: 10px; padding: 13px 15px; margin-bottom: 8px;
    transition: border-color 0.15s ease, transform 0.15s ease;
}}
.term-card:hover {{ border-color: {v['accent']}; transform: translateX(3px); }}
.term-title {{ color: {v['accent']}; font-size: 0.8rem; font-weight: 600; margin-bottom: 4px; }}
.term-desc  {{ color: {v['text_m']}; font-size: 0.8rem; line-height: 1.55; }}
</style>

<!-- Store chart colors for Python to read -->
<script>
window._chartColors = {{
    plot_bg:    "{v['chart_bg']}",
    paper_bg:   "{v['chart_paper']}",
    font_color: "{v['chart_font']}",
    grid_color: "{v['chart_grid']}",
    line_color: "{v['chart_line']}",
    fill_color: "{v['chart_fill']}",
}};
</script>"""


# ── Data classes ──────────────────────────────────────────────────────────────
@dataclass
class UserProfile:
    experience: str; risk_level: str; horizon: str
    goal: str; preference: str; amount: float


@dataclass
class DomainRecommendation:
    name: str; match_score: float; role: str
    plain_intro: str; suitable_for: str; return_source: str
    key_risks: str; common_mistake: str
    category_keywords: Tuple[str, ...]; risk_hint: str


DOMAIN_LIBRARY: Dict[str, DomainRecommendation] = {
    "宽基核心": DomainRecommendation(
        name="宽基核心", match_score=0, role="核心仓位",
        plain_intro="先用一篮子代表性股票打底，不把胜负押在单一行业上。",
        suitable_for="适合新手、长期配置、希望先把组合底座搭稳的投资者。",
        return_source="主要来自市场整体增长、指数成分股盈利改善和估值修复。",
        key_risks="市场整体下跌时也会回撤，不能理解成保本工具。",
        common_mistake="只看短期涨幅排名，忽略宽基更适合长期、分批和纪律化持有。",
        category_keywords=("宽基","核心"), risk_hint="适合作为组合底座，但仍要控制买入节奏和总仓位。",
    ),
    "红利/低波": DomainRecommendation(
        name="红利/低波", match_score=0, role="稳健核心或防守仓位",
        plain_intro="更关注分红、现金流和相对稳健的公司，少刺激，多纪律。",
        suitable_for="适合不喜欢大起大落、希望降低组合波动的新手或保守型投资者。",
        return_source="主要来自股息、企业稳定盈利和低估值修复。",
        key_risks="红利不等于无风险，可能有行业集中、周期下行和股息下降风险。",
        common_mistake="把高股息率简单理解成高收益，忽略股价下跌可能抵消分红。",
        category_keywords=("红利","低波","价值","金融"),
        risk_hint="当前样本池红利ETF不足，demo会用价值/金融类作为近似候选展示。",
    ),
    "行业主题": DomainRecommendation(
        name="行业主题", match_score=0, role="卫星仓位",
        plain_intro="集中投向某个行业或主题，适合表达观点，不适合新手重仓。",
        suitable_for="适合能承受较大波动、理解行业周期、愿意控制在小比例仓位的投资者。",
        return_source="主要来自行业景气度上行、政策催化、估值扩张或盈利改善。",
        key_risks="行业热度高时容易追涨，景气反转时回撤也可能很深。",
        common_mistake="因为最近涨得多就买入，却没想清楚退出规则和仓位上限。",
        category_keywords=("行业","主题","成长"),
        risk_hint="更适合作为小比例卫星仓位，不建议保守投资者作核心仓位。",
    ),
}


# ── Data loading ──────────────────────────────────────────────────────────────
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns={
        "证券代码":"code","证券名称":"name","基金类型":"fund_type",
        "单位净值\n[交易日期] 最新[单位]元":"nav",
        "基金份额\n[交易日期] 最新\n[单位]  份":"shares",
        "投资风格\n[年度] 2023\n[报告期] 中报":"style",
        "基金管理人":"manager","基金托管人":"custodian",
        "管理费率[单位]%":"management_fee","托管费率[单位]%":"custody_fee",
        "基金经理(现任)":"fund_manager","基金存续期[单位]年":"duration_years",
        "基金成立日":"inception_date","基金到期日":"maturity_date",
    })

def parse_inception(value: Any) -> Optional[pd.Timestamp]:
    if pd.isna(value): return None
    text = str(value).split(".")[0].strip()
    if len(text) != 8: return None
    t = pd.to_datetime(text, format="%Y%m%d", errors="coerce")
    return None if pd.isna(t) else t

def classify_etf(row: pd.Series) -> str:
    t = f"{row.get('name','')} {row.get('style','')}"
    if any(k in t for k in ["沪深300","上证50","中证500","国企"]): return "宽基/核心"
    if any(k in t for k in ["消费","金融"]): return "行业/主题"
    if "成长" in t: return "成长风格"
    if "价值" in t: return "价值风格"
    return "其他"

def load_profile_from_excel() -> pd.DataFrame:
    df = pd.read_excel(EXCEL_PATH).head(SAMPLE_SIZE)
    df = normalize_columns(df)
    keep = ["code","name","fund_type","nav","shares","style","manager",
            "custodian","management_fee","custody_fee","fund_manager","inception_date"]
    df = df[keep].copy()
    df["inception"]   = df["inception_date"].apply(parse_inception)
    df["age_years"]   = df["inception"].apply(
        lambda x: np.nan if x is None else max((pd.Timestamp.today()-x).days/365.25, 0))
    df["total_fee"]   = (pd.to_numeric(df["management_fee"], errors="coerce").fillna(0)
                       + pd.to_numeric(df["custody_fee"],    errors="coerce").fillna(0))
    df["scale_proxy"] = (pd.to_numeric(df["nav"],    errors="coerce")
                       * pd.to_numeric(df["shares"], errors="coerce"))
    df["category"]    = df.apply(classify_etf, axis=1)
    return df

def load_or_create_profile(refresh: bool = False) -> pd.DataFrame:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    if PROFILE_CACHE.exists() and not refresh:
        df = pd.read_csv(PROFILE_CACHE)
        if "inception" in df.columns:
            df["inception"] = pd.to_datetime(df["inception"], errors="coerce")
        if "category" not in df.columns:
            df["category"] = df.apply(classify_etf, axis=1)
        return df
    df = load_profile_from_excel()
    df.to_csv(PROFILE_CACHE, index=False, encoding="utf-8-sig")
    return df

def fetch_prices_from_yfinance(
    codes: Iterable[str], progress_placeholder: Any = None
) -> Tuple[pd.DataFrame, Optional[str]]:
    try:
        import yfinance as yf
    except ImportError:
        return pd.DataFrame(), "未安装 yfinance。"
    code_list = list(codes); frames: List[pd.DataFrame] = []; errors: List[str] = []
    for i, code in enumerate(code_list):
        if progress_placeholder:
            progress_placeholder.info(f"正在拉取 {i+1}/{len(code_list)}：{code} …")
        try:
            clean  = str(code).split(".")[0]
            suffix = str(code).split(".")[-1].upper() if "." in str(code) else "SH"
            hist   = yf.Ticker(clean + (".SS" if suffix=="SH" else ".SZ")).history(period="1y")
            if hist.empty: errors.append(f"{code}: 无数据"); continue
            frames.append(pd.DataFrame({
                "code": code,
                "date": pd.to_datetime(hist.index.tz_localize(None)),
                "close": hist["Close"].values,
            }).dropna(subset=["date","close"]).sort_values("date"))
        except Exception as e:
            errors.append(f"{code}: {type(e).__name__}"); continue
    if not frames: return pd.DataFrame(), f"拉取失败：{'; '.join(errors[:3])}"
    prices = pd.concat(frames, ignore_index=True)
    prices["fetch_time"] = pd.Timestamp.now(tz="UTC").tz_convert(CST).strftime("%Y-%m-%d %H:%M:%S")
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    prices.to_csv(PRICE_CACHE, index=False, encoding="utf-8-sig")
    return prices, ("; ".join(errors) if errors else None)

def load_prices(
    codes: Iterable[str], refresh: bool = False, progress_placeholder: Any = None
) -> Tuple[pd.DataFrame, Optional[str]]:
    if PRICE_CACHE.exists() and not refresh:
        prices = pd.read_csv(PRICE_CACHE)
        prices["date"] = pd.to_datetime(prices["date"], errors="coerce")
        return prices, None
    return fetch_prices_from_yfinance(codes, progress_placeholder=progress_placeholder)


# ── Metrics & scoring ─────────────────────────────────────────────────────────
def max_drawdown(c: pd.Series) -> float:
    c = c.dropna(); return np.nan if c.empty else float((c/c.cummax()-1).min())

def annual_volatility(c: pd.Series) -> float:
    r = c.pct_change().dropna(); return np.nan if r.empty else float(r.std()*np.sqrt(252))

def one_year_return(c: pd.Series) -> float:
    c = c.dropna(); return np.nan if len(c)<2 else float(c.iloc[-1]/c.iloc[0]-1)

def add_metrics(profile: pd.DataFrame, prices: pd.DataFrame) -> pd.DataFrame:
    df = profile.copy()
    rows = []
    for code in df["code"]:
        c = prices.loc[prices["code"]==code,"close"] if not prices.empty else pd.Series(dtype=float)
        rows.append({"code":code,"max_drawdown":max_drawdown(c),"volatility":annual_volatility(c),
                     "one_year_return":one_year_return(c),"price_points":int(c.dropna().shape[0])})
    df = df.merge(pd.DataFrame(rows), on="code", how="left")
    scores = pd.DataFrame([score_etf(r) for _,r in df.iterrows()])
    return pd.concat([df.reset_index(drop=True), scores], axis=1)

def score_size(v):
    if pd.isna(v): return 8
    if v>=50e9: return 25
    if v>=10e9: return 22
    if v>=1e9:  return 16
    if v>=1e8:  return 9
    return 4

def score_fee(v):
    if pd.isna(v): return 8
    if v<=0.20: return 20
    if v<=0.30: return 17
    if v<=0.60: return 12
    if v<=1.00: return 7
    return 3

def score_age(v):
    if pd.isna(v): return 6
    if v>=5: return 20
    if v>=3: return 16
    if v>=1: return 10
    if v>=0.5: return 6
    return 3

def score_risk(dd, vol):
    if pd.isna(dd) or pd.isna(vol): return 8
    s = 20
    if dd<-0.35: s-=8
    elif dd<-0.25: s-=5
    elif dd<-0.15: s-=2
    if vol>0.35: s-=7
    elif vol>0.25: s-=4
    elif vol>0.18: s-=2
    return max(s,3)

def score_data(pts, missing):
    s = 15
    if pts<120: s-=6
    elif pts<220: s-=3
    s -= min(missing*2, 6)
    return max(s,3)

def score_etf(row: pd.Series) -> Dict[str, float]:
    miss = sum(pd.isna(row.get(c)) for c in ["nav","shares","management_fee","custody_fee","inception"])
    sz=score_size(row.get("scale_proxy")); fe=score_fee(row.get("total_fee"))
    ag=score_age(row.get("age_years"));   ri=score_risk(row.get("max_drawdown"),row.get("volatility"))
    da=score_data(int(row.get("price_points",0) or 0), miss)
    return {"quality_score":round(sz+fe+ag+ri+da,1),
            "size_score":sz,"fee_score":fe,"age_score":ag,"risk_score":ri,"data_score":da}


# ── Recommendation logic ──────────────────────────────────────────────────────
def infer_risk_level(ans, exp, hor):
    if "不确定" in ans: return "低风险" if (exp=="我是新手" or hor=="1年以内") else "中风险"
    if "不太能接受" in ans: return "低风险"
    if "阶段性波动"  in ans: return "中风险"
    return "高风险"

def recommend_domains(user: UserProfile) -> List[DomainRecommendation]:
    s = {"宽基核心":50.0,"红利/低波":45.0,"行业主题":30.0}
    if user.risk_level=="低风险":  s["红利/低波"]+=24;s["宽基核心"]+=18;s["行业主题"]-=20
    elif user.risk_level=="中风险": s["宽基核心"]+=20;s["红利/低波"]+=10;s["行业主题"]+=4
    else:                           s["行业主题"]+=24;s["宽基核心"]+=10
    if user.experience=="我是新手":                  s["宽基核心"]+=14;s["红利/低波"]+=8;s["行业主题"]-=8
    if user.horizon=="1年以内":                      s["红利/低波"]+=12;s["行业主题"]-=12
    elif user.horizon=="3年以上":                    s["宽基核心"]+=10;s["行业主题"]+=3
    if user.goal=="希望稳一点，少一些大起大落":       s["红利/低波"]+=20;s["行业主题"]-=10
    elif user.goal=="长期配置，慢慢积累":             s["宽基核心"]+=20
    elif user.goal=="想抓住某些行业机会":             s["行业主题"]+=20
    if user.preference=="我想先要一个稳一点的底座":   s["宽基核心"]+=16
    elif user.preference=="我更关注分红和稳健":       s["红利/低波"]+=16
    elif user.preference=="我想了解行业主题机会":     s["行业主题"]+=16
    result = [DomainRecommendation(**{**b.__dict__,"match_score":round(s[n],1)})
              for n,b in DOMAIN_LIBRARY.items()]
    return sorted(result, key=lambda x: x.match_score, reverse=True)

def etf_matches_domain(row, domain):
    t = f"{row.get('name','')} {row.get('style','')} {row.get('category','')}"
    return any(k in t for k in domain.category_keywords)

def candidate_score(row, user, domain):
    sc = float(row["quality_score"])
    if etf_matches_domain(row,domain): sc+=18
    if user.risk_level=="低风险":
        if row.get("total_fee",99)<=0.20: sc+=8
        if pd.notna(row.get("max_drawdown")) and row["max_drawdown"]<-0.25: sc-=12
        if pd.notna(row.get("volatility"))   and row["volatility"]>0.25:    sc-=10
        if "行业" in str(row.get("category","")): sc-=6
    elif user.risk_level=="高风险" and "行业" in str(row.get("category","")): sc+=5
    if user.horizon=="1年以内" and pd.notna(row.get("volatility")) and row["volatility"]>0.25: sc-=8
    return round(sc,1)

def rank_candidates(df, user, domain):
    r = df.copy()
    r["candidate_score"] = r.apply(lambda x: candidate_score(x,user,domain), axis=1)
    r["domain_match"]    = r.apply(lambda x: etf_matches_domain(x,domain), axis=1)
    return r.sort_values(["domain_match","candidate_score"],ascending=[False,False]).head(3)

def risk_flags(row: pd.Series) -> List[str]:
    f = []
    if row.get("age_years",99)<1: f.append("成立时间较短，成立以来收益率参考价值有限。")
    if row.get("total_fee",0)>0.60: f.append("费率偏高，长期持有成本会持续累积。")
    if row.get("scale_proxy",np.inf)<1e9: f.append("规模偏小，需关注流动性和清盘风险。")
    if pd.notna(row.get("max_drawdown")) and row["max_drawdown"]<-0.25: f.append("近一年最大回撤较大，需确认能否承受。")
    if pd.notna(row.get("volatility"))   and row["volatility"]>0.25:    f.append("近一年波动率较高，不适合保守型核心仓位。")
    return f or ["未发现突出单项风险，但仍需控制仓位。"]

def self_check_questions(row, domain=None):
    cat = domain.name if domain else row.get("category","这类ETF")
    return [
        f'我是否理解这只ETF为什么归到"{cat}"，主要风险来源是什么？',
        "如果短期回撤20%-30%，我是否还能按原计划持有？",
        "我买入它是长期配置需要，还是被近期涨幅或热度吸引？",
        "我是否比较过同类ETF的规模、费率、成立时间和流动性？",
        "这只ETF在我的总资产中应该是核心仓位还是小比例卫星仓位？",
    ]

def build_fallback_narrative(user, domain, candidates):
    names = "、".join(candidates["name"].head(3).tolist()) if not candidates.empty else "暂无合适候选"
    return (f'根据你的偏好，当前更适合先从"{domain.name}"看起。'
            f"定位：{domain.role}——{domain.plain_intro}"
            f"可先重点比较：{names}。推荐只是缩小观察范围，不是买入建议。")


# ── AI functions ──────────────────────────────────────────────────────────────
def ai_available(): return bool(st.session_state.get("ai_api_key") or os.environ.get("DEEPSEEK_API_KEY"))

def get_ai_config():
    k = st.session_state.get("ai_api_key") or os.environ.get("DEEPSEEK_API_KEY","")
    u = (st.session_state.get("ai_base_url") or DEEPSEEK_BASE_URL).rstrip("/")
    m = st.session_state.get("ai_model") or DEEPSEEK_MODEL
    return k, u, m

def to_jsonable(v):
    if isinstance(v,dict):  return {str(k):to_jsonable(x) for k,x in v.items()}
    if isinstance(v,(list,tuple)): return [to_jsonable(x) for x in v]
    if isinstance(v,np.integer):   return int(v)
    if isinstance(v,np.floating):  return None if pd.isna(v) else float(v)
    if isinstance(v,pd.Timestamp): return v.strftime("%Y-%m-%d")
    if isinstance(v,float) and pd.isna(v): return None
    return v

def stream_ai_analysis(row, user=None, domain=None) -> Generator[str,None,None]:
    api_key, base_url, model = get_ai_config()
    if not api_key: yield "请先在左侧侧边栏填写 AI 服务商和 API Key。"; return
    ctx: Dict[str,Any] = {
        "etf_name":row.get("name"),"etf_code":row.get("code"),
        "category":row.get("category"),"quality_score":row.get("quality_score"),
        "total_fee_pct":row.get("total_fee"),"age_years":row.get("age_years"),
        "max_drawdown":row.get("max_drawdown"),"volatility":row.get("volatility"),
        "one_year_return":row.get("one_year_return"),"scale":row.get("scale_proxy"),
        "risk_flags":risk_flags(row),
    }
    if user:   ctx.update({"user_risk":user.risk_level,"horizon":user.horizon,"exp":user.experience})
    if domain: ctx["domain"] = domain.name
    body = json.dumps({
        "model": model,
        "messages": [
            {"role":"system","content":"你是谨慎的ETF投资教育助手。只基于提供数据分析，不编造数据，不预测收益，不给买卖建议。中文，面向新手。"},
            {"role":"user","content":"对以下ETF做三部分Markdown分析：\n**1. 综合质量判断**（约80字）\n**2. 主要风险提示**（约80字，结合用户风险偏好）\n**3. 购前自查清单**（3-5条）\n\n数据："+json.dumps(to_jsonable(ctx),ensure_ascii=False)},
        ],
        "temperature":0.4,"max_tokens":700,"stream":True,
    }, ensure_ascii=False).encode("utf-8")
    try:
        with requests.post(f"{base_url}/chat/completions",
            headers={"Authorization":f"Bearer {api_key}","Content-Type":"application/json; charset=utf-8"},
            data=body, stream=True, timeout=30) as resp:
            if not resp.ok: yield f"API错误 {resp.status_code}"; return
            for line in resp.iter_lines():
                if not line or not line.startswith(b"data: "): continue
                chunk = line[6:]
                if chunk.strip()==b"[DONE]": return
                try:
                    d = json.loads(chunk.decode("utf-8"))["choices"][0]["delta"].get("content","")
                    if d: yield d
                except Exception: continue
    except Exception as e: yield f"\n\n（中断：{e}）"

def call_ai_narrative(payload):
    api_key, base_url, model = get_ai_config()
    if not api_key: return None,"未设置 API Key"
    try:
        body = json.dumps({
            "model":model,
            "messages":[
                {"role":"system","content":"谨慎的ETF投资教育助手，只基于事实，不编造、不预测、不给建议，中文，面向新手。"},
                {"role":"user","content":"把以下事实润色为展示解释，结构：1）为什么先看这个方向；2）候选ETF怎么比较；3）风险和自查提醒。180字以内。\n\n"+json.dumps(to_jsonable(payload),ensure_ascii=False)},
            ],
            "temperature":0.3,"max_tokens":500,
        }, ensure_ascii=False).encode("utf-8")
        resp = requests.post(f"{base_url}/chat/completions",
            headers={"Authorization":f"Bearer {api_key}","Content-Type":"application/json; charset=utf-8"},
            data=body, timeout=20)
        if not resp.ok: return None,f"HTTP {resp.status_code}"
        return json.loads(resp.content.decode("utf-8"))["choices"][0]["message"]["content"].strip(), None
    except Exception as e:
        return None, f"{st.session_state.get('ai_provider','AI')} 调用失败：{e}"

def make_narrative(user, domain, candidates):
    payload = {
        "user_profile":user.__dict__,
        "domain":{"name":domain.name,"role":domain.role,"intro":domain.plain_intro,
                  "suitable_for":domain.suitable_for,"key_risks":domain.key_risks},
        "candidates":[{"code":r["code"],"name":r["name"],"quality_score":r["quality_score"],
                       "fee":r.get("total_fee"),"risk_flags":risk_flags(r)}
                      for _,r in candidates.head(3).iterrows()],
    }
    text, warn = call_ai_narrative(payload)
    return (text,None) if text else (build_fallback_narrative(user,domain,candidates), warn)


# ── Formatting ────────────────────────────────────────────────────────────────
def fp(v):  return "—" if pd.isna(v) else f"{v*100:.2f}%"
def fm(v):
    if pd.isna(v): return "—"
    v = float(v)
    if v>=1e8:  return f"{v/1e8:.1f}亿"
    if v>=1e4:  return f"{v/1e4:.1f}万"
    return f"{v:.0f}"

def _cv():
    """Chart color values — read from session state."""
    dark = st.session_state.get("dark_mode", True)
    if dark:
        return {"bg":"#0e1628","paper":"#0e1628","font":"#4e6080",
                "grid":"#192236","line":"#4f8ef7","fill":"rgba(79,142,247,0.08)"}
    return {"bg":"#ffffff","paper":"#f5f6f8","font":"#9ca3af",
            "grid":"#f0f1f3","line":"#2563eb","fill":"rgba(37,99,235,0.05)"}


# ── UI helpers ────────────────────────────────────────────────────────────────
def stitle(title: str, sub: str = "") -> None:
    s = f'<p class="stitle-s">{sub}</p>' if sub else ""
    st.markdown(f'<div class="stitle"><p class="stitle-h">{title}</p>{s}</div>',
                unsafe_allow_html=True)

def etf_card(row: pd.Series, highlight: bool = False) -> None:
    score = float(row.get("quality_score",0))
    sc    = "#10b981" if score>=70 else "#f59e0b" if score>=50 else "#ef4444"
    cls   = 'card etf-card card--accent' if highlight else 'card etf-card'
    mtag  = '<span class="mtag">方向匹配</span>' if highlight else ""
    fee   = row.get("total_fee", float("nan"))
    dd    = row.get("max_drawdown", float("nan"))
    vol   = row.get("volatility",   float("nan"))
    age   = row.get("age_years",    float("nan"))
    scl   = row.get("scale_proxy",  float("nan"))
    ddc   = "#ef4444" if not pd.isna(dd) and dd<-0.2 else "var(--etf-val-color, inherit)"

    def cell(lbl, val, color=""):
        cs = f' style="color:{color}"' if color else ""
        return (f'<div><div class="etf-lbl">{lbl}</div>'
                f'<div class="etf-val"{cs}>{val}</div></div>')

    st.markdown(
        f'<div class="{cls}" style="padding:16px 20px;margin-bottom:10px">'
        f'<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:0">'
        f'<div style="min-width:0;flex:1">'
        f'  <div class="etf-name">{row.get("name","")}</div>'
        f'  <div class="etf-meta">'
        f'    <span>{row.get("code","")}</span>'
        f'    <span>·</span>'
        f'    <span>{row.get("category","")}</span>'
        f'    {mtag}'
        f'  </div>'
        f'</div>'
        f'<div class="etf-score" style="background:{sc}13;color:{sc};border:1px solid {sc}2e">'
        f'  {score:.0f}<span style="font-size:0.58rem;opacity:0.5;font-weight:400"> /100</span>'
        f'</div>'
        f'</div>'
        f'<div class="etf-grid">'
        f'{cell("费率", f"{fee:.2f}%" if not pd.isna(fee) else "—")}'
        f'{cell("最大回撤", fp(dd), ddc)}'
        f'{cell("年化波动", fp(vol))}'
        f'{cell("成立年限", f"{age:.1f}年" if not pd.isna(age) else "—")}'
        f'{cell("规模", fm(scl))}'
        f'</div></div>',
        unsafe_allow_html=True,
    )

def score_bars(row: pd.Series) -> None:
    items = [("规模",row["size_score"],25),("费率",row["fee_score"],20),
             ("成立年限",row["age_score"],20),("历史风险",row["risk_score"],20),
             ("数据完整性",row["data_score"],15)]
    total = float(row["quality_score"])
    tc    = "#10b981" if total>=70 else "#f59e0b" if total>=50 else "#ef4444"
    bars  = ""
    for lbl,sc,mx in items:
        pct = sc/mx*100
        col = "#10b981" if pct>=75 else "#4f8ef7" if pct>=50 else "#f59e0b"
        bars += (
            f'<div class="sbar-row">'
            f'<div class="sbar-hd">'
            f'<span class="sbar-lbl">{lbl}</span>'
            f'<span class="sbar-num">{sc:.0f}<span>/{mx}</span></span>'
            f'</div>'
            f'<div class="sbar-track">'
            f'<div class="sbar-fill" style="width:{pct:.0f}%;background:{col}"></div>'
            f'</div></div>'
        )
    st.markdown(
        f'<div class="card" style="padding:20px 22px">'
        f'<div style="display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:20px">'
        f'<div>'
        f'<div class="sbar-lbl" style="margin-bottom:4px">综合质量分</div>'
        f'<div style="color:{tc};font-size:2.4rem;font-weight:800;line-height:1">'
        f'{total:.0f}<span style="font-size:0.8rem;font-weight:400;opacity:0.5"> /100</span>'
        f'</div>'
        f'</div></div>'
        f'{bars}</div>',
        unsafe_allow_html=True,
    )

def price_chart_fig(price_df: pd.DataFrame) -> None:
    cv = _cv()
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=price_df["date"], y=price_df["close"], mode="lines",
        line=dict(color=cv["line"], width=2),
        fill="tozeroy", fillcolor=cv["fill"],
        hovertemplate="%{x|%Y-%m-%d}  <b>%{y:.3f}</b><extra></extra>",
    ))
    mn = price_df["close"].min()*0.99; mx = price_df["close"].max()*1.01
    fig.update_layout(
        plot_bgcolor=cv["bg"], paper_bgcolor=cv["paper"],
        font=dict(color=cv["font"], size=11, family="-apple-system,BlinkMacSystemFont,sans-serif"),
        height=240, margin=dict(l=42,r=8,t=6,b=6),
        xaxis=dict(showgrid=False, zeroline=False, showline=False,
                   tickfont=dict(size=10), tickformat="%m月"),
        yaxis=dict(showgrid=True, gridcolor=cv["grid"], gridwidth=0.5,
                   zeroline=False, showline=False, tickfont=dict(size=10), range=[mn,mx]),
        hovermode="x unified",
        hoverlabel=dict(bgcolor=cv["bg"], bordercolor=cv["grid"], font_size=11),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})


# ── Sidebar ───────────────────────────────────────────────────────────────────
def sidebar_controls() -> Tuple[bool, bool]:
    with st.sidebar:
        # Theme toggle
        st.toggle("🌙  深色模式", value=st.session_state.get("dark_mode", True), key="dark_mode")
        st.markdown("<hr style='margin:10px 0'>", unsafe_allow_html=True)

        st.markdown('<p class="sb-label">AI 解释引擎</p>', unsafe_allow_html=True)
        selected = st.selectbox("服务商", list(AI_PROVIDERS.keys()), index=0,
                                key="ai_provider", label_visibility="collapsed")
        cfg = AI_PROVIDERS[selected]

        key_val = st.text_input("API Key", type="password", placeholder="填入 API Key",
            value=st.session_state.get("ai_api_key", os.environ.get("DEEPSEEK_API_KEY","")),
            key="_ai_key_input")
        st.session_state["ai_api_key"] = key_val

        mdl_val = st.text_input("模型", placeholder=cfg["default_model"],
            value=st.session_state.get("ai_model","") or cfg["default_model"],
            key="_ai_model_input")
        st.session_state["ai_model"]    = mdl_val or cfg["default_model"]
        st.session_state["ai_base_url"] = cfg["base_url"]

        if ai_available():
            st.success(f"✓  {st.session_state['ai_model']}")
        else:
            st.caption("未填写 API Key，使用规则模板")

        st.markdown("<hr style='margin:10px 0'>", unsafe_allow_html=True)
        st.markdown('<p class="sb-label">行情数据</p>', unsafe_allow_html=True)
        rp = st.button("拉取最新行情", use_container_width=True)
        rr = st.button("重建静态缓存", use_container_width=True)
        st.caption("Yahoo Finance · 自动缓存")
        return rp, rr


# ── Tab: 偏好问卷 ─────────────────────────────────────────────────────────────
def render_profile_form() -> UserProfile:
    stitle("偏好问卷", "用自己的真实感受来选，不需要先懂专业术语")
    l, r = st.columns(2, gap="large")
    with l:
        exp  = st.selectbox("投资经验", ["我是新手","有一些经验","比较熟悉ETF"])
        risk = st.selectbox("账户短期亏损时，你的感受更接近？",
            ["我不太能接受亏损","能接受阶段性波动","愿意承受较大波动换取机会","我不确定"])
        hor  = st.selectbox("这笔钱预计多久不用？", ["1年以内","1-3年","3年以上"])
    with r:
        goal = st.selectbox("你更接近哪种目标？",
            ["希望稳一点，少一些大起大落","长期配置，慢慢积累","想抓住某些行业机会","还没想清楚"])
        pref = st.selectbox("你现在更想先了解什么？",
            ["我想先要一个稳一点的底座","我更关注分红和稳健","我想了解行业主题机会","我不确定"])
        amt  = st.number_input("计划投入金额（元）", min_value=1000.0, value=50000.0, step=1000.0)

    rl   = infer_risk_level(risk, exp, hor)
    user = UserProfile(exp, rl, hor, goal, pref, amt)
    st.session_state["user_profile"] = user

    icon   = {"低风险":"🟢","中风险":"🟡","高风险":"🔴"}.get(rl,"⚪")
    driver = f'由【亏损感受】决定：你选了「{risk}」' if "不确定" not in risk else '你选了"我不确定"，系统保守处理'
    st.info(f"{icon} **风险层级：{rl}**　{driver}\n\n不太能接受→低 / 阶段性波动→中 / 愿意承受→高")
    return user


# ── Tab: 方向推荐 ─────────────────────────────────────────────────────────────
def render_domain_recommendations(user: UserProfile) -> List[DomainRecommendation]:
    domains = recommend_domains(user)
    stitle("适合你的 ETF 方向", "先确定方向，再比较具体产品")
    cols   = st.columns(3, gap="medium")
    colors = ["#4f8ef7","#10b981","#f59e0b"]
    for col, d, clr in zip(cols, domains, colors):
        with col:
            st.markdown(
                f'<div class="card domain-card">'
                f'<div style="color:{clr};font-size:1.6rem;font-weight:800;line-height:1;margin-bottom:2px">{d.match_score:.0f}</div>'
                f'<div class="etf-lbl" style="margin-bottom:10px">匹配分</div>'
                f'<div class="etf-name" style="margin-bottom:2px">{d.name}</div>'
                f'<div class="etf-lbl" style="text-transform:none;letter-spacing:0;font-size:0.72rem;margin-bottom:10px">{d.role}</div>'
                f'<div style="font-size:0.82rem;line-height:1.55">{d.plain_intro}</div>'
                f'</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    top = domains[0]
    with st.expander(f"为什么优先推荐「{top.name}」？", expanded=True):
        c1,c2 = st.columns(2)
        c1.markdown(f"**适合人群**\n\n{top.suitable_for}")
        c1.markdown(f"**收益来源**\n\n{top.return_source}")
        c2.markdown(f"**主要风险**\n\n{top.key_risks}")
        c2.markdown(f"**常见误区**\n\n{top.common_mistake}")
    with st.expander("当前用户画像"):
        r1,r2,r3 = st.columns(3)
        r1.metric("投资经验",user.experience); r1.metric("风险层级",user.risk_level)
        r2.metric("持有期限",user.horizon)
        r2.metric("投资目标",user.goal[:8]+"…" if len(user.goal)>8 else user.goal)
        r3.metric("偏好方向",user.preference[:8]+"…" if len(user.preference)>8 else user.preference)
        r3.metric("计划金额",f"{user.amount:,.0f} 元")
    return domains


# ── Tab: ETF 对比 ─────────────────────────────────────────────────────────────
def render_candidates(df, user, domains):
    stitle("ETF 候选对比", "先选方向，再在同方向内比较规模、费率、回撤")
    sel    = st.selectbox("选择方向", [d.name for d in domains], label_visibility="collapsed")
    domain = next(d for d in domains if d.name==sel)
    cands  = rank_candidates(df, user, domain)
    for _, row in cands.iterrows():
        etf_card(row, highlight=bool(row.get("domain_match")))
    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("AI 方向解释", expanded=False):
        cb, cn = st.columns([2,5])
        with cb:
            if st.button("生成 AI 解释", type="primary", key="btn_narrative"):
                nt, nw = make_narrative(user, domain, cands)
                st.session_state["narrative_text"] = nt
                st.session_state["narrative_warning"] = nw
        with cn: st.caption("AI 根据偏好和候选数据生成解释，不构成投资建议")
        if "narrative_text" in st.session_state:
            if st.session_state.get("narrative_warning"): st.caption(st.session_state["narrative_warning"])
            st.markdown(st.session_state["narrative_text"])
    with st.expander("完整对比数据"):
        st.dataframe(cands[["code","name","category","quality_score","candidate_score",
                              "total_fee","age_years","max_drawdown","volatility"]],
                     use_container_width=True)


# ── Tab: 质量评估 ─────────────────────────────────────────────────────────────
def render_quality_detail(df, prices, domains):
    stitle("单只 ETF 质量评估", "质量分不是买卖信号，是对规模、费率、风险和数据完整性的系统检查")
    sel    = st.selectbox("选择 ETF", df["code"]+"  "+df["name"], label_visibility="collapsed")
    code   = sel.split("  ")[0]
    row    = df.loc[df["code"]==code].iloc[0]
    domain = domains[0] if domains else None

    # 4 metrics
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("最大回撤",   fp(row.get("max_drawdown")))
    m2.metric("年化波动率", fp(row.get("volatility")))
    m3.metric("近一年收益", fp(row.get("one_year_return")))
    m4.metric("成立年限",   f"{row.get('age_years',0):.1f} 年" if pd.notna(row.get("age_years")) else "—")

    st.markdown("<br>", unsafe_allow_html=True)

    # Score bars + Price chart
    lc, rc = st.columns([5,7], gap="large")
    with lc:
        score_bars(row)
    with rc:
        price_df = prices[prices["code"]==code].copy() if not prices.empty else pd.DataFrame()
        st.markdown(
            '<div class="card" style="padding:16px 18px">'
            '<div class="etf-lbl" style="margin-bottom:12px">近一年收盘价走势</div>',
            unsafe_allow_html=True)
        if not price_df.empty:
            price_chart_fig(price_df)
        else:
            st.info("暂无行情数据，请点击左侧「拉取最新行情」")
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Risk flags
    flags = risk_flags(row)
    rows_html = "".join(
        f'<div class="flag-row"><span class="flag-icon">⚠</span>'
        f'<span class="flag-text">{f}</span></div>' for f in flags)
    st.markdown(
        f'<div class="card" style="padding:16px 20px;margin-bottom:14px">'
        f'<div class="etf-lbl" style="margin-bottom:12px">风险提示</div>'
        f'{rows_html}</div>',
        unsafe_allow_html=True)

    # AI analysis
    st.markdown('<div class="card" style="padding:20px 22px">', unsafe_allow_html=True)
    ac, nc = st.columns([2,4])
    with ac:
        run_ai = st.button("🤖 一键 AI 深度分析", type="primary", key=f"ai_{code}", use_container_width=True)
    with nc:
        if ai_available(): st.caption(f"{st.session_state.get('ai_provider','AI')} · {st.session_state.get('ai_model','')} 流式输出")
        else: st.caption("请在左侧填写 AI 服务商信息")
    ai_key      = f"ai_result_{code}"
    user_profile = st.session_state.get("user_profile")
    if run_ai:
        st.session_state.pop(ai_key, None)
        result = st.write_stream(stream_ai_analysis(row, user_profile, domain))
        st.session_state[ai_key] = result
    elif ai_key in st.session_state:
        st.markdown(st.session_state[ai_key])
    else:
        qs = self_check_questions(row, domain)
        st.markdown(
            '<div style="margin-bottom:10px;font-size:0.8rem">点击按钮后 AI 将流式生成分析报告。以下是预设自查问题：</div>'
            + "".join(f'<div style="font-size:0.84rem;margin-bottom:6px">· {q}</div>' for q in qs),
            unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)


# ── Tab: 数据 & 术语 ──────────────────────────────────────────────────────────
def render_universe_and_terms(df):
    stitle("样本池 & 术语", "10 只上证 ETF · 数据来自 Excel + Yahoo Finance 缓存")
    d2 = df[["code","name","category","style","quality_score","total_fee","age_years","scale_proxy","price_points"]].copy()
    d2["scale_proxy"] = d2["scale_proxy"].apply(fm)
    st.dataframe(d2, use_container_width=True)
    st.markdown("<br>", unsafe_allow_html=True)
    stitle("术语说明")
    cols = st.columns(2)
    for i,(term,desc) in enumerate(HELP_TEXT.items()):
        with cols[i%2]:
            st.markdown(
                f'<div class="term-card">'
                f'<div class="term-title">{term}</div>'
                f'<div class="term-desc">{desc}</div>'
                f'</div>', unsafe_allow_html=True)


# ── Tab: 调试 ─────────────────────────────────────────────────────────────────
def render_debug(profile, prices, df):
    stitle("调试信息","开发用")
    def mt(p): return (pd.Timestamp(p.stat().st_mtime,unit="s",tz="UTC")
                        .tz_convert(CST).strftime("%Y-%m-%d %H:%M:%S"))
    c1,c2 = st.columns(2)
    with c1:
        if PROFILE_CACHE.exists(): st.success(f"profile 缓存 · {mt(PROFILE_CACHE)}")
        else: st.error("profile 缓存不存在")
    with c2:
        if PRICE_CACHE.exists(): st.success(f"prices 缓存 · {mt(PRICE_CACHE)}")
        else: st.error("prices 缓存不存在")
    if not prices.empty and "fetch_time" in prices.columns and prices["fetch_time"].notna().any():
        st.info(f"行情最近拉取：{prices['fetch_time'].dropna().iloc[0]}")
    if not prices.empty:
        st.dataframe(prices.groupby("code").agg(
            数据点数=("close","count"),最早日期=("date","min"),
            最新日期=("date","max"),最新收盘价=("close","last")).reset_index(),
            use_container_width=True)
    with st.expander("ETF 基本信息"): st.dataframe(profile, use_container_width=True)
    with st.expander("行情缓存（最近20条）"):
        if not prices.empty: st.dataframe(prices.sort_values("date",ascending=False).head(20), use_container_width=True)
    with st.expander("综合打分表"): st.dataframe(df, use_container_width=True)


# ── Main ──────────────────────────────────────────────────────────────────────
def main() -> None:
    st.set_page_config(
        page_title="ETF 投资助手", layout="wide", page_icon="📊",
        initial_sidebar_state="expanded",
    )

    # Sidebar first (sets dark_mode in session_state via toggle)
    refresh_prices, refresh_profile = sidebar_controls()

    # Inject CSS based on current theme
    dark = st.session_state.get("dark_mode", True)
    st.markdown(build_css(dark), unsafe_allow_html=True)

    # Data
    profile = load_or_create_profile(refresh=refresh_profile)
    if refresh_prices:
        prog = st.empty()
        with st.spinner("正在从 Yahoo Finance 拉取行情…"):
            prices, warn = load_prices(profile["code"], refresh=True, progress_placeholder=prog)
        prog.empty()
        if warn: st.error(f"部分拉取失败：{warn}")
        else: st.success("行情更新成功")
    else:
        prices, warn = load_prices(profile["code"], refresh=False)
        if warn: st.warning(warn)

    df      = add_metrics(profile, prices)
    default = UserProfile("我是新手","低风险","3年以上","长期配置，慢慢积累","我想先要一个稳一点的底座",50000.0)
    user    = st.session_state.get("user_profile", default)
    domains = recommend_domains(user)

    # Page header
    st.markdown(
        '<div style="padding:0 0 20px">'
        '<div style="font-size:1.65rem;font-weight:800;letter-spacing:-0.02em">ETF 投资助手</div>'
        '<div style="font-size:0.82rem;margin-top:4px;opacity:0.55">'
        '基于规则评分 + AI 解释 · 仅用于学习，不构成投资建议</div></div>',
        unsafe_allow_html=True)

    tab_a, tab_b, tab_c, tab_d, tab_e, tab_f = st.tabs(
        ["偏好问卷","方向推荐","ETF 对比","质量评估","数据 & 术语","🔧 调试"])

    with tab_a: user    = render_profile_form();         domains = recommend_domains(user)
    with tab_b: domains = render_domain_recommendations(user)
    with tab_c: render_candidates(df, user, domains)
    with tab_d: render_quality_detail(df, prices, domains)
    with tab_e: render_universe_and_terms(df)
    with tab_f: render_debug(profile, prices, df)


if __name__ == "__main__":
    main()
