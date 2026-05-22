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
BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "上证ETF(1).xlsx"
CACHE_DIR = BASE_DIR / "data" / "cache"
PROFILE_CACHE = CACHE_DIR / "etf_sample_profile.csv"
PRICE_CACHE = CACHE_DIR / "etf_sample_prices.csv"
SAMPLE_SIZE = 10
CST = "Asia/Shanghai"

DEEPSEEK_BASE_URL = os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com").rstrip("/")
DEEPSEEK_MODEL    = os.environ.get("DEEPSEEK_MODEL", "deepseek-chat")

AI_PROVIDERS: Dict[str, Dict[str, str]] = {
    "DeepSeek":       {"base_url": "https://api.deepseek.com",                              "default_model": "deepseek-chat"},
    "Gemini (Google)":{"base_url": "https://generativelanguage.googleapis.com/v1beta/openai","default_model": "gemini-2.0-flash"},
    "OpenAI":         {"base_url": "https://api.openai.com/v1",                             "default_model": "gpt-4o-mini"},
    "Moonshot (Kimi)":{"base_url": "https://api.moonshot.cn/v1",                            "default_model": "moonshot-v1-8k"},
    "智谱 GLM":        {"base_url": "https://open.bigmodel.cn/api/paas/v4",                  "default_model": "glm-4-flash"},
    "硅基流动":        {"base_url": "https://api.siliconflow.cn/v1",                         "default_model": "Qwen/Qwen2.5-7B-Instruct"},
}

# ── CSS ───────────────────────────────────────────────────────────────────────
BASE_CSS = """<style>
/* ── Variables: dark defaults ── */
:root {
  --bg-page:     #080d18;
  --bg-sidebar:  #0c1322;
  --bg-card:     #0d1526;
  --border-s:    #162035;
  --border-c:    #1e2d4a;
  --text-h:      #e8f0ff;
  --text-p:      #c8d6e8;
  --text-sec:    #94a3b8;
  --text-m:      #64748b;
  --text-d:      #334155;
  --accent:      #60a5fa;
  --accent-btn:  #2563eb;
  --card-shadow: none;
  --badge-bg:    rgba(96,165,250,0.12);
  --badge-bdr:   rgba(96,165,250,0.3);
}

/* ── Hide chrome, NOT the header element (sidebar toggle lives there) ── */
footer                                { visibility: hidden !important; }
#MainMenu                             { visibility: hidden !important; }
[data-testid="stToolbar"]             { visibility: hidden !important; }
[data-testid="stDecoration"]          { display: none !important; }
[data-testid="stHeader"]              {
  background: transparent !important;
  border-bottom: none !important;
  box-shadow: none !important;
}

/* ── Sidebar collapsed arrow: make it obvious ── */
[data-testid="collapsedControl"] {
  visibility: visible !important;
  background: var(--bg-card) !important;
  border: 1px solid var(--border-c) !important;
  border-left: none !important;
  border-radius: 0 8px 8px 0 !important;
  box-shadow: 4px 0 12px rgba(0,0,0,0.25) !important;
  z-index: 999 !important;
}
[data-testid="collapsedControl"] svg {
  color: var(--accent) !important;
  fill: var(--accent) !important;
}

/* ── Layout ── */
.stApp                    { background-color: var(--bg-page) !important; }
[data-testid="stSidebar"] {
  background-color: var(--bg-sidebar) !important;
  border-right: 1px solid var(--border-s);
}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
  gap: 0; border-bottom: 1px solid var(--border-s);
  background: transparent; padding: 0;
}
.stTabs [data-baseweb="tab"] {
  background: transparent !important;
  color: var(--text-m);
  height: 42px; padding: 0 20px; font-size: 0.875rem;
  border-radius: 0; border-bottom: 2px solid transparent;
  margin-bottom: -1px;
  transition: color 0.18s ease;
}
.stTabs [aria-selected="true"] {
  color: var(--accent) !important;
  border-bottom-color: var(--accent) !important;
  background: transparent !important;
}
.stTabs [data-baseweb="tab-panel"] { padding-top: 20px; }

/* ── Buttons ── */
div[data-testid="stButton"] > button {
  background: var(--bg-card); color: var(--text-sec);
  border: 1px solid var(--border-c); border-radius: 8px;
  font-size: 0.875rem;
  transition: all 0.18s cubic-bezier(0.4, 0, 0.2, 1);
}
div[data-testid="stButton"] > button:hover {
  border-color: var(--accent); color: var(--accent);
  background: var(--bg-card); transform: translateY(-1px);
}
div[data-testid="stButton"] > button[kind="primary"] {
  background: var(--accent-btn); color: #fff; border: none; font-weight: 600;
  box-shadow: 0 1px 4px rgba(37,99,235,0.35);
}
div[data-testid="stButton"] > button[kind="primary"]:hover {
  background: #1d4ed8; color: #fff; border: none;
  transform: translateY(-1px);
  box-shadow: 0 4px 14px rgba(37,99,235,0.45);
}

/* ── Metrics ── */
[data-testid="stMetricValue"] { font-size: 1.45rem; font-weight: 700; color: var(--text-h); }
[data-testid="stMetricLabel"] {
  font-size: 0.7rem; color: var(--text-m);
  text-transform: uppercase; letter-spacing: 0.06em;
}

/* ── Inputs ── */
.stSelectbox > div > div {
  background: var(--bg-card) !important;
  border-color: var(--border-c) !important; border-radius: 8px !important;
  transition: border-color 0.15s;
}
.stTextInput > div > div > input,
[data-testid="stNumberInput"] input {
  background: var(--bg-card) !important;
  border-color: var(--border-c) !important; border-radius: 8px !important;
  color: var(--text-p) !important; transition: border-color 0.15s;
}

/* ── Misc ── */
.stAlert   { border-radius: 10px !important; }
hr         { border-color: var(--border-s) !important; }
[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }
details > summary {
  background: var(--bg-card) !important; border: 1px solid var(--border-c) !important;
  border-radius: 8px !important; padding: 10px 16px !important;
  color: var(--text-sec) !important; font-size: 0.875rem !important;
  cursor: pointer; transition: border-color 0.15s;
}
details > summary:hover { border-color: var(--accent) !important; }
h1         { color: var(--text-h) !important; }
h2, h3, h4 { color: var(--text-p) !important; }
p, [data-testid="stMarkdownContainer"] p { color: var(--text-sec); }

/* ── Component classes ── */

/* Card base */
.c-card {
  background: var(--bg-card);
  border: 1px solid var(--border-c);
  border-radius: 12px;
  box-shadow: var(--card-shadow);
  transition: transform 0.2s cubic-bezier(0.4,0,0.2,1),
              box-shadow 0.2s ease;
}
.c-card:hover {
  transform: translateY(-2px);
  box-shadow: 0 8px 28px rgba(0,0,0,0.18);
}
.c-card--accent { border-color: var(--accent-btn); }

/* ETF card */
.etf-card          { padding: 18px 22px; margin-bottom: 10px; }
.etf-name          { font-size: 1rem; font-weight: 600; color: var(--text-h); margin-bottom: 4px; }
.etf-meta          { color: var(--text-m); font-size: 0.78rem;
                     display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
.etf-score         { padding: 4px 14px; border-radius: 20px;
                     font-weight: 700; font-size: 0.95rem;
                     white-space: nowrap; flex-shrink: 0; margin-left: 12px; }
.etf-metrics       { display: grid; grid-template-columns: repeat(5,1fr);
                     gap: 10px; border-top: 1px solid var(--border-s); padding-top: 12px;
                     margin-top: 14px; }
.etf-cell-label    { color: var(--text-d); font-size: 0.63rem;
                     text-transform: uppercase; letter-spacing: .05em; margin-bottom: 3px; }
.etf-cell-val      { color: var(--text-p); font-size: 0.85rem; font-weight: 500; }

/* Match tag (inline, no overlap) */
.match-tag {
  display: inline-flex; align-items: center;
  background: var(--badge-bg); color: var(--accent);
  padding: 1px 7px; border-radius: 99px;
  font-size: 0.63rem; font-weight: 600;
  border: 1px solid var(--badge-bdr);
}

/* Domain card */
.domain-card { padding: 20px; height: 100%; }

/* Score card */
.score-card { padding: 22px 24px; }

/* Term card */
.term-card {
  background: var(--bg-card); border: 1px solid var(--border-c);
  border-radius: 10px; padding: 14px 16px; margin-bottom: 8px;
  transition: border-color 0.15s ease, transform 0.15s ease;
}
.term-card:hover { border-color: var(--accent); transform: translateX(3px); }

/* AI box */
.ai-box { padding: 20px 22px; }

/* Section title */
.sec-title { margin-bottom: 24px; }
.sec-title h2 {
  margin: 0; font-size: 1.2rem; font-weight: 700;
  color: var(--text-h); letter-spacing: -0.01em;
}
.sec-title p {
  margin: 4px 0 0; font-size: 0.82rem; color: var(--text-m) !important;
}
</style>"""

# Light mode — lobehub-inspired: white cards + shadow, clean gray hierarchy
LIGHT_CSS = """<style>
:root {
  --bg-page:     #f9fafb;
  --bg-sidebar:  #ffffff;
  --bg-card:     #ffffff;
  --border-s:    #f3f4f6;
  --border-c:    #e5e7eb;
  --text-h:      #111827;
  --text-p:      #1f2937;
  --text-sec:    #374151;
  --text-m:      #6b7280;
  --text-d:      #9ca3af;
  --accent:      #2563eb;
  --accent-btn:  #2563eb;
  --card-shadow: 0 1px 3px rgba(0,0,0,0.07), 0 1px 2px rgba(0,0,0,0.04);
  --badge-bg:    rgba(37,99,235,0.07);
  --badge-bdr:   rgba(37,99,235,0.2);
}
.stApp        { color: #1f2937 !important; }
.stApp *      { font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",system-ui,sans-serif !important; }
[data-testid="stMarkdownContainer"] p { color: #374151; }
.stSelectbox > div > div { color: #1f2937 !important; }
.stTextInput > div > div > input,
[data-testid="stNumberInput"] input   { color: #1f2937 !important; }
[data-testid="stHeader"]              { box-shadow: 0 1px 0 #f3f4f6 !important; }
[data-testid="stSidebar"]             { box-shadow: 1px 0 0 #f3f4f6 !important; }
.c-card:hover                         { box-shadow: 0 8px 24px rgba(0,0,0,0.09) !important; }
[data-testid="collapsedControl"]      {
  box-shadow: 3px 0 10px rgba(0,0,0,0.08) !important;
  background: #ffffff !important;
  border-color: #e5e7eb !important;
}
[data-testid="collapsedControl"] svg  {
  color: #2563eb !important; fill: #2563eb !important;
}
</style>"""

HELP_TEXT = {
    "宽基":   "宽基ETF通常跟踪沪深300、上证50、中证500这类大盘或综合指数，买的是一篮子股票，更适合作为长期配置的核心仓位。",
    "红利":   "红利ETF偏向高分红、现金流较稳定的公司，常被保守投资者关注，但也会受利率、周期和行业集中影响。",
    "行业主题":"行业主题ETF集中投向某个行业或主题，例如消费、金融、半导体。机会更集中，波动和回撤通常也更大。",
    "回撤":   "回撤可以理解为从阶段高点跌下来多少。比如最大回撤-30%，意味着过去某段时间里最高点买入会一度亏约30%。",
    "波动率": "波动率衡量价格上下震荡的剧烈程度。波动率高不等于一定不好，但更考验持有耐心和仓位控制。",
    "费率":   "费率是长期持有时持续付出的管理费、托管费。单年看不大，但多年复利下来会影响实际收益。",
    "规模":   "规模越大通常流动性和稳定性更好。规模很小的ETF要额外关注成交不活跃或清盘风险。",
    "流动性": "流动性指买卖是否方便、价差是否大。流动性差时，买入或卖出可能付出额外隐性成本。",
}


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
        category_keywords=("宽基","核心"),
        risk_hint="适合作为组合底座，但仍要控制买入节奏和总仓位。",
    ),
    "红利/低波": DomainRecommendation(
        name="红利/低波", match_score=0, role="稳健核心或防守仓位",
        plain_intro="更关注分红、现金流和相对稳健的公司，目标是少一点刺激，多一点纪律。",
        suitable_for="适合不喜欢大起大落、希望降低组合波动的新手或保守型投资者。",
        return_source="主要来自股息、企业稳定盈利和低估值修复。",
        key_risks="红利不等于无风险，可能有行业集中、周期下行和股息下降风险。",
        common_mistake="把高股息率简单理解成高收益，忽略股价下跌可能抵消分红。",
        category_keywords=("红利","低波","价值","金融"),
        risk_hint="当前样本池红利ETF不足，demo会用价值/金融类作为近似候选展示。",
    ),
    "行业主题": DomainRecommendation(
        name="行业主题", match_score=0, role="卫星仓位",
        plain_intro="集中投向某个行业或主题，适合表达观点，但不适合新手一上来重仓。",
        suitable_for="适合能承受较大波动、理解行业周期、愿意把它控制在小比例仓位的投资者。",
        return_source="主要来自行业景气度上行、政策催化、估值扩张或盈利改善。",
        key_risks="行业热度高时容易追涨，景气反转时回撤也可能很深。",
        common_mistake="因为最近涨得多就买入，却没有想清楚退出规则和仓位上限。",
        category_keywords=("行业","主题","成长"),
        risk_hint="更适合作为小比例卫星仓位，不建议作为保守投资者的核心仓位。",
    ),
}


# ── Data loading ──────────────────────────────────────────────────────────────
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {
        "证券代码":"code","证券名称":"name","基金类型":"fund_type",
        "单位净值\n[交易日期] 最新[单位]元":"nav",
        "基金份额\n[交易日期] 最新\n[单位]  份":"shares",
        "投资风格\n[年度] 2023\n[报告期] 中报":"style",
        "基金管理人":"manager","基金托管人":"custodian",
        "管理费率[单位]%":"management_fee","托管费率[单位]%":"custody_fee",
        "基金经理(现任)":"fund_manager","基金存续期[单位]年":"duration_years",
        "基金成立日":"inception_date","基金到期日":"maturity_date",
    }
    return df.rename(columns=mapping)

def parse_inception(value: Any) -> Optional[pd.Timestamp]:
    if pd.isna(value): return None
    text = str(value).split(".")[0].strip()
    if len(text) != 8: return None
    parsed = pd.to_datetime(text, format="%Y%m%d", errors="coerce")
    return None if pd.isna(parsed) else parsed

def classify_etf(row: pd.Series) -> str:
    text = f"{row.get('name','')} {row.get('style','')}"
    if any(k in text for k in ["沪深300","上证50","中证500","国企"]): return "宽基/核心"
    if any(k in text for k in ["消费","金融"]): return "行业/主题"
    if "成长" in text: return "成长风格"
    if "价值" in text: return "价值风格"
    return "其他"

def load_profile_from_excel() -> pd.DataFrame:
    df = pd.read_excel(EXCEL_PATH).head(SAMPLE_SIZE)
    df = normalize_columns(df)
    keep = ["code","name","fund_type","nav","shares","style","manager",
            "custodian","management_fee","custody_fee","fund_manager","inception_date"]
    df = df[keep].copy()
    df["inception"]   = df["inception_date"].apply(parse_inception)
    df["age_years"]   = df["inception"].apply(
        lambda x: np.nan if x is None else max((pd.Timestamp.today()-x).days/365.25,0))
    df["total_fee"]   = (pd.to_numeric(df["management_fee"],errors="coerce").fillna(0)
                       + pd.to_numeric(df["custody_fee"],errors="coerce").fillna(0))
    df["scale_proxy"] = pd.to_numeric(df["nav"],errors="coerce") * pd.to_numeric(df["shares"],errors="coerce")
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

def fetch_prices_from_yfinance(codes: Iterable[str], progress_placeholder: Any = None) -> Tuple[pd.DataFrame, Optional[str]]:
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

def load_prices(codes: Iterable[str], refresh: bool = False, progress_placeholder: Any = None) -> Tuple[pd.DataFrame, Optional[str]]:
    if PRICE_CACHE.exists() and not refresh:
        prices = pd.read_csv(PRICE_CACHE)
        prices["date"] = pd.to_datetime(prices["date"], errors="coerce")
        return prices, None
    return fetch_prices_from_yfinance(codes, progress_placeholder=progress_placeholder)


# ── Metrics & scoring ─────────────────────────────────────────────────────────
def max_drawdown(close: pd.Series) -> float:
    close = close.dropna()
    return np.nan if close.empty else float((close/close.cummax()-1).min())

def annual_volatility(close: pd.Series) -> float:
    r = close.pct_change().dropna()
    return np.nan if r.empty else float(r.std()*np.sqrt(252))

def one_year_return(close: pd.Series) -> float:
    close = close.dropna()
    return np.nan if len(close)<2 else float(close.iloc[-1]/close.iloc[0]-1)

def add_metrics(profile: pd.DataFrame, prices: pd.DataFrame) -> pd.DataFrame:
    df = profile.copy()
    rows = []
    for code in df["code"]:
        close = prices.loc[prices["code"]==code,"close"] if not prices.empty else pd.Series(dtype=float)
        rows.append({"code":code,"max_drawdown":max_drawdown(close),
                     "volatility":annual_volatility(close),"one_year_return":one_year_return(close),
                     "price_points":int(close.dropna().shape[0])})
    df = df.merge(pd.DataFrame(rows), on="code", how="left")
    scores = pd.DataFrame([score_etf(row) for _,row in df.iterrows()])
    return pd.concat([df.reset_index(drop=True), scores], axis=1)

def score_size(scale):
    if pd.isna(scale): return 8
    if scale>=50_000_000_000: return 25
    if scale>=10_000_000_000: return 22
    if scale>=1_000_000_000:  return 16
    if scale>=100_000_000:    return 9
    return 4

def score_fee(f):
    if pd.isna(f): return 8
    if f<=0.20: return 20
    if f<=0.30: return 17
    if f<=0.60: return 12
    if f<=1.00: return 7
    return 3

def score_age(a):
    if pd.isna(a): return 6
    if a>=5: return 20
    if a>=3: return 16
    if a>=1: return 10
    if a>=0.5: return 6
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

def score_etf(row: pd.Series) -> Dict[str,float]:
    static_missing = sum(pd.isna(row.get(c)) for c in ["nav","shares","management_fee","custody_fee","inception"])
    sz = score_size(row.get("scale_proxy")); fe = score_fee(row.get("total_fee"))
    ag = score_age(row.get("age_years"));    ri = score_risk(row.get("max_drawdown"),row.get("volatility"))
    da = score_data(int(row.get("price_points",0) or 0), static_missing)
    return {"quality_score":round(sz+fe+ag+ri+da,1),
            "size_score":sz,"fee_score":fe,"age_score":ag,"risk_score":ri,"data_score":da}


# ── Recommendation logic ──────────────────────────────────────────────────────
def infer_risk_level(risk_answer, experience, horizon):
    if "不确定" in risk_answer:
        return "低风险" if (experience=="我是新手" or horizon=="1年以内") else "中风险"
    if "不太能接受" in risk_answer: return "低风险"
    if "阶段性波动"  in risk_answer: return "中风险"
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
    result = [DomainRecommendation(**{**base.__dict__,"match_score":round(s[n],1)})
              for n,base in DOMAIN_LIBRARY.items()]
    return sorted(result, key=lambda x: x.match_score, reverse=True)

def etf_matches_domain(row, domain):
    text = f"{row.get('name','')} {row.get('style','')} {row.get('category','')}"
    return any(k in text for k in domain.category_keywords)

def candidate_score(row, user, domain):
    score = float(row["quality_score"])
    if etf_matches_domain(row,domain): score+=18
    if user.risk_level=="低风险":
        if row.get("total_fee",99)<=0.20: score+=8
        if pd.notna(row.get("max_drawdown")) and row.get("max_drawdown")<-0.25: score-=12
        if pd.notna(row.get("volatility")) and row.get("volatility")>0.25: score-=10
        if "行业" in str(row.get("category","")): score-=6
    elif user.risk_level=="高风险" and "行业" in str(row.get("category","")):
        score+=5
    if user.horizon=="1年以内" and pd.notna(row.get("volatility")) and row.get("volatility")>0.25:
        score-=8
    return round(score,1)

def rank_candidates(df, user, domain):
    ranked = df.copy()
    ranked["candidate_score"] = ranked.apply(lambda r: candidate_score(r,user,domain), axis=1)
    ranked["domain_match"]    = ranked.apply(lambda r: etf_matches_domain(r,domain), axis=1)
    return ranked.sort_values(["domain_match","candidate_score"],ascending=[False,False]).head(3)

def risk_flags(row: pd.Series) -> List[str]:
    flags = []
    if row.get("age_years",99)<1: flags.append("成立时间较短，成立以来收益率参考价值有限。")
    if row.get("total_fee",0)>0.60: flags.append("费率偏高，长期持有成本会持续累积。")
    if row.get("scale_proxy",np.inf)<1_000_000_000: flags.append("规模偏小，需关注流动性和清盘风险。")
    if pd.notna(row.get("max_drawdown")) and row.get("max_drawdown")<-0.25:
        flags.append("近一年最大回撤较大，需确认能否承受。")
    if pd.notna(row.get("volatility")) and row.get("volatility")>0.25:
        flags.append("近一年波动率较高，不适合保守型核心仓位。")
    if not flags: flags.append("未发现突出单项风险，但仍需控制仓位。")
    return flags

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
            f"这个方向的定位是{domain.role}：{domain.plain_intro}"
            f"样本池里可以先重点比较：{names}。"
            f"推荐只是帮助缩小观察范围，不是买入建议；决策前还要看回撤、费率、规模和能否长期持有。")


# ── AI functions ──────────────────────────────────────────────────────────────
def ai_available(): return bool(st.session_state.get("ai_api_key") or os.environ.get("DEEPSEEK_API_KEY"))

def get_ai_config():
    api_key  = st.session_state.get("ai_api_key") or os.environ.get("DEEPSEEK_API_KEY","")
    base_url = (st.session_state.get("ai_base_url") or DEEPSEEK_BASE_URL).rstrip("/")
    model    = st.session_state.get("ai_model") or DEEPSEEK_MODEL
    return api_key, base_url, model

def to_jsonable(v):
    if isinstance(v, dict):  return {str(k):to_jsonable(x) for k,x in v.items()}
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
    if user:   ctx.update({"user_risk_level":user.risk_level,"user_horizon":user.horizon,"user_experience":user.experience})
    if domain: ctx["recommended_domain"] = domain.name
    body = json.dumps({
        "model":model,
        "messages":[
            {"role":"system","content":"你是一个谨慎的ETF投资教育助手。只基于用户提供的数据写分析，不得编造数据，不得预测收益，不得给出买入卖出建议。用中文输出，面向新手。"},
            {"role":"user","content":("请对以下ETF做三部分分析，使用Markdown格式：\n"
                "**1. 综合质量判断**（约80字）\n**2. 主要风险提示**（约80字）\n**3. 购前自查清单**（3-5条）\n\n"
                f"数据：{json.dumps(to_jsonable(ctx),ensure_ascii=False)}")},
        ],
        "temperature":0.4,"max_tokens":700,"stream":True,
    }, ensure_ascii=False).encode("utf-8")
    try:
        with requests.post(f"{base_url}/chat/completions",
            headers={"Authorization":f"Bearer {api_key}","Content-Type":"application/json; charset=utf-8"},
            data=body, stream=True, timeout=30) as resp:
            if not resp.ok: yield f"API 错误 {resp.status_code}: {resp.content.decode('utf-8',errors='replace')[:200]}"; return
            for line in resp.iter_lines():
                if not line or not line.startswith(b"data: "): continue
                chunk_data = line[6:]
                if chunk_data.strip()==b"[DONE]": return
                try:
                    delta = json.loads(chunk_data.decode("utf-8"))["choices"][0]["delta"].get("content","")
                    if delta: yield delta
                except Exception: continue
    except Exception as e: yield f"\n\n（输出中断：{e}）"

def call_ai_narrative(payload):
    api_key, base_url, model = get_ai_config()
    if not api_key: return None,"未设置 API Key，使用规则模板。"
    try:
        body = json.dumps({
            "model":model,
            "messages":[
                {"role":"system","content":"你是一个谨慎的ETF投资教育助手。只基于提供的事实写解释，不得编造数据，不得预测收益，不得给出投资建议。用中文输出，面向新手。"},
                {"role":"user","content":"请把以下事实润色成一段适合展示的解释。结构：1）为什么先看这个方向；2）候选ETF怎么比较；3）风险和自查提醒。控制在180字以内。\n\n"+json.dumps(to_jsonable(payload),ensure_ascii=False)},
            ],
            "temperature":0.3,"max_tokens":500,
        }, ensure_ascii=False).encode("utf-8")
        resp = requests.post(f"{base_url}/chat/completions",
            headers={"Authorization":f"Bearer {api_key}","Content-Type":"application/json; charset=utf-8"},
            data=body, timeout=20)
        if not resp.ok: return None,f"HTTP {resp.status_code}: {resp.content.decode('utf-8',errors='replace')[:200]}"
        return json.loads(resp.content.decode("utf-8"))["choices"][0]["message"]["content"].strip(), None
    except Exception as exc:
        return None, f"{st.session_state.get('ai_provider','AI')} 调用失败：{exc}"

def make_narrative(user, domain, candidates):
    payload = {
        "user_profile":user.__dict__,
        "domain":{"name":domain.name,"role":domain.role,"intro":domain.plain_intro,
                  "suitable_for":domain.suitable_for,"return_source":domain.return_source,
                  "key_risks":domain.key_risks,"common_mistake":domain.common_mistake},
        "candidates":[{"code":r["code"],"name":r["name"],"quality_score":r["quality_score"],
                       "fee":r.get("total_fee"),"age_years":r.get("age_years"),
                       "max_drawdown":r.get("max_drawdown"),"volatility":r.get("volatility"),
                       "risk_flags":risk_flags(r)} for _,r in candidates.head(3).iterrows()],
        "guardrail":"仅用于学习和决策辅助，不构成投资建议。",
    }
    text, warn = call_ai_narrative(payload)
    return (text,None) if text else (build_fallback_narrative(user,domain,candidates), warn)


# ── Formatting helpers ────────────────────────────────────────────────────────
def fmt_pct(v):  return "—" if pd.isna(v) else f"{v*100:.2f}%"
def fmt_money(v):
    if pd.isna(v): return "—"
    v = float(v)
    if v>=100_000_000: return f"{v/100_000_000:.1f} 亿"
    if v>=10_000: return f"{v/10_000:.1f} 万"
    return f"{v:.0f}"

def _chart_style():
    dark = st.session_state.get("dark_mode", True)
    if dark: return {"plot_bg":"#0d1526","paper_bg":"#0d1526","font":"#64748b","grid":"#162035","line":"#3b82f6","fill":"rgba(59,130,246,0.07)"}
    return   {"plot_bg":"#ffffff",       "paper_bg":"#f9fafb","font":"#6b7280","grid":"#f3f4f6","line":"#2563eb","fill":"rgba(37,99,235,0.05)"}


# ── UI primitives ─────────────────────────────────────────────────────────────
def section_title(title: str, subtitle: str = "") -> None:
    sub = f'<p>{subtitle}</p>' if subtitle else ""
    st.markdown(f'<div class="sec-title"><h2>{title}</h2>{sub}</div>', unsafe_allow_html=True)

def etf_card_html(row: pd.Series, highlight: bool = False) -> None:
    score = float(row.get("quality_score",0))
    sc    = "#10b981" if score>=70 else "#f59e0b" if score>=50 else "#ef4444"
    accent_cls = " c-card--accent" if highlight else ""
    # inline badge — no overlap with score
    match_tag = '<span class="match-tag">方向匹配</span>' if highlight else ""
    fee = row.get("total_fee", float("nan"))
    dd  = row.get("max_drawdown", float("nan"))
    vol = row.get("volatility", float("nan"))
    age = row.get("age_years", float("nan"))
    scl = row.get("scale_proxy", float("nan"))
    dd_color = "#ef4444" if not pd.isna(dd) and dd<-0.2 else "var(--text-p)"

    def cell(lbl, val, color="var(--text-p)"):
        return (f'<div><div class="etf-cell-label">{lbl}</div>'
                f'<div class="etf-cell-val" style="color:{color}">{val}</div></div>')

    st.markdown(
        f'<div class="c-card etf-card{accent_cls}">'
        # header row: name+meta on left, score on right — no absolute positioning
        f'<div style="display:flex;justify-content:space-between;align-items:flex-start">'
        f'  <div style="min-width:0;flex:1">'
        f'    <div class="etf-name">{row.get("name","")}</div>'
        f'    <div class="etf-meta">'
        f'      <span>{row.get("code","")}</span>'
        f'      <span style="color:var(--border-c)">·</span>'
        f'      <span>{row.get("category","")}</span>'
        f'      {match_tag}'
        f'    </div>'
        f'  </div>'
        # score stays to the right, no absolute
        f'  <div class="etf-score" style="background:{sc}14;color:{sc};border:1px solid {sc}30">'
        f'    {score:.0f}<span style="font-size:0.58rem;opacity:0.55;font-weight:400"> /100</span>'
        f'  </div>'
        f'</div>'
        f'<div class="etf-metrics">'
        f'{cell("费率", f"{fee:.2f}%" if not pd.isna(fee) else "—")}'
        f'{cell("最大回撤", fmt_pct(dd), dd_color)}'
        f'{cell("年化波动", fmt_pct(vol))}'
        f'{cell("成立年限", f"{age:.1f} 年" if not pd.isna(age) else "—")}'
        f'{cell("规模", fmt_money(scl))}'
        f'</div></div>',
        unsafe_allow_html=True,
    )

def score_bars_html(row: pd.Series) -> None:
    """Render quality score breakdown as clean HTML progress bars."""
    items = [("规模",row["size_score"],25),("费率",row["fee_score"],20),
             ("成立年限",row["age_score"],20),("历史风险",row["risk_score"],20),
             ("数据完整性",row["data_score"],15)]
    total = float(row["quality_score"])
    tc    = "#10b981" if total>=70 else "#f59e0b" if total>=50 else "#ef4444"

    bars = ""
    for lbl, sc, mx in items:
        pct  = sc/mx*100
        col  = "#10b981" if pct>=75 else "#3b82f6" if pct>=50 else "#f59e0b"
        bars += (
            f'<div style="margin-bottom:13px">'
            f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:5px">'
            f'  <span style="color:var(--text-sec);font-size:0.82rem">{lbl}</span>'
            f'  <span style="color:var(--text-p);font-size:0.8rem;font-weight:600">'
            f'    {sc:.0f}<span style="color:var(--text-d);font-weight:400">/{mx}</span></span>'
            f'</div>'
            f'<div style="background:var(--border-s);border-radius:99px;height:5px;overflow:hidden">'
            f'  <div style="width:{pct:.0f}%;height:100%;background:{col};border-radius:99px"></div>'
            f'</div></div>'
        )

    st.markdown(
        f'<div class="c-card score-card">'
        f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:22px">'
        f'  <div>'
        f'    <div style="color:var(--text-m);font-size:0.7rem;text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px">综合质量分</div>'
        f'    <div style="color:{tc};font-size:2.6rem;font-weight:800;line-height:1">'
        f'      {total:.0f}<span style="color:var(--text-d);font-size:0.85rem;font-weight:400"> /100</span>'
        f'    </div>'
        f'  </div>'
        f'</div>'
        f'{bars}'
        f'</div>',
        unsafe_allow_html=True,
    )

def price_chart(price_df: pd.DataFrame) -> None:
    cs = _chart_style()
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=price_df["date"], y=price_df["close"],
        mode="lines",
        line=dict(color=cs["line"], width=2),
        fill="tozeroy", fillcolor=cs["fill"],
        hovertemplate="%{x|%Y-%m-%d}  <b>%{y:.3f}</b><extra></extra>",
    ))
    mn, mx = price_df["close"].min()*0.99, price_df["close"].max()*1.01
    fig.update_layout(
        plot_bgcolor=cs["plot_bg"], paper_bgcolor=cs["paper_bg"],
        font=dict(color=cs["font"], size=11),
        height=248, margin=dict(l=40,r=8,t=8,b=8),
        xaxis=dict(showgrid=False, zeroline=False, showline=False,
                   tickfont=dict(size=10), tickformat="%m月"),
        yaxis=dict(showgrid=True, gridcolor=cs["grid"], gridwidth=0.5,
                   zeroline=False, showline=False, tickfont=dict(size=10), range=[mn,mx]),
        hovermode="x unified",
        hoverlabel=dict(bgcolor=cs["plot_bg"], bordercolor=cs["grid"], font_size=12),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})


# ── Sidebar ───────────────────────────────────────────────────────────────────
def sidebar_controls() -> Tuple[bool,bool]:
    with st.sidebar:
        st.toggle("深色模式", value=st.session_state.get("dark_mode",True), key="dark_mode")
        st.markdown(
            '<div style="padding:4px 0 20px">'
            '<div style="font-size:1.05rem;font-weight:700;color:var(--text-h)">ETF 投资助手</div>'
            '<div style="font-size:0.73rem;color:var(--text-d);margin-top:2px">仅用于学习，不构成投资建议</div>'
            '</div>', unsafe_allow_html=True)

        st.markdown('<div style="color:var(--text-m);font-size:0.68rem;text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px">AI 解释引擎</div>', unsafe_allow_html=True)
        selected_provider = st.selectbox("服务商", list(AI_PROVIDERS.keys()), index=0, key="ai_provider", label_visibility="collapsed")
        provider_cfg = AI_PROVIDERS[selected_provider]
        api_key_input = st.text_input("API Key", type="password", placeholder="填入 API Key",
            value=st.session_state.get("ai_api_key", os.environ.get("DEEPSEEK_API_KEY","")), key="_ai_key_input")
        st.session_state["ai_api_key"] = api_key_input
        model_input = st.text_input("模型", placeholder=provider_cfg["default_model"],
            value=st.session_state.get("ai_model","") or provider_cfg["default_model"], key="_ai_model_input")
        st.session_state["ai_model"]    = model_input or provider_cfg["default_model"]
        st.session_state["ai_base_url"] = provider_cfg["base_url"]
        if ai_available(): st.success(f"已启用 · {st.session_state['ai_model']}", icon="✓")
        else: st.caption("未填写 API Key，解释使用规则模板")

        st.divider()
        st.markdown('<div style="color:var(--text-m);font-size:0.68rem;text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px">行情数据</div>', unsafe_allow_html=True)
        refresh_prices  = st.button("拉取最新行情",  use_container_width=True)
        refresh_profile = st.button("重建静态缓存",  use_container_width=True)
        st.caption("数据源：Yahoo Finance · 自动缓存")
        return refresh_prices, refresh_profile


# ── Tab: 偏好问卷 ─────────────────────────────────────────────────────────────
def render_profile_form() -> UserProfile:
    section_title("偏好问卷", "尽量用自己的真实感受来选，不需要先懂专业术语")
    left, right = st.columns(2, gap="large")
    with left:
        experience   = st.selectbox("投资经验", ["我是新手","有一些经验","比较熟悉ETF"])
        risk_answer  = st.selectbox("账户短期亏损时，你的感受更接近？",
            ["我不太能接受亏损","能接受阶段性波动","愿意承受较大波动换取机会","我不确定"])
        horizon      = st.selectbox("这笔钱预计多久不用？",["1年以内","1-3年","3年以上"])
    with right:
        goal         = st.selectbox("你更接近哪种目标？",
            ["希望稳一点，少一些大起大落","长期配置，慢慢积累","想抓住某些行业机会","还没想清楚"])
        preference   = st.selectbox("你现在更想先了解什么？",
            ["我想先要一个稳一点的底座","我更关注分红和稳健","我想了解行业主题机会","我不确定"])
        amount       = st.number_input("计划投入金额（元）", min_value=1000.0, value=50000.0, step=1000.0)

    risk_level = infer_risk_level(risk_answer, experience, horizon)
    user = UserProfile(experience=experience, risk_level=risk_level, horizon=horizon,
                       goal=goal, preference=preference, amount=amount)
    st.session_state["user_profile"] = user

    icon   = {"低风险":"🟢","中风险":"🟡","高风险":"🔴"}.get(risk_level,"⚪")
    driver = (f'由【亏损感受】决定：你选了「{risk_answer}」' if "不确定" not in risk_answer
              else '你选了"我不确定"，系统按经验和期限保守处理')
    st.info(f"{icon} **风险层级：{risk_level}**　　{driver}\n\n"
            "想改变结果，请修改上方【亏损感受】那道题。"
            "（不太能接受 → 低 / 阶段性波动 → 中 / 愿意承受较大波动 → 高）")
    return user


# ── Tab: 方向推荐 ─────────────────────────────────────────────────────────────
def render_domain_recommendations(user: UserProfile) -> List[DomainRecommendation]:
    domains = recommend_domains(user)
    section_title("适合你的 ETF 方向", "先确定方向，再比较具体产品")

    cols   = st.columns(3, gap="medium")
    colors = ["#3b82f6","#10b981","#f59e0b"]
    for col, d, color in zip(cols, domains, colors):
        with col:
            st.markdown(
                f'<div class="c-card domain-card">'
                f'<div style="color:{color};font-size:1.7rem;font-weight:800;line-height:1;margin-bottom:2px">{d.match_score:.0f}</div>'
                f'<div style="color:var(--text-d);font-size:0.65rem;text-transform:uppercase;letter-spacing:.08em;margin-bottom:10px">匹配分</div>'
                f'<div style="color:var(--text-h);font-weight:700;font-size:1rem;margin-bottom:3px">{d.name}</div>'
                f'<div style="color:var(--text-m);font-size:0.72rem;margin-bottom:10px">{d.role}</div>'
                f'<div style="color:var(--text-sec);font-size:0.82rem;line-height:1.55">{d.plain_intro}</div>'
                f'</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    top = domains[0]
    with st.expander(f"为什么优先推荐「{top.name}」？", expanded=True):
        c1, c2 = st.columns(2)
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
def render_candidates(df: pd.DataFrame, user: UserProfile, domains: List[DomainRecommendation]) -> None:
    section_title("ETF 候选对比", "先选方向，再在同方向内比较规模、费率、回撤")
    selected_name = st.selectbox("选择方向", [d.name for d in domains], label_visibility="collapsed")
    domain = next(d for d in domains if d.name==selected_name)
    candidates = rank_candidates(df, user, domain)

    for _, row in candidates.iterrows():
        etf_card_html(row, highlight=bool(row.get("domain_match")))

    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("AI 方向解释", expanded=False):
        col_btn, col_note = st.columns([2,5])
        with col_btn:
            if st.button("生成 AI 解释", type="primary", key="btn_narrative"):
                narrative, warning = make_narrative(user, domain, candidates)
                st.session_state["narrative_text"]    = narrative
                st.session_state["narrative_warning"] = warning
        with col_note:
            st.caption("AI 根据你的偏好和候选数据生成解释，不构成投资建议")
        if "narrative_text" in st.session_state:
            if st.session_state.get("narrative_warning"): st.caption(st.session_state["narrative_warning"])
            st.markdown(st.session_state["narrative_text"])
    with st.expander("查看完整对比数据"):
        st.dataframe(candidates[["code","name","category","quality_score","candidate_score",
                                  "total_fee","age_years","max_drawdown","volatility"]],
                     use_container_width=True)


# ── Tab: 质量评估 ─────────────────────────────────────────────────────────────
def render_quality_detail(df: pd.DataFrame, prices: pd.DataFrame, domains: List[DomainRecommendation]) -> None:
    section_title("单只 ETF 质量评估", "质量分不是买卖信号，只是系统性检查规模、费率、风险和数据完整性")

    selected = st.selectbox("选择 ETF", df["code"]+"  "+df["name"], label_visibility="collapsed")
    code = selected.split("  ")[0]
    row  = df.loc[df["code"]==code].iloc[0]
    domain = domains[0] if domains else None

    # ── 4 metric chips ──
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("最大回撤",    fmt_pct(row.get("max_drawdown")))
    m2.metric("年化波动率",  fmt_pct(row.get("volatility")))
    m3.metric("近一年收益",  fmt_pct(row.get("one_year_return")))
    m4.metric("成立年限",    f"{row.get('age_years',0):.1f} 年" if pd.notna(row.get("age_years")) else "—")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Score bars (left) + Price chart (right) ──
    left, right = st.columns([5, 7], gap="large")

    with left:
        score_bars_html(row)

    with right:
        price_df = prices[prices["code"]==code].copy() if not prices.empty else pd.DataFrame()
        st.markdown(
            '<div class="c-card" style="padding:16px 18px">'
            '<div style="color:var(--text-m);font-size:0.7rem;text-transform:uppercase;letter-spacing:.06em;margin-bottom:12px">近一年收盘价</div>',
            unsafe_allow_html=True)
        if not price_df.empty:
            price_chart(price_df)
        else:
            st.info("暂无行情，点击侧边栏「拉取最新行情」")
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Risk flags ──
    flags = risk_flags(row)
    flag_rows = "".join(
        f'<div style="display:flex;gap:8px;align-items:flex-start;margin-bottom:8px">'
        f'<span style="color:#f59e0b;font-size:0.9rem;flex-shrink:0;margin-top:1px">⚠</span>'
        f'<span style="color:var(--text-sec);font-size:0.85rem;line-height:1.5">{f}</span></div>'
        for f in flags)
    st.markdown(
        f'<div class="c-card" style="padding:16px 20px;margin-bottom:14px">'
        f'<div style="color:var(--text-p);font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:.06em;margin-bottom:12px">风险提示</div>'
        f'{flag_rows}</div>',
        unsafe_allow_html=True)

    # ── AI analysis ──
    st.markdown('<div class="c-card ai-box">', unsafe_allow_html=True)
    ai_col, note_col = st.columns([2,4])
    with ai_col:
        run_ai = st.button("🤖 一键 AI 深度分析", type="primary", key=f"ai_{code}", use_container_width=True)
    with note_col:
        if ai_available(): st.caption(f"使用 {st.session_state.get('ai_provider','AI')} · {st.session_state.get('ai_model','')} 流式生成")
        else: st.caption("请先在左侧侧边栏填写 AI 服务商信息")

    ai_key = f"ai_result_{code}"
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
            '<div style="color:var(--text-m);font-size:0.8rem;margin-bottom:10px">点击按钮后 AI 将流式生成分析报告。以下是预设自查问题：</div>'
            + "".join(f'<div style="color:var(--text-m);font-size:0.84rem;margin-bottom:6px;padding-left:2px">· {q}</div>' for q in qs),
            unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)


# ── Tab: 数据 & 术语 ──────────────────────────────────────────────────────────
def render_universe_and_terms(df: pd.DataFrame) -> None:
    section_title("样本池 & 术语", "10 只上证 ETF 样本 · 数据来自 Excel + Yahoo Finance 缓存")
    display = df[["code","name","category","style","quality_score","total_fee","age_years","scale_proxy","price_points"]].copy()
    display["scale_proxy"] = display["scale_proxy"].apply(fmt_money)
    st.dataframe(display, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    section_title("术语说明")
    cols = st.columns(2)
    for i,(term,desc) in enumerate(HELP_TEXT.items()):
        with cols[i%2]:
            st.markdown(
                f'<div class="term-card">'
                f'<div style="color:var(--accent);font-size:0.8rem;font-weight:600;margin-bottom:4px">{term}</div>'
                f'<div style="color:var(--text-m);font-size:0.82rem;line-height:1.55">{desc}</div>'
                f'</div>', unsafe_allow_html=True)


# ── Tab: 调试 ─────────────────────────────────────────────────────────────────
def render_debug(profile, prices, df):
    section_title("调试信息","开发用，后续可隐藏")
    def fmt_mtime(p): return (pd.Timestamp(p.stat().st_mtime,unit="s",tz="UTC")
                               .tz_convert(CST).strftime("%Y-%m-%d %H:%M:%S（北京时间）"))
    c1,c2 = st.columns(2)
    with c1:
        if PROFILE_CACHE.exists(): st.success(f"profile 缓存 · {fmt_mtime(PROFILE_CACHE)}")
        else: st.error("profile 缓存不存在")
    with c2:
        if PRICE_CACHE.exists(): st.success(f"prices 缓存 · {fmt_mtime(PRICE_CACHE)}")
        else: st.error("prices 缓存不存在")
    if not prices.empty and "fetch_time" in prices.columns and prices["fetch_time"].notna().any():
        st.info(f"行情最近拉取时间：{prices['fetch_time'].dropna().iloc[0]}")
    else: st.info("行情来自本地缓存（未记录拉取时间）")
    st.markdown("#### 各 ETF 数据点统计")
    if not prices.empty:
        st.dataframe(prices.groupby("code").agg(数据点数=("close","count"),最早日期=("date","min"),
            最新日期=("date","max"),最新收盘价=("close","last")).reset_index(), use_container_width=True)
    with st.expander("ETF 基本信息（完整）"): st.dataframe(profile, use_container_width=True)
    with st.expander("行情缓存（最近 20 条）"):
        if not prices.empty: st.dataframe(prices.sort_values("date",ascending=False).head(20), use_container_width=True)
    with st.expander("综合打分表（完整）"): st.dataframe(df, use_container_width=True)


# ── Main ──────────────────────────────────────────────────────────────────────
def main() -> None:
    st.set_page_config(page_title="ETF 投资助手", layout="wide", page_icon="📊")

    refresh_prices, refresh_profile = sidebar_controls()

    # Inject CSS after toggle so dark_mode state is settled for this run
    st.markdown(BASE_CSS, unsafe_allow_html=True)
    if not st.session_state.get("dark_mode", True):
        st.markdown(LIGHT_CSS, unsafe_allow_html=True)

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

    df = add_metrics(profile, prices)
    default_user = UserProfile("我是新手","低风险","3年以上","长期配置，慢慢积累","我想先要一个稳一点的底座",50000.0)
    user    = st.session_state.get("user_profile", default_user)
    domains = recommend_domains(user)

    # Page header
    st.markdown(
        '<div style="padding:6px 0 24px">'
        '<div style="font-size:1.8rem;font-weight:800;color:var(--text-h);letter-spacing:-0.02em">ETF 投资助手</div>'
        '<div style="color:var(--text-d);font-size:0.85rem;margin-top:5px">'
        '基于规则评分 + AI 解释 · 仅用于学习和决策辅助，不构成投资建议</div></div>',
        unsafe_allow_html=True)

    tab_profile, tab_domain, tab_candidates, tab_detail, tab_data, tab_debug = st.tabs(
        ["偏好问卷","方向推荐","ETF 对比","质量评估","数据 & 术语","🔧 调试"])

    with tab_profile:
        user    = render_profile_form()
        domains = recommend_domains(user)
    with tab_domain:
        domains = render_domain_recommendations(user)
    with tab_candidates:
        render_candidates(df, user, domains)
    with tab_detail:
        render_quality_detail(df, prices, domains)
    with tab_data:
        render_universe_and_terms(df)
    with tab_debug:
        render_debug(profile, prices, df)


if __name__ == "__main__":
    main()
