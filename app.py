"""ETF Investment Assistant"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Generator, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
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
DEEPSEEK_MODEL = os.environ.get("DEEPSEEK_MODEL", "deepseek-chat")

AI_PROVIDERS: Dict[str, Dict[str, str]] = {
    "DeepSeek": {"base_url": "https://api.deepseek.com", "default_model": "deepseek-chat"},
    "Gemini (Google)": {
        "base_url": "https://generativelanguage.googleapis.com/v1beta/openai",
        "default_model": "gemini-2.0-flash",
    },
    "OpenAI": {"base_url": "https://api.openai.com/v1", "default_model": "gpt-4o-mini"},
    "Moonshot (Kimi)": {"base_url": "https://api.moonshot.cn/v1", "default_model": "moonshot-v1-8k"},
    "智谱 GLM": {"base_url": "https://open.bigmodel.cn/api/paas/v4", "default_model": "glm-4-flash"},
    "硅基流动": {
        "base_url": "https://api.siliconflow.cn/v1",
        "default_model": "Qwen/Qwen2.5-7B-Instruct",
    },
}

# ── CSS ───────────────────────────────────────────────────────────────────────
# Base CSS: dark-mode defaults as CSS custom properties + component classes.
BASE_CSS = """<style>
:root {
  --bg-page:    #080d18;
  --bg-sidebar: #0c1322;
  --bg-card:    #0d1526;
  --border-s:   #162035;
  --border-c:   #1e2d4a;
  --text-h:     #e8f0ff;
  --text-p:     #c8d6e8;
  --text-sec:   #94a3b8;
  --text-m:     #475569;
  --text-d:     #3d5166;
  --accent:     #60a5fa;
  --accent-btn: #1d4ed8;
}

#MainMenu, footer { visibility: hidden; }
header { visibility: hidden; }
/* Keep sidebar expand arrow visible even when header is hidden */
[data-testid="collapsedControl"] { visibility: visible !important; }

.stApp { background-color: var(--bg-page); }
[data-testid="stSidebar"] {
  background-color: var(--bg-sidebar);
  border-right: 1px solid var(--border-s);
}

.stTabs [data-baseweb="tab-list"] {
  gap: 0; border-bottom: 1px solid var(--border-s);
  background: transparent; padding: 0 2px;
}
.stTabs [data-baseweb="tab"] {
  background: transparent !important; color: var(--text-m);
  height: 44px; padding: 0 22px; font-size: 0.88rem;
  border-radius: 0; border-bottom: 2px solid transparent; margin-bottom: -1px;
}
.stTabs [aria-selected="true"] {
  color: var(--accent) !important;
  border-bottom-color: var(--accent) !important;
  background: transparent !important;
}
.stTabs [data-baseweb="tab-panel"] { padding-top: 24px; }

div[data-testid="stButton"] > button {
  background: var(--bg-card); color: var(--text-sec);
  border: 1px solid var(--border-c); border-radius: 8px;
  font-size: 0.88rem; transition: all 0.18s;
}
div[data-testid="stButton"] > button:hover {
  border-color: var(--accent); color: var(--accent); background: var(--bg-card);
}
div[data-testid="stButton"] > button[kind="primary"] {
  background: linear-gradient(135deg, var(--accent-btn), #2563eb);
  color: #fff; border: none; font-weight: 600;
}
div[data-testid="stButton"] > button[kind="primary"]:hover {
  opacity: 0.88; border: none; color: #fff;
}

[data-testid="stMetricValue"] { font-size: 1.4rem; font-weight: 700; color: var(--text-h); }
[data-testid="stMetricLabel"] {
  font-size: 0.72rem; color: var(--text-m);
  text-transform: uppercase; letter-spacing: 0.06em;
}

.stSelectbox > div > div {
  background: var(--bg-card) !important;
  border-color: var(--border-c) !important; border-radius: 8px !important;
}
.stTextInput > div > div > input,
[data-testid="stNumberInput"] input {
  background: var(--bg-card) !important; border-color: var(--border-c) !important;
  border-radius: 8px !important; color: var(--text-p) !important;
}

.stAlert { border-radius: 8px !important; }
hr { border-color: var(--border-s) !important; opacity: 0.6; }
[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }
details > summary {
  background: var(--bg-card) !important; border: 1px solid var(--border-c) !important;
  border-radius: 8px !important; padding: 10px 16px !important;
  color: var(--text-sec) !important; font-size: 0.88rem !important;
}
h1 { color: var(--text-h) !important; letter-spacing: -0.025em; }
h2, h3 { color: var(--text-p) !important; }
p { color: var(--text-sec); }

/* ── Component classes ── */
.etf-card {
  position: relative; background: var(--bg-card);
  border: 1px solid var(--border-c); border-radius: 12px;
  padding: 18px 22px; margin-bottom: 10px;
}
.etf-card--hl { border-color: var(--accent-btn) !important; }
.etf-card__name { font-size: 1.0rem; font-weight: 600; color: var(--text-h); }
.etf-card__meta { color: var(--text-d); font-size: 0.8rem; margin-top: 3px; }
.etf-card__badge {
  position: absolute; top: 14px; right: 14px;
  background: rgba(29,78,216,0.08); color: var(--accent);
  padding: 2px 10px; border-radius: 4px; font-size: 0.7rem;
  font-weight: 600; border: 1px solid rgba(29,78,216,0.25);
}
.etf-card__metrics {
  display: grid; grid-template-columns: repeat(5, 1fr);
  gap: 10px; border-top: 1px solid var(--border-s); padding-top: 12px;
}
.etf-card__label {
  color: var(--text-d); font-size: 0.67rem;
  text-transform: uppercase; letter-spacing: .06em; margin-bottom: 3px;
}
.etf-card__val { color: var(--text-p); font-size: 0.88rem; font-weight: 500; }
.domain-card {
  background: var(--bg-card); border: 1px solid var(--border-c);
  border-radius: 12px; padding: 20px;
}
.info-box {
  background: var(--bg-card); border: 1px solid var(--border-c);
  border-radius: 10px; padding: 16px 20px; margin-bottom: 16px;
}
.term-card {
  background: var(--bg-card); border: 1px solid var(--border-c);
  border-radius: 8px; padding: 12px 16px; margin-bottom: 8px;
}
.ai-box {
  background: var(--bg-card); border: 1px solid var(--border-c);
  border-radius: 10px; padding: 20px 22px;
}
</style>"""

# Light mode: override CSS variables only; all component classes stay the same.
LIGHT_CSS = """<style>
:root {
  --bg-page:    #f8fafc;
  --bg-sidebar: #f1f5f9;
  --bg-card:    #ffffff;
  --border-s:   #e2e8f0;
  --border-c:   #cbd5e1;
  --text-h:     #0f172a;
  --text-p:     #1e293b;
  --text-sec:   #334155;
  --text-m:     #64748b;
  --text-d:     #94a3b8;
  --accent:     #2563eb;
  --accent-btn: #1d4ed8;
}
.stApp { color: #1e293b; }
[data-testid="stMarkdownContainer"] p { color: #334155; }
.stSelectbox > div > div { color: #1e293b !important; }
.stTextInput > div > div > input,
[data-testid="stNumberInput"] input { color: #1e293b !important; }
</style>"""

HELP_TEXT = {
    "宽基": "宽基ETF通常跟踪沪深300、上证50、中证500这类大盘或综合指数，买的是一篮子股票，更适合作为长期配置的核心仓位。",
    "红利": "红利ETF偏向高分红、现金流较稳定的公司，常被保守投资者关注，但也会受利率、周期和行业集中影响。",
    "行业主题": "行业主题ETF集中投向某个行业或主题，例如消费、金融、半导体。机会更集中，波动和回撤通常也更大。",
    "回撤": "回撤可以理解为从阶段高点跌下来多少。比如最大回撤-30%，意味着过去某段时间里最高点买入会一度亏约30%。",
    "波动率": "波动率衡量价格上下震荡的剧烈程度。波动率高不等于一定不好，但更考验持有耐心和仓位控制。",
    "费率": "费率是长期持有时持续付出的管理费、托管费。单年看不大，但多年复利下来会影响实际收益。",
    "规模": "规模越大通常流动性和稳定性更好。规模很小的ETF要额外关注成交不活跃或清盘风险。",
    "流动性": "流动性指买卖是否方便、价差是否大。流动性差时，买入或卖出可能付出额外隐性成本。",
}


# ── Data classes ──────────────────────────────────────────────────────────────
@dataclass
class UserProfile:
    experience: str
    risk_level: str
    horizon: str
    goal: str
    preference: str
    amount: float


@dataclass
class DomainRecommendation:
    name: str
    match_score: float
    role: str
    plain_intro: str
    suitable_for: str
    return_source: str
    key_risks: str
    common_mistake: str
    category_keywords: Tuple[str, ...]
    risk_hint: str


DOMAIN_LIBRARY: Dict[str, DomainRecommendation] = {
    "宽基核心": DomainRecommendation(
        name="宽基核心", match_score=0, role="核心仓位",
        plain_intro="先用一篮子代表性股票打底，不把胜负押在单一行业上。",
        suitable_for="适合新手、长期配置、希望先把组合底座搭稳的投资者。",
        return_source="主要来自市场整体增长、指数成分股盈利改善和估值修复。",
        key_risks="市场整体下跌时也会回撤，不能理解成保本工具。",
        common_mistake="只看短期涨幅排名，忽略宽基更适合长期、分批和纪律化持有。",
        category_keywords=("宽基", "核心"), risk_hint="适合作为组合底座，但仍要控制买入节奏和总仓位。",
    ),
    "红利/低波": DomainRecommendation(
        name="红利/低波", match_score=0, role="稳健核心或防守仓位",
        plain_intro="更关注分红、现金流和相对稳健的公司，目标是少一点刺激，多一点纪律。",
        suitable_for="适合不喜欢大起大落、希望降低组合波动的新手或保守型投资者。",
        return_source="主要来自股息、企业稳定盈利和低估值修复。",
        key_risks="红利不等于无风险，可能有行业集中、周期下行和股息下降风险。",
        common_mistake="把高股息率简单理解成高收益，忽略股价下跌可能抵消分红。",
        category_keywords=("红利", "低波", "价值", "金融"),
        risk_hint="当前样本池红利ETF不足，demo会用价值/金融类作为近似候选展示。",
    ),
    "行业主题": DomainRecommendation(
        name="行业主题", match_score=0, role="卫星仓位",
        plain_intro="集中投向某个行业或主题，适合表达观点，但不适合新手一上来重仓。",
        suitable_for="适合能承受较大波动、理解行业周期、愿意把它控制在小比例仓位的投资者。",
        return_source="主要来自行业景气度上行、政策催化、估值扩张或盈利改善。",
        key_risks="行业热度高时容易追涨，景气反转时回撤也可能很深。",
        common_mistake="因为最近涨得多就买入，却没有想清楚退出规则和仓位上限。",
        category_keywords=("行业", "主题", "成长"),
        risk_hint="更适合作为小比例卫星仓位，不建议作为保守投资者的核心仓位。",
    ),
}


# ── Data loading ──────────────────────────────────────────────────────────────
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {
        "证券代码": "code", "证券名称": "name", "基金类型": "fund_type",
        "单位净值\n[交易日期] 最新[单位]元": "nav",
        "基金份额\n[交易日期] 最新\n[单位]  份": "shares",
        "投资风格\n[年度] 2023\n[报告期] 中报": "style",
        "基金管理人": "manager", "基金托管人": "custodian",
        "管理费率[单位]%": "management_fee", "托管费率[单位]%": "custody_fee",
        "基金经理(现任)": "fund_manager", "基金存续期[单位]年": "duration_years",
        "基金成立日": "inception_date", "基金到期日": "maturity_date",
    }
    return df.rename(columns=mapping)


def parse_inception(value: Any) -> Optional[pd.Timestamp]:
    if pd.isna(value):
        return None
    text = str(value).split(".")[0].strip()
    if len(text) != 8:
        return None
    parsed = pd.to_datetime(text, format="%Y%m%d", errors="coerce")
    return None if pd.isna(parsed) else parsed


def classify_etf(row: pd.Series) -> str:
    text = f"{row.get('name', '')} {row.get('style', '')}"
    if any(k in text for k in ["沪深300", "上证50", "中证500", "国企"]):
        return "宽基/核心"
    if any(k in text for k in ["消费", "金融"]):
        return "行业/主题"
    if "成长" in text:
        return "成长风格"
    if "价值" in text:
        return "价值风格"
    return "其他"


def load_profile_from_excel() -> pd.DataFrame:
    df = pd.read_excel(EXCEL_PATH).head(SAMPLE_SIZE)
    df = normalize_columns(df)
    keep = ["code", "name", "fund_type", "nav", "shares", "style", "manager",
            "custodian", "management_fee", "custody_fee", "fund_manager", "inception_date"]
    df = df[keep].copy()
    df["inception"] = df["inception_date"].apply(parse_inception)
    df["age_years"] = df["inception"].apply(
        lambda x: np.nan if x is None else max((pd.Timestamp.today() - x).days / 365.25, 0)
    )
    df["total_fee"] = (
        pd.to_numeric(df["management_fee"], errors="coerce").fillna(0)
        + pd.to_numeric(df["custody_fee"], errors="coerce").fillna(0)
    )
    df["scale_proxy"] = pd.to_numeric(df["nav"], errors="coerce") * pd.to_numeric(df["shares"], errors="coerce")
    df["category"] = df.apply(classify_etf, axis=1)
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
    code_list = list(codes)
    frames: List[pd.DataFrame] = []
    errors: List[str] = []
    for i, code in enumerate(code_list):
        if progress_placeholder is not None:
            progress_placeholder.info(f"正在拉取 {i + 1}/{len(code_list)}：{code} …")
        try:
            clean = str(code).split(".")[0]
            suffix = str(code).split(".")[-1].upper() if "." in str(code) else "SH"
            yf_code = clean + (".SS" if suffix == "SH" else ".SZ")
            hist = yf.Ticker(yf_code).history(period="1y")
            if hist.empty:
                errors.append(f"{code}: 无数据")
                continue
            frame = pd.DataFrame({
                "code": code,
                "date": pd.to_datetime(hist.index.tz_localize(None)),
                "close": hist["Close"].values,
            }).dropna(subset=["date", "close"]).sort_values("date")
            frames.append(frame)
        except Exception as e:
            errors.append(f"{code}: {type(e).__name__}")
            continue
    if not frames:
        return pd.DataFrame(), f"拉取失败：{'; '.join(errors[:3])}"
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
def max_drawdown(close: pd.Series) -> float:
    close = close.dropna()
    if close.empty:
        return np.nan
    return float((close / close.cummax() - 1).min())


def annual_volatility(close: pd.Series) -> float:
    r = close.pct_change().dropna()
    return np.nan if r.empty else float(r.std() * np.sqrt(252))


def one_year_return(close: pd.Series) -> float:
    close = close.dropna()
    return np.nan if len(close) < 2 else float(close.iloc[-1] / close.iloc[0] - 1)


def add_metrics(profile: pd.DataFrame, prices: pd.DataFrame) -> pd.DataFrame:
    df = profile.copy()
    metric_rows = []
    for code in df["code"]:
        close = prices.loc[prices["code"] == code, "close"] if not prices.empty else pd.Series(dtype=float)
        metric_rows.append({
            "code": code,
            "max_drawdown": max_drawdown(close),
            "volatility": annual_volatility(close),
            "one_year_return": one_year_return(close),
            "price_points": int(close.dropna().shape[0]),
        })
    df = df.merge(pd.DataFrame(metric_rows), on="code", how="left")
    scores = pd.DataFrame([score_etf(row) for _, row in df.iterrows()])
    return pd.concat([df.reset_index(drop=True), scores], axis=1)


def score_size(scale: float) -> float:
    if pd.isna(scale): return 8
    if scale >= 50_000_000_000: return 25
    if scale >= 10_000_000_000: return 22
    if scale >= 1_000_000_000: return 16
    if scale >= 100_000_000: return 9
    return 4


def score_fee(total_fee: float) -> float:
    if pd.isna(total_fee): return 8
    if total_fee <= 0.20: return 20
    if total_fee <= 0.30: return 17
    if total_fee <= 0.60: return 12
    if total_fee <= 1.00: return 7
    return 3


def score_age(age: float) -> float:
    if pd.isna(age): return 6
    if age >= 5: return 20
    if age >= 3: return 16
    if age >= 1: return 10
    if age >= 0.5: return 6
    return 3


def score_risk(drawdown: float, volatility: float) -> float:
    if pd.isna(drawdown) or pd.isna(volatility): return 8
    score = 20
    if drawdown < -0.35: score -= 8
    elif drawdown < -0.25: score -= 5
    elif drawdown < -0.15: score -= 2
    if volatility > 0.35: score -= 7
    elif volatility > 0.25: score -= 4
    elif volatility > 0.18: score -= 2
    return max(score, 3)


def score_data(points: int, static_missing: int) -> float:
    score = 15
    if points < 120: score -= 6
    elif points < 220: score -= 3
    score -= min(static_missing * 2, 6)
    return max(score, 3)


def score_etf(row: pd.Series) -> Dict[str, float]:
    static_cols = ["nav", "shares", "management_fee", "custody_fee", "inception"]
    static_missing = sum(pd.isna(row.get(col)) for col in static_cols)
    size = score_size(row.get("scale_proxy"))
    fee = score_fee(row.get("total_fee"))
    age = score_age(row.get("age_years"))
    risk = score_risk(row.get("max_drawdown"), row.get("volatility"))
    data = score_data(int(row.get("price_points", 0) or 0), static_missing)
    return {
        "quality_score": round(size + fee + age + risk + data, 1),
        "size_score": size, "fee_score": fee, "age_score": age,
        "risk_score": risk, "data_score": data,
    }


# ── Recommendation logic ──────────────────────────────────────────────────────
def infer_risk_level(risk_answer: str, experience: str, horizon: str) -> str:
    if "不确定" in risk_answer:
        return "低风险" if (experience == "我是新手" or horizon == "1年以内") else "中风险"
    if "不太能接受" in risk_answer:
        return "低风险"
    if "阶段性波动" in risk_answer:
        return "中风险"
    return "高风险"


def recommend_domains(user: UserProfile) -> List[DomainRecommendation]:
    scores = {"宽基核心": 50.0, "红利/低波": 45.0, "行业主题": 30.0}
    if user.risk_level == "低风险":
        scores["红利/低波"] += 24; scores["宽基核心"] += 18; scores["行业主题"] -= 20
    elif user.risk_level == "中风险":
        scores["宽基核心"] += 20; scores["红利/低波"] += 10; scores["行业主题"] += 4
    else:
        scores["行业主题"] += 24; scores["宽基核心"] += 10
    if user.experience == "我是新手":
        scores["宽基核心"] += 14; scores["红利/低波"] += 8; scores["行业主题"] -= 8
    if user.horizon == "1年以内":
        scores["红利/低波"] += 12; scores["行业主题"] -= 12
    elif user.horizon == "3年以上":
        scores["宽基核心"] += 10; scores["行业主题"] += 3
    if user.goal == "希望稳一点，少一些大起大落":
        scores["红利/低波"] += 20; scores["行业主题"] -= 10
    elif user.goal == "长期配置，慢慢积累":
        scores["宽基核心"] += 20
    elif user.goal == "想抓住某些行业机会":
        scores["行业主题"] += 20
    if user.preference == "我想先要一个稳一点的底座":
        scores["宽基核心"] += 16
    elif user.preference == "我更关注分红和稳健":
        scores["红利/低波"] += 16
    elif user.preference == "我想了解行业主题机会":
        scores["行业主题"] += 16
    result = [
        DomainRecommendation(**{**base.__dict__, "match_score": round(scores[name], 1)})
        for name, base in DOMAIN_LIBRARY.items()
    ]
    return sorted(result, key=lambda x: x.match_score, reverse=True)


def etf_matches_domain(row: pd.Series, domain: DomainRecommendation) -> bool:
    text = f"{row.get('name', '')} {row.get('style', '')} {row.get('category', '')}"
    return any(k in text for k in domain.category_keywords)


def candidate_score(row: pd.Series, user: UserProfile, domain: DomainRecommendation) -> float:
    score = float(row["quality_score"])
    if etf_matches_domain(row, domain): score += 18
    if user.risk_level == "低风险":
        if row.get("total_fee", 99) <= 0.20: score += 8
        if pd.notna(row.get("max_drawdown")) and row.get("max_drawdown") < -0.25: score -= 12
        if pd.notna(row.get("volatility")) and row.get("volatility") > 0.25: score -= 10
        if "行业" in str(row.get("category", "")): score -= 6
    elif user.risk_level == "高风险" and "行业" in str(row.get("category", "")):
        score += 5
    if user.horizon == "1年以内" and pd.notna(row.get("volatility")) and row.get("volatility") > 0.25:
        score -= 8
    return round(score, 1)


def rank_candidates(
    df: pd.DataFrame, user: UserProfile, domain: DomainRecommendation
) -> pd.DataFrame:
    ranked = df.copy()
    ranked["candidate_score"] = ranked.apply(lambda r: candidate_score(r, user, domain), axis=1)
    ranked["domain_match"] = ranked.apply(lambda r: etf_matches_domain(r, domain), axis=1)
    return ranked.sort_values(["domain_match", "candidate_score"], ascending=[False, False]).head(3)


def risk_flags(row: pd.Series) -> List[str]:
    flags = []
    if row.get("age_years", 99) < 1: flags.append("成立时间较短，成立以来收益率参考价值有限。")
    if row.get("total_fee", 0) > 0.60: flags.append("费率偏高，长期持有成本会持续累积。")
    if row.get("scale_proxy", np.inf) < 1_000_000_000: flags.append("规模偏小，需关注流动性和清盘风险。")
    if pd.notna(row.get("max_drawdown")) and row.get("max_drawdown") < -0.25:
        flags.append("近一年最大回撤较大，需确认能否承受。")
    if pd.notna(row.get("volatility")) and row.get("volatility") > 0.25:
        flags.append("近一年波动率较高，不适合保守型核心仓位。")
    if not flags: flags.append("未发现突出单项风险，但仍需控制仓位。")
    return flags


def self_check_questions(row: pd.Series, domain: Optional[DomainRecommendation] = None) -> List[str]:
    category = domain.name if domain else row.get("category", "这类ETF")
    return [
        f'我是否理解这只ETF为什么归到"{category}"，主要风险来源是什么？',
        "如果短期回撤20%-30%，我是否还能按原计划持有？",
        "我买入它是长期配置需要，还是被近期涨幅或热度吸引？",
        "我是否比较过同类ETF的规模、费率、成立时间和流动性？",
        "这只ETF在我的总资产中应该是核心仓位还是小比例卫星仓位？",
    ]


def build_fallback_narrative(
    user: UserProfile, domain: DomainRecommendation, candidates: pd.DataFrame
) -> str:
    names = "、".join(candidates["name"].head(3).tolist()) if not candidates.empty else "暂无合适候选"
    return (
        f'根据你的偏好，当前更适合先从"{domain.name}"看起。'
        f"这个方向的定位是{domain.role}：{domain.plain_intro}"
        f"样本池里可以先重点比较：{names}。"
        f"推荐只是帮助缩小观察范围，不是买入建议；决策前还要看回撤、费率、规模和能否长期持有。"
    )


# ── AI functions ──────────────────────────────────────────────────────────────
def ai_available() -> bool:
    return bool(st.session_state.get("ai_api_key") or os.environ.get("DEEPSEEK_API_KEY"))


def get_ai_config() -> Tuple[str, str, str]:
    api_key = st.session_state.get("ai_api_key") or os.environ.get("DEEPSEEK_API_KEY", "")
    base_url = (st.session_state.get("ai_base_url") or DEEPSEEK_BASE_URL).rstrip("/")
    model = st.session_state.get("ai_model") or DEEPSEEK_MODEL
    return api_key, base_url, model


def to_jsonable(value: Any) -> Any:
    if isinstance(value, dict): return {str(k): to_jsonable(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)): return [to_jsonable(v) for v in value]
    if isinstance(value, np.integer): return int(value)
    if isinstance(value, np.floating): return None if pd.isna(value) else float(value)
    if isinstance(value, pd.Timestamp): return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and pd.isna(value): return None
    return value


def stream_ai_analysis(
    row: pd.Series,
    user: Optional[UserProfile] = None,
    domain: Optional[DomainRecommendation] = None,
) -> Generator[str, None, None]:
    api_key, base_url, model = get_ai_config()
    if not api_key:
        yield "请先在左侧侧边栏填写 AI 服务商和 API Key。"
        return

    context: Dict[str, Any] = {
        "etf_name": row.get("name"), "etf_code": row.get("code"),
        "category": row.get("category"), "quality_score": row.get("quality_score"),
        "total_fee_pct": row.get("total_fee"), "age_years": row.get("age_years"),
        "max_drawdown": row.get("max_drawdown"), "volatility": row.get("volatility"),
        "one_year_return": row.get("one_year_return"), "scale": row.get("scale_proxy"),
        "risk_flags": risk_flags(row),
    }
    if user:
        context["user_risk_level"] = user.risk_level
        context["user_horizon"] = user.horizon
        context["user_experience"] = user.experience
    if domain:
        context["recommended_domain"] = domain.name

    system_prompt = (
        "你是一个谨慎的ETF投资教育助手。只基于用户提供的数据写分析，"
        "不得编造数据，不得预测收益，不得给出买入卖出建议。用中文输出，面向新手。"
    )
    user_prompt = (
        "请对以下ETF做三部分分析，使用Markdown格式：\n"
        "**1. 综合质量判断**（约80字，基于评分和各维度指标）\n"
        "**2. 主要风险提示**（约80字，结合用户风险偏好）\n"
        "**3. 购前自查清单**（3-5条简洁条目）\n\n"
        f"数据：{json.dumps(to_jsonable(context), ensure_ascii=False)}"
    )

    body = json.dumps({
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.4,
        "max_tokens": 700,
        "stream": True,
    }, ensure_ascii=False).encode("utf-8")

    try:
        with requests.post(
            f"{base_url}/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json; charset=utf-8",
            },
            data=body, stream=True, timeout=30,
        ) as resp:
            if not resp.ok:
                yield f"API 错误 {resp.status_code}: {resp.content.decode('utf-8', errors='replace')[:200]}"
                return
            for line in resp.iter_lines():
                if not line:
                    continue
                if line.startswith(b"data: "):
                    chunk_data = line[6:]
                    if chunk_data.strip() == b"[DONE]":
                        return
                    try:
                        chunk = json.loads(chunk_data.decode("utf-8"))
                        delta = chunk["choices"][0]["delta"].get("content", "")
                        if delta:
                            yield delta
                    except Exception:
                        continue
    except Exception as e:
        yield f"\n\n（输出中断：{e}）"


def call_ai_narrative(payload: Dict[str, Any]) -> Tuple[Optional[str], Optional[str]]:
    api_key, base_url, model = get_ai_config()
    if not api_key:
        return None, "未设置 API Key，使用规则模板。"
    system_prompt = (
        "你是一个谨慎的ETF投资教育助手。只基于提供的事实写解释，"
        "不得编造数据，不得预测收益，不得给出投资建议。用中文输出，面向新手。"
    )
    user_prompt = (
        "请把以下事实润色成一段适合展示的解释。"
        "结构：1）为什么先看这个方向；2）候选ETF怎么比较；3）风险和自查提醒。控制在180字以内。\n\n"
        + json.dumps(to_jsonable(payload), ensure_ascii=False)
    )
    try:
        body = json.dumps({
            "model": model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "temperature": 0.3, "max_tokens": 500,
        }, ensure_ascii=False).encode("utf-8")
        resp = requests.post(
            f"{base_url}/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json; charset=utf-8",
            },
            data=body, timeout=20,
        )
        if not resp.ok:
            return None, f"HTTP {resp.status_code}: {resp.content.decode('utf-8', errors='replace')[:200]}"
        data = json.loads(resp.content.decode("utf-8"))
        return data["choices"][0]["message"]["content"].strip(), None
    except Exception as exc:
        provider = st.session_state.get("ai_provider", "AI")
        return None, f"{provider} 调用失败：{exc}"


def make_narrative(
    user: UserProfile, domain: DomainRecommendation, candidates: pd.DataFrame
) -> Tuple[str, Optional[str]]:
    candidate_facts = [
        {
            "code": r["code"], "name": r["name"], "quality_score": r["quality_score"],
            "fee": r.get("total_fee"), "age_years": r.get("age_years"),
            "max_drawdown": r.get("max_drawdown"), "volatility": r.get("volatility"),
            "risk_flags": risk_flags(r),
        }
        for _, r in candidates.head(3).iterrows()
    ]
    payload = {
        "user_profile": user.__dict__,
        "domain": {
            "name": domain.name, "role": domain.role, "intro": domain.plain_intro,
            "suitable_for": domain.suitable_for, "return_source": domain.return_source,
            "key_risks": domain.key_risks, "common_mistake": domain.common_mistake,
        },
        "candidates": candidate_facts,
        "guardrail": "仅用于学习和决策辅助，不构成投资建议。",
    }
    text, warning = call_ai_narrative(payload)
    return (text, None) if text else (build_fallback_narrative(user, domain, candidates), warning)


# ── Formatting ────────────────────────────────────────────────────────────────
def format_percent(value: Any) -> str:
    return "—" if pd.isna(value) else f"{value * 100:.2f}%"


def format_money(value: Any) -> str:
    if pd.isna(value): return "—"
    v = float(value)
    if v >= 100_000_000: return f"{v / 100_000_000:.1f} 亿"
    if v >= 10_000: return f"{v / 10_000:.1f} 万"
    return f"{v:.0f}"


# ── Theme helper ──────────────────────────────────────────────────────────────
def _chart_style() -> Dict[str, str]:
    dark = st.session_state.get("dark_mode", True)
    if dark:
        return {"plot_bg": "#0d1526", "paper_bg": "#0d1526", "font": "#94a3b8", "grid": "#162035"}
    return {"plot_bg": "#ffffff", "paper_bg": "#f8fafc", "font": "#475569", "grid": "#e2e8f0"}


# ── UI primitives ─────────────────────────────────────────────────────────────
def section_title(title: str, subtitle: str = "") -> None:
    sub = (
        f'<p style="color:var(--text-m);margin:5px 0 0;font-size:0.85rem">{subtitle}</p>'
        if subtitle else ""
    )
    st.markdown(
        f'<div style="margin-bottom:24px">'
        f'<h2 style="margin:0;font-size:1.25rem;color:var(--text-h);font-weight:700">{title}</h2>'
        f'{sub}</div>',
        unsafe_allow_html=True,
    )


def etf_card_html(row: pd.Series, highlight: bool = False) -> None:
    score = float(row.get("quality_score", 0))
    score_color = "#10b981" if score >= 70 else "#f59e0b" if score >= 50 else "#ef4444"
    extra_class = " etf-card--hl" if highlight else ""
    badge = '<span class="etf-card__badge">方向匹配</span>' if highlight else ""
    fee = row.get("total_fee", float("nan"))
    dd = row.get("max_drawdown", float("nan"))
    vol = row.get("volatility", float("nan"))
    age = row.get("age_years", float("nan"))
    scale = row.get("scale_proxy", float("nan"))
    dd_color = "#ef4444" if not pd.isna(dd) and dd < -0.2 else "var(--text-p)"

    def cell(label: str, val: str, color: str = "var(--text-p)") -> str:
        return (
            f'<div><div class="etf-card__label">{label}</div>'
            f'<div class="etf-card__val" style="color:{color}">{val}</div></div>'
        )

    st.markdown(
        f'<div class="etf-card{extra_class}">'
        f'{badge}'
        f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px">'
        f'<div>'
        f'<div class="etf-card__name">{row.get("name", "")}</div>'
        f'<div class="etf-card__meta">{row.get("code", "")} &nbsp;·&nbsp; {row.get("category", "")}</div>'
        f'</div>'
        f'<div style="background:{score_color}18;color:{score_color};border:1px solid {score_color}40;'
        f'padding:5px 16px;border-radius:16px;font-weight:700;font-size:1.05rem;white-space:nowrap">'
        f'{score:.0f}<span style="font-size:0.6rem;opacity:0.7"> /100</span></div>'
        f'</div>'
        f'<div class="etf-card__metrics">'
        f'{cell("费率", f"{fee:.2f}%" if not pd.isna(fee) else "—")}'
        f'{cell("最大回撤", format_percent(dd), dd_color)}'
        f'{cell("年化波动", format_percent(vol))}'
        f'{cell("成立年限", f"{age:.1f} 年" if not pd.isna(age) else "—")}'
        f'{cell("规模", format_money(scale))}'
        f'</div></div>',
        unsafe_allow_html=True,
    )


# ── Sidebar ───────────────────────────────────────────────────────────────────
def sidebar_controls() -> Tuple[bool, bool]:
    with st.sidebar:
        # Theme toggle at the very top
        st.toggle("深色模式", value=st.session_state.get("dark_mode", True), key="dark_mode")

        st.markdown(
            '<div style="padding:4px 0 20px">'
            '<div style="font-size:1.1rem;font-weight:700;color:var(--text-h)">ETF 投资助手</div>'
            '<div style="font-size:0.75rem;color:var(--text-d);margin-top:3px">仅用于学习，不构成投资建议</div>'
            '</div>',
            unsafe_allow_html=True,
        )

        st.markdown(
            '<div style="color:var(--text-m);font-size:0.72rem;text-transform:uppercase;'
            'letter-spacing:.08em;margin-bottom:8px">AI 解释引擎</div>',
            unsafe_allow_html=True,
        )

        provider_names = list(AI_PROVIDERS.keys())
        selected_provider = st.selectbox(
            "服务商", provider_names, index=0, key="ai_provider", label_visibility="collapsed"
        )
        provider_cfg = AI_PROVIDERS[selected_provider]

        api_key_input = st.text_input(
            "API Key", type="password", placeholder="填入 API Key",
            value=st.session_state.get("ai_api_key", os.environ.get("DEEPSEEK_API_KEY", "")),
            key="_ai_key_input",
        )
        st.session_state["ai_api_key"] = api_key_input

        model_input = st.text_input(
            "模型", placeholder=provider_cfg["default_model"],
            value=st.session_state.get("ai_model", "") or provider_cfg["default_model"],
            key="_ai_model_input",
        )
        st.session_state["ai_model"] = model_input or provider_cfg["default_model"]
        st.session_state["ai_base_url"] = provider_cfg["base_url"]

        if ai_available():
            st.success(f"已启用 · {st.session_state['ai_model']}", icon="✓")
        else:
            st.caption("未填写 API Key，解释使用规则模板")

        st.divider()
        st.markdown(
            '<div style="color:var(--text-m);font-size:0.72rem;text-transform:uppercase;'
            'letter-spacing:.08em;margin-bottom:8px">行情数据</div>',
            unsafe_allow_html=True,
        )
        refresh_prices = st.button("拉取最新行情", use_container_width=True)
        refresh_profile = st.button("重建静态缓存", use_container_width=True)
        st.caption("数据源：Yahoo Finance · 自动缓存")
        return refresh_prices, refresh_profile


# ── Tab: 偏好问卷 ─────────────────────────────────────────────────────────────
def render_profile_form() -> UserProfile:
    section_title("偏好问卷", "尽量用自己的真实感受来选，不需要先懂专业术语")

    left, right = st.columns(2, gap="large")
    with left:
        experience = st.selectbox("投资经验", ["我是新手", "有一些经验", "比较熟悉ETF"])
        risk_answer = st.selectbox(
            "账户短期亏损时，你的感受更接近？",
            ["我不太能接受亏损", "能接受阶段性波动", "愿意承受较大波动换取机会", "我不确定"],
        )
        horizon = st.selectbox("这笔钱预计多久不用？", ["1年以内", "1-3年", "3年以上"])
    with right:
        goal = st.selectbox(
            "你更接近哪种目标？",
            ["希望稳一点，少一些大起大落", "长期配置，慢慢积累", "想抓住某些行业机会", "还没想清楚"],
        )
        preference = st.selectbox(
            "你现在更想先了解什么？",
            ["我想先要一个稳一点的底座", "我更关注分红和稳健", "我想了解行业主题机会", "我不确定"],
        )
        amount = st.number_input("计划投入金额（元）", min_value=1000.0, value=50000.0, step=1000.0)

    risk_level = infer_risk_level(risk_answer, experience, horizon)
    user = UserProfile(
        experience=experience, risk_level=risk_level, horizon=horizon,
        goal=goal, preference=preference, amount=amount,
    )
    st.session_state["user_profile"] = user

    risk_icon = {"低风险": "🟢", "中风险": "🟡", "高风险": "🔴"}.get(risk_level, "⚪")
    if "不确定" not in risk_answer:
        driver = f'由【亏损感受】决定：你选了「{risk_answer}」'
    else:
        driver = '你选了"我不确定"，系统按经验和期限保守处理'
    hint = "不太能接受 → 低 / 阶段性波动 → 中 / 愿意承受较大波动 → 高"
    st.info(
        f"{risk_icon} **风险层级：{risk_level}**　　{driver}\n\n"
        f"想改变结果，请修改上方【亏损感受】那道题。（{hint}）"
    )
    return user


# ── Tab: 领域推荐 ─────────────────────────────────────────────────────────────
def render_domain_recommendations(user: UserProfile) -> List[DomainRecommendation]:
    domains = recommend_domains(user)
    section_title("适合你的 ETF 方向", "先确定方向，再比较具体产品")

    cols = st.columns(3, gap="medium")
    colors = ["#3b82f6", "#10b981", "#f59e0b"]
    for col, domain, color in zip(cols, domains, colors):
        with col:
            st.markdown(
                f'<div class="domain-card">'
                f'<div style="color:{color};font-size:1.6rem;font-weight:800;margin-bottom:2px">'
                f'{domain.match_score:.0f}</div>'
                f'<div style="color:var(--text-m);font-size:0.7rem;text-transform:uppercase;'
                f'letter-spacing:.08em">匹配分</div>'
                f'<div style="color:var(--text-h);font-weight:700;font-size:1.05rem;margin:10px 0 4px">'
                f'{domain.name}</div>'
                f'<div style="color:var(--text-m);font-size:0.75rem;margin-bottom:8px">{domain.role}</div>'
                f'<div style="color:var(--text-sec);font-size:0.82rem;line-height:1.5">'
                f'{domain.plain_intro}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)
    top = domains[0]
    with st.expander(f"为什么优先推荐「{top.name}」？", expanded=True):
        c1, c2 = st.columns(2)
        c1.markdown(f"**适合人群**\n\n{top.suitable_for}")
        c1.markdown(f"**收益来源**\n\n{top.return_source}")
        c2.markdown(f"**主要风险**\n\n{top.key_risks}")
        c2.markdown(f"**常见误区**\n\n{top.common_mistake}")

    with st.expander("当前用户画像（确认偏好是否生效）"):
        r1, r2, r3 = st.columns(3)
        r1.metric("投资经验", user.experience)
        r1.metric("风险层级", user.risk_level)
        r2.metric("持有期限", user.horizon)
        r2.metric("投资目标", user.goal[:8] + "…" if len(user.goal) > 8 else user.goal)
        r3.metric("偏好方向", user.preference[:8] + "…" if len(user.preference) > 8 else user.preference)
        r3.metric("计划金额", f"{user.amount:,.0f} 元")

    return domains


# ── Tab: ETF候选 ──────────────────────────────────────────────────────────────
def render_candidates(
    df: pd.DataFrame, user: UserProfile, domains: List[DomainRecommendation]
) -> None:
    section_title("ETF 候选对比", "先选方向，再在同方向内比较规模、费率、回撤")

    domain_names = [d.name for d in domains]
    selected_name = st.selectbox("选择方向", domain_names, label_visibility="collapsed")
    domain = next(d for d in domains if d.name == selected_name)
    candidates = rank_candidates(df, user, domain)

    for _, row in candidates.iterrows():
        etf_card_html(row, highlight=bool(row.get("domain_match")))

    st.markdown("<br>", unsafe_allow_html=True)

    with st.expander("AI 方向解释", expanded=False):
        col_btn, col_note = st.columns([2, 5])
        with col_btn:
            if st.button("生成 AI 解释", type="primary", key="btn_narrative"):
                narrative, warning = make_narrative(user, domain, candidates)
                st.session_state["narrative_text"] = narrative
                st.session_state["narrative_warning"] = warning
        with col_note:
            st.caption("AI 根据你的偏好和候选数据生成解释，不构成投资建议")
        if "narrative_text" in st.session_state:
            if st.session_state.get("narrative_warning"):
                st.caption(st.session_state["narrative_warning"])
            st.markdown(st.session_state["narrative_text"])

    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("查看完整对比数据"):
        st.dataframe(
            candidates[[
                "code", "name", "category", "quality_score", "candidate_score",
                "total_fee", "age_years", "max_drawdown", "volatility",
            ]],
            use_container_width=True,
        )


# ── Tab: 质量评估 ─────────────────────────────────────────────────────────────
def render_quality_detail(
    df: pd.DataFrame, prices: pd.DataFrame, domains: List[DomainRecommendation]
) -> None:
    section_title("单只 ETF 质量评估", "质量分不是买卖信号，只是系统性检查规模、费率、风险和数据完整性")

    selected = st.selectbox("选择 ETF", df["code"] + "  " + df["name"], label_visibility="collapsed")
    code = selected.split("  ")[0]
    row = df.loc[df["code"] == code].iloc[0]
    domain = domains[0] if domains else None

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("质量总分", f"{row['quality_score']:.1f} / 100")
    m2.metric("最大回撤", format_percent(row.get("max_drawdown")))
    m3.metric("年化波动率", format_percent(row.get("volatility")))
    m4.metric("近一年收益", format_percent(row.get("one_year_return")))

    st.markdown("<br>", unsafe_allow_html=True)

    cs = _chart_style()
    left_col, right_col = st.columns([1, 1], gap="large")

    with left_col:
        score_df = pd.DataFrame({
            "维度": ["规模", "费率", "成立时间", "历史风险", "数据完整性"],
            "得分": [row["size_score"], row["fee_score"], row["age_score"], row["risk_score"], row["data_score"]],
            "满分": [25, 20, 20, 20, 15],
        })
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=score_df["维度"], y=score_df["满分"],
            marker_color=cs["grid"], name="满分", showlegend=False,
        ))
        fig.add_trace(go.Bar(
            x=score_df["维度"], y=score_df["得分"],
            marker_color="#3b82f6", name="得分", showlegend=False,
        ))
        fig.update_layout(
            barmode="overlay",
            plot_bgcolor=cs["plot_bg"], paper_bgcolor=cs["paper_bg"],
            font_color=cs["font"], height=240,
            margin=dict(l=0, r=0, t=8, b=0),
            xaxis=dict(gridcolor=cs["grid"]),
            yaxis=dict(gridcolor=cs["grid"]),
        )
        st.plotly_chart(fig, use_container_width=True)

    with right_col:
        price_df = prices[prices["code"] == code].copy() if not prices.empty else pd.DataFrame()
        if not price_df.empty:
            fig2 = px.line(price_df, x="date", y="close", color_discrete_sequence=["#3b82f6"])
            fig2.update_layout(
                plot_bgcolor=cs["plot_bg"], paper_bgcolor=cs["paper_bg"],
                font_color=cs["font"], height=240,
                margin=dict(l=0, r=0, t=8, b=0),
                xaxis=dict(gridcolor=cs["grid"], title=""),
                yaxis=dict(gridcolor=cs["grid"], title="收盘价"),
            )
            fig2.update_traces(line_width=1.8)
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("暂无历史行情，点击侧边栏「拉取最新行情」获取数据")

    st.markdown("<br>", unsafe_allow_html=True)

    flags = risk_flags(row)
    flag_html = "".join(
        f'<div style="display:flex;gap:8px;align-items:flex-start;margin-bottom:7px">'
        f'<span style="color:#f59e0b;margin-top:1px">⚠</span>'
        f'<span style="color:var(--text-sec);font-size:0.87rem">{f}</span></div>'
        for f in flags
    )
    st.markdown(
        f'<div class="info-box">'
        f'<div style="color:var(--text-p);font-size:0.82rem;font-weight:600;margin-bottom:10px">风险提示</div>'
        f'{flag_html}</div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="ai-box">', unsafe_allow_html=True)
    ai_col, note_col = st.columns([2, 4])
    user_profile = st.session_state.get("user_profile")
    with ai_col:
        run_ai = st.button(
            "🤖 一键 AI 深度分析", type="primary", key=f"ai_{code}", use_container_width=True
        )
    with note_col:
        if ai_available():
            st.caption(
                f"使用 {st.session_state.get('ai_provider', 'AI')} · "
                f"{st.session_state.get('ai_model', '')} 流式生成"
            )
        else:
            st.caption("请先在左侧侧边栏填写 AI 服务商信息")

    ai_key = f"ai_result_{code}"
    if run_ai:
        st.session_state.pop(ai_key, None)
        result = st.write_stream(stream_ai_analysis(row, user_profile, domain))
        st.session_state[ai_key] = result
    elif ai_key in st.session_state:
        st.markdown(st.session_state[ai_key])
    else:
        questions = self_check_questions(row, domain)
        q_html = "".join(
            f'<div style="color:var(--text-m);font-size:0.85rem;margin-bottom:6px">· {q}</div>'
            for q in questions
        )
        st.markdown(
            f'<div style="color:var(--text-m);font-size:0.8rem;margin-bottom:10px">'
            f'点击按钮后 AI 将流式生成分析报告。以下是预设自查问题：</div>'
            f'{q_html}',
            unsafe_allow_html=True,
        )

    st.markdown("</div>", unsafe_allow_html=True)


# ── Tab: 数据与术语 ───────────────────────────────────────────────────────────
def render_universe_and_terms(df: pd.DataFrame) -> None:
    section_title("样本池 & 术语", "10 只上证 ETF 样本 · 数据来自 Excel + Yahoo Finance 缓存")

    display = df[[
        "code", "name", "category", "style", "quality_score",
        "total_fee", "age_years", "scale_proxy", "price_points",
    ]].copy()
    display["scale_proxy"] = display["scale_proxy"].apply(format_money)
    st.dataframe(display, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    section_title("术语说明")
    cols = st.columns(2)
    for idx, (term, desc) in enumerate(HELP_TEXT.items()):
        with cols[idx % 2]:
            st.markdown(
                f'<div class="term-card">'
                f'<div style="color:var(--accent);font-size:0.82rem;font-weight:600;margin-bottom:4px">'
                f'{term}</div>'
                f'<div style="color:var(--text-m);font-size:0.83rem;line-height:1.5">{desc}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )


# ── Tab: 调试 ─────────────────────────────────────────────────────────────────
def render_debug(profile: pd.DataFrame, prices: pd.DataFrame, df: pd.DataFrame) -> None:
    section_title("调试信息", "开发用，后续可隐藏")

    def fmt_mtime(path: Path) -> str:
        return (
            pd.Timestamp(path.stat().st_mtime, unit="s", tz="UTC")
            .tz_convert(CST)
            .strftime("%Y-%m-%d %H:%M:%S（北京时间）")
        )

    c1, c2 = st.columns(2)
    with c1:
        if PROFILE_CACHE.exists():
            st.success(f"profile 缓存存在 · 修改于 {fmt_mtime(PROFILE_CACHE)}")
        else:
            st.error("profile 缓存不存在")
    with c2:
        if PRICE_CACHE.exists():
            st.success(f"prices 缓存存在 · 修改于 {fmt_mtime(PRICE_CACHE)}")
        else:
            st.error("prices 缓存不存在")

    if not prices.empty and "fetch_time" in prices.columns and prices["fetch_time"].notna().any():
        st.info(f"行情最近拉取时间：{prices['fetch_time'].dropna().iloc[0]}")
    else:
        st.info("行情来自本地缓存（未记录拉取时间）")

    st.markdown("#### 各 ETF 数据点统计")
    if not prices.empty:
        st.dataframe(
            prices.groupby("code").agg(
                数据点数=("close", "count"),
                最早日期=("date", "min"),
                最新日期=("date", "max"),
                最新收盘价=("close", "last"),
            ).reset_index(),
            use_container_width=True,
        )

    with st.expander("ETF 基本信息（完整）"):
        st.dataframe(profile, use_container_width=True)
    with st.expander("行情缓存（最近 20 条）"):
        if not prices.empty:
            st.dataframe(
                prices.sort_values("date", ascending=False).head(20), use_container_width=True
            )
    with st.expander("综合打分表（完整）"):
        st.dataframe(df, use_container_width=True)


# ── Main ──────────────────────────────────────────────────────────────────────
def main() -> None:
    st.set_page_config(page_title="ETF 投资助手", layout="wide", page_icon="📊")

    # sidebar_controls reads/writes dark_mode in session_state via st.toggle
    refresh_prices, refresh_profile = sidebar_controls()

    # Inject CSS after sidebar so dark_mode state is settled
    st.markdown(BASE_CSS, unsafe_allow_html=True)
    if not st.session_state.get("dark_mode", True):
        st.markdown(LIGHT_CSS, unsafe_allow_html=True)

    profile = load_or_create_profile(refresh=refresh_profile)

    if refresh_prices:
        prog = st.empty()
        with st.spinner("正在从 Yahoo Finance 拉取行情…"):
            prices, warn = load_prices(profile["code"], refresh=True, progress_placeholder=prog)
        prog.empty()
        if warn:
            st.error(f"部分拉取失败：{warn}")
        else:
            st.success("行情更新成功")
    else:
        prices, warn = load_prices(profile["code"], refresh=False)
        if warn:
            st.warning(warn)

    df = add_metrics(profile, prices)

    default_user = UserProfile(
        "我是新手", "低风险", "3年以上", "长期配置，慢慢积累", "我想先要一个稳一点的底座", 50000.0
    )
    user = st.session_state.get("user_profile", default_user)
    domains = recommend_domains(user)

    st.markdown(
        '<div style="padding:8px 0 28px">'
        '<div style="font-size:1.9rem;font-weight:800;color:var(--text-h);letter-spacing:-0.03em">'
        'ETF 投资助手</div>'
        '<div style="color:var(--text-d);font-size:0.88rem;margin-top:6px">'
        '基于规则评分 + AI 解释 · 仅用于学习和决策辅助，不构成投资建议</div></div>',
        unsafe_allow_html=True,
    )

    tab_profile, tab_domain, tab_candidates, tab_detail, tab_data, tab_debug = st.tabs(
        ["偏好问卷", "方向推荐", "ETF 对比", "质量评估", "数据 & 术语", "🔧 调试"]
    )

    with tab_profile:
        user = render_profile_form()
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
