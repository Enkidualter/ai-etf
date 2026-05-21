"""ETF investment assistant demo v2.

Run with:
    streamlit run app.py --browser.gatherUsageStats false
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import requests
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "上证ETF(1).xlsx"
CACHE_DIR = BASE_DIR / "data" / "cache"
PROFILE_CACHE = CACHE_DIR / "etf_sample_profile.csv"
PRICE_CACHE = CACHE_DIR / "etf_sample_prices.csv"
SAMPLE_SIZE = 10

DEEPSEEK_BASE_URL = os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com").rstrip("/")
DEEPSEEK_MODEL = os.environ.get("DEEPSEEK_MODEL", "deepseek-v4-flash")


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
        name="宽基核心",
        match_score=0,
        role="核心仓位",
        plain_intro="先用一篮子代表性股票打底，不把胜负押在单一行业上。",
        suitable_for="适合新手、长期配置、希望先把组合底座搭稳的投资者。",
        return_source="主要来自市场整体增长、指数成分股盈利改善和估值修复。",
        key_risks="市场整体下跌时也会回撤，不能理解成保本工具。",
        common_mistake="只看短期涨幅排名，忽略宽基更适合长期、分批和纪律化持有。",
        category_keywords=("宽基", "核心"),
        risk_hint="适合作为组合底座，但仍要控制买入节奏和总仓位。",
    ),
    "红利/低波": DomainRecommendation(
        name="红利/低波",
        match_score=0,
        role="稳健核心或防守仓位",
        plain_intro="更关注分红、现金流和相对稳健的公司，目标是少一点刺激，多一点纪律。",
        suitable_for="适合不喜欢大起大落、希望降低组合波动的新手或保守型投资者。",
        return_source="主要来自股息、企业稳定盈利和低估值修复。",
        key_risks="红利不等于无风险，可能有行业集中、周期下行和股息下降风险。",
        common_mistake="把高股息率简单理解成高收益，忽略股价下跌可能抵消分红。",
        category_keywords=("红利", "低波", "价值", "金融"),
        risk_hint="当前样本池红利ETF不足，demo会用价值/金融类作为近似候选展示。",
    ),
    "行业主题": DomainRecommendation(
        name="行业主题",
        match_score=0,
        role="卫星仓位",
        plain_intro="集中投向某个行业或主题，适合表达观点，但不适合新手一上来重仓。",
        suitable_for="适合能承受较大波动、理解行业周期、愿意把它控制在小比例仓位的投资者。",
        return_source="主要来自行业景气度上行、政策催化、估值扩张或盈利改善。",
        key_risks="行业热度高时容易追涨，景气反转时回撤也可能很深。",
        common_mistake="因为最近涨得多就买入，却没有想清楚退出规则和仓位上限。",
        category_keywords=("行业", "主题", "成长"),
        risk_hint="更适合作为小比例卫星仓位，不建议作为保守投资者的核心仓位。",
    ),
}


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {
        "证券代码": "code",
        "证券名称": "name",
        "基金类型": "fund_type",
        "单位净值\n[交易日期] 最新[单位]元": "nav",
        "基金份额\n[交易日期] 最新\n[单位]  份": "shares",
        "投资风格\n[年度] 2023\n[报告期] 中报": "style",
        "基金管理人": "manager",
        "基金托管人": "custodian",
        "管理费率[单位]%": "management_fee",
        "托管费率[单位]%": "custody_fee",
        "基金经理(现任)": "fund_manager",
        "基金存续期[单位]年": "duration_years",
        "基金成立日": "inception_date",
        "基金到期日": "maturity_date",
    }
    return df.rename(columns=mapping)


def parse_inception(value: Any) -> Optional[pd.Timestamp]:
    if pd.isna(value):
        return None
    text = str(value).split(".")[0].strip()
    if len(text) != 8:
        return None
    parsed = pd.to_datetime(text, format="%Y%m%d", errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed


def classify_etf(row: pd.Series) -> str:
    text = f"{row.get('name', '')} {row.get('style', '')}"
    if any(key in text for key in ["沪深300", "上证50", "中证500", "国企"]):
        return "宽基/核心"
    if any(key in text for key in ["消费", "金融"]):
        return "行业/主题"
    if "成长" in text:
        return "成长风格"
    if "价值" in text:
        return "价值风格"
    return "其他"


def load_profile_from_excel() -> pd.DataFrame:
    df = pd.read_excel(EXCEL_PATH).head(SAMPLE_SIZE)
    df = normalize_columns(df)
    keep = [
        "code",
        "name",
        "fund_type",
        "nav",
        "shares",
        "style",
        "manager",
        "custodian",
        "management_fee",
        "custody_fee",
        "fund_manager",
        "inception_date",
    ]
    df = df[keep].copy()
    df["inception"] = df["inception_date"].apply(parse_inception)
    df["age_years"] = df["inception"].apply(
        lambda x: np.nan if x is None else max((pd.Timestamp.today() - x).days / 365.25, 0)
    )
    df["total_fee"] = pd.to_numeric(df["management_fee"], errors="coerce").fillna(0) + pd.to_numeric(
        df["custody_fee"], errors="coerce"
    ).fillna(0)
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
    codes: Iterable[str],
    progress_placeholder: Any = None,
) -> Tuple[pd.DataFrame, Optional[str]]:
    try:
        import yfinance as yf
    except ImportError:
        return pd.DataFrame(), "未安装 yfinance，请执行 pip install yfinance。"

    code_list = list(codes)
    frames: List[pd.DataFrame] = []
    errors: List[str] = []
    successes: List[str] = []

    for i, code in enumerate(code_list):
        if progress_placeholder is not None:
            progress_placeholder.info(f"正在拉取 {i+1}/{len(code_list)}：{code} …")
        try:
            clean = str(code).split(".")[0]
            suffix = str(code).split(".")[-1].upper() if "." in str(code) else "SH"
            yf_code = clean + (".SS" if suffix == "SH" else ".SZ")
            hist = yf.Ticker(yf_code).history(period="1y")
            if hist.empty:
                errors.append(f"{code} → {yf_code}：无数据")
                continue
            frame = pd.DataFrame({
                "code": code,
                "date": pd.to_datetime(hist.index.tz_localize(None)),
                "close": hist["Close"].values,
            }).dropna(subset=["date", "close"]).sort_values("date")
            frames.append(frame)
            successes.append(f"{code}（{len(frame)}条）")
        except Exception as e:
            errors.append(f"{code}：{type(e).__name__}: {e}")
            continue

    if not frames:
        detail = "\n".join(errors)
        return pd.DataFrame(), f"所有代码拉取失败：\n{detail}"

    prices = pd.concat(frames, ignore_index=True)
    prices["fetch_time"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    prices.to_csv(PRICE_CACHE, index=False, encoding="utf-8-sig")
    summary = f"成功：{', '.join(successes)}"
    if errors:
        summary += f"\n失败：{', '.join(errors)}"
    return prices, None if not errors else summary


def load_prices(
    codes: Iterable[str],
    refresh: bool = False,
    progress_placeholder: Any = None,
) -> Tuple[pd.DataFrame, Optional[str]]:
    if PRICE_CACHE.exists() and not refresh:
        prices = pd.read_csv(PRICE_CACHE)
        prices["date"] = pd.to_datetime(prices["date"], errors="coerce")
        return prices, None
    return fetch_prices_from_yfinance(codes, progress_placeholder=progress_placeholder)


def max_drawdown(close: pd.Series) -> float:
    close = close.dropna()
    if close.empty:
        return np.nan
    drawdown = close / close.cummax() - 1
    return float(drawdown.min())


def annual_volatility(close: pd.Series) -> float:
    returns = close.pct_change().dropna()
    if returns.empty:
        return np.nan
    return float(returns.std() * np.sqrt(252))


def one_year_return(close: pd.Series) -> float:
    close = close.dropna()
    if len(close) < 2:
        return np.nan
    return float(close.iloc[-1] / close.iloc[0] - 1)


def add_metrics(profile: pd.DataFrame, prices: pd.DataFrame) -> pd.DataFrame:
    df = profile.copy()
    metric_rows = []
    for code in df["code"]:
        close = prices.loc[prices["code"] == code, "close"] if not prices.empty else pd.Series(dtype=float)
        metric_rows.append(
            {
                "code": code,
                "max_drawdown": max_drawdown(close),
                "volatility": annual_volatility(close),
                "one_year_return": one_year_return(close),
                "price_points": int(close.dropna().shape[0]),
            }
        )
    df = df.merge(pd.DataFrame(metric_rows), on="code", how="left")
    scores = pd.DataFrame([score_etf(row) for _, row in df.iterrows()])
    return pd.concat([df.reset_index(drop=True), scores], axis=1)


def score_size(scale: float) -> float:
    if pd.isna(scale):
        return 8
    if scale >= 50_000_000_000:
        return 25
    if scale >= 10_000_000_000:
        return 22
    if scale >= 1_000_000_000:
        return 16
    if scale >= 100_000_000:
        return 9
    return 4


def score_fee(total_fee: float) -> float:
    if pd.isna(total_fee):
        return 8
    if total_fee <= 0.20:
        return 20
    if total_fee <= 0.30:
        return 17
    if total_fee <= 0.60:
        return 12
    if total_fee <= 1.00:
        return 7
    return 3


def score_age(age: float) -> float:
    if pd.isna(age):
        return 6
    if age >= 5:
        return 20
    if age >= 3:
        return 16
    if age >= 1:
        return 10
    if age >= 0.5:
        return 6
    return 3


def score_risk(drawdown: float, volatility: float) -> float:
    if pd.isna(drawdown) or pd.isna(volatility):
        return 8
    score = 20
    if drawdown < -0.35:
        score -= 8
    elif drawdown < -0.25:
        score -= 5
    elif drawdown < -0.15:
        score -= 2
    if volatility > 0.35:
        score -= 7
    elif volatility > 0.25:
        score -= 4
    elif volatility > 0.18:
        score -= 2
    return max(score, 3)


def score_data(points: int, static_missing: int) -> float:
    score = 15
    if points < 120:
        score -= 6
    elif points < 220:
        score -= 3
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
        "size_score": size,
        "fee_score": fee,
        "age_score": age,
        "risk_score": risk,
        "data_score": data,
    }


def infer_risk_level(risk_answer: str, experience: str, horizon: str) -> str:
    if "不确定" in risk_answer:
        if experience == "我是新手" or horizon == "1年以内":
            return "低风险"
        return "中风险"
    if "不太能接受" in risk_answer:
        return "低风险"
    if "阶段性波动" in risk_answer:
        return "中风险"
    return "高风险"


def recommend_domains(user: UserProfile) -> List[DomainRecommendation]:
    scores = {"宽基核心": 50.0, "红利/低波": 45.0, "行业主题": 30.0}

    if user.risk_level == "低风险":
        scores["红利/低波"] += 24
        scores["宽基核心"] += 18
        scores["行业主题"] -= 20
    elif user.risk_level == "中风险":
        scores["宽基核心"] += 20
        scores["红利/低波"] += 10
        scores["行业主题"] += 4
    else:
        scores["行业主题"] += 24
        scores["宽基核心"] += 10

    if user.experience == "我是新手":
        scores["宽基核心"] += 14
        scores["红利/低波"] += 8
        scores["行业主题"] -= 8
    if user.horizon == "1年以内":
        scores["红利/低波"] += 12
        scores["行业主题"] -= 12
    elif user.horizon == "3年以上":
        scores["宽基核心"] += 10
        scores["行业主题"] += 3

    if user.goal == "希望稳一点，少一些大起大落":
        scores["红利/低波"] += 20
        scores["行业主题"] -= 10
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

    result = []
    for name, base in DOMAIN_LIBRARY.items():
        item = DomainRecommendation(**{**base.__dict__, "match_score": round(scores[name], 1)})
        result.append(item)
    return sorted(result, key=lambda item: item.match_score, reverse=True)


def etf_matches_domain(row: pd.Series, domain: DomainRecommendation) -> bool:
    text = f"{row.get('name', '')} {row.get('style', '')} {row.get('category', '')}"
    return any(keyword in text for keyword in domain.category_keywords)


def candidate_score(row: pd.Series, user: UserProfile, domain: DomainRecommendation) -> float:
    score = float(row["quality_score"])
    if etf_matches_domain(row, domain):
        score += 18
    if user.risk_level == "低风险":
        if row.get("total_fee", 99) <= 0.20:
            score += 8
        if pd.notna(row.get("max_drawdown")) and row.get("max_drawdown") < -0.25:
            score -= 12
        if pd.notna(row.get("volatility")) and row.get("volatility") > 0.25:
            score -= 10
        if "行业" in str(row.get("category", "")):
            score -= 6
    elif user.risk_level == "高风险" and "行业" in str(row.get("category", "")):
        score += 5
    if user.horizon == "1年以内" and pd.notna(row.get("volatility")) and row.get("volatility") > 0.25:
        score -= 8
    return round(score, 1)


def rank_candidates(df: pd.DataFrame, user: UserProfile, domain: DomainRecommendation) -> pd.DataFrame:
    ranked = df.copy()
    ranked["candidate_score"] = ranked.apply(lambda row: candidate_score(row, user, domain), axis=1)
    ranked["domain_match"] = ranked.apply(lambda row: etf_matches_domain(row, domain), axis=1)
    ranked = ranked.sort_values(["domain_match", "candidate_score"], ascending=[False, False])
    return ranked.head(3)


def risk_flags(row: pd.Series) -> List[str]:
    flags = []
    if row.get("age_years", 99) < 1:
        flags.append("成立时间较短，成立以来收益率参考价值有限。")
    if row.get("total_fee", 0) > 0.60:
        flags.append("费率相对偏高，长期持有时成本会持续累积。")
    if row.get("scale_proxy", np.inf) < 1_000_000_000:
        flags.append("规模代理偏小，需关注流动性和清盘风险。")
    if pd.notna(row.get("max_drawdown")) and row.get("max_drawdown") < -0.25:
        flags.append("近一年最大回撤较大，需要先确认能否承受波动。")
    if pd.notna(row.get("volatility")) and row.get("volatility") > 0.25:
        flags.append("近一年波动率较高，不适合作为保守型核心仓位。")
    if not flags:
        flags.append("未发现特别突出的单项风险，但仍需结合自身期限和仓位控制。")
    return flags


def self_check_questions(row: pd.Series, domain: Optional[DomainRecommendation] = None) -> List[str]:
    category = domain.name if domain else row.get("category", "这类 ETF")
    return [
        f"我是否理解这只ETF为什么被归到“{category}”，主要风险来源是什么？",
        "如果短期回撤20%-30%，我是否还能按原计划持有？",
        "我买入它是长期配置需要，还是被近期涨幅或热度吸引？",
        "我是否比较过同类ETF的规模、费率、成立时间和流动性？",
        "这只ETF在我的总资产中应该是核心仓位还是小比例卫星仓位？",
    ]


def explain_etf(row: pd.Series, domain: Optional[DomainRecommendation] = None) -> Tuple[str, str, List[str]]:
    seen = (
        f"{row['name']} 当前质量分 {row['quality_score']:.1f}/100，"
        f"费率合计约 {row.get('total_fee', np.nan):.2f}%，"
        f"成立约 {row.get('age_years', np.nan):.1f} 年，"
        f"规模代理约 {format_money(row.get('scale_proxy'))}。"
    )
    ignored = "；".join(risk_flags(row))
    return seen, ignored, self_check_questions(row, domain)


def build_fallback_narrative(
    user: UserProfile,
    domain: DomainRecommendation,
    candidates: pd.DataFrame,
) -> str:
    names = "、".join(candidates["name"].head(3).tolist()) if not candidates.empty else "暂无合适候选"
    return (
        f"根据你的风险承受能力和持有期限，当前更适合先从“{domain.name}”看起。"
        f"这个方向的定位是{domain.role}：{domain.plain_intro}"
        f"在当前样本池里，可以先重点比较：{names}。"
        f"需要注意，推荐只是帮助你缩小观察范围，不是买入建议；真正决策前还要看回撤、费率、规模和自己能否长期持有。"
    )


def deepseek_available() -> bool:
    return bool(os.environ.get("DEEPSEEK_API_KEY"))


def call_deepseek(payload: Dict[str, Any]) -> Tuple[Optional[str], Optional[str]]:
    api_key = os.environ.get("DEEPSEEK_API_KEY")
    if not api_key:
        return None, "未设置 DEEPSEEK_API_KEY，使用规则模板文案。"

    system_prompt = (
        "你是一个谨慎的ETF投资教育助手。你只能基于用户提供的结构化事实写解释，"
        "不得编造数据，不得预测收益，不得给出买入、卖出、保证收益等投资建议。"
        "请用中文输出，面向新手，但保留必要专业词并解释。"
    )
    user_prompt = (
        "请把以下事实润色成一段适合展示在ETF投资助手页面里的解释。"
        "结构：1）为什么先看这个方向；2）候选ETF怎么比较；3）风险和自查提醒。"
        "控制在180字以内。\n\n"
        + json.dumps(to_jsonable(payload), ensure_ascii=False)
    )
    try:
        response = requests.post(
            f"{DEEPSEEK_BASE_URL}/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={
                "model": DEEPSEEK_MODEL,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                "temperature": 0.3,
                "max_tokens": 500,
            },
            timeout=20,
        )
        response.raise_for_status()
        data = response.json()
        text = data["choices"][0]["message"]["content"].strip()
        return text, None
    except Exception as exc:
        return None, f"DeepSeek 调用失败，已回退模板：{exc}"


def to_jsonable(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(key): to_jsonable(item) for key, item in value.items()}
    if isinstance(value, list):
        return [to_jsonable(item) for item in value]
    if isinstance(value, tuple):
        return [to_jsonable(item) for item in value]
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating,)):
        if pd.isna(value):
            return None
        return float(value)
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if pd.isna(value):
        return None
    return value


def make_narrative(user: UserProfile, domain: DomainRecommendation, candidates: pd.DataFrame) -> Tuple[str, Optional[str]]:
    candidate_facts = []
    for _, row in candidates.head(3).iterrows():
        candidate_facts.append(
            {
                "code": row["code"],
                "name": row["name"],
                "quality_score": row["quality_score"],
                "fee": row.get("total_fee"),
                "age_years": row.get("age_years"),
                "max_drawdown": row.get("max_drawdown"),
                "volatility": row.get("volatility"),
                "risk_flags": risk_flags(row),
            }
        )
    payload = {
        "user_profile": user.__dict__,
        "domain": {
            "name": domain.name,
            "role": domain.role,
            "intro": domain.plain_intro,
            "suitable_for": domain.suitable_for,
            "return_source": domain.return_source,
            "key_risks": domain.key_risks,
            "common_mistake": domain.common_mistake,
        },
        "candidates": candidate_facts,
        "guardrail": "仅用于学习和决策辅助，不构成投资建议。",
    }
    text, warning = call_deepseek(payload)
    if text:
        return text, None
    return build_fallback_narrative(user, domain, candidates), warning


def format_percent(value: Any) -> str:
    if pd.isna(value):
        return "暂无"
    return f"{value * 100:.2f}%"


def format_money(value: Any) -> str:
    if pd.isna(value):
        return "暂无"
    value = float(value)
    if value >= 100_000_000:
        return f"{value / 100_000_000:.1f} 亿元"
    if value >= 10_000:
        return f"{value / 10_000:.1f} 万元"
    return f"{value:.0f} 元"


def show_header() -> None:
    st.set_page_config(page_title="ETF 投资助手 Demo", layout="wide")
    st.title("ETF 投资助手 Demo v2")
    st.caption("先理解适合自己的ETF方向，再比较具体产品。仅用于学习和决策辅助，不构成投资建议。")


def sidebar_controls() -> Tuple[bool, bool]:
    with st.sidebar:
        st.header("数据与模型")

        if deepseek_available():
            st.success(f"已启用 AI 解释：{DEEPSEEK_MODEL}")
        else:
            st.info("未配置 DeepSeek，解释文案使用规则模板。")

        st.divider()
        refresh_prices = st.button("🔄 拉取最新行情（Yahoo Finance）")
        refresh_profile = st.button("重建ETF静态缓存")
        st.caption("行情数据来自 Yahoo Finance，自动缓存到 data/cache。")
        return refresh_prices, refresh_profile


def render_profile_form() -> UserProfile:
    st.subheader("投资者偏好问卷")
    st.write("尽量用自己的真实感受来选，不需要先懂专业术语。")

    left, right = st.columns(2)
    with left:
        experience = st.selectbox(
            "你的投资经验",
            ["我是新手", "有一些经验", "比较熟悉ETF"],
            help="这个选项会影响系统解释的详细程度，也会影响是否优先推荐更稳健的方向。",
        )
        risk_answer = st.selectbox(
            "如果账户短期亏损，你的感受更接近哪一种？",
            ["我不太能接受亏损", "能接受阶段性波动", "愿意承受较大波动换取机会", "我不确定"],
            help=f"{HELP_TEXT['回撤']} 系统会把“我不确定”默认按更保守的方式处理。",
        )
        horizon = st.selectbox(
            "这笔钱预计多久不用？",
            ["1年以内", "1-3年", "3年以上"],
            help="投资期限越短，越不适合承受高波动ETF。短钱不适合拿来押行业主题。",
        )
    with right:
        goal = st.selectbox(
            "你更接近哪种目标？",
            ["希望稳一点，少一些大起大落", "长期配置，慢慢积累", "想抓住某些行业机会", "还没想清楚"],
            help="目标不同，适合的ETF方向不同。稳健、长期、主题机会不是同一种问题。",
        )
        preference = st.selectbox(
            "你现在更想先了解什么？",
            ["我想先要一个稳一点的底座", "我更关注分红和稳健", "我想了解行业主题机会", "我不确定"],
            help=f"底座通常接近{HELP_TEXT['宽基']} 分红和稳健通常接近{HELP_TEXT['红利']}",
        )
        amount = st.number_input(
            "计划投入金额（元）",
            min_value=1000.0,
            value=50000.0,
            step=1000.0,
            help="demo暂不直接给仓位建议，只用金额判断解释口径。实际投资还要看总资产和现金流。",
        )

    risk_level = infer_risk_level(risk_answer, experience, horizon)
    user = UserProfile(
        experience=experience,
        risk_level=risk_level,
        horizon=horizon,
        goal=goal,
        preference=preference,
        amount=amount,
    )
    st.session_state["user_profile"] = user
    st.info(f"系统识别的风险层级：{risk_level}。这只用于推荐排序和风险提示，不代表正式风险测评。")
    return user


def render_domain_recommendations(user: UserProfile) -> List[DomainRecommendation]:
    domains = recommend_domains(user)
    st.subheader("适合你的ETF方向")
    st.write("第一步先看方向，而不是直接跳到某一只ETF。方向对了，后面的产品比较才有意义。")

    cols = st.columns(3)
    for col, domain in zip(cols, domains):
        with col:
            st.metric(domain.name, f"{domain.match_score:.0f}分", help=domain.risk_hint)
            st.write(f"定位：{domain.role}")
            st.write(domain.plain_intro)
            st.caption(domain.risk_hint)

    top_domain = domains[0]
    with st.expander(f"为什么优先看：{top_domain.name}", expanded=True):
        c1, c2 = st.columns(2)
        c1.write(f"**适合人群：** {top_domain.suitable_for}")
        c1.write(f"**收益来源：** {top_domain.return_source}")
        c2.write(f"**主要风险：** {top_domain.key_risks}")
        c2.write(f"**常见误区：** {top_domain.common_mistake}")
    return domains


def render_candidates(df: pd.DataFrame, user: UserProfile, domains: List[DomainRecommendation]) -> None:
    st.subheader("该方向下的具体ETF候选")
    domain_names = [domain.name for domain in domains]
    selected_name = st.selectbox("选择一个方向查看ETF候选", domain_names, help="先选方向，再在同方向内比较规模、费率、回撤和流动性。")
    domain = next(item for item in domains if item.name == selected_name)
    candidates = rank_candidates(df, user, domain)
    narrative, warning = make_narrative(user, domain, candidates)
    if warning:
        st.caption(warning)

    st.markdown("#### 方向解释")
    st.write(narrative)

    cols = st.columns(3)
    for col, (_, row) in zip(cols, candidates.iterrows()):
        with col:
            st.metric(row["name"], f"{row['quality_score']:.1f}/100", help=row["code"])
            st.write(f"候选分：{row['candidate_score']:.1f}")
            st.write(f"类型：{row['category']}")
            st.write(f"费率：{row.get('total_fee', np.nan):.2f}%")
            st.write(f"规模代理：{format_money(row.get('scale_proxy'))}")
            st.write(f"最大回撤：{format_percent(row.get('max_drawdown'))}")
            st.caption(risk_flags(row)[0])

    st.dataframe(
        candidates[
            [
                "code",
                "name",
                "category",
                "quality_score",
                "candidate_score",
                "domain_match",
                "total_fee",
                "age_years",
                "max_drawdown",
                "volatility",
            ]
        ],
        use_container_width=True,
    )


def render_quality_detail(df: pd.DataFrame, prices: pd.DataFrame, domains: List[DomainRecommendation]) -> None:
    st.subheader("单只ETF质量评估")
    selected = st.selectbox("选择ETF", df["code"] + " - " + df["name"], help="质量分不是买卖建议，只是帮你系统检查规模、费率、成立时间、风险和数据完整性。")
    code = selected.split(" - ")[0]
    row = df.loc[df["code"] == code].iloc[0]
    domain = domains[0] if domains else None
    seen, ignored, questions = explain_etf(row, domain)

    top = st.columns([1, 1, 1, 1])
    top[0].metric("质量总分", f"{row['quality_score']:.1f}/100")
    top[1].metric("最大回撤", format_percent(row.get("max_drawdown")), help=HELP_TEXT["回撤"])
    top[2].metric("年化波动率", format_percent(row.get("volatility")), help=HELP_TEXT["波动率"])
    top[3].metric("近一年收益", format_percent(row.get("one_year_return")), help="过去一年表现不代表未来收益，也可能受起止日期影响。")

    score_df = pd.DataFrame(
        {
            "维度": ["规模", "费率", "成立时间", "历史风险", "数据完整性"],
            "得分": [row["size_score"], row["fee_score"], row["age_score"], row["risk_score"], row["data_score"]],
            "满分": [25, 20, 20, 20, 15],
        }
    )
    st.plotly_chart(px.bar(score_df, x="维度", y="得分", text="得分", range_y=[0, 25]), use_container_width=True)

    price_df = prices[prices["code"] == code].copy() if not prices.empty else pd.DataFrame()
    if not price_df.empty:
        st.plotly_chart(px.line(price_df, x="date", y="close", title="近一年收盘价"), use_container_width=True)
    else:
        st.info("暂无历史行情缓存。设置 iFinD 账号后点击侧边栏“重新拉取 iFinD 行情”。")

    st.markdown("### 三段式解释")
    st.write(f"**你看到的指标：** {seen}")
    st.write(f"**容易忽略的问题：** {ignored}")
    st.write("**投资前自查问题：**")
    for question in questions:
        st.write(f"- {question}")


def render_terms() -> None:
    st.subheader("术语小抄")
    cols = st.columns(2)
    for idx, (term, desc) in enumerate(HELP_TEXT.items()):
        with cols[idx % 2]:
            st.write(f"**{term}**：{desc}")


def render_universe(df: pd.DataFrame) -> None:
    st.subheader("10只ETF样本池")
    display = df[
        [
            "code",
            "name",
            "category",
            "style",
            "quality_score",
            "total_fee",
            "age_years",
            "scale_proxy",
            "price_points",
        ]
    ].copy()
    display["scale_proxy"] = display["scale_proxy"].apply(format_money)
    st.dataframe(display, use_container_width=True)


def render_debug(profile: pd.DataFrame, prices: pd.DataFrame, df: pd.DataFrame) -> None:
    st.subheader("调试信息")

    # 缓存文件状态
    st.markdown("#### 缓存文件状态")
    c1, c2 = st.columns(2)
    with c1:
        if PROFILE_CACHE.exists():
            mtime = pd.Timestamp(PROFILE_CACHE.stat().st_mtime, unit="s").strftime("%Y-%m-%d %H:%M:%S")
            st.success(f"etf_sample_profile.csv 存在\n最后修改：{mtime}")
        else:
            st.error("etf_sample_profile.csv 不存在")
    with c2:
        if PRICE_CACHE.exists():
            mtime = pd.Timestamp(PRICE_CACHE.stat().st_mtime, unit="s").strftime("%Y-%m-%d %H:%M:%S")
            st.success(f"etf_sample_prices.csv 存在\n最后修改：{mtime}")
        else:
            st.error("etf_sample_prices.csv 不存在")

    # 行情缓存中的更新时间（如果有）
    if not prices.empty and "fetch_time" in prices.columns:
        fetch_time = prices["fetch_time"].dropna().iloc[0] if prices["fetch_time"].notna().any() else "未知"
        st.info(f"行情最近一次从 Yahoo Finance 拉取时间：{fetch_time}")
    else:
        st.info("行情来自缓存文件（未记录拉取时间，或尚未从 Yahoo Finance 拉取过）")

    # 每只ETF行情点数
    st.markdown("#### 各ETF行情数据点数")
    if not prices.empty:
        summary = prices.groupby("code").agg(
            数据点数=("close", "count"),
            最早日期=("date", "min"),
            最新日期=("date", "max"),
            最新收盘价=("close", "last"),
        ).reset_index()
        st.dataframe(summary, use_container_width=True)
    else:
        st.warning("暂无行情数据")

    # 完整profile表
    st.markdown("#### ETF基本信息（完整）")
    st.dataframe(profile, use_container_width=True)

    # 完整prices表（最近20条）
    st.markdown("#### 行情缓存（最近20条）")
    if not prices.empty:
        st.dataframe(prices.sort_values("date", ascending=False).head(20), use_container_width=True)
    else:
        st.warning("暂无行情数据")

    # 完整打分表
    st.markdown("#### 综合打分表（完整）")
    st.dataframe(df, use_container_width=True)


def main() -> None:
    show_header()
    refresh_prices, refresh_profile = sidebar_controls()

    profile = load_or_create_profile(refresh=refresh_profile)

    if refresh_prices:
        progress_placeholder = st.empty()
        with st.spinner("正在从 Yahoo Finance 拉取行情，请稍候…"):
            prices, price_warning = load_prices(profile["code"], refresh=True, progress_placeholder=progress_placeholder)
        progress_placeholder.empty()
        if price_warning:
            st.error(f"拉取结果：\n\n{price_warning}")
        else:
            st.success("行情拉取成功，已更新缓存！")
    else:
        prices, price_warning = load_prices(profile["code"], refresh=False)
        if price_warning:
            st.warning(price_warning)

    df = add_metrics(profile, prices)

    default_user = UserProfile("我是新手", "低风险", "3年以上", "长期配置，慢慢积累", "我想先要一个稳一点的底座", 50000.0)
    user = st.session_state.get("user_profile", default_user)
    domains = recommend_domains(user)

    tab_profile, tab_domain, tab_candidates, tab_detail, tab_data, tab_debug = st.tabs(
        ["偏好问卷", "领域推荐", "ETF候选", "质量评估", "数据与术语", "🔧 调试"]
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
        render_universe(df)
        render_terms()
        st.caption("数据来源：本地Excel样本 + Yahoo Finance行情。仅用于demo，不构成投资建议。")

    with tab_debug:
        render_debug(profile, prices, df)


if __name__ == "__main__":
    main()
