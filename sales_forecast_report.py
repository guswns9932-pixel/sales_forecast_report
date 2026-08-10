# -*- coding: utf-8 -*-
"""
sales_forecast_report.py
=========================
RAWDATA 엑셀(매출일/모델/반입라인/공정/메이커/수량/단가/매출액 포함)을 읽어
분기별 매출을 예측하고 HTML + PDF 리포트를 생성하는 독립 실행 스크립트입니다.

사용법:
    python sales_forecast_report.py RAWDATA.xlsx
    python sales_forecast_report.py RAWDATA.xlsx --forecast-start 2026Q1 --forecast-end 2026Q2
    python sales_forecast_report.py RAWDATA.xlsx --train-end 2024Q4 --forecast-start 2025Q1 --forecast-end 2025Q4

자세한 옵션은:
    python sales_forecast_report.py --help
"""
import argparse
import base64
import io
import re
import sys
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

warnings.filterwarnings("ignore")
import logging
logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)

BRAND_MAGENTA = "#C8007C"
BRAND_GRAY = "#7F7F7F"
BRAND_MAGENTA_LIGHT = "#E9A0CE"

plt.rcParams["font.family"] = ["Malgun Gothic", "AppleGothic", "NanumGothic", "Noto Sans CJK KR",
                                "Noto Sans CJK JP", "Noto Sans KR", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False

# 시스템에 따라 한글 폰트가 기본 등록되어 있지 않은 경우(특히 .ttc 컬렉션 폰트)를 위해
# 흔한 경로를 찾아 matplotlib 폰트 매니저에 직접 등록을 시도한다. 실패해도 무시하고 진행.
def _register_korean_fonts():
    import matplotlib.font_manager as fm
    candidates = [
        "C:/Windows/Fonts/malgun.ttf",
        "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansKR-Regular.otf",
    ]
    for p in candidates:
        if Path(p).exists():
            try:
                fm.fontManager.addfont(p)
            except Exception:
                pass


_register_korean_fonts()


# =====================================================================
# 분기 유틸리티
# =====================================================================
def parse_quarter(label):
    """'2026Q1' -> (2026, 1)"""
    m = re.match(r"^\s*(\d{4})\s*Q\s*([1-4])\s*$", str(label), re.IGNORECASE)
    if not m:
        raise ValueError(f"분기 형식이 올바르지 않습니다: {label!r} (예: 2026Q1)")
    return int(m.group(1)), int(m.group(2))


def quarter_label(y, q):
    return f"{y}Q{q}"


def quarter_index(y, q, base_year=2000):
    return (y - base_year) * 4 + (q - 1)


def index_to_quarter(idx, base_year=2000):
    y = base_year + idx // 4
    q = idx % 4 + 1
    return y, q


def quarter_range(start_label, end_label):
    """start_label ~ end_label 사이 모든 분기 라벨 리스트 (양끝 포함)"""
    sy, sq = parse_quarter(start_label)
    ey, eq = parse_quarter(end_label)
    si, ei = quarter_index(sy, sq), quarter_index(ey, eq)
    if ei < si:
        raise ValueError(f"종료분기({end_label})가 시작분기({start_label})보다 앞섭니다")
    out = []
    for i in range(si, ei + 1):
        y, q = index_to_quarter(i)
        out.append(quarter_label(y, q))
    return out


def add_quarters(label, n):
    y, q = parse_quarter(label)
    idx = quarter_index(y, q) + n
    y2, q2 = index_to_quarter(idx)
    return quarter_label(y2, q2)


# =====================================================================
# 데이터 로딩
# =====================================================================
REQUIRED_COLS = ["매출일", "모델", "반입라인", "공정", "메이커", "수량", "단가(KRW)", "매출액(KRW)"]


def load_rawdata(path):
    """RAWDATA 시트를 읽는다. 표준 양식(1행 헤더, 2~3행 안내, 4행부터 데이터)과
    일반적인 단순 표(1행 헤더, 2행부터 데이터) 둘 다 지원한다."""
    xls = pd.ExcelFile(path)
    sheet = "RAWDATA" if "RAWDATA" in xls.sheet_names else xls.sheet_names[0]

    raw_preview = pd.read_excel(path, sheet_name=sheet, header=0, nrows=3)
    tag_row_present = raw_preview.iloc[0].astype(str).str.contains("필수|권장|선택|자동").any()
    skiprows = [1, 2] if tag_row_present else None

    df = pd.read_excel(path, sheet_name=sheet, header=0, skiprows=skiprows)
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"RAWDATA에 다음 컬럼이 없습니다: {missing}\n실제 컬럼: {list(df.columns)}")

    df = df.dropna(subset=["매출일", "공정", "모델"]).copy()
    df["매출일"] = pd.to_datetime(df["매출일"])
    df["연도"] = df["매출일"].dt.year
    df["분기"] = df["매출일"].dt.quarter
    df["연분기"] = df["연도"].astype(str) + "Q" + df["분기"].astype(str)
    df["매출액(KRW)"] = pd.to_numeric(df["매출액(KRW)"], errors="coerce").fillna(0)
    df["수량"] = pd.to_numeric(df["수량"], errors="coerce").fillna(0)
    return df


# =====================================================================
# 예측 방법 (계절성 ETS / 간헐수요 Croston-SBA / 단순평균)
# =====================================================================
def croston_sba(values, alpha=0.1):
    """Croston's method + SBA 편의보정. 간헐적(0이 많은) 수요 시계열에 적합.
    반환값: 분기당 평균 수요율(모든 미래 분기에 동일하게 적용되는 flat 예측치)"""
    values = np.asarray(values, dtype=float)
    nz = np.nonzero(values)[0]
    if len(nz) == 0:
        return 0.0
    z_est = values[nz[0]]
    q_est = nz[0] + 1
    last = nz[0]
    for t in nz[1:]:
        interval = t - last
        z_est = alpha * values[t] + (1 - alpha) * z_est
        q_est = alpha * interval + (1 - alpha) * q_est
        last = t
    rate = z_est / q_est if q_est > 0 else 0.0
    return max(0.0, rate * (1 - alpha / 2))  # SBA bias correction


def seasonal_ets_forecast(values, n_ahead):
    """statsmodels Holt-Winters(가법 추세+가법 계절성, damped)로 향후 n_ahead 분기 예측.
    실패하면 None 반환(호출부에서 폴백 처리)."""
    from statsmodels.tsa.holtwinters import ExponentialSmoothing
    values = np.asarray(values, dtype=float)
    try:
        model = ExponentialSmoothing(
            values, trend="add", damped_trend=True,
            seasonal="add", seasonal_periods=4,
            initialization_method="estimated",
        )
        fit = model.fit(optimized=True)
        fc = fit.forecast(n_ahead)
        return np.maximum(fc, 0.0)
    except Exception:
        return None


def linear_trend_forecast(values, n_ahead):
    """계절성 ETS가 실패했을 때의 폴백: 단순 선형회귀 추세 연장(0 하한)."""
    values = np.asarray(values, dtype=float)
    x = np.arange(len(values))
    if len(values) < 2 or np.all(values == values[0]):
        base = float(np.mean(values)) if len(values) else 0.0
        return np.full(n_ahead, max(0.0, base))
    slope, intercept = np.polyfit(x, values, 1)
    future_x = np.arange(len(values), len(values) + n_ahead)
    fc = slope * future_x + intercept
    return np.maximum(fc, 0.0)


def forecast_series(values, n_ahead, min_regular_quarters=8, min_activity_ratio=0.5, min_intermittent_obs=3):
    """시계열 특성에 따라 방법을 자동 선택해 n_ahead분기를 예측한다.
    반환: (forecast_array, method_name)"""
    values = np.asarray(values, dtype=float)
    n = len(values)
    nz_count = int(np.count_nonzero(values))
    activity_ratio = nz_count / n if n else 0

    if n >= min_regular_quarters and activity_ratio >= min_activity_ratio:
        fc = seasonal_ets_forecast(values, n_ahead)
        if fc is not None:
            return fc, "계절성 지수평활(Holt-Winters)"
        fc = linear_trend_forecast(values, n_ahead)
        return fc, "선형추세(ETS 적합 실패 → 대체)"

    if nz_count >= min_intermittent_obs:
        rate = croston_sba(values)
        return np.full(n_ahead, rate), "간헐수요모델(Croston-SBA)"

    avg = float(np.mean(values[values > 0])) if nz_count > 0 else 0.0
    return np.full(n_ahead, avg), "표본부족(평균유지)"


# =====================================================================
# 집계 & 예측 실행
# =====================================================================
def build_quarter_axis(df, extra_future_quarters=0):
    quarters = sorted(df["연분기"].unique(), key=lambda s: quarter_index(*parse_quarter(s)))
    if extra_future_quarters:
        last = quarters[-1]
        for i in range(1, extra_future_quarters + 1):
            quarters.append(add_quarters(last, i))
    return quarters


def quarterly_pivot(df, group_cols, value_col, quarters):
    idx_q = pd.CategoricalDtype(quarters, ordered=True)
    tmp = df.copy()
    tmp["연분기"] = tmp["연분기"].astype(idx_q)
    piv = tmp.pivot_table(index="연분기", columns=group_cols, values=value_col,
                           aggfunc="sum", fill_value=0, observed=False)
    piv = piv.reindex(quarters, fill_value=0)
    return piv


def run_forecast(df, train_start, train_end, forecast_start, forecast_end,
                  exclude_processes=("미확인",)):
    all_quarters_hist = build_quarter_axis(df)
    train_quarters = quarter_range(train_start, train_end)
    for q in train_quarters:
        if q not in all_quarters_hist:
            pass  # 학습기간에 실적 없는 분기는 0으로 취급 (아래 reindex가 처리)
    fc_quarters = quarter_range(forecast_start, forecast_end)
    n_ahead = len(fc_quarters)

    combos = sorted(df.groupby(["공정", "모델"]).size().index.tolist())
    processes = sorted(df["공정"].unique().tolist())

    amt_piv = quarterly_pivot(df, ["공정", "모델"], "매출액(KRW)", train_quarters)
    qty_piv = quarterly_pivot(df, ["공정", "모델"], "수량", train_quarters)

    results = []
    for (proc, model) in combos:
        amt_hist = amt_piv[(proc, model)].values if (proc, model) in amt_piv.columns else np.zeros(len(train_quarters))
        qty_hist = qty_piv[(proc, model)].values if (proc, model) in qty_piv.columns else np.zeros(len(train_quarters))

        excluded = proc in exclude_processes
        if excluded:
            amt_fc = np.zeros(n_ahead)
            qty_fc = np.zeros(n_ahead)
            method = "제외(돌발성 매출)"
        else:
            amt_fc, method = forecast_series(amt_hist, n_ahead)
            qty_fc, _ = forecast_series(qty_hist, n_ahead)

        baseline_len = min(n_ahead, len(amt_hist))
        baseline_amt = float(np.sum(amt_hist[-baseline_len:])) if baseline_len else 0.0

        results.append({
            "공정": proc, "모델": model, "excluded": excluded, "method": method,
            "amt_hist": amt_hist, "qty_hist": qty_hist,
            "amt_fc": amt_fc, "qty_fc": qty_fc,
            "amt_fc_total": float(np.sum(amt_fc)), "qty_fc_total": float(np.sum(qty_fc)),
            "baseline_amt": baseline_amt,
            "diff_amt": float(np.sum(amt_fc)) - baseline_amt,
            "nz_count": int(np.count_nonzero(amt_hist)),
            "txn_count": int(((df["공정"] == proc) & (df["모델"] == model) &
                               (df["연분기"].isin(train_quarters))).sum()),
        })

    detail = pd.DataFrame(results)

    # 전사/공정 top-down (계절성 반영) — bottom-up과 교차검증용
    total_hist = amt_piv.sum(axis=1).values
    total_fc_topdown, total_method_topdown = forecast_series(total_hist, n_ahead)

    proc_topdown = {}
    for proc in processes:
        if proc in exclude_processes:
            continue
        cols = [c for c in amt_piv.columns if c[0] == proc]
        if not cols:
            proc_topdown[proc] = (np.zeros(n_ahead), "데이터없음")
            continue
        series = amt_piv[cols].sum(axis=1).values
        fc, m = forecast_series(series, n_ahead)
        proc_topdown[proc] = (fc, m)

    return {
        "train_quarters": train_quarters,
        "fc_quarters": fc_quarters,
        "detail": detail,
        "total_hist": total_hist,
        "total_fc_topdown": total_fc_topdown,
        "total_method_topdown": total_method_topdown,
        "proc_topdown": proc_topdown,
        "amt_piv": amt_piv,
    }


# =====================================================================
# 차트
# =====================================================================
def _fig_to_base64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=140, bbox_inches="tight")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def chart_total_trend(train_quarters, total_hist, fc_quarters, total_fc_bottomup, total_fc_topdown):
    fig, ax = plt.subplots(figsize=(9, 4))
    ax.plot(range(len(train_quarters)), total_hist, color=BRAND_GRAY, marker="o", markersize=3,
            linewidth=1.6, label="실적")
    n_hist = len(train_quarters)
    fc_x = range(n_hist, n_hist + len(fc_quarters))
    ax.plot(fc_x, total_fc_bottomup, color=BRAND_MAGENTA, marker="D", markersize=5,
            linewidth=1.8, label="예측(상향식: 조합별 합산)")
    ax.plot(fc_x, total_fc_topdown, color="#1f77b4", marker="s", markersize=5,
            linewidth=1.4, linestyle="--", label="예측(하향식: 전사 계절성 모델)")
    # connect last actual to forecast start
    ax.plot([n_hist - 1, n_hist], [total_hist[-1], total_fc_bottomup[0]], color=BRAND_MAGENTA, linewidth=1, alpha=0.5)
    ax.plot([n_hist - 1, n_hist], [total_hist[-1], total_fc_topdown[0]], color="#1f77b4", linewidth=1, alpha=0.5, linestyle="--")

    all_labels = list(train_quarters) + list(fc_quarters)
    step = max(1, len(all_labels) // 16)
    ax.set_xticks(range(0, len(all_labels), step))
    ax.set_xticklabels([all_labels[i] for i in range(0, len(all_labels), step)], rotation=45, ha="right", fontsize=8)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, p: f"{v/1e8:,.0f}억"))
    ax.set_title("전사 분기별 매출 추이 및 예측", fontsize=12, color=BRAND_MAGENTA, fontweight="bold")
    ax.legend(fontsize=8, loc="upper left")
    ax.grid(alpha=0.25)
    fig.tight_layout()
    return _fig_to_base64(fig)


def chart_process_bar(proc_summary):
    fig, ax = plt.subplots(figsize=(9, 4.5))
    procs = proc_summary["공정"].tolist()
    x = np.arange(len(procs))
    w = 0.38
    ax.bar(x - w / 2, proc_summary["직전동기간실적"] / 1e8, width=w, color=BRAND_GRAY, label="직전 동기간 실적")
    ax.bar(x + w / 2, proc_summary["예측합계"] / 1e8, width=w, color=BRAND_MAGENTA, label="예측")
    ax.set_xticks(x)
    ax.set_xticklabels(procs, rotation=30, ha="right", fontsize=9)
    ax.set_ylabel("억원")
    ax.set_title("공정별 실적 대비 예측", fontsize=12, color=BRAND_MAGENTA, fontweight="bold")
    ax.legend(fontsize=9)
    ax.grid(alpha=0.25, axis="y")
    fig.tight_layout()
    return _fig_to_base64(fig)


def chart_top_contributors(detail, top_n=10):
    d = detail[~detail["excluded"]].copy()
    d = d.reindex(d["diff_amt"].abs().sort_values(ascending=False).index).head(top_n)
    d = d.sort_values("diff_amt")
    labels = [f"{p}-{m}" for p, m in zip(d["공정"], d["모델"])]
    colors = [BRAND_MAGENTA if v >= 0 else "#4472C4" for v in d["diff_amt"]]
    fig, ax = plt.subplots(figsize=(9, 5))
    ax.barh(labels, d["diff_amt"] / 1e8, color=colors)
    ax.set_xlabel("증감 기여액 (억원)")
    ax.set_title(f"증감 기여도 상위 {top_n}개 조합", fontsize=12, color=BRAND_MAGENTA, fontweight="bold")
    ax.grid(alpha=0.25, axis="x")
    fig.tight_layout()
    return _fig_to_base64(fig)


# =====================================================================
# HTML 리포트
# =====================================================================
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ko"><head><meta charset="utf-8">
<title>매출 예측 리포트</title>
<style>
body {{ font-family: 'Malgun Gothic','Apple SD Gothic Neo',sans-serif; margin:0; padding:0; background:#fafafa; color:#222; }}
.wrap {{ max-width: 980px; margin: 0 auto; padding: 32px 24px 80px; }}
h1 {{ color:{magenta}; font-size:24px; margin-bottom:4px;}}
h2 {{ color:{magenta}; font-size:18px; margin-top:40px; border-bottom:2px solid {magenta}; padding-bottom:6px;}}
.meta {{ color:{gray}; font-size:13px; margin-bottom:24px;}}
.note {{ color:#C00000; font-size:13px; background:#FFF4F4; padding:10px 14px; border-radius:6px; margin:12px 0;}}
table {{ border-collapse: collapse; width:100%; margin:14px 0; font-size:13px;}}
th {{ background:{magenta}; color:white; padding:8px 10px; text-align:center;}}
td {{ padding:7px 10px; border-bottom:1px solid #eee; text-align:right;}}
td:first-child, td:nth-child(2) {{ text-align:left; }}
tr:hover {{ background:#fafafa; }}
.pos {{ color:#1F7A1F; font-weight:bold; }}
.neg {{ color:#C00000; font-weight:bold; }}
img {{ max-width:100%; border-radius:8px; box-shadow:0 1px 6px rgba(0,0,0,0.08); }}
.card {{ background:white; border-radius:10px; padding:20px 24px; margin:16px 0; box-shadow:0 1px 6px rgba(0,0,0,0.06);}}
.summary-box {{ display:flex; gap:16px; flex-wrap:wrap; margin:16px 0;}}
.stat {{ background:white; border-radius:10px; padding:16px 20px; flex:1; min-width:180px; box-shadow:0 1px 6px rgba(0,0,0,0.06);}}
.stat .label {{ font-size:12px; color:{gray}; }}
.stat .value {{ font-size:20px; font-weight:bold; color:{magenta}; margin-top:4px;}}
.methodbadge {{ display:inline-block; font-size:11px; padding:2px 8px; border-radius:10px; background:#EEE; color:#555;}}
</style></head>
<body><div class="wrap">
<h1>매출 예측 리포트</h1>
<div class="meta">학습기간: {train_start} ~ {train_end} &nbsp;|&nbsp; 예측기간: {fc_start} ~ {fc_end} &nbsp;|&nbsp; 생성 파일: {src_name}</div>
{exclude_note}

<div class="summary-box">
  <div class="stat"><div class="label">직전 동기간 실적 (전사, 미확인 제외)</div><div class="value">{total_base_fmt}원</div></div>
  <div class="stat"><div class="label">예측 합계 (상향식)</div><div class="value">{total_fc_fmt}원</div></div>
  <div class="stat"><div class="label">증감률</div><div class="value">{total_pct_fmt}</div></div>
  <div class="stat"><div class="label">하향식(전사 계절성모델) 예측</div><div class="value">{total_topdown_fmt}원</div><div class="methodbadge">{topdown_method}</div></div>
</div>

<div class="card"><h2 style="margin-top:0;border:none;">전사 매출 추이 및 예측</h2>
<img src="data:image/png;base64,{chart_total}"/>
<p style="font-size:13px;color:{gray}">실선(회색)은 실적, 마름모(마젠타)는 조합별 예측을 합산한 상향식 결과, 파란 점선은 전사 데이터 자체에 계절성 모델을 적용한 하향식 결과입니다. 두 방식이 크게 어긋나면 개별 조합의 이상치를 의심해볼 수 있습니다.</p>
</div>

<div class="card"><h2 style="margin-top:0;border:none;">공정별 실적 대비 예측</h2>
<img src="data:image/png;base64,{chart_proc}"/>
{proc_table}
</div>

<div class="card"><h2 style="margin-top:0;border:none;">증감 기여도 상위 조합</h2>
<img src="data:image/png;base64,{chart_top}"/>
<p style="font-size:13px;color:{gray}">{narrative}</p>
</div>

<h2>공정×모델 상세 예측 (전체 {n_combo}개 조합)</h2>
{detail_table}

<h2>예측 방법론</h2>
<div class="card" style="font-size:13px; line-height:1.7;">
<b>{ets_count}개 조합</b>: 학습기간 8분기 이상 &amp; 활동분기 비율 50% 이상 → <b>계절성 지수평활(Holt-Winters)</b> 적용. 추세와 분기별 반복 패턴을 함께 반영합니다.<br>
<b>{croston_count}개 조합</b>: 거래는 있지만 간헐적(활동분기 적음) → <b>Croston-SBA 간헐수요모델</b> 적용. 언제/얼마나 발생하는지의 확률적 패턴을 반영한 분기당 평균 발생률로 예측합니다.<br>
<b>{avg_count}개 조합</b>: 표본이 매우 적음 → 과거 평균을 그대로 유지(보수적 접근)<br>
<b>{excl_count}개 조합</b>: 공정이 '미확인'으로 돌발성(일회성) 매출로 분류되어 예측에서 제외(과거 실적은 참고용으로만 표시)<br>
모든 예측치는 0 미만이 되지 않도록 하한을 적용했습니다.
</div>

</div></body></html>"""


def fmt_krw(v):
    return f"{v:,.0f}"


def build_html_report(df, result, src_name, exclude_processes):
    detail = result["detail"]
    train_quarters, fc_quarters = result["train_quarters"], result["fc_quarters"]
    n_ahead = len(fc_quarters)

    active = detail[~detail["excluded"]]
    total_base = active["baseline_amt"].sum()
    total_fc = active["amt_fc_total"].sum()
    total_pct = (total_fc - total_base) / total_base if total_base else None
    total_topdown = float(np.sum(result["total_fc_topdown"]))

    proc_rows = []
    for proc, (fc, method) in result["proc_topdown"].items():
        sub = active[active["공정"] == proc]
        base = sub["baseline_amt"].sum()
        fc_sum = sub["amt_fc_total"].sum()
        proc_rows.append({"공정": proc, "직전동기간실적": base, "예측합계": fc_sum,
                           "증감액": fc_sum - base, "증감률": (fc_sum - base) / base if base else np.nan})
    proc_summary = pd.DataFrame(proc_rows).sort_values("예측합계", ascending=False)

    total_fc_bottomup_arr = np.sum(np.stack(active["amt_fc"].values), axis=0) if len(active) else np.zeros(n_ahead)
    chart_total = chart_total_trend(train_quarters, result["total_hist"], fc_quarters,
                                     total_fc_bottomup_arr, result["total_fc_topdown"])
    chart_proc = chart_process_bar(proc_summary)
    chart_top = chart_top_contributors(detail)

    def pct_cell(v):
        if pd.isna(v):
            return '<span class="methodbadge">N/A</span>'
        cls = "pos" if v >= 0 else "neg"
        sign = "+" if v >= 0 else ""
        return f'<span class="{cls}">{sign}{v*100:.1f}%</span>'

    proc_table_rows = "\n".join(
        f"<tr><td>{r['공정']}</td><td>{fmt_krw(r['직전동기간실적'])}</td><td>{fmt_krw(r['예측합계'])}</td>"
        f"<td>{fmt_krw(r['증감액'])}</td><td>{pct_cell(r['증감률'])}</td></tr>"
        for _, r in proc_summary.iterrows()
    )
    proc_table = (f"<table><tr><th>공정</th><th>직전동기간실적</th><th>예측합계</th><th>증감액</th><th>증감률</th></tr>"
                  f"{proc_table_rows}</table>")

    top3 = detail[~detail["excluded"]].reindex(
        detail[~detail["excluded"]]["diff_amt"].abs().sort_values(ascending=False).index).head(3)
    if total_pct is not None and len(top3):
        narrative = (f"설정한 예측기간({fc_quarters[0]}~{fc_quarters[-1]}) 전사 합계 예측은 {fmt_krw(total_fc)}원으로, "
                     f"직전 동일 길이 기간 실적({fmt_krw(total_base)}원) 대비 {'+' if total_pct>=0 else ''}{total_pct*100:.1f}% "
                     f"{'증가' if total_pct>=0 else '감소'}입니다. 가장 큰 요인은 " +
                     ", ".join(f"{r['공정']}-{r['모델']}({'+' if r['diff_amt']>=0 else ''}{fmt_krw(r['diff_amt'])}원)"
                               for _, r in top3.iterrows()) + " 입니다.")
    else:
        narrative = ""

    det_sorted = detail.sort_values("amt_fc_total", ascending=False)
    detail_rows = []
    for _, r in det_sorted.iterrows():
        cls = "" if r["excluded"] else ("pos" if r["diff_amt"] >= 0 else "neg")
        detail_rows.append(
            f"<tr><td>{r['공정']}</td><td>{r['모델']}</td><td>{fmt_krw(r['baseline_amt'])}</td>"
            f"<td>{fmt_krw(r['amt_fc_total'])}</td><td>{r['qty_fc_total']:.1f}</td>"
            f"<td><span class='methodbadge'>{r['method']}</span></td></tr>")
    detail_table = (f"<table><tr><th>공정</th><th>모델</th><th>직전동기간실적(금액)</th>"
                     f"<th>예측합계(금액)</th><th>예측합계(수량)</th><th>적용방법</th></tr>"
                     f"{''.join(detail_rows)}</table>")

    exclude_note = ""
    if exclude_processes:
        exclude_note = (f'<div class="note">공정이 {", ".join(exclude_processes)} 인 매출은 돌발성(일회성) '
                         f"매출로 간주되어 예측 및 위 합계 지표에서 제외되었습니다. (과거 실적은 상세표에서 참고용으로 확인 가능)</div>")

    method_counts = detail[~detail["excluded"]]["method"].value_counts()

    html = HTML_TEMPLATE.format(
        magenta=BRAND_MAGENTA, gray=BRAND_GRAY,
        train_start=train_quarters[0], train_end=train_quarters[-1],
        fc_start=fc_quarters[0], fc_end=fc_quarters[-1], src_name=src_name,
        exclude_note=exclude_note,
        total_base_fmt=fmt_krw(total_base), total_fc_fmt=fmt_krw(total_fc),
        total_pct_fmt=(f"{'+' if total_pct>=0 else ''}{total_pct*100:.1f}%" if total_pct is not None else "N/A"),
        total_topdown_fmt=fmt_krw(total_topdown), topdown_method=result["total_method_topdown"],
        chart_total=chart_total, chart_proc=chart_proc, chart_top=chart_top,
        proc_table=proc_table, narrative=narrative,
        n_combo=len(detail), detail_table=detail_table,
        ets_count=int(method_counts.get("계절성 지수평활(Holt-Winters)", 0)),
        croston_count=int(method_counts.get("간헐수요모델(Croston-SBA)", 0)),
        avg_count=int(method_counts.get("표본부족(평균유지)", 0)) + int(method_counts.get("선형추세(ETS 적합 실패 → 대체)", 0)),
        excl_count=int(detail["excluded"].sum()),
    )
    return html, proc_summary, total_base, total_fc, total_pct, total_topdown


# =====================================================================
# PDF 리포트 (요약 위주, 가벼운 버전)
# =====================================================================
def build_pdf_report(outpath, df, result, src_name, exclude_processes,
                      proc_summary, total_base, total_fc, total_pct, total_topdown,
                      chart_total_png_b64, chart_proc_png_b64, chart_top_png_b64):
    from fpdf import FPDF

    train_quarters, fc_quarters = result["train_quarters"], result["fc_quarters"]
    detail = result["detail"]

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    FONT_CANDIDATES = [
        # Windows
        "C:/Windows/Fonts/malgun.ttf", "C:/Windows/Fonts/malgunbd.ttf",
        # macOS
        "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
        "/Library/Fonts/AppleGothic.ttf",
        # Linux (Ubuntu/Debian with fonts-nanum or fonts-noto-cjk installed)
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansKR-Regular.otf",
    ]
    font_path = next((p for p in FONT_CANDIDATES if Path(p).exists()), None)
    if font_path is None:
        # 마지막 수단: 시스템 폰트 디렉터리를 뒤져서 한글이 포함될 만한 폰트를 찾는다
        for d in ["/usr/share/fonts", "C:/Windows/Fonts", "/System/Library/Fonts",
                  str(Path.home() / "Library/Fonts"), str(Path.home() / ".fonts")]:
            p = Path(d)
            if p.exists():
                hits = list(p.rglob("*gothic*")) + list(p.rglob("*Gothic*")) + \
                       list(p.rglob("*nanum*")) + list(p.rglob("*Nanum*")) + \
                       list(p.rglob("*malgun*"))
                if hits:
                    font_path = str(hits[0])
                    break

    pdf.add_page()
    if font_path:
        try:
            pdf.add_font("Kr", "", font_path)
            pdf.set_font("Kr", "", 18)
        except Exception:
            font_path = None
    if not font_path:
        pdf.set_font("Helvetica", "B", 18)
    pdf.set_text_color(200, 0, 124)
    pdf.cell(0, 12, "매출 예측 리포트", ln=True)
    pdf.set_font(pdf.font_family, "", 10)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 8, f"학습기간 {train_quarters[0]}~{train_quarters[-1]}  |  "
                    f"예측기간 {fc_quarters[0]}~{fc_quarters[-1]}  |  원본: {src_name}", ln=True)
    pdf.ln(2)

    if exclude_processes:
        pdf.set_text_color(192, 0, 0)
        pdf.set_font(pdf.font_family, "", 9)
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(0, 6, f"※ 공정이 {', '.join(exclude_processes)}인 매출은 돌발성(일회성)으로 간주해 예측/합계에서 제외했습니다.")
        pdf.ln(1)

    pdf.set_text_color(30, 30, 30)
    pdf.set_font(pdf.font_family, "", 11)
    pct_txt = f"{'+' if total_pct is not None and total_pct>=0 else ''}{total_pct*100:.1f}%" if total_pct is not None else "N/A"
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(0, 7, f"전사 직전동기간 실적: {fmt_krw(total_base)}원\n"
                         f"전사 예측 합계(상향식): {fmt_krw(total_fc)}원  ({pct_txt})\n"
                         f"전사 예측(하향식/계절성모델): {fmt_krw(total_topdown)}원")
    pdf.ln(3)

    img1 = io.BytesIO(base64.b64decode(chart_total_png_b64))
    pdf.image(img1, w=180)
    pdf.ln(2)

    pdf.add_page()
    pdf.set_font(pdf.font_family, "", 13)
    pdf.set_text_color(200, 0, 124)
    pdf.cell(0, 10, "공정별 실적 대비 예측", ln=True)
    img2 = io.BytesIO(base64.b64decode(chart_proc_png_b64))
    pdf.image(img2, w=180)
    pdf.ln(2)
    pdf.set_font(pdf.font_family, "", 9)
    pdf.set_text_color(30, 30, 30)
    pdf.set_fill_color(200, 0, 124)
    pdf.set_text_color(255, 255, 255)
    headers = ["공정", "직전실적", "예측합계", "증감률"]
    widths = [50, 45, 45, 40]
    for h, w in zip(headers, widths):
        pdf.cell(w, 8, h, border=1, align="C", fill=True)
    pdf.ln()
    pdf.set_text_color(30, 30, 30)
    for _, r in proc_summary.iterrows():
        pct = r["증감률"]
        pct_s = f"{'+' if pd.notna(pct) and pct>=0 else ''}{pct*100:.1f}%" if pd.notna(pct) else "N/A"
        pdf.cell(widths[0], 7, str(r["공정"]), border=1)
        pdf.cell(widths[1], 7, fmt_krw(r["직전동기간실적"]), border=1, align="R")
        pdf.cell(widths[2], 7, fmt_krw(r["예측합계"]), border=1, align="R")
        pdf.cell(widths[3], 7, pct_s, border=1, align="R")
        pdf.ln()

    pdf.add_page()
    pdf.set_font(pdf.font_family, "", 13)
    pdf.set_text_color(200, 0, 124)
    pdf.cell(0, 10, "증감 기여도 상위 조합", ln=True)
    img3 = io.BytesIO(base64.b64decode(chart_top_png_b64))
    pdf.image(img3, w=180)

    pdf.add_page()
    pdf.set_font(pdf.font_family, "", 13)
    pdf.set_text_color(200, 0, 124)
    pdf.cell(0, 10, "예측 방법론 요약", ln=True)
    pdf.set_font(pdf.font_family, "", 10)
    pdf.set_text_color(30, 30, 30)
    method_counts = detail[~detail["excluded"]]["method"].value_counts()
    lines = [
        f"- 계절성 지수평활(Holt-Winters): {int(method_counts.get('계절성 지수평활(Holt-Winters)', 0))}개 조합 "
        "(학습기간 8분기 이상 & 활동분기 비율 50% 이상)",
        f"- 간헐수요모델(Croston-SBA): {int(method_counts.get('간헐수요모델(Croston-SBA)', 0))}개 조합 (거래는 있으나 간헐적)",
        f"- 표본부족(평균유지): {int(method_counts.get('표본부족(평균유지)', 0)) + int(method_counts.get('선형추세(ETS 적합 실패 → 대체)', 0))}개 조합",
        f"- 제외(돌발성 매출): {int(detail['excluded'].sum())}개 조합",
        "- 모든 예측치는 0 미만이 되지 않도록 하한을 적용했습니다.",
        "- 상세 조합별 결과와 그래프는 함께 생성된 HTML 리포트에서 확인하실 수 있습니다.",
    ]
    for ln_txt in lines:
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(0, 7, ln_txt)

    pdf.output(str(outpath))


# =====================================================================
# 메인
# =====================================================================
def main():
    ap = argparse.ArgumentParser(description="RAWDATA 엑셀로부터 분기별 매출 예측 리포트를 생성합니다.")
    ap.add_argument("input", help="RAWDATA 엑셀 파일 경로")
    ap.add_argument("--train-start", default=None, help="학습 시작분기 (예: 2012Q1). 기본값: 데이터의 첫 분기")
    ap.add_argument("--train-end", default=None, help="학습 종료분기 (예: 2025Q4). 기본값: 데이터의 마지막 분기")
    ap.add_argument("--forecast-start", default=None, help="예측 시작분기. 기본값: 학습종료분기 다음 분기")
    ap.add_argument("--forecast-end", default=None, help="예측 종료분기. 기본값: 예측시작분기+3(4개 분기)")
    ap.add_argument("--exclude-process", nargs="*", default=["미확인"],
                     help="돌발성 매출로 간주해 예측에서 제외할 공정명 목록 (기본값: 미확인). 없애려면 --exclude-process 를 빈 값으로")
    ap.add_argument("--outdir", default="./output", help="결과 저장 폴더 (기본값: ./output)")
    args = ap.parse_args()

    src_path = Path(args.input)
    if not src_path.exists():
        sys.exit(f"[오류] 파일을 찾을 수 없습니다: {src_path}")

    print(f"[1/5] RAWDATA 읽는 중... ({src_path.name})")
    df = load_rawdata(src_path)
    print(f"      -> {len(df):,}건, {df['매출일'].min().date()} ~ {df['매출일'].max().date()}")

    all_quarters = build_quarter_axis(df)
    train_start = args.train_start or all_quarters[0]
    train_end = args.train_end or all_quarters[-1]
    forecast_start = args.forecast_start or add_quarters(train_end, 1)
    forecast_end = args.forecast_end or add_quarters(forecast_start, 3)

    print(f"[2/5] 예측 계산 중... 학습 {train_start}~{train_end} / 예측 {forecast_start}~{forecast_end}")
    result = run_forecast(df, train_start, train_end, forecast_start, forecast_end,
                           exclude_processes=tuple(args.exclude_process or []))

    print("[3/5] 그래프/리포트 생성 중...")
    html, proc_summary, total_base, total_fc, total_pct, total_topdown = build_html_report(
        df, result, src_path.name, args.exclude_process or [])

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    html_path = outdir / "매출예측_리포트.html"
    html_path.write_text(html, encoding="utf-8")
    print(f"      -> {html_path}")

    print("[4/5] PDF 생성 중...")
    # HTML에서 이미 만든 차트를 재사용하기 위해 다시 계산(가벼운 연산이라 재실행)
    active = result["detail"][~result["detail"]["excluded"]]
    total_fc_bottomup_arr = np.sum(np.stack(active["amt_fc"].values), axis=0) if len(active) else np.zeros(len(result["fc_quarters"]))
    c1 = chart_total_trend(result["train_quarters"], result["total_hist"], result["fc_quarters"],
                            total_fc_bottomup_arr, result["total_fc_topdown"])
    c2 = chart_process_bar(proc_summary)
    c3 = chart_top_contributors(result["detail"])
    pdf_path = outdir / "매출예측_리포트.pdf"
    build_pdf_report(pdf_path, df, result, src_path.name, args.exclude_process or [],
                      proc_summary, total_base, total_fc, total_pct, total_topdown, c1, c2, c3)
    print(f"      -> {pdf_path}")

    print("[5/5] 완료!")
    print(f"\n전사 직전동기간 실적: {fmt_krw(total_base)}원")
    print(f"전사 예측 합계(상향식): {fmt_krw(total_fc)}원 "
          f"({'+' if total_pct is not None and total_pct>=0 else ''}{total_pct*100:.1f}%)" if total_pct is not None else "")
    print(f"전사 예측(하향식/계절성모델): {fmt_krw(total_topdown)}원")


if __name__ == "__main__":
    main()
