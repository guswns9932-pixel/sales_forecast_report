# -*- coding: utf-8 -*-
"""
sales_forecast_report.py
=========================
RAWDATA 엑셀(매출일/모델/반입라인/공정/메이커/수량/단가/매출액 포함)을 읽어
분기/월/년 단위로 매출을 예측하고 HTML + PDF 리포트를 생성하는 독립 실행 스크립트입니다.

주요 기능:
  - 더블클릭(또는 인자 없이 실행) 시 콘솔(DOS) 창 없이 하나의 윈도우 창에서
    파일 선택 → 시트/기간단위 선택 → 학습·예측기간 선택 → 실행 → 로그 확인까지 진행
  - 예측기간 단위(분기/월/년)를 직접 선택하고, 학습·예측 시작/종료 시점을 직접 지정
  - 공정×메이커(설비)×모델 조합 단위로 예측하고, 공정별/메이커(설비)별/모델별로
    묶어서 볼 수 있는 리포트 생성
  - 여러 예측 알고리즘(계절성 지수평활, Croston-SBA 간헐수요모델, 선형추세,
    이동평균, 계절성 단순모형, 평균유지)을 실제 로우데이터로 백테스트하여
    조합마다 가장 정합성 높은(오차가 가장 작은) 알고리즘을 자동 선택
  - 리포트 그래프는 값이 겹쳐 읽기 어렵지 않도록 단순하게 유지하고, 정확한 값은 그래프 아래
    표에서 과거 실적까지 포함해 확인
  - 리포트의 모든 금액은 "685.35억원"처럼 억원 단위로 축약 표기
  - 실행 전 과정(백테스트 점수, 선택 근거, 데이터 처리 내역 등)을 메모장으로
    바로 열어볼 수 있는 텍스트 로그 파일로 저장

사용법 (콘솔/자동화용):
    python sales_forecast_report.py                                (창이 뜹니다)
    python sales_forecast_report.py RAWDATA.xlsx
    python sales_forecast_report.py RAWDATA.xlsx --granularity 월 --forecast-start 2026-01 --forecast-end 2026-06
    python sales_forecast_report.py RAWDATA.xlsx --granularity 분기 --train-end 2024Q4 --forecast-start 2025Q1 --forecast-end 2025Q4

자세한 옵션은:
    python sales_forecast_report.py --help
"""
import argparse
import base64
import io
import logging
import os
import queue
import re
import subprocess
import sys
import threading
import warnings
from datetime import datetime
from pathlib import Path


def _pause_and_exit(message, code=1):
    """더블클릭 실행 시 창이 즉시 닫혀버리는 것을 막기 위해, 종료 전 사용자 입력을 기다린다."""
    print(message, file=sys.stderr)
    try:
        input("\n엔터 키를 누르면 창을 닫습니다...")
    except (EOFError, KeyboardInterrupt):
        pass
    sys.exit(code)


try:
    import numpy as np
    import pandas as pd
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.ticker as mticker
except ImportError as e:
    _pause_and_exit(
        f"[오류] 필요한 라이브러리가 설치되어 있지 않습니다: {e}\n\n"
        "아래 명령을 실행해 필요한 패키지를 설치한 뒤 다시 실행해주세요:\n"
        "    pip install pandas numpy matplotlib openpyxl statsmodels fpdf2"
    )

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except Exception:
    tk = None

warnings.filterwarnings("ignore")
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


def _hide_console_window():
    """Windows에서 python.exe로 실행되어 함께 뜬 DOS 콘솔 창을 숨긴다.
    (GUI 창만 보이도록 하기 위함. 실패해도 무시하고 진행)"""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE
    except Exception:
        pass


# =====================================================================
# 로깅 (메모장으로 바로 열어볼 수 있는 텍스트 로그)
# =====================================================================
class QueueLogHandler(logging.Handler):
    """로그 레코드를 큐에 넣어 GUI 창(다른 스레드)에서 실시간으로 표시할 수 있게 한다."""

    def __init__(self, q):
        super().__init__()
        self.q = q

    def emit(self, record):
        try:
            self.q.put(("log", self.format(record)))
        except Exception:
            pass


def setup_logger(outdir, extra_handler=None, ts=None):
    """콘솔과 텍스트 파일(UTF-8 BOM, 메모장 호환) 양쪽에 동시에 기록하는 로거를 만든다.
    ts를 지정하면 같은 실행에서 생성되는 HTML/PDF 파일명과 시각을 통일할 수 있다."""
    outdir.mkdir(parents=True, exist_ok=True)
    ts = ts or datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = outdir / f"실행로그_{ts}.txt"

    logger = logging.getLogger("sales_forecast")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    logger.propagate = False

    file_fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    fh = logging.FileHandler(log_path, encoding="utf-8-sig")
    fh.setFormatter(file_fmt)
    logger.addHandler(fh)

    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(sh)

    if extra_handler is not None:
        extra_handler.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(extra_handler)

    return logger, log_path


# =====================================================================
# 기간(분기/월/년) 유틸리티
# =====================================================================
GRANULARITIES = ["분기", "월", "년"]

GRANULARITY_META = {
    "분기": {"season": 4, "window": 4, "example": "2026Q1", "default_horizon": 4},
    "월": {"season": 12, "window": 12, "example": "2026-01", "default_horizon": 12},
    "년": {"season": 1, "window": 3, "example": "2026", "default_horizon": 3},
}


def parse_period(label, gran):
    """'2026Q1'/'2026-01'/'2026' -> (year, sub) : gran에 맞는 형식만 허용."""
    s = str(label).strip()
    if gran == "분기":
        m = re.match(r"^(\d{4})\s*Q\s*([1-4])$", s, re.IGNORECASE)
        if not m:
            raise ValueError(f"분기 형식이 올바르지 않습니다: {label!r} (예: 2026Q1)")
        return int(m.group(1)), int(m.group(2))
    if gran == "월":
        m = re.match(r"^(\d{4})-(\d{1,2})$", s)
        if not m or not (1 <= int(m.group(2)) <= 12):
            raise ValueError(f"월 형식이 올바르지 않습니다: {label!r} (예: 2026-01)")
        return int(m.group(1)), int(m.group(2))
    if gran == "년":
        m = re.match(r"^(\d{4})$", s)
        if not m:
            raise ValueError(f"년 형식이 올바르지 않습니다: {label!r} (예: 2026)")
        return int(m.group(1)), 1
    raise ValueError(f"알 수 없는 기간 단위입니다: {gran!r} (분기/월/년 중 하나여야 합니다)")


def period_label(y, sub, gran):
    if gran == "분기":
        return f"{y}Q{sub}"
    if gran == "월":
        return f"{y}-{sub:02d}"
    if gran == "년":
        return f"{y}"
    raise ValueError(f"알 수 없는 기간 단위입니다: {gran!r}")


def period_index(y, sub, gran, base_year=2000):
    n = GRANULARITY_META[gran]["season"]
    return (y - base_year) * n + (sub - 1)


def index_to_period(idx, gran, base_year=2000):
    n = GRANULARITY_META[gran]["season"]
    y = base_year + idx // n
    sub = idx % n + 1
    return y, sub


def period_range(start_label, end_label, gran):
    """start_label ~ end_label 사이 모든 기간 라벨 리스트 (양끝 포함)"""
    sy, ss = parse_period(start_label, gran)
    ey, es = parse_period(end_label, gran)
    si, ei = period_index(sy, ss, gran), period_index(ey, es, gran)
    if ei < si:
        raise ValueError(f"종료기간({end_label})이 시작기간({start_label})보다 앞섭니다")
    out = []
    for i in range(si, ei + 1):
        y, s = index_to_period(i, gran)
        out.append(period_label(y, s, gran))
    return out


def add_periods(label, n, gran):
    y, s = parse_period(label, gran)
    idx = period_index(y, s, gran) + n
    y2, s2 = index_to_period(idx, gran)
    return period_label(y2, s2, gran)


# =====================================================================
# 데이터 로딩
# =====================================================================
REQUIRED_COLS = ["매출일", "모델", "반입라인", "공정", "메이커", "수량", "단가(KRW)", "매출액(KRW)"]

COLUMN_GUIDE = {
    "매출일": "매출(거래)이 발생한 날짜입니다. (예: 2025-03-15)",
    "모델": "판매된 장비/부품의 모델명입니다. 예측 시 '모델별' 기준이 됩니다.",
    "반입라인": "장비가 투입되는 고객사 생산라인 구분입니다.",
    "공정": "해당 매출이 속한 공정 단계입니다. (예: 증착, 식각, 세정 등) 예측 시 '공정별' 기준이 됩니다.",
    "메이커": "장비/부품 제조사(설비 메이커)입니다. 예측 시 '메이커(설비)별' 기준이 됩니다.",
    "수량": "판매 수량입니다. 숫자만 입력되어야 합니다.",
    "단가(KRW)": "개당 단가(원화)입니다.",
    "매출액(KRW)": "총 매출액(원화)입니다. 보통 수량 × 단가로 계산됩니다.",
}


def column_guide_text():
    lines = ["[RAWDATA 필수 컬럼 안내] 아래 8개 컬럼이 선택한 시트에 정확한 이름으로 있어야 합니다:"]
    for col, desc in COLUMN_GUIDE.items():
        lines.append(f"  - {col}: {desc}")
    lines.append("컬럼명은 대소문자/띄어쓰기까지 정확히 일치해야 합니다.")
    lines.append("1행 헤더 다음 2~3행에 '필수/권장/선택/자동' 같은 안내문구가 있는 표준 양식과,")
    lines.append("안내문구 없이 2행부터 바로 데이터가 시작하는 단순 표 양식을 모두 지원합니다.")
    return "\n".join(lines)


def detect_default_sheet(sheet_names):
    return "RAWDATA" if "RAWDATA" in sheet_names else sheet_names[0]


def load_rawdata(path, sheet, logger=None):
    """지정한 시트를 읽는다. 표준 양식(1행 헤더, 2~3행 안내, 4행부터 데이터)과
    일반적인 단순 표(1행 헤더, 2행부터 데이터) 둘 다 지원한다."""
    raw_preview = pd.read_excel(path, sheet_name=sheet, header=0, nrows=3)
    tag_row_present = raw_preview.iloc[0].astype(str).str.contains("필수|권장|선택|자동").any()
    skiprows = [1, 2] if tag_row_present else None

    df = pd.read_excel(path, sheet_name=sheet, header=0, skiprows=skiprows)
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(
            f"'{sheet}' 시트에 다음 필수 컬럼이 없습니다: {missing}\n"
            f"실제 컬럼: {list(df.columns)}\n\n{column_guide_text()}"
        )

    n_before = len(df)
    df = df.dropna(subset=["매출일", "공정", "메이커", "모델"]).copy()
    n_dropped = n_before - len(df)
    if logger and n_dropped:
        logger.warning(f"매출일/공정/메이커/모델 중 결측값이 있는 {n_dropped}건을 제외했습니다.")

    df["매출일"] = pd.to_datetime(df["매출일"])

    amt_num = pd.to_numeric(df["매출액(KRW)"], errors="coerce")
    n_bad_amt = int(amt_num.isna().sum() - df["매출액(KRW)"].isna().sum())
    df["매출액(KRW)"] = amt_num.fillna(0)

    qty_num = pd.to_numeric(df["수량"], errors="coerce")
    n_bad_qty = int(qty_num.isna().sum() - df["수량"].isna().sum())
    df["수량"] = qty_num.fillna(0)

    if logger and (n_bad_amt or n_bad_qty):
        logger.warning(f"숫자로 변환할 수 없는 값을 0으로 처리했습니다: 매출액 {n_bad_amt}건, 수량 {n_bad_qty}건")

    return df


def add_period_column(df, gran):
    """선택한 기간 단위(분기/월/년)에 따라 '기간' 컬럼을 추가한다."""
    df = df.copy()
    if gran == "분기":
        df["기간"] = df["매출일"].dt.year.astype(str) + "Q" + df["매출일"].dt.quarter.astype(str)
    elif gran == "월":
        df["기간"] = df["매출일"].dt.strftime("%Y-%m")
    elif gran == "년":
        df["기간"] = df["매출일"].dt.year.astype(str)
    else:
        raise ValueError(f"알 수 없는 기간 단위입니다: {gran!r}")
    return df


def build_period_axis(df, gran, extra_future_periods=0):
    periods = sorted(df["기간"].unique(), key=lambda s: period_index(*parse_period(s, gran), gran))
    if extra_future_periods:
        last = periods[-1]
        for i in range(1, extra_future_periods + 1):
            periods.append(add_periods(last, i, gran))
    return periods


def detect_annual_only_years(df, gran):
    """gran이 '년'이 아닐 때, 실제 거래가 한 해 안에서 단 하나의 하위 기간에만 몰려 있는
    연도를 찾는다. (예: 오래된 데이터가 연간 합계 한 건으로만 입력된 경우) 이런 연도는
    그래프/표에서 선택한 기간단위 대신 '년' 단위로 요약해 보여준다."""
    if gran == "년":
        return set()
    years = df["매출일"].dt.year
    annual_years = set()
    for y in years.unique():
        n_periods = df.loc[years == y, "기간"].nunique()
        if n_periods <= 1:
            annual_years.add(int(y))
    return annual_years


def collapse_annual_periods(periods, values, gran, annual_only_years):
    """실적(과거) 축에서 annual_only_years에 해당하는 연속 구간을 연 단위 합계 하나로 합친다.
    예측(미래) 구간은 그대로 두고, 실적 표시에만 적용한다."""
    if gran == "년" or not annual_only_years:
        return list(periods), list(values)
    out_periods, out_values = [], []
    i, n = 0, len(periods)
    while i < n:
        y, _ = parse_period(periods[i], gran)
        if y in annual_only_years:
            total = 0.0
            j = i
            while j < n and parse_period(periods[j], gran)[0] == y:
                total += values[j]
                j += 1
            out_periods.append(str(y))
            out_values.append(total)
            i = j
        else:
            out_periods.append(periods[i])
            out_values.append(values[i])
            i += 1
    return out_periods, out_values


# =====================================================================
# 예측 알고리즘 후보
# =====================================================================
def croston_sba(values, alpha=0.1):
    """Croston's method + SBA 편의보정. 간헐적(0이 많은) 수요 시계열에 적합.
    반환값: 기간당 평균 수요율(모든 미래 기간에 동일하게 적용되는 flat 예측치)"""
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


def seasonal_ets_forecast(values, n_ahead, season_periods):
    """statsmodels Holt-Winters(가법 추세+가법 계절성, damped)로 향후 n_ahead 기간 예측.
    계절 주기가 2 미만이거나 적합에 실패하면 None 반환(호출부에서 폴백 처리)."""
    if not season_periods or season_periods < 2:
        return None
    from statsmodels.tsa.holtwinters import ExponentialSmoothing
    values = np.asarray(values, dtype=float)
    try:
        model = ExponentialSmoothing(
            values, trend="add", damped_trend=True,
            seasonal="add", seasonal_periods=season_periods,
            initialization_method="estimated",
        )
        fit = model.fit(optimized=True)
        fc = fit.forecast(n_ahead)
        return np.maximum(fc, 0.0)
    except Exception:
        return None


def linear_trend_forecast(values, n_ahead):
    """단순 선형회귀 추세 연장(0 하한)."""
    values = np.asarray(values, dtype=float)
    x = np.arange(len(values))
    if len(values) < 2 or np.all(values == values[0]):
        base = float(np.mean(values)) if len(values) else 0.0
        return np.full(n_ahead, max(0.0, base))
    slope, intercept = np.polyfit(x, values, 1)
    future_x = np.arange(len(values), len(values) + n_ahead)
    fc = slope * future_x + intercept
    return np.maximum(fc, 0.0)


def moving_average_forecast(values, n_ahead, window):
    """최근 window개 기간의 평균을 그대로 미래 예측치로 사용."""
    values = np.asarray(values, dtype=float)
    w = min(window, len(values)) if len(values) else 0
    base = float(np.mean(values[-w:])) if w else 0.0
    return np.full(n_ahead, max(0.0, base))


def seasonal_naive_forecast(values, n_ahead, season):
    """1년 전(직전 season개 기간) 같은 시점의 실적을 그대로 사용."""
    values = np.asarray(values, dtype=float)
    n = len(values)
    if not season or n < season:
        base = float(np.mean(values)) if n else 0.0
        return np.full(n_ahead, max(0.0, base))
    tail = values[-season:]
    fc = np.array([tail[i % season] for i in range(n_ahead)], dtype=float)
    return np.maximum(fc, 0.0)


def simple_mean_forecast(values, n_ahead):
    """전체 기간 평균을 그대로 유지."""
    values = np.asarray(values, dtype=float)
    base = float(np.mean(values)) if len(values) else 0.0
    return np.full(n_ahead, max(0.0, base))


def build_candidate_methods(season, window):
    """기간 단위(계절 주기 season, 이동평균 창 window)에 맞는 예측 알고리즘 후보를 구성한다.
    년 단위처럼 계절성을 정의할 수 없는(season < 2) 경우 계절성 알고리즘은 제외한다."""
    methods = {
        "선형추세": lambda v, h: linear_trend_forecast(v, h),
        "이동평균": lambda v, h: moving_average_forecast(v, h, window),
        "평균유지": lambda v, h: simple_mean_forecast(v, h),
        "간헐수요모델(Croston-SBA)": lambda v, h: np.full(h, croston_sba(v)),
    }
    if season and season >= 2:
        methods["계절성 지수평활(Holt-Winters)"] = lambda v, h: seasonal_ets_forecast(v, h, season)
        methods["계절성 단순모형(전년동기)"] = lambda v, h: seasonal_naive_forecast(v, h, season)
    return methods


METHOD_DESCRIPTIONS = {
    "계절성 지수평활(Holt-Winters)": "과거 성장 추세와 {unit}별로 반복되는 계절성 패턴을 함께 반영하는 통계 모델입니다.",
    "간헐수요모델(Croston-SBA)": "거래가 간헐적으로 발생하는 품목에 특화된 모델로, 발생 확률과 발생 시 평균 규모를 함께 추정합니다.",
    "선형추세": "과거 추세선을 직선으로 연장하는 단순한 방법입니다.",
    "이동평균": "가장 최근 {window}개 {unit}의 평균값을 다음 {unit} 예측치로 그대로 사용합니다.",
    "계절성 단순모형(전년동기)": "1년 전 같은 {unit}의 실적을 그대로 사용합니다. 추세는 약하지만 계절성이 뚜렷할 때 유리합니다.",
    "평균유지": "과거 전체 평균값을 보수적으로 유지합니다.",
    "표본부족(평균유지)": "유효한 거래 데이터가 너무 적어 백테스트를 수행할 수 없어, 과거 평균을 보수적으로 유지했습니다.",
    "제외(돌발성 매출)": "돌발성·일회성 매출로 분류되어 예측 대상에서 제외되었습니다(과거 실적은 참고용으로만 표시).",
}


def method_description(name, gran, window):
    tmpl = METHOD_DESCRIPTIONS.get(name, "선정된 통계적 방법으로 예측했습니다.")
    return tmpl.format(unit=gran, window=window)


def reader_guide_items(n_ahead, unit, method_count):
    """리포트 상단의 "읽는 법" 안내 내용. HTML과 PDF가 같은 문구를 공유해서
    두 결과물의 내용이 어긋나지 않도록 한다. 반환: [(제목, 설명), ...]"""
    return [
        ("직전 동기간 실적",
         f"예측기간과 길이가 같은 가장 최근 실적기간의 실제 매출입니다. "
         f"(예: {n_ahead}개 {unit}를 예측하면 직전 {n_ahead}개 {unit}의 실제 실적과 비교합니다)"),
        ("조합별 합산 예측(기본)",
         "공정×메이커×모델 조합 하나하나에 대해, 백테스트로 가장 정확했던 알고리즘을 적용해 예측한 값을 "
         "모두 더한 값입니다. 이 리포트의 기본 예측치입니다. (예전 표현: 상향식)"),
        ("전사 통합 예측(참고)",
         "조합별로 나누지 않고 전사 매출 전체를 하나로 보고 계절성 모델을 적용한 값입니다. "
         "'조합별 합산 예측'과 크게 차이가 나면 특정 조합에 이상치가 섞였을 가능성을 점검해볼 수 있습니다. (예전 표현: 하향식)"),
        ("두 예측 중 어느 쪽이 더 큰가요?",
         "정해진 규칙은 없습니다. 두 방식은 서로 다른 독립적인 계산이라 기간에 따라 어느 한쪽이 더 클 수도, "
         "작을 수도 있습니다. 항상 조합별 합산 예측이 더 크다거나 전사 통합 예측이 더 작다는 보장은 없으니, "
         "상세 데이터표에서 기간별로 직접 비교해보세요."),
        ("백테스트",
         f"과거 데이터를 학습구간과 검증구간으로 나눈 뒤, 학습구간만으로 검증구간을 예측해보고 실제값과 "
         f"얼마나 차이 나는지(오차율, sMAPE) 계산하는 절차입니다. 이 리포트는 {method_count}가지 예측 알고리즘을 "
         f"모두 백테스트해서 조합마다 오차가 가장 작은 알고리즘을 자동으로 선택합니다."),
        ("예측신뢰도",
         "백테스트 오차율 기준입니다 — 10% 이하 높음, 10~25% 보통, 25% 초과 낮음(참고용). 데이터가 너무 적어 "
         "백테스트 자체가 불가능했던 조합은 'N/A'로 표시하고 보수적으로 과거 평균을 유지합니다."),
        ("메이커(설비)", "RAWDATA의 '메이커' 컬럼을 설비 제조사 기준 구분으로 사용해 집계했습니다."),
        ("금액 표기", "모든 금액은 억원 단위로 축약해 표기합니다. (예: 68,535,079,090원 → 685.35억원)"),
        ("상한적용",
         f"예측신뢰도가 '낮음(참고용)'인 조합은 통계 모델이 불안정한 추세·계절성을 과도하게 연장해 "
         f"예측치가 비정상적으로 튈 수 있습니다. 이를 막기 위해 이런 조합은 한 {unit}의 예측치가 과거 최대 실적의 "
         f"{LOW_CONFIDENCE_CAP_MULTIPLIER:.0f}배를 넘지 않도록 자동으로 상한을 적용하고, 상세표에 '상한적용' "
         f"배지로 표시합니다."),
    ]


def _smape(actual, pred):
    """대칭 평균절대백분율오차. 0이 섞인 매출 데이터에서도 안정적으로 동작한다."""
    actual = np.asarray(actual, dtype=float)
    pred = np.asarray(pred, dtype=float)
    denom = np.abs(actual) + np.abs(pred)
    out = np.where(denom == 0, 0.0, 2 * np.abs(actual - pred) / np.where(denom == 0, 1, denom))
    return float(np.mean(out))


def confidence_label(scores, method):
    """백테스트 오차율(sMAPE) 기준의 예측 신뢰도 등급."""
    if not scores or method not in scores:
        return "N/A"
    s = scores[method]
    if s <= 0.10:
        return "높음"
    if s <= 0.25:
        return "보통"
    return "낮음(참고용)"


CONFIDENCE_RANK = {"높음": 0, "보통": 1, "낮음(참고용)": 2, "N/A": 3}

# 신뢰도가 "낮음(참고용)"인 조합은 백테스트 오차가 25%를 넘어 알고리즘이 불안정한 추세/계절성을
# 과도하게 연장했을 가능성이 있다. 이런 예측이 과거 최대 실적 대비 비정상적으로 튀어 리포트
# 총계를 왜곡하지 않도록, 한 기간의 예측치가 과거 최대 실적의 이 배수를 넘지 않게 상한을 둔다.
LOW_CONFIDENCE_CAP_MULTIPLIER = 2.0


def _cap_low_confidence_forecast(fc, values, method, scores, logger=None, series_name=""):
    """신뢰도 낮음(참고용) 조합의 예측치가 과거 최대 실적 대비 튀는 것을 막는다.
    반환: (조정된 예측 array, 상한이 실제로 적용됐는지 여부)"""
    if confidence_label(scores, method) != "낮음(참고용)":
        return fc, False
    values = np.asarray(values, dtype=float)
    nz = values[values > 0]
    if len(nz) == 0:
        return fc, False
    cap = float(np.max(nz)) * LOW_CONFIDENCE_CAP_MULTIPLIER
    if cap <= 0 or not np.any(fc > cap):
        return fc, False
    capped = np.minimum(fc, cap)
    if logger:
        logger.info(
            f"[{series_name}] 신뢰도 낮음(오차 {scores[method]*100:.1f}%) + 예측치가 과거 최대 실적의 "
            f"{LOW_CONFIDENCE_CAP_MULTIPLIER:.0f}배({cap:,.0f})를 초과해 상한을 적용했습니다."
        )
    return capped, True


def sort_detail_table(detail):
    """공정×메이커×모델 상세 예측표 정렬 기준: 예측수량 많음→적음, 동률이면 예측신뢰도 높음→낮음."""
    d = detail.copy()
    conf_rank = d.apply(
        lambda r: CONFIDENCE_RANK.get(confidence_label(r["backtest_scores"], r["method"]), 3), axis=1)
    d = d.assign(_conf_rank=conf_rank)
    return d.sort_values(["qty_fc_total", "_conf_rank"], ascending=[False, True]).drop(columns="_conf_rank")


def forecast_series(values, n_ahead, season=4, window=4, logger=None, series_name="", min_intermittent_obs=3):
    """여러 예측 알고리즘을 백테스트(과거 구간을 학습/검증으로 나눠 실제값과 비교)하여
    오차(sMAPE)가 가장 낮은, 즉 가장 정합성 높은 방법으로 최종 예측한다.
    데이터가 너무 짧아 백테스트가 불가능하면 표본 크기에 따른 보수적 방법으로 대체한다.
    신뢰도가 낮음(참고용)인 결과는 과거 최대 실적 대비 튀지 않도록 상한을 적용한다.
    반환: (forecast_array, method_name, backtest_scores dict[method_name -> smape], capped)"""
    values = np.asarray(values, dtype=float)
    n = len(values)
    nz_count = int(np.count_nonzero(values))

    if n == 0 or nz_count == 0:
        if logger:
            logger.info(f"[{series_name}] 거래 이력이 없어 예측치를 0으로 둡니다.")
        return np.zeros(n_ahead), "표본부족(평균유지)", {}, False

    methods = build_candidate_methods(season, window)
    test_h = max(1, min(n_ahead, 4))
    min_train = 4
    origins = sorted({o for o in (n - test_h, n - 2 * test_h) if o >= min_train}, reverse=True)

    scores = {}
    if origins:
        for name, fn in methods.items():
            errs = []
            for origin in origins:
                train_v = values[:origin]
                test_v = values[origin:origin + test_h]
                try:
                    pred = fn(train_v, test_h)
                except Exception:
                    pred = None
                if pred is None:
                    continue
                pred = np.asarray(pred, dtype=float)
                if len(pred) != len(test_v):
                    continue
                errs.append(_smape(test_v, pred))
            if errs:
                scores[name] = float(np.mean(errs))

    if not scores:
        if nz_count >= min_intermittent_obs:
            rate = croston_sba(values)
            if logger:
                logger.info(f"[{series_name}] 데이터가 짧아(n={n}) 백테스트 불가 → 간헐수요모델(Croston-SBA) 적용")
            return np.full(n_ahead, rate), "간헐수요모델(Croston-SBA)", {}, False
        avg = float(np.mean(values[values > 0])) if nz_count > 0 else 0.0
        if logger:
            logger.info(f"[{series_name}] 표본 부족(n={n}, 유효거래 {nz_count}건) → 평균유지 적용")
        return np.full(n_ahead, avg), "표본부족(평균유지)", {}, False

    ranked = sorted(scores.items(), key=lambda kv: kv[1])
    if logger:
        score_str = ", ".join(f"{name} 오차 {s*100:.1f}%" for name, s in ranked)
        logger.info(f"[{series_name}] 백테스트(n={n}, 검증창 {test_h}개 x {len(origins)}회) 결과 → {score_str}")

    for name, _ in ranked:
        fc = methods[name](values, n_ahead)
        if fc is not None:
            fc = np.maximum(np.asarray(fc, dtype=float), 0.0)
            fc, capped = _cap_low_confidence_forecast(fc, values, name, scores, logger, series_name)
            if logger:
                logger.info(f"[{series_name}] 최종 선택 알고리즘: {name}" + (" (상한 적용됨)" if capped else ""))
            return fc, name, scores, capped

    avg = float(np.mean(values[values > 0])) if nz_count > 0 else 0.0
    return np.full(n_ahead, avg), "표본부족(평균유지)", scores, False


# =====================================================================
# 집계 & 예측 실행 (공정 × 메이커(설비) × 모델 단위)
# =====================================================================
COMBO_COLS = ["공정", "메이커", "모델"]


def period_pivot(df, group_cols, value_col, periods):
    idx_p = pd.CategoricalDtype(periods, ordered=True)
    tmp = df.copy()
    tmp["기간"] = tmp["기간"].astype(idx_p)
    piv = tmp.pivot_table(index="기간", columns=group_cols, values=value_col,
                           aggfunc="sum", fill_value=0, observed=False)
    piv = piv.reindex(periods, fill_value=0)
    return piv


def run_forecast(df, gran, train_start, train_end, forecast_start, forecast_end,
                  exclude_processes=("미확인",), logger=None):
    season = GRANULARITY_META[gran]["season"]
    window = GRANULARITY_META[gran]["window"]

    train_periods = period_range(train_start, train_end, gran)
    fc_periods = period_range(forecast_start, forecast_end, gran)
    n_ahead = len(fc_periods)

    combos = sorted(df.groupby(COMBO_COLS).size().index.tolist())
    processes = sorted(df["공정"].unique().tolist())

    amt_piv = period_pivot(df, COMBO_COLS, "매출액(KRW)", train_periods)
    qty_piv = period_pivot(df, COMBO_COLS, "수량", train_periods)

    if logger:
        logger.info(f"공정×메이커×모델 조합 {len(combos)}개에 대해 조합별로 예측을 계산합니다. (기간단위: {gran})")

    results = []
    for (proc, maker, model) in combos:
        key = (proc, maker, model)
        amt_hist = amt_piv[key].values if key in amt_piv.columns else np.zeros(len(train_periods))
        qty_hist = qty_piv[key].values if key in qty_piv.columns else np.zeros(len(train_periods))
        series_name = f"{proc}/{maker}/{model}"

        excluded = proc in exclude_processes
        if excluded:
            amt_fc = np.zeros(n_ahead)
            qty_fc = np.zeros(n_ahead)
            method = "제외(돌발성 매출)"
            scores = {}
            capped = False
        else:
            amt_fc, method, scores, capped = forecast_series(amt_hist, n_ahead, season=season, window=window,
                                                               logger=logger, series_name=f"{series_name} [금액]")
            qty_fc, _, _, _ = forecast_series(qty_hist, n_ahead, season=season, window=window,
                                               logger=logger, series_name=f"{series_name} [수량]")

        baseline_len = min(n_ahead, len(amt_hist))
        baseline_amt = float(np.sum(amt_hist[-baseline_len:])) if baseline_len else 0.0

        results.append({
            "공정": proc, "메이커": maker, "모델": model, "excluded": excluded, "method": method,
            "backtest_scores": scores, "capped": capped,
            "amt_hist": amt_hist, "qty_hist": qty_hist,
            "amt_fc": amt_fc, "qty_fc": qty_fc,
            "amt_fc_total": float(np.sum(amt_fc)), "qty_fc_total": float(np.sum(qty_fc)),
            "baseline_amt": baseline_amt,
            "diff_amt": float(np.sum(amt_fc)) - baseline_amt,
            "nz_count": int(np.count_nonzero(amt_hist)),
            "txn_count": int(((df["공정"] == proc) & (df["메이커"] == maker) & (df["모델"] == model) &
                               (df["기간"].isin(train_periods))).sum()),
        })

    detail = pd.DataFrame(results)

    # 전사/공정 top-down (계절성 반영) — bottom-up과 교차검증용
    total_hist = amt_piv.sum(axis=1).values
    total_fc_topdown, total_method_topdown, _, _ = forecast_series(
        total_hist, n_ahead, season=season, window=window, logger=logger, series_name="전사 합계(전사통합)")

    proc_topdown = {}
    for proc in processes:
        if proc in exclude_processes:
            continue
        cols = [c for c in amt_piv.columns if c[0] == proc]
        if not cols:
            proc_topdown[proc] = (np.zeros(n_ahead), "데이터없음", {})
            continue
        series = amt_piv[cols].sum(axis=1).values
        fc, m, sc, _ = forecast_series(series, n_ahead, season=season, window=window,
                                        logger=logger, series_name=f"공정 합계(전사통합): {proc}")
        proc_topdown[proc] = (fc, m, sc)

    return {
        "gran": gran, "season": season, "window": window,
        "train_periods": train_periods,
        "fc_periods": fc_periods,
        "detail": detail,
        "total_hist": total_hist,
        "total_fc_topdown": total_fc_topdown,
        "total_method_topdown": total_method_topdown,
        "proc_topdown": proc_topdown,
        "amt_piv": amt_piv,
    }


def dimension_summary(detail, dim_col):
    """예측이 제외되지 않은 조합들을 지정한 차원(공정/메이커/모델) 기준으로 다시 합산한다."""
    active = detail[~detail["excluded"]]
    if not len(active):
        return pd.DataFrame(columns=[dim_col, "직전동기간실적", "예측합계", "증감액", "증감률"])
    g = active.groupby(dim_col, observed=True).agg(
        직전동기간실적=("baseline_amt", "sum"),
        예측합계=("amt_fc_total", "sum"),
    ).reset_index()
    g["증감액"] = g["예측합계"] - g["직전동기간실적"]
    g["증감률"] = np.where(g["직전동기간실적"] != 0, g["증감액"] / g["직전동기간실적"], np.nan)
    return g.sort_values("예측합계", ascending=False).reset_index(drop=True)


# =====================================================================
# 금액 표기 (억원 단위 축약)
# =====================================================================
def fmt_eok(v, decimals=2):
    """68,535,079,090 -> '685.35억원' 형태로 축약 표기한다."""
    return f"{v/1e8:,.{decimals}f}억원"


def fmt_eok_num(v, decimals=2):
    """단위(원) 없이 억 단위 숫자만: 68,535,079,090 -> '685.35'"""
    return f"{v/1e8:,.{decimals}f}"


# =====================================================================
# 차트 (그래프는 값이 겹치지 않도록 단순하게 유지하고, 정확한 수치는 아래 표로 제공)
# =====================================================================
def _fig_to_base64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=140, bbox_inches="tight")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def chart_total_trend(train_periods, total_hist, fc_periods, total_fc_bottomup, total_fc_topdown, unit):
    fig, ax = plt.subplots(figsize=(9, 4))
    ax.plot(range(len(train_periods)), total_hist, color=BRAND_GRAY, marker="o", markersize=3,
            linewidth=1.6, label="실적")
    n_hist = len(train_periods)
    fc_x = list(range(n_hist, n_hist + len(fc_periods)))
    ax.plot(fc_x, total_fc_bottomup, color=BRAND_MAGENTA, marker="D", markersize=5,
            linewidth=1.8, label="예측 · 조합별 합산(기본)")
    ax.plot(fc_x, total_fc_topdown, color="#1f77b4", marker="s", markersize=5,
            linewidth=1.4, linestyle="--", label="예측 · 전사 통합(참고)")
    # connect last actual to forecast start
    ax.plot([n_hist - 1, n_hist], [total_hist[-1], total_fc_bottomup[0]], color=BRAND_MAGENTA, linewidth=1, alpha=0.5)
    ax.plot([n_hist - 1, n_hist], [total_hist[-1], total_fc_topdown[0]], color="#1f77b4", linewidth=1, alpha=0.5, linestyle="--")

    # 기간이 많아 그래프에 값을 전부 표기하면 겹쳐서 읽기 어려우므로, 데이터 레이블 대신
    # 아래 상세 데이터표(모든 기간의 실적/예측 값)로 정확한 수치를 제공한다.
    all_labels = list(train_periods) + list(fc_periods)
    step = max(1, len(all_labels) // 16)
    ax.set_xticks(range(0, len(all_labels), step))
    ax.set_xticklabels([all_labels[i] for i in range(0, len(all_labels), step)], rotation=45, ha="right", fontsize=8)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, p: f"{v/1e8:,.0f}억"))
    ax.set_title(f"전사 {unit}별 매출 추이 및 예측", fontsize=12, color=BRAND_MAGENTA, fontweight="bold")
    ax.legend(fontsize=8, loc="upper left")
    ax.grid(alpha=0.25)
    fig.tight_layout()
    return _fig_to_base64(fig)


def chart_dimension_bar(summary, dim_col, title, top_n=15):
    s = summary.copy()
    if len(s) > top_n:
        head = s.iloc[:top_n]
        rest = s.iloc[top_n:]
        etc = pd.DataFrame([{
            dim_col: f"기타 {len(rest)}건",
            "직전동기간실적": rest["직전동기간실적"].sum(),
            "예측합계": rest["예측합계"].sum(),
        }])
        s = pd.concat([head, etc], ignore_index=True)

    fig, ax = plt.subplots(figsize=(9, 4.5))
    labels = s[dim_col].astype(str).tolist()
    x = np.arange(len(labels))
    w = 0.38
    h1 = (s["직전동기간실적"] / 1e8).tolist()
    h2 = (s["예측합계"] / 1e8).tolist()
    ax.bar(x - w / 2, h1, width=w, color=BRAND_GRAY, label="직전 동기간 실적")
    ax.bar(x + w / 2, h2, width=w, color=BRAND_MAGENTA, label="예측")
    # 데이터 레이블 대신 아래 표에서 정확한 수치를 확인할 수 있다.
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=30, ha="right", fontsize=9)
    ax.set_ylabel("억원")
    ax.set_title(title, fontsize=12, color=BRAND_MAGENTA, fontweight="bold")
    ax.legend(fontsize=9)
    ax.grid(alpha=0.25, axis="y")
    fig.tight_layout()
    return _fig_to_base64(fig)


def chart_top_contributors(detail, top_n=10):
    d = detail[~detail["excluded"]].copy()
    d = d.reindex(d["diff_amt"].abs().sort_values(ascending=False).index).head(top_n)
    d = d.sort_values("diff_amt")
    labels = [f"{p}-{k}-{m}" for p, k, m in zip(d["공정"], d["메이커"], d["모델"])]
    values = (d["diff_amt"] / 1e8).tolist()
    colors = [BRAND_MAGENTA if v >= 0 else "#4472C4" for v in values]
    fig, ax = plt.subplots(figsize=(9, 5))
    ax.barh(labels, values, color=colors)
    # 데이터 레이블 대신 아래 표에서 정확한 수치를 확인할 수 있다.
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
.tablewrap {{ overflow-x:auto; }}
.scrolltable {{ overflow:auto; max-height:420px; border:1px solid #eee; border-radius:6px; }}
.scrolltable table {{ margin:0; }}
.scrolltable th {{ position:sticky; top:0; z-index:1; }}
table {{ border-collapse: collapse; width:100%; margin:14px 0; font-size:13px;}}
th {{ background:{magenta}; color:white; padding:8px 10px; text-align:center; white-space:nowrap;}}
td {{ padding:7px 10px; border-bottom:1px solid #eee; text-align:right; white-space:nowrap;}}
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
.capbadge {{ display:inline-block; font-size:11px; padding:2px 8px; border-radius:10px; background:#FDECEA; color:#B3261E; margin-left:4px;}}
.conf-high {{ color:#1F7A1F; font-weight:bold; }}
.conf-mid {{ color:#B8860B; font-weight:bold; }}
.conf-low {{ color:#C00000; font-weight:bold; }}
code {{ background:#f0f0f0; padding:1px 5px; border-radius:4px; }}
</style></head>
<body><div class="wrap">
<h1>매출 예측 리포트</h1>
<div class="meta">학습기간: {train_start} ~ {train_end} &nbsp;|&nbsp; 예측기간: {fc_start} ~ {fc_end} &nbsp;|&nbsp; 기간단위: {unit} &nbsp;|&nbsp; 생성 파일: {src_name} ({sheet_name} 시트)</div>
{exclude_note}

<div class="card">
<h2 style="margin-top:0;border:none;">이 리포트 읽는 법</h2>
<div style="font-size:13px; line-height:1.8; color:#333;">
{guide_section}
</div>
</div>

<div class="summary-box">
  <div class="stat"><div class="label">직전 동기간 실적 (전사, 미확인 제외)</div><div class="value">{total_base_fmt}</div></div>
  <div class="stat"><div class="label">조합별 합산 예측 (기본)</div><div class="value">{total_fc_fmt}</div></div>
  <div class="stat"><div class="label">증감률</div><div class="value">{total_pct_fmt}</div></div>
  <div class="stat"><div class="label">전사 통합 예측 (참고)</div><div class="value">{total_topdown_fmt}</div><div class="methodbadge">{topdown_method}</div></div>
</div>

<div class="card"><h2 style="margin-top:0;border:none;">전사 매출 추이 및 예측</h2>
<img src="data:image/png;base64,{chart_total}"/>
<p style="font-size:13px;color:{gray}">실선(회색)은 실적, 마름모(마젠타)는 조합별로 예측해 합산한 <b>조합별 합산 예측(기본)</b> 결과, 파란 점선은 전사 데이터 전체에 계절성 모델을 적용한 <b>전사 통합 예측(참고)</b> 결과입니다. 두 방식이 크게 어긋나면 개별 조합의 이상치를 의심해볼 수 있습니다. 어느 한쪽이 항상 더 크거나 작다는 규칙은 없으며, 정확한 수치는 그래프 대신 아래 표에서 확인합니다.{annual_note}</p>
{trend_table}
</div>

<div class="card"><h2 style="margin-top:0;border:none;">공정별 실적 대비 예측</h2>
<img src="data:image/png;base64,{chart_proc}"/>
<div class="tablewrap">{proc_table}</div>
</div>

<div class="card"><h2 style="margin-top:0;border:none;">메이커(설비)별 실적 대비 예측</h2>
<img src="data:image/png;base64,{chart_maker}"/>
<div class="tablewrap">{maker_table}</div>
</div>

<div class="card"><h2 style="margin-top:0;border:none;">모델별 실적 대비 예측</h2>
<img src="data:image/png;base64,{chart_model}"/>
<div class="tablewrap">{model_table}</div>
</div>

<div class="card"><h2 style="margin-top:0;border:none;">증감 기여도 상위 조합</h2>
<img src="data:image/png;base64,{chart_top}"/>
<p style="font-size:13px;color:{gray}">{narrative}</p>
<div class="tablewrap">{contributors_table}</div>
</div>

<h2>공정×메이커×모델 상세 예측 (전체 {n_combo}개 조합)</h2>
<div class="tablewrap">{detail_table}</div>

<h2>예측 방법론 (다중 알고리즘 백테스트)</h2>
<div class="card" style="font-size:13px; line-height:1.7;">
이 리포트는 각 조합마다 아래 알고리즘들을 모두 후보로 놓고, 과거 데이터로 백테스트(학습/검증 분리 검증)를 수행해 오차(sMAPE)가 가장 작은 알고리즘을 자동으로 선택합니다. 검증할 데이터가 너무 짧은 조합은 표본 크기에 맞는 보수적인 방법으로 대체합니다.<br><br>
{method_section}
<br>모든 예측치는 0 미만이 되지 않도록 하한을 적용했습니다.
</div>

</div></body></html>"""


def _pct_cell_html(v):
    if pd.isna(v):
        return '<span class="methodbadge">N/A</span>'
    cls = "pos" if v >= 0 else "neg"
    sign = "+" if v >= 0 else ""
    return f'<span class="{cls}">{sign}{v*100:.1f}%</span>'


def _guide_html(items, log_name):
    """reader_guide_items()의 내용을 HTML로 렌더링한다 (예측신뢰도 항목만 색상 강조)."""
    out = []
    for title, text in items:
        if title == "예측신뢰도":
            text = (text.replace("10% 이하 높음", '10% 이하 <span class="conf-high">높음</span>')
                        .replace("10~25% 보통", '10~25% <span class="conf-mid">보통</span>')
                        .replace("25% 초과 낮음(참고용)", '25% 초과 <span class="conf-low">낮음(참고용)</span>'))
        out.append(f"<b>{title}</b>: {text}<br>")
    out.append(f"실행 과정 전체와 조합별 백테스트 점수 등 모든 상세 로그는 함께 생성된 텍스트 파일"
               f"(<code>{log_name}</code>)에서 확인할 수 있습니다(메모장으로 바로 열립니다).")
    return "".join(out)


def _dimension_table_html(summary, dim_col, dim_label):
    rows = "\n".join(
        f"<tr><td>{r[dim_col]}</td><td>{fmt_eok(r['직전동기간실적'])}</td><td>{fmt_eok(r['예측합계'])}</td>"
        f"<td>{fmt_eok(r['증감액'])}</td><td>{_pct_cell_html(r['증감률'])}</td></tr>"
        for _, r in summary.iterrows()
    )
    return (f"<table><tr><th>{dim_label}</th><th>직전동기간실적</th><th>예측합계</th><th>증감액</th><th>증감률</th></tr>"
            f"{rows}</table>")


def _trend_table_html(train_periods, total_hist, fc_periods, total_fc_bottomup, total_fc_topdown, unit):
    """전사 추이 그래프는 기간이 많아 그래프 위에 값을 전부 표기하기 어려우므로,
    그래프 바로 아래에 과거 실적을 포함한 기간별 상세 수치표를 항상 표시한다."""
    rows = []
    for period, v in zip(train_periods, total_hist):
        rows.append(f"<tr><td>{period}</td><td>{fmt_eok(v)}</td><td>-</td><td>-</td></tr>")
    for period, bu, td in zip(fc_periods, total_fc_bottomup, total_fc_topdown):
        rows.append(f"<tr><td>{period}</td><td>-</td><td>{fmt_eok(bu)}</td><td>{fmt_eok(td)}</td></tr>")
    n_total = len(train_periods) + len(fc_periods)
    table = (f"<table><tr><th>{unit}</th><th>실적</th><th>조합별 합산 예측(기본)</th><th>전사 통합 예측(참고)</th></tr>"
             f"{''.join(rows)}</table>")
    return (f'<div style="font-size:12px;color:#888;margin:10px 0 4px;">{unit}별 상세 데이터표 (전체 {n_total}개 구간)</div>'
            f'<div class="scrolltable">{table}</div>')


def _contributors_table_html(detail, top_n=10):
    """증감 기여도 상위 조합 그래프에 대응하는 상세 수치표."""
    d = detail[~detail["excluded"]].copy()
    d = d.reindex(d["diff_amt"].abs().sort_values(ascending=False).index).head(top_n)
    rows = "\n".join(
        f"<tr><td>{r['공정']}</td><td>{r['메이커']}</td><td style='text-align:left'>{r['모델']}</td>"
        f"<td>{fmt_eok(r['baseline_amt'])}</td><td>{fmt_eok(r['amt_fc_total'])}</td>"
        f"<td class=\"{'pos' if r['diff_amt']>=0 else 'neg'}\">{'+' if r['diff_amt']>=0 else ''}{fmt_eok(r['diff_amt'])}</td></tr>"
        for _, r in d.iterrows()
    )
    return (f"<table><tr><th>공정</th><th>메이커(설비)</th><th>모델</th><th>직전동기간실적</th>"
            f"<th>예측합계</th><th>증감액</th></tr>{rows}</table>")


def build_html_report(df, result, src_name, sheet_name, exclude_processes, log_path):
    detail = result["detail"]
    train_periods, fc_periods = result["train_periods"], result["fc_periods"]
    gran, window = result["gran"], result["window"]
    n_ahead = len(fc_periods)

    active = detail[~detail["excluded"]]
    total_base = active["baseline_amt"].sum()
    total_fc = active["amt_fc_total"].sum()
    total_pct = (total_fc - total_base) / total_base if total_base else None
    total_topdown = float(np.sum(result["total_fc_topdown"]))

    proc_summary = dimension_summary(detail, "공정")
    maker_summary = dimension_summary(detail, "메이커")
    model_summary = dimension_summary(detail, "모델")

    total_fc_bottomup_arr = np.sum(np.stack(active["amt_fc"].values), axis=0) if len(active) else np.zeros(n_ahead)

    annual_only_years = detect_annual_only_years(df, gran)
    display_train_periods, display_total_hist = collapse_annual_periods(
        train_periods, list(result["total_hist"]), gran, annual_only_years)

    chart_total = chart_total_trend(display_train_periods, display_total_hist, fc_periods,
                                     total_fc_bottomup_arr, result["total_fc_topdown"], gran)
    trend_table = _trend_table_html(display_train_periods, display_total_hist, fc_periods,
                                     total_fc_bottomup_arr, result["total_fc_topdown"], gran)
    chart_proc = chart_dimension_bar(proc_summary, "공정", "공정별 실적 대비 예측")
    chart_maker = chart_dimension_bar(maker_summary, "메이커", "메이커(설비)별 실적 대비 예측")
    chart_model = chart_dimension_bar(model_summary, "모델", "모델별 실적 대비 예측")
    chart_top = chart_top_contributors(detail)

    proc_table = _dimension_table_html(proc_summary, "공정", "공정")
    maker_table = _dimension_table_html(maker_summary, "메이커", "메이커(설비)")
    model_table = _dimension_table_html(model_summary, "모델", "모델")
    contributors_table = _contributors_table_html(detail)

    top3 = active.reindex(active["diff_amt"].abs().sort_values(ascending=False).index).head(3)
    conf_counts = active.apply(lambda r: confidence_label(r["backtest_scores"], r["method"]), axis=1).value_counts() \
        if len(active) else pd.Series(dtype=int)
    capped_count = int(active["capped"].sum()) if len(active) else 0
    conf_txt = (f" 조합별 예측 신뢰도는 높음 {int(conf_counts.get('높음', 0))}건, "
                f"보통 {int(conf_counts.get('보통', 0))}건, "
                f"낮음(참고용) {int(conf_counts.get('낮음(참고용)', 0))}건, "
                f"산정불가(N/A) {int(conf_counts.get('N/A', 0))}건입니다." +
                (f" 이 중 {capped_count}건은 신뢰도가 낮아 과거 최대 실적 대비 예측치가 과도하게 튀지 않도록 "
                 f"상한을 적용했습니다." if capped_count else ""))

    if total_pct is not None and len(top3):
        narrative = (f"설정한 예측기간({fc_periods[0]}~{fc_periods[-1]}) 전사 합계 예측은 {fmt_eok(total_fc)}로, "
                     f"직전 동일 길이 기간 실적({fmt_eok(total_base)}) 대비 {'+' if total_pct>=0 else ''}{total_pct*100:.1f}% "
                     f"{'증가' if total_pct>=0 else '감소'}입니다. 가장 큰 요인은 " +
                     ", ".join(f"{r['공정']}-{r['메이커']}-{r['모델']}({'+' if r['diff_amt']>=0 else ''}{fmt_eok(r['diff_amt'])})"
                               for _, r in top3.iterrows()) + " 입니다." + conf_txt)
    else:
        narrative = conf_txt

    def conf_span(scores, method):
        label = confidence_label(scores, method)
        cls = {"높음": "conf-high", "보통": "conf-mid", "낮음(참고용)": "conf-low"}.get(label, "")
        return f'<span class="{cls}">{label}</span>' if cls else label

    cap_badge = ' <span class="capbadge">상한적용</span>'

    det_sorted = sort_detail_table(detail)
    detail_rows = []
    for _, r in det_sorted.iterrows():
        detail_rows.append(
            f"<tr><td>{r['공정']}</td><td>{r['메이커']}</td><td style='text-align:left'>{r['모델']}</td>"
            f"<td>{fmt_eok(r['baseline_amt'])}</td>"
            f"<td>{fmt_eok(r['amt_fc_total'])}</td><td>{r['qty_fc_total']:.1f}</td>"
            f"<td><span class='methodbadge'>{r['method']}</span>"
            f"{cap_badge if r['capped'] else ''}</td>"
            f"<td>{conf_span(r['backtest_scores'], r['method'])}</td></tr>")
    detail_table = (f"<table><tr><th>공정</th><th>메이커(설비)</th><th>모델</th><th>직전동기간실적(금액)</th>"
                     f"<th>예측합계(금액)</th><th>예측합계(수량)</th><th>적용알고리즘</th><th>예측신뢰도</th></tr>"
                     f"{''.join(detail_rows)}</table>")

    exclude_note = ""
    if exclude_processes:
        exclude_note = (f'<div class="note">공정이 {", ".join(exclude_processes)} 인 매출은 돌발성(일회성) '
                         f"매출로 간주되어 예측 및 위 합계 지표에서 제외되었습니다. (과거 실적은 상세표에서 참고용으로 확인 가능)</div>")

    method_counts = detail[~detail["excluded"]]["method"].value_counts()
    method_lines = []
    for name, cnt in method_counts.items():
        desc = method_description(name, gran, window)
        method_lines.append(f"<b>{name}</b> ({int(cnt)}개 조합): {desc}")
    excl_count = int(detail["excluded"].sum())
    if excl_count:
        method_lines.append(f"<b>제외(돌발성 매출)</b> ({excl_count}개 조합): {METHOD_DESCRIPTIONS['제외(돌발성 매출)']}")
    method_section = "<br>".join(method_lines)
    method_count = len(build_candidate_methods(result["season"], window))
    guide_section = _guide_html(reader_guide_items(n_ahead, gran, method_count), log_path.name)

    annual_note = ""
    if annual_only_years:
        years_txt = ", ".join(str(y) for y in sorted(annual_only_years))
        annual_note = (f" 다만 {years_txt}년은 실제 데이터가 연간 합계 형태로만 존재해 "
                        f"선택한 기간단위({gran}) 대신 <b>년</b> 단위로 요약해 표시했습니다.")

    html = HTML_TEMPLATE.format(
        magenta=BRAND_MAGENTA, gray=BRAND_GRAY,
        train_start=train_periods[0], train_end=train_periods[-1],
        fc_start=fc_periods[0], fc_end=fc_periods[-1], src_name=src_name, sheet_name=sheet_name,
        unit=gran, n_ahead=n_ahead, method_count=method_count, guide_section=guide_section,
        annual_note=annual_note,
        exclude_note=exclude_note, log_name=log_path.name,
        total_base_fmt=fmt_eok(total_base), total_fc_fmt=fmt_eok(total_fc),
        total_pct_fmt=(f"{'+' if total_pct>=0 else ''}{total_pct*100:.1f}%" if total_pct is not None else "N/A"),
        total_topdown_fmt=fmt_eok(total_topdown), topdown_method=result["total_method_topdown"],
        chart_total=chart_total, chart_proc=chart_proc, chart_model=chart_model, chart_maker=chart_maker,
        chart_top=chart_top, trend_table=trend_table,
        proc_table=proc_table, model_table=model_table, maker_table=maker_table,
        contributors_table=contributors_table,
        narrative=narrative,
        n_combo=len(detail), detail_table=detail_table,
        method_section=method_section,
    )

    ctx = {
        "proc_summary": proc_summary, "model_summary": model_summary, "maker_summary": maker_summary,
        "total_base": total_base, "total_fc": total_fc, "total_pct": total_pct, "total_topdown": total_topdown,
        "chart_total": chart_total, "chart_proc": chart_proc, "chart_model": chart_model,
        "chart_maker": chart_maker, "chart_top": chart_top,
        "fc_bottomup_arr": total_fc_bottomup_arr, "narrative": narrative,
        "display_train_periods": display_train_periods, "display_total_hist": display_total_hist,
        "annual_only_years": annual_only_years,
    }
    return html, ctx


# =====================================================================
# PDF 리포트 (HTML 리포트와 동일한 내용을 동일한 순서로 담는다)
# =====================================================================
def build_pdf_report(outpath, df, result, src_name, sheet_name, exclude_processes, ctx, log_path):
    from fpdf import FPDF
    from fpdf.fonts import FontFace

    train_periods, fc_periods = result["train_periods"], result["fc_periods"]
    gran, window = result["gran"], result["window"]
    detail = result["detail"]
    total_base, total_fc = ctx["total_base"], ctx["total_fc"]
    total_pct, total_topdown = ctx["total_pct"], ctx["total_topdown"]
    n_ahead = len(fc_periods)
    HEADER_FILL = (200, 0, 124)

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

    def section_title(text):
        pdf.set_font(pdf.font_family, "", 13)
        pdf.set_text_color(*HEADER_FILL)
        pdf.set_x(pdf.l_margin)
        pdf.cell(0, 10, text, ln=True)

    def body_text(text, size=9, color=(60, 60, 60)):
        pdf.set_font(pdf.font_family, "", size)
        pdf.set_text_color(*color)
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(0, 6, text, align="L")

    def draw_table(rows, col_widths, aligns):
        pdf.set_font(pdf.font_family, "", 8)
        pdf.set_text_color(30, 30, 30)
        with pdf.table(rows, text_align=aligns, col_widths=col_widths,
                       headings_style=FontFace(family=pdf.font_family, color=255, fill_color=HEADER_FILL)):
            pass
        pdf.ln(2)

    # ---- 제목 ----
    pdf.set_text_color(*HEADER_FILL)
    pdf.set_font(pdf.font_family, "", 18)
    pdf.cell(0, 12, "매출 예측 리포트", ln=True)
    pdf.set_font(pdf.font_family, "", 10)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 8, f"학습기간 {train_periods[0]}~{train_periods[-1]}  |  "
                    f"예측기간 {fc_periods[0]}~{fc_periods[-1]} ({gran})  |  원본: {src_name} ({sheet_name} 시트)", ln=True)
    pdf.ln(2)

    if exclude_processes:
        body_text(f"※ 공정이 {', '.join(exclude_processes)}인 매출은 돌발성(일회성)으로 간주해 예측/합계에서 제외했습니다.",
                   size=9, color=(192, 0, 0))
        pdf.ln(1)

    pct_txt = f"{'+' if total_pct is not None and total_pct>=0 else ''}{total_pct*100:.1f}%" if total_pct is not None else "N/A"
    body_text(f"전사 직전동기간 실적: {fmt_eok(total_base)}\n"
              f"전사 조합별 합산 예측(기본): {fmt_eok(total_fc)}  ({pct_txt})\n"
              f"전사 통합 예측(참고): {fmt_eok(total_topdown)}", size=11, color=(30, 30, 30))
    pdf.ln(3)

    # ---- 이 리포트 읽는 법 (HTML과 동일한 문구를 공유) ----
    pdf.add_page()
    section_title("이 리포트 읽는 법")
    method_count = len(build_candidate_methods(result["season"], window))
    for title, text in reader_guide_items(n_ahead, gran, method_count):
        pdf.set_font(pdf.font_family, "", 10)
        pdf.set_text_color(*HEADER_FILL)
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(0, 6, title)
        body_text(text, size=9, color=(60, 60, 60))
        pdf.ln(1)
    body_text(f"실행 과정 전체와 조합별 백테스트 점수 등 모든 상세 로그는 함께 생성된 텍스트 파일 "
              f"({log_path.name})에서 확인할 수 있습니다(메모장으로 바로 열립니다).", size=9, color=(120, 120, 120))

    # ---- 전사 매출 추이 및 예측 (과거 실적을 포함한 전체 기간 표) ----
    pdf.add_page()
    section_title("전사 매출 추이 및 예측")
    img1 = io.BytesIO(base64.b64decode(ctx["chart_total"]))
    pdf.image(img1, w=190)
    pdf.ln(2)
    annual_years = ctx.get("annual_only_years") or set()
    annual_note = ""
    if annual_years:
        years_txt = ", ".join(str(y) for y in sorted(annual_years))
        annual_note = (f" 다만 {years_txt}년은 실제 데이터가 연간 합계 형태로만 존재해 "
                        f"선택한 기간단위({gran}) 대신 년 단위로 요약해 표시했습니다.")
    body_text("실선(회색)은 실적, 마름모(마젠타)는 조합별로 예측해 합산한 조합별 합산 예측(기본) 결과, 파란 점선은 전사 "
              "데이터 전체에 계절성 모델을 적용한 전사 통합 예측(참고) 결과입니다. 두 방식이 크게 어긋나면 개별 조합의 "
              "이상치를 의심해볼 수 있습니다. 어느 한쪽이 항상 더 크거나 작다는 규칙은 없으며, 정확한 수치는 아래 표에서 "
              "확인합니다." + annual_note)
    pdf.ln(1)
    trend_rows = [[gran, "실적", "조합별 합산 예측(기본)", "전사 통합 예측(참고)"]]
    for period, v in zip(ctx["display_train_periods"], ctx["display_total_hist"]):
        trend_rows.append([str(period), fmt_eok(v), "-", "-"])
    for period, bu, td in zip(fc_periods, ctx["fc_bottomup_arr"], result["total_fc_topdown"]):
        trend_rows.append([str(period), "-", fmt_eok(bu), fmt_eok(td)])
    draw_table(trend_rows, col_widths=(30, 53, 53, 54), aligns=("LEFT", "RIGHT", "RIGHT", "RIGHT"))

    # ---- 공정별 -> 메이커(설비)별 -> 모델별 실적 대비 예측 ----
    for title, summary, dim_col, chart_key, header_label in [
        ("공정별 실적 대비 예측", ctx["proc_summary"], "공정", "chart_proc", "공정"),
        ("메이커(설비)별 실적 대비 예측", ctx["maker_summary"], "메이커", "chart_maker", "메이커(설비)"),
        ("모델별 실적 대비 예측", ctx["model_summary"], "모델", "chart_model", "모델"),
    ]:
        pdf.add_page()
        section_title(title)
        img = io.BytesIO(base64.b64decode(ctx[chart_key]))
        pdf.image(img, w=190)
        pdf.ln(2)
        rows = [[header_label, "직전동기간실적", "예측합계", "증감액", "증감률"]]
        for _, r in summary.iterrows():
            pct = r["증감률"]
            pct_s = f"{'+' if pd.notna(pct) and pct>=0 else ''}{pct*100:.1f}%" if pd.notna(pct) else "N/A"
            rows.append([str(r[dim_col]), fmt_eok(r["직전동기간실적"]), fmt_eok(r["예측합계"]),
                         fmt_eok(r["증감액"]), pct_s])
        draw_table(rows, col_widths=(40, 38, 38, 38, 36), aligns=("LEFT", "RIGHT", "RIGHT", "RIGHT", "RIGHT"))

    # ---- 증감 기여도 상위 조합 ----
    pdf.add_page()
    section_title("증감 기여도 상위 조합")
    img3 = io.BytesIO(base64.b64decode(ctx["chart_top"]))
    pdf.image(img3, w=190)
    pdf.ln(2)
    body_text(ctx["narrative"])
    pdf.ln(1)
    top_detail = detail[~detail["excluded"]].copy()
    top_detail = top_detail.reindex(top_detail["diff_amt"].abs().sort_values(ascending=False).index).head(10)
    rows = [["공정", "메이커(설비)", "모델", "직전동기간실적", "예측합계", "증감액"]]
    for _, r in top_detail.iterrows():
        sign = "+" if r["diff_amt"] >= 0 else ""
        rows.append([str(r["공정"]), str(r["메이커"]), str(r["모델"]), fmt_eok(r["baseline_amt"]),
                     fmt_eok(r["amt_fc_total"]), f"{sign}{fmt_eok(r['diff_amt'])}"])
    draw_table(rows, col_widths=(28, 28, 28, 35, 35, 36),
               aligns=("LEFT", "LEFT", "LEFT", "RIGHT", "RIGHT", "RIGHT"))

    # ---- 공정×메이커×모델 상세 예측 (전체 조합, HTML과 동일하게 생략 없이 전부 표시) ----
    pdf.add_page()
    section_title(f"공정×메이커×모델 상세 예측 (전체 {len(detail)}개 조합)")
    det_sorted = sort_detail_table(detail)
    rows = [["공정", "메이커", "모델", "직전실적(금액)", "예측합계(금액)", "예측합계(수량)", "적용알고리즘", "신뢰도"]]
    for _, r in det_sorted.iterrows():
        conf = confidence_label(r["backtest_scores"], r["method"])
        method_s = r["method"] + (" (상한적용)" if r["capped"] else "")
        rows.append([str(r["공정"]), str(r["메이커"]), str(r["모델"]), fmt_eok(r["baseline_amt"]),
                     fmt_eok(r["amt_fc_total"]), f"{r['qty_fc_total']:.1f}", method_s, conf])
    draw_table(rows, col_widths=(20, 20, 20, 26, 26, 18, 44, 16),
               aligns=("LEFT", "LEFT", "LEFT", "RIGHT", "RIGHT", "RIGHT", "LEFT", "CENTER"))

    # ---- 예측 방법론 (다중 알고리즘 백테스트) ----
    pdf.add_page()
    section_title("예측 방법론 (다중 알고리즘 백테스트)")
    body_text("이 리포트는 각 조합마다 아래 알고리즘들을 모두 후보로 놓고, 과거 데이터로 백테스트(학습/검증 분리 검증)를 "
              "수행해 오차(sMAPE)가 가장 작은 알고리즘을 자동으로 선택합니다. 검증할 데이터가 너무 짧은 조합은 표본 "
              "크기에 맞는 보수적인 방법으로 대체합니다.")
    pdf.ln(1)
    method_counts = detail[~detail["excluded"]]["method"].value_counts()
    for name, cnt in method_counts.items():
        body_text(f"- {name} ({int(cnt)}개 조합): {method_description(name, gran, window)}")
    excl_count = int(detail["excluded"].sum())
    if excl_count:
        body_text(f"- 제외(돌발성 매출) ({excl_count}개 조합): {METHOD_DESCRIPTIONS['제외(돌발성 매출)']}")
    pdf.ln(1)
    body_text(f"모든 예측치는 0 미만이 되지 않도록 하한을 적용했습니다. 조합별 백테스트 상세 점수는 실행 로그 파일"
              f"({log_path.name})에서 확인할 수 있습니다.")

    pdf.output(str(outpath))


# =====================================================================
# 파이프라인 (CLI/GUI 공용)
# =====================================================================
def _ensure_writable_dir(outdir):
    """outdir에 실제로 쓰기가 가능한지 확인하고, 권한이 없으면(예: 접근이 제한된 공유/네트워크
    드라이브) 사용자 문서 폴더 하위로 대체한다. 반환: (실제 사용할 경로, 안내 메시지 또는 None)"""
    try:
        outdir.mkdir(parents=True, exist_ok=True)
        probe = outdir / ".__write_test__.tmp"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        return outdir, None
    except (PermissionError, OSError):
        fallback = Path.home() / "Documents" / "매출예측_리포트" / outdir.name
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback, f"'{outdir}' 폴더에 쓰기 권한이 없어 결과를 '{fallback}'에 대신 저장합니다."


def _safe_write_file(write_fn, path, logger=None):
    """write_fn(path)로 저장을 시도한다. 파일이 다른 프로그램(PDF/엑셀 뷰어 등)에서 열려있어
    PermissionError가 나면, 같은 폴더에 시각을 붙인 다른 이름으로 대신 저장한다."""
    try:
        write_fn(path)
        return path
    except PermissionError:
        ts = datetime.now().strftime("%H%M%S")
        alt_path = path.parent / f"{path.stem}_{ts}{path.suffix}"
        if logger:
            logger.warning(
                f"'{path.name}' 파일을 저장하지 못했습니다 (다른 프로그램에서 열려있거나 접근 권한이 없는 것으로 보입니다). "
                f"'{alt_path.name}' 이름으로 대신 저장합니다. "
                f"(파일이 열려 있었다면 닫은 뒤 다시 실행하면 원래 이름으로 저장됩니다)"
            )
        write_fn(alt_path)
        return alt_path


def run_pipeline(input_path, sheet, gran, train_start, train_end, forecast_start, forecast_end,
                  exclude_processes, outdir, extra_log_handler=None):
    """RAWDATA 로드부터 HTML/PDF/로그 생성까지 전체 과정을 실행한다.
    CLI 콘솔 모드와 GUI 모드가 이 함수를 공유한다."""
    input_path = Path(input_path)
    outdir, outdir_fallback_note = _ensure_writable_dir(Path(outdir))
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    logger, log_path = setup_logger(outdir, extra_handler=extra_log_handler, ts=ts)
    if outdir_fallback_note:
        logger.warning(outdir_fallback_note)
    logger.info("=" * 70)
    logger.info("매출 예측 리포트 생성을 시작합니다")
    logger.info(f"입력 파일: {input_path}")
    logger.info(f"사용 시트: {sheet}")
    logger.info(f"기간 단위: {gran}")
    logger.info(f"결과 폴더: {outdir}")
    logger.info("=" * 70)
    logger.info(column_guide_text())

    logger.info(f"[1/5] RAWDATA 읽는 중... ({input_path.name} / 시트: {sheet})")
    df = load_rawdata(input_path, sheet, logger=logger)
    df = add_period_column(df, gran)
    logger.info(f"      -> {len(df):,}건, {df['매출일'].min().date()} ~ {df['매출일'].max().date()}")

    logger.info(f"[2/5] 예측 계산 중 (알고리즘별 백테스트 수행)... "
                f"학습 {train_start}~{train_end} / 예측 {forecast_start}~{forecast_end} ({gran} 단위)")
    result = run_forecast(df, gran, train_start, train_end, forecast_start, forecast_end,
                           exclude_processes=tuple(exclude_processes or []), logger=logger)

    logger.info("[3/5] 그래프/리포트 생성 중...")
    html, ctx = build_html_report(df, result, input_path.name, sheet, exclude_processes or [], log_path)
    html_path = _safe_write_file(lambda p: p.write_text(html, encoding="utf-8"),
                                  outdir / f"매출예측_리포트_{ts}.html", logger)
    logger.info(f"      -> {html_path}")

    logger.info("[4/5] PDF 생성 중...")
    pdf_path = _safe_write_file(
        lambda p: build_pdf_report(p, df, result, input_path.name, sheet, exclude_processes or [], ctx, log_path),
        outdir / f"매출예측_리포트_{ts}.pdf", logger)
    logger.info(f"      -> {pdf_path}")

    logger.info("[5/5] 완료!")
    logger.info(f"전사 직전동기간 실적: {fmt_eok(ctx['total_base'])}")
    if ctx["total_pct"] is not None:
        sign = "+" if ctx["total_pct"] >= 0 else ""
        logger.info(f"전사 조합별 합산 예측(기본): {fmt_eok(ctx['total_fc'])} ({sign}{ctx['total_pct']*100:.1f}%)")
    else:
        logger.info(f"전사 조합별 합산 예측(기본): {fmt_eok(ctx['total_fc'])}")
    logger.info(f"전사 통합 예측(참고): {fmt_eok(ctx['total_topdown'])}")
    logger.info(f"실행 로그 저장 위치: {log_path}")

    return {"html": html_path, "pdf": pdf_path, "log": log_path, "outdir": outdir}


# =====================================================================
# GUI
# =====================================================================
def pick_sheet_gui(sheet_names, default_sheet):
    """RAWDATA가 들어있는 시트를 GUI 창에서 선택한다. 필수 컬럼 가이드도 함께 보여준다.
    tkinter/디스플레이가 없거나 취소하면 None 반환."""
    if tk is None:
        return None
    try:
        root = tk.Tk()
        root.title("시트 선택")
        root.attributes("-topmost", True)
        selected = {"sheet": None}

        tk.Label(root, text="RAWDATA가 들어있는 시트를 선택하세요", font=("", 11, "bold"),
                 anchor="w", justify="left").pack(padx=14, pady=(14, 6), anchor="w")

        guide_text = "[필수 컬럼 안내]\n" + "\n".join(f"- {c}: {d}" for c, d in COLUMN_GUIDE.items())
        tk.Label(root, text=guide_text, font=("", 9), fg="#555", justify="left",
                 anchor="w", wraplength=440).pack(padx=14, pady=(0, 10), anchor="w")

        listbox = tk.Listbox(root, height=min(8, len(sheet_names)), exportselection=False)
        for name in sheet_names:
            listbox.insert(tk.END, name)
        default_idx = sheet_names.index(default_sheet) if default_sheet in sheet_names else 0
        listbox.selection_set(default_idx)
        listbox.activate(default_idx)
        listbox.pack(padx=14, pady=(0, 12), fill="x")

        def on_ok():
            sel = listbox.curselection()
            selected["sheet"] = sheet_names[sel[0]] if sel else None
            root.destroy()

        def on_cancel():
            selected["sheet"] = None
            root.destroy()

        btn_frame = tk.Frame(root)
        btn_frame.pack(padx=14, pady=(0, 14), anchor="e")
        tk.Button(btn_frame, text="취소", width=8, command=on_cancel).pack(side="right", padx=(6, 0))
        tk.Button(btn_frame, text="확인", width=8, command=on_ok).pack(side="right")

        root.mainloop()
        return selected["sheet"]
    except Exception:
        return None


def pick_sheet_console(sheet_names, default_sheet):
    """콘솔에서 시트를 선택받는다. 필수 컬럼 가이드를 함께 출력한다."""
    if len(sheet_names) == 1:
        return sheet_names[0]
    print("\n" + column_guide_text())
    print("\n엑셀 시트 목록:")
    for i, name in enumerate(sheet_names, 1):
        marker = " (기본값)" if name == default_sheet else ""
        print(f"  {i}. {name}{marker}")
    raw = input("\n사용할 시트 번호 또는 이름을 입력하세요 (엔터 시 기본값 사용): ").strip()
    if not raw:
        return default_sheet
    if raw.isdigit():
        idx = int(raw) - 1
        if 0 <= idx < len(sheet_names):
            return sheet_names[idx]
    if raw in sheet_names:
        return raw
    print(f"[안내] '{raw}'는 올바른 시트가 아니어서 기본값을 사용합니다: {default_sheet}")
    return default_sheet


def _open_path(p):
    if not p:
        return
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(p))  # noqa: S606 (Windows 전용 API)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(p)])
        else:
            subprocess.Popen(["xdg-open", str(p)])
    except Exception as e:
        if tk is not None:
            messagebox.showerror("오류", f"열 수 없습니다: {e}")


class SalesForecastApp:
    """DOS 콘솔 없이, 파일 선택부터 결과 확인까지 하나의 창에서 진행하는 GUI 애플리케이션."""

    def __init__(self, root):
        self.root = root
        self.root.title("매출 예측 리포트 생성기")
        self.root.geometry("820x820")
        self.msg_queue = queue.Queue()
        self.sheet_names = []
        self.current_gran = None
        self.excl_vars = {}
        self.worker_running = False
        self._result_paths = {}
        self._build_widgets()
        self.root.after(150, self._poll_queue)

    # ---------------- 위젯 구성 ----------------
    def _build_widgets(self):
        pad = {"padx": 10, "pady": 6}

        frm_file = ttk.LabelFrame(self.root, text="1. RAWDATA 엑셀 파일")
        frm_file.pack(fill="x", **pad)
        self.file_var = tk.StringVar()
        ttk.Entry(frm_file, textvariable=self.file_var, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(8, 4), pady=8)
        ttk.Button(frm_file, text="찾아보기...", command=self._on_browse_file).pack(side="left", padx=(0, 8), pady=8)

        frm_sheet = ttk.LabelFrame(self.root, text="2. 시트 및 예측기간 단위")
        frm_sheet.pack(fill="x", **pad)
        ttk.Label(frm_sheet, text="시트:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(frm_sheet, textvariable=self.sheet_var, state="readonly", width=22)
        self.sheet_combo.grid(row=0, column=1, sticky="w", padx=4, pady=6)

        ttk.Label(frm_sheet, text="기간단위:").grid(row=0, column=2, sticky="w", padx=(16, 4), pady=6)
        self.gran_var = tk.StringVar(value="분기")
        for i, g in enumerate(GRANULARITIES):
            ttk.Radiobutton(frm_sheet, text=g, variable=self.gran_var, value=g).grid(
                row=0, column=3 + i, sticky="w")

        ttk.Button(frm_sheet, text="데이터 불러오기 / 기간 확인", command=self._on_load_periods).grid(
            row=1, column=0, columnspan=6, sticky="we", padx=8, pady=(4, 8))
        self.data_info_var = tk.StringVar(value="파일과 시트를 선택한 뒤 '데이터 불러오기'를 눌러주세요.")
        ttk.Label(frm_sheet, textvariable=self.data_info_var, foreground="#555", wraplength=760,
                  justify="left").grid(row=2, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 8))

        frm_period = ttk.LabelFrame(self.root, text="3. 학습기간 / 예측기간 선택")
        frm_period.pack(fill="x", **pad)
        self.train_start_var = tk.StringVar()
        self.train_end_var = tk.StringVar()
        self.fc_start_var = tk.StringVar()
        self.fc_end_var = tk.StringVar()
        labels = ["학습 시작", "학습 종료", "예측 시작", "예측 종료"]
        period_vars = [self.train_start_var, self.train_end_var, self.fc_start_var, self.fc_end_var]
        self.period_combos = []
        for i, (lbl, var) in enumerate(zip(labels, period_vars)):
            ttk.Label(frm_period, text=lbl + ":").grid(row=0, column=2 * i, sticky="w",
                                                         padx=(8 if i == 0 else 10, 2), pady=8)
            combo = ttk.Combobox(frm_period, textvariable=var, state="readonly", width=11)
            combo.grid(row=0, column=2 * i + 1, sticky="w", padx=(0, 4), pady=8)
            self.period_combos.append(combo)

        frm_excl = ttk.LabelFrame(self.root, text="4. 돌발성(제외) 공정 선택")
        frm_excl.pack(fill="x", **pad)
        self.excl_frame_inner = ttk.Frame(frm_excl)
        self.excl_frame_inner.pack(fill="x", padx=8, pady=8)

        frm_out = ttk.LabelFrame(self.root, text="5. 결과 저장 폴더")
        frm_out.pack(fill="x", **pad)
        self.outdir_var = tk.StringVar()
        ttk.Entry(frm_out, textvariable=self.outdir_var).pack(side="left", fill="x", expand=True, padx=(8, 4), pady=8)
        ttk.Button(frm_out, text="찾아보기...", command=self._on_browse_outdir).pack(side="left", padx=(0, 8), pady=8)

        frm_run = ttk.Frame(self.root)
        frm_run.pack(fill="x", **pad)
        self.run_btn = ttk.Button(frm_run, text="예측 리포트 생성", command=self._on_run)
        self.run_btn.pack(side="left")
        self.progress = ttk.Progressbar(frm_run, mode="indeterminate", length=220)
        self.progress.pack(side="left", padx=12)

        frm_log = ttk.LabelFrame(self.root, text="실행 로그")
        frm_log.pack(fill="both", expand=True, **pad)
        self.log_text = scrolledtext.ScrolledText(frm_log, height=14, state="disabled", font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True, padx=8, pady=8)

        frm_done = ttk.Frame(self.root)
        frm_done.pack(fill="x", **pad)
        self.open_html_btn = ttk.Button(frm_done, text="HTML 리포트 열기", command=lambda: _open_path(self._result_paths.get("html")), state="disabled")
        self.open_html_btn.pack(side="left", padx=(0, 6))
        self.open_pdf_btn = ttk.Button(frm_done, text="PDF 리포트 열기", command=lambda: _open_path(self._result_paths.get("pdf")), state="disabled")
        self.open_pdf_btn.pack(side="left", padx=(0, 6))
        self.open_folder_btn = ttk.Button(frm_done, text="결과 폴더 열기", command=lambda: _open_path(self._result_paths.get("outdir")), state="disabled")
        self.open_folder_btn.pack(side="left")

    # ---------------- 이벤트 핸들러 ----------------
    def _log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _on_browse_file(self):
        path = filedialog.askopenfilename(
            title="RAWDATA 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")],
        )
        if not path:
            return
        self.file_var.set(path)
        try:
            self.sheet_names = pd.ExcelFile(path).sheet_names
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 파일을 열 수 없습니다: {e}")
            return
        self.sheet_combo["values"] = self.sheet_names
        self.sheet_var.set(detect_default_sheet(self.sheet_names))
        self.outdir_var.set(str(Path(path).parent / "output"))
        self.data_info_var.set("시트/기간단위를 확인한 뒤 '데이터 불러오기'를 눌러주세요.")

    def _on_browse_outdir(self):
        path = filedialog.askdirectory(title="결과 저장 폴더 선택")
        if path:
            self.outdir_var.set(path)

    def _on_load_periods(self):
        path = self.file_var.get()
        sheet = self.sheet_var.get()
        if not path or not sheet:
            messagebox.showwarning("안내", "먼저 RAWDATA 파일과 시트를 선택해주세요.")
            return
        gran = self.gran_var.get()
        try:
            df = load_rawdata(path, sheet)
            df = add_period_column(df, gran)
        except Exception as e:
            messagebox.showerror("오류", str(e))
            return

        self.current_gran = gran
        periods = build_period_axis(df, gran)
        default_h = GRANULARITY_META[gran]["default_horizon"]
        future_periods = [add_periods(periods[-1], i, gran) for i in range(1, default_h + 8)]
        fc_values = periods + future_periods

        self.train_start_var.set(periods[0])
        self.train_end_var.set(periods[-1])
        self.period_combos[0]["values"] = periods
        self.period_combos[1]["values"] = periods

        fc_start = add_periods(periods[-1], 1, gran)
        fc_end = add_periods(fc_start, default_h - 1, gran)
        self.period_combos[2]["values"] = fc_values
        self.period_combos[3]["values"] = fc_values
        self.fc_start_var.set(fc_start)
        self.fc_end_var.set(fc_end)

        n = len(df)
        dmin, dmax = df["매출일"].min().date(), df["매출일"].max().date()
        self.data_info_var.set(f"불러온 데이터: {n:,}건 ({dmin} ~ {dmax}), {gran} 단위로 {len(periods)}개 구간이 있습니다.")

        for w in self.excl_frame_inner.winfo_children():
            w.destroy()
        self.excl_vars = {}
        for proc in sorted(df["공정"].dropna().unique().tolist()):
            var = tk.BooleanVar(value=(proc == "미확인"))
            ttk.Checkbutton(self.excl_frame_inner, text=proc, variable=var).pack(side="left", padx=6)
            self.excl_vars[proc] = var

    def _on_run(self):
        if self.worker_running:
            return
        path = self.file_var.get()
        sheet = self.sheet_var.get()
        gran = self.gran_var.get()
        if not path or not sheet or self.current_gran is None:
            messagebox.showwarning("안내", "먼저 '데이터 불러오기 / 기간 확인'을 눌러주세요.")
            return
        if gran != self.current_gran:
            messagebox.showwarning("안내", "기간 단위를 변경하셨습니다. '데이터 불러오기'를 다시 눌러주세요.")
            return

        train_start, train_end = self.train_start_var.get(), self.train_end_var.get()
        fc_start, fc_end = self.fc_start_var.get(), self.fc_end_var.get()
        try:
            si = period_index(*parse_period(train_start, gran), gran)
            ei = period_index(*parse_period(train_end, gran), gran)
            if ei < si:
                raise ValueError("학습 종료는 학습 시작보다 앞설 수 없습니다.")
            fsi = period_index(*parse_period(fc_start, gran), gran)
            fei = period_index(*parse_period(fc_end, gran), gran)
            if fei < fsi:
                raise ValueError("예측 종료는 예측 시작보다 앞설 수 없습니다.")
        except ValueError as e:
            messagebox.showerror("오류", str(e))
            return

        outdir = self.outdir_var.get().strip()
        if not outdir:
            messagebox.showwarning("안내", "결과 저장 폴더를 지정해주세요.")
            return

        exclude = [p for p, v in self.excl_vars.items() if v.get()]

        self.run_btn.configure(state="disabled")
        self.progress.start(12)
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        for b in (self.open_html_btn, self.open_pdf_btn, self.open_folder_btn):
            b.configure(state="disabled")

        self.worker_running = True
        t = threading.Thread(
            target=self._worker,
            args=(path, sheet, gran, train_start, train_end, fc_start, fc_end, exclude, outdir),
            daemon=True,
        )
        t.start()

    def _worker(self, path, sheet, gran, train_start, train_end, fc_start, fc_end, exclude, outdir):
        try:
            result_paths = run_pipeline(
                input_path=path, sheet=sheet, gran=gran,
                train_start=train_start, train_end=train_end,
                forecast_start=fc_start, forecast_end=fc_end,
                exclude_processes=tuple(exclude), outdir=outdir,
                extra_log_handler=QueueLogHandler(self.msg_queue),
            )
            self.msg_queue.put(("done", result_paths))
        except Exception:
            import traceback
            self.msg_queue.put(("error", traceback.format_exc()))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "log":
                    self._log(payload)
                elif kind == "done":
                    self._on_worker_done(payload)
                elif kind == "error":
                    self._on_worker_error(payload)
        except queue.Empty:
            pass
        self.root.after(150, self._poll_queue)

    def _on_worker_done(self, result_paths):
        self.worker_running = False
        self.progress.stop()
        self.run_btn.configure(state="normal")
        self._result_paths = result_paths
        for b in (self.open_html_btn, self.open_pdf_btn, self.open_folder_btn):
            b.configure(state="normal")
        self._log("\n생성이 완료되었습니다.")
        messagebox.showinfo("완료", "매출 예측 리포트 생성이 완료되었습니다.")

    def _on_worker_error(self, tb_text):
        self.worker_running = False
        self.progress.stop()
        self.run_btn.configure(state="normal")
        self._log("\n[오류]\n" + tb_text)
        messagebox.showerror("오류", "리포트 생성 중 오류가 발생했습니다. 아래 로그를 확인해주세요.")


def launch_gui_app():
    _hide_console_window()
    root = tk.Tk()
    SalesForecastApp(root)
    root.mainloop()


# =====================================================================
# 메인
# =====================================================================
def main():
    ap = argparse.ArgumentParser(description="RAWDATA 엑셀로부터 예측 리포트를 생성합니다 (분기/월/년 단위 선택 가능).")
    ap.add_argument("input", nargs="?", default=None,
                     help="RAWDATA 엑셀 파일 경로 (지정하지 않으면 GUI 창이 뜹니다)")
    ap.add_argument("--granularity", choices=GRANULARITIES, default="분기",
                     help="예측기간 단위: 분기/월/년 (기본값: 분기)")
    ap.add_argument("--train-start", default=None,
                     help="학습 시작기간. 형식은 --granularity에 따라 다름 (분기 예: 2012Q1, 월 예: 2012-01, 년 예: 2012). 기본값: 데이터의 첫 기간")
    ap.add_argument("--train-end", default=None,
                     help="학습 종료기간 (형식은 --granularity와 동일). 기본값: 데이터의 마지막 기간")
    ap.add_argument("--forecast-start", default=None, help="예측 시작기간. 기본값: 학습종료기간 다음 기간")
    ap.add_argument("--forecast-end", default=None, help="예측 종료기간. 기본값: 예측시작기간 + 1년 분량")
    ap.add_argument("--exclude-process", nargs="*", default=["미확인"],
                     help="돌발성 매출로 간주해 예측에서 제외할 공정명 목록 (기본값: 미확인). 없애려면 --exclude-process 를 빈 값으로")
    ap.add_argument("--outdir", default=None, help="결과 저장 폴더 (기본값: 입력 파일과 같은 위치의 output 폴더)")
    ap.add_argument("--sheet", default=None, help="RAWDATA가 들어있는 시트 이름 (지정하지 않으면 자동 감지하거나 선택창이 뜹니다)")
    ap.add_argument("--no-gui", action="store_true", help="GUI 창을 띄우지 않고 콘솔 입력만 사용합니다")
    ap.add_argument("--column-guide", action="store_true", help="RAWDATA에 필요한 필수 컬럼 안내를 출력하고 종료합니다")
    args = ap.parse_args()

    if args.column_guide:
        print(column_guide_text())
        sys.exit(0)

    gran = args.granularity

    if not args.input:
        # 인자 없이 실행 + GUI 사용 가능 -> DOS 콘솔 대신 하나의 GUI 창에서 전 과정을 진행
        if not args.no_gui and tk is not None:
            try:
                launch_gui_app()
                sys.exit(0)
            except SystemExit:
                raise
            except Exception:
                # 디스플레이가 없는 등 GUI를 띄울 수 없는 환경 -> 콘솔 입력으로 대체
                print("[안내] GUI 창을 열 수 없어 콘솔 입력으로 진행합니다.")
        print("RAWDATA 엑셀 파일 경로가 지정되지 않았습니다.")
        print("사용법: python sales_forecast_report.py RAWDATA.xlsx\n")
        args.input = input("RAWDATA 엑셀 파일 경로를 입력하세요 (탐색기에서 파일을 이 창으로 끌어놓아도 됩니다): ")
        args.input = args.input.strip().strip('"').strip("'")
        if not args.input:
            _pause_and_exit("[오류] 입력 파일이 지정되지 않았습니다.")

    src_path = Path(args.input)
    if not src_path.exists():
        _pause_and_exit(f"[오류] 파일을 찾을 수 없습니다: {src_path}")

    try:
        sheet_names = pd.ExcelFile(src_path).sheet_names
    except Exception as e:
        _pause_and_exit(f"[오류] 엑셀 파일을 열 수 없습니다: {e}")

    default_sheet = detect_default_sheet(sheet_names)
    if args.sheet:
        if args.sheet not in sheet_names:
            _pause_and_exit(f"[오류] 지정한 시트를 찾을 수 없습니다: {args.sheet}\n사용 가능한 시트: {sheet_names}")
        chosen_sheet = args.sheet
    elif len(sheet_names) == 1:
        chosen_sheet = sheet_names[0]
    else:
        chosen_sheet = None if args.no_gui else pick_sheet_gui(sheet_names, default_sheet)
        if not chosen_sheet:
            chosen_sheet = pick_sheet_console(sheet_names, default_sheet)

    try:
        df_probe = load_rawdata(src_path, chosen_sheet)
        df_probe = add_period_column(df_probe, gran)
    except Exception as e:
        _pause_and_exit(f"[오류] {e}")

    all_periods = build_period_axis(df_probe, gran)
    train_start = args.train_start or all_periods[0]
    train_end = args.train_end or all_periods[-1]
    forecast_start = args.forecast_start or add_periods(train_end, 1, gran)
    default_h = GRANULARITY_META[gran]["default_horizon"]
    forecast_end = args.forecast_end or add_periods(forecast_start, default_h - 1, gran)

    outdir = Path(args.outdir) if args.outdir else (src_path.parent / "output")

    run_pipeline(
        input_path=src_path, sheet=chosen_sheet, gran=gran,
        train_start=train_start, train_end=train_end,
        forecast_start=forecast_start, forecast_end=forecast_end,
        exclude_processes=args.exclude_process or [], outdir=outdir,
    )


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except ModuleNotFoundError as e:
        _pause_and_exit(
            f"[오류] 필요한 라이브러리가 설치되어 있지 않습니다: {e}\n\n"
            "아래 명령을 실행해 필요한 패키지를 설치한 뒤 다시 실행해주세요:\n"
            "    pip install pandas numpy matplotlib openpyxl statsmodels fpdf2"
        )
    except Exception:
        import traceback
        traceback.print_exc()
        _pause_and_exit("\n[오류] 실행 중 문제가 발생했습니다. 위 내용을 확인해주세요.")
    else:
        try:
            input("\n엔터 키를 누르면 창을 닫습니다...")
        except (EOFError, KeyboardInterrupt):
            pass
