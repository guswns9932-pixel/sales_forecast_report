# -*- coding: utf-8 -*-
"""
sales_forecast_report.py
=========================
RAWDATA 엑셀(매출일/모델/반입라인/공정/메이커/수량/단가/매출액 포함)을 읽어
분기별 매출을 예측하고 HTML + PDF 리포트를 생성하는 독립 실행 스크립트입니다.

주요 기능:
  - 더블클릭(또는 인자 없이 실행) 시 GUI 창에서 RAWDATA 엑셀 파일을 직접 선택
  - 공정×모델×메이커(설비) 조합 단위로 예측하고, 모델별/설비(메이커)별/공정별로
    묶어서 볼 수 있는 리포트 생성
  - 여러 예측 알고리즘(계절성 지수평활, Croston-SBA 간헐수요모델, 선형추세,
    이동평균, 계절성 단순모형, 평균유지)을 실제 로우데이터로 백테스트하여
    조합마다 가장 정합성 높은(오차가 가장 작은) 알고리즘을 자동 선택
  - 실행 전 과정(백테스트 점수, 선택 근거, 데이터 처리 내역 등)을 메모장으로
    바로 열어볼 수 있는 텍스트 로그 파일로 저장
  - 결과 수치에 대한 해석 가이드와 신뢰도 표시를 리포트에 포함

사용법:
    python sales_forecast_report.py                (파일 선택 창이 뜹니다)
    python sales_forecast_report.py RAWDATA.xlsx
    python sales_forecast_report.py RAWDATA.xlsx --forecast-start 2026Q1 --forecast-end 2026Q2
    python sales_forecast_report.py RAWDATA.xlsx --train-end 2024Q4 --forecast-start 2025Q1 --forecast-end 2025Q4

자세한 옵션은:
    python sales_forecast_report.py --help
"""
import argparse
import base64
import io
import logging
import re
import sys
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


# =====================================================================
# 로깅 (메모장으로 바로 열어볼 수 있는 텍스트 로그)
# =====================================================================
def setup_logger(outdir):
    """콘솔과 텍스트 파일(UTF-8 BOM, 메모장 호환) 양쪽에 동시에 기록하는 로거를 만든다."""
    outdir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
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

    return logger, log_path


# =====================================================================
# GUI 파일/폴더 선택
# =====================================================================
def pick_file_gui():
    """RAWDATA 엑셀 파일을 GUI 창에서 선택한다. tkinter/디스플레이가 없으면 None 반환."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return None
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="RAWDATA 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")],
        )
        root.destroy()
        return path or None
    except Exception:
        return None


def show_done_gui(message):
    """완료 안내를 메시지 박스로 띄운다. 실패해도 프로그램 흐름에는 영향 없음."""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showinfo("매출 예측 리포트", message)
        root.destroy()
    except Exception:
        pass


def pick_sheet_gui(sheet_names, default_sheet):
    """RAWDATA가 들어있는 시트를 GUI 창에서 선택한다. 필수 컬럼 가이드도 함께 보여준다.
    tkinter/디스플레이가 없거나 취소하면 None 반환."""
    try:
        import tkinter as tk
    except Exception:
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
    df = df.dropna(subset=["매출일", "공정", "모델", "메이커"]).copy()
    n_dropped = n_before - len(df)
    if logger and n_dropped:
        logger.warning(f"매출일/공정/모델/메이커 중 결측값이 있는 {n_dropped}건을 제외했습니다.")

    df["매출일"] = pd.to_datetime(df["매출일"])
    df["연도"] = df["매출일"].dt.year
    df["분기"] = df["매출일"].dt.quarter
    df["연분기"] = df["연도"].astype(str) + "Q" + df["분기"].astype(str)

    amt_num = pd.to_numeric(df["매출액(KRW)"], errors="coerce")
    n_bad_amt = int(amt_num.isna().sum() - df["매출액(KRW)"].isna().sum())
    df["매출액(KRW)"] = amt_num.fillna(0)

    qty_num = pd.to_numeric(df["수량"], errors="coerce")
    n_bad_qty = int(qty_num.isna().sum() - df["수량"].isna().sum())
    df["수량"] = qty_num.fillna(0)

    if logger and (n_bad_amt or n_bad_qty):
        logger.warning(f"숫자로 변환할 수 없는 값을 0으로 처리했습니다: 매출액 {n_bad_amt}건, 수량 {n_bad_qty}건")

    return df


# =====================================================================
# 예측 알고리즘 후보
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


def moving_average_forecast(values, n_ahead, window=4):
    """최근 window개 분기의 평균을 그대로 미래 예측치로 사용."""
    values = np.asarray(values, dtype=float)
    w = min(window, len(values)) if len(values) else 0
    base = float(np.mean(values[-w:])) if w else 0.0
    return np.full(n_ahead, max(0.0, base))


def seasonal_naive_forecast(values, n_ahead, season=4):
    """1년 전(직전 season개 분기) 같은 분기의 실적을 그대로 사용."""
    values = np.asarray(values, dtype=float)
    n = len(values)
    if n < season:
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


CANDIDATE_METHODS = {
    "계절성 지수평활(Holt-Winters)": lambda v, h: seasonal_ets_forecast(v, h),
    "간헐수요모델(Croston-SBA)": lambda v, h: np.full(h, croston_sba(v)),
    "선형추세": lambda v, h: linear_trend_forecast(v, h),
    "이동평균(최근4분기)": lambda v, h: moving_average_forecast(v, h, 4),
    "계절성 단순모형(전년동분기)": lambda v, h: seasonal_naive_forecast(v, h, 4),
    "평균유지": lambda v, h: simple_mean_forecast(v, h),
}

METHOD_DESCRIPTIONS = {
    "계절성 지수평활(Holt-Winters)": "과거 성장 추세와 분기별로 반복되는 계절성 패턴을 함께 반영하는 통계 모델입니다.",
    "간헐수요모델(Croston-SBA)": "거래가 간헐적으로 발생하는 품목에 특화된 모델로, 발생 확률과 발생 시 평균 규모를 함께 추정합니다.",
    "선형추세": "과거 추세선을 직선으로 연장하는 단순한 방법입니다.",
    "이동평균(최근4분기)": "가장 최근 4개 분기의 평균값을 다음 분기 예측치로 그대로 사용합니다.",
    "계절성 단순모형(전년동분기)": "1년 전 같은 분기의 실적을 그대로 사용합니다. 추세는 약하지만 계절성이 뚜렷할 때 유리합니다.",
    "평균유지": "과거 전체 평균값을 보수적으로 유지합니다.",
    "표본부족(평균유지)": "유효한 거래 데이터가 너무 적어 백테스트를 수행할 수 없어, 과거 평균을 보수적으로 유지했습니다.",
    "제외(돌발성 매출)": "돌발성·일회성 매출로 분류되어 예측 대상에서 제외되었습니다(과거 실적은 참고용으로만 표시).",
}


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


def forecast_series(values, n_ahead, logger=None, series_name="", min_intermittent_obs=3):
    """여러 예측 알고리즘을 백테스트(과거 구간을 학습/검증으로 나눠 실제값과 비교)하여
    오차(sMAPE)가 가장 낮은, 즉 가장 정합성 높은 방법으로 최종 예측한다.
    데이터가 너무 짧아 백테스트가 불가능하면 표본 크기에 따른 보수적 방법으로 대체한다.
    반환: (forecast_array, method_name, backtest_scores dict[method_name -> smape])"""
    values = np.asarray(values, dtype=float)
    n = len(values)
    nz_count = int(np.count_nonzero(values))

    if n == 0 or nz_count == 0:
        if logger:
            logger.info(f"[{series_name}] 거래 이력이 없어 예측치를 0으로 둡니다.")
        return np.zeros(n_ahead), "표본부족(평균유지)", {}

    test_h = max(1, min(n_ahead, 4))
    min_train = 4
    origins = sorted({o for o in (n - test_h, n - 2 * test_h) if o >= min_train}, reverse=True)

    scores = {}
    if origins:
        for name, fn in CANDIDATE_METHODS.items():
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
            return np.full(n_ahead, rate), "간헐수요모델(Croston-SBA)", {}
        avg = float(np.mean(values[values > 0])) if nz_count > 0 else 0.0
        if logger:
            logger.info(f"[{series_name}] 표본 부족(n={n}, 유효거래 {nz_count}건) → 평균유지 적용")
        return np.full(n_ahead, avg), "표본부족(평균유지)", {}

    ranked = sorted(scores.items(), key=lambda kv: kv[1])
    if logger:
        score_str = ", ".join(f"{name} 오차 {s*100:.1f}%" for name, s in ranked)
        logger.info(f"[{series_name}] 백테스트(n={n}, 검증창 {test_h}분기 x {len(origins)}회) 결과 → {score_str}")

    for name, _ in ranked:
        fc = CANDIDATE_METHODS[name](values, n_ahead)
        if fc is not None:
            fc = np.maximum(np.asarray(fc, dtype=float), 0.0)
            if logger:
                logger.info(f"[{series_name}] 최종 선택 알고리즘: {name}")
            return fc, name, scores

    avg = float(np.mean(values[values > 0])) if nz_count > 0 else 0.0
    return np.full(n_ahead, avg), "표본부족(평균유지)", scores


# =====================================================================
# 집계 & 예측 실행 (공정 × 모델 × 메이커(설비) 단위)
# =====================================================================
COMBO_COLS = ["공정", "모델", "메이커"]


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
                  exclude_processes=("미확인",), logger=None):
    train_quarters = quarter_range(train_start, train_end)
    fc_quarters = quarter_range(forecast_start, forecast_end)
    n_ahead = len(fc_quarters)

    combos = sorted(df.groupby(COMBO_COLS).size().index.tolist())
    processes = sorted(df["공정"].unique().tolist())

    amt_piv = quarterly_pivot(df, COMBO_COLS, "매출액(KRW)", train_quarters)
    qty_piv = quarterly_pivot(df, COMBO_COLS, "수량", train_quarters)

    if logger:
        logger.info(f"공정×모델×메이커 조합 {len(combos)}개에 대해 조합별로 예측을 계산합니다.")

    results = []
    for (proc, model, maker) in combos:
        key = (proc, model, maker)
        amt_hist = amt_piv[key].values if key in amt_piv.columns else np.zeros(len(train_quarters))
        qty_hist = qty_piv[key].values if key in qty_piv.columns else np.zeros(len(train_quarters))
        series_name = f"{proc}/{model}/{maker}"

        excluded = proc in exclude_processes
        if excluded:
            amt_fc = np.zeros(n_ahead)
            qty_fc = np.zeros(n_ahead)
            method = "제외(돌발성 매출)"
            scores = {}
        else:
            amt_fc, method, scores = forecast_series(amt_hist, n_ahead, logger=logger,
                                                       series_name=f"{series_name} [금액]")
            qty_fc, _, _ = forecast_series(qty_hist, n_ahead, logger=logger,
                                            series_name=f"{series_name} [수량]")

        baseline_len = min(n_ahead, len(amt_hist))
        baseline_amt = float(np.sum(amt_hist[-baseline_len:])) if baseline_len else 0.0

        results.append({
            "공정": proc, "모델": model, "메이커": maker, "excluded": excluded, "method": method,
            "backtest_scores": scores,
            "amt_hist": amt_hist, "qty_hist": qty_hist,
            "amt_fc": amt_fc, "qty_fc": qty_fc,
            "amt_fc_total": float(np.sum(amt_fc)), "qty_fc_total": float(np.sum(qty_fc)),
            "baseline_amt": baseline_amt,
            "diff_amt": float(np.sum(amt_fc)) - baseline_amt,
            "nz_count": int(np.count_nonzero(amt_hist)),
            "txn_count": int(((df["공정"] == proc) & (df["모델"] == model) & (df["메이커"] == maker) &
                               (df["연분기"].isin(train_quarters))).sum()),
        })

    detail = pd.DataFrame(results)

    # 전사/공정 top-down (계절성 반영) — bottom-up과 교차검증용
    total_hist = amt_piv.sum(axis=1).values
    total_fc_topdown, total_method_topdown, _ = forecast_series(
        total_hist, n_ahead, logger=logger, series_name="전사 합계(하향식)")

    proc_topdown = {}
    for proc in processes:
        if proc in exclude_processes:
            continue
        cols = [c for c in amt_piv.columns if c[0] == proc]
        if not cols:
            proc_topdown[proc] = (np.zeros(n_ahead), "데이터없음", {})
            continue
        series = amt_piv[cols].sum(axis=1).values
        fc, m, sc = forecast_series(series, n_ahead, logger=logger, series_name=f"공정 합계(하향식): {proc}")
        proc_topdown[proc] = (fc, m, sc)

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


def dimension_summary(detail, dim_col):
    """예측이 제외되지 않은 조합들을 지정한 차원(공정/모델/메이커) 기준으로 다시 합산한다."""
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
    ax.bar(x - w / 2, s["직전동기간실적"] / 1e8, width=w, color=BRAND_GRAY, label="직전 동기간 실적")
    ax.bar(x + w / 2, s["예측합계"] / 1e8, width=w, color=BRAND_MAGENTA, label="예측")
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
    labels = [f"{p}-{m}-{k}" for p, m, k in zip(d["공정"], d["모델"], d["메이커"])]
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
.tablewrap {{ overflow-x:auto; }}
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
.conf-high {{ color:#1F7A1F; font-weight:bold; }}
.conf-mid {{ color:#B8860B; font-weight:bold; }}
.conf-low {{ color:#C00000; font-weight:bold; }}
details > summary {{ cursor:pointer; font-weight:bold; color:{magenta}; }}
code {{ background:#f0f0f0; padding:1px 5px; border-radius:4px; }}
</style></head>
<body><div class="wrap">
<h1>매출 예측 리포트</h1>
<div class="meta">학습기간: {train_start} ~ {train_end} &nbsp;|&nbsp; 예측기간: {fc_start} ~ {fc_end} &nbsp;|&nbsp; 생성 파일: {src_name} ({sheet_name} 시트)</div>
{exclude_note}

<div class="card">
<details>
<summary>이 리포트 읽는 법 (클릭하여 펼치기/접기)</summary>
<div style="font-size:13px; line-height:1.8; margin-top:10px; color:#333;">
<b>직전 동기간 실적</b>: 예측기간과 길이가 같은 가장 최근 실적기간의 실제 매출입니다. (예: 4개 분기를 예측하면 직전 4개 분기의 실제 실적과 비교합니다)<br>
<b>예측합계(상향식)</b>: 공정×모델×메이커 조합 하나하나에 대해, 아래 "백테스트"로 가장 정확했던 알고리즘을 적용해 예측한 값을 모두 더한 값입니다. 이 리포트의 기본 예측치입니다.<br>
<b>하향식(전사 계절성모델)</b>: 조합별로 나누지 않고 전사 매출 전체에 계절성 모델 하나를 적용한 값입니다. 상향식과 크게 차이가 나면 특정 조합에 이상치가 섞였을 가능성을 점검해볼 수 있습니다.<br>
<b>백테스트</b>: 과거 데이터를 학습구간과 검증구간으로 나눈 뒤, 학습구간만으로 검증구간을 예측해보고 실제값과 얼마나 차이 나는지(오차율, sMAPE) 계산하는 절차입니다. 이 리포트는 6가지 예측 알고리즘을 모두 백테스트해서 조합마다 오차가 가장 작은 알고리즘을 자동으로 선택합니다.<br>
<b>예측신뢰도</b>: 백테스트 오차율 기준입니다 — 10% 이하 <span class="conf-high">높음</span>, 10~25% <span class="conf-mid">보통</span>, 25% 초과 <span class="conf-low">낮음(참고용)</span>. 데이터가 너무 적어 백테스트 자체가 불가능했던 조합은 'N/A'로 표시하고 보수적으로 과거 평균을 유지합니다.<br>
<b>메이커(설비)</b>: RAWDATA의 '메이커' 컬럼을 설비 제조사 기준 구분으로 사용해 집계했습니다.<br>
실행 과정 전체와 조합별 백테스트 점수 등 모든 상세 로그는 함께 생성된 텍스트 파일(<code>{log_name}</code>)에서 확인할 수 있습니다(메모장으로 바로 열립니다).
</div>
</details>
</div>

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
<div class="tablewrap">{proc_table}</div>
</div>

<div class="card"><h2 style="margin-top:0;border:none;">모델별 실적 대비 예측</h2>
<img src="data:image/png;base64,{chart_model}"/>
<div class="tablewrap">{model_table}</div>
</div>

<div class="card"><h2 style="margin-top:0;border:none;">메이커(설비)별 실적 대비 예측</h2>
<img src="data:image/png;base64,{chart_maker}"/>
<div class="tablewrap">{maker_table}</div>
</div>

<div class="card"><h2 style="margin-top:0;border:none;">증감 기여도 상위 조합</h2>
<img src="data:image/png;base64,{chart_top}"/>
<p style="font-size:13px;color:{gray}">{narrative}</p>
</div>

<h2>공정×모델×메이커 상세 예측 (전체 {n_combo}개 조합)</h2>
<div class="tablewrap">{detail_table}</div>

<h2>예측 방법론 (다중 알고리즘 백테스트)</h2>
<div class="card" style="font-size:13px; line-height:1.7;">
이 리포트는 각 조합마다 아래 알고리즘들을 모두 후보로 놓고, 과거 데이터로 백테스트(학습/검증 분리 검증)를 수행해 오차(sMAPE)가 가장 작은 알고리즘을 자동으로 선택합니다. 검증할 데이터가 너무 짧은 조합은 표본 크기에 맞는 보수적인 방법으로 대체합니다.<br><br>
{method_section}
<br>모든 예측치는 0 미만이 되지 않도록 하한을 적용했습니다.
</div>

</div></body></html>"""


def fmt_krw(v):
    return f"{v:,.0f}"


def _dimension_table_html(summary, dim_col, dim_label):
    def pct_cell(v):
        if pd.isna(v):
            return '<span class="methodbadge">N/A</span>'
        cls = "pos" if v >= 0 else "neg"
        sign = "+" if v >= 0 else ""
        return f'<span class="{cls}">{sign}{v*100:.1f}%</span>'

    rows = "\n".join(
        f"<tr><td>{r[dim_col]}</td><td>{fmt_krw(r['직전동기간실적'])}</td><td>{fmt_krw(r['예측합계'])}</td>"
        f"<td>{fmt_krw(r['증감액'])}</td><td>{pct_cell(r['증감률'])}</td></tr>"
        for _, r in summary.iterrows()
    )
    return (f"<table><tr><th>{dim_label}</th><th>직전동기간실적</th><th>예측합계</th><th>증감액</th><th>증감률</th></tr>"
            f"{rows}</table>")


def build_html_report(df, result, src_name, sheet_name, exclude_processes, log_path):
    detail = result["detail"]
    train_quarters, fc_quarters = result["train_quarters"], result["fc_quarters"]
    n_ahead = len(fc_quarters)

    active = detail[~detail["excluded"]]
    total_base = active["baseline_amt"].sum()
    total_fc = active["amt_fc_total"].sum()
    total_pct = (total_fc - total_base) / total_base if total_base else None
    total_topdown = float(np.sum(result["total_fc_topdown"]))

    proc_summary = dimension_summary(detail, "공정")
    model_summary = dimension_summary(detail, "모델")
    maker_summary = dimension_summary(detail, "메이커")

    total_fc_bottomup_arr = np.sum(np.stack(active["amt_fc"].values), axis=0) if len(active) else np.zeros(n_ahead)
    chart_total = chart_total_trend(train_quarters, result["total_hist"], fc_quarters,
                                     total_fc_bottomup_arr, result["total_fc_topdown"])
    chart_proc = chart_dimension_bar(proc_summary, "공정", "공정별 실적 대비 예측")
    chart_model = chart_dimension_bar(model_summary, "모델", "모델별 실적 대비 예측")
    chart_maker = chart_dimension_bar(maker_summary, "메이커", "메이커(설비)별 실적 대비 예측")
    chart_top = chart_top_contributors(detail)

    proc_table = _dimension_table_html(proc_summary, "공정", "공정")
    model_table = _dimension_table_html(model_summary, "모델", "모델")
    maker_table = _dimension_table_html(maker_summary, "메이커", "메이커(설비)")

    top3 = active.reindex(active["diff_amt"].abs().sort_values(ascending=False).index).head(3)
    conf_counts = active.apply(lambda r: confidence_label(r["backtest_scores"], r["method"]), axis=1).value_counts() \
        if len(active) else pd.Series(dtype=int)
    conf_txt = (f" 조합별 예측 신뢰도는 높음 {int(conf_counts.get('높음', 0))}건, "
                f"보통 {int(conf_counts.get('보통', 0))}건, "
                f"낮음(참고용) {int(conf_counts.get('낮음(참고용)', 0))}건, "
                f"산정불가(N/A) {int(conf_counts.get('N/A', 0))}건입니다.")

    if total_pct is not None and len(top3):
        narrative = (f"설정한 예측기간({fc_quarters[0]}~{fc_quarters[-1]}) 전사 합계 예측은 {fmt_krw(total_fc)}원으로, "
                     f"직전 동일 길이 기간 실적({fmt_krw(total_base)}원) 대비 {'+' if total_pct>=0 else ''}{total_pct*100:.1f}% "
                     f"{'증가' if total_pct>=0 else '감소'}입니다. 가장 큰 요인은 " +
                     ", ".join(f"{r['공정']}-{r['모델']}-{r['메이커']}({'+' if r['diff_amt']>=0 else ''}{fmt_krw(r['diff_amt'])}원)"
                               for _, r in top3.iterrows()) + " 입니다." + conf_txt)
    else:
        narrative = conf_txt

    def conf_span(scores, method):
        label = confidence_label(scores, method)
        cls = {"높음": "conf-high", "보통": "conf-mid", "낮음(참고용)": "conf-low"}.get(label, "")
        return f'<span class="{cls}">{label}</span>' if cls else label

    det_sorted = detail.sort_values("amt_fc_total", ascending=False)
    detail_rows = []
    for _, r in det_sorted.iterrows():
        detail_rows.append(
            f"<tr><td>{r['공정']}</td><td>{r['모델']}</td><td style='text-align:left'>{r['메이커']}</td>"
            f"<td>{fmt_krw(r['baseline_amt'])}</td>"
            f"<td>{fmt_krw(r['amt_fc_total'])}</td><td>{r['qty_fc_total']:.1f}</td>"
            f"<td><span class='methodbadge'>{r['method']}</span></td>"
            f"<td>{conf_span(r['backtest_scores'], r['method'])}</td></tr>")
    detail_table = (f"<table><tr><th>공정</th><th>모델</th><th>메이커(설비)</th><th>직전동기간실적(금액)</th>"
                     f"<th>예측합계(금액)</th><th>예측합계(수량)</th><th>적용알고리즘</th><th>예측신뢰도</th></tr>"
                     f"{''.join(detail_rows)}</table>")

    exclude_note = ""
    if exclude_processes:
        exclude_note = (f'<div class="note">공정이 {", ".join(exclude_processes)} 인 매출은 돌발성(일회성) '
                         f"매출로 간주되어 예측 및 위 합계 지표에서 제외되었습니다. (과거 실적은 상세표에서 참고용으로 확인 가능)</div>")

    method_counts = detail[~detail["excluded"]]["method"].value_counts()
    method_lines = []
    for name, cnt in method_counts.items():
        desc = METHOD_DESCRIPTIONS.get(name, "선정된 통계적 방법으로 예측했습니다.")
        method_lines.append(f"<b>{name}</b> ({int(cnt)}개 조합): {desc}")
    excl_count = int(detail["excluded"].sum())
    if excl_count:
        method_lines.append(f"<b>제외(돌발성 매출)</b> ({excl_count}개 조합): {METHOD_DESCRIPTIONS['제외(돌발성 매출)']}")
    method_section = "<br>".join(method_lines)

    html = HTML_TEMPLATE.format(
        magenta=BRAND_MAGENTA, gray=BRAND_GRAY,
        train_start=train_quarters[0], train_end=train_quarters[-1],
        fc_start=fc_quarters[0], fc_end=fc_quarters[-1], src_name=src_name, sheet_name=sheet_name,
        exclude_note=exclude_note, log_name=log_path.name,
        total_base_fmt=fmt_krw(total_base), total_fc_fmt=fmt_krw(total_fc),
        total_pct_fmt=(f"{'+' if total_pct>=0 else ''}{total_pct*100:.1f}%" if total_pct is not None else "N/A"),
        total_topdown_fmt=fmt_krw(total_topdown), topdown_method=result["total_method_topdown"],
        chart_total=chart_total, chart_proc=chart_proc, chart_model=chart_model, chart_maker=chart_maker,
        chart_top=chart_top,
        proc_table=proc_table, model_table=model_table, maker_table=maker_table,
        narrative=narrative,
        n_combo=len(detail), detail_table=detail_table,
        method_section=method_section,
    )

    ctx = {
        "proc_summary": proc_summary, "model_summary": model_summary, "maker_summary": maker_summary,
        "total_base": total_base, "total_fc": total_fc, "total_pct": total_pct, "total_topdown": total_topdown,
        "chart_total": chart_total, "chart_proc": chart_proc, "chart_model": chart_model,
        "chart_maker": chart_maker, "chart_top": chart_top,
    }
    return html, ctx


# =====================================================================
# PDF 리포트 (요약 위주, 가벼운 버전)
# =====================================================================
def _pdf_dimension_table(pdf, summary, dim_col, header_label):
    pdf.set_font(pdf.font_family, "", 9)
    pdf.set_fill_color(200, 0, 124)
    pdf.set_text_color(255, 255, 255)
    headers = [header_label, "직전실적", "예측합계", "증감률"]
    widths = [50, 45, 45, 40]
    for h, w in zip(headers, widths):
        pdf.cell(w, 8, h, border=1, align="C", fill=True)
    pdf.ln()
    pdf.set_text_color(30, 30, 30)
    for _, r in summary.iterrows():
        pct = r["증감률"]
        pct_s = f"{'+' if pd.notna(pct) and pct>=0 else ''}{pct*100:.1f}%" if pd.notna(pct) else "N/A"
        pdf.cell(widths[0], 7, str(r[dim_col])[:22], border=1)
        pdf.cell(widths[1], 7, fmt_krw(r["직전동기간실적"]), border=1, align="R")
        pdf.cell(widths[2], 7, fmt_krw(r["예측합계"]), border=1, align="R")
        pdf.cell(widths[3], 7, pct_s, border=1, align="R")
        pdf.ln()


def build_pdf_report(outpath, df, result, src_name, sheet_name, exclude_processes, ctx, log_path):
    from fpdf import FPDF

    train_quarters, fc_quarters = result["train_quarters"], result["fc_quarters"]
    detail = result["detail"]
    total_base, total_fc = ctx["total_base"], ctx["total_fc"]
    total_pct, total_topdown = ctx["total_pct"], ctx["total_topdown"]

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
                    f"예측기간 {fc_quarters[0]}~{fc_quarters[-1]}  |  원본: {src_name} ({sheet_name} 시트)", ln=True)
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

    img1 = io.BytesIO(base64.b64decode(ctx["chart_total"]))
    pdf.image(img1, w=180)

    for title, summary, dim_col, chart_key, header_label in [
        ("공정별 실적 대비 예측", ctx["proc_summary"], "공정", "chart_proc", "공정"),
        ("모델별 실적 대비 예측", ctx["model_summary"], "모델", "chart_model", "모델"),
        ("메이커(설비)별 실적 대비 예측", ctx["maker_summary"], "메이커", "chart_maker", "메이커(설비)"),
    ]:
        pdf.add_page()
        pdf.set_font(pdf.font_family, "", 13)
        pdf.set_text_color(200, 0, 124)
        pdf.cell(0, 10, title, ln=True)
        img = io.BytesIO(base64.b64decode(ctx[chart_key]))
        pdf.image(img, w=180)
        pdf.ln(2)
        _pdf_dimension_table(pdf, summary, dim_col, header_label)

    pdf.add_page()
    pdf.set_font(pdf.font_family, "", 13)
    pdf.set_text_color(200, 0, 124)
    pdf.cell(0, 10, "증감 기여도 상위 조합", ln=True)
    img3 = io.BytesIO(base64.b64decode(ctx["chart_top"]))
    pdf.image(img3, w=180)

    pdf.add_page()
    pdf.set_font(pdf.font_family, "", 13)
    pdf.set_text_color(200, 0, 124)
    pdf.cell(0, 10, "예측 방법론 요약 (다중 알고리즘 백테스트)", ln=True)
    pdf.set_font(pdf.font_family, "", 10)
    pdf.set_text_color(30, 30, 30)
    method_counts = detail[~detail["excluded"]]["method"].value_counts()
    lines = [f"- {name}: {int(cnt)}개 조합 - {METHOD_DESCRIPTIONS.get(name, '')}"
             for name, cnt in method_counts.items()]
    excl_count = int(detail["excluded"].sum())
    if excl_count:
        lines.append(f"- 제외(돌발성 매출): {excl_count}개 조합")
    lines += [
        "- 조합마다 6개 예측 알고리즘을 과거 데이터로 백테스트하여 오차(sMAPE)가 가장 낮은 알고리즘을 자동 선택했습니다.",
        "- 모든 예측치는 0 미만이 되지 않도록 하한을 적용했습니다.",
        f"- 조합별 백테스트 상세 점수는 실행 로그 파일({log_path.name})에서 확인할 수 있습니다.",
        "- 상세 조합별 결과와 그래프는 함께 생성된 HTML 리포트에서도 확인하실 수 있습니다.",
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
    ap.add_argument("input", nargs="?", default=None,
                     help="RAWDATA 엑셀 파일 경로 (지정하지 않으면 파일 선택 창이 뜹니다)")
    ap.add_argument("--train-start", default=None, help="학습 시작분기 (예: 2012Q1). 기본값: 데이터의 첫 분기")
    ap.add_argument("--train-end", default=None, help="학습 종료분기 (예: 2025Q4). 기본값: 데이터의 마지막 분기")
    ap.add_argument("--forecast-start", default=None, help="예측 시작분기. 기본값: 학습종료분기 다음 분기")
    ap.add_argument("--forecast-end", default=None, help="예측 종료분기. 기본값: 예측시작분기+3(4개 분기)")
    ap.add_argument("--exclude-process", nargs="*", default=["미확인"],
                     help="돌발성 매출로 간주해 예측에서 제외할 공정명 목록 (기본값: 미확인). 없애려면 --exclude-process 를 빈 값으로")
    ap.add_argument("--outdir", default=None, help="결과 저장 폴더 (기본값: 입력 파일과 같은 위치의 output 폴더)")
    ap.add_argument("--sheet", default=None, help="RAWDATA가 들어있는 시트 이름 (지정하지 않으면 자동 감지하거나 선택창이 뜹니다)")
    ap.add_argument("--no-gui", action="store_true", help="파일/시트 선택 창을 띄우지 않고 콘솔 입력만 사용합니다")
    ap.add_argument("--column-guide", action="store_true", help="RAWDATA에 필요한 필수 컬럼 안내를 출력하고 종료합니다")
    args = ap.parse_args()

    if args.column_guide:
        print(column_guide_text())
        sys.exit(0)

    used_gui = False
    if not args.input:
        picked = None if args.no_gui else pick_file_gui()
        if picked:
            args.input = picked
            used_gui = True
        else:
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

    outdir = Path(args.outdir) if args.outdir else (src_path.parent / "output")
    logger, log_path = setup_logger(outdir)
    logger.info("=" * 70)
    logger.info("매출 예측 리포트 생성을 시작합니다")
    logger.info(f"입력 파일: {src_path}")
    logger.info(f"사용 시트: {chosen_sheet} (엑셀 내 시트 목록: {sheet_names})")
    logger.info(f"결과 폴더: {outdir}")
    logger.info("=" * 70)
    logger.info(column_guide_text())

    logger.info(f"[1/5] RAWDATA 읽는 중... ({src_path.name} / 시트: {chosen_sheet})")
    df = load_rawdata(src_path, chosen_sheet, logger=logger)
    logger.info(f"      -> {len(df):,}건, {df['매출일'].min().date()} ~ {df['매출일'].max().date()}")

    all_quarters = build_quarter_axis(df)
    train_start = args.train_start or all_quarters[0]
    train_end = args.train_end or all_quarters[-1]
    forecast_start = args.forecast_start or add_quarters(train_end, 1)
    forecast_end = args.forecast_end or add_quarters(forecast_start, 3)

    logger.info(f"[2/5] 예측 계산 중 (알고리즘별 백테스트 수행)... 학습 {train_start}~{train_end} / 예측 {forecast_start}~{forecast_end}")
    result = run_forecast(df, train_start, train_end, forecast_start, forecast_end,
                           exclude_processes=tuple(args.exclude_process or []), logger=logger)

    logger.info("[3/5] 그래프/리포트 생성 중...")
    html, ctx = build_html_report(df, result, src_path.name, chosen_sheet, args.exclude_process or [], log_path)

    html_path = outdir / "매출예측_리포트.html"
    html_path.write_text(html, encoding="utf-8")
    logger.info(f"      -> {html_path}")

    logger.info("[4/5] PDF 생성 중...")
    pdf_path = outdir / "매출예측_리포트.pdf"
    build_pdf_report(pdf_path, df, result, src_path.name, chosen_sheet, args.exclude_process or [], ctx, log_path)
    logger.info(f"      -> {pdf_path}")

    logger.info("[5/5] 완료!")
    logger.info(f"전사 직전동기간 실적: {fmt_krw(ctx['total_base'])}원")
    if ctx["total_pct"] is not None:
        sign = "+" if ctx["total_pct"] >= 0 else ""
        logger.info(f"전사 예측 합계(상향식): {fmt_krw(ctx['total_fc'])}원 ({sign}{ctx['total_pct']*100:.1f}%)")
    else:
        logger.info(f"전사 예측 합계(상향식): {fmt_krw(ctx['total_fc'])}원")
    logger.info(f"전사 예측(하향식/계절성모델): {fmt_krw(ctx['total_topdown'])}원")
    logger.info(f"실행 로그 저장 위치: {log_path}")

    if used_gui:
        show_done_gui(
            "매출 예측 리포트 생성이 완료되었습니다.\n\n"
            f"HTML: {html_path}\n"
            f"PDF: {pdf_path}\n"
            f"실행 로그: {log_path}"
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
