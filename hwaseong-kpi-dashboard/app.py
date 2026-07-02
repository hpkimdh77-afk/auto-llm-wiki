# coding: utf-8

import os
import io
import json
import uuid
import base64
from pathlib import Path
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import streamlit as st

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


# =====================================================
# Streamlit 기본 설정
# =====================================================

st.set_page_config(page_title="2026 사고 KPI", layout="wide")


# =====================================================
# 경로 및 기본 설정
# =====================================================

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
REPORT_DIR = APP_DIR / "reports"
UPLOAD_DIR = APP_DIR / "uploads"

for p in [DATA_DIR, REPORT_DIR, UPLOAD_DIR]:
    p.mkdir(parents=True, exist_ok=True)

HISTORY_FILE = DATA_DIR / "analysis_history.json"

ACCIDENT_FILE = os.getenv(
    "ACCIDENT_FILE",
    "/config/work/sharedworkspace/ITS등록사고_streamlit.xlsx",
)

KPI_FILE = os.getenv(
    "KPI_FILE",
    "/config/work/sharedworkspace/화성KPI_streamlit.xlsx",
)

CURRENT_YEAR = int(os.getenv("KPI_CURRENT_YEAR", "2026"))
KPI_YEAR = CURRENT_YEAR % 100
CURRENT_MONTH = datetime.now().month
CURRENT_WEEK = datetime.now().isocalendar().week
WEEK_LABEL = f"{max(CURRENT_WEEK - 1, 1)}W"


# =====================================================
# DS OSS API 기본값
# =====================================================

DEFAULT_API_BASE_URL = "http://apigw.samsungds.net:8000/gpt-oss-playground/1/v1"
DEFAULT_USER_ID = "dongrami.kim"
DEFAULT_MODEL = "gpt-oss-120b"
DEFAULT_MAX_TOKENS = 12000
DEFAULT_TIMEOUT = 300
DEFAULT_RETRY = 2


# =====================================================
# 그래프 폰트 크기
# =====================================================

GRAPH_TITLE_SIZE = 15
GRAPH_TICK_SIZE = 11
GRAPH_LABEL_SIZE = 11
GRAPH_TEXT_SIZE = 14
GRAPH_LEGEND_SIZE = 11


# =====================================================
# Matplotlib 한글 폰트 설정
# =====================================================

KOREAN_FONT_PROP = None
KOREAN_FONT_NAME = None


def find_korean_font_path():
    font_dir = Path("/config/work/sharedworkspace/fonts")

    candidates = [
        font_dir / "NotoSansKR-Regular.ttf",
        font_dir / "NotoSansKR-VariableFont_wght.ttf",
        font_dir / "NotoSansKR-Medium.ttf",
        font_dir / "NotoSansKR-Bold.ttf",
        Path("/config/work/sharedworkspace/NotoSansKR-Regular.ttf"),
        Path("/config/work/sharedworkspace/NotoSansKR-VariableFont_wght.ttf"),
    ]

    for path in candidates:
        if path.exists():
            return str(path)

    if font_dir.exists():
        patterns = [
            "*Noto*Sans*KR*.ttf",
            "*Noto*Sans*KR*.otf",
            "*NotoSansKR*.ttf",
            "*NotoSansKR*.otf",
            "*Noto*.ttf",
            "*Noto*.otf",
        ]

        for pattern in patterns:
            matched = sorted(font_dir.glob(pattern))
            if matched:
                return str(matched[0])

    return None


def set_korean_font():
    global KOREAN_FONT_PROP
    global KOREAN_FONT_NAME

    font_path = find_korean_font_path()

    if font_path and os.path.exists(font_path):
        try:
            fm.fontManager.addfont(font_path)
            KOREAN_FONT_PROP = fm.FontProperties(fname=font_path)
            KOREAN_FONT_NAME = KOREAN_FONT_PROP.get_name()

            plt.rcParams["font.family"] = KOREAN_FONT_NAME
            plt.rcParams["font.sans-serif"] = [KOREAN_FONT_NAME]
            plt.rcParams["axes.unicode_minus"] = False
            return

        except Exception:
            KOREAN_FONT_PROP = None
            KOREAN_FONT_NAME = None

    KOREAN_FONT_PROP = None
    KOREAN_FONT_NAME = None
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["axes.unicode_minus"] = False


def apply_korean_font_to_axes(ax):
    if KOREAN_FONT_PROP is None:
        return

    ax.title.set_fontproperties(KOREAN_FONT_PROP)
    ax.xaxis.label.set_fontproperties(KOREAN_FONT_PROP)
    ax.yaxis.label.set_fontproperties(KOREAN_FONT_PROP)

    for label in ax.get_xticklabels():
        label.set_fontproperties(KOREAN_FONT_PROP)

    for label in ax.get_yticklabels():
        label.set_fontproperties(KOREAN_FONT_PROP)

    legend = ax.get_legend()
    if legend is not None:
        for text in legend.get_texts():
            text.set_fontproperties(KOREAN_FONT_PROP)


set_korean_font()


# =====================================================
# CSS
# =====================================================

st.markdown(
    """
<style>
.block-container { padding-top: 1.4rem; padding-bottom: 2rem; }

.kpi-title {
    text-align: center;
    font-size: 34px;
    font-weight: 900;
    margin-bottom: 18px;
}

.metric-card {
    padding: 14px;
    border: 1px solid #e2e6ea;
    border-radius: 14px;
    background: #ffffff;
    min-height: 112px;
    box-shadow: 0 1px 3px rgba(0,0,0,.04);
}

.metric-title {
    font-size: 15px;
    font-weight: 900;
    color: #222;
    margin-bottom: 4px;
    text-align: center;
}

.metric-value {
    font-size: 16px;
    font-weight: 900;
    color: #111;
    text-align: center;
}

.metric-unit {
    display: block;
    text-align: center;
    font-size: 13px;
    color: #666;
    margin-top: 2px;
}

.metric-count {
    text-align: center;
    font-size: 18px;
    font-weight: 900;
    color: #111;
    margin-top: 2px;
}

.kpi-section-title {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-top: 28px;
    margin-bottom: 10px;
}

.kpi-section-title h3 {
    margin: 0;
    font-size: 22px;
    font-weight: 900;
}

.kpi-badge {
    display: inline-block;
    padding: 5px 12px;
    border-radius: 999px;
    background: #f1f3f5;
    color: #222;
    font-size: 14px;
    font-weight: 900;
    border: 1px solid #d9dee3;
}

.safe-table-wrap {
    width: 100%;
    overflow-x: auto;
    border: 1px solid #dee2e6;
    border-radius: 8px;
    margin-bottom: 14px;
}

.safe-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
}

.safe-table th {
    background: #f1f3f5;
    color: #222;
    font-weight: 900;
    border: 1px solid #dee2e6;
    padding: 7px;
    text-align: left;
    white-space: nowrap;
}

.safe-table td {
    border: 1px solid #dee2e6;
    padding: 6px;
    vertical-align: top;
    word-break: break-word;
    max-width: 420px;
}

.safe-table tr:nth-child(even) {
    background: #fafafa;
}

.safe-link-button {
    display: block;
    text-align: center;
    padding: 0.55rem 0.8rem;
    border: 1px solid #d9dee3;
    border-radius: 0.5rem;
    text-decoration: none !important;
    font-weight: 800;
    color: #222 !important;
    background: #f8f9fa;
}
</style>
""",
    unsafe_allow_html=True,
)


# =====================================================
# 설정값 로드
# =====================================================

def get_config_value(key: str, default: str = "") -> str:
    env_val = os.getenv(key)

    if env_val not in (None, ""):
        return str(env_val).strip()

    try:
        secret_val = st.secrets.get(key, None)
        if secret_val not in (None, ""):
            return str(secret_val).strip()
    except Exception:
        pass

    return str(default).strip()


DS_OSS_API_BASE_URL = get_config_value("DS_OSS_API_BASE_URL", DEFAULT_API_BASE_URL).rstrip("/")
DS_OSS_API_KEY = get_config_value("DS_OSS_API_KEY", "")
DS_OSS_USER_ID = get_config_value("DS_OSS_USER_ID", DEFAULT_USER_ID)
DS_OSS_MODEL = get_config_value("DS_OSS_MODEL", DEFAULT_MODEL)
DS_OSS_MAX_TOKENS = int(get_config_value("DS_OSS_MAX_TOKENS", str(DEFAULT_MAX_TOKENS)))
DS_OSS_TIMEOUT = int(get_config_value("DS_OSS_TIMEOUT", str(DEFAULT_TIMEOUT)))
DS_OSS_RETRY = int(get_config_value("DS_OSS_RETRY", str(DEFAULT_RETRY)))


# =====================================================
# 공통 유틸
# =====================================================

def rerun_app():
    try:
        st.rerun()
    except Exception:
        st.experimental_rerun()


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.replace("\n", "", regex=False)
        .str.replace("\r", "", regex=False)
        .str.strip()
    )
    return df


def normalize_col_key(value) -> str:
    return (
        str(value)
        .replace("\xa0", "")
        .replace(" ", "")
        .replace("\n", "")
        .replace("\r", "")
        .strip()
    )


def add_month_num(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "월번호" in df.columns:
        df["월번호"] = pd.to_numeric(df["월번호"], errors="coerce").astype("Int64")
        return df

    if "월" not in df.columns:
        df["월번호"] = pd.NA
        return df

    df["월번호"] = (
        df["월"]
        .astype(str)
        .str.replace("월", "", regex=False)
        .str.strip()
        .replace("", np.nan)
    )

    df["월번호"] = pd.to_numeric(df["월번호"], errors="coerce").astype("Int64")
    return df


def clean_text_series(s: pd.Series) -> pd.Series:
    return (
        s.fillna("")
        .astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )


def first_existing_col(df: pd.DataFrame, candidates: list):
    if df is None or df.empty:
        return None

    for col in candidates:
        if col in df.columns:
            return col

    return None


def find_col_fuzzy(df: pd.DataFrame, candidates: list):
    if df is None or df.empty:
        return None

    col_map = {normalize_col_key(c): c for c in df.columns}

    for cand in candidates:
        key = normalize_col_key(cand)
        if key in col_map:
            return col_map[key]

    for cand in candidates:
        key = normalize_col_key(cand)

        for norm_col, original_col in col_map.items():
            if key and key in norm_col:
                return original_col

    return None


def read_excel_sheet(path: str, preferred_names: list, fallback_index: int):
    xls = pd.ExcelFile(path)
    sheet_map = {normalize_col_key(s): s for s in xls.sheet_names}

    for name in preferred_names:
        key = normalize_col_key(name)
        if key in sheet_map:
            return normalize_columns(pd.read_excel(path, sheet_name=sheet_map[key]))

    return normalize_columns(pd.read_excel(path, sheet_name=fallback_index))


def get_goal_actual_cols(df: pd.DataFrame):
    goal_candidates = [
        "26년 목표",
        "26년목표",
        "2026년 목표",
        "2026년목표",
        f"{KPI_YEAR}년 목표",
        f"{KPI_YEAR}년목표",
        f"{CURRENT_YEAR}년 목표",
        f"{CURRENT_YEAR}년목표",
        "목표",
        "목표건수",
        "목표 건수",
        "목표(건)",
        "목표 건",
    ]

    actual_candidates = [
        "26년 실적",
        "26년실적",
        "2026년 실적",
        "2026년실적",
        f"{KPI_YEAR}년 실적",
        f"{KPI_YEAR}년실적",
        f"{CURRENT_YEAR}년 실적",
        f"{CURRENT_YEAR}년실적",
        "실적",
        "실적건수",
        "실적 건수",
        "실적(건)",
        "실적 건",
    ]

    goal_col = find_col_fuzzy(df, goal_candidates)
    actual_col = find_col_fuzzy(df, actual_candidates)

    return goal_col, actual_col


def get_sheet2_goal_actual_cols(df: pd.DataFrame):
    goal_col = find_col_fuzzy(
        df,
        [
            "25년",
            "25 년",
            "2025년",
            "25년 실적",
            "25년실적",
            "2025년 실적",
            "2025년실적",
        ],
    )

    actual_col = find_col_fuzzy(
        df,
        [
            "26년",
            "26 년",
            "2026년",
            "26년 실적",
            "26년실적",
            "2026년 실적",
            "2026년실적",
        ],
    )

    return goal_col, actual_col


def to_numeric_clean(value):
    if pd.isna(value):
        return np.nan

    try:
        text = str(value).strip()

        if text == "":
            return np.nan

        text = (
            text.replace(",", "")
            .replace("%", "")
            .replace("건/Mm2", "")
            .replace("건", "")
            .replace("명", "")
            .replace("만인률", "")
            .replace("만인율", "")
            .strip()
        )

        return pd.to_numeric(text, errors="coerce")

    except Exception:
        return np.nan


def fmt_int(value, default: str = "-") -> str:
    try:
        if pd.isna(value):
            return default

        return f"{int(float(value))}"

    except Exception:
        return default


def fmt_float(value, digits: int = 2, default: str = "-") -> str:
    try:
        if pd.isna(value):
            return default

        return f"{float(value):.{digits}f}"

    except Exception:
        return default


def fmt_no_decimal(value, default: str = "-") -> str:
    try:
        if pd.isna(value):
            return default

        return f"{float(value):.0f}"

    except Exception:
        return default


def img_to_base64(path: str) -> str:
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()

    except Exception:
        return ""


def signal_image(actual, target) -> str:
    green_img = "/config/work/sharedworkspace/초록신호등.png"
    orange_img = "/config/work/sharedworkspace/주황신호등.png"
    red_img = "/config/work/sharedworkspace/빨간신호등.png"

    if pd.isna(actual) or pd.isna(target):
        return orange_img

    actual = float(actual)
    target = float(target)

    gap = abs(actual - target)

    if gap <= 0.2000001:
        return orange_img

    if actual > target:
        return red_img

    return green_img


def status_text(actual, target) -> str:
    if pd.isna(actual) or pd.isna(target):
        return "확인필요"

    actual = float(actual)
    target = float(target)

    if abs(actual - target) <= 0.2000001:
        return "주의"

    if actual > target:
        return "미달성"

    return "달성"


def make_kpi_card_text(actual, target, status):
    return (
        f"{CURRENT_MONTH}월 실적 {fmt_no_decimal(actual)} / 목표 {fmt_no_decimal(target)} / {status}"
        "<span class='metric-unit'>(건)</span>"
    )


def df_to_csv_text(df: pd.DataFrame, max_rows: int = 30) -> str:
    if df is None or df.empty:
        return "데이터 없음"

    out = df.head(max_rows).copy()
    out = out.astype(object).where(pd.notna(out), "")

    for col in out.columns:
        out[col] = out[col].astype(str)

    return out.to_csv(index=False)


def save_history(title: str, prompt: str, result: str):
    history = []

    if HISTORY_FILE.exists():
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                history = json.load(f)
        except Exception:
            history = []

    history.append(
        {
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "title": title,
            "prompt": prompt,
            "result": result,
        }
    )

    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


# =====================================================
# 데이터 로드
# =====================================================

@st.cache_data(show_spinner=False)
def load_kpi_data(path: str):
    df_high = add_month_num(
        read_excel_sheet(path, ["고위험", "고위험사고", "Sheet0"], 0)
    )

    df_grade = add_month_num(
        read_excel_sheet(path, ["등급", "등급사고", "Sheet1"], 1)
    )

    df_leak = add_month_num(
        read_excel_sheet(path, ["LEAK", "Leak", "누출", "Sheet2"], 2)
    )

    df_ert = add_month_num(
        read_excel_sheet(path, ["ERT", "ERT고등급", "ERT출동", "Sheet3"], 3)
    )

    return df_high, df_grade, df_leak, df_ert


@st.cache_data(show_spinner=False)
def load_accident_data(path: str):
    df_main = normalize_columns(pd.read_excel(path, sheet_name=0))

    try:
        df_ert = normalize_columns(pd.read_excel(path, sheet_name="ERT출동"))
    except Exception:
        df_ert = pd.DataFrame()

    return df_main, df_ert


# =====================================================
# KPI Series
# =====================================================

def safe_series(df: pd.DataFrame, col: str) -> list:
    vals = [np.nan] * 12

    if df is None or df.empty or not col:
        return vals

    if "월번호" not in df.columns:
        df = add_month_num(df)

    if "월번호" not in df.columns or col not in df.columns:
        return vals

    d = df.copy()

    if "연도" in d.columns:
        year_num = pd.to_numeric(d["연도"], errors="coerce")
        d_year = d[(year_num == KPI_YEAR) | (year_num == CURRENT_YEAR)]

        if not d_year.empty:
            d = d_year

    for m in range(1, 13):
        v = d.loc[d["월번호"] == m, col]

        if not v.empty:
            vals[m - 1] = to_numeric_clean(v.iloc[0])

    return vals


def get_ltir_series(df_grade: pd.DataFrame):
    if df_grade is None or df_grade.empty:
        return [np.nan] * 12, None

    d = df_grade.copy()

    if "월번호" not in d.columns:
        d = add_month_num(d)

    ltir_col = None

    for col in d.columns:
        if normalize_col_key(col) == normalize_col_key("LTIR 0.020 만인률"):
            ltir_col = col
            break

    if ltir_col is None:
        for col in d.columns:
            if normalize_col_key(col) == normalize_col_key("LTIR 0.020 만인율"):
                ltir_col = col
                break

    if ltir_col is None:
        ltir_col = find_col_fuzzy(
            d,
            [
                "LTIR 0.020 만인률",
                "LTIR0.020만인률",
                "LTIR 0.020 만인율",
                "LTIR0.020만인율",
                "LTIR",
                "만인률",
                "만인율",
            ],
        )

    if ltir_col is None:
        return [np.nan] * 12, None

    vals = [np.nan] * 12

    if "월번호" not in d.columns:
        return vals, ltir_col

    for m in range(1, 13):
        row = d[d["월번호"] == m]

        if not row.empty:
            vals[m - 1] = to_numeric_clean(row.iloc[0][ltir_col])

    return vals, ltir_col


def current_value(df: pd.DataFrame, target_col: str, actual_col: str):
    if df is None or df.empty:
        return np.nan, np.nan

    if not target_col or not actual_col:
        return np.nan, np.nan

    if "월번호" not in df.columns:
        df = add_month_num(df)

    if "월번호" not in df.columns or target_col not in df.columns or actual_col not in df.columns:
        return np.nan, np.nan

    d = df.copy()

    if "연도" in d.columns:
        year_num = pd.to_numeric(d["연도"], errors="coerce")
        d_year = d[(year_num == KPI_YEAR) | (year_num == CURRENT_YEAR)]

        if not d_year.empty:
            d = d_year

    row = d[d["월번호"] == CURRENT_MONTH]

    if row.empty:
        return np.nan, np.nan

    row = row.iloc[0]

    return (
        to_numeric_clean(row[target_col]),
        to_numeric_clean(row[actual_col]),
    )


def get_month_target_count(df: pd.DataFrame, fallback_col: str):
    if df is None or df.empty:
        return np.nan

    if "월번호" not in df.columns:
        df = add_month_num(df)

    if "월번호" not in df.columns:
        return np.nan

    d = df.copy()

    if "연도" in d.columns:
        year_num = pd.to_numeric(d["연도"], errors="coerce")
        d_year = d[(year_num == KPI_YEAR) | (year_num == CURRENT_YEAR)]

        if not d_year.empty:
            d = d_year

    row = d[d["월번호"] == CURRENT_MONTH]

    if row.empty:
        return np.nan

    row = row.iloc[0]

    candidates = [
        "26년 목표건수",
        "26년 목표 건수",
        "26년 목표(건)",
        "26년 목표 건",
        "26년 목표",
        "26년목표",
        "2026년 목표건수",
        "2026년 목표 건수",
        "2026년 목표(건)",
        "2026년 목표 건",
        "2026년 목표",
        "2026년목표",
        "25년",
        "2025년",
        "목표건수",
        "목표 건수",
        "목표(건)",
        "목표 건",
        "건수목표",
        "건수 목표",
        "목표",
        fallback_col,
    ]

    col = find_col_fuzzy(df, candidates)

    if col and col in df.columns:
        val = to_numeric_clean(row[col])

        if pd.notna(val):
            return val

    return np.nan


# =====================================================
# 차트
# =====================================================

def style_chart(ax):
    ax.yaxis.tick_right()
    ax.yaxis.set_label_position("right")

    ax.spines["top"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["right"].set_visible(False)

    ax.grid(axis="y", linestyle="--", alpha=0.3)

    ax.tick_params(axis="x", labelsize=GRAPH_TICK_SIZE)
    ax.tick_params(axis="y", labelsize=GRAPH_TICK_SIZE)

    apply_korean_font_to_axes(ax)


def draw_bar_chart(
    title: str,
    y_goal: list,
    y_actual: list,
    goal_label: str = "목표",
    actual_label: str = "실적",
    line_values=None,
    line_label: str = "",
    value_digits: int = 0,
):
    labels = list(range(1, 13))
    x = np.arange(12)
    width = 0.35
    month_idx = CURRENT_MONTH - 1

    y_goal = pd.to_numeric(pd.Series(y_goal), errors="coerce").tolist()
    y_actual = pd.to_numeric(pd.Series(y_actual), errors="coerce").tolist()

    fig, ax = plt.subplots(figsize=(4.2, 3.6))

    ax.bar(x - width / 2, y_goal, width, label=goal_label, color="#4C72B0")
    ax.bar(x + width / 2, y_actual, width, label=actual_label, color="#55A868")

    ax.set_title(
        title,
        fontweight="bold",
        fontsize=GRAPH_TITLE_SIZE,
        fontproperties=KOREAN_FONT_PROP,
    )

    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=GRAPH_TICK_SIZE, fontproperties=KOREAN_FONT_PROP)

    style_chart(ax)

    goal_now = y_goal[month_idx] if month_idx < len(y_goal) else np.nan
    actual_now = y_actual[month_idx] if month_idx < len(y_actual) else np.nan

    if pd.notna(goal_now):
        ax.text(
            month_idx - width / 2,
            goal_now,
            f"{goal_now:.{value_digits}f}",
            ha="center",
            va="bottom",
            fontweight="bold",
            fontsize=GRAPH_TEXT_SIZE,
            fontproperties=KOREAN_FONT_PROP,
        )

    if pd.notna(actual_now):
        ax.text(
            month_idx + width / 2,
            actual_now,
            f"{actual_now:.{value_digits}f}",
            ha="center",
            va="bottom",
            fontweight="bold",
            fontsize=GRAPH_TEXT_SIZE,
            fontproperties=KOREAN_FONT_PROP,
        )

    if line_values is not None:
        ax2 = ax.twinx()
        line_clean = pd.to_numeric(pd.Series(line_values), errors="coerce").tolist()

        ax2.plot(
            x,
            line_clean,
            color="lightcoral",
            marker="o",
            linewidth=1.8,
            alpha=0.9,
            label=line_label or "LTIR",
        )

        ax2.spines["top"].set_visible(False)
        ax2.spines["left"].set_visible(False)
        ax2.spines["right"].set_visible(False)
        ax2.tick_params(axis="y", labelsize=GRAPH_TICK_SIZE)
        apply_korean_font_to_axes(ax2)

        line_now = line_clean[month_idx] if month_idx < len(line_clean) else np.nan

        if pd.notna(line_now):
            ax2.text(
                month_idx,
                line_now,
                f"{line_now:.3f}",
                color="red",
                ha="center",
                va="bottom",
                fontweight="bold",
                fontsize=GRAPH_TEXT_SIZE,
                fontproperties=KOREAN_FONT_PROP,
            )

        h1, l1 = ax.get_legend_handles_labels()
        h2, l2 = ax2.get_legend_handles_labels()

        ax.legend(
            h1 + h2,
            l1 + l2,
            frameon=False,
            fontsize=GRAPH_LEGEND_SIZE,
            prop=KOREAN_FONT_PROP,
        )

    else:
        ax.legend(
            frameon=False,
            fontsize=GRAPH_LEGEND_SIZE,
            prop=KOREAN_FONT_PROP,
        )

    apply_korean_font_to_axes(ax)

    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)


def make_monthly_count(df: pd.DataFrame, date_col: str = "발생일", years: int = 3) -> pd.DataFrame:
    if df is None or df.empty or date_col not in df.columns:
        return pd.DataFrame(columns=["년월", "건수"])

    out = df.copy()
    out[date_col] = pd.to_datetime(out[date_col], errors="coerce")
    out = out.dropna(subset=[date_col])

    min_year = CURRENT_YEAR - years + 1
    out = out[out[date_col].dt.year.between(min_year, CURRENT_YEAR)]

    if out.empty:
        return pd.DataFrame(columns=["년월", "건수"])

    out["년월"] = out[date_col].dt.strftime("%Y-%m")

    return (
        out.groupby("년월")
        .size()
        .reset_index(name="건수")
        .sort_values("년월")
    )


def make_top_count(df: pd.DataFrame, col: str, top_n: int = 5) -> pd.DataFrame:
    if df is None or df.empty or not col or col not in df.columns:
        return pd.DataFrame(columns=[col or "항목", "건수"])

    out = df.copy()
    out[col] = clean_text_series(out[col])
    out = out[out[col] != ""]

    if out.empty:
        return pd.DataFrame(columns=[col, "건수"])

    return (
        out.groupby(col)
        .size()
        .reset_index(name="건수")
        .sort_values("건수", ascending=False)
        .head(top_n)
    )


def draw_count_line_chart(title: str, df: pd.DataFrame, x_col: str = "년월", y_col: str = "건수"):
    if df is None or df.empty or x_col not in df.columns or y_col not in df.columns:
        st.info(f"{title}: 표시할 데이터가 없습니다.")
        return

    fig, ax = plt.subplots(figsize=(7, 3.2))

    ax.plot(
        df[x_col],
        df[y_col],
        marker="o",
        linewidth=2.0,
        color="#4C72B0",
    )

    ax.set_title(title, fontweight="bold", fontsize=GRAPH_TITLE_SIZE, fontproperties=KOREAN_FONT_PROP)
    ax.set_xlabel("년월", fontsize=GRAPH_LABEL_SIZE, fontproperties=KOREAN_FONT_PROP)
    ax.set_ylabel("건수", fontsize=GRAPH_LABEL_SIZE, fontproperties=KOREAN_FONT_PROP)
    ax.tick_params(axis="x", labelrotation=45, labelsize=GRAPH_TICK_SIZE)
    ax.tick_params(axis="y", labelsize=GRAPH_TICK_SIZE)

    for x_val, y_val in zip(df[x_col], df[y_col]):
        ax.text(
            x_val,
            y_val,
            str(int(y_val)),
            ha="center",
            va="bottom",
            fontsize=GRAPH_TEXT_SIZE,
            fontweight="bold",
            fontproperties=KOREAN_FONT_PROP,
        )

    style_chart(ax)
    apply_korean_font_to_axes(ax)

    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)


def draw_top_bar_chart(title: str, df: pd.DataFrame, name_col: str, value_col: str = "건수"):
    if df is None or df.empty or name_col not in df.columns or value_col not in df.columns:
        st.info(f"{title}: 표시할 데이터가 없습니다.")
        return

    plot_df = df.sort_values(value_col, ascending=True)

    fig, ax = plt.subplots(figsize=(5.8, 3.2))

    ax.barh(
        plot_df[name_col],
        plot_df[value_col],
        color="#55A868",
    )

    ax.set_title(title, fontweight="bold", fontsize=GRAPH_TITLE_SIZE, fontproperties=KOREAN_FONT_PROP)
    ax.set_xlabel("건수", fontsize=GRAPH_LABEL_SIZE, fontproperties=KOREAN_FONT_PROP)
    ax.set_ylabel("", fontsize=GRAPH_LABEL_SIZE, fontproperties=KOREAN_FONT_PROP)

    for i, v in enumerate(plot_df[value_col]):
        ax.text(
            v,
            i,
            f" {int(v)}",
            va="center",
            fontweight="bold",
            fontsize=GRAPH_TEXT_SIZE,
            fontproperties=KOREAN_FONT_PROP,
        )

    style_chart(ax)
    apply_korean_font_to_axes(ax)

    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)


# =====================================================
# 안전 HTML Table
# =====================================================

def render_safe_html_table(df: pd.DataFrame, max_rows: int = 300):
    if df is None or df.empty:
        st.info("표시할 데이터가 없습니다.")
        return

    out = df.copy().head(max_rows)
    out = out.astype(object).where(pd.notna(out), "")

    for col in out.columns:
        out[col] = out[col].astype(str)

    html = out.to_html(
        index=False,
        escape=True,
        border=0,
        classes="safe-table",
    )

    st.markdown(
        f"""
<div class="safe-table-wrap">
{html}
</div>
""",
        unsafe_allow_html=True,
    )


# =====================================================
# Sidebar
# =====================================================

def render_sidebar():
    with st.sidebar:
        st.markdown("### 사고분석 TOOL")

        st.markdown(
            """
<a href="https://ehs-analyzer--hsacci-prod.cdep1.samsungds.net/"
target="_blank" class="safe-link-button">
사고분석 TOOL 바로가기
</a>
""",
            unsafe_allow_html=True,
        )

        st.markdown("---")
        st.markdown("### 연관사이트")

        dash_icon = img_to_base64("/config/work/sharedworkspace/icons/대시보드.png")
        acc_icon = img_to_base64("/config/work/sharedworkspace/icons/환경안전사고.png")
        eam_icon = img_to_base64("/config/work/sharedworkspace/icons/EAM.png")
        idps_icon = img_to_base64("/config/work/sharedworkspace/icons/IDPS.png")

        st.markdown(
            f"""
<div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
<a href="https://ehsdxdep--hsdash-prod.cdep.samsungds.net/" target="_blank">
<img src="data:image/png;base64,{dash_icon}" style="width:100%"></a>

<a href="https://confluence.samsungds.net/spaces/HSEHS84000/pages/3196138433/26%EB%85%84+%EC%82%AC%EA%B3%A0+%ED%98%84%ED%99%A9" target="_blank">
<img src="data:image/png;base64,{acc_icon}" style="width:100%"></a>

<a href="https://eam.sec.samsung.net:4443/workplace/portal/eamMain.do" target="_blank">
<img src="data:image/png;base64,{eam_icon}" style="width:100%"></a>

<a href="https://idps.samsungds.net:9055/idps/main" target="_blank">
<img src="data:image/png;base64,{idps_icon}" style="width:100%"></a>
</div>
""",
            unsafe_allow_html=True,
        )


# =====================================================
# UI Components
# =====================================================

def render_kpi_card(icon: str, title: str, text: str):
    c1, c2 = st.columns([1, 5])

    with c1:
        try:
            st.image(icon, width=30)
        except Exception:
            st.write("")

    with c2:
        st.markdown(
            f"""
<div class="metric-card">
    <div class="metric-title">{title}</div>
    <div class="metric-value">{text}</div>
</div>
""",
            unsafe_allow_html=True,
        )


def render_section_header(title: str, count=None, target=None, unit: str = "건"):
    if count is None:
        badge = ""
    elif target is None or pd.isna(target):
        badge = f"<span class='kpi-badge'>실적 {count}{unit}</span>"
    else:
        badge = f"<span class='kpi-badge'>실적 {count}{unit} / 목표 {fmt_int(target)}{unit}</span>"

    st.markdown(
        f"""
<div class="kpi-section-title">
    <h3>{title}</h3>
    {badge}
</div>
""",
        unsafe_allow_html=True,
    )


def show_table(df_view: pd.DataFrame, key: str):
    render_safe_html_table(df_view, max_rows=300)


def prepare_accident_table(df_view: pd.DataFrame, columns: list) -> pd.DataFrame:
    if df_view is None or df_view.empty:
        return pd.DataFrame(columns=[])

    out = df_view.copy()

    if "발생일" in out.columns:
        out["발생일"] = pd.to_datetime(out["발생일"], errors="coerce")
        out = out.sort_values("발생일", ascending=False)
        out["발생일"] = out["발생일"].dt.strftime("%Y-%m-%d").fillna("")

    if "컨플링크" in out.columns:
        out["컨플링크"] = out["컨플링크"].fillna("").astype(str)

    exist_cols = [c for c in columns if c in out.columns]
    out = out[exist_cols].copy()
    out = out.astype(object).where(pd.notna(out), "")

    for col in out.columns:
        out[col] = out[col].astype(str)

    return out


# =====================================================
# DS OSS
# =====================================================

def make_client():
    if OpenAI is None:
        raise ImportError("openai 패키지가 없습니다. python -m pip install openai 실행 필요")

    if not DS_OSS_API_KEY:
        raise ValueError("DS_OSS_API_KEY가 비어 있습니다.")

    headers = {
        "x-dep-ticket": DS_OSS_API_KEY,
        "Send-System-Name": "playground",
        "User-Id": DS_OSS_USER_ID,
        "User-Type": "AD_ID",
        "Prompt-Msg-Id": str(uuid.uuid4()),
        "Completion-Msg-Id": str(uuid.uuid4()),
    }

    return OpenAI(
        api_key="dummy",
        base_url=DS_OSS_API_BASE_URL,
        default_headers=headers,
        timeout=DS_OSS_TIMEOUT,
        max_retries=DS_OSS_RETRY,
    )


def call_ds_oss(prompt: str, system_prompt: str = "", temperature: float = 0.2) -> str:
    last_error = None

    for _ in range(DS_OSS_RETRY + 1):
        try:
            client = make_client()
            messages = []

            if system_prompt.strip():
                messages.append({"role": "system", "content": system_prompt.strip()})

            messages.append({"role": "user", "content": prompt})

            response = client.chat.completions.create(
                model=DS_OSS_MODEL,
                messages=messages,
                temperature=temperature,
                max_tokens=DS_OSS_MAX_TOKENS,
            )

            return response.choices[0].message.content

        except Exception as e:
            last_error = e

    raise RuntimeError(f"DS OSS API 호출 실패: {last_error}")


# =====================================================
# AI 분석용 데이터/그래프
# =====================================================

def build_ai_analysis_frame(
    analysis_type: str,
    df_acc: pd.DataFrame,
    df_ert_list: pd.DataFrame,
):
    if analysis_type == "ERT 고등급 출동 분석":
        if df_ert_list is None or df_ert_list.empty:
            return pd.DataFrame(), "발생일", "담당부서", "건물"

        out = df_ert_list.copy()

        if "RI등급" in out.columns:
            out["RI등급"] = clean_text_series(out["RI등급"]).str.upper().str.replace(" ", "", regex=False)
            out = out[out["RI등급"].isin(["A", "B", "C+"])].copy()

        date_col = first_existing_col(out, ["발생일", "등록일", "출동일"]) or "발생일"
        dept_col = first_existing_col(out, ["담당부서", "부서명", "부서"])
        building_col = first_existing_col(out, ["건물", "건물명"])

        return out, date_col, dept_col, building_col

    if df_acc is None or df_acc.empty:
        return pd.DataFrame(), "발생일", "부서명", "건물명"

    out = df_acc.copy()

    if analysis_type == "고위험 및 재발사고 TREND":
        masks = []

        if {"변경 등급", "고위험/일반/생활사고"}.issubset(out.columns):
            masks.append(
                (out["변경 등급"].isin(["등급", "경미"]))
                & (out["고위험/일반/생활사고"] == "고위험")
            )

        if "재발사고" in out.columns:
            masks.append(clean_text_series(out["재발사고"]).str.contains("재발", na=False))

        if masks:
            mask = masks[0]

            for m in masks[1:]:
                mask = mask | m

            out = out[mask].copy()

    elif analysis_type == "LEAK 사고 분석":
        if "재해유형" in out.columns:
            out = out[out["재해유형"] == "누출/접촉/흡입"].copy()

        if "변경 등급" in out.columns:
            out = out[out["변경 등급"].isin(["경미", "등급"])].copy()

    date_col = first_existing_col(out, ["발생일", "등록일", "사고일"]) or "발생일"
    dept_col = first_existing_col(out, ["부서명", "담당부서", "부서"])
    building_col = first_existing_col(out, ["건물명", "건물", "장소"])

    return out, date_col, dept_col, building_col


def get_ai_graph_data(
    analysis_type: str,
    df_acc: pd.DataFrame,
    df_ert_list: pd.DataFrame,
):
    target_df, date_col, dept_col, building_col = build_ai_analysis_frame(
        analysis_type,
        df_acc,
        df_ert_list,
    )

    if target_df is None or target_df.empty:
        return target_df, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), dept_col, building_col

    monthly_df = make_monthly_count(target_df, date_col=date_col, years=3)
    dept_top = make_top_count(target_df, dept_col, 5) if dept_col else pd.DataFrame()
    building_top = make_top_count(target_df, building_col, 5) if building_col else pd.DataFrame()

    return target_df, monthly_df, dept_top, building_top, dept_col, building_col


def render_ai_charts(
    analysis_type: str,
    df_acc: pd.DataFrame,
    df_ert_list: pd.DataFrame,
):
    target_df, monthly_df, dept_top, building_top, dept_col, building_col = get_ai_graph_data(
        analysis_type,
        df_acc,
        df_ert_list,
    )

    if target_df is None or target_df.empty:
        st.info("선택한 분석 유형에 해당하는 데이터가 없습니다.")
        return target_df, monthly_df, dept_top, building_top

    st.markdown("#### 📊 AI 분석용 그래프")

    g1, g2, g3 = st.columns([1.5, 1, 1])

    with g1:
        draw_count_line_chart(f"{analysis_type} 월별 Trend", monthly_df)

    with g2:
        if dept_col:
            draw_top_bar_chart("다발 부서명 TOP5", dept_top, dept_col)
        else:
            st.info("부서명 컬럼이 없습니다.")

    with g3:
        if building_col:
            draw_top_bar_chart("다발 건물명 TOP5", building_top, building_col)
        else:
            st.info("건물명 컬럼이 없습니다.")

    return target_df, monthly_df, dept_top, building_top


def render_ai_panel(
    high_table: pd.DataFrame,
    repeat_table: pd.DataFrame,
    grade_table: pd.DataFrame,
    leak_table: pd.DataFrame,
    ert_table: pd.DataFrame,
    summary: dict,
    df_acc: pd.DataFrame,
    df_ert_list: pd.DataFrame,
):
    st.markdown("---")
    st.markdown("### 🤖 DS OSS AI 사고 KPI 분석")

    analysis_type = st.selectbox(
        "AI 분석 메뉴",
        [
            "고위험 및 재발사고 TREND",
            "LEAK 사고 분석",
            "ERT 고등급 출동 분석",
            "월별 사고 예측(3개년 데이터기반)",
            "사고 경향성 그래프 분석",
        ],
        key="ds_oss_analysis_type",
    )

    target_df, monthly_df, dept_top, building_top, dept_col, building_col = get_ai_graph_data(
        analysis_type,
        df_acc,
        df_ert_list,
    )

    if st.button("AI 분석용 그래프 보기", key="show_ai_graph_btn"):
        render_ai_charts(
            analysis_type,
            df_acc,
            df_ert_list,
        )

    with st.expander("DS OSS AI 분석 실행", expanded=False):
        user_extra_prompt = st.text_area(
            "추가 분석 지시사항",
            placeholder="예: 전월 대비 악화 요인과 즉시 실행 가능한 Action Item 중심으로 정리해줘.",
            height=100,
            key="ds_oss_user_extra_prompt",
        )

        if not DS_OSS_API_KEY:
            st.info(
                "DS OSS API Key가 설정되지 않았습니다. AI 분석을 사용하려면 DS_OSS_API_KEY를 환경변수 또는 secrets.toml에 설정하세요."
            )

        menu_instruction_map = {
            "고위험 및 재발사고 TREND": """
분석 초점:
- 고위험사고와 재발사고의 월별 증감 추세를 해석한다.
- 반복 발생 부서와 건물의 공통 원인을 도출한다.
- SOP, 작업 전 위험성 검토, 안전조치 이행 기준 관점의 Action Item을 제시한다.
""",
            "LEAK 사고 분석": """
분석 초점:
- LEAK 사고의 월별 추세와 다발 부서/건물을 해석한다.
- 배관, 밸브, 체결부, 잔압, 잔류 약액, 중성화, 철거/교체 작업 관점으로 원인을 분석한다.
- 반복 발생 가능성이 높은 설비/공정 관리 포인트를 제시한다.
""",
            "ERT 고등급 출동 분석": """
분석 초점:
- RI A/B/C+ 고등급 출동 추세를 해석한다.
- 다발 부서와 건물 기준으로 비상대응 취약 지점을 도출한다.
- 출동 빈도 저감, 초기 대응, 현장 통제, 재발 방지 관점의 Action Item을 제시한다.
""",
            "월별 사고 예측(3개년 데이터기반)": """
분석 초점:
- 최근 3개년 월별 사고 건수를 기준으로 계절성, 반복성, 증가/감소 경향을 분석한다.
- 다음 월 사고 발생 가능성을 정성적으로 예측한다.
- 예측 근거와 선제 관리 포인트를 제시한다.
""",
            "사고 경향성 그래프 분석": """
분석 초점:
- 월별 Trend, 다발 부서, 다발 건물 그래프를 종합 해석한다.
- 증가/감소 경향과 특정 조직/장소 쏠림 현상을 설명한다.
- 관리자가 바로 볼 수 있는 그래프 기반 보고 문안을 작성한다.
""",
        }

        menu_instruction = menu_instruction_map.get(analysis_type, "")

        if st.button(
            "DS OSS AI 분석 실행",
            type="primary",
            key="run_ds_oss_analysis",
        ):
            if not DS_OSS_API_KEY:
                st.error("DS_OSS_API_KEY가 설정되지 않았습니다.")
                return

            actual_high, target_high, status_high = summary.get("고위험사고", (np.nan, np.nan, ""))
            actual_grade, target_grade, status_grade = summary.get("등급사고", (np.nan, np.nan, ""))
            actual_leak, target_leak, status_leak = summary.get("LEAK 경미이상", (np.nan, np.nan, ""))
            actual_ert, target_ert, status_ert = summary.get("ERT 고등급 출동", (np.nan, np.nan, ""))

            prompt = f"""
다음은 {CURRENT_YEAR}년 화성사업장 사고 KPI 데이터다.

[선택 분석 메뉴]
{analysis_type}

[분석 메뉴별 세부 지시사항]
{menu_instruction}

[KPI 현황]
- 고위험사고: 실적 {actual_high}, 목표 {target_high}, 상태 {status_high}
- 등급사고: 실적 {actual_grade}, 목표 {target_grade}, 상태 {status_grade}
- LEAK 경미이상: 실적 {actual_leak}, 목표 {target_leak}, 상태 {status_leak}
- ERT 고등급 출동: 실적 {actual_ert}, 목표 {target_ert}, 상태 {status_ert}, 총 {summary.get('ERT 총 건수', 0)}건
- 재발사고 건수: {summary.get('재발사고 건수', 0)}건

[선택 메뉴 기준 월별 Trend]
{df_to_csv_text(monthly_df, max_rows=50)}

[선택 메뉴 기준 다발 부서명 TOP5]
{df_to_csv_text(dept_top, max_rows=10)}

[선택 메뉴 기준 다발 건물명 TOP5]
{df_to_csv_text(building_top, max_rows=10)}

[고위험사고 LIST]
{df_to_csv_text(high_table)}

[재발사고 LIST]
{df_to_csv_text(repeat_table)}

[등급사고 LIST]
{df_to_csv_text(grade_table)}

[LEAK 경미이상 LIST]
{df_to_csv_text(leak_table)}

[ERT 고등급 출동 LIST]
{df_to_csv_text(ert_table)}

[사용자 추가 지시사항]
{user_extra_prompt}

아래 형식으로 작성하라.

■ 0. 핵심 결론
- 3줄 이내

■ 1. 그래프 기반 경향성 해석
- 월별 Trend
- 다발 부서명
- 다발 건물명

■ 2. 주요 리스크
- 고위험/재발/LEAK/ERT 관점

■ 3. 원인 추정 및 추가 확인 필요사항
- 데이터상 확인되는 원인
- 추가 확인 필요사항

■ 4. Action Item
- 즉시 조치
- 단기 조치
- 월간 Tracking 항목

■ 5. 보고용 문안
- 파트장/그룹장 보고용으로 바로 붙여넣을 수 있는 문장
"""

            system_prompt = """
당신은 삼성전자 DS부문 화성사업장 EHS 사고 KPI 분석 전문가다.
답변은 한국어 보고서 문체로 작성한다.
사실과 추정을 구분한다.
정량 데이터 기반으로 간결하게 작성한다.
불필요한 미사여구는 제외한다.
"""

            with st.spinner("DS OSS API로 분석 중..."):
                try:
                    result = call_ds_oss(
                        prompt=prompt,
                        system_prompt=system_prompt,
                        temperature=0.2,
                    )

                    st.success("DS OSS AI 분석 완료")
                    st.markdown(result)

                    save_history(analysis_type, prompt, result)

                    st.download_button(
                        label="AI 분석 결과 다운로드",
                        data=result.encode("utf-8"),
                        file_name=f"DS_OSS_AI_분석_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                        mime="text/plain",
                    )

                except Exception as e:
                    st.error(f"DS OSS API 호출 중 오류 발생: {e}")


# =====================================================
# Main KPI Page
# =====================================================

def render_kpi_page():
    st.markdown(
        f"<div class='kpi-title'>🚧 {CURRENT_YEAR} 화성사업장 {CURRENT_MONTH}월 ({WEEK_LABEL}) 사고 KPI</div>",
        unsafe_allow_html=True,
    )

    refresh_cols = st.columns([1, 5])

    with refresh_cols[0]:
        if st.button("🔄 KPI 새로고침"):
            st.cache_data.clear()
            rerun_app()

    if KOREAN_FONT_NAME:
        st.caption(f"그래프 폰트 적용: {KOREAN_FONT_NAME}")
    else:
        st.warning(
            "그래프 한글 폰트를 찾지 못했습니다. "
            "/config/work/sharedworkspace/fonts/ 경로에 NotoSansKR-Regular.ttf 또는 NotoSansKR-VariableFont_wght.ttf 파일을 넣어주세요."
        )

    try:
        df_high, df_grade, df_leak, df_ert_kpi = load_kpi_data(KPI_FILE)
        df_acc, df_ert_list = load_accident_data(ACCIDENT_FILE)

    except Exception as e:
        st.error(f"엑셀 데이터 로드 실패: {e}")
        st.stop()

    high_goal_col, high_actual_col = get_goal_actual_cols(df_high)
    grade_goal_col, grade_actual_col = get_goal_actual_cols(df_grade)
    leak_goal_col, leak_actual_col = get_sheet2_goal_actual_cols(df_leak)
    ert_goal_col, ert_actual_col = get_goal_actual_cols(df_ert_kpi)

    y_ltir, ltir_col = get_ltir_series(df_grade)

    y_goal_high = safe_series(df_high, high_goal_col)
    y_act_high = safe_series(df_high, high_actual_col)

    y_goal_grade = safe_series(df_grade, grade_goal_col)
    y_act_grade = safe_series(df_grade, grade_actual_col)

    y_goal_leak = safe_series(df_leak, leak_goal_col)
    y_act_leak = safe_series(df_leak, leak_actual_col)

    y_goal_ert = safe_series(df_ert_kpi, ert_goal_col)
    y_act_ert = safe_series(df_ert_kpi, ert_actual_col)

    target_high, actual_high = current_value(df_high, high_goal_col, high_actual_col)
    target_grade, actual_grade = current_value(df_grade, grade_goal_col, grade_actual_col)
    target_leak, actual_leak = current_value(df_leak, leak_goal_col, leak_actual_col)
    target_ert, actual_ert = current_value(df_ert_kpi, ert_goal_col, ert_actual_col)

    missing_cols = []

    if not high_goal_col or not high_actual_col:
        missing_cols.append(f"Sheet0 고위험: 목표={high_goal_col}, 실적={high_actual_col}")

    if not grade_goal_col or not grade_actual_col:
        missing_cols.append(f"Sheet1 등급: 목표={grade_goal_col}, 실적={grade_actual_col}")

    if ltir_col is None:
        missing_cols.append("Sheet 등급: LTIR 0.020 만인률 컬럼 없음")

    if not leak_goal_col or not leak_actual_col:
        missing_cols.append(f"Sheet2 LEAK: 목표={leak_goal_col}, 실적={leak_actual_col}")

    if not ert_goal_col or not ert_actual_col:
        missing_cols.append(f"Sheet3 ERT: 목표={ert_goal_col}, 실적={ert_actual_col}")

    if missing_cols:
        st.warning("KPI 컬럼 확인 필요: " + " / ".join(missing_cols))

    with st.expander("🔍 PROD 반영 확인용 디버그", expanded=False):
        st.write("현재 실행 파일:", __file__)
        st.write("KPI_FILE:", KPI_FILE)
        st.write("ACCIDENT_FILE:", ACCIDENT_FILE)

        try:
            st.write("KPI 파일 수정시간:", datetime.fromtimestamp(os.path.getmtime(KPI_FILE)))
        except Exception as e:
            st.write("KPI 파일 수정시간 확인 실패:", e)

        try:
            st.write("KPI Sheet 목록:", pd.ExcelFile(KPI_FILE).sheet_names)
        except Exception as e:
            st.write("Sheet 확인 실패:", e)

        st.write("고위험 KPI 컬럼:", list(df_high.columns))
        st.write("등급 KPI 컬럼:", list(df_grade.columns))
        st.write("LEAK KPI 컬럼:", list(df_leak.columns))
        st.write("ERT KPI 컬럼:", list(df_ert_kpi.columns))

        st.write("ERT 목표 컬럼:", ert_goal_col)
        st.write("ERT 실적 컬럼:", ert_actual_col)
        st.write("현재월:", CURRENT_MONTH)
        st.write("현재월 ERT 목표:", target_ert)
        st.write("현재월 ERT 실적:", actual_ert)
        st.dataframe(df_ert_kpi)

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_high.to_excel(writer, sheet_name="고위험", index=False)
        df_grade.to_excel(writer, sheet_name="등급", index=False)
        df_leak.to_excel(writer, sheet_name="LEAK", index=False)
        df_ert_kpi.to_excel(writer, sheet_name="ERT", index=False)

    st.download_button(
        label="KPI 엑셀 다운로드",
        data=output.getvalue(),
        file_name="KPI_dashboard.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    list_target_high = get_month_target_count(df_high, high_goal_col)
    list_target_grade = get_month_target_count(df_grade, grade_goal_col)
    list_target_leak = get_month_target_count(df_leak, leak_goal_col)

    ert_total_count = 0

    if not df_ert_list.empty and {"RI등급", "발생일"}.issubset(df_ert_list.columns):
        temp_ert = df_ert_list.copy()
        temp_ert["RI등급"] = clean_text_series(temp_ert["RI등급"]).str.upper().str.replace(" ", "", regex=False)
        temp_ert["발생일"] = pd.to_datetime(temp_ert["발생일"], errors="coerce")

        ert_total_count = len(
            temp_ert[
                (temp_ert["RI등급"].isin(["A", "B", "C+"]))
                & (temp_ert["발생일"].dt.year == CURRENT_YEAR)
            ]
        )

    repeat_col = "재발사고"
    repeat_count = 0

    if repeat_col in df_acc.columns and "발생년도" in df_acc.columns:
        repeat_count = len(
            df_acc[
                clean_text_series(df_acc[repeat_col]).str.contains("재발", na=False)
                & (pd.to_numeric(df_acc["발생년도"], errors="coerce") == CURRENT_YEAR)
            ]
        )

    status_high = status_text(actual_high, target_high)
    status_grade = status_text(actual_grade, target_grade)
    status_leak = status_text(actual_leak, target_leak)
    status_ert = status_text(actual_ert, target_ert)

    leak_count_card = fmt_no_decimal(actual_leak)

    card_cols = st.columns(4)

    with card_cols[0]:
        render_kpi_card(
            signal_image(actual_high, target_high),
            "고위험사고 KPI",
            make_kpi_card_text(actual_high, target_high, status_high),
        )

    with card_cols[1]:
        render_kpi_card(
            signal_image(actual_grade, target_grade),
            "등급사고 KPI",
            make_kpi_card_text(actual_grade, target_grade, status_grade),
        )

    with card_cols[2]:
        render_kpi_card(
            signal_image(actual_leak, target_leak),
            "LEAK사고 KPI",
            (
                f"{CURRENT_MONTH}월 실적 {fmt_no_decimal(actual_leak)} / "
                f"목표 {fmt_no_decimal(target_leak)} / {status_leak}"
                "<span class='metric-unit'>(건)</span>"
                f"<br><div class='metric-count'>{leak_count_card}건</div>"
            ),
        )

    with card_cols[3]:
        render_kpi_card(
            signal_image(actual_ert, target_ert),
            "ERT고등급출동 KPI",
            (
                f"{CURRENT_MONTH}월 실적 {fmt_no_decimal(actual_ert)} / "
                f"목표 {fmt_no_decimal(target_ert)} / {status_ert}, 총 {ert_total_count}건"
                "<span class='metric-unit'>(건)</span>"
            ),
        )

    graph_cols = st.columns(4)

    with graph_cols[0]:
        draw_bar_chart(
            "고위험사고",
            y_goal_high,
            y_act_high,
            value_digits=0,
        )

    with graph_cols[1]:
        draw_bar_chart(
            "등급사건",
            y_goal_grade,
            y_act_grade,
            goal_label="등급사건 목표",
            actual_label="등급사건 실적",
            line_values=y_ltir,
            line_label="LTIR",
            value_digits=0,
        )

        if ltir_col is None:
            st.caption("※ Sheet '등급'에서 'LTIR 0.020 만인률' 컬럼을 찾지 못했습니다.")

    with graph_cols[2]:
        draw_bar_chart(
            "LEAK(경미 이상)",
            y_goal_leak,
            y_act_leak,
            goal_label="25년",
            actual_label="26년",
            value_digits=0,
        )

    with graph_cols[3]:
        draw_bar_chart(
            "ERT출동건수(고등급)",
            y_goal_ert,
            y_act_ert,
            value_digits=0,
        )

    base_cols = [
        "재발사고",
        "발생일",
        "사고단계",
        "EHS팀 저감대책",
        "변경 등급",
        "변경 대분류",
        "고위험/일반/생활사고",
        "건물명",
        "사고명",
        "재해유형",
        "사업부",
        "부서명",
        "재발인자(조직/사고유형/기인물/세부원인)",
        "컨플링크",
    ]

    if {"변경 등급", "고위험/일반/생활사고", "발생년도"}.issubset(df_acc.columns):
        high_df = df_acc[
            (df_acc["변경 등급"].isin(["등급", "경미"]))
            & (df_acc["고위험/일반/생활사고"] == "고위험")
            & (pd.to_numeric(df_acc["발생년도"], errors="coerce") == CURRENT_YEAR)
        ].copy()

        high_table = prepare_accident_table(high_df, base_cols)

    else:
        high_table = pd.DataFrame()

    if repeat_col in df_acc.columns and "발생년도" in df_acc.columns:
        repeat_df = df_acc.copy()
        repeat_df[repeat_col] = clean_text_series(repeat_df[repeat_col])

        repeat_df = repeat_df[
            repeat_df[repeat_col].str.contains("재발", na=False)
            & (pd.to_numeric(repeat_df["발생년도"], errors="coerce") == CURRENT_YEAR)
        ].copy()

        repeat_cols = [
            "재발사고",
            "사고명",
            "재발인자(조직/사고유형/기인물/세부원인)",
            "발생일",
            "사고단계",
            "부서명",
            "EHS팀 저감대책",
            "변경 등급",
            "변경 대분류",
            "고위험/일반/생활사고",
            "건물명",
            "재해유형",
            "사업부",
            "컨플링크",
        ]

        repeat_table = prepare_accident_table(repeat_df, repeat_cols)

    else:
        repeat_table = pd.DataFrame()

    if {"고위험/일반/생활사고", "변경 등급", "발생년도"}.issubset(df_acc.columns):
        grade_df = df_acc[
            (df_acc["고위험/일반/생활사고"].isin(["고위험", "일반"]))
            & (df_acc["변경 등급"] == "등급")
            & (pd.to_numeric(df_acc["발생년도"], errors="coerce") == CURRENT_YEAR)
        ].copy()

        grade_table = prepare_accident_table(grade_df, base_cols)

    else:
        grade_table = pd.DataFrame()

    if {"발생년도", "변경 등급", "재해유형"}.issubset(df_acc.columns):
        leak_df = df_acc[
            (pd.to_numeric(df_acc["발생년도"], errors="coerce") == CURRENT_YEAR)
            & (df_acc["변경 등급"].isin(["경미", "등급"]))
            & (df_acc["재해유형"] == "누출/접촉/흡입")
        ].copy()

        leak_table = prepare_accident_table(leak_df, base_cols)

    else:
        leak_table = pd.DataFrame()

    if not df_ert_list.empty and {"RI등급", "발생일"}.issubset(df_ert_list.columns):
        ert_df = df_ert_list.copy()
        ert_df["RI등급"] = clean_text_series(ert_df["RI등급"]).str.upper().str.replace(" ", "", regex=False)
        ert_df["발생일"] = pd.to_datetime(ert_df["발생일"], errors="coerce")

        ert_df = ert_df[
            (ert_df["RI등급"].isin(["A", "B", "C+"]))
            & (ert_df["발생일"].dt.year == CURRENT_YEAR)
        ].copy()

        ert_df = ert_df.sort_values("발생일", ascending=False)
        ert_df["발생일"] = ert_df["발생일"].dt.strftime("%Y-%m-%d").fillna("")

        ert_cols = [
            c
            for c in [
                "사업부",
                "담당부서",
                "발생일",
                "RI등급",
                "제목",
                "사업장",
                "건물",
                "층",
            ]
            if c in ert_df.columns
        ]

        ert_df = ert_df.astype(object).where(pd.notna(ert_df), "")

        for c in ert_df.columns:
            ert_df[c] = ert_df[c].astype(str)

    else:
        ert_df = pd.DataFrame()
        ert_cols = []

    summary = {
        "고위험사고": (actual_high, target_high, status_high),
        "등급사고": (actual_grade, target_grade, status_grade),
        "LEAK 경미이상": (actual_leak, target_leak, status_leak),
        "ERT 고등급 출동": (actual_ert, target_ert, status_ert),
        "ERT 총 건수": ert_total_count,
        "재발사고 건수": repeat_count,
    }

    render_ai_panel(
        high_table=high_table,
        repeat_table=repeat_table,
        grade_table=grade_table,
        leak_table=leak_table,
        ert_table=ert_df,
        summary=summary,
        df_acc=df_acc,
        df_ert_list=df_ert_list,
    )

    render_section_header("🚨 고위험사고 LIST", len(high_table), list_target_high)
    if not high_table.empty:
        show_table(high_table, "high_risk_list")
    else:
        st.warning("고위험사고 LIST 생성에 필요한 데이터가 없습니다.")

    render_section_header("🔁 고위험 재발 사고 LIST", len(repeat_table), None)
    if not repeat_table.empty:
        show_table(repeat_table, "repeat_accident_list")
    else:
        st.warning("재발 사고 LIST 생성에 필요한 데이터가 없습니다.")

    render_section_header("⚠️ 등급사고 LIST", len(grade_table), list_target_grade)
    if not grade_table.empty:
        show_table(grade_table, "grade_accident_list")
    else:
        st.warning("등급사고 LIST 생성에 필요한 데이터가 없습니다.")

    render_section_header("💧 LEAK 경미이상 LIST", len(leak_table), list_target_leak)
    if not leak_table.empty:
        show_table(leak_table, "leak_minor_list")
    else:
        st.warning("LEAK LIST 생성에 필요한 데이터가 없습니다.")

    render_section_header("🚒 ERT고등급 출동 LIST", len(ert_df), None)
    if not ert_df.empty:
        ert_view = ert_df[ert_cols].copy() if ert_cols else ert_df.copy()
        render_safe_html_table(ert_view, max_rows=300)
    else:
        st.warning("ERT 고등급 출동 LIST 생성에 필요한 데이터가 없습니다.")

    st.markdown("---")
    st.caption(
        "※ DS OSS Credential은 코드에 직접 입력하지 말고 환경변수 또는 .streamlit/secrets.toml로 관리하세요."
    )


# =====================================================
# Entry
# =====================================================

def main():
    render_sidebar()
    render_kpi_page()


if __name__ == "__main__":
    main()
