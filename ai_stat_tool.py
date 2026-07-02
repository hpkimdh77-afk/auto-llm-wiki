"""
ai_stat_tool.py

Streamlit 기반 AI 사고 통합검색 & 통계 분석 도구입니다.

DS-OSS GPT 호출 구조는 기존 방식 그대로 유지합니다.

v2 개편 내용:
- "사고 통합검색"을 메인 기능으로 승격 (구글처럼 자연어로 물어보면 바로 답이 나오는 구조)
- 검색어에서 연도/기간 표현 자동 인식 (예: "2023년 이후", "최근 3년", "21~23년")
- 동의어 확장 + 관련도 스코어링 검색 (화재↔발화, 협착↔끼임 등 / AND 우선, 부분일치 폴백)
- 데이터 기준 기간 2021~2026년 고정 스코프
- AI 답변은 매칭 데이터 요약 + 표본 + 전체 연도 컨텍스트에만 근거하도록 프롬프트 강화
- 검색 UI를 검색 히어로 + 추천 검색어 칩 + 답변 카드 형태로 개선
"""

import os
import re
import uuid
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from openai import OpenAI


# ============================================================================
# 기본 설정
# ============================================================================

# 데이터 파일 경로: 기존 환경과 동일
ACCIDENT_FILE = "/config/work/sharedworkspace/ITS등록사고_streamlit.xlsx"

# DS-OSS / GPT OSS 설정: 기존 기본값 유지
DEFAULT_API_BASE_URL = "http://apigw.samsungds.net:8000/gpt-oss-playground/1/v1"
DEFAULT_MODEL = "gpt-oss-120b"

# 보안상 토큰과 User ID는 코드에 하드코딩하지 않음
# 아래 우선순위로 읽음:
# 1) Streamlit secrets
# 2) 환경변수
# 3) 사이드바 직접 입력
DEFAULT_CREDENTIAL_KEY = ""
DEFAULT_USER_ID = ""

DEFAULT_TIMEOUT = 300
DEFAULT_RETRY = 2

# 분석 대상 데이터 기준 기간
DATA_YEAR_MIN = 2021
DATA_YEAR_MAX = 2026


# ============================================================================
# 통합검색용 언어 자원 (동의어 / 불용어)
# ============================================================================

# 같은 그룹 안의 단어는 서로 동의어로 취급하여 검색 범위를 넓힘 (재현율 우선)
SYNONYM_GROUPS: List[List[str]] = [
    ["화재", "발화", "연소", "소손", "발연", "연기", "불티", "불꽃", "스파크"],
    ["폭발", "파열", "폭주"],
    ["낙상", "넘어짐", "전도", "미끄러짐", "미끄러"],
    ["추락", "떨어짐", "낙하"],
    ["협착", "끼임", "말림", "눌림"],
    ["감전", "누전", "합선"],
    ["누출", "유출", "누수", "리크", "누액", "leak"],
    ["화학", "약품", "약액", "케미컬", "chemical"],
    ["질식", "산소결핍", "중독"],
    ["충돌", "부딪힘", "부딪침", "접촉"],
    ["베임", "절단", "찔림", "자상", "절상"],
    ["화상", "데임", "열상", "고온"],
    ["배관", "파이프", "piping"],
    ["정비", "보수", "수리", "maintenance"],
    ["설비", "장비", "장치"],
]

# 단어 -> 동의어 그룹 조회 테이블
SYNONYM_LOOKUP: Dict[str, List[str]] = {}
for _group in SYNONYM_GROUPS:
    for _word in _group:
        SYNONYM_LOOKUP[_word] = _group

# 별칭: 그룹으로 확장은 하되, 그 단어 자체는 검색 패턴으로 쓰지 않음
# (예: "불"은 "불량" 같은 무관한 단어까지 매칭되므로 화재 그룹으로만 치환)
SYNONYM_ALIASES: Dict[str, str] = {
    "불": "화재",
    "소화": "화재",
}

# 검색 키워드에서 제외할 단어 (질문 조사는 LLM에 원문으로 전달되므로 검색에서는 제거)
STOPWORDS = {
    "사고", "재해", "알려줘", "알려주세요", "알려", "정리", "정리해줘", "해줘", "주세요",
    "관련", "관련된", "현황", "건수", "몇", "건", "건이야", "건이고", "건인지",
    "뭐야", "뭐가", "무엇", "무엇이", "어떤", "어떻게", "어디서", "언제", "누가", "왜",
    "년", "년도", "연도", "최근", "이후", "이전", "부터", "까지", "동안", "사이",
    "및", "그리고", "또", "또는", "대해", "대한", "대해서", "좀", "그", "이", "저",
    "발생", "발생한", "발생했어", "있어", "있었어", "있는", "있나", "없어",
    "궁금", "궁금해", "보여줘", "찾아줘", "검색", "분석", "분석해줘", "통계",
    "케이스", "사례", "원인", "원인은", "원인이", "대책", "대책도", "패턴",
    "많이", "가장", "제일", "주요", "중", "중에", "중에서", "요약",
    "연도별", "월별", "일별", "추이", "트렌드", "경향", "변화", "변하고", "변했어",
    "증가", "감소", "비율", "순위", "분포", "비교",
}

# 토큰 끝에 붙는 조사 문자 (반복 제거)
PARTICLE_CHARS = "은는이가을를의도에로와과서만"

# 추천 검색어 (구글 스타일 예시 칩)
EXAMPLE_QUERIES = [
    "2023년 이후 화재 사고 몇 건이고 주요 원인은 뭐야?",
    "최근 3년 협착(끼임) 사고 패턴 알려줘",
    "감전 사고는 어떤 작업 단계에서 많이 발생했어?",
    "2021~2026년 고위험 사고 중 가장 반복되는 유형은?",
    "화학물질 누출 사고 대표 사례와 대책 정리해줘",
    "연도별 사고 추이가 어떻게 변하고 있어?",
]


# ============================================================================
# 공통 유틸
# ============================================================================

def _is_text_column(series: pd.Series) -> bool:
    """
    텍스트 컬럼 여부를 판단합니다.
    pandas 2.x(object dtype)와 3.x(str dtype) 모두에서 동작하도록 처리합니다.
    """
    try:
        return pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series)
    except Exception:
        return series.dtype == "object"


def get_secret_or_env(secret_key: str, env_key: str, default: str = "") -> str:
    """
    Streamlit secrets 또는 환경변수에서 값을 가져옵니다.
    secrets.toml이 없거나 접근 불가한 경우에도 앱이 죽지 않도록 처리합니다.
    """
    value = ""

    try:
        value = st.secrets.get(secret_key, "")
    except Exception:
        value = ""

    if not value:
        value = os.getenv(env_key, default)

    return value


def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    사고 데이터프레임 컬럼명과 주요 타입을 정리합니다.
    기본 파일과 업로드 파일 모두 동일 전처리를 적용합니다.
    발생년도가 비어 있으면 발생일에서 보완하여 연도 기반 분석에서 빠지는 행을 줄입니다.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()

    df.columns = (
        df.columns.astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.replace("\n", "", regex=False)
        .str.replace("\r", "", regex=False)
        .str.strip()
    )

    if "발생일" in df.columns:
        df["발생일"] = pd.to_datetime(df["발생일"], errors="coerce")

    if "발생년도" in df.columns:
        df["발생년도"] = pd.to_numeric(df["발생년도"], errors="coerce")
        if "발생일" in df.columns:
            df["발생년도"] = df["발생년도"].fillna(df["발생일"].dt.year)
    elif "발생일" in df.columns:
        df["발생년도"] = df["발생일"].dt.year

    for col in df.columns:
        if _is_text_column(df[col]):
            df[col] = df[col].fillna("").astype(str).str.strip()

    return df


def apply_year_scope(df: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
    """
    데이터를 기준 기간(2021~2026년)으로 제한합니다.
    연도를 알 수 없는 행은 데이터 유실 방지를 위해 유지하고, 그 개수를 함께 반환합니다.
    """
    if df is None or df.empty or "발생년도" not in df.columns:
        return df, 0

    years = pd.to_numeric(df["발생년도"], errors="coerce")
    in_scope = years.between(DATA_YEAR_MIN, DATA_YEAR_MAX)
    unknown = years.isna()

    scoped = df[in_scope | unknown]
    return scoped, int(unknown.sum())


def filter_by_year(df: pd.DataFrame, year_from: int, year_to: int) -> pd.DataFrame:
    """
    발생년도 기준으로 데이터를 특정 기간으로 필터링합니다.
    """
    if df is None or df.empty or "발생년도" not in df.columns:
        return df

    years = pd.to_numeric(df["발생년도"], errors="coerce")
    return df[years.between(year_from, year_to)]


def validate_ascii_headers(headers: Dict[str, str]) -> None:
    """
    HTTP header 값에 비 ASCII 문자가 들어가면 OpenAI client에서 오류가 날 수 있어 사전 검증합니다.
    """
    for key, value in headers.items():
        if isinstance(value, str):
            try:
                value.encode("ascii")
            except UnicodeEncodeError:
                raise ValueError(f"HTTP header '{key}' contains non-ASCII characters.")


def to_display_date(df: pd.DataFrame, col: str = "발생일") -> pd.DataFrame:
    """
    날짜 컬럼을 화면 표시용 date로 변환합니다.
    """
    display_df = df.copy()

    if col in display_df.columns:
        display_df[col] = pd.to_datetime(display_df[col], errors="coerce").dt.date

    return display_df


def build_value_count_table(df: pd.DataFrame, col: str, top_n: int = 10) -> pd.DataFrame:
    """
    특정 컬럼의 상위 빈도표를 생성합니다.
    """
    if df is None or df.empty or col not in df.columns:
        return pd.DataFrame(columns=[col, "건수"])

    s = df[col].dropna().astype(str).str.strip()
    s = s[s != ""]

    if s.empty:
        return pd.DataFrame(columns=[col, "건수"])

    return (
        s.value_counts()
        .head(top_n)
        .rename_axis(col)
        .reset_index(name="건수")
    )


# ============================================================================
# 데이터 로드
# ============================================================================

@st.cache_data(ttl=900)
def load_default_accident_data() -> pd.DataFrame:
    """
    기본 사고 데이터를 로드합니다.
    """
    if not Path(ACCIDENT_FILE).exists():
        return pd.DataFrame()

    df = pd.read_excel(ACCIDENT_FILE, sheet_name="Sheet2")
    return normalize_dataframe(df)


@st.cache_data(ttl=900)
def load_default_ert_data() -> pd.DataFrame:
    """
    기본 파일에서 ERT 출동 데이터를 로드합니다.
    """
    if not Path(ACCIDENT_FILE).exists():
        return pd.DataFrame()

    try:
        df = pd.read_excel(ACCIDENT_FILE, sheet_name="ERT출동")
        return normalize_dataframe(df)
    except Exception:
        return pd.DataFrame()


def load_uploaded_data(upload_file) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    업로드된 파일에서 사고 데이터와 ERT 데이터를 읽습니다.
    xlsx의 경우 Sheet2와 ERT출동 시트를 우선 사용합니다.
    csv의 경우 사고 데이터만 읽고 ERT는 빈 데이터프레임으로 반환합니다.
    """
    if upload_file is None:
        return pd.DataFrame(), pd.DataFrame()

    try:
        file_name = upload_file.name.lower()

        if file_name.endswith(".csv"):
            df = pd.read_csv(upload_file)
            return normalize_dataframe(df), pd.DataFrame()

        excel_file = pd.ExcelFile(upload_file)
        sheet_names = excel_file.sheet_names

        accident_sheet = "Sheet2" if "Sheet2" in sheet_names else sheet_names[0]
        df = pd.read_excel(excel_file, sheet_name=accident_sheet)
        df = normalize_dataframe(df)

        if "ERT출동" in sheet_names:
            ert_df = pd.read_excel(excel_file, sheet_name="ERT출동")
            ert_df = normalize_dataframe(ert_df)
        else:
            ert_df = pd.DataFrame()

        return df, ert_df

    except Exception as e:
        st.error(f"업로드된 파일을 읽을 수 없습니다: {e}")
        return pd.DataFrame(), pd.DataFrame()


# ============================================================================
# DS-OSS GPT Client
# ============================================================================

def make_client(api_base_url: str, token: str, user_id: str) -> OpenAI:
    """
    DS-OSS GPT API 클라이언트를 생성합니다.
    기존 DS-OSS 호출 방식은 유지합니다.
    """
    headers = {
        "x-dep-ticket": token,
        "Send-System-Name": "ai-stat-tool",
        "User-Id": user_id,
        "User-Type": "AD_ID",
        "Prompt-Msg-Id": str(uuid.uuid4()),
        "Completion-Msg-Id": str(uuid.uuid4()),
    }

    validate_ascii_headers(headers)

    return OpenAI(
        api_key="dummy",
        base_url=api_base_url,
        default_headers=headers,
        timeout=DEFAULT_TIMEOUT,
        max_retries=DEFAULT_RETRY,
    )


def call_llm(
    prompt: str,
    api_base_url: str,
    token: str,
    user_id: str,
    model_name: str,
    temperature: float = 0.2,
    max_tokens: int = 2048,
) -> str:
    """
    DS-OSS GPT 모델을 호출하여 결과를 반환합니다.
    """
    if not token:
        return "API 호출 오류:\n\nGPT OSS 토큰이 입력되지 않았습니다."

    if not user_id:
        return "API 호출 오류:\n\nUser ID가 입력되지 않았습니다."

    if not api_base_url:
        return "API 호출 오류:\n\nAPI Base URL이 입력되지 않았습니다."

    system_prompt = (
        "너는 사내 환경안전 사고 데이터 분석을 지원하는 AI 도우미다. "
        f"분석 대상 데이터의 기준 기간은 {DATA_YEAR_MIN}년부터 {DATA_YEAR_MAX}년까지다. "
        "반드시 사용자가 제공한 데이터와 요약 정보에 근거해서 답변해야 한다. "
        "데이터에 없는 사실, 건수, 원인, 출처, 날짜, 링크는 지어내지 않는다. "
        "확인되지 않은 내용은 '제공된 데이터만으로는 확인 불가' 또는 '추가 확인 필요'라고 표시한다. "
        "건수와 비율을 말할 때는 제공된 집계표의 숫자를 그대로 사용한다. "
        "개선방안은 현장 실행 중심으로 구체적이고 실무적으로 제시한다."
    )

    last_error = None

    for attempt in range(DEFAULT_RETRY + 1):
        try:
            client = make_client(api_base_url, token, user_id)

            response = client.chat.completions.create(
                model=model_name,
                messages=[
                    {
                        "role": "system",
                        "content": system_prompt,
                    },
                    {
                        "role": "user",
                        "content": prompt,
                    },
                ],
                temperature=temperature,
                max_tokens=max_tokens,
            )

            return response.choices[0].message.content

        except Exception as e:
            last_error = e
            time.sleep(1.5)

    return f"API 호출 오류:\n\n{last_error}"


# ============================================================================
# 통합검색: 질의 해석
# ============================================================================

def _norm_year(raw: Optional[str]) -> Optional[int]:
    """
    '2023', '23' 같은 연도 문자열을 4자리 연도로 정규화합니다.
    """
    if not raw:
        return None

    try:
        y = int(raw)
    except ValueError:
        return None

    if y < 100:
        y += 2000

    if 2000 <= y <= 2099:
        return y

    return None


def _strip_particles(token: str) -> str:
    """
    토큰 끝의 조사(이/가/은/는/을/를 등)를 제거합니다. 최소 2글자는 남깁니다.
    """
    while len(token) > 2 and token[-1] in PARTICLE_CHARS:
        token = token[:-1]
    return token


def _expand_token(token: str) -> List[str]:
    """
    하나의 검색 토큰을 변형(조사 제거, '사고' 접미어 제거)과 동의어로 확장합니다.
    반환된 리스트 안의 단어들은 OR 조건으로 매칭됩니다.
    """
    variants = {token}

    stripped = _strip_particles(token)
    variants.add(stripped)

    for base in list(variants):
        for suffix in ("사고", "관련", "사례"):
            if base.endswith(suffix) and len(base) > len(suffix):
                variants.add(base[: -len(suffix)])

    expanded = set(variants)
    for v in list(variants):
        if v in SYNONYM_LOOKUP:
            expanded.update(SYNONYM_LOOKUP[v])
        if v in SYNONYM_ALIASES:
            expanded.discard(v)
            target = SYNONYM_ALIASES[v]
            expanded.update(SYNONYM_LOOKUP.get(target, [target]))

    return sorted(expanded, key=len, reverse=True)


def parse_search_query(query: str) -> Dict:
    """
    자연어 검색 질의에서 기간 조건과 검색 키워드 그룹을 추출합니다.

    반환 값:
    - raw: 원본 질의
    - year_from / year_to: 인식된 기간 (없으면 None)
    - tokens: 핵심 키워드 (표시용)
    - token_groups: 동의어까지 확장된 키워드 그룹 (검색용)
    """
    raw = query.strip()
    work = raw
    year_from: Optional[int] = None
    year_to: Optional[int] = None

    # 1) "최근 N년"
    m = re.search(r"최근\s*(\d{1,2})\s*년", work)
    if m:
        n = max(1, int(m.group(1)))
        year_to = DATA_YEAR_MAX
        year_from = max(DATA_YEAR_MIN, DATA_YEAR_MAX - n + 1)
        work = work.replace(m.group(0), " ")

    # 2) 기간 범위: "2021~2023", "21~23년", "2021-2023년"
    m = re.search(r"(20\d{2}|\d{2})\s*년?\s*[~∼\-]\s*(20\d{2}|\d{2})\s*년?", work)
    if m:
        y1 = _norm_year(m.group(1))
        y2 = _norm_year(m.group(2))
        if y1 and y2:
            year_from, year_to = min(y1, y2), max(y1, y2)
            work = work.replace(m.group(0), " ")

    # 3) 단일 연도: "2023년", "2023", "23년"
    singles: List[int] = []
    for sm in re.finditer(r"(20\d{2})\s*년?|(?<!\d)(2[0-9])\s*년", work):
        y = _norm_year(sm.group(1) or sm.group(2))
        if y:
            singles.append(y)

    if singles:
        work = re.sub(r"(20\d{2})\s*년?|(?<!\d)(2[0-9])\s*년", " ", work)
        if year_from is None:
            year_from, year_to = min(singles), max(singles)

    # 4) "이후/부터", "이전/까지" 보정 (단일 연도가 지정된 경우)
    if year_from is not None and year_from == year_to:
        if re.search(r"이후|부터", raw):
            year_to = DATA_YEAR_MAX
        elif re.search(r"이전|까지", raw):
            year_from = DATA_YEAR_MIN

    # 5) 키워드 토큰 추출
    tokens: List[str] = []
    for t in re.findall(r"[가-힣a-zA-Z0-9]+", work):
        t = t.strip()
        if not t or t.isdigit():
            continue

        base = _strip_particles(t)

        # '감전사고' -> '감전' 처럼 접미어를 뗀 형태를 대표 키워드로 사용
        for suffix in ("사고", "관련", "사례"):
            if base.endswith(suffix) and len(base) > len(suffix):
                base = base[: -len(suffix)]
                break

        if t in STOPWORDS or base in STOPWORDS:
            continue
        if len(base) < 2 and base not in SYNONYM_LOOKUP and base not in SYNONYM_ALIASES:
            continue

        if base not in tokens:
            tokens.append(base)

    token_groups = [_expand_token(t) for t in tokens]

    return {
        "raw": raw,
        "year_from": year_from,
        "year_to": year_to,
        "tokens": tokens,
        "token_groups": token_groups,
    }


# ============================================================================
# 통합검색: 데이터 검색
# ============================================================================

def smart_search(
    df: pd.DataFrame,
    token_groups: List[List[str]],
    tokens: Optional[List[str]] = None,
) -> Tuple[pd.DataFrame, str]:
    """
    동의어 확장된 키워드 그룹으로 사고 데이터를 검색합니다.

    - 모든 텍스트 컬럼을 대상으로 검색해 누락(빈틈)을 최소화합니다.
    - 데이터에 전혀 없는 키워드는 자동 제외하여 결과가 0건으로 끊기지 않게 합니다.
    - 남은 키워드가 모두 일치하는 행(AND)을 우선 반환하고,
      결과가 없으면 일부만 일치하는 행을 관련도 순으로 반환합니다.

    반환: (검색 결과, 매칭 모드 설명)
    """
    if df is None or df.empty:
        return pd.DataFrame(), "데이터 없음"

    if not token_groups:
        return df.copy(), "키워드 없음 (기간/전체 기준)"

    text_cols = [c for c in df.columns if _is_text_column(df[c])]
    if not text_cols:
        return pd.DataFrame(), "검색 가능한 텍스트 컬럼 없음"

    combined = (
        df[text_cols]
        .astype(str)
        .agg(" | ".join, axis=1)
        .str.lower()
    )

    group_masks: List[pd.Series] = []

    for group in token_groups:
        group_mask = pd.Series(False, index=df.index)

        for variant in group:
            group_mask = group_mask | combined.str.contains(
                variant.lower(),
                regex=False,
                na=False,
            )

        group_masks.append(group_mask)

    # 데이터에 전혀 등장하지 않는 키워드는 제외 (전체 결과가 0건이 되는 것을 방지)
    labels = tokens if tokens and len(tokens) == len(token_groups) else [
        g[0] for g in token_groups
    ]
    effective_masks = [m for m in group_masks if m.any()]
    ignored_labels = [
        labels[i] for i, m in enumerate(group_masks) if not m.any()
    ]

    if not effective_masks:
        return df.copy(), "키워드 매칭 없음 (전체 데이터 기준)"

    matched_groups = pd.Series(0, index=df.index)
    for m in effective_masks:
        matched_groups = matched_groups + m.astype(int)

    n_groups = len(effective_masks)

    strict = df[matched_groups == n_groups]

    if not strict.empty:
        result = strict.copy()
        result["_score"] = n_groups
        mode = "전체 키워드 일치"
    else:
        result = df[matched_groups > 0].copy()
        result["_score"] = matched_groups[result.index]
        mode = "부분 일치 (일부 키워드만 포함된 결과)"

    if ignored_labels:
        mode += f" · 데이터에 없는 키워드 제외: {', '.join(ignored_labels)}"

    sort_cols = ["_score"]
    ascending = [False]

    if "발생일" in result.columns:
        sort_cols.append("발생일")
        ascending.append(False)

    result = result.sort_values(sort_cols, ascending=ascending).drop(columns=["_score"])

    return result, mode


# ============================================================================
# 사고 데이터 요약
# ============================================================================

def build_data_summary(df: pd.DataFrame, top_n: int = 10) -> str:
    """
    사고 데이터프레임에서 AI 분석용 요약 텍스트를 생성합니다.
    """
    if df is None or df.empty:
        return "데이터 없음"

    lines: List[str] = []
    lines.append(f"[총 건수]\n{len(df)}건")

    if "발생년도" in df.columns:
        year_series = pd.to_numeric(df["발생년도"], errors="coerce").dropna()

        if not year_series.empty:
            year_counts = year_series.astype(int).value_counts().sort_index()
            lines.append("\n[연도별 사고 건수]")
            lines.append(year_counts.to_string())

    if "발생일" in df.columns:
        date_series = pd.to_datetime(df["발생일"], errors="coerce").dropna()

        if not date_series.empty:
            monthly_counts = date_series.dt.to_period("M").astype(str).value_counts().sort_index()
            lines.append("\n[월별 사고 건수]")
            lines.append(monthly_counts.to_string())

    summary_cols = [
        ("고위험/일반/생활사고", "사고구분별 건수"),
        ("재해유형", "재해유형 상위"),
        ("재해유형(상세)", "재해유형 상세 상위"),
        ("사고명", "사고명 상위"),
        ("사고단계", "사고단계 상위"),
        ("사고원인상세", "사고원인상세 상위"),
        ("사고대책", "사고대책 상위"),
    ]

    for col, title in summary_cols:
        if col not in df.columns:
            continue

        s = df[col].dropna().astype(str).str.strip()
        s = s[s != ""]

        if s.empty:
            continue

        counts = s.value_counts().head(top_n)
        lines.append(f"\n[{title} {top_n}]")
        lines.append(counts.to_string())

    if "재해유형" in df.columns and "고위험/일반/생활사고" in df.columns:
        try:
            cross = pd.crosstab(
                df["재해유형"].astype(str).str.strip(),
                df["고위험/일반/생활사고"].astype(str).str.strip(),
            )

            if not cross.empty:
                lines.append("\n[재해유형 x 사고구분 교차표]")
                lines.append(cross.to_string())

        except Exception:
            pass

    return "\n".join(lines)


def build_ert_summary(df: pd.DataFrame) -> str:
    """
    ERT 출동 데이터 요약을 문자열로 생성합니다.
    """
    if df is None or df.empty:
        return "ERT 데이터 없음"

    lines: List[str] = []
    lines.append(f"[총 고등급 출동 건수]\n{len(df)}건")

    if "RI등급" in df.columns:
        counts = df["RI등급"].dropna().astype(str).str.strip().value_counts().sort_index()
        if not counts.empty:
            lines.append("\n[RI등급별 출동건수]")
            lines.append(counts.to_string())

    if "발생일" in df.columns:
        date_series = pd.to_datetime(df["발생일"], errors="coerce").dropna()

        if not date_series.empty:
            year_counts = date_series.dt.year.value_counts().sort_index()
            lines.append("\n[연도별 출동건수]")
            lines.append(year_counts.to_string())

            monthly_counts = date_series.dt.to_period("M").astype(str).value_counts().sort_index()
            lines.append("\n[월별 출동건수]")
            lines.append(monthly_counts.to_string())

    for col in ["출동유형", "사고명", "재해유형", "발생장소", "조치내용"]:
        if col not in df.columns:
            continue

        s = df[col].dropna().astype(str).str.strip()
        s = s[s != ""]

        if s.empty:
            continue

        counts = s.value_counts().head(10)
        lines.append(f"\n[{col} 상위 10]")
        lines.append(counts.to_string())

    return "\n".join(lines)


def build_compact_table(df: pd.DataFrame, max_rows: int = 20, max_cell_len: int = 120) -> str:
    """
    AI 입력용 표본 데이터를 CSV 형태로 반환합니다.
    전체 컬럼을 다 보내지 않고 주요 컬럼만 사용하며, 긴 셀은 잘라 토큰을 절약합니다.
    """
    if df is None or df.empty:
        return "데이터 없음"

    preferred_cols = [
        "발생일",
        "발생년도",
        "고위험/일반/생활사고",
        "재해유형",
        "재해유형(상세)",
        "사고명",
        "사고단계",
        "사고원인상세",
        "사고대책",
    ]

    cols = [col for col in preferred_cols if col in df.columns]

    if not cols:
        cols = list(df.columns[:10])

    try:
        tmp = df[cols].head(max_rows).copy()

        for col in tmp.columns:
            if pd.api.types.is_datetime64_any_dtype(tmp[col]):
                tmp[col] = tmp[col].dt.strftime("%Y-%m-%d")
            else:
                tmp[col] = tmp[col].astype(str).str.slice(0, max_cell_len)

        return tmp.to_csv(index=False)

    except Exception:
        return "데이터를 요약하는 과정에서 오류가 발생했습니다."


def filter_accident_by_risk(
    df: pd.DataFrame,
    risk_type: str,
    selected_type: Optional[str],
) -> pd.DataFrame:
    """
    사고구분과 재해유형 기준으로 데이터를 필터링합니다.

    risk_type:
    - highrisk: 고위험만
    - general: 고위험 제외
    - all: 전체
    """
    if df is None or df.empty:
        return pd.DataFrame()

    result = df.copy()

    if "고위험/일반/생활사고" in result.columns:
        risk_col = result["고위험/일반/생활사고"].astype(str).str.strip()

        if risk_type == "highrisk":
            result = result[risk_col == "고위험"]
        elif risk_type == "general":
            result = result[risk_col != "고위험"]

    if selected_type and selected_type != "(전체)" and "재해유형" in result.columns:
        result = result[result["재해유형"].astype(str).str.strip() == selected_type]

    return result


def prepare_ert_high_grade_data(ert_data: pd.DataFrame) -> pd.DataFrame:
    """
    ERT 데이터에서 고등급 RI 등급 A, B, C+ 출동만 필터링합니다.
    """
    if ert_data is None or ert_data.empty:
        return pd.DataFrame()

    df_ert = ert_data.copy()

    if "발생일" in df_ert.columns:
        df_ert["발생일"] = pd.to_datetime(df_ert["발생일"], errors="coerce")

    if "RI등급" in df_ert.columns:
        df_ert["RI등급"] = (
            df_ert["RI등급"]
            .fillna("")
            .astype(str)
            .str.upper()
            .str.replace(" ", "", regex=False)
            .str.strip()
        )

        df_ert = df_ert[df_ert["RI등급"].isin(["A", "B", "C+"])]

    return df_ert


# ============================================================================
# 통합검색: AI 프롬프트 & 실행
# ============================================================================

def build_unified_prompt(
    question: str,
    parsed: Dict,
    matched_df: pd.DataFrame,
    scoped_df: pd.DataFrame,
    match_mode: str,
) -> str:
    """
    통합검색용 AI 프롬프트를 생성합니다.
    질문 + 검색 조건 해석 + 전체 컨텍스트 + 매칭 데이터 요약 + 표본을 모두 포함해
    빈틈없는(그러나 데이터에 근거한) 답변을 유도합니다.
    """
    if parsed["year_from"] is not None:
        period_text = f"{parsed['year_from']}년 ~ {parsed['year_to']}년"
    else:
        period_text = f"{DATA_YEAR_MIN}년 ~ {DATA_YEAR_MAX}년 (전체)"

    if parsed["tokens"]:
        keyword_text = ", ".join(
            f"{t} (확장: {'/'.join(g[:4])})"
            for t, g in zip(parsed["tokens"], parsed["token_groups"])
        )
    else:
        keyword_text = "(키워드 없음 - 기간/전체 데이터 기준 질문)"

    # 전체 데이터 연도별 컨텍스트 (매칭 결과와 비교 근거)
    overall_lines = [f"기준 기간 전체 사고 건수: {len(scoped_df)}건"]
    if "발생년도" in scoped_df.columns:
        year_series = pd.to_numeric(scoped_df["발생년도"], errors="coerce").dropna()
        if not year_series.empty:
            overall_lines.append("연도별 전체 사고 건수:")
            overall_lines.append(
                year_series.astype(int).value_counts().sort_index().to_string()
            )
    overall_context = "\n".join(overall_lines)

    if matched_df is not None and not matched_df.empty:
        matched_summary = build_data_summary(matched_df)
        sample_table = build_compact_table(matched_df, max_rows=30)
        matched_info = f"{len(matched_df)}건 ({match_mode})"
    else:
        matched_summary = "키워드 매칭 결과 없음"
        sample_table = build_compact_table(scoped_df, max_rows=30)
        matched_info = f"0건 ({match_mode})"

    return (
        f"사내 사고 데이터({DATA_YEAR_MIN}~{DATA_YEAR_MAX}년)에 대한 통합검색 질문입니다.\n"
        "검색 시스템이 아래와 같이 질문을 해석하고 관련 데이터를 추출했습니다.\n\n"
        "[사용자 질문]\n"
        f"{question}\n\n"
        "[검색 조건 해석]\n"
        f"- 적용 기간: {period_text}\n"
        f"- 검색 키워드: {keyword_text}\n"
        f"- 매칭 결과: {matched_info}\n\n"
        "[전체 데이터 컨텍스트]\n"
        f"{overall_context}\n\n"
        "[매칭 데이터 요약]\n"
        f"{matched_summary}\n\n"
        "[매칭 표본 데이터 (관련도 순 최대 30건)]\n"
        f"{sample_table}\n\n"
        "위 데이터에만 근거하여, 검색엔진 최상단 요약 카드처럼 질문에 대한 답을 "
        "가장 먼저 제시하는 답변을 작성하세요.\n\n"
        "답변 구성 (질문과 무관한 섹션은 생략 가능):\n"
        "## 💡 핵심 답변\n"
        "- 질문에 대한 직접적인 답 2~4문장. 반드시 구체적 수치(건수, 연도, 유형)를 포함.\n"
        "## 📊 근거 데이터\n"
        "- 연도별/유형별/원인별 수치를 마크다운 표 또는 리스트로 정리.\n"
        "## 📌 대표 사례\n"
        "- 표본 데이터의 실제 행에서만 3~5건 (발생일 · 사고명 · 원인 요약). 표본에 없는 사례 금지.\n"
        "## 🔎 패턴과 시사점\n"
        "- 반복 패턴, 취약 공정/단계, 현장에서 실행 가능한 개선 포인트.\n"
        "## ⚠️ 확인 한계\n"
        "- 이 데이터만으로 답할 수 없는 부분을 명시. 없으면 '없음'.\n\n"
        "규칙:\n"
        f"1. {DATA_YEAR_MIN}~{DATA_YEAR_MAX}년 제공 데이터만 근거로 답한다.\n"
        "2. 요약/표본에 없는 수치·사례·원인·출처는 절대 만들지 않는다.\n"
        "3. 매칭 결과가 0건이면 그 사실을 핵심 답변에서 먼저 밝히고, "
        "전체 데이터 기준으로 답할 수 있는 부분만 답한다.\n"
        "4. 부분 일치 결과라면 어떤 키워드 기준의 결과인지 답변에서 밝힌다.\n"
        "5. 모호한 표현 대신 데이터의 숫자를 직접 인용한다."
    )


def execute_unified_search(
    query: str,
    df: pd.DataFrame,
    token: str,
    api_base_url: str,
    user_id: str,
    model_name: str,
) -> Dict:
    """
    통합검색을 실행합니다: 질의 해석 -> 기간 필터 -> 스마트 검색 -> AI 답변 생성.
    """
    parsed = parse_search_query(query)

    scoped = df
    if parsed["year_from"] is not None:
        scoped = filter_by_year(df, parsed["year_from"], parsed["year_to"])

    matched, match_mode = smart_search(scoped, parsed["token_groups"], parsed["tokens"])

    answer = None
    if token and user_id:
        prompt = build_unified_prompt(query, parsed, matched, scoped, match_mode)
        answer = call_llm(
            prompt=prompt,
            api_base_url=api_base_url,
            token=token,
            user_id=user_id,
            model_name=model_name,
            temperature=0.2,
            max_tokens=2500,
        )

    return {
        "query": query,
        "parsed": parsed,
        "scoped_count": len(scoped),
        "matched": matched,
        "match_mode": match_mode,
        "answer": answer,
    }


# ============================================================================
# 화면 렌더링 함수
# ============================================================================

def render_style() -> None:
    """
    Streamlit 화면 스타일을 적용합니다.
    """
    st.markdown(
        """
        <style>
        [data-testid="stSidebar"] {
            background: linear-gradient(135deg, #f9f9f9 0%, #e5ecf6 100%);
            color: #2c2c2c;
        }

        .hero-card {
            background: linear-gradient(135deg, #4f46e5 0%, #3b82f6 100%);
            color: white;
            border-radius: 16px;
            padding: 28px 32px;
            margin-bottom: 20px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }

        .hero-title {
            font-size: 34px;
            font-weight: 800;
            margin: 0;
            padding: 0;
        }

        .hero-subtitle {
            font-size: 15px;
            font-weight: 400;
            margin-top: 8px;
            opacity: 0.95;
        }

        .hero-badge {
            display: inline-block;
            background: rgba(255,255,255,0.2);
            border-radius: 999px;
            padding: 4px 12px;
            font-size: 12px;
            font-weight: 600;
            margin-right: 8px;
            margin-top: 4px;
        }

        /* 통합검색 폼: 구글 스타일 둥근 검색바 */
        [data-testid="stForm"] {
            border: none;
            padding: 0;
        }

        [data-testid="stForm"] [data-testid="stTextInput"] input {
            border-radius: 999px;
            border: 1.5px solid #dfe1e5;
            padding: 14px 22px;
            font-size: 16px;
            box-shadow: 0 1px 6px rgba(32,33,36,0.12);
        }

        [data-testid="stForm"] [data-testid="stTextInput"] input:focus {
            border-color: #4f46e5;
            box-shadow: 0 2px 10px rgba(79,70,229,0.22);
        }

        [data-testid="stForm"] button[kind="primaryFormSubmit"],
        [data-testid="stForm"] button {
            border-radius: 999px;
            font-weight: 700;
        }

        .search-condition {
            display: inline-block;
            background: #eef2ff;
            color: #3730a3;
            border-radius: 999px;
            padding: 4px 12px;
            font-size: 13px;
            font-weight: 600;
            margin-right: 6px;
            margin-bottom: 6px;
        }

        .small-caption {
            color: #666;
            font-size: 13px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_hero(total_count: int) -> None:
    """
    상단 히어로 카드 렌더링.
    """
    st.markdown(
        f"""
        <div class="hero-card">
            <div class="hero-title">🔍 AI 사고 통합검색</div>
            <div class="hero-subtitle">
                {DATA_YEAR_MIN}~{DATA_YEAR_MAX}년 사고 데이터 {total_count:,}건 기반 ·
                구글처럼 질문하면 데이터에 근거한 답을 바로 제시합니다.
            </div>
            <div style="margin-top:12px;">
                <span class="hero-badge">자연어 통합검색</span>
                <span class="hero-badge">연도·동의어 자동 인식</span>
                <span class="hero-badge">데이터 근거 답변</span>
                <span class="hero-badge">고위험·ERT 분석</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_basic_metrics(df: pd.DataFrame) -> None:
    """
    사고 데이터 기본 메트릭 표시.
    """
    total_cnt = len(df)

    years = []
    if "발생년도" in df.columns:
        years = sorted([
            int(y)
            for y in pd.to_numeric(df["발생년도"], errors="coerce").dropna().unique()
        ])

    top_type_val = "-"
    if "재해유형" in df.columns and not df.empty:
        s = df["재해유형"].dropna().astype(str).str.strip()
        s = s[s != ""]
        if not s.empty:
            top_type_val = s.value_counts().index[0]

    top_name_val = "-"
    if "사고명" in df.columns and not df.empty:
        s = df["사고명"].dropna().astype(str).str.strip()
        s = s[s != ""]
        if not s.empty:
            top_name_val = s.value_counts().index[0]

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.metric("총 사고 건수", total_cnt)

    with c2:
        st.metric("연도 범위", f"{min(years)} - {max(years)}" if years else "-")

    with c3:
        st.metric("최다 재해유형", top_type_val)

    with c4:
        st.metric("최다 사고명", top_name_val)


def render_top_tables(df: pd.DataFrame) -> None:
    """
    사고명, 재해유형, 원인 등 주요 상위 테이블 표시.
    """
    cols = st.columns(2)

    with cols[0]:
        if "사고명" in df.columns:
            st.subheader("사고명 상위 10")
            st.table(build_value_count_table(df, "사고명", 10))

    with cols[1]:
        if "재해유형" in df.columns:
            st.subheader("재해유형 상위 10")
            st.table(build_value_count_table(df, "재해유형", 10))

    cols = st.columns(2)

    with cols[0]:
        if "사고단계" in df.columns:
            st.subheader("사고단계 상위 10")
            st.table(build_value_count_table(df, "사고단계", 10))

    with cols[1]:
        if "사고원인상세" in df.columns:
            st.subheader("사고원인상세 상위 10")
            st.table(build_value_count_table(df, "사고원인상세", 10))


def render_unified_search_section(
    df: pd.DataFrame,
    token: str,
    api_base_url: str,
    user_id: str,
    model_name: str,
) -> None:
    """
    사고 통합검색 메인 섹션 (구글 스타일 검색 + AI 통합 답변).
    """
    # ------------------------------------------------------------------
    # 검색바 (Enter로 바로 검색되도록 form 사용)
    # ------------------------------------------------------------------
    with st.form("unified_search_form", clear_on_submit=False):
        col_input, col_button = st.columns([5, 1])

        with col_input:
            query = st.text_input(
                "검색어",
                placeholder="예: 2023년 이후 화재 사고 몇 건이고 주요 원인은 뭐야?",
                key="unified_query_input",
                label_visibility="collapsed",
            )

        with col_button:
            submitted = st.form_submit_button("🔍 검색", use_container_width=True)

    # ------------------------------------------------------------------
    # 추천 검색어 칩
    # ------------------------------------------------------------------
    st.caption("이렇게 물어보세요")

    clicked_example = None
    chip_cols = st.columns(3)

    for i, example in enumerate(EXAMPLE_QUERIES):
        if chip_cols[i % 3].button(example, key=f"example_chip_{i}", use_container_width=True):
            clicked_example = example

    # ------------------------------------------------------------------
    # 검색 실행
    # ------------------------------------------------------------------
    final_query = None

    if submitted and query and query.strip():
        final_query = query.strip()
    elif clicked_example:
        final_query = clicked_example

    if final_query:
        if not token or not user_id:
            st.info(
                "사이드바에 GPT OSS 토큰과 User ID를 입력하면 AI 통합 답변까지 생성됩니다. "
                "지금은 데이터 검색 결과만 표시합니다."
            )

        with st.spinner("검색 및 AI 통합 답변 생성 중..."):
            st.session_state["unified_result"] = execute_unified_search(
                query=final_query,
                df=df,
                token=token,
                api_base_url=api_base_url,
                user_id=user_id,
                model_name=model_name,
            )

    # ------------------------------------------------------------------
    # 결과 렌더링
    # ------------------------------------------------------------------
    if "unified_result" not in st.session_state:
        st.markdown("---")
        st.caption(
            f"💡 {DATA_YEAR_MIN}~{DATA_YEAR_MAX}년 사고 데이터를 대상으로 "
            "자연어로 질문하면 기간·키워드를 자동 인식해 데이터 근거 답변을 생성합니다."
        )
        return

    result = st.session_state["unified_result"]
    parsed = result["parsed"]
    matched = result["matched"]

    st.markdown("---")

    # 적용된 검색 조건 표시
    if parsed["year_from"] is not None:
        period_chip = f"📅 {parsed['year_from']}~{parsed['year_to']}년"
    else:
        period_chip = f"📅 {DATA_YEAR_MIN}~{DATA_YEAR_MAX}년 전체"

    keyword_chips = "".join(
        f'<span class="search-condition">🔑 {t}</span>' for t in parsed["tokens"]
    )

    st.markdown(
        f'<span class="search-condition">{period_chip}</span>'
        f"{keyword_chips}"
        f'<span class="search-condition">📄 매칭 {len(matched):,}건 · {result["match_mode"]}</span>',
        unsafe_allow_html=True,
    )

    st.caption(f'검색어: "{result["query"]}"')

    # AI 통합 답변 (검색 결과 최상단, 구글 요약 카드 위치)
    if result["answer"]:
        with st.container(border=True):
            st.markdown("#### 🤖 AI 통합 답변")
            st.markdown(result["answer"])
            st.caption(
                f"※ {DATA_YEAR_MIN}~{DATA_YEAR_MAX}년 등록 사고 데이터에만 근거한 답변입니다. "
                "데이터에 없는 내용은 '확인 불가'로 표시됩니다."
            )

    # 매칭 데이터 통계
    if matched is None or matched.empty:
        st.warning(
            "키워드와 일치하는 사고 데이터가 없습니다. "
            "검색어를 바꾸거나 기간을 넓혀서 다시 검색해 보세요."
        )
        return

    render_basic_metrics(matched)

    chart_cols = st.columns(2)

    with chart_cols[0]:
        if "발생년도" in matched.columns:
            year_counts = (
                pd.to_numeric(matched["발생년도"], errors="coerce")
                .dropna()
                .astype(int)
                .value_counts()
                .sort_index()
            )
            if not year_counts.empty:
                st.subheader("연도별 매칭 건수")
                st.bar_chart(year_counts)

    with chart_cols[1]:
        if "재해유형" in matched.columns:
            type_counts = (
                matched["재해유형"].dropna().astype(str).str.strip()
            )
            type_counts = type_counts[type_counts != ""].value_counts().head(10)
            if not type_counts.empty:
                st.subheader("재해유형 상위 10")
                st.bar_chart(type_counts)

    with st.expander(f"📋 검색 결과 상세 데이터 ({len(matched):,}건, 관련도·최신순)", expanded=False):
        st.dataframe(
            to_display_date(matched),
            use_container_width=True,
            height=400,
        )

        csv_bytes = matched.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ 검색 결과 CSV 다운로드",
            data=csv_bytes,
            file_name="사고_통합검색_결과.csv",
            mime="text/csv",
            key="unified_download_csv",
        )

    with st.expander("📊 검색 결과 통계 상세", expanded=False):
        render_top_tables(matched)


def render_ai_summary_section(
    title: str,
    df: pd.DataFrame,
    token: str,
    api_base_url: str,
    user_id: str,
    model_name: str,
    default_external_query: str,
    key_prefix: str,
) -> None:
    """
    일반사고/고위험사고 AI 요약 공통 섹션.
    """
    st.header(title)

    if df is None or df.empty:
        st.info("해당 조건에 맞는 사고 데이터가 없습니다.")
        return

    render_basic_metrics(df)
    render_top_tables(df)

    st.markdown("---")

    if st.button(f"{title} 생성", key=f"{key_prefix}_generate_ai_summary"):
        summary_text = build_data_summary(df)

        prompt = (
            f"다음은 '{title}' 대상 사고 데이터({DATA_YEAR_MIN}~{DATA_YEAR_MAX}년) 요약입니다.\n\n"
            "아래 데이터에 근거하여 안전 분석 보고서를 작성하세요.\n\n"
            "작성 기준:\n"
            "1. 제공된 데이터에 있는 내용만 근거로 삼으세요.\n"
            "2. 데이터에 없는 원인, 수치, 출처, 날짜는 지어내지 마세요.\n"
            "3. 확인 불가한 내용은 '제공된 데이터만으로는 확인 불가'라고 표시하세요.\n"
            "4. 주요 트렌드, 반복 발생 패턴, 취약 공정/단계, 재발 가능성이 높은 요인을 정리하세요.\n"
            "5. 개선방안은 현장에서 실행 가능한 조치 중심으로 작성하세요.\n"
            "6. 보고서 형식은 다음 구조를 따르세요.\n"
            "   - 핵심 요약\n"
            "   - 주요 통계 및 패턴\n"
            "   - 반복 원인 및 취약점\n"
            "   - 우선관리 대상\n"
            "   - 개선방안\n"
            "   - 추가 확인 필요사항\n\n"
            "[데이터 요약]\n"
            f"{summary_text}"
        )

        with st.spinner("AI 요약 생성 중..."):
            result = call_llm(
                prompt=prompt,
                api_base_url=api_base_url,
                token=token,
                user_id=user_id,
                model_name=model_name,
                temperature=0.2,
                max_tokens=1800,
            )

        st.markdown("### AI 요약 결과")
        st.markdown(result)

    st.markdown("---")
    st.subheader("외부 사고 사례 정리")

    st.caption(
        "주의: DS-OSS 모델에 실제 웹 검색 기능이 연결되어 있지 않은 경우, "
        "AI가 출처를 확인하지 못할 수 있습니다. 확인되지 않은 링크나 사례는 실제 사례로 단정하지 않도록 제한했습니다."
    )

    ext_query = st.text_input(
        "외부 사고 검색어",
        default_external_query,
        key=f"{key_prefix}_external_query",
    )

    ext_num = st.slider(
        "사례 개수",
        1,
        5,
        3,
        key=f"{key_prefix}_external_num",
    )

    if st.button("외부 사고 사례 정리", key=f"{key_prefix}_external_search"):
        ext_prompt = (
            f"'{ext_query}'와 관련된 사고 사례 {ext_num}건을 정리하세요.\n\n"
            "작성 기준:\n"
            "1. 실제 출처를 확인할 수 없는 사례는 실제 사례처럼 단정하지 마세요.\n"
            "2. 출처 URL을 모르면 '출처 확인 필요'라고 표시하세요.\n"
            "3. 가짜 기관명, 가짜 보고서명, 가짜 URL, 확인되지 않은 날짜를 만들지 마세요.\n"
            "4. 확인된 사례와 일반적 예방 교훈을 명확히 구분하세요.\n"
            "5. 각 사례는 다음 형식으로 작성하세요.\n"
            "   - 사고 개요\n"
            "   - 주요 원인\n"
            "   - 피해 또는 영향\n"
            "   - 예방 교훈\n"
            "   - 참고 출처 또는 출처 확인 필요\n"
        )

        with st.spinner("외부 사고 사례 정리 중..."):
            cases = call_llm(
                prompt=ext_prompt,
                api_base_url=api_base_url,
                token=token,
                user_id=user_id,
                model_name=model_name,
                temperature=0.2,
                max_tokens=2000,
            )

        st.markdown("#### 외부 사고 사례 정리 결과")
        st.markdown(cases)


def render_ert_section(
    ert_data: pd.DataFrame,
    token: str,
    api_base_url: str,
    user_id: str,
    model_name: str,
) -> None:
    """
    ERT 고등급 출동 분석 섹션.
    """
    st.header("ERT 고등급 출동 분석")

    if ert_data is None or ert_data.empty:
        st.info("ERT 출동 데이터가 없습니다.")
        return

    filtered_ert = prepare_ert_high_grade_data(ert_data)

    if filtered_ert.empty:
        st.info("RI등급 A, B, C+에 해당하는 ERT 출동 데이터가 없습니다.")
        return

    st.metric("고등급 출동 건수", len(filtered_ert))

    if "RI등급" in filtered_ert.columns:
        st.subheader("RI등급별 출동 건수")
        ert_counts = filtered_ert["RI등급"].value_counts().sort_index()
        st.bar_chart(ert_counts)

    st.subheader("ERT 데이터 테이블")
    st.dataframe(
        to_display_date(filtered_ert),
        use_container_width=True,
        height=300,
    )

    st.subheader("ERT 세부 데이터 조회")

    if "RI등급" in filtered_ert.columns:
        grade_options = ["(전체)"] + sorted(
            filtered_ert["RI등급"].dropna().astype(str).unique().tolist()
        )
        selected_grade = st.selectbox(
            "RI등급 선택",
            grade_options,
            key="ert_grade_select",
        )

        ert_details = filtered_ert.copy()

        if selected_grade != "(전체)":
            ert_details = ert_details[ert_details["RI등급"] == selected_grade]

        st.dataframe(
            to_display_date(ert_details),
            use_container_width=True,
            height=250,
        )

    if st.button("ERT 데이터 AI 분석", key="ert_ai_analysis_button"):
        ert_summary = build_ert_summary(filtered_ert)

        prompt_ert = (
            "다음은 ERT 고등급 출동 데이터 요약 정보입니다.\n\n"
            "아래 데이터에 근거하여 ERT 고등급 출동의 주요 패턴과 개선방안을 분석하세요.\n\n"
            "작성 기준:\n"
            "1. 데이터에 없는 원인이나 수치를 단정하지 마세요.\n"
            "2. RI등급별 특징, 반복 출동 유형, 시간/연도별 패턴을 중심으로 분석하세요.\n"
            "3. 고등급 출동을 줄이기 위한 선제 예방활동을 제시하세요.\n"
            "4. 추가 확인이 필요한 데이터 항목도 함께 제시하세요.\n\n"
            "[ERT 데이터 요약]\n"
            f"{ert_summary}"
        )

        with st.spinner("ERT 데이터 AI 분석 중..."):
            result_ert = call_llm(
                prompt=prompt_ert,
                api_base_url=api_base_url,
                token=token,
                user_id=user_id,
                model_name=model_name,
                temperature=0.2,
                max_tokens=1800,
            )

        st.markdown("#### AI 분석 결과")
        st.markdown(result_ert)


def select_injury_type(df: pd.DataFrame, key: str) -> str:
    """
    재해유형 선택 위젯을 렌더링하고 선택값을 반환합니다.
    """
    if "재해유형" not in df.columns:
        st.warning("데이터에 '재해유형' 컬럼이 없어 재해유형 필터를 적용할 수 없습니다.")
        return "(전체)"

    type_values = (
        df["재해유형"]
        .dropna()
        .astype(str)
        .str.strip()
    )
    type_values = type_values[type_values != ""]

    type_options = ["(전체)"] + sorted(type_values.unique().tolist())

    return st.selectbox(
        "재해유형 선택",
        type_options,
        key=key,
    )


# ============================================================================
# Main
# ============================================================================

def main() -> None:
    st.set_page_config(
        page_title="AI 사고 통합검색",
        page_icon="🔍",
        layout="wide",
    )

    render_style()

    # ------------------------------------------------------------------------
    # 기본값 로드
    # ------------------------------------------------------------------------
    default_api_base_url = get_secret_or_env(
        "GPT_OSS_BASE_URL",
        "GPT_OSS_BASE_URL",
        DEFAULT_API_BASE_URL,
    )

    default_token = get_secret_or_env(
        "GPT_OSS_TOKEN",
        "GPT_OSS_TOKEN",
        DEFAULT_CREDENTIAL_KEY,
    )

    default_user_id = get_secret_or_env(
        "GPT_OSS_USER_ID",
        "GPT_OSS_USER_ID",
        DEFAULT_USER_ID,
    )

    default_model = get_secret_or_env(
        "GPT_OSS_MODEL",
        "GPT_OSS_MODEL",
        DEFAULT_MODEL,
    )

    # ------------------------------------------------------------------------
    # 사이드바
    # ------------------------------------------------------------------------
    with st.sidebar:
        st.header("DS-OSS GPT 설정")

        api_base_url = st.text_input(
            "API Base URL",
            value=default_api_base_url,
            help="DS-OSS 기본 URL은 기존 값으로 유지했습니다.",
        )

        token = st.text_input(
            "GPT OSS 토큰",
            value=default_token,
            type="password",
            help="토큰은 코드에 저장하지 말고 secrets 또는 환경변수 사용을 권장합니다.",
        )

        user_id = st.text_input(
            "User ID",
            value=default_user_id,
            placeholder="AD ID 입력",
        )

        model_name = st.text_input(
            "모델 이름",
            value=default_model,
        )

        st.markdown("---")
        st.header("데이터 옵션")

        upload_file = st.file_uploader(
            "사고 데이터 업로드",
            type=["xlsx", "csv"],
            help="업로드하지 않으면 기본 경로의 ITS등록사고_streamlit.xlsx를 사용합니다.",
        )

        year_range = st.slider(
            "분석 대상 연도",
            min_value=DATA_YEAR_MIN,
            max_value=DATA_YEAR_MAX,
            value=(DATA_YEAR_MIN, DATA_YEAR_MAX),
            help=f"기본 분석 범위는 {DATA_YEAR_MIN}~{DATA_YEAR_MAX}년입니다.",
        )

        st.markdown("---")
        st.caption("토큰과 User ID는 소스코드에 하드코딩하지 않는 것을 권장합니다.")

    # ------------------------------------------------------------------------
    # 데이터 로드
    # ------------------------------------------------------------------------
    if upload_file:
        df, uploaded_ert_data = load_uploaded_data(upload_file)
        ert_data = uploaded_ert_data

        if not df.empty:
            st.success("사용자 사고 데이터가 성공적으로 업로드되었습니다.")

        if ert_data.empty:
            st.info("업로드 파일에서 ERT출동 시트를 찾지 못했습니다. ERT 분석에는 기본 ERT 데이터를 사용하지 않습니다.")

    else:
        df = load_default_accident_data()
        ert_data = load_default_ert_data()

    if df is None or df.empty:
        st.warning("분석할 사고 데이터가 없습니다. 기본 파일 경로 또는 업로드 파일을 확인하세요.")
        st.stop()

    # 2021~2026년 기준 기간 스코프 적용 (연도 미상 행은 유지)
    df, unknown_year_count = apply_year_scope(df)

    # 사이드바 연도 범위 추가 적용
    if year_range != (DATA_YEAR_MIN, DATA_YEAR_MAX) and "발생년도" in df.columns:
        df = filter_by_year(df, year_range[0], year_range[1])

    if df.empty:
        st.warning("선택한 연도 범위에 해당하는 사고 데이터가 없습니다.")
        st.stop()

    df_original = df.copy()

    render_hero(len(df_original))

    if unknown_year_count > 0:
        st.caption(
            f"※ 발생 연도를 확인할 수 없는 {unknown_year_count}건은 데이터 유실 방지를 위해 "
            "검색 대상에 포함되어 있습니다 (연도 필터에는 걸리지 않음)."
        )

    # ------------------------------------------------------------------------
    # 탭 구성: 통합검색이 메인
    # ------------------------------------------------------------------------
    tab_search, tab_highrisk, tab_general, tab_ert, tab_data = st.tabs([
        "🔍 사고 통합검색",
        "🔴 고위험 AI 요약",
        "🟡 등급사고 AI 요약",
        "🚨 ERT 고등급 분석",
        "📁 데이터 현황",
    ])

    with tab_search:
        render_unified_search_section(
            df=df_original,
            token=token,
            api_base_url=api_base_url,
            user_id=user_id,
            model_name=model_name,
        )

    with tab_highrisk:
        selected_type = select_injury_type(df_original, key="highrisk_injury_type")

        ai_df = filter_accident_by_risk(
            df=df_original,
            risk_type="highrisk",
            selected_type=selected_type,
        )

        render_ai_summary_section(
            title="고위험사고 AI 요약",
            df=ai_df,
            token=token,
            api_base_url=api_base_url,
            user_id=user_id,
            model_name=model_name,
            default_external_query="폭발 사고",
            key_prefix="highrisk_ai",
        )

    with tab_general:
        selected_type = select_injury_type(df_original, key="general_injury_type")

        ai_df = filter_accident_by_risk(
            df=df_original,
            risk_type="general",
            selected_type=selected_type,
        )

        render_ai_summary_section(
            title="등급사고 AI 요약",
            df=ai_df,
            token=token,
            api_base_url=api_base_url,
            user_id=user_id,
            model_name=model_name,
            default_external_query="낙상 사고",
            key_prefix="general_ai",
        )

    with tab_ert:
        render_ert_section(
            ert_data=ert_data,
            token=token,
            api_base_url=api_base_url,
            user_id=user_id,
            model_name=model_name,
        )

    with tab_data:
        st.header("데이터 현황")
        render_basic_metrics(df_original)

        if "발생년도" in df_original.columns:
            year_counts = (
                pd.to_numeric(df_original["발생년도"], errors="coerce")
                .dropna()
                .astype(int)
                .value_counts()
                .sort_index()
            )
            if not year_counts.empty:
                st.subheader("연도별 사고 건수")
                st.bar_chart(year_counts)

        st.subheader("데이터 미리보기 (상위 100건)")
        preview_df = to_display_date(df_original.head(100))
        st.dataframe(preview_df, use_container_width=True, height=300)

        render_top_tables(df_original)


if __name__ == "__main__":
    main()
