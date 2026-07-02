# 2026 화성사업장 사고 KPI 대시보드

Streamlit 기반 사고 KPI 대시보드입니다. 사고/KPI 엑셀 데이터를 읽어 KPI 카드, 월별 목표·실적 그래프, 사고 LIST를 표시하고, DS OSS API를 통한 AI 사고 분석 기능을 제공합니다.

> 이 앱은 이 레포의 Astro 사이트와 무관한 독립 앱입니다. Vercel 배포 대상이 아니며, 별도 환경에서 직접 실행해야 합니다.

## 실행 방법

```bash
cd hwaseong-kpi-dashboard
pip install -r requirements.txt
streamlit run app.py
```

## 환경변수

환경변수 또는 `.streamlit/secrets.toml`로 설정합니다. (API Key는 절대 코드에 직접 넣지 마세요.)

| 변수 | 기본값 | 설명 |
|---|---|---|
| `ACCIDENT_FILE` | `/config/work/sharedworkspace/ITS등록사고_streamlit.xlsx` | ITS 등록사고 엑셀 경로 |
| `KPI_FILE` | `/config/work/sharedworkspace/화성KPI_streamlit.xlsx` | 화성 KPI 엑셀 경로 |
| `KPI_CURRENT_YEAR` | `2026` | 기준 연도 |
| `DS_OSS_API_KEY` | (없음) | DS OSS API Key — 미설정 시 AI 분석 비활성 |
| `DS_OSS_API_BASE_URL` | 사내 기본 URL | DS OSS API Base URL |
| `DS_OSS_USER_ID` | 기본값 있음 | DS OSS User ID |
| `DS_OSS_MODEL` | `gpt-oss-120b` | 사용 모델 |
| `DS_OSS_MAX_TOKENS` | `12000` | 최대 토큰 |
| `DS_OSS_TIMEOUT` | `300` | 타임아웃(초) |
| `DS_OSS_RETRY` | `2` | 재시도 횟수 |

## 참고

- 데이터/폰트/아이콘 기본 경로는 사내 `/config/work/sharedworkspace/` 환경 기준입니다. 다른 환경에서는 `ACCIDENT_FILE`, `KPI_FILE`을 재지정하고, 한글 폰트(NotoSansKR)를 해당 경로에 두거나 코드의 폰트 경로를 수정하세요.
- 실행 시 앱 폴더 아래 `data/`, `reports/`, `uploads/` 디렉토리가 자동 생성됩니다. (git 추적 제외)
- AI 분석 이력은 `data/analysis_history.json`에 저장됩니다.
