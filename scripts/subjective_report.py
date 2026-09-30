"""사내 조직문화 진단 "주관식 응답" 분석 파이프라인.

입력 (엑셀 3종, DRM 보호 — 반드시 xlwings로 엶):
  (A) 잘하는점 (30컬럼, A~AD, 소분류/중분류/대분류/확률이 4개 슬롯 _1.._4)
  (B) 노력해야할점 (16컬럼, A~P, 소분류/중분류/대분류/확률이 각 1개 컬럼이고
      콤마로 여러 값 표시, 부서 컬럼이 맨 뒤에 있음) — (A)와 컬럼 레이아웃이 다름
  (C) 부서장에게 하고싶은말 (21컬럼, A~U)
모두 "진단부서"(또는 "진단부서명")로 부서를 식별하며, 로딩 시 이 둘을
"진단부서" 컬럼 하나로 통일합니다(unify_dept_column).

처리: xlwings로 읽기 -> pandas로 위치 기반 컬럼명 부여/집계 -> 로컬 vLLM으로
조직×카테고리 테마/부서장 의견/조직별 총평·종합 분석 요약 -> python-docx +
matplotlib로 조직별 임원보고용 .docx 생성. 조직명은 segment_loader.py/
qualitative_loader.py와 동일한 load_dept_code_lookup()으로 dept_code에
매칭되고, 그 dept_code로 객관식 설문 지표(data/survey_data.json 등)를
찾아 총평/종합 분석 LLM 프롬프트에 근거로 포함시킵니다. 최종적으로 조직별
분석 결과 전체를 dept_code로 키를 맞춰 data/subjective_analysis.json 하나에
모아 쓰는데, 이 파일을 dashboard/index.html의 "주관식분석" 탭에서 불러오면
이미 로드되어 있는 객관식 지표와 한 화면에서 통합해 볼 수 있습니다(조직별로
따로 파일을 만들지 않음 — 대시보드가 그 역할을 대신함).

실행:
    python scripts/subjective_report.py

경로/모델 등은 전부 아래 CONFIG에서 조정합니다. BASE_URL을 비워두면 LLM 요약
관련 섹션만 생략되고 나머지(집계표/차트/보고서 뼈대)는 정상적으로 생성됩니다.

테스트: xlwings 로딩 없이 전체 파이프라인을 검증하려면 main()에 넘기는
config와 함께 load_strength_weakness()/load_weakness()/load_leader_comments()를
원하는 DataFrame을 돌려주는 함수로 바꿔치기하면 됩니다 (아래 세 함수는 모듈
전역 이름으로 참조되므로 monkeypatch가 그대로 적용됩니다):

    import subjective_report as sr
    sr.load_strength_weakness = lambda path, sheet_name=None: my_fake_strength_df
    sr.load_weakness = lambda path, sheet_name=None: my_fake_weakness_df
    sr.load_leader_comments = lambda path, sheet_name=None: my_fake_leader_df
    sr.main({**sr.CONFIG, "OUT_DIR": "test_output", "BASE_URL": ""})

이렇게 하면 집계/LLM 캐시/차트/.docx 생성 로직은 실제 실행과 완전히
동일하게 동작합니다 (실행 방법과 검증 결과는 이 스크립트를 만든 커밋의
메시지에 정리되어 있습니다).
"""
from __future__ import annotations

import hashlib
import json
import random
import re
import shutil
import sys
from pathlib import Path

import pandas as pd

from segment_loader import load_dept_code_lookup

DATA_DIR = Path(__file__).resolve().parent.parent / "data"

# =============================================================
# 0. CONFIG — 여기만 고치면 됩니다
# =============================================================
CONFIG = {
    "BASE_URL": "",          # 사내 vLLM OpenAI 호환 엔드포인트(예: "http://10.x.x.x:8000/v1"). 비워두면 LLM 요약 전부 건너뜀.
    "MODEL": "thinkingcap",
    "API_KEY": "EMPTY",      # curl 테스트로 확인된 값 — "Authorization: Bearer EMPTY" 헤더가 있어야 통과됨.
                             # 비워두면("") Authorization 헤더 자체를 안 보냄(그 경우 403이 남).

    # 실제 파일명이 매번 달라서(예: "2026_1차_잘하는점_전체.xlsx") 고정 파일명 대신
    # EXCEL_DIR 안에서 파일명에 키워드가 들어간 .xlsx를 패턴으로 찾습니다
    # (아래 FILE_KEYWORDS 참고). 파일 경로를 직접 지정하고 싶으면 EXCEL_STRENGTH/
    # WEAKNESS/LEADER에 실제 경로를 넣으세요 — 넣으면 그 경로를 그대로 쓰고
    # 패턴 탐색은 하지 않습니다.
    "EXCEL_DIR": ".",
    "EXCEL_STRENGTH": None,   # (A) 잘하는점 — None이면 EXCEL_DIR에서 자동 탐색
    "EXCEL_WEAKNESS": None,   # (B) 노력해야할점 — None이면 EXCEL_DIR에서 자동 탐색
    "EXCEL_LEADER": None,     # (C) 부서장에게 하고싶은말 — None이면 EXCEL_DIR에서 자동 탐색
    "SHEET": None,  # None이면 각 워크북의 첫 번째 시트 사용

    "ORG_LEVEL": "진단부서",  # "사업부" | "실" | "팀" | "진단부서" — 보고서를 어느 조직 단위로
                              # 쪼갤지. 3개 엑셀 파일 모두 로딩 시 "진단부서" 컬럼으로 통일되므로
                              # (unify_dept_column 참고) 기본값은 "진단부서". 객관식 지표
                              # (OBJECTIVE_DATA_PATH)와 연동하려면 그 파일의 dept_name과 값이
                              # 정확히 일치해야 합니다.
    "MIN_N": 5,             # 조직 단위 응답 수가 이 미만이면 익명성 보호를 위해 해당 섹션 비표시
    "MIN_CAT_N": 3,         # 조직×카테고리 응답 수가 이 미만이면 그 카테고리는 LLM 요약 생략
    "PROB_MIN": 0.3,        # 이 확률 미만인 분류는 무시
    "SAMPLE_N": 15,         # LLM에 보여줄 응답 샘플 개수
    "TOP_K": 5,             # 보고서에 표시할 강점/개선 Top-K
    "TOP_K_AMBIVALENT": 8,  # 양가 이슈 판정에 쓰는 Top-N (요청 명세: 상위 8)
    "RANDOM_SEED": 42,      # 부서장 의견 샘플링 고정 시드

    "OUT_DIR": "output",

    "LLM_TIMEOUT_SEC": 300,
    "LLM_TEMPERATURE": 0.2,

    # 객관식 설문 집계 데이터(dept_code/dept_name/area_scores/area_yoy/response_rate
    # 등을 담은 JSON 배열) 경로. None이면 data/survey_data.json -> data/
    # dummy_survey_data.json 순으로 자동 탐색(대시보드가 쓰는 것과 동일 파일).
    # dept_code 기준으로 매칭되므로(load_dept_code_lookup 참고) 조직명 표기가
    # 살짝 달라도 안전하게 연결됩니다. 지정하면 총평/종합 분석 LLM 프롬프트에
    # 객관식 점수가 근거로 같이 들어감.
    "OBJECTIVE_DATA_PATH": None,

    # 대시보드 통합용 결과 JSON 경로. None이면 data/subjective_analysis.json에
    # 씀 — dashboard/index.html의 "주관식분석" 탭 "주관식 AI 분석 불러오기"
    # 버튼으로 이 파일을 불러오면, 조직별 종합 분석/Top5/부서장 의견 요약이
    # 이미 로드되어 있는 객관식 지표와 한 화면에 같이 나타납니다.
    "DASHBOARD_JSON_PATH": None,
}

# EXCEL_STRENGTH/WEAKNESS/LEADER가 None일 때 EXCEL_DIR 안에서 파일명을 찾는 데
# 쓰는 키워드 — 파일명 전체가 아니라 이 키워드가 "포함"되어 있는지만 봅니다
# (예: "2026_1차_노력해야할점_전체.xlsx", "노력해야할점_08연구소.xlsx" 전부 매칭).
FILE_KEYWORDS = {
    "strength": "잘하는점",
    "weakness": "노력해야할점",
    "leader": "부서장하고싶은말",
}


def resolve_excel_path(explicit_path, keyword, excel_dir):
    """explicit_path가 주어졌으면 그대로 쓰고, 아니면 excel_dir에서 파일명에
    keyword가 들어간 .xlsx를 찾습니다. 실제 파일명이 CONFIG에 적어둔 이름과
    달라도(부서/회차마다 파일명이 바뀌는 경우) 이 키워드 매칭 하나로 대응됩니다.
    여러 개가 매칭되면 가장 최근에 수정된 파일을 쓰고 나머지는 경고로 알려주며,
    하나도 없으면 찾을 수 있게 무엇을 찾고 있었는지 에러 메시지에 그대로 남깁니다."""
    if explicit_path:
        return Path(explicit_path)

    candidates = [
        p for p in Path(excel_dir).glob("*.xlsx")
        if keyword in p.stem and not p.name.startswith("~$")  # ~$는 엑셀이 열려있을 때 생기는 잠금 파일
    ]
    if not candidates:
        raise FileNotFoundError(
            f"'{excel_dir}' 폴더에서 파일명에 '{keyword}'가 들어간 .xlsx를 찾지 못했습니다. "
            f"CONFIG의 EXCEL_DIR을 확인하거나, 해당 항목에 실제 파일 경로를 직접 지정하세요."
        )
    if len(candidates) > 1:
        candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        print(f"  warning: '{keyword}' 키워드에 맞는 파일이 {len(candidates)}개 발견됨 — "
              f"가장 최근에 수정된 '{candidates[0].name}'을 사용합니다 "
              f"(나머지: {[p.name for p in candidates[1:]]})", file=sys.stderr)
    return candidates[0]


def load_objective_records(path=None):
    """객관식 설문 집계 JSON(부서별 area_scores/area_yoy/response_rate 등이 담긴
    배열 — 대시보드가 쓰는 data/survey_data.json과 동일한 구조)을 읽어
    {dept_code: record} 딕셔너리로 반환합니다. path를 안 주면 대시보드/
    segment_loader.py/qualitative_loader.py와 동일하게 data/survey_data.json을
    먼저 찾고 없으면 data/dummy_survey_data.json으로 대체합니다(load_dept_code_
    lookup()과 같은 탐색 순서 — 주관식 조직명 -> dept_code 매칭도 이 파일 기준
    이므로, 객관식 지표도 같은 파일에서 가져와야 서로 어긋나지 않습니다).
    파일이 없으면 빈 딕셔너리를 반환하고 경고만 출력합니다 — 이 경우 객관식
    지표 없이 주관식 분석만 정상적으로 진행됩니다."""
    if path:
        p = Path(path)
        if not p.exists():
            print(f"  warning: 객관식 지표 파일을 찾을 수 없습니다 — {path} (해당 섹션 생략)", file=sys.stderr)
            return {}
    else:
        p = None
        for candidate in (DATA_DIR / "survey_data.json", DATA_DIR / "dummy_survey_data.json"):
            if candidate.exists():
                p = candidate
                break
        if p is None:
            print(f"  warning: 객관식 지표 파일을 찾을 수 없습니다 — "
                  f"{DATA_DIR / 'survey_data.json'} (해당 섹션 생략)", file=sys.stderr)
            return {}
    with open(p, "r", encoding="utf-8") as f:
        records = json.load(f)
    return {r["dept_code"]: r for r in records if r.get("dept_code")}


# 색상 (Word 보고서 공통)
COLOR_NAVY = "0B1F3A"
COLOR_CORAL = "FF6B57"
COLOR_GOLD = "D4A24C"
FONT_KR = "맑은 고딕"

# 분류 체계: 중분류 코드의 첫 글자가 곧 대분류 코드입니다(예: "E1.효율" -> "E").
# 전체 대분류/중분류 라벨은 참고용으로만 들고 있고, 실제 그룹핑은 코드 문자열
# 자체("E1.효율" 등, 이미 라벨까지 포함된 형태로 들어온다는 명세 기준)를 그대로
# 키로 사용합니다 — 실제 시트의 표기가 조금 달라도(예: 코드만 있고 라벨이 없는
# 경우) 첫 글자만으로 대분류를 판정하므로 major_of()만 바꾸면 대응 가능합니다.
MAJOR_LABELS = {
    "A": "복리후생·근무여건", "B": "보상·인사", "C": "성장·교육",
    "D": "조직문화·소통", "E": "일하는 방식", "F": "사업·기술·전략",
    "G": "사회적책임", "H": "반어·냉소", "I": "비주제·무의견",
}
EXCLUDED_MAJOR = "I"    # 순위/요약에서 제외
IRONIC_MAJOR = "H"      # 순위에는 포함하되 별도 KPI 비율 산출

# =============================================================
# 1. 컬럼 위치 -> 이름 매핑 (열 순서만 여기서 바꾸면 전체에 반영됩니다)
# =============================================================
# (A) 잘하는점: A~AD, 30개 컬럼, 분류가 소분류_1..4/중분류_1..4/확률_1..4
# 4개 슬롯으로 나뉘어 있음(슬롯 내부는 콤마로 추가 분리 가능).
STRENGTH_COLUMNS = [
    "사업부", "실", "팀", "진단부서명", "raw_text", "text_norm", "status", "char_len",
    "dup_key", "dup_count",
    "소분류코드_1", "소분류코드_2", "소분류코드_3", "소분류코드_4",
    "소분류_1", "소분류_2", "소분류_3", "소분류_4",
    "중분류_1", "중분류_2", "중분류_3", "중분류_4",
    "대분류_1", "대분류_2", "대분류_3", "대분류_4",
    "확률_1", "확률_2", "확률_3", "확률_4",
]

# (B) 노력해야할점: A~P, 16개 컬럼 — 잘하는점과 컬럼 구성이 다름(슬롯 없이
# 소분류코드/소분류/중분류/대분류/확률이 각 1개 컬럼이고, 한 응답이 여러
# 카테고리에 해당하면 그 컬럼 안에서 콤마로 나열됨). 부서 컬럼도 맨 앞이
# 아니라 맨 뒤(사업부/실/팀/진단부서)에 있고, "진단부서" 컬럼이 별도로 있음.
WEAKNESS_COLUMNS = [
    "num", "raw_text", "text_norm", "status", "char_len", "dup_key", "dup_count",
    "소분류코드", "소분류", "중분류", "대분류", "확률",
    "사업부", "실", "팀", "진단부서",
]

# (C) 부서장에게 하고싶은말: A~U, 21개 컬럼
LEADER_COLUMNS = [
    "id", "연도", "진단부서코드", "진단부서명", "사업부", "실", "팀", "원문", "글자수",
    "동일응답수", "응답상태", "심각도", "심각도명", "점검대상", "점검대상점수",
    "즉시확인", "즉시확인점수", "코드", "유형", "대분류", "코드확률",
]


def assign_columns_by_position(df, column_names):
    """행 위치 기준 DataFrame(헤더 없이 읽은 raw 값)에 컬럼명을 부여합니다.
    실제 열 순서가 명세와 다르면 이 함수를 호출하는 쪽의 column_names 리스트만
    고치면 되고, 그 아래 로직은 전부 컬럼 "이름" 기준이라 영향받지 않습니다.
    시트에 컬럼이 명세보다 많으면 잘라내고, 적으면 없는 컬럼을 NaN으로 채웁니다
    (실제 파일이 명세와 살짝 다를 가능성에 대비)."""
    n = len(column_names)
    if df.shape[1] > n:
        df = df.iloc[:, :n]
    elif df.shape[1] < n:
        for _ in range(n - df.shape[1]):
            df[df.shape[1]] = None
    df = df.copy()
    df.columns = column_names
    return df


# =============================================================
# 2. 엑셀 로딩 (xlwings — DRM 때문에 openpyxl/pandas.read_excel 직접 사용 불가)
# =============================================================
def open_workbook_readonly(path):
    """xlwings.App(visible=False)로 읽기 전용으로 엽니다. 호출부에서
    반드시 finally에서 app.quit()하세요(또는 close_workbook_app 사용)."""
    import xlwings as xw
    app = xw.App(visible=False)
    try:
        book = app.books.open(str(path), read_only=True, update_links=False)
    except Exception:
        app.quit()
        raise
    return app, book


def close_workbook_app(app, book):
    try:
        book.close()
    except Exception:
        pass
    try:
        app.quit()
    except Exception:
        pass


def read_sheet_raw_values(path, sheet_name=None):
    """워크북 하나를 열어 지정 시트(없으면 첫 시트)의 used_range를 2차원
    리스트(raw values, 헤더 포함)로 반환하고 앱을 종료합니다."""
    app, book = open_workbook_readonly(path)
    try:
        sheet = book.sheets[sheet_name] if sheet_name else book.sheets[0]
        values = sheet.used_range.value
        if values is None:
            return []
        if not isinstance(values[0], list):  # 행이 1개뿐이면 xlwings가 1차원 리스트로 줌
            values = [values]
        return values
    finally:
        close_workbook_app(app, book)


def raw_values_to_dataframe(values, column_names, header_rows=1):
    """엑셀 raw values(1행 헤더 포함)를 pandas DataFrame으로 변환하고
    위치 기반으로 컬럼명을 부여합니다."""
    body = values[header_rows:]
    df = pd.DataFrame(body)
    return assign_columns_by_position(df, column_names)


def normalize_id_value(v):
    """조직 코드/이름 등 "그룹핑에 쓰이는" 컬럼 값을 문자열로 통일합니다.
    실제 엑셀에서는 같은 컬럼인데도 일부 행은 "08연구소"처럼 텍스트로,
    일부 행은 셀 서식이 숫자라 8(또는 8.0)로 들어오는 경우가 흔합니다 —
    이 상태로 그대로 두면 df.groupby(org_level)가 그룹 키를 정렬하려다
    "'<' not supported between instances of 'float' and 'str'"로 죽습니다.
    여기서 전부 문자열로 맞춰서 그 문제를 막습니다. 정수처럼 보이는
    float(예: 8.0)는 ".0"을 떼고 "8"로 변환하지만, 원래 엑셀 셀이
    "08"처럼 앞자리 0이 있었는데 숫자 서식이라 0이 사라진 경우까지는
    복원할 수 없습니다(그 정보 손실은 엑셀 자체에서 이미 발생한 것) —
    조직 코드 앞자리 0이 실제로 의미가 있다면 해당 컬럼을 엑셀에서
    텍스트 서식으로 입력해 두는 쪽이 근본적인 해결책입니다."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


ID_COLUMNS_TO_NORMALIZE = ["사업부", "실", "팀", "진단부서명", "진단부서", "진단부서코드", "dup_key"]


def normalize_id_columns(df):
    for col in ID_COLUMNS_TO_NORMALIZE:
        if col in df.columns:
            df[col] = df[col].apply(normalize_id_value)
    return df


def unify_dept_column(df):
    """레이아웃마다 부서 컬럼명이 다릅니다 — 노력해야할점(신규)은 "진단부서",
    잘하는점/부서장의견(기존)은 "진단부서명". ORG_LEVEL="진단부서"를 세 파일
    공통 집계 기준으로 쓸 수 있도록, "진단부서"가 없으면 "진단부서명"에서
    복사해 항상 "진단부서" 컬럼이 존재하도록 통일합니다."""
    if "진단부서" not in df.columns and "진단부서명" in df.columns:
        df["진단부서"] = df["진단부서명"]
    return df


def _load_sw_common(path, sheet_name, column_names):
    values = read_sheet_raw_values(path, sheet_name)
    if not values:
        return pd.DataFrame(columns=column_names)
    df = raw_values_to_dataframe(values, column_names)
    df = df[df["status"].astype(str).str.startswith("10", na=False)].reset_index(drop=True)
    df = unify_dept_column(df)
    return normalize_id_columns(df)


def load_strength_weakness(path, sheet_name=None):
    """(A) 잘하는점 로더 — 30컬럼 4-슬롯 레이아웃. status가 "10"으로 시작하는
    유효 응답만 남깁니다."""
    return _load_sw_common(path, sheet_name, STRENGTH_COLUMNS)


def load_weakness(path, sheet_name=None):
    """(B) 노력해야할점 로더 — 16컬럼 단일 컬럼(콤마 구분) 레이아웃. 잘하는점과
    컬럼 구성이 달라서 별도 함수로 분리했습니다. status가 "10"으로 시작하는
    유효 응답만 남깁니다."""
    return _load_sw_common(path, sheet_name, WEAKNESS_COLUMNS)


def load_leader_comments(path, sheet_name=None):
    """(C) 로더. 응답상태가 "10"으로 시작하는 유효 응답만 남깁니다."""
    values = read_sheet_raw_values(path, sheet_name)
    if not values:
        return pd.DataFrame(columns=LEADER_COLUMNS)
    df = raw_values_to_dataframe(values, LEADER_COLUMNS)
    df = df[df["응답상태"].astype(str).str.startswith("10", na=False)].reset_index(drop=True)
    df = unify_dept_column(df)
    return normalize_id_columns(df)


# =============================================================
# 3. 콤마 분리 / 카테고리 추출 (순수 함수 — 단위 테스트하기 쉽게 분리)
# =============================================================
def split_cell_list(cell):
    """셀 값을 ","로 분리한 문자열 리스트로. 빈 값/NaN이면 빈 리스트."""
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return []
    s = str(cell).strip()
    if s == "" or s.lower() == "nan":
        return []
    return [v.strip() for v in s.split(",") if v.strip() != ""]


def parse_prob_list(cell, n_expected):
    """확률 셀을 n_expected개의 float 리스트로 정규화합니다.
    - 콤마로 나뉜 값 개수가 n_expected와 다르면 "첫 값을 전체에 사용"합니다.
    - 값이 1을 넘으면 %로 보고 100으로 나눕니다.
    - 파싱 불가한 값은 None으로 채웁니다."""
    if n_expected <= 0:
        return []
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return [None] * n_expected

    if isinstance(cell, (int, float)):
        raw_vals = [float(cell)]
    else:
        parts = split_cell_list(cell)
        raw_vals = []
        for p in parts:
            try:
                raw_vals.append(float(p))
            except ValueError:
                raw_vals.append(None)

    raw_vals = [(v / 100.0 if (v is not None and v > 1) else v) for v in raw_vals]

    if len(raw_vals) == n_expected:
        return raw_vals
    first = raw_vals[0] if raw_vals else None
    return [first] * n_expected


def major_of(midcat):
    """중분류 문자열(예: "E1.효율")의 첫 글자로 대분류 코드를 판정합니다."""
    if not midcat:
        return None
    return midcat.strip()[0].upper()


def extract_categories_from_row(row, prob_min):
    """응답 한 행에서 {중분류: 확률} 딕셔너리를 만듭니다. 두 레이아웃을 모두
    지원합니다: (1) 소분류_n/중분류_n/확률_n 4개 슬롯 레이아웃(잘하는점, 슬롯
    내부 콤마 분리까지 처리), (2) 중분류/확률이 각각 1개 컬럼이고 콤마로 여러
    값이 들어있는 단일 컬럼 레이아웃(노력해야할점). 같은 중분류가 여러
    슬롯/콤마 항목에 걸쳐 중복되면 최댓값만 남기고, PROB_MIN 미만은 제외합니다."""
    best = {}
    if "중분류_1" in row:
        for n in (1, 2, 3, 4):
            mids = split_cell_list(row.get(f"중분류_{n}"))
            if not mids:
                continue
            probs = parse_prob_list(row.get(f"확률_{n}"), len(mids))
            for mid, prob in zip(mids, probs):
                if prob is None or prob < prob_min:
                    continue
                if mid not in best or prob > best[mid]:
                    best[mid] = prob
        return best

    mids = split_cell_list(row.get("중분류"))
    if mids:
        probs = parse_prob_list(row.get("확률"), len(mids))
        for mid, prob in zip(mids, probs):
            if prob is None or prob < prob_min:
                continue
            if mid not in best or prob > best[mid]:
                best[mid] = prob
    return best


# =============================================================
# 4. 집계 (pandas — LLM 사용 안 함)
# =============================================================
def explode_categories(df, org_level, prob_min, source):
    """(A)/(B) DataFrame을 "응답 1건 x 카테고리 1개"로 펼친 long-format
    DataFrame으로 변환합니다. 컬럼: org, dup_key, char_len, text, midcat,
    major, prob, source. 카테고리가 하나도 없는 행(전부 PROB_MIN 미만 등)은
    빠지므로, 응답 수 집계는 이 결과가 아니라 원본 df 기준으로 해야 합니다."""
    records = []
    for _, row in df.iterrows():
        cats = extract_categories_from_row(row, prob_min)
        for mid, prob in cats.items():
            records.append({
                "org": row.get(org_level),
                "dup_key": row.get("dup_key"),
                "char_len": row.get("char_len"),
                "text": row.get("text_norm") or row.get("raw_text"),
                "midcat": mid,
                "major": major_of(mid),
                "prob": prob,
                "source": source,
            })
    return pd.DataFrame(records, columns=["org", "dup_key", "char_len", "text", "midcat", "major", "prob", "source"])


def compute_company_baseline(df, org_level, prob_min):
    """카테고리별 전사 언급 비율(= 전사 전체에서 그 카테고리가 매겨진
    "응답 수" / 전사 전체 유효 응답 수). org별 비율의 평균이 아니라 응답
    풀 전체를 하나로 본 비율입니다."""
    total_n = len(df)
    if total_n == 0:
        return {}
    counts = {}
    for _, row in df.iterrows():
        cats = extract_categories_from_row(row, prob_min)
        for mid in cats:
            counts[mid] = counts.get(mid, 0) + 1
    return {mid: cnt / total_n for mid, cnt in counts.items()}


def aggregate_org_strength_weakness(df, org_level, prob_min, min_n, company_baseline):
    """조직별 {org: {"n":, "counts": {midcat:count}, "pct": {midcat:pct},
    "deviation_pp": {midcat: %p 편차}, "low_sample": bool,
    "midcat_dup_keys": {midcat: [dup_key,...]}}}."""
    result = {}
    for org, g in df.groupby(org_level, sort=False):
        n = len(g)
        counts = {}
        dup_keys_by_mid = {}
        for _, row in g.iterrows():
            cats = extract_categories_from_row(row, prob_min)
            for mid in cats:
                counts[mid] = counts.get(mid, 0) + 1
                dup_keys_by_mid.setdefault(mid, []).append(row.get("dup_key"))
        pct = {mid: cnt / n for mid, cnt in counts.items()} if n else {}
        deviation_pp = {mid: (pct[mid] - company_baseline.get(mid, 0.0)) * 100 for mid in pct}
        result[org] = {
            "n": n,
            "counts": counts,
            "pct": pct,
            "deviation_pp": deviation_pp,
            "low_sample": n < min_n,
            "dup_keys_by_mid": dup_keys_by_mid,
        }
    return result


def top_categories(pct_dict, top_k, exclude_major=EXCLUDED_MAJOR):
    """대분류 I(비주제)를 제외하고 비율 내림차순 Top-K 카테고리 리스트
    ([(midcat, pct), ...])를 반환합니다."""
    items = [(mid, p) for mid, p in pct_dict.items() if major_of(mid) != exclude_major]
    items.sort(key=lambda x: x[1], reverse=True)
    return items[:top_k]


def compute_ambivalent_issues(strength_pct, weakness_pct, top_n):
    """같은 중분류가 강점 상위 top_n과 개선 상위 top_n에 동시에 있으면 양가 이슈."""
    strong_set = {mid for mid, _ in top_categories(strength_pct, top_n)}
    weak_set = {mid for mid, _ in top_categories(weakness_pct, top_n)}
    return sorted(strong_set & weak_set)


def compute_ironic_rate(weakness_df, org_level, prob_min):
    """조직별 "개선 응답 중 H(반어·냉소) 분류가 1개 이상 있는 응답 비율"."""
    result = {}
    for org, g in weakness_df.groupby(org_level, sort=False):
        n = len(g)
        if n == 0:
            result[org] = 0.0
            continue
        ironic = 0
        for _, row in g.iterrows():
            cats = extract_categories_from_row(row, prob_min)
            if any(major_of(mid) == IRONIC_MAJOR for mid in cats):
                ironic += 1
        result[org] = ironic / n
    return result


def aggregate_leader_stats(leader_df, org_level, min_n):
    """조직별 부서장 의견 통계: 응답 수, 심각도 0~3 분포, 점검대상/즉시확인
    건수, 심각도>=1인 응답의 유형별 상위 3."""
    result = {}
    for org, g in leader_df.groupby(org_level, sort=False):
        n = len(g)
        severity = pd.to_numeric(g["심각도"], errors="coerce").fillna(0).astype(int)
        severity_dist = {lvl: int((severity == lvl).sum()) for lvl in (0, 1, 2, 3)}
        check_target_n = int((g["점검대상"].astype(str).str.strip().str.upper() == "O").sum())
        immediate_n = int((g["즉시확인"].astype(str).str.strip().str.upper() == "O").sum())

        high_sev = g[severity >= 1]
        type_counts = high_sev["유형"].dropna().astype(str).value_counts()
        top_types = list(type_counts.head(3).items())

        result[org] = {
            "n": n,
            "severity_dist": severity_dist,
            "check_target_n": check_target_n,
            "immediate_n": immediate_n,
            "top_types": top_types,
            "low_sample": n < min_n,
        }
    return result


# =============================================================
# 5. LLM (요약·총평 전용 — 분류는 이미 끝난 데이터를 그대로 사용)
# =============================================================
LLM_SYSTEM_PROMPT = (
    "당신은 사내 조직문화 진단 결과를 임원에게 보고하는 HR 담당자를 돕는 어시스턴트입니다. "
    "반드시 제공된 응답/수치에 있는 내용만 사용하고, 추측하거나 새로운 사실을 만들지 마세요. "
    "숫자는 제공된 값을 그대로 인용하고 스스로 계산하거나 새로 만들지 마세요. "
    "표면적인 문장 재요약에 그치지 말고, 제공된 응답들에서 반복되는 구체적 패턴(어떤 상황/대상/행동이 "
    "반복 언급되는지)을 근거로 짚고, 그 패턴이 조직 운영에 시사하는 바까지 한 단계 더 들어가 분석하세요. "
    "단, 근거로 삼을 수 있는 것은 어디까지나 제공된 응답과 수치뿐이며, 원인을 단정하지 말고 "
    "'~로 보임', '~일 가능성' 등 응답에 근거한 추정으로 표현하세요. "
    "개인이 특정될 수 있는 표현(이름, 특정 사건의 세부 정황 등)은 일반화해서 표현하세요. "
    "한국어로, 개조식(명사형 종결)으로 작성하되 각 항목은 근거-시사점이 드러나도록 1~2문장으로 구체적으로 쓰세요."
)

THINK_TAG_RE = re.compile(r"<think>.*?</think>", re.DOTALL | re.IGNORECASE)


def strip_think_tags(text):
    return THINK_TAG_RE.sub("", text or "").strip()


def _cache_path(out_dir):
    return Path(out_dir) / "_llm_cache.json"


def load_llm_cache(out_dir):
    path = _cache_path(out_dir)
    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_llm_cache(out_dir, cache):
    path = _cache_path(out_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def prompt_hash(system, user, model):
    h = hashlib.sha256()
    h.update(system.encode("utf-8"))
    h.update(b"\x00")
    h.update(user.encode("utf-8"))
    h.update(b"\x00")
    h.update(model.encode("utf-8"))
    return h.hexdigest()


def is_degenerate_llm_output(text):
    """LLM 응답이 정상적인 문장이 아니라 같은 문자/토큰이 반복되는 깨진
    출력(서버의 채팅 템플릿 미적용, 모델 로딩 오류 등으로 발생)인지 휴리스틱으로
    판별합니다. API 호출 자체는 성공(예외 없음)해도 내용이 이런 경우가 있어서
    별도로 걸러줘야 합니다."""
    stripped = text.strip()
    if len(stripped) < 5:
        return False
    if len(set(stripped)) / len(stripped) < 0.15:
        return True
    if re.search(r"(.)\1{9,}", stripped):
        return True
    return False


def call_llm(system, user, cache, config=CONFIG):
    """LLM 호출 — BASE_URL이 비어있거나 호출/생성에 실패하면 빈 문자열을 반환합니다
    (이 함수를 쓰는 쪽에서 반드시 "빈 문자열 = 요약 없음, 섹션 생략"으로
    처리해야 합니다). 프롬프트(system+user+model) 해시로 캐시해서 같은
    입력에 대해 재실행 시 다시 호출하지 않습니다."""
    if not config["BASE_URL"]:
        return ""

    key = prompt_hash(system, user, config["MODEL"])
    if key in cache:
        return cache[key]

    try:
        import requests
        headers = {"Content-Type": "application/json"}
        if config["API_KEY"]:
            headers["Authorization"] = f"Bearer {config['API_KEY']}"
        payload = {
            "model": config["MODEL"],
            "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}],
            "temperature": config["LLM_TEMPERATURE"],
        }
        # 사내망 내부 주소(예: 10.x.x.x)인데도 윈도우에 설정된 시스템 프록시를
        # requests가 자동으로 타면서 그 프록시가 내부 IP 접근을 막아 403을
        # 내는 경우가 있음(curl.exe는 기본적으로 시스템 프록시를 안 써서 이
        # 문제가 안 생김). proxies={}를 명시해서 이 요청만큼은 프록시를 절대
        # 타지 않도록 강제함.
        resp = requests.post(f"{config['BASE_URL']}/chat/completions", headers=headers,
                              json=payload, timeout=config["LLM_TIMEOUT_SEC"],
                              proxies={"http": None, "https": None})
        if not resp.ok:
            raise RuntimeError(f"{resp.status_code} {resp.reason} — 응답 본문: {resp.text[:500]}")
        text = strip_think_tags(resp.json()["choices"][0]["message"]["content"])
        if is_degenerate_llm_output(text):
            print(f"  warning: LLM 응답이 비정상(반복/깨짐)으로 판단되어 무시합니다: {text[:60]!r}",
                  file=sys.stderr)
            text = ""
    except Exception as e:
        print(f"  warning: LLM 호출 실패 — {e}", file=sys.stderr)
        text = ""

    cache[key] = text
    return text


def sample_texts_for_theme(long_df, org, midcat, sample_n):
    """조직×카테고리 테마 요약용 샘플: 확률 내림차순 -> dup_key 중복 제거
    -> 글자수 8~300 -> 상위 sample_n건."""
    sub = long_df[(long_df["org"] == org) & (long_df["midcat"] == midcat)].copy()
    sub = sub.sort_values("prob", ascending=False)
    sub = sub.drop_duplicates(subset="dup_key", keep="first")
    sub = sub[sub["char_len"].apply(lambda v: isinstance(v, (int, float)) and 8 <= v <= 300)]
    return sub["text"].head(sample_n).tolist()


def summarize_category_theme(org, midcat, texts, cache, config=CONFIG):
    if not texts:
        return ""
    user = (
        f"조직: {org}\n카테고리: {midcat}\n"
        f"아래는 이 조직의 해당 카테고리로 분류된 실제 응답입니다 (총 {len(texts)}건):\n"
        + "\n".join(f"- {t}" for t in texts)
        + "\n\n위 응답들을 분석해 반복되는 핵심 테마를 최대 3개 뽑고, 각 테마마다 다음 3가지를 "
          "한 항목으로 묶어 작성하세요: (1) 테마 내용, (2) 응답에서 반복적으로 등장하는 구체적 표현/상황 "
          "(근거), (3) 이 테마가 조직 운영 측면에서 시사하는 바. 형식:\n"
          "- [테마명] 내용 요약 — 근거: ... / 시사점: ..."
    )
    return call_llm(LLM_SYSTEM_PROMPT, user, cache, config)


def summarize_leader_opinions(org, rows, config=CONFIG, cache=None, seed=None):
    """심각도 1~2, 즉시확인!=O인 응답만 사용(심각도 3/즉시확인 건은 이 함수에
    절대 전달하지 마세요 — 호출부에서 이미 필터링된 rows만 넘겨야 합니다)."""
    if not rows:
        return ""
    texts = list(dict.fromkeys(rows))  # 원문 중복 제거(순서 보존)
    texts = [t for t in texts if isinstance(t, str) and 8 <= len(t) <= 300]
    if not texts:
        return ""
    rng = random.Random(seed if seed is not None else config["RANDOM_SEED"])
    sample = texts if len(texts) <= config["SAMPLE_N"] else rng.sample(texts, config["SAMPLE_N"])
    user = (
        f"조직: {org}\n아래는 부서장에게 하고 싶은 말 중 심각도 1~2(즉시확인 대상 제외)만 모은 "
        f"실제 응답입니다 (총 {len(sample)}건):\n"
        + "\n".join(f"- {t}" for t in sample)
        + "\n\n유형별 핵심 요청을 최대 3개 뽑고, 각 항목마다 (1) 요청 내용, (2) 응답에 드러난 배경/근본 "
          "원인으로 보이는 것, (3) 시급성이나 우선순위에 대한 판단을 함께 작성하세요. 형식:\n"
          "- [유형] 요청 내용 — 배경: ... / 우선순위: ..."
    )
    return call_llm(LLM_SYSTEM_PROMPT, user, cache or {}, config)


def summarize_org_overview(org, facts_text, config=CONFIG, cache=None):
    user = (
        f"조직: {org}\n아래는 이 조직의 확정된 집계 수치와 요약입니다(주관식 응답 집계 및, 있는 경우 "
        f"객관식 설문 점수 포함) — 여기 없는 숫자나 사실은 만들지 말고 그대로만 활용하세요:\n{facts_text}\n\n"
        "이 내용을 바탕으로 아래 4개 항목을 작성해 주세요. 객관식 점수가 함께 제공된 경우, 주관식에서 "
        "드러난 패턴과 객관식 점수/전년비 흐름을 연결해서 설명하되 근거 없는 인과관계 단정은 피하고 "
        "'~와 함께 나타남', '~와 궤를 같이함' 등으로 표현하세요:\n"
        "총평: 이 조직의 현재 상태를 종합한 2~3문장\n"
        "구체적 근거: 총평의 근거가 된 수치/패턴을 2~3개, 각 한 줄\n"
        "시사점: 이 수치/패턴이 의미하는 조직적 함의 3개, 각 1~2문장(근거 포함)\n"
        "제언: 실행 가능한 수준의 제언 3개, 각 한 줄\n"
        "형식:\n총평: ...\n구체적 근거:\n- ...\n시사점:\n- ...\n제언:\n- ..."
    )
    return call_llm(LLM_SYSTEM_PROMPT, user, cache or {}, config)


def summarize_integrated_narrative(org, facts_text, config=CONFIG, cache=None):
    """주관식 분석 결과와 객관식 설문 점수를 하나로 합쳐, 개조식/목록이 아니라
    자연스럽게 읽히는 5문장짜리 서술형 문단으로 요약합니다. build_org_report()
    (.docx) 최상단과 write_subjective_analysis_json() 출력(대시보드 연동)
    양쪽에 "종합 분석"으로 노출되는, 이 조직에 대한 가장 압축된 한 장 요약입니다."""
    user = (
        f"조직: {org}\n아래는 이 조직의 주관식 응답 집계와(있는 경우) 객관식 설문 점수를 합친 "
        f"확정 수치입니다 — 여기 없는 숫자나 사실은 만들지 말고 그대로만 활용하세요:\n{facts_text}\n\n"
        "이 내용을 주관식(강점/개선/부서장 의견)과 객관식(설문 점수) 결과를 모두 아우르는 하나의 "
        "종합 분석으로, 정확히 5개 문장으로 구성된 서술형 문단으로 작성해 주세요. 개조식이나 "
        "목록(- 등)은 쓰지 말고 문장과 문장이 자연스럽게 이어지는 글로 쓰세요. 각 문장은 서로 다른 "
        "근거(수치/패턴)를 담아야 하고, 객관식 점수가 있다면 주관식 패턴과 연결해서 설명하되 근거 "
        "없는 인과관계 단정은 피하세요. 마지막 문장은 이 조직에 대한 종합적 시사점으로 마무리하세요."
    )
    return call_llm(LLM_SYSTEM_PROMPT, user, cache or {}, config)


def build_overview_facts_text(org, stats, objective=None):
    """조직별 임원 총평 프롬프트에 넣을 "확정 수치" 텍스트를 만듭니다. objective가
    주어지면(load_objective_records() 결과에서 해당 조직명을 찾은 레코드) 객관식
    설문 점수/전년비/참여율도 함께 근거로 포함시킵니다."""
    lines = [f"강점 응답 수: {stats['strength']['n']}건", f"개선 응답 수: {stats['weakness']['n']}건"]
    lines.append("강점 Top: " + ", ".join(f"{mid}({pct*100:.1f}%, 전사대비 {dev:+.1f}%p)"
                 for mid, pct in stats["strength_top"]
                 for dev in [stats["strength"]["deviation_pp"].get(mid, 0.0)]))
    lines.append("개선 Top: " + ", ".join(f"{mid}({pct*100:.1f}%, 전사대비 {dev:+.1f}%p)"
                 for mid, pct in stats["weakness_top"]
                 for dev in [stats["weakness"]["deviation_pp"].get(mid, 0.0)]))
    if stats["ambivalent"]:
        lines.append("양가 이슈: " + ", ".join(stats["ambivalent"]))
    lines.append(f"반어·냉소 비율: {stats['ironic_rate']*100:.1f}%")
    leader = stats.get("leader")
    if leader and not leader["low_sample"]:
        lines.append(f"부서장 의견 수: {leader['n']}건, 즉시확인 {leader['immediate_n']}건, "
                      f"점검대상 {leader['check_target_n']}건")
        lines.append("심각도 분포: " + ", ".join(f"{k}단계 {v}건" for k, v in leader["severity_dist"].items()))
    if objective:
        area_scores = objective.get("area_scores") or {}
        area_yoy = objective.get("area_yoy") or {}
        if area_scores:
            lines.append("객관식 설문 영역별 점수: " + ", ".join(
                f"{k} {v:.1f}점" + (f"(전년비 {area_yoy[k]:+.1f})" if area_yoy.get(k) is not None else "")
                for k, v in area_scores.items()))
        if objective.get("response_rate") is not None:
            lines.append(f"객관식 설문 참여율: {objective['response_rate']*100:.1f}%")
    return "\n".join(lines)


# =============================================================
# 6. 차트 (matplotlib)
# =============================================================
def set_korean_font_for_matplotlib():
    """맑은 고딕을 우선 시도하고, 없는 환경(예: 이 저장소의 테스트/CI 환경)
    에서는 얻을 수 있는 아무 sans-serif로 조용히 대체합니다 — 실제 배포
    환경(Windows)에는 맑은 고딕이 있으므로 거기서는 정상 렌더링됩니다."""
    import matplotlib
    candidates = ["Malgun Gothic", "맑은 고딕", "NanumGothic", "AppleGothic", "Noto Sans CJK KR"]
    available = {f.name for f in matplotlib.font_manager.fontManager.ttflist}
    for name in candidates:
        if name in available:
            matplotlib.rcParams["font.family"] = name
            return
    matplotlib.rcParams["font.family"] = "sans-serif"
    matplotlib.rcParams["axes.unicode_minus"] = False


def make_top_k_bar_chart(items, deviation_pp, out_path, color_hex):
    """items: [(label, pct), ...] (이미 Top-K로 잘라서 전달). 가로 막대,
    막대 옆에 비율과 전사 대비 편차(%p)를 표기합니다."""
    import matplotlib.pyplot as plt
    set_korean_font_for_matplotlib()

    labels = [mid for mid, _ in items][::-1]
    values = [pct * 100 for _, pct in items][::-1]
    devs = [deviation_pp.get(mid, 0.0) for mid, _ in items][::-1]

    fig, ax = plt.subplots(figsize=(6.2, 0.55 * len(items) + 0.6))
    bars = ax.barh(labels, values, color=f"#{color_hex}")
    ax.set_xlabel("응답 비율 (%)")
    ax.set_xlim(0, max(values + [1]) * 1.35)
    for bar, val, dev in zip(bars, values, devs):
        ax.text(bar.get_width() + max(values) * 0.02, bar.get_y() + bar.get_height() / 2,
                 f"{val:.1f}% (전사대비 {dev:+.1f}%p)", va="center", fontsize=9)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    return out_path


# =============================================================
# 7. Word 보고서 (python-docx)
# =============================================================
def _set_run_korean_font(run, name=FONT_KR):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts")
    if rfonts is None:
        from docx.oxml.ns import qn
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    from docx.oxml.ns import qn
    rfonts.set(qn("w:eastAsia"), name)


def add_kr_paragraph(doc, text="", size=None, bold=False, color_hex=None, align=None, italic=False):
    from docx.shared import Pt, RGBColor
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    run = p.add_run(text)
    _set_run_korean_font(run)
    if size:
        run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color_hex:
        run.font.color.rgb = RGBColor.from_string(color_hex)
    return p


def add_kr_heading(doc, text, level=1):
    from docx.shared import Pt
    p = doc.add_paragraph()
    run = p.add_run(text)
    _set_run_korean_font(run)
    run.font.bold = True
    run.font.size = Pt(16 if level == 1 else 13)
    run.font.color.rgb = None
    return p


def set_cell_background(cell, color_hex):
    from docx.oxml.ns import qn
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.makeelement(qn("w:shd"), {qn("w:val"): "clear", qn("w:color"): "auto", qn("w:fill"): color_hex})
    tc_pr.append(shd)


def set_page_layout(doc):
    from docx.shared import Cm
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)


def safe_filename(name):
    return re.sub(r'[\\/:*?"<>|]', "_", str(name))


def build_org_report(org, stats, tmp_dir, out_path, config=CONFIG):
    """조직 하나의 임원보고용 .docx를 생성합니다."""
    import docx
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = docx.Document()
    set_page_layout(doc)

    add_kr_heading(doc, str(org), level=1)
    add_kr_paragraph(doc, "주관식 응답 분석 · 임원보고용", size=11, color_hex="808080")

    # --- KPI 표 ---
    kpi_labels = ["잘하는점 응답수", "노력할점 응답수", "부서장 의견수", "즉시확인 건수", "반어·냉소 비율"]
    leader = stats.get("leader") or {"n": 0, "immediate_n": 0}
    kpi_values = [
        str(stats["strength"]["n"]),
        str(stats["weakness"]["n"]),
        str(leader.get("n", 0)),
        str(leader.get("immediate_n", 0)),
        f"{stats['ironic_rate'] * 100:.1f}%",
    ]
    table = doc.add_table(rows=2, cols=len(kpi_labels))
    table.style = "Table Grid"
    for i, label in enumerate(kpi_labels):
        cell = table.cell(0, i)
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(label)
        _set_run_korean_font(run)
        run.font.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = docx.shared.RGBColor.from_string("FFFFFF")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_background(cell, COLOR_NAVY)
    for i, val in enumerate(kpi_values):
        cell = table.cell(1, i)
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(val)
        _set_run_korean_font(run)
        run.font.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = docx.shared.RGBColor.from_string(COLOR_GOLD)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_background(cell, COLOR_NAVY)

    # --- 종합 분석 (주관식+객관식 통합 5문장 서술형, LLM 결과 없으면 섹션 생략) ---
    if stats.get("narrative_summary"):
        add_kr_heading(doc, "종합 분석", level=2)
        add_kr_paragraph(doc, stats["narrative_summary"].strip(), size=11, italic=True)

    # --- 총평 (LLM 결과 없으면 섹션 생략) ---
    if stats.get("overview_summary"):
        add_kr_heading(doc, "총평", level=2)
        for line in stats["overview_summary"].splitlines():
            if line.strip():
                add_kr_paragraph(doc, line.strip(), size=10.5)

    # --- 강점 Top5 ---
    add_kr_heading(doc, f"잘하는 점 Top{config['TOP_K']}", level=2)
    if stats["strength"]["low_sample"]:
        add_kr_paragraph(doc, "표본 부족(익명성 기준)", size=10.5, color_hex="999999")
    elif not stats["strength_top"]:
        add_kr_paragraph(doc, "집계된 항목이 없습니다.", size=10.5, color_hex="999999")
    else:
        chart_path = Path(tmp_dir) / f"{safe_filename(org)}_strength.png"
        make_top_k_bar_chart(stats["strength_top"], stats["strength"]["deviation_pp"], chart_path, COLOR_NAVY)
        doc.add_picture(str(chart_path), width=Cm(15))
        for mid, _ in stats["strength_top"]:
            theme = stats["theme_summaries"].get(("strength", mid), "")
            if theme:
                add_kr_paragraph(doc, f"[{mid}] {theme}", size=9, color_hex="555555")

    # --- 개선 Top5 ---
    add_kr_heading(doc, f"노력해야 할 점 Top{config['TOP_K']}", level=2)
    if stats["weakness"]["low_sample"]:
        add_kr_paragraph(doc, "표본 부족(익명성 기준)", size=10.5, color_hex="999999")
    elif not stats["weakness_top"]:
        add_kr_paragraph(doc, "집계된 항목이 없습니다.", size=10.5, color_hex="999999")
    else:
        chart_path = Path(tmp_dir) / f"{safe_filename(org)}_weakness.png"
        make_top_k_bar_chart(stats["weakness_top"], stats["weakness"]["deviation_pp"], chart_path, COLOR_CORAL)
        doc.add_picture(str(chart_path), width=Cm(15))
        for mid, _ in stats["weakness_top"]:
            theme = stats["theme_summaries"].get(("weakness", mid), "")
            if theme:
                add_kr_paragraph(doc, f"[{mid}] {theme}", size=9, color_hex="555555")

    # --- 양가 이슈 ---
    add_kr_heading(doc, "양가 이슈", level=2)
    if stats["ambivalent"]:
        for mid in stats["ambivalent"]:
            add_kr_paragraph(doc, f"· {mid} — 강점과 개선 양쪽 상위권에 동시 언급", size=10.5)
    else:
        add_kr_paragraph(doc, "해당 없음", size=10.5, color_hex="999999")

    # --- 부서장 의견 ---
    add_kr_heading(doc, "부서장에게 하고 싶은 말", level=2)
    if leader.get("low_sample", True):
        add_kr_paragraph(doc, "표본 부족(익명성 기준)", size=10.5, color_hex="999999")
    else:
        sev_table = doc.add_table(rows=2, cols=4)
        sev_table.style = "Table Grid"
        for i, lvl in enumerate((0, 1, 2, 3)):
            c = sev_table.cell(0, i)
            r = c.paragraphs[0].add_run(f"심각도 {lvl}")
            _set_run_korean_font(r)
            r.font.bold = True
            r.font.size = Pt(9)
        for i, lvl in enumerate((0, 1, 2, 3)):
            c = sev_table.cell(1, i)
            r = c.paragraphs[0].add_run(str(leader["severity_dist"].get(lvl, 0)))
            _set_run_korean_font(r)
            r.font.size = Pt(10)
        if leader["top_types"]:
            add_kr_paragraph(doc, "주요 유형: " + ", ".join(f"{t}({c}건)" for t, c in leader["top_types"]), size=10)
        if stats.get("leader_summary"):
            for line in stats["leader_summary"].splitlines():
                if line.strip():
                    add_kr_paragraph(doc, line.strip(), size=10)

    # --- 각주 ---
    doc.add_paragraph()
    add_kr_paragraph(doc, "※ 비율은 유효 응답 대비 언급 비율이며, 한 응답이 여러 카테고리로 "
                          "중복 집계될 수 있습니다.", size=8, color_hex="888888")
    add_kr_paragraph(doc, "※ AI 요약은 초안이므로 원문 확인이 필요합니다.", size=8, color_hex="888888")
    add_kr_paragraph(doc, "※ 개별 원문과 즉시확인 건은 본 보고서에 포함하지 않았습니다.",
                      size=8, color_hex="888888")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


def build_company_report(all_stats, out_path):
    """전사_종합.docx — 조직별 KPI 비교표 + 조직별 총평 요약."""
    import docx
    from docx.shared import Pt

    doc = docx.Document()
    set_page_layout(doc)
    add_kr_heading(doc, "전사 종합", level=1)
    add_kr_paragraph(doc, "주관식 응답 분석 · 임원보고용", size=11, color_hex="808080")

    headers = ["조직", "응답수", "즉시확인", "반어·냉소 비율", "강점 1위", "개선 1위"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        r = cell.paragraphs[0].add_run(h)
        _set_run_korean_font(r)
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = docx.shared.RGBColor.from_string("FFFFFF")
        set_cell_background(cell, COLOR_NAVY)

    for org, stats in all_stats.items():
        leader = stats.get("leader") or {"n": 0, "immediate_n": 0}
        total_n = stats["strength"]["n"] + stats["weakness"]["n"] + leader.get("n", 0)
        top1_strength = stats["strength_top"][0][0] if stats["strength_top"] else "-"
        top1_weakness = stats["weakness_top"][0][0] if stats["weakness_top"] else "-"
        row = table.add_row()
        values = [str(org), str(total_n), str(leader.get("immediate_n", 0)),
                  f"{stats['ironic_rate'] * 100:.1f}%", top1_strength, top1_weakness]
        for i, v in enumerate(values):
            r = row.cells[i].paragraphs[0].add_run(v)
            _set_run_korean_font(r)
            r.font.size = Pt(10)

    doc.add_paragraph()
    add_kr_heading(doc, "조직별 종합 분석", level=2)
    for org, stats in all_stats.items():
        add_kr_paragraph(doc, str(org), size=11, bold=True, color_hex=COLOR_NAVY)
        narrative = stats.get("narrative_summary") or "(종합 분석 없음 — LLM 미사용 또는 생성 실패)"
        add_kr_paragraph(doc, narrative.strip(), size=10, italic=True)

    doc.add_paragraph()
    add_kr_heading(doc, "조직별 총평 요약", level=2)
    for org, stats in all_stats.items():
        add_kr_paragraph(doc, str(org), size=11, bold=True, color_hex=COLOR_NAVY)
        summary = stats.get("overview_summary") or "(총평 없음 — LLM 미사용 또는 생성 실패)"
        for line in summary.splitlines():
            if line.strip():
                add_kr_paragraph(doc, line.strip(), size=10)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


# =============================================================
# 7. 대시보드 통합 JSON — 조직마다 .html을 따로 만드는 대신, dept_code로
#    키를 맞춘 JSON 하나만 만들어서 dashboard/index.html의 "주관식분석" 탭이
#    직접 불러와 객관식 지표(이미 대시보드에 로드돼 있음)와 한 화면에서
#    보여줄 수 있게 합니다. 파일 수를 줄이고, dept_code 기준 매칭이라
#    이름 표기 차이로 인한 연동 실패도 없습니다.
# =============================================================
def write_subjective_analysis_json(all_stats, out_path):
    """dept_code로 매칭된 조직만 담아 {dept_code: {...}} JSON을 씁니다.
    dept_code가 없는 조직(진단부서 표기가 survey_data.json의 dept_name과
    안 맞는 경우)은 build_all_org_stats()에서 이미 경고를 출력했으므로
    여기서는 조용히 건너뜁니다 — 그런 조직도 .docx 리포트에는 여전히 포함됩니다."""
    payload = {}
    skipped = []
    for org, stats in all_stats.items():
        dept_code = stats.get("dept_code")
        if not dept_code:
            skipped.append(org)
            continue
        payload[dept_code] = {
            "org_name": org,
            "narrative_summary": stats.get("narrative_summary") or "",
            "overview_summary": stats.get("overview_summary") or "",
            "strength_top": stats.get("strength_top") or [],
            "weakness_top": stats.get("weakness_top") or [],
            "ambivalent": stats.get("ambivalent") or [],
            "ironic_rate": stats.get("ironic_rate", 0.0),
            "leader": stats.get("leader"),
            "leader_summary": stats.get("leader_summary") or "",
            "theme_summaries": {f"{kind}::{mid}": text for (kind, mid), text in stats.get("theme_summaries", {}).items()},
        }
    if skipped:
        print(f"  warning: dept_code 매칭 실패로 대시보드 연동 JSON에서 제외된 조직: {skipped}", file=sys.stderr)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=str)
    return out_path


# =============================================================
# 8. 기타 산출물
# =============================================================
def write_results_json(all_stats, out_path):
    def _clean(obj):
        if isinstance(obj, dict):
            return {str(k): _clean(v) for k, v in obj.items()}
        if isinstance(obj, (list, tuple)):
            return [_clean(v) for v in obj]
        return obj

    payload = {str(org): _clean({k: v for k, v in stats.items() if k != "dup_keys_by_mid"})
               for org, stats in all_stats.items()}
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=str)


def write_hr_only_immediate_csv(leader_df, out_path):
    """즉시확인=O 행만 별도 CSV로 — 원문은 이 파일에만 남기고 보고서(.docx)에는
    절대 넣지 않습니다."""
    mask = leader_df["즉시확인"].astype(str).str.strip().str.upper() == "O"
    cols = ["사업부", "실", "팀", "진단부서", "심각도", "즉시확인점수", "원문"]
    sub = leader_df.loc[mask, cols]
    out_path.parent.mkdir(parents=True, exist_ok=True)
    sub.to_csv(out_path, index=False, encoding="utf-8-sig")


# =============================================================
# 9. 조직별 통계 구성 (LLM 포함) — main()에서 호출
# =============================================================
def build_all_org_stats(strength_df, weakness_df, leader_df, config=CONFIG):
    org_level = config["ORG_LEVEL"]
    prob_min = config["PROB_MIN"]
    min_n = config["MIN_N"]
    min_cat_n = config["MIN_CAT_N"]
    top_k = config["TOP_K"]
    top_k_amb = config["TOP_K_AMBIVALENT"]
    objective_records = load_objective_records(config.get("OBJECTIVE_DATA_PATH"))
    dept_code_by_name = load_dept_code_lookup()

    company_strength_baseline = compute_company_baseline(strength_df, org_level, prob_min)
    company_weakness_baseline = compute_company_baseline(weakness_df, org_level, prob_min)

    strength_org = aggregate_org_strength_weakness(strength_df, org_level, prob_min, min_n, company_strength_baseline)
    weakness_org = aggregate_org_strength_weakness(weakness_df, org_level, prob_min, min_n, company_weakness_baseline)
    ironic_rate = compute_ironic_rate(weakness_df, org_level, prob_min)
    leader_org = aggregate_leader_stats(leader_df, org_level, min_n) if len(leader_df) else {}

    strength_long = explode_categories(strength_df, org_level, prob_min, "strength")
    weakness_long = explode_categories(weakness_df, org_level, prob_min, "weakness")

    cache = load_llm_cache(config["OUT_DIR"])
    all_orgs = sorted(set(strength_org) | set(weakness_org) | set(leader_org))
    all_stats = {}

    for org in all_orgs:
        s_stats = strength_org.get(org, {"n": 0, "counts": {}, "pct": {}, "deviation_pp": {}, "low_sample": True})
        w_stats = weakness_org.get(org, {"n": 0, "counts": {}, "pct": {}, "deviation_pp": {}, "low_sample": True})
        l_stats = leader_org.get(org)

        strength_top = [] if s_stats["low_sample"] else top_categories(s_stats["pct"], top_k)
        weakness_top = [] if w_stats["low_sample"] else top_categories(w_stats["pct"], top_k)
        ambivalent = [] if (s_stats["low_sample"] or w_stats["low_sample"]) else compute_ambivalent_issues(
            s_stats["pct"], w_stats["pct"], top_k_amb)

        theme_summaries = {}
        for mid, _ in strength_top:
            if s_stats["counts"].get(mid, 0) < min_cat_n:
                continue
            texts = sample_texts_for_theme(strength_long, org, mid, config["SAMPLE_N"])
            theme_summaries[("strength", mid)] = summarize_category_theme(org, mid, texts, cache, config)
        for mid, _ in weakness_top:
            if w_stats["counts"].get(mid, 0) < min_cat_n:
                continue
            texts = sample_texts_for_theme(weakness_long, org, mid, config["SAMPLE_N"])
            theme_summaries[("weakness", mid)] = summarize_category_theme(org, mid, texts, cache, config)

        leader_summary = ""
        if l_stats and not l_stats["low_sample"]:
            g = leader_df[leader_df[org_level] == org]
            sev = pd.to_numeric(g["심각도"], errors="coerce").fillna(0).astype(int)
            eligible = g[(sev.isin([1, 2])) & (g["즉시확인"].astype(str).str.strip().str.upper() != "O")]
            leader_summary = summarize_leader_opinions(org, eligible["원문"].dropna().tolist(), config, cache)

        dept_code = dept_code_by_name.get(org)
        if dept_code is None:
            print(f"  warning: '{org}'에 매칭되는 dept_code를 찾지 못했습니다 "
                  f"(data/survey_data.json의 dept_name과 표기가 다를 수 있음) — "
                  f"객관식 지표 연동 및 대시보드 통합 없이 진행합니다.", file=sys.stderr)
        objective = objective_records.get(dept_code) if dept_code else None
        stats = {
            "strength": s_stats, "weakness": w_stats, "leader": l_stats,
            "strength_top": strength_top, "weakness_top": weakness_top,
            "ambivalent": ambivalent, "ironic_rate": ironic_rate.get(org, 0.0),
            "theme_summaries": theme_summaries, "leader_summary": leader_summary,
            "objective": objective, "dept_code": dept_code,
        }
        facts_text = build_overview_facts_text(org, stats, objective)
        stats["overview_summary"] = summarize_org_overview(org, facts_text, config, cache)
        stats["narrative_summary"] = summarize_integrated_narrative(org, facts_text, config, cache)
        all_stats[org] = stats

    save_llm_cache(config["OUT_DIR"], cache)
    return all_stats


# =============================================================
# 10. main
# =============================================================
def main(config=CONFIG):
    out_dir = Path(config["OUT_DIR"])
    tmp_dir = out_dir / "_tmp"
    out_dir.mkdir(parents=True, exist_ok=True)
    tmp_dir.mkdir(parents=True, exist_ok=True)

    strength_path = resolve_excel_path(config["EXCEL_STRENGTH"], FILE_KEYWORDS["strength"], config["EXCEL_DIR"])
    weakness_path = resolve_excel_path(config["EXCEL_WEAKNESS"], FILE_KEYWORDS["weakness"], config["EXCEL_DIR"])
    leader_path = resolve_excel_path(config["EXCEL_LEADER"], FILE_KEYWORDS["leader"], config["EXCEL_DIR"])
    print(f"엑셀 로딩 중... (강점: {strength_path.name}, 개선: {weakness_path.name}, "
          f"부서장: {leader_path.name})")
    strength_df = load_strength_weakness(strength_path, config["SHEET"])
    weakness_df = load_weakness(weakness_path, config["SHEET"])
    leader_df = load_leader_comments(leader_path, config["SHEET"])
    print(f"  잘하는점 {len(strength_df)}건, 노력해야할점 {len(weakness_df)}건, "
          f"부서장의견 {len(leader_df)}건 (유효 응답)")

    if not config["BASE_URL"]:
        print("BASE_URL이 비어있어 LLM 요약을 건너뜁니다 (집계·보고서는 정상 생성).")

    print("집계 및 요약 생성 중...")
    all_stats = build_all_org_stats(strength_df, weakness_df, leader_df, config)

    print(f"{len(all_stats)}개 조직 보고서 생성 중...")
    for org, stats in all_stats.items():
        out_path = out_dir / f"{safe_filename(org)}.docx"
        build_org_report(org, stats, tmp_dir, out_path, config)
        print(f"  {out_path}")

    company_path = out_dir / "전사_종합.docx"
    build_company_report(all_stats, company_path)
    print(f"  {company_path}")

    write_results_json(all_stats, out_dir / "results.json")
    if len(leader_df):
        write_hr_only_immediate_csv(leader_df, out_dir / "HR_only_즉시확인.csv")

    dashboard_json_path = Path(config.get("DASHBOARD_JSON_PATH") or (DATA_DIR / "subjective_analysis.json"))
    write_subjective_analysis_json(all_stats, dashboard_json_path)
    print(f"  {dashboard_json_path}  <- 대시보드 '주관식분석' 탭에서 이 파일을 불러오세요")

    shutil.rmtree(tmp_dir, ignore_errors=True)
    print("완료.")


if __name__ == "__main__":
    main()
