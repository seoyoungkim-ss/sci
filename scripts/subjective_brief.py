"""사내 진단 "부서장에게 하고 싶은 말" 응답을 분석해 우리 사업부 토론
회의용 Word(.docx) 브리프를 만드는 단일 스크립트.

전사 데이터가 없는 전제이므로 비교 기준은 항상 "사업부 내 평균"과
"하위 조직(실/팀) 간 비교"뿐입니다 — 전사 비교는 어디서도 하지 않습니다.

입력 (엑셀 A~U, 21컬럼, DRM 보호 — 반드시 xlwings로 엶):
  A:id B:연도 C:진단부서코드 D:진단부서명 E:사업부 F:실 G:팀 H:원문 I:글자수
  J:동일응답수 K:응답상태 L:심각도 M:심각도명 N:점검대상 O:점검대상점수
  P:즉시확인 Q:즉시확인점수 R:코드 S:유형 T:대분류 U:코드확률
헤더명이 이 이름 그대로 있으면 헤더명으로 매핑하고, 없거나 순서가 다르면
위 A~U 고정 위치로 매핑합니다(build_column_mapping 참고).

처리: xlwings로 읽기 -> pandas로 집계(사업부 전체/유형별/실별/품질/익명성)
-> 로컬 vLLM(OpenAI 호환, 모델 "thinkingcap")으로 원문 구조화 추출 ->
이슈 정규화 -> 이슈 집계·요약 -> 브리프용 핵심 3줄/토론 안건 생성 ->
python-docx + matplotlib로 1페이지 브리프(+부록) .docx 생성.

실행:
    python scripts/subjective_brief.py

경로/모델 등은 전부 아래 CONFIG에서 조정합니다. BASE_URL을 비워두면 LLM
사용 섹션(원문 구조화 추출 이후 전부)만 생략되고, pandas 집계 기반 섹션은
정상적으로 생성됩니다.

테스트: xlwings 로딩 없이 전체 파이프라인을 검증하려면 main()에 넘기는
config와 함께 load_raw_dataframe()을 원하는 DataFrame을 돌려주는 함수로
바꿔치기하면 됩니다(아래 함수는 모듈 전역 이름으로 참조되므로 monkeypatch가
그대로 적용됩니다):

    import subjective_brief as sb
    sb.load_raw_dataframe = lambda path, sheet=None: my_fake_df
    sb.main({**sb.CONFIG, "OUT_DIR": "test_output", "BASE_URL": ""})

이렇게 하면 집계/LLM 캐시/차트/.docx 생성 로직은 실제 실행과 완전히
동일하게 동작합니다. LLM까지 포함해 전체 파이프라인을 끝까지 돌려보려면
sb.call_llm도 가짜 응답을 돌려주는 함수로 바꿔치면 됩니다.
"""
from __future__ import annotations

import hashlib
import json
import random
import re
import shutil
import sys
from collections import Counter
from pathlib import Path

import pandas as pd

# =============================================================
# 0. CONFIG — 여기만 고치면 됩니다
# =============================================================
CONFIG = {
    "BASE_URL": "",          # 사내 vLLM OpenAI 호환 엔드포인트. 비워두면 LLM 사용 섹션 전부 건너뜀.
    "MODEL": "thinkingcap",
    "API_KEY": "EMPTY",

    "FILE_PATH": None,   # 엑셀 경로 (xlwings로 읽음) — 실행 전 반드시 지정
    "SHEET": None,       # None이면 첫 번째 시트

    "DIV_NAME": "",      # 우리 사업부명. 비워두면 파일의 "사업부" 열 최빈값을 사용
    "UNIT_LEVEL": "실",  # "실" | "팀" — 하위 조직 비교 기준
    "MIN_N": 5,          # 이 미만 응답의 조직은 "기타(소규모 조직 합산)"로 묶음(익명성)
    "SAMPLE_N": 15,      # 이슈별 요약에 쓰는 대표 응답 샘플 개수

    "OUT_DIR": "output",

    "LLM_TIMEOUT_SEC": 300,
    "LLM_TEMPERATURE": 0.2,
    "LLM_MAX_TOKENS": 8192,  # thinkingcap은 사고 과정(<think>...</think>)이 길어서 생각만 하다 끝나는
                             # 경우가 있어(토큰 부족 -> 실제 JSON 답변을 못 냄 -> 전부 "미분류") 넉넉히 잡음.
                             # 그래도 미분류 비율이 높게 나오면 더 올려보세요.
    "RANDOM_SEED": 42,
}

# 색상 (Word 보고서 공통)
COLOR_NAVY = "0B1F3A"
COLOR_CORAL = "FF6B57"
COLOR_GOLD = "D4A24C"
FONT_KR = "맑은 고딕"

NO_DATA_TEXT_VARIANTS = {"-", "–", "—", "N/A", "NA", "없음"}


# =============================================================
# 1. 컬럼 매핑 + 엑셀 로딩
# =============================================================
# 열 위치(A~U) 기준 기대 순서 — 실제 열 순서가 다르면 이 리스트만 고치면 됩니다.
EXPECTED_COLUMNS = [
    "id", "연도", "진단부서코드", "진단부서명", "사업부", "실", "팀", "원문", "글자수",
    "동일응답수", "응답상태", "심각도", "심각도명", "점검대상", "점검대상점수",
    "즉시확인", "즉시확인점수", "코드", "유형", "대분류", "코드확률",
]
# 헤더명으로 찾을 때 허용하는 표기(대소문자/공백 변형 없이 정확히 일치하는 경우만 —
# 실제 헤더 표기가 이와 다르면 폴백으로 위치 매핑이 쓰이므로 안전함).
HEADER_NAME_ALIASES = {
    "ID": "id", "id": "id",
}


def build_column_mapping(header_row):
    """1행(헤더로 추정되는 행)이 EXPECTED_COLUMNS의 이름들과 얼마나 맞는지
    보고, 70% 이상 맞으면 "헤더명 -> 실제 위치"로 매핑하고, 아니면 고정
    위치(A~U 순서 그대로)로 폴백합니다. 반환값은 EXPECTED_COLUMNS와 같은
    길이의 리스트로, i번째 값이 "EXPECTED_COLUMNS[i]가 들어있는 실제 열
    인덱스(0-based)"입니다."""
    name_to_pos = {}
    for idx, cell in enumerate(header_row or []):
        if not isinstance(cell, str):
            continue
        key = cell.strip()
        key = HEADER_NAME_ALIASES.get(key, key)
        if key in EXPECTED_COLUMNS and key not in name_to_pos:
            name_to_pos[key] = idx

    if len(name_to_pos) >= len(EXPECTED_COLUMNS) * 0.7:
        return [name_to_pos.get(name, i) for i, name in enumerate(EXPECTED_COLUMNS)]
    return list(range(len(EXPECTED_COLUMNS)))


def open_workbook_readonly(path):
    """xlwings.App(visible=False)로 읽기 전용으로 엽니다. 호출부에서
    반드시 finally에서 app.quit()하세요."""
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
    """워크북을 열어 지정 시트(없으면 첫 시트)의 used_range를 2차원
    리스트(raw values, 헤더 포함)로 반환하고 앱을 종료합니다."""
    app, book = open_workbook_readonly(path)
    try:
        sheet = book.sheets[sheet_name] if sheet_name else book.sheets[0]
        values = sheet.used_range.value
        if values is None:
            return []
        if not isinstance(values[0], list):  # 행이 1개뿐이면 xlwings가 1차원으로 줌
            values = [values]
        return values
    finally:
        close_workbook_app(app, book)


def load_raw_dataframe(path, sheet=None):
    """엑셀을 읽어 EXPECTED_COLUMNS 이름이 부여된 DataFrame으로 반환합니다
    (응답상태 필터링 전의 raw 상태 — 필터링은 filter_valid_responses에서)."""
    values = read_sheet_raw_values(path, sheet)
    if not values:
        return pd.DataFrame(columns=EXPECTED_COLUMNS)
    mapping = build_column_mapping(values[0])
    body = values[1:]
    n = len(EXPECTED_COLUMNS)
    rows = []
    for row in body:
        if not row:
            continue
        rows.append([row[pos] if pos < len(row) else None for pos in mapping])
    df = pd.DataFrame(rows, columns=EXPECTED_COLUMNS[:n])
    return df


def to_number(v):
    """None/빈값/"-" 류/수식 에러는 None, 그 외엔 float로. 퍼센트 기호·콤마·
    공백을 허용합니다(이 저장소의 다른 로더들과 동일한 관용 규칙)."""
    import math
    if v is None or v == "":
        return None
    if isinstance(v, str):
        s = v.strip()
        if s == "" or s in NO_DATA_TEXT_VARIANTS or s.startswith("#"):
            return None
        s = s.replace(",", "").rstrip("%").strip()
        try:
            v = float(s)
        except ValueError:
            return None
    else:
        try:
            v = float(v)
        except (TypeError, ValueError):
            return None
    return None if math.isnan(v) else v


def filter_valid_responses(df):
    """응답상태가 "10"으로 시작하는 행만 남기고, 심각도를 숫자로 변환합니다."""
    df = df[df["응답상태"].astype(str).str.startswith("10", na=False)].copy()
    df["심각도"] = pd.to_numeric(df["심각도"], errors="coerce").fillna(0).astype(int)
    df["점검대상점수"] = df["점검대상점수"].apply(to_number)
    df["즉시확인점수"] = df["즉시확인점수"].apply(to_number)
    df["코드확률"] = df["코드확률"].apply(lambda v: v if v is None else str(v))  # explode에서 다시 파싱
    df["글자수"] = df["글자수"].apply(to_number)
    df["동일응답수"] = df["동일응답수"].apply(to_number)
    df["즉시확인"] = df["즉시확인"].fillna("").astype(str).str.strip().str.upper()
    df["점검대상"] = df["점검대상"].fillna("").astype(str).str.strip().str.upper()
    return df.reset_index(drop=True)


def split_cell_list(cell):
    """셀 값을 ","로 분리한 문자열 리스트로. 빈 값/NaN이면 빈 리스트."""
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return []
    s = str(cell).strip()
    if s == "" or s.lower() == "nan":
        return []
    return [v.strip() for v in s.split(",") if v.strip() != ""]


def parse_prob_list(cell, n_expected):
    """코드확률 셀을 n_expected개의 float 리스트로. 값이 1을 넘으면 %로
    보고 100으로 나눔. 개수가 안 맞으면 첫 값을 전체에 사용."""
    if n_expected <= 0:
        return []
    raw_vals = []
    for p in split_cell_list(cell):
        try:
            raw_vals.append(float(p))
        except ValueError:
            raw_vals.append(None)
    raw_vals = [(v / 100.0 if (v is not None and v > 1) else v) for v in raw_vals]
    if len(raw_vals) == n_expected:
        return raw_vals
    first = raw_vals[0] if raw_vals else None
    return [first] * n_expected


def explode_type_major_code(df):
    """유형/대분류/코드가 콤마로 여러 개일 수 있는 행을, 응답 1건 x 항목
    1개로 펼친 long-format DataFrame으로 변환합니다. 컬럼: id, 실, 팀,
    심각도, 점검대상점수, 유형, 대분류, 코드, 코드확률. 세 컬럼(유형/대분류/
    코드) 중 길이가 다르면 가장 긴 것 기준으로 맞추고 모자란 자리는 None."""
    records = []
    for _, row in df.iterrows():
        types = split_cell_list(row.get("유형"))
        majors = split_cell_list(row.get("대분류"))
        codes = split_cell_list(row.get("코드"))
        n = max(len(types), len(majors), len(codes), 1)
        probs = parse_prob_list(row.get("코드확률"), n)
        types = types + [None] * (n - len(types))
        majors = majors + [None] * (n - len(majors))
        codes = codes + [None] * (n - len(codes))
        for i in range(n):
            if types[i] is None and majors[i] is None and codes[i] is None:
                continue
            records.append({
                "id": row.get("id"), "실": row.get("실"), "팀": row.get("팀"),
                "심각도": row.get("심각도"), "점검대상점수": row.get("점검대상점수"),
                "유형": types[i], "대분류": majors[i], "코드": codes[i],
                "코드확률": probs[i] if i < len(probs) else None,
            })
    return pd.DataFrame(records)


# =============================================================
# 2. 집계 (pandas — LLM 사용 안 함)
# =============================================================
def compute_division_summary(df):
    """사업부 전체 KPI: 유효 응답 수, 심각도 0~3 건수·비율, 점검대상/즉시확인
    건수, 평균 점검대상점수."""
    n = len(df)
    sev_counts = df["심각도"].value_counts().to_dict()
    sev_counts = {lvl: int(sev_counts.get(lvl, 0)) for lvl in (0, 1, 2, 3)}
    sev_pct = {lvl: (sev_counts[lvl] / n if n else 0.0) for lvl in (0, 1, 2, 3)}
    check_target_n = int((df["점검대상"] == "O").sum())
    immediate_n = int((df["즉시확인"] == "O").sum())
    check_scores = df["점검대상점수"].dropna()
    return {
        "n_valid": n,
        "severity_counts": sev_counts,
        "severity_pct": sev_pct,
        "check_target_count": check_target_n,
        "immediate_count": immediate_n,
        "check_target_score_mean": float(check_scores.mean()) if len(check_scores) else None,
    }


def compute_type_summary(long_df, n_valid):
    """유형별: 건수, 비율(유효 응답 대비), 심각도≥2 비율, 평균 점검대상점수,
    평균 코드확률. 대분류별 분포도 함께."""
    by_type = {}
    for typ, g in long_df.groupby("유형", sort=False):
        if not typ:
            continue
        n = len(g)
        severe = (pd.to_numeric(g["심각도"], errors="coerce").fillna(0) >= 2).sum()
        check_scores = g["점검대상점수"].dropna()
        probs = g["코드확률"].dropna()
        by_type[typ] = {
            "n": int(n),
            "pct": n / n_valid if n_valid else 0.0,
            "severe_ge2_pct": float(severe / n) if n else 0.0,
            "check_target_score_mean": float(check_scores.mean()) if len(check_scores) else None,
            "code_prob_mean": float(probs.mean()) if len(probs) else None,
        }
    major_dist = Counter(m for m in long_df["대분류"].tolist() if m)
    return by_type, dict(major_dist)


def compute_unit_summary(df, long_df, unit_level, min_n, division_sev_pct):
    """UNIT_LEVEL(실/팀)별: 응답 수, 심각도 1/2/3 비율, 점검대상 비율,
    즉시확인 건수, 유형 Top3, 사업부 평균 대비 편차(%p). MIN_N 미만인
    조직은 "기타(소규모 조직 합산)"로 묶어서 개별 노출하지 않습니다
    (익명성 보호 — 요청 Top/요약/원문도 이 합산 그룹에선 만들지 않음)."""
    unit_counts = df[unit_level].value_counts()
    small_units = set(unit_counts[unit_counts < min_n].index)

    df = df.copy()
    df["_unit_group"] = df[unit_level].apply(
        lambda u: "기타(소규모 조직 합산)" if u in small_units or not u else u)
    long_df = long_df.copy()
    long_df["_unit_group"] = long_df[unit_level if unit_level in long_df.columns else "실"].apply(
        lambda u: "기타(소규모 조직 합산)" if u in small_units or not u else u)

    by_unit = {}
    for unit, g in df.groupby("_unit_group", sort=False):
        n = len(g)
        sev = pd.to_numeric(g["심각도"], errors="coerce").fillna(0)
        sev_pct = {lvl: float((sev == lvl).sum() / n) if n else 0.0 for lvl in (1, 2, 3)}
        check_pct = float((g["점검대상"] == "O").sum() / n) if n else 0.0
        immediate_n = int((g["즉시확인"] == "O").sum())
        deviation_pp = {lvl: (sev_pct[lvl] - division_sev_pct.get(lvl, 0.0)) * 100 for lvl in (1, 2, 3)}

        top_types = []
        if unit != "기타(소규모 조직 합산)":
            sub_long = long_df[long_df["_unit_group"] == unit]
            counts = sub_long["유형"].value_counts()
            top_types = [(t, int(c)) for t, c in counts.head(3).items() if t]

        by_unit[unit] = {
            "n": n, "severity_pct": sev_pct, "check_target_pct": check_pct,
            "immediate_count": immediate_n, "top_types": top_types,
            "deviation_pp": deviation_pp,
            "is_merged_small": unit == "기타(소규모 조직 합산)",
        }
    return by_unit


def build_heatmap_unit_severity(by_unit):
    """실(또는 팀) x 심각도(1/2/3) 비율 매트릭스 — docx 히트맵용."""
    units = [u for u in by_unit if not by_unit[u]["is_merged_small"]]
    return {u: by_unit[u]["severity_pct"] for u in units}


def build_heatmap_unit_type_top5(long_df, by_unit, unit_level):
    """실(또는 팀) x 유형 Top5 건수 매트릭스(부록용) — 소규모 합산 조직은 제외."""
    real_units = [u for u in by_unit if not by_unit[u]["is_merged_small"]]
    top5_types = [t for t, _ in Counter(
        t for t in long_df["유형"].tolist() if t).most_common(5)]
    matrix = {}
    for unit in real_units:
        sub = long_df[long_df[unit_level] == unit]
        counts = sub["유형"].value_counts()
        matrix[unit] = {t: int(counts.get(t, 0)) for t in top5_types}
    return matrix, top5_types


def compute_quality_metrics(long_df, df):
    """코드확률 <0.5 비율, 동일응답수가 큰(>=3) 응답 비중."""
    probs = long_df["코드확률"].dropna()
    low_prob_pct = float((probs < 0.5).sum() / len(probs)) if len(probs) else 0.0
    dup = df["동일응답수"].dropna()
    dup_heavy_pct = float((dup >= 3).sum() / len(dup)) if len(dup) else 0.0
    return {"low_prob_pct": low_prob_pct, "dup_heavy_pct": dup_heavy_pct}


def compute_year_trend(df):
    """연도 열에 2개 이상 값이 있으면 연도별 KPI 추이를, 1개뿐이면 None을
    반환합니다(그 경우 docx에서 이 섹션 자체를 생략)."""
    years = sorted({y for y in df["연도"].tolist() if y is not None and y != ""})
    if len(years) < 2:
        return None
    trend = {}
    for y in years:
        g = df[df["연도"] == y]
        n = len(g)
        sev = pd.to_numeric(g["심각도"], errors="coerce").fillna(0)
        trend[str(y)] = {
            "n": n,
            "severity_pct": {lvl: float((sev == lvl).sum() / n) if n else 0.0 for lvl in (0, 1, 2, 3)},
            "check_target_pct": float((g["점검대상"] == "O").sum() / n) if n else 0.0,
        }
    return trend


# =============================================================
# 3. LLM (요약·추출 전용 — 집계는 이미 끝난 데이터를 그대로 사용)
# =============================================================
LLM_SYSTEM_PROMPT = (
    "당신은 사내 조직문화 진단의 서술형 응답을 분석해 임원/그룹장 토론 자료를 돕는 "
    "HR 담당자의 어시스턴트입니다. 반드시 제공된 응답/수치에 있는 내용만 사용하고, "
    "추측하거나 새로운 사실을 만들지 마세요. 숫자는 제공된 값을 그대로 인용하고 스스로 "
    "계산하거나 새로 만들지 마세요. 개인, 특정 조직, 특정 직책이 식별될 수 있는 표현은 "
    "일반화해서 표현하세요. 응답 내용을 비난조로 해석하지 말고 '요청 사항', '개선 요구'의 "
    "관점으로 서술하세요. 한국어로, 개조식(명사형 종결)으로 간결하게 작성하세요."
)

THINK_TAG_RE = re.compile(r"<think>.*?</think>", re.DOTALL | re.IGNORECASE)


def strip_think_tags(text):
    return THINK_TAG_RE.sub("", text or "").strip()


def is_degenerate_llm_output(text):
    """같은 문자/토큰이 반복되는 깨진 출력(모델 서빙 문제 등으로 발생)인지
    휴리스틱으로 판별 — API 호출 자체는 성공해도 이런 경우가 있어 걸러냄."""
    stripped = text.strip()
    if len(stripped) < 5:
        return False
    if len(set(stripped)) / len(stripped) < 0.15:
        return True
    if re.search(r"(.)\1{9,}", stripped):
        return True
    return False


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


def call_llm(system, user, cache, config=CONFIG, guided_json_schema=None):
    """LLM 호출 — BASE_URL이 비어있거나 호출/생성에 실패하면 빈 문자열을
    반환합니다(호출부는 반드시 "빈 문자열 = 생략/미분류"로 처리). 프롬프트
    (system+user+model) 해시로 캐시해서 같은 입력에 대해 재실행 시 다시
    호출하지 않습니다. guided_json_schema가 주어지면 vLLM의 guided_json
    (extra_body)을 먼저 시도하고, 서버가 지원하지 않으면(예외 발생) 일반
    호출로 자동 폴백합니다."""
    if not config["BASE_URL"]:
        return ""

    key = prompt_hash(system, user + (json.dumps(guided_json_schema) if guided_json_schema else ""),
                       config["MODEL"])
    if key in cache:
        return cache[key]

    text = ""
    try:
        from openai import OpenAI
        # 사내망 내부 주소인데 시스템 프록시를 타면서 막히는 사례가 있었음
        # (교훈: subjective_report.py 작업 당시 겪음) — 여기서도 같은 문제를
        # 피하기 위해 httpx 클라이언트가 환경변수/시스템 프록시 설정을 아예
        # 안 읽도록 trust_env=False로 끔. (주의: httpx.Client(proxies=...)는
        # 최신 httpx에서 제거된 kwarg라 TypeError가 나므로 쓰지 않음 —
        # trust_env는 버전에 관계없이 안정적으로 지원됨.)
        try:
            import httpx
            http_client = httpx.Client(trust_env=False, timeout=config["LLM_TIMEOUT_SEC"])
        except Exception as e:
            print(f"  warning: httpx 클라이언트 생성 실패(프록시 우회 미적용) — {e}", file=sys.stderr)
            http_client = None
        client = OpenAI(base_url=config["BASE_URL"], api_key=config["API_KEY"],
                         timeout=config["LLM_TIMEOUT_SEC"], http_client=http_client)

        kwargs = dict(
            model=config["MODEL"],
            messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
            temperature=config["LLM_TEMPERATURE"],
            max_tokens=config["LLM_MAX_TOKENS"],
        )
        if guided_json_schema:
            try:
                resp = client.chat.completions.create(
                    extra_body={"guided_json": guided_json_schema}, **kwargs)
            except Exception:
                resp = client.chat.completions.create(**kwargs)  # 서버가 guided_json 미지원 -> 폴백
        else:
            resp = client.chat.completions.create(**kwargs)

        text = strip_think_tags(resp.choices[0].message.content)
        if is_degenerate_llm_output(text):
            print(f"  warning: LLM 응답이 비정상(반복/깨짐)으로 판단되어 무시합니다: {text[:60]!r}",
                  file=sys.stderr)
            text = ""
    except Exception as e:
        print(f"  warning: LLM 호출 실패 — {e}", file=sys.stderr)
        text = ""

    cache[key] = text
    return text


TITLE_SUFFIXES = ["팀장", "과장", "부장", "차장", "대리", "주임", "그룹장", "실장",
                   "본부장", "사원", "매니저", "리더", "수석", "책임"]
_NAME_TITLE_RE = re.compile(r"[가-힣]{2,4}(?:" + "|".join(TITLE_SUFFIXES) + r")")
_HONORIFIC_RE = re.compile(r"[가-힣]{2,4}님")


def mask_identifying_info(text):
    """이름+직책, "OOO님" 패턴을 일반화된 표현으로 바꿉니다. 정규식 기반의
    best-effort 처리로, 완벽한 개인정보 마스킹을 보장하지 않습니다 — 사람
    검토가 필요한 민감 정보가 있다면 별도로 확인하세요."""
    text = _NAME_TITLE_RE.sub("[직책자]", text)
    text = _HONORIFIC_RE.sub("[특정인]", text)
    return text


EXTRACTION_SCHEMA = {
    "type": "object",
    "properties": {
        "이슈": {"type": "string"},
        "성격": {"type": "string", "enum": ["요청", "불만", "제안", "칭찬"]},
        "요청내용": {"type": "string"},
        "행동영역": {"type": "string", "enum": ["업무지시", "소통", "평가", "리더십", "조직운영", "기타"]},
    },
    "required": ["이슈", "성격", "요청내용", "행동영역"],
}

UNCLASSIFIED = {"이슈": "미분류", "성격": "불만", "요청내용": "", "행동영역": "기타"}


def extract_structured_one(resp_id, text, cache, config=CONFIG):
    """응답 1건을 구조화된 JSON으로 추출합니다(1회 호출, 실패 시 1회
    재시도, 그래도 실패하면 "미분류"). 8자 미만 응답은 호출 전에 걸러서
    바로 미분류 처리(호출부에서 이미 걸렀다고 가정하지 않고 여기서도 확인).
    재시도는 1차와 다른 프롬프트 문자열을 보내서(캐시 키가 달라짐) 실제로
    다시 호출되도록 합니다 — 같은 문자열로 재시도하면 call_llm의 프롬프트
    해시 캐시에 걸려 "재시도"가 사실상 같은 실패 응답을 다시 읽는 것에
    불과해지는 문제가 있었음."""
    if not isinstance(text, str) or len(text.strip()) < 8:
        return dict(UNCLASSIFIED), None
    masked = mask_identifying_info(text.strip())
    base_user = (
        f"아래 응답을 분석해 JSON으로만 답하세요. 다른 설명 없이 JSON 객체 하나만 출력하세요.\n"
        f"스키마: {{\"이슈\": \"10자 내외 주제\", \"성격\": \"요청|불만|제안|칭찬\", "
        f"\"요청내용\": \"한 줄로 일반화한 문장\", \"행동영역\": \"업무지시|소통|평가|리더십|조직운영|기타\"}}\n"
        f"응답: {masked}"
    )
    last_raw = None
    for attempt in range(2):
        user = base_user if attempt == 0 else (
            base_user + "\n\n(다시 요청합니다: 설명이나 생각 과정 없이 JSON 객체 하나만 출력하세요.)")
        raw = call_llm(LLM_SYSTEM_PROMPT, user, cache, config, guided_json_schema=EXTRACTION_SCHEMA)
        last_raw = raw
        if not raw:
            break
        try:
            start, end = raw.index("{"), raw.rindex("}") + 1
            parsed = json.loads(raw[start:end])
            if all(k in parsed for k in EXTRACTION_SCHEMA["required"]):
                return parsed, None
        except (ValueError, json.JSONDecodeError):
            pass
    return dict(UNCLASSIFIED), last_raw


def extract_all(df, config=CONFIG, cache=None):
    """심각도 3/즉시확인=O 응답은 여기 들어오기 전에 반드시 제외되어 있어야
    합니다(main()에서 필터링). 순차 호출, 캐시 적용. 미분류 비율이 높으면
    (파싱 실패가 많다는 신호) 원인을 바로 알 수 있게 실패 샘플과 함께
    경고를 출력합니다 — 조용히 전부 "미분류"로만 끝나서 원인을 알 수
    없었던 문제를 막기 위함."""
    cache = cache if cache is not None else {}
    extracted = {}
    failures = []
    for _, row in df.iterrows():
        result, failed_raw = extract_structured_one(row["id"], row.get("원문"), cache, config)
        extracted[row["id"]] = result
        if failed_raw is not None:
            failures.append((row["id"], failed_raw))

    n = len(extracted)
    n_unclassified = sum(1 for v in extracted.values() if v["이슈"] == "미분류")
    if n and n_unclassified / n > 0.3:
        print(f"  warning: {n}건 중 {n_unclassified}건({n_unclassified/n*100:.0f}%)이 미분류로 "
              f"처리됐습니다 — LLM 응답 파싱이 많이 실패하고 있다는 신호입니다.", file=sys.stderr)
        for resp_id, raw in failures[:3]:
            snippet = raw[:200].replace("\n", " ") if raw else "(빈 응답)"
            print(f"    예시 [{resp_id}] 원시 응답: {snippet!r}", file=sys.stderr)
        print("    -> <think> 태그가 안 닫혀 있으면(생각만 하다 끝남) LLM_MAX_TOKENS를 "
              "더 늘려보세요. 빈 응답이 많으면 output/_llm_cache.json에서 실제 호출 "
              "실패 이유(에러 메시지)를 확인하세요.", file=sys.stderr)
    return extracted


ISSUE_MAPPING_PATH_NAME = "이슈_매핑.json"


def normalize_issues(extracted, out_dir, config=CONFIG, cache=None):
    """추출된 이슈 전체 목록을 LLM에 한 번 전달해 8~12개 표준 이슈로 병합한
    매핑표를 만듭니다. output/이슈_매핑.json이 이미 있으면(사람이 수정한
    파일) 그것을 그대로 우선 사용하고 LLM은 호출하지 않습니다."""
    mapping_path = Path(out_dir) / ISSUE_MAPPING_PATH_NAME
    if mapping_path.exists():
        with open(mapping_path, "r", encoding="utf-8") as f:
            mapping = json.load(f)
        print(f"  사람이 수정한 {mapping_path} 를 그대로 사용합니다(LLM 재호출 안 함).")
        return mapping

    raw_issues = sorted({v["이슈"] for v in extracted.values() if v["이슈"] and v["이슈"] != "미분류"})
    if not raw_issues or not config["BASE_URL"]:
        return {issue: issue for issue in raw_issues}

    cache = cache if cache is not None else {}
    user = (
        "아래는 서술형 응답에서 추출된 이슈 주제 목록입니다. 의미가 같거나 매우 비슷한 "
        "것끼리 묶어서 8~12개의 표준 이슈명으로 정리하고, 원래 이슈 각각이 어떤 표준 "
        "이슈에 해당하는지 매핑표를 JSON으로만 답하세요(다른 설명 없이).\n"
        "형식: {\"원래이슈1\": \"표준이슈A\", \"원래이슈2\": \"표준이슈A\", ...}\n"
        "원래 이슈 목록:\n" + "\n".join(f"- {i}" for i in raw_issues)
    )
    raw = call_llm(LLM_SYSTEM_PROMPT, user, cache, config)
    mapping = {}
    if raw:
        try:
            start, end = raw.index("{"), raw.rindex("}") + 1
            mapping = json.loads(raw[start:end])
        except (ValueError, json.JSONDecodeError):
            mapping = {}
    # LLM이 일부 이슈를 빠뜨렸으면 자기 자신으로 폴백(표준 이슈 목록에서
    # 누락되는 원본 이슈가 없도록).
    for issue in raw_issues:
        mapping.setdefault(issue, issue)

    mapping_path.parent.mkdir(parents=True, exist_ok=True)
    with open(mapping_path, "w", encoding="utf-8") as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)
    return mapping


def apply_issue_mapping(extracted, mapping):
    """추출 결과 각각에 표준 이슈를 부여합니다."""
    result = {}
    for resp_id, v in extracted.items():
        v = dict(v)
        v["표준이슈"] = mapping.get(v["이슈"], v["이슈"]) if v["이슈"] != "미분류" else "미분류"
        result[resp_id] = v
    return result


def aggregate_issues(df, extracted, unit_level, min_n):
    """표준 이슈 x 성격: 건수/비율/평균 점검대상점수/심각도≥2 비율.
    표준 이슈 x 실(unit_level)은 n>=min_n인 셀만 표시."""
    rows = []
    id_to_row = df.set_index("id")
    for resp_id, v in extracted.items():
        if resp_id not in id_to_row.index:
            continue
        r = id_to_row.loc[resp_id]
        rows.append({
            "id": resp_id, "표준이슈": v["표준이슈"], "성격": v["성격"],
            "심각도": r["심각도"], "점검대상점수": r["점검대상점수"], "unit": r[unit_level],
        })
    issue_df = pd.DataFrame(rows)
    if issue_df.empty:
        return {}, {}, []

    n_total = len(issue_df)
    issue_agg = {}
    for issue, g in issue_df.groupby("표준이슈", sort=False):
        if issue == "미분류":
            continue
        n = len(g)
        severe = (pd.to_numeric(g["심각도"], errors="coerce").fillna(0) >= 2).sum()
        scores = g["점검대상점수"].dropna()
        personality = g["성격"].value_counts().to_dict()
        issue_agg[issue] = {
            "n": int(n), "pct": n / n_total if n_total else 0.0,
            "severe_ge2_pct": float(severe / n) if n else 0.0,
            "check_target_score_mean": float(scores.mean()) if len(scores) else None,
            "personality": {k: int(v) for k, v in personality.items()},
            "severe_ge2_n": int(severe),
        }

    issue_unit = {}
    for (issue, unit), g in issue_df[issue_df["표준이슈"] != "미분류"].groupby(["표준이슈", "unit"], sort=False):
        if len(g) >= min_n:
            sev = pd.to_numeric(g["심각도"], errors="coerce").fillna(0)
            issue_unit.setdefault(issue, {})[unit] = {
                "n": int(len(g)),
                "severe_ge2_pct": float((sev >= 2).sum() / len(g)),
            }

    top5 = sorted(issue_agg.items(),
                  key=lambda kv: (kv[1]["severe_ge2_n"], kv[1]["check_target_score_mean"] or 0),
                  reverse=True)[:5]
    top5_issues = [k for k, _ in top5]

    return issue_agg, issue_unit, top5_issues


def sample_texts_for_issue(df, extracted, issue, sample_n, seed):
    """표준 이슈별 대표 응답 샘플: 점검대상점수 높은 순 -> 원문 중복 제거 ->
    8~300자 -> 상위 sample_n건."""
    ids = [rid for rid, v in extracted.items() if v["표준이슈"] == issue]
    sub = df[df["id"].isin(ids)].copy()
    sub = sub.sort_values("점검대상점수", ascending=False, na_position="last")
    sub = sub.drop_duplicates(subset="원문", keep="first")
    sub = sub[sub["원문"].apply(lambda t: isinstance(t, str) and 8 <= len(t.strip()) <= 300)]
    return list(zip(sub["id"].head(sample_n).tolist(), sub["원문"].head(sample_n).tolist()))


def summarize_issue(issue, samples, config=CONFIG, cache=None):
    """표준 이슈 하나에 대해 (1) 원문에서 반복되는 패턴·배경을 분석한
    3~4문장 서술형 '분석', (2) 핵심 요청 최대 3개(각 한 줄), (3) 대표 의견
    2개(일반화한 한 문장)를 만들고, 요청/의견 각 문장에 근거 응답 ID를
    붙입니다. "분석"은 단순 재요약이 아니라 응답에서 실제로 반복되는
    구체적 상황/표현을 근거로 들고 그게 뭘 시사하는지까지 담아야 합니다
    (LLM_SYSTEM_PROMPT에 이미 같은 지시가 있지만 여기서도 명시)."""
    if not samples:
        return {"analysis": "", "requests": [], "quotes": []}
    cache = cache if cache is not None else {}
    masked_samples = [(rid, mask_identifying_info(text)) for rid, text in samples]
    listing = "\n".join(f"[{rid}] {text}" for rid, text in masked_samples)
    user = (
        f"이슈: {issue}\n아래는 이 이슈로 분류된 실제 응답입니다(각 줄 앞 [ID]는 근거 "
        f"인용용 응답 ID, 총 {len(samples)}건):\n{listing}\n\n"
        "다음을 JSON으로만 답하세요(다른 설명 없이):\n"
        "{\"analysis\": \"이 응답들에서 반복되는 구체적 상황/표현이 무엇인지, 그것이 "
        "조직 운영에 시사하는 바가 무엇인지를 3~4문장으로 분석(단순 재요약 금지, "
        "근거가 된 패턴을 구체적으로 언급)\", "
        "\"requests\": [\"핵심 요청 한 줄\", ...최대 3개], "
        "\"quotes\": [{\"text\": \"대표 의견을 일반화한 한 문장(원문 그대로 쓰지 말고 "
        "개인 식별 요소 제거)\", \"evidence_ids\": [\"근거로 쓴 응답 ID\"]}, ...최대 2개]}"
    )
    raw = call_llm(LLM_SYSTEM_PROMPT, user, cache, config)
    if not raw:
        return {"analysis": "", "requests": [], "quotes": []}
    try:
        start, end = raw.index("{"), raw.rindex("}") + 1
        parsed = json.loads(raw[start:end])
        return {
            "analysis": parsed.get("analysis", ""),
            "requests": parsed.get("requests", [])[:3],
            "quotes": parsed.get("quotes", [])[:2],
        }
    except (ValueError, json.JSONDecodeError):
        return {"analysis": "", "requests": [], "quotes": []}


def sample_texts_for_unit(df, extracted, unit, unit_level, sample_n, seed):
    """실(또는 팀) 하나의 토론 포인트 생성용 대표 응답 샘플: 그 조직 소속이면서
    표준 이슈가 부여된(미분류 제외) 응답 중 점검대상점수 높은 순 -> 원문
    중복 제거 -> 8~300자 -> 상위 sample_n건."""
    classified_ids = {rid for rid, v in extracted.items() if v.get("표준이슈") not in (None, "미분류")}
    sub = df[(df[unit_level] == unit) & (df["id"].isin(classified_ids))].copy()
    sub = sub.sort_values("점검대상점수", ascending=False, na_position="last")
    sub = sub.drop_duplicates(subset="원문", keep="first")
    sub = sub[sub["원문"].apply(lambda t: isinstance(t, str) and 8 <= len(t.strip()) <= 300)]
    return list(zip(sub["id"].head(sample_n).tolist(), sub["원문"].head(sample_n).tolist()))


def summarize_unit_talking_point(unit, unit_stats, samples, config=CONFIG, cache=None):
    """실(또는 팀) 하나에 대한 "조직별 핵심 포인트" — 사업부 평균 대비 그
    조직만의 편차와 대표 응답을 근거로 2~3문장. 소규모 합산 그룹("기타")에는
    호출하지 않습니다(익명성 — 호출부에서 걸러서 넘겨야 함)."""
    if not samples:
        return ""
    cache = cache if cache is not None else {}
    masked_samples = [(rid, mask_identifying_info(text)) for rid, text in samples]
    listing = "\n".join(f"[{rid}] {text}" for rid, text in masked_samples)
    dev = unit_stats["deviation_pp"]
    facts = (f"이 조직 응답 {unit_stats['n']}건, 사업부 평균 대비 편차: 심각도1 {dev[1]:+.1f}%p, "
             f"심각도2 {dev[2]:+.1f}%p, 심각도3 {dev[3]:+.1f}%p, 점검대상 비율 {unit_stats['check_target_pct']*100:.1f}%")
    user = (
        f"조직: {unit}\n확정 수치(전사 데이터는 없으니 전사와 비교하지 말고 사업부 평균 대비로만 "
        f"비교하세요): {facts}\n대표 응답(각 줄 앞 [ID]는 근거 인용용):\n{listing}\n\n"
        "이 조직만의 특징적인 토론 포인트를 2~3문장으로 작성해 주세요. 사업부 평균과 비교해 "
        "이 조직에서 두드러지는 점이 무엇인지, 응답 내용과 수치에 근거해서만 쓰고 과장하지 마세요."
    )
    return call_llm(LLM_SYSTEM_PROMPT, user, cache, config)


def generate_unit_talking_points(df, extracted, by_unit, unit_level, sample_n, seed, config=CONFIG, cache=None):
    """소규모 합산 그룹("기타")을 제외한 실제 조직마다 토론 포인트를 생성해
    {unit: text} 딕셔너리로 반환합니다."""
    cache = cache if cache is not None else {}
    points = {}
    for unit, stats in by_unit.items():
        if stats["is_merged_small"]:
            continue
        samples = sample_texts_for_unit(df, extracted, unit, unit_level, sample_n, seed)
        points[unit] = summarize_unit_talking_point(unit, stats, samples, config, cache)
    return points


def build_brief_facts_text(division_summary, issue_agg, top5_issues, by_unit):
    """핵심 3줄/토론 안건 프롬프트에 넣을 확정 수치 텍스트."""
    lines = [
        f"유효 응답 수: {division_summary['n_valid']}건",
        "심각도 분포: " + ", ".join(
            f"{lvl}단계 {division_summary['severity_counts'][lvl]}건"
            f"({division_summary['severity_pct'][lvl]*100:.1f}%)" for lvl in (0, 1, 2, 3)),
        f"점검대상 {division_summary['check_target_count']}건, "
        f"즉시확인 {division_summary['immediate_count']}건",
    ]
    if top5_issues:
        lines.append("우선 점검 이슈 Top5: " + ", ".join(
            f"{issue}({issue_agg[issue]['n']}건, 심각도≥2 {issue_agg[issue]['severe_ge2_pct']*100:.1f}%)"
            for issue in top5_issues))
    unit_lines = []
    for unit, stats in by_unit.items():
        if stats["is_merged_small"]:
            continue
        dev = stats["deviation_pp"]
        unit_lines.append(f"{unit}(심각도2 {dev[2]:+.1f}%p, 심각도3 {dev[3]:+.1f}%p, 사업부 평균 대비)")
    if unit_lines:
        lines.append("하위 조직 간 편차(사업부 평균 대비): " + ", ".join(unit_lines))
    return "\n".join(lines)


def generate_brief_three_lines(facts_text, config=CONFIG, cache=None):
    user = (
        f"아래는 우리 사업부의 확정된 집계 수치입니다 — 여기 없는 숫자나 사실은 만들지 "
        f"말고 그대로만 활용하세요(전사 데이터는 없으니 전사와 비교하지 마세요):\n{facts_text}\n\n"
        "다음 3줄을 각 한 문장으로 작성해 주세요. 형식:\n"
        "가장 큰 이슈: ...\n하위 조직 간 차이: ...\n회의 제안 방향: ..."
    )
    raw = call_llm(LLM_SYSTEM_PROMPT, user, cache or {}, config)
    lines = {"biggest_issue": "", "unit_diff": "", "suggested_direction": ""}
    if not raw:
        return lines
    for line in raw.splitlines():
        line = line.strip()
        if line.startswith("가장 큰 이슈:"):
            lines["biggest_issue"] = line.split(":", 1)[1].strip()
        elif line.startswith("하위 조직 간 차이:"):
            lines["unit_diff"] = line.split(":", 1)[1].strip()
        elif line.startswith("회의 제안 방향:"):
            lines["suggested_direction"] = line.split(":", 1)[1].strip()
    return lines


def generate_discussion_agenda(facts_text, config=CONFIG, cache=None):
    user = (
        f"아래는 우리 사업부의 확정된 집계 수치입니다(전사 비교 금지):\n{facts_text}\n\n"
        "다음을 JSON으로만 답하세요(다른 설명 없이):\n"
        "{\"agenda\": [\"공유할 안건\", ...3개], \"expected_questions\": [\"예상 질문\", ...3개], "
        "\"cross_division_questions\": [\"타 사업부에 물어볼 질문\", ...3개]}"
    )
    raw = call_llm(LLM_SYSTEM_PROMPT, user, cache or {}, config)
    empty = {"agenda": [], "expected_questions": [], "cross_division_questions": []}
    if not raw:
        return empty
    try:
        start, end = raw.index("{"), raw.rindex("}") + 1
        parsed = json.loads(raw[start:end])
        return {
            "agenda": parsed.get("agenda", [])[:3],
            "expected_questions": parsed.get("expected_questions", [])[:3],
            "cross_division_questions": parsed.get("cross_division_questions", [])[:3],
        }
    except (ValueError, json.JSONDecodeError):
        return empty


# =============================================================
# 4. 차트 (matplotlib)
# =============================================================
def set_korean_font_for_matplotlib():
    """맑은 고딕을 우선 시도하고, 없는 환경(이 저장소의 테스트/CI 환경 등)
    에서는 얻을 수 있는 아무 sans-serif로 조용히 대체합니다."""
    import matplotlib
    candidates = ["Malgun Gothic", "맑은 고딕", "NanumGothic", "AppleGothic", "Noto Sans CJK KR"]
    available = {f.name for f in matplotlib.font_manager.fontManager.ttflist}
    for name in candidates:
        if name in available:
            matplotlib.rcParams["font.family"] = name
            return
    matplotlib.rcParams["font.family"] = "sans-serif"
    matplotlib.rcParams["axes.unicode_minus"] = False


def make_severity_and_type_chart(division_summary, by_type, out_path):
    """심각도 분포 막대 + 유형 Top5(비율, 심각도≥2 비율) — 한 이미지에
    2개 서브플롯."""
    import matplotlib.pyplot as plt
    set_korean_font_for_matplotlib()

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(9, 3.2))

    levels = [0, 1, 2, 3]
    counts = [division_summary["severity_counts"][lvl] for lvl in levels]
    ax1.bar([str(lvl) for lvl in levels], counts, color=f"#{COLOR_NAVY}")
    ax1.set_title("심각도 분포", fontsize=10)
    ax1.set_xlabel("심각도")
    for i, c in enumerate(counts):
        ax1.text(i, c, str(c), ha="center", va="bottom", fontsize=8)

    top5 = sorted(by_type.items(), key=lambda kv: kv[1]["n"], reverse=True)[:5]
    labels = [k for k, _ in top5]
    pcts = [v["pct"] * 100 for _, v in top5]
    ax2.barh(labels[::-1], pcts[::-1], color=f"#{COLOR_CORAL}")
    ax2.set_title("유형 Top5 (비율 %)", fontsize=10)
    for i, (label, v) in enumerate(zip(labels[::-1], [v for _, v in top5][::-1])):
        ax2.text(v["pct"] * 100, i, f"  심각도≥2 {v['severe_ge2_pct']*100:.0f}%", va="center", fontsize=7)

    fig.tight_layout()
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    return out_path


def make_unit_heatmap(heatmap_data, out_path, division_sev_pct):
    """실(또는 팀) x 심각도 1/2/3 비율 히트맵. 사업부 평균보다 높은 셀은
    글자를 굵게 강조."""
    import matplotlib.pyplot as plt
    import numpy as np
    set_korean_font_for_matplotlib()

    units = list(heatmap_data.keys())
    levels = [1, 2, 3]
    if not units:
        fig, ax = plt.subplots(figsize=(6, 1))
        ax.text(0.5, 0.5, "표시할 조직 데이터 없음", ha="center", va="center")
        ax.axis("off")
        fig.savefig(out_path, dpi=150)
        plt.close(fig)
        return out_path

    matrix = np.array([[heatmap_data[u][lvl] * 100 for lvl in levels] for u in units])
    fig, ax = plt.subplots(figsize=(6, 0.5 * len(units) + 1.2))
    im = ax.imshow(matrix, cmap="Oranges", aspect="auto")
    ax.set_xticks(range(len(levels)))
    ax.set_xticklabels([f"심각도{lvl}" for lvl in levels], fontsize=9)
    ax.set_yticks(range(len(units)))
    ax.set_yticklabels(units, fontsize=9)
    for i, u in enumerate(units):
        for j, lvl in enumerate(levels):
            val = matrix[i, j]
            is_high = val / 100 > division_sev_pct.get(lvl, 0.0)
            ax.text(j, i, f"{val:.0f}%", ha="center", va="center", fontsize=8,
                    fontweight="bold" if is_high else "normal",
                    color="black" if val < 60 else "white")
    fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
    fig.tight_layout()
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    return out_path


# =============================================================
# 5. Word 보고서 (python-docx)
# =============================================================
def _set_run_korean_font(run, name=FONT_KR):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts")
    if rfonts is None:
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        rfonts = OxmlElement("w:rFonts")
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


def add_kr_heading(doc, text, level=1, size=None):
    from docx.shared import Pt, RGBColor
    p = doc.add_paragraph()
    run = p.add_run(text)
    _set_run_korean_font(run)
    run.font.bold = True
    run.font.size = Pt(size or (16 if level == 1 else 12))
    run.font.color.rgb = RGBColor.from_string(COLOR_NAVY)
    return p


def set_cell_background(cell, color_hex):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color_hex)
    cell._tc.get_or_add_tcPr().append(shd)


def set_page_layout(doc):
    from docx.shared import Cm
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)


def safe_filename(name):
    return re.sub(r'[\\/:*?"<>|]', "_", str(name))


def build_kpi_table(doc, division_summary):
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    import docx as docx_module

    labels = ["유효 응답수", "심각도 1 비율", "심각도 2 비율", "점검대상 건수", "즉시확인 건수"]
    values = [
        str(division_summary["n_valid"]),
        f"{division_summary['severity_pct'][1]*100:.1f}%",
        f"{division_summary['severity_pct'][2]*100:.1f}%",
        str(division_summary["check_target_count"]),
        str(division_summary["immediate_count"]),
    ]
    table = doc.add_table(rows=2, cols=len(labels))
    table.style = "Table Grid"
    for i, label in enumerate(labels):
        cell = table.cell(0, i)
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(label)
        _set_run_korean_font(run)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = docx_module.shared.RGBColor.from_string("FFFFFF")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_background(cell, COLOR_NAVY)
    for i, val in enumerate(values):
        cell = table.cell(1, i)
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(val)
        _set_run_korean_font(run)
        run.font.bold = True
        run.font.size = Pt(13)
        is_immediate = i == len(values) - 1
        run.font.color.rgb = docx_module.shared.RGBColor.from_string(COLOR_CORAL if is_immediate else COLOR_GOLD)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_background(cell, COLOR_NAVY)


def add_issue_unit_severity_table(doc, issue_agg, issue_unit, config=CONFIG):
    """표준 이슈 x 조직(실/팀) 심각도 매트릭스 표. 셀 = "n건(심각도≥2 XX%)",
    MIN_N 미만이라 집계에서 빠진 조합은 "-"로 표시(익명성 — aggregate_issues가
    이미 걸러서 넘긴 issue_unit 그대로 사용하므로 여기선 추가 필터링 불필요)."""
    import docx
    from docx.shared import Pt

    units = sorted({u for per_unit in issue_unit.values() for u in per_unit})
    issues = [i for i in issue_agg if i in issue_unit]
    if not units or not issues:
        add_kr_paragraph(doc, f"조직별로 표시할 만큼 응답이 모인 이슈가 없습니다"
                               f"(조직당 n<{config['MIN_N']}).", size=9, color_hex="999999")
        return

    table = doc.add_table(rows=1 + len(issues), cols=1 + len(units))
    table.style = "Table Grid"
    header_cells = ["표준 이슈"] + units
    for i, h in enumerate(header_cells):
        cell = table.cell(0, i)
        r = cell.paragraphs[0].add_run(h)
        _set_run_korean_font(r)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = docx.shared.RGBColor.from_string("FFFFFF")
        set_cell_background(cell, COLOR_NAVY)

    for row_i, issue in enumerate(issues, start=1):
        cell = table.cell(row_i, 0)
        r = cell.paragraphs[0].add_run(issue)
        _set_run_korean_font(r)
        r.font.bold = True
        r.font.size = Pt(8.5)
        for col_i, unit in enumerate(units, start=1):
            stats = issue_unit.get(issue, {}).get(unit)
            cell = table.cell(row_i, col_i)
            text = f"{stats['n']}건({stats['severe_ge2_pct']*100:.0f}%)" if stats else "-"
            r = cell.paragraphs[0].add_run(text)
            _set_run_korean_font(r)
            r.font.size = Pt(8.5)
            if stats and stats["severe_ge2_pct"] >= 0.5:
                r.font.color.rgb = docx.shared.RGBColor.from_string(COLOR_CORAL)
                r.font.bold = True


def build_brief_docx(div_name, division_summary, by_type, major_dist, by_unit,
                      heatmap_severity, issue_agg, issue_unit, top5_issues, issue_summaries,
                      three_lines, agenda, quality, year_trend, heatmap_type_top5,
                      unit_talking_points, tmp_dir, out_path, config=CONFIG):
    import docx
    from docx.shared import Cm

    doc = docx.Document()
    set_page_layout(doc)

    # ---------- 1페이지: 브리프 ----------
    add_kr_heading(doc, div_name, level=1, size=18)
    add_kr_paragraph(doc, "부서장에게 하고 싶은 말 · 토론용 브리프", size=10, color_hex="808080")

    build_kpi_table(doc, division_summary)
    doc.add_paragraph()

    if any(three_lines.values()):
        add_kr_heading(doc, "핵심 3줄 (AI 초안)", level=2, size=11)
        if three_lines["biggest_issue"]:
            add_kr_paragraph(doc, f"· 가장 큰 이슈: {three_lines['biggest_issue']}", size=9.5)
        if three_lines["unit_diff"]:
            add_kr_paragraph(doc, f"· 하위 조직 간 차이: {three_lines['unit_diff']}", size=9.5)
        if three_lines["suggested_direction"]:
            add_kr_paragraph(doc, f"· 회의 제안 방향: {three_lines['suggested_direction']}", size=9.5)

    chart_path = Path(tmp_dir) / "severity_type_chart.png"
    make_severity_and_type_chart(division_summary, by_type, chart_path)
    doc.add_picture(str(chart_path), width=Cm(16.5))

    if heatmap_severity:
        add_kr_heading(doc, f"{config['UNIT_LEVEL']}별 심각도 히트맵 (사업부 평균 대비 강조)", level=2, size=10.5)
        heatmap_path = Path(tmp_dir) / "unit_heatmap.png"
        make_unit_heatmap(heatmap_severity, heatmap_path, division_summary["severity_pct"])
        doc.add_picture(str(heatmap_path), width=Cm(15))

    if top5_issues:
        add_kr_heading(doc, "우선 점검 이슈 Top5 — 핵심 요청", level=2, size=10.5)
        for issue in top5_issues:
            stats = issue_agg[issue]
            add_kr_paragraph(doc, f"[{issue}] {stats['n']}건, 심각도≥2 {stats['severe_ge2_pct']*100:.0f}%",
                              size=8.5, bold=True)
            for req in issue_summaries.get(issue, {}).get("requests", []):
                add_kr_paragraph(doc, f"  - {req}", size=8)

    if any(agenda.values()):
        add_kr_heading(doc, "토론 안건 후보", level=2, size=10.5)
        for a in agenda.get("agenda", []):
            add_kr_paragraph(doc, f"· {a}", size=9)

    # ---------- 부록 (2페이지 이후) ----------
    doc.add_page_break()
    add_kr_heading(doc, "부록", level=1, size=14)

    add_kr_heading(doc, "표준 이슈 전체 요약", level=2)
    for issue, stats in issue_agg.items():
        add_kr_paragraph(doc, f"{issue} — {stats['n']}건 ({stats['pct']*100:.1f}%), "
                               f"심각도≥2 {stats['severe_ge2_pct']*100:.1f}%", size=9.5, bold=True)
        summary = issue_summaries.get(issue, {})
        if summary.get("analysis"):
            add_kr_paragraph(doc, f"  {summary['analysis']}", size=9, italic=True, color_hex="333333")
        for req in summary.get("requests", []):
            add_kr_paragraph(doc, f"  - {req}", size=9)
        for quote in summary.get("quotes", []):
            ids = ", ".join(quote.get("evidence_ids", []))
            add_kr_paragraph(doc, f"  \"{quote.get('text', '')}\" (근거: {ids})", size=8.5, italic=True,
                              color_hex="555555")

    add_kr_heading(doc, "표준 이슈 x 조직별 심각도", level=2)
    add_kr_paragraph(doc, f"셀 = 응답 건수(심각도≥2 비율). 조직당 응답 n<{config['MIN_N']}인 "
                           f"조합은 \"-\"로 표시(익명성 보호).", size=8, color_hex="999999")
    add_issue_unit_severity_table(doc, issue_agg, issue_unit, config)

    add_kr_heading(doc, f"{config['UNIT_LEVEL']}별 상세 ({config['UNIT_LEVEL']} x 유형 Top3) 및 조직별 핵심 포인트",
                   level=2)
    for unit, stats in by_unit.items():
        types_txt = ", ".join(f"{t}({c}건)" for t, c in stats["top_types"]) or "-"
        add_kr_paragraph(doc, f"{unit}: 응답 {stats['n']}건, 점검대상 {stats['check_target_pct']*100:.1f}%, "
                               f"즉시확인 {stats['immediate_count']}건, 유형 Top3: {types_txt}", size=9, bold=True)
        talking_point = unit_talking_points.get(unit)
        if talking_point:
            add_kr_paragraph(doc, f"  {talking_point}", size=8.5, italic=True, color_hex="555555")

    add_kr_heading(doc, "대분류 분포", level=2)
    add_kr_paragraph(doc, ", ".join(f"{m}({c}건)" for m, c in major_dist.items()) or "-", size=9)

    add_kr_heading(doc, "품질 지표", level=2)
    add_kr_paragraph(doc, f"코드확률 낮은 건(<0.5) 비율: {quality['low_prob_pct']*100:.1f}%", size=9)
    add_kr_paragraph(doc, f"동일응답수 과다(>=3) 비중: {quality['dup_heavy_pct']*100:.1f}%", size=9)

    if year_trend:
        add_kr_heading(doc, "연도별 추이", level=2)
        for year, stats in year_trend.items():
            add_kr_paragraph(doc, f"{year}년: 응답 {stats['n']}건, 점검대상 {stats['check_target_pct']*100:.1f}%",
                              size=9)

    if agenda.get("expected_questions") or agenda.get("cross_division_questions"):
        add_kr_heading(doc, "예상 질문 / 타 사업부 질문", level=2)
        for q in agenda.get("expected_questions", []):
            add_kr_paragraph(doc, f"(예상 질문) {q}", size=9)
        for q in agenda.get("cross_division_questions", []):
            add_kr_paragraph(doc, f"(타 사업부 질문) {q}", size=9)

    doc.add_paragraph()
    add_kr_paragraph(doc, "※ 비율은 유효 응답 기준입니다.", size=8, color_hex="888888")
    add_kr_paragraph(doc, "※ AI 요약은 초안이므로 원문 검증이 필요합니다.", size=8, color_hex="888888")
    add_kr_paragraph(doc, "※ 즉시확인·심각도 3 원문은 본 자료에 포함하지 않았습니다.", size=8, color_hex="888888")
    add_kr_paragraph(doc, f"※ 응답 n<{config['MIN_N']}인 조직은 합산 표시했습니다.", size=8, color_hex="888888")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


# =============================================================
# 6. 기타 산출물
# =============================================================
def write_results_json(payload, out_path):
    def _clean(obj):
        if isinstance(obj, dict):
            return {str(k): _clean(v) for k, v in obj.items()}
        if isinstance(obj, (list, tuple)):
            return [_clean(v) for v in obj]
        return obj
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(_clean(payload), f, ensure_ascii=False, indent=2, default=str)


def write_hr_only_immediate_csv(df, out_path):
    """즉시확인=O 또는 심각도=3 행만 별도 CSV로 — 원문은 이 파일에만 남기고
    브리프(.docx)에는 절대 넣지 않습니다."""
    mask = (df["즉시확인"] == "O") | (df["심각도"] == 3)
    cols = ["id", "실", "팀", "심각도", "즉시확인점수", "원문"]
    sub = df.loc[mask, cols]
    out_path.parent.mkdir(parents=True, exist_ok=True)
    sub.to_csv(out_path, index=False, encoding="utf-8-sig")


def write_verification_sample_xlsx(df, extracted, out_path, seed=42):
    """표준 이슈별 무작위 5건의 (응답ID, 원문, 추출 JSON, 표준 이슈) — 사람이
    분류를 대조할 수 있게 함. 브리프에는 포함하지 않음."""
    rng = random.Random(seed)
    by_issue = {}
    for resp_id, v in extracted.items():
        by_issue.setdefault(v["표준이슈"], []).append(resp_id)

    id_to_text = dict(zip(df["id"], df["원문"]))
    rows = []
    for issue, ids in by_issue.items():
        sample_ids = ids if len(ids) <= 5 else rng.sample(ids, 5)
        for rid in sample_ids:
            v = extracted[rid]
            rows.append({
                "응답ID": rid, "원문": id_to_text.get(rid, ""),
                "추출JSON": json.dumps(v, ensure_ascii=False), "표준이슈": issue,
            })
    out_df = pd.DataFrame(rows)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_df.to_excel(out_path, index=False)


# =============================================================
# 7. main
# =============================================================
def main(config=CONFIG):
    out_dir = Path(config["OUT_DIR"])
    tmp_dir = out_dir / "_tmp"
    out_dir.mkdir(parents=True, exist_ok=True)
    tmp_dir.mkdir(parents=True, exist_ok=True)

    print(f"엑셀 로딩 중... ({config['FILE_PATH']})")
    raw_df = load_raw_dataframe(config["FILE_PATH"], config["SHEET"])
    df = filter_valid_responses(raw_df)
    print(f"  유효 응답 {len(df)}건")

    div_name = config["DIV_NAME"] or df["사업부"].mode().iloc[0]
    df = df[df["사업부"] == div_name].reset_index(drop=True)
    print(f"  사업부 '{div_name}' 응답 {len(df)}건")

    long_df = explode_type_major_code(df)

    division_summary = compute_division_summary(df)
    by_type, major_dist = compute_type_summary(long_df, division_summary["n_valid"])
    by_unit = compute_unit_summary(df, long_df, config["UNIT_LEVEL"], config["MIN_N"],
                                    division_summary["severity_pct"])
    heatmap_severity = build_heatmap_unit_severity(by_unit)
    heatmap_type_top5, top5_types = build_heatmap_unit_type_top5(long_df, by_unit, config["UNIT_LEVEL"])
    quality = compute_quality_metrics(long_df, df)
    year_trend = compute_year_trend(df)

    if not config["BASE_URL"]:
        print("BASE_URL이 비어있어 LLM 사용 섹션(이슈 추출/요약/브리프 문구)을 건너뜁니다 "
              "(집계 기반 섹션은 정상 생성).")

    # 심각도 3 / 즉시확인=O는 LLM에 절대 전달하지 않음.
    llm_eligible = df[(df["심각도"] < 3) & (df["즉시확인"] != "O")]

    cache = load_llm_cache(config["OUT_DIR"])
    print("이슈 추출 중...")
    extracted = extract_all(llm_eligible, config, cache)
    mapping = normalize_issues(extracted, config["OUT_DIR"], config, cache)
    extracted = apply_issue_mapping(extracted, mapping)
    issue_agg, issue_unit, top5_issues = aggregate_issues(llm_eligible, extracted,
                                                           config["UNIT_LEVEL"], config["MIN_N"])

    print("이슈별 요약 생성 중...")
    issue_summaries = {}
    for issue in issue_agg:
        samples = sample_texts_for_issue(llm_eligible, extracted, issue, config["SAMPLE_N"],
                                          config["RANDOM_SEED"])
        issue_summaries[issue] = summarize_issue(issue, samples, config, cache)

    print("조직별 핵심 포인트 생성 중...")
    unit_talking_points = generate_unit_talking_points(llm_eligible, extracted, by_unit,
                                                        config["UNIT_LEVEL"], config["SAMPLE_N"],
                                                        config["RANDOM_SEED"], config, cache)

    facts_text = build_brief_facts_text(division_summary, issue_agg, top5_issues, by_unit)
    three_lines = generate_brief_three_lines(facts_text, config, cache)
    agenda = generate_discussion_agenda(facts_text, config, cache)

    save_llm_cache(config["OUT_DIR"], cache)

    print("Word 브리프 생성 중...")
    out_path = out_dir / f"{safe_filename(div_name)}_부서장의견_브리프.docx"
    build_brief_docx(div_name, division_summary, by_type, major_dist, by_unit,
                      heatmap_severity, issue_agg, issue_unit, top5_issues, issue_summaries,
                      three_lines, agenda, quality, year_trend, heatmap_type_top5,
                      unit_talking_points, tmp_dir, out_path, config)
    print(f"  {out_path}")

    results_payload = {
        "division_summary": division_summary, "by_type": by_type, "major_dist": major_dist,
        "by_unit": by_unit, "issue_agg": issue_agg, "issue_unit": issue_unit,
        "top5_issues": top5_issues, "issue_summaries": issue_summaries,
        "unit_talking_points": unit_talking_points,
        "three_lines": three_lines, "agenda": agenda, "quality": quality,
        "year_trend": year_trend,
    }
    write_results_json(results_payload, out_dir / "results.json")
    write_hr_only_immediate_csv(df, out_dir / "HR_only_즉시확인.csv")
    write_verification_sample_xlsx(llm_eligible, extracted, out_dir / "검증_샘플.xlsx",
                                    config["RANDOM_SEED"])

    shutil.rmtree(tmp_dir, ignore_errors=True)
    print("완료.")


if __name__ == "__main__":
    main()
