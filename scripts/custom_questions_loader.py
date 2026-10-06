"""부서별 "특화문항"(사업부/조직마다 다를 수 있는 맞춤형 객관식 문항) 점수
워크북을 읽어 data/custom_questions_data.json을 씁니다. 다른 로더들과 같은
방식: DRM 때문에 xlwings로 열고, dept_code는 load_dept_code_lookup()으로
매칭합니다(파일 자체에 부서코드 컬럼이 있으면 그걸 우선 사용).

실제 컬럼 레이아웃(위치 고정, 1~2행이 헤더 — 1,2행이 병합된 컬럼도 있음):
  A: 부서명        (1~2행 병합)
  B: 부서코드      (1~2행 병합)
  C: 대상인원      (1행 "인원" 그룹 헤더, 2행 "대상인원")
  D: 응답인원      (2행 "응답인원")
  E: 응답률        (2행 "응답률")
  F~M: 특화문항 8개 (1행 "특화문항" 그룹 헤더, 2행에 각 문항의 실제 문장 —
       하드코딩하지 않고 이 문장을 그대로 읽어 결과 JSON의 키로 씁니다)
  N: 사업부

병합된 셀은 xlwings로 읽으면 맨 위/맨 왼쪽 셀에만 값이 들어오고 나머지는
None으로 옵니다. 그래서 각 컬럼의 실제 라벨은 "2행 값이 있으면 그것(그룹
안의 세부 라벨), 없으면 1행 값(단독으로 세로 병합된 컬럼)"으로 판정합니다.

실제 열 순서가 이 설명과 다르면 아래 COL_* 상수만 고치면 됩니다(위치 기반
컬럼 매핑 — 이 저장소의 다른 로더들과 동일한 방식).

분기(시트)별 데이터: 한 워크북 안에 여러 시트가 있고 각 시트가 서로 다른
분기(예: 1Q/2Q/3Q)를 나타내는 경우를 지원합니다 — 시트명을 그대로 분기
라벨로 써서 data/custom_questions_data.json을 { "quarters": { "<시트명>":
[레코드, ...] } } 형태로 씁니다(시트마다 같은 부서가 다시 나와도 서로
다른 분기로 보고 덮어쓰지 않음 — 예전엔 전체를 dept_code 하나로만 묶어서
같은 부서가 여러 시트에 있으면 마지막 시트 것만 남는 문제가 있었음).
분기마다 문항 문구가 바뀌거나(완전히 같은 문항이 아닐 수 있음) 조직
구성이 달라질 수 있어서, 분기 간 문항/조직을 어떻게 매칭할지는 이 스크립트가
결정하지 않고 대시보드의 "특화문항" 탭 관리자 매칭 UI에서 사람이 처리합니다
— 그래서 이 스크립트는 raw 데이터를 분기별로 그대로 넘기기만 합니다.

부서코드를 못 구한 행(파일에 부서코드 컬럼도 비어있고 survey_data.json
에도 이 이름이 없는 경우)도 건너뛰지 않고 dept_code: null로 그대로
포함시킵니다 — 조직 구성이 분기마다 달라 자동 매칭이 안 되는 경우를
대시보드의 관리자 매칭 UI에서 사람이 수동으로 연결할 수 있게 하기 위함
(경고는 그대로 출력해서 몇 건이나 수동 처리가 필요한지 바로 알 수 있음).

Usage:
    python scripts/custom_questions_loader.py "C:\\path\\to\\특화문항.xlsx"
    python scripts/custom_questions_loader.py "C:\\folder" --recursive
"""
import argparse
import json
import math
import sys
from pathlib import Path

from segment_loader import load_dept_code_lookup, resolve_workbook_paths
from xlwings_utils import close_or_detach, open_or_attach

DATA_DIR = Path(__file__).resolve().parent.parent / "data"

# 컬럼 위치(0-based) — 실제 열 순서가 다르면 여기만 고치면 됩니다.
COL_DEPT_NAME = 0        # A
COL_DEPT_CODE = 1        # B
COL_TARGET_COUNT = 2     # C
COL_RESPONSE_COUNT = 3   # D
COL_RESPONSE_RATE = 4    # E
COL_CUSTOM_START = 5     # F (특화문항 1번째)
COL_CUSTOM_END = 12      # M (특화문항 8번째, 0-based inclusive)
COL_BUSINESS_UNIT = 13   # N
HEADER_ROWS = 2          # 1~2행이 헤더, 데이터는 3행부터

NO_DATA_TEXT_VARIANTS = {"-", "–", "—", "N/A", "NA", "없음"}


def to_number(v):
    """segment_loader.py의 to_number()와 동일한 관용 규칙(퍼센트 기호/콤마/
    "점" 단위/수식 에러 허용) — 이 저장소 전체에서 같은 기준으로 셀을
    숫자로 바꾸기 위해 그대로 재사용합니다."""
    if v is None or v == "":
        return None
    if isinstance(v, str):
        s = v.strip()
        if s == "" or s in NO_DATA_TEXT_VARIANTS or s.startswith("#"):
            return None
        s = s.replace(",", "").rstrip("%").rstrip("점").strip()
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


def parse_response_rate(v):
    """응답률은 대시보드 전체에서 0~1 사이 소수로 저장하는 게 규칙인데
    (예: data/dummy_survey_data.json의 response_rate: 0.867), 이 파일의
    셀이 텍스트 "86.7%"로 들어오면 to_number()가 86.7을 그대로 돌려줘서
    100배 차이가 남. 1을 넘는 값은 퍼센트로 보고 100으로 나눕니다
    (subjective_report.py의 parse_prob_list()와 동일한 방어 로직)."""
    n = to_number(v)
    if n is None:
        return None
    return n / 100.0 if n > 1 else n


def parse_header_labels(values):
    """특화문항 각 컬럼(F~M)의 실제 문항 문장을 1~2행에서 뽑습니다.
    2행 값이 있으면 그걸 쓰고(그룹 헤더 밑 세부 라벨), 없으면 1행 값으로
    대체합니다(세로 병합된 단독 컬럼의 경우)."""
    row1 = values[0] if len(values) > 0 else []
    row2 = values[1] if len(values) > 1 else []

    def label_at(col):
        v2 = row2[col] if col < len(row2) else None
        if isinstance(v2, str) and v2.strip():
            return v2.strip()
        v1 = row1[col] if col < len(row1) else None
        return v1.strip() if isinstance(v1, str) else None

    labels = {}
    for col in range(COL_CUSTOM_START, COL_CUSTOM_END + 1):
        labels[col] = label_at(col) or f"특화문항_{col - COL_CUSTOM_START + 1}"
    return labels


def parse_one_sheet(values, sheet_label, dept_code_by_name, verbose=False):
    """시트 하나(한 분기)의 raw values를 부서별 레코드 리스트로 변환합니다.
    부서코드를 못 구해도 건너뛰지 않고 dept_code: None으로 포함시킵니다
    (대시보드 관리자 매칭 UI에서 수동 연결 대상)."""
    if not values or len(values) <= HEADER_ROWS:
        return []
    custom_labels = parse_header_labels(values)

    records = []
    unmatched = 0
    for row in values[HEADER_ROWS:]:
        if not row:
            continue
        dept_name = row[COL_DEPT_NAME] if COL_DEPT_NAME < len(row) else None
        if not isinstance(dept_name, str) or not dept_name.strip():
            continue
        dept_name = dept_name.strip()

        raw_code = row[COL_DEPT_CODE] if COL_DEPT_CODE < len(row) else None
        if raw_code not in (None, ""):
            dept_code = str(int(raw_code)) if isinstance(raw_code, float) and raw_code.is_integer() else str(raw_code).strip()
        else:
            dept_code = dept_code_by_name.get(dept_name)
        if not dept_code:
            unmatched += 1
            dept_code = None

        custom_questions = {}
        for col, label in custom_labels.items():
            score = to_number(row[col]) if col < len(row) else None
            if score is not None:
                custom_questions[label] = score

        biz_unit = row[COL_BUSINESS_UNIT] if COL_BUSINESS_UNIT < len(row) else None
        records.append({
            "dept_code": dept_code,
            "dept_name": dept_name,
            "사업부": biz_unit.strip() if isinstance(biz_unit, str) else None,
            "target_count": to_number(row[COL_TARGET_COUNT]) if COL_TARGET_COUNT < len(row) else None,
            "response_count": to_number(row[COL_RESPONSE_COUNT]) if COL_RESPONSE_COUNT < len(row) else None,
            "response_rate": parse_response_rate(row[COL_RESPONSE_RATE]) if COL_RESPONSE_RATE < len(row) else None,
            "custom_questions": custom_questions,
        })
    if unmatched:
        print(f"  warning: [{sheet_label}] {unmatched}개 부서의 dept_code를 자동으로 못 구해서 "
              f"dept_code: null로 포함시켰습니다 — 대시보드 '특화문항' 탭 관리자 매칭에서 "
              f"수동으로 연결하세요.", file=sys.stderr)
    if verbose:
        print(f"  [{sheet_label}] {len(records)}개 부서 행 발견 (미매칭 {unmatched}건)")
    return records


def parse_one_workbook(path, dept_code_by_name, verbose=False):
    """워크북 하나를 열어 {시트명(분기 라벨): [레코드, ...]}를 반환합니다.
    시트마다 독립적인 분기로 취급하므로, 같은 부서가 여러 시트(분기)에
    나와도 서로 덮어쓰지 않습니다."""
    quarters = {}
    app, wb, owns_app = open_or_attach(path)
    try:
        fname = Path(path).name
        for sheet in wb.sheets:
            values = sheet.used_range.value
            if not values:
                continue
            if not isinstance(values[0], list):  # 행이 1개뿐이면 xlwings가 1차원으로 줌
                values = [values]
            sheet_label = sheet.name
            records = parse_one_sheet(values, f"{fname}:{sheet_label}", dept_code_by_name, verbose=verbose)
            if records:
                quarters.setdefault(sheet_label, []).extend(records)
    finally:
        close_or_detach(app, wb, owns_app)
    return quarters


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("workbooks", nargs="+",
                         help="One or more .xlsx paths AND/OR folder paths (폴더는 --recursive로 하위 폴더까지).")
    parser.add_argument("--out", default=None, help="Output JSON path (default: data/custom_questions_data.json)")
    parser.add_argument("--recursive", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    dept_code_by_name = load_dept_code_lookup()
    if not dept_code_by_name:
        print("warning: no data/survey_data.json or dummy_survey_data.json found — "
              "부서코드 컬럼이 비어있는 행은 전부 dept_code: null로 들어갑니다(관리자 매칭 필요). "
              "메인 로더(또는 generate_dummy_data.py)를 먼저 돌려두면 자동 매칭률이 올라갑니다.",
              file=sys.stderr)

    workbook_paths = resolve_workbook_paths(args.workbooks, recursive=args.recursive)
    print(f"{len(workbook_paths)}개 파일 처리 시작")

    all_quarters = {}
    for path in workbook_paths:
        for quarter, records in parse_one_workbook(path, dept_code_by_name, verbose=args.verbose).items():
            all_quarters.setdefault(quarter, []).extend(records)

    out_path = Path(args.out) if args.out else DATA_DIR / "custom_questions_data.json"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"quarters": all_quarters}, f, ensure_ascii=False, indent=2)
    total = sum(len(v) for v in all_quarters.values())
    print(f"wrote {out_path} (분기 {len(all_quarters)}개: {list(all_quarters.keys())}, 총 {total}개 부서x분기 행)")


if __name__ == "__main__":
    main()
