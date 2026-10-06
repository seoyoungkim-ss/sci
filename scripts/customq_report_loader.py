"""별도 특화문항 보고서(dashboard/customq_report.html)용 엑셀 -> JSON 변환기.

메인 대시보드의 custom_questions_loader.py와는 **완전히 별개의 간단한
표 형식**을 읽습니다(부서코드/대상인원/응답인원/사업부 컬럼 없음, 조직
계층도 이 파일 자체엔 없음 — 산하 부서 구성은 dashboard/customq_report.html
안에서 관리자가 직접 지정합니다).

워크북 시트 = 분기(예: 시트1 "3Q", 시트2 "1Q") — 시트명을 그대로 분기
라벨로 씁니다.

컬럼 레이아웃(1행 헤더, 데이터는 2행부터, 위치 고정):
  A: 부서명
  B: 응답률
  C~G: 특화문항 5개 — 고객 / 협업 / 기술경쟁력 / 실행 / 윤리
  H~J: 특화문항 3개 — 보고간소화 / 회의결론 / 자료작성

빈 셀/"-"(데이터 없음 표기)는 모두 null로 저장됩니다(화면에 "-"로
보여주는 건 HTML 쪽 책임).

DRM 때문에 이 저장소의 다른 로더들과 동일하게 xlwings로 엽니다.

Usage:
    python scripts/customq_report_loader.py "C:\\path\\to\\특화문항.xlsx"
    python scripts/customq_report_loader.py "C:\\folder" --recursive
"""
import argparse
import json
import math
import sys
from pathlib import Path

from segment_loader import resolve_workbook_paths
from xlwings_utils import close_or_detach, open_or_attach

DATA_DIR = Path(__file__).resolve().parent.parent / "data"

COL_DEPT_NAME = 0   # A
COL_RESPONSE_RATE = 1  # B
QUESTION_COLS = [
    (2, "고객"), (3, "협업"), (4, "기술경쟁력"), (5, "실행"), (6, "윤리"),
    (7, "보고간소화"), (8, "회의결론"), (9, "자료작성"),
]
HEADER_ROWS = 1  # 1행이 헤더, 데이터는 2행부터

NO_DATA_TEXT_VARIANTS = {"-", "–", "—", "N/A", "NA", "없음"}


def to_number(v):
    """셀 값을 숫자로. 빈 칸/"-"류 표기/숫자로 못 바꾸는 값은 전부 None —
    이 저장소의 다른 로더들(예: custom_questions_loader.py)과 동일한
    관용 규칙."""
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
    """응답률은 0~1 소수가 기준(다른 로더들과 동일) — 텍스트 "86.7%"처럼
    1을 넘는 값이 나오면 퍼센트로 보고 100으로 나눕니다."""
    n = to_number(v)
    if n is None:
        return None
    return n / 100.0 if n > 1 else n


def parse_one_sheet(values, sheet_label, verbose=False):
    """시트 하나(한 분기)의 raw values를 부서별 레코드 리스트로 변환."""
    if not values or len(values) <= HEADER_ROWS:
        return []
    records = []
    for row in values[HEADER_ROWS:]:
        if not row:
            continue
        dept_name = row[COL_DEPT_NAME] if COL_DEPT_NAME < len(row) else None
        if not isinstance(dept_name, str) or not dept_name.strip():
            continue
        dept_name = dept_name.strip()

        scores = {}
        for col, label in QUESTION_COLS:
            scores[label] = to_number(row[col]) if col < len(row) else None

        records.append({
            "dept_name": dept_name,
            "response_rate": parse_response_rate(row[COL_RESPONSE_RATE]) if COL_RESPONSE_RATE < len(row) else None,
            "scores": scores,
        })
    if verbose:
        print(f"  [{sheet_label}] {len(records)}개 부서 행 발견")
    return records


def parse_one_workbook(path, verbose=False):
    """워크북 하나를 열어 {시트명(분기 라벨): [레코드, ...]}를 반환."""
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
            records = parse_one_sheet(values, f"{fname}:{sheet.name}", verbose=verbose)
            if records:
                quarters.setdefault(sheet.name, []).extend(records)
    finally:
        close_or_detach(app, wb, owns_app)
    return quarters


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("workbooks", nargs="+",
                         help="One or more .xlsx paths AND/OR folder paths (폴더는 --recursive로 하위 폴더까지).")
    parser.add_argument("--out", default=None, help="Output JSON path (default: data/customq_report_data.json)")
    parser.add_argument("--recursive", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    workbook_paths = resolve_workbook_paths(args.workbooks, recursive=args.recursive)
    print(f"{len(workbook_paths)}개 파일 처리 시작")

    all_quarters = {}
    for path in workbook_paths:
        for quarter, records in parse_one_workbook(path, verbose=args.verbose).items():
            all_quarters.setdefault(quarter, []).extend(records)

    out_path = Path(args.out) if args.out else DATA_DIR / "customq_report_data.json"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"quarters": all_quarters}, f, ensure_ascii=False, indent=2)
    total = sum(len(v) for v in all_quarters.values())
    print(f"wrote {out_path} (분기 {len(all_quarters)}개: {list(all_quarters.keys())}, 총 {total}개 부서x분기 행)")


if __name__ == "__main__":
    main()
