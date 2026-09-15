"""Loads each department's segment-detail sheet (근속/성별/연령/직군/경력
개발단계/학력/사업부 비교 — see the dashboard's "세그먼트비교" tab) out of
one or more DRM-protected workbooks, and writes data/segment_data.json
in the shape the dashboard's "세그먼트 데이터 불러오기" file picker expects:

    { "departments": { "<dept_code>": { "dept_name", "response_count",
      "segments": { "<세그먼트명>": { "categories": [...],
      "areas": {...}, "items": {...} } } } } }

Usage (a single multi-sheet workbook — every sheet scanned):
    python scripts/segment_loader.py "C:\\path\\to\\survey.xlsx" [--out out.json] [--sheets "1팀" "2팀"]

Usage (44개 조직 각각 별도 파일인 경우 — 전사 포함, glob으로 한 번에):
    python scripts/segment_loader.py C:\\segments\\*.xlsx --expected-count 44

Multiple paths are always accepted (nargs="+") and every sheet in
every one of them is scanned independently — this covers "한 워크북에
여러 시트" and "조직마다 별도 파일" both, since it wasn't clear from the
request alone which shape the real export takes. --expected-count lets
you pass the total department count you expect (전사 최상위조직 포함)
and get an explicit OK/mismatch summary at the end instead of having to
notice a silent gap yourself.

WHY THIS IS SEPARATE FROM xlwings_loader.py: that script reads ONE
sheet with a FIXED column layout (every column letter hardcoded in
config/schema.json) because every department is a ROW in that same
sheet. This script instead expects ONE SHEET PER DEPARTMENT (title in
row 1-2 reading "문항별 결과: <조직명>"), and each sheet's segment
columns (얼마나 많은 사업부와 비교하는지 등) can differ department to
department — so column letters can't be hardcoded here. Everything is
derived at runtime by reading the row3/row4 headers, per the request
this was built from.

ASSUMPTIONS — please validate against a real workbook and tell me
exactly what's off rather than hand-patching the output, since this
was written from a text description of the layout, not a sample file:
  - Row 1 or 2 contains literal text "문항별 결과: <조직명>" identifying
    which department a sheet belongs to. dept_name is matched back to
    a dept_code via whatever survey_data.json (or the dummy data, as a
    fallback) is already on disk — load/refresh that first.
  - Row 3 = 대분류 (연도별결과추이/전년대비/긍정응답률/보통응답률/근속/
    성별/연령및세대/직군/경력개발단계/학력/사업장), merged across each
    category's columns; Row 4 = 소분류 label, merged across one
    category's 2 columns; Row 6 = 응답인원. Merged cells report their
    value only in the top-left cell when read via xlwings' used_range
    — every other cell in the merge comes back None — so both header
    rows are read with a "forward-fill" pass (carry the last non-None
    value rightward across the row).
  - Each 소분류 spans exactly 2 physical columns: a score column then a
    YoY column. Like the master sheet's own separate yoy_col, the 2nd
    column is assumed to already BE the precomputed delta — not a 2nd
    raw year that would need subtracting. If row 5 actually labels the
    pair as two literal years (e.g. "2025"/"2026") rather than a
    score/delta pair, this assumption is wrong and this script needs a
    small fix (subtract the two instead of reading col 2 as-is).
  - Columns D-E ("전체", 세그먼트 미적용 원점수) are skipped — the
    master sheet already covers company-wide numbers. Column scanning
    starts at F (index 6, 1-based) by default.
  - Within the area/item block, a row is read by which of columns A/B/C
    is non-empty (area name / item name / question text respectively)
    rather than assumed fixed row numbers or a fixed item count per
    area — the spec's row ranges (e.g. "Row9-18 즐거운일") didn't
    reconcile cleanly with the item/question counts already used
    elsewhere in this repo (config/schema.json), so reading the labels
    directly is the safer bet. Question-level rows are still parsed
    and kept in the raw pass but not included in reshape_for_dashboard
    output, since the dashboard's 세그먼트비교 tab only shows area/item
    level (문항 단위 세부는 생략하기로 한 이전 요청과 동일).
  - "사업장" 대분류는 명시적으로 무시합니다 (요청에 명시됨).
"""
import argparse
import json
import re
import sys
from pathlib import Path

from schema_utils import NO_DATA_MARKER

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
TITLE_RE = re.compile(r"문항별\s*결과\s*[:：]\s*(.+)")
IGNORED_MAJOR_CATEGORIES = {"사업장"}


def forward_fill(row):
    """Merged header cells report their value only once (top-left cell)
    when read via xlwings' used_range.value — repeat it across every
    column the merge actually spans, so a plain column-by-column scan
    sees the right 대분류/소분류 for every column."""
    filled = []
    last = None
    for v in row or []:
        if v not in (None, ""):
            last = v
        filled.append(last)
    return filled


def find_dept_name(values):
    for row in values[:2]:
        if not row:
            continue
        for cell in row:
            if isinstance(cell, str):
                m = TITLE_RE.search(cell)
                if m:
                    return m.group(1).strip()
    return None


def to_number(v):
    if v is None or v == "" or (isinstance(v, str) and v.strip() == NO_DATA_MARKER):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def build_column_map(values, header_row_1idx=3, sub_row_1idx=4, first_data_col_1idx=6):
    """Returns a list of {major, category, score_col, yoy_col} (0-based
    column indices) for every 소분류 column pair from first_data_col_1idx
    onward, skipping IGNORED_MAJOR_CATEGORIES and any column whose
    대분류 can't be read at all."""
    major_row = forward_fill(values[header_row_1idx - 1] if len(values) >= header_row_1idx else [])
    sub_row = forward_fill(values[sub_row_1idx - 1] if len(values) >= sub_row_1idx else [])
    n = max(len(major_row), len(sub_row))

    columns = []
    col = first_data_col_1idx - 1  # 0-based
    while col < n:
        major = major_row[col] if col < len(major_row) else None
        sub = sub_row[col] if col < len(sub_row) else None
        if major is None or str(major).strip() in IGNORED_MAJOR_CATEGORIES:
            col += 1
            continue
        columns.append({
            "major": str(major).strip(),
            "category": str(sub).strip() if sub not in (None, "") else f"col{col + 1}",
            "score_col": col,
            "yoy_col": col + 1,
        })
        col += 2
    return columns


def parse_area_item_rows(values, columns, row_start_1idx=9):
    """Walks rows top-to-bottom from row_start_1idx, using columns A/B/C
    (0-based indices 0/1/2) to tell area-total / item / question rows
    apart by whichever of the three is non-empty on that row — see the
    module docstring for why this isn't hardcoded to fixed row numbers.
    Returns (areas, items, questions), each {label: {category: {"score", "yoy"}}}."""
    areas, items, questions = {}, {}, {}

    for row in values[row_start_1idx - 1:]:
        if not row or all((c in (None, "") for c in row[:3])):
            continue  # blank separator row
        a = row[0] if len(row) > 0 else None
        b = row[1] if len(row) > 1 else None
        c = row[2] if len(row) > 2 else None

        if a not in (None, ""):
            target = areas.setdefault(str(a).strip(), {})
        elif b not in (None, ""):
            target = items.setdefault(str(b).strip(), {})
        elif c not in (None, ""):
            target = questions.setdefault(str(c).strip(), {})
        else:
            continue

        for col_def in columns:
            score = to_number(row[col_def["score_col"]] if col_def["score_col"] < len(row) else None)
            yoy = to_number(row[col_def["yoy_col"]] if col_def["yoy_col"] < len(row) else None)
            if score is None and yoy is None:
                continue
            target[col_def["category"]] = {"score": score, "yoy": yoy}

    return areas, items, questions


def reshape_for_dashboard(bucket, categories_in_order):
    """dashboard/index.html expects {label: {"score": [...], "yoy": [...]}}
    aligned by index to one shared `categories` array — not a dict
    keyed by category name — reshape here."""
    out = {}
    for label, by_cat in bucket.items():
        out[label] = {
            "score": [by_cat.get(cat, {}).get("score") for cat in categories_in_order],
            "yoy": [by_cat.get(cat, {}).get("yoy") for cat in categories_in_order],
        }
    return out


def parse_segment_sheet(values):
    """values: 2D list from sheet.used_range.value. Returns
    (dept_name, {대분류: {"categories": [...], "areas": {...}, "items": {...}}})
    or (None, {}) if this doesn't look like a 문항별 결과 sheet at all
    (e.g. the unrelated 4th "마스터 요약" sheet xlwings_loader.py reads)."""
    dept_name = find_dept_name(values)
    if dept_name is None:
        return None, {}

    columns = build_column_map(values)
    if not columns:
        print(f"  warning: '{dept_name}' — no segment columns found (row 3/4 headers unreadable?), skipping",
              file=sys.stderr)
        return dept_name, {}

    areas, items, _questions = parse_area_item_rows(values, columns)

    categories_by_major = {}
    for col_def in columns:
        categories_by_major.setdefault(col_def["major"], []).append(col_def["category"])

    segments = {}
    for major, categories in categories_by_major.items():
        segments[major] = {
            "categories": categories,
            "areas": reshape_for_dashboard(areas, categories),
            "items": reshape_for_dashboard(items, categories),
        }
    return dept_name, segments


def load_dept_code_lookup():
    """dept_name -> dept_code, from whatever survey data is already on
    disk (real data first, dummy data as fallback) — sheets identify
    departments by name only, so this bridges back to dept_code."""
    for candidate in (DATA_DIR / "survey_data.json", DATA_DIR / "dummy_survey_data.json"):
        if candidate.exists():
            with open(candidate, "r", encoding="utf-8") as f:
                records = json.load(f)
            return {r["dept_name"]: r["dept_code"] for r in records}
    return {}


def open_workbook(path):
    import xlwings as xw
    app = xw.App(visible=False)
    wb = app.books.open(path)
    return app, wb


def parse_one_workbook(path, dept_code_by_name, sheet_filter=None):
    """Opens a single .xlsx (one department's own file, or a multi-sheet
    workbook holding several departments — both are supported, since a
    "44개 조직별 상세 시트" description could mean either 44 sheets in
    one file or 44 separate files) and returns {dept_code: {...}} for
    every sheet in it that looks like a 문항별 결과 sheet. Prints one
    line per sheet it either parses or explicitly skips, so a run over
    many files/sheets is easy to audit against an expected total."""
    departments = {}
    app, wb = open_workbook(path)
    try:
        targets = [wb.sheets[s] for s in sheet_filter] if sheet_filter else list(wb.sheets)
        for sheet in targets:
            values = sheet.used_range.value
            if not values:
                continue
            dept_name, segments = parse_segment_sheet(values)
            if dept_name is None:
                continue  # not a 문항별 결과 sheet — silently skip (e.g. a master summary sheet)
            dept_code = dept_code_by_name.get(dept_name)
            if dept_code is None:
                print(f"  warning: '{dept_name}' (sheet '{sheet.name}' in {Path(path).name}) has no "
                      f"matching dept_code in survey data — skipped. Load/refresh survey data first, "
                      f"or check the name matches exactly.", file=sys.stderr)
                continue
            response_row = values[5] if len(values) > 5 else []  # Row6 (0-based index 5) = 응답인원
            response_count = next((int(v) for v in response_row if isinstance(v, (int, float))), None)
            departments[dept_code] = {
                "dept_name": dept_name,
                "response_count": response_count,
                "segments": segments,
            }
            print(f"  parsed {Path(path).name}:'{sheet.name}' -> {dept_name} ({dept_code}): {len(segments)} segment(s)")
    finally:
        wb.close()
        app.quit()
    return departments


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("workbooks", nargs="+",
                         help="One or more .xlsx paths — a single multi-sheet workbook (every sheet in it "
                              "scanned for a '문항별 결과: ...' title), several separate per-department "
                              "files, or a mix of both.")
    parser.add_argument("--out", default=None, help="Output JSON path (default: data/segment_data.json)")
    parser.add_argument("--sheets", nargs="*", default=None,
                         help="Restrict to these specific sheet names, applied to EVERY workbook given "
                              "(only meaningful with a single multi-sheet workbook). Default: scan all sheets.")
    parser.add_argument("--expected-count", type=int, default=None,
                         help="Total department count you expect across all files (전사 포함), e.g. 44 — "
                              "prints a clear mismatch warning naming which survey-data departments never "
                              "got a matching sheet, instead of leaving you to notice a silent gap yourself.")
    args = parser.parse_args()

    dept_code_by_name = load_dept_code_lookup()
    if not dept_code_by_name:
        print("warning: no data/survey_data.json or dummy_survey_data.json found — "
              "every sheet will be skipped since dept_name can't be matched to a dept_code. "
              "Run the main loader (or generate_dummy_data.py) first.", file=sys.stderr)

    departments = {}
    for path in args.workbooks:
        departments.update(parse_one_workbook(path, dept_code_by_name, args.sheets))

    out_path = Path(args.out) if args.out else DATA_DIR / "segment_data.json"
    payload = {"departments": departments}
    if out_path.exists():
        with open(out_path, "r", encoding="utf-8") as f:
            existing = json.load(f)
        existing.setdefault("departments", {}).update(departments)
        payload = existing

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    total = len(payload["departments"])
    print(f"wrote {out_path} ({total}개 부서 누적)")

    if args.expected_count is not None:
        if total == args.expected_count:
            print(f"OK: {total}/{args.expected_count}개 부서 모두 파싱됨")
        else:
            missing_names = sorted(set(dept_code_by_name) - {d["dept_name"] for d in payload["departments"].values()})
            print(f"WARNING: {total}/{args.expected_count}개만 파싱됨 (누락 {args.expected_count - total}개)",
                  file=sys.stderr)
            if missing_names:
                print(f"  survey 데이터엔 있지만 이번 실행에서 못 찾은 부서명: {', '.join(missing_names)}",
                      file=sys.stderr)


if __name__ == "__main__":
    main()
