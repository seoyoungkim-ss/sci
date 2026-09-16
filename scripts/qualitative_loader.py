"""Loads free-text ("주관식") survey responses out of one or more
DRM-protected workbooks, and writes data/qualitative_data.json in the
shape the dashboard's "주관식 데이터 불러오기" file picker expects:

    { "records": [ { "dept_code", "dept_name", "category", "text" }, ... ] }

Usage:
    python scripts/qualitative_loader.py "C:\\path\\to\\주관식.xlsx"
    python scripts/qualitative_loader.py "C:\\qualitative_folder" --recursive

Per the request this was built from, the source data is a single flat
table (可能 spread across a few files/sheets, not one-sheet-per-department
like segment_loader.py) with exactly 3 columns, identified by header
text rather than a fixed column letter since the column order wasn't
specified precisely:
  - "구분": one of 잘하고있는점 / 노력해야할점 / 부서장에게하고싶은말
    (kept as free text, not hardcoded to just these 3, in case a real
    file uses slightly different wording)
  - "부서명": matched back to a dept_code via whatever survey_data.json
    (or the dummy data, as a fallback) is already on disk — load/refresh
    that first, same requirement as segment_loader.py.
  - "내용": the free-text response itself.

The header row can be anywhere in the first few rows of a sheet (not
assumed to be row 1), and every sheet in every given workbook is
scanned independently, so this covers both "한 시트에 다 있다" and
"여러 시트/파일에 나눠져 있다" without needing to know which shape the
real export takes. Rows are simply concatenated across every workbook
given in one run (there's no natural per-row unique key to merge
against an existing output file the way segment_loader.py's per-
department JSON can, so re-running this OVERWRITES data/qualitative_data.json
rather than accumulating into it — pass every source file in one
invocation).

Keyword extraction itself happens in the dashboard (client-side JS),
not here — this script only normalizes the raw rows into JSON.
"""
import argparse
import json
import sys
from pathlib import Path

from segment_loader import load_dept_code_lookup, resolve_workbook_paths
from xlwings_utils import close_or_detach, open_or_attach

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
EXPECTED_HEADERS = {"구분": "category", "부서명": "dept_name", "내용": "text"}
HEADER_SCAN_ROWS = 5  # header row can be anywhere in the first few rows of a sheet


def find_header_row(values):
    """Returns (row_index_0based, {field_name: col_index_0based}) for the
    first row (within the first HEADER_SCAN_ROWS rows) containing all 3
    expected headers, or (None, {}) if no such row is found."""
    for row_idx, row in enumerate(values[:HEADER_SCAN_ROWS]):
        if not row:
            continue
        col_by_field = {}
        for col_idx, cell in enumerate(row):
            if not isinstance(cell, str):
                continue
            field = EXPECTED_HEADERS.get(cell.strip())
            if field:
                col_by_field[field] = col_idx
        if len(col_by_field) == len(EXPECTED_HEADERS):
            return row_idx, col_by_field
    return None, {}


def parse_sheet_rows(values, verbose=False, sheet_label=""):
    """Returns a list of {category, dept_name, text} dicts (dept_code not
    yet resolved) from one sheet's raw values, or [] if this sheet
    doesn't look like a 주관식 데이터 table at all."""
    header_row, col_by_field = find_header_row(values)
    if header_row is None:
        if verbose:
            print(f"  skip {sheet_label}: 구분/부서명/내용 헤더를 첫 {HEADER_SCAN_ROWS}행에서 찾지 못함", file=sys.stderr)
        return []

    rows = []
    for row in values[header_row + 1:]:
        if not row:
            continue
        category = row[col_by_field["category"]] if col_by_field["category"] < len(row) else None
        dept_name = row[col_by_field["dept_name"]] if col_by_field["dept_name"] < len(row) else None
        text = row[col_by_field["text"]] if col_by_field["text"] < len(row) else None
        if not isinstance(text, str) or not text.strip():
            continue  # blank response row -- nothing to extract keywords from
        if not isinstance(dept_name, str) or not dept_name.strip():
            continue  # can't attribute this response to any department
        rows.append({
            "category": category.strip() if isinstance(category, str) else "",
            "dept_name": dept_name.strip(),
            "text": text.strip(),
        })
    if verbose:
        print(f"  [{sheet_label}] {len(rows)}건 응답 발견 (헤더: 행 {header_row + 1})")
    return rows


def parse_one_workbook(path, dept_code_by_name, verbose=False):
    """Opens a single .xlsx and returns a list of resolved
    {dept_code, dept_name, category, text} records, scanning every
    sheet independently (see module docstring for why)."""
    records = []
    app, wb, owns_app = open_or_attach(path)
    try:
        fname = Path(path).name
        for sheet in wb.sheets:
            values = sheet.used_range.value
            if not values:
                continue
            for row in parse_sheet_rows(values, verbose=verbose, sheet_label=f"{fname}:{sheet.name}"):
                dept_code = dept_code_by_name.get(row["dept_name"])
                if dept_code is None:
                    print(f"  warning: '{row['dept_name']}' ({fname}:{sheet.name}) has no matching dept_code in "
                          f"survey data — skipped. Load/refresh survey data first, or check the name matches exactly.",
                          file=sys.stderr)
                    continue
                records.append({
                    "dept_code": dept_code,
                    "dept_name": row["dept_name"],
                    "category": row["category"],
                    "text": row["text"],
                })
    finally:
        close_or_detach(app, wb, owns_app)
    return records


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("workbooks", nargs="+",
                         help="One or more .xlsx paths AND/OR folder paths. A folder is expanded to every "
                              ".xlsx directly inside it (see --recursive for subfolders too).")
    parser.add_argument("--out", default=None, help="Output JSON path (default: data/qualitative_data.json)")
    parser.add_argument("--recursive", action="store_true",
                         help="When a given path is a folder, also scan its subfolders for .xlsx files.")
    parser.add_argument("--verbose", action="store_true",
                         help="Print a per-sheet breakdown of how many responses were found and where the "
                              "header row was located.")
    args = parser.parse_args()

    dept_code_by_name = load_dept_code_lookup()
    if not dept_code_by_name:
        print("warning: no data/survey_data.json or dummy_survey_data.json found — "
              "every row will be skipped since dept_name can't be matched to a dept_code. "
              "Run the main loader (or generate_dummy_data.py) first.", file=sys.stderr)

    workbook_paths = resolve_workbook_paths(args.workbooks, recursive=args.recursive)
    print(f"{len(workbook_paths)}개 파일 처리 시작")

    records = []
    for path in workbook_paths:
        records.extend(parse_one_workbook(path, dept_code_by_name, verbose=args.verbose))

    out_path = Path(args.out) if args.out else DATA_DIR / "qualitative_data.json"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"records": records}, f, ensure_ascii=False, indent=2)
    print(f"wrote {out_path} ({len(records)}건 응답)")


if __name__ == "__main__":
    main()
