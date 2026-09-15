"""Load the department health-survey sheet out of a DRM-protected Excel
workbook using xlwings (openpyxl/pandas.read_excel cannot open DRM
files, since they only read the raw zip/XML — xlwings drives the real
Excel application via COM/AppleScript instead).

TIP if this errors with "읽기 전용이거나 손상되었거나 암호화되어
있습니다" even though the file opens fine by double-clicking: this is
automatically retried a couple of times with the "읽기 전용으로
여시겠습니까?" prompt suppressed at the COM layer (the usual cause of
an automated open hanging/failing on a file that opens fine
interactively), and falls back to attaching to an already-open copy of
the same file if you happen to have opened it manually. See
xlwings_utils.open_or_attach() for the full explanation and what to
try if it still fails.

Usage:
    python scripts/xlwings_loader.py "C:\\path\\to\\survey.xlsx" [--out out.json]

The 4th sheet (wb.sheets[3]) is expected to have:
    row 1        merged title
    rows 2-3     two-row category header
    row 4+       one row per department, per config/schema.json

Output: a JSON array of department records in the same shape produced
by generate_dummy_data.py, ready to drop into data/ and load from the
dashboard's "부서 데이터 불러오기" file picker.
"""
import argparse
import json
import sys
from pathlib import Path

from schema_utils import NO_DATA_MARKER, col_to_index, iter_item_defs, load_schema
from xlwings_utils import close_or_detach, open_or_attach

DATA_DIR = Path(__file__).resolve().parent.parent / "data"


def open_sheet(workbook_path, sheet_index):
    # See xlwings_utils.open_or_attach()'s docstring: if the file
    # opens fine by double-clicking but a script gets Excel's generic
    # "읽기 전용이거나 손상되었거나 암호화되어 있습니다" error, this
    # attaches to an already-open copy instead of asking Excel to open
    # a fresh one via COM (which some DRM plugins can't decrypt through).
    app, wb, owns_app = open_or_attach(workbook_path)
    try:
        sheet = wb.sheets[sheet_index]
        used = sheet.used_range
        values = used.value  # 2D list, row-major, 1-based offsets tracked separately
        return values
    finally:
        close_or_detach(app, wb, owns_app)


def cell(row_values, col_letter, first_col_index):
    idx = col_to_index(col_letter) - first_col_index
    if idx < 0 or idx >= len(row_values):
        return None
    return row_values[idx]


def to_float(value):
    if value is None or value == "":
        return None
    if isinstance(value, str) and value.strip() == NO_DATA_MARKER:
        return NO_DATA_MARKER
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def parse_rows(all_values, schema):
    header_rows = schema["header_rows"]
    cols = schema["columns"]
    data_rows = all_values[header_rows:]
    first_col_index = 1  # used_range starts at column A in this script's assumption

    records = []
    for row in data_rows:
        if row is None:
            continue
        dept_code = cell(row, cols["dept_code"], first_col_index)
        if dept_code in (None, ""):
            continue  # blank trailing row

        target_count = to_float(cell(row, cols["target_count"], first_col_index)) or 0
        response_count = to_float(cell(row, cols["response_count"], first_col_index)) or 0
        response_rate = to_float(cell(row, cols["response_rate"], first_col_index))
        if response_rate is None and target_count:
            response_rate = response_count / target_count

        area_scores = {}
        area_yoy = {}
        for area in schema["areas"]:
            area_scores[area["key"]] = to_float(cell(row, area["score_col"], first_col_index))
            area_yoy[area["key"]] = to_float(cell(row, area["yoy_col"], first_col_index))

        item_scores = {}
        question_scores = {}
        for area_key, item_key, score_col, question_cols in iter_item_defs(schema):
            item_scores[item_key] = to_float(cell(row, score_col, first_col_index))
            for i, q_col in enumerate(question_cols, start=1):
                question_scores[f"{item_key}_{i}"] = to_float(cell(row, q_col, first_col_index))

        records.append({
            "dept_level": int(cell(row, cols["dept_level"], first_col_index) or 0),
            "dept_name": cell(row, cols["dept_name"], first_col_index),
            "dept_code": dept_code,
            "target_count": int(target_count),
            "response_count": int(response_count),
            "response_rate": response_rate,
            "leader_id": cell(row, cols["leader_id"], first_col_index),
            "leader_name": cell(row, cols["leader_name"], first_col_index),
            "area_scores": area_scores,
            "area_yoy": area_yoy,
            "item_scores": item_scores,
            "question_scores": question_scores,
        })
    return records


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("workbook", help="Path to the DRM-protected .xlsx file")
    parser.add_argument("--out", default=None, help="Output JSON path (default: data/survey_data.json)")
    args = parser.parse_args()

    schema = load_schema()
    all_values = open_sheet(args.workbook, schema["sheet_index"])
    if not all_values:
        print("No data found on the target sheet.", file=sys.stderr)
        sys.exit(1)

    records = parse_rows(all_values, schema)

    out_path = Path(args.out) if args.out else DATA_DIR / "survey_data.json"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    print(f"wrote {out_path} ({len(records)} departments)")


if __name__ == "__main__":
    main()
