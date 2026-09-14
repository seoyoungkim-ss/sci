"""Shared helpers for reading the survey column schema and building
per-department JSON records that both the xlwings loader and the dummy
data generator emit. Keep this in sync with config/schema.json and with
the SCHEMA constant embedded in dashboard/index.html.
"""
import json
from pathlib import Path

SCHEMA_PATH = Path(__file__).resolve().parent.parent / "config" / "schema.json"
NO_DATA_MARKER = "-"


def load_schema():
    with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def col_to_index(col: str) -> int:
    """Convert a spreadsheet column letter ('A', 'AA', 'BD', ...) to a
    1-based column index."""
    idx = 0
    for ch in col.strip().upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def iter_item_defs(schema):
    """Yield (area_key, item_key, score_col, question_cols) for every
    item defined in the schema, in table order."""
    for area in schema["areas"]:
        for item in area["items"]:
            yield area["key"], item["key"], item["score_col"], item["question_cols"]


def build_record(dept_level, dept_name, dept_code, target_count, response_count,
                  response_rate, leader_id, leader_name, area_scores, area_yoy,
                  item_scores, question_scores):
    """Assemble a department record in the canonical JSON shape used
    throughout the pipeline and the dashboard."""
    return {
        "dept_level": dept_level,
        "dept_name": dept_name,
        "dept_code": dept_code,
        "target_count": target_count,
        "response_count": response_count,
        "response_rate": response_rate,
        "leader_id": leader_id,
        "leader_name": leader_name,
        "area_scores": area_scores,
        "area_yoy": area_yoy,
        "item_scores": item_scores,
        "question_scores": question_scores,
    }


def weighted_average(values_and_weights):
    """values_and_weights: iterable of (value, weight). Returns None if
    total weight is 0 (nothing to average)."""
    total_weight = 0.0
    total = 0.0
    for value, weight in values_and_weights:
        if value is None or weight is None:
            continue
        total += value * weight
        total_weight += weight
    if total_weight == 0:
        return None
    return total / total_weight


def rollup(dept_code, records_by_code, children_by_parent, _cache=None):
    """Recursively compute the effective (rolled-up) record for
    dept_code, given:
      - records_by_code: {dept_code: raw record} for every department
        that has its own raw survey row
      - children_by_parent: {parent_code: [child_code, ...]} built from
        the org hierarchy JSON

    A department with no children in the hierarchy uses its own raw
    row as-is. A department with children ignores its own scores and
    is recomputed as the response-count-weighted average of its
    (recursively rolled-up) children. Leader info and YoY are always
    taken from the department's own raw row when present.
    """
    if _cache is None:
        _cache = {}
    if dept_code in _cache:
        return _cache[dept_code]

    children = children_by_parent.get(dept_code, [])
    own = records_by_code.get(dept_code)

    if not children:
        result = own
        _cache[dept_code] = result
        return result

    child_results = [rollup(c, records_by_code, children_by_parent, _cache) for c in children]
    child_results = [c for c in child_results if c is not None]
    if not child_results:
        result = own
        _cache[dept_code] = result
        return result

    target_count = sum(c["target_count"] for c in child_results)
    response_count = sum(c["response_count"] for c in child_results)
    response_rate = (response_count / target_count) if target_count else None

    area_scores = {}
    for area_key in child_results[0]["area_scores"].keys():
        area_scores[area_key] = weighted_average(
            (c["area_scores"].get(area_key), c["response_count"]) for c in child_results
        )

    item_scores = {}
    for item_key in child_results[0]["item_scores"].keys():
        item_scores[item_key] = weighted_average(
            (c["item_scores"].get(item_key), c["response_count"]) for c in child_results
        )

    question_scores = {}
    for q_key in child_results[0]["question_scores"].keys():
        question_scores[q_key] = weighted_average(
            (c["question_scores"].get(q_key), c["response_count"]) for c in child_results
        )

    if own is not None:
        area_yoy = own.get("area_yoy", {})
    else:
        area_yoy = {k: NO_DATA_MARKER for k in area_scores.keys()}

    result = build_record(
        dept_level=(own or child_results[0])["dept_level"],
        dept_name=(own or child_results[0])["dept_name"],
        dept_code=dept_code,
        target_count=target_count,
        response_count=response_count,
        response_rate=response_rate,
        leader_id=(own or {}).get("leader_id"),
        leader_name=(own or {}).get("leader_name"),
        area_scores=area_scores,
        area_yoy=area_yoy,
        item_scores=item_scores,
        question_scores=question_scores,
    )
    _cache[dept_code] = result
    return result
