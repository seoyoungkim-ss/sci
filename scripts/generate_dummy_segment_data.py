"""Generates a small mock data/segment_data.json — the per-department
demographic segment detail (근속/성별/연령/직군/경력개발단계/학력) the
dashboard's "세그먼트비교" tab reads — so that feature has something to
demo/test with before segment_loader.py has a real workbook to run
against. Mirrors generate_dummy_data.py's approach (deterministic
random seed, scores centered on the department's own dummy area score).

Usage:
    python scripts/generate_dummy_segment_data.py

Writes: data/segment_data.json
"""
import json
import random
from pathlib import Path

random.seed(7)

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
AREA_ITEMS = {
    "즐거운일": ["몰입", "효율", "성장", "직무만족"],
    "함께하는동료": ["협력", "존중", "성과", "부서만족"],
    "자랑스러운회사": ["신뢰", "소통", "공정", "회사만족"],
}

SEGMENT_CATEGORIES = {
    "근속": ["3년미만", "5년미만", "10년미만", "15년미만", "15년이상"],
    "성별": ["남성", "여성"],
    "연령및세대": ["25세미만", "25-34세", "35-44세", "45-55세", "56세이상"],
    "직군": ["기획", "개발", "영업", "지원"],
    "경력개발단계": ["CL1", "CL2", "CL3", "CL4", "임원"],
    "학력": ["고졸", "초대졸", "대졸", "석사", "박사"],
}


def scored_series(base, n, spread=6):
    scores, yoys = [], []
    for _ in range(n):
        # ~8% chance a given category/area cell has no respondents at all
        if random.random() < 0.08:
            scores.append(None)
            yoys.append(None)
            continue
        scores.append(round(max(40.0, min(98.0, random.gauss(base, spread))), 1))
        yoys.append(round(random.uniform(-4.0, 4.0), 1) if random.random() > 0.1 else None)
    return scores, yoys


def build_department_segments(area_scores):
    segments = {}
    for seg_key, categories in SEGMENT_CATEGORIES.items():
        areas_out, items_out = {}, {}
        for area_key, base in area_scores.items():
            score, yoy = scored_series(base, len(categories))
            areas_out[area_key] = {"score": score, "yoy": yoy}
            for item_key in AREA_ITEMS[area_key]:
                iscore, iyoy = scored_series(base + random.uniform(-3, 3), len(categories))
                items_out[item_key] = {"score": iscore, "yoy": iyoy}
        segments[seg_key] = {"categories": categories, "areas": areas_out, "items": items_out}
    return segments


def main():
    survey_path = DATA_DIR / "dummy_survey_data.json"
    with open(survey_path, "r", encoding="utf-8") as f:
        records = json.load(f)

    departments = {}
    for r in records:
        if not r.get("target_count"):
            continue  # pure rollup placeholder, no leaf-level respondents of its own
        departments[r["dept_code"]] = {
            "dept_name": r["dept_name"],
            "response_count": r["response_count"],
            "segments": build_department_segments(r["area_scores"]),
        }

    out_path = DATA_DIR / "segment_data.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"departments": departments}, f, ensure_ascii=False, indent=2)
    print(f"wrote {out_path} ({len(departments)}개 부서)")


if __name__ == "__main__":
    main()
