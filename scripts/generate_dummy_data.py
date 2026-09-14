"""Generate mock survey data + a sample "previous round" org hierarchy
so the dashboard can be built and demoed before real data is available.

Usage:
    python scripts/generate_dummy_data.py

Writes:
    data/dummy_survey_data.json   - one raw record per department
    data/org_map_2026Q2.json      - sample org hierarchy (previous round)
"""
import json
import random
from pathlib import Path

from schema_utils import build_record, iter_item_defs, load_schema, NO_DATA_MARKER

random.seed(42)

DATA_DIR = Path(__file__).resolve().parent.parent / "data"

SURNAMES = ["김", "이", "박", "최", "정", "강", "조", "윤"]
GIVEN = ["민준", "서연", "지훈", "하은", "도윤", "수아", "예준", "지우", "시우", "채원"]


def random_name():
    return random.choice(SURNAMES) + random.choice(GIVEN)


# (dept_code, dept_name, dept_level, parent_code or None, target_count range)
ORG_TREE = [
    ("C000", "전사", 1, None, (260, 300)),
    ("H100", "제품본부", 2, "C000", (90, 100)),
    ("H200", "영업본부", 2, "C000", (70, 80)),
    ("H300", "지원본부", 2, "C000", (40, 50)),
    ("H400", "감사본부", 2, "C000", (8, 12)),  # level-2 dept with no level-3 children
    ("R110", "개발실", 3, "H100", (55, 65)),
    ("R120", "디자인실", 3, "H100", (20, 25)),
    ("R210", "국내영업실", 3, "H200", (40, 45)),
    ("R220", "해외영업실", 3, "H200", (25, 30)),
    ("R310", "경영지원실", 3, "H300", (40, 50)),
    ("T111", "백엔드팀", 4, "R110", (25, 30)),
    ("T112", "프론트팀", 4, "R110", (15, 20)),
    ("T121", "UX팀", 4, "R120", (20, 25)),
    ("T211", "1팀", 4, "R210", (18, 22)),
    ("T212", "2팀", 4, "R210", (18, 22)),
    ("T221", "아시아팀", 4, "R220", (25, 30)),
    ("T311", "인사팀", 4, "R310", (18, 22)),
    ("T312", "재무팀", 4, "R310", (18, 22)),
    ("P1111", "플랫폼파트", 5, "T111", (12, 15)),
    ("P1112", "데이터파트", 5, "T111", (10, 13)),
]

# codes that own raw survey data directly (leaves of ORG_TREE, i.e. the
# departments that actually run the survey rather than being pure rollups)
CHILD_CODES = {parent for *_, parent, _ in ORG_TREE if parent}
LEAF_CODES = {code for code, *_ in ORG_TREE if code not in CHILD_CODES}


def random_score(base, spread=6):
    return round(max(40.0, min(98.0, random.gauss(base, spread))), 1)


def random_yoy():
    if random.random() < 0.15:
        return NO_DATA_MARKER
    return round(random.uniform(-5.0, 5.0), 1)


def build_leaf_record(schema, dept_code, dept_name, dept_level, target_range):
    target_count = random.randint(*target_range)
    response_count = round(target_count * random.uniform(0.65, 1.0))
    response_count = min(response_count, target_count)
    response_rate = round(response_count / target_count, 3) if target_count else 0.0

    area_base = {area["key"]: random.uniform(68, 88) for area in schema["areas"]}
    area_scores = {key: random_score(base, 4) for key, base in area_base.items()}
    area_yoy = {key: random_yoy() for key in area_base.keys()}

    item_scores = {}
    question_scores = {}
    for area_key, item_key, _score_col, question_cols in iter_item_defs(schema):
        item_base = area_base[area_key] + random.uniform(-4, 4)
        item_score = random_score(item_base, 5)
        item_scores[item_key] = item_score
        for i, _col in enumerate(question_cols, start=1):
            question_scores[f"{item_key}_{i}"] = random_score(item_score, 6)

    return build_record(
        dept_level=dept_level,
        dept_name=dept_name,
        dept_code=dept_code,
        target_count=target_count,
        response_count=response_count,
        response_rate=response_rate,
        leader_id=f"E{random.randint(10000, 99999)}",
        leader_name=random_name(),
        area_scores=area_scores,
        area_yoy=area_yoy,
        item_scores=item_scores,
        question_scores=question_scores,
    )


def build_parent_placeholder(dept_code, dept_name, dept_level):
    """Non-leaf departments still need a leader on record, but their
    numeric survey stats are always recomputed via rollup, so they are
    stored as zeros/blank here and never read directly by the rollup
    logic (see scripts/schema_utils.py: rollup())."""
    return build_record(
        dept_level=dept_level,
        dept_name=dept_name,
        dept_code=dept_code,
        target_count=0,
        response_count=0,
        response_rate=0.0,
        leader_id=f"E{random.randint(10000, 99999)}",
        leader_name=random_name(),
        area_scores={},
        area_yoy={},
        item_scores={},
        question_scores={},
    )


def main():
    schema = load_schema()
    records = []
    for dept_code, dept_name, dept_level, _parent, target_range in ORG_TREE:
        if dept_code in LEAF_CODES:
            records.append(build_leaf_record(schema, dept_code, dept_name, dept_level, target_range))
        else:
            records.append(build_parent_placeholder(dept_code, dept_name, dept_level))

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    survey_path = DATA_DIR / "dummy_survey_data.json"
    with open(survey_path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    print(f"wrote {survey_path} ({len(records)} departments)")

    org_map = {code: parent for code, _name, _level, parent, _t in ORG_TREE}
    org_payload = {"round": "2026Q2", "based_on": None, "map": org_map}
    org_path = DATA_DIR / "org_map_2026Q2.json"
    with open(org_path, "w", encoding="utf-8") as f:
        json.dump(org_payload, f, ensure_ascii=False, indent=2)
    print(f"wrote {org_path} ({len(org_map)} departments)")


if __name__ == "__main__":
    main()
