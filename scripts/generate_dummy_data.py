"""Generate mock survey data + a sample "previous round" org hierarchy
so the dashboard can be built and demoed before real data is available.

Usage:
    python scripts/generate_dummy_data.py

Writes:
    data/dummy_survey_data.json   - one raw record per department
    data/org_map_2026Q2.json      - sample org hierarchy (previous round)

The org tree mirrors the company's real shape: 전사(1) - 본부(2) -
플랫폼/팀(3, the main reporting unit — 9 of them) - 하위팀/파트(4~5).
Most 플랫폼/팀 units still break down into level-4 sub-teams so there
are enough real respondent-level rows for the key-driver regression to
be numerically stable (see MIN_LEAF_SAMPLE_FOR_REGRESSION in
dashboard/index.html) — the dashboard's report mode can filter down to
just the 9 level-3 rows whenever a "main unit" view is needed. A couple
of deliberately tiny/low-response departments are kept in to exercise
the "low response count" / "표본 부족" paths.
"""
import json
import random
from pathlib import Path

from schema_utils import build_record, iter_item_defs, load_schema, NO_DATA_MARKER

random.seed(42)

DATA_DIR = Path(__file__).resolve().parent.parent / "data"

SURNAMES = ["김", "이", "박", "최", "정", "강", "조", "윤"]
GIVEN = ["민준", "서연", "지훈", "하은", "도윤", "수아", "예준", "지우", "시우", "채원"]

# 회사만족(Z) is deliberately generated as a function of a few "true"
# driver items plus noise, so the dashboard's key-driver regression has
# a real signal to recover (rather than pure noise across 11 predictors).
FINAL_SATISFACTION_ITEM = "회사만족"
FINAL_SATISFACTION_WEIGHTS = {"신뢰": 0.30, "소통": 0.22, "성과": 0.18, "직무만족": 0.12}

# code -> (leader_id, leader_name) overrides, used to make one leader
# officially own two sibling teams (국내영업실장 겸임 사례) so the
# leadership-comparison feature has a real multi-department example.
SHARED_LEADER = ("E50001", "한지원")


def random_name():
    return random.choice(SURNAMES) + random.choice(GIVEN)


# Nested org tree. Nodes without "children" are leaves that run their
# own survey (get real random scores); nodes with "children" are pure
# rollups (placeholder row, scores always recomputed from children).
#
# Level 3 ("플랫폼/팀") is the company's real main reporting unit — 9
# of them total, spread under 3 본부 (level 2), under 전사 (level 1).
# Most still break down into level-4 sub-teams (so there's enough real
# respondent-level data for the key-driver regression to be stable —
# see MIN_LEAF_SAMPLE_FOR_REGRESSION), but 파트너십팀 and 법무팀 are
# themselves leaves with no level-4 children, to keep exercising the
# "no children below this level" rollup case.
ORG_STRUCTURE = {
    "code": "C000", "name": "전사", "level": 1, "children": [
        {"code": "H100", "name": "제품본부", "level": 2, "children": [
            {"code": "R110", "name": "커머스플랫폼", "level": 3, "children": [
                {"code": "T111", "name": "프론트파트", "level": 4, "target_range": (15, 20)},
                {"code": "T112", "name": "백엔드파트", "level": 4, "children": [
                    {"code": "P1121", "name": "플랫폼셀", "level": 5, "target_range": (10, 13)},
                    {"code": "P1122", "name": "데이터셀", "level": 5, "target_range": (8, 11)},
                ]},
                {"code": "T113", "name": "QA파트", "level": 4, "target_range": (10, 14)},
            ]},
            {"code": "R120", "name": "결제플랫폼", "level": 3, "children": [
                {"code": "T121", "name": "결제개발팀", "level": 4, "target_range": (14, 18)},
                {"code": "T122", "name": "정산팀", "level": 4, "target_range": (10, 14)},
                {"code": "T123", "name": "PG연동팀", "level": 4, "target_range": (8, 12)},
            ]},
            {"code": "R130", "name": "데이터플랫폼", "level": 3, "children": [
                {"code": "T131", "name": "데이터엔지니어링팀", "level": 4, "target_range": (12, 16)},
                {"code": "T132", "name": "데이터분석팀", "level": 4, "target_range": (10, 14)},
                {"code": "T133", "name": "ML팀", "level": 4, "target_range": (8, 12)},
            ]},
        ]},
        {"code": "H200", "name": "영업본부", "level": 2, "children": [
            {"code": "R210", "name": "국내영업팀", "level": 3, "children": [
                {"code": "T211", "name": "1팀", "level": 4, "target_range": (16, 22), "leader": SHARED_LEADER},
                {"code": "T212", "name": "2팀", "level": 4, "target_range": (16, 22), "leader": SHARED_LEADER},
                {"code": "T213", "name": "3팀", "level": 4, "target_range": (16, 22)},
            ]},
            {"code": "R220", "name": "해외영업팀", "level": 3, "children": [
                {"code": "T221", "name": "아시아팀", "level": 4, "target_range": (20, 28)},
                {"code": "T222", "name": "미주팀", "level": 4, "target_range": (14, 18)},
                {"code": "T223", "name": "유럽팀", "level": 4, "target_range": (10, 15)},
            ]},
            {"code": "R230", "name": "파트너십팀", "level": 3, "target_range": (8, 12)},  # level-3 leaf, no level-4 children
        ]},
        {"code": "H300", "name": "지원본부", "level": 2, "children": [
            {"code": "R310", "name": "인사팀", "level": 3, "children": [
                {"code": "T311", "name": "채용팀", "level": 4, "target_range": (10, 14)},
                {"code": "T312", "name": "인사운영팀", "level": 4, "target_range": (12, 16)},
                {"code": "T313", "name": "조직문화팀", "level": 4, "target_range": (10, 14), "response_rate_override": 0.35},  # <50%: exercises the "경고" confidence badge
            ]},
            {"code": "R320", "name": "재무팀", "level": 3, "children": [
                {"code": "T321", "name": "회계팀", "level": 4, "target_range": (10, 14)},
                {"code": "T322", "name": "재무기획팀", "level": 4, "target_range": (8, 12)},
                {"code": "T323", "name": "세무팀", "level": 4, "target_range": (4, 5)},  # tiny: <5 responses expected
            ]},
            {"code": "R330", "name": "법무팀", "level": 3, "target_range": (8, 12)},  # level-3 leaf, no level-4 children
        ]},
    ]
}


def random_score(base, spread=6):
    return round(max(40.0, min(98.0, random.gauss(base, spread))), 1)


def random_yoy():
    if random.random() < 0.15:
        return NO_DATA_MARKER
    return round(random.uniform(-5.0, 5.0), 1)


def build_leaf_record(schema, node):
    target_count = random.randint(*node["target_range"])
    rate_draw = node.get("response_rate_override", random.uniform(0.65, 1.0))
    response_count = round(target_count * rate_draw)
    response_count = min(response_count, target_count)
    response_rate = round(response_count / target_count, 3) if target_count else 0.0

    area_base = {area["key"]: random.uniform(68, 88) for area in schema["areas"]}

    item_scores = {}
    question_scores = {}
    for area_key, item_key, _score_col, question_cols in iter_item_defs(schema):
        if item_key == FINAL_SATISFACTION_ITEM:
            continue  # derived below, as a function of the key-driver items
        item_base = area_base[area_key] + random.uniform(-4, 4)
        item_score = random_score(item_base, 5)
        item_scores[item_key] = item_score
        for i, _col in enumerate(question_cols, start=1):
            question_scores[f"{item_key}_{i}"] = random_score(item_score, 6)

    driven = sum(FINAL_SATISFACTION_WEIGHTS[k] * item_scores[k] for k in FINAL_SATISFACTION_WEIGHTS)
    noise_weight = 1 - sum(FINAL_SATISFACTION_WEIGHTS.values())
    noise = random.gauss(area_base["자랑스러운회사"], 6)
    company_sat = round(max(40.0, min(98.0, driven + noise_weight * noise)), 1)
    item_scores[FINAL_SATISFACTION_ITEM] = company_sat
    question_scores[f"{FINAL_SATISFACTION_ITEM}_1"] = random_score(company_sat, 4)

    area_scores = {}
    for area in schema["areas"]:
        keys = [it["key"] for it in area["items"]]
        area_scores[area["key"]] = round(sum(item_scores[k] for k in keys) / len(keys), 1)
    area_yoy = {area["key"]: random_yoy() for area in schema["areas"]}

    leader_id, leader_name = node.get("leader", (f"E{random.randint(10000, 99999)}", random_name()))

    return build_record(
        dept_level=node["level"],
        dept_name=node["name"],
        dept_code=node["code"],
        target_count=target_count,
        response_count=response_count,
        response_rate=response_rate,
        leader_id=leader_id,
        leader_name=leader_name,
        area_scores=area_scores,
        area_yoy=area_yoy,
        item_scores=item_scores,
        question_scores=question_scores,
    )


def build_parent_placeholder(node):
    """Non-leaf departments still need a leader on record, but their
    numeric survey stats are always recomputed via rollup, so they are
    stored as zeros/blank here and never read directly by the rollup
    logic (see scripts/schema_utils.py: rollup())."""
    return build_record(
        dept_level=node["level"],
        dept_name=node["name"],
        dept_code=node["code"],
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


def walk(node, parent_code, schema, records, org_map):
    org_map[node["code"]] = parent_code
    children = node.get("children")
    if children:
        records.append(build_parent_placeholder(node))
        for child in children:
            walk(child, node["code"], schema, records, org_map)
    else:
        records.append(build_leaf_record(schema, node))


def main():
    schema = load_schema()
    records = []
    org_map = {}
    walk(ORG_STRUCTURE, None, schema, records, org_map)

    leaf_count = sum(1 for r in records if r["target_count"] > 0)
    print(f"{len(records)} departments total, {leaf_count} leaf departments with survey data")

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    survey_path = DATA_DIR / "dummy_survey_data.json"
    with open(survey_path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    print(f"wrote {survey_path} ({len(records)} departments)")

    org_payload = {"round": "2026Q2", "based_on": None, "map": org_map}
    org_path = DATA_DIR / "org_map_2026Q2.json"
    with open(org_path, "w", encoding="utf-8") as f:
        json.dump(org_payload, f, ensure_ascii=False, indent=2)
    print(f"wrote {org_path} ({len(org_map)} departments)")


if __name__ == "__main__":
    main()
