# 부서 건강도 설문 분석 대시보드

사내 부서 건강도 설문 데이터를 조직 계층에 따라 집계/조회하는 대시보드입니다.
현재 단계는 **구조/파이프라인 목업**이며, 실제 설문 데이터는 추후 사내에서
`xlwings` 로더로 채워 넣습니다.

## 구성

```
config/schema.json          엑셀 컬럼 <-> JSON 필드 매핑 (단일 소스)
scripts/schema_utils.py     스키마 로딩 + 롤업(가중평균) 계산 로직
scripts/generate_dummy_data.py   더미 데이터 생성 (5레벨 조직 예시)
scripts/xlwings_loader.py   DRM 엑셀 4번째 시트 -> JSON 변환 (실 데이터용)
data/dummy_survey_data.json 생성된 더미 부서별 설문 데이터
data/org_map_2026Q2.json    샘플 "이전 회차" 조직 계층 JSON
dashboard/index.html        단일 HTML 대시보드 (조회모드 + 관리자모드)
```

## 데이터 파이프라인

1. **엑셀 -> JSON**: `python scripts/xlwings_loader.py "경로/설문.xlsx"`
   - DRM 파일은 openpyxl 등으로 열 수 없으므로 xlwings로 실제 Excel
     애플리케이션을 구동해 4번째 시트(`wb.sheets[3]`)를 읽습니다.
   - `config/schema.json`의 컬럼 매핑(A~BD)에 따라 부서레벨/인원/영역·항목·문항
     점수를 파싱하여 `data/survey_data.json`으로 저장합니다.
   - 전년비(L~N)가 `-`인 경우 문자열 `"-"` 그대로 유지되며, 대시보드에서
     "데이터 없음"으로 표시됩니다.

2. **조직 계층 JSON**: 대시보드 관리자 모드에서 부서 노드를 드래그앤드롭으로
   재배치한 뒤 "현재 회차로 저장" 버튼을 누르면 `org_map_<회차명>.json`
   (`{ "round": ..., "map": { "부서코드": "상위부서코드|null" } }`) 파일이
   다운로드됩니다. 다음 회차 작업 시 "이전 회차 불러오기"로 이 파일을 불러와
   변경된 부분만 수정하면 됩니다.
   - 엑셀 A컬럼의 부서레벨(1~5)은 참고용 숫자일 뿐이며, 실제 상하위 관계와
     롤업 계산의 기준은 이 JSON 트리입니다.

3. **롤업(집계)**: 하위 부서 -> 상위 부서로 점수를 올릴 때는 응답인원 가중평균을
   사용합니다 (`scripts/schema_utils.py`의 `rollup()`과 `dashboard/index.html`의
   동일 로직 `rollup()`이 서로 미러링되어 있습니다). 조직 트리 상 자식이 없는
   노드(예: 하위 조직이 없는 레벨2 부서)는 자기 자신의 원본 응답 데이터를
   그대로 사용합니다.

## 대시보드 실행

`dashboard/index.html` 을 브라우저로 열면 바로 동작합니다 (별도 서버/빌드 불필요).
처음에는 내장된 더미 데이터로 렌더링되며, 상단의 "부서 데이터 불러오기" /
관리자 모드의 "이전 회차 불러오기" 버튼으로 실제 JSON 파일을 불러올 수 있습니다.

- **조회 모드**: 좌측 조직도에서 부서 선택 -> 브레드크럼/드릴다운으로 이동.
  영역별 점수(전년비 포함), 항목별 점수(클릭 시 문항 펼치기), 응답률,
  부서장 정보를 확인할 수 있습니다.
- **관리자 모드**: 부서 노드를 드래그해서 다른 노드 위에 놓으면 하위 조직으로
  이동합니다. 최상위로 만들려면 상단 드롭존을 사용하세요. 순환 참조(상위 부서를
  자신의 하위 부서 밑으로 이동)는 자동으로 차단됩니다.

## 더미 데이터 재생성

```
cd scripts
python generate_dummy_data.py
```

`data/dummy_survey_data.json`, `data/org_map_2026Q2.json`을 다시 생성합니다.
대시보드에 내장된 더미 데이터를 갱신하려면 이 두 파일의 내용을
`dashboard/index.html`의 `DUMMY_SURVEY_DATA` / `DUMMY_ORG_MAP` 상수에
다시 붙여넣으세요.

## 다음 단계 (실 데이터 연결)

1. 사내 PC에서 `pip install -r requirements.txt` 후 `scripts/xlwings_loader.py`
   실행 (Excel + DRM 플러그인이 설치되어 있어야 함)
2. 생성된 `data/survey_data.json`을 대시보드에서 "부서 데이터 불러오기"로 로드
3. 관리자 모드에서 실제 조직 계층을 드래그앤드롭으로 구성 후 회차명으로 저장
