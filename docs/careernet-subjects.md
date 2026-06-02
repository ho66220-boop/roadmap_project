# 커리어넷 학과별 선택과목 데이터

## 목적

커리어넷 학과정보의 2022 개정 교육과정 선택과목 자료를 수집해 기존 로드맵 DB의 학과-과목 그래프를 보강한다.

기존 로드맵 DB는 대학별 학과 트랙과 학교 편제표를 연결하는 데 강점이 있고, 커리어넷 자료는 학과별 일반 선택, 진로 선택, 융합 선택 권장 과목을 폭넓게 보완하는 역할을 한다.

## 출처

- 목록 API: `https://www.career.go.kr/cloud/api/major/uSearch`
- 상세 API: `https://www.career.go.kr/cloud/api/major/uView?seq={seq}`
- 원문 페이지 예시: `https://www.career.go.kr/cloud/w/major/uView?seq=20`

2026-06-02 기준 목록 API에서 501개 학과가 확인되었다.

## 산출물

수집 결과는 `data/careernet_major_subjects.csv`에 저장한다.

주요 컬럼:

- `career_seq`: 커리어넷 학과 seq
- `career_major_name`: 커리어넷 학과명
- `career_major_category`: 커리어넷 학과 분류
- `general_selection_subjects`: 일반 선택 관련 과목
- `career_selection_subjects`: 진로 선택 관련 과목
- `convergence_selection_subjects`: 융합 선택 관련 과목
- `source_text`: 커리어넷 상세 페이지에 표시된 원 출처 문구
- `source_url`: 커리어넷 상세 페이지 URL

## 재수집 방법

```powershell
python scripts\collect_careernet_major_subjects.py --delay 0.15 --output data\careernet_major_subjects.csv
```

샘플만 확인할 때:

```powershell
python scripts\collect_careernet_major_subjects.py --limit 5 --delay 0.1 --output data\careernet_major_subjects_sample.csv
```

## Graph RAG 반영 방식

`GraphRAGEngine`은 기본 실행 시 `data/careernet_major_subjects.csv`가 있으면 자동으로 읽는다.

- 일반 선택 과목: `권장`, 우선순위 40
- 진로 선택 과목: `핵심`, 우선순위 20
- 융합 선택 과목: `권장`, 우선순위 60

커리어넷 학과명은 기존 로드맵 DB의 학과명과 `학과`, `전공`, `계열`, `학부`, `과`, `부` 접미사를 제거한 기준으로 먼저 매칭한다. 정확히 매칭되지 않는 학과는 커리어넷 학과명으로 새 그래프 노드를 만든다.

## 주의사항

커리어넷 자료는 학과별 일반 권장 과목에 가깝고, 특정 고등학교의 실제 개설 여부를 보장하지 않는다. 챗봇 응답에서는 학교 편제표 안에서 확인되는 과목과, 목표 학과 근거에는 있지만 편제표에서 바로 확인되지 않는 과목을 분리해 안내해야 한다.
