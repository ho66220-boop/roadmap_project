# 울산 고교 과목 로드맵 기반 진로 상담 MVP

2026학년도 울산 지역 고교 편제표와 학과별 추천 과목 데이터를 연결해, 예비 고1~고3 학생의 과목 선택과 대학/학과 탐색을 돕는 진로 상담 MVP입니다.

**현재 단계는 외부 LLM이나 그래프 DB 없이, 엑셀/CSV 데이터를 메모리에 올려 규칙 기반으로 과목을 매칭·추천하는 프로토타입입니다.** Neo4j/NetworkX 그래프 저장소와 LLM 답변 체인을 붙이는 Graph RAG 구조로 확장할 수 있도록 노드/엣지 응답 형태를 미리 맞춰 두었습니다(자세한 발전 방향은 [향후 발전 방향](#향후-발전-방향) 참고).

이 프로젝트는 팀 프로젝트 주제를 바탕으로 하지만, 데이터 구조화와 MVP 설계, 검색/추천 흐름, 웹앱 프로토타입 구상은 개인 개발 가능성을 검증하기 위해 별도 정리했습니다.

## 문제 정의

고교학점제와 2028학년도 대입 개편 환경에서는 학생이 학교별 편제표 안에서 어떤 과목을 선택해야 목표 학과와 연결되는지 판단하기 어렵습니다. 특히 같은 학과라도 대학별 권장 과목, 핵심 과목, 전형 반영 방식이 다르고, 학교마다 실제 개설 과목과 선택 슬롯이 다릅니다.

이 프로젝트는 다음 질문에 답하는 것을 목표로 합니다.

- 고1·고2: 현재까지 선택한 과목과 목표 대학/학과를 기준으로 남은 학기 과목 선택 로드맵을 제안할 수 있는가?
- 고3: 학교, 내신, 수강 편제표, 목표 학과를 기준으로 지원 가능한 대학/학과와 유사 학과 대안을 제안할 수 있는가?
- 상담 결과가 단순 LLM 답변이 아니라 편제표, 시행계획, 과목 DB에 근거한 설명으로 제공될 수 있는가?

## 지금 할 수 있는 것 (현재 구현)

엑셀/CSV 데이터를 읽어 메모리 그래프를 만들고, 규칙 기반으로 과목을 매칭·추천합니다. **외부 LLM 호출이나 별도 DB 없이 동작합니다.**

- 엑셀 로드맵 DB 로딩 (`openpyxl`)
- 학교별 개설 과목 추출 (편제표 시트의 학기별 이수 단위 기준)
- 학과별 핵심/권장/참고 과목 매칭 및 과목명 정규화(별칭, 로마숫자 변환, 인코딩 깨짐 복구)
- 학생 조건(학교·학년·이수 과목·목표 학과) 기반 과목 추천 — 충족/부족/대체 분류
- 학기별 과목 로드맵 생성
- 내신 등급 입력 시 입시결과 샘플 CSV와 비교해 안정/적정/상향 밴드 분류 (※ 공개 저장소에는 입시결과 CSV가 포함되지 않아 비활성, 아래 [데이터와 공개 범위](#데이터와-공개-범위) 참고)
- 커리어넷 학과별 2022 개정 선택과목 CSV 연동
- 규칙 기반 상담 문장 생성 (LLM 미사용, 계산 결과를 템플릿으로 서술)
- 정적 HTML + `fetch` 기반 상담 리포트 웹 UI

> 참고: 추천 점수·필터·밴드 분류는 모두 추천 엔진이 계산합니다. 상담 문장 생성기는 그 결과를 문장으로 옮기기만 하므로, 추후 LLM을 붙여도 점수를 임의로 만들어내지 않는 구조입니다.

## 아직 안 되는 것 / 알려진 한계

README가 과대표현이 되지 않도록 현재 구현의 경계를 명시합니다.

- **Graph RAG 미구현**: 실제로는 Neo4j/NetworkX나 벡터 인덱스가 아니라 메모리 내 `dict` 그래프 + 규칙 엔진입니다.
- **LLM 답변 체인 없음**: 답변은 규칙 기반 템플릿으로 생성됩니다.
- **선택 슬롯(택N) 규칙 미반영**: 현재 로더는 편제표 시트의 학기 컬럼만 읽고, 학교별 선택 묶음(택N·교차선택) 규칙은 사용하지 않습니다. `docs/schema.md`·`docs/architecture.md`의 슬롯 스키마는 향후 목표 설계입니다.
- **테스트 범위가 좁음**: 현재 자동 테스트는 학과명 매칭 위주입니다.
- **입시결과 데이터 비공개**: 내신 기반 대학 후보 추천은 별도 CSV가 있을 때만 활성화됩니다.

## 데이터와 공개 범위

| 구분 | 내용 |
| --- | --- |
| 과목 로드맵 DB | `울산고교_과목로드맵_수정.xlsx` 기반. 학교 편제표, 대학트랙, 학과트랙, 과목 별칭, 공식 과목명, 로드맵 집계 시트 포함 |
| 학교 편제표 | 울산 지역 고교 2026학년도 신입생/전학년 교육과정 편제표 79개 |
| 대학 시행계획 | 2028학년도 대학입학전형 시행계획 PDF/HWP 82개 |
| 과목 선택 자료 | 고려대, 연세대, 이화여대, 경북대 과목 선택 자료 4개 |
| 기존 웹앱 | Google Sheets + Apps Script + HTML 기반 과목 로드맵 조회 웹앱 코드 |

원본 XLSX/PDF/HWP 데이터는 저작권 및 개인정보 검토 문제로 공개하지 않습니다. 공개 저장소에는 샘플 엑셀(`data/sample_roadmap.xlsx`), 공개 가능한 커리어넷 선택과목 CSV, 구조 설명, 변환/수집 스크립트 중심으로 제공합니다.

> **입시결과 CSV는 공개 저장소에 포함되지 않습니다.** `.gitignore`가 `data/admission_results_*.csv`를 제외하므로, 공개 저장소를 그대로 클론하면 내신 기반 대학 후보 추천은 비활성 상태로 동작합니다. 이 기능을 시연하려면 별도로 준비한 입시결과 CSV를 `data/admission_results_core.csv`에 두어야 합니다.

자세한 데이터 현황은 [docs/data-inventory.md](docs/data-inventory.md)에 정리했습니다.

## 실행 방법

### 요구 사항

- **Python 3.10 이상** (코드가 `X | None`, `tuple[str, ...]` 등 최신 타입 문법을 사용합니다)
- 의존성: `openpyxl` 하나뿐입니다. 웹 서버는 표준 라이브러리(`http.server`)로 동작하며 외부 프레임워크가 필요 없습니다.

```powershell
pip install -r requirements.txt
```

### 공개 샘플 데이터로 실행

공개 저장소에는 원본 로드맵 엑셀 대신 `data/sample_roadmap.xlsx`가 포함되어 있습니다.

```powershell
python server.py
```

브라우저에서 `http://localhost:8000` 접속.

샘플 엑셀을 다시 만들고 싶다면:

```powershell
python scripts\create_sample_workbook.py
```

### 원본 엑셀 데이터로 실행

개인 로컬 환경에서 원본 `울산고교_과목로드맵_수정.xlsx`를 사용하는 경우 경로를 직접 넘깁니다.

```powershell
python server.py --workbook ..\울산고교_과목로드맵_수정.xlsx
```

원본 파일이 없으면 서버는 “원본 로드맵 엑셀 파일을 찾을 수 없습니다. README의 데이터 준비 방법을 확인하세요.”라는 안내를 출력하고 종료합니다.

### 테스트 실행

```powershell
python -m unittest discover -s tests
```

### 시연 화면

![진로 상담 챗봇 시연 화면](assets/demo-overview.png)

학교, 학년, 내신, 목표 학과, 이수 과목을 입력해 과목 로드맵과 대학 후보 상담 리포트를 생성하는 화면입니다.

## API

`server.py`가 엑셀 로드맵 DB를 읽어 메모리 그래프를 만들고, `app/index.html`이 아래 엔드포인트를 호출합니다.

### `GET /api/meta`

로딩된 학교/학과 목록과 데이터 통계를 반환합니다. UI가 시작 시 호출해 입력 후보를 채웁니다.

### `POST /api/chat`

상담 질의를 받아 추천 리포트를 반환합니다. `message`(자유 입력)와 `profile`(구조화 입력)을 받으며, `profile`이 없으면 `message`에서 값을 추출합니다.

`profile` 필드:

| 필드 | 설명 |
| --- | --- |
| `school` | 학교명 |
| `grade` | 학년(1~3) |
| `major` | 목표 학과/계열 |
| `taken` | 이수 과목 (쉼표 구분 문자열 또는 배열) |
| `school_grade` (별칭 `gpa`) | 내신 등급. 입력 시 입시결과 밴드 분류 활성화 |
| `admission_focus` | `교과` / `종합` / 전체 |

```powershell
Invoke-WebRequest -Uri http://localhost:8000/api/chat `
  -Method POST `
  -ContentType "application/json; charset=utf-8" `
  -Body '{"message":"울산고 고1이고 컴퓨터공학 희망인데 과목 추천해줘","profile":{"school":"울산고","grade":1,"major":"컴퓨터공학","taken":"공통수학1, 공통수학2, 통합과학1"}}'
```

## 저장소 구조

```text
roadmap_project/
  app/                  # 상담 리포트 웹 UI (정적 HTML 단일 파일)
  roadmap_rag/          # 엑셀 기반 메모리 그래프 로더와 규칙 기반 추천 엔진
  server.py             # 표준 라이브러리 기반 챗봇 API 서버
  data/                 # 공개 가능한 샘플/메타 데이터 (sample_roadmap.xlsx, careernet CSV 등)
  docs/                 # 설계 문서
  scripts/              # 데이터 점검 및 수집/변환 스크립트
  tests/                # 단위 테스트
  assets/               # 시연 이미지
```

## 문서

- [docs/architecture.md](docs/architecture.md) — 목표 Graph RAG 아키텍처와 검색 흐름
- [docs/data-inventory.md](docs/data-inventory.md) — 원자료 인벤토리와 공개 범위
- [docs/schema.md](docs/schema.md) — 표준 데이터 스키마 초안(목표)
- [docs/mvp-plan.md](docs/mvp-plan.md) — 단계별 구현 계획
- [docs/admission-results.md](docs/admission-results.md) — 입시결과 데이터 설명
- [docs/careernet-subjects.md](docs/careernet-subjects.md) — 커리어넷 선택과목 데이터 설명
- [docs/github-publish.md](docs/github-publish.md) — GitHub 업로드 절차

> `docs/architecture.md`와 `docs/schema.md`는 **앞으로 도달하고자 하는 목표 설계**이며, 일부(그래프 DB, 선택 슬롯 스키마, LLM 계층)는 아직 구현되지 않았습니다. 현재 구현 경계는 위의 [아직 안 되는 것 / 알려진 한계](#아직-안-되는-것--알려진-한계)를 기준으로 봐 주세요.

## 향후 발전 방향

현재 규칙 기반 MVP를 기반으로 다음 순서로 발전시킬 계획입니다.

1. **데이터 정합성 강화** — 선택 슬롯(택N) 규칙 반영, 과목명 정규화·매칭률 개선, 변환 실패 케이스 정리
2. **추천 엔진 검증 확대** — 추천/학기 배치/내신 밴드 로직에 대한 테스트 보강
3. **그래프 저장소 도입** — 메모리 `dict` 그래프를 Neo4j 또는 NetworkX로 이전 (노드/엣지 응답 구조는 이미 호환되도록 유지)
4. **API 현대화** — `http.server`에서 FastAPI로 전환
5. **LLM 답변 체인 연결** — 추천 계산은 현재 엔진이 담당하고, LLM은 질문 의도 파악·누락 입력 확인·상담 문장 재구성·근거 요약에만 사용 (점수 생성은 LLM이 하지 않음)
6. **내신컷 데이터 확장** — 최근 3개년 대학/학과별 내신컷을 확보해 지원권/상향/안정 후보 분류 정확도 향상

향후 그래프는 다음과 같은 구조를 목표로 합니다.

```text
School -> CurriculumSlot -> Subject
Subject -> MajorTrack -> Major
University -> AdmissionPlan -> Major
Major -> SimilarMajor -> Major
StudentProfile -> TakenSubject / TargetMajor / GradeBand
```

## 한계와 윤리

본 서비스는 입시 결과를 보장하지 않습니다. 대학별 시행계획은 변경될 수 있고, 내신컷은 연도·전형·모집단위에 따라 달라집니다. 따라서 결과 화면에는 항상 데이터 기준일, 사용한 근거, 불확실성, 최종 확인 필요 문구를 함께 표시해야 합니다.

## 라이선스 / 데이터 사용

코드와 원본 데이터의 라이선스는 별도 정합니다(현재 미정). 원본 편제표·시행계획·로드맵 엑셀은 저작권 및 개인정보 검토 대상이므로 공개 저장소에 포함하지 않으며, 재배포하지 마세요. GitHub 업로드 절차는 [docs/github-publish.md](docs/github-publish.md)를 참고하세요.
