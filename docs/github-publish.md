# GitHub 업로드 절차

`roadmap_project` 폴더만 하나의 포트폴리오 저장소로 올리는 것을 권장합니다. 상위 `Road` 폴더에는 원본 PDF/HWP/XLSX가 많아서 저작권, 용량, 개인정보 검토 없이 통째로 공개하기에는 부담이 큽니다.

## 1. 저장소 초기화

```powershell
cd C:\Users\Administrator\project\Road\roadmap_project
git init
git add .
git commit -m "Add roadmap Graph RAG portfolio"
```

## 2. GitHub 원격 저장소 연결

GitHub에서 빈 저장소를 만든 뒤 아래 명령을 실행합니다.

```powershell
git branch -M main
git remote add origin https://github.com/<계정명>/<저장소명>.git
git push -u origin main
```

## 3. 공개 전 확인

- 원본 시행계획 PDF/HWP와 고교 편제표 원본을 그대로 올리지 않았는지 확인
- 내신컷 등 민감하거나 출처 확인이 필요한 데이터는 공개 범위를 검토
- README의 한계와 윤리 문구를 유지
- 정적 데모가 `app/index.html`에서 열리는지 확인

## 4. 포트폴리오 설명 문장

> 울산 지역 고교 편제표와 2028학년도 대학입학전형 시행계획을 연결해, 고교학점제 환경에서 학생별 과목 선택 로드맵과 대학/학과 탐색을 지원하는 Graph RAG 기반 진로 상담 챗봇을 설계했습니다. 기존 Google Sheets 웹앱 데이터를 분석해 그래프 스키마, 검색 흐름, MVP UI를 개인 주도로 구체화했습니다.
