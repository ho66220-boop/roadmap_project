# GitHub 업로드 절차

`roadmap_project` 폴더만 하나의 포트폴리오 저장소로 올리는 것을 권장합니다. 상위 `Road` 폴더에는 원본 PDF/HWP/XLSX가 많아서 저작권, 용량, 개인정보 검토 없이 통째로 공개하기에는 부담이 큽니다.

## 1. 저장소 초기화

```powershell
cd C:\Users\Administrator\project\Road\roadmap_project
git init
git add .
git commit -m "Add roadmap counseling MVP portfolio"
```

## 2. GitHub 원격 저장소 연결

GitHub에서 빈 저장소를 만든 뒤 아래 명령을 실행합니다.

```powershell
git branch -M main
git remote add origin https://github.com/<계정명>/<저장소명>.git
git push -u origin main
```

## 3. 공개 전 확인

- 원본 시행계획 PDF/HWP 파일이 포함되지 않았는지 확인
- 고교 편제표 원본 XLSX 파일이 포함되지 않았는지 확인
- 학생 개인정보, 학교 내부 메모, 상담 기록, API 키가 포함되지 않았는지 확인
- 내신컷 등 출처 확인과 재배포 검토가 필요한 데이터는 샘플 또는 스크립트만 공개했는지 확인
- 공개 저장소에는 `data/sample_roadmap.xlsx`, 공개 가능한 CSV, 구조 설명 문서만 포함했는지 확인
- README에서 현재 구현 범위와 향후 구현 계획이 명확히 분리되어 있는지 확인
- README의 한계와 윤리 문구를 유지했는지 확인
- 샘플 데이터 기준 `python server.py` 실행이 가능한지 확인

## 4. 포트폴리오 설명 문장

> 울산 지역 고교 편제표와 학과별 추천 과목 DB를 연결해, 고교학점제 환경에서 학생별 과목 선택 로드맵을 제안하는 진로 상담 MVP를 구현했습니다. 현재는 엑셀 기반 데이터 로딩과 규칙 기반 추천, 웹 UI 시연까지 구현했으며, 향후 Neo4j/NetworkX와 LLM 답변 체인을 연결하는 Graph RAG 구조로 확장할 계획입니다.
