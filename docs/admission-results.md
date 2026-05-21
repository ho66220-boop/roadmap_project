# 대학어디가 입시결과 수집 메모

## 확인한 접근 방식

대학어디가 검색 페이지는 내부적으로 공개 JSON 엔드포인트를 호출합니다.

- 검색 엔드포인트: `https://www.adiga.kr/man/sch/majorInfo2.do`
- 입시결과 팝업: `https://www.adiga.kr/ucp/cls/uni/classUnivAdmssPopup.do`

검색 결과에서 `UNV_CD`, `RCU_CD`, 학과명, 대학명을 얻고, 입시결과 팝업에 `searchSyr`, `unvCd`, `ruCd`를 전달하면 전년도 입시결과 표를 확인할 수 있습니다.

## 생성한 로컬 산출물

아래 CSV는 로컬 검증용 산출물이며, 저작권 정책을 고려해 기본적으로 Git 추적에서 제외했습니다.

| 파일 | 내용 |
| --- | --- |
| `data/admission_results_sample.csv` | 엑셀 전공 목록 앞부분 기준 샘플. 178개 전형 행 |
| `data/admission_results_core.csv` | 컴퓨터공학, 생명공학, 기계공학 중심 샘플. 543개 전형 행 |

## CSV 스키마

| 필드 | 설명 |
| --- | --- |
| result_year | 입시결과 기준 학년도 |
| university | 대학명 |
| major | 모집단위명 |
| area | 지역 |
| admission_period | 수시/정시 등 모집 시기 |
| admission_type | 학생부위주(교과), 학생부위주(종합), 수능위주 등 |
| track_name | 세부 전형명 |
| recruit_count | 모집인원 |
| additional_pass_count | 충원인원 |
| competition_rate | 경쟁률 |
| score_50 | 학생부 환산점수 50% cut |
| score_70 | 학생부 환산점수 70% cut |
| grade_50 | 학생부 환산등급 50% cut |
| grade_70 | 학생부 환산등급 70% cut |
| source_url | 대학어디가 입시결과 팝업 출처 URL |
| collected_at | 수집 시각 |

## 실행 예시

```powershell
python scripts\collect_adiga_admission_results.py `
  --terms "컴퓨터공학,생명공학,기계공학,AI 인공지능" `
  --search-limit 80 `
  --max-popups 80 `
  --delay 0.5 `
  --output data\admission_results_core.csv
```

## 주의 사항

- 대학어디가 입시결과는 대학이 제공한 참고용 자료이며, 대학별 산출 방식이 다릅니다.
- 50%/70% cut은 지원 가능성을 보장하지 않습니다.
- 모집단위별 선발 인원이 매우 적거나 대학이 제공하지 않은 항목은 `-` 또는 공란으로 표시될 수 있습니다.
- 공개 GitHub에는 전체 수집 CSV를 올리기보다, 수집 스크립트와 스키마, 출처 URL 저장 방식만 공개하는 편이 안전합니다.
