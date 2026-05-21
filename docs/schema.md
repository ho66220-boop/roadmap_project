# 표준 데이터 스키마 초안

## subjects

| 필드 | 설명 |
| --- | --- |
| subject_code | 공식 과목 코드 |
| official_name | 공식 과목명 |
| subject_group | 교과군 |
| subject_type | 공통/일반/진로/융합 등 |
| default_credit | 기본 학점 |
| aliases | 원본 표기와 별칭 목록 |

## curriculum_slots

| 필드 | 설명 |
| --- | --- |
| slot_id | 학교별 선택 슬롯 식별자 |
| school_name | 학교명 |
| cohort_year | 입학생 기준 연도 |
| grade | 학년 |
| semester | 학기 |
| choice_type | 필수, 택N, 선택수불명 등 |
| min_select | 최소 선택 수 |
| max_select | 최대 선택 수 |
| source_file | 원본 편제표 파일명 |
| confidence | 해석 신뢰도 |

## curriculum_slot_subjects

| 필드 | 설명 |
| --- | --- |
| slot_id | 선택 슬롯 식별자 |
| subject_order | 슬롯 내 과목 순서 |
| raw_subject_name | 원본 과목명 |
| normalized_subject_name | 정규화 과목명 |
| subject_code | 공식 과목 코드 |
| match_status | 확정/검토필요/미매칭 |

## university_major_tracks

| 필드 | 설명 |
| --- | --- |
| university_name | 대학명 |
| major_name | 모집단위명 |
| normalized_major | 표준 학과명 |
| subject_code | 추천 과목 코드 |
| recommend_type | 핵심/권장/참고 |
| priority | 우선순위 |
| evidence | 시행계획 또는 과목 선택 자료 근거 |

## student_profiles

| 필드 | 설명 |
| --- | --- |
| profile_id | 상담 세션 식별자 |
| school_name | 학교명 |
| grade | 현재 학년 |
| target_major | 목표 학과 |
| grade_band | 내신 구간 |
| taken_subjects | 이수 과목 코드 목록 |
| preferred_region | 희망 지역 |

## recommendation_results

| 필드 | 설명 |
| --- | --- |
| result_id | 추천 결과 식별자 |
| profile_id | 상담 세션 식별자 |
| missing_core_subjects | 부족한 핵심 과목 |
| available_subjects | 학교 편제표 안에서 선택 가능한 추천 과목 |
| unavailable_subjects | 목표 학과에는 중요하지만 해당 학교 편제표에 없는 과목 |
| evidence_nodes | 답변 근거 그래프 노드 |
| caution | 데이터 한계 및 최종 확인 안내 |
