# -*- coding: utf-8 -*-
"""공용 데이터 모델과 20열 출력 스키마.

앞 18열(인덱스 0~17)은 v1 산출물과 바이트 단위로 동일한 의미다. 소비자
(graph_rag._load_curriculum=row[0..15], Code.gs=row[10+i], export_for_gsheet=전열)가
인덱스로 읽으므로 순서를 절대 바꾸지 말 것. 말미 2열[선택군ID, 택N]만 신규.
"""
from __future__ import annotations

from dataclasses import dataclass, field

COHORT_YEAR = 2026

# 학기 키 — 출력 열 10~15에 이 순서로 대응
SEMESTER_KEYS = ["1-1", "1-2", "2-1", "2-2", "3-1", "3-2"]

# 20열 출력 헤더. 0~17은 v1과 동일, 18~19가 신규.
HEADERS = [
    "학교", "학년도", "구분", "교과군", "과목유형", "과목명", "공식과목명", "매칭",
    "기준학점", "운영학점", *SEMESTER_KEYS, "원본파일", "원본시트",
    "선택군ID", "택N",
]


@dataclass
class ParsedRow:
    """편제표 과목 한 줄(분리·정제·매칭·그룹 배정 완료)."""
    school: str
    section: str            # 구분(학교지정/학생선택 등, 정규화됨)
    subject_group: str      # 교과(군)
    subject_type: str       # 공통/일반/진로/융합 등
    raw_name: str           # 원본 과목명(마커 포함)
    official_name: str      # 과목마스터 매칭 결과(미매칭 시 "")
    base_credit: float
    run_credit: float
    credits: dict           # {"1-1": 4.0, ...} 값 없는 학기는 미포함/0
    source_file: str
    source_sheet: str
    choice_id: str = ""     # 선택군ID(택N 그룹). 미해당 시 ""
    take_n: int = 0         # 택N의 N. 미해당 시 0
    cross: bool = False     # 교차이수(↔) 여부

    def to_output_row(self) -> list:
        """20열 리스트로 직렬화(HEADERS 순서)."""
        return [
            self.school, COHORT_YEAR, self.section, self.subject_group,
            self.subject_type, self.raw_name,
            self.official_name or self.raw_name, "O" if self.official_name else "X",
            self.base_credit or None, self.run_credit or None,
            *[self.credits.get(k) or None for k in SEMESTER_KEYS],
            self.source_file, self.source_sheet,
            self.choice_id or None, self.take_n or None,
        ]


@dataclass
class CleanLog:
    """정제/병합분배 감사 로그 한 줄."""
    school: str = ""
    coord: str = ""         # 셀 좌표 또는 병합범위(예: I34, I34:I38)
    raw: str = ""           # 원본값
    cleaned: str = ""       # 정제값
    flag: str = ""          # range_value/paren_note/multi_credit/glyph/nfc/stack_mismatch 등

    def as_row(self) -> list:
        return [self.school, self.coord, self.raw, self.cleaned, self.flag]


@dataclass
class ChoiceGroup:
    """택N 선택군."""
    choice_id: str
    take_n: int
    semester: str           # 그룹 귀속 학기 키(기타열 마커면 "")
    members: list = field(default_factory=list)  # ParsedRow 리스트(참조)


@dataclass
class SheetResult:
    """단일 시트 파싱 결과."""
    rows: list = field(default_factory=list)        # ParsedRow
    logs: list = field(default_factory=list)        # CleanLog
    groups: list = field(default_factory=list)      # ChoiceGroup
    meta_notes: list = field(default_factory=list)  # 하단 안내문 등 보존 텍스트
    issues: list = field(default_factory=list)      # 파싱 이슈 메시지
