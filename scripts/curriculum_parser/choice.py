# -*- coding: utf-8 -*-
"""택N 그룹 전개 · 선택군ID 부여 (v2 핵심 신규).

v1 의 암묵 상속(block_semester)을 1급 그룹 객체로 승격한다.

그룹 경계(설계):
  ① 학기열 마커(택N) → 그 열이 그룹 학기. 마커 행부터, 같은 학기열에 학점이
     이어지는 연속 행이 멤버. 다음 택N/구분 변경/소계에서 종료. 멤버 중 자체
     학점 없는 행은 그룹 학기에 운영학점(없으면 기준학점)으로 채운다.
  ② 과목명 접미 "(택N)" → 같은 행 학기열이 그룹 학기(단일 행 그룹).
  ③ 기타열 마커 → take_n 만 취하고 멤버 각자 학기 유지(단일 행).

선택군ID: {학교}-{학기키}-{연번:02d}  예) 신정고-2-1-01
"""
from __future__ import annotations

from dataclasses import dataclass, field

from .model import ChoiceGroup


@dataclass
class RawRow:
    """파싱 중간 표현(그리드 데이터 행 1개)."""
    idx: int                        # 그리드 행 인덱스(0-based)
    section: str = ""
    group: str = ""
    subtype: str = ""
    name_raw: str = ""
    base: float = 0.0
    run: float = 0.0
    sem_num: dict = field(default_factory=dict)      # semkey -> float(>0) 직접 학점
    sem_marker: dict = field(default_factory=dict)   # semkey -> take_n(int) (택N)
    sem_nonanchor: set = field(default_factory=set)  # 병합 비앵커인 학기셀 키(전파된 마커)
    cross_grade: str = ""            # 교차집중 학년("1"/"2"/"3")
    name_suffix_take: int = 0        # 과목명 접미 (택N)
    other_take: int = 0              # 기타열 택N
    nonanchor_name: bool = False     # 병합 비앵커 과목명(중복 방지)
    is_skip: bool = False            # 소계/합계/창체 등
    # 배정 결과
    choice_id: str = ""
    take_n: int = 0
    credits: dict = field(default_factory=dict)      # 최종 학기 학점

    @property
    def emittable(self) -> bool:
        return bool(self.name_raw) and not self.nonanchor_name and not self.is_skip


def _fallback_credit(row: "RawRow") -> float:
    return row.run or row.base or 0.0


def assign_choice_groups(raw_rows: list, semester_keys: list, school: str,
                         seq_start: int = 1) -> tuple[list, int]:
    """raw_rows 를 순회하며 택N 그룹을 배정한다.

    각 RawRow 의 credits 를 확정하고 choice_id/take_n 을 채운다.
    반환: (groups, next_seq)
    """
    groups: list = []
    seq = seq_start
    n = len(raw_rows)

    # 1) 직접 학점 반영 + 교차집중(양 학기) 처리
    for row in raw_rows:
        for k in semester_keys:
            v = row.sem_num.get(k, 0.0)
            if v > 0:
                row.credits[k] = v
        if row.cross_grade:
            fb = _fallback_credit(row)
            row.credits.setdefault(f"{row.cross_grade}-1", fb)
            row.credits.setdefault(f"{row.cross_grade}-2", fb)

    # 2) 학기열 마커 그룹(①) — 단일 패스, 마커 행부터 다음 마커/구분변경/소계까지
    #    (설계 경계 ③). 자체 학점 없는 멤버(마커 행·부가 선택지)는 그룹 학기로 상속.
    def _anchor_marker(r):
        for k in semester_keys:
            take = r.sem_marker.get(k, 0)
            if take and k not in r.sem_nonanchor:
                return k, take
        return None, 0

    i = 0
    while i < n:
        row = raw_rows[i]
        if row.is_skip:
            i += 1
            continue
        semkey, take = _anchor_marker(row)
        if not semkey:
            i += 1
            continue
        start = i
        end = i
        j = i + 1
        while j < n:
            nxt = raw_rows[j]
            if nxt.is_skip:
                break
            if nxt.section and nxt.section != row.section:
                break
            if _anchor_marker(nxt)[0]:              # 다음 택N 앵커 → 새 그룹
                break
            end = j
            j += 1

        choice_id = f"{school}-{semkey}-{seq:02d}"
        seq += 1
        grp = ChoiceGroup(choice_id=choice_id, take_n=take, semester=semkey)
        for k in range(start, end + 1):
            m = raw_rows[k]
            if m.choice_id or not m.emittable:
                continue
            m.choice_id = choice_id
            m.take_n = take
            if not any(v > 0 for v in m.credits.values()):   # 자체 학점 없는 행 상속
                m.credits[semkey] = _fallback_credit(m)
            grp.members.append(m)
        groups.append(grp)
        i = end + 1

    # 3) 과목명 접미(②) & 기타열(③)
    for row in raw_rows:
        if row.choice_id or not row.emittable:
            continue
        take = row.name_suffix_take or row.other_take
        if not take:
            continue
        semkey = next((k for k in semester_keys if row.credits.get(k, 0.0) > 0), "")
        if row.name_suffix_take and semkey and row.credits.get(semkey, 0.0) <= 0:
            row.credits[semkey] = _fallback_credit(row)
        tag = semkey if semkey else "x"
        choice_id = f"{school}-{tag}-{seq:02d}"
        seq += 1
        row.choice_id = choice_id
        row.take_n = take
        groups.append(ChoiceGroup(choice_id=choice_id, take_n=take,
                                  semester=semkey, members=[row]))

    return groups, seq
