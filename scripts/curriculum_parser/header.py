# -*- coding: utf-8 -*-
"""헤더 3행 / 학기라벨 5행 탐지 (v1 로직 이관, 밀집 그리드 기반).

grid: list[list]  — grid[r][c] 는 0-based (r=시트행-1, c=시트열-1).
"""
from __future__ import annotations

import re


def _cell(row: list, idx: int):
    return row[idx] if 0 <= idx < len(row) else None


def find_header(grid: list) -> int | None:
    """헤더 행(0-based)을 찾는다.

    허용 형태: '세부과목' 포함 / '구분'+'교과' / '교과'+'과목'+'학년'(구분열 없는
    출력본형, 예: 울산고운고).
    """
    for i, row in enumerate(grid[:12]):
        joined = " ".join(str(c) for c in row if c)
        if "세부과목" in joined or ("구분" in joined and "교과" in joined):
            return i
        if "교과" in joined and "과목" in joined and "학년" in joined:
            return i
    return None


def build_column_map(grid: list, header_idx: int) -> dict | None:
    """헤더 블록에서 열 위치를 해석한다.

    반환 dict 키: section/group/type/name/base/run/semesters/data_start
    semesters: {"1-1": col_idx, ...}
    실패 시 None.
    """
    header = grid[header_idx]
    col = {"section": None, "group": None, "type": None, "name": None,
           "base": None, "run": None}
    grade_anchor: dict[int, int] = {}

    for idx, cell in enumerate(header):
        text = re.sub(r"\s", "", str(cell)) if cell else ""
        if not text:
            continue
        if text == "구분" and col["section"] is None:
            col["section"] = idx
        elif "교과" in text and col["group"] is None:
            col["group"] = idx
        elif "세부과목" in text and col["type"] is None:
            col["type"], col["name"] = idx, idx + 1
        elif "과목유형" in text:
            col["type"] = idx
        elif text == "과목" and col["name"] is None:
            col["name"] = idx
        elif col["name"] is None and (text.startswith("세부") or text in ("과목명", "교과목")):
            # '세부교과목'(과목유형 열이 별도로 있는 편제표) 등을 과목명 열로 인식
            col["name"] = idx
        elif ("기준" in text or "기본" in text) and "학점" in text.replace("단위", "학점"):
            col["base"] = idx
        elif ("운영" in text or "편성" in text) and col["run"] is None:
            col["run"] = idx
        else:
            match = re.match(r"([123])학년", text)
            if match:
                # 병합 복원으로 '1학년'이 여러 열에 전파될 수 있으므로 첫 열만 채택
                grade_anchor.setdefault(int(match.group(1)), idx)

    if col["name"] is None or col["group"] is None or not grade_anchor:
        return None

    # 학기 행: 헤더 아래 3행 안에서 '1학기/2학기'가 4개 이상인 행
    semester_cols: dict[str, int] = {}
    for offset in range(1, 4):
        if header_idx + offset >= len(grid):
            break
        row = grid[header_idx + offset]
        hits = [
            (idx, re.sub(r"\s", "", str(c)))
            for idx, c in enumerate(row)
            if c and re.fullmatch(r"[12]학기", re.sub(r"\s", "", str(c)))
        ]
        if len(hits) >= 4:
            for grade, start in grade_anchor.items():
                for idx, text in hits:
                    if start <= idx <= start + 1:
                        semester_cols[f"{grade}-{text[0]}"] = idx
            col["data_start"] = header_idx + offset + 1
            break

    if not semester_cols:
        # 학기 행이 없으면 학년 앵커 기준 2열씩 배정
        for grade, start in grade_anchor.items():
            semester_cols[f"{grade}-1"] = start
            semester_cols[f"{grade}-2"] = start + 1
        col["data_start"] = header_idx + 1

    col["semesters"] = semester_cols
    return col
