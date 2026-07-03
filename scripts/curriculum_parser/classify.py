# -*- coding: utf-8 -*-
"""행 분류: subject / meta / legend / cross / skip / empty.

설계:
  - 과목명이 서술문(20자+ & 이수|시기별|경우 패턴) & 학점 없음 → meta(본표 제외+보존)
  - '↔' 포함 → cross(양쪽 개설)
  - 범례/표기법 행 → legend(스킵)
  - 소계/합계/창체 → skip
  - 그 외 → subject
"""
from __future__ import annotations

import re

from .cells import is_legend, is_skip_name

META_PATTERN = re.compile(
    r"이수\s*시기|시기별|경우에|운영할\s*수\s*있다|희망하는\s*학생|별도로\s*계획|참고\s*사항"
)


def classify_row(name_raw: str, base: float, run: float, credits_all_zero: bool) -> str:
    name = (name_raw or "").strip()
    if not name:
        return "empty"
    if is_skip_name(name):
        return "skip"
    if is_legend(name):
        return "legend"
    if "↔" in name:
        return "cross"
    if (len(name) >= 20 and META_PATTERN.search(name)
            and credits_all_zero and not base and not run):
        return "meta"
    return "subject"
