# -*- coding: utf-8 -*-
"""셀 값 정제기 · 과목명 분리 · 마커/범례 · 유니코드 정규화.

graph_rag.repair_text/norm 을 재사용(복제 금지). NFC 정규화만 이 층에서 추가.
"""
from __future__ import annotations

import re
import sys
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from roadmap_rag.graph_rag import norm, repair_text  # noqa: E402,F401  (재사용)

from .model import CleanLog  # noqa: E402

# 학교별 각주 마커(공동교육과정·소인수·순증 표시 등) — 매칭 전에 제거
# ▪▫◾◽(작은 사각형)은 성신고 등이 과목명 접미 마커로 사용 → 반드시 포함
FOOTNOTE_MARKERS = re.compile(r"[■□◆◇▲△▶▷●○◦∘▪▫◾◽★☆†‡*]+")

# 편제표 하단/범례 행("☆고시외 과목", "표시 예 : *특목고" 등) — 과목이 아니므로 제외
LEGEND_ROW = re.compile(r"^[\s■□◆◇▲△●○★☆†‡*:]+|표시\s*예|표기\s*방법|표기법")

# 과목명 뒤에 붙은 선택 규칙 표기("중국어(택1)") — 매칭 전에 제거(그룹 정보는 별도 보존)
CHOICE_SUFFIX = re.compile(r"\(\s*택\s*\d+\s*\)\s*$")

# 소계/합계/창체 등 과목 아닌 행
SKIP_NAME_PATTERNS = re.compile(
    r"^(소\s*계|합\s*계|총|계|이수학점\s*소계)$|창의적|체험활동|자율.?자치|동아리|"
    r"봉사활동|진로활동|학기별\s*총|이수\s*학점\s*소계|필수\s*이수"
)

# '택N' 감지(어디에 있든)
TAKE_N = re.compile(r"택\s*(\d+)")

# 교차집중 마커(학기열에 위치, 해당 학년 두 학기에 배정)
CROSS_CONCENTRATE = re.compile(r"교\s*차\s*집\s*중")

# 공동교육과정 등에서 학기열에 서술로 박힌 '(1학년 2학기, 2학년 2학기)' → 학기키 추출
_JOINT_SEM = re.compile(r"([123])\s*학년\s*([12])\s*학기")


def parse_joint_semesters(text) -> list[str]:
    """'공동교육과정 (2학년 1학기)' 같은 서술에서 학기키(예: '2-1')를 뽑는다."""
    if not text:
        return []
    return [f"{g}-{s}" for g, s in _JOINT_SEM.findall(nfc(str(text)))]

# 학교 표기와 공식 과목명이 다른 알려진 별칭
NAME_ALIASES = {
    "동아시아사 역사 기행": "동아시아 역사 기행",
}

# 앞머리 숫자 토큰(옵션 괄호 허용): "4", "(70)", "29", "3.5"
_LEAD_NUM = re.compile(r"^\(?\s*(\d+(?:\.\d+)?)")
# 범위값 "29~30" / "29-30"(하이픈은 과목명 혼동 위험이 커서 물결만 처리)
_RANGE = re.compile(r"^(\d+)\s*~\s*(\d+)")


def nfc(text: str) -> str:
    """NFD 등으로 분해된 한글을 NFC로 합성."""
    if not text:
        return text
    return unicodedata.normalize("NFC", text)


def numeric_tokens(value) -> tuple[list[str], list[str]]:
    """줄바꿈 분리 토큰과 그중 앞머리 숫자 토큰 목록을 반환."""
    toks = [t.strip() for t in str(value).split("\n") if t.strip()]
    nums = [t for t in toks if _LEAD_NUM.match(t)]
    return toks, nums


def clean_value(raw):
    """학점/학기 셀 값을 정제한다.

    반환: (value, flag)
      - value: 숫자면 float, 비숫자면 정제된 문자열, 빈값이면 None
      - flag: 정제 종류(range_value/paren_note/multi_credit/glyph/nfc) 또는 ""

    정책표: 3'→3(glyph), 3º→3(glyph), 3\\n3(4)→3(multi_credit),
            16\\n(70)→16(paren_note), 29~30→29(range_value)
    """
    if raw is None:
        return None, ""
    text0 = str(raw)
    text = nfc(text0)
    flag = ""
    if text != text0:
        flag = "nfc"
    stripped = text.strip()
    if stripped == "":
        return None, ""

    lines = [ln.strip() for ln in stripped.split("\n") if ln.strip()]
    first = lines[0] if lines else stripped

    # 다중 학점(줄바꿈 적층이 단일 셀에 남은 경우)
    if len(lines) >= 2:
        rest = lines[1]
        if rest.startswith("("):
            flag = "paren_note"
        elif _LEAD_NUM.match(rest):
            flag = "multi_credit"

    # 범위값
    m = _RANGE.match(first)
    if m:
        return float(m.group(1)), "range_value"

    # 앞머리 숫자 추출(3', 3º, (70), 4 등)
    m = _LEAD_NUM.match(first)
    if m:
        val = float(m.group(1))
        # glyph 잔여(따옴표/도기호 등) 감지
        residue = first[m.end():]
        if flag in ("", "nfc") and re.search(r"['′°º]", residue):
            flag = "glyph"
        return val, flag

    # 비숫자 텍스트(택N/교차집중/공동교육과정 등)는 원문 유지
    return first, flag


def to_credit(raw) -> float:
    """학점 셀을 float로. 비숫자/빈값은 0.0."""
    val, _ = clean_value(raw)
    return float(val) if isinstance(val, (int, float)) else 0.0


def split_subject_cell(text: str) -> list[str]:
    """선택군 셀('물리학, 화학, 생명과학')을 개별 과목으로 분리.

    '기술·가정'의 가운뎃점은 구분자로 취급하지 않는다. '↔'는 교차이수 전용이라
    여기서 분리하지 않는다(classify에서 처리). ▪▫(작은 사각형)은 성신고가 과목
    접미 마커 겸 구분자로 써서 '물질과 에너지▪ 생물의 유전▪ ...' 형태가 나오므로
    분리자로 함께 취급한다(다른 학교엔 등장하지 않아 안전).
    """
    parts = re.split(r"[,/\n|▪▫◾◽]+", text)
    return [p.strip() for p in parts if p.strip()]


def strip_markers(name: str) -> str:
    """각주 마커/선택 접미를 제거하고 별칭을 적용한 정제 과목명."""
    cleaned = nfc(repair_text(name))
    cleaned = FOOTNOTE_MARKERS.sub("", cleaned).strip()
    cleaned = CHOICE_SUFFIX.sub("", cleaned).strip()
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return NAME_ALIASES.get(cleaned, cleaned)


def is_legend(name: str) -> bool:
    return bool(LEGEND_ROW.search(name.strip()))


def is_skip_name(name: str) -> bool:
    compact = name.replace(" ", "")
    return bool(SKIP_NAME_PATTERNS.search(compact))


def normalize_group(text: str) -> str:
    """교과군 표기 변형('과 학', '사회\\n(역사/도덕 포함)')을 표준형으로."""
    compact = re.sub(r"\s+", "", nfc(repair_text(text))).replace("⋅", "·").replace("ㆍ", "·")
    return compact.split("(")[0] if "(" in compact else compact


def normalize_section(text: str) -> str:
    """'학 교 지 정 교 육 과 정' 같은 변형을 표준 구분값으로."""
    compact = re.sub(r"[\s●○]+", "", nfc(repair_text(text)))
    if not compact:
        return ""
    if "지정" in compact:
        return "학교지정"
    if "선택" in compact:
        return "학생선택"
    if "공동" in compact:
        return "공동교육과정"
    if "소인수" in compact:
        return "소인수"
    if "순증" in compact:
        return "순증"
    if "계절" in compact:
        return "계절수업"
    return compact
