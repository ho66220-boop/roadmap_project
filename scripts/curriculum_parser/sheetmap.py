# -*- coding: utf-8 -*-
"""시트맵 로드.

원장부 CSV(school, file_keyword, sheet_name, sheet_confirmed, note) 또는
주입 dict 로 학교→(파일키워드, 시트명) 매핑을 제공한다. 테스트/스모크는
임시 CSV 없이 dict 주입으로 사용한다.
"""
from __future__ import annotations

import csv
from pathlib import Path


def load_sheetmap(path: Path | None = None, injected: dict | None = None) -> list[dict]:
    """시트맵 레코드 목록을 반환.

    각 레코드: {school, file_keyword, sheet_name, sheet_confirmed, note}
    - injected: {school: {"file_keyword":..., "sheet_name":..., ...}} 형태 dict 주입
    - path: CSV 경로(utf-8-sig)
    injected 가 있으면 그것만 사용(테스트/스모크용).
    """
    records: list[dict] = []
    if injected:
        for school, cfg in injected.items():
            records.append({
                "school": school,
                "file_keyword": cfg.get("file_keyword", school),
                "sheet_name": cfg.get("sheet_name", ""),
                "sheet_confirmed": cfg.get("sheet_confirmed", "N"),
                "note": cfg.get("note", ""),
            })
        return records
    if path is None or not Path(path).exists():
        return records
    with Path(path).open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            records.append({
                "school": (row.get("school") or "").strip(),
                "file_keyword": (row.get("file_keyword") or "").strip(),
                "sheet_name": (row.get("sheet_name") or "").strip(),
                "sheet_confirmed": (row.get("sheet_confirmed") or "N").strip(),
                "note": (row.get("note") or "").strip(),
            })
    return records
