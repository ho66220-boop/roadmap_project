# -*- coding: utf-8 -*-
"""편제표 전처리 v2 진입점(얇은 CLI). 로직은 scripts/curriculum_parser 패키지.

병합셀 복원·택N 그룹·정제 정책을 적용해 20열 표준 스키마 엑셀을 생성한다.
앞 18열은 v1과 동일 의미(소비자 무수정 호환), 말미 [선택군ID, 택N] 추가.

사용 예:
    python scripts/build_curriculum.py
    python scripts/build_curriculum.py --config data/curriculum_sheetmap.csv \
        --out data/curriculum_g1_2026.xlsx
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from curriculum_parser import parse_all              # noqa: E402
from curriculum_parser.matching import load_subject_master  # noqa: E402
from curriculum_parser.report import build_workbook  # noqa: E402
from curriculum_parser.sheetmap import load_sheetmap  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SOURCE = ROOT.parent / "울산 고교 편제표"
DEFAULT_MASTER = ROOT.parent / "울산고교_과목로드맵_수정.xlsx"
DEFAULT_CONFIG = ROOT / "data" / "curriculum_sheetmap.csv"
DEFAULT_OUT = ROOT / "data" / "curriculum_g1_2026.xlsx"


def main() -> int:
    ap = argparse.ArgumentParser(description="편제표 전처리 v2(고1, 신입생 3개년)")
    ap.add_argument("--source-dir", default=str(DEFAULT_SOURCE))
    ap.add_argument("--master", default=str(DEFAULT_MASTER))
    ap.add_argument("--config", default=str(DEFAULT_CONFIG),
                    help="시트맵/타깃 CSV (school,file_keyword,sheet_name[,sheet_confirmed,note])")
    ap.add_argument("--out", default=str(DEFAULT_OUT))
    args = ap.parse_args()

    source_dir = Path(args.source_dir)
    if not source_dir.exists():
        print(f"원자료 폴더가 없습니다: {source_dir}", file=sys.stderr)
        return 1

    master = load_subject_master(Path(args.master))
    print(f"과목마스터 로딩: {len(set(master.values()))}개 공식과목")

    records = load_sheetmap(Path(args.config))
    if not records:
        print(f"시트맵/타깃 CSV가 비어있습니다: {args.config}", file=sys.stderr)
        return 1

    all_rows, results = parse_all(source_dir, master, records)

    out = build_workbook(all_rows, results)
    out_path = Path(args.out)
    out.save(out_path)

    print(f"\n저장: {out_path}  (총 {len(all_rows)}행)")
    for e in results:
        rate = round(e["matched"] / e["rows"] * 100, 1) if e["rows"] else 0
        flag = " !!" if e["issues"] or e["rows"] == 0 else ""
        print(f"  {e['school']:10s} {e['rows']:4d}행  매칭 {rate:5.1f}%  "
              f"학기 {e['sem_cov']:4d}  택N {e['groups']:3d}  "
              f"{'; '.join(e['issues'])}{flag}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
