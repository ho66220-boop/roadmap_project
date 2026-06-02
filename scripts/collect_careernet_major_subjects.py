"""Collect CareerNet 2022 curriculum elective subjects for university majors.

The CareerNet university-major pages are rendered from public JSON endpoints:

    POST /cloud/api/major/uSearch
    GET  /cloud/api/major/uView?seq=...

This script keeps requests conservative and writes one CSV row per major.
"""

from __future__ import annotations

import argparse
import csv
from datetime import datetime, timezone
import json
from pathlib import Path
import re
import time
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


BASE_URL = "https://www.career.go.kr"
DEFAULT_OUTPUT = Path(__file__).resolve().parents[1] / "data" / "careernet_major_subjects.csv"
USER_AGENT = "roadmap-project/0.1 (+https://github.com/ho66220-boop)"

CATEGORY_LABELS = {
    "일반": "general_selection_subjects",
    "진로": "career_selection_subjects",
    "융합": "convergence_selection_subjects",
}


def normalize_space(text: Any) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def request_json(url: str, *, method: str = "GET", payload: dict[str, Any] | None = None) -> dict[str, Any]:
    body = None
    headers = {
        "Accept": "application/json",
        "User-Agent": USER_AGENT,
    }
    if payload is not None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json; charset=utf-8"
    request = Request(url, data=body, headers=headers, method=method)
    with urlopen(request, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))


def collect_major_list(page_size: int = 200) -> list[dict[str, Any]]:
    payload = {"keyword": "", "selectedtab": [], "selectedSchool": [], "selectedEmp": []}
    sort = "updt_dt,desc&sort=regist_dt,desc&sort=major_nm,asc&sort=rdcnt,desc&sort=likecnt,desc"
    first = request_json(
        f"{BASE_URL}/cloud/api/major/uSearch?size={page_size}&page=0&sort={sort}",
        method="POST",
        payload=payload,
    )
    rows = list(first.get("content") or [])
    total_pages = int(first.get("totalPages") or 1)
    for page in range(1, total_pages):
        data = request_json(
            f"{BASE_URL}/cloud/api/major/uSearch?size={page_size}&page={page}&sort={sort}",
            method="POST",
            payload=payload,
        )
        rows.extend(data.get("content") or [])
    return rows


def split_subjects(text: str) -> list[str]:
    text = normalize_space(text)
    text = re.sub(r"\s*등$", "", text)
    items = [normalize_space(item) for item in re.split(r",|ㆍ|·|/", text)]
    subjects = []
    for item in items:
        if not item or item == "등":
            continue
        if item in {"Ⅰ", "Ⅱ", "I", "II"} and subjects:
            stem = re.sub(r"\s*(?:Ⅰ|Ⅱ|I|II)$", "", subjects[-1]).strip()
            if stem:
                item = f"{stem} {item}"
        if len(item) == 1:
            continue
        subjects.append(item)
    return subjects


def parse_relate_subjects_2022(items: list[dict[str, Any]]) -> dict[str, str]:
    parsed = {
        "general_selection_subjects": "",
        "career_selection_subjects": "",
        "convergence_selection_subjects": "",
        "source_text": "",
    }
    current = ""
    buckets: dict[str, list[str]] = {value: [] for value in CATEGORY_LABELS.values()}

    for item in sorted(items or [], key=lambda row: int(row.get("sort_ORDER") or 0)):
        description = normalize_space(item.get("subject_DESCRIPTION"))
        if not description:
            continue
        if description.startswith("[출처"):
            parsed["source_text"] = description.strip("[]")
            continue
        if "일반" in description and "선택" in description and len(description) <= 20:
            current = "일반"
            continue
        if "진로" in description and "선택" in description and len(description) <= 20:
            current = "진로"
            continue
        if "융합" in description and "선택" in description and len(description) <= 20:
            current = "융합"
            continue
        if current in CATEGORY_LABELS:
            buckets[CATEGORY_LABELS[current]].extend(split_subjects(description))

    for field, subjects in buckets.items():
        seen = set()
        deduped = []
        for subject in subjects:
            key = re.sub(r"\s+", "", subject).lower()
            if key in seen:
                continue
            seen.add(key)
            deduped.append(subject)
        parsed[field] = "|".join(deduped)
    return parsed


def collect_detail(seq: str) -> dict[str, Any]:
    return request_json(f"{BASE_URL}/cloud/api/major/uView?seq={seq}")


def build_row(detail: dict[str, Any]) -> dict[str, str]:
    parsed_subjects = parse_relate_subjects_2022(detail.get("relateSubject2022") or [])
    seq = normalize_space(detail.get("seq"))
    return {
        "career_seq": seq,
        "career_major_name": normalize_space(detail.get("major_NM")),
        "career_major_category": normalize_space(detail.get("major_CLNM")),
        "updated_at": normalize_space(detail.get("updt")),
        **parsed_subjects,
        "source_url": f"{BASE_URL}/cloud/w/major/uView?seq={seq}",
        "collected_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }


def write_csv(rows: list[dict[str, str]], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    fields = [
        "career_seq",
        "career_major_name",
        "career_major_category",
        "updated_at",
        "general_selection_subjects",
        "career_selection_subjects",
        "convergence_selection_subjects",
        "source_text",
        "source_url",
        "collected_at",
    ]
    with output.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    parser = argparse.ArgumentParser(description="Collect CareerNet 2022 elective subjects.")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--limit", type=int, default=0, help="Maximum detail pages to collect. 0 means all.")
    parser.add_argument("--delay", type=float, default=0.15, help="Delay between detail requests in seconds.")
    parser.add_argument("--start-page", type=int, default=0, help="Reserved for manual retry workflows.")
    args = parser.parse_args()

    majors = collect_major_list()
    if args.limit:
        majors = majors[: args.limit]

    rows = []
    failures = []
    for index, major in enumerate(majors, start=1):
        seq = normalize_space(major.get("seq"))
        if not seq:
            continue
        try:
            detail = collect_detail(seq)
            rows.append(build_row(detail))
            print(f"[{index}/{len(majors)}] {seq} {rows[-1]['career_major_name']}")
        except (HTTPError, URLError, TimeoutError, json.JSONDecodeError) as exc:
            failures.append((seq, str(exc)))
            print(f"[warn] failed seq={seq}: {exc}")
        if args.delay:
            time.sleep(args.delay)

    write_csv(rows, args.output)
    print(f"Wrote {len(rows)} rows to {args.output}")
    if failures:
        print(f"Failures: {len(failures)}")
        for seq, error in failures[:10]:
            print(f"  {seq}: {error}")
    return 0 if rows else 1


if __name__ == "__main__":
    raise SystemExit(main())
