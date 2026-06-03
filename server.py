from __future__ import annotations

import argparse
import json
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import sys
from urllib.parse import urlparse

from roadmap_rag.graph_rag import GraphRAGEngine


ROOT = Path(__file__).resolve().parent
APP_DIR = ROOT / "app"
DEFAULT_SAMPLE_WORKBOOK = ROOT / "data" / "sample_roadmap.xlsx"
DEFAULT_PRIVATE_WORKBOOK = ROOT.parent / "울산고교_과목로드맵_수정.xlsx"


class ChatbotHandler(SimpleHTTPRequestHandler):
    engine: GraphRAGEngine

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(APP_DIR), **kwargs)

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if path == "/api/meta":
            self._send_json(self.engine.meta())
            return
        if path == "/":
            self.path = "/index.html"
        super().do_GET()

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        if path != "/api/chat":
            self.send_error(404, "Not found")
            return
        length = int(self.headers.get("Content-Length", "0"))
        payload = json.loads(self.rfile.read(length).decode("utf-8") or "{}")
        result = self.engine.chat(
            message=str(payload.get("message", "")),
            profile=payload.get("profile") or {},
        )
        self._send_json(result)

    def _send_json(self, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main() -> int:
    parser = argparse.ArgumentParser(description="Roadmap counseling MVP server")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument(
        "--workbook",
        default=str(DEFAULT_SAMPLE_WORKBOOK if DEFAULT_SAMPLE_WORKBOOK.exists() else DEFAULT_PRIVATE_WORKBOOK),
        help="로드맵 엑셀 파일 경로. 공개 저장소에서는 data/sample_roadmap.xlsx를 기본으로 사용합니다.",
    )
    args = parser.parse_args()

    try:
        ChatbotHandler.engine = GraphRAGEngine(Path(args.workbook))
    except FileNotFoundError:
        print(
            "원본 로드맵 엑셀 파일을 찾을 수 없습니다. README의 데이터 준비 방법을 확인하세요.",
            file=sys.stderr,
        )
        print(f"요청한 경로: {args.workbook}", file=sys.stderr)
        return 1

    server = ThreadingHTTPServer((args.host, args.port), ChatbotHandler)
    print(f"Roadmap chatbot: http://{args.host}:{args.port}")
    print(f"Workbook: {args.workbook}")
    server.serve_forever()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
