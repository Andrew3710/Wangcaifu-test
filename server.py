#!/usr/bin/env python3

from __future__ import annotations

import argparse
import cgi
import json
import sys
import tempfile
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote


SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from report_stats import generate_circulation_export, generate_report_payload  # noqa: E402


class ReportStatsHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(SCRIPT_DIR), **kwargs)

    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def do_GET(self) -> None:
        if self.path == "/api/health":
            self._write_json({"ok": True})
            return
        super().do_GET()

    def do_POST(self) -> None:
        if self.path not in {"/api/report", "/api/export-circulation"}:
            self.send_error(HTTPStatus.NOT_FOUND, "接口不存在")
            return

        temp_path: Path | None = None
        try:
            form = self._parse_form()
            if "file" not in form:
                raise ValueError("请先上传 Excel 文件。")

            file_item = form["file"]
            if not getattr(file_item, "filename", ""):
                raise ValueError("上传文件缺少文件名。")

            upload_bytes = file_item.file.read()
            if not upload_bytes:
                raise ValueError("上传文件为空。")

            suffix = Path(file_item.filename).suffix or ".xlsx"
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
                temp_file.write(upload_bytes)
                temp_path = Path(temp_file.name)

            mode = form.getfirst("mode", "week")
            payload_kwargs = {"excel_path": temp_path}
            if mode == "custom":
                payload_kwargs["start_date"] = form.getfirst("start_date", "").strip()
                payload_kwargs["end_date"] = form.getfirst("end_date", "").strip()
            else:
                payload_kwargs["anchor_date"] = form.getfirst("anchor_date", "").strip()
            payload_kwargs["display_name"] = file_item.filename

            if self.path == "/api/export-circulation":
                workbook_bytes, rows = generate_circulation_export(**payload_kwargs)
                filename = f"报表流转明细_{sanitize_filename(file_item.filename)}.xlsx"
                self._write_file(
                    workbook_bytes,
                    filename,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    rows_count=len(rows),
                )
            else:
                payload = generate_report_payload(**payload_kwargs)
                payload["filename"] = file_item.filename
                self._write_json({"ok": True, "data": payload})
        except SystemExit as exc:
            self._write_json({"ok": False, "error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
        except ValueError as exc:
            self._write_json({"ok": False, "error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
        except Exception as exc:  # pragma: no cover
            self._write_json(
                {"ok": False, "error": f"服务处理失败：{exc}"},
                status=HTTPStatus.INTERNAL_SERVER_ERROR,
            )
        finally:
            if temp_path and temp_path.exists():
                temp_path.unlink(missing_ok=True)

    def _parse_form(self) -> cgi.FieldStorage:
        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            raise ValueError("请求格式错误，请重新上传文件。")

        environ = {
            "REQUEST_METHOD": "POST",
            "CONTENT_TYPE": content_type,
            "CONTENT_LENGTH": self.headers.get("Content-Length", "0"),
        }
        return cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ=environ,
            keep_blank_values=True,
        )

    def _write_json(self, payload: dict[str, object], status: int = HTTPStatus.OK) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _write_file(
        self,
        body: bytes,
        filename: str,
        content_type: str,
        rows_count: int,
    ) -> None:
        ascii_filename = sanitize_ascii_filename(filename)
        encoded_filename = quote(filename, safe="")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header(
            "Content-Disposition",
            f"attachment; filename=\"{ascii_filename}\"; filename*=UTF-8''{encoded_filename}",
        )
        self.send_header("X-Circulation-Rows", str(rows_count))
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="启动报表统计本地网页版。")
    parser.add_argument("--host", default="127.0.0.1", help="监听地址，默认 127.0.0.1")
    parser.add_argument("--port", type=int, default=8765, help="监听端口，默认 8765")
    return parser.parse_args()


def sanitize_filename(name: str) -> str:
    stem = Path(name).stem
    return "".join(ch for ch in stem if ch not in '\\/:*?"<>|').strip() or "report"


def sanitize_ascii_filename(name: str) -> str:
    path = Path(name)
    stem = sanitize_filename(path.stem)
    ascii_stem = "".join(ch for ch in stem if ord(ch) < 128).strip(" ._-")
    if not ascii_stem:
        ascii_stem = "report"
    suffix = path.suffix or ".bin"
    return f"{ascii_stem}{suffix}"


def main() -> int:
    args = parse_args()
    server = ThreadingHTTPServer((args.host, args.port), ReportStatsHandler)
    url = f"http://{args.host}:{args.port}"
    print(f"报表统计网页版已启动: {url}")
    print("按 Ctrl+C 停止服务。")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止。")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
