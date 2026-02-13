from __future__ import annotations

import json
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from . import matching, proposal, sheet_sync, storage, tasks_engine


class LedgerAPIHandler(BaseHTTPRequestHandler):
    server_version = "LedgerAPI/3.1"

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length", "0") or "0")
        if length <= 0:
            return {}
        raw = self.rfile.read(length)
        if not raw:
            return {}
        return json.loads(raw.decode("utf-8"))

    def _send(self, status: int, payload: dict | list) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_error(self, status: int, message: str) -> None:
        self._send(status, {"error": message})

    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)

        try:
            if path == "/health":
                self._send(HTTPStatus.OK, {"status": "ok"})
                return

            if path == "/properties":
                tab = qs.get("tab", [None])[0]
                include_deleted = qs.get("include_deleted", ["false"])[0].lower() == "true"
                self._send(HTTPStatus.OK, storage.list_properties(tab=tab, include_deleted=include_deleted))
                return

            if path == "/customers":
                include_deleted = qs.get("include_deleted", ["false"])[0].lower() == "true"
                self._send(HTTPStatus.OK, storage.list_customers(include_deleted=include_deleted))
                return

            if path == "/tasks":
                status = qs.get("status", [None])[0]
                self._send(HTTPStatus.OK, storage.list_tasks(status=status))
                return

            if path.startswith("/matching/"):
                customer_id = int(path.split("/")[-1])
                customer = storage.get_customer(customer_id)
                if not customer:
                    self._send_error(HTTPStatus.NOT_FOUND, "Customer not found")
                    return
                limit = int(qs.get("limit", [30])[0])
                result = matching.match_properties(customer, storage.list_properties(include_deleted=False), limit=limit)
                self._send(
                    HTTPStatus.OK,
                    {
                        "customer": customer,
                        "matches": [
                            {"property": item.property_row, "score": item.score, "reasons": item.reasons}
                            for item in result
                        ],
                    },
                )
                return

            self._send_error(HTTPStatus.NOT_FOUND, "Not found")
        except Exception as exc:  # defensive response
            self._send_error(HTTPStatus.BAD_REQUEST, str(exc))

    def do_POST(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        path = parsed.path

        try:
            payload = self._read_json()

            if path == "/properties":
                property_id = storage.add_property(payload)
                tasks_engine.reconcile_auto_tasks()
                self._send(HTTPStatus.CREATED, {"id": property_id, "property": storage.get_property(property_id, include_deleted=True)})
                return

            if path == "/customers":
                customer_id = storage.add_customer(payload)
                self._send(HTTPStatus.CREATED, {"id": customer_id, "customer": storage.get_customer(customer_id, include_deleted=True)})
                return

            if path == "/viewings":
                required = ["property_id", "start_at", "end_at", "title"]
                missing = [k for k in required if not payload.get(k)]
                if missing:
                    self._send_error(HTTPStatus.BAD_REQUEST, f"Missing required fields: {', '.join(missing)}")
                    return
                viewing_id = storage.add_viewing(
                    property_id=int(payload["property_id"]),
                    customer_id=payload.get("customer_id"),
                    start_at=str(payload["start_at"]),
                    end_at=str(payload["end_at"]),
                    title=str(payload["title"]),
                    memo=str(payload.get("memo", "")),
                    status=str(payload.get("status", "예정")),
                )
                tasks_engine.reconcile_auto_tasks()
                self._send(HTTPStatus.CREATED, {"id": viewing_id, "viewing": storage.get_viewing(viewing_id)})
                return

            if path == "/tasks/reconcile":
                open_count = tasks_engine.reconcile_auto_tasks()
                self._send(HTTPStatus.OK, {"open_auto_tasks": open_count, "tasks": storage.list_auto_tasks(include_done=False)})
                return

            if path == "/sync/export":
                sync_dir = Path(payload.get("sync_dir", "")).expanduser() if payload.get("sync_dir") else None
                settings = sheet_sync.SyncSettings(
                    webhook_url=str(payload.get("webhook_url", "")).strip(),
                    sync_dir=sync_dir if sync_dir else Path(storage.get_setting("sync_dir", str(sheet_sync.DEFAULT_SYNC_DIR))),
                )
                ok, message = sheet_sync.upload_visible_data(
                    storage.list_properties(include_deleted=False),
                    storage.list_customers(include_deleted=False),
                    photos=storage.list_photos_all(),
                    viewings=storage.list_viewings(),
                    tasks=storage.list_tasks(status="OPEN"),
                    settings=settings,
                )
                self._send(HTTPStatus.OK, {"ok": ok, "message": message, "sync_dir": str(settings.sync_dir)})
                return

            if path.startswith("/proposal/message/"):
                customer_id = int(path.split("/")[-1])
                customer = storage.get_customer(customer_id)
                if not customer:
                    self._send_error(HTTPStatus.NOT_FOUND, "Customer not found")
                    return
                property_ids = payload.get("property_ids") or []
                if property_ids:
                    selected = [storage.get_property(int(pid)) for pid in property_ids]
                    properties = [p for p in selected if p]
                else:
                    matches = matching.match_properties(customer, storage.list_properties(include_deleted=False), limit=10)
                    properties = [m.property_row for m in matches]
                message = proposal.build_kakao_message(customer, properties)
                self._send(HTTPStatus.OK, {"customer_id": customer_id, "count": len(properties), "message": message})
                return

            self._send_error(HTTPStatus.NOT_FOUND, "Not found")
        except Exception as exc:
            self._send_error(HTTPStatus.BAD_REQUEST, str(exc))

    def do_PATCH(self) -> None:  # noqa: N802
        path = urlparse(self.path).path
        if not path.startswith("/properties/"):
            self._send_error(HTTPStatus.NOT_FOUND, "Not found")
            return
        try:
            property_id = int(path.split("/")[-1])
            payload = self._read_json()
            storage.update_property(property_id, payload)
            tasks_engine.reconcile_auto_tasks()
            self._send(HTTPStatus.OK, {"property": storage.get_property(property_id, include_deleted=True)})
        except ValueError as exc:
            self._send_error(HTTPStatus.NOT_FOUND, str(exc))
        except Exception as exc:
            self._send_error(HTTPStatus.BAD_REQUEST, str(exc))


def run_api_server(host: str = "0.0.0.0", port: int = 8000) -> None:
    storage.init_db()
    server = ThreadingHTTPServer((host, port), LedgerAPIHandler)
    print(f"Ledger API server running on http://{host}:{port}")
    server.serve_forever()
