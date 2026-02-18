"""Sync/Export layer.

요구사항
--------
- Google API(Drive/Sheets/Calendar) 없이도 100% 운영 가능해야 함.
- (선택) Webhook(Apps Script Web App) URL이 있으면 즉시 업로드.
- Webhook이 없어도 Drive 동기화 폴더에 CSV/JSON/ICS를 생성해서
  어디서나 열람(Gemini/Sheets/Drive) 가능하도록 함.

표준 라이브러리만 사용.
"""

from __future__ import annotations

import csv
import json
import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib import request


DEFAULT_SYNC_DIR = Path.home() / "RealEstateLedgerSync"


def _visible_only(rows: list[dict]) -> list[dict]:
    # hidden=1이면 제외. deleted=1이면 제외.
    out = []
    for r in rows:
        if r.get("hidden"):
            continue
        if r.get("deleted"):
            continue
        out.append(r)
    return out


@dataclass
class SyncSettings:
    webhook_url: str = ""
    sync_dir: Path = DEFAULT_SYNC_DIR
    export_mode: str = "single"  # single / advanced
    export_ics: bool = False

    @staticmethod
    def from_env() -> "SyncSettings":
        webhook_url = os.getenv("GOOGLE_SHEETS_WEBHOOK_URL", "").strip()
        sync_dir = Path(os.getenv("GOOGLE_DRIVE_SYNC_DIR", str(DEFAULT_SYNC_DIR))).expanduser()
        export_mode = os.getenv("EXPORT_MODE", "single").strip() or "single"
        export_ics = os.getenv("EXPORT_ICS", "0").strip() in {"1", "true", "True", "on", "ON"}
        return SyncSettings(webhook_url=webhook_url, sync_dir=sync_dir, export_mode=export_mode, export_ics=export_ics)




def _last4_phone(value: Any) -> str:
    digits = "".join(ch for ch in str(value or "") if ch.isdigit())
    return digits[-4:] if digits else ""


def _anonymize_value(key: str, value: Any) -> Any:
    k = (key or "").lower()

    # 이름 계열은 전부 제거
    if "name" in k:
        return ""

    # 전화번호 계열은 뒷자리 4자리만 유지
    if any(token in k for token in ("phone", "tel", "mobile")):
        return _last4_phone(value)

    return value


def _anonymize_rows(rows: list[dict]) -> list[dict]:
    out: list[dict] = []
    for row in rows:
        cleaned = {k: _anonymize_value(k, v) for k, v in row.items()}
        out.append(cleaned)
    return out


def upload_visible_data(
    properties: list[dict],
    customers: list[dict],
    *,
    photos: list[dict] | None = None,
    viewings: list[dict] | None = None,
    tasks: list[dict] | None = None,
    settings: SyncSettings | None = None,
) -> tuple[bool, str]:
    """Export(항상) + optional webhook push.

    Returns
    -------
    (ok, message)
    """

    settings = settings or SyncSettings.from_env()
    settings.sync_dir.mkdir(parents=True, exist_ok=True)

    visible_props = _visible_only(properties)
    visible_customers = _visible_only(customers)
    photos = photos or []
    viewings = viewings or []
    tasks = tasks or []

    # 개인정보 마스킹(이름 제거 + 전화번호 뒷4자리)
    export_props = _anonymize_rows(visible_props)
    export_customers = _anonymize_rows(visible_customers)
    export_photos = _anonymize_rows(photos)
    export_viewings = _anonymize_rows(viewings)
    export_tasks = _anonymize_rows(tasks)

    # 1) Always export to files (API-less fallback)
    exported = export_all(
        sync_dir=settings.sync_dir,
        properties=export_props,
        customers=export_customers,
        photos=export_photos,
        viewings=export_viewings,
        tasks=export_tasks,
        mode=settings.export_mode,
        export_ics=settings.export_ics,
    )

    # 2) Optional webhook upload
    webhook_ok = None
    webhook_msg = ""
    if settings.webhook_url:
        webhook_ok, webhook_msg = _post_webhook(
            settings.webhook_url,
            {
                "uploaded_at": datetime.now().isoformat(timespec="seconds"),
                "properties": export_props,
                "customers": export_customers,
                "viewings": export_viewings,
                "tasks": export_tasks,
            },
        )

    # 메시지 조립
    msg = f"파일 내보내기 완료: {exported['base_dir']}"
    if exported.get("snapshot_json"):
        msg += f"\n- snapshot: {exported['snapshot_json'].name}"
    if exported.get("properties_csv"):
        msg += "\n- CSV: properties/customers/tasks"
    if exported.get("ics"):
        msg += f"\n- calendar: {exported['ics'].name}"

    if settings.webhook_url:
        if webhook_ok:
            msg += "\n웹훅 업로드: 성공"
        else:
            msg += f"\n웹훅 업로드: 실패 ({webhook_msg})"

    return True, msg


def export_all(
    *,
    sync_dir: Path,
    properties: list[dict],
    customers: list[dict],
    photos: list[dict],
    viewings: list[dict],
    tasks: list[dict],
    mode: str = "single",
    export_ics: bool = False,
) -> dict[str, Any]:
    """Create exports into sync_dir.

    mode == "single":
      exports/ledger_snapshot.json only (기본)
    mode == "advanced":
      snapshot + CSV + optional ICS
    """

    base = sync_dir / "exports"
    base.mkdir(parents=True, exist_ok=True)

    snapshot_json = base / "ledger_snapshot.json"
    snapshot = {
        "exported_at": datetime.now().isoformat(timespec="seconds"),
        "properties": properties,
        "customers": customers,
        "photos": photos,
        "viewings": viewings,
        "tasks": tasks,
    }
    snapshot_json.write_text(json.dumps(snapshot, ensure_ascii=False, indent=2), encoding="utf-8")

    out: dict[str, Any] = {
        "base_dir": base,
        "snapshot_json": snapshot_json,
    }

    if mode == "advanced":
        prop_csv = base / "visible_properties.csv"
        cust_csv = base / "visible_customers.csv"
        tasks_csv = base / "open_tasks.csv"
        _write_csv(prop_csv, properties)
        _write_csv(cust_csv, customers)
        _write_csv(tasks_csv, tasks)
        out["properties_csv"] = prop_csv
        out["customers_csv"] = cust_csv
        out["tasks_csv"] = tasks_csv

    if export_ics:
        ics_dir = base / ("_calendar" if mode == "single" else "")
        ics_dir.mkdir(parents=True, exist_ok=True)
        ics_file = ics_dir / "viewings.ics"
        ics_file.write_text(_to_ics(viewings), encoding="utf-8")
        out["ics"] = ics_file

    return out


def _write_csv(path: Path, rows: list[dict]) -> None:
    keys = sorted({k for row in rows for k in row.keys()}) if rows else ["empty"]
    with path.open("w", newline="", encoding="utf-8-sig") as fp:
        writer = csv.DictWriter(fp, fieldnames=keys)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def _post_webhook(url: str, payload_obj: dict) -> tuple[bool, str]:
    payload = json.dumps(payload_obj, ensure_ascii=False).encode("utf-8")
    req = request.Request(
        url,
        data=payload,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )
    try:
        with request.urlopen(req, timeout=20) as res:
            code = getattr(res, "status", 200)
        if 200 <= code < 300:
            return True, "OK"
        return False, f"HTTP {code}"
    except Exception as exc:
        return False, str(exc)


def _to_ics(viewings: list[dict]) -> str:
    """Basic iCalendar (.ics) output.

    - Google Calendar로 가져오기(import) 가능한 최소 스펙.
    - API 없는 운영 시 '캘린더 공유'의 현실적 대안.
    """

    def esc(s: Any) -> str:
        s = "" if s is None else str(s)
        return (
            s.replace("\\", "\\\\")
            .replace(";", "\\;")
            .replace(",", "\\,")
            .replace("\n", "\\n")
        )

    def dtstamp() -> str:
        return datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

    def to_utc_dt(iso_or_text: str) -> str:
        # 저장 포맷이 다양할 수 있어 최대한 관대하게 처리
        s = (iso_or_text or "").strip()
        if not s:
            return dtstamp()
        try:
            # 1) ISO8601
            dt = datetime.fromisoformat(s)
        except Exception:
            # 2) "YYYY-MM-DD HH:MM" fallback
            try:
                dt = datetime.strptime(s, "%Y-%m-%d %H:%M")
            except Exception:
                return dtstamp()

        # naive -> local(Asia/Seoul)로 가정 후 UTC로 변환
        try:
            from zoneinfo import ZoneInfo

            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=ZoneInfo("Asia/Seoul"))
            dt_utc = dt.astimezone(ZoneInfo("UTC"))
            return dt_utc.strftime("%Y%m%dT%H%M%SZ")
        except Exception:
            # zoneinfo가 없거나 실패하면 naive를 UTC로 간주
            return dt.strftime("%Y%m%dT%H%M%SZ")

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//RealEstateLedger//KO",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
    ]

    now_stamp = dtstamp()
    for v in viewings:
        vid = v.get("id") or v.get("viewing_id") or "0"
        uid = f"VW_{vid}@realestateledger.local"
        summary = v.get("title") or "현장/상담 일정"
        desc = v.get("memo") or ""
        start = to_utc_dt(v.get("start_at") or v.get("start_iso") or "")
        end = to_utc_dt(v.get("end_at") or v.get("end_iso") or "")

        lines.extend(
            [
                "BEGIN:VEVENT",
                f"UID:{esc(uid)}",
                f"DTSTAMP:{now_stamp}",
                f"DTSTART:{start}",
                f"DTEND:{end}",
                f"SUMMARY:{esc(summary)}",
                f"DESCRIPTION:{esc(desc)}",
                "END:VEVENT",
            ]
        )

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"
