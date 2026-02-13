import csv
import json
import os
from datetime import datetime
from pathlib import Path
from urllib import request


def _visible_only(rows: list[dict]) -> list[dict]:
    return [r for r in rows if not r.get("hidden")]


def upload_visible_data(properties: list[dict], customers: list[dict]) -> tuple[bool, str]:
    visible_props = _visible_only(properties)
    visible_customers = _visible_only(customers)

    webhook_url = os.getenv("GOOGLE_SHEETS_WEBHOOK_URL", "").strip()
    if webhook_url:
        payload = json.dumps(
            {
                "uploaded_at": datetime.now().isoformat(timespec="seconds"),
                "properties": visible_props,
                "customers": visible_customers,
            },
            ensure_ascii=False,
        ).encode("utf-8")
        req = request.Request(
            webhook_url,
            data=payload,
            headers={"Content-Type": "application/json; charset=utf-8"},
            method="POST",
        )
        try:
            with request.urlopen(req, timeout=20) as res:
                code = getattr(res, "status", 200)
            if 200 <= code < 300:
                return True, "Google Sheets 업로드 성공(웹훅)."
            return False, f"Google Sheets 업로드 실패: HTTP {code}"
        except Exception as exc:
            return False, f"Google Sheets 업로드 실패: {exc}"

    sync_dir = Path(os.getenv("GOOGLE_DRIVE_SYNC_DIR", str(Path.home() / "Google Drive")))
    sync_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    prop_file = sync_dir / f"visible_properties_{timestamp}.csv"
    cust_file = sync_dir / f"visible_customers_{timestamp}.csv"

    _write_csv(prop_file, visible_props)
    _write_csv(cust_file, visible_customers)

    return True, (
        "웹훅 미설정: Google Drive 동기화 폴더에 CSV 저장 완료. "
        "(Gemini/Google Sheets에서 해당 파일 열람 가능)"
    )


def _write_csv(path: Path, rows: list[dict]) -> None:
    keys = sorted({k for row in rows for k in row.keys()}) if rows else ["empty"]
    with path.open("w", newline="", encoding="utf-8-sig") as fp:
        writer = csv.DictWriter(fp, fieldnames=keys)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)
