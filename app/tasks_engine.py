"""Auto 'Next Action' tasks engine.

목표
----
- 엄마가 '오늘 뭐 해야 하지?'를 생각하지 않게.
- 데이터(물건/일정)가 바뀌면 자동으로 할 일이 생성/해제.
- Google API 없이도 100% 동작.

원칙
----
- AUTO_* 태스크는 unique_key로 중복 방지.
- 조건이 해소되면 OPEN → DONE 처리(자동).

주의
----
- 여기서 만드는 태스크는 "권장"입니다.
  (필요시 MANUAL 태스크로 별도 관리 가능)
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Any

from .storage import (
    list_auto_tasks,
    list_photos_all,
    list_properties,
    list_viewings,
    set_task_status,
    upsert_task_unique,
)


def _parse_dt(text: str) -> datetime | None:
    """Parse common datetime formats used in this project.

    Supports
    - ISO8601: 2026-02-13T14:30:00
    - Local text: 2026-02-13 14:30
    - Local text with seconds: 2026-02-13 14:30:00
    """
    s = (text or "").strip()
    if not s:
        return None
    try:
        return datetime.fromisoformat(s)
    except Exception:
        pass
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None


def _fmt_local(dt: datetime) -> str:
    return dt.strftime("%Y-%m-%d %H:%M")


@dataclass
class DesiredTask:
    unique_key: str
    kind: str
    entity_type: str
    entity_id: int | None
    title: str
    due_at: str | None = None
    note: str = ""
    status: str = "OPEN"


def compute_desired_auto_tasks(
    *,
    properties: list[dict[str, Any]],
    photos: list[dict[str, Any]],
    viewings: list[dict[str, Any]],
    now: datetime | None = None,
) -> dict[str, DesiredTask]:
    now = now or datetime.now()

    # photo count by property
    photo_count: dict[int, int] = {}
    for ph in photos:
        pid = ph.get("property_id")
        if pid is None:
            continue
        try:
            pid_i = int(pid)
        except Exception:
            continue
        photo_count[pid_i] = photo_count.get(pid_i, 0) + 1

    desired: dict[str, DesiredTask] = {}

    # --- Properties: missing core info / missing photos
    for p in properties:
        if p.get("deleted") or p.get("hidden"):
            continue
        try:
            pid = int(p.get("id"))
        except Exception:
            continue

        complex_name = str(p.get("complex_name") or "").strip()
        address = str(p.get("address_detail") or "").strip()
        unit_type = str(p.get("unit_type") or "").strip()

        name_hint = " ".join([x for x in [complex_name, address, unit_type] if x]).strip()
        if not name_hint:
            name_hint = f"물건 {pid}"

        missing = []
        if not address:
            missing.append("동/호")
        if p.get("area") in (None, ""):
            missing.append("면적")
        if not unit_type and str(p.get("tab")) in ("봉담자이 프라이드시티", "힐스테이트봉담프라이드시티"):
            missing.append("타입")
        if not str(p.get("floor") or "").strip():
            missing.append("층")

        if missing:
            key = f"AUTO_PROP_INFO:{pid}"
            desired[key] = DesiredTask(
                unique_key=key,
                kind="AUTO_PROP_INFO",
                entity_type="PROPERTY",
                entity_id=pid,
                title=f"[정보 보완] {name_hint} (누락: {', '.join(missing)})",
                due_at=None,
                note="핵심 항목이 누락되어 있습니다.",
            )

        status = str(p.get("status") or "")
        if status != "거래완료" and (photo_count.get(pid, 0) == 0 or status == "사진필요"):
            key = f"AUTO_PROP_PHOTO:{pid}"
            desired[key] = DesiredTask(
                unique_key=key,
                kind="AUTO_PROP_PHOTO",
                entity_type="PROPERTY",
                entity_id=pid,
                title=f"[사진 등록] {name_hint}",
                due_at=None,
                note="고객 제안/광고용 사진이 필요합니다.",
            )

    # --- Viewings: upcoming soon / missing result memo
    for v in viewings:
        try:
            vid = int(v.get("id") or v.get("viewing_id") or 0)
        except Exception:
            continue
        status = str(v.get("status") or "").strip()
        start_dt = _parse_dt(str(v.get("start_at") or v.get("start_iso") or ""))
        title = str(v.get("title") or "현장/상담").strip()

        if status == "예정" and start_dt:
            if start_dt < now - timedelta(hours=1):
                key = f"AUTO_VIEWING_OVERDUE:{vid}"
                desired[key] = DesiredTask(
                    unique_key=key,
                    kind="AUTO_VIEWING_OVERDUE",
                    entity_type="VIEWING",
                    entity_id=vid,
                    title=f"[지난 일정 확인] {title} ({_fmt_local(start_dt)})",
                    due_at=_fmt_local(start_dt),
                    note="일정이 지났는데 '예정'으로 남아있습니다. 완료/취소/메모를 정리하세요.",
                )
            elif start_dt <= now + timedelta(hours=24):
                key = f"AUTO_VIEWING_PREP:{vid}"
                desired[key] = DesiredTask(
                    unique_key=key,
                    kind="AUTO_VIEWING_PREP",
                    entity_type="VIEWING",
                    entity_id=vid,
                    title=f"[일정 준비] {title} ({_fmt_local(start_dt)})",
                    due_at=_fmt_local(start_dt),
                    note="고객 연락/주소/주차/열쇠/세입자 등 사전 체크.",
                )

        if status == "완료":
            memo = str(v.get("memo") or "").strip()
            if not memo:
                key = f"AUTO_VIEWING_RESULT:{vid}"
                desired[key] = DesiredTask(
                    unique_key=key,
                    kind="AUTO_VIEWING_RESULT",
                    entity_type="VIEWING",
                    entity_id=vid,
                    title=f"[결과 기록] {title} (일정 {vid})",
                    due_at=None,
                    note="완료된 일정인데 메모가 비어 있습니다. 결과/피드백을 남기세요.",
                )

    return desired


def reconcile_auto_tasks() -> int:
    """Create/refresh AUTO_* tasks and resolve tasks that are no longer needed.

    Returns
    -------
    int
        Number of currently OPEN auto tasks.
    """
    properties = list_properties(include_deleted=False)
    photos = list_photos_all()
    viewings = list_viewings()

    desired = compute_desired_auto_tasks(properties=properties, photos=photos, viewings=viewings)
    desired_keys = set(desired.keys())

    existing = list_auto_tasks(include_done=True)
    existing_by_key = {str(t.get("unique_key")): t for t in existing if t.get("unique_key")}

    # 1) Upsert desired
    for key, t in desired.items():
        upsert_task_unique(
            unique_key=t.unique_key,
            kind=t.kind,
            entity_type=t.entity_type,
            entity_id=t.entity_id,
            title=t.title,
            due_at=t.due_at,
            status="OPEN",
            note=t.note,
        )

    # 2) Resolve no-longer-needed
    for key, row in existing_by_key.items():
        if key and key not in desired_keys and str(row.get("status")) == "OPEN":
            try:
                set_task_status(int(row.get("id")), "DONE")
            except Exception:
                continue

    open_now = list_auto_tasks(include_done=False)
    return len(open_now)
