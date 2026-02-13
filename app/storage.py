"""Local storage layer (SQLite).

Design goals
-------------
1) Offline-first: 앱은 Google API/네트워크 없이도 100% 동작.
2) Easy migrations: 스키마가 바뀌어도 기존 DB를 최대한 보존.
3) "삭제"는 소프트 삭제(복구 가능) 기본.

Note
----
표준 라이브러리만 사용합니다.
"""

from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any


BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "ledger.db"


# UI 탭(요구사항 그대로)
PROPERTY_TABS = ["봉담자이 프라이드시티", "힐스테이트봉담프라이드시티", "상가", "단독주택"]


# 탭 -> 고정 단지명(아파트단지는 고정)
TAB_COMPLEX_NAME = {
    "봉담자이 프라이드시티": "봉담자이 프라이드시티",
    "힐스테이트봉담프라이드시티": "힐스테이트봉담프라이드시티",
}

LEGACY_TAB_MAP = {
    "아파트단지1": "봉담자이 프라이드시티",
    "아파트단지2": "힐스테이트봉담프라이드시티",
}


def connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _now_ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def init_db() -> None:
    """Create tables + lightweight migrations."""

    conn = connect()
    cur = conn.cursor()

    # 1) Base tables (create if missing)
    cur.executescript(
        """
        CREATE TABLE IF NOT EXISTS properties (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tab TEXT NOT NULL,
            complex_name TEXT,
            unit_type TEXT,
            area REAL,
            pyeong REAL,
            address_detail TEXT,
            floor TEXT,
            total_floor TEXT,
            view TEXT,
            orientation TEXT,
            condition TEXT,
            repair_needed INTEGER DEFAULT 0,
            tenant_info TEXT,
            naver_link TEXT,
            special_notes TEXT,
            note TEXT,
            status TEXT DEFAULT '신규등록',
            hidden INTEGER DEFAULT 0,
            deleted INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS customer_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_name TEXT NOT NULL,
            phone TEXT,
            preferred_tab TEXT,
            preferred_area TEXT,
            preferred_pyeong TEXT,
            budget TEXT,
            move_in_period TEXT,
            view_preference TEXT,
            location_preference TEXT,
            floor_preference TEXT,
            extra_needs TEXT,
            status TEXT DEFAULT '진행',
            hidden INTEGER DEFAULT 0,
            deleted INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS photos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            property_id INTEGER NOT NULL,
            file_path TEXT NOT NULL,
            tag TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(property_id) REFERENCES properties(id)
        );

        CREATE TABLE IF NOT EXISTS viewings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            property_id INTEGER NOT NULL,
            customer_id INTEGER,
            start_at TEXT NOT NULL,
            end_at TEXT NOT NULL,
            title TEXT,
            memo TEXT,
            status TEXT DEFAULT '예정',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(property_id) REFERENCES properties(id),
            FOREIGN KEY(customer_id) REFERENCES customer_requests(id)
        );

        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        );

        
CREATE TABLE IF NOT EXISTS tasks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    unique_key TEXT,
    kind TEXT DEFAULT 'MANUAL',
    entity_type TEXT,
    entity_id INTEGER,
    title TEXT NOT NULL,
    due_at TEXT,
    status TEXT DEFAULT 'OPEN',
    note TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_tasks_unique_key ON tasks(unique_key);

CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entity_type TEXT NOT NULL,
            entity_id INTEGER NOT NULL,
            action TEXT NOT NULL,
            before_json TEXT,
            after_json TEXT,
            ts TEXT DEFAULT CURRENT_TIMESTAMP
        );
        """
    )

    # 2) Lightweight migrations (ALTER TABLE add missing columns)
    _ensure_columns(
        conn,
        "properties",
        {
            "address_detail": "TEXT",
            "total_floor": "TEXT",
            "view": "TEXT",
            "orientation": "TEXT",
            "status": "TEXT DEFAULT '신규등록'",
            "deleted": "INTEGER DEFAULT 0",
            "updated_at": "TEXT DEFAULT CURRENT_TIMESTAMP",
        },
    )
    _ensure_columns(
        conn,
        "customer_requests",
        {
            "status": "TEXT DEFAULT '진행'",
            "deleted": "INTEGER DEFAULT 0",
            "updated_at": "TEXT DEFAULT CURRENT_TIMESTAMP",
        },
    )

    # Legacy tab values migration
    for old_name, new_name in LEGACY_TAB_MAP.items():
        conn.execute("UPDATE properties SET tab=? WHERE tab=?", (new_name, old_name))
        conn.execute("UPDATE customer_requests SET preferred_tab=? WHERE preferred_tab=?", (new_name, old_name))

    conn.commit()
    conn.close()


def _ensure_columns(conn: sqlite3.Connection, table: str, columns: dict[str, str]) -> None:
    existing = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()}
    for name, ddl in columns.items():
        if name in existing:
            continue
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {ddl}")


def _audit(conn: sqlite3.Connection, entity_type: str, entity_id: int, action: str, before: dict | None, after: dict | None) -> None:
    conn.execute(
        "INSERT INTO audit_log(entity_type, entity_id, action, before_json, after_json, ts) VALUES(?,?,?,?,?,?)",
        (
            entity_type,
            entity_id,
            action,
            json.dumps(before, ensure_ascii=False) if before else None,
            json.dumps(after, ensure_ascii=False) if after else None,
            _now_ts(),
        ),
    )


# ----------------------------
# Settings
# ----------------------------
def get_setting(key: str, default: str = "") -> str:
    conn = connect()
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    conn.close()
    return row["value"] if row and row["value"] is not None else default


def set_setting(key: str, value: str) -> None:
    conn = connect()
    conn.execute(
        "INSERT INTO settings(key,value,updated_at) VALUES(?,?,?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value, updated_at=excluded.updated_at",
        (key, value, _now_ts()),
    )
    conn.commit()
    conn.close()


# ----------------------------
# Properties
# ----------------------------
def add_property(data: dict[str, Any]) -> int:
    conn = connect()
    cur = conn.cursor()
    tab = data.get("tab", "")
    complex_name = data.get("complex_name") or TAB_COMPLEX_NAME.get(tab, "")

    cur.execute(
        """
        INSERT INTO properties
        (tab, complex_name, unit_type, area, pyeong, address_detail, floor, total_floor,
         view, orientation, condition, repair_needed, tenant_info, naver_link,
         special_notes, note, status, hidden, deleted, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?)
        """,
        (
            tab,
            complex_name,
            data.get("unit_type", ""),
            _to_float_or_none(data.get("area")),
            _to_float_or_none(data.get("pyeong")),
            data.get("address_detail", ""),
            data.get("floor", ""),
            data.get("total_floor", ""),
            data.get("view", ""),
            data.get("orientation", ""),
            data.get("condition", ""),
            1 if data.get("repair_needed") else 0,
            data.get("tenant_info", ""),
            data.get("naver_link", ""),
            data.get("special_notes", ""),
            data.get("note", ""),
            data.get("status", "신규등록"),
            1 if data.get("hidden") else 0,
            _now_ts(),
            _now_ts(),
        ),
    )
    pid = int(cur.lastrowid)
    after = get_property(pid, include_deleted=True)
    _audit(conn, "PROPERTY", pid, "CREATE", None, after)
    conn.commit()
    conn.close()
    return pid


def update_property(property_id: int, data: dict[str, Any]) -> None:
    conn = connect()
    before = get_property(property_id, include_deleted=True)
    if not before:
        conn.close()
        raise ValueError("Property not found")

    tab = data.get("tab", before.get("tab"))
    complex_name = data.get("complex_name") or TAB_COMPLEX_NAME.get(tab, before.get("complex_name", ""))

    conn.execute(
        """
        UPDATE properties SET
            tab=?,
            complex_name=?,
            unit_type=?,
            area=?,
            pyeong=?,
            address_detail=?,
            floor=?,
            total_floor=?,
            view=?,
            orientation=?,
            condition=?,
            repair_needed=?,
            tenant_info=?,
            naver_link=?,
            special_notes=?,
            note=?,
            status=?,
            updated_at=?
        WHERE id=?
        """,
        (
            tab,
            complex_name,
            data.get("unit_type", before.get("unit_type", "")),
            _to_float_or_none(data.get("area")) if "area" in data else before.get("area"),
            _to_float_or_none(data.get("pyeong")) if "pyeong" in data else before.get("pyeong"),
            data.get("address_detail", before.get("address_detail", "")),
            data.get("floor", before.get("floor", "")),
            data.get("total_floor", before.get("total_floor", "")),
            data.get("view", before.get("view", "")),
            data.get("orientation", before.get("orientation", "")),
            data.get("condition", before.get("condition", "")),
            1 if data.get("repair_needed") else 0,
            data.get("tenant_info", before.get("tenant_info", "")),
            data.get("naver_link", before.get("naver_link", "")),
            data.get("special_notes", before.get("special_notes", "")),
            data.get("note", before.get("note", "")),
            data.get("status", before.get("status", "신규등록")),
            _now_ts(),
            property_id,
        ),
    )

    after = get_property(property_id, include_deleted=True)
    _audit(conn, "PROPERTY", property_id, "UPDATE", before, after)
    conn.commit()
    conn.close()


def get_property(property_id: int, include_deleted: bool = False) -> dict[str, Any] | None:
    conn = connect()
    if include_deleted:
        row = conn.execute("SELECT * FROM properties WHERE id=?", (property_id,)).fetchone()
    else:
        row = conn.execute("SELECT * FROM properties WHERE id=? AND deleted=0", (property_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def list_properties(tab: str | None = None, *, include_deleted: bool = False) -> list[dict[str, Any]]:
    conn = connect()
    where = []
    params: list[Any] = []

    if tab:
        where.append("tab=?")
        params.append(tab)
    if not include_deleted:
        where.append("deleted=0")

    sql = "SELECT * FROM properties"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY hidden ASC, updated_at DESC, created_at DESC"

    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def toggle_property_hidden(property_id: int) -> None:
    conn = connect()
    before = get_property(property_id, include_deleted=True)
    conn.execute(
        "UPDATE properties SET hidden=CASE WHEN hidden=1 THEN 0 ELSE 1 END, updated_at=? WHERE id=?",
        (_now_ts(), property_id),
    )
    after = get_property(property_id, include_deleted=True)
    _audit(conn, "PROPERTY", property_id, "TOGGLE_HIDDEN", before, after)
    conn.commit()
    conn.close()


def soft_delete_property(property_id: int) -> None:
    conn = connect()
    before = get_property(property_id, include_deleted=True)
    conn.execute("UPDATE properties SET deleted=1, updated_at=? WHERE id=?", (_now_ts(), property_id))
    after = get_property(property_id, include_deleted=True)
    _audit(conn, "PROPERTY", property_id, "SOFT_DELETE", before, after)
    conn.commit()
    conn.close()


def restore_property(property_id: int) -> None:
    conn = connect()
    before = get_property(property_id, include_deleted=True)
    conn.execute("UPDATE properties SET deleted=0, updated_at=? WHERE id=?", (_now_ts(), property_id))
    after = get_property(property_id, include_deleted=True)
    _audit(conn, "PROPERTY", property_id, "RESTORE", before, after)
    conn.commit()
    conn.close()


# ----------------------------
# Customers
# ----------------------------
def add_customer(data: dict[str, Any]) -> int:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO customer_requests
        (customer_name, phone, preferred_tab, preferred_area, preferred_pyeong,
         budget, move_in_period, view_preference, location_preference,
         floor_preference, extra_needs, status, hidden, deleted, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?)
        """,
        (
            data.get("customer_name", ""),
            data.get("phone", ""),
            data.get("preferred_tab", ""),
            data.get("preferred_area", ""),
            data.get("preferred_pyeong", ""),
            data.get("budget", ""),
            data.get("move_in_period", ""),
            data.get("view_preference", ""),
            data.get("location_preference", ""),
            data.get("floor_preference", ""),
            data.get("extra_needs", ""),
            data.get("status", "진행"),
            1 if data.get("hidden") else 0,
            _now_ts(),
            _now_ts(),
        ),
    )
    cid = int(cur.lastrowid)
    after = get_customer(cid, include_deleted=True)
    _audit(conn, "CUSTOMER", cid, "CREATE", None, after)
    conn.commit()
    conn.close()
    return cid


def update_customer(customer_id: int, data: dict[str, Any]) -> None:
    conn = connect()
    before = get_customer(customer_id, include_deleted=True)
    if not before:
        conn.close()
        raise ValueError("Customer not found")

    conn.execute(
        """
        UPDATE customer_requests SET
            customer_name=?,
            phone=?,
            preferred_tab=?,
            preferred_area=?,
            preferred_pyeong=?,
            budget=?,
            move_in_period=?,
            view_preference=?,
            location_preference=?,
            floor_preference=?,
            extra_needs=?,
            status=?,
            updated_at=?
        WHERE id=?
        """,
        (
            data.get("customer_name", before.get("customer_name", "")),
            data.get("phone", before.get("phone", "")),
            data.get("preferred_tab", before.get("preferred_tab", "")),
            data.get("preferred_area", before.get("preferred_area", "")),
            data.get("preferred_pyeong", before.get("preferred_pyeong", "")),
            data.get("budget", before.get("budget", "")),
            data.get("move_in_period", before.get("move_in_period", "")),
            data.get("view_preference", before.get("view_preference", "")),
            data.get("location_preference", before.get("location_preference", "")),
            data.get("floor_preference", before.get("floor_preference", "")),
            data.get("extra_needs", before.get("extra_needs", "")),
            data.get("status", before.get("status", "진행")),
            _now_ts(),
            customer_id,
        ),
    )

    after = get_customer(customer_id, include_deleted=True)
    _audit(conn, "CUSTOMER", customer_id, "UPDATE", before, after)
    conn.commit()
    conn.close()


def get_customer(customer_id: int, include_deleted: bool = False) -> dict[str, Any] | None:
    conn = connect()
    if include_deleted:
        row = conn.execute("SELECT * FROM customer_requests WHERE id=?", (customer_id,)).fetchone()
    else:
        row = conn.execute("SELECT * FROM customer_requests WHERE id=? AND deleted=0", (customer_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def list_customers(*, include_deleted: bool = False) -> list[dict[str, Any]]:
    conn = connect()
    if include_deleted:
        rows = conn.execute("SELECT * FROM customer_requests ORDER BY hidden ASC, updated_at DESC").fetchall()
    else:
        rows = conn.execute(
            "SELECT * FROM customer_requests WHERE deleted=0 ORDER BY hidden ASC, updated_at DESC"
        ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def toggle_customer_hidden(customer_id: int) -> None:
    conn = connect()
    before = get_customer(customer_id, include_deleted=True)
    conn.execute(
        "UPDATE customer_requests SET hidden=CASE WHEN hidden=1 THEN 0 ELSE 1 END, updated_at=? WHERE id=?",
        (_now_ts(), customer_id),
    )
    after = get_customer(customer_id, include_deleted=True)
    _audit(conn, "CUSTOMER", customer_id, "TOGGLE_HIDDEN", before, after)
    conn.commit()
    conn.close()


def soft_delete_customer(customer_id: int) -> None:
    conn = connect()
    before = get_customer(customer_id, include_deleted=True)
    conn.execute("UPDATE customer_requests SET deleted=1, updated_at=? WHERE id=?", (_now_ts(), customer_id))
    after = get_customer(customer_id, include_deleted=True)
    _audit(conn, "CUSTOMER", customer_id, "SOFT_DELETE", before, after)
    conn.commit()
    conn.close()


def restore_customer(customer_id: int) -> None:
    conn = connect()
    before = get_customer(customer_id, include_deleted=True)
    conn.execute("UPDATE customer_requests SET deleted=0, updated_at=? WHERE id=?", (_now_ts(), customer_id))
    after = get_customer(customer_id, include_deleted=True)
    _audit(conn, "CUSTOMER", customer_id, "RESTORE", before, after)
    conn.commit()
    conn.close()


# ----------------------------
# Photos
# ----------------------------
def add_photo(property_id: int, file_path: str, tag: str = "") -> int:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO photos(property_id,file_path,tag,created_at) VALUES(?,?,?,?)",
        (property_id, file_path, tag, _now_ts()),
    )
    pid = int(cur.lastrowid)
    after = {"id": pid, "property_id": property_id, "file_path": file_path, "tag": tag}
    _audit(conn, "PHOTO", pid, "ADD", None, after)
    conn.commit()
    conn.close()
    return pid


def list_photos(property_id: int) -> list[dict[str, Any]]:
    conn = connect()
    rows = conn.execute(
        "SELECT * FROM photos WHERE property_id=? ORDER BY created_at DESC", (property_id,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def list_photos_all() -> list[dict[str, Any]]:
    conn = connect()
    rows = conn.execute("SELECT * FROM photos ORDER BY created_at DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def delete_photo(photo_id: int) -> None:
    """DB 레코드만 삭제(파일은 남김)."""
    conn = connect()
    row = conn.execute("SELECT * FROM photos WHERE id=?", (photo_id,)).fetchone()
    before = dict(row) if row else None
    conn.execute("DELETE FROM photos WHERE id=?", (photo_id,))
    _audit(conn, "PHOTO", photo_id, "DELETE", before, None)
    conn.commit()
    conn.close()


# ----------------------------
# Viewings (Schedules)
# ----------------------------
def add_viewing(
    *,
    property_id: int,
    customer_id: int | None,
    start_at: str,
    end_at: str,
    title: str,
    memo: str = "",
    status: str = "예정",
) -> int:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO viewings(property_id,customer_id,start_at,end_at,title,memo,status,created_at)
        VALUES(?,?,?,?,?,?,?,?)
        """,
        (property_id, customer_id, start_at, end_at, title, memo, status, _now_ts()),
    )
    vid = int(cur.lastrowid)
    after = get_viewing(vid)
    _audit(conn, "VIEWING", vid, "CREATE", None, after)
    conn.commit()
    conn.close()
    return vid


def get_viewing(viewing_id: int) -> dict[str, Any] | None:
    conn = connect()
    row = conn.execute("SELECT * FROM viewings WHERE id=?", (viewing_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def list_viewings(*, property_id: int | None = None, customer_id: int | None = None) -> list[dict[str, Any]]:
    conn = connect()
    where = []
    params: list[Any] = []
    if property_id is not None:
        where.append("property_id=?")
        params.append(property_id)
    if customer_id is not None:
        where.append("customer_id=?")
        params.append(customer_id)

    sql = "SELECT * FROM viewings"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY start_at DESC"
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def update_viewing_status(viewing_id: int, status: str) -> None:
    conn = connect()
    before = get_viewing(viewing_id)
    conn.execute("UPDATE viewings SET status=? WHERE id=?", (status, viewing_id))
    after = get_viewing(viewing_id)
    _audit(conn, "VIEWING", viewing_id, "STATUS_UPDATE", before, after)
    conn.commit()
    conn.close()


def delete_viewing(viewing_id: int) -> None:
    conn = connect()
    before = get_viewing(viewing_id)
    conn.execute("DELETE FROM viewings WHERE id=?", (viewing_id,))
    _audit(conn, "VIEWING", viewing_id, "DELETE", before, None)
    conn.commit()
    conn.close()



# ----------------------------
# Tasks (Next Actions)
# ----------------------------
def add_task(
    *,
    title: str,
    due_at: str | None = None,
    entity_type: str | None = None,
    entity_id: int | None = None,
    note: str = "",
    kind: str = "MANUAL",
    unique_key: str | None = None,
    status: str = "OPEN",
) -> int:
    """Create a task.

    Parameters
    ----------
    unique_key:
        If provided, task becomes idempotent(upsert-friendly). Useful for AUTO_* tasks.
    """
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO tasks(unique_key,kind,entity_type,entity_id,title,due_at,status,note,created_at,updated_at)
        VALUES(?,?,?,?,?,?,?,?,?,?)
        """,
        (
            unique_key,
            kind,
            entity_type,
            entity_id,
            title,
            due_at,
            status,
            note,
            _now_ts(),
            _now_ts(),
        ),
    )
    tid = int(cur.lastrowid)
    after = get_task(tid)
    _audit(conn, "TASK", tid, "CREATE", None, after)
    conn.commit()
    conn.close()
    return tid


def upsert_task_unique(
    *,
    unique_key: str,
    title: str,
    due_at: str | None = None,
    entity_type: str | None = None,
    entity_id: int | None = None,
    note: str = "",
    kind: str = "AUTO",
    status: str = "OPEN",
) -> int:
    """Upsert a task by unique_key (used for AUTO tasks)."""
    conn = connect()
    before = conn.execute("SELECT * FROM tasks WHERE unique_key=?", (unique_key,)).fetchone()
    before_obj = dict(before) if before else None

    conn.execute(
        """
        INSERT INTO tasks(unique_key,kind,entity_type,entity_id,title,due_at,status,note,created_at,updated_at)
        VALUES(?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(unique_key) DO UPDATE SET
            kind=excluded.kind,
            entity_type=excluded.entity_type,
            entity_id=excluded.entity_id,
            title=excluded.title,
            due_at=excluded.due_at,
            status=excluded.status,
            note=excluded.note,
            updated_at=excluded.updated_at
        """,
        (
            unique_key,
            kind,
            entity_type,
            entity_id,
            title,
            due_at,
            status,
            note,
            _now_ts(),
            _now_ts(),
        ),
    )

    row = conn.execute("SELECT id FROM tasks WHERE unique_key=?", (unique_key,)).fetchone()
    task_id = int(row["id"]) if row else -1
    after = get_task(task_id) if task_id != -1 else None
    _audit(conn, "TASK", task_id, "UPSERT", before_obj, after)
    conn.commit()
    conn.close()
    return task_id


def get_task(task_id: int) -> dict[str, Any] | None:
    conn = connect()
    row = conn.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def list_tasks(
    *,
    include_done: bool = False,
    kind_prefix: str | None = None,
    entity_type: str | None = None,
    entity_id: int | None = None,
    limit: int | None = None,
) -> list[dict[str, Any]]:
    conn = connect()
    where = []
    params: list[Any] = []

    if not include_done:
        where.append("status='OPEN'")
    if kind_prefix:
        where.append("kind LIKE ?")
        params.append(f"{kind_prefix}%")
    if entity_type:
        where.append("entity_type=?")
        params.append(entity_type)
    if entity_id is not None:
        where.append("entity_id=?")
        params.append(entity_id)

    sql = "SELECT * FROM tasks"
    if where:
        sql += " WHERE " + " AND ".join(where)

    # due_at이 있는 것이 먼저, 그 다음 오래된 것
    sql += " ORDER BY CASE WHEN due_at IS NULL OR due_at='' THEN 1 ELSE 0 END, due_at ASC, created_at ASC"
    if limit:
        sql += f" LIMIT {int(limit)}"

    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def set_task_status(task_id: int, status: str) -> None:
    conn = connect()
    before = get_task(task_id)
    conn.execute("UPDATE tasks SET status=?, updated_at=? WHERE id=?", (status, _now_ts(), task_id))
    after = get_task(task_id)
    _audit(conn, "TASK", task_id, f"STATUS_{status}", before, after)
    conn.commit()
    conn.close()


def delete_task(task_id: int) -> None:
    conn = connect()
    before = get_task(task_id)
    conn.execute("DELETE FROM tasks WHERE id=?", (task_id,))
    _audit(conn, "TASK", task_id, "DELETE", before, None)
    conn.commit()
    conn.close()


def list_auto_tasks(*, include_done: bool = False) -> list[dict[str, Any]]:
    return list_tasks(include_done=include_done, kind_prefix="AUTO_")


def mark_task_done(task_id: int) -> None:
    set_task_status(task_id, "DONE")

# ----------------------------
# Helpers
# ----------------------------
def _to_float_or_none(v: Any) -> float | None:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return None
    try:
        return float(s)
    except Exception:
        return None
