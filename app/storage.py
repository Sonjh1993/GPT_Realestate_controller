import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "ledger.db"

PROPERTY_TABS = ["아파트단지1", "아파트단지2", "상가", "단독주택"]


def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    conn = connect()
    cur = conn.cursor()
    cur.executescript(
        """
        CREATE TABLE IF NOT EXISTS properties (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tab TEXT NOT NULL,
            complex_name TEXT,
            unit_type TEXT,
            area REAL,
            pyeong REAL,
            floor TEXT,
            condition TEXT,
            repair_needed INTEGER DEFAULT 0,
            tenant_info TEXT,
            naver_link TEXT,
            special_notes TEXT,
            note TEXT,
            hidden INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
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
            hidden INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        """
    )
    conn.commit()
    conn.close()


def add_property(data: dict) -> None:
    conn = connect()
    conn.execute(
        """
        INSERT INTO properties
        (tab, complex_name, unit_type, area, pyeong, floor, condition, repair_needed,
         tenant_info, naver_link, special_notes, note)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            data.get("tab", ""),
            data.get("complex_name", ""),
            data.get("unit_type", ""),
            data.get("area") or None,
            data.get("pyeong") or None,
            data.get("floor", ""),
            data.get("condition", ""),
            1 if data.get("repair_needed") else 0,
            data.get("tenant_info", ""),
            data.get("naver_link", ""),
            data.get("special_notes", ""),
            data.get("note", ""),
        ),
    )
    conn.commit()
    conn.close()


def list_properties(tab: str | None = None):
    conn = connect()
    if tab:
        rows = conn.execute(
            "SELECT * FROM properties WHERE tab=? ORDER BY hidden ASC, created_at DESC", (tab,)
        ).fetchall()
    else:
        rows = conn.execute("SELECT * FROM properties ORDER BY hidden ASC, created_at DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def toggle_property_hidden(property_id: int) -> None:
    conn = connect()
    conn.execute(
        "UPDATE properties SET hidden=CASE WHEN hidden=1 THEN 0 ELSE 1 END WHERE id=?", (property_id,)
    )
    conn.commit()
    conn.close()


def delete_property(property_id: int) -> None:
    conn = connect()
    conn.execute("DELETE FROM properties WHERE id=?", (property_id,))
    conn.commit()
    conn.close()


def add_customer(data: dict) -> None:
    conn = connect()
    conn.execute(
        """
        INSERT INTO customer_requests
        (customer_name, phone, preferred_tab, preferred_area, preferred_pyeong,
         budget, move_in_period, view_preference, location_preference,
         floor_preference, extra_needs)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
        ),
    )
    conn.commit()
    conn.close()


def list_customers():
    conn = connect()
    rows = conn.execute(
        "SELECT * FROM customer_requests ORDER BY hidden ASC, created_at DESC"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def toggle_customer_hidden(customer_id: int) -> None:
    conn = connect()
    conn.execute(
        "UPDATE customer_requests SET hidden=CASE WHEN hidden=1 THEN 0 ELSE 1 END WHERE id=?", (customer_id,)
    )
    conn.commit()
    conn.close()


def delete_customer(customer_id: int) -> None:
    conn = connect()
    conn.execute("DELETE FROM customer_requests WHERE id=?", (customer_id,))
    conn.commit()
    conn.close()
