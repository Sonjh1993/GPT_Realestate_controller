from __future__ import annotations

import csv
from functools import lru_cache
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent.parent

CSV_BY_TAB = {
    "봉담자이 프라이드시티": BASE_DIR / "자이.csv",
    "힐스테이트봉담프라이드시티": BASE_DIR / "힐스.csv",
}


def _to_int(value: Any) -> int | None:
    try:
        text = str(value).strip()
        if not text:
            return None
        return int(float(text))
    except Exception:
        return None


def _sort_numeric(values: list[str]) -> list[str]:
    def key(v: str):
        n = _to_int("".join(ch for ch in v if ch.isdigit()))
        return (n is None, n if n is not None else v)

    return sorted(values, key=key)


@lru_cache(maxsize=8)
def _load_rows(complex_tab: str) -> list[dict[str, str]]:
    csv_path = CSV_BY_TAB.get(complex_tab)
    if not csv_path or not csv_path.exists():
        return []

    last_error: Exception | None = None
    for enc in ("utf-8-sig", "cp949"):
        try:
            with csv_path.open("r", encoding=enc, newline="") as f:
                return list(csv.DictReader(f))
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"CSV 로드 실패: {csv_path} ({last_error})")


def get_dongs(complex_tab: str) -> list[str]:
    rows = _load_rows(complex_tab)
    dongs = {str(r.get("dong") or "").strip() for r in rows if str(r.get("dong") or "").strip()}
    return _sort_numeric(list(dongs))


def get_floors(complex_tab: str, dong: str) -> list[int]:
    rows = _load_rows(complex_tab)
    floors = {
        _to_int(r.get("floor"))
        for r in rows
        if str(r.get("dong") or "").strip() == str(dong).strip()
    }
    return sorted(f for f in floors if f is not None)


def get_hos(complex_tab: str, dong: str, floor: int) -> list[str]:
    rows = _load_rows(complex_tab)
    hos = {
        str(r.get("ho") or "").strip()
        for r in rows
        if str(r.get("dong") or "").strip() == str(dong).strip()
        and _to_int(r.get("floor")) == _to_int(floor)
        and str(r.get("ho") or "").strip()
    }
    return _sort_numeric(list(hos))


def get_unit_info(complex_tab: str, dong: str, floor: int, ho: str | int) -> dict[str, float | str]:
    rows = _load_rows(complex_tab)
    for r in rows:
        if (
            str(r.get("dong") or "").strip() == str(dong).strip()
            and _to_int(r.get("floor")) == _to_int(floor)
            and str(r.get("ho") or "").strip() == str(ho).strip()
        ):
            unit_type = str(r.get("type") or "").strip()
            supply_m2 = float(r.get("supply_m2") or 0)
            pyeong = float(r.get("pyeong") or 0)
            return {"type": unit_type, "supply_m2": supply_m2, "pyeong": pyeong}
    return {"type": "", "supply_m2": 0.0, "pyeong": 0.0}


def get_total_floor(complex_tab: str, dong: str) -> int:
    rows = _load_rows(complex_tab)
    floors = [
        _to_int(r.get("floor"))
        for r in rows
        if str(r.get("dong") or "").strip() == str(dong).strip()
    ]
    valid = [f for f in floors if f is not None]
    return max(valid) if valid else 0


def has_master(complex_tab: str) -> bool:
    return complex_tab in CSV_BY_TAB
