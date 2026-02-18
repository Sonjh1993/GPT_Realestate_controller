from __future__ import annotations

from typing import Any


def to_int(text: Any, default: int = 0) -> int:
    try:
        s = str(text).strip()
        if s == "":
            return default
        return int(float(s))
    except Exception:
        return default


def ten_million_to_eok_che(n_10m: Any) -> tuple[int, int]:
    n = max(0, to_int(n_10m, 0))
    return n // 10, n % 10


def eok_che_to_ten_million(eok: Any, che: Any) -> int:
    return max(0, to_int(eok, 0) * 10 + to_int(che, 0))


def ten_man_to_man(n_10man: Any) -> int:
    return max(0, to_int(n_10man, 0) * 10)


def man_to_ten_man(man: Any) -> int:
    return max(0, to_int(man, 0) // 10)


def fmt_10m(n_10m: Any) -> str:
    return f"{max(0, to_int(n_10m, 0))}천만원"


def fmt_10man(n_10man: Any) -> str:
    # 내부 저장은 10만원 단위이지만, 사용자 표기는 만원 단위로 보여준다.
    return f"{ten_man_to_man(n_10man)}만원"


def property_price_summary(row: dict[str, Any]) -> str:
    parts: list[str] = []
    if row.get("deal_sale"):
        n10 = eok_che_to_ten_million(row.get("price_sale_eok"), row.get("price_sale_che"))
        parts.append(f"매매 {fmt_10m(n10)}")
    if row.get("deal_jeonse"):
        n10 = eok_che_to_ten_million(row.get("price_jeonse_eok"), row.get("price_jeonse_che"))
        parts.append(f"전세 {fmt_10m(n10)}")
    if row.get("deal_wolse"):
        dep10 = eok_che_to_ten_million(row.get("wolse_deposit_eok"), row.get("wolse_deposit_che"))
        rent10 = man_to_ten_man(row.get("wolse_rent_man"))
        parts.append(f"월세 {fmt_10m(dep10)} / {fmt_10man(rent10)}")
    return " / ".join(parts)
