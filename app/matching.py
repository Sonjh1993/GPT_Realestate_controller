"""Rule-based matching (Customer ↔ Property).

목표
----
AI 없이도 "고객 요구사항 → 후보 물건"을 즉시 뽑아주는 생산성 기능.
Gemini/LLM은 나중에 "요약"을 담당하도록 분리.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Iterable

from . import money_utils


def _to_text(v: Any) -> str:
    return "" if v is None else str(v)


def _parse_deal_types(text: str) -> set[str]:
    allowed = {"매매", "전세", "월세"}
    return {token.strip() for token in _to_text(text).split(",") if token.strip() in allowed}


def parse_range(text: str) -> tuple[float | None, float | None]:
    """Parse simple range like "80~90", "80-90", "84".

    Returns (min, max).
    """
    s = (text or "").strip().replace("㎡", "").replace("평", "")
    if not s:
        return (None, None)

    # 80~90, 80-90
    m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*[~\-]\s*(\d+(?:\.\d+)?)\s*$", s)
    if m:
        return (float(m.group(1)), float(m.group(2)))

    # single
    m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*$", s)
    if m:
        v = float(m.group(1))
        return (v, v)

    return (None, None)


def parse_floor(text: str) -> int | None:
    """Try to parse floor number from strings like "15", "15/25", "15층"."""
    s = (text or "").strip()
    if not s:
        return None
    m = re.search(r"(\d+)", s)
    return int(m.group(1)) if m else None


@dataclass
class MatchResult:
    property_id: int
    score: int
    reasons: list[str]
    property_row: dict[str, Any]


def match_properties(customer: dict[str, Any], properties: Iterable[dict[str, Any]], *, limit: int = 30) -> list[MatchResult]:
    """Return ranked list of properties for a customer.

    Notes
    -----
    - hidden/deleted는 호출자가 이미 걸러준다고 가정.
    - 점수는 단순 규칙 기반(설명 가능성 > 복잡도).
    """
    pref_tab_text = _to_text(customer.get("preferred_tab")).strip()
    pref_tabs = {t.strip() for t in pref_tab_text.split(",") if t.strip()}
    pref_area = _to_text(customer.get("preferred_area")).strip()
    pref_pyeong = _to_text(customer.get("preferred_pyeong")).strip()
    floor_pref = _to_text(customer.get("floor_preference")).strip()
    view_pref = _to_text(customer.get("view_preference")).strip()
    loc_pref = _to_text(customer.get("location_preference")).strip()
    deal_types = _parse_deal_types(_to_text(customer.get("deal_type")))
    budget_10m = money_utils.to_int(customer.get("budget_10m"), 0)
    rent_budget_10man = money_utils.to_int(customer.get("wolse_rent_10man"), 0)

    a_min, a_max = parse_range(pref_area)
    p_min, p_max = parse_range(pref_pyeong)

    results: list[MatchResult] = []

    for p in properties:
        # 탭 필터
        prop_tab = _to_text(p.get("tab")).strip()
        if pref_tabs and prop_tab and prop_tab not in pref_tabs:
            continue

        property_deal_types = {
            deal for deal, enabled in (("매매", p.get("deal_sale")), ("전세", p.get("deal_jeonse")), ("월세", p.get("deal_wolse"))) if enabled
        }
        deal_overlap = deal_types & property_deal_types
        if deal_types and not deal_overlap:
            continue

        score = 0
        reasons: list[str] = []
        if deal_overlap:
            reasons.append("거래유형 일치")
            score += 15

        # 면적/평형
        area = p.get("area")
        pyeong = p.get("pyeong")
        if a_min is not None and a_max is not None and isinstance(area, (int, float)):
            if a_min <= float(area) <= a_max:
                score += 30
                reasons.append("면적 범위 일치")
            else:
                # 멀어질수록 감점(완만)
                dist = min(abs(float(area) - a_min), abs(float(area) - a_max))
                score -= int(min(15, dist / 2))
        else:
            score += 5

        if p_min is not None and p_max is not None and isinstance(pyeong, (int, float)):
            if p_min <= float(pyeong) <= p_max:
                score += 25
                reasons.append("평형 범위 일치")
            else:
                dist = min(abs(float(pyeong) - p_min), abs(float(pyeong) - p_max))
                score -= int(min(12, dist))
        else:
            score += 5

        # 층수 선호
        fl = parse_floor(_to_text(p.get("floor")))
        if floor_pref:
            if "고" in floor_pref and fl is not None and fl >= 15:
                score += 10
                reasons.append("고층 선호")
            elif "저" in floor_pref and fl is not None and fl <= 5:
                score += 10
                reasons.append("저층 선호")
            elif "무관" in floor_pref:
                score += 3

        # 뷰/위치 키워드
        hay = " ".join(
            [
                _to_text(p.get("view")),
                _to_text(p.get("special_notes")),
                _to_text(p.get("address_detail")),
                _to_text(p.get("note")),
            ]
        ).lower()
        if view_pref and view_pref.lower() in hay:
            score += 6
            reasons.append("뷰 키워드")
        if loc_pref and loc_pref.lower() in hay:
            score += 6
            reasons.append("위치 키워드")

        # 컨디션 가중치(상/중/하)
        cond = _to_text(p.get("condition")).strip()
        if cond == "상":
            score += 4
        elif cond == "중":
            score += 2

        if (deal_overlap & {"매매", "전세"}) and budget_10m > 0:
            candidates: list[int] = []
            if "매매" in deal_overlap:
                candidates.append(money_utils.eok_che_to_ten_million(p.get("price_sale_eok"), p.get("price_sale_che")))
            if "전세" in deal_overlap:
                candidates.append(money_utils.eok_che_to_ten_million(p.get("price_jeonse_eok"), p.get("price_jeonse_che")))
            if candidates:
                price_10m = min(candidates)
                if price_10m <= budget_10m:
                    score += 18
                    reasons.append("예산 범위 일치")
                else:
                    score -= min(20, max(0, price_10m - budget_10m) // 2)
        elif "월세" in deal_overlap and rent_budget_10man > 0:
            p_rent_10man = money_utils.man_to_ten_man(p.get("wolse_rent_man"))
            if p_rent_10man <= rent_budget_10man:
                score += 16
                reasons.append("월세 예산 범위 일치")
            else:
                score -= min(20, max(0, p_rent_10man - rent_budget_10man))

        results.append(MatchResult(property_id=int(p.get("id")), score=score, reasons=reasons, property_row=p))

    results.sort(key=lambda x: x.score, reverse=True)
    return results[:limit]
