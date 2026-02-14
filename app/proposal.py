"""Proposal (매물 제안서) generator.

요구사항
--------
- Google API 없어도 동작해야 함.
- 고객에게 '전문적으로' 보이는 출력물(PDF) 1클릭 생성.
- Drive 동기화 폴더에 저장되면 어디서나 열람 가능.

구현 전략
---------
- PDF는 reportlab 사용(없으면 안내 메시지).
- 한글은 reportlab CID 폰트(HYSMyeongJo-Medium)로 처리.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

# reportlab는 선택 의존성 (PDF 생성 시에만 필요)
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.platypus import (
        Image,
        PageBreak,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.lib.utils import ImageReader

    _REPORTLAB_OK = True
except Exception:  # pragma: no cover
    _REPORTLAB_OK = False


def _ensure_reportlab() -> None:
    if not _REPORTLAB_OK:
        raise RuntimeError("PDF 기능을 사용하려면 'reportlab' 설치가 필요합니다.\n예) pip install reportlab")


def _safe_filename(name: str) -> str:
    s = (name or "proposal").strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^0-9A-Za-z_\-가-힣]", "", s)
    return s[:80] if s else "proposal"


def _yn(v: Any) -> str:
    if v in (1, True, "1", "Y", "y", "yes", "Yes", "예"):
        return "Y"
    if v in (0, False, "0", "N", "n", "no", "No", "아니오"):
        return "N"
    return ""


def build_kakao_message(customer: dict[str, Any], properties: list[dict[str, Any]], *, include_links: bool = True) -> str:
    cname = str(customer.get("customer_name") or "").strip()
    phone = str(customer.get("phone") or "").strip()
    header = f"안녕하세요{(' ' + cname + '님') if cname else ''}.\n요청 조건 기준 추천 매물 보내드립니다.\n"
    if phone:
        header += f"(연락처: {phone})\n"
    lines = [header]

    for i, p in enumerate(properties, start=1):
        complex_name = str(p.get("complex_name") or "").strip()
        addr = str(p.get("address_detail") or "").strip()
        unit = str(p.get("unit_type") or "").strip()
        floor = str(p.get("floor") or "").strip()
        total_floor = str(p.get("total_floor") or "").strip()
        cond = str(p.get("condition") or "").strip()
        view = str(p.get("view") or "").strip()
        ori = str(p.get("orientation") or "").strip()
        note = str(p.get("special_notes") or "").strip()
        link = str(p.get("naver_link") or "").strip()

        title = " ".join([x for x in [complex_name, addr, unit] if x]).strip()
        if not title:
            title = f"물건ID {p.get('id')}"

        detail = []
        if floor or total_floor:
            detail.append(f"층: {floor}/{total_floor}".strip("/"))
        if cond:
            detail.append(f"컨디션:{cond}")
        if ori or view:
            ov = " / ".join([x for x in [ori, view] if x])
            detail.append(f"향/뷰:{ov}")
        if note:
            detail.append(f"특이:{note}")

        lines.append(f"{i}) {title}")
        if detail:
            lines.append("   - " + " | ".join(detail))
        if include_links and link:
            lines.append(f"   - 링크: {link}")
        lines.append("")  # blank line

    lines.append("원하시면 조건(층/뷰/예산/기간) 조금 더 구체화해서 더 정확히 추려드릴게요.")
    return "\n".join(lines).strip()


@dataclass
class ProposalOutput:
    pdf_path: Path
    txt_path: Path


def generate_proposal_pdf(
    *,
    customer: dict[str, Any],
    properties: list[dict[str, Any]],
    photos_by_property: dict[int, list[dict[str, str] | str]] | None,
    output_dir: Path,
    title: str = "매물 제안서",
    max_photos_per_property: int = 4,
) -> ProposalOutput:
    """Generate proposal PDF + txt message into output_dir."""
    _ensure_reportlab()
    output_dir.mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    cname = _safe_filename(str(customer.get("customer_name") or "고객"))
    base = f"proposal_{cname}_{ts}"

    pdf_path = output_dir / f"{base}.pdf"
    txt_path = output_dir / f"{base}.txt"

    # 1) Text message for Kakao/문자
    txt = build_kakao_message(customer, properties, include_links=True)
    txt_path.write_text(txt, encoding="utf-8")

    # 2) PDF
    # Register Korean CID fonts
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("HYSMyeongJo-Medium"))
    except Exception:
        pass
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("HYGothic-Medium"))
    except Exception:
        pass

    styles = getSampleStyleSheet()
    style_title = ParagraphStyle(
        "title",
        parent=styles["Title"],
        fontName="HYGothic-Medium" if "HYGothic-Medium" in pdfmetrics.getRegisteredFontNames() else "HYSMyeongJo-Medium",
        fontSize=18,
        leading=22,
        spaceAfter=10,
    )
    style_h2 = ParagraphStyle(
        "h2",
        parent=styles["Heading2"],
        fontName="HYGothic-Medium" if "HYGothic-Medium" in pdfmetrics.getRegisteredFontNames() else "HYSMyeongJo-Medium",
        fontSize=12.5,
        leading=16,
        spaceBefore=8,
        spaceAfter=4,
    )
    style_body = ParagraphStyle(
        "body",
        parent=styles["BodyText"],
        fontName="HYSMyeongJo-Medium" if "HYSMyeongJo-Medium" in pdfmetrics.getRegisteredFontNames() else "Helvetica",
        fontSize=10.5,
        leading=14,
    )
    style_small = ParagraphStyle(
        "small",
        parent=styles["BodyText"],
        fontName="HYSMyeongJo-Medium" if "HYSMyeongJo-Medium" in pdfmetrics.getRegisteredFontNames() else "Helvetica",
        fontSize=9.5,
        leading=12,
        textColor=colors.grey,
    )

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
    )

    elements: list[Any] = []
    elements.append(Paragraph(title, style_title))

    cname_raw = str(customer.get("customer_name") or "").strip()
    phone = str(customer.get("phone") or "").strip()
    meta = f"고객: {cname_raw}" + (f" / {phone}" if phone else "")
    elements.append(Paragraph(meta, style_body))
    elements.append(Paragraph(f"생성일: {datetime.now().strftime('%Y-%m-%d %H:%M')}", style_small))
    elements.append(Spacer(1, 8))

    photos_by_property = photos_by_property or {}

    # helper: scale image to fit box
    def build_image(path: str, max_w_mm: float = 44, max_h_mm: float = 36):
        try:
            reader = ImageReader(path)
            w, h = reader.getSize()
            max_w = max_w_mm * mm
            max_h = max_h_mm * mm
            if w <= 0 or h <= 0:
                return Spacer(1, 1)
            scale = min(max_w / w, max_h / h)
            iw, ih = w * scale, h * scale
            img = Image(path, width=iw, height=ih)
            return img
        except Exception:
            return Spacer(1, 1)

    for idx, p in enumerate(properties, start=1):
        pid = int(p.get("id") or 0)
        complex_name = str(p.get("complex_name") or "").strip()
        addr = str(p.get("address_detail") or "").strip()
        unit = str(p.get("unit_type") or "").strip()

        title_line = " ".join([x for x in [complex_name, addr, unit] if x]).strip()
        if not title_line:
            title_line = f"물건ID {pid}"

        elements.append(Paragraph(f"{idx}. {title_line}", style_h2))

        info_lines = []
        area = p.get("area")
        pyeong = p.get("pyeong")
        if area:
            info_lines.append(f"면적: {area}㎡" + (f" ({pyeong}평)" if pyeong else ""))
        if p.get("floor") or p.get("total_floor"):
            info_lines.append(f"층: {p.get('floor','')}/{p.get('total_floor','')}".strip("/"))
        if p.get("orientation") or p.get("view"):
            ov = " / ".join([str(x).strip() for x in [p.get("orientation"), p.get("view")] if str(x or "").strip()])
            info_lines.append(f"향/뷰: {ov}")
        if p.get("condition") or p.get("repair_needed") is not None:
            rn = _yn(p.get("repair_needed"))
            cn = str(p.get("condition") or "").strip()
            part = " / ".join([x for x in [f"컨디션:{cn}" if cn else "", f"수리:{rn}" if rn else ""] if x])
            if part:
                info_lines.append(part)
        if p.get("special_notes"):
            info_lines.append(f"특이: {p.get('special_notes')}")
        if p.get("naver_link"):
            info_lines.append(f"링크: {p.get('naver_link')}")

        detail_text = "<br/>".join([str(x).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;") for x in info_lines])
        elements.append(Paragraph(detail_text, style_body))
        elements.append(Spacer(1, 4))

        # 사진 우선순위: 거실 > 안방 > 작은방 > 화장실 > 주방 > 나머지
        raw_photos = photos_by_property.get(pid, [])
        normalized: list[dict[str, str]] = []
        for it in raw_photos:
            if isinstance(it, dict):
                fp = str(it.get("file_path") or "").strip()
                tg = str(it.get("tag") or "").strip()
            else:
                fp = str(it or "").strip()
                tg = ""
            if fp:
                normalized.append({"file_path": fp, "tag": tg})

        priority = {"거실": 0, "안방": 1, "작은방": 2, "화장실": 3, "주방": 4}
        normalized.sort(key=lambda x: (priority.get(x.get("tag", ""), 99), x.get("tag", "")))
        chosen = normalized[:max_photos_per_property]

        if chosen:
            img_cells = [build_image(str(it.get("file_path") or "")) for it in chosen]
            cap_cells = [Paragraph(str(it.get("tag") or "사진"), style_small) for it in chosen]
            img_table = Table([img_cells, cap_cells], colWidths=[45 * mm] * len(img_cells))
            img_table.setStyle(
                TableStyle(
                    [
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 3),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                        ("TOPPADDING", (0, 0), (-1, -1), 2),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                    ]
                )
            )
            elements.append(img_table)
            elements.append(Spacer(1, 8))
        else:
            elements.append(Spacer(1, 6))

        # page break every ~3 items for readability
        if idx % 3 == 0 and idx != len(properties):
            elements.append(PageBreak())

    doc.build(elements)

    return ProposalOutput(pdf_path=pdf_path, txt_path=txt_path)
