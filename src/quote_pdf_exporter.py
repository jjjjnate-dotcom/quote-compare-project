from __future__ import annotations

from dataclasses import dataclass
from datetime import date
import math
from pathlib import Path

from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle

FONT_KR = "HYSMyeongJo-Medium"


@dataclass
class QuoteLine:
    name: str
    qty: float
    unit_price: float
    supply_amount: float
    vat_amount: float


def _to_float(value) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _round_to_hundred_half_up(value: float) -> float:
    factor = 100.0
    if value >= 0:
        return math.floor(value / factor + 0.5) * factor
    return -math.floor(abs(value) / factor + 0.5) * factor


def _money(value: float) -> str:
    return f"{int(round(value)):,}"


def _load_source_items(source_quote_path: Path) -> list[tuple[str, float, float]]:
    wb = load_workbook(source_quote_path, data_only=True)
    ws = wb.worksheets[0]
    items: list[tuple[str, float, float]] = []
    blank_streak = 0

    for row in range(2, ws.max_row + 1):
        name = ws.cell(row, 1).value
        qty = ws.cell(row, 2).value
        price = ws.cell(row, 3).value

        if name in (None, "") and qty in (None, "") and price in (None, ""):
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue

        blank_streak = 0
        qty_num = _to_float(qty)
        price_num = _to_float(price)
        name_text = str(name).strip() if name is not None else ""
        if not name_text or qty_num is None or price_num is None:
            continue
        if qty_num <= 0 or price_num <= 0:
            continue

        items.append((name_text, qty_num, price_num))

    return items


def _build_company_lines(source_quote_path: Path, rate: float, vat_rate: float) -> list[QuoteLine]:
    source_items = _load_source_items(source_quote_path)
    lines: list[QuoteLine] = []
    for name, qty, base_price in source_items:
        unit_price = _round_to_hundred_half_up(base_price * (1 + rate))
        supply = unit_price * qty
        vat = supply * vat_rate
        lines.append(
            QuoteLine(
                name=name,
                qty=qty,
                unit_price=unit_price,
                supply_amount=supply,
                vat_amount=vat,
            )
        )
    return lines


def export_company_quote_pdf(
    source_quote_path: Path,
    output_pdf_path: Path,
    company_name: str,
    rate: float,
    vat_rate: float,
) -> Path:
    # Keep margins close to the original quotation PDF layout.
    left_margin = 18 * mm
    right_margin = 18 * mm
    top_margin = 18 * mm
    bottom_margin = 18 * mm

    pdfmetrics.registerFont(UnicodeCIDFont(FONT_KR))

    lines = _build_company_lines(source_quote_path, rate, vat_rate)
    total_supply = sum(line.supply_amount for line in lines)
    total_vat = sum(line.vat_amount for line in lines)
    total_amount = total_supply + total_vat

    output_pdf_path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(output_pdf_path), pagesize=A4)
    page_width, page_height = A4
    content_width = page_width - left_margin - right_margin

    y = page_height - top_margin
    c.setFont(FONT_KR, 20)
    c.drawCentredString(page_width / 2, y, "견적서")
    y -= 12 * mm

    c.setFont(FONT_KR, 10)
    today = date.today().isoformat()
    c.drawString(left_margin, y, f"수신: 발주처")
    c.drawString(left_margin + 75 * mm, y, f"공급사: {company_name}")
    c.drawRightString(page_width - right_margin, y, f"견적일자: {today}")
    y -= 6 * mm
    c.drawString(left_margin, y, f"가산율/할인율: {rate * 100:.1f}%")
    c.drawString(left_margin + 75 * mm, y, f"부가세율: {vat_rate * 100:.1f}%")
    y -= 8 * mm

    table_data = [["No", "품목명", "수량", "단가", "공급가액", "부가세"]]
    for idx, line in enumerate(lines, start=1):
        qty_text = int(line.qty) if float(line.qty).is_integer() else line.qty
        table_data.append(
            [
                str(idx),
                line.name,
                str(qty_text),
                _money(line.unit_price),
                _money(line.supply_amount),
                _money(line.vat_amount),
            ]
        )
    table_data.append(["", "합계", "", "", _money(total_supply), _money(total_vat)])
    table_data.append(["", "총액(공급가액+부가세)", "", "", "", _money(total_amount)])

    col_widths = [12 * mm, 86 * mm, 18 * mm, 28 * mm, 28 * mm, 22 * mm]
    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), FONT_KR),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f3f6fb")),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),
                ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -3), 0.4, colors.HexColor("#b7c5db")),
                ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#7f8ea8")),
                ("LINEABOVE", (0, -2), (-1, -2), 0.8, colors.HexColor("#7f8ea8")),
                ("SPAN", (1, -1), (4, -1)),
                ("ALIGN", (1, -1), (4, -1), "RIGHT"),
            ]
        )
    )

    table_width, table_height = table.wrap(content_width, page_height)
    if y - table_height < bottom_margin:
        c.showPage()
        y = page_height - top_margin
    table.drawOn(c, left_margin, y - table_height)

    c.setFont(FONT_KR, 9)
    c.drawString(left_margin, bottom_margin - 2 * mm, "※ 본 문서는 비교견적 자동 생성 시스템에서 생성되었습니다.")
    c.save()
    return output_pdf_path
