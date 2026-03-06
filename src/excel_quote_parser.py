from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

from openpyxl import Workbook, load_workbook


class ExcelQuoteParseError(Exception):
    pass


@dataclass
class QuoteItem:
    name: str
    qty: float
    unit_price: float


def _normalize_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip().replace(" ", "")


def _to_number(value) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = re.sub(r"[^0-9.\-]", "", text.replace(",", ""))
    if text in {"", "-", ".", "-."}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _is_total_row(name_value) -> bool:
    text = _normalize_text(name_value)
    return any(k in text for k in ("총합계", "합계", "계"))


def _find_header_columns(ws) -> tuple[int, int, int, int] | None:
    max_row = min(ws.max_row, 80)
    max_col = min(ws.max_column, 30)

    for r in range(1, max_row + 1):
        name_col = qty_col = price_col = None
        for c in range(1, max_col + 1):
            text = _normalize_text(ws.cell(r, c).value)
            if not text:
                continue

            if name_col is None and ("품목명" in text or "품명" in text):
                name_col = c
            if qty_col is None and "수량" in text:
                qty_col = c
            if price_col is None and "단가" in text:
                price_col = c

        if name_col and qty_col and price_col:
            return r, name_col, qty_col, price_col

    return None


def _extract_from_header_table(ws) -> list[QuoteItem]:
    header = _find_header_columns(ws)
    if not header:
        return []

    header_row, name_col, qty_col, price_col = header
    items: list[QuoteItem] = []
    blank_streak = 0

    for r in range(header_row + 1, ws.max_row + 1):
        name = ws.cell(r, name_col).value
        qty = ws.cell(r, qty_col).value
        unit_price = ws.cell(r, price_col).value

        if _is_total_row(name):
            break

        if name in (None, "") and qty in (None, "") and unit_price in (None, ""):
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue

        blank_streak = 0
        qty_num = _to_number(qty)
        price_num = _to_number(unit_price)
        name_text = str(name).strip() if name is not None else ""

        if not name_text or qty_num is None or price_num is None:
            continue
        if qty_num <= 0 or price_num <= 0:
            continue

        items.append(QuoteItem(name=name_text, qty=qty_num, unit_price=price_num))

    return items


def _extract_from_simple_columns(ws) -> list[QuoteItem]:
    items: list[QuoteItem] = []
    blank_streak = 0

    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 1).value
        qty = ws.cell(r, 2).value
        unit_price = ws.cell(r, 3).value

        if name in (None, "") and qty in (None, "") and unit_price in (None, ""):
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue

        blank_streak = 0
        if _is_total_row(name):
            break

        qty_num = _to_number(qty)
        price_num = _to_number(unit_price)
        name_text = str(name).strip() if name is not None else ""

        if not name_text or qty_num is None or price_num is None:
            continue
        if qty_num <= 0 or price_num <= 0:
            continue

        items.append(QuoteItem(name=name_text, qty=qty_num, unit_price=price_num))

    return items


def extract_items_from_excel(source_path: Path) -> list[QuoteItem]:
    wb = load_workbook(source_path, data_only=True)
    ws = wb.worksheets[0]

    items = _extract_from_header_table(ws)
    if not items:
        items = _extract_from_simple_columns(ws)

    if not items:
        raise ExcelQuoteParseError(
            "엑셀에서 품목/수량/단가를 찾지 못했습니다. 표 헤더(품목명, 수량, 단가)를 확인해 주세요."
        )

    return items


def convert_excel_to_source_workbook(source_path: Path, output_xlsx_path: Path) -> Path:
    items = extract_items_from_excel(source_path)

    wb = Workbook()
    ws = wb.active
    ws.title = "본견적"
    ws.cell(1, 1).value = "품목명"
    ws.cell(1, 2).value = "수량"
    ws.cell(1, 3).value = "단가"

    for i, item in enumerate(items, start=2):
        ws.cell(i, 1).value = item.name
        ws.cell(i, 2).value = item.qty
        ws.cell(i, 3).value = item.unit_price

    output_xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_xlsx_path)
    return output_xlsx_path
