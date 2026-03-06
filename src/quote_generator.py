from __future__ import annotations

from copy import copy
from dataclasses import dataclass
from datetime import date
import math
from pathlib import Path
import re
from zipfile import BadZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from .excel_utils import (
    SOURCE_SHEET_NAME,
    apply_row_merges,
    clear_row_values,
    copy_row_style,
    copy_sheet_content,
    detect_item_count,
)
from .quote_source_metadata import QuoteSourceMetadata, read_metadata_sheet


class QuoteGenerationError(Exception):
    pass


@dataclass
class CompareSheetSpec:
    sheet_name: str
    item_start_row: int
    template_capacity: int
    template_total_row: int
    style_row: int
    max_col: int


INVALID_SHEET_TITLE_CHARS = re.compile(r"[\\/*?:\[\]]")


class QuoteGenerator:
    def __init__(self, template_path: Path):
        self.template_path = Path(template_path)
        if not self.template_path.exists():
            raise QuoteGenerationError(f"Template file not found: {self.template_path}")

    def generate(
        self,
        source_quote_path: Path,
        output_path: Path,
        company1_name: str,
        company2_name: str,
        company1_rate: float,
        company2_rate: float,
        vat_rate: float,
    ) -> Path:
        try:
            template_wb = load_workbook(self.template_path)
        except (BadZipFile, InvalidFileException) as exc:
            raise QuoteGenerationError(f"Cannot open template file: {self.template_path.name}") from exc

        try:
            source_wb = load_workbook(source_quote_path, data_only=False)
        except (BadZipFile, InvalidFileException) as exc:
            raise QuoteGenerationError("Cannot open uploaded workbook. Check .xlsx/.xlsm format and file integrity.") from exc

        source_ws = source_wb.worksheets[0]
        source_metadata = read_metadata_sheet(source_wb)
        item_count = detect_item_count(source_ws)
        if item_count == 0:
            raise QuoteGenerationError("No items found in the first sheet of uploaded workbook.")

        self._replace_source_sheet(template_wb, source_ws)

        if len(template_wb.sheetnames) < 3:
            raise QuoteGenerationError("Template workbook structure is invalid.")

        sheet1_name = template_wb.sheetnames[1]
        sheet2_name = template_wb.sheetnames[2]

        company1_title = self._sanitize_sheet_title(company1_name, "Company1")
        company2_title = self._sanitize_sheet_title(company2_name, "Company2")
        company1_title, company2_title = self._make_distinct_titles(company1_title, company2_title)

        template_wb[sheet1_name].title = company1_title
        template_wb[sheet2_name].title = company2_title

        self._fill_geoseong_sheet(
            template_wb[company1_title],
            source_ws,
            source_metadata,
            item_count,
            company1_rate,
            vat_rate,
        )
        self._fill_haegwang_sheet(
            template_wb[company2_title],
            source_ws,
            source_metadata,
            item_count,
            company2_rate,
            vat_rate,
        )

        output_path.parent.mkdir(parents=True, exist_ok=True)
        template_wb.save(output_path)
        return output_path

    @staticmethod
    def _sanitize_sheet_title(raw_title: str, fallback: str) -> str:
        title = (raw_title or "").strip()
        title = INVALID_SHEET_TITLE_CHARS.sub("_", title)
        title = title.strip("'")
        if not title:
            title = fallback
        return title[:31]

    @staticmethod
    def _make_distinct_titles(title1: str, title2: str) -> tuple[str, str]:
        if title1 != title2:
            return title1, title2

        suffix = " (2)"
        max_base_len = 31 - len(suffix)
        title2 = f"{title2[:max_base_len]}{suffix}"
        if title2 == title1:
            title2 = "Company2"
        return title1, title2

    @staticmethod
    def _to_float(value) -> float | None:
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            cleaned = value.replace(",", "").strip()
            if cleaned == "":
                return None
            try:
                return float(cleaned)
            except ValueError:
                return None
        return None

    @staticmethod
    def _round_to_hundred_half_up(value: float) -> float:
        factor = 100.0
        if value >= 0:
            return math.floor(value / factor + 0.5) * factor
        return -math.floor(abs(value) / factor + 0.5) * factor

    def _replace_source_sheet(self, wb: Workbook, source_ws) -> None:
        old_ws = wb[wb.sheetnames[0]]
        wb.remove(old_ws)
        new_ws = wb.create_sheet(title=SOURCE_SHEET_NAME, index=0)
        copy_sheet_content(source_ws, new_ws)

    def _apply_recipient_to_geoseong(self, ws, metadata: QuoteSourceMetadata) -> None:
        if not metadata.has_values():
            return

        lines: list[str] = []
        if metadata.recipient_name:
            lines.append(metadata.recipient_name)

        contact_parts: list[str] = []
        if metadata.recipient_phone:
            contact_parts.append(f"TEL {metadata.recipient_phone}")
        if metadata.recipient_fax:
            contact_parts.append(f"FAX {metadata.recipient_fax}")
        if contact_parts:
            lines.append(" / ".join(contact_parts))

        if not lines:
            return

        cell = ws.cell(6, 2)
        cell.value = "\n".join(lines)
        alignment = copy(cell.alignment)
        alignment.wrap_text = True
        cell.alignment = alignment
        current_height = ws.row_dimensions[6].height or 15.0
        ws.row_dimensions[6].height = max(current_height, 30.0 if len(lines) > 1 else 20.0)

    def _apply_recipient_to_haegwang(self, ws, metadata: QuoteSourceMetadata) -> None:
        if not metadata.has_values():
            return

        if metadata.recipient_name:
            ws.cell(2, 6).value = metadata.recipient_name

        ws.cell(5, 6).value = metadata.recipient_phone or ""
        ws.cell(6, 6).value = metadata.recipient_fax or ""

    def _fill_geoseong_sheet(
        self,
        ws,
        source_ws,
        source_metadata: QuoteSourceMetadata,
        item_count: int,
        rate: float,
        vat_rate: float,
    ) -> None:
        spec = CompareSheetSpec(
            sheet_name=ws.title,
            item_start_row=11,
            template_capacity=23,
            template_total_row=34,
            style_row=11,
            max_col=13,
        )
        extra_rows = max(0, item_count - spec.template_capacity)
        if extra_rows:
            ws.insert_rows(spec.template_total_row, extra_rows)
            for row_idx in range(spec.template_total_row, spec.template_total_row + extra_rows):
                copy_row_style(ws, spec.style_row, row_idx, spec.max_col)
                apply_row_merges(ws, row_idx, [(2, 5), (9, 10), (11, 12)])

        total_row = spec.template_total_row + extra_rows
        last_item_row = spec.item_start_row + item_count - 1
        supply_total = 0.0

        for row_idx in range(spec.item_start_row, last_item_row + 1):
            source_row = row_idx - spec.item_start_row + 2
            source_name = source_ws.cell(source_row, 1).value
            source_qty = source_ws.cell(source_row, 2).value
            source_price = source_ws.cell(source_row, 3).value

            qty_num = self._to_float(source_qty)
            price_num = self._to_float(source_price)
            adjusted_price = self._round_to_hundred_half_up(price_num * (1 + rate)) if price_num is not None else None
            supply_amount = adjusted_price * qty_num if adjusted_price is not None and qty_num is not None else None
            vat_amount = supply_amount * vat_rate if supply_amount is not None else None

            ws.cell(row_idx, 1).value = row_idx - spec.item_start_row + 1
            ws.cell(row_idx, 2).value = source_name
            ws.cell(row_idx, 6).value = source_qty
            ws.cell(row_idx, 8).value = adjusted_price
            ws.cell(row_idx, 9).value = supply_amount
            ws.cell(row_idx, 11).value = vat_amount
            ws.cell(row_idx, 13).value = None

            if supply_amount is not None:
                supply_total += supply_amount

        for row_idx in range(last_item_row + 1, total_row):
            clear_row_values(ws, row_idx, [1, 2, 6, 8, 9, 11, 13])

        ws.cell(total_row, 1).value = "TOTAL"
        ws.cell(total_row, 9).value = supply_total
        ws.cell(total_row, 11).value = supply_total * vat_rate
        ws.cell(9, 10).value = ws.cell(total_row, 9).value + ws.cell(total_row, 11).value
        ws.cell(4, 2).value = date.today()
        ws.cell(9, 6).value = ws.cell(9, 10).value
        self._apply_recipient_to_geoseong(ws, source_metadata)

    def _fill_haegwang_sheet(
        self,
        ws,
        source_ws,
        source_metadata: QuoteSourceMetadata,
        item_count: int,
        rate: float,
        vat_rate: float,
    ) -> None:
        spec = CompareSheetSpec(
            sheet_name=ws.title,
            item_start_row=15,
            template_capacity=20,
            template_total_row=36,
            style_row=15,
            max_col=32,
        )
        extra_rows = max(0, item_count - spec.template_capacity)
        if extra_rows:
            ws.insert_rows(spec.template_total_row, extra_rows)
            for row_idx in range(spec.template_total_row, spec.template_total_row + extra_rows):
                copy_row_style(ws, spec.style_row, row_idx, spec.max_col)
                apply_row_merges(ws, row_idx, [(4, 15), (16, 17), (19, 21), (22, 25), (26, 28), (29, 32)])

        total_row = spec.template_total_row + extra_rows
        last_item_row = spec.item_start_row + item_count - 1
        supply_total = 0.0

        for row_idx in range(spec.item_start_row, last_item_row + 1):
            source_row = row_idx - spec.item_start_row + 2
            source_name = source_ws.cell(source_row, 1).value
            source_qty = source_ws.cell(source_row, 2).value
            source_price = source_ws.cell(source_row, 3).value

            qty_num = self._to_float(source_qty)
            price_num = self._to_float(source_price)
            adjusted_price = self._round_to_hundred_half_up(price_num * (1 + rate)) if price_num is not None else None
            supply_amount = adjusted_price * qty_num if adjusted_price is not None and qty_num is not None else None
            vat_amount = supply_amount * vat_rate if supply_amount is not None else None

            ws.cell(row_idx, 3).value = row_idx - spec.item_start_row + 1
            ws.cell(row_idx, 4).value = source_name
            ws.cell(row_idx, 18).value = source_qty
            ws.cell(row_idx, 19).value = adjusted_price
            ws.cell(row_idx, 22).value = supply_amount
            ws.cell(row_idx, 26).value = vat_amount
            ws.cell(row_idx, 29).value = None

            if supply_amount is not None:
                supply_total += supply_amount

        for row_idx in range(last_item_row + 1, total_row):
            clear_row_values(ws, row_idx, [3, 4, 18, 19, 22, 26, 29])

        ws.cell(total_row, 3).value = "TOTAL"
        ws.cell(total_row, 5).value = supply_total
        ws.cell(total_row, 10).value = "VAT"
        ws.cell(total_row, 14).value = supply_total * vat_rate
        ws.cell(total_row, 19).value = "SUM(Total)"
        ws.cell(total_row, 25).value = ws.cell(total_row, 5).value + ws.cell(total_row, 14).value
        ws.cell(8, 6).value = date.today()
        self._apply_recipient_to_haegwang(ws, source_metadata)
