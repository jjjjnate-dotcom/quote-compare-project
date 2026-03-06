from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook, load_workbook

from .excel_utils import (
    SOURCE_SHEET_NAME,
    apply_row_merges,
    clear_row_values,
    copy_row_style,
    copy_sheet_content,
    detect_item_count,
)


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


class QuoteGenerator:
    def __init__(self, template_path: Path):
        self.template_path = Path(template_path)
        if not self.template_path.exists():
            raise QuoteGenerationError(f"템플릿 파일을 찾을 수 없습니다: {self.template_path}")

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
        template_wb = load_workbook(self.template_path)
        source_wb = load_workbook(source_quote_path, data_only=False)

        source_ws = source_wb.worksheets[0]
        item_count = detect_item_count(source_ws)
        if item_count == 0:
            raise QuoteGenerationError("본견적 첫 번째 시트에서 품목을 찾지 못했습니다.")

        self._replace_source_sheet(template_wb, source_ws)

        sheet1_name = template_wb.sheetnames[1]
        sheet2_name = template_wb.sheetnames[2]
        template_wb[sheet1_name].title = company1_name
        template_wb[sheet2_name].title = company2_name

        self._fill_geoseong_sheet(template_wb[company1_name], item_count, company1_rate, vat_rate)
        self._fill_haegwang_sheet(template_wb[company2_name], item_count, company2_rate, vat_rate)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        template_wb.save(output_path)
        return output_path

    def _replace_source_sheet(self, wb: Workbook, source_ws) -> None:
        old_ws = wb[wb.sheetnames[0]]
        wb.remove(old_ws)
        new_ws = wb.create_sheet(title=SOURCE_SHEET_NAME, index=0)
        copy_sheet_content(source_ws, new_ws)

    def _fill_geoseong_sheet(self, ws, item_count: int, rate: float, vat_rate: float) -> None:
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
            for row in range(spec.template_total_row, spec.template_total_row + extra_rows):
                copy_row_style(ws, spec.style_row, row, spec.max_col)
                apply_row_merges(ws, row, [(2, 5), (9, 10), (11, 12)])

        total_row = spec.template_total_row + extra_rows
        last_item_row = spec.item_start_row + item_count - 1

        for row in range(spec.item_start_row, last_item_row + 1):
            source_row = row - spec.item_start_row + 2
            ws.cell(row, 1).value = row - spec.item_start_row + 1
            ws.cell(row, 2).value = f"={SOURCE_SHEET_NAME}!A{source_row}"
            ws.cell(row, 6).value = f"={SOURCE_SHEET_NAME}!B{source_row}"
            ws.cell(row, 8).value = f"=ROUND({SOURCE_SHEET_NAME}!C{source_row}*(1+{rate}),-2)"
            ws.cell(row, 9).value = f"=H{row}*F{row}"
            ws.cell(row, 11).value = f"=I{row}*{vat_rate}"
            ws.cell(row, 13).value = None

        for row in range(last_item_row + 1, total_row):
            clear_row_values(ws, row, [1, 2, 6, 8, 9, 11, 13])

        ws.cell(total_row, 1).value = "합 계"
        ws.cell(total_row, 9).value = f"=SUM(I{spec.item_start_row}:I{last_item_row})"
        ws.cell(total_row, 11).value = f"=I{total_row}*{vat_rate}"
        ws.cell(9, 10).value = f"=I{total_row}+K{total_row}"

    def _fill_haegwang_sheet(self, ws, item_count: int, rate: float, vat_rate: float) -> None:
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
            for row in range(spec.template_total_row, spec.template_total_row + extra_rows):
                copy_row_style(ws, spec.style_row, row, spec.max_col)
                apply_row_merges(ws, row, [(4, 15), (16, 17), (19, 21), (22, 25), (26, 28), (29, 32)])

        total_row = spec.template_total_row + extra_rows
        last_item_row = spec.item_start_row + item_count - 1

        for row in range(spec.item_start_row, last_item_row + 1):
            source_row = row - spec.item_start_row + 2
            ws.cell(row, 3).value = row - spec.item_start_row + 1
            ws.cell(row, 4).value = f"={SOURCE_SHEET_NAME}!A{source_row}"
            ws.cell(row, 18).value = f"={SOURCE_SHEET_NAME}!B{source_row}"
            ws.cell(row, 19).value = f"=ROUND({SOURCE_SHEET_NAME}!C{source_row}*(1+{rate}),-2)"
            ws.cell(row, 22).value = f"=R{row}*S{row}"
            ws.cell(row, 26).value = f"=V{row}*{vat_rate}"
            ws.cell(row, 29).value = None

        for row in range(last_item_row + 1, total_row):
            clear_row_values(ws, row, [3, 4, 18, 19, 22, 26, 29])

        ws.cell(total_row, 3).value = "공급가액"
        ws.cell(total_row, 5).value = f"=SUM(V{spec.item_start_row}:V{last_item_row})"
        ws.cell(total_row, 10).value = "세액"
        ws.cell(total_row, 14).value = f"=E{total_row}*{vat_rate}"
        ws.cell(total_row, 19).value = "합계(Total)"
        ws.cell(total_row, 25).value = f"=E{total_row}+N{total_row}"
