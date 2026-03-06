from __future__ import annotations

from dataclasses import dataclass


META_SHEET_NAME = "_meta"


@dataclass
class QuoteSourceMetadata:
    recipient_name: str | None = None
    recipient_phone: str | None = None
    recipient_fax: str | None = None

    def has_values(self) -> bool:
        return any((self.recipient_name, self.recipient_phone, self.recipient_fax))


def _clean(value) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def write_metadata_sheet(workbook, metadata: QuoteSourceMetadata) -> None:
    if META_SHEET_NAME in workbook.sheetnames:
        workbook.remove(workbook[META_SHEET_NAME])

    ws = workbook.create_sheet(META_SHEET_NAME)
    ws.sheet_state = "hidden"
    ws["A1"] = "recipient_name"
    ws["B1"] = _clean(metadata.recipient_name)
    ws["A2"] = "recipient_phone"
    ws["B2"] = _clean(metadata.recipient_phone)
    ws["A3"] = "recipient_fax"
    ws["B3"] = _clean(metadata.recipient_fax)


def read_metadata_sheet(workbook) -> QuoteSourceMetadata:
    if META_SHEET_NAME not in workbook.sheetnames:
        return QuoteSourceMetadata()

    ws = workbook[META_SHEET_NAME]
    data = {}
    for row in range(1, ws.max_row + 1):
        key = _clean(ws.cell(row, 1).value)
        value = _clean(ws.cell(row, 2).value)
        if key:
            data[key] = value

    return QuoteSourceMetadata(
        recipient_name=data.get("recipient_name"),
        recipient_phone=data.get("recipient_phone"),
        recipient_fax=data.get("recipient_fax"),
    )
