from src.pdf_quote_parser import _parse_item_line


def test_parse_item_line_standard_format():
    item = _parse_item_line("1 배송료 1 20,000 20,000 2,000")

    assert item is not None
    assert item.name == "배송료"
    assert item.qty == 1
    assert item.unit_price == 20000


def test_parse_item_line_recovers_attached_quantity():
    item = _parse_item_line("2 상수도소화전(주정차금지) 안전표지판/원형600∮ [반사지인쇄]2 50,000 100,000 10,000")

    assert item is not None
    assert item.name == "상수도소화전(주정차금지) 안전표지판/원형600∮ [반사지인쇄]"
    assert item.qty == 2
    assert item.unit_price == 50000
