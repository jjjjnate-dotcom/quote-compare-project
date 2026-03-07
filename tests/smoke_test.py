import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

from src.quote_generator import QuoteGenerator, SupplierInfo


def run_smoke_test():
    generator = QuoteGenerator(BASE_DIR / "resources" / "comparison_template.xlsx")
    output_path = BASE_DIR / "tests" / "_smoke_output.xlsx"
    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        generator.generate(
            source_quote_path=BASE_DIR / "resources" / "comparison_template.xlsx",
            output_path=output_path,
            company1_name="거성",
            company2_name="해광",
            company1_rate=0.15,
            company2_rate=0.20,
            vat_rate=0.10,
            include_company3=True,
            company3_name="업체3",
            company3_rate=0.05,
            company3_supplier=SupplierInfo(
                trade_name="테스트상사",
                representative="홍길동",
                business_number="123-45-67890",
                address="서울시 강남구 테스트로 1",
                tel="02-1234-5678",
                fax="02-9876-5432",
            ),
        )
        print(f"generated: {output_path}")
    finally:
        if output_path.exists():
            output_path.unlink()


if __name__ == "__main__":
    run_smoke_test()
