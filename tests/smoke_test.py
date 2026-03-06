import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

from src.quote_generator import QuoteGenerator


BASE_DIR = Path(__file__).resolve().parents[1]


def run_smoke_test():
    generator = QuoteGenerator(BASE_DIR / "resources" / "comparison_template.xlsx")
    output_path = BASE_DIR / "tests" / "smoke_output.xlsx"
    generator.generate(
        source_quote_path=BASE_DIR / "resources" / "comparison_template.xlsx",
        output_path=output_path,
        company1_name="거성",
        company2_name="해광",
        company1_rate=0.15,
        company2_rate=0.20,
        vat_rate=0.10,
    )
    print(f"generated: {output_path}")


if __name__ == "__main__":
    run_smoke_test()