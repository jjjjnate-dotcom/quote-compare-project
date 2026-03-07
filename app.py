import os
from pathlib import Path
import tempfile

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from src.excel_quote_parser import ExcelQuoteParseError, convert_excel_to_source_workbook
from src.pdf_quote_parser import PdfQuoteParseError, convert_pdf_to_source_workbook
from src.quote_generator import QuoteGenerationError, QuoteGenerator, SupplierInfo

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "resources" / "comparison_template.xlsx"
TRUTHY_VALUES = {"1", "true", "on", "yes"}

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "comparison-quote-secret-key")
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in {"xlsx", "xlsm", "pdf"}


def make_safe_upload_name(original_name: str) -> str:
    safe_name = secure_filename(original_name)
    original_suffix = Path(original_name).suffix.lower()

    if "." not in safe_name:
        if original_suffix in {".xlsx", ".xlsm", ".pdf"}:
            return f"upload{original_suffix}"
        return "upload.xlsx"

    return safe_name


def is_checked(value: str | None) -> bool:
    return str(value or "").strip().lower() in TRUTHY_VALUES


def parse_rate(value: str | None, label: str) -> float:
    try:
        return float(value or "0")
    except ValueError as exc:
        raise ValueError(f"{label}은 숫자로 입력해 주세요.") from exc


def get_text(form_key: str, fallback: str) -> str:
    return request.form.get(form_key, fallback).strip() or fallback


def get_required_text(form_key: str, label: str) -> str:
    value = request.form.get(form_key, "").strip()
    if not value:
        raise ValueError(f"{label}을 입력해 주세요.")
    return value


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
@app.route("/generate/", methods=["POST"])
def generate():
    uploaded = request.files.get("quote_file")
    if not uploaded or uploaded.filename == "":
        flash("본견적 파일을 선택해 주세요.")
        return redirect(url_for("index"))

    if not allowed_file(uploaded.filename):
        flash("엑셀 또는 PDF 파일(.xlsx, .xlsm, .pdf)만 업로드할 수 있습니다.")
        return redirect(url_for("index"))

    try:
        company1 = get_text("company1", "거성")
        company2 = get_text("company2", "해광")
        rate1 = parse_rate(request.form.get("rate1", "15"), "업체1 가산율/할인율")
        rate2 = parse_rate(request.form.get("rate2", "20"), "업체2 가산율/할인율")
        vat_rate = parse_rate(request.form.get("vat_rate", "10"), "부가세율")
    except ValueError as exc:
        flash(str(exc))
        return redirect(url_for("index"))

    include_company3 = is_checked(request.form.get("include_company3"))
    company3_name: str | None = None
    company3_rate: float | None = None
    company3_supplier: SupplierInfo | None = None

    if include_company3:
        try:
            company3_name = get_text("company3_name", "업체3")
            company3_rate = parse_rate(request.form.get("rate3", "0"), "업체3 가산율/할인율")
            company3_supplier = SupplierInfo(
                trade_name=get_required_text("supplier_trade_name", "업체3 상호"),
                representative=get_required_text("supplier_representative", "업체3 대표"),
                business_number=get_required_text("supplier_business_number", "업체3 사업자번호"),
                address=get_required_text("supplier_address", "업체3 주소"),
                tel=get_required_text("supplier_tel", "업체3 TEL"),
                fax=get_required_text("supplier_fax", "업체3 FAX"),
            )
        except ValueError as exc:
            flash(str(exc))
            return redirect(url_for("index"))

    with tempfile.TemporaryDirectory() as tmp_dir_name:
        tmp_dir = Path(tmp_dir_name)
        upload_path = tmp_dir / make_safe_upload_name(uploaded.filename)
        uploaded.save(upload_path)

        source_quote_path = tmp_dir / f"{upload_path.stem}_normalized.xlsx"
        if upload_path.suffix.lower() == ".pdf":
            source_quote_path = tmp_dir / f"{upload_path.stem}_parsed.xlsx"
            try:
                convert_pdf_to_source_workbook(upload_path, source_quote_path)
            except PdfQuoteParseError as exc:
                flash(str(exc))
                return redirect(url_for("index"))
        else:
            try:
                convert_excel_to_source_workbook(upload_path, source_quote_path)
            except ExcelQuoteParseError as exc:
                flash(str(exc))
                return redirect(url_for("index"))

        output_path = tmp_dir / f"비교견적_{upload_path.stem}.xlsx"
        try:
            generator = QuoteGenerator(TEMPLATE_PATH)
            generator.generate(
                source_quote_path=source_quote_path,
                output_path=output_path,
                company1_name=company1,
                company2_name=company2,
                company1_rate=rate1 / 100,
                company2_rate=rate2 / 100,
                vat_rate=vat_rate / 100,
                include_company3=include_company3,
                company3_name=company3_name,
                company3_rate=company3_rate / 100 if company3_rate is not None else None,
                company3_supplier=company3_supplier,
            )
        except QuoteGenerationError as exc:
            flash(str(exc))
            return redirect(url_for("index"))
        except Exception as exc:
            app.logger.exception("Unhandled error while generating quote file")
            flash(f"파일 생성 중 오류가 발생했습니다. 상세: {exc}")
            return redirect(url_for("index"))

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_path.name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
