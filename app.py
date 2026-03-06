import os
from pathlib import Path
import tempfile

from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename

from src.quote_generator import QuoteGenerator, QuoteGenerationError

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "resources" / "comparison_template.xlsx"

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "comparison-quote-secret-key")
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in {"xlsx", "xlsm"}


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    uploaded = request.files.get("quote_file")
    if not uploaded or uploaded.filename == "":
        flash("본견적 엑셀 파일을 선택해 주세요.")
        return redirect(url_for("index"))

    if not allowed_file(uploaded.filename):
        flash("엑셀 파일(.xlsx, .xlsm)만 업로드할 수 있습니다.")
        return redirect(url_for("index"))

    company1 = request.form.get("company1", "거성").strip() or "거성"
    company2 = request.form.get("company2", "해광").strip() or "해광"

    try:
        rate1 = float(request.form.get("rate1", "15"))
        rate2 = float(request.form.get("rate2", "20"))
        vat_rate = float(request.form.get("vat_rate", "10"))
    except ValueError:
        flash("가산율/할인율과 부가세율은 숫자로 입력해 주세요.")
        return redirect(url_for("index"))

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir = Path(tmp_dir)
        upload_path = tmp_dir / secure_filename(uploaded.filename)
        uploaded.save(upload_path)

        output_path = tmp_dir / f"비교견적_{upload_path.stem}.xlsx"

        try:
            generator = QuoteGenerator(TEMPLATE_PATH)
            generator.generate(
                source_quote_path=upload_path,
                output_path=output_path,
                company1_name=company1,
                company2_name=company2,
                company1_rate=rate1 / 100,
                company2_rate=rate2 / 100,
                vat_rate=vat_rate / 100,
            )
        except QuoteGenerationError as exc:
            flash(str(exc))
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
