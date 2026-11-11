import hashlib
import shutil
import tempfile
from pathlib import Path
import re
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

try:
    from .seo_visibility_report import build, property_name, property_name_from_xlsx
except ImportError:
    from seo_visibility_report import build, property_name, property_name_from_xlsx
from pypdf import PdfReader

app = Flask(__name__)
CORS(app, expose_headers=["Content-Disposition"])
CACHE_DIR = Path(tempfile.gettempdir()) / "seo_visibility_report_cache"
CACHE_DIR.mkdir(parents=True, exist_ok=True)


def _download_filename(name: str) -> str:
    base = re.sub(r'[\\/:*?"<>|]+', "_", name).strip(" .") or "Property_Name"
    return f"seo_visibility_report_{base}.pptx"


def _file_signature(path: Path | None) -> str:
    if not path or not path.is_file():
        return ""
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _cache_key(template_path: Path, xlsx_path: Path, pdf_path: Path | None = None) -> str:
    report_script = Path(__file__).resolve().parent / "seo_visibility_report.py"
    script_hash = hashlib.sha256(report_script.read_bytes()).hexdigest() if report_script.is_file() else "v11"
    payload = "|".join(
        [
            script_hash,
            _file_signature(template_path),
            _file_signature(xlsx_path),
            _file_signature(pdf_path),
        ]
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


@app.route('/')
def home():
    root_index = Path(__file__).resolve().parent.parent / 'index.html'
    if root_index.is_file():
        return send_file(root_index)
    return send_file('index.html')


@app.route('/generate', methods=['POST'])
def generate_report():
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()

        template = request.files.get('template')
        xlsx = request.files.get('xlsx')
        pdf = request.files.get('pdf')

        if not template or not xlsx:
            return jsonify({"error": "PowerPoint template and Excel file are required."}), 400

        template_path = Path(temp_dir) / template.filename
        xlsx_path = Path(temp_dir) / xlsx.filename
        template.save(template_path)
        xlsx.save(xlsx_path)

        pdf_path = None
        name = property_name_from_xlsx(xlsx_path)

        if pdf and pdf.filename:
            pdf_path = Path(temp_dir) / pdf.filename
            pdf.save(pdf_path)
            try:
                name = property_name(PdfReader(str(pdf_path)), pdf_path)
            except Exception:
                pass

        cache_key = _cache_key(template_path, xlsx_path, pdf_path)
        cached_path = CACHE_DIR / f"{cache_key}.pptx"
        download_name = _download_filename(name)

        if cached_path.exists():
            return send_file(
                cached_path,
                as_attachment=True,
                download_name=download_name,
                mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

        output_path = cached_path
        build(template_path, pdf_path, xlsx_path, output_path, full_list_url="")

        response = send_file(
            output_path,
            as_attachment=True,
            download_name=download_name,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        response.headers["X-Accel-Buffering"] = "no"
        return response

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == '__main__':
    app.run(debug=True, port=8080)
