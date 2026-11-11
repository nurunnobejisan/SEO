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

@app.route('/')
def home():
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

        safe_name = re.sub(r'[\\/:*?"<>|]+', "_", name).strip(" .") or "Property_Name"
        output_path = Path(temp_dir) / f"seo_visibility_report_{safe_name}.pptx"

        build(template_path, pdf_path, xlsx_path, output_path, full_list_url="")

        return send_file(output_path, as_attachment=True, download_name=output_path.name)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=8080)
