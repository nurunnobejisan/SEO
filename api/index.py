from __future__ import annotations

import hashlib
import shutil
import tempfile
from pathlib import Path
import re
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

try:
    from .seo_visibility_report import (
        SERVICE_LABELS,
        build,
        crawl_target_website,
        detect_services_in_text,
        extract_workbook_profile,
        property_name,
        property_name_from_xlsx,
    )
except ImportError:
    from seo_visibility_report import (
        SERVICE_LABELS,
        build,
        crawl_target_website,
        detect_services_in_text,
        extract_workbook_profile,
        property_name,
        property_name_from_xlsx,
    )
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


def _cache_key(
    template_path: Path,
    xlsx_path: Path,
    pdf_path: Path | None = None,
    target_url: str = "",
) -> str:
    report_script = Path(__file__).resolve().parent / "seo_visibility_report.py"
    script_hash = hashlib.sha256(report_script.read_bytes()).hexdigest() if report_script.is_file() else "v11"
    payload = "|".join(
        [
            script_hash,
            _file_signature(template_path),
            _file_signature(xlsx_path),
            _file_signature(pdf_path),
            (target_url or "").strip().lower(),
        ]
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


@app.route('/')
def home():
    root_index = Path(__file__).resolve().parent.parent / 'index.html'
    if root_index.is_file():
        return send_file(root_index)
    return send_file('index.html')


@app.route('/crawl', methods=['POST'])
def crawl_website_endpoint():
    temp_dir = None
    try:
        data = request.get_json(silent=True) or {}
        target_url = (data.get('target_url') or request.form.get('target_url') or '').strip()
        wb_urls_in = data.get('workbook_urls') or []
        wb_kws_in = data.get('workbook_keywords') or []
        prop_title = (data.get('property_title') or '').strip()
        xlsx = request.files.get('xlsx')

        wb_profile = {
            "services": set(),
            "keywords": [str(k).strip() for k in wb_kws_in if str(k).strip()],
            "primary_url": wb_urls_in[0].split("?")[0] if wb_urls_in else "",
            "urls": [str(u).strip() for u in wb_urls_in if str(u).strip()],
        }
        if wb_profile["keywords"]:
            wb_profile["services"] = detect_services_in_text("\n".join(wb_profile["keywords"]))

        if xlsx and xlsx.filename:
            temp_dir = tempfile.mkdtemp()
            xlsx_path = Path(temp_dir) / xlsx.filename
            xlsx.save(xlsx_path)
            wb_profile = extract_workbook_profile(xlsx_path)
            prop_title = property_name_from_xlsx(xlsx_path)

        effective_url = target_url or wb_profile.get("primary_url", "")
        if not effective_url:
            return jsonify({"error": "Please enter a Target Website URL or upload a Keywords Workbook with a URL column."}), 400

        crawl_info = crawl_target_website(
            effective_url,
            fallback_text=prop_title,
            workbook_urls=wb_profile.get("urls"),
            workbook_keywords=wb_profile.get("keywords"),
        )
        web_svcs = crawl_info["services"]
        wb_svcs = wb_profile["services"]
        matched_svcs = (web_svcs & wb_svcs) if wb_svcs else web_svcs
        if "lodging" in wb_svcs:
            matched_svcs.add("lodging")
        unmatched_wb_svcs = (wb_svcs - web_svcs) if wb_svcs else set()

        return jsonify(
            {
                "target_url": crawl_info["target_url"],
                "page_title": crawl_info["page_title"] or prop_title or crawl_info["target_url"],
                "crawled_pages": crawl_info["crawled_pages"],
                "crawled_urls": crawl_info.get("crawled_urls", []),
                "sitemap_urls": crawl_info.get("sitemap_urls", 0),
                "page_topics": crawl_info.get("page_topics", []),
                "website_services": [SERVICE_LABELS.get(s, s) for s in sorted(web_svcs)],
                "workbook_services": [SERVICE_LABELS.get(s, s) for s in sorted(wb_svcs)],
                "matched_services": [SERVICE_LABELS.get(s, s) for s in sorted(matched_svcs)],
                "unmatched_services": [SERVICE_LABELS.get(s, s) for s in sorted(unmatched_wb_svcs)],
                "workbook_keyword_count": len(wb_profile["keywords"]),
            }
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)


@app.route('/generate', methods=['POST'])
def generate_report():
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()

        template = request.files.get('template')
        xlsx = request.files.get('xlsx')
        pdf = request.files.get('pdf')
        target_url = (request.form.get('target_url') or '').strip()

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

        cache_key = _cache_key(template_path, xlsx_path, pdf_path, target_url=target_url)
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
        build(template_path, pdf_path, xlsx_path, output_path, full_list_url="", target_url=target_url)

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
