from __future__ import annotations

import base64
import hashlib
import io
import ipaddress
import json
import logging
import os
import re
import shutil
import tempfile
import threading
import time
import urllib.parse
import uuid
from pathlib import Path

from flask import Flask, g, jsonify, request, send_file
from flask_cors import CORS
from pptx import Presentation
from pypdf import PdfReader
from werkzeug.utils import secure_filename

try:
    from .seo_visibility_report import (
        SERVICE_LABELS,
        build,
        crawl_target_website,
        detect_services_in_text,
        extract_workbook_profile,
        property_name,
        property_name_from_xlsx,
        pymupdf,
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
        pymupdf,
    )

LOG_LEVEL = os.environ.get("LOG_LEVEL", "INFO").upper()
logging.basicConfig(level=LOG_LEVEL, format="%(message)s")
logger = logging.getLogger("seo_report_app")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 12 * 1024 * 1024  # 12 MB request ceiling
CORS(app, expose_headers=["Content-Disposition", "X-Request-Id", "X-Report-Summary"])

MAX_FILE_BYTES = 5 * 1024 * 1024  # 5 MB per individual file after client slimming
MAX_CACHE_ENTRIES = 10
_BUILD_SEMAPHORE = threading.BoundedSemaphore(value=3)
_CACHE_LOCK = threading.Lock()
CACHE_DIR = Path(tempfile.gettempdir()) / "seo_visibility_report_cache"
CACHE_DIR.mkdir(parents=True, exist_ok=True)


class ValidationError(ValueError):
    """Raised when user-supplied files or URLs fail validation."""


def _log_event(event: str, **fields) -> None:
    payload = {
        "event": event,
        "request_id": getattr(g, "request_id", None),
        **fields,
    }
    logger.info(json.dumps(payload, default=str))


@app.before_request
def _start_request() -> None:
    g.request_id = uuid.uuid4().hex[:12]
    g.start_time = time.monotonic()


@app.after_request
def _finish_request(response):
    duration_ms = round((time.monotonic() - getattr(g, "start_time", time.monotonic())) * 1000, 1)
    response.headers["X-Request-Id"] = getattr(g, "request_id", "")
    _log_event(
        "http_request",
        method=request.method,
        path=request.path,
        status=response.status_code,
        duration_ms=duration_ms,
    )
    return response


@app.errorhandler(400)
def _handle_400(exc):
    return jsonify({"error": str(exc), "request_id": getattr(g, "request_id", "")}), 400


@app.errorhandler(404)
def _handle_404(_exc):
    return jsonify({"error": "Endpoint not found.", "request_id": getattr(g, "request_id", "")}), 404


@app.errorhandler(413)
def _handle_413(_exc):
    return (
        jsonify(
            {
                "error": "Uploaded files exceed the 12 MB server request limit. Please use a smaller workbook or PDF.",
                "request_id": getattr(g, "request_id", ""),
            }
        ),
        413,
    )


@app.errorhandler(500)
def _handle_500(exc):
    _log_event("unhandled_error", error=str(exc))
    return jsonify({"error": "Internal server error while processing report.", "request_id": getattr(g, "request_id", "")}), 500


def _download_filename(name: str) -> str:
    base = re.sub(r'[\\/:*?"<>|]+', "_", name).strip(" .") or "Property_Name"
    return f"seo_visibility_report_{base}.pptx"


def _validate_target_url(raw_url: str) -> str:
    url = (raw_url or "").strip()
    if not url:
        return ""
    if not re.match(r"^https?://", url, re.I):
        if "://" in url:
            raise ValidationError("Target Website URL must use http:// or https://.")
        url = "https://" + url
    parsed = urllib.parse.urlparse(url)
    host = (parsed.hostname or "").lower().strip()
    if not host or "." not in host:
        raise ValidationError("Please enter a valid Target Website domain (e.g. https://thenines.com/).")
    if host in {"localhost", "127.0.0.1", "0.0.0.0", "::1"} or host.endswith(".local"):
        raise ValidationError("Local or internal hostnames are not allowed as Target Website URLs.")
    try:
        ip = ipaddress.ip_address(host)
        if ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved:
            raise ValidationError("Private or reserved IP addresses are not allowed as Target Website URLs.")
    except ValueError:
        pass
    return url


def _save_validated_upload(
    file_storage,
    temp_dir: Path,
    expected_ext: str,
    magic_prefixes: tuple[bytes, ...],
    label: str,
    max_bytes: int = MAX_FILE_BYTES,
) -> Path:
    if not file_storage or not file_storage.filename:
        raise ValidationError(f"{label} file is required.")
    safe_name = secure_filename(Path(file_storage.filename).name) or f"upload{expected_ext}"
    if not safe_name.lower().endswith(expected_ext):
        raise ValidationError(f"{label} must be a {expected_ext} file.")
    dest = temp_dir / safe_name
    file_storage.save(dest)
    size = dest.stat().st_size
    if size == 0:
        raise ValidationError(f"{label} file is empty.")
    if size > max_bytes:
        raise ValidationError(f"{label} exceeds the {max_bytes // (1024 * 1024)} MB per-file limit.")
    with dest.open("rb") as fh:
        header = fh.read(8)
    if not any(header.startswith(prefix) for prefix in magic_prefixes):
        raise ValidationError(f"{label} is not a valid {expected_ext} file (invalid file signature).")
    return dest


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
    script_hash = hashlib.sha256(report_script.read_bytes()).hexdigest() if report_script.is_file() else "v12"
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


def _prune_cache() -> None:
    try:
        entries = sorted(CACHE_DIR.glob("*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)
        for stale in entries[MAX_CACHE_ENTRIES:]:
            stale.unlink(missing_ok=True)
    except Exception:
        pass


def _read_cache_bytes(cache_key: str) -> bytes | None:
    cached_path = CACHE_DIR / f"{cache_key}.pptx"
    with _CACHE_LOCK:
        if cached_path.is_file() and cached_path.stat().st_size > 0:
            try:
                return cached_path.read_bytes()
            except OSError:
                return None
    return None


def _write_cache_bytes(cache_key: str, pptx_bytes: bytes) -> None:
    cached_path = CACHE_DIR / f"{cache_key}.pptx"
    tmp_path = CACHE_DIR / f"{cache_key}.{uuid.uuid4().hex[:8]}.tmp"
    with _CACHE_LOCK:
        try:
            tmp_path.write_bytes(pptx_bytes)
            tmp_path.replace(cached_path)
            _prune_cache()
        except OSError:
            tmp_path.unlink(missing_ok=True)


def _extract_report_summary_header(pptx_bytes: bytes, prop_name: str, pdf_uploaded: bool) -> str:
    """Build a compact base64-encoded JSON summary from the generated PPTX for UI preview."""
    try:
        prs = Presentation(io.BytesIO(pptx_bytes))
        slide_count = len(prs.slides)
        is_brand_hosted = slide_count < 5

        tracked = []
        if slide_count >= 2:
            t_tables = [sh.table for sh in prs.slides[1].shapes if sh.has_table]
            if t_tables:
                tbl = t_tables[0]
                for r in range(1, len(tbl.rows)):
                    kw = tbl.cell(r, 0).text.strip()
                    if not kw:
                        continue
                    # Strip leading "1. " numbering if present
                    kw_clean = re.sub(r"^\d+\.\s*", "", kw)
                    tracked.append(
                        {
                            "keyword": kw_clean,
                            "start": tbl.cell(r, 1).text.strip(),
                            "current": tbl.cell(r, 2).text.strip(),
                            "delta": tbl.cell(r, 3).text.strip(),
                        }
                    )

        ancillary = []
        buckets = {"top3": 0, "pos4_10": 0, "pos11_20": 0, "pos21_40": 0}
        if slide_count >= 3:
            a_tables = [sh.table for sh in prs.slides[2].shapes if sh.has_table]
            if a_tables:
                tbl = a_tables[0]
                for r in range(1, len(tbl.rows)):
                    kw = tbl.cell(r, 0).text.strip()
                    if not kw:
                        continue
                    rank_str = tbl.cell(r, 1).text.strip()
                    vol_str = tbl.cell(r, 2).text.strip()
                    ancillary.append({"keyword": kw, "rank": rank_str, "volume": vol_str})
                    try:
                        rk = int(rank_str)
                        if rk <= 3:
                            buckets["top3"] += 1
                        elif rk <= 10:
                            buckets["pos4_10"] += 1
                        elif rk <= 20:
                            buckets["pos11_20"] += 1
                        elif rk <= 40:
                            buckets["pos21_40"] += 1
                    except ValueError:
                        pass

        top10_count = buckets["top3"] + buckets["pos4_10"]
        for item in tracked:
            try:
                if 1 <= int(item["current"]) <= 10:
                    top10_count += 1
            except ValueError:
                pass

        summary = {
            "property_name": prop_name,
            "slide_count": slide_count,
            "is_brand_hosted": is_brand_hosted,
            "pdf_included": pdf_uploaded,
            "tracked_count": len(tracked),
            "ancillary_count": len(ancillary),
            "top10_count": top10_count,
            "rank_buckets": buckets,
            "tracked_preview": tracked[:6],
            "ancillary_preview": ancillary[:10],
        }
        return base64.b64encode(json.dumps(summary).encode("utf-8")).decode("ascii")
    except Exception:
        return ""


@app.route("/health", methods=["GET"])
def health_check():
    return jsonify(
        {
            "status": "ok",
            "pymupdf_available": pymupdf is not None,
            "max_request_mb": app.config["MAX_CONTENT_LENGTH"] // (1024 * 1024),
            "max_file_mb": MAX_FILE_BYTES // (1024 * 1024),
        }
    )


@app.route("/")
def home():
    root_index = Path(__file__).resolve().parent.parent / "index.html"
    if root_index.is_file():
        return send_file(root_index)
    return send_file("index.html")


@app.route("/crawl", methods=["POST"])
def crawl_website_endpoint():
    temp_dir = None
    try:
        data = request.get_json(silent=True) or {}
        raw_target_url = (data.get("target_url") or request.form.get("target_url") or "").strip()
        target_url = _validate_target_url(raw_target_url) if raw_target_url else ""
        wb_urls_in = data.get("workbook_urls") or []
        wb_kws_in = data.get("workbook_keywords") or []
        prop_title = (data.get("property_title") or "").strip()
        xlsx = request.files.get("xlsx")

        wb_profile = {
            "services": set(),
            "keywords": [str(k).strip() for k in wb_kws_in[:10000] if str(k).strip()],
            "primary_url": str(wb_urls_in[0]).split("?")[0] if wb_urls_in else "",
            "urls": [str(u).strip() for u in wb_urls_in[:500] if str(u).strip()],
        }
        if wb_profile["keywords"]:
            wb_profile["services"] = detect_services_in_text("\n".join(wb_profile["keywords"]))

        if xlsx and xlsx.filename:
            temp_dir = Path(tempfile.mkdtemp())
            xlsx_path = _save_validated_upload(xlsx, temp_dir, ".xlsx", (b"PK\x03\x04",), "Keywords Workbook")
            wb_profile = extract_workbook_profile(xlsx_path)
            prop_title = property_name_from_xlsx(xlsx_path)

        effective_url = target_url or _validate_target_url(wb_profile.get("primary_url", ""))
        if not effective_url:
            raise ValidationError("Please enter a Target Website URL or upload a Keywords Workbook with a URL column.")

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
    except ValidationError as exc:
        return jsonify({"error": str(exc), "request_id": getattr(g, "request_id", "")}), 400
    except Exception as exc:
        _log_event("crawl_error", error=str(exc))
        return jsonify({"error": str(exc), "request_id": getattr(g, "request_id", "")}), 500
    finally:
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)


@app.route("/generate", methods=["POST"])
def generate_report():
    temp_dir = None
    acquired = False
    try:
        temp_dir = Path(tempfile.mkdtemp())

        template = request.files.get("template")
        xlsx = request.files.get("xlsx")
        pdf = request.files.get("pdf")
        raw_target_url = (request.form.get("target_url") or "").strip()
        target_url = _validate_target_url(raw_target_url) if raw_target_url else ""

        template_path = _save_validated_upload(template, temp_dir, ".pptx", (b"PK\x03\x04",), "PowerPoint Template")
        xlsx_path = _save_validated_upload(xlsx, temp_dir, ".xlsx", (b"PK\x03\x04",), "Keywords Workbook")

        pdf_path = None
        name = property_name_from_xlsx(xlsx_path)

        if pdf and pdf.filename:
            pdf_path = _save_validated_upload(pdf, temp_dir, ".pdf", (b"%PDF-",), "Semrush Audit PDF")
            try:
                name = property_name(PdfReader(str(pdf_path)), pdf_path)
            except Exception:
                pass

        cache_key = _cache_key(template_path, xlsx_path, pdf_path, target_url=target_url)
        download_name = _download_filename(name)

        cached_bytes = _read_cache_bytes(cache_key)
        if cached_bytes is not None:
            _log_event("cache_hit", cache_key=cache_key[:12])
            resp = send_file(
                io.BytesIO(cached_bytes),
                as_attachment=True,
                download_name=download_name,
                mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            summary_b64 = _extract_report_summary_header(cached_bytes, name, bool(pdf_path))
            if summary_b64:
                resp.headers["X-Report-Summary"] = summary_b64
            return resp

        acquired = _BUILD_SEMAPHORE.acquire(timeout=15.0)
        if not acquired:
            return (
                jsonify(
                    {
                        "error": "Server is currently busy generating other reports. Please retry in a few seconds.",
                        "request_id": getattr(g, "request_id", ""),
                    }
                ),
                503,
            )

        out_temp = temp_dir / "generated_report.pptx"
        build(template_path, pdf_path, xlsx_path, out_temp, full_list_url="", target_url=target_url)
        pptx_bytes = out_temp.read_bytes()
        _write_cache_bytes(cache_key, pptx_bytes)

        response = send_file(
            io.BytesIO(pptx_bytes),
            as_attachment=True,
            download_name=download_name,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        response.headers["X-Accel-Buffering"] = "no"
        summary_b64 = _extract_report_summary_header(pptx_bytes, name, bool(pdf_path))
        if summary_b64:
            response.headers["X-Report-Summary"] = summary_b64
        return response

    except ValidationError as exc:
        return jsonify({"error": str(exc), "request_id": getattr(g, "request_id", "")}), 400
    except ValueError as exc:
        return jsonify({"error": str(exc), "request_id": getattr(g, "request_id", "")}), 400
    except Exception as exc:
        _log_event("generate_error", error=str(exc))
        return jsonify({"error": str(exc), "request_id": getattr(g, "request_id", "")}), 500
    finally:
        if acquired:
            _BUILD_SEMAPHORE.release()
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=int(os.environ.get("PORT", "8080")))
