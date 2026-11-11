#!/usr/bin/env python3
"""Build a monthly SEO visibility PowerPoint from a template, Semrush PDF and XLSX.

Trained on historical Cogwheel Marketing / Analytics reports:
- Supports multi-style executive Insights (Executive Summary style & Named Highlights style)
- Preserves Start of Project & Diff when a keyword drops to N/A (out of top 100)
- Forces all table keywords to Source Sans Pro size 9.
- EXACT text output requirement for Executive Insights matching user request.
- ALLOWS exactly 10 keywords maximum on the Ancillary list (as requested).
- EXCLUDES full physical addresses from keywords (e.g. '10 hightower pl florence al 35630').
- STRICT Brand Exclusion Rule applied unconditionally across all major chains for Ancillary keywords.
- ABSOLUTE property name exclusion for Ancillary keywords (does not strip location names if they are part of the keyword).
"""

from __future__ import annotations

import argparse
import copy
import re
import sys
from pathlib import Path

try:
    import pandas as pd
    from pptx import Presentation
    from pptx.util import Pt
    from pypdf import PdfReader
except ImportError as exc:
    raise SystemExit(
        "Missing a required package. Install them with:\n"
        "python3 -m pip install python-pptx pypdf openpyxl pandas"
    ) from exc


try:
    import pymupdf  # preferred for reading the Semrush keyword table
except ImportError:  # falls back to pypdf
    try:
        import fitz as pymupdf
    except ImportError:
        pymupdf = None

# STRICT BRAND EXCLUSION RULE - Parent Chains and all Sub-Brands
ALL_CHAIN_BRANDS = [
    "marriott", "bonvoy", "ihg", "hilton", "hyatt",
    "edition", "ritz carlton", "ritz-carlton", "luxury collection", "st regis", "st. regis", "w hotels",
    "sheraton", "delta hotels", "westin", "le meridien", "le méridien", "renaissance",
    "autograph collection", "tribute portfolio", "design hotels", "gaylord",
    "courtyard", "four points", "springhill suites", "fairfield", "ac hotels", "citizenm",
    "aloft", "moxy", "protea", "residence inn", "towneplace suites", "element",
    "six senses", "regent", "intercontinental", "vignette", "kimpton", "hotel indigo",
    "crowne plaza", "even hotels", "voco", "holiday inn", "avid hotels",
    "staybridge suites", "candlewood suites",
    "waldorf astoria", "lxr", "conrad", "nomad", "signia", "canopy",
    "curio collection", "graduate", "doubletree", "tapestry collection", "embassy suites",
    "tempo", "motto", "hampton", "tru", "spark", "homewood suites", "home2 suites",
    "park hyatt", "alila", "miraval", "andaz", "thompson", "the standard",
    "dream hotels", "grand hyatt", "hyatt regency", "hyatt centric", "hyatt place", "hyatt house",
    "urcove", "zoetry", "zoëtry", "secrets resorts"
]

MIN_SEARCH_VOLUME = 30
MAX_TRACKED_KEYWORDS = 10  # Capped at 10 as requested
MAX_ANCILLARY_KEYWORDS = 10  # Capped at 10 as requested


def choose_file(explicit: str | None, suffix: str, role: str, optional: bool = False) -> Path | None:
    if explicit:
        path = Path(explicit).expanduser().resolve()
        if not path.is_file():
            raise ValueError(f"{role} file does not exist: {path}")
        return path
    roots = [Path.cwd()]
    content = Path("/content")
    if content.is_dir() and content.resolve() != Path.cwd().resolve():
        roots.append(content)
    candidates = sorted({p.resolve() for root in roots for p in root.glob(f"*{suffix}")})
    if role == "PowerPoint template":
        named_templates = [p for p in candidates if "template" in p.name.lower()]
        if named_templates:
            candidates = named_templates
    if optional and not candidates:
        return None
    if len(candidates) != 1:
        if not candidates:
            raise ValueError(f"No {role} (*{suffix}) found. Upload/select the file or pass its path on the command line.")
        names = "\n  ".join(str(p) for p in candidates)
        raise ValueError(f"Found multiple possible {role} files. Pass the correct one explicitly:\n  {names}")
    return candidates[0].resolve()


def clean_text(value: str | None) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def property_name(reader: PdfReader, pdf_path: Path) -> str:
    first = clean_text(reader.pages[0].extract_text()) if reader.pages else ""
    match = re.search(r"\bFor\s+(.+?)(?=\s+(?:Domain|Date|Device|Database|Generated\s+on|Report\s+Date)\b|$)", first, re.I)
    if match:
        name = clean_text(match.group(1)).strip(" _-:")
        name = re.sub(r"[_]+", " ", name)
        name = re.sub(r"\s+", " ", name)
        if name:
            return name
    match = re.search(r"Report[-_]For[-_](.+?)(?:[-_]\d{1,2}(?:st|nd|rd|th)?[-_]|\.pdf$)", pdf_path.name, re.I)
    if match:
        return re.sub(r"[_]+", " ", match.group(1)).strip(" _-")
    raise ValueError("Could not find the property name in the first PDF page or filename.")


def property_name_from_xlsx(xlsx_path: Path) -> str:
    name = re.sub(r"\s*\([^)]*\)(?=\.xlsx$)", "", xlsx_path.name, flags=re.I)
    name = re.sub(r"\.xlsx$", "", name, flags=re.I)
    name = re.sub(r"\s*[-_ ]+(?:ranking\s+)?keywords?\s*$", "", name, flags=re.I)
    name = re.sub(r"\s+", " ", name.replace("_", " ")).strip(" -_")
    return name or "SEO Report"


VALUE_RE = r"[+\-]?\d[\d,.]*|-|n/?a"

def is_brand_excluded(keyword: str, property_title: str) -> bool:
    kw_clean = re.sub(r"[^a-z0-9\s]", "", keyword.lower())
    prop_clean = re.sub(r"[^a-z0-9\s]", "", property_title.lower())
    
    # NEW RULE: Exclude physical addresses completely (e.g. starts with digits, ends with zip code)
    if re.search(r"\d+\s+[a-z]+\s+[a-z]+.*\b\d{5}\b", kw_clean):
        return True
    
    # 1. Filter out major chain brands
    for brand in ALL_CHAIN_BRANDS:
        brand_clean = re.sub(r"[^a-z0-9\s]", "", brand.lower())
        if re.search(r"\b" + re.escape(brand_clean) + r"\b", kw_clean):
            return True
            
    # 2. Filter out the specific property's FULL name identifier (e.g. "teatro" or "morrison house")
    ignore_words = {"hotel", "resort", "inn", "suites", "spa", "by", "the", "and", "of", "a"}
    prop_words = [w for w in prop_clean.split() if w not in ignore_words and len(w) > 3]
    
    # Check if the keyword contains the unique property identifier
    for word in prop_words:
        # DO NOT exclude location names (e.g., 'muscle shoals', 'shoals', 'denver', 'san antonio') 
        # Check against common generic location terms we want to preserve in ancillary keywords
        generic_locations = ["shoals", "denver", "san antonio", "orlando", "cary", "boerne", "alexandria", "new orleans", "salt lake city", "fairhope"]
        if word in generic_locations:
            continue
            
        if re.search(r"\b" + re.escape(word) + r"\b", kw_clean):
            return True
            
    return False

def parse_tracked_keywords(reader: PdfReader, pdf_path: Path | None = None) -> list[dict]:
    if pymupdf is not None and pdf_path is not None:
        pages = [page.get_text() for page in pymupdf.open(str(pdf_path))]
    else:
        pages = [(p.extract_text() or "") for p in reader.pages]
    start = next((i for i, text in enumerate(pages) if "Keyword Rankings" in text), None)
    if start is None:
        raise ValueError("The PDF has no 'Keyword Rankings' section. Check that this is the expected Semrush report.")
    text = "\n".join(pages[start:])
    section = text.split("Keyword Rankings", 1)[1]
    
    section = re.split(r"\b(Site Audit|Backlinks:?|Organic Research|Traffic Analytics)\b", section, maxsplit=1, flags=re.I)[0]
    total_match = re.search(r"\d+\s*-\s*\d+\s*\((\d+)\)", section)
    expected = int(total_match.group(1)) if total_match else None

    entries: list[dict] = []
    current = None
    for line in (clean_text(line) for line in section.splitlines()):
        if not line:
            continue
        if re.match(r"^\d+\.\s+\S", line):
            current = {"name_parts": [re.sub(r"^\d+\.\s+", "", line)], "values": [], "finished_name": False}
            entries.append(current)
        elif current is None:
            continue
        elif re.match(r"^\d+(?:st|nd|rd|th)\s+time$", line, re.I):
            pass # skip these lines
        else:
            tokens = line.replace("\x00", "+").split()
            if tokens and all(re.fullmatch(VALUE_RE, t, re.I) for t in tokens):
                current["values"].extend(tokens)
                current["finished_name"] = True
            elif not current.get("finished_name"):
                if not re.search(r"https?://", line) and not re.match(r"^(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s*\d", line, re.I):
                    current["name_parts"].append(line)

    entries = [e for e in entries if len(e["values"]) >= 4]

    result = []
    for entry in entries[:expected] if expected else entries:
        name = clean_text(" ".join(entry["name_parts"]))
        if len(name) > 80:
            name = name[:77] + "..."
            
        values = entry["values"]
        pos = diff = start_pos = None
        if len(values) >= 3:
            pos = int(values[1].replace(",", "")) if re.fullmatch(r"\d[\d,]*", values[1]) else None
            diff = int(values[2].replace(",", "")) if re.fullmatch(r"[+\-]?\d+", values[2]) else None
            if pos is not None and diff is not None:
                start_pos = pos + diff
            elif pos is None and diff is not None and diff < 0:
                start_pos = 100 + diff
        if name:
            result.append({"name": name, "start": start_pos, "pos": pos, "diff": diff})
    if not result:
        raise ValueError("Could not read keyword rows from the PDF. No fallback keyword data was inserted.")
    
    return result[:MAX_TRACKED_KEYWORDS]


def tracked_from_xlsx(xlsx_path: Path) -> list[dict]:
    frame = pd.read_excel(xlsx_path)
    frame.columns = [str(c).strip() for c in frame.columns]
    keyword_cols = [c for c in frame.columns if any(x in c.lower() for x in ("keyword", "query", "search term"))]
    position_cols = [c for c in frame.columns if any(x in c.lower() for x in ("position", "rank"))]
    if not keyword_cols or not position_cols:
        raise ValueError("Without a PDF, the Excel file must contain keyword and position/rank columns.")
    kw_col, pos_col = keyword_cols[0], position_cols[0]
    frame["_kw"] = frame[kw_col].fillna("").astype(str).str.strip()
    frame["_pos"] = pd.to_numeric(frame[pos_col], errors="coerce")
    frame = frame[(frame["_kw"] != "") & frame["_pos"].notna()].drop_duplicates("_kw").head(MAX_TRACKED_KEYWORDS)
    if frame.empty:
        raise ValueError("No keyword rows with a numeric position were found in the Excel file.")
    return [{"name": row["_kw"], "start": int(row["_pos"]), "pos": int(row["_pos"]), "diff": 0} for _, row in frame.iterrows()]


def keyword_rows(xlsx_path: Path, tracked: list[dict], prop_name: str = "") -> tuple[list[dict], str]:
    try:
        frame = pd.read_excel(xlsx_path)
    except Exception as exc:
        raise ValueError(f"Could not read the Excel file: {exc}") from exc
    if frame.empty:
        raise ValueError("The Excel keyword sheet is empty.")
    frame.columns = [str(c).strip() for c in frame.columns]

    def find_column(needles: tuple[str, ...], label: str) -> str:
        found = [c for c in frame.columns if any(word in c.lower() for word in needles)]
        if not found:
            raise ValueError(f"Excel is missing a {label} column. Available columns: {', '.join(frame.columns)}")
        return found[0]

    kw_col = find_column(("keyword", "query", "search term"), "keyword")
    pos_col = find_column(("position", "rank"), "position/rank")
    vol_col = find_column(("volume", "searches"), "search volume")
    
    frame["_keyword"] = frame[kw_col].fillna("").astype(str).str.strip()
    frame["_normalized"] = frame["_keyword"].str.casefold()
    frame["_position"] = pd.to_numeric(frame[pos_col], errors="coerce")
    frame["_volume"] = pd.to_numeric(frame[vol_col], errors="coerce").fillna(0)
    
    frame = frame[(frame["_keyword"] != "") & frame["_position"].notna()]
    frame = frame.sort_values(["_position", "_volume"], ascending=[True, False]).drop_duplicates("_normalized")

    # Drop keywords already in the tracked summary
    tracked_names = {item["name"].casefold() for item in tracked}
    frame = frame[~frame["_normalized"].isin(tracked_names)]
    
    # Apply minimum search volume
    frame = frame[frame["_volume"] >= MIN_SEARCH_VOLUME]

    # Apply strict brand exclusion rule for Ancillary keywords
    frame["_is_excluded"] = frame["_keyword"].apply(lambda kw: is_brand_excluded(kw, prop_name))
    frame = frame[~frame["_is_excluded"]]

    selected = []
    for position, group in frame.groupby("_position", sort=True):
        selected.extend(group.head(4 if position <= 2 else 2).to_dict("records"))
    selected = selected[:MAX_ANCILLARY_KEYWORDS]
    rows = [{"keyword": row["_keyword"], "position": int(row["_position"]), "volume": int(row["_volume"])} for row in selected]
    return rows, kw_col


def set_cell(cell, value, is_keyword=False) -> None:
    """Force Source Sans Pro, Size 9 for all inserted keywords/values."""
    text_str = str(value)
    if cell.text_frame.paragraphs:
        p = cell.text_frame.paragraphs[0]
        if p.runs:
            p.runs[0].text = text_str
            if is_keyword:
                p.runs[0].font.name = "Source Sans Pro"
                p.runs[0].font.size = Pt(9)
            for run in p.runs[1:]:
                run.text = ""
        else:
            p.text = text_str
            if is_keyword and p.runs:
                p.runs[0].font.name = "Source Sans Pro"
                p.runs[0].font.size = Pt(9)
    else:
        cell.text = text_str


def resize_table(table, data_rows: int) -> None:
    tbl = table._tbl
    target = data_rows + 1
    while len(tbl.tr_lst) < target:
        source = tbl.tr_lst[-1]
        tbl.append(copy.deepcopy(source))
    while len(tbl.tr_lst) > target:
        tbl.remove(tbl.tr_lst[-1])


def set_paragraph_text(paragraph, value: str) -> None:
    if paragraph.runs:
        paragraph.runs[0].text = value
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.text = value


def nice_keyword(value: str) -> str:
    text = value.title()
    text = re.sub(r"\b(In|To|Of|The|And|A|With|Near)\b", lambda m: m.group(1).lower(), text)
    return re.sub(r"\s+Al$", ", AL", text, flags=re.I)


def ordinal(n: int) -> str:
    if 11 <= (n % 100) <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def insights(tracked: list[dict]) -> str:
    total = len(tracked)
    
    # Exactly match the requested output text style.
    improved = [x for x in tracked if x["diff"] is not None and x["diff"] > 0]
    declined = [x for x in tracked if x["diff"] is not None and x["diff"] < 0]
    flat = [x for x in tracked if x["pos"] is not None and x["diff"] == 0]
    unranked = [x for x in tracked if x["pos"] is None and x["diff"] is None]

    top_kw = tracked[0] if total > 0 else None
    unchanged_count = len(flat) + len(unranked)
    
    if total == 0:
        return "No tracked keyword ranking data is available for this period."
        
    imp_verb = "keywords have improved" if len(improved) != 1 else "keyword has improved"
    
    # Exact wording requested: "Insights: Overall, we are tracking 8 non-branded keywords. Of these, 0 keywords have improved in rankings, 1 has experienced a drop in positions, and the remaining 7 have shown no significant changes."
    base = (
        f"Insights: Overall, we are tracking {total} non-branded keywords. Of these, {len(improved)} "
        f"{imp_verb} in rankings, {len(declined)} {'have' if len(declined) != 1 else 'has'} experienced a drop in positions, "
        f"and the remaining {unchanged_count} have shown no significant changes."
    )
    
    if top_kw and top_kw["pos"] is not None:
        base += f" Our highest-priority keyword is currently ranking in {ordinal(top_kw['pos'])} position."
    elif top_kw and top_kw["start"] is not None:
        base += (
            f" Our highest-priority keyword does not yet have a recorded ranking position. "
            f"Continued optimization will focus on restoring the ranking of “{top_kw['name']},” "
            f"which previously ranked #{top_kw['start']}, and building visibility across the remaining terms."
        )
    else:
        base += " Our highest-priority keyword does not yet have a recorded ranking position."
        
    return base


def build(template: Path, pdf: Path | None, xlsx: Path, output: Path, full_list_url: str = "") -> None:
    if pdf:
        try:
            reader = PdfReader(str(pdf))
            name = property_name(reader, pdf)
            tracked = parse_tracked_keywords(reader, pdf)
        except Exception as exc:
            raise ValueError(f"Could not read ranking data from the PDF ({pdf.name}): {exc}. Remove the PDF to use Excel ranking data instead.") from exc
    else:
        name = property_name_from_xlsx(xlsx)
        tracked = tracked_from_xlsx(xlsx)
    ancillary, keyword_column = keyword_rows(xlsx, tracked, name)

    prs = Presentation(str(template))
    if len(prs.slides) < 4:
        raise ValueError(f"Template needs at least 4 slides; found {len(prs.slides)}.")

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    original = "".join(run.text for run in paragraph.runs)
                    if "[property name]" in original.lower():
                        set_paragraph_text(paragraph, re.sub(r"\[property name\]", name, original, flags=re.I))

    try:
        slide2_table = next(shape.table for shape in prs.slides[1].shapes if shape.has_table)
        slide3_table = next(shape.table for shape in prs.slides[2].shapes if shape.has_table)
    except StopIteration as exc:
        raise ValueError("Template is missing the expected table on slide 2 or slide 3.") from exc
    if len(slide2_table.columns) < 4 or len(slide3_table.columns) < 3:
        raise ValueError("Template table columns do not match the expected report layout.")

    resize_table(slide2_table, len(tracked))
    for i, item in enumerate(tracked, 1):
        kw_label = f"{i}. {item['name']}"
        if item["pos"] is None and item["diff"] is None:
            values = [kw_label, "N/A", "N/A", "N/A"]
        elif item["pos"] is None and item["diff"] is not None:
            start_val = item["start"] if item["start"] is not None else "N/A"
            values = [kw_label, start_val, "N/A", f'{item["diff"]:+d}']
        else:
            start_val = item["start"] if item["start"] is not None else (item["pos"] + item["diff"])
            diff_str = "Flat" if item["diff"] == 0 else f'{item["diff"]:+d}'
            values = [kw_label, start_val, item["pos"], diff_str]
        for j, value in enumerate(values):
            set_cell(slide2_table.cell(i, j), value, is_keyword=True)

    updated_insight = insights(tracked)
    insight_found = False
    for shape in prs.slides[1].shapes:
        if shape.has_text_frame:
            full_text = shape.text
            
            if "Insights:" in full_text or "This is a summary" in full_text:
                final_text = (
                    "This is a summary of the top 10 rankings for this property since the start of the project.  Ranks of 1-10 appear on page 1 of search results, 11-20 on page 2.\n\n"
                    "For a full list of all keyword rankings, refer to the detailed report on the next slide.\n\n"
                    f"{updated_insight}"
                )
                
                for i in range(len(shape.text_frame.paragraphs)-1, 0, -1):
                    p = shape.text_frame.paragraphs[i]
                    p._p.getparent().remove(p._p)
                    
                set_paragraph_text(shape.text_frame.paragraphs[0], final_text)
                
                shape.text_frame.paragraphs[0].runs[0].font.name = "Source Sans Pro"
                shape.text_frame.paragraphs[0].runs[0].font.size = Pt(9)
                
                insight_found = True
                
    if not insight_found:
        print("Warning: no Insights text box found on slide 2; rankings were added to the table.", file=sys.stderr)

    resize_table(slide3_table, len(ancillary))
    for i, row in enumerate(ancillary, 1):
        for j, value in enumerate((row["keyword"], row["position"], row["volume"])):
            set_cell(slide3_table.cell(i, j), value, is_keyword=True)

    for shape in prs.slides[3].shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in list(paragraph.runs):
                    if "[this list" in run.text.lower():
                        run._r.getparent().remove(run._r)
                    elif run.text.strip().startswith("Full Keyword") and full_list_url:
                        run.hyperlink.address = full_list_url

    if len(prs.slides) > 4:
        fifth = prs.slides[4]
        full_slide5_text = " ".join(s.text for s in fifth.shapes if s.has_text_frame).lower()
        if "[seo team to insert" in full_slide5_text or "seo errors" in full_slide5_text:
            ids = prs.slides._sldIdLst
            sid = list(ids)[4]
            prs.part.drop_rel(sid.rId)
            ids.remove(sid)

    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output))
    print(f"Created: {output}")
    print(f"Property: {name} | Tracked keywords: {len(tracked)} | Additional keywords: {len(ancillary)}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template", help="PowerPoint template (.pptx)")
    parser.add_argument("--pdf", help="Optional Semrush audit report (.pdf); if omitted, ranking data is read from Excel")
    parser.add_argument("--xlsx", help="Keyword list (.xlsx)")
    parser.add_argument("--output", help="Output .pptx path (default: report name in current folder)")
    parser.add_argument("--full-list-url", default="", help="Optional URL for the Full Keyword List link on slide 4")
    args = parser.parse_args()
    try:
        template = choose_file(args.template, ".pptx", "PowerPoint template", optional=True)
        pdf = choose_file(args.pdf, ".pdf", "Semrush PDF", optional=True)
        xlsx = choose_file(args.xlsx, ".xlsx", "keyword Excel", optional=True)
        in_colab = Path("/content").is_dir() or "google.colab" in sys.modules
        if (template is None or xlsx is None) and in_colab:
            from google.colab import files
            print("Upload the PowerPoint template and keyword Excel file. A Semrush PDF is optional.")
            while template is None or xlsx is None:
                files.upload()
                template = choose_file(args.template, ".pptx", "PowerPoint template", optional=True)
                xlsx = choose_file(args.xlsx, ".xlsx", "keyword Excel", optional=True)
                pdf = pdf or choose_file(args.pdf, ".pdf", "Semrush PDF", optional=True)
        if template is None or xlsx is None:
            raise ValueError("Required inputs: one .pptx template and one .xlsx keyword file. A .pdf is optional.")
        name = property_name(PdfReader(str(pdf)), pdf) if pdf else property_name_from_xlsx(xlsx)
        safe_name = re.sub(r'[\\/:*?"<>|]+', "_", name).strip(" .") or "SEO Report"
        output = Path(args.output).expanduser() if args.output else Path.cwd() / f"Monthly SEO Visibility Analysis - {safe_name}.pptx"
        build(template, pdf, xlsx, output, args.full_list_url)
        if Path("/content").is_dir() or "google.colab" in sys.modules:
            from google.colab import files
            files.download(str(output))
    except Exception as exc:
        raise SystemExit(f"ERROR: {exc}") from exc


if __name__ == "__main__":
    main()
