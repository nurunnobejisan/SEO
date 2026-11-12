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
import io
import re
import sys
import urllib.parse
import urllib.request
from difflib import SequenceMatcher
from pathlib import Path

try:
    import pandas as pd
    from PIL import Image, ImageDraw, ImageFont
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    from pypdf import PdfReader
except ImportError as exc:
    raise SystemExit(
        "Missing a required package. Install them with:\n"
        "python3 -m pip install python-pptx pypdf openpyxl pandas Pillow"
    ) from exc


try:
    import pymupdf  # preferred for reading the Semrush keyword table
except ImportError:  # falls back to pypdf
    try:
        import fitz as pymupdf
    except ImportError:
        pymupdf = None

# STRICT BRAND EXCLUSION RULE - Parent Chains, Sub-Brands, and Common Misspellings
ALL_CHAIN_BRANDS = [
    "marriott", "mariott", "marriot", "bonvoy", "ihg", "hilton", "hyatt",
    "edition", "ritz carlton", "ritz-carlton", "luxury collection", "st regis", "st. regis", "w hotels", "w hotel",
    "sheraton", "delta hotels", "westin", "le meridien", "le méridien", "renaissance", "renassance", "renaisance",
    "autograph collection", "tribute portfolio", "design hotels", "gaylord",
    "courtyard", "four points", "springhill suites", "fairfield", "ac hotels", "ac hotel", "citizenm",
    "aloft", "moxy", "protea", "residence inn", "towneplace suites", "element",
    "six senses", "regent", "intercontinental", "vignette", "kimpton", "hotel indigo",
    "crowne plaza", "even hotels", "voco", "holiday inn", "avid hotels",
    "staybridge suites", "candlewood suites",
    "waldorf astoria", "lxr", "conrad", "nomad", "signia", "canopy",
    "curio collection", "graduate", "doubletree", "tapestry collection", "embassy suites",
    "tempo", "motto", "hampton", "tru", "spark", "homewood suites", "home2 suites", "hilton garden inn",
    "park hyatt", "alila", "miraval", "andaz", "thompson", "the standard",
    "dream hotels", "grand hyatt", "hyatt regency", "hyatt centric", "hyatt place", "hyatt house",
    "urcove", "zoetry", "zoëtry", "secrets resorts",
    "best western", "wyndham", "baymont", "days inn", "super 8", "la quinta", "ramada",
    "comfort inn", "comfort suites", "quality inn", "sleep inn", "clarion", "econo lodge", "rodeway inn",
    "cambria", "radisson", "country inn", "sonesta", "red roof", "motel 6", "drury", "omni", "loews",
    "four seasons", "fairmont", "sofitel", "novotel", "ibis", "mercure", "mandarin oriental",
    "peninsula", "rosewood", "langham", "hard rock", "margaritaville", "virgin hotels", "ace hotel",
    "1 hotel", "proper hotel", "pendry", "montage", "auberge", "provenance", "mcmenamins", "oxford suites", "shilo inn"
]

KNOWN_LOCATIONS = {
    "alabama", "alaska", "arizona", "arkansas", "california", "colorado", "connecticut", "delaware",
    "florida", "georgia", "hawaii", "idaho", "illinois", "indiana", "iowa", "kansas", "kentucky",
    "louisiana", "maine", "maryland", "massachusetts", "michigan", "minnesota", "mississippi",
    "missouri", "montana", "nebraska", "nevada", "new hampshire", "new jersey", "new mexico",
    "new york", "north carolina", "north dakota", "ohio", "oklahoma", "oregon", "pennsylvania",
    "rhode island", "south carolina", "south dakota", "tennessee", "texas", "utah", "vermont",
    "virginia", "washington", "west virginia", "wisconsin", "wyoming",
    "florence", "portland", "denver", "san antonio", "orlando", "cary", "boerne", "alexandria",
    "new orleans", "salt lake city", "fairhope", "omaha", "bloomington", "detroit", "novi",
    "arlington", "tahoe", "atlanta", "austin", "baltimore", "boston", "charleston", "charlotte",
    "chicago", "cincinnati", "cleveland", "columbus", "dallas", "fort worth", "houston",
    "indianapolis", "jacksonville", "kansas city", "las vegas", "los angeles", "louisville",
    "memphis", "miami", "milwaukee", "minneapolis", "nashville", "newark", "oakland",
    "oklahoma city", "philadelphia", "phoenix", "pittsburgh", "raleigh", "richmond",
    "sacramento", "salt lake", "san diego", "san francisco", "san jose", "savannah", "seattle",
    "st louis", "tampa", "tucson", "virginia beach", "washington", "scottsdale", "sedona",
    "napa", "sonoma", "monterey", "carmel", "santa barbara", "palm springs", "anaheim",
    "irvine", "newport beach", "santa monica", "beverly hills", "pasadena", "long beach",
    "honolulu", "waikiki", "maui", "kauai", "key west", "fort lauderdale", "west palm beach",
    "boca raton", "naples", "sarasota", "clearwater", "st petersburg", "daytona", "st augustine",
    "destin", "pensacola", "panama city", "myrtle beach", "hilton head", "wilmington",
    "asheville", "knoxville", "chattanooga", "gatlinburg", "pigeon forge", "lexington",
    "albuquerque", "santa fe", "boise", "spokane", "tacoma", "bellevue", "eugene", "bend", "reno",
    "park city", "vail", "aspen", "breckenridge", "boulder", "colorado springs", "annapolis",
    "providence", "newport", "hartford", "new haven", "burlington", "portsmouth"
}

US_STATE_CODES = {
    "al", "ak", "az", "ar", "ca", "co", "ct", "de", "fl", "ga", "hi", "id", "il", "in", "ia",
    "ks", "ky", "la", "me", "md", "ma", "mi", "mn", "ms", "mo", "mt", "ne", "nv", "nh", "nj",
    "nm", "ny", "nc", "nd", "oh", "ok", "or", "pa", "ri", "sc", "sd", "tn", "tx", "ut", "vt",
    "va", "wa", "wv", "wi", "wy", "dc"
}

AIRPORT_CODES = {
    "pdx", "dca", "iad", "bwi", "den", "atl", "ord", "lax", "dfw", "mco", "msy", "slc", "sat",
    "sfo", "sea", "bos", "mia", "phx", "las", "msp", "dtw", "phl", "clt", "ewr", "jfk", "lga",
    "san", "tpa", "bna", "aus", "rdu", "stl", "mci", "cvg", "cmh", "ind", "mke", "oma"
}

HOSPITALITY_WORDS = {
    "hotel", "hotels", "resort", "resorts", "spa", "spas", "inn", "inns", "suite", "suites",
    "lodging", "accommodation", "accommodations", "stay", "stays", "motel", "motels", "retreat",
    "retreats", "getaway", "getaways", "vacation", "vacations", "place", "places", "boutique",
    "luxury", "luxurious", "modern", "best", "top", "good", "nice", "great", "cheap", "affordable",
    "budget", "discount", "deal", "deals", "package", "packages", "special", "specials", "offer",
    "offers", "rate", "rates", "all", "inclusive", "romantic", "family", "friendly", "pet", "dog",
    "dogs", "adults", "only", "kid", "kids", "couples", "group", "groups", "corporate", "business",
    "executive", "extended", "weekend", "night", "day", "overnight", "star", "stars", "rated",
    "upscale", "unique", "trendy", "cozy", "quiet", "room", "rooms", "bedroom", "bedrooms", "bed",
    "beds", "king", "queen", "double", "view", "views", "balcony", "balconies", "patio", "terrace",
    "rooftop", "pool", "pools", "indoor", "outdoor", "heated", "jacuzzi", "hot", "tub", "tubs",
    "whirlpool", "sauna", "massage", "facial", "wellness", "fitness", "gym", "breakfast", "free",
    "complimentary", "parking", "valet", "shuttle", "wifi", "kitchen", "fireplace", "restaurant",
    "restaurants", "dining", "wedding", "weddings", "venue", "venues", "reception", "receptions",
    "banquet", "ballroom", "ceremony", "bridal", "meeting", "meetings", "conference", "conferences",
    "convention", "event", "events", "space", "spaces", "party", "parties", "things", "thing", "do",
    "see", "visit", "attractions", "activities", "amenities", "amenity", "lodge", "lodges", "villa",
    "villas", "cottage", "cottages", "cabin", "cabins", "trip", "trips", "travel", "guide"
}

GEO_DESCRIPTORS = {
    "downtown", "midtown", "uptown", "old", "town", "historic", "district", "airport",
    "international", "center", "centre", "central", "square", "park", "beach", "lake", "lakeside",
    "river", "riverfront", "riverwalk", "waterfront", "oceanfront", "ocean", "bay", "harbor",
    "island", "mountain", "mountains", "valley", "canyon", "springs", "falls", "coast", "coastal",
    "north", "south", "east", "west", "metro", "area", "local", "city", "village", "university",
    "college", "campus", "stadium", "arena", "theater", "theatre", "museum", "zoo", "aquarium",
    "gardens", "station", "port", "pier", "boardwalk", "strip", "quarter", "landing"
}

CONNECTORS = {
    "in", "near", "by", "at", "on", "to", "for", "with", "of", "the", "a", "an", "and", "or",
    "from", "around", "close", "walking", "distance"
}

MIN_SEARCH_VOLUME = 30
MAX_TRACKED_KEYWORDS = 10  # Capped at 10 as requested
MAX_ANCILLARY_KEYWORDS = 10  # Capped at 10 as requested

SERVICE_PATTERNS: dict[str, str] = {
    "lodging": r"\b(?:hotels?|inns?|motels?|lodging|stays?|places\s+to\s+stay|rooms?)\b",
    "accommodations": r"\b(?:accommodations?)\b",
    "resort": r"\b(?:resorts?|retreats?|getaways?)\b",
    "spa": r"\b(?:spas?|massages?|facials?|saunas?|wellness|day\s+spas?)\b",
    "pool": r"\b(?:pools?|swim(?:ming)?|waterparks?|lazy\s+river)\b",
    "hot_tub": r"\b(?:jacuzzis?|hot\s+tubs?|whirlpools?)\b",
    "suites": r"\b(?:suites?|extended\s+stay|kitchenettes?)\b",
    "boutique": r"\b(?:boutique)\b",
    "luxury": r"\b(?:luxury|luxurious|5\s+star|five\s+star|4\s+star|four\s+star|upscale)\b",
    "pet_friendly": r"\b(?:pets?|pet\s+friendly|dogs?|dog\s+friendly)\b",
    "wedding": r"\b(?:weddings?|bridal|ceremon(?:y|ies)|receptions?)\b",
    "meetings_events": r"\b(?:meetings?|conferences?|conventions?|ballrooms?|banquets?|events?)\b",
    "venues": r"\b(?:venues?)\b",
    "dining": r"\b(?:restaurants?|dining|bistros?|steakhouses?)\b",
    "breakfast": r"\b(?:breakfast|brunch)\b",
    "things_to_do": r"\b(?:things?\s+to\s+do|attractions?|activities|sightseeing)\b",
    "waterfront": r"\b(?:waterfront|riverfront|riverwalk|beachfront|oceanfront|lakefront|lakeside)\b",
    "golf": r"\b(?:golf(?:ing)?)\b",
    "casino": r"\b(?:casinos?|gaming|slots?)\b",
    "ski": r"\b(?:ski(?:ing)?|snowboard(?:ing)?)\b",
    "all_inclusive": r"\b(?:all[\s\-]+inclusive)\b",
    "family": r"\b(?:famil(?:y|ies)|kids?)\b",
    "romantic": r"\b(?:romantic|couples?|honeymoon)\b",
    "packages_deals": r"\b(?:packages?|deals?|specials?|offers?|discounts?)\b",
}

SERVICE_LABELS: dict[str, str] = {
    "lodging": "Hotel & Lodging",
    "accommodations": "Accommodations",
    "resort": "Resort & Getaways",
    "spa": "Spa & Wellness",
    "pool": "Pool & Aquatics",
    "hot_tub": "Hot Tub / Jacuzzi",
    "suites": "Suites",
    "boutique": "Boutique Hotel",
    "luxury": "Luxury / Upscale",
    "pet_friendly": "Pet-Friendly",
    "wedding": "Weddings",
    "meetings_events": "Meetings & Events",
    "venues": "Event Venues",
    "dining": "Restaurants & Dining",
    "breakfast": "Breakfast",
    "things_to_do": "Local Attractions / Things to Do",
    "waterfront": "Waterfront / Riverfront",
    "golf": "Golf",
    "casino": "Casino",
    "ski": "Ski",
    "all_inclusive": "All-Inclusive",
    "family": "Family-Friendly",
    "romantic": "Romantic / Couples",
    "packages_deals": "Packages & Deals",
}


def detect_services_in_text(text: str) -> set[str]:
    cleaned = re.sub(r"[^a-z0-9\s\-]", " ", (text or "").lower())
    found: set[str] = set()
    for svc, pat in SERVICE_PATTERNS.items():
        if re.search(pat, cleaned, re.I):
            found.add(svc)
    return found


def extract_workbook_profile(xlsx_path: Path) -> dict:
    """Extract the services, ranked keywords, and primary URL from the Keywords Workbook."""
    try:
        frame = pd.read_excel(xlsx_path)
    except Exception:
        return {"services": set(), "keywords": [], "primary_url": "", "urls": []}
    if frame.empty:
        return {"services": set(), "keywords": [], "primary_url": "", "urls": []}
    frame.columns = [str(c).strip() for c in frame.columns]
    kw_cols = [c for c in frame.columns if any(w in c.lower() for w in ("keyword", "query", "search term"))]
    pos_cols = [c for c in frame.columns if any(w in c.lower() for w in ("position", "rank"))]
    url_cols = [c for c in frame.columns if c.lower() in ("url", "landing page", "page url")]

    if kw_cols and pos_cols:
        kw_col, pos_col = kw_cols[0], pos_cols[0]
        frame["_kw"] = frame[kw_col].fillna("").astype(str).str.strip()
        frame["_pos"] = pd.to_numeric(frame[pos_col], errors="coerce")
        valid = frame[(frame["_kw"] != "") & frame["_pos"].notna()]
        keywords = valid["_kw"].tolist()
    elif kw_cols:
        keywords = frame[kw_cols[0]].fillna("").astype(str).str.strip().tolist()
    else:
        keywords = []

    urls: list[str] = []
    primary_url = ""
    if url_cols:
        raw_urls = frame[url_cols[0]].dropna().astype(str).str.strip().tolist()
        urls = [u for u in raw_urls if u.startswith("http")]
        if urls:
            # Strip query strings for the cleanest canonical target URL
            clean_urls = [u.split("?")[0] for u in urls]
            primary_url = pd.Series(clean_urls).mode().iloc[0] if clean_urls else urls[0]

    combined_kw_text = "\n".join(keywords)
    wb_services = detect_services_in_text(combined_kw_text)
    return {
        "services": wb_services,
        "keywords": keywords,
        "primary_url": primary_url,
        "urls": urls,
    }


def _fetch_html(url: str, timeout: float = 5.0) -> tuple[str, str]:
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read(600_000).decode("utf-8", errors="ignore")
        return raw, resp.geturl()


def _html_to_crawl_text(html: str) -> tuple[str, str]:
    title_m = re.search(r"<title[^>]*>(.*?)</title>", html, re.I | re.S)
    title = clean_text(re.sub(r"<[^>]+>", " ", title_m.group(1))) if title_m else ""
    meta_parts = re.findall(
        r'<meta[^>]+(?:name|property)=["\'](?:description|og:title|og:description|keywords)["\'][^>]+content=["\']([^"\']+)["\']',
        html,
        re.I,
    )
    meta_parts += re.findall(
        r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+(?:name|property)=["\'](?:description|og:title|og:description|keywords)["\']',
        html,
        re.I,
    )
    cleaned_html = re.sub(r"<(script|style|noscript|svg)[^>]*>.*?</\1>", " ", html, flags=re.I | re.S)
    visible = clean_text(re.sub(r"<[^>]+>", " ", cleaned_html))
    return title, " ".join([title] + meta_parts + [visible])


def crawl_target_website(
    target_url: str,
    fallback_text: str = "",
    workbook_urls: list[str] | None = None,
    workbook_keywords: list[str] | None = None,
) -> dict:
    """Crawl the target hotel website (and key service subpages) to identify supported services."""
    url = (target_url or "").strip()
    if url and not re.match(r"^https?://", url, re.I):
        url = "https://" + url

    crawled_pages = 0
    page_title = ""
    collected_chunks: list[str] = []

    if url:
        # Always include the URL path slugs themselves (e.g., 'renaissance-shoals-resort-and-spa')
        parsed = urllib.parse.urlparse(url)
        slug_text = re.sub(r"[^a-z0-9]+", " ", parsed.path.lower())
        collected_chunks.append(slug_text)

        try:
            html, final_url = _fetch_html(url, timeout=5.0)
            if "access denied" not in html.lower()[:600]:
                t, txt = _html_to_crawl_text(html)
                page_title = t
                collected_chunks.append(txt)
                crawled_pages += 1

                # Discover up to 3 internal service links on the same domain
                base_netloc = urllib.parse.urlparse(final_url).netloc.lower()
                hrefs = re.findall(r'href=["\']([^"\'#]+)["\']', html, re.I)
                sub_candidates: list[str] = []
                service_slug_re = re.compile(
                    r"(?:amenit|spa|dining|restaurant|meeting|wedding|event|room|suite|pool|accommodat|overview|experience)",
                    re.I,
                )
                for h in hrefs:
                    full_u = urllib.parse.urljoin(final_url, h).split("?")[0].rstrip("/")
                    u_parsed = urllib.parse.urlparse(full_u)
                    if (
                        u_parsed.scheme in ("http", "https")
                        and u_parsed.netloc.lower() == base_netloc
                        and full_u != final_url.rstrip("/")
                        and service_slug_re.search(u_parsed.path)
                        and not re.search(r"\.(?:pdf|jpg|jpeg|png|gif|webp|css|js)$", u_parsed.path, re.I)
                        and full_u not in sub_candidates
                    ):
                        sub_candidates.append(full_u)
                        collected_chunks.append(re.sub(r"[^a-z0-9]+", " ", u_parsed.path.lower()))

                for sub_url in sub_candidates[:3]:
                    try:
                        sub_html, _ = _fetch_html(sub_url, timeout=3.5)
                        if "access denied" not in sub_html.lower()[:600]:
                            _, sub_txt = _html_to_crawl_text(sub_html)
                            collected_chunks.append(sub_txt)
                            crawled_pages += 1
                    except Exception:
                        pass
        except Exception:
            pass

    if workbook_urls:
        for wu in workbook_urls[:100]:
            try:
                p_path = urllib.parse.urlparse(wu).path.lower()
                collected_chunks.append(re.sub(r"[^a-z0-9]+", " ", p_path))
            except Exception:
                pass

    # If the live HTTP request was blocked by WAF (crawled_pages == 0) and workbook URLs match target_url's domain,
    # include the indexed workbook keywords for that target domain as proven website content
    if crawled_pages == 0 and workbook_keywords:
        collected_chunks.extend(workbook_keywords)

    if fallback_text:
        collected_chunks.append(fallback_text)

    combined = "\n".join(collected_chunks)
    website_services = detect_services_in_text(combined)
    if url or combined.strip():
        website_services.add("lodging")

    return {
        "target_url": url,
        "page_title": page_title,
        "crawled_pages": crawled_pages,
        "services": website_services,
    }


def keyword_matches_services(
    keyword: str,
    allowed_services: set[str],
    workbook_keywords: list[str] | None = None,
    require_workbook_match: bool = False,
) -> bool:
    """Return True if all services referenced in `keyword` match `allowed_services` (website ∩ workbook)."""
    kw_clean = re.sub(r"[^a-z0-9\s\-]", " ", keyword.lower()).strip()
    kw_services = detect_services_in_text(kw_clean)

    # If the keyword mentions any service category, every mentioned service must be supported
    if kw_services:
        if not kw_services.issubset(allowed_services):
            return False
    elif require_workbook_match:
        return False

    if require_workbook_match and workbook_keywords:
        wb_norm = {re.sub(r"\s+", " ", k.lower()).strip() for k in workbook_keywords}
        if kw_clean in wb_norm:
            return True
        # Check if any workbook keyword shares the same service category in the workbook
        wb_services_all = detect_services_in_text("\n".join(workbook_keywords))
        if not kw_services.issubset(wb_services_all):
            return False

    return True


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
    first = ""
    if pymupdf is not None and pdf_path is not None:
        try:
            doc = pymupdf.open(str(pdf_path))
            if len(doc) > 0:
                first = clean_text(doc[0].get_text())
        except Exception:
            pass
    if not first and reader.pages:
        first = clean_text(reader.pages[0].extract_text())
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


def extract_property_brand_and_locations(
    property_title: str,
    tracked: list[dict] | None = None,
    df_keywords: list[str] | None = None,
    from_pdf: bool = True,
) -> tuple[set[str], set[str]]:
    prop_clean = re.sub(r"[^a-z0-9\s]", " ", property_title.lower())
    prop_tokens = [w for w in prop_clean.split() if len(w) > 1]

    title_locs: set[str] = set()
    for loc in KNOWN_LOCATIONS:
        if re.search(r"\b" + re.escape(loc) + r"\b", prop_clean):
            for w in loc.split():
                title_locs.add(w)

    # Also detect if any token in property_title is used as a city in 'hotels in <city>' / 'downtown <city>'
    sample_kws: list[str] = []
    if from_pdf and tracked:
        sample_kws.extend(item["name"] for item in tracked)
    if df_keywords:
        sample_kws.extend(str(k) for k in df_keywords[:500])
    if sample_kws:
        joined_samples = "\n".join(re.sub(r"[^a-z0-9\s]", " ", k.lower()) for k in sample_kws)
        for w in prop_tokens:
            if len(w) >= 3 and w not in HOSPITALITY_WORDS and w not in CONNECTORS:
                if re.search(
                    r"\b(?:hotels?\s+in|stays?\s+in|places\s+to\s+stay\s+in|things\s+to\s+do\s+in|downtown)\s+"
                    + re.escape(w)
                    + r"\b",
                    joined_samples,
                ):
                    title_locs.add(w)

    ignore_prop = HOSPITALITY_WORDS | CONNECTORS | GEO_DESCRIPTORS | title_locs
    prop_brand_words: set[str] = set()
    for w in prop_tokens:
        if w not in ignore_prop and len(w) >= 3:
            prop_brand_words.add(w)

    tracked_locs: set[str] = set()
    if from_pdf and tracked:
        for item in tracked:
            kw_clean = re.sub(r"[^a-z0-9\s]", " ", item["name"].lower())
            for w in kw_clean.split():
                if (
                    w not in HOSPITALITY_WORDS
                    and w not in CONNECTORS
                    and w not in GEO_DESCRIPTORS
                    and w not in prop_brand_words
                    and len(w) >= 2
                ):
                    tracked_locs.add(w)

    inferred_locs: set[str] = set()
    if df_keywords:
        loc_counts: dict[str, int] = {}
        for kw in df_keywords:
            kw_c = re.sub(r"[^a-z0-9\s]", " ", str(kw).lower())
            words = kw_c.split()
            for w in words:
                if (w in KNOWN_LOCATIONS or w in AIRPORT_CODES) and w not in prop_brand_words:
                    loc_counts[w] = loc_counts.get(w, 0) + 1
        for w, cnt in sorted(loc_counts.items(), key=lambda x: -x[1])[:4]:
            if cnt >= 2:
                inferred_locs.add(w)

        known_cities = (tracked_locs | title_locs | inferred_locs) - US_STATE_CODES
        for kw in df_keywords:
            kw_c = re.sub(r"[^a-z0-9\s]", " ", str(kw).lower())
            words = kw_c.split()
            for idx in range(len(words) - 1):
                if words[idx] in known_cities and words[idx + 1] in US_STATE_CODES:
                    inferred_locs.add(words[idx + 1])

    target_locations = (tracked_locs | title_locs | inferred_locs) - prop_brand_words - CONNECTORS
    return prop_brand_words, target_locations


def is_brand_excluded(
    keyword: str,
    property_title: str,
    prop_brand_words: set[str] | None = None,
    target_locations: set[str] | None = None,
    strict_vocab: bool = True,
) -> bool:
    """Return True if the keyword should be EXCLUDED from non-branded location-based Ancillary keywords."""
    kw_clean = re.sub(r"[^a-z0-9\s]", " ", keyword.lower())
    kw_clean = re.sub(r"\s+", " ", kw_clean).strip()
    kw_nospace = re.sub(r"[^a-z0-9]", "", keyword.lower())
    tokens = kw_clean.split()
    if not tokens:
        return True

    if prop_brand_words is None or target_locations is None:
        prop_brand_words, target_locations = extract_property_brand_and_locations(property_title)

    # 1. Exclude physical street addresses or 5-digit zip codes
    if (
        re.search(r"\d+\s+[a-z]+\s+[a-z]+.*\b\d{5}\b", kw_clean)
        or re.search(r"^\d+\s+[a-z0-9]+\s+(?:st|street|ave|avenue|rd|road|blvd|pl|place|dr|drive|ln|lane|way|pkwy|hwy)\b", kw_clean)
        or re.search(r"\b\d{5}\b", kw_clean)
    ):
        return True

    # 2. Exclude major hotel chain brands, spaceless concatenations, and typos
    for brand in ALL_CHAIN_BRANDS:
        brand_clean = re.sub(r"[^a-z0-9\s]", " ", brand.lower()).strip()
        brand_nospace = re.sub(r"[^a-z0-9]", "", brand.lower())
        if re.search(r"\b" + re.escape(brand_clean) + r"\b", kw_clean):
            return True
        if len(brand_nospace) >= 5 and brand_nospace in kw_nospace:
            return True
        if " " not in brand_clean and len(brand_clean) >= 5:
            for t in tokens:
                if len(t) >= 4 and SequenceMatcher(None, t, brand_clean).ratio() >= 0.80:
                    return True

    # 3. Exclude property-specific brand words, spaceless variants (e.g. 'thenines'), and typos (e.g. 'tiatro', 'treato')
    for bw in prop_brand_words:
        if re.search(r"\b" + re.escape(bw) + r"\b", kw_clean):
            return True
        if len(bw) >= 4 and bw in kw_nospace:
            return True
        for t in tokens:
            if len(t) >= 4 and len(bw) >= 4 and SequenceMatcher(None, t, bw).ratio() >= 0.80:
                return True

    # 4. Exclude restaurant menus, bars/grills, reviews, photos, room counts, and '+' lounge names
    bad_terms = {
        "menu", "menus", "reviews", "review", "photos", "photo", "pictures", "count",
        "address", "phone", "number", "careers", "jobs", "login", "bar and grill",
        "bar and grille", "steakhouse",
    }
    for bt in bad_terms:
        if re.search(r"\b" + re.escape(bt) + r"\b", kw_clean):
            return True
    if "+" in keyword:
        return True

    # 5. Require a core hospitality / travel / event intent word
    core_intent = {
        "hotel", "hotels", "resort", "resorts", "spa", "spas", "inn", "inns", "suite", "suites",
        "lodging", "accommodation", "accommodations", "stay", "stays", "motel", "motels",
        "retreat", "retreats", "getaway", "getaways", "boutique", "wedding", "weddings",
        "venue", "venues", "meeting", "meetings", "conference", "package", "packages",
        "deal", "deals", "things", "places", "restaurants",
    }
    if not any(t in core_intent for t in tokens):
        return True

    # Only allow restaurant keywords if paired with hotel(s) or general discovery ('restaurants in/near', 'rooftop restaurants')
    if ("restaurant" in tokens or "restaurants" in tokens) and not ("hotel" in tokens or "hotels" in tokens):
        if not re.search(r"\b(?:restaurants\s+(?:in|near)|rooftop\s+restaurants?)\b", kw_clean):
            return True

    # 6. Must be LOCATION-BASED (contain a target location word for this property, or a 'near <landmark>' phrase)
    has_target_loc = any(t in target_locations for t in tokens)
    has_near = "near" in tokens and tokens[-1] != "near" and "near me" not in kw_clean
    if not (has_target_loc or has_near):
        return True

    # 7. Reject unknown proper nouns (on-site restaurant client names like 'departure', 'urban farmer', 'nickel',
    # building/venue names like 'meier & frank', 'fantinos', competitor names, or unrelated cities)
    if strict_vocab:
        allowed_vocab = HOSPITALITY_WORDS | GEO_DESCRIPTORS | CONNECTORS | target_locations
        check_tokens = tokens[:tokens.index("near")] if has_near else tokens
        for t in check_tokens:
            if t not in allowed_vocab:
                return True
    else:
        # Tier-2 fallback for local neighborhood modifiers: require both a target location word and a strict lodging/venue word
        strict_lodging = {
            "hotel", "hotels", "resort", "resorts", "spa", "spas", "inn", "inns",
            "suite", "suites", "lodging", "stay", "stays", "boutique", "wedding", "weddings", "venue", "venues",
        }
        if not (has_target_loc and any(t in strict_lodging for t in tokens)):
            return True

    return False


def _stem_word(w: str) -> str:
    if w.endswith("ies") and len(w) > 4:
        return w[:-3] + "y"
    if w.endswith("s") and not w.endswith("ss") and len(w) > 3:
        return w[:-1]
    return w


def _keyword_topic_signature(kw: str, target_locs: set[str]) -> tuple:
    tokens = re.sub(r"[^a-z0-9\s]", " ", kw.lower()).split()
    if "near" in tokens:
        near_idx = tokens.index("near")
        after = [_stem_word(t) for t in tokens[near_idx + 1:] if t not in CONNECTORS and t not in target_locs]
        return ("near", tuple(sorted(after)) or ("loc",))
    core = sorted({_stem_word(t) for t in tokens if t not in CONNECTORS and t not in target_locs})
    return tuple(core)

def _decode_semrush_pua(text: str) -> str:
    out = []
    for ch in text or "":
        code = ord(ch)
        if 0xE071 <= code <= 0xE07A:
            out.append(chr(ord("0") + (code - 0xE071)))
        else:
            out.append(ch)
    return "".join(out)


def _extract_pdf_pages(reader: PdfReader, pdf_path: Path | None = None, prefer_pymupdf: bool = True) -> list[str]:
    """Extract text per page using pymupdf first (preferred for Semrush tables), falling back to pypdf."""
    if prefer_pymupdf and pymupdf is not None and pdf_path is not None:
        try:
            pages = [_decode_semrush_pua(page.get_text() or "") for page in pymupdf.open(str(pdf_path))]
            if any(p.strip() for p in pages):
                return pages
        except Exception:
            pass
    return [_decode_semrush_pua(p.extract_text() or "") for p in reader.pages]


def _parse_tracked_from_pages(pages: list[str]) -> list[dict]:
    start = next((i for i, text in enumerate(pages) if "Keyword Rankings" in text), None)
    if start is None:
        raise ValueError("The PDF has no 'Keyword Rankings' section. Check that this is the expected Semrush report.")
    text = "\n".join(pages[start:])
    section = text.split("Keyword Rankings", 1)[1]
    
    section = re.split(r"\b(Site Audit|Backlinks:?|Organic Research|Traffic Analytics)\b", section, maxsplit=1, flags=re.I)[0]
    total_match = (
        re.search(r"\d+\s*-\s*\d+\s*\((\d+)\)", section)
        or re.search(r"^\s*\d+\s+\d+\s+(\d+)\b", section)
        or re.search(r"Keyword\s+Rankings\s+\d+\s+\d+\s+(\d+)\b", section, re.I)
    )
    expected = int(total_match.group(1)) if total_match else None

    page_break_re = re.compile(
        r"^(?:"
        r"800\s+BOYLSTON\s+STREET.*"
        r"|WWW\.SEMRUSH\.COM.*"
        r"|Generated\s+on\s+.*"
        r"|.*The\s+report\s+data\s+is\s+taken\s+from\s+Semrush.*"
        r")$",
        re.I,
    )
    header_noise_re = re.compile(
        r"^(?:"
        r"(?:Keyword\s+Rankings\s*)?(?:\d+\s*-\s*\d+\s*\(\d+\)|\d+\s+\d+\s+\d+)?"
        r"|Keyword(?:\s+SERP\s+Features)?(?:\s+KD%?)?"
        r"|SERP\s+Features(?:\s+KD%?)?"
        r"|KD|%?"
        r"|(?:Position|Diff|URL|KD%|CPC|Volume)(?:\s+(?:Position|Diff|URL|KD%|CPC|Volume))*"
        r")$",
        re.I,
    )

    entries: list[dict] = []
    current = None
    in_page_break = False
    for line in (clean_text(line) for line in section.splitlines()):
        if not line:
            continue
        if re.match(r"^\d+\.\s+\S", line):
            current = {"name_parts": [re.sub(r"^\d+\.\s+", "", line)], "values": [], "finished_name": False}
            entries.append(current)
            in_page_break = False
        elif current is None:
            continue
        elif page_break_re.match(line):
            current["finished_name"] = True
            in_page_break = True
            continue
        elif header_noise_re.match(line):
            current["finished_name"] = True
            in_page_break = False
            continue
        elif re.match(r"^\d+(?:st|nd|rd|th)\s+time$", line, re.I):
            current["finished_name"] = True
            in_page_break = False
            continue
        else:
            tokens = line.replace("\x00", "+").split()
            if tokens and all(re.fullmatch(VALUE_RE, t, re.I) for t in tokens):
                # Ignore a standalone page number right after a page footer/header or after values are complete
                if len(tokens) == 1 and (in_page_break or len(current["values"]) >= 5):
                    in_page_break = False
                    continue
                in_page_break = False
                if len(current["values"]) < 6:
                    current["values"].extend(tokens)
                    current["finished_name"] = True
            elif not current.get("finished_name"):
                if (
                    re.search(r"https?://|www\.|\.com\b|\.org\b|\.net\b|/", line, re.I)
                    or re.match(r"^(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s*\d", line, re.I)
                ):
                    current["finished_name"] = True
                else:
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
    return result


def parse_tracked_keywords(reader: PdfReader, pdf_path: Path | None = None) -> list[dict]:
    pages = _extract_pdf_pages(reader, pdf_path, prefer_pymupdf=True)
    result = _parse_tracked_from_pages(pages)
    if not result and pymupdf is not None and pdf_path is not None:
        fallback_pages = _extract_pdf_pages(reader, pdf_path, prefer_pymupdf=False)
        result = _parse_tracked_from_pages(fallback_pages)
    if not result:
        raise ValueError("Could not read keyword rows from the PDF. No fallback keyword data was inserted.")
    return result


def tracked_from_xlsx(
    xlsx_path: Path,
    prop_name: str = "",
    allowed_services: set[str] | None = None,
) -> list[dict]:
    frame = pd.read_excel(xlsx_path)
    frame.columns = [str(c).strip() for c in frame.columns]
    keyword_cols = [c for c in frame.columns if any(x in c.lower() for x in ("keyword", "query", "search term"))]
    position_cols = [c for c in frame.columns if any(x in c.lower() for x in ("position", "rank"))]
    if not keyword_cols or not position_cols:
        raise ValueError("Without a PDF, the Excel file must contain keyword and position/rank columns.")
    kw_col, pos_col = keyword_cols[0], position_cols[0]
    frame["_kw"] = frame[kw_col].fillna("").astype(str).str.strip()
    frame["_pos"] = pd.to_numeric(frame[pos_col], errors="coerce")
    frame = frame[(frame["_kw"] != "") & frame["_pos"].notna()].drop_duplicates("_kw")
    if prop_name:
        prop_brands, target_locs = extract_property_brand_and_locations(
            prop_name, None, frame["_kw"].tolist(), from_pdf=False
        )
        non_branded = frame[
            ~frame["_kw"].apply(lambda kw: is_brand_excluded(kw, prop_name, prop_brands, target_locs))
        ]
        if not non_branded.empty:
            frame = non_branded
    if allowed_services:
        svc_matched = frame[
            frame["_kw"].apply(lambda kw: keyword_matches_services(kw, allowed_services))
        ]
        if not svc_matched.empty:
            frame = svc_matched
    frame = frame.head(MAX_TRACKED_KEYWORDS)
    if frame.empty:
        raise ValueError("No keyword rows with a numeric position were found in the Excel file.")
    return [{"name": row["_kw"], "start": int(row["_pos"]), "pos": int(row["_pos"]), "diff": 0} for _, row in frame.iterrows()]


def keyword_rows(
    xlsx_path: Path,
    tracked: list[dict],
    prop_name: str = "",
    from_pdf: bool = True,
    allowed_services: set[str] | None = None,
) -> tuple[list[dict], str]:
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

    # Extract property brand identifiers and target location terms
    prop_brands, target_locs = extract_property_brand_and_locations(
        prop_name, tracked, frame["_keyword"].tolist(), from_pdf=from_pdf
    )

    # Optional URL-column check to block dedicated restaurant/bar/menu PDF URLs for non-hotel queries
    url_cols = [c for c in frame.columns if c.lower() in ("url", "landing page", "page url")]
    if url_cols:
        url_col = url_cols[0]
        frame["_url"] = frame[url_col].fillna("").astype(str).str.lower()
        frame = frame[
            ~(
                frame["_url"].str.contains(r"\.pdf\b|/menu\b|/menus\b", regex=True)
                | (
                    frame["_url"].str.contains(r"/dining/|/restaurant/|/restaurants/", regex=True)
                    & ~frame["_keyword"].str.lower().str.contains(r"\b(?:hotels?|resorts?|spas?)\b", regex=True)
                )
            )
        ]

    # Apply strict brand, client/restaurant name, and location-based filtering for Ancillary keywords
    strict_frame = frame[
        ~frame["_keyword"].apply(
            lambda kw: is_brand_excluded(kw, prop_name, prop_brands, target_locs, strict_vocab=True)
        )
    ]

    if len(strict_frame) < MAX_ANCILLARY_KEYWORDS:
        fallback_frame = frame[
            ~frame["_keyword"].apply(
                lambda kw: is_brand_excluded(kw, prop_name, prop_brands, target_locs, strict_vocab=False)
            )
        ]
        frame = pd.concat([strict_frame, fallback_frame]).drop_duplicates("_normalized")
    else:
        frame = strict_frame

    if allowed_services:
        svc_filtered = frame[
            frame["_keyword"].apply(lambda kw: keyword_matches_services(kw, allowed_services))
        ]
        if not svc_filtered.empty:
            frame = svc_filtered

    # Select diverse location-based keywords (capping identical topic stems at 2 so 'near' & diverse terms appear)
    candidates = []
    for position, group in frame.groupby("_position", sort=True):
        candidates.extend(group.head(4 if position <= 2 else 2).to_dict("records"))

    topic_counts: dict[tuple, int] = {}
    selected = []
    overflow = []
    for row in candidates:
        sig = _keyword_topic_signature(row["_keyword"], target_locs)
        if topic_counts.get(sig, 0) < 2:
            topic_counts[sig] = topic_counts.get(sig, 0) + 1
            selected.append(row)
        else:
            overflow.append(row)

    selected = selected[:MAX_ANCILLARY_KEYWORDS]
    if len(selected) < MAX_ANCILLARY_KEYWORDS:
        selected.extend(overflow[: MAX_ANCILLARY_KEYWORDS - len(selected)])

    # Ensure at least one 'near' location keyword is included when available in the filtered set
    near_rows = [
        r for r in frame.to_dict("records")
        if re.search(r"\bnear\b", r["_keyword"].lower())
    ]
    if near_rows and not any(re.search(r"\bnear\b", r["_keyword"].lower()) for r in selected):
        if len(selected) >= MAX_ANCILLARY_KEYWORDS:
            selected[-1] = near_rows[0]
        else:
            selected.append(near_rows[0])

    selected.sort(key=lambda r: (int(r["_position"]), -int(r["_volume"])))
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


def duplicate_slide(prs: Presentation, source_index: int, insert_index: int):
    """Clone prs.slides[source_index] and insert the new slide at insert_index."""
    source_slide = prs.slides[source_index]
    new_slide = prs.slides.add_slide(source_slide.slide_layout)

    sp_tree = new_slide.shapes._spTree
    for child in list(sp_tree)[2:]:
        sp_tree.remove(child)

    rId_map = {}
    for rel in source_slide.part.rels.values():
        if "slideLayout" in rel.reltype or "notesSlide" in rel.reltype:
            continue
        if rel.is_external:
            new_rId = new_slide.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            new_rId = new_slide.part.rels.get_or_add(rel.reltype, rel.target_part)
        rId_map[rel.rId] = new_rId

    next_shape_id = 100
    for child in list(source_slide.shapes._spTree)[2:]:
        new_el = copy.deepcopy(child)
        for el in new_el.iter():
            if el.tag.endswith("}cNvPr") and "id" in el.attrib:
                el.set("id", str(next_shape_id))
                next_shape_id += 1
            for attr_name, attr_val in list(el.attrib.items()):
                if attr_val in rId_map:
                    el.set(attr_name, rId_map[attr_val])
        sp_tree.append(new_el)

    sldIdLst = prs.slides._sldIdLst
    new_sldId = sldIdLst[-1]
    sldIdLst.remove(new_sldId)
    sldIdLst.insert(insert_index, new_sldId)
    return prs.slides[insert_index]


def populate_tracked_table(table, items: list[dict], start_number: int = 1) -> None:
    resize_table(table, len(items))
    for i, item in enumerate(items, 1):
        kw_label = f"{start_number + i - 1}. {item['name']}"
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
            set_cell(table.cell(i, j), value, is_keyword=True)


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


def parse_site_audit(reader: PdfReader, pdf_path: Path | None = None) -> dict | None:
    pages = _extract_pdf_pages(reader, pdf_path)
    full_doc = "\n".join(pages)
    url_match = re.search(r"https?://[^\s|)\x00]+", full_doc)
    audited_url = url_match.group(0).rstrip(".,;)") if url_match else ""

    start = next((i for i, t in enumerate(pages) if "Site Audit" in t or "Top Issues" in t), None)
    if start is None:
        return None

    section = "\n".join(pages[start:])
    if "Site Audit" in section:
        section = section.split("Site Audit", 1)[1]
    section = re.split(r"\b(Backlinks:?|Organic Research|Traffic Analytics)\b", section, maxsplit=1, flags=re.I)[0]

    health_m = re.search(r"Site\s+Health\s+(\d+%)", section, re.I)
    errors_m = re.search(r"\bErrors\s+(\d[\d,]*)", section, re.I)
    warnings_m = re.search(r"\bWarnings\s+(\d[\d,]*)", section, re.I)
    notices_m = re.search(r"\bNotices\s+(\d[\d,]*)", section, re.I)
    total_pages_m = re.search(r"\bTotal\s+(\d[\d,]*)", section, re.I)

    site_health = health_m.group(1) if health_m else "N/A"
    errors_count = int(errors_m.group(1).replace(",", "")) if errors_m else 0
    warnings_count = int(warnings_m.group(1).replace(",", "")) if warnings_m else 0
    notices_count = int(notices_m.group(1).replace(",", "")) if notices_m else 0
    crawled_pages = int(total_pages_m.group(1).replace(",", "")) if total_pages_m else 0

    top_issues: list[dict] = []
    if "Number Of Issues" in section:
        issues_part = section.split("Number Of Issues", 1)[1]
    elif "Status Of Issues" in section:
        issues_part = section.split("Status Of Issues", 1)[1]
    else:
        issues_part = ""

    noise_re = re.compile(
        r"^(?:"
        r"800\s+BOYLSTON\s+STREET.*"
        r"|WWW\.SEMRUSH\.COM.*"
        r"|Generated\s+on\s+.*"
        r"|.*The\s+report\s+data\s+is\s+taken\s+from\s+Semrush.*"
        r"|Top\s+Issues(?:\s+Status\s+Of\s+Issues)?(?:\s+Number\s+Of\s+Issues)?"
        r"|Status\s+Of\s+Issues|Number\s+Of\s+Issues"
        r"|\d+"
        r")$",
        re.I,
    )

    raw_lines = [clean_text(ln) for ln in issues_part.splitlines() if clean_text(ln)]
    idx = 0
    desc_parts: list[str] = []
    while idx < len(raw_lines):
        ln = raw_lines[idx]
        # Case 1: 'Error 1' or 'Warning 10' on one line
        m_inline = re.match(r"^(Error|Warning|Notice)s?\s+(\d[\d,]*)(?:\s+([+\-]?\d[\d,]*))?$", ln, re.I)
        # Case 2: 'Error' or 'Warning' alone on one line followed by count on next line
        m_alone = re.match(r"^(Error|Warning|Notice)s?$", ln, re.I)

        if m_inline and desc_parts:
            status = m_inline.group(1).capitalize()
            count = m_inline.group(2)
            diff = m_inline.group(3) or "0"
            if not m_inline.group(3) and idx + 1 < len(raw_lines) and re.fullmatch(r"[+\-]?\d[\d,]*", raw_lines[idx + 1]):
                diff = raw_lines[idx + 1]
                idx += 1
            top_issues.append(
                {
                    "issue": clean_text(" ".join(desc_parts)),
                    "status": status,
                    "count": count,
                    "diff": diff,
                }
            )
            desc_parts = []
        elif m_alone and desc_parts and idx + 1 < len(raw_lines) and re.fullmatch(r"\d[\d,]*", raw_lines[idx + 1]):
            status = m_alone.group(1).capitalize()
            count = raw_lines[idx + 1]
            idx += 1
            diff = "0"
            if idx + 1 < len(raw_lines) and re.fullmatch(r"[+\-]?\d[\d,]*", raw_lines[idx + 1]):
                diff = raw_lines[idx + 1]
                idx += 1
            top_issues.append(
                {
                    "issue": clean_text(" ".join(desc_parts)),
                    "status": status,
                    "count": count,
                    "diff": diff,
                }
            )
            desc_parts = []
        elif noise_re.match(ln):
            pass
        else:
            desc_parts.append(ln)
        idx += 1

    return {
        "site_health": site_health,
        "errors": errors_count,
        "warnings": warnings_count,
        "notices": notices_count,
        "crawled_pages": crawled_pages,
        "top_issues": top_issues,
        "audited_url": audited_url,
    }


def _load_pil_font(size: int, bold: bool = False):
    candidates = (
        [
            "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
            "/System/Library/Fonts/Helvetica.ttc",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        ]
        if bold
        else [
            "/System/Library/Fonts/Supplemental/Arial.ttf",
            "/System/Library/Fonts/Helvetica.ttc",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        ]
    )
    for path in candidates:
        if Path(path).exists():
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    try:
        return ImageFont.load_default(size=size)
    except TypeError:
        return ImageFont.load_default()


def crop_top_issues_from_pdf(pdf_path: Path | None) -> tuple[bytes, int, int] | None:
    """Crop the actual 'Top Issues' table directly from the Semrush PDF at 250 DPI using PyMuPDF."""
    if pymupdf is None or pdf_path is None or not pdf_path.is_file():
        return None
    try:
        doc = pymupdf.open(str(pdf_path))
        strips: list[Image.Image] = []
        started = False
        for page in doc:
            blocks = page.get_text("blocks")
            if started and any("Backlinks" in b[4] for b in blocks if b[1] < 200):
                break
            hdr_block = next(
                (b for b in blocks if "Top Issues" in b[4] and "Status Of Issues" in b[4]),
                None,
            )
            if hdr_block is None:
                continue
            row_blocks = [b for b in blocks if b[1] > hdr_block[3] and b[3] < 805]
            if not row_blocks:
                continue
            max_y = max(b[3] for b in row_blocks) + 9.0
            if not started:
                clip_top = hdr_block[1] - 8.5
                started = True
            else:
                clip_top = hdr_block[3] + 7.5
            rect = pymupdf.Rect(26.0, clip_top, 570.0, max_y)
            pix = page.get_pixmap(dpi=250, clip=rect, alpha=False)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            strips.append(img)
        if not strips:
            return None
        total_w = strips[0].width
        total_h = sum(s.height for s in strips)
        combined = Image.new("RGB", (total_w, total_h), (255, 255, 255))
        y = 0
        for s in strips:
            combined.paste(s, (0, y))
            y += s.height
        buf = io.BytesIO()
        combined.save(buf, format="PNG")
        return buf.getvalue(), total_w, total_h
    except Exception:
        return None


def render_top_issues_image(top_issues: list[dict]) -> tuple[bytes, int, int]:
    """Render a Semrush 'Top Issues' table screenshot matching the template style and size."""
    rows = top_issues or [
        {"issue": "No critical top issues reported in this audit period", "status": "Notice", "count": "0", "diff": "0"}
    ]
    scale = 3
    width = 512 * scale
    header_h = 28 * scale
    row_h = 37 * scale
    height = header_h + row_h * len(rows) + 3 * scale

    img = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    font_hdr = _load_pil_font(10 * scale, bold=False)
    font_issue = _load_pil_font(11 * scale, bold=False)
    font_badge = _load_pil_font(10 * scale, bold=False)
    font_count = _load_pil_font(11 * scale, bold=False)
    font_diff = _load_pil_font(9 * scale, bold=False)

    draw.rectangle([0, 0, width, header_h], fill=(246, 248, 250))
    draw.line([0, header_h, width, header_h], fill=(230, 234, 240), width=scale)

    col1_x = 10 * scale
    col2_x = 250 * scale
    col3_right = width - 12 * scale

    draw.text((col1_x, 8 * scale), "Top Issues", fill=(65, 75, 88), font=font_hdr)
    draw.text((col2_x, 8 * scale), "Status Of Issues", fill=(65, 75, 88), font=font_hdr)
    hdr3_w = draw.textbbox((0, 0), "Number Of Issues", font=font_hdr)[2]
    draw.text((col3_right - hdr3_w, 8 * scale), "Number Of Issues", fill=(65, 75, 88), font=font_hdr)

    y = header_h
    for item in rows:
        words = item["issue"].split()
        lines: list[str] = []
        cur = ""
        for w in words:
            test = (cur + " " + w).strip()
            if draw.textbbox((0, 0), test, font=font_issue)[2] <= 232 * scale:
                cur = test
            else:
                if cur:
                    lines.append(cur)
                cur = w
        if cur:
            lines.append(cur)
        lines = lines[:2]

        if len(lines) == 1:
            draw.text((col1_x, y + 11 * scale), lines[0], fill=(34, 40, 49), font=font_issue)
        else:
            draw.text((col1_x, y + 5 * scale), lines[0], fill=(34, 40, 49), font=font_issue)
            draw.text((col1_x, y + 19 * scale), lines[1], fill=(34, 40, 49), font=font_issue)

        st = item["status"]
        if st.lower() == "error":
            bg_col = (253, 232, 232)
            fg_col = (224, 82, 82)
        elif st.lower() == "warning":
            bg_col = (254, 240, 230)
            fg_col = (232, 124, 62)
        else:
            bg_col = (232, 244, 253)
            fg_col = (43, 125, 233)

        tw = draw.textbbox((0, 0), st, font=font_badge)[2]
        pad_x = 8 * scale
        pill_w = tw + pad_x * 2
        pill_h = 16 * scale
        pill_y = y + 10 * scale
        draw.rounded_rectangle([col2_x, pill_y, col2_x + pill_w, pill_y + pill_h], radius=8 * scale, fill=bg_col)
        draw.text((col2_x + pad_x, pill_y + 2 * scale), st, fill=fg_col, font=font_badge)

        cnt_str = str(item.get("count", "0"))
        diff_str = str(item.get("diff", "0"))
        cnt_w = draw.textbbox((0, 0), cnt_str, font=font_count)[2]
        draw.text((col3_right - cnt_w, y + 6 * scale), cnt_str, fill=(34, 40, 49), font=font_count)
        diff_w = draw.textbbox((0, 0), diff_str, font=font_diff)[2]
        draw.text((col3_right - diff_w, y + 21 * scale), diff_str, fill=(136, 146, 160), font=font_diff)

        y += row_h
        draw.line([0, y, width, y], fill=(235, 238, 243), width=scale)

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue(), width, height


def ensure_seo_errors_slide(prs: Presentation):
    """Return existing SEO Errors Report slide or create one from Slide 4 if the template only has 4 slides."""
    if len(prs.slides) > 4:
        return prs.slides[4]
    new_slide = duplicate_slide(prs, 3, 4)
    for shape in list(new_slide.shapes):
        if shape.shape_type == 14 and shape.has_text_frame:
            set_paragraph_text(shape.text_frame.paragraphs[0], "SEO Errors Report")
        elif shape.shape_type == 17 and shape.top > 1000000:
            shape._element.getparent().remove(shape._element)

    card = new_slide.shapes.add_shape(1, 5551900, 1402100, 2781900, 3164700)
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(0xEE, 0xEE, 0xEE)
    card.line.fill.background()

    note_box = new_slide.shapes.add_textbox(1299350, 830300, 4683900, 548700)
    note_box.text_frame.word_wrap = True
    note_box.text_frame.paragraphs[0].text = "NOTE FOR SEO TEAM:"

    exp_box = new_slide.shapes.add_textbox(5655600, 1550675, 2613600, 2993700)
    exp_box.text_frame.word_wrap = True
    exp_box.text_frame.paragraphs[0].text = "[SEO Team to insert a short boilerplate explanation]"
    return new_slide


def _strip_highlights_and_spacing(shape) -> None:
    """Remove template yellow <a:highlight> markers and excessive paragraph spacing from a text shape."""
    for el in list(shape._element.iter()):
        if el.tag.endswith("}highlight") or el.tag.endswith("}spcBef") or el.tag.endswith("}spcAft"):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)


def populate_seo_errors_slide(
    errors_slide,
    audit: dict | None,
    full_list_url: str = "",
    pdf_path: Path | None = None,
) -> None:
    """Populate Slide 5 (SEO Errors Report) with the follow-up note, Top Issues screenshot, and client summary."""
    note_text = (
        "NOTE FOR SEO TEAM:  This slide only applies to stand-alone sites for soft brands or independents.  "
        "Remove this slide for brand hotels."
    )

    if audit:
        err_cnt = audit["errors"]
        warn_cnt = audit["warnings"]
        health = audit["site_health"]
        health_line = (
            f"Site Health is currently at {health} with {err_cnt} Errors and {warn_cnt} Warnings identified."
            if health and health != "N/A"
            else f"The current audit identified {err_cnt} Errors and {warn_cnt} Warnings."
        )
    else:
        health_line = "Our technical SEO audit monitors site health across all crawled pages for Errors and Warnings."

    para1_text = (
        f"{health_line} “Errors” are high-priority technical issues (such as broken images or incorrect hreflang "
        f"links) that directly impact search engine crawlability and indexation, while “Warnings” are on-page "
        f"items (such as missing alt or <h> tags, lengthy title tags, or large JS/CSS files) that affect page "
        f"speed and keyword relevance."
    )
    para2_text = (
        "What We Fix & Impact: For hosted sites, our SEO team directly resolves all technical errors and "
        "warnings—including adding missing <h> headings and image alt attributes—to improve crawl efficiency and "
        "rankings. When a warning requires new page copy or applies to an externally hosted site (such as low "
        "text-to-HTML ratio or code compression), our Account Manager coordinates directly with the site host."
    )
    report_url = full_list_url or (audit.get("audited_url", "") if audit else "")

    for shape in list(errors_slide.shapes):
        if shape.has_text_frame:
            txt_lower = shape.text.lower()
            if "note for seo team" in txt_lower:
                _strip_highlights_and_spacing(shape)
                set_paragraph_text(shape.text_frame.paragraphs[0], note_text)
                if shape.text_frame.paragraphs[0].runs:
                    shape.text_frame.paragraphs[0].runs[0].font.name = "Source Sans Pro"
                    shape.text_frame.paragraphs[0].runs[0].font.size = Pt(10.5)
            elif "boilerplate" in txt_lower or ("errors" in txt_lower and "warnings" in txt_lower and "seo team" in txt_lower):
                for i in range(len(shape.text_frame.paragraphs) - 1, 0, -1):
                    p = shape.text_frame.paragraphs[i]
                    p._p.getparent().remove(p._p)

                p0 = shape.text_frame.paragraphs[0]
                for child in list(p0._p):
                    if child.tag.endswith("}br"):
                        p0._p.remove(child)
                set_paragraph_text(p0, f"{para1_text}\n\n{para2_text}")
                _strip_highlights_and_spacing(shape)
                if p0.runs:
                    p0.runs[0].font.name = "Source Sans Pro"
                    p0.runs[0].font.size = Pt(9)
                    p0.runs[0].font.color.rgb = RGBColor(0x22, 0x22, 0x22)

                p_link = shape.text_frame.add_paragraph()
                link_run = p_link.add_run()
                link_run.text = "\nFull Error & Warning Audit Report"
                link_run.font.name = "Source Sans Pro"
                link_run.font.size = Pt(9)
                link_run.font.bold = True
                link_run.font.color.rgb = RGBColor(0x4B, 0x28, 0x82)
                if report_url:
                    link_run.hyperlink.address = report_url

    # Replace or add the Top Issues screenshot picture on Slide 5
    target_pic = None
    for shape in errors_slide.shapes:
        if shape.shape_type == 13 and shape.left < 2000000 and shape.top > 1200000:
            target_pic = shape
            break

    cropped = crop_top_issues_from_pdf(pdf_path)
    if cropped is not None:
        png_bytes, img_w, img_h = cropped
    elif audit and audit.get("top_issues"):
        png_bytes, img_w, img_h = render_top_issues_image(audit["top_issues"])
    else:
        png_bytes = None

    if png_bytes is not None:
        if target_pic is not None:
            left, top, width = target_pic.left, target_pic.top, target_pic.width
            target_pic._element.getparent().remove(target_pic._element)
        else:
            left, top, width = 696125, 1827925, 4683900
        new_height = int(width * (img_h / img_w))
        max_height = 5143500 - top - 250000
        if new_height > max_height:
            new_height = max_height
        errors_slide.shapes.add_picture(io.BytesIO(png_bytes), left, top, width=width, height=new_height)


def build(
    template: Path,
    pdf: Path | None,
    xlsx: Path,
    output: Path,
    full_list_url: str = "",
    target_url: str = "",
) -> None:
    audit_data = None
    used_pdf_tracked = False
    wb_profile = extract_workbook_profile(xlsx)

    if pdf:
        reader = PdfReader(str(pdf))
        try:
            name = property_name(reader, pdf)
        except Exception:
            name = property_name_from_xlsx(xlsx)
        try:
            audit_data = parse_site_audit(reader, pdf)
        except Exception as exc:
            print(f"Warning: PDF Site Audit parsing failed ({exc}).", file=sys.stderr)
            audit_data = None
    else:
        name = property_name_from_xlsx(xlsx)

    effective_target_url = (
        (target_url or "").strip()
        or (audit_data.get("audited_url", "") if audit_data else "")
        or wb_profile.get("primary_url", "")
    )
    crawl_info = crawl_target_website(
        effective_target_url,
        fallback_text=name,
        workbook_urls=wb_profile.get("urls"),
        workbook_keywords=wb_profile.get("keywords"),
    )
    if crawl_info["services"] and wb_profile["services"]:
        allowed_services = crawl_info["services"] & wb_profile["services"]
        if "lodging" in wb_profile["services"]:
            allowed_services.add("lodging")
    else:
        allowed_services = wb_profile["services"] or crawl_info["services"]

    if pdf:
        try:
            tracked = parse_tracked_keywords(reader, pdf)
            used_pdf_tracked = True
            if allowed_services:
                matched_tracked = [
                    item
                    for item in tracked
                    if keyword_matches_services(
                        item["name"],
                        allowed_services,
                        workbook_keywords=wb_profile["keywords"],
                        require_workbook_match=True,
                    )
                ]
                if matched_tracked:
                    tracked = matched_tracked
                else:
                    tracked = tracked_from_xlsx(xlsx, name, allowed_services=allowed_services)
                    used_pdf_tracked = False
        except Exception as exc:
            print(f"Warning: PDF keyword parsing failed ({exc}); falling back to Excel.", file=sys.stderr)
            tracked = tracked_from_xlsx(xlsx, name, allowed_services=allowed_services)
    else:
        tracked = tracked_from_xlsx(xlsx, name, allowed_services=allowed_services)

    ancillary, keyword_column = keyword_rows(
        xlsx,
        tracked,
        name,
        from_pdf=used_pdf_tracked,
        allowed_services=allowed_services,
    )

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
        ancillary_table = next(shape.table for shape in prs.slides[2].shapes if shape.has_table)
    except StopIteration as exc:
        raise ValueError("Template is missing the expected table on slide 2 or slide 3.") from exc
    if len(slide2_table.columns) < 4 or len(ancillary_table.columns) < 3:
        raise ValueError("Template table columns do not match the expected report layout.")

    benefits_slide = prs.slides[3]
    errors_slide = ensure_seo_errors_slide(prs)
    populate_seo_errors_slide(errors_slide, audit_data, full_list_url, pdf_path=pdf)

    # Update Insights on Slide 2 before duplicating for overflow slides
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

    # Paginate tracked keywords across Slide 2 (and Slide 3+ if len(tracked) > 10)
    tracked_chunks = [
        tracked[idx : idx + MAX_TRACKED_KEYWORDS]
        for idx in range(0, len(tracked), MAX_TRACKED_KEYWORDS)
    ] or [[]]

    for page_idx, chunk in enumerate(tracked_chunks):
        start_num = page_idx * MAX_TRACKED_KEYWORDS + 1
        if page_idx == 0:
            populate_tracked_table(slide2_table, chunk, start_number=start_num)
        else:
            overflow_slide = duplicate_slide(prs, 1, 1 + page_idx)
            overflow_table = next(shape.table for shape in overflow_slide.shapes if shape.has_table)
            populate_tracked_table(overflow_table, chunk, start_number=start_num)

    resize_table(ancillary_table, len(ancillary))
    for i, row in enumerate(ancillary, 1):
        for j, value in enumerate((row["keyword"], row["position"], row["volume"])):
            set_cell(ancillary_table.cell(i, j), value, is_keyword=True)

    for shape in benefits_slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in list(paragraph.runs):
                    if "[this list" in run.text.lower():
                        run._r.getparent().remove(run._r)
                    elif run.text.strip().startswith("Full Keyword") and full_list_url:
                        run.hyperlink.address = full_list_url

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
    parser.add_argument("--target-url", default="", help="Optional Target Website URL to crawl and match services against the Keywords Workbook")
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
        build(template, pdf, xlsx, output, args.full_list_url, target_url=args.target_url)
        if Path("/content").is_dir() or "google.colab" in sys.modules:
            from google.colab import files
            files.download(str(output))
    except Exception as exc:
        raise SystemExit(f"ERROR: {exc}") from exc


if __name__ == "__main__":
    main()
