import io
import os
import re
import sys
import time
import subprocessfrom concurrent.futures import ThreadPoolExecutor, as_completedfrom urllib.parse 
import quote, urljoin, urlparse

import pandas as pd
import requests
import streamlit as stfrom pypdf import PdfReader, PdfWriter

try:from playwright.sync_api import sync_playwright as _sync_playwright_PLAYWRIGHT_AVAILABLE = True

# Streamlit Cloud installs the Playwright Python package from requirements.txt,
# but the Chromium browser may still need to be downloaded.
install_result = subprocess.run(
    [sys.executable, "-m", "playwright", "install", "chromium"],
    capture_output=True,
    text=True,
)

if install_result.returncode != 0:
    _PLAYWRIGHT_AVAILABLE = False
    st.error("Playwright Chromium install failed.")
    st.code(install_result.stdout + "\n" + install_result.stderr)

except ImportError as e:_PLAYWRIGHT_AVAILABLE = Falsest.error(f"Playwright is not installed: {e}")

============================================================

Configuration

============================================================

MAX_WORKERS = 2

ZAMBELIS_URL_PATTERNS = ["https://www.zambelislights.gr/image/catalog/sopranos/pdfs/Datasheet_{code}.pdf","https://www.zambelislights.gr/image/catalog/sopranos/pdfs/Datasheet...%20{code}.pdf","https://www.zambelislights.gr/image/catalog/sopranos/pdfs/Datasheet%20...%20{code}.pdf",]

DEFAULT_TIMEOUT = (20, 40)  # connect timeout, read timeoutMAX_WORKERS = 8

============================================================

Helpers

============================================================

def normalize_code(value: str) -> str:"""Clean product code without removing spaces inside the code.

Expected prefixes:
PHL = Philips / Signify product
ZMB = Zambelis product
"""
if value is None:
    return ""

value = str(value).strip()

# Keep letters, numbers, spaces, dash, underscore, and dot
value = re.sub(r"[^A-Za-z0-9 \-_\.]", "", value)

return value.upper()

def get_product_type(code: str) -> str:"""Detect product type from prefix."""if code.startswith("PHL"):return "philips"

if code.startswith("ZMB"):
    return "zambelis"

return "unknown"

def strip_product_prefix(code: str) -> str:"""Remove PHL or ZMB prefix before searching vendor websites.Keeps spaces inside the actual product code."""if code.startswith("PHL"):cleaned = code[3:]elif code.startswith("ZMB"):cleaned = code[3:]else:cleaned = code

cleaned = cleaned.lstrip("-_.")

return cleaned

def extract_codes_from_text(text: str) -> list[str]:"""Extract product codes from manual text input.Supports newline, comma, semicolon, and tab separators.Spaces are kept inside product codes."""if not text:return []

raw_items = re.split(r"[\n,;\t]+", text)
codes = []

for item in raw_items:
    code = normalize_code(item)
    if code:
        codes.append(code)

return codes

def extract_codes_from_excel(uploaded_file, selected_column: str) -> list[str]:"""Extract product codes from the selected Excel column."""df = pd.read_excel(uploaded_file)

if selected_column not in df.columns:
    return []

codes = []

for value in df[selected_column].dropna():
    code = normalize_code(value)
    if code:
        codes.append(code)

return codes

def dedupe_preserve_order(items: list[str]) -> list[str]:"""Remove duplicates while preserving the first appearance order."""seen = set()result = []

for item in items:
    if item not in seen:
        seen.add(item)
        result.append(item)

return result

def is_pdf_bytes(content: bytes) -> bool:"""Check if the downloaded file starts with PDF signature."""return content[:5] == b"%PDF-"

def validate_pdf_content(content: bytes):"""Validate downloaded PDF content."""if not content:raise ValueError("Empty response")

if not is_pdf_bytes(content):
    raise ValueError("Downloaded file does not start with %PDF")

PdfReader(io.BytesIO(content))

def download_philips_datasheet(code: str) -> dict:"""Download one Philips / Signify product leaflet using Playwright.

Improved workflow:
  1. Open Signify and accept cookies.
  2. Search for the stripped Philips code.
  3. Collect multiple possible product pages.
  4. Skip only exact generic category pages, not all indoor-luminaires pages.
  5. Try each product page until a valid PDF leaflet is found.
  6. Validate that the downloaded content is really a PDF.
"""
if not _PLAYWRIGHT_AVAILABLE:
    return {
        "code": code,
        "brand": "Philips",
        "success": False,
        "url": "",
        "error": (
            "Playwright is not installed or Chromium could not be installed."
        ),
        "content": None,
    }

search_code = strip_product_prefix(code)
encoded_code = quote(search_code, safe="")
compact_code = re.sub(r"[^a-z0-9]", "", search_code.lower())

search_urls = [
    f"https://www.signify.com/global/en/search#q={encoded_code}&t=All",
    f"https://www.signify.com/global/search?query={encoded_code}",
]

home_url = "https://www.signify.com/global/en"
product_url = ""
last_error = ""

def absolute_url(raw_url: str, base_url: str = "https://www.signify.com") -> str:
    if not raw_url:
        return ""

    raw_url = raw_url.strip()

    if raw_url.startswith("javascript:") or raw_url.startswith("mailto:") or raw_url.startswith("#"):
        return ""

    if raw_url.startswith("//"):
        return "https:" + raw_url

    if raw_url.startswith("http"):
        return raw_url

    return urljoin(base_url, raw_url)

def compact(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (value or "").lower())

def is_exact_generic_prof_page(url: str) -> bool:
    """
    Reject only exact category pages.
    Do NOT reject every URL containing /indoor-luminaires,
    because real product pages can be below that path.
    """
    path = urlparse(url).path.rstrip("/").lower()

    exact_bad_paths = {
        "/global/prof",
        "/global/prof/indoor-luminaires",
        "/global/prof/outdoor-luminaires",
        "/global/prof/lamps",
        "/global/prof/products",
        "/global/prof/support",
    }

    return path in exact_bad_paths

def looks_like_product_page(url: str) -> bool:
    if not url:
        return False

    parsed = urlparse(url)
    path = parsed.path.lower()
    host = parsed.netloc.lower()

    if "signify.com" not in host:
        return False

    if "/prof/" not in path:
        return False

    if is_exact_generic_prof_page(url):
        return False

    # Product pages are usually deeper than /global/prof/indoor-luminaires
    parts = [part for part in path.split("/") if part]
    if len(parts) < 4:
        return False

    return True

def add_unique(items: list[str], value: str):
    if value and value not in items:
        items.append(value)

def collect_product_candidates(page) -> list[str]:
    candidates = []

    try:
        links = page.locator("a[href]").all()
    except Exception:
        return []

    for link in links[:500]:
        try:
            href = link.get_attribute("href", timeout=1_000)
            text = link.inner_text(timeout=1_000).strip()
        except Exception:
            continue

        full = absolute_url(href)

        if not looks_like_product_page(full):
            continue

        score = 0
        compact_url = compact(full)
        compact_text = compact(text)

        if compact_code and compact_code in compact_url:
            score += 100

        if compact_code and compact_code in compact_text:
            score += 100

        if "product" in compact_text:
            score += 5

        candidates.append((score, full))

    candidates.sort(key=lambda item: item[0], reverse=True)

    result = []
    for _, url in candidates:
        add_unique(result, url)

    return result[:10]

def is_download_like(url: str) -> bool:
    low = (url or "").lower()

    return (
        ".pdf" in low
        or "product_leaflet" in low
        or "product-leaflet" in low
        or "leaflet" in low
        or "datasheet" in low
        or "data-sheet" in low
        or "assets.signify.com" in low
        or "/api/assets" in low
        or "/is/content/signify" in low
    )

def collect_urls_from_text(raw_text: str, base_url: str) -> list[str]:
    urls = []

    if not raw_text:
        return urls

    # Absolute URLs
    for match in re.findall(r"https?://[^\"'<>\s)]+", raw_text):
        add_unique(urls, match)

    # Relative URLs that look like assets or PDFs
    for match in re.findall(
        r"/[^\"'<>\s)]*(?:product_leaflet|product-leaflet|leaflet|datasheet|data-sheet|api/assets|\.pdf)[^\"'<>\s)]*",
        raw_text,
        flags=re.IGNORECASE,
    ):
        add_unique(urls, urljoin(base_url, match))

    return urls

def collect_download_candidates(page, base_url: str) -> list[str]:
    candidates = []

    # 1. Normal links
    try:
        links = page.locator("a[href]").all()
        for link in links[:500]:
            try:
                href = link.get_attribute("href", timeout=1_000)
                text = link.inner_text(timeout=1_000).strip()
            except Exception:
                continue

            full = absolute_url(href, base_url)

            if is_download_like(full) or "leaflet" in text.lower() or "datasheet" in text.lower():
                add_unique(candidates, full)
    except Exception:
        pass

    # 2. Data attributes and onclick values
    for sel in ["button", "[data-href]", "[data-url]", "[data-download-url]"]:
        try:
            nodes = page.locator(sel).all()
            for node in nodes[:300]:
                for attr in ["href", "data-href", "data-url", "data-download-url", "onclick"]:
                    try:
                        value = node.get_attribute(attr, timeout=500)
                    except Exception:
                        value = None

                    if not value:
                        continue

                    for found in collect_urls_from_text(value, base_url):
                        if is_download_like(found):
                            add_unique(candidates, found)
        except Exception:
            pass

    # 3. URLs hidden in page HTML / scripts
    try:
        html = page.content()
        for found in collect_urls_from_text(html, base_url):
            if is_download_like(found):
                add_unique(candidates, found)
    except Exception:
        pass

    return candidates

def pdf_mentions_code(content: bytes) -> bool:
    """
    Optional safety check to avoid downloading the wrong product leaflet.
    If text extraction fails, this returns False but does not crash.
    """
    try:
        reader = PdfReader(io.BytesIO(content))
        text_parts = []

        for page_index, pdf_page in enumerate(reader.pages[:3]):
            text_parts.append(pdf_page.extract_text() or "")

        pdf_text = compact("\n".join(text_parts))
        return compact_code in pdf_text
    except Exception:
        return False

def try_pdf_urls(candidate_urls: list[str], referer: str, page_matches_code: bool):
    last = "No candidate download URLs found."

    for original_url in candidate_urls:
        attempts = [original_url]

        # Some Signify asset URLs include region/language prefixes.
        # Try a few common variants.
        if "assets.signify.com/is/content/" in original_url:
            for region in ("EU.en_AA", "global.en_AA", "US.en_US"):
                variant = re.sub(
                    r"(assets\.signify\.com/is/content/Signify/)[^.]+\.[^.]+\.",
                    rf"\1{region}.",
                    original_url,
                )
                if variant != original_url and variant not in attempts:
                    attempts.append(variant)

        for attempt_url in attempts:
            try:
                response = context.request.get(
                    attempt_url,
                    headers={
                        "Accept": "application/pdf,*/*",
                        "Referer": referer,
                        "User-Agent": (
                            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                            "AppleWebKit/537.36"
                        ),
                    },
                    timeout=40_000,
                )

                content = response.body() if response.ok else b""

                if not response.ok:
                    last = f"HTTP {response.status} for {attempt_url}"
                    continue

                if not content or not is_pdf_bytes(content):
                    last = f"Response is not a PDF for {attempt_url}"
                    continue

                validate_pdf_content(content)

                # Prefer PDFs that are clearly connected to the searched code.
                # If the page itself matched the code, accept the PDF.
                # If not, accept only if the PDF text mentions the code.
                if page_matches_code or pdf_mentions_code(content) or compact_code in compact(attempt_url):
                    return {
                        "code": code,
                        "brand": "Philips",
                        "success": True,
                        "url": attempt_url,
                        "error": "",
                        "content": content,
                    }, ""

                last = (
                    "PDF was found, but the product code was not found "
                    "in the page/PDF, so it was skipped to avoid wrong datasheet."
                )

            except Exception as e:
                last = str(e)

    return None, last

with _sync_playwright() as p:
    browser = p.chromium.launch(
        headless=True,
        args=[
            "--no-sandbox",
            "--disable-dev-shm-usage",
        ],
    )

    context = browser.new_context(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        locale="en-US",
    )

    page = context.new_page()

    try:
        # Step 1: open Signify and accept cookies
        page.goto(home_url, wait_until="domcontentloaded", timeout=30_000)

        for accept_sel in [
            "#onetrust-accept-btn-handler",
            'button:has-text("Accept all")',
            'button:has-text("Accept All")',
            'button:has-text("Accept")',
        ]:
            try:
                page.click(accept_sel, timeout=4_000)
                break
            except Exception:
                pass

        product_candidates = []

        # Step 2: search Signify using both known search URL formats
        for search_url in search_urls:
            try:
                page.goto(search_url, wait_until="domcontentloaded", timeout=45_000)

                try:
                    page.wait_for_load_state("networkidle", timeout=15_000)
                except Exception:
                    pass

                page.wait_for_timeout(5_000)

                # Sometimes download/asset links appear directly in search HTML
                search_downloads = collect_download_candidates(page, search_url)
                result, last_error = try_pdf_urls(
                    search_downloads,
                    referer=search_url,
                    page_matches_code=True,
                )

                if result:
                    return result

                for candidate in collect_product_candidates(page):
                    add_unique(product_candidates, candidate)

            except Exception as e:
                last_error = str(e)

        if not product_candidates:
            return {
                "code": code,
                "brand": "Philips",
                "success": False,
                "url": " | ".join(search_urls),
                "error": (
                    f"No Philips product page candidates found for code: {search_code}. "
                    f"Last error: {last_error}"
                ),
                "content": None,
            }

        # Step 3: try each candidate product page
        for product_url in product_candidates:
            try:
                page.goto(product_url, wait_until="domcontentloaded", timeout=45_000)

                try:
                    page.wait_for_load_state("networkidle", timeout=15_000)
                except Exception:
                    pass

                page.wait_for_timeout(2_000)

                page_text = ""
                try:
                    page_text = page.inner_text("body", timeout=5_000)
                except Exception:
                    pass

                page_matches_code = compact_code in compact(page_text)

                # Open Downloads tab/section if present
                for tab_sel in [
                    'button:has-text("Downloads")',
                    'a:has-text("Downloads")',
                    '[role="tab"]:has-text("Downloads")',
                    'button:has-text("Download")',
                    'a:has-text("Download")',
                ]:
                    try:
                        page.click(tab_sel, timeout=3_000)
                        page.wait_for_timeout(1_500)
                        break
                    except Exception:
                        pass

                download_candidates = collect_download_candidates(page, product_url)

                result, last_error = try_pdf_urls(
                    download_candidates,
                    referer=product_url,
                    page_matches_code=page_matches_code,
                )

                if result:
                    return result

            except Exception as e:
                last_error = str(e)
                continue

        return {
            "code": code,
            "brand": "Philips",
            "success": False,
            "url": " | ".join(product_candidates[:5]),
            "error": (
                f"Tried {len(product_candidates)} Philips product page candidate(s), "
                f"but no valid matching datasheet PDF was found. "
                f"Last error: {last_error}"
            ),
            "content": None,
        }

    except Exception as e:
        return {
            "code": code,
            "brand": "Philips",
            "success": False,
            "url": product_url or " | ".join(search_urls),
            "error": str(e),
            "content": None,
        }

    finally:
        browser.close()

def download_zambelis_datasheet(code: str) -> dict:"""Download one Zambelis datasheet.

The user enters code starting with ZMB.
The ZMB prefix is removed before searching Zambelis.

Search order:
1. Datasheet_{code}.pdf
2. Datasheet...%20{code}.pdf
3. Datasheet%20...%20{code}.pdf
"""
search_code = strip_product_prefix(code)

# Encode only unsafe characters in the product code.
# Keep common product separators safe.
encoded_code = quote(search_code, safe="-_.")

attempted_urls = []
last_error = ""

for pattern in ZAMBELIS_URL_PATTERNS:
    url = pattern.format(code=encoded_code)
    attempted_urls.append(url)

    try:
        response = requests.get(
            url,
            timeout=DEFAULT_TIMEOUT,
            headers={
                "User-Agent": "Mozilla/5.0",
                "Accept": "application/pdf,*/*",
            },
        )

        content_type = response.headers.get("Content-Type", "").lower()
        content = response.content or b""

        if response.status_code != 200:
            last_error = f"HTTP {response.status_code}"
            continue

        if "application/pdf" not in content_type and not is_pdf_bytes(content):
            last_error = f"Not a PDF. Content-Type: {content_type}"
            continue

        validate_pdf_content(content)

        return {
            "code": code,
            "brand": "Zambelis",
            "success": True,
            "url": url,
            "error": "",
            "content": content,
        }

    except Exception as e:
        last_error = str(e)
        continue

return {
    "code": code,
    "brand": "Zambelis",
    "success": False,
    "url": " | ".join(attempted_urls),
    "error": f"Not found after trying all 3 Zambelis links. Last error: {last_error}",
    "content": None,
}

def download_datasheet(code: str) -> dict:"""Route product code to the correct vendor downloader."""product_type = get_product_type(code)

if product_type == "philips":
    return download_philips_datasheet(code)

if product_type == "zambelis":
    return download_zambelis_datasheet(code)

return {
    "code": code,
    "brand": "Unknown",
    "success": False,
    "url": "",
    "error": "Unknown code prefix. Code must start with PHL or ZMB.",
    "content": None,
}

def merge_pdfs(downloads: list[dict]) -> bytes:"""Merge all successful PDFs into one PDF."""writer = PdfWriter()

for item in downloads:
    if not item["success"]:
        continue

    reader = PdfReader(io.BytesIO(item["content"]))

    for page in reader.pages:
        writer.add_page(page)

output = io.BytesIO()
writer.write(output)
return output.getvalue()

============================================================

Streamlit Page Setup

============================================================

st.set_page_config(page_title="Datasheet Pack Builder",page_icon="💡",layout="wide",)

============================================================

Philips + Zambelis Professional CSS

============================================================

st.markdown(""" {/* Philips */--philips-blue: #035ED8;--philips-bright-blue: #0B5ED7;--philips-deep-blue: #003B79;--philips-light-blue: #EAF3FF;

        /* Zambelis inspired premium lighting palette */
        --zambelis-black: #111111;
        --zambelis-charcoal: #252525;
        --zambelis-warm-gray: #6F6A62;
        --zambelis-gold: #C8A45D;
        --zambelis-soft-gold: #F4E8CC;
        --zambelis-cream: #FAF7F0;

        /* Shared UI */
        --app-bg: #F6F8FB;
        --card-bg: #FFFFFF;
        --text-main: #102033;
        --text-muted: #64748B;
        --border-soft: #DDE6F0;
        --success-green: #0F9F6E;
        --warning-orange: #F59E0B;
        --danger-red: #DC2626;
    }

    .stApp {
        background:
            radial-gradient(circle at top left, rgba(3, 94, 216, 0.13), transparent 30%),
            radial-gradient(circle at top right, rgba(200, 164, 93, 0.16), transparent 28%),
            linear-gradient(180deg, #ffffff 0%, var(--app-bg) 46%, var(--zambelis-cream) 100%);
        color: var(--text-main);
    }

    .block-container {
        padding-top: 85px;
        padding-bottom: 60px;
        max-width: 1180px;
    }

    .brand-topbar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 22px;
        gap: 18px;
    }

    .brand-logos {
        display: flex;
        align-items: center;
        gap: 12px;
        flex-wrap: wrap;
    }

    .philips-logo {
        background: #ffffff;
        color: var(--philips-blue);
        border: 2px solid var(--philips-blue);
        border-radius: 4px;
        padding: 8px 15px;
        font-weight: 800;
        font-size: 22px;
        letter-spacing: 1.5px;
        line-height: 1;
        box-shadow: 0 8px 20px rgba(0, 59, 121, 0.10);
    }

    .brand-divider {
        width: 1px;
        height: 32px;
        background: linear-gradient(
            180deg,
            transparent,
            rgba(37, 37, 37, 0.35),
            transparent
        );
    }

    .zambelis-logo {
        background: var(--zambelis-black);
        color: var(--zambelis-gold);
        border: 1px solid rgba(200, 164, 93, 0.65);
        border-radius: 999px;
        padding: 9px 18px;
        font-weight: 700;
        font-size: 18px;
        letter-spacing: 2px;
        line-height: 1;
        box-shadow: 0 8px 22px rgba(17, 17, 17, 0.16);
    }

    .brand-badge {
        color: var(--zambelis-charcoal);
        background: linear-gradient(135deg, var(--philips-light-blue), var(--zambelis-soft-gold));
        border: 1px solid var(--border-soft);
        padding: 8px 14px;
        border-radius: 999px;
        font-size: 13px;
        font-weight: 700;
        box-shadow: 0 8px 18px rgba(15, 23, 42, 0.06);
    }

    .hero {
        position: relative;
        overflow: hidden;
        padding: 34px;
        border-radius: 30px;
        background:
            linear-gradient(
                135deg,
                var(--philips-deep-blue) 0%,
                var(--philips-blue) 44%,
                var(--zambelis-charcoal) 72%,
                var(--zambelis-black) 100%
            );
        color: white;
        margin-bottom: 28px;
        box-shadow:
            0 24px 50px rgba(0, 59, 121, 0.22),
            0 12px 30px rgba(17, 17, 17, 0.18);
    }

    .hero::after {
        content: "";
        position: absolute;
        right: -90px;
        top: -90px;
        width: 260px;
        height: 260px;
        background: rgba(200, 164, 93, 0.22);
        border-radius: 50%;
    }

    .hero::before {
        content: "";
        position: absolute;
        right: 120px;
        bottom: -80px;
        width: 190px;
        height: 190px;
        background: rgba(255, 255, 255, 0.10);
        border-radius: 50%;
    }

    .hero-content {
        position: relative;
        z-index: 2;
        max-width: 780px;
    }

    .hero-kicker {
        text-transform: uppercase;
        letter-spacing: 1.9px;
        font-size: 12px;
        font-weight: 800;
        color: var(--zambelis-soft-gold);
        margin-bottom: 10px;
    }

    .hero h1 {
        margin: 0 0 12px 0;
        font-size: 42px;
        line-height: 1.08;
        font-weight: 850;
    }

    .hero p {
        margin: 0;
        font-size: 17px;
        line-height: 1.6;
        opacity: 0.94;
    }

    .tool-card {
        background: rgba(255, 255, 255, 0.90);
        border: 1px solid var(--border-soft);
        border-radius: 24px;
        padding: 22px;
        box-shadow: 0 16px 40px rgba(15, 23, 42, 0.07);
        backdrop-filter: blur(10px);
        margin-bottom: 18px;
    }

    .section-title {
        font-size: 19px;
        font-weight: 850;
        color: var(--philips-deep-blue);
        margin-bottom: 6px;
    }

    .section-subtitle {
        color: var(--text-muted);
        font-size: 14px;
        margin-bottom: 14px;
    }

    .info-strip {
        background: #ffffff;
        border: 1px solid var(--border-soft);
        border-left: 5px solid var(--zambelis-gold);
        border-radius: 20px;
        padding: 16px 18px;
        margin: 18px 0;
        color: var(--text-muted);
        box-shadow: 0 10px 28px rgba(15, 23, 42, 0.05);
    }

    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid var(--border-soft);
        border-radius: 22px;
        padding: 18px;
        box-shadow: 0 12px 30px rgba(15, 23, 42, 0.06);
    }

    div[data-testid="stMetric"] label {
        color: var(--text-muted) !important;
        font-weight: 650;
    }

    div[data-testid="stMetricValue"] {
        color: var(--philips-deep-blue);
        font-weight: 850;
    }

    .stTextArea textarea,
    .stTextInput input {
        border-radius: 16px !important;
        border-color: var(--border-soft) !important;
    }

    .stTextArea textarea:focus,
    .stTextInput input:focus {
        border-color: var(--philips-blue) !important;
        box-shadow: 0 0 0 1px var(--philips-blue) !important;
    }

    .stSelectbox div[data-baseweb="select"] {
        border-radius: 16px !important;
        border-color: var(--border-soft) !important;
    }

    .stFileUploader {
        background: rgba(255, 255, 255, 0.70);
        border-radius: 18px;
    }

    .stButton > button {
        background:
            linear-gradient(
                135deg,
                var(--philips-blue) 0%,
                var(--philips-deep-blue) 55%,
                var(--zambelis-black) 100%
            );
        color: white;
        border: 0;
        border-radius: 999px;
        padding: 13px 28px;
        font-weight: 850;
        box-shadow:
            0 14px 28px rgba(3, 94, 216, 0.28),
            0 8px 18px rgba(17, 17, 17, 0.15);
        transition: all 0.2s ease;
    }

    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow:
            0 18px 34px rgba(3, 94, 216, 0.32),
            0 10px 22px rgba(17, 17, 17, 0.18);
        color: white;
    }

    .stDownloadButton > button {
        background: #ffffff;
        color: var(--zambelis-black);
        border: 1px solid var(--zambelis-gold);
        border-radius: 999px;
        padding: 13px 28px;
        font-weight: 850;
    }

    .stDownloadButton > button:hover {
        background: var(--zambelis-soft-gold);
        color: var(--zambelis-black);
        border: 1px solid var(--zambelis-gold);
    }

    .small-note {
        color: var(--text-muted);
        font-size: 13px;
    }

    hr {
        border-color: var(--border-soft);
    }

    @media screen and (max-width: 768px) {
        .brand-topbar {
            flex-direction: column;
            align-items: flex-start;
            gap: 12px;
        }

        .brand-logos {
            gap: 10px;
        }

        .brand-divider {
            display: none;
        }

        .hero {
            padding: 26px;
            border-radius: 24px;
        }

        .hero h1 {
            font-size: 32px;
        }

        .hero p {
            font-size: 15px;
        }
    }
</style>
""",
unsafe_allow_html=True,

)

============================================================

Header

============================================================

st.markdown("""PHILIPSZAMBELISPhilips & Zambelis Datasheet Automation

<div class="hero">
    <div class="hero-content">
        <div class="hero-kicker">Product documentation tool</div>
        <h1>Datasheet Pack Builder</h1>
        <p>
            Paste product codes, upload Excel lists, download official datasheets,
            and merge everything into one clean PDF pack.
            Use PHL for Philips products and ZMB for Zambelis products.
        </p>
    </div>
</div>
""",
unsafe_allow_html=True,

)

============================================================

Input Section

============================================================

left_col, right_col = st.columns([1, 1], gap="large")

with left_col:st.markdown("""Paste product codesAdd one or multiple product codes. Use PHL for Philips and ZMB for Zambelis.""",unsafe_allow_html=True,)

manual_codes_text = st.text_area(
    "Product codes",
    placeholder="Example:\nPHL046677568283\nZMB12345",
    height=65,
    label_visibility="collapsed",
)

st.markdown("</div>", unsafe_allow_html=True)

with right_col:st.markdown("""Upload Excel fileSelect the column that contains the product codes.""",unsafe_allow_html=True,)

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xls"],
    label_visibility="collapsed",
)

excel_codes = []

if uploaded_file:
    df_preview = pd.read_excel(uploaded_file)

    st.caption("Excel preview")
    st.dataframe(df_preview.head(), use_container_width=True)

    selected_column = st.selectbox(
        "Choose product code column",
        options=list(df_preview.columns),
    )

    excel_codes = extract_codes_from_excel(uploaded_file, selected_column)

st.markdown("</div>", unsafe_allow_html=True)

============================================================

Options Section

============================================================

st.markdown("""Export settingsChoose the final PDF filename and how failed downloads should be handled.""",unsafe_allow_html=True,)

settings_col_1, settings_col_2 = st.columns([2, 1], gap="large")

with settings_col_1:output_filename = st.text_input("Output PDF filename",value="datasheets pack.pdf",)

with settings_col_2:skip_failed = st.checkbox("Skip failed codes and continue",value=True,)

============================================================

Code Summary

============================================================

manual_codes = extract_codes_from_text(manual_codes_text)all_codes = dedupe_preserve_order(manual_codes + excel_codes)

philips_count = len([code for code in all_codes if get_product_type(code) == "philips"])zambelis_count = len([code for code in all_codes if get_product_type(code) == "zambelis"])unknown_count = len([code for code in all_codes if get_product_type(code) == "unknown"])

st.markdown("### Summary before download")

metric_1, metric_2, metric_3 = st.columns(3)

with metric_1:st.metric("Manual codes", len(manual_codes))

with metric_2:st.metric("Excel codes", len(excel_codes))

with metric_3:st.metric("Unique total", len(all_codes))

brand_metric_1, brand_metric_2, brand_metric_3 = st.columns(3)

with brand_metric_1:st.metric("Philips codes", philips_count)

with brand_metric_2:st.metric("Zambelis codes", zambelis_count)

with brand_metric_3:st.metric("Unknown prefix", unknown_count)

if all_codes:with st.expander("View detected codes"):st.write(all_codes)

============================================================

Download + Merge Action

============================================================

download_button = st.button("Download and merge datasheets",type="primary",disabled=len(all_codes) == 0,)

if download_button:start_time = time.time()

st.info("Downloading datasheets...")

results_by_code = {}
progress_bar = st.progress(0)
status_text = st.empty()

with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
    future_map = {
        executor.submit(download_datasheet, code): code
        for code in all_codes
    }

    completed = 0

    for future in as_completed(future_map):
        code = future_map[future]

        try:
            result = future.result()
        except Exception as e:
            product_type = get_product_type(code)

            if product_type == "philips":
                brand = "Philips"
            elif product_type == "zambelis":
                brand = "Zambelis"
            else:
                brand = "Unknown"

            result = {
                "code": code,
                "brand": brand,
                "success": False,
                "url": "",
                "error": str(e),
                "content": None,
            }

        results_by_code[code] = result

        completed += 1
        progress_bar.progress(completed / len(all_codes))
        status_text.write(f"Processed {completed} / {len(all_codes)}")

# Rebuild results in the same order as the original input
results = [
    results_by_code[code]
    for code in all_codes
    if code in results_by_code
]

successful = [item for item in results if item["success"]]
failed = [item for item in results if not item["success"]]

st.divider()

result_col_1, result_col_2, result_col_3 = st.columns(3)

with result_col_1:
    st.metric("Submitted", len(all_codes))

with result_col_2:
    st.metric("Downloaded", len(successful))

with result_col_3:
    st.metric("Failed", len(failed))

if failed:
    st.warning("Some codes failed.")

    failed_table = pd.DataFrame(
        [
            {
                "Code": item["code"],
                "Brand": item.get("brand", ""),
                "URL": item["url"],
                "Error": item["error"],
            }
            for item in failed
        ]
    )

    st.dataframe(failed_table, use_container_width=True)

    if not skip_failed:
        st.error("Process stopped because failed codes were found.")
        st.stop()

if not successful:
    st.error("No valid PDF datasheets were downloaded.")
    st.stop()

try:
    merged_pdf = merge_pdfs(successful)

    if not output_filename.lower().endswith(".pdf"):
        output_filename += ".pdf"

    elapsed = round(time.time() - start_time, 2)

    st.success(f"PDF pack created successfully in {elapsed} seconds.")

    st.download_button(
        label="Download merged PDF",
        data=merged_pdf,
        file_name=output_filename,
        mime="application/pdf",
    )

except Exception as e:
    st.error(f"Failed to merge PDFs: {e}")
