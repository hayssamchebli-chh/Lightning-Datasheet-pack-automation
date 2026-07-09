import io
import os
import re
import sys
import time
import threading
import subprocess
from html import unescape
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import quote, urljoin, urlparse

import pandas as pd
import requests
import streamlit as st
from pypdf import PdfReader, PdfWriter
from pypdf.annotations import Link
from pypdf.generic import Fit
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas as pdf_canvas

# ============================================================
# Streamlit Page Setup - must be the first Streamlit command
# ============================================================

st.set_page_config(
    page_title="Datasheet Pack Builder",
    page_icon="💡",
    layout="wide",
)

# ============================================================
# Optional Playwright setup
# ============================================================

try:
    from playwright.sync_api import sync_playwright as _sync_playwright
    _PLAYWRIGHT_AVAILABLE = True
    _PLAYWRIGHT_IMPORT_ERROR = ""
except ImportError as e:
    _sync_playwright = None
    _PLAYWRIGHT_AVAILABLE = False
    _PLAYWRIGHT_IMPORT_ERROR = str(e)


@st.cache_resource(show_spinner=False)
def ensure_playwright_chromium() -> tuple[bool, str]:
    """Ensure Chromium is available for Playwright.

    Streamlit Cloud can install the Playwright Python package from requirements.txt,
    while the Chromium browser binary may still need to be installed separately.
    """
    if not _PLAYWRIGHT_AVAILABLE:
        return False, f"Playwright is not installed: {_PLAYWRIGHT_IMPORT_ERROR}"

    try:
        install_result = subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            capture_output=True,
            text=True,
            timeout=180,
            check=False,
        )
    except Exception as e:
        return False, f"Playwright Chromium install failed: {e}"

    if install_result.returncode != 0:
        details = (install_result.stdout or "") + "\n" + (install_result.stderr or "")
        return False, "Playwright Chromium install failed.\n" + details.strip()

    return True, ""


PLAYWRIGHT_READY, PLAYWRIGHT_ERROR = ensure_playwright_chromium()
if not PLAYWRIGHT_READY:
    st.warning(PLAYWRIGHT_ERROR)

# ============================================================
# Configuration
# ============================================================

MAX_WORKERS = 4
DEFAULT_TIMEOUT = (8, 15)  # connect timeout, read timeout

ZAMBELIS_URL_PATTERNS = [
    "https://www.zambelislights.gr/image/catalog/sopranos/pdfs/Datasheet_{code}.pdf",
    "https://www.zambelislights.gr/image/catalog/sopranos/pdfs/Datasheet...%20{code}.pdf",
    "https://www.zambelislights.gr/image/catalog/sopranos/pdfs/Datasheet%20...%20{code}.pdf",
]

SIGNIFY_SEARCH_API_URL = "https://api.microservices.signify.com/api/product/v1/smc/en_AA/search"

# Cover page inserted before each item's datasheet in the merged PDF
COVER_TEMPLATE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "item_type_template.pdf",
)
COVER_TEXT_COLOR = "#1F4EA1"  # same blue as the Harb Electric logo
COVER_TEXT_X = 42  # left aligned with the logo
COVER_TEXT_TOP_OFFSET = 170  # distance of the first line from the top of the page
COVER_TEXT_MAX_WIDTH = 340  # keep the text inside the white area
COVER_TEXT_FONT = "Helvetica-Bold"
COVER_TEXT_FONT_SIZE = 34
COVER_TEXT_MIN_FONT_SIZE = 18

# Table of contents at the beginning of the merged PDF
TOC_LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "toc_logo.png")
TOC_ACCENT_COLOR = "#1F4EA1"
TOC_TITLE_COLOR = "#102033"
TOC_DOTS_COLOR = "#9AA7B5"
TOC_MARGIN_X = 48
TOC_ENTRY_SPACING = 28
TOC_ENTRIES_FIRST_PAGE = 18
TOC_ENTRIES_LATER_PAGES = 22

FUMAGALLI_DOWNLOADS_URL = "https://www.fumagalli.it/en/downloads/"
FUMAGALLI_CATALOG_TTL = 3600  # refresh the cached product list every hour
FUMAGALLI_MAX_PAGES = 6

# Description words that never appear in Fumagalli catalogue names
FUMAGALLI_NOISE_TOKENS = {
    "mod", "model", "cm", "mm", "d", "diam", "diameter", "h",
    "grey", "gray", "black", "white", "green", "brown", "beige",
    "anthracite", "rust", "antique", "opal", "clear", "smoked",
}

# ============================================================
# Helpers
# ============================================================


def normalize_code(value: str) -> str:
    """Clean product code without removing spaces inside the code.

    Expected prefixes:
    PHL = Philips / Signify product
    ZMB = Zambelis product
    """
    if value is None:
        return ""

    value = str(value).strip()
    value = re.sub(r"[^A-Za-z0-9 \-_\.]", "", value)
    return value.upper()


def get_product_type(code: str) -> str:
    """Detect product type from prefix."""
    if code.startswith("PHL"):
        return "philips"
    if code.startswith("ZMB"):
        return "zambelis"
    return "unknown"


def strip_product_prefix(code: str) -> str:
    """Remove PHL or ZMB prefix before searching vendor websites."""
    if code.startswith(("PHL", "ZMB")):
        cleaned = code[3:]
    else:
        cleaned = code

    return cleaned.lstrip("-_.")


def extract_codes_from_text(text: str) -> list[str]:
    """Extract product codes from manual text input."""
    if not text:
        return []

    raw_items = re.split(r"[\n,;\t]+", text)
    codes = []

    for item in raw_items:
        code = normalize_code(item)
        if code:
            codes.append(code)

    return codes


def extract_items_from_excel(uploaded_file) -> tuple[list[dict], str]:
    """Parse Type / Code / Description rows from the uploaded Excel file.

    Expected columns (matched by name, falling back to column position):
    1. Type        - written on the cover page before the item's datasheet
    2. Code        - PHL/ZMB codes search by code; FUM codes use the Description
    3. Description - FUMAGALLI product name / full description

    Returns (items, error_message). Each item is a dict with:
    kind ("code" or "fumagalli"), value (what to search), type, display.
    """
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file)

    if df.empty:
        return [], "The Excel file has no rows."

    columns = list(df.columns)

    def find_column(keyword: str, fallback_index: int):
        for column in columns:
            if keyword in str(column).strip().lower():
                return column
        if len(columns) > fallback_index:
            return columns[fallback_index]
        return None

    type_col = find_column("type", 0)
    code_col = find_column("code", 1)
    desc_col = find_column("desc", 2)

    if code_col is None:
        return [], "Could not find a Code column in the Excel file."

    def cell_text(row, column) -> str:
        if column is None:
            return ""
        value = row.get(column)
        if value is None or pd.isna(value):
            return ""
        return re.sub(r"\s+", " ", str(value)).strip()

    items = []

    for _, row in df.iterrows():
        code = normalize_code(cell_text(row, code_col))
        type_text = cell_text(row, type_col)
        description = normalize_product_name(cell_text(row, desc_col))

        if not code:
            continue

        if code.startswith("FUM"):
            items.append(
                {
                    "kind": "fumagalli",
                    "value": description,
                    "type": type_text,
                    "display": f"{code} - {description}" if description else code,
                }
            )
        else:
            items.append(
                {
                    "kind": "code",
                    "value": code,
                    "type": type_text,
                    "display": code,
                }
            )

    return items, ""


def normalize_product_name(value: str) -> str:
    """Clean a FUMAGALLI product name without changing its casing."""
    if value is None:
        return ""

    value = str(value).strip()
    value = re.sub(r"[^A-Za-z0-9 \-_\.&/]", "", value)
    return re.sub(r"\s+", " ", value).strip()


def extract_names_from_text(text: str) -> list[str]:
    """Extract FUMAGALLI product names from manual text input."""
    if not text:
        return []

    raw_items = re.split(r"[\n;\t]+", text)
    names = []

    for item in raw_items:
        name = normalize_product_name(item)
        if name:
            names.append(name)

    return names


def dedupe_names_preserve_order(items: list[str]) -> list[str]:
    """Remove duplicate names case-insensitively, preserving first appearance."""
    seen = set()
    result = []

    for item in items:
        key = item.casefold()
        if key not in seen:
            seen.add(key)
            result.append(item)

    return result


def drop_untyped_duplicates(items: list[dict]) -> list[dict]:
    """Apply the duplicates rule to the item list.

    Items WITH a Type always keep their own cover page + datasheet, even
    when several items share the same code/description. Items WITHOUT a
    Type are included only once: repeated untyped occurrences are dropped,
    and an untyped occurrence is also dropped when the same product appears
    elsewhere with a Type (its datasheet is already in the pack).
    """
    typed_keys = {
        (item["kind"], str(item["value"]).casefold())
        for item in items
        if item.get("type", "").strip()
    }

    seen_untyped = set()
    result = []

    for item in items:
        key = (item["kind"], str(item["value"]).casefold())

        if not item.get("type", "").strip():
            if key in typed_keys or key in seen_untyped:
                continue
            seen_untyped.add(key)

        result.append(item)

    return result


def dedupe_preserve_order(items: list[str]) -> list[str]:
    """Remove duplicates while preserving the first appearance order."""
    seen = set()
    result = []

    for item in items:
        if item not in seen:
            seen.add(item)
            result.append(item)

    return result


def is_pdf_bytes(content: bytes) -> bool:
    """Check if the downloaded file starts with PDF signature."""
    return bool(content) and content[:5] == b"%PDF-"


def validate_pdf_content(content: bytes) -> None:
    """Validate downloaded PDF content."""
    if not content:
        raise ValueError("Empty response")

    if not is_pdf_bytes(content):
        raise ValueError("Downloaded file does not start with %PDF")

    PdfReader(io.BytesIO(content))


def download_philips_datasheet_api(code: str) -> dict | None:
    """Fast path: resolve a Philips code through the public Signify product API.

    This is the same API the signify.com search page uses. It matches both
    12NC order codes and EAN/UPC codes, and returns a direct product leaflet
    PDF URL, so no browser is needed. Returns None when the code cannot be
    resolved confidently, so the caller can fall back to the browser flow.
    """
    search_code = strip_product_prefix(code)
    compact_code = re.sub(r"[^a-z0-9]", "", search_code.lower())

    if not compact_code:
        return None

    try:
        response = requests.get(
            SIGNIFY_SEARCH_API_URL,
            params={"query": search_code},
            timeout=DEFAULT_TIMEOUT,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Accept": "application/json",
                "Origin": "https://www.signify.com",
                "Referer": "https://www.signify.com/",
            },
        )

        if response.status_code != 200:
            return None

        payload = response.json()
    except Exception:
        return None

    def field(result: dict, key: str):
        value = result.get(key)
        if isinstance(value, dict):
            return value.get("value")
        return value

    def compact(value) -> str:
        return re.sub(r"[^a-z0-9]", "", str(value or "").lower())

    matches = []
    for result in payload.get("results") or []:
        codes = field(result, "product_codes") or []
        if isinstance(codes, str):
            codes = [codes]

        known_codes = {compact(c) for c in codes}
        known_codes.add(compact(field(result, "sku")))

        if compact_code in known_codes:
            matches.append(result)

    for result in matches:
        leaflet_url = field(result, "leaflet") or ""
        if not leaflet_url:
            continue

        try:
            pdf_response = requests.get(
                leaflet_url,
                timeout=DEFAULT_TIMEOUT,
                headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                    "Accept": "application/pdf,*/*",
                    "Referer": "https://www.signify.com/",
                },
            )

            if pdf_response.status_code != 200:
                continue

            content = pdf_response.content or b""
            validate_pdf_content(content)

            return {
                "code": code,
                "brand": "Philips",
                "success": True,
                "url": leaflet_url,
                "error": "",
                "content": content,
            }
        except Exception:
            continue

    return None


def download_philips_datasheet(code: str) -> dict:
    """Download one Philips / Signify product leaflet.

    Tries the fast Signify product API first; falls back to the original
    Playwright browser flow when the API cannot resolve the code.
    """
    api_result = download_philips_datasheet_api(code)
    if api_result is not None:
        return api_result

    if not PLAYWRIGHT_READY or _sync_playwright is None:
        return {
            "code": code,
            "brand": "Philips",
            "success": False,
            "url": "",
            "error": PLAYWRIGHT_ERROR or "Playwright is not available.",
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

        if raw_url.startswith(("javascript:", "mailto:", "#")):
            return ""
        if raw_url.startswith("//"):
            return "https:" + raw_url
        if raw_url.startswith("http"):
            return raw_url

        return urljoin(base_url, raw_url)

    def compact(value: str) -> str:
        return re.sub(r"[^a-z0-9]", "", (value or "").lower())

    def is_exact_generic_prof_page(url: str) -> bool:
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

        parts = [part for part in path.split("/") if part]
        return len(parts) >= 4

    def add_unique(items: list[str], value: str) -> None:
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

        return result[:3]

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

        for match in re.findall(r"https?://[^\"'<>\s)]+", raw_text):
            add_unique(urls, match)

        for match in re.findall(
            r"/[^\"'<>\s)]*(?:product_leaflet|product-leaflet|leaflet|datasheet|data-sheet|api/assets|\.pdf)[^\"'<>\s)]*",
            raw_text,
            flags=re.IGNORECASE,
        ):
            add_unique(urls, urljoin(base_url, match))

        return urls

    def collect_download_candidates(page, base_url: str) -> list[str]:
        candidates = []

        try:
            links = page.locator("a[href]").all()
            for link in links[:500]:
                try:
                    href = link.get_attribute("href", timeout=1_000)
                    text = link.inner_text(timeout=1_000).strip()
                except Exception:
                    continue

                full = absolute_url(href, base_url)
                text_low = text.lower()

                if is_download_like(full) or "leaflet" in text_low or "datasheet" in text_low:
                    add_unique(candidates, full)
        except Exception:
            pass

        for selector in ["button", "[data-href]", "[data-url]", "[data-download-url]"]:
            try:
                nodes = page.locator(selector).all()
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

        try:
            html = page.content()
            for found in collect_urls_from_text(html, base_url):
                if is_download_like(found):
                    add_unique(candidates, found)
        except Exception:
            pass

        return candidates

    def pdf_mentions_code(content: bytes) -> bool:
        try:
            reader = PdfReader(io.BytesIO(content))
            text_parts = []

            for pdf_page in reader.pages[:3]:
                text_parts.append(pdf_page.extract_text() or "")

            pdf_text = compact("\n".join(text_parts))
            return compact_code in pdf_text
        except Exception:
            return False

    def try_pdf_urls(candidate_urls: list[str], referer: str, page_matches_code: bool, context):
        last = "No candidate download URLs found."

        for original_url in candidate_urls[:6]:
            attempts = [original_url]

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
                        timeout=15_000,
                    )

                    content = response.body() if response.ok else b""

                    if not response.ok:
                        last = f"HTTP {response.status} for {attempt_url}"
                        continue

                    if not content or not is_pdf_bytes(content):
                        last = f"Response is not a PDF for {attempt_url}"
                        continue

                    validate_pdf_content(content)

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
            args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"],
        )

        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            locale="en-US",
        )

        try:
            context.route(
                "**/*",
                lambda route: route.abort()
                if route.request.resource_type in {"image", "font", "media"}
                else route.continue_(),
            )
        except Exception:
            pass

        page = context.new_page()

        try:
            page.goto(home_url, wait_until="domcontentloaded", timeout=15_000)

            for accept_sel in [
                "#onetrust-accept-btn-handler",
                'button:has-text("Accept all")',
                'button:has-text("Accept All")',
                'button:has-text("Accept")',
            ]:
                try:
                    page.click(accept_sel, timeout=1_500)
                    break
                except Exception:
                    pass

            product_candidates = []

            for search_url in search_urls:
                try:
                    page.goto(search_url, wait_until="domcontentloaded", timeout=20_000)

                    try:
                        page.wait_for_load_state("load", timeout=5_000)
                    except Exception:
                        pass

                    page.wait_for_timeout(1_000)

                    search_downloads = collect_download_candidates(page, search_url)
                    result, last_error = try_pdf_urls(
                        search_downloads,
                        referer=search_url,
                        page_matches_code=True,
                        context=context,
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

            for product_url in product_candidates:
                try:
                    page.goto(product_url, wait_until="domcontentloaded", timeout=20_000)

                    try:
                        page.wait_for_load_state("load", timeout=5_000)
                    except Exception:
                        pass

                    page.wait_for_timeout(750)

                    try:
                        page_text = page.inner_text("body", timeout=2_000)
                    except Exception:
                        page_text = ""

                    page_matches_code = compact_code in compact(page_text)

                    for tab_sel in [
                        'button:has-text("Downloads")',
                        'a:has-text("Downloads")',
                        '[role="tab"]:has-text("Downloads")',
                        'button:has-text("Download")',
                        'a:has-text("Download")',
                    ]:
                        try:
                            page.click(tab_sel, timeout=1_000)
                            page.wait_for_timeout(500)
                            break
                        except Exception:
                            pass

                    download_candidates = collect_download_candidates(page, product_url)
                    result, last_error = try_pdf_urls(
                        download_candidates,
                        referer=product_url,
                        page_matches_code=page_matches_code,
                        context=context,
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
                    f"Tried {len(product_candidates)} fastest Philips product page candidate(s), "
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


def download_zambelis_datasheet(code: str) -> dict:
    """Download one Zambelis datasheet."""
    search_code = strip_product_prefix(code)
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
                headers={"User-Agent": "Mozilla/5.0", "Accept": "application/pdf,*/*"},
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


def parse_fumagalli_products(html: str) -> list[dict]:
    """Parse product entries from the Fumagalli downloads page.

    Each product is a block like:
    <div class="prodotto-download ...">
        <h3><a href="...">Product Name</a></h3>
        <div class="download-files ...">
            <a href="....pdf" class="download-file">... Technical Details ...</a>
            ...
        </div>
    </div>
    """
    products = []

    for block in re.split(r'class="prodotto-download', html)[1:]:
        name_match = re.search(r"<h3[^>]*>\s*<a[^>]*>([^<]+)</a>", block)
        if not name_match:
            continue

        name = re.sub(r"\s+", " ", name_match.group(1)).strip()
        technical_pdf = ""

        for href, inner in re.findall(
            r'<a\s+href="([^"]+)"\s+class="download-file"[^>]*>(.*?)</a>',
            block,
            flags=re.DOTALL,
        ):
            inner_text = re.sub(r"<[^>]+>", " ", inner)
            if "technical" in inner_text.lower() and href.lower().endswith(".pdf"):
                technical_pdf = href
                break

        products.append({"name": name, "technical_pdf": technical_pdf})

    return products


def search_fumagalli_products(query: str) -> tuple[list[dict], str]:
    """Search the Fumagalli downloads page and return parsed product entries."""
    try:
        response = requests.post(
            FUMAGALLI_DOWNLOADS_URL,
            data={"search": query},
            timeout=DEFAULT_TIMEOUT,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                "Accept": "text/html,*/*",
            },
        )

        if response.status_code != 200:
            return [], f"HTTP {response.status_code} while searching Fumagalli downloads"

        return parse_fumagalli_products(response.text), ""

    except Exception as e:
        return [], str(e)


_FUMAGALLI_CATALOG_CACHE: dict = {"timestamp": 0.0, "products": []}
_FUMAGALLI_CATALOG_LOCK = threading.Lock()


def fetch_fumagalli_catalog() -> list[dict]:
    """Fetch the full Fumagalli product list from all downloads pages (cached)."""
    with _FUMAGALLI_CATALOG_LOCK:
        age = time.time() - _FUMAGALLI_CATALOG_CACHE["timestamp"]
        if _FUMAGALLI_CATALOG_CACHE["products"] and age < FUMAGALLI_CATALOG_TTL:
            return _FUMAGALLI_CATALOG_CACHE["products"]

        products = []
        seen = set()

        for page in range(1, FUMAGALLI_MAX_PAGES + 1):
            url = FUMAGALLI_DOWNLOADS_URL if page == 1 else f"{FUMAGALLI_DOWNLOADS_URL}page/{page}/"

            try:
                response = requests.get(
                    url,
                    timeout=DEFAULT_TIMEOUT,
                    headers={
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                        "Accept": "text/html,*/*",
                    },
                )
            except Exception:
                break

            if response.status_code != 200:
                break

            page_products = parse_fumagalli_products(response.text)
            if not page_products:
                break

            for product in page_products:
                key = product["name"].casefold()
                if key not in seen:
                    seen.add(key)
                    products.append(product)

        if products:
            _FUMAGALLI_CATALOG_CACHE["products"] = products
            _FUMAGALLI_CATALOG_CACHE["timestamp"] = time.time()

        return products


def fumagalli_tokens(text: str) -> list[str]:
    """Tokenize a product name or free-form description for matching.

    Attribute words that never appear in catalogue names (colors, wattage,
    color temperature, IP rating) are dropped, and dimensions in cm are
    converted to the mm numbers used by catalogue names (D 6 CM -> 60).
    """
    text = unescape(str(text)).lower()
    text = text.replace("/", " ").replace("-", " ").replace("_", " ")

    extra_tokens = []
    for match in re.finditer(r"\b(\d+(?:[.,]\d+)?)\s*cm\b", text):
        value = float(match.group(1).replace(",", "."))
        extra_tokens.append(str(int(round(value * 10))))

    text = re.sub(r"\b\d+(?:[.,]\d+)?\s*w\b", " ", text)  # wattage: 8.5W
    text = re.sub(r"\b\d{3,4}\s*k\b", " ", text)  # color temperature: 3000K
    text = re.sub(r"\bip\s*\d+\b", " ", text)  # IP rating
    text = re.sub(r"\b\d+(?:[.,]\d+)?\s*cm\b", " ", text)  # converted dimensions

    tokens = re.findall(r"[a-z]+\d+[a-z0-9]*|[a-z]+|\d+(?:\.\d+)?", text)
    return [t for t in tokens if t not in FUMAGALLI_NOISE_TOKENS] + extra_tokens


def fumagalli_name_tokens(name: str) -> list[str]:
    """Tokens of a catalogue product name ('Range' is decorative, not matching)."""
    return [t for t in fumagalli_tokens(name) if t != "range"]


def resolve_fumagalli_product(description: str, products: list[dict]) -> tuple[dict | None, str]:
    """Match a product name or free-form description to one catalogue product.

    Returns (product, note). product is None when there is no confident match,
    and the note explains why (ambiguous, unknown family, ...).
    """
    desc_norm = re.sub(r"\s+", " ", unescape(str(description))).strip().casefold()
    usable = [p for p in products if p["technical_pdf"]]

    for product in usable:
        if unescape(product["name"]).strip().casefold() == desc_norm:
            return product, "exact name match"

    desc_tokens = set(fumagalli_tokens(description))
    scored = []

    for product in usable:
        name_tokens = set(fumagalli_name_tokens(product["name"]))
        if not name_tokens:
            continue

        family_token = fumagalli_name_tokens(product["name"])[0]
        if family_token not in desc_tokens:
            continue

        matched = len(name_tokens & desc_tokens)
        scored.append((matched, len(name_tokens), product))

    if not scored:
        return None, "no catalogue product matches this description"

    # Products whose name words are ALL present in the description.
    full = [(total, product) for matched, total, product in scored if matched == total]
    if full:
        full.sort(key=lambda item: item[0], reverse=True)
        best_total = full[0][0]
        best = [product for total, product in full if total == best_total]

        if len(best) == 1:
            return best[0], "matched all name words"

        return None, "ambiguous between: " + ", ".join(p["name"] for p in best)

    # Otherwise pick the product with the most matched words, if unique.
    scored.sort(key=lambda item: item[0], reverse=True)
    best_matched = scored[0][0]
    best = [product for matched, total, product in scored if matched == best_matched]

    if best_matched >= 2 and len(best) == 1:
        return best[0], "closest catalogue match"

    candidates = ", ".join(dict.fromkeys(product["name"] for _, _, product in scored))
    return None, f"description is not specific enough; possible products: {candidates}"


def fetch_fumagalli_pdf(display_name: str, product: dict) -> dict:
    """Download and validate the Technical Details PDF of a matched product."""
    pdf_url = product["technical_pdf"]

    try:
        response = requests.get(
            pdf_url,
            timeout=DEFAULT_TIMEOUT,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                "Accept": "application/pdf,*/*",
                "Referer": FUMAGALLI_DOWNLOADS_URL,
            },
        )

        if response.status_code != 200:
            raise ValueError(f"HTTP {response.status_code}")

        content = response.content or b""
        validate_pdf_content(content)

        return {
            "code": display_name,
            "brand": "Fumagalli",
            "success": True,
            "url": pdf_url,
            "error": "",
            "content": content,
        }

    except Exception as e:
        return {
            "code": display_name,
            "brand": "Fumagalli",
            "success": False,
            "url": pdf_url,
            "error": f"Matched product '{product['name']}' but the PDF download failed: {e}",
            "content": None,
        }


def download_fumagalli_datasheet(name: str) -> dict:
    """Download the Technical Details PDF for one FUMAGALLI name or description."""
    query = normalize_product_name(name)

    if not query:
        return {
            "code": name,
            "brand": "Fumagalli",
            "success": False,
            "url": "",
            "error": (
                "Empty FUMAGALLI product description. FUM items need the product "
                "name or description in the Description column."
            ),
            "content": None,
        }

    catalog = fetch_fumagalli_catalog()

    if catalog:
        product, note = resolve_fumagalli_product(query, catalog)

        if product is None:
            return {
                "code": name,
                "brand": "Fumagalli",
                "success": False,
                "url": FUMAGALLI_DOWNLOADS_URL,
                "error": f"Could not match '{query}' to a FUMAGALLI catalogue product: {note}",
                "content": None,
            }

        return fetch_fumagalli_pdf(name, product)

    # Catalogue unavailable - fall back to the on-site search.
    key = query.casefold()

    products, last_error = search_fumagalli_products(query)

    # If nothing came back for the full name, retry with the first word only.
    if not products and " " in query:
        products, last_error = search_fumagalli_products(query.split(" ")[0])

    match = None
    exact = [p for p in products if p["name"].casefold() == key]
    prefix = [p for p in products if p["name"].casefold().startswith(key)]
    contains = [p for p in products if key in p["name"].casefold()]

    for group in (exact, prefix, contains):
        with_pdf = [p for p in group if p["technical_pdf"]]
        if with_pdf:
            match = with_pdf[0]
            break

    if match is None:
        if products:
            found_names = ", ".join(p["name"] for p in products[:10])
            error = (
                f"No matching Fumagalli product with a Technical Details PDF "
                f"was found for '{query}'. Products returned by the search: {found_names}"
            )
        else:
            error = (
                f"No Fumagalli products found for '{query}'. "
                f"Last error: {last_error or 'empty search result'}"
            )

        return {
            "code": name,
            "brand": "Fumagalli",
            "success": False,
            "url": FUMAGALLI_DOWNLOADS_URL,
            "error": error,
            "content": None,
        }

    return fetch_fumagalli_pdf(name, match)


def download_datasheet(code: str) -> dict:
    """Route product code to the correct vendor downloader."""
    product_type = get_product_type(code)

    if product_type == "philips":
        return download_philips_datasheet(code)
    if product_type == "zambelis":
        return download_zambelis_datasheet(code)

    return {
        "code": code,
        "brand": "Unknown",
        "success": False,
        "url": "",
        "error": (
            "Unknown code prefix. Codes must start with PHL or ZMB. "
            "FUM items are searched by their Description (FUMAGALLI box or Excel Description column)."
        ),
        "content": None,
    }


def load_cover_template_bytes() -> bytes | None:
    """Load the cover page template PDF shipped with the app."""
    try:
        with open(COVER_TEMPLATE_PATH, "rb") as f:
            return f.read()
    except Exception:
        return None


def build_type_overlay(type_text: str, page_width: float, page_height: float) -> bytes:
    """Draw the item type in the blank space under the logo of the cover page."""
    buffer = io.BytesIO()
    overlay = pdf_canvas.Canvas(buffer, pagesize=(page_width, page_height))
    overlay.setFillColor(HexColor(COVER_TEXT_COLOR))

    def wrap_lines(font_size: int) -> list[str]:
        lines = []
        current = ""
        for word in type_text.split():
            candidate = f"{current} {word}".strip()
            if overlay.stringWidth(candidate, COVER_TEXT_FONT, font_size) <= COVER_TEXT_MAX_WIDTH:
                current = candidate
            else:
                if current:
                    lines.append(current)
                current = word
        if current:
            lines.append(current)
        return lines

    font_size = COVER_TEXT_FONT_SIZE
    lines = wrap_lines(font_size)

    while font_size > COVER_TEXT_MIN_FONT_SIZE and (
        len(lines) > 3
        or any(
            overlay.stringWidth(line, COVER_TEXT_FONT, font_size) > COVER_TEXT_MAX_WIDTH
            for line in lines
        )
    ):
        font_size -= 2
        lines = wrap_lines(font_size)

    overlay.setFont(COVER_TEXT_FONT, font_size)
    y = page_height - COVER_TEXT_TOP_OFFSET

    for line in lines:
        overlay.drawString(COVER_TEXT_X, y, line)
        y -= font_size * 1.3

    overlay.save()
    return buffer.getvalue()


def build_cover_page(template_bytes: bytes, type_text: str):
    """Return the cover template page, with the item type written on it."""
    template_reader = PdfReader(io.BytesIO(template_bytes))
    page = template_reader.pages[0]

    if type_text:
        overlay_bytes = build_type_overlay(
            type_text,
            float(page.mediabox.width),
            float(page.mediabox.height),
        )
        overlay_reader = PdfReader(io.BytesIO(overlay_bytes))
        page.merge_page(overlay_reader.pages[0])

    return page


def toc_pages_needed(entry_count: int) -> int:
    """Number of pages the table of contents itself will occupy."""
    if entry_count <= TOC_ENTRIES_FIRST_PAGE:
        return 1

    remaining = entry_count - TOC_ENTRIES_FIRST_PAGE
    extra_pages = -(-remaining // TOC_ENTRIES_LATER_PAGES)  # ceiling division
    return 1 + extra_pages


def build_toc_pdf(
    entries: list[dict],
    page_width: float,
    page_height: float,
) -> tuple[bytes, list[dict]]:
    """Draw the table of contents pages.

    entries: [{"title": str, "target_page": int}] where target_page is the
    0-based page index of the item's cover page in the final document.

    Returns (pdf_bytes, link_boxes). link_boxes hold the clickable rectangle
    of every entry: [{"page": toc_page_index, "rect": (x0,y0,x1,y1), "target": int}].
    """
    buffer = io.BytesIO()
    toc = pdf_canvas.Canvas(buffer, pagesize=(page_width, page_height))

    accent = HexColor(TOC_ACCENT_COLOR)
    title_color = HexColor(TOC_TITLE_COLOR)
    dots_color = HexColor(TOC_DOTS_COLOR)

    number_x = TOC_MARGIN_X
    title_x = TOC_MARGIN_X + 34
    page_num_right = page_width - TOC_MARGIN_X
    max_title_width = page_num_right - title_x - 60

    def truncate(text: str, font: str, size: float) -> str:
        if toc.stringWidth(text, font, size) <= max_title_width:
            return text
        while text and toc.stringWidth(text + "...", font, size) > max_title_width:
            text = text[:-1]
        return text.rstrip() + "..."

    def draw_first_page_header() -> float:
        """Draw logo + title, return the y where entries start."""
        y_top = page_height - 52

        try:
            from reportlab.lib.utils import ImageReader

            logo = ImageReader(TOC_LOGO_PATH)
            logo_w, logo_h = logo.getSize()
            draw_h = 26
            draw_w = logo_w * draw_h / logo_h
            toc.drawImage(
                logo,
                TOC_MARGIN_X,
                y_top - draw_h,
                width=draw_w,
                height=draw_h,
                mask="auto",
            )
        except Exception:
            pass

        title_y = y_top - 64
        toc.setFillColor(title_color)
        toc.setFont("Helvetica-Bold", 27)
        toc.drawString(TOC_MARGIN_X, title_y, "Table of Contents")

        toc.setFillColor(accent)
        toc.rect(TOC_MARGIN_X, title_y - 14, 64, 4, stroke=0, fill=1)

        return title_y - 52

    def draw_later_page_header() -> float:
        toc.setFillColor(dots_color)
        toc.setFont("Helvetica", 11)
        toc.drawString(TOC_MARGIN_X, page_height - 56, "Table of Contents (continued)")
        toc.setFillColor(accent)
        toc.rect(TOC_MARGIN_X, page_height - 64, 42, 2.6, stroke=0, fill=1)
        return page_height - 100

    link_boxes = []
    toc_page_index = 0
    y = draw_first_page_header()
    capacity = TOC_ENTRIES_FIRST_PAGE
    drawn_on_page = 0

    for position, entry in enumerate(entries, start=1):
        if drawn_on_page >= capacity:
            toc.showPage()
            toc_page_index += 1
            y = draw_later_page_header()
            capacity = TOC_ENTRIES_LATER_PAGES
            drawn_on_page = 0

        title = truncate(entry["title"], "Helvetica-Bold", 12.5)
        page_label = str(entry["target_page"] + 1)

        toc.setFillColor(accent)
        toc.setFont("Helvetica-Bold", 10.5)
        toc.drawString(number_x, y, f"{position:02d}")

        toc.setFillColor(title_color)
        toc.setFont("Helvetica-Bold", 12.5)
        toc.drawString(title_x, y, title)

        toc.setFont("Helvetica-Bold", 11.5)
        toc.setFillColor(accent)
        toc.drawRightString(page_num_right, y, page_label)

        title_end = title_x + toc.stringWidth(title, "Helvetica-Bold", 12.5) + 8
        num_start = page_num_right - toc.stringWidth(page_label, "Helvetica-Bold", 11.5) - 8
        if num_start > title_end + 12:
            toc.setFillColor(dots_color)
            toc.setFont("Helvetica", 10)
            dot = "."
            dot_width = toc.stringWidth(dot, "Helvetica", 10) + 3.2
            x = title_end
            while x < num_start:
                toc.drawString(x, y + 1, dot)
                x += dot_width

        link_boxes.append(
            {
                "page": toc_page_index,
                "rect": (TOC_MARGIN_X - 6, y - 8, page_num_right + 6, y + 14),
                "target": entry["target_page"],
            }
        )

        y -= TOC_ENTRY_SPACING
        drawn_on_page += 1

    toc.save()
    return buffer.getvalue(), link_boxes


def merge_pdfs(items: list[dict], template_bytes: bytes | None) -> bytes:
    """Merge every item's datasheet into one PDF.

    The document starts with a clickable table of contents listing each
    item's Type. Every successful item then contributes a cover page (the
    template with the item's Type written on it) followed by its datasheet.
    Items are kept in order and duplicates are NOT removed: every item gets
    its own cover and datasheet even when two items share the same file.
    """
    prepared = []
    for item in items:
        result = item["result"]
        if not result["success"]:
            continue
        reader = PdfReader(io.BytesIO(result["content"]))
        prepared.append((item, reader))

    if not prepared:
        return b""

    cover_pages = 1 if template_bytes else 0
    toc_page_count = toc_pages_needed(len(prepared))

    if template_bytes:
        template_page = PdfReader(io.BytesIO(template_bytes)).pages[0]
        page_width = float(template_page.mediabox.width)
        page_height = float(template_page.mediabox.height)
    else:
        first_page = prepared[0][1].pages[0]
        page_width = float(first_page.mediabox.width)
        page_height = float(first_page.mediabox.height)

    entries = []
    cursor = toc_page_count
    for item, reader in prepared:
        title = item.get("type", "").strip() or item.get("display", "") or "Item"
        entries.append({"title": title, "target_page": cursor})
        cursor += cover_pages + len(reader.pages)

    toc_bytes, link_boxes = build_toc_pdf(entries, page_width, page_height)

    writer = PdfWriter()

    for page in PdfReader(io.BytesIO(toc_bytes)).pages:
        writer.add_page(page)

    for item, reader in prepared:
        if template_bytes:
            writer.add_page(build_cover_page(template_bytes, item.get("type", "")))
        for page in reader.pages:
            writer.add_page(page)

    for box in link_boxes:
        writer.add_annotation(
            page_number=box["page"],
            annotation=Link(
                rect=box["rect"],
                target_page_index=box["target"],
                fit=Fit(fit_type="/Fit"),
            ),
        )

    for entry in entries:
        writer.add_outline_item(entry["title"], entry["target_page"])

    output = io.BytesIO()
    writer.write(output)
    return output.getvalue()

# ============================================================
# Philips + Zambelis Professional CSS
# ============================================================

st.markdown(
    """
<style>
    :root {
        --philips-blue: #035ED8;
        --philips-bright-blue: #0B5ED7;
        --philips-deep-blue: #003B79;
        --philips-light-blue: #EAF3FF;
        --zambelis-black: #111111;
        --zambelis-charcoal: #252525;
        --zambelis-warm-gray: #6F6A62;
        --zambelis-gold: #C8A45D;
        --zambelis-soft-gold: #F4E8CC;
        --zambelis-cream: #FAF7F0;
        --fumagalli-green: #1E5B3A;
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
        background: linear-gradient(180deg, transparent, rgba(37, 37, 37, 0.35), transparent);
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

    .fumagalli-logo {
        background: #ffffff;
        color: var(--fumagalli-green);
        border: 2px solid var(--fumagalli-green);
        border-radius: 4px;
        padding: 9px 15px;
        font-weight: 800;
        font-size: 18px;
        letter-spacing: 2px;
        line-height: 1;
        box-shadow: 0 8px 20px rgba(30, 91, 58, 0.14);
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
        background: linear-gradient(135deg, var(--philips-deep-blue) 0%, var(--philips-blue) 44%, var(--zambelis-charcoal) 72%, var(--zambelis-black) 100%);
        color: white;
        margin-bottom: 28px;
        box-shadow: 0 24px 50px rgba(0, 59, 121, 0.22), 0 12px 30px rgba(17, 17, 17, 0.18);
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
        background: linear-gradient(135deg, var(--philips-blue) 0%, var(--philips-deep-blue) 55%, var(--zambelis-black) 100%);
        color: white;
        border: 0;
        border-radius: 999px;
        padding: 13px 28px;
        font-weight: 850;
        box-shadow: 0 14px 28px rgba(3, 94, 216, 0.28), 0 8px 18px rgba(17, 17, 17, 0.15);
        transition: all 0.2s ease;
    }

    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 18px 34px rgba(3, 94, 216, 0.32), 0 10px 22px rgba(17, 17, 17, 0.18);
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

# ============================================================
# Header
# ============================================================

st.markdown(
    """
<div class="brand-topbar">
    <div class="brand-logos">
        <div class="philips-logo">PHILIPS</div>
        <div class="brand-divider"></div>
        <div class="zambelis-logo">ZAMBELIS</div>
        <div class="brand-divider"></div>
        <div class="fumagalli-logo">FUMAGALLI</div>
    </div>
    <div class="brand-badge">Philips, Zambelis &amp; FUMAGALLI Datasheet Automation</div>
</div>

<div class="hero">
    <div class="hero-content">
        <div class="hero-kicker">Product documentation tool</div>
        <h1>Datasheet Pack Builder</h1>
        <p>
            Paste product codes, upload Excel lists, download official datasheets,
            and merge everything into one PDF.
        </p>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

# ============================================================
# Input Section
# ============================================================

left_col, middle_col, right_col = st.columns([1, 1, 1], gap="large")

with left_col:
    st.markdown(
        """
<div class="tool-card">
    <div class="section-title">Paste product codes</div>
    <div class="section-subtitle">Add one or multiple product codes. Use PHL for Philips and ZMB for Zambelis.</div>
""",
        unsafe_allow_html=True,
    )

    manual_codes_text = st.text_area(
        "Product codes",
        placeholder="Example:\nPHL046677568283\nZMB12345",
        height=90,
        label_visibility="collapsed",
    )

    st.markdown("</div>", unsafe_allow_html=True)

with middle_col:
    st.markdown(
        """
<div class="tool-card">
    <div class="section-title">FUMAGALLI product names</div>
    <div class="section-subtitle">Paste FUMAGALLI product names or full descriptions, one per line.</div>
""",
        unsafe_allow_html=True,
    )

    fumagalli_names_text = st.text_area(
        "FUMAGALLI product names",
        placeholder="Example:\nCarlo\nMod. Abram 190 Grey 8.5W 3000K\nMod Livia D 6 CM Grey 1.7W 4000K",
        height=90,
        label_visibility="collapsed",
    )

    st.markdown("</div>", unsafe_allow_html=True)

with right_col:
    st.markdown(
        """
<div class="tool-card">
    <div class="section-title">Upload Excel file</div>
    <div class="section-subtitle">Columns: Type, Code, and Description. FUMAGALLI codes use the Description.</div>
""",
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader(
        "Upload Excel file",
        type=["xlsx", "xls"],
        label_visibility="collapsed",
    )

    st.markdown("</div>", unsafe_allow_html=True)

# Full-width Excel preview below the input row
excel_items = []

if uploaded_file:
    try:
        uploaded_file.seek(0)
        df_preview = pd.read_excel(uploaded_file)

        st.caption("Excel preview")
        st.dataframe(df_preview.head(10), use_container_width=True)

        excel_items, excel_error = extract_items_from_excel(uploaded_file)
        if excel_error:
            st.error(excel_error)
    except Exception as e:
        st.error(f"Could not read Excel file: {e}")

# ============================================================
# Options Section
# ============================================================

st.markdown(
    """
<div class="tool-card">
    <div class="section-title">Export settings</div>
    <div class="section-subtitle">Choose the final PDF filename and how failed downloads should be handled.</div>
""",
    unsafe_allow_html=True,
)

settings_col_1, settings_col_2 = st.columns([2, 1], gap="large")

with settings_col_1:
    output_filename = st.text_input(
        "Output PDF filename",
        value="datasheets pack.pdf",
    )

with settings_col_2:
    skip_failed = st.checkbox(
        "Skip failed codes and continue",
        value=True,
    )

st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# Code Summary
# ============================================================

manual_codes = dedupe_preserve_order(extract_codes_from_text(manual_codes_text))
fumagalli_names = dedupe_names_preserve_order(extract_names_from_text(fumagalli_names_text))

all_items = drop_untyped_duplicates(
    [{"kind": "code", "value": code, "type": "", "display": code} for code in manual_codes]
    + [{"kind": "fumagalli", "value": name, "type": "", "display": name} for name in fumagalli_names]
    + excel_items
)

philips_count = len(
    [i for i in all_items if i["kind"] == "code" and get_product_type(i["value"]) == "philips"]
)
zambelis_count = len(
    [i for i in all_items if i["kind"] == "code" and get_product_type(i["value"]) == "zambelis"]
)
fumagalli_count = len([i for i in all_items if i["kind"] == "fumagalli"])
unknown_count = len(all_items) - philips_count - zambelis_count - fumagalli_count

st.markdown("### Summary before download")

metric_1, metric_2, metric_3, metric_4 = st.columns(4)

with metric_1:
    st.metric("Manual codes", len(manual_codes))

with metric_2:
    st.metric("FUMAGALLI names", len(fumagalli_names))

with metric_3:
    st.metric("Excel items", len(excel_items))

with metric_4:
    st.metric("Total items", len(all_items))

brand_metric_1, brand_metric_2, brand_metric_3, brand_metric_4 = st.columns(4)

with brand_metric_1:
    st.metric("Philips", philips_count)

with brand_metric_2:
    st.metric("Zambelis", zambelis_count)

with brand_metric_3:
    st.metric("FUMAGALLI", fumagalli_count)

with brand_metric_4:
    st.metric("Unknown prefix", unknown_count)

if all_items:
    with st.expander("View detected items"):
        fumagalli_items = [i for i in all_items if i["kind"] == "fumagalli"]

        catalog_preview = []
        if fumagalli_items:
            try:
                catalog_preview = fetch_fumagalli_catalog()
            except Exception:
                catalog_preview = []

        overview_rows = []
        for item in all_items:
            if item["kind"] == "fumagalli":
                brand = "FUMAGALLI"
                matched = ""
                if catalog_preview:
                    matched_product, match_note = resolve_fumagalli_product(
                        normalize_product_name(item["value"]),
                        catalog_preview,
                    )
                    matched = matched_product["name"] if matched_product else f"No match ({match_note})"
            else:
                product_type = get_product_type(item["value"])
                brand = product_type.capitalize() if product_type != "unknown" else "Unknown"
                matched = ""

            overview_rows.append(
                {
                    "Item": item["display"],
                    "Brand": brand,
                    "Type (cover page)": item.get("type", ""),
                    "Matched FUMAGALLI product": matched,
                }
            )

        st.dataframe(pd.DataFrame(overview_rows), use_container_width=True)

# ============================================================
# Download + Merge Action
# ============================================================

download_button = st.button(
    "Download and merge datasheets",
    type="primary",
    disabled=len(all_items) == 0,
)

if download_button:
    start_time = time.time()

    st.info("Downloading datasheets...")

    # Download each unique (kind, value) once, then map results back to
    # every item, so repeated codes/descriptions do not download twice but
    # still each get their own cover page and datasheet in the merged PDF.
    unique_jobs = list(dict.fromkeys((item["kind"], item["value"]) for item in all_items))

    results_by_job = {}
    progress_bar = st.progress(0)
    status_text = st.empty()

    def run_download_job(kind: str, value: str) -> dict:
        if kind == "fumagalli":
            return download_fumagalli_datasheet(value)
        return download_datasheet(value)

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_map = {
            executor.submit(run_download_job, kind, value): (kind, value)
            for kind, value in unique_jobs
        }
        completed = 0

        for future in as_completed(future_map):
            kind, value = future_map[future]

            try:
                result = future.result()
            except Exception as e:
                if kind == "fumagalli":
                    brand = "Fumagalli"
                else:
                    product_type = get_product_type(value)
                    if product_type == "philips":
                        brand = "Philips"
                    elif product_type == "zambelis":
                        brand = "Zambelis"
                    else:
                        brand = "Unknown"

                result = {
                    "code": value,
                    "brand": brand,
                    "success": False,
                    "url": "",
                    "error": str(e),
                    "content": None,
                }

            results_by_job[(kind, value)] = result
            completed += 1
            progress_bar.progress(completed / len(unique_jobs))
            status_text.write(f"Processed {completed} / {len(unique_jobs)}")

    for item in all_items:
        item["result"] = results_by_job[(item["kind"], item["value"])]

    successful = [item for item in all_items if item["result"]["success"]]
    failed = [item for item in all_items if not item["result"]["success"]]

    st.divider()

    result_col_1, result_col_2, result_col_3 = st.columns(3)

    with result_col_1:
        st.metric("Submitted", len(all_items))

    with result_col_2:
        st.metric("Downloaded", len(successful))

    with result_col_3:
        st.metric("Failed", len(failed))

    if failed:
        st.warning("Some items failed.")

        failed_table = pd.DataFrame(
            [
                {
                    "Item": item["display"],
                    "Brand": item["result"].get("brand", ""),
                    "Type": item.get("type", ""),
                    "URL": item["result"]["url"],
                    "Error": item["result"]["error"],
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
        cover_template_bytes = load_cover_template_bytes()
        if cover_template_bytes is None:
            st.warning(
                "Cover page template (item_type_template.pdf) was not found. "
                "Datasheets were merged without cover pages."
            )

        merged_pdf = merge_pdfs(successful, cover_template_bytes)

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
