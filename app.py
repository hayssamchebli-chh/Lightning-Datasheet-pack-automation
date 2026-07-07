import hashlib
import io
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


def extract_codes_from_excel(uploaded_file, selected_column: str) -> list[str]:
    """Extract product codes from the selected Excel column."""
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file)

    if selected_column not in df.columns:
        return []

    codes = []
    for value in df[selected_column].dropna():
        code = normalize_code(value)
        if code:
            codes.append(code)

    return codes


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
        "error": "Unknown code prefix. Code must start with PHL or ZMB.",
        "content": None,
    }


def merge_pdfs(downloads: list[dict]) -> tuple[bytes, int]:
    """Merge all successful PDFs into one PDF, skipping duplicate files.

    Different inputs can lead to the same datasheet (for example two color
    variants of the same FUMAGALLI product). Identical PDF files are merged
    only once. Returns (merged_pdf_bytes, number_of_duplicates_skipped).
    """
    writer = PdfWriter()
    seen_hashes = set()
    duplicates_skipped = 0

    for item in downloads:
        if not item["success"]:
            continue

        digest = hashlib.md5(item["content"]).hexdigest()
        if digest in seen_hashes:
            duplicates_skipped += 1
            continue

        seen_hashes.add(digest)

        reader = PdfReader(io.BytesIO(item["content"]))
        for page in reader.pages:
            writer.add_page(page)

    output = io.BytesIO()
    writer.write(output)
    return output.getvalue(), duplicates_skipped

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
    <div class="section-subtitle">Select the column that contains the product codes.</div>
""",
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader(
        "Upload Excel file",
        type=["xlsx", "xls"],
        label_visibility="collapsed",
    )

    excel_codes = []

    if uploaded_file:
        try:
            uploaded_file.seek(0)
            df_preview = pd.read_excel(uploaded_file)

            st.caption("Excel preview")
            st.dataframe(df_preview.head(), use_container_width=True)

            selected_column = st.selectbox(
                "Choose product code column",
                options=list(df_preview.columns),
            )

            excel_codes = extract_codes_from_excel(uploaded_file, selected_column)
        except Exception as e:
            st.error(f"Could not read Excel file: {e}")

    st.markdown("</div>", unsafe_allow_html=True)

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

manual_codes = extract_codes_from_text(manual_codes_text)
all_codes = dedupe_preserve_order(manual_codes + excel_codes)
fumagalli_names = dedupe_names_preserve_order(extract_names_from_text(fumagalli_names_text))

philips_count = len([code for code in all_codes if get_product_type(code) == "philips"])
zambelis_count = len([code for code in all_codes if get_product_type(code) == "zambelis"])
unknown_count = len([code for code in all_codes if get_product_type(code) == "unknown"])

st.markdown("### Summary before download")

metric_1, metric_2, metric_3, metric_4 = st.columns(4)

with metric_1:
    st.metric("Manual codes", len(manual_codes))

with metric_2:
    st.metric("Excel codes", len(excel_codes))

with metric_3:
    st.metric("FUMAGALLI names", len(fumagalli_names))

with metric_4:
    st.metric("Unique total", len(all_codes) + len(fumagalli_names))

brand_metric_1, brand_metric_2, brand_metric_3, brand_metric_4 = st.columns(4)

with brand_metric_1:
    st.metric("Philips codes", philips_count)

with brand_metric_2:
    st.metric("Zambelis codes", zambelis_count)

with brand_metric_3:
    st.metric("FUMAGALLI names", len(fumagalli_names))

with brand_metric_4:
    st.metric("Unknown prefix", unknown_count)

if all_codes or fumagalli_names:
    with st.expander("View detected items"):
        if all_codes:
            st.write("Product codes:", all_codes)

        if fumagalli_names:
            st.write("FUMAGALLI items:")

            try:
                catalog_preview = fetch_fumagalli_catalog()
            except Exception:
                catalog_preview = []

            if catalog_preview:
                preview_rows = []
                for fumagalli_name in fumagalli_names:
                    matched_product, match_note = resolve_fumagalli_product(
                        normalize_product_name(fumagalli_name),
                        catalog_preview,
                    )
                    preview_rows.append(
                        {
                            "Input": fumagalli_name,
                            "Matched product": matched_product["name"] if matched_product else "No match",
                            "Note": "" if matched_product else match_note,
                        }
                    )

                st.dataframe(pd.DataFrame(preview_rows), use_container_width=True)
            else:
                st.write(fumagalli_names)

# ============================================================
# Download + Merge Action
# ============================================================

download_jobs = [("code", code) for code in all_codes] + [
    ("fumagalli", name) for name in fumagalli_names
]

download_button = st.button(
    "Download and merge datasheets",
    type="primary",
    disabled=len(download_jobs) == 0,
)

if download_button:
    start_time = time.time()

    st.info("Downloading datasheets...")

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
            for kind, value in download_jobs
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
            progress_bar.progress(completed / len(download_jobs))
            status_text.write(f"Processed {completed} / {len(download_jobs)}")

    results = [results_by_job[job] for job in download_jobs if job in results_by_job]
    successful = [item for item in results if item["success"]]
    failed = [item for item in results if not item["success"]]

    st.divider()

    result_col_1, result_col_2, result_col_3 = st.columns(3)

    with result_col_1:
        st.metric("Submitted", len(download_jobs))

    with result_col_2:
        st.metric("Downloaded", len(successful))

    with result_col_3:
        st.metric("Failed", len(failed))

    if failed:
        st.warning("Some codes failed.")

        failed_table = pd.DataFrame(
            [
                {
                    "Code / Name": item["code"],
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
        merged_pdf, duplicates_skipped = merge_pdfs(successful)

        if not output_filename.lower().endswith(".pdf"):
            output_filename += ".pdf"

        elapsed = round(time.time() - start_time, 2)
        st.success(f"PDF pack created successfully in {elapsed} seconds.")

        if duplicates_skipped:
            st.info(
                f"{duplicates_skipped} item(s) shared the same datasheet as another item, "
                f"so each datasheet was included only once."
            )

        st.download_button(
            label="Download merged PDF",
            data=merged_pdf,
            file_name=output_filename,
            mime="application/pdf",
        )

    except Exception as e:
        st.error(f"Failed to merge PDFs: {e}")
