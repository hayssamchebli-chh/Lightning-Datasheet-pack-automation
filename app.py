import io
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests
import streamlit as st
from pypdf import PdfReader, PdfWriter


# ============================================================
# Configuration
# ============================================================

BASE_URL = "https://www.assets.signify.com/is/content/Signify/US.en_US.{code}"

DEFAULT_TIMEOUT = (20, 40)  # connect timeout, read timeout
MAX_WORKERS = 8


# ============================================================
# Helpers
# ============================================================

def normalize_code(value: str) -> str:
    """
    Clean Philips / Signify product code.

    Supports values like:
    HE-046677590543 -> 046677590543
    HE-046677591670 -> 046677591670
    046677568283 -> 046677568283
    """
    if value is None:
        return ""

    value = str(value).strip()

    # If the value contains a dash, keep only the part after the first dash
    if "-" in value:
        value = value.split("-", 1)[1]

    # Remove spaces, commas, semicolons, tabs
    value = re.sub(r"[\s,;]+", "", value)

    # Keep only letters and numbers
    value = re.sub(r"[^A-Za-z0-9]", "", value)

    return value


def extract_codes_from_text(text: str) -> list[str]:
    """
    Extract product codes from manual text input.
    Supports newline, space, comma, semicolon, and tab separators.
    """
    if not text:
        return []

    raw_items = re.split(r"[\n,;\t ]+", text)
    codes = []

    for item in raw_items:
        code = normalize_code(item)
        if code:
            codes.append(code)

    return codes


def extract_codes_from_excel(uploaded_file, selected_column: str) -> list[str]:
    """
    Extract product codes from the selected Excel column.
    """
    df = pd.read_excel(uploaded_file)

    if selected_column not in df.columns:
        return []

    codes = []

    for value in df[selected_column].dropna():
        code = normalize_code(value)
        if code:
            codes.append(code)

    return codes


def dedupe_preserve_order(items: list[str]) -> list[str]:
    """
    Remove duplicates while preserving the first appearance order.
    """
    seen = set()
    result = []

    for item in items:
        if item not in seen:
            seen.add(item)
            result.append(item)

    return result


def is_pdf_bytes(content: bytes) -> bool:
    """
    Check if the downloaded file starts with PDF signature.
    """
    return content[:5] == b"%PDF-"


def download_datasheet(code: str) -> dict:
    """
    Download one Philips / Signify datasheet.
    The URL does not end with .pdf, so we verify using Content-Type and %PDF signature.
    """
    url = BASE_URL.format(code=code)

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
            return {
                "code": code,
                "success": False,
                "url": url,
                "error": f"HTTP {response.status_code}",
                "content": None,
            }

        if "application/pdf" not in content_type and not is_pdf_bytes(content):
            return {
                "code": code,
                "success": False,
                "url": url,
                "error": f"Not a PDF. Content-Type: {content_type}",
                "content": None,
            }

        if not is_pdf_bytes(content):
            return {
                "code": code,
                "success": False,
                "url": url,
                "error": "Downloaded file does not start with %PDF",
                "content": None,
            }

        # Validate that pypdf can read it
        PdfReader(io.BytesIO(content))

        return {
            "code": code,
            "success": True,
            "url": url,
            "error": "",
            "content": content,
        }

    except Exception as e:
        return {
            "code": code,
            "success": False,
            "url": url,
            "error": str(e),
            "content": None,
        }


def merge_pdfs(downloads: list[dict]) -> bytes:
    """
    Merge all successful PDFs into one PDF.
    """
    writer = PdfWriter()

    for item in downloads:
        if not item["success"]:
            continue

        reader = PdfReader(io.BytesIO(item["content"]))

        for page in reader.pages:
            writer.add_page(page)

    output = io.BytesIO()
    writer.write(output)
    return output.getvalue()


# ============================================================
# Streamlit Page Setup
# ============================================================

st.set_page_config(
    page_title="Philips Datasheet Pack Builder",
    page_icon="💡",
    layout="wide",
)


# ============================================================
# Philips-Style CSS
# ============================================================

st.markdown(
    """
    <style>
        :root {
            --philips-blue: #0b5ed7;
            --philips-deep-blue: #003b79;
            --philips-light-blue: #eaf3ff;
            --philips-bg: #f7f9fc;
            --philips-card: #ffffff;
            --philips-text: #102033;
            --philips-muted: #64748b;
            --philips-border: #dbe7f5;
            --success-green: #0f9f6e;
            --warning-orange: #f59e0b;
            --danger-red: #dc2626;
        }

        .stApp {
            background:
                radial-gradient(circle at top left, rgba(11, 94, 215, 0.10), transparent 32%),
                linear-gradient(180deg, #ffffff 0%, var(--philips-bg) 45%, #eef4fb 100%);
            color: var(--philips-text);
        }

        .block-container {
            padding-top: 28px;
            padding-bottom: 60px;
            max-width: 1180px;
        }

        .philips-topbar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 22px;
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
            box-shadow: 0 8px 20px rgba(0, 59, 121, 0.08);
        }

        .philips-badge {
            color: var(--philips-deep-blue);
            background: var(--philips-light-blue);
            border: 1px solid var(--philips-border);
            padding: 8px 14px;
            border-radius: 999px;
            font-size: 13px;
            font-weight: 600;
        }

        .hero {
            position: relative;
            overflow: hidden;
            padding: 34px;
            border-radius: 30px;
            background:
                linear-gradient(135deg, #003b79 0%, #0b5ed7 55%, #47a3ff 100%);
            color: white;
            margin-bottom: 28px;
            box-shadow: 0 24px 50px rgba(0, 59, 121, 0.24);
        }

        .hero::after {
            content: "";
            position: absolute;
            right: -90px;
            top: -90px;
            width: 260px;
            height: 260px;
            background: rgba(255, 255, 255, 0.16);
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
            max-width: 760px;
        }

        .hero-kicker {
            text-transform: uppercase;
            letter-spacing: 1.7px;
            font-size: 12px;
            font-weight: 700;
            opacity: 0.86;
            margin-bottom: 10px;
        }

        .hero h1 {
            margin: 0 0 12px 0;
            font-size: 42px;
            line-height: 1.08;
            font-weight: 800;
        }

        .hero p {
            margin: 0;
            font-size: 17px;
            line-height: 1.6;
            opacity: 0.92;
        }

        .tool-card {
            background: rgba(255, 255, 255, 0.88);
            border: 1px solid var(--philips-border);
            border-radius: 24px;
            padding: 22px;
            box-shadow: 0 16px 40px rgba(15, 23, 42, 0.07);
            backdrop-filter: blur(10px);
            margin-bottom: 18px;
        }

        .section-title {
            font-size: 19px;
            font-weight: 800;
            color: var(--philips-deep-blue);
            margin-bottom: 6px;
        }

        .section-subtitle {
            color: var(--philips-muted);
            font-size: 14px;
            margin-bottom: 14px;
        }

        .info-strip {
            background: #ffffff;
            border: 1px solid var(--philips-border);
            border-radius: 20px;
            padding: 16px 18px;
            margin: 18px 0;
            color: var(--philips-muted);
            box-shadow: 0 10px 28px rgba(15, 23, 42, 0.05);
        }

        div[data-testid="stMetric"] {
            background: white;
            border: 1px solid var(--philips-border);
            border-radius: 22px;
            padding: 18px;
            box-shadow: 0 12px 30px rgba(15, 23, 42, 0.06);
        }

        div[data-testid="stMetric"] label {
            color: var(--philips-muted) !important;
            font-weight: 600;
        }

        div[data-testid="stMetricValue"] {
            color: var(--philips-deep-blue);
            font-weight: 800;
        }

        .stTextArea textarea,
        .stTextInput input {
            border-radius: 16px !important;
            border-color: var(--philips-border) !important;
        }

        .stSelectbox div[data-baseweb="select"] {
            border-radius: 16px !important;
            border-color: var(--philips-border) !important;
        }

        .stFileUploader {
            background: rgba(255, 255, 255, 0.65);
            border-radius: 18px;
        }

        .stButton > button {
            background: linear-gradient(135deg, var(--philips-blue), var(--philips-deep-blue));
            color: white;
            border: 0;
            border-radius: 999px;
            padding: 13px 28px;
            font-weight: 800;
            box-shadow: 0 14px 28px rgba(11, 94, 215, 0.28);
            transition: all 0.2s ease;
        }

        .stButton > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 18px 34px rgba(11, 94, 215, 0.32);
            color: white;
        }

        .stDownloadButton > button {
            background: #ffffff;
            color: var(--philips-blue);
            border: 1px solid var(--philips-blue);
            border-radius: 999px;
            padding: 13px 28px;
            font-weight: 800;
        }

        .stDownloadButton > button:hover {
            background: var(--philips-light-blue);
            color: var(--philips-deep-blue);
            border: 1px solid var(--philips-deep-blue);
        }

        .small-note {
            color: var(--philips-muted);
            font-size: 13px;
        }

        hr {
            border-color: var(--philips-border);
        }

        @media screen and (max-width: 768px) {
            .philips-topbar {
                flex-direction: column;
                align-items: flex-start;
                gap: 12px;
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
    <div class="philips-topbar">
        <div class="philips-logo">PHILIPS</div>
        <div class="philips-badge">Signify Datasheet Automation</div>
    </div>

    <div class="hero">
        <div class="hero-content">
            <div class="hero-kicker">Product documentation tool</div>
            <h1>Philips Datasheet Pack Builder</h1>
            <p>
                Paste Philips / Signify product codes, upload Excel lists,
                download official datasheets, and merge everything into one clean PDF pack.
            </p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# Input Section
# ============================================================

left_col, right_col = st.columns([1, 1], gap="large")

with left_col:
    st.markdown(
        """
        <div class="tool-card">
            <div class="section-title">Paste product codes</div>
            <div class="section-subtitle">
                Add one or multiple Philips / Signify product codes.
            </div>
        """,
        unsafe_allow_html=True,
    )

    manual_codes_text = st.text_area(
        "Product codes",
        placeholder="Example:\n046677568283",
        height=65,
        label_visibility="collapsed",
    )

    st.markdown("</div>", unsafe_allow_html=True)


with right_col:
    st.markdown(
        """
        <div class="tool-card">
            <div class="section-title">Upload Excel file</div>
            <div class="section-subtitle">
                Select the column that contains the Philips product codes.
            </div>
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
        df_preview = pd.read_excel(uploaded_file)

        st.caption("Excel preview")
        st.dataframe(df_preview.head(), use_container_width=True)

        selected_column = st.selectbox(
            "Choose product code column",
            options=list(df_preview.columns),
        )

        excel_codes = extract_codes_from_excel(uploaded_file, selected_column)

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# Options Section
# ============================================================

st.markdown(
    """
    <div class="tool-card">
        <div class="section-title">Export settings</div>
        <div class="section-subtitle">
            Choose the final PDF filename and how failed downloads should be handled.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

settings_col_1, settings_col_2 = st.columns([2, 1], gap="large")

with settings_col_1:
    output_filename = st.text_input(
        "Output PDF filename",
        value="philips_datasheets_pack.pdf",
    )

with settings_col_2:
    skip_failed = st.checkbox(
        "Skip failed codes and continue",
        value=True,
    )


# ============================================================
# Code Summary
# ============================================================

manual_codes = extract_codes_from_text(manual_codes_text)
all_codes = dedupe_preserve_order(manual_codes + excel_codes)



st.markdown("### Summary before download")

metric_1, metric_2, metric_3 = st.columns(3)

with metric_1:
    st.metric("Manual codes", len(manual_codes))

with metric_2:
    st.metric("Excel codes", len(excel_codes))

with metric_3:
    st.metric("Unique total", len(all_codes))


if all_codes:
    with st.expander("View detected codes"):
        st.write(all_codes)


# ============================================================
# Download + Merge Action
# ============================================================

download_button = st.button(
    "Download and merge datasheets",
    type="primary",
    disabled=len(all_codes) == 0,
)


if download_button:
    start_time = time.time()

    st.info("Downloading Philips / Signify datasheets...")

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
        result = future.result()

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
