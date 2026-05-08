"""
Product Feature Scraper — Streamlit App (Pattern-branded)
Scrapes product sections (ingredients, how to use, key benefits, etc.)
from brand websites and exports a formatted Excel file.
"""

import streamlit as st
import time
import pandas as pd
from pathlib import Path
from scraper import ProductScraper, parse_input, build_excel

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Product Feature Scraper",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────────────────────
# PATTERN LOGO SVG
# ─────────────────────────────────────────────────────────────────────────────

PATTERN_LOGO_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 675 135.7" height="48" aria-label="Pattern">
<g>
  <path fill="#009BFF" d="M81.55,0.99L0.99,81.55c-1.32,1.32-1.32,3.47,0,4.8l19.84,19.84c1.32,1.32,3.47,1.32,4.8,0l80.56-80.56c1.32-1.32,1.32-3.47,0-4.8L86.35,0.99C85.02-0.33,82.88-0.33,81.55,0.99z"/>
  <path fill="#009BFF" d="M114.73,34.17L67.37,81.54c-1.32,1.32-1.32,3.47,0,4.8l19.84,19.84c1.32,1.32,3.47,1.32,4.8,0l47.36-47.36c1.32-1.32,1.32-3.47,0-4.8l-19.84-19.84C118.2,32.85,116.05,32.85,114.73,34.17z"/>
  <path fill="#F2F2F2" d="M254.36,64.21c0,24.35-18.47,42.98-40.69,42.98c-12.74,0-22.39-5.23-28.6-13.73v40.25c0,1.1-0.89,2-2,2h-13.65c-1.1,0-2-0.89-2-2V25.35c0-1.1,0.89-2,2-2h13.65c1.1,0,2,0.89,2,2v9.77c6.21-8.66,15.85-13.89,28.6-13.89C235.9,21.23,254.36,40.02,254.36,64.21z M236.71,64.21c0-15.2-11.11-26.15-25.82-26.15c-14.71,0-25.82,10.95-25.82,26.15c0,15.2,11.11,26.15,25.82,26.15C225.6,90.35,236.71,79.4,236.71,64.21z"/>
  <path fill="#F2F2F2" d="M347.84,25.35v77.71c0,1.1-0.89,2-2,2h-13.65c-1.1,0-2-0.89-2-2v-9.77c-6.21,8.66-15.85,13.89-28.6,13.89c-22.22,0-40.69-18.79-40.69-42.98c0-24.35,18.46-42.98,40.69-42.98c12.74,0,22.39,5.23,28.6,13.73v-9.6c0-1.1,0.89-2,2-2h13.65C346.95,23.35,347.84,24.25,347.84,25.35z M330.19,64.21c0-15.2-11.11-26.15-25.82-26.15s-25.82,10.95-25.82,26.15c0,15.2,11.11,26.15,25.82,26.15S330.19,79.4,330.19,64.21z"/>
  <path fill="#F2F2F2" d="M493.81,91.01c9.04,0,15.99-3.75,20.09-8.81c0.61-0.75,1.69-0.91,2.52-0.42l11.12,6.5c1.02,0.59,1.33,1.95,0.62,2.89c-7.59,10.04-19.37,16.03-34.51,16.03c-26.96,0-44.45-18.46-44.45-42.98c0-24.18,17.48-42.98,43.14-42.98c24.35,0,41.02,19.61,41.02,43.14c0,2.45-0.33,5.07-0.65,7.35h-65.04C470.44,84.47,480.74,91.01,493.81,91.01z M515.54,57.34c-2.45-14.05-12.75-20.1-23.37-20.1c-13.24,0-22.22,7.84-24.67,20.1H515.54z"/>
  <path fill="#F2F2F2" d="M675,54.89v48.17c0,1.1-0.89,2-2,2h-13.65c-1.1,0-2-0.89-2-2V56.69c0-12.42-7.19-18.96-18.3-18.96c-11.6,0-20.75,6.86-20.75,23.53v41.8c0,1.1-0.89,2-2,2h-13.65c-1.1,0-2-0.89-2-2V25.35c0-1.1,0.89-2,2-2h13.65c1.1,0,2,0.89,2,2v8.46c5.39-8.5,14.22-12.58,25.33-12.58C661.93,21.23,675,33.65,675,54.89z"/>
  <path fill="#F2F2F2" d="M583.58,21.88c-10.29,0-20.26,4.09-25.16,15.2V25.35c0-1.1-0.89-2-2-2h-13.65c-1.1,0-2,0.89-2,2v77.71c0,1.1,0.89,2,2,2h13.65c1.1,0,2-0.89,2-2V63.88c0-18.3,13.28-22.88,25.16-22.88h6.15c1.1,0,2-0.89,2-2V23.88c0-1.1-0.89-2-2-2H583.58z"/>
  <path fill="#F2F2F2" d="M397.4,40.35c1.1,0,2-0.89,2-2v-13c0-1.1-0.89-2-2-2h-21V2c0-1.1-0.89-2-2-2h-13.65c-1.1,0-2,0.89-2,2v78.96c0,16.77,8.09,24.83,24.97,24.83l13.68-0.01c1.1,0,2-0.9,2-2V91.42c0-1.1-0.9-2-2-2c-2.84,0-8.3-0.01-10.5-0.01c-8.05,0-10.5-2.09-10.5-9.85V40.35H397.4z"/>
  <path fill="#F2F2F2" d="M445.33,40.35c1.1,0,2-0.89,2-2v-13c0-1.1-0.89-2-2-2h-21V2c0-1.1-0.89-2-2-2h-13.65c-1.1,0-2,0.89-2,2v78.96c0,16.77,8.09,24.83,24.97,24.83l13.68-0.01c1.1,0,1.99-0.9,1.99-2V91.42c0-1.1-0.9-2-2-2c-2.84,0-8.3-0.01-10.5-0.01c-8.05,0-10.5-2.09-10.5-9.85V40.35H445.33z"/>
</g>
</svg>"""

# ─────────────────────────────────────────────────────────────────────────────
# FONT + CSS
# ─────────────────────────────────────────────────────────────────────────────

st.markdown(
    '<link href="https://fonts.googleapis.com/css2?family=Wix+Madefor+Display:wght@400;500;600;700;800&display=swap" rel="stylesheet">',
    unsafe_allow_html=True,
)

st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Wix+Madefor+Display:wght@400;500;600;700;800&display=swap');

  /* ── Hide Streamlit chrome ── */
  header[data-testid="stHeader"], #MainMenu,
  [data-testid="stToolbar"], [data-testid="stDecoration"],
  [data-testid="collapsedControl"], [data-testid="stSidebarCollapsedControl"],
  footer { display: none !important; }

  /* ── Global font — target Streamlit containers without breaking icon fonts ── */
  html, body { font-family: 'Wix Madefor Display', Inter, system-ui, sans-serif !important; }
  [data-testid="stMarkdownContainer"],
  [data-testid="stMarkdownContainer"] p,
  [data-testid="stMarkdownContainer"] li,
  [data-testid="stMarkdownContainer"] h1,
  [data-testid="stMarkdownContainer"] h2,
  [data-testid="stMarkdownContainer"] h3,
  .stMarkdown, .stMarkdown p, .stMarkdown li,
  .stTextInput input, .stTextArea textarea,
  .stButton > button, .stDownloadButton > button,
  [data-testid="stMetricLabel"], [data-testid="stMetricValue"],
  [data-testid="stCaptionContainer"] p {
    font-family: 'Wix Madefor Display', Inter, system-ui, sans-serif !important;
    -webkit-font-smoothing: antialiased;
  }

  /* ── Dark background #090A0F ── */
  .stApp, .main, .stApp > div, [data-testid="stAppViewContainer"] {
    background-color: #090A0F !important;
  }

  /* ── 4px gradient accent bar ── */
  .stApp::before {
    content: ''; position: fixed; top: 0; left: 0; right: 0;
    height: 4px; background: linear-gradient(90deg, #0096FA 0%, #770BFF 100%);
    z-index: 9999;
  }

  /* ── Content container ── */
  .main .block-container {
    background-color: #090A0F !important;
    padding-top: 3rem; padding-bottom: 4rem; max-width: 840px;
  }

  /* ── Headings ── */
  h1, h2, h3, h4, h5, h6 {
    font-family: 'Wix Madefor Display', Inter, system-ui, sans-serif !important;
    color: #FFFFFF !important;
  }
  h3 { font-size: 30px !important; font-weight: 600 !important; letter-spacing: -0.5px !important; line-height: 34px !important; }

  /* ── Body text ── */
  .stMarkdown p, .stMarkdown li {
    font-size: 16px !important; font-weight: 400 !important;
    line-height: 24px !important; color: #FFFFFF;
  }
  .stCaption, [data-testid="stCaptionContainer"] p, small {
    font-size: 13px !important; color: #C4D3E3 !important;
  }

  /* ── Input labels — overline ── */
  .stTextInput label, .stTextArea label {
    font-family: 'Wix Madefor Display', Inter, system-ui, sans-serif !important;
    font-size: 12px !important; font-weight: 500 !important;
    text-transform: uppercase !important; letter-spacing: 1.5px !important;
    color: #C4D3E3 !important;
  }

  /* ── Section overline + heading ── */
  .section-overline {
    font-family: 'Wix Madefor Display', Inter, system-ui, sans-serif;
    font-size: 12px; font-weight: 500; text-transform: uppercase;
    letter-spacing: 1.5px; color: #84C9F7; margin: 0 0 4px 0; line-height: 1;
  }
  .section-heading {
    font-family: 'Wix Madefor Display', Inter, system-ui, sans-serif;
    font-size: 30px; font-weight: 600; letter-spacing: -0.5px;
    line-height: 34px; color: #FFFFFF; margin: 0 0 8px 0;
  }

  /* ── Text inputs ── */
  .stTextInput input, .stTextArea textarea {
    background-color: #10121A !important;
    border: 1px solid rgba(196,211,227,0.15) !important;
    color: #FFFFFF !important; border-radius: 8px !important;
    font-family: 'Wix Madefor Display', Inter, system-ui, sans-serif !important;
    font-size: 15px !important; padding: 10px 14px !important;
  }
  .stTextInput input:focus, .stTextArea textarea:focus {
    border-color: #3A55FF !important;
    box-shadow: 0 0 0 2px rgba(58,85,255,0.2) !important; outline: none !important;
  }
  .stTextInput input::placeholder, .stTextArea textarea::placeholder {
    color: #C4D3E3 !important; opacity: 0.45;
  }

  /* ── Tooltip ── */
  div[data-baseweb="tooltip"] div, div[role="tooltip"] {
    background-color: #10121A !important; color: #FFFFFF !important;
    border: 1px solid rgba(196,211,227,0.15) !important;
    border-radius: 8px !important; font-size: 13px !important;
  }

  /* ── Run Scraper button — Pattern Purple ── */
  .stButton > button[kind="primary"] {
    background: #770BFF !important;
    border: none !important; color: #FFFFFF !important;
    font-family: 'Wix Madefor Display', Inter, system-ui, sans-serif !important;
    font-size: 12px !important; font-weight: 600 !important;
    text-transform: uppercase; letter-spacing: 1.5px;
    border-radius: 8px !important; padding: 12px 24px !important;
    transition: opacity 0.15s ease !important;
  }
  .stButton > button[kind="primary"]:hover { opacity: 0.85 !important; }
  .stButton > button[kind="primary"]:disabled { opacity: 0.28 !important; cursor: not-allowed !important; }

  /* ── Download scraped data button — CTA gradient ── */
  .stDownloadButton > button {
    background: linear-gradient(95deg, #009BFF 0%, #3A55FF 100%) !important;
    border: none !important; color: #FFFFFF !important;
    font-family: 'Wix Madefor Display', Inter, system-ui, sans-serif !important;
    font-size: 12px !important; font-weight: 600 !important;
    text-transform: uppercase; letter-spacing: 1.5px;
    border-radius: 8px !important; padding: 12px 24px !important;
  }
  .stDownloadButton > button:hover { opacity: 0.88 !important; }

  /* ── Template download — ghost pill ── */
  .template-link .stDownloadButton > button {
    background: transparent !important;
    border: 1px solid rgba(196,211,227,0.2) !important;
    color: #84C9F7 !important; font-size: 12px !important;
    font-weight: 500 !important; text-transform: none !important;
    letter-spacing: 0 !important; padding: 7px 16px !important;
    border-radius: 20px !important; box-shadow: none !important;
  }
  .template-link .stDownloadButton > button:hover {
    border-color: #3A55FF !important; color: #FFFFFF !important; opacity: 1 !important;
  }

  /* ── File uploader ── */
  [data-testid="stFileUploader"] {
    background-color: #10121A !important;
    border: 1px dashed rgba(196,211,227,0.2) !important;
    border-radius: 12px !important; padding: 0.75rem 1rem !important;
  }
  [data-testid="stFileUploader"]:hover { border-color: #3A55FF !important; }
  [data-testid="stFileUploader"] label,
  [data-testid="stFileUploader"] p,
  [data-testid="stFileUploader"] span {
    color: #C4D3E3 !important; text-transform: none !important;
    letter-spacing: 0 !important; font-size: 14px !important;
  }

  /* ── Progress bar — brand gradient ── */
  .stProgress > div > div > div > div {
    background: linear-gradient(90deg, #0096FA 0%, #770BFF 100%) !important; border-radius: 4px;
  }
  .stProgress > div > div {
    background-color: rgba(196,211,227,0.1) !important; border-radius: 4px; height: 6px !important;
  }

  /* ── Summary result card ── */
  .result-summary {
    background-color: #10121A;
    border: 1px solid rgba(196,211,227,0.12);
    border-radius: 12px; padding: 20px 24px; margin-bottom: 8px;
  }
  .result-headline {
    font-family: 'Wix Madefor Display', Inter, system-ui, sans-serif;
    font-size: 24px; font-weight: 600; color: #FFFFFF;
    letter-spacing: -0.5px; margin: 0 0 4px 0;
  }
  .result-sub {
    font-family: 'Wix Madefor Display', Inter, system-ui, sans-serif;
    font-size: 13px; color: #C4D3E3; margin: 0 0 16px 0;
  }
  .result-pills { display: flex; gap: 10px; flex-wrap: wrap; }
  .pill {
    font-family: 'Wix Madefor Display', Inter, system-ui, sans-serif;
    font-size: 12px; font-weight: 500; border-radius: 20px;
    padding: 5px 14px; display: inline-flex; align-items: center; gap: 6px;
  }
  .pill-success { background: rgba(132,201,247,0.12); color: #84C9F7; }
  .pill-warn    { background: rgba(255,200,87,0.12);  color: #FFC857; }
  .pill-error   { background: rgba(255,107,107,0.12); color: #FF6B6B; }
  .pill-miss    { background: rgba(196,211,227,0.08); color: #C4D3E3; }

  /* ── Alerts ── */
  [data-testid="stAlert"] {
    background-color: rgba(58,85,255,0.1) !important;
    border: 1px solid rgba(196,211,227,0.12) !important;
    border-left: 3px solid #3A55FF !important;
    border-radius: 8px !important; color: #FFFFFF !important;
  }
  [data-testid="stAlert"] p { font-size: 14px !important; color: #FFFFFF !important; }

  /* ── Expanders ── */
  [data-testid="stExpander"] {
    background-color: #10121A !important;
    border: 1px solid rgba(196,211,227,0.12) !important;
    border-radius: 12px !important;
  }
  [data-testid="stExpander"] summary {
    color: #C4D3E3 !important; font-weight: 500 !important; font-size: 14px !important;
  }
  [data-testid="stExpander"] summary:hover { color: #FFFFFF !important; }

  /* ── Divider ── */
  hr {
    border: none !important; height: 1px !important;
    background: linear-gradient(90deg, #0096FA 0%, #770BFF 100%) !important;
    opacity: 0.25; margin: 2rem 0 !important;
  }

  /* ── Spinner ── */
  .stSpinner > div { border-top-color: #770BFF !important; }

  /* ── Page hero (centred) ── */
  .page-hero {
    text-align: center !important;
    display: block !important;
    width: 100% !important;
    padding: 16px 0 32px;
  }
  .page-hero-logo { text-align: center !important; margin-bottom: 20px; }
  .page-hero-title {
    font-family: 'Wix Madefor Display', Inter, system-ui, sans-serif !important;
    font-size: 36px; font-weight: 600; color: #FFFFFF;
    letter-spacing: -0.5px; line-height: 42px; margin: 0 0 12px 0;
    text-align: center !important; display: block !important;
  }
  .page-hero-desc {
    font-family: 'Wix Madefor Display', Inter, system-ui, sans-serif !important;
    font-size: 16px; font-weight: 400; color: #C4D3E3;
    line-height: 26px; margin: 0 auto; max-width: 520px;
    text-align: center !important; display: block !important;
  }

</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# PAGE HEADER
# ─────────────────────────────────────────────────────────────────────────────

st.markdown(
    f"""
    <div class="page-hero" style="text-align:center;width:100%;display:block;">
      <div class="page-hero-logo" style="text-align:center;">{PATTERN_LOGO_SVG}</div>
      <p class="page-hero-title" style="text-align:center;display:block;font-family:'Wix Madefor Display',Inter,sans-serif;">Product Feature Scraper</p>
      <p class="page-hero-desc" style="text-align:center;display:block;margin:0 auto;font-family:'Wix Madefor Display',Inter,sans-serif;">
        Upload a product list. Configure the brand details. then hit Run.<br>A content-ready catalogue ready in no time.
      </p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.divider()


# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — UPLOAD
# ─────────────────────────────────────────────────────────────────────────────

_template_path = Path(__file__).parent / "input_products_TEMPLATE.xlsx"

_label_col, _tmpl_col = st.columns([5, 2])
with _label_col:
    st.markdown(
        '<p class="section-overline">Step 1</p>'
        '<p class="section-heading">Upload product list</p>',
        unsafe_allow_html=True,
    )
with _tmpl_col:
    if _template_path.exists():
        st.markdown('<div class="template-link">', unsafe_allow_html=True)
        st.download_button(
            label="⬇  Download template",
            data=_template_path.read_bytes(),
            file_name="input_products_TEMPLATE.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="template_download",
        )
        st.markdown('</div>', unsafe_allow_html=True)

st.caption(
    "Column A: Product name or title  |  "
    "Column B: UPC / barcode (optional)  |  "
    "Column C: Direct product URL (optional)"
)

uploaded_file = st.file_uploader(
    "Upload product list",
    type=["xlsx", "xls", "csv"],
    label_visibility="collapsed",
)

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            preview_df = pd.read_csv(uploaded_file, header=None, nrows=5, dtype=str)
        else:
            preview_df = pd.read_excel(
                uploaded_file, header=None, nrows=5, dtype=str, engine="openpyxl"
            )
        uploaded_file.seek(0)
        preview_df.columns = [f"Col {chr(65+i)}" for i in range(len(preview_df.columns))]
        with st.expander("Preview (first 5 rows)", expanded=True):
            st.table(preview_df)
    except Exception as e:
        st.warning(f"Could not preview file: {e}")

st.divider()


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 — BRAND CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

st.markdown(
    '<p class="section-overline">Step 2</p>'
    '<p class="section-heading">Brand configuration</p>'
    '<p style="font-size:14px;color:#C4D3E3;margin-bottom:20px;'
    'font-family:\'Wix Madefor Display\',Inter,sans-serif;">'
    'Configure once per brand — update any time before running.</p>',
    unsafe_allow_html=True,
)

cfg_col1, cfg_col2 = st.columns(2)

with cfg_col1:
    brand_name = st.text_input(
        "Brand name",
        value="",
        placeholder="e.g. K18",
        help="Used to name the output file.",
    )

with cfg_col2:
    store_url = st.text_input(
        "Store URL",
        value="",
        placeholder="e.g. https://k18hair.com.au",
        help="Paste the brand's website URL and click Check Site.",
    )

# ── Site check ────────────────────────────────────────────────────────────────
if "site_check_url"    not in st.session_state: st.session_state.site_check_url    = ""
if "site_is_shopify"   not in st.session_state: st.session_state.site_is_shopify   = None
if "site_search_url"   not in st.session_state: st.session_state.site_search_url   = ""

check_col, status_col = st.columns([1, 4])

with check_col:
    check_clicked = st.button(
        "🔍 Check site",
        disabled=not store_url,
        use_container_width=True,
    )

if check_clicked and store_url:
    with st.spinner("Checking site…"):
        _scraper_tmp = ProductScraper({
            "store_url": store_url,
            "sections": [],
            "brand_name": brand_name or "check",
        })
        _is_shopify = _scraper_tmp.is_shopify
        _has_catalogue = bool(_scraper_tmp._load_shopify_catalogue()) if _is_shopify else False

    st.session_state.site_check_url  = store_url
    st.session_state.site_is_shopify = _is_shopify
    # Auto-suggest search URL for non-Shopify or Shopify with blocked catalogue
    if not _is_shopify or not _has_catalogue:
        from urllib.parse import urlparse as _up
        _p = _up(store_url.strip())
        _root = f"{_p.scheme}://{_p.netloc}" if _p.netloc else store_url.strip().rstrip("/")
        st.session_state.site_search_url = f"{_root}/search?q={{query}}"
    else:
        st.session_state.site_search_url = ""

with status_col:
    if st.session_state.site_check_url == store_url and st.session_state.site_is_shopify is not None:
        if st.session_state.site_is_shopify and not st.session_state.site_search_url:
            st.success("✅ Shopify store detected — ready to scrape")
        elif st.session_state.site_is_shopify:
            st.warning("⚠️ Shopify store detected but product catalogue is restricted. A search URL has been pre-filled below — verify it's correct before running.")
        else:
            st.warning("⚠️ Non-Shopify site detected. A search URL has been pre-filled below — verify it's correct before running.")
    elif store_url and st.session_state.site_check_url != store_url:
        st.caption("Click **Check site** to verify this URL before running.")

# ── Search URL (shown when needed) ────────────────────────────────────────────
_show_search = (
    st.session_state.site_check_url == store_url
    and st.session_state.site_search_url
)

if _show_search:
    st.markdown(
        '<p style="font-size:12px;font-weight:500;text-transform:uppercase;letter-spacing:1.5px;'
        'color:#C4D3E3;margin:16px 0 4px 0;font-family:\'Wix Madefor Display\',Inter,sans-serif;">'
        'Search URL</p>',
        unsafe_allow_html=True,
    )
    st.caption("We've pre-filled this based on common search patterns. Edit if the site uses a different format.")
    search_url_pattern = st.text_input(
        "Search URL pattern",
        value=st.session_state.site_search_url,
        label_visibility="collapsed",
    )
else:
    # Hidden — only visible via manual expander fallback
    with st.expander("Advanced settings"):
        st.caption(
            "Only needed for non-Shopify sites. Click **Check site** above to auto-detect and pre-fill this."
        )
        search_url_pattern = st.text_input(
            "Search URL pattern (optional)",
            value=st.session_state.site_search_url,
            placeholder="https://brand.com/search?q={query}",
        )

st.markdown(
    '<p style="font-size:12px;font-weight:500;text-transform:uppercase;letter-spacing:1.5px;'
    'color:#C4D3E3;margin:16px 0 4px 0;font-family:\'Wix Madefor Display\',Inter,sans-serif;">'
    'Sections to extract</p>',
    unsafe_allow_html=True,
)
st.caption("One per line — copy the exact header name from the product page (e.g. how to use, ingredients).")
sections_raw = st.text_area(
    label="Sections to extract",
    value="",
    placeholder="how to use\nkey benefits\ningredients\nmore to know",
    height=120,
    label_visibility="collapsed",
)
sections = [s.strip() for s in sections_raw.splitlines() if s.strip()]

run_col, _ = st.columns([1, 4])
run_button = run_col.button(
    "▶  Run Scraper",
    type="primary",
    disabled=uploaded_file is None,
    use_container_width=True,
)

st.divider()


# ─────────────────────────────────────────────────────────────────────────────
# SCRAPING
# ─────────────────────────────────────────────────────────────────────────────

if run_button and uploaded_file:
    if not store_url.startswith("http"):
        st.error("Store URL must start with https://")
        st.stop()
    if not sections:
        st.error("Please add at least one section name above.")
        st.stop()

    config = {
        "brand_name":          brand_name,
        "store_url":           store_url,
        "sections":            sections,
        "search_url_pattern":  search_url_pattern,
        "accordion_selector":  ".accordion",
        "min_match_score":     0.6,
    }

    try:
        products = parse_input(uploaded_file, uploaded_file.name)
    except Exception as e:
        st.error(f"Could not read input file: {e}")
        st.stop()

    if not products:
        st.warning("No products found in the uploaded file.")
        st.stop()

    scraper = ProductScraper(config)

    with st.spinner("Detecting platform..."):
        platform = scraper.platform_label()
        if scraper.is_shopify:
            n = scraper.catalogue_size()
            st.info(
                f"**Shopify detected** — matched against {n} products in the {brand_name} store catalogue"
            )
        else:
            st.info(f"**Non-Shopify site** — fetching product pages via Jina AI Reader")

    total = len(products)
    st.markdown(f"### Scraping {total} product{'s' if total != 1 else ''}...")

    progress_bar   = st.progress(0)
    status_text    = st.empty()
    log_placeholder = st.empty()   # single placeholder — updated in-place (no duplicates)

    results   = []
    log_lines = []
    start_time = time.time()

    STATUS_ICON = {
        "OK":             "✅",
        "LOW CONFIDENCE": "⚠️",
        "NOT FOUND":      "❌",
        "ERROR":          "🔴",
    }

    for i, (name, upc, direct_url) in enumerate(products):
        status_text.markdown(f"Scraping **{i+1} of {total}:** {name[:80]}")
        result = scraper.scrape_product(name, upc=upc, direct_url=direct_url)
        results.append(result)

        icon     = STATUS_ICON.get(result["status"], "❓")
        title    = result.get("matched_title") or name
        log_line = f"{icon} **{title}**"
        if result["status"] != "OK":
            log_line += f" — {result['status']}"
            note = result.get("note", "")
            if note:
                log_line += f"  \n&nbsp;&nbsp;&nbsp;&nbsp;*{note}*"

        log_lines.append(log_line)
        log_placeholder.markdown("\n\n".join(log_lines), unsafe_allow_html=False)
        progress_bar.progress((i + 1) / total)
        time.sleep(0.3)

    elapsed = time.time() - start_time
    status_text.empty()
    progress_bar.empty()

    # ── Summary ──────────────────────────────────────────────────────────────
    ok        = sum(1 for r in results if r["status"] == "OK")
    low_conf  = sum(1 for r in results if r["status"] == "LOW CONFIDENCE")
    not_found = sum(1 for r in results if r["status"] == "NOT FOUND")
    errors    = sum(1 for r in results if r["status"] == "ERROR")

    pills_html = ""
    if ok:
        pills_html += f'<span class="pill pill-success">✅ {ok}/{total} scraped successfully</span>'
    if low_conf:
        pills_html += f'<span class="pill pill-warn">⚠️ {low_conf}/{total} low confidence</span>'
    if not_found:
        pills_html += f'<span class="pill pill-miss">❌ {not_found}/{total} not found</span>'
    if errors:
        pills_html += f'<span class="pill pill-error">🔴 {errors}/{total} errors</span>'

    st.markdown(
        f"""
        <div class="result-summary">
          <p class="result-headline">{ok} of {total} products scraped</p>
          <p class="result-sub">Completed in {elapsed:.0f}s</p>
          <div class="result-pills">{pills_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Download ──────────────────────────────────────────────────────────────
    st.divider()
    try:
        excel_buf = build_excel(results, sections)
        filename  = f"{brand_name.replace(' ', '_')}_product_features.xlsx"
        st.download_button(
            label="⬇  Download scraped data",
            data=excel_buf,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption(f"File: {filename}")
    except Exception as e:
        st.error(f"Could not build output file: {e}")

    # ── Results preview ───────────────────────────────────────────────────────
    with st.expander("Results preview", expanded=False):
        preview_cols = ["input_name", "matched_title", "status", "note"] + sections[:2]
        df_results   = pd.DataFrame(results)
        df_preview   = df_results[[c for c in preview_cols if c in df_results.columns]]
        st.table(df_preview)
