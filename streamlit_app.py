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
# PATTERN BRAND CSS
# ─────────────────────────────────────────────────────────────────────────────

PATTERN_LOGO_SVG = """<svg version="1.1" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 675 135.7" height="28" aria-label="Pattern">
<style>.pl0{fill:#009BFF;}.pl3{fill:#F2F2F2;}</style>
<g>
  <g>
    <path class="pl0" d="M81.55,0.99L0.99,81.55c-1.32,1.32-1.32,3.47,0,4.8l19.84,19.84c1.32,1.32,3.47,1.32,4.8,0l80.56-80.56c1.32-1.32,1.32-3.47,0-4.8L86.35,0.99C85.02-0.33,82.88-0.33,81.55,0.99z"/>
    <path class="pl0" d="M114.73,34.17L67.37,81.54c-1.32,1.32-1.32,3.47,0,4.8l19.84,19.84c1.32,1.32,3.47,1.32,4.8,0l47.36-47.36c1.32-1.32,1.32-3.47,0-4.8l-19.84-19.84C118.2,32.85,116.05,32.85,114.73,34.17z"/>
  </g>
  <path class="pl3" d="M254.36,64.21c0,24.35-18.47,42.98-40.69,42.98c-12.74,0-22.39-5.23-28.6-13.73v40.25c0,1.1-0.89,2-2,2h-13.65c-1.1,0-2-0.89-2-2V25.35c0-1.1,0.89-2,2-2h13.65c1.1,0,2,0.89,2,2v9.77c6.21-8.66,15.85-13.89,28.6-13.89C235.9,21.23,254.36,40.02,254.36,64.21z M236.71,64.21c0-15.2-11.11-26.15-25.82-26.15c-14.71,0-25.82,10.95-25.82,26.15c0,15.2,11.11,26.15,25.82,26.15C225.6,90.35,236.71,79.4,236.71,64.21z"/>
  <path class="pl3" d="M347.84,25.35v77.71c0,1.1-0.89,2-2,2h-13.65c-1.1,0-2-0.89-2-2v-9.77c-6.21,8.66-15.85,13.89-28.6,13.89c-22.22,0-40.69-18.79-40.69-42.98c0-24.35,18.46-42.98,40.69-42.98c12.74,0,22.39,5.23,28.6,13.73v-9.6c0-1.1,0.89-2,2-2h13.65C346.95,23.35,347.84,24.25,347.84,25.35z M330.19,64.21c0-15.2-11.11-26.15-25.82-26.15s-25.82,10.95-25.82,26.15c0,15.2,11.11,26.15,25.82,26.15S330.19,79.4,330.19,64.21z"/>
  <path class="pl3" d="M493.81,91.01c9.04,0,15.99-3.75,20.09-8.81c0.61-0.75,1.69-0.91,2.52-0.42l11.12,6.5c1.02,0.59,1.33,1.95,0.62,2.89c-7.59,10.04-19.37,16.03-34.51,16.03c-26.96,0-44.45-18.46-44.45-42.98c0-24.18,17.48-42.98,43.14-42.98c24.35,0,41.02,19.61,41.02,43.14c0,2.45-0.33,5.07-0.65,7.35h-65.04C470.44,84.47,480.74,91.01,493.81,91.01z M515.54,57.34c-2.45-14.05-12.75-20.1-23.37-20.1c-13.24,0-22.22,7.84-24.67,20.1H515.54z"/>
  <path class="pl3" d="M675,54.89v48.17c0,1.1-0.89,2-2,2h-13.65c-1.1,0-2-0.89-2-2V56.69c0-12.42-7.19-18.96-18.3-18.96c-11.6,0-20.75,6.86-20.75,23.53v41.8c0,1.1-0.89,2-2,2h-13.65c-1.1,0-2-0.89-2-2V25.35c0-1.1,0.89-2,2-2h13.65c1.1,0,2,0.89,2,2v8.46c5.39-8.5,14.22-12.58,25.33-12.58C661.93,21.23,675,33.65,675,54.89z"/>
  <path class="pl3" d="M583.58,21.88c-10.29,0-20.26,4.09-25.16,15.2V25.35c0-1.1-0.89-2-2-2h-13.65c-1.1,0-2,0.89-2,2v77.71c0,1.1,0.89,2,2,2h13.65c1.1,0,2-0.89,2-2V63.88c0-18.3,13.28-22.88,25.16-22.88h6.15c1.1,0,2-0.89,2-2V23.88c0-1.1-0.89-2-2-2H583.58z"/>
  <path class="pl3" d="M397.4,40.35c1.1,0,2-0.89,2-2v-13c0-1.1-0.89-2-2-2h-21V2c0-1.1-0.89-2-2-2h-13.65c-1.1,0-2,0.89-2,2v78.96c0,16.77,8.09,24.83,24.97,24.83l13.68-0.01c1.1,0,2-0.9,2-2V91.42c0-1.1-0.9-2-2-2c-2.84,0-8.3-0.01-10.5-0.01c-8.05,0-10.5-2.09-10.5-9.85V40.35H397.4z"/>
  <path class="pl3" d="M445.33,40.35c1.1,0,2-0.89,2-2v-13c0-1.1-0.89-2-2-2h-21V2c0-1.1-0.89-2-2-2h-13.65c-1.1,0-2,0.89-2,2v78.96c0,16.77,8.09,24.83,24.97,24.83l13.68-0.01c1.1,0,1.99-0.9,1.99-2V91.42c0-1.1-0.9-2-2-2c-2.84,0-8.3-0.01-10.5-0.01c-8.05,0-10.5-2.09-10.5-9.85V40.35H445.33z"/>
</g>
</svg>"""

st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Wix+Madefor+Display:wght@400;500;600;700;800&display=swap');

  /* ── Design tokens (Pattern brand system) ── */
  :root {{
    --brand-blue:      #3A55FF;
    --bright-blue:     #009BFF;
    --brand-purple:    #770BFF;
    --brand-cyan:      #0096FA;
    --light-cyan:      #84C9F7;
    --dark-primary:    #090A0F;
    --dark-secondary:  #10121A;
    --light-gray:      #EBF0F5;
    --muted-gray:      #C4D3E3;
    --white:           #FFFFFF;
    --gradient-brand:  linear-gradient(90deg, #0096FA 0%, #770BFF 100%);
    --gradient-cta:    linear-gradient(95deg, #009BFF 0%, #3A55FF 100%);
    --shadow-base:     0 2px 3px rgba(7,7,8,0.5);
    --shadow-lg:       0 2px 3px rgba(7,7,8,0.5), 0 14px 14px rgba(7,7,8,0.5), 0 45px 45px rgba(7,7,8,0.5);
    --radius-card:     12px;
    --radius-button:   8px;
    --radius-pill:     20px;
    --font-display:    'Wix Madefor Display', Inter, system-ui, sans-serif;
    --border-subtle:   1px solid rgba(196, 211, 227, 0.12);
  }}

  /* ── Hide Streamlit chrome ── */
  header[data-testid="stHeader"],
  #MainMenu,
  [data-testid="stToolbar"],
  [data-testid="stDecoration"],
  [data-testid="collapsedControl"],
  [data-testid="stSidebarCollapsedControl"],
  footer {{ display: none !important; }}

  /* ── Global font ── */
  html, body, [class*="css"] {{
    font-family: var(--font-display) !important;
    -webkit-font-smoothing: antialiased;
  }}

  /* ── App + page background ── */
  .stApp, .main {{
    background-color: var(--dark-primary) !important;
  }}

  /* ── 4px gradient accent bar pinned to top ── */
  .stApp::before {{
    content: '';
    position: fixed;
    top: 0; left: 0; right: 0;
    height: 4px;
    background: var(--gradient-brand);
    z-index: 9999;
  }}

  /* ── Content container ── */
  .main .block-container {{
    background-color: var(--dark-primary) !important;
    padding-top: 3rem;
    padding-bottom: 4rem;
    max-width: 840px;
  }}

  /* ── Type scale (Pattern spec) ── */
  h1, h2, h3, h4, h5, h6 {{
    font-family: var(--font-display) !important;
    color: var(--white) !important;
  }}
  /* H3 — 30px SemiBold -0.5px */
  h3 {{
    font-size: 30px !important;
    font-weight: 600 !important;
    letter-spacing: -0.5px !important;
    line-height: 34px !important;
  }}
  /* H4 — 24px Medium -0.5px */
  h4 {{
    font-size: 24px !important;
    font-weight: 500 !important;
    letter-spacing: -0.5px !important;
    line-height: 28px !important;
  }}
  /* Body Copy 2 — 16px Regular */
  p, li {{
    font-family: var(--font-display) !important;
    font-size: 16px !important;
    font-weight: 400 !important;
    line-height: 24px !important;
    color: var(--white);
  }}
  /* Caption / Body 3 — 14px Regular */
  .stCaption, [data-testid="stCaptionContainer"] p, small {{
    font-size: 14px !important;
    font-weight: 400 !important;
    line-height: 22px !important;
    color: var(--muted-gray) !important;
  }}

  /* ── Overline labels ── */
  /* Used on input field labels and custom .overline class */
  .stTextInput label,
  .stTextArea label,
  .overline {{
    font-family: var(--font-display) !important;
    font-size: 12px !important;
    font-weight: 500 !important;
    text-transform: uppercase !important;
    letter-spacing: 1.5px !important;
    color: var(--muted-gray) !important;
  }}

  /* Section heading with overline pattern */
  .section-overline {{
    font-size: 12px;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: var(--light-cyan);
    margin-bottom: 4px;
    line-height: 1;
  }}
  .section-heading {{
    font-size: 30px;
    font-weight: 600;
    letter-spacing: -0.5px;
    line-height: 34px;
    color: var(--white);
    margin-bottom: 6px;
  }}

  /* ── Text inputs ── */
  .stTextInput input, .stTextArea textarea {{
    background-color: var(--dark-secondary) !important;
    border: var(--border-subtle) !important;
    color: var(--white) !important;
    border-radius: var(--radius-button) !important;
    font-family: var(--font-display) !important;
    font-size: 16px !important;
    font-weight: 400 !important;
    padding: 10px 14px !important;
    box-shadow: var(--shadow-base) !important;
  }}
  .stTextInput input:focus, .stTextArea textarea:focus {{
    border-color: var(--brand-blue) !important;
    box-shadow: 0 0 0 2px rgba(58,85,255,0.2) !important;
    outline: none !important;
  }}
  .stTextInput input::placeholder, .stTextArea textarea::placeholder {{
    color: var(--muted-gray) !important;
    opacity: 0.5;
  }}

  /* ── Tooltip ── */
  [data-testid="stTooltipHoverTarget"] svg {{
    fill: var(--muted-gray) !important;
    opacity: 0.5;
  }}
  div[data-baseweb="tooltip"] div,
  div[role="tooltip"] {{
    background-color: var(--dark-secondary) !important;
    color: var(--white) !important;
    border: var(--border-subtle) !important;
    border-radius: var(--radius-button) !important;
    font-family: var(--font-display) !important;
    font-size: 13px !important;
    box-shadow: var(--shadow-lg) !important;
  }}

  /* ── CTA button — Run Scraper ── */
  .stButton > button[kind="primary"] {{
    background: var(--gradient-cta) !important;
    border: none !important;
    color: var(--white) !important;
    font-family: var(--font-display) !important;
    font-size: 12px !important;
    font-weight: 600 !important;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    border-radius: var(--radius-button) !important;
    padding: 12px 24px !important;
    transition: opacity 0.15s ease !important;
    box-shadow: var(--shadow-base) !important;
  }}
  .stButton > button[kind="primary"]:hover {{ opacity: 0.88 !important; }}
  .stButton > button[kind="primary"]:disabled {{
    opacity: 0.28 !important;
    cursor: not-allowed !important;
  }}

  /* ── Download button (Excel results) ── */
  .stDownloadButton > button {{
    background: var(--gradient-cta) !important;
    border: none !important;
    color: var(--white) !important;
    font-family: var(--font-display) !important;
    font-size: 12px !important;
    font-weight: 600 !important;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    border-radius: var(--radius-button) !important;
    padding: 12px 24px !important;
    box-shadow: var(--shadow-base) !important;
  }}
  .stDownloadButton > button:hover {{ opacity: 0.88 !important; }}

  /* ── Template download (ghost pill button) ── */
  .template-link .stDownloadButton > button {{
    background: transparent !important;
    border: var(--border-subtle) !important;
    color: var(--light-cyan) !important;
    font-size: 12px !important;
    font-weight: 500 !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
    padding: 7px 16px !important;
    border-radius: var(--radius-pill) !important;
    box-shadow: none !important;
  }}
  .template-link .stDownloadButton > button:hover {{
    border-color: var(--brand-blue) !important;
    color: var(--white) !important;
    opacity: 1 !important;
  }}

  /* ── Config section card ── */
  .config-section {{
    background: var(--dark-secondary);
    border: var(--border-subtle);
    border-radius: var(--radius-card);
    padding: 28px 32px 32px;
    margin-bottom: 8px;
    box-shadow: var(--shadow-base);
  }}

  /* ── File uploader ── */
  [data-testid="stFileUploader"] {{
    background-color: var(--dark-secondary) !important;
    border: 1px dashed rgba(196,211,227,0.2) !important;
    border-radius: var(--radius-card) !important;
    padding: 0.75rem 1rem !important;
    transition: border-color 0.15s ease;
    box-shadow: var(--shadow-base) !important;
  }}
  [data-testid="stFileUploader"]:hover {{
    border-color: var(--brand-blue) !important;
  }}
  [data-testid="stFileUploader"] label,
  [data-testid="stFileUploader"] p,
  [data-testid="stFileUploader"] span {{
    color: var(--muted-gray) !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
    font-size: 14px !important;
  }}

  /* ── Progress bar ── */
  .stProgress > div > div > div > div {{
    background: var(--gradient-brand) !important;
    border-radius: 4px;
  }}
  .stProgress > div > div {{
    background-color: rgba(196,211,227,0.12) !important;
    border-radius: 4px;
    height: 6px !important;
  }}

  /* ── Metric cards ── */
  [data-testid="stMetric"] {{
    background-color: var(--dark-secondary) !important;
    border: var(--border-subtle) !important;
    border-radius: var(--radius-card) !important;
    padding: 1.25rem 1.5rem !important;
    box-shadow: var(--shadow-base) !important;
  }}
  [data-testid="stMetricLabel"] {{
    color: var(--muted-gray) !important;
    font-size: 11px !important;
    font-weight: 500 !important;
    text-transform: uppercase !important;
    letter-spacing: 1.5px !important;
  }}
  [data-testid="stMetricValue"] {{
    color: var(--white) !important;
    font-size: 48px !important;
    font-weight: 600 !important;
    letter-spacing: -2px !important;
  }}

  /* ── Alerts ── */
  [data-testid="stAlert"] {{
    background-color: rgba(58,85,255,0.1) !important;
    border: var(--border-subtle) !important;
    border-left: 3px solid var(--brand-blue) !important;
    border-radius: var(--radius-button) !important;
    color: var(--white) !important;
  }}
  [data-testid="stAlert"] p {{ font-size: 14px !important; }}

  /* ── Expanders ── */
  [data-testid="stExpander"] {{
    background-color: var(--dark-secondary) !important;
    border: var(--border-subtle) !important;
    border-radius: var(--radius-card) !important;
    box-shadow: var(--shadow-base) !important;
  }}
  [data-testid="stExpander"] summary {{
    color: var(--muted-gray) !important;
    font-weight: 500 !important;
    font-size: 14px !important;
  }}
  [data-testid="stExpander"] summary:hover {{ color: var(--white) !important; }}

  /* ── Divider — thin gradient line ── */
  hr {{
    border: none !important;
    height: 1px !important;
    background: var(--gradient-brand) !important;
    opacity: 0.25;
    margin: 2rem 0 !important;
  }}

  /* ── Spinner ── */
  .stSpinner > div {{ border-top-color: var(--brand-blue) !important; }}

  /* ── Dataframe ── */
  [data-testid="stDataFrame"] {{ border-radius: var(--radius-card) !important; overflow: hidden; }}
  .dvn-scroller {{ background-color: var(--dark-secondary) !important; }}

  /* ── Page header row ── */
  .page-header {{
    display: flex;
    align-items: center;
    gap: 16px;
    margin-bottom: 6px;
    padding-top: 4px;
  }}
  .page-title {{
    font-family: var(--font-display);
    font-size: 30px;
    font-weight: 600;
    color: var(--white);
    letter-spacing: -0.5px;
    line-height: 34px;
    margin: 0;
  }}
  .page-divider {{
    width: 1px;
    height: 26px;
    background: rgba(196,211,227,0.2);
    display: inline-block;
  }}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# PAGE HEADER (logo + title)
# ─────────────────────────────────────────────────────────────────────────────

st.markdown(
    f"""
    <div class="page-header">
      {PATTERN_LOGO_SVG}
      <span class="page-divider"></span>
      <p class="page-title">Product Feature Scraper</p>
    </div>
    <p style="font-size:16px;font-weight:400;color:#C4D3E3;line-height:24px;margin:10px 0 0;">
      Fill in the brand details below, upload your product list, then hit Run.
      Works with Shopify stores and any other brand website.
    </p>
    """,
    unsafe_allow_html=True,
)

st.divider()


# ─────────────────────────────────────────────────────────────────────────────
# BRAND CONFIGURATION (main body)
# ─────────────────────────────────────────────────────────────────────────────

st.markdown(
    '<p class="section-overline">Step 1</p>'
    '<p class="section-heading">Brand configuration</p>'
    '<p style="font-size:14px;color:#C4D3E3;margin-bottom:20px;">Configure once per brand — update any time before running.</p>',
    unsafe_allow_html=True,
)

cfg_col1, cfg_col2 = st.columns(2)

with cfg_col1:
    brand_name = st.text_input(
        "Brand name",
        value="K18",
        help="Used to name the output file.",
    )

with cfg_col2:
    store_url = st.text_input(
        "Store URL",
        value="https://k18hair.com.au",
        help="The brand's website. The scraper auto-detects whether it's Shopify.",
    )

st.markdown(
    '<p class="field-label">Sections to extract</p>',
    unsafe_allow_html=True,
)
st.caption("One section per line — copy the exact header name from the product page (e.g. how to use, ingredients).")
sections_raw = st.text_area(
    label="Sections to extract",
    value="how to use\nkey benefits\ningredients\nmore to know",
    height=120,
    label_visibility="collapsed",
)
sections = [s.strip() for s in sections_raw.splitlines() if s.strip()]

# Advanced — non-Shopify search pattern (hidden by default)
with st.expander("Advanced settings"):
    st.caption(
        "Only needed for non-Shopify sites where you don't have direct product URLs "
        "in column C of your input file. Paste the site's search URL and replace the "
        "search query part with {query} — e.g. https://brand.com/search?q={query}"
    )
    search_url_pattern = st.text_input(
        "Search URL pattern (optional)",
        value="",
        placeholder="https://brand.com/search?q={query}",
    )

st.divider()


# ─────────────────────────────────────────────────────────────────────────────
# FILE UPLOAD
# ─────────────────────────────────────────────────────────────────────────────

_template_path = Path(__file__).parent / "input_products_TEMPLATE.xlsx"

_label_col, _tmpl_col = st.columns([5, 2])
with _label_col:
    st.markdown(
        '<p class="section-overline">Step 2</p>'
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
    "Column A: Product name or title &nbsp;|&nbsp; "
    "Column B: UPC / barcode (optional) &nbsp;|&nbsp; "
    "Column C: Direct product URL (optional)"
)

uploaded_file = st.file_uploader(
    "Upload product list",
    type=["xlsx", "xls", "csv"],
    label_visibility="collapsed",
)

# Preview uploaded file
if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            preview_df = pd.read_csv(uploaded_file, header=None, nrows=5, dtype=str)
        else:
            preview_df = pd.read_excel(uploaded_file, header=None, nrows=5, dtype=str)
        uploaded_file.seek(0)
        with st.expander("Preview (first 5 rows)", expanded=False):
            st.dataframe(preview_df, use_container_width=True, hide_index=True)
    except Exception:
        pass

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
            st.info(f"✅ **{platform}** — {n} products loaded from store catalogue")
        else:
            st.info(f"🌐 **{platform}** — will use Jina AI Reader for all requests")

    st.markdown(f"### Processing {len(products)} product(s)")
    progress_bar  = st.progress(0)
    status_text   = st.empty()
    log_container = st.container()

    results    = []
    log_lines  = []
    start_time = time.time()

    STATUS_ICON = {
        "OK":             "✅",
        "LOW CONFIDENCE": "⚠️",
        "NOT FOUND":      "❌",
        "ERROR":          "🔴",
    }

    for i, (name, upc, direct_url) in enumerate(products):
        status_text.markdown(f"**Processing {i+1} / {len(products)}:** {name[:80]}")
        result = scraper.scrape_product(name, upc=upc, direct_url=direct_url)
        results.append(result)

        icon     = STATUS_ICON.get(result["status"], "❓")
        title    = result.get("matched_title") or name
        log_line = f"{icon} **{title}** — {result['status']}"
        if result["status"] != "OK":
            log_line += f"  \n&nbsp;&nbsp;&nbsp;&nbsp;_{result.get('note', '')}_"

        log_lines.append(log_line)
        with log_container:
            st.markdown("\n\n".join(log_lines), unsafe_allow_html=False)

        progress_bar.progress((i + 1) / len(products))
        time.sleep(0.3)

    elapsed = time.time() - start_time
    status_text.empty()
    progress_bar.empty()

    ok        = sum(1 for r in results if r["status"] == "OK")
    low_conf  = sum(1 for r in results if r["status"] == "LOW CONFIDENCE")
    not_found = sum(1 for r in results if r["status"] == "NOT FOUND")
    errors    = sum(1 for r in results if r["status"] == "ERROR")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("✅ OK",             ok)
    c2.metric("⚠️ Low confidence", low_conf)
    c3.metric("❌ Not found",      not_found)
    c4.metric("🔴 Errors",         errors)

    st.caption(f"Completed in {elapsed:.0f}s")

    st.divider()
    try:
        excel_buf = build_excel(results, sections)
        filename  = f"{brand_name.replace(' ','_')}_product_features.xlsx"
        st.download_button(
            label="⬇  Download Excel",
            data=excel_buf,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption(f"File: {filename}")
    except Exception as e:
        st.error(f"Could not build Excel file: {e}")

    with st.expander("Results preview", expanded=False):
        preview_cols = ["input_name", "matched_title", "status", "note"] + sections[:2]
        df_preview   = pd.DataFrame(results)[
            [c for c in preview_cols if c in pd.DataFrame(results).columns]
        ]
        st.dataframe(df_preview, use_container_width=True, hide_index=True)
