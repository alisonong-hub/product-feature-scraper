"""
scraper.py — Core scraping logic for the Product Feature Scraper.
Imported by streamlit_app.py. No UI or file I/O here.
"""

import requests
import re
import time
import io
from difflib import SequenceMatcher
from urllib.parse import urlparse
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

JINA_BASE = "https://r.jina.ai/"


# ─────────────────────────────────────────────────────────────────────────────
# SCRAPER CLASS
# ─────────────────────────────────────────────────────────────────────────────

class ProductScraper:
    _STOP_WORDS = {
        "k18","hair","the","and","for","with","from","into","your","that","this",
        "size","pack","set","kit","new","best","top","oz","ml","fl","ounce",
        "care","style","use","day","free","safe","non","after","while","each",
        "all","add","per","also",
    }
    _DEPRIORITISE_WORDS = {
        "salon","offer","deal","valued","trio","bundle","intro","pro service"
    }

    def __init__(self, config):
        self.config    = config
        # Normalise to root domain — strip any path the user may have pasted
        # e.g. https://amika.com/collections/all → https://amika.com
        raw_url = config["store_url"].strip().rstrip("/")
        parsed  = urlparse(raw_url)
        self.store_url = f"{parsed.scheme}://{parsed.netloc}" if parsed.netloc else raw_url
        self.session   = requests.Session()
        self.session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept-Language": "en-AU,en;q=0.9",
        })
        self._product_cache = None
        self._is_shopify    = None

    # ── Platform detection ────────────────────────────────────────────────────

    @property
    def is_shopify(self):
        if self._is_shopify is None:
            self._is_shopify = self._detect_shopify()
        return self._is_shopify

    def platform_label(self):
        return "Shopify" if self.is_shopify else "Non-Shopify (Jina mode)"

    def _detect_shopify(self):
        # Check 1: standard products.json endpoint
        try:
            resp = self.session.get(
                f"{self.store_url}/products.json?limit=1", timeout=8
            )
            if resp.status_code == 200 and "products" in resp.json():
                return True
        except Exception:
            pass

        # Check 2: HTML fingerprint — some Shopify stores block products.json
        # but still load cdn.shopify.com assets in their page source
        try:
            resp = self.session.get(self.store_url, timeout=10)
            html = resp.text.lower()
            shopify_signals = [
                "cdn.shopify.com",
                "shopify.com/s/",
                "myshopify.com",
                "shopify.theme",
                '"shopify"',
            ]
            if any(sig in html for sig in shopify_signals):
                return True
        except Exception:
            pass

        return False

    # ── Jina AI Reader ────────────────────────────────────────────────────────

    def _fetch_via_jina(self, url):
        resp = self.session.get(
            JINA_BASE + url,
            timeout=30,
            headers={"Accept": "text/markdown", "X-No-Cache": "true"},
        )
        resp.raise_for_status()
        return resp.text

    def _parse_sections_from_markdown(self, markdown):
        sections, current_key, lines = {}, None, []
        for line in markdown.split("\n"):
            h = re.match(r"^#{1,4}\s+(.+)$", line)
            if h:
                if current_key:
                    sections[current_key] = "\n".join(lines).strip()
                current_key = h.group(1).strip().lower().rstrip(":")
                lines = []
            elif current_key is not None:
                lines.append(line)
        if current_key and lines:
            sections[current_key] = "\n".join(lines).strip()
        return sections

    # ── Scoring helpers ───────────────────────────────────────────────────────

    def _fuzzy_score(self, a, b):
        return SequenceMatcher(None, a.lower().strip(), b.lower().strip()).ratio()

    def _keyword_score(self, search_term, product_title):
        search_text = re.sub(r"[^a-z0-9\s]", " ", search_term.lower())
        title_words = re.sub(r"[^a-z0-9\s]", " ", product_title.lower()).split()
        meaningful  = [w for w in title_words if len(w) >= 3 and w not in self._STOP_WORDS]
        if not meaningful:
            return 0.0
        return sum(1 for w in meaningful if w in search_text) / len(meaningful)

    def _retail_penalty(self, title):
        t = title.lower()
        return 0.15 if any(w in t for w in self._DEPRIORITISE_WORDS) else 0.0

    # ── Shopify product discovery ─────────────────────────────────────────────

    def _load_shopify_catalogue(self):
        if self._product_cache is not None:
            return self._product_cache
        products, page = [], 1
        while True:
            try:
                resp  = self.session.get(
                    f"{self.store_url}/products.json?limit=250&page={page}", timeout=15
                )
                batch = resp.json().get("products", [])
            except Exception:
                break
            if not batch:
                break
            products.extend(batch)
            page += 1
            time.sleep(0.3)
        self._product_cache = products
        return products

    def catalogue_size(self):
        """Return number of products loaded (triggers load if needed)."""
        return len(self._load_shopify_catalogue())

    def _find_shopify_product(self, search_term, upc=None):
        products = self._load_shopify_catalogue()

        if upc:
            upc_str = str(upc).strip().split(".")[0]
            for p in products:
                for v in p.get("variants", []):
                    if str(v.get("barcode", "")).strip() == upc_str:
                        return (
                            f"{self.store_url}/products/{p['handle']}",
                            p["title"], "HIGH", f"Matched by UPC {upc_str}",
                        )

        search_clean = search_term.strip().lower()
        for p in products:
            if p["title"].strip().lower() == search_clean:
                return (
                    f"{self.store_url}/products/{p['handle']}",
                    p["title"], "HIGH", "Exact title match",
                )

        best_score, best = 0, None
        for p in products:
            score = (
                self._keyword_score(search_term, p["title"])
                + 0.05 * self._fuzzy_score(search_term, p["title"])
                - self._retail_penalty(p["title"])
            )
            if score > best_score:
                best_score, best = score, p

        if best and best_score >= 0.65:
            display_score = min(best_score, 1.0)
            reason = f"Keyword match ({display_score:.0%})"
            if upc:
                reason += f" — UPC {upc} not found in AU store (may be a US barcode)"
            return (
                f"{self.store_url}/products/{best['handle']}",
                best["title"], "HIGH", reason,
            )

        if best and best_score >= 0.35:
            kw = self._keyword_score(search_term, best["title"])
            parts = []
            if upc:
                parts.append(f"UPC {upc} not in AU store")
            if kw < 0.5:
                parts.append(f"Low keyword overlap ({kw:.0%}) — may be a bundle or US-only listing")
            else:
                parts.append(f"Score {best_score:.0%} below confidence threshold")
            return (
                f"{self.store_url}/products/{best['handle']}",
                best["title"], "LOW", "; ".join(parts),
            )

        reason = "No matching product found"
        if upc:
            reason += f" (UPC {upc} not in catalogue)"
        return None, None, None, reason

    # ── Non-Shopify discovery ─────────────────────────────────────────────────

    def _find_jina_product(self, search_term, upc=None, direct_url=None):
        if direct_url:
            return direct_url, search_term, "HIGH", "Direct URL provided"

        pattern = self.config.get("search_url_pattern", "")
        if not pattern:
            return (
                None, None, None,
                "No direct URL or search_url_pattern set. "
                "Add a product URL in column C of your input file.",
            )

        search_url = pattern.replace(
            "{query}", requests.utils.quote(
                re.sub(r"[^a-z0-9\s]", " ", search_term.lower()).strip()
            )
        )
        try:
            markdown = self._fetch_via_jina(search_url)
        except Exception as e:
            return None, None, None, f"Search page fetch failed: {e}"

        links = re.findall(r"\[([^\]]+)\]\((https?://[^\)]+)\)", markdown)
        if not links:
            return None, None, None, "No product links found in search results"

        best_score, best_link = 0, None
        for text, url in links:
            score = self._keyword_score(search_term, text) + 0.05 * self._fuzzy_score(search_term, text)
            if score > best_score:
                best_score, best_link = score, (text, url)

        if best_link and best_score >= 0.4:
            return best_link[1], best_link[0], "HIGH", f"Search match ({min(best_score, 1.0):.0%})"

        return None, None, None, f"No confident search result (best score {min(best_score, 1.0):.0%})"

    # ── Unified product finder ────────────────────────────────────────────────

    def find_product(self, search_term, upc=None, direct_url=None):
        if self.is_shopify:
            return self._find_shopify_product(search_term, upc)
        else:
            return self._find_jina_product(search_term, upc, direct_url)

    # ── HTML content extraction ───────────────────────────────────────────────

    def _fetch_html(self, url):
        resp = self.session.get(url, timeout=15)
        resp.raise_for_status()
        return resp.text

    def _parse_html_sections(self, html):
        soup     = BeautifulSoup(html, "lxml")
        sections = {}

        accordion = soup.select_one(self.config.get("accordion_selector", ".accordion"))
        if accordion:
            for details in accordion.find_all("details"):
                summary = details.find("summary")
                if not summary:
                    continue
                parts = []
                for node in summary.children:
                    if not getattr(node, "name", None):
                        t = str(node).strip()
                        if t:
                            parts.append(t)
                    elif node.name not in ("img", "span", "svg"):
                        t = node.get_text(strip=True)
                        if t:
                            parts.append(t)
                header = " ".join(parts).strip().rstrip(":").strip()
                content_div = details.find(class_="details-content") or details
                content     = self._extract_html_text(content_div)
                if header:
                    sections[header.lower()] = content

        if not sections:
            for tag in soup.find_all(["h2", "h3", "h4", "h5", "strong"]):
                header = tag.get_text(strip=True)
                if not header or len(header) > 60:
                    continue
                parts = []
                for sib in tag.find_next_siblings():
                    if sib.name in ("h2", "h3", "h4", "h5"):
                        break
                    t = sib.get_text(separator="\n", strip=True)
                    if t:
                        parts.append(t)
                if parts:
                    sections[header.lower()] = "\n".join(parts)

        return sections

    def _extract_html_text(self, element):
        lines = []

        def process(node):
            if not getattr(node, "name", None):
                return
            if node.name == "ol":
                for i, li in enumerate(node.find_all("li", recursive=False), 1):
                    lines.append(f"{i}. {li.get_text(separator=' ', strip=True)}")
            elif node.name == "ul":
                for li in node.find_all("li", recursive=False):
                    lines.append("• " + li.get_text(separator=" ", strip=True))
            elif node.name in ("p", "div", "span"):
                if node.find(["ol", "ul"]):
                    for child in node.children:
                        process(child)
                else:
                    t = node.get_text(separator=" ", strip=True)
                    if t:
                        lines.append(t)
            elif node.name in ("h2", "h3", "h4", "h5", "h6", "strong", "b"):
                t = node.get_text(strip=True)
                if t:
                    lines.append(t + ":")
            elif hasattr(node, "children"):
                for child in node.children:
                    process(child)

        root = element.find(class_="metafield-rich_text_field") or element
        for child in root.children:
            process(child)

        result = "\n".join(lines).strip()
        return result if result else element.get_text(separator="\n", strip=True)

    def _get_sections(self, url):
        """Return (sections_dict, method_label). Tries HTML first for Shopify, Jina otherwise."""
        if self.is_shopify:
            try:
                sections = self._parse_html_sections(self._fetch_html(url))
                if any(sections.values()):
                    return sections, "HTML"
            except Exception:
                pass  # fall through to Jina

        markdown = self._fetch_via_jina(url)
        return self._parse_sections_from_markdown(markdown), "Jina"

    def _match_section(self, target, parsed):
        t = target.lower().strip()
        if t in parsed:
            return parsed[t]
        best_score, best_key = 0, None
        for key in parsed:
            score = self._fuzzy_score(t, key)
            if t in key or key in t:
                score = max(score, 0.8)
            if score > best_score:
                best_score, best_key = score, key
        return parsed[best_key] if best_key and best_score >= 0.6 else ""

    # ── Main scrape ───────────────────────────────────────────────────────────

    def scrape_product(self, product_name, upc=None, direct_url=None):
        result = {
            "input_name":    product_name,
            "matched_title": "",
            "product_url":   "",
            "status":        "",
            "note":          "",
        }
        for s in self.config["sections"]:
            result[s] = ""

        url, matched_title, confidence, reason = self.find_product(
            product_name, upc=upc, direct_url=direct_url
        )

        if not url:
            result["status"] = "NOT FOUND"
            result["note"]   = reason
            return result

        result["matched_title"] = matched_title or product_name
        result["product_url"]   = url

        if confidence == "LOW":
            result["status"] = "LOW CONFIDENCE"
            result["note"]   = reason
            return result

        try:
            sections, method = self._get_sections(url)
        except Exception as e:
            result["status"] = "ERROR"
            result["note"]   = str(e)
            return result

        for s in self.config["sections"]:
            result[s] = self._match_section(s, sections)

        result["status"] = "OK"
        result["note"]   = f"{reason} · via {method}"
        return result


# ─────────────────────────────────────────────────────────────────────────────
# INPUT PARSING
# ─────────────────────────────────────────────────────────────────────────────

def parse_input(file_obj, filename):
    """
    Parse an uploaded file (file-like object) into a list of (name, upc, url).
    Supports .xlsx, .xls, .csv.
    """
    ext = filename.rsplit(".", 1)[-1].lower()
    if ext in ("xlsx", "xls"):
        df = pd.read_excel(file_obj, header=None, dtype=str)
    elif ext == "csv":
        df = pd.read_csv(file_obj, header=None, dtype=str)
    else:
        raise ValueError(f"Unsupported file type: .{ext}")

    first = str(df.iloc[0, 0]).lower().strip()
    if any(kw in first for kw in ("product","name","title","sku","upc","barcode","item")):
        df = df.iloc[1:].reset_index(drop=True)

    UPC_RE = re.compile(r"\(UPC[:\s]+(\d+)\)", re.IGNORECASE)
    rows = []
    for _, row in df.iterrows():
        name = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        upc  = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else None
        url  = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else None

        if upc:
            upc = upc.split(".")[0]
            if upc.lower() in ("nan", "", "none"):
                upc = None
        if url and url.lower() in ("nan", "", "none"):
            url = None

        if name:
            m = UPC_RE.search(name)
            if m:
                if not upc:
                    upc = m.group(1)
                name = UPC_RE.sub("", name).strip(" ,|–-")

        if name and name.lower() not in ("nan", "", "none"):
            rows.append((name, upc, url))

    return rows


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL BUILDER  (returns BytesIO — no disk writes)
# ─────────────────────────────────────────────────────────────────────────────

def build_excel(results, sections):
    wb = Workbook()
    ws = wb.active
    ws.title = "Product Features"

    hdr_font  = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    hdr_fill  = PatternFill("solid", fgColor="1F3864")
    meta_fill = PatternFill("solid", fgColor="D9E1F2")
    ok_fill   = PatternFill("solid", fgColor="E2EFDA")
    warn_fill = PatternFill("solid", fgColor="FFF2CC")
    err_fill  = PatternFill("solid", fgColor="FCE4D6")
    c_align   = Alignment(horizontal="center", vertical="top", wrap_text=True)
    l_align   = Alignment(horizontal="left",   vertical="top", wrap_text=True)
    border    = Border(*[Side(style="thin")] * 0,
                       left=Side(style="thin"), right=Side(style="thin"),
                       top=Side(style="thin"),  bottom=Side(style="thin"))

    columns = ["Input Name","Matched Title","Product URL","Status","Note"] + sections
    for ci, col in enumerate(columns, 1):
        c = ws.cell(1, ci, col)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = c_align; c.border = border

    for ri, r in enumerate(results, 2):
        row = [
            r.get("input_name",""), r.get("matched_title",""),
            r.get("product_url",""), r.get("status",""), r.get("note",""),
        ] + [r.get(s,"") for s in sections]
        for ci, val in enumerate(row, 1):
            c = ws.cell(ri, ci, val)
            c.alignment = l_align; c.border = border
            if ci == 4:
                if val == "OK":
                    c.fill = ok_fill;   c.font = Font(name="Calibri", bold=True, color="375623")
                elif val == "LOW CONFIDENCE":
                    c.fill = warn_fill; c.font = Font(name="Calibri", bold=True, color="7F6000")
                elif val not in ("", None):
                    c.fill = err_fill;  c.font = Font(name="Calibri", bold=True, color="9C0006")
            elif ci <= 5:
                c.fill = meta_fill

    widths = {1:35, 2:35, 3:55, 4:14, 5:55}
    for i in range(len(sections)):
        widths[6+i] = 55
    for ci, w in widths.items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 22
    for ri in range(2, len(results)+2):
        ws.row_dimensions[ri].height = 80
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
