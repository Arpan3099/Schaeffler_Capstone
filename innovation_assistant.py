import streamlit as st
import anthropic
import json
import plotly.graph_objects as go
import time
import requests
from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment as XlAlign

# Load API key — Streamlit secrets take priority, fallback to hardcoded for local dev

def _mval(field):
    """Get display string from market field — handles both new dict and legacy string format."""
    if isinstance(field, dict):
        return field.get("value", "N/A")
    return str(field) if field else "N/A"

TAVILY_KEY = ""  # optional

st.set_page_config(page_title="Schaeffler Innovation Assistant", page_icon="🟢", layout="centered", initial_sidebar_state="expanded")

# ── API key — loaded from Streamlit secrets ───────────────────────────────────
try:
    ANTHROPIC_KEY = st.secrets["ANTHROPIC_API_KEY"]
    if not ANTHROPIC_KEY or not ANTHROPIC_KEY.startswith("sk-"):
        raise ValueError("Key looks invalid")
except Exception:
    st.error(
        "⚠️ **Anthropic API key not found or invalid.**\n\n"
        "Add your key to Streamlit secrets:\n"
        "1. Open your app on Streamlit Cloud → **Settings → Secrets**\n"
        "2. Add: `ANTHROPIC_API_KEY = \"sk-ant-...your-key...\"`\n"
        "3. Save and reboot the app.\n\n"
        "For local dev, create `.streamlit/secrets.toml` with the same line."
    )
    st.stop()

# ── Google Sheets — Idea Log ──────────────────────────────────────────────────
SHEET_ID = "1Ya-z55BtzRS7NYiKiM8U8E0-NChueTVprJovvUrvZ6s"

_SA_INFO = {
    "type": "service_account",
    "project_id": "absolute-dahlia-450007-r8",
    "private_key_id": "292e1cb7c9d249e34ad3f89ab2f180e865ea1576",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQCbuv2KGVKo3RN+\nY2vvVAsjGRBFxu8j8p8tGYxrvaxc0g0wefRZq0hOJbh9zBZu+dL4JwWf5SAJQ+Yp\nLdqctTIl+M6vBJNL1J8iNZxtmOmKrQeGvI+O+CnkU/4ZVMLk+lhOUrSVyJ8rPXOl\nT4E5ar8j+ZpWtueKtODgdKVgYHpTX4kjYj+/ochlFG6DQfhSWmJDf52Gfd/nldZt\n13+n3/uHYmLI1wKNFEpVaq3Qd2r191bcVAY8/NrOtRaqtHsoINAP2Sn6rEuGthaa\niZECrIkYiMKkjTWk8E8XIv9Zv2JnBOWPErIHRbGGHV6WBz5p0Z0L1GLGDiuhvha8\nvq6AiVsXAgMBAAECggEAHCGClDQv1NYeo9mU0UY2vs/Tuy8M2ssEivKPBZVdMeU1\nwbh99ca1iHxS39KCiOhy/iWaZABRMatEw9KHJ4Cpvuc7eq0SaIPPfS//AmM5aLYJ\n4oJkUlisxJSRlYTUseUxF3DkMxxq+DYhEk8S0krgnUCE6z4eBFXZO2KGzyqOXknf\nCS0ThZ2ky0/ytuIoj8OeCkanuBQZXj219zH0Kiamqt7AWs7qSUoafnejkkQGJ5xR\nKrYr+Y0D4L9LFmZEbidlfoHZOcz7vrHvkx3RrL+RA2QJU7XolzR6sl7ecwzLGifs\nvsCxzD52+7pOvwNUp/hHpol/2bD8W0xkDcRpqCFrKQKBgQDKJ6V12ElUixKz9djx\n3X8SLpQTJihAqIts5gesaNp1mAG7/3nAq1cThprwKaYPukkPYbwGjN/IPXwZKfZr\nfyHxFBA/a5ooz2bJoEm4w4w8dIrqxbDd9K2HuqKSnW0ouU/Hustv7gE+zkyGFRIx\nJ08hJwQaMuAfulzt/7vhbOvjDwKBgQDFNcvjltsm8gw80jjCxh3+6r+UDFPncVDL\niL5P/q6pfCzulsdxFVZUfx9yh9mCrxGopbdZnmzQW86GuhfnPC+AQpYdltGTOUdx\n+calCeSAmtfoztI3P+CmR4DDiZ4c2rfPlqOwwHueJXo62u8984y7NsGBdCRCrBSG\njVzFWFBneQKBgHTF/BUbsBhPEam0rPHh0cJN96ksFHptIcTxB6O3GeJtwSq4w7rg\n/ra/vYZXeJ6DLCrfeP6Lp8UCh0n97GNiF9grj8siu/UxAR4dIhjBlKNjas99DNLZ\nwNeznq90kpbAnO4x38wzPrLp9lhJma2dGF99Kyh7FO4e+Ale/UeVZJlPAoGAB8N6\nZ1dFAV9+A9byzRgnjiWHrThfBTl8yMZ1V4jbL2joC+x7pYQFhgYLIuMeOPrTYyRC\n95A5EGrM0pj43+2KoS394uRRE86pdV8z5sNg738pCM07kVk+as1d0FTWmKQzoER5\n5TdupmcrTK3ZxUKVQ7mAHKyJ0OYdWL6v7ETxxWECgYA2RXMn67Pb0CtfX1Gs7ztB\n8N3PMOtYdz/TQbOK0u4waWrzbnFUCvsH39yoqiDofYdUDN/E/31IXoTzCi3N4F2U\ncvJ3oid6ZubsBA2gJrHTYvd/yljV+o8kuzUV4d2/fxTuphhh2TyoYa3BUD+W7alr\nDs4l7RfxIUIgMviRmVpAdw==\n-----END PRIVATE KEY-----\n",
    "client_email": "schaeffler-capstone@absolute-dahlia-450007-r8.iam.gserviceaccount.com",
    "client_id": "105367783720778854419",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/schaeffler-capstone%40absolute-dahlia-450007-r8.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}

SHEET_COLUMNS = [
    "Date", "Submitter Name", "Position", "Department",
    "Full Idea Description", "Clarifying Q&A",
    "Quadrant", "Innovation Cluster", "Product Family",
    "Market Score", "Patent Score", "Feasibility Score", "Org Readiness Score", "IPI Score",
    "Recommendation", "Key Concerns", "Next Steps",
    "Market Name", "Market Size 2024", "CAGR",
    "TRL Level", "Build Strategy"
]

def _sheets_client():
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(_SA_INFO, scopes=scopes)
    return gspread.authorize(creds)

def load_past_ideas():
    try:
        ws = _sheets_client().open_by_key(SHEET_ID).sheet1
        return ws.get_all_records()
    except Exception:
        return []

def save_idea_to_sheets(row_data: dict):
    try:
        sh = _sheets_client().open_by_key(SHEET_ID)
        ws = sh.sheet1
        existing = ws.get_all_values()
        if not existing:
            ws.append_row(SHEET_COLUMNS)
        row = [str(row_data.get(col, "")) for col in SHEET_COLUMNS]
        ws.append_row(row)
        return True
    except Exception as e:
        return False

def check_similar_ideas(new_idea, past_ideas):
    if not past_ideas:
        return []
    past_summaries = "\n".join([
        f"- [{r.get('Date','')}] {r.get('Submitter Name','Unknown')} ({r.get('Department','')}): {str(r.get('Full Idea Description',''))[:200]}"
        for r in past_ideas[:30]
    ])
    result = call_claude(
        'You compare innovation ideas. Return ONLY valid JSON: {"similar": [{"date": "...", "submitter": "...", "department": "...", "idea_snippet": "...", "quadrant": "...", "ipi": "...", "recommendation": "...", "similarity": "High/Medium", "reason": "one sentence"}]}. Return empty similar array if nothing is genuinely similar.',
        f"New idea: {new_idea}\n\nPast researched ideas:\n{past_summaries}",
        max_tokens=600
    )
    try:
        raw = result.strip().replace("```json","").replace("```","").strip()
        fb = raw.find("{"); lb = raw.rfind("}") + 1
        if fb >= 0: raw = raw[fb:lb]
        return json.loads(raw).get("similar", [])
    except Exception:
        return []

def generate_ideas_excel():
    records = load_past_ideas()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Innovation Ideas"
    hdr_fill  = PatternFill("solid", fgColor="1F3864")
    hdr_font  = Font(bold=True, color="FFFFFF", size=10)
    alt_fill  = PatternFill("solid", fgColor="EAF1FB")
    if not records:
        ws.append(["No ideas recorded yet — complete a full pipeline assessment to populate this log."])
        ws["A1"].font = Font(italic=True, color="555555")
    else:
        headers = list(records[0].keys())
        ws.append(headers)
        for cell in ws[1]:
            cell.font = hdr_font
            cell.fill = hdr_fill
            cell.alignment = XlAlign(horizontal="center", wrap_text=True)
        for i, record in enumerate(records, start=2):
            ws.append([record.get(h, "") for h in headers])
            if i % 2 == 0:
                for cell in ws[i]:
                    cell.fill = alt_fill
        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Styling — sidebar identity + sidebar-toggle arrow fix ────
st.markdown("""
<style>
/* ── Sidebar: Schaeffler green identity ── */
section[data-testid="stSidebar"] {
    background-color: #007A3D !important;
    border-right: none !important;
}
section[data-testid="stSidebar"] * {
    color: #FFFFFF !important;
}
section[data-testid="stSidebar"] .stButton > button {
    background: transparent !important;
    color: rgba(255,255,255,0.75) !important;
    border: none !important;
    border-radius: 4px !important;
    font-size: 12px !important;
    font-family: 'Arial','Helvetica Neue',Helvetica,sans-serif !important;
    font-weight: 400 !important;
    letter-spacing: 0.3px !important;
    padding: 8px 12px !important;
    text-align: left !important;
}
section[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(255,255,255,0.12) !important;
    color: #FFFFFF !important;
}

/* ── Sidebar collapse/expand toggle — keep it functional ── */
[data-testid="collapsedControl"],
[data-testid="stSidebarCollapsedControl"],
div[class*="collapsedControl"] {
    display: flex !important;
    visibility: visible !important;
    pointer-events: auto !important;
}
[data-testid="collapsedControl"] button,
[data-testid="stSidebarCollapsedControl"] button,
button[data-testid="baseButton-headerNoPadding"] {
    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
    visibility: visible !important;
    pointer-events: auto !important;
    font-size: 0 !important; /* hide broken icon ligature text fallback */
}
[data-testid="collapsedControl"] button span,
[data-testid="collapsedControl"] button p,
[data-testid="stSidebarCollapsedControl"] button span,
[data-testid="stSidebarCollapsedControl"] button p,
button[data-testid="baseButton-headerNoPadding"] span,
button[data-testid="baseButton-headerNoPadding"] p {
    font-size: 0 !important;
    color: transparent !important;
}
[data-testid="collapsedControl"] button::before {
    content: "◀";
    font-size: 14px !important;
    line-height: 1 !important;
    color: rgba(255,255,255,0.9) !important;
}
[data-testid="stSidebarCollapsedControl"] button::before {
    content: "▶";
    font-size: 14px !important;
    line-height: 1 !important;
    color: rgba(255,255,255,0.9) !important;
}

/* ── Fix expander arrow rendering as icon-name text ── */
/* Same root cause: Material Icons ligature not loading → hide the icon
   span entirely and use a CSS triangle that is font-independent. */
[data-testid="stExpander"] summary {
    position: relative !important;
    padding-left: 28px !important;
    list-style: none !important;
}
[data-testid="stExpander"] summary::-webkit-details-marker { display: none !important; }
/* Hide every potential icon element — SVG, span, p inside summary */
[data-testid="stExpander"] summary > svg,
[data-testid="stExpander"] summary > span[data-testid],
[data-testid="stExpander"] summary > div > svg {
    display: none !important;
}
/* Zero out any orphan text node that is a Material icon name */
[data-testid="stExpander"] summary > span:not([class]),
[data-testid="stExpander"] summary > p {
    font-size: 0 !important;
    line-height: 0 !important;
}
/* Inject a CSS-only chevron arrow — totally font-independent */
[data-testid="stExpander"] summary::before {
    content: "" !important;
    position: absolute !important;
    left: 8px !important;
    top: 50% !important;
    transform: translateY(-50%) rotate(-90deg) !important;
    width: 0 !important;
    height: 0 !important;
    border-left: 5px solid transparent !important;
    border-right: 5px solid transparent !important;
    border-top: 7px solid #60a5fa !important;
    transition: transform 0.2s ease !important;
    display: block !important;
    visibility: visible !important;
}
details[open] [data-testid="stExpander"] summary::before {
    transform: translateY(-50%) rotate(0deg) !important;
}

/* ── Sidebar selectbox — match sidebar style ── */
section[data-testid="stSidebar"] .stSelectbox label {
    color: rgba(255,255,255,0.55) !important;
    font-size: 10px !important;
    letter-spacing: 1.5px !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
}
section[data-testid="stSidebar"] .stSelectbox > div > div {
    background: transparent !important;
    border: 1px solid rgba(255,255,255,0.2) !important;
    color: rgba(255,255,255,0.8) !important;
    font-size: 11px !important;
    font-family: 'Arial','Helvetica Neue',Helvetica,sans-serif !important;
    border-radius: 4px !important;
}
section[data-testid="stSidebar"] .stSelectbox > div > div:hover {
    border-color: rgba(255,255,255,0.45) !important;
}
section[data-testid="stSidebar"] .stSelectbox svg {
    fill: rgba(255,255,255,0.6) !important;
    color: rgba(255,255,255,0.6) !important;
}

/* ── Language + Ideas Log fixed at bottom of sidebar ── */
.ideas-log-fixed {
    position: fixed !important;
    bottom: 0 !important;
    left: 0 !important;
    width: var(--sidebar-width, 246px) !important;
    max-width: 260px !important;
    background: #007A3D !important;
    border-top: 1px solid rgba(255,255,255,0.18) !important;
    padding: 10px 14px 16px 14px !important;
    z-index: 9999 !important;
    box-sizing: border-box !important;
}
.ideas-log-fixed .stButton > button,
.ideas-log-fixed .stDownloadButton > button {
    background: transparent !important;
    border: none !important;
    color: rgba(255,255,255,0.5) !important;
    font-size: 10px !important;
    font-family: 'Arial','Helvetica Neue',Helvetica,sans-serif !important;
    font-weight: 400 !important;
    letter-spacing: 0.3px !important;
    padding: 4px 0 !important;
    text-align: left !important;
    width: 100% !important;
}
.ideas-log-fixed .stButton > button:hover,
.ideas-log-fixed .stDownloadButton > button:hover {
    color: rgba(255,255,255,0.85) !important;
    background: transparent !important;
}
.ideas-log-fixed .stSelectbox label {
    color: rgba(255,255,255,0.45) !important;
    font-size: 9px !important;
    letter-spacing: 1.5px !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    margin-bottom: 2px !important;
}
.ideas-log-fixed .stSelectbox > div > div {
    background: transparent !important;
    border: 1px solid rgba(255,255,255,0.18) !important;
    color: rgba(255,255,255,0.65) !important;
    font-size: 10px !important;
    font-family: 'Arial','Helvetica Neue',Helvetica,sans-serif !important;
    border-radius: 3px !important;
    padding: 2px 6px !important;
}
.ideas-log-fixed .stSelectbox svg {
    fill: rgba(255,255,255,0.45) !important;
}

/* ── Fix expander arrow overlapping label text ── */
[data-testid="stExpander"] summary {
    padding-left: 28px !important;
}
[data-testid="stExpander"] summary svg {
    left: 6px !important;
    position: absolute !important;
    flex-shrink: 0 !important;
}
[data-testid="stExpander"] summary p,
[data-testid="stExpander"] summary span {
    margin-left: 4px !important;
}

/* ── Source tag pill ── */
.source-tag {
    background:#1e3a5f; color:#93c5fd;
    font-size:11px; padding:2px 8px;
    border-radius:3px; margin-left:6px;
    font-family: monospace;
}

/* ── Global font consistency — Arial/Helvetica Neue throughout ── */
html, body, [class*="css"],
.stApp, .stMarkdown, .stText,
.stButton > button,
.stTextInput > div > div > input,
.stTextArea > div > div > textarea,
.stSelectbox > div > div,
.stMetric, .stChatMessage,
.stCaption, .stAlert,
h1, h2, h3, h4, h5, h6, p, div, span, label, li {
    font-family: 'Arial', 'Helvetica Neue', Helvetica, sans-serif !important;
}
/* Keep monospace only for code blocks */
code, pre, .stCode {
    font-family: 'Courier New', Courier, monospace !important;
}
</style>

<script>
// Inject favicon dynamically as an SVG data URI with Schaeffler S logo
(function() {
    var svgFavicon = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect width='32' height='32' rx='5' fill='%23007A3D'/%3E%3Ctext x='16' y='24' font-family='Arial,Helvetica,sans-serif' font-size='22' font-weight='700' fill='white' text-anchor='middle'%3ES%3C/text%3E%3C/svg%3E";
    var existing = document.querySelector("link[rel*='icon']");
    if (existing) {
        existing.href = svgFavicon;
    } else {
        var link = document.createElement('link');
        link.rel = 'icon';
        link.type = 'image/svg+xml';
        link.href = svgFavicon;
        document.head.appendChild(link);
    }
    var shortcut = document.querySelector("link[rel='shortcut icon']");
    if (shortcut) shortcut.href = svgFavicon;
})();

// ── Keep Streamlit sidebar toggle clickable + icon-safe ───────────
(function fixSidebarToggleControl() {
    var SELECTORS = [
        '[data-testid="collapsedControl"] button',
        '[data-testid="stSidebarCollapsedControl"] button',
        'button[data-testid="baseButton-headerNoPadding"]'
    ];

    function setImportantStyle(el, prop, value) {
        if (!el || !el.style) return;
        el.style.setProperty(prop, value, "important");
    }

    function patchOneButton(btn) {
        if (!btn) return;

        // Force functional click target
        setImportantStyle(btn, "display", "inline-flex");
        setImportantStyle(btn, "visibility", "visible");
        setImportantStyle(btn, "pointer-events", "auto");
        setImportantStyle(btn, "align-items", "center");
        setImportantStyle(btn, "justify-content", "center");
        setImportantStyle(btn, "font-size", "0");
        setImportantStyle(btn, "color", "transparent");

        // Hide broken ligature fallback text (e.g., keyboard_double_arrow_left)
        btn.querySelectorAll("span, p").forEach(function(node) {
            setImportantStyle(node, "font-size", "0");
            setImportantStyle(node, "color", "transparent");
            setImportantStyle(node, "line-height", "0");
        });
    }

    function patchAll() {
        SELECTORS.forEach(function(sel) {
            document.querySelectorAll(sel).forEach(patchOneButton);
        });
    }

    patchAll();
    new MutationObserver(patchAll).observe(document.body, { childList: true, subtree: true });
})();

</script>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────
BG    = "#0f1e35"
BLUE  = "#60a5fa"
DIM   = "#4a6fa5"
WHITE = "#e2e8f0"
NAVY  = "#1F3864"

# ── Language system ────────────────────────────────────────────
_LANG = {
    "en": {
        "pipeline":  "INNOVATION PIPELINE",
        "idea_cap":  "Idea",
        "quad_cap":  "Quadrant",
        "lang_label":"Language / Sprache",
        "stages": [
            (1,"01 · Quadrant Classifier"),
            (2,"02 · Market Intelligence"),
            (3,"03 · Patent Intelligence"),
            (4,"04 · Technical Feasibility"),
            (5,"05 · Organisational Readiness"),
            (6,"06 · Scoring & Synthesis"),
        ],
        "dl_market":  "⬇️  Download Market Intelligence Report",
        "dl_patent":  "⬇️  Download Patent Intelligence Report",
        "dl_feasib":  "⬇️  Download Technical Feasibility Report",
        "dl_org":     "⬇️  Download Organisational Readiness Report",
        "dl_master":  "⬇️  Download Full Innovation Assessment Report",
        "dl_ideas":   "↓  Download Ideas Log",
        "dl_spinner": "Generating report — please wait…",
        "dl_caption": "Covers all 5 stages: Market · Patent · Feasibility · Org Readiness · IPI",
        "claude_suffix": "",
        # Stage 01
        "s1_title": "Stage 01 · Quadrant Classifier",
        "s1_what": "Maps your idea onto Schaeffler's modified Ansoff matrix — Exploit, Extend, Radical, or Disrupt. Ideas in Radical and Disrupt proceed through the full pipeline. Others are redirected to the right Schaeffler product division.",
        "s1_what_label": "WHAT THIS STAGE DOES",
        "s1_you_get": "<b style='color:#e2e8f0;'>You get:</b> Quadrant classification · Schaeffler Motion product family fit · Strategic trend alignment · Innovation pathway (Start-Up Mode vs Innovation Factory)",
        "s1_step1": "Step 1 — Who are you & describe your idea",
        "s1_name": "Your name *", "s1_name_ph": "e.g. Anna Müller",
        "s1_role": "Position / Role *", "s1_role_ph": "e.g. Senior Engineer",
        "s1_dept": "Department *", "s1_dept_ph": "e.g. E-Mobility R&D",
        "s1_idea_label": "What is your innovation idea?",
        "s1_idea_ph": "e.g. A self-lubricating bearing system that uses micro-reservoirs embedded within the bearing material to release lubricant automatically based on temperature and load sensing...",
        "s1_submit": "Submit idea",
        "s1_warn_identity": "Please fill in your name, position, and department before continuing.",
        "s1_warn_empty": "Please enter your idea first.",
        "s1_warn_brief": "A bit brief — can you add more detail? What does it do, and for whom?",
        "s1_step2": "Step 2 — Three quick questions",
        "s1_step2_caption": "These answers help refine the classification. The idea description itself drives the result — use these to add context, not to override what the idea clearly is.",
        "s1_q1": "**1. Has the core technology behind this idea been demonstrated anywhere — in a lab, a startup, a research paper, or a competitor product?**",
        "s1_q1a": "Yes — it has been demonstrated somewhere (even if not commercialised)",
        "s1_q1b": "No — the underlying technology is genuinely novel or theoretical",
        "s1_q1_detail": "Optional — where has it been demonstrated, or what makes it genuinely new?",
        "s1_q1_ph": "e.g. MIT lab prototype, or 'no known demonstration of this mechanism'",
        "s1_q2": "**2. Does this idea target markets or applications that Schaeffler currently operates in?**",
        "s1_q2_caption": "Schaeffler's current markets: automotive (ICE & EV), industrial machinery, rail, aerospace, energy, two-wheelers. If the idea fits any of these, select yes.",
        "s1_q2a": "Yes — it targets automotive, industrial, rail, aerospace, energy or adjacent sectors Schaeffler already serves",
        "s1_q2b": "No — it targets a market genuinely outside Schaeffler's current scope (e.g. consumer electronics, healthcare devices, retail)",
        "s1_q2_detail": "Optional — which specific market or application area?",
        "s1_q2_ph": "e.g. EV drivetrain OEMs, or 'medical implant manufacturers — entirely new for Schaeffler'",
        "s1_q3": "**3. Is the problem this idea solves already recognised and being worked on by the industry?**",
        "s1_q3a": "Yes — the problem is well known and others are actively trying to solve it",
        "s1_q3b": "No — the problem itself is new, underappreciated, or not yet widely recognised",
        "s1_q3_detail": "Optional — describe the problem in one sentence",
        "s1_q3_ph": "e.g. Bearing failure in EV drivetrains due to high-frequency current leakage",
        "s1_classify_btn": "Classify my idea →",
        "s1_result": "Result",
        "s1_qualifies": "✓ This idea qualifies for the full Innovation pipeline.",
        "s1_continue": "Continue to Stage 02: Market Intelligence →",
        "s1_full_run": "⚡ Run Full Analysis — all 5 stages",
        "s1_chat_header": "💬 Questions about this classification?",
        "s1_chat_ph": "Ask about the classification...",
        "s1_startover": "← Start over",
        "s1_spinner_check": "Checking your idea...",
        "s1_spinner_similar": "Checking against past ideas...",
        "s1_spinner_classify": "Classifying your idea...",
        "s1_similar_warn": "⚠️ **{n} similar idea(s) previously researched** — review before proceeding.",
        # Stage 02
        "s2_title": "Stage 02 · Market Intelligence",
        "s2_what": "Analyses the commercial opportunity behind your idea — how big the market is, how fast it is growing, who the competitors are, and how well the idea fits across Schaeffler's 10 customer sector clusters.",
        "s2_you_get": "<b style='color:#e2e8f0;'>You get:</b> Market size & CAGR with sources · Sector cluster fit chart · Competitor landscape · Market Intelligence Score (0–10)",
        "s2_run_btn": "Run Market Intelligence →",
        "s2_continue": "Continue to Stage 03: Patent Intelligence →",
        "s2_rerun": "← Re-run analysis",
        "s2_success": "✓ Market Intelligence complete. Final score: **{score}/10**",
        "s2_chat_header": "💬 Questions about the market analysis?",
        "s2_chat_ph": "Ask about the market, sectors, competitors...",
        # Stage 03
        "s3_title": "Stage 03 · Patent Intelligence",
        "s3_what": "Maps the patent landscape for your idea's core technology — who is filing, whether they are competitors or potential customers, where the IP white spaces are, and how Schaeffler's existing patent portfolio relates to the idea.",
        "s3_you_get": "<b style='color:#e2e8f0;'>You get:</b> Patent Ansoff map with all key filers plotted · IP white spaces · Schaeffler IP gap analysis · Patent Intelligence Score (0–10)",
        "s3_run_btn": "Run Patent Intelligence →",
        "s3_continue": "Continue to Stage 04: Technical Feasibility →",
        "s3_rerun": "← Re-run analysis",
        "s3_success": "✓ Patent Intelligence complete. Final score: **{score}/10**",
        "s3_chat_header": "💬 Questions about the patent analysis?",
        "s3_chat_ph": "Ask about patents, IP risks, white spaces...",
        # Stage 04
        "s4_title": "Stage 04 · Technical Feasibility",
        "s4_what": "Assesses whether the core technology of your idea actually exists and how mature it is — using the Technology Readiness Level (TRL) framework used by NASA, the EU, and industrial R&D organisations.",
        "s4_you_get": "<b style='color:#e2e8f0;'>You get:</b> Technology existence verdict · TRL level (1–9) with rationale · Schaeffler entry readiness · Technical Feasibility Score (0–10)",
        "s4_run_btn": "Run Technical Feasibility →",
        "s4_continue": "Continue to Stage 05: Organisational Readiness →",
        "s4_rerun": "← Re-run analysis",
        "s4_success": "✓ Technical Feasibility complete. Final score: **{score}/10**",
        "s4_chat_header": "💬 Questions about the feasibility analysis?",
        "s4_chat_ph": "Ask about TRL, technology gaps, development timeline...",
        # Stage 05
        "s5_title": "Stage 05 · Organisational Readiness",
        "s5_what": "Assesses whether Schaeffler has the organisational capability to develop and commercialise this idea — using the P³ formula: Performance = Portfolio × People × Process.",
        "s5_you_get": "<b style='color:#e2e8f0;'>You get:</b> P³ readiness scores · Build vs Partner recommendation · Key capability gaps · Organisational Readiness Score (0–10)",
        "s5_run_btn": "Run Organisational Readiness →",
        "s5_continue": "Continue to Stage 06: Scoring & Synthesis →",
        "s5_rerun": "← Re-run analysis",
        "s5_success": "✓ Organisational Readiness complete. Final score: **{score}/10**",
        "s5_chat_header": "💬 Questions about the org readiness analysis?",
        "s5_chat_ph": "Ask about P³ scores, capability gaps, partnerships...",
        # Stage 06
        "s6_title": "Stage 06 · Scoring & Synthesis",
        "s6_what": "Synthesises all five pipeline dimensions into a single Innovation Potential Index (IPI) score and strategic recommendation — with full narrative, risk analysis, and concrete next steps.",
        "s6_you_get": "<b style='color:#e2e8f0;'>You get:</b> IPI score (0–10) · Strategic recommendation · Radar chart · Key concerns & next steps · Full narrative synthesis",
        "s6_run_btn": "Run Scoring & Synthesis →",
        "s6_rerun": "← Adjust weights and re-run",
        "s6_chat_header": "💬 Questions about the overall assessment?",
        "s6_chat_ph": "Ask about the IPI score, recommendation, or next steps...",
        "s6_image_btn": "🖼️ Generate Solution Image",
        "s6_image_redo": "🔄 Generate different image",
    },
    "de": {
        "pipeline":  "INNOVATIONS-PIPELINE",
        "idea_cap":  "Idee",
        "quad_cap":  "Quadrant",
        "lang_label":"Sprache",
        "stages": [
            (1,"01 · Quadrant-Klassifikator"),
            (2,"02 · Marktintelligenz"),
            (3,"03 · Patentintelligenz"),
            (4,"04 · Technische Machbarkeit"),
            (5,"05 · Organisatorische Bereitschaft"),
            (6,"06 · Bewertung & Synthese"),
        ],
        "dl_market":  "⬇️  Marktintelligenz-Bericht herunterladen",
        "dl_patent":  "⬇️  Patentintelligenz-Bericht herunterladen",
        "dl_feasib":  "⬇️  Technischen Machbarkeitsbericht herunterladen",
        "dl_org":     "⬇️  Org. Bereitschaftsbericht herunterladen",
        "dl_master":  "⬇️  Vollständigen Innovationsbericht herunterladen",
        "dl_ideas":   "↓  Ideen-Log herunterladen",
        "dl_spinner": "Bericht wird erstellt — bitte warten…",
        "dl_caption": "Umfasst alle 5 Stufen: Markt · Patente · Machbarkeit · Org. Bereitschaft · IPI",
        # Stage 01
        "s1_title": "Stufe 01 · Quadrant-Klassifikator",
        "s1_what": "Ordnet Ihre Idee in die modifizierte Ansoff-Matrix von Schaeffler ein — Exploit, Extend, Radical oder Disrupt. Ideen in Radical und Disrupt durchlaufen die gesamte Pipeline. Andere werden an die passende Produktsparte weitergeleitet.",
        "s1_what_label": "WAS DIESE STUFE TUT",
        "s1_you_get": "<b style='color:#e2e8f0;'>Sie erhalten:</b> Quadrant-Klassifikation · Passung zur Schaeffler Motion-Produktfamilie · Strategische Trendausrichtung · Innovationspfad (Start-Up-Modus vs. Innovationsfabrik)",
        "s1_step1": "Schritt 1 — Wer sind Sie & beschreiben Sie Ihre Idee",
        "s1_name": "Ihr Name *", "s1_name_ph": "z.B. Anna Müller",
        "s1_role": "Position / Rolle *", "s1_role_ph": "z.B. Senior Engineer",
        "s1_dept": "Abteilung *", "s1_dept_ph": "z.B. E-Mobilität F&E",
        "s1_idea_label": "Was ist Ihre Innovationsidee?",
        "s1_idea_ph": "z.B. Ein selbstschmierendes Lagersystem, das Mikro-Reservoirs im Lagermaterial nutzt, um Schmiermittel automatisch basierend auf Temperatur- und Lastsensoren freizusetzen...",
        "s1_submit": "Idee einreichen",
        "s1_warn_identity": "Bitte füllen Sie Name, Position und Abteilung aus, bevor Sie fortfahren.",
        "s1_warn_empty": "Bitte geben Sie zuerst Ihre Idee ein.",
        "s1_warn_brief": "Etwas kurz — können Sie mehr Details hinzufügen? Was macht es, und für wen?",
        "s1_step2": "Schritt 2 — Drei kurze Fragen",
        "s1_step2_caption": "Diese Antworten helfen bei der Verfeinerung der Klassifikation. Die Ideenbeschreibung selbst bestimmt das Ergebnis — nutzen Sie diese nur zur Ergänzung.",
        "s1_q1": "**1. Wurde die Kerntechnologie dieser Idee irgendwo demonstriert — in einem Labor, einem Startup, einer Forschungsarbeit oder einem Konkurrenzprodukt?**",
        "s1_q1a": "Ja — sie wurde irgendwo demonstriert (auch wenn nicht kommerzialisiert)",
        "s1_q1b": "Nein — die zugrunde liegende Technologie ist wirklich neu oder theoretisch",
        "s1_q1_detail": "Optional — wo wurde sie demonstriert, oder was macht sie wirklich neu?",
        "s1_q1_ph": "z.B. MIT-Labor-Prototyp, oder 'kein bekannter Nachweis dieses Mechanismus'",
        "s1_q2": "**2. Zielt diese Idee auf Märkte oder Anwendungen, in denen Schaeffler derzeit tätig ist?**",
        "s1_q2_caption": "Aktuelle Märkte von Schaeffler: Automotive (Verbrenner & EV), Industriemaschinen, Bahn, Luft- und Raumfahrt, Energie, Zweiräder.",
        "s1_q2a": "Ja — sie zielt auf Automotive, Industrie, Bahn, Luft- und Raumfahrt, Energie oder angrenzende Sektoren",
        "s1_q2b": "Nein — sie zielt auf einen Markt außerhalb des aktuellen Schaeffler-Portfolios (z.B. Unterhaltungselektronik, Medizintechnik)",
        "s1_q2_detail": "Optional — welcher spezifische Markt oder Anwendungsbereich?",
        "s1_q2_ph": "z.B. EV-Antriebsstrang-OEMs, oder 'Medizinimplantat-Hersteller — völlig neu für Schaeffler'",
        "s1_q3": "**3. Wird das Problem, das diese Idee löst, bereits von der Branche erkannt und bearbeitet?**",
        "s1_q3a": "Ja — das Problem ist bekannt und andere versuchen es aktiv zu lösen",
        "s1_q3b": "Nein — das Problem selbst ist neu, unterschätzt oder noch nicht weit verbreitet",
        "s1_q3_detail": "Optional — beschreiben Sie das Problem in einem Satz",
        "s1_q3_ph": "z.B. Lagerversagen in EV-Antriebssträngen durch hochfrequente Stromableitung",
        "s1_classify_btn": "Meine Idee klassifizieren →",
        "s1_result": "Ergebnis",
        "s1_qualifies": "✓ Diese Idee qualifiziert sich für die vollständige Innovationspipeline.",
        "s1_continue": "Weiter zu Stufe 02: Marktintelligenz →",
        "s1_full_run": "⚡ Vollanalyse — alle 5 Stufen",
        "s1_chat_header": "💬 Fragen zur Klassifikation?",
        "s1_chat_ph": "Fragen zur Klassifikation stellen...",
        "s1_startover": "← Von vorne beginnen",
        "s1_spinner_check": "Idee wird geprüft...",
        "s1_spinner_similar": "Vergleich mit früheren Ideen...",
        "s1_spinner_classify": "Idee wird klassifiziert...",
        "s1_similar_warn": "⚠️ **{n} ähnliche Idee(n) wurden bereits untersucht** — bitte vor dem Fortfahren prüfen.",
        # Stage 02
        "s2_title": "Stufe 02 · Marktintelligenz",
        "s2_what": "Analysiert die kommerzielle Chance hinter Ihrer Idee — wie groß der Markt ist, wie schnell er wächst, wer die Wettbewerber sind, und wie gut die Idee zu Schaefflers 10 Kundensektorclustern passt.",
        "s2_you_get": "<b style='color:#e2e8f0;'>Sie erhalten:</b> Marktgröße & CAGR mit Quellen · Sektorcluster-Passungsdiagramm · Wettbewerbslandschaft · Marktintelligenz-Score (0–10)",
        "s2_run_btn": "Marktintelligenz starten →",
        "s2_continue": "Weiter zu Stufe 03: Patentintelligenz →",
        "s2_rerun": "← Analyse wiederholen",
        "s2_success": "✓ Marktintelligenz abgeschlossen. Endergebnis: **{score}/10**",
        "s2_chat_header": "💬 Fragen zur Marktanalyse?",
        "s2_chat_ph": "Fragen zu Markt, Sektoren, Wettbewerbern...",
        # Stage 03
        "s3_title": "Stufe 03 · Patentintelligenz",
        "s3_what": "Kartiert die Patentlandschaft für die Kerntechnologie Ihrer Idee — wer anmeldet, ob es Wettbewerber oder potenzielle Kunden sind, wo die IP-Weißräume liegen, und wie Schaefflers bestehendes Patentportfolio mit der Idee zusammenhängt.",
        "s3_you_get": "<b style='color:#e2e8f0;'>Sie erhalten:</b> Patent-Ansoff-Karte · IP-Weißräume · Schaeffler-IP-Lückenanalyse · Patentintelligenz-Score (0–10)",
        "s3_run_btn": "Patentintelligenz starten →",
        "s3_continue": "Weiter zu Stufe 04: Technische Machbarkeit →",
        "s3_rerun": "← Analyse wiederholen",
        "s3_success": "✓ Patentintelligenz abgeschlossen. Endergebnis: **{score}/10**",
        "s3_chat_header": "💬 Fragen zur Patentanalyse?",
        "s3_chat_ph": "Fragen zu Patenten, IP-Risiken, Weißräumen...",
        # Stage 04
        "s4_title": "Stufe 04 · Technische Machbarkeit",
        "s4_what": "Bewertet, ob die Kerntechnologie Ihrer Idee tatsächlich existiert und wie ausgereift sie ist — unter Verwendung des TRL-Rahmens (Technology Readiness Level) von NASA, EU und industriellen F&E-Organisationen.",
        "s4_you_get": "<b style='color:#e2e8f0;'>Sie erhalten:</b> Technologie-Existenzbewertung · TRL-Stufe (1–9) mit Begründung · Schaeffler-Eintrittsvorbereitung · Technische Machbarkeits-Score (0–10)",
        "s4_run_btn": "Technische Machbarkeit starten →",
        "s4_continue": "Weiter zu Stufe 05: Organisatorische Bereitschaft →",
        "s4_rerun": "← Analyse wiederholen",
        "s4_success": "✓ Technische Machbarkeit abgeschlossen. Endergebnis: **{score}/10**",
        "s4_chat_header": "💬 Fragen zur Machbarkeitsanalyse?",
        "s4_chat_ph": "Fragen zu TRL, Technologielücken, Entwicklungszeitplan...",
        # Stage 05
        "s5_title": "Stufe 05 · Organisatorische Bereitschaft",
        "s5_what": "Bewertet, ob Schaeffler die organisatorische Fähigkeit hat, diese Idee zu entwickeln und zu kommerzialisieren — mit der P³-Formel: Leistung = Portfolio × People × Process.",
        "s5_you_get": "<b style='color:#e2e8f0;'>Sie erhalten:</b> P³-Bereitschaftsscores · Build-vs-Partner-Empfehlung · Wichtige Kompetenzlücken · Organisatorischer Bereitschafts-Score (0–10)",
        "s5_run_btn": "Organisatorische Bereitschaft starten →",
        "s5_continue": "Weiter zu Stufe 06: Bewertung & Synthese →",
        "s5_rerun": "← Analyse wiederholen",
        "s5_success": "✓ Organisatorische Bereitschaft abgeschlossen. Endergebnis: **{score}/10**",
        "s5_chat_header": "💬 Fragen zur Org.-Bereitschaftsanalyse?",
        "s5_chat_ph": "Fragen zu P³-Scores, Kompetenzlücken, Partnerschaften...",
        # Stage 06
        "s6_title": "Stufe 06 · Bewertung & Synthese",
        "s6_what": "Synthetisiert alle fünf Pipeline-Dimensionen zu einem einzigen Innovationspotenzialindex (IPI) und einer strategischen Empfehlung — mit vollständiger Erzählung, Risikoanalyse und konkreten nächsten Schritten.",
        "s6_you_get": "<b style='color:#e2e8f0;'>Sie erhalten:</b> IPI-Score (0–10) · Strategische Empfehlung · Radardiagramm · Wichtige Bedenken & nächste Schritte · Vollständige narrative Synthese",
        "s6_run_btn": "Bewertung & Synthese starten →",
        "s6_rerun": "← Gewichtungen anpassen und neu starten",
        "s6_chat_header": "💬 Fragen zur Gesamtbewertung?",
        "s6_chat_ph": "Fragen zu IPI-Score, Empfehlung oder nächsten Schritten...",
        "s6_image_btn": "🖼️ Lösungsbild generieren",
        "s6_image_redo": "🔄 Anderes Bild generieren",
        "claude_suffix": (
            "\n\nWICHTIG: Antworte AUSSCHLIESSLICH auf Deutsch. "
            "Alle Analysen, Beschreibungen, Begründungen, Zusammenfassungen und sonstigen Texte "
            "müssen vollständig auf Deutsch verfasst sein — auch alle String-Werte in JSON-Antworten "
            "(z.B. market_name, rationale, focus, summary, growth_drivers usw.). "
            "JSON-Schlüssel bleiben auf Englisch. Eigennamen, Marken und internationale "
            "Fachbegriffe dürfen auf Englisch bleiben."
        ),
    },
}

def T(key):
    lang = st.session_state.get("ui_lang", "en")
    return _LANG.get(lang, _LANG["en"]).get(key, _LANG["en"].get(key, key))

def _lang_suffix():
    lang = st.session_state.get("ui_lang", "en")
    return _LANG.get(lang, _LANG["en"]).get("claude_suffix", "")

def _one_click_dl(label, gen_fn, filename):
    """Single-click: show progress bar, generate report, auto-download via JS."""
    import base64
    import streamlit.components.v1 as _stc
    MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    col_btn, _ = st.columns([3,1])
    with col_btn:
        if st.button(label, type="primary", use_container_width=True):
            prog = st.progress(0, text=T("dl_spinner"))
            try:
                prog.progress(20, text=T("dl_spinner"))
                buf = gen_fn()
                prog.progress(80, text=T("dl_spinner"))
                b64 = base64.b64encode(buf.getvalue()).decode()
                prog.progress(100, text="✓ Ready — downloading…")
                time.sleep(0.3)
                prog.empty()
                _stc.html(f"""
<a id="_sdl" href="data:{MIME};base64,{b64}" download="{filename}"
   style="display:inline-block;margin:6px 0 2px;padding:10px 22px;
          background:#007A3D;color:#fff;font-weight:700;border-radius:5px;
          text-decoration:none;font-family:Arial,sans-serif;font-size:14px;">
  ⬇️ {filename}
</a>
<script>(function(){{var a=document.getElementById('_sdl');if(a)a.click();}})();</script>
""", height=55)
            except Exception as exc:
                prog.empty()
                st.error(f"Report error: {exc}")

# ── Helpers ───────────────────────────────────────────────────
def call_claude(system, user, max_tokens=2000):
    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)
    full_sys = system + _lang_suffix()
    for attempt in range(3):
        try:
            msg = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=max_tokens,
                system=full_sys,
                messages=[{"role": "user", "content": user}]
            )
            return msg.content[0].text
        except Exception as e:
            err = str(e)
            if "overloaded" in err.lower() and attempt < 2:
                time.sleep(3)
            elif "auth" in err.lower() or "api_key" in err.lower() or "401" in err:
                st.error(f"API key error: {err}. Check your Streamlit secrets.")
                st.stop()
            else:
                raise e

def call_claude_chat(system, history, max_tokens=500):
    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)
    full_sys = system + _lang_suffix()
    for attempt in range(3):
        try:
            msg = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=max_tokens,
                system=full_sys,
                messages=history
            )
            return msg.content[0].text
        except Exception as e:
            if "overloaded" in str(e).lower() and attempt < 2:
                time.sleep(3)
            else:
                return "The API is briefly overloaded — please try again in a moment."

def tavily_search(query):
    if not TAVILY_KEY:
        return []
    try:
        resp = requests.post("https://api.tavily.com/search", json={
            "api_key": TAVILY_KEY, "query": query,
            "search_depth": "basic", "max_results": 5
        }, timeout=10)
        return [{"title": r.get("title",""), "url": r.get("url",""),
                 "content": r.get("content","")} for r in resp.json().get("results",[])]
    except:
        return []


import re as _re_global

def _parse_json(raw: str) -> dict:
    """Robustly parse JSON from a Claude response, handling common failure modes:
    1. Preamble/postamble text around the JSON block
    2. ```json ... ``` fences
    3. Inline rubric annotations on numeric fields, e.g.
       "market_score": 7 (strong, $2-10bn)  →  "market_score": 7
    4. Trailing commas before } or ]
    """
    if not raw:
        raise ValueError("Empty response")
    text = raw.strip()
    # Strip markdown fences
    text = _re_global.sub(r"```json\s*", "", text)
    text = _re_global.sub(r"```\s*", "", text).strip()
    # Extract outermost { … }
    fb = text.find("{"); lb = text.rfind("}")
    if fb >= 0 and lb > fb:
        text = text[fb:lb+1]
    # Strip inline rubric annotations on numeric values:
    #   "field": 7 (some text)  →  "field": 7
    text = _re_global.sub(r'("[\w_]+"\s*:\s*)(\d+(?:\.\d+)?)\s*\([^)]*\)', r'\1\2', text)
    # Remove trailing commas before } or ]
    text = _re_global.sub(r',\s*([}\]])', r'\1', text)
    return json.loads(text)


# ── Session state ─────────────────────────────────────────────
defaults = {
    "active_stage": 1,
    "ui_lang": "en",
    # User identity
    "user_name": "",
    "user_position": "",
    "user_dept": "",
    # Stage 01
    "s1_step": 1,
    "s1_idea": "",
    "s1_questions": [],
    "s1_answers": [],
    "s1_classification": {},
    "s1_chat": [],
    "s1_similar_ideas": [],
    # Stage 02
    "s2_step": "intro",
    "s2_data": {},
    "s2_chat": [],
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── Sidebar ───────────────────────────────────────────────────
with st.sidebar:
    # ── Logo ──────────────────────────────────────────────────
    st.markdown(f"""
<div style="padding:20px 12px 12px 12px;">
  <div style="font-family:'Arial','Helvetica Neue',Helvetica,sans-serif;font-size:20px;font-weight:700;letter-spacing:3px;color:#FFFFFF;line-height:1;">SCHAEFFLER</div>
  <div style="font-family:'Arial','Helvetica Neue',Helvetica,sans-serif;font-size:8px;letter-spacing:3.5px;color:rgba(255,255,255,0.7);margin-top:3px;font-weight:400;">WE PIONEER MOTION</div>
  <div style="background:rgba(255,255,255,0.2);height:1px;margin:16px 0 12px 0;"></div>
  <div style="font-family:'Arial','Helvetica Neue',Helvetica,sans-serif;font-size:9px;letter-spacing:2px;color:rgba(255,255,255,0.65);font-weight:600;">{T("pipeline")}</div>
</div>
""", unsafe_allow_html=True)

    st.markdown('<div style="background:rgba(255,255,255,0.15);height:1px;margin:4px 0 10px;"></div>', unsafe_allow_html=True)

    # ── Stage navigation ───────────────────────────────────────
    stages = T("stages")
    completed = set()
    if st.session_state.get("s1_classification"): completed.add(1)
    if st.session_state.get("s2_data"):           completed.add(2)
    if st.session_state.get("s3_data"):           completed.add(3)
    if st.session_state.get("s4_data"):           completed.add(4)
    if st.session_state.get("s5_data"):           completed.add(5)
    if st.session_state.get("s6_data"):           completed.add(6)

    for num, label in stages:
        active = st.session_state.active_stage
        if num == active:
            st.markdown(f"""<div style="font-family:Arial,sans-serif;background:rgba(255,255,255,0.15);border-radius:4px;padding:8px 12px;margin:3px 0;font-size:12px;font-weight:700;color:#FFFFFF;border-left:3px solid #FFFFFF;letter-spacing:0.3px;">&#9658; {label}</div>""", unsafe_allow_html=True)
        elif num in completed:
            if st.button(f"✓  {label}", key=f"nav_{num}", use_container_width=True):
                st.session_state.active_stage = num
                st.rerun()
        else:
            st.markdown(f"""<div style="font-family:Arial,sans-serif;padding:8px 12px;margin:3px 0;font-size:12px;color:rgba(255,255,255,0.45);">&#9675; {label}</div>""", unsafe_allow_html=True)

    if st.session_state.s1_idea:
        st.markdown("---")
        st.caption(f"**{T('idea_cap')}:** {st.session_state.s1_idea[:60]}...")
        if st.session_state.s1_classification:
            st.caption(f"**{T('quad_cap')}:** {st.session_state.s1_classification.get('quadrant','')}")

    # ── Language + Ideas Log — fixed at bottom of sidebar ──
    st.markdown('<div class="ideas-log-fixed">', unsafe_allow_html=True)
    # Language selector
    _lang_opts   = ["English", "Deutsch"]
    _lang_to_key = {"English":"en","Deutsch":"de"}
    _key_to_lang = {"en":"English","de":"Deutsch"}
    _sel = st.selectbox(
        T("lang_label"), _lang_opts,
        index=_lang_opts.index(_key_to_lang.get(st.session_state.ui_lang,"English")),
        key="_lang_sel"
    )
    if _lang_to_key[_sel] != st.session_state.ui_lang:
        st.session_state.ui_lang = _lang_to_key[_sel]
        # Clear all cached stage outputs so every stage re-runs in the new language
        for _k in ["s2_data","s3_data","s4_data","s5_data","s6_data",
                   "s2_step","s3_step","s4_step","s5_step","s6_step",
                   "s2_chat","s3_chat","s4_chat","s5_chat","s6_chat",
                   "s1_classification","s1_chat","s1_similar_ideas"]:
            if _k in st.session_state:
                del st.session_state[_k]
        st.session_state.active_stage = 1
        st.rerun()
    st.markdown('<div style="background:rgba(255,255,255,0.12);height:1px;margin:6px 0 4px;"></div>', unsafe_allow_html=True)
    import base64 as _b64, streamlit.components.v1 as _stcv1
    _XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    if st.button(T("dl_ideas"), key="sidebar_xl", use_container_width=True):
        with st.spinner(""):
            try:
                _xl_buf = generate_ideas_excel()
                _xl_b64 = _b64.b64encode(_xl_buf.getvalue()).decode()
                _xl_fn  = f"Schaeffler_Ideas_Log_{datetime.now().strftime('%Y%m%d')}.xlsx"
                _stcv1.html(f"""
<a id="_xdl" href="data:{_XLSX_MIME};base64,{_xl_b64}" download="{_xl_fn}"
   style="display:inline-block;padding:7px 14px;background:#005a2d;color:#fff;
          font-size:12px;font-weight:600;border-radius:4px;text-decoration:none;
          font-family:Arial,sans-serif;">{_xl_fn}</a>
<script>(function(){{var a=document.getElementById('_xdl');if(a)a.click();}})();</script>
""", height=44)
            except Exception as exc:
                st.error(str(exc))
    st.markdown('</div>', unsafe_allow_html=True)

# ── Ansoff chart helper ───────────────────────────────────────
def ansoff_chart(quadrant, tech_score, market_score):
    q_cols = {"EXPLOIT":"#1a2d45","EXTEND":"#1e3a5f","RADICAL":"#1F3864","DISRUPT":"#0d2137","DISRUPTIVE":"#0d2137"}
    text_col = "#e2e8f0"; dim_col = "#4a6fa5"; grid_col = "#2a4a70"

    fig = go.Figure()
    # Schaeffler convention: X=Technology (left=Established, right=New)
    #                        Y=Market (bottom=Established, top=New)
    # EXPLOIT=bottom-left, EXTEND=top-left, DISRUPTIVE=bottom-right, RADICAL=top-right
    # Schaeffler convention (from paper):
    # X axis = Market (left=Established → right=New to the World)
    # Y axis = Technology (bottom=Established → top=New to the World)
    # EXPLOIT = bottom-left (existing tech, existing market)
    # EXTEND  = top-left   (new tech, existing market) — Schaeffler calls this differently
    # RADICAL = top-right  (new tech, new market) — appears top-right
    # DISRUPTIVE = bottom-right (existing market novelty, new tech) 
    # Per Lau 2023: EXPLOIT(low/low) EXTEND(low/high) RADICAL(high/high) DISRUPTIVE(high/low)
    # X=Technology, Y=Market in paper. Radical+Disruptive on right side of X.
    for q in [
        dict(x=[0,5,5,0],   y=[0,0,5,5],   name="EXPLOIT",    lx=2.5, ly=2.5),
        dict(x=[0,5,5,0],   y=[5,5,10,10], name="EXTEND",     lx=2.5, ly=7.5),
        dict(x=[5,10,10,5], y=[5,5,10,10], name="RADICAL",    lx=7.5, ly=7.5),
        dict(x=[5,10,10,5], y=[0,0,5,5],   name="DISRUPTIVE", lx=7.5, ly=2.5),
    ]:
        fig.add_trace(go.Scatter(
            x=q["x"]+[q["x"][0]], y=q["y"]+[q["y"][0]],
            fill="toself", fillcolor=q_cols[q["name"]],
            line=dict(color=grid_col, width=1),
            mode="lines", showlegend=False, hoverinfo="skip"
        ))
        fig.add_annotation(x=q["lx"], y=q["ly"], text=f"<b>{q['name']}</b>",
            showarrow=False,
            font=dict(size=14, color=BLUE if q["name"]==quadrant else dim_col))

    fig.add_shape(type="line",x0=5,x1=5,y0=0,y1=10,line=dict(color=grid_col,width=1.5,dash="dot"))
    fig.add_shape(type="line",x0=0,x1=10,y0=5,y1=5,line=dict(color=grid_col,width=1.5,dash="dot"))

    # 4-level axis tick labels (X=Technology, Y=Market — Schaeffler convention)
    for x_pos, x_label in [(1.25,"Established"),(3.75,"Adjacent"),(6.25,"New to Schaeffler"),(8.75,"New to the World")]:
        fig.add_annotation(x=x_pos, y=-1.0, text=x_label, showarrow=False,
            font=dict(size=9, color=dim_col), textangle=0)
    for y_pos, y_label in [(1.25,"Established"),(3.75,"Adjacent"),(6.25,"New to\nSchaeffler"),(8.75,"New to\nthe World")]:
        fig.add_annotation(x=-1.3, y=y_pos, text=y_label, showarrow=False,
            font=dict(size=9, color=dim_col), textangle=-90)
    # Axis spine labels
    fig.add_annotation(x=5, y=-1.8, text="← Technology Dimension (Newness) →", showarrow=False,
        font=dict(size=10, color=text_col))
    fig.add_annotation(x=-2.4, y=5, text="← Market Dimension (Newness) →", showarrow=False,
        font=dict(size=10, color=text_col), textangle=-90)
    # Dividing lines at each level boundary (2.5, 5, 7.5)
    for v in [2.5, 5.0, 7.5]:
        fig.add_shape(type="line", x0=v, x1=v, y0=0, y1=10,
            line=dict(color=grid_col, width=0.8, dash="dot"))
        fig.add_shape(type="line", x0=0, x1=10, y0=v, y1=v,
            line=dict(color=grid_col, width=0.8, dash="dot"))

    fig.add_trace(go.Scatter(
        x=[tech_score], y=[market_score], mode="markers+text",
        marker=dict(size=20, color=BLUE, line=dict(color="white",width=2)),
        text=["  Your idea"], textposition="middle right",
        textfont=dict(size=12,color="white",family="Arial Bold"),
        showlegend=False,
        hovertemplate=f"<b>{quadrant}</b><br>Tech: {tech_score}/10<br>Market: {market_score}/10<extra></extra>"
    ))
    fig.update_layout(
        title=dict(text="Schaeffler Innovation Framework — Modified Ansoff Matrix",
                   font=dict(size=13,color=text_col),x=0.5),
        xaxis=dict(range=[-1.6,11],showticklabels=False,showgrid=False,zeroline=False,
                   title="Technology Dimension →",title_font=dict(size=11,color=dim_col)),
        yaxis=dict(range=[-1.6,11],showticklabels=False,showgrid=False,zeroline=False,
                   title="Market Dimension →",title_font=dict(size=11,color=dim_col)),
        plot_bgcolor=BG, paper_bgcolor=BG, height=420,
        margin=dict(l=110,r=30,t=50,b=70), font=dict(color=text_col)
    )
    return fig

def get_dot_position(quadrant, confidence):
    # X=Technology, Y=Market (Schaeffler convention)
    # base = (tech_x, market_y)
    base = {
        "EXPLOIT":    (2.5, 2.5),   # low tech, low market  — bottom-left
        "EXTEND":     (2.5, 7.5),   # low tech, high market — top-left
        "RADICAL":    (7.5, 7.5),   # high tech, high market — top-right
        "DISRUPTIVE": (7.5, 2.5),   # high tech, low market — bottom-right
        "DISRUPT":    (7.5, 2.5),   # alias
    }
    nudge = {"High":1.5,"Medium":1.0,"Low":0.5}.get(confidence,1.0)
    t, m = base.get(quadrant,(5,5))
    if quadrant in ("EXPLOIT",):      t-=nudge; m-=nudge
    elif quadrant in ("EXTEND",):     t-=nudge; m+=nudge
    elif quadrant in ("RADICAL",):    t+=nudge; m+=nudge
    elif quadrant in ("DISRUPTIVE","DISRUPT"): t+=nudge; m-=nudge
    return round(max(0.5,min(9.5,t)),1), round(max(0.5,min(9.5,m)),1)

# ════════════════════════════════════════════════════════════
# STAGE 01 — QUADRANT CLASSIFIER
# ════════════════════════════════════════════════════════════




def generate_org_report(idea, quadrant, s1c, s5d):
    """Generate a Stage 05 Organisational Readiness Word report."""
    org_data   = s5d.get("org_data", {})
    portfolio  = org_data.get("p3_portfolio", {})
    people     = org_data.get("p3_people", {})
    process    = org_data.get("p3_process", {})
    bop        = org_data.get("build_or_partner", {})
    gaps       = org_data.get("org_gaps", [])
    partners   = org_data.get("partnership_candidates", [])

    org_ctx = (
        f"Idea: {idea}\nQuadrant: {quadrant}\n"
        f"P3 Portfolio score: {s5d.get('p_portfolio',5)}/10  Rationale: {portfolio.get('rationale','')}\n"
        f"Cluster fit: {portfolio.get('cluster_fit','')}  Strengths: {chr(44).join(portfolio.get('strengths',[])[:3])}\n"
        f"P3 People score: {s5d.get('p_people',5)}/10  Rationale: {people.get('rationale','')}\n"
        f"Matched competencies: {chr(44).join(people.get('matched_competencies',[])[:4])}\n"
        f"Critical gap: {people.get('competency_gap','')}  Closure route: {people.get('sourcing_route','')}\n"
        f"P3 Process score: {s5d.get('p_process',5)}/10  Rationale: {process.get('rationale','')}\n"
        f"Applicable assets: {chr(44).join(process.get('applicable_assets',[])[:3])}\n"
        f"Investment required: {process.get('investment_required','')}  Time to close: {process.get('time_to_close','')}\n"
        f"Build strategy: {bop.get('recommendation','')}  Rationale: {bop.get('rationale','')}\n"
        f"Time internal: {bop.get('time_to_trl6_internal','')}  Time with partner: {bop.get('time_to_trl6_partner','')}\n"
        f"Overall org readiness score: {s5d.get('final_score',5)}/10"
    )
    extended = call_claude(
        '''You are a senior Schaeffler innovation strategist writing a detailed Organisational Readiness report.
Write specific, substantive content referencing Schaeffler P3 formula (Performance = Portfolio x People x Process).
Return ONLY valid JSON, no markdown backticks:
{
  "executive_summary": "3-4 full paragraphs: overall organisational readiness verdict, key P3 strengths and gaps, build-or-partner recommendation, and strategic rationale.",
  "portfolio_analysis": "2-3 full paragraphs on strategic portfolio fit — how this idea aligns with Schaeffler innovation clusters, trends, product families, and the broader electrification strategy post-Vitesco.",
  "people_analysis": "2-3 full paragraphs on human capital readiness — which Schaeffler competencies match, what is critically missing, and the most realistic route to close the gap (hire, upskill, acquire, partner).",
  "process_analysis": "2-3 full paragraphs on process and infrastructure readiness — which Schaeffler assets and processes apply directly, what needs to be built, and estimated investment and timeline.",
  "partnership_strategy": "2 full paragraphs on the recommended partnership strategy — who Schaeffler should partner with, what type of arrangement, and how to structure it.",
  "risks": ["Org risk 1 with mitigation", "Org risk 2 with mitigation", "Org risk 3", "Org risk 4"],
  "recommendations": ["Concrete org action 1 with timeline", "Concrete action 2", "Concrete action 3", "Concrete action 4"]
}''',
        org_ctx,
        max_tokens=3500
    )
    raw_e = extended.strip().replace("```json","").replace("```","").strip()
    fb = raw_e.find("{"); lb = raw_e.rfind("}") + 1
    if fb >= 0: raw_e = raw_e[fb:lb]
    try:
        ext = json.loads(raw_e)
    except:
        ext = {
            "executive_summary": f"Schaeffler\'s organisational readiness for this idea scores {s5d.get('final_score',5)}/10. The P3 assessment shows Portfolio fit at {s5d.get('p_portfolio',5)}/10, People readiness at {s5d.get('p_people',5)}/10, and Process readiness at {s5d.get('p_process',5)}/10. Recommended build strategy: {bop.get('recommendation','Co-develop')}. {bop.get('rationale','')}",
            "portfolio_analysis": f"Portfolio fit: {portfolio.get('rationale','')} Cluster fit: {portfolio.get('cluster_fit','')}",
            "people_analysis": f"People readiness: {people.get('rationale','')} Critical gap: {people.get('competency_gap','')} Closure route: {people.get('sourcing_route','')}",
            "process_analysis": f"Process readiness: {process.get('rationale','')} Investment required: {process.get('investment_required','')} Time to close: {process.get('time_to_close','')}",
            "partnership_strategy": f"Partnership strategy: {bop.get('recommendation','Co-develop')}. Time to TRL6 internally: {bop.get('time_to_trl6_internal','')}. With partner: {bop.get('time_to_trl6_partner','')}.",
            "risks": ["Competency gap requires immediate sourcing action — delay increases time to TRL6", "Partnership negotiations can be slow — initiate early", "Internal process gaps may create bottlenecks in development", "Build vs buy decision requires board-level sign-off"],
            "recommendations": [f"Initiate {bop.get('recommendation','co-development')} process within 30 days", f"Address critical gap: {people.get('competency_gap','key competency')} via {people.get('sourcing_route','targeted hiring')}", "Map applicable Schaeffler assets to innovation project plan", "Establish steering committee for cross-divisional coordination"]
        }

    NAVY=RGBColor(0x1F,0x38,0x64); BLUE=RGBColor(0x2E,0x75,0xB6)
    WHITE=RGBColor(0xFF,0xFF,0xFF); GREY=RGBColor(0x55,0x55,0x55)
    BLACK=RGBColor(0x00,0x00,0x00); LBLUE=RGBColor(0x60,0xA5,0xFA)

    def set_bg(cell, hx):
        tc=cell._tc; tcPr=tc.get_or_add_tcPr(); shd=OxmlElement("w:shd")
        shd.set(qn("w:val"),"clear"); shd.set(qn("w:color"),"auto"); shd.set(qn("w:fill"),hx); tcPr.append(shd)
    def h1(doc, text):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(14); p.paragraph_format.space_after=Pt(4)
        pPr=p._p.get_or_add_pPr(); pBdr=OxmlElement("w:pBdr"); bot=OxmlElement("w:bottom")
        bot.set(qn("w:val"),"single"); bot.set(qn("w:sz"),"8"); bot.set(qn("w:space"),"3"); bot.set(qn("w:color"),"2E75B6")
        pBdr.append(bot); pPr.append(pBdr); r=p.add_run(text)
        r.bold=True; r.font.size=Pt(14); r.font.color.rgb=NAVY
    def h2(doc, text):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(8); p.paragraph_format.space_after=Pt(2)
        r=p.add_run(text); r.bold=True; r.font.size=Pt(11); r.font.color.rgb=NAVY
    def body(doc, text):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(3); p.paragraph_format.space_after=Pt(3)
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; r=p.add_run(text); r.font.size=Pt(10.5)
    def kv(doc, label, value):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
        r1=p.add_run(f"{label}: "); r1.bold=True; r1.font.size=Pt(10.5); r1.font.color.rgb=NAVY
        r2=p.add_run(str(value)); r2.font.size=Pt(10.5)
    def bul(doc, text):
        p=doc.add_paragraph(style="List Bullet"); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
        r=p.add_run(str(text)); r.font.size=Pt(10.5)

    doc=DocxDocument()
    for sec in doc.sections:
        sec.top_margin=Cm(2.0); sec.bottom_margin=Cm(2.0); sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    t=doc.add_table(rows=1,cols=1); t.style="Table Grid"; c=t.cell(0,0); set_bg(c,"1F3864")
    p=c.paragraphs[0]; p.paragraph_format.space_before=Pt(10); p.paragraph_format.space_after=Pt(2)
    r=p.add_run("ORGANISATIONAL READINESS REPORT"); r.bold=True; r.font.size=Pt(9); r.font.color.rgb=WHITE
    p2=c.add_paragraph(); p2.paragraph_format.space_before=Pt(0); p2.paragraph_format.space_after=Pt(10)
    r2=p2.add_run("Schaeffler AI Innovation Research Assistant  ·  Stage 05"); r2.font.size=Pt(9); r2.font.color.rgb=RGBColor(0x93,0xC5,0xFD)
    doc.add_paragraph()
    p=doc.add_paragraph(); r=p.add_run("Organisational Readiness Assessment")
    r.bold=True; r.font.size=Pt(18); r.font.color.rgb=NAVY
    p2=doc.add_paragraph(); r2=p2.add_run(f"Score: {s5d.get('final_score',5)}/10  ·  Strategy: {bop.get('recommendation','')}  ·  Quadrant: {quadrant}  ·  {datetime.now().strftime('%d %B %Y')}"); r2.font.size=Pt(10); r2.italic=True; r2.font.color.rgb=GREY
    doc.add_paragraph()

    h1(doc,"P³ Score Summary")
    p3_tbl=doc.add_table(rows=1,cols=4); p3_tbl.style="Table Grid"
    for i,h in enumerate(["Dimension","Score","Weight","Description"]):
        c=p3_tbl.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
        r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
    for i,(dim,score,wt,desc) in enumerate([
        ("Portfolio",f"{s5d.get('p_portfolio',5):.1f}/10","35%",portfolio.get('cluster_fit','')),
        ("People",f"{s5d.get('p_people',5):.1f}/10","40%",people.get('competency_gap','')),
        ("Process",f"{s5d.get('p_process',5):.1f}/10","25%",process.get('investment_required','')),
    ]):
        row=p3_tbl.add_row(); fill="EAF1FB" if i%2==0 else "FFFFFF"
        for c in row.cells: set_bg(c,fill)
        for j,val in enumerate([dim,score,wt,desc[:80] if desc else ""]):
            r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.bold=(j==0)
    fr=p3_tbl.add_row()
    for c in fr.cells: set_bg(c,"1F3864")
    r=fr.cells[0].paragraphs[0].add_run("OVERALL READINESS"); r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
    r2=fr.cells[1].paragraphs[0].add_run(f"{s5d.get('final_score',5)}/10"); r2.bold=True; r2.font.size=Pt(11); r2.font.color.rgb=LBLUE
    doc.add_paragraph()

    h1(doc,"Executive Summary"); body(doc, ext.get("executive_summary",""))
    h1(doc,"P³ Portfolio — Strategic Fit"); body(doc, ext.get("portfolio_analysis",""))
    if portfolio.get("strengths"):
        h2(doc,"Portfolio Strengths")
        for s in portfolio["strengths"]: bul(doc, f"✓ {s}")
    if portfolio.get("gaps"):
        h2(doc,"Portfolio Gaps")
        for g in portfolio["gaps"]: bul(doc, f"✗ {g}")
    h1(doc,"P³ People — Competency Readiness"); body(doc, ext.get("people_analysis",""))
    if people.get("matched_competencies"):
        h2(doc,"Matched Competencies")
        for c in people["matched_competencies"]: bul(doc, f"✓ {c}")
    kv(doc,"Critical competency gap", people.get("competency_gap",""))
    kv(doc,"Closure route", people.get("sourcing_route",""))
    h1(doc,"P³ Process — Infrastructure & Assets"); body(doc, ext.get("process_analysis",""))
    if process.get("applicable_assets"):
        h2(doc,"Applicable Assets")
        for a in process["applicable_assets"]: bul(doc, a)
    kv(doc,"Investment required", process.get("investment_required",""))
    kv(doc,"Estimated time to close", process.get("time_to_close",""))

    if gaps:
        h1(doc,"Organisational Gaps Register")
        gt=doc.add_table(rows=1,cols=4); gt.style="Table Grid"
        for i,h in enumerate(["Gap","Severity","Closure Route","Timeline"]):
            c=gt.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,g in enumerate(gaps):
            row=gt.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            sev_fill={"High":"FFE4E4","Medium":"FFF8E4","Low":"E4FFE9"}.get(g.get("severity",""),"FFFFFF")
            for c in row.cells: set_bg(c,sev_fill if idx==0 else fill)
            for j,val in enumerate([g.get("gap",""),g.get("severity",""),g.get("closure_route",""),g.get("timeline","")]):
                r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.bold=(j==0)

    if partners:
        h1(doc,"Partnership Candidates"); body(doc, ext.get("partnership_strategy",""))
        pt=doc.add_table(rows=1,cols=4); pt.style="Table Grid"
        for i,h in enumerate(["Organisation","Type","Rationale","Route"]):
            c=pt.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,p in enumerate(partners):
            row=pt.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            for c in row.cells: set_bg(c,fill)
            for j,val in enumerate([p.get("name",""),p.get("type",""),p.get("rationale",""),p.get("route","")]):
                r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.bold=(j==0)

    h1(doc,"Build-or-Partner Recommendation")
    kv(doc,"Recommendation", bop.get("recommendation",""))
    kv(doc,"Time to TRL 6 — Internal", bop.get("time_to_trl6_internal",""))
    kv(doc,"Time to TRL 6 — With Partner", bop.get("time_to_trl6_partner",""))
    body(doc, bop.get("rationale",""))
    h1(doc,"Risks"); 
    for risk in ext.get("risks",[]): bul(doc, risk)
    h1(doc,"Recommendations")
    for rec in ext.get("recommendations",[]): bul(doc, rec)

    doc.add_paragraph()
    ft=doc.add_table(rows=1,cols=1); ft.style="Table Grid"; fc=ft.cell(0,0); set_bg(fc,"1F3864")
    fp=fc.paragraphs[0]; fp.paragraph_format.space_before=Pt(6); fp.paragraph_format.space_after=Pt(6)
    fr=fp.add_run(f"Schaeffler AI Innovation Research Assistant  ·  Stage 05: Organisational Readiness  ·  {datetime.now().strftime('%d %B %Y')}  ·  Capstone Project — Arpan Chowdhury, EBS Universität")
    fr.font.size=Pt(8); fr.font.color.rgb=RGBColor(0x93,0xC5,0xFD)
    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf


def generate_master_report(idea, quadrant, s1c, s2d, s3d, s4d, s5d_org, s6d):
    """Generate the full comprehensive Innovation Assessment Word report covering all stages."""
    ipi       = s6d.get("ipi", 0)
    weights   = s6d.get("weights", {"market":35,"patent":25,"feasibility":25,"org":15})
    synthesis = s6d.get("synthesis", {})
    scores    = s6d.get("scores", {})
    market    = s2d.get("market", {})
    comp      = s2d.get("comp", {})
    sectors   = s2d.get("sectors", {})
    landscape = s3d.get("landscape", {})
    ansoff_d  = s3d.get("ansoff_data", {})
    existence = s4d.get("existence", {})
    trl       = s4d.get("trl", {})
    org_data  = s5d_org.get("org_data", {})
    portfolio = org_data.get("p3_portfolio", {})
    people    = org_data.get("p3_people", {})
    process   = org_data.get("p3_process", {})
    bop       = org_data.get("build_or_partner", {})

    # Generate enriched narrative for the master report
    master_ctx = (
        f"Idea: {idea}\nQuadrant: {quadrant}\nIPI: {ipi}/10\nRecommendation: {synthesis.get('recommendation','')}\n\n"
        f"STAGE 02 — Market ({weights.get('market',35)}%): {scores.get('market',5)}/10\n"
        f"Market: {market.get('market_name','')}  Size 2024: {_mval(market.get('market_size_current') or market.get('market_size_2024',''))}  CAGR: {_mval(market.get('cagr',''))}\n"
        f"Primary sectors: {', '.join(s2d.get('sectors',{}).get('primary_sectors',[]))}\n"
        f"Competition: {comp.get('competitive_intensity','')}  White space: {comp.get('white_space','')}\n\n"
        f"STAGE 03 — Patent ({weights.get('patent',25)}%): {scores.get('patent',5)}/10\n"
        f"Activity: {landscape.get('activity_level','')}  Trend: {landscape.get('filing_trend','')}\n"
        f"Novelty: {ansoff_d.get('novelty_signal','')}  IP risk: {ansoff_d.get('ip_risk','')}\n"
        f"White spaces: {'; '.join(landscape.get('white_spaces',[])[:3])}\n\n"
        f"STAGE 04 — Feasibility ({weights.get('feasibility',25)}%): {scores.get('feasibility',5)}/10\n"
        f"TRL: {trl.get('trl_level',3)} — {trl.get('trl_label','')}\n"
        f"Existence: {existence.get('existence_verdict','')}  Entry readiness: {trl.get('schaeffler_entry_readiness','')}\n"
        f"Time to production: {existence.get('time_to_readiness','')}\n\n"
        f"STAGE 05 — Org Readiness ({weights.get('org',15)}%): {scores.get('org',5)}/10\n"
        f"Portfolio: {s5d_org.get('p_portfolio',5)}/10  People: {s5d_org.get('p_people',5)}/10  Process: {s5d_org.get('p_process',5)}/10\n"
        f"Strategy: {bop.get('recommendation','')}  Time with partner: {bop.get('time_to_trl6_partner','')}\n"
        f"Critical gap: {people.get('competency_gap','')}\n\n"
        f"Strongest signals: {'; '.join(synthesis.get('strongest_signals',[])[:3])}\n"
        f"Key concerns: {'; '.join(synthesis.get('key_concerns',[])[:3])}\n"
        f"Strategic fit: {synthesis.get('strategic_fit','')}"
    )
    enriched = call_claude(
        '''You are a senior Schaeffler innovation strategist writing a comprehensive multi-stage Innovation Assessment Report.
Write rich, specific, analytical content. Return ONLY valid JSON, no markdown backticks:
{
  "executive_summary": "4-5 full paragraphs: the headline verdict, what the data shows across all 4 dimensions (market, IP, feasibility, org), the single most compelling reason to proceed, the single most important risk, and the concrete next step.",
  "strategic_narrative": "4-5 full paragraphs synthesising the opportunity holistically — how the market signals, IP landscape, technology readiness, and org capability interact. Reference Schaeffler electrification strategy, Vitesco merger, E-Mobility growth, OEM relationships and P3 formula.",
  "market_highlights": "2 full paragraphs on the top 3 market signals — most important size/growth data point, the most relevant sector fit, and the most significant competitive dynamic.",
  "ip_highlights": "2 full paragraphs on the top IP insights — novelty position, the most threatening filer, and the most valuable white space to capture.",
  "feasibility_highlights": "2 full paragraphs on the technology readiness picture — TRL rationale, most convincing evidence found, and the one critical gap that needs solving before TRL6.",
  "org_highlights": "2 full paragraphs on organisational readiness — the strongest existing capability, the most critical gap, and the recommended build-partner path.",
  "risk_synthesis": ["Top cross-cutting risk 1 with specific mitigation", "Risk 2 with mitigation", "Risk 3 with mitigation", "Risk 4 with mitigation", "Risk 5 with mitigation"],
  "action_plan": ["Immediate action (0-30 days): specific step 1", "Short-term (1-3 months): specific step 2", "Medium-term (3-6 months): specific step 3", "Longer-term (6-12 months): specific step 4", "Strategic (12+ months): specific step 5"]
}''',
        master_ctx, max_tokens=4000
    )
    raw_e = enriched.strip().replace("```json","").replace("```","").strip()
    fb = raw_e.find("{"); lb = raw_e.rfind("}") + 1
    if fb >= 0: raw_e = raw_e[fb:lb]
    try:
        enr = json.loads(raw_e)
    except:
        enr = {
            "executive_summary": f"This innovation idea ({idea[:80]}) has been assessed across four dimensions of Schaeffler\'s Innovation Pipeline, yielding an Innovation Potential Index (IPI) of {ipi}/10. The recommendation is: {synthesis.get('recommendation','PROCEED WITH CONDITIONS')}. {synthesis.get('recommendation_rationale','')} The strongest signal is: {synthesis.get('strongest_signals',['market opportunity'])[0] if synthesis.get('strongest_signals') else 'market opportunity'}. The primary concern is: {synthesis.get('key_concerns',['technical maturity'])[0] if synthesis.get('key_concerns') else 'technical maturity'}.",
            "strategic_narrative": synthesis.get("narrative", f"The idea targets {market.get('market_name','a growing market')} and sits in Schaeffler\'s {quadrant} innovation quadrant. {synthesis.get('strategic_fit','')}"),
            "market_highlights": f"Market size: {_mval(market.get('market_size_current') or market.get('market_size_2024',''))} (2024), growing to {_mval(market.get('market_size_forecast') or market.get('market_size_2030',''))} by 2030 at {_mval(market.get('cagr',''))} CAGR. Primary sector fit: {', '.join(s2d.get('sectors',{}).get('primary_sectors',[]))}. Competitive intensity: {comp.get('competitive_intensity','')}.",
            "ip_highlights": f"Novelty signal: {ansoff_d.get('novelty_signal','Moderate')}. IP risk: {ansoff_d.get('ip_risk','Medium')}. White spaces: {'; '.join(landscape.get('white_spaces',[])[:2])}.",
            "feasibility_highlights": f"TRL {trl.get('trl_level',3)}: {trl.get('trl_label','')}. Existence: {existence.get('existence_verdict','')}. Time to production: {existence.get('time_to_readiness','')}.",
            "org_highlights": f"Org readiness: {scores.get('org',5)}/10. Strategy: {bop.get('recommendation','Co-develop')}. Critical gap: {people.get('competency_gap','')}.",
            "risk_synthesis": synthesis.get("risks", ["IP risk requires FTO analysis", "Technical maturity requires R&D investment", "Competitive dynamics require monitoring", "Org readiness gaps require targeted hiring/partnering"]),
            "action_plan": synthesis.get("next_steps", ["Commission engineering feasibility review", "Conduct FTO IP analysis", "Identify pilot customer", "Present to innovation steering committee"])
        }

    NAVY=RGBColor(0x1F,0x38,0x64); BLUE=RGBColor(0x2E,0x75,0xB6)
    WHITE=RGBColor(0xFF,0xFF,0xFF); GREY=RGBColor(0x55,0x55,0x55)
    BLACK=RGBColor(0x00,0x00,0x00); LBLUE=RGBColor(0x60,0xA5,0xFA)
    GREEN=RGBColor(0x22,0xC5,0x5E); AMBER=RGBColor(0xF5,0x9E,0x0B)

    def set_bg(cell, hx):
        tc=cell._tc; tcPr=tc.get_or_add_tcPr(); shd=OxmlElement("w:shd")
        shd.set(qn("w:val"),"clear"); shd.set(qn("w:color"),"auto"); shd.set(qn("w:fill"),hx); tcPr.append(shd)
    def h1(doc, text):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(14); p.paragraph_format.space_after=Pt(4)
        pPr=p._p.get_or_add_pPr(); pBdr=OxmlElement("w:pBdr"); bot=OxmlElement("w:bottom")
        bot.set(qn("w:val"),"single"); bot.set(qn("w:sz"),"8"); bot.set(qn("w:space"),"3"); bot.set(qn("w:color"),"2E75B6")
        pBdr.append(bot); pPr.append(pBdr); r=p.add_run(text)
        r.bold=True; r.font.size=Pt(14); r.font.color.rgb=NAVY
    def h2(doc, text):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(8); p.paragraph_format.space_after=Pt(2)
        r=p.add_run(text); r.bold=True; r.font.size=Pt(11); r.font.color.rgb=NAVY
    def body(doc, text):
        if not text: return
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(3); p.paragraph_format.space_after=Pt(3)
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; r=p.add_run(str(text)); r.font.size=Pt(10.5)
    def kv(doc, label, value):
        if not value: return
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
        r1=p.add_run(f"{label}: "); r1.bold=True; r1.font.size=Pt(10.5); r1.font.color.rgb=NAVY
        r2=p.add_run(str(value)); r2.font.size=Pt(10.5)
    def bul(doc, text):
        if not text: return
        p=doc.add_paragraph(style="List Bullet"); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
        r=p.add_run(str(text)); r.font.size=Pt(10.5)

    doc=DocxDocument()
    for sec in doc.sections:
        sec.top_margin=Cm(2.0); sec.bottom_margin=Cm(2.0); sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    # ── Cover header ──────────────────────────────────────────────
    t=doc.add_table(rows=1,cols=1); t.style="Table Grid"; c=t.cell(0,0); set_bg(c,"1F3864")
    p=c.paragraphs[0]; p.paragraph_format.space_before=Pt(14); p.paragraph_format.space_after=Pt(2)
    r=p.add_run("SCHAEFFLER INNOVATION ASSESSMENT REPORT"); r.bold=True; r.font.size=Pt(12); r.font.color.rgb=WHITE
    p2=c.add_paragraph(); p2.paragraph_format.space_before=Pt(0); p2.paragraph_format.space_after=Pt(14)
    r2=p2.add_run("AI Innovation Research Assistant  ·  Full Pipeline Assessment  ·  Stages 01–06")
    r2.font.size=Pt(9); r2.font.color.rgb=RGBColor(0x93,0xC5,0xFD)
    doc.add_paragraph()

    # ── Title ──────────────────────────────────────────────────────
    p=doc.add_paragraph(); r=p.add_run(market.get("market_name", idea[:80]) or idea[:80])
    r.bold=True; r.font.size=Pt(22); r.font.color.rgb=NAVY
    p2=doc.add_paragraph()
    rec_text = synthesis.get("recommendation","")
    r2=p2.add_run(f"IPI: {ipi}/10  ·  {rec_text}  ·  Quadrant: {quadrant}  ·  {datetime.now().strftime('%d %B %Y')}")
    r2.font.size=Pt(10); r2.italic=True; r2.font.color.rgb=GREY

    # ── Idea box ──────────────────────────────────────────────────
    doc.add_paragraph()
    tb=doc.add_table(rows=1,cols=2); tb.style="Table Grid"
    c1=tb.cell(0,0); c2=tb.cell(0,1); set_bg(c1,"1F3864"); set_bg(c2,"EAF1FB"); c1.width=Inches(0.12)
    c1.paragraphs[0].add_run("")
    rp=c2.paragraphs[0]; rp.paragraph_format.space_before=Pt(8); rp.paragraph_format.space_after=Pt(2)
    rb=rp.add_run("Innovation Idea"); rb.bold=True; rb.font.size=Pt(9); rb.font.color.rgb=NAVY
    rp2=c2.add_paragraph(); rp2.paragraph_format.space_before=Pt(0); rp2.paragraph_format.space_after=Pt(8)
    ri=rp2.add_run(idea); ri.font.size=Pt(10); ri.italic=True
    doc.add_paragraph()

    # ── IPI Score table ───────────────────────────────────────────
    h1(doc,"Innovation Potential Index (IPI)")
    ipi_tbl=doc.add_table(rows=1,cols=4); ipi_tbl.style="Table Grid"
    for i,h in enumerate(["Stage","Score","Weight","Contribution"]):
        c=ipi_tbl.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
        r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
    for i,(stage,score,wt,contrib) in enumerate([
        ("02 · Market Intelligence",    f"{scores.get('market',5):.1f}/10",      f"{weights.get('market',35)}%",      f"{scores.get('market',5)*weights.get('market',35)/100:.2f}"),
        ("03 · Patent Intelligence",    f"{scores.get('patent',5):.1f}/10",      f"{weights.get('patent',25)}%",      f"{scores.get('patent',5)*weights.get('patent',25)/100:.2f}"),
        ("04 · Technical Feasibility",  f"{scores.get('feasibility',5):.1f}/10", f"{weights.get('feasibility',25)}%", f"{scores.get('feasibility',5)*weights.get('feasibility',25)/100:.2f}"),
        ("05 · Organisational Readiness",f"{scores.get('org',5):.1f}/10",        f"{weights.get('org',15)}%",         f"{scores.get('org',5)*weights.get('org',15)/100:.2f}"),
    ]):
        row=ipi_tbl.add_row(); fill="EAF1FB" if i%2==0 else "FFFFFF"
        for c in row.cells: set_bg(c,fill)
        for j,val in enumerate([stage,score,wt,contrib]):
            r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(10); r.bold=(j==0)
    fr=ipi_tbl.add_row()
    for c in fr.cells: set_bg(c,"1F3864")
    r=fr.cells[0].paragraphs[0].add_run("INNOVATION POTENTIAL INDEX"); r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
    r2=fr.cells[3].paragraphs[0].add_run(f"{ipi:.1f}/10"); r2.bold=True; r2.font.size=Pt(13); r2.font.color.rgb=LBLUE
    doc.add_paragraph()

    # ── Recommendation panel ──────────────────────────────────────
    h1(doc,"Recommendation")
    rec_col_map = {"PROCEED":"EAF9F0","PROCEED WITH CONDITIONS":"FFF8E4","DEFER":"FFF0E4","REJECT":"FFE9E9"}
    rec_fill = rec_col_map.get(rec_text, "EAF1FB").lstrip("#")
    rec_tbl=doc.add_table(rows=1,cols=1); rec_tbl.style="Table Grid"
    rc=rec_tbl.cell(0,0); set_bg(rc, rec_fill)
    p=rc.paragraphs[0]; p.paragraph_format.space_before=Pt(8); p.paragraph_format.space_after=Pt(4)
    r=p.add_run(rec_text); r.bold=True; r.font.size=Pt(14); r.font.color.rgb=NAVY
    p2=rc.add_paragraph(); p2.paragraph_format.space_before=Pt(0); p2.paragraph_format.space_after=Pt(8)
    r2=p2.add_run(synthesis.get("recommendation_rationale","")); r2.font.size=Pt(10)
    doc.add_paragraph()

    if synthesis.get("strongest_signals"):
        h2(doc,"Strongest Signals")
        for s in synthesis["strongest_signals"]: bul(doc, f"✓  {s}")
    if synthesis.get("key_concerns"):
        h2(doc,"Key Concerns")
        for c in synthesis["key_concerns"]: bul(doc, f"⚠  {c}")
    if synthesis.get("conditions"):
        h2(doc,"Conditions to Proceed")
        for c in synthesis["conditions"]: bul(doc, f"→  {c}")

    # ── Executive Summary ─────────────────────────────────────────
    h1(doc,"Executive Summary")
    body(doc, enr.get("executive_summary",""))

    # ── Strategic Narrative ───────────────────────────────────────
    h1(doc,"Strategic Narrative & Full Synthesis")
    body(doc, enr.get("strategic_narrative",""))
    if synthesis.get("strategic_fit"):
        doc.add_paragraph()
        body(doc, synthesis.get("strategic_fit",""))

    # ── Stage 02 — Market Intelligence ───────────────────────────
    h1(doc,"Stage 02 · Market Intelligence  ·  Score: " + str(scores.get("market",5)) + "/10")
    body(doc, enr.get("market_highlights",""))
    doc.add_paragraph()
    kv(doc,"Market", market.get("market_name",""))
    kv(doc,"Size (2024)", _mval(market.get("market_size_current") or market.get("market_size_2024","")))
    kv(doc,"Projected (2030)", _mval(market.get("market_size_forecast") or market.get("market_size_2030","")))
    kv(doc,"CAGR", market.get("cagr",""))
    kv(doc,"Maturity", market.get("market_maturity",""))
    kv(doc,"Geography", market.get("geographic_focus",""))
    kv(doc,"Competitive intensity", comp.get("competitive_intensity",""))
    kv(doc,"White space", comp.get("white_space",""))
    kv(doc,"Schaeffler advantage", comp.get("schaeffler_advantage",""))
    if market.get("growth_drivers"):
        h2(doc,"Growth Drivers")
        for d in market["growth_drivers"][:4]: bul(doc, d)
    if comp.get("competitors"):
        h2(doc,"Key Competitors")
        ct=doc.add_table(rows=1,cols=3); ct.style="Table Grid"
        for i,h in enumerate(["Company","Type","Relevance"]):
            c=ct.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,ci in enumerate(comp["competitors"]):
            row=ct.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            for c in row.cells: set_bg(c,fill)
            for j,val in enumerate([ci.get("name",""),ci.get("type",""),ci.get("relevance","")+" "+ci.get("source","")]):
                r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.bold=(j==0)
    if sectors.get("sector_scores"):
        h2(doc,"Schaeffler Sector Cluster Scores")
        primary=sectors.get("primary_sectors",[])
        st2=doc.add_table(rows=1,cols=3); st2.style="Table Grid"
        for i,h in enumerate(["Sector","Score","Rationale"]):
            c=st2.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,(sec,data) in enumerate(sectors["sector_scores"].items()):
            row=st2.add_row(); fill="EAF5EA" if sec in primary else ("EAF1FB" if idx%2==0 else "FFFFFF")
            for c in row.cells: set_bg(c,fill)
            r0=row.cells[0].paragraphs[0].add_run(sec); r0.font.size=Pt(10); r0.bold=(sec in primary)
            r1=row.cells[1].paragraphs[0].add_run(f"{data.get('score',0)}/10"); r1.font.size=Pt(10); r1.bold=True
            r2=row.cells[2].paragraphs[0].add_run(data.get("rationale","")); r2.font.size=Pt(9.5)

    # ── Stage 03 — Patent Intelligence ───────────────────────────
    h1(doc,"Stage 03 · Patent Intelligence  ·  Score: " + str(scores.get("patent",5)) + "/10")
    body(doc, enr.get("ip_highlights",""))
    doc.add_paragraph()
    kv(doc,"Filing activity", landscape.get("activity_level",""))
    kv(doc,"Filing trend", landscape.get("filing_trend",""))
    kv(doc,"Trend rationale", landscape.get("filing_trend_rationale",""))
    kv(doc,"Novelty signal", ansoff_d.get("novelty_signal",""))
    kv(doc,"Novelty rationale", ansoff_d.get("novelty_rationale",""))
    kv(doc,"IP risk", ansoff_d.get("ip_risk",""))
    kv(doc,"IP risk rationale", ansoff_d.get("ip_risk_rationale",""))
    sp=ansoff_d.get("schaeffler_position",{})
    kv(doc,"Schaeffler existing IP", sp.get("existing_ip",""))
    kv(doc,"IP gap addressed", sp.get("gap",""))
    if landscape.get("white_spaces"):
        h2(doc,"IP White Spaces")
        for ws in landscape["white_spaces"]: bul(doc, ws)
    if landscape.get("technology_keywords"):
        h2(doc,"Key Technology Search Terms")
        body(doc, "  ·  ".join(landscape["technology_keywords"]))
    if landscape.get("key_filers"):
        h2(doc,"Key Patent Filers")
        ft=doc.add_table(rows=1,cols=4); ft.style="Table Grid"
        for i,h in enumerate(["Company","Type","Threat","Filing Focus"]):
            c=ft.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,fi in enumerate(landscape["key_filers"]):
            row=ft.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            for c in row.cells: set_bg(c,fill)
            for j,val in enumerate([fi.get("company",""),fi.get("type",""),fi.get("threat_level",""),fi.get("focus","")]):
                r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.bold=(j==0)
        if landscape.get("landscape_summary"):
            doc.add_paragraph(); body(doc, landscape["landscape_summary"])

    # ── Stage 04 — Technical Feasibility ─────────────────────────
    h1(doc,"Stage 04 · Technical Feasibility  ·  Score: " + str(scores.get("feasibility",5)) + "/10")
    body(doc, enr.get("feasibility_highlights",""))
    doc.add_paragraph()
    kv(doc,"TRL Level", trl.get("trl_label",""))
    kv(doc,"Existence verdict", existence.get("existence_verdict",""))
    kv(doc,"Schaeffler entry readiness", trl.get("schaeffler_entry_readiness",""))
    kv(doc,"Time to production readiness", existence.get("time_to_readiness",""))
    kv(doc,"Technology core", existence.get("technology_core",""))
    if trl.get("trl_rationale"):
        doc.add_paragraph(); body(doc, trl["trl_rationale"])
    if existence.get("existence_summary"):
        body(doc, existence["existence_summary"])
    if existence.get("evidence"):
        h2(doc,"Evidence (Academic, Startups, Pilots, Programmes)")
        ev_tbl=doc.add_table(rows=1,cols=4); ev_tbl.style="Table Grid"
        for i,h in enumerate(["Type","Title & Description","Relevance","Source"]):
            c=ev_tbl.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,ev in enumerate(existence["evidence"]):
            row=ev_tbl.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            for c in row.cells: set_bg(c,fill)
            for j,val in enumerate([ev.get("type",""), ev.get("title","")+" — "+ev.get("description",""), ev.get("relevance","")+" / "+ev.get("confidence",""), ev.get("source","")]):
                r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.bold=(j==1)
    if existence.get("technology_gaps"):
        h2(doc,"Technology Gaps to Bridge")
        for gap in existence["technology_gaps"]: bul(doc, gap)
    if existence.get("keywords"):
        h2(doc,"Technology Keyword Map")
        body(doc, "  ·  ".join(existence["keywords"]))
    if trl.get("key_technical_risks"):
        h2(doc,"Key Technical Risks")
        rt=doc.add_table(rows=1,cols=3); rt.style="Table Grid"
        for i,h in enumerate(["Risk","Severity","Mitigation"]):
            c=rt.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,risk in enumerate(trl["key_technical_risks"]):
            row=rt.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            for c in row.cells: set_bg(c,fill)
            for j,val in enumerate([risk.get("risk",""),risk.get("severity",""),risk.get("mitigation","")]):
                r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.bold=(j==0)

    # ── Stage 05 — Organisational Readiness ──────────────────────
    h1(doc,"Stage 05 · Organisational Readiness  ·  Score: " + str(scores.get("org",5)) + "/10")
    body(doc, enr.get("org_highlights",""))
    doc.add_paragraph()
    kv(doc,"P³ Portfolio", f"{s5d_org.get('p_portfolio',5)}/10")
    kv(doc,"P³ People", f"{s5d_org.get('p_people',5)}/10")
    kv(doc,"P³ Process", f"{s5d_org.get('p_process',5)}/10")
    kv(doc,"Build strategy", bop.get("recommendation",""))
    kv(doc,"Time to TRL6 — Internal", bop.get("time_to_trl6_internal",""))
    kv(doc,"Time to TRL6 — With Partner", bop.get("time_to_trl6_partner",""))
    kv(doc,"Critical competency gap", people.get("competency_gap",""))
    kv(doc,"Closure route", people.get("sourcing_route",""))
    if portfolio.get("rationale"): body(doc, portfolio["rationale"])
    if people.get("rationale"): body(doc, people["rationale"])
    if process.get("rationale"): body(doc, process["rationale"])
    if process.get("applicable_assets"):
        h2(doc,"Applicable Schaeffler Assets")
        for a in process["applicable_assets"]: bul(doc, a)
    if people.get("matched_competencies"):
        h2(doc,"Matched Competencies")
        for c in people["matched_competencies"]: bul(doc, f"✓  {c}")
    if org_data.get("org_gaps"):
        h2(doc,"Organisational Gaps")
        gt=doc.add_table(rows=1,cols=4); gt.style="Table Grid"
        for i,h in enumerate(["Gap","Severity","Closure Route","Timeline"]):
            c=gt.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,g in enumerate(org_data["org_gaps"]):
            row=gt.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            for c in row.cells: set_bg(c,fill)
            for j,val in enumerate([g.get("gap",""),g.get("severity",""),g.get("closure_route",""),g.get("timeline","")]):
                r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.bold=(j==0)
    if org_data.get("partnership_candidates"):
        h2(doc,"Partnership Candidates")
        pt=doc.add_table(rows=1,cols=4); pt.style="Table Grid"
        for i,h in enumerate(["Organisation","Type","Rationale","Route"]):
            c=pt.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,p in enumerate(org_data["partnership_candidates"]):
            row=pt.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            for c in row.cells: set_bg(c,fill)
            for j,val in enumerate([p.get("name",""),p.get("type",""),p.get("rationale",""),p.get("route","")]):
                r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.bold=(j==0)

    # ── Integrated Risk Register ──────────────────────────────────
    h1(doc,"Integrated Risk Register")
    if enr.get("risk_synthesis"):
        rt2=doc.add_table(rows=1,cols=2); rt2.style="Table Grid"
        for i,h in enumerate(["Risk","Description & Mitigation"]):
            c=rt2.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,risk in enumerate(enr["risk_synthesis"]):
            row=rt2.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            for c in row.cells: set_bg(c,fill)
            parts=str(risk).split(" — ",1) if " — " in str(risk) else [str(risk),"See full analysis"]
            r0=row.cells[0].paragraphs[0].add_run(parts[0]); r0.font.size=Pt(9.5); r0.bold=True
            r1=row.cells[1].paragraphs[0].add_run(parts[1] if len(parts)>1 else ""); r1.font.size=Pt(9.5)

    # ── Action Plan ───────────────────────────────────────────────
    h1(doc,"Recommended Action Plan")
    if enr.get("action_plan"):
        for i, step in enumerate(enr["action_plan"], 1):
            p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(3); p.paragraph_format.space_after=Pt(3)
            r=p.add_run(f"{i}.  {step}"); r.font.size=Pt(10.5)
    if synthesis.get("next_steps"):
        h2(doc,"Additional Next Steps")
        for step in synthesis["next_steps"]: bul(doc, step)

    # ── Quadrant classification summary ──────────────────────────
    h1(doc,"Stage 01 · Innovation Classification")
    kv(doc,"Quadrant", quadrant)
    kv(doc,"Technology level", s1c.get("technology_level",""))
    kv(doc,"Market level", s1c.get("market_level",""))
    kv(doc,"Technology novelty", s1c.get("technology_novelty",""))
    kv(doc,"Market position", s1c.get("market_position",""))
    kv(doc,"Innovation cluster", s1c.get("innovation_cluster",""))
    kv(doc,"Product family", s1c.get("product_family",""))
    if s1c.get("trend_alignment"):
        kv(doc,"Trend alignment", ", ".join(s1c.get("trend_alignment",[])))
    kv(doc,"Project type", s1c.get("project_type",""))
    kv(doc,"Innovation model", s1c.get("innovation_model",""))
    kv(doc,"Reasoning", s1c.get("reasoning",""))

    # ── Footer ────────────────────────────────────────────────────
    doc.add_paragraph()
    ft=doc.add_table(rows=1,cols=1); ft.style="Table Grid"; fc=ft.cell(0,0); set_bg(fc,"1F3864")
    fp=fc.paragraphs[0]; fp.paragraph_format.space_before=Pt(6); fp.paragraph_format.space_after=Pt(6)
    fr=fp.add_run(f"Schaeffler AI Innovation Research Assistant  ·  Full Innovation Assessment  ·  {datetime.now().strftime('%d %B %Y')}  ·  Capstone Project — Arpan Chowdhury, EBS Universität")
    fr.font.size=Pt(8); fr.font.color.rgb=RGBColor(0x93,0xC5,0xFD)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf
def generate_feasibility_report(idea, quadrant, s1c, existence, trl, scores):
    """Generate a Technical Feasibility Word report."""
    feas_ctx = (
        f"Idea: {idea}\nQuadrant: {quadrant}\n"
        f"TRL level: {trl.get('trl_level',3)} — {trl.get('trl_label','')}\n"
        f"TRL rationale: {trl.get('trl_rationale','')}\n"
        f"Existence verdict: {existence.get('existence_verdict','')}\n"
        f"Existence summary: {existence.get('existence_summary','')}\n"
        f"Schaeffler entry readiness: {trl.get('schaeffler_entry_readiness','')}\n"
        f"Time to production readiness: {existence.get('time_to_readiness','')}\n"
        f"Technology core: {existence.get('technology_core','')}\n"
        f"Technology gaps: {chr(59).join(existence.get('technology_gaps',[])[:4])}\n"
        f"Key evidence: {chr(59).join([e.get('title','') + ' (' + e.get('source','') + ')' for e in existence.get('evidence',[])[:4]])}\n"
        f"Analogous Schaeffler tech: {trl.get('analogous_schaeffler_technologies','') or trl.get('analogous_schaeffler_tech','')}\n"
        f"Key technical risks: {chr(59).join([r.get('risk','') for r in trl.get('key_technical_risks',[])[:3]])}\n"
        f"Score: {scores['final_score']}/10"
    )
    extended = call_claude(
        '''You are a Schaeffler R&D director writing a detailed technical feasibility assessment.
Write specific, evidence-based content. Return ONLY valid JSON, no markdown backticks:
{
  "executive_summary": "3-4 full paragraphs: overall feasibility verdict, TRL assessment rationale, key evidence found, and what Schaeffler needs to do to advance the technology.",
  "technology_analysis": "3-4 full paragraphs on the core technology mechanism, current state of the art globally, what has been demonstrated vs what remains theoretical, and the critical technical gaps to bridge.",
  "schaeffler_readiness": "2-3 full paragraphs on Schaeffler-specific technical readiness — which existing competencies (bearings, mechatronics, Vitesco power electronics, tribology) are applicable, what is genuinely missing, and how analogous it is to existing Schaeffler product lines.",
  "development_pathway": "2-3 full paragraphs: specific recommended development pathway from current TRL to TRL 6 and beyond — what milestones, what resources, what partnerships, and realistic timelines.",
  "risks": ["Technical risk 1 with specific mitigation approach", "Technical risk 2 with mitigation", "Technical risk 3", "Technical risk 4"],
  "recommendations": ["Concrete R&D action 1 with timeline", "Concrete action 2", "Concrete action 3", "Concrete action 4"]
}''',
        feas_ctx,
        max_tokens=3500
    )
    raw_e = extended.strip().replace("```json","").replace("```","").strip()
    fb = raw_e.find("{"); lb = raw_e.rfind("}") + 1
    if fb >= 0: raw_e = raw_e[fb:lb]
    try:
        ext = json.loads(raw_e)
    except:
        ext = {
            "executive_summary": f"This innovation idea is assessed at TRL {trl.get('trl_level',3)} — {trl.get('trl_label','')}. Existence verdict: {existence.get('existence_verdict','Research Stage')}. {existence.get('existence_summary','')} Schaeffler entry readiness: {trl.get('schaeffler_entry_readiness','Ready for Innovation')}. Estimated time to production readiness: {existence.get('time_to_readiness','3-5 years')}.",
            "technology_analysis": f"Core technology: {existence.get('technology_core','')}. TRL rationale: {trl.get('trl_rationale','')} Technology gaps to bridge: {chr(44).join(existence.get('technology_gaps',[])[:3])}.",
            "schaeffler_readiness": f"Analogous Schaeffler technology: {trl.get('analogous_schaeffler_technologies','') or trl.get('analogous_schaeffler_tech','')}. {trl.get('entry_rationale','')}",
            "development_pathway": f"From current TRL {trl.get('trl_level',3)}, the recommended development pathway progresses through lab validation, prototype testing in relevant environment, and pilot deployment. Estimated timeline: {existence.get('time_to_readiness','3-5 years')}.",
            "risks": [r.get('risk','') + ' — ' + r.get('mitigation','') for r in trl.get('key_technical_risks',[])[:4]] or ["Technology gaps require systematic R&D investment", "Manufacturing scalability not yet demonstrated", "Integration with existing Schaeffler systems needs validation"],
            "recommendations": ["Commission proof-of-concept study with Schaeffler R&D within 60 days", f"Target TRL {min(trl.get('trl_level',3)+2, 6)} within 18 months via innovation project", "Identify university or research institute partner for early-stage development", "Conduct internal engineering workshop to assess competency gaps"]
        }

    NAVY=RGBColor(0x1F,0x38,0x64); WHITE=RGBColor(0xFF,0xFF,0xFF)
    GREY=RGBColor(0x55,0x55,0x55); BLACK=RGBColor(0x00,0x00,0x00)
    LBLUE=RGBColor(0x60,0xA5,0xFA)

    def set_bg(cell, hx):
        tc=cell._tc; tcPr=tc.get_or_add_tcPr(); shd=OxmlElement("w:shd")
        shd.set(qn("w:val"),"clear"); shd.set(qn("w:color"),"auto"); shd.set(qn("w:fill"),hx); tcPr.append(shd)

    def h1(doc, text):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(14); p.paragraph_format.space_after=Pt(4)
        pPr=p._p.get_or_add_pPr(); pBdr=OxmlElement("w:pBdr"); bot=OxmlElement("w:bottom")
        bot.set(qn("w:val"),"single"); bot.set(qn("w:sz"),"8"); bot.set(qn("w:space"),"3"); bot.set(qn("w:color"),"2E75B6")
        pBdr.append(bot); pPr.append(pBdr); r=p.add_run(text)
        r.bold=True; r.font.size=Pt(14); r.font.color.rgb=NAVY

    def body(doc, text):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(3); p.paragraph_format.space_after=Pt(3)
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; r=p.add_run(text); r.font.size=Pt(10.5)

    def kv(doc, label, value):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
        r1=p.add_run(f"{label}: "); r1.bold=True; r1.font.size=Pt(10.5); r1.font.color.rgb=NAVY
        r2=p.add_run(value); r2.font.size=Pt(10.5)

    def bul(doc, text):
        p=doc.add_paragraph(style="List Bullet"); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
        r=p.add_run(text); r.font.size=Pt(10.5)

    doc=DocxDocument()
    for sec in doc.sections:
        sec.top_margin=Cm(2.0); sec.bottom_margin=Cm(2.0); sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    # Header
    t=doc.add_table(rows=1,cols=1); t.style="Table Grid"; c=t.cell(0,0); set_bg(c,"1F3864")
    p=c.paragraphs[0]; p.paragraph_format.space_before=Pt(10); p.paragraph_format.space_after=Pt(2)
    r=p.add_run("TECHNICAL FEASIBILITY REPORT"); r.bold=True; r.font.size=Pt(9); r.font.color.rgb=WHITE
    p2=c.add_paragraph(); p2.paragraph_format.space_before=Pt(0); p2.paragraph_format.space_after=Pt(10)
    r2=p2.add_run("Schaeffler AI Innovation Research Assistant  ·  Stage 04"); r2.font.size=Pt(9); r2.font.color.rgb=RGBColor(0x93,0xC5,0xFD)
    doc.add_paragraph()

    # Title
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2)
    r=p.add_run("Technical Feasibility Assessment"); r.bold=True; r.font.size=Pt(18); r.font.color.rgb=NAVY
    p2=doc.add_paragraph()
    r2=p2.add_run(f"Score: {scores['final_score']}/10  ·  TRL {trl.get('trl_level','')}  ·  {datetime.now().strftime('%d %B %Y')}")
    r2.font.size=Pt(10); r2.italic=True; r2.font.color.rgb=GREY

    # Idea box
    doc.add_paragraph()
    tb=doc.add_table(rows=1,cols=2); tb.style="Table Grid"
    c1=tb.cell(0,0); c2=tb.cell(0,1); set_bg(c1,"1F3864"); set_bg(c2,"EAF1FB"); c1.width=Inches(0.12)
    c1.paragraphs[0].add_run("")
    rp=c2.paragraphs[0]; rp.paragraph_format.space_before=Pt(8); rp.paragraph_format.space_after=Pt(2)
    rb=rp.add_run("Innovation Idea"); rb.bold=True; rb.font.size=Pt(9); rb.font.color.rgb=NAVY
    rp2=c2.add_paragraph(); rp2.paragraph_format.space_before=Pt(0); rp2.paragraph_format.space_after=Pt(8)
    ri=rp2.add_run(idea); ri.font.size=Pt(10); ri.italic=True
    doc.add_paragraph()

    # Score summary
    h1(doc,"Score Summary")
    st2=doc.add_table(rows=1,cols=3); st2.style="Table Grid"
    for i,h in enumerate(["Dimension","Score","Weight"]):
        c=st2.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
        r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
    for i,(dim,score,wt) in enumerate([
        ("TRL Score",f"{scores['trl_score']:.1f}/10","50%"),
        ("Existence Quality",f"{scores['existence_score']:.1f}/10","30%"),
        ("Risk Profile",f"{scores['risk_score']:.1f}/10","20%"),
    ]):
        row=st2.add_row(); fill="EAF1FB" if i%2==0 else "FFFFFF"
        for c in row.cells: set_bg(c,fill)
        for j,val in enumerate([dim,score,wt]):
            r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(10); r.bold=(j==0)
    fr=st2.add_row()
    for c in fr.cells: set_bg(c,"1F3864")
    r=fr.cells[0].paragraphs[0].add_run("FINAL SCORE"); r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
    r2=fr.cells[2].paragraphs[0].add_run(f"{scores['final_score']}/10"); r2.bold=True; r2.font.size=Pt(11); r2.font.color.rgb=LBLUE
    doc.add_paragraph()

    # Executive summary
    h1(doc,"Executive Summary"); body(doc,ext.get("executive_summary",""))

    # TRL assessment
    h1(doc,"Technology Readiness Level Assessment")
    kv(doc,"TRL Level",trl.get("trl_label",""))
    kv(doc,"Existence verdict",existence.get("existence_verdict",""))
    kv(doc,"Schaeffler entry readiness",trl.get("schaeffler_entry_readiness",""))
    kv(doc,"Time to production readiness",existence.get("time_to_readiness",""))
    doc.add_paragraph(); body(doc,trl.get("trl_rationale",""))

    # TRL scale table
    doc.add_paragraph()
    trl_tbl=doc.add_table(rows=1,cols=3); trl_tbl.style="Table Grid"
    for i,h in enumerate(["TRL","Level","Description"]):
        c=trl_tbl.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
        r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
    trl_rows=[
        (1,"Basic principles observed","Theoretical concept only"),
        (2,"Technology concept formulated","Application identified, no testing"),
        (3,"Experimental proof of concept","Lab demonstration, key functions validated"),
        (4,"Technology validated in lab","Component tested, controlled environment"),
        (5,"Validated in relevant environment","Prototype tested, industrial-like conditions"),
        (6,"Demonstrated in relevant environment","System prototype demonstrated"),
        (7,"Prototype in operational environment","Field trial or industrial pilot"),
        (8,"System complete and qualified","Limited production run"),
        (9,"Proven in operational environment","Commercial deployment at scale"),
    ]
    current_trl = trl.get("trl_level",3)
    for lvl,lbl,desc in trl_rows:
        row=trl_tbl.add_row()
        fill="EAF5EA" if lvl==current_trl else ("EAF1FB" if lvl%2==0 else "FFFFFF")
        for c in row.cells: set_bg(c,fill)
        r0=row.cells[0].paragraphs[0].add_run(f"TRL {lvl}"); r0.font.size=Pt(10); r0.bold=(lvl==current_trl)
        r1=row.cells[1].paragraphs[0].add_run(lbl); r1.font.size=Pt(10); r1.bold=(lvl==current_trl)
        r2=row.cells[2].paragraphs[0].add_run(desc); r2.font.size=Pt(9.5)

    # Technology analysis
    h1(doc,"Technology Analysis"); body(doc,ext.get("technology_analysis",""))

    # Evidence table
    h1(doc,"Existing Evidence")
    body(doc,existence.get("existence_summary",""))
    evidence=existence.get("evidence",[])
    if evidence:
        doc.add_paragraph()
        ev_tbl=doc.add_table(rows=1,cols=4); ev_tbl.style="Table Grid"
        for i,h in enumerate(["Type","Title","Relevance","Source"]):
            c=ev_tbl.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,ev in enumerate(evidence):
            row=ev_tbl.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            for c in row.cells: set_bg(c,fill)
            for j,val in enumerate([ev.get("type",""),ev.get("title",""),ev.get("relevance",""),ev.get("source","")]):
                r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.bold=(j==1)
        # Description column
        doc.add_paragraph()
        for ev in evidence:
            p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
            r1=p.add_run(f"{ev.get('title','')}: "); r1.bold=True; r1.font.size=Pt(10)
            r2=p.add_run(ev.get("description","")); r2.font.size=Pt(10)

    # Technology gaps
    if existence.get("technology_gaps"):
        h1(doc,"Technology Gaps to Bridge")
        for gap in existence["technology_gaps"]: bul(doc,gap)

    # Technical risks
    h1(doc,"Key Technical Risks")
    risks=trl.get("key_technical_risks",[])
    if risks:
        rt=doc.add_table(rows=1,cols=3); rt.style="Table Grid"
        for i,h in enumerate(["Risk","Severity","Mitigation"]):
            c=rt.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,risk in enumerate(risks):
            row=rt.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            for c in row.cells: set_bg(c,fill)
            for j,val in enumerate([risk.get("risk",""),risk.get("severity",""),risk.get("mitigation","")]):
                r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.bold=(j==0)

    # Schaeffler readiness & development pathway
    h1(doc,"Schaeffler Readiness"); body(doc,ext.get("schaeffler_readiness",""))
    h1(doc,"Development Pathway"); body(doc,ext.get("development_pathway",""))
    h1(doc,"Recommendations")
    for rec in ext.get("recommendations",[]): bul(doc,rec)

    # Footer
    doc.add_paragraph()
    ft=doc.add_table(rows=1,cols=1); ft.style="Table Grid"; fc=ft.cell(0,0); set_bg(fc,"1F3864")
    fp=fc.paragraphs[0]; fp.paragraph_format.space_before=Pt(6); fp.paragraph_format.space_after=Pt(6)
    fr=fp.add_run(f"Schaeffler AI Innovation Research Assistant  ·  Stage 04: Technical Feasibility  ·  {datetime.now().strftime('%d %B %Y')}  ·  Capstone Project — Arpan Chowdhury, EBS Universität")
    fr.font.size=Pt(8); fr.font.color.rgb=RGBColor(0x93,0xC5,0xFD)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf


def generate_patent_report(idea, quadrant, s1c, landscape, ansoff_data, scores):
    """Generate a Patent Intelligence Word report."""
    pat_ctx = (
        f"Idea: {idea}\nQuadrant: {quadrant}\n"
        f"Activity level: {landscape.get('activity_level','')}  Trend: {landscape.get('filing_trend','')}\n"
        f"Filing trend rationale: {landscape.get('filing_trend_rationale','')}\n"
        f"Novelty signal: {ansoff_data.get('novelty_signal','')}  IP risk: {ansoff_data.get('ip_risk','')}\n"
        f"Novelty rationale: {ansoff_data.get('novelty_rationale','')}\n"
        f"IP risk rationale: {ansoff_data.get('ip_risk_rationale','')}\n"
        f"Key filers: {chr(44).join([f.get('company','') for f in landscape.get('key_filers',[])[:6]])}\n"
        f"White spaces: {chr(59).join(landscape.get('white_spaces',[]))}\n"
        f"Schaeffler existing IP: {ansoff_data.get('schaeffler_position',{}).get('existing_ip','')}\n"
        f"IP gap addressed: {ansoff_data.get('schaeffler_position',{}).get('gap','')}\n"
        f"Landscape summary: {landscape.get('landscape_summary','')}\nScore: {scores['final_score']}/10"
    )
    extended = call_claude(
        '''You are a senior patent intelligence analyst writing a detailed IP report for Schaeffler Group.
Write specific, actionable content based on the data provided. Return ONLY valid JSON, no markdown backticks:
{
  "executive_summary": "3-4 full paragraphs: overall IP landscape assessment, novelty of the idea, key risks and opportunities, and headline recommendation for Schaeffler IP strategy.",
  "landscape_analysis": "3-4 full paragraphs covering filing activity levels, who is filing and why, filing trends and what they indicate, which quadrant of IP activity is most active, and what it means for new entrants.",
  "ip_strategy": "2-3 full paragraphs: specific recommended IP strategy for Schaeffler — whether to file broadly or narrowly, which white spaces to target, which claims to prioritise, and how to defend against existing filers.",
  "risks": ["Specific IP risk 1 with mitigation strategy", "Specific risk 2 with action", "Specific risk 3 with action", "Specific risk 4"],
  "recommendations": ["Concrete IP action 1 with timeline", "Concrete action 2", "Concrete action 3", "Concrete action 4"]
}''',
        pat_ctx,
        max_tokens=3500
    )
    raw_e = extended.strip().replace("```json","").replace("```","").strip()
    fb = raw_e.find("{"); lb = raw_e.rfind("}") + 1
    if fb >= 0: raw_e = raw_e[fb:lb]
    try:
        ext = json.loads(raw_e)
    except:
        sp = ansoff_data.get('schaeffler_position',{})
        ext = {
            "executive_summary": f"The patent landscape for this innovation shows {landscape.get('activity_level','moderate')} activity with a {landscape.get('filing_trend','stable')} trend. Novelty signal is {ansoff_data.get('novelty_signal','Moderate')} and IP risk is assessed as {ansoff_data.get('ip_risk','Medium')}. {landscape.get('landscape_summary','')} Schaeffler's existing IP: {sp.get('existing_ip','see analysis')}.",
            "landscape_analysis": f"Filing trend is {landscape.get('filing_trend','stable')}. {landscape.get('filing_trend_rationale','')} Key filers include {chr(44).join([f.get('company','') for f in landscape.get('key_filers',[])[:4]])}. White spaces identified: {chr(59).join(landscape.get('white_spaces',[])[:3])}.",
            "ip_strategy": f"Schaeffler should target the identified IP white spaces to establish a defensible position. The gap this idea addresses: {sp.get('gap','see analysis')}. Novelty rationale: {ansoff_data.get('novelty_rationale','')}",
            "risks": [f"IP risk level {ansoff_data.get('ip_risk','Medium')}: {ansoff_data.get('ip_risk_rationale','conduct FTO analysis before R&D commitment')}", "Freedom-to-operate analysis required before external disclosure", "Monitor filing activity of identified key players for blocking patents", "Defensive filing strategy needed to protect white space positions"],
            "recommendations": ["Commission formal FTO (freedom-to-operate) analysis within 60 days", "File provisional patent applications in identified white spaces", "Monitor competitor filing activity via patent watch service", "Engage Schaeffler patent counsel to assess portfolio gap"]
        }

    NAVY  = RGBColor(0x1F,0x38,0x64); BLUE = RGBColor(0x2E,0x75,0xB6)
    WHITE = RGBColor(0xFF,0xFF,0xFF); GREY = RGBColor(0x55,0x55,0x55)
    BLACK = RGBColor(0x00,0x00,0x00); LBLUE= RGBColor(0x60,0xA5,0xFA)

    def set_bg(cell, hex_col):
        tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        shd=OxmlElement("w:shd"); shd.set(qn("w:val"),"clear")
        shd.set(qn("w:color"),"auto"); shd.set(qn("w:fill"),hex_col); tcPr.append(shd)

    def h1(doc, text):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(14); p.paragraph_format.space_after=Pt(4)
        pPr=p._p.get_or_add_pPr(); pBdr=OxmlElement("w:pBdr"); bot=OxmlElement("w:bottom")
        bot.set(qn("w:val"),"single"); bot.set(qn("w:sz"),"8"); bot.set(qn("w:space"),"3"); bot.set(qn("w:color"),"2E75B6")
        pBdr.append(bot); pPr.append(pBdr)
        r=p.add_run(text); r.bold=True; r.font.size=Pt(14); r.font.color.rgb=NAVY

    def body(doc, text):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(3); p.paragraph_format.space_after=Pt(3)
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        r=p.add_run(text); r.font.size=Pt(10.5)

    def kv(doc, label, value):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
        r1=p.add_run(f"{label}: "); r1.bold=True; r1.font.size=Pt(10.5); r1.font.color.rgb=NAVY
        r2=p.add_run(value); r2.font.size=Pt(10.5)

    def bul(doc, text):
        p=doc.add_paragraph(style="List Bullet"); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
        r=p.add_run(text); r.font.size=Pt(10.5)

    doc = DocxDocument()
    for sec in doc.sections:
        sec.top_margin=Cm(2.0); sec.bottom_margin=Cm(2.0); sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    # Header
    t=doc.add_table(rows=1,cols=1); t.style="Table Grid"; c=t.cell(0,0); set_bg(c,"1F3864")
    p=c.paragraphs[0]; p.paragraph_format.space_before=Pt(10); p.paragraph_format.space_after=Pt(2)
    r=p.add_run("PATENT INTELLIGENCE REPORT"); r.bold=True; r.font.size=Pt(9); r.font.color.rgb=WHITE
    p2=c.add_paragraph(); p2.paragraph_format.space_before=Pt(0); p2.paragraph_format.space_after=Pt(10)
    r2=p2.add_run("Schaeffler AI Innovation Research Assistant  ·  Stage 03"); r2.font.size=Pt(9); r2.font.color.rgb=RGBColor(0x93,0xC5,0xFD)
    doc.add_paragraph()

    # Title
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2)
    r=p.add_run("Patent Intelligence Analysis"); r.bold=True; r.font.size=Pt(18); r.font.color.rgb=NAVY
    p2=doc.add_paragraph()
    r2=p2.add_run(f"Score: {scores['final_score']}/10  ·  Quadrant: {quadrant}  ·  {datetime.now().strftime('%d %B %Y')}")
    r2.font.size=Pt(10); r2.italic=True; r2.font.color.rgb=GREY

    # Idea box
    doc.add_paragraph()
    tb=doc.add_table(rows=1,cols=2); tb.style="Table Grid"
    c1=tb.cell(0,0); c2=tb.cell(0,1); set_bg(c1,"1F3864"); set_bg(c2,"EAF1FB"); c1.width=Inches(0.12)
    c1.paragraphs[0].add_run("")
    rp=c2.paragraphs[0]; rp.paragraph_format.space_before=Pt(8); rp.paragraph_format.space_after=Pt(2)
    rb=rp.add_run("Innovation Idea"); rb.bold=True; rb.font.size=Pt(9); rb.font.color.rgb=NAVY
    rp2=c2.add_paragraph(); rp2.paragraph_format.space_before=Pt(0); rp2.paragraph_format.space_after=Pt(8)
    ri=rp2.add_run(idea); ri.font.size=Pt(10); ri.italic=True
    doc.add_paragraph()

    # Score summary
    h1(doc,"Score Summary")
    st2=doc.add_table(rows=1,cols=3); st2.style="Table Grid"
    for i,h in enumerate(["Dimension","Score","Weight"]):
        c=st2.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
        r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
    for i,(dim,score,wt) in enumerate([
        ("Landscape Openness",f"{scores['landscape_score']:.1f}/10","40%"),
        ("Novelty Signal",f"{scores['novelty_score']:.1f}/10","35%"),
        ("IP Risk (inverted)",f"{scores['ip_score']:.1f}/10","25%"),
    ]):
        row=st2.add_row(); fill="EAF1FB" if i%2==0 else "FFFFFF"
        for j,c in enumerate(row.cells): set_bg(c,fill)
        for j,val in enumerate([dim,score,wt]):
            r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(10); r.bold=(j==0)
    fr=st2.add_row()
    for c in fr.cells: set_bg(c,"1F3864")
    r=fr.cells[0].paragraphs[0].add_run("FINAL PATENT INTELLIGENCE SCORE"); r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
    r2=fr.cells[2].paragraphs[0].add_run(f"{scores['final_score']}/10"); r2.bold=True; r2.font.size=Pt(11); r2.font.color.rgb=LBLUE
    doc.add_paragraph()

    # Executive summary
    h1(doc,"Executive Summary"); body(doc,ext.get("executive_summary",""))

    # Landscape overview
    h1(doc,"Patent Landscape Overview")
    kv(doc,"Activity level",landscape.get("activity_level",""))
    kv(doc,"Filing trend",landscape.get("filing_trend","") + " — " + landscape.get("filing_trend_rationale",""))
    kv(doc,"Novelty signal",ansoff_data.get("novelty_signal","") + " — " + ansoff_data.get("novelty_rationale",""))
    kv(doc,"IP risk",ansoff_data.get("ip_risk","") + " — " + ansoff_data.get("ip_risk_rationale",""))
    if landscape.get("technology_keywords"):
        kv(doc,"Key technology terms",", ".join(landscape["technology_keywords"]))
    doc.add_paragraph(); body(doc,ext.get("landscape_analysis",""))

    # Key filers table
    h1(doc,"Key Patent Filers")
    filers=landscape.get("key_filers",[])
    if filers:
        ft=doc.add_table(rows=1,cols=4); ft.style="Table Grid"
        for i,h in enumerate(["Company","Type","Threat","Focus"]):
            c=ft.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,fi in enumerate(filers):
            row=ft.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            for c in row.cells: set_bg(c,fill)
            for j,val in enumerate([fi.get("company",""),fi.get("type",""),fi.get("threat_level",""),fi.get("focus","")+" "+fi.get("source","")]):
                r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.bold=(j==0)

    # Schaeffler IP position
    h1(doc,"Schaeffler IP Position")
    sp=ansoff_data.get("schaeffler_position",{})
    kv(doc,"Existing IP",sp.get("existing_ip",""))
    kv(doc,"IP gap addressed",sp.get("gap",""))
    if landscape.get("white_spaces"):
        doc.add_paragraph()
        p=doc.add_paragraph(); r=p.add_run("IP White Spaces"); r.bold=True; r.font.size=Pt(11); r.font.color.rgb=NAVY
        for ws in landscape["white_spaces"]: bul(doc,ws)
    doc.add_paragraph(); body(doc,ext.get("ip_strategy",""))

    # Risks & recommendations
    h1(doc,"IP Risks & Mitigations")
    for risk in ext.get("risks",[]): bul(doc,risk)
    h1(doc,"Recommendations")
    for rec in ext.get("recommendations",[]): bul(doc,rec)

    # Footer
    doc.add_paragraph()
    ft2=doc.add_table(rows=1,cols=1); ft2.style="Table Grid"; fc=ft2.cell(0,0); set_bg(fc,"1F3864")
    fp=fc.paragraphs[0]; fp.paragraph_format.space_before=Pt(6); fp.paragraph_format.space_after=Pt(6)
    fr=fp.add_run(f"Schaeffler AI Innovation Research Assistant  ·  Stage 03: Patent Intelligence  ·  {datetime.now().strftime('%d %B %Y')}  ·  Capstone Project — Arpan Chowdhury, EBS Universität")
    fr.font.size=Pt(8); fr.font.color.rgb=RGBColor(0x93,0xC5,0xFD)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf


def generate_market_report(idea, quadrant, s1c, market, comp, sectors, weights, final_score):
    """Generate a formatted Word document market intelligence report."""
    mkt_ctx = (
        f"Idea: {idea}\nQuadrant: {quadrant}\nMarket: {market.get('market_name','')}\n"
        f"Size 2024: {_mval(market.get('market_size_current') or market.get('market_size_2024',''))}  Size 2030: {_mval(market.get('market_size_forecast') or market.get('market_size_2030',''))}\n"
        f"CAGR: {_mval(market.get('cagr',''))}  Maturity: {market.get('market_maturity','')}\n"
        f"Geography: {market.get('geographic_focus','')}\n"
        f"Growth drivers: {chr(59).join(market.get('growth_drivers',[])[:4])}\n"
        f"Competitive intensity: {comp.get('competitive_intensity','')}  White space: {comp.get('white_space','')}\n"
        f"Schaeffler advantage: {comp.get('schaeffler_advantage','')}\n"
        f"Key competitors: {', '.join([c.get('name','') for c in comp.get('competitors',[])[:6]])}\n"
        f"Primary sectors: {', '.join(sectors.get('primary_sectors',[]))}\n"
        f"Sector fit rationale: {sectors.get('sector_fit_rationale','')}\nFinal score: {final_score}/10"
    )
    extended = call_claude(
        '''You are a senior market analyst writing a comprehensive market intelligence report for Schaeffler Group.
Write substantive, specific content using the data provided. Return ONLY valid JSON, no markdown backticks:
{
  "executive_summary": "3-4 full paragraphs summarising market opportunity, competitive position, and strategic fit. Reference exact figures provided.",
  "market_deep_dive": "3-4 full paragraphs covering market dynamics, demand signals, geographic breakdown, growth trajectory and 5-year outlook.",
  "competitive_analysis": "3-4 full paragraphs on key players, their strategies, where white space exists, and how Schaeffler is positioned.",
  "schaeffler_fit": "2-3 full paragraphs on strategic fit — reference Schaeffler electrification agenda, Vitesco merger, E-Mobility growth, and P3 portfolio formula.",
  "risks": ["Specific risk 1 with concrete mitigation", "Specific risk 2 with mitigation", "Specific risk 3 with mitigation", "Specific risk 4"],
  "recommendations": ["Concrete action 1 with timeline", "Concrete action 2", "Concrete action 3", "Concrete action 4"]
}''',
        mkt_ctx,
        max_tokens=3500
    )
    raw_e = extended.strip().replace("```json","").replace("```","").strip()
    fb = raw_e.find("{"); lb = raw_e.rfind("}") + 1
    if fb >= 0: raw_e = raw_e[fb:lb]
    try:
        ext = json.loads(raw_e)
    except:
        ext = {
            "executive_summary": f"The {market.get('market_name','target market')} represents a {market.get('market_maturity','growing')} opportunity. Market size is estimated at {market.get('market_size_2024','significant')} in 2024, projected to reach {market.get('market_size_2030','substantial')} by 2030 at a CAGR of {market.get('cagr','strong growth')}. Competitive intensity is {comp.get('competitive_intensity','moderate')}, with white space identified: {comp.get('white_space','see analysis')}.",
            "market_deep_dive": f"The market is characterised by {market.get('market_maturity','growth')} dynamics. Key growth drivers include: {chr(10).join(market.get('growth_drivers',['strong demand','digital transformation','electrification'])[:3])}. Geographic focus: {market.get('geographic_focus','global')}.",
            "competitive_analysis": f"The competitive landscape shows {comp.get('competitive_intensity','moderate')} intensity. Schaeffler advantage: {comp.get('schaeffler_advantage','precision engineering and OEM relationships')}. White space: {comp.get('white_space','niche segments underserved')}.",
            "schaeffler_fit": f"Primary sector fit is strongest in {chr(44).join(sectors.get('primary_sectors',[]))}, which aligns with Schaeffler's post-Vitesco portfolio strategy. {sectors.get('sector_fit_rationale','See sector analysis for detail.')}",
            "risks": ["Competitive intensity requires freedom-to-operate analysis before R&D commitment — conduct FTO within 60 days", "Market timing risk: validate demand with target OEMs before scaling investment", "Sector fit assumptions require validation with Schaeffler divisional teams", "IP landscape may restrict entry — patent clearance essential"],
            "recommendations": ["Commission internal engineering feasibility review within 30 days", "Conduct FTO IP analysis with Schaeffler patent counsel", f"Identify pilot customer in {chr(44).join(sectors.get('primary_sectors',['target sector'])[:1])} for co-development conversation", "Present to Innovation steering committee with this report as supporting material"]
        }

    NAVY  = RGBColor(0x1F,0x38,0x64)
    BLUE  = RGBColor(0x2E,0x75,0xB6)
    WHITE = RGBColor(0xFF,0xFF,0xFF)
    GREY  = RGBColor(0x55,0x55,0x55)
    BLACK = RGBColor(0x00,0x00,0x00)
    LBLUE = RGBColor(0x60,0xA5,0xFA)

    def set_bg(cell, hex_col):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"),"clear")
        shd.set(qn("w:color"),"auto")
        shd.set(qn("w:fill"), hex_col)
        tcPr.append(shd)

    def h1(doc, text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after  = Pt(4)
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bot = OxmlElement("w:bottom")
        bot.set(qn("w:val"),"single"); bot.set(qn("w:sz"),"8")
        bot.set(qn("w:space"),"3"); bot.set(qn("w:color"),"2E75B6")
        pBdr.append(bot); pPr.append(pBdr)
        r = p.add_run(text)
        r.bold=True; r.font.size=Pt(14); r.font.color.rgb=NAVY

    def h2(doc, text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after  = Pt(2)
        r = p.add_run(text)
        r.bold=True; r.font.size=Pt(11); r.font.color.rgb=NAVY

    def body(doc, text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after  = Pt(3)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r = p.add_run(text)
        r.font.size=Pt(10.5)

    def kv(doc, label, value):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(2)
        r1 = p.add_run(f"{label}: ")
        r1.bold=True; r1.font.size=Pt(10.5); r1.font.color.rgb=NAVY
        r2 = p.add_run(value)
        r2.font.size=Pt(10.5)

    def bul(doc, text):
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
        r = p.add_run(text); r.font.size=Pt(10.5)

    def hdr_row(tbl, headers):
        row = tbl.rows[0]
        for i, h in enumerate(headers):
            c = row.cells[i]; set_bg(c,"1F3864")
            r = c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE

    doc = DocxDocument()
    for sec in doc.sections:
        sec.top_margin=Cm(2.0); sec.bottom_margin=Cm(2.0)
        sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    # ── Header banner ─────────────────────────────────────────
    t = doc.add_table(rows=1,cols=1); t.style="Table Grid"
    c = t.cell(0,0); set_bg(c,"1F3864")
    p = c.paragraphs[0]
    p.paragraph_format.space_before=Pt(10); p.paragraph_format.space_after=Pt(2)
    r = p.add_run("MARKET INTELLIGENCE REPORT")
    r.bold=True; r.font.size=Pt(9); r.font.color.rgb=WHITE
    p2 = c.add_paragraph()
    p2.paragraph_format.space_before=Pt(0); p2.paragraph_format.space_after=Pt(10)
    r2 = p2.add_run("Schaeffler AI Innovation Research Assistant  ·  Stage 02")
    r2.font.size=Pt(9); r2.font.color.rgb=RGBColor(0x93,0xC5,0xFD)
    doc.add_paragraph()

    # ── Title ─────────────────────────────────────────────────
    p = doc.add_paragraph()
    p.paragraph_format.space_after=Pt(2)
    r = p.add_run(market.get("market_name", idea[:80]))
    r.bold=True; r.font.size=Pt(18); r.font.color.rgb=NAVY
    p2 = doc.add_paragraph()
    r2 = p2.add_run(f"Score: {final_score}/10  ·  Quadrant: {quadrant}  ·  {datetime.now().strftime('%d %B %Y')}")
    r2.font.size=Pt(10); r2.italic=True; r2.font.color.rgb=GREY

    # ── Idea box ──────────────────────────────────────────────
    doc.add_paragraph()
    tb = doc.add_table(rows=1,cols=2); tb.style="Table Grid"
    c1=tb.cell(0,0); c2=tb.cell(0,1)
    set_bg(c1,"1F3864"); set_bg(c2,"EAF1FB")
    c1.width=Inches(0.12)
    c1.paragraphs[0].add_run("")
    rp=c2.paragraphs[0]
    rp.paragraph_format.space_before=Pt(8); rp.paragraph_format.space_after=Pt(2)
    rp.add_run("Innovation Idea").bold=True
    rb=rp.runs[0]; rb.font.size=Pt(9); rb.font.color.rgb=NAVY
    rp2=c2.add_paragraph()
    rp2.paragraph_format.space_before=Pt(0); rp2.paragraph_format.space_after=Pt(8)
    ri=rp2.add_run(idea); ri.font.size=Pt(10); ri.italic=True
    doc.add_paragraph()

    # ── Score summary ─────────────────────────────────────────
    h1(doc,"Score Summary")
    st2=doc.add_table(rows=1,cols=4); st2.style="Table Grid"
    hdr_row(st2,["Dimension","Score","Weight","Weighted"])
    rows_data=[
        ("Market Attractiveness",f"{weights['Market Attractiveness'][0]:.1f}/10","40%",f"{weights['Market Attractiveness'][0]*0.4:.1f}"),
        ("Sector Fit",f"{weights['Sector Fit'][0]:.1f}/10","35%",f"{weights['Sector Fit'][0]*0.35:.1f}"),
        ("Competition Opportunity",f"{weights['Competition Opportunity'][0]:.1f}/10","25%",f"{weights['Competition Opportunity'][0]*0.25:.1f}"),
    ]
    for i,(dim,score,weight,weighted) in enumerate(rows_data):
        row=st2.add_row(); fill="EAF1FB" if i%2==0 else "FFFFFF"
        for j,val in enumerate([dim,score,weight,weighted]):
            c=row.cells[j]; set_bg(c,fill)
            r=c.paragraphs[0].add_run(val); r.font.size=Pt(10); r.bold=(j==0)
    fr=st2.add_row()
    for c in fr.cells: set_bg(c,"1F3864")
    r=fr.cells[0].paragraphs[0].add_run("FINAL SCORE")
    r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
    r2=fr.cells[3].paragraphs[0].add_run(f"{final_score}/10")
    r2.bold=True; r2.font.size=Pt(11); r2.font.color.rgb=LBLUE
    doc.add_paragraph()

    # ── Executive summary ─────────────────────────────────────
    h1(doc,"Executive Summary")
    body(doc, ext.get("executive_summary",""))

    # ── Market size ───────────────────────────────────────────
    h1(doc,"Market Size & Growth")
    kv(doc,"Market size (2024)",_mval(market.get("market_size_current") or market.get("market_size_2024","N/A")))
    kv(doc,"Projected (2030)",_mval(market.get("market_size_forecast") or market.get("market_size_2030","N/A")))
    kv(doc,"CAGR",_mval(market.get("cagr","N/A")))
    kv(doc,"Maturity",market.get("market_maturity","N/A"))
    kv(doc,"Geography",market.get("geographic_focus","N/A"))
    doc.add_paragraph()
    body(doc, ext.get("market_deep_dive",""))
    if market.get("growth_drivers"):
        h2(doc,"Growth Drivers")
        for d in market["growth_drivers"]: bul(doc,d)

    # ── Sector clusters ───────────────────────────────────────
    h1(doc,"Schaeffler Sector Cluster Fit")
    primary=sectors.get("primary_sectors",[])
    body(doc,f"Primary sectors: {', '.join(primary)}")
    body(doc,sectors.get("sector_fit_rationale",""))
    doc.add_paragraph()
    st3=doc.add_table(rows=1,cols=3); st3.style="Table Grid"
    hdr_row(st3,["Sector","Score","Rationale"])
    for idx,(sector,data) in enumerate(sectors.get("sector_scores",{}).items()):
        row=st3.add_row()
        fill="EAF5EA" if sector in primary else ("EAF1FB" if idx%2==0 else "FFFFFF")
        for c in row.cells: set_bg(c,fill)
        r0=row.cells[0].paragraphs[0].add_run(sector)
        r0.font.size=Pt(10); r0.bold=(sector in primary)
        r1=row.cells[1].paragraphs[0].add_run(f"{data.get('score',0)}/10")
        r1.font.size=Pt(10); r1.bold=True
        r1.font.color.rgb=BLUE if sector in primary else BLACK
        r2=row.cells[2].paragraphs[0].add_run(data.get("rationale",""))
        r2.font.size=Pt(9.5)

    # ── Competitive landscape ─────────────────────────────────
    h1(doc,"Competitive Landscape")
    kv(doc,"Competitive intensity",comp.get("competitive_intensity",""))
    kv(doc,"White space",comp.get("white_space",""))
    kv(doc,"Schaeffler advantage",comp.get("schaeffler_advantage",""))
    doc.add_paragraph()
    body(doc, ext.get("competitive_analysis",""))
    if comp.get("competitors"):
        h2(doc,"Key Players")
        ct=doc.add_table(rows=1,cols=3); ct.style="Table Grid"
        hdr_row(ct,["Company","Type","Relevance"])
        for idx,ci in enumerate(comp["competitors"]):
            row=ct.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            for c in row.cells: set_bg(c,fill)
            for j,val in enumerate([ci.get("name",""),ci.get("type",""),ci.get("relevance","")+" "+ci.get("source","")]):
                r=row.cells[j].paragraphs[0].add_run(val)
                r.font.size=Pt(9.5); r.bold=(j==0)

    # ── Strategic fit ─────────────────────────────────────────
    h1(doc,"Schaeffler Strategic Fit")
    body(doc, ext.get("schaeffler_fit",""))

    # ── Risks & recommendations ───────────────────────────────
    h1(doc,"Risks & Mitigations")
    risks_list = [
        f"Competitive intensity is {comp.get('competitive_intensity','unknown')} — monitor key players and conduct freedom-to-operate analysis before committing R&D budget",
        "Market timing risk — validate demand with target customers before scaling investment",
        "Sector fit assumptions should be validated with Schaeffler divisional teams before resource allocation"
    ]
    for risk in risks_list: bul(doc, risk)
    h1(doc,"Recommendations")
    recs_list = [
        "Commission internal engineering feasibility review within 30 days",
        "Conduct freedom-to-operate IP analysis with Schaeffler patent team",
        f"Identify pilot customer in {', '.join(sectors.get('primary_sectors', ['target sector'])[:1])} for co-development",
        "Present to Innovation steering committee with this report as supporting material"
    ]
    for rec in recs_list: bul(doc, rec)

    # ── Footer ────────────────────────────────────────────────
    doc.add_paragraph()
    ft=doc.add_table(rows=1,cols=1); ft.style="Table Grid"
    fc=ft.cell(0,0); set_bg(fc,"1F3864")
    fp=fc.paragraphs[0]
    fp.paragraph_format.space_before=Pt(6); fp.paragraph_format.space_after=Pt(6)
    fr=fp.add_run(f"Schaeffler AI Innovation Research Assistant  ·  Stage 02: Market Intelligence  ·  {datetime.now().strftime('%d %B %Y')}  ·  Capstone Project — Arpan Chowdhury, EBS Universität")
    fr.font.size=Pt(8); fr.font.color.rgb=RGBColor(0x93,0xC5,0xFD)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf



# ════════════════════════════════════════════════════════════
# RUN FULL ANALYSIS — sequential pipeline helper
# ════════════════════════════════════════════════════════════
def run_stage2(idea, quadrant, s1c):
    """Run Stage 02 Market Intelligence and store results in session state."""
    web_ctx = ""
    system_market = """You are a senior market analyst. Analyse the market for this innovation idea.
RULES:
- Use only the most credible sources: McKinsey Global Institute, Gartner, Frost & Sullivan, BloombergNEF, IEA, Roland Berger, Statista, MarketsandMarkets, Grand View Research, Allied Market Research, Mordor Intelligence, Fortune Business Insights, IDC, Wood Mackenzie, S&P Global.
- For each market figure, provide structured source objects. Each source MUST have its own entry in the sources array.
- For URLs: provide the EXACT deep-link URL. If you do not know it, omit the url field entirely.
- market_score: integer 1-10. Guide: 9-10=large fast-growing; 7-8=strong; 5-6=moderate; 3-4=niche; 1-2=tiny/declining.
Return ONLY valid JSON with NO extra text, NO inline comments:
{"market_name":"string",
"market_size_current":{"value":"$X.XB","year":"2024","sources":[{"org":"OrgName","title":"Report Title Year","year":"2024","url":"https://url-or-omit-field"}]},
"market_size_forecast":{"value":"$X.XB","year":"2030","sources":[{"org":"OrgName","title":"Report Title Year","year":"2024","url":"https://url-or-omit-field"}]},
"cagr":{"value":"X%","period":"2024-2030","sources":[{"org":"OrgName","title":"Report Title Year","year":"2024","url":"https://url-or-omit-field"}]},
"growth_drivers":["driver 1","driver 2","driver 3"],"market_maturity":"Emerging/Growing/Mature/Declining",
"geographic_focus":"string","market_score":7,"market_score_rationale":"2 sentences"}"""
    raw = call_claude(system_market, f"Idea: {idea}\nQuadrant: {quadrant}\nTech: {s1c.get('technology_novelty','')}", max_tokens=1400)
    try:
        market = _parse_json(raw)
    except Exception:
        market = {"market_name":"N/A","market_size_current":{"value":"N/A","year":"2024","sources":[]},"market_size_forecast":{"value":"N/A","year":"2030","sources":[]},"cagr":{"value":"N/A","period":"","sources":[]},"growth_drivers":[],"market_maturity":"N/A","geographic_focus":"N/A","market_score":5,"market_score_rationale":""}

    system_comp = """You are a competitive intelligence analyst. Identify key competitors for this idea.
Every company must have a source. Return ONLY valid JSON with NO inline comments:
{"competitors":[{"name":"string","type":"Incumbent/Startup/Research","relevance":"one sentence","source":"Source: X, Y"}],
"competitive_intensity":"Low/Medium/High/Very High","white_space":"one sentence","schaeffler_advantage":"one sentence",
"competition_score":7,"competition_score_rationale":"2 sentences"}"""
    raw = call_claude(system_comp, f"Idea: {idea}\nMarket: {market.get('market_name','')}\nQuadrant: {quadrant}", max_tokens=1000)
    try:
        comp = _parse_json(raw)
    except Exception:
        comp = {"competitors":[],"competitive_intensity":"N/A","white_space":"N/A","schaeffler_advantage":"N/A","competition_score":5,"competition_score_rationale":""}

    system_sectors = """You are a Schaeffler strategist. Score fit against Schaeffler's 10 sector clusters (0-10 each).
Clusters: Passenger Cars, Commercial Vehicles, Industrial Machinery, Rail, Aerospace, Two-Wheelers, Construction & Agriculture, Medical Equipment, Conventional Energy, Renewable Energy.
Return ONLY valid JSON:
{"sector_scores":{"Passenger Cars":{"score":0-10,"rationale":"one sentence"},"Commercial Vehicles":{"score":0-10,"rationale":"one sentence"},"Industrial Machinery":{"score":0-10,"rationale":"one sentence"},"Rail":{"score":0-10,"rationale":"one sentence"},"Aerospace":{"score":0-10,"rationale":"one sentence"},"Two-Wheelers":{"score":0-10,"rationale":"one sentence"},"Construction & Agriculture":{"score":0-10,"rationale":"one sentence"},"Medical Equipment":{"score":0-10,"rationale":"one sentence"},"Conventional Energy":{"score":0-10,"rationale":"one sentence"},"Renewable Energy":{"score":0-10,"rationale":"one sentence"}},
"primary_sectors":["top 2-3 sector names"],"sector_fit_score":0-10,"sector_fit_rationale":"2 sentences"}"""
    raw = call_claude(system_sectors, f"Idea: {idea}\nQuadrant: {quadrant}", max_tokens=1000)
    try:
        rc = raw.strip().replace("```json","").replace("```","").strip()
        fb = rc.find("{"); lb = rc.rfind("}") + 1
        if fb >= 0: rc = rc[fb:lb]
        sectors = json.loads(rc)
    except:
        sectors = {"sector_scores":{},"primary_sectors":[],"sector_fit_score":5,"sector_fit_rationale":""}

    ms = float(market.get("market_score",5))
    ss = float(sectors.get("sector_fit_score",5))
    cs = float(comp.get("competition_score",5))
    final = round(ms*0.40 + ss*0.35 + cs*0.25, 1)

    st.session_state.s2_data = {
        "market": market, "comp": comp, "sectors": sectors,
        "weights": {"Market Attractiveness":(ms,0.40),"Sector Fit":(ss,0.35),"Competition Opportunity":(cs,0.25)},
        "final_score": final, "web_results": []
    }
    st.session_state.s2_step = "done"


def run_stage3(idea, quadrant, s1c):
    """Run Stage 03 Patent Intelligence and store results in session state."""
    system_landscape = """You are a patent intelligence analyst. Analyse the external patent landscape.
Return ONLY valid JSON:
{"technology_keywords":["3-5 terms"],"landscape_summary":"2-3 sentences","activity_level":"Low/Moderate/High/Very High","filing_trend":"Increasing/Stable/Decreasing","filing_trend_rationale":"one sentence","patent_landscape_score":1-10,
"key_filers":[{"company":"name","type":"Competitor/Customer/Research Institution/Adjacent Player","focus":"one sentence","threat_level":"Low/Medium/High","schaeffler_relationship":"string","source":"Source: X, Year"}],
"white_spaces":["white space 1","white space 2","white space 3"]}"""
    raw = call_claude(system_landscape, f"Idea: {idea}\nQuadrant: {quadrant}\nTech: {s1c.get('technology_novelty','')}", max_tokens=1500)
    try:
        raw_c = raw.strip().replace("```json","").replace("```","").strip()
        fb = raw_c.find("{"); lb = raw_c.rfind("}") + 1
        if fb >= 0: raw_c = raw_c[fb:lb]
        landscape = json.loads(raw_c)
    except:
        landscape = {"technology_keywords":[],"landscape_summary":"N/A","activity_level":"N/A","filing_trend":"N/A","filing_trend_rationale":"","patent_landscape_score":5,"key_filers":[],"white_spaces":[]}

    key_filers_run3 = landscape.get("key_filers", [])
    filers_full_run3 = json.dumps([
        {"company": f.get("company",""), "type": f.get("type",""), "focus": f.get("focus",""), "threat_level": f.get("threat_level","")}
        for f in key_filers_run3
    ])
    system_ansoff = """You are a Schaeffler patent strategist. Map ALL listed filers onto Schaeffler's Ansoff matrix.
IMPORTANT: Every filer in the input list MUST appear in filer_positions — do not skip any.

Matrix axes (X=Technology Dimension, Y=Market Dimension — same as Schaeffler's Stage 1):
- x_score: 0-10 (0=existing technology, 10=new to the world technology)
- y_score: 0-10 (0=existing market, 10=new to the world market)
Quadrants: EXPLOIT(x 0-5,y 0-5)=bottom-left, EXTEND(x 0-5,y 5-10)=top-left,
           RADICAL(x 5-10,y 5-10)=top-right, DISRUPT(x 5-10,y 0-5)=bottom-right

Return ONLY valid JSON:
{"filer_positions":[{"company":"name","matrix_position":"EXPLOIT/EXTEND/RADICAL/DISRUPT","x_score":0-10,"y_score":0-10,"rationale":"one sentence"}],
"schaeffler_position":{"matrix_position":"EXPLOIT/EXTEND/RADICAL/DISRUPT","x_score":0-10,"y_score":0-10,"existing_ip":"one sentence","gap":"one sentence"},
"idea_position":{"x_score":0-10,"y_score":0-10},"novelty_signal":"Strong/Moderate/Weak","novelty_rationale":"one sentence","ip_risk":"Low/Medium/High","ip_risk_rationale":"one sentence"}"""
    raw2 = call_claude(system_ansoff,
        f"Idea: {idea}\nQuadrant: {quadrant}\nMap ALL {len(key_filers_run3)} filers: {filers_full_run3}",
        max_tokens=max(1800, len(key_filers_run3) * 200 + 800))
    try:
        raw2_c = raw2.strip().replace("```json","").replace("```","").strip()
        fb2 = raw2_c.find("{"); lb2 = raw2_c.rfind("}") + 1
        if fb2 >= 0: raw2_c = raw2_c[fb2:lb2]
        ansoff_data = json.loads(raw2_c)
    except:
        ansoff_data = {"filer_positions":[],"schaeffler_position":{"matrix_position":"EXPLOIT","x_score":2,"y_score":2,"existing_ip":"N/A","gap":"N/A"},"idea_position":{"x_score":7,"y_score":7},"novelty_signal":"Moderate","novelty_rationale":"","ip_risk":"Medium","ip_risk_rationale":""}

    # Guarantee every key_filer has a position
    # X=Technology, Y=Market — EXPLOIT(low x,low y), EXTEND(low x,high y),
    # RADICAL(high x,high y), DISRUPT(high x,low y)
    positioned_run3 = {fp.get("company","").lower() for fp in ansoff_data.get("filer_positions", [])}
    type_defaults_run3 = {
        "Competitor":          ("EXPLOIT", 3.0, 3.0),   # established tech + established market
        "Customer":            ("EXTEND",  2.5, 6.5),   # established tech + new market
        "Research Institution":("RADICAL", 7.0, 7.5),  # new tech + new market
        "Adjacent Player":     ("DISRUPT", 6.5, 3.5),  # new tech + established market
        "Patent Troll":        ("EXPLOIT", 2.0, 2.0),  # established tech + established market
    }
    for i, f in enumerate(key_filers_run3):
        name = f.get("company","")
        if name.lower() not in positioned_run3 and name:
            quad, bx, by = type_defaults_run3.get(f.get("type","Adjacent Player"), ("EXPLOIT", 4.0, 4.0))
            nudge_x = ((i * 0.7) % 2.0) - 1.0
            nudge_y = ((i * 1.1) % 2.0) - 1.0
            ansoff_data.setdefault("filer_positions", []).append({
                "company": name, "type": f.get("type","Adjacent Player"),
                "matrix_position": quad,
                "x_score": round(min(9.5, max(0.5, bx + nudge_x)), 1),
                "y_score": round(min(9.5, max(0.5, by + nudge_y)), 1),
                "rationale": f.get("focus","Auto-placed based on filer type")
            })

    landscape_score = float(landscape.get("patent_landscape_score",5))
    novelty_score = {"Strong":9,"Moderate":6,"Weak":3}.get(ansoff_data.get("novelty_signal","Moderate"),6)
    ip_score      = {"Low":8,"Medium":5,"High":2}.get(ansoff_data.get("ip_risk","Medium"),5)
    final_patent  = round(landscape_score*0.40 + novelty_score*0.35 + ip_score*0.25, 1)

    st.session_state.s3_data = {
        "landscape": landscape, "ansoff_data": ansoff_data,
        "novelty_score": novelty_score, "ip_score": ip_score,
        "landscape_score": landscape_score, "final_score": final_patent
    }
    st.session_state.s3_step = "done"


def run_stage4(idea, quadrant, s1c):
    """Run Stage 04 Technical Feasibility and store results in session state."""
    system_existence = """You are a technology analyst. Assess whether this technology exists.
Return ONLY valid JSON:
{"technology_core":"one sentence","existence_verdict":"Demonstrated/Partially Demonstrated/Research Stage/Theoretical",
"existence_summary":"2-3 sentences","evidence":[{"type":"Academic Paper/Startup/Pilot/Industry Report/Patent","title":"string","description":"one sentence","relevance":"Direct/Adjacent/Analogous","confidence":"High/Medium/Low","source":"org or URL"}],
"technology_gaps":["gap 1","gap 2","gap 3"],"time_to_readiness":"e.g. 3-5 years","keywords":["6-10 key technical terms from this domain"]}"""
    raw = call_claude(system_existence, f"Idea: {idea}\nQuadrant: {quadrant}\nTech: {s1c.get('technology_novelty','')}", max_tokens=2000)
    try:
        raw_clean = raw.strip().replace("```json","").replace("```","").strip()
        fb = raw_clean.find("{"); lb = raw_clean.rfind("}") + 1
        if fb >= 0: raw_clean = raw_clean[fb:lb]
        existence = json.loads(raw_clean)
    except:
        existence = {"technology_core":"N/A","existence_verdict":"Research Stage","existence_summary":"N/A","evidence":[],"technology_gaps":[],"time_to_readiness":"Not yet estimated","keywords":[]}

    system_trl = """You are a Schaeffler TRL expert. Rate using Schaeffler-adapted TRL 1-9.
TRL 1-2=Theoretical, TRL 3-5=Innovation territory, TRL 6-7=Borderline, TRL 8-9=Product Development.
Return ONLY valid JSON:
{"trl_level":1-9,"trl_label":"TRL X — label","trl_rationale":"2-3 sentences","schaeffler_entry_readiness":"Too Early/Ready for Innovation/Ready for Product Development",
"key_technical_risks":[{"risk":"string","severity":"High/Medium/Low","mitigation":"one sentence"}],
"analogous_schaeffler_technologies":"one sentence on which Schaeffler Motion Product Family this is closest to",
"trl_score":1-10}"""
    raw2 = call_claude(system_trl, f"Idea: {idea}\nExistence: {existence.get('existence_verdict','')}\nEvidence count: {len(existence.get('evidence',[]))}\nGaps: {existence.get('technology_gaps','')}", max_tokens=1200)
    try:
        raw2_c = raw2.strip().replace("```json","").replace("```","").strip()
        fb2 = raw2_c.find("{"); lb2 = raw2_c.rfind("}") + 1
        if fb2 >= 0: raw2_c = raw2_c[fb2:lb2]
        trl = json.loads(raw2_c)
    except:
        trl = {"trl_level":3,"trl_label":"TRL 3 — Experimental proof of concept","trl_rationale":"","schaeffler_entry_readiness":"Too Early","key_technical_risks":[],"analogous_schaeffler_technologies":"","trl_score":3}

    trl_score  = float(trl.get("trl_score", round((trl.get("trl_level",3) / 9) * 10, 1)))
    ev_map = {"Demonstrated":9,"Partially Demonstrated":6,"Research Stage":3,"Theoretical":1}
    existence_score = ev_map.get(existence.get("existence_verdict","Research Stage"),3)
    risks = trl.get("key_technical_risks",[])
    sev_map = {"High":8,"Medium":5,"Low":2}
    risk_score = round(10 - (sum(sev_map.get(r.get("severity","Medium"),5) for r in risks[:3]) / max(len(risks[:3]),1)), 1) if risks else 7.0
    final_feasibility = round(trl_score*0.50 + existence_score*0.30 + risk_score*0.20, 1)

    st.session_state.s4_data = {
        "existence": existence, "trl": trl,
        "trl_score": trl_score, "existence_score": existence_score, "risk_score": risk_score,
        "final_score": final_feasibility
    }
    st.session_state.s4_step = "done"


def run_stage5(idea, quadrant, s1c):
    """Run Stage 05 Organisational Readiness and store results in session state."""
    s3_landscape = st.session_state.get("s3_data",{}).get("landscape",{})
    s4_existence = st.session_state.get("s4_data",{}).get("existence",{})
    s4_trl       = st.session_state.get("s4_data",{}).get("trl",{})
    prior_filers = [f.get("company","") for f in s3_landscape.get("key_filers",[])]
    prior_evidence_sources = [e.get("source","") for e in s4_existence.get("evidence",[])]
    trl_level = s4_trl.get("trl_level",3)
    innovation_cluster = s1c.get("innovation_cluster","")
    product_family     = s1c.get("product_family","")
    trend_alignment    = s1c.get("trend_alignment",[])

    system_readiness = f"""You are a senior Schaeffler innovation strategist assessing internal organisational readiness.
Schaeffler P³: Performance = Portfolio × People × Process.
Innovation cluster: {innovation_cluster}, Product family: {product_family}, Trends: {', '.join(trend_alignment)}, TRL: {trl_level}
Known filers: {', '.join(prior_filers[:6])}, Evidence sources: {', '.join(prior_evidence_sources[:5])}
Schaeffler competencies: Precision bearings, mechatronics, power electronics (Vitesco), tribology, EV drivetrains, embedded sensors, ASPICE/ISO 26262, OEM Tier 1.
Return ONLY valid JSON:
{{"p3_portfolio":{{"score":0-10,"rationale":"2 sentences","cluster_fit":"one sentence","strengths":["s1","s2"],"gaps":["g1"]}},
"p3_people":{{"score":0-10,"rationale":"2 sentences","matched_competencies":["c1","c2","c3"],"competency_gap":"string","sourcing_route":"string"}},
"p3_process":{{"score":0-10,"rationale":"2 sentences","applicable_assets":["a1","a2"],"investment_required":"string","time_to_close":"string"}},
"partnership_candidates":[{{"name":"string","type":"Startup/University/Customer/Supplier","rationale":"string","route":"Co-develop/Acquire/License/JDA"}}],
"org_gaps":[{{"gap":"string","severity":"High/Medium/Low","closure_route":"string","timeline":"string"}}],
"build_or_partner":{{"recommendation":"string","rationale":"2-3 sentences","time_to_trl6_internal":"string","time_to_trl6_partner":"string"}},
"org_readiness_score":0-10}}"""

    raw = call_claude(system_readiness, f"Innovation idea: {idea}\nQuadrant: {quadrant}\nTRL: {trl_level}", max_tokens=2500)
    try:
        raw_clean = raw.strip().replace("```json","").replace("```","").strip()
        fb = raw_clean.find("{"); lb = raw_clean.rfind("}") + 1
        if fb >= 0: raw_clean = raw_clean[fb:lb]
        org_data = json.loads(raw_clean)
    except:
        org_data = {"p3_portfolio":{"score":5,"rationale":"N/A","cluster_fit":"N/A","strengths":[],"gaps":[]},"p3_people":{"score":5,"rationale":"N/A","matched_competencies":[],"competency_gap":"N/A","sourcing_route":"N/A"},"p3_process":{"score":5,"rationale":"N/A","applicable_assets":[],"investment_required":"N/A","time_to_close":"N/A"},"partnership_candidates":[],"org_gaps":[],"build_or_partner":{"recommendation":"Co-develop","rationale":"N/A","time_to_trl6_internal":"N/A","time_to_trl6_partner":"N/A"},"org_readiness_score":5}

    p_portfolio = float(org_data.get("p3_portfolio",{}).get("score",5))
    p_people    = float(org_data.get("p3_people",{}).get("score",5))
    p_process   = float(org_data.get("p3_process",{}).get("score",5))
    final_org   = round(p_portfolio*0.35 + p_people*0.40 + p_process*0.25, 1)

    st.session_state.s5_data = {
        "org_data": org_data,
        "p_portfolio": p_portfolio,
        "p_people": p_people,
        "p_process": p_process,
        "final_score": final_org
    }
    st.session_state.s5_step = "done"


def run_stage6_synthesis(idea, quadrant, s1c):
    """Run Stage 06 synthesis with default weights and store results."""
    market_score      = st.session_state.s2_data.get("final_score", 5.0)
    patent_score      = st.session_state.s3_data.get("final_score", 5.0)
    feasibility_score = st.session_state.s4_data.get("final_score", 5.0)
    org_score         = st.session_state.s5_data.get("final_score", 5.0)

    weights = {"market":35,"patent":25,"feasibility":25,"org":15}
    wm = weights["market"]/100; wp = weights["patent"]/100
    wf = weights["feasibility"]/100; wo = weights["org"]/100
    ipi = round(market_score*wm + patent_score*wp + feasibility_score*wf + org_score*wo, 1)

    s2d = st.session_state.s2_data
    s3d = st.session_state.s3_data
    s4d = st.session_state.s4_data
    s5d = st.session_state.s5_data
    org_d = s5d.get("org_data",{})

    synthesis_context = f"""Idea: {idea}\nQuadrant: {quadrant}\nIPI: {ipi}/10
Market ({weights['market']}%): {market_score}/10 — {s2d.get('market',{}).get('market_name','')}
Patent ({weights['patent']}%): {patent_score}/10 — Novelty: {s3d.get('ansoff_data',{}).get('novelty_signal','')} IP risk: {s3d.get('ansoff_data',{}).get('ip_risk','')}
Feasibility ({weights['feasibility']}%): {feasibility_score}/10 — TRL {s4d.get('trl',{}).get('trl_level','')} {s4d.get('trl',{}).get('schaeffler_entry_readiness','')}
Org Readiness ({weights['org']}%): {org_score}/10 — {org_d.get('build_or_partner',{}).get('recommendation','')}"""

    system_structured = """You are a senior Schaeffler innovation strategist. Return ONLY valid JSON:
{"headline":"one direct sentence","recommendation":"PROCEED or PROCEED WITH CONDITIONS or DEFER or REJECT",
"recommendation_rationale":"2-3 sentences","strongest_signals":["signal 1","signal 2","signal 3"],
"key_concerns":["concern 1","concern 2","concern 3"],"conditions":["condition 1","condition 2"],
"strategic_fit":"2-3 sentences referencing Schaeffler P³, electrification, Vitesco merger",
"risks":["risk 1 with mitigation","risk 2","risk 3"],
"next_steps":["action 1","action 2","action 3","action 4"]}"""

    raw1 = call_claude(system_structured, synthesis_context, max_tokens=1000)
    raw1_clean = raw1.strip().replace("```json","").replace("```","").strip()
    fb = raw1_clean.find("{"); lb = raw1_clean.rfind("}") + 1
    if fb >= 0: raw1_clean = raw1_clean[fb:lb]
    try:
        synthesis_structured = json.loads(raw1_clean)
    except:
        synthesis_structured = {"headline":f"IPI {ipi}/10","recommendation":"PROCEED WITH CONDITIONS" if ipi>=5 else "DEFER","recommendation_rationale":"Based on pipeline analysis.","strongest_signals":[],"key_concerns":[],"conditions":[],"strategic_fit":"","risks":[],"next_steps":[]}

    system_narrative = "Write a 4-paragraph narrative synthesis for this Schaeffler innovation assessment. Flowing prose, no bullets. Cover: market opportunity, IP landscape, technical maturity, organisational readiness, and recommendation. Reference Schaeffler P³ formula and electrification context."
    raw2 = call_claude(system_narrative, synthesis_context + f"\nRecommendation: {synthesis_structured.get('recommendation','')}\nIPI: {ipi}/10", max_tokens=600)
    narrative_text = raw2.strip().replace("```","").strip()

    synthesis = {**synthesis_structured, "narrative": narrative_text}

    st.session_state.s6_data = {
        "ipi": ipi, "weights": weights, "synthesis": synthesis,
        "scores": {"market":market_score,"patent":patent_score,"feasibility":feasibility_score,"org":org_score}
    }
    st.session_state.s6_step = "done"


if st.session_state.active_stage == 1:
    st.markdown(f"## {T('s1_title')}")
    st.markdown(f"""<div style="background:#1a2d45;border-radius:6px;padding:14px 18px;margin-bottom:16px;border-left:4px solid #2E75B6;">
<div style="color:#60a5fa;font-size:11px;letter-spacing:1px;font-weight:600;">{T('s1_what_label')}</div>
<div style="color:#e2e8f0;font-size:13px;margin-top:6px;">{T('s1_what')}</div>
<div style="color:#94a3b8;font-size:12px;margin-top:8px;">{T('s1_you_get')}</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    # Step 1 — Input
    if st.session_state.s1_step == 1:
        st.subheader(T("s1_step1"))

        ucol1, ucol2, ucol3 = st.columns(3)
        with ucol1:
            user_name = st.text_input(T("s1_name"), value=st.session_state.user_name, placeholder=T("s1_name_ph"))
        with ucol2:
            user_position = st.text_input(T("s1_role"), value=st.session_state.user_position, placeholder=T("s1_role_ph"))
        with ucol3:
            user_dept = st.text_input(T("s1_dept"), value=st.session_state.user_dept, placeholder=T("s1_dept_ph"))

        st.markdown("---")
        idea = st.text_area(T("s1_idea_label"), height=150, placeholder=T("s1_idea_ph"))

        if st.button(T("s1_submit"), type="primary"):
            if not user_name.strip() or not user_position.strip() or not user_dept.strip():
                st.warning(T("s1_warn_identity"))
            elif not idea.strip():
                st.warning(T("s1_warn_empty"))
            elif len(idea.split()) < 15:
                st.warning(T("s1_warn_brief"))
            else:
                # Save user identity
                st.session_state.user_name     = user_name.strip()
                st.session_state.user_position = user_position.strip()
                st.session_state.user_dept     = user_dept.strip()
                with st.spinner(T("s1_spinner_check")):
                    check = call_claude(
                        'Check if this innovation idea has enough detail to classify. Reply ONLY with JSON: {"sufficient": true/false, "missing": "one short sentence or empty string"}',
                        idea
                    )
                    try:
                        result = json.loads(check.strip().replace("```json","").replace("```","").strip())
                    except:
                        result = {"sufficient": True, "missing": ""}
                if not result.get("sufficient", True):
                    st.warning(f"A bit more detail needed: {result.get('missing','')}")
                else:
                    st.session_state.s1_idea = idea
                    # Check for similar past ideas
                    with st.spinner(T("s1_spinner_similar")):
                        past = load_past_ideas()
                        similar = check_similar_ideas(idea, past)
                        st.session_state.s1_similar_ideas = similar
                    st.session_state.s1_step = 2
                    st.rerun()

    # Step 2 — Forced-choice enrichment
    if st.session_state.s1_step == 2:
        st.markdown("---")
        st.subheader(T("s1_step2"))
        st.info(f"**Your idea:** {st.session_state.s1_idea}")

        # ── Similar ideas alert ───────────────────────────────
        similar = st.session_state.get("s1_similar_ideas", [])
        if similar:
            st.warning(T("s1_similar_warn").format(n=len(similar)))
            for s in similar:
                sim_col = "#f59e0b" if s.get("similarity") == "High" else "#60a5fa"
                st.markdown(f"""
<div style="background:#1a2d45;border-left:4px solid {sim_col};border-radius:4px;padding:12px 16px;margin:6px 0;">
  <div style="color:{sim_col};font-size:11px;font-weight:600;letter-spacing:1px;">{s.get('similarity','').upper()} SIMILARITY — {s.get('date','')} · {s.get('submitter','')} ({s.get('department','')})</div>
  <div style="color:#e2e8f0;font-size:13px;margin-top:4px;">{s.get('idea_snippet','')}</div>
  <div style="color:#94a3b8;font-size:12px;margin-top:4px;">Quadrant: <b>{s.get('quadrant','')}</b> · IPI: <b>{s.get('ipi','')}</b> · Outcome: <b>{s.get('recommendation','')}</b></div>
  <div style="color:#94a3b8;font-size:12px;margin-top:2px;">Similarity reason: {s.get('reason','')}</div>
</div>
""", unsafe_allow_html=True)
            st.markdown("---")
        st.caption(T("s1_step2_caption"))

        # Q1 — Technology novelty
        st.markdown(T("s1_q1"))
        q1_choice = st.radio("", [T("s1_q1a"), T("s1_q1b")], key="q1_radio", label_visibility="collapsed")
        q1_detail = st.text_input(T("s1_q1_detail"), key="q1_detail", placeholder=T("s1_q1_ph"))

        st.markdown("---")
        # Q2 — Market familiarity
        st.markdown(T("s1_q2"))
        st.caption(T("s1_q2_caption"))
        q2_choice = st.radio("", [T("s1_q2a"), T("s1_q2b")], key="q2_radio", label_visibility="collapsed")
        q2_detail = st.text_input(T("s1_q2_detail"), key="q2_detail", placeholder=T("s1_q2_ph"))

        st.markdown("---")
        # Q3 — Problem clarity
        st.markdown(T("s1_q3"))
        q3_choice = st.radio("", [T("s1_q3a"), T("s1_q3b")], key="q3_radio", label_visibility="collapsed")
        q3_detail = st.text_input(T("s1_q3_detail"), key="q3_detail", placeholder=T("s1_q3_ph"))

        if st.button(T("s1_classify_btn"), type="primary"):
            a1 = q1_choice + (f" — {q1_detail}" if q1_detail.strip() else "")
            a2 = q2_choice + (f" — {q2_detail}" if q2_detail.strip() else "")
            a3 = q3_choice + (f" — {q3_detail}" if q3_detail.strip() else "")
            st.session_state.s1_questions = [
                "Has the core technology been demonstrated anywhere?",
                "Does this idea target markets Schaeffler currently operates in?",
                "Is the problem this idea solves already recognised by industry?"
            ]
            st.session_state.s1_answers = [a1, a2, a3]
            st.session_state.s1_step = 3
            st.rerun()

    # Step 3 — Classify
    if st.session_state.s1_step == 3 and not st.session_state.s1_classification:
        with st.spinner(T("s1_spinner_classify")):
            q = st.session_state.s1_questions
            a = st.session_state.s1_answers
            system = """You are a senior innovation strategist at Schaeffler Group.
Classify using Schaeffler's Modified Innovation Matrix (Lau et al., ISPIM 2023).

The matrix uses 4 levels on BOTH axes (NOT a simple 2x2):
- Technology axis: Established → Adjacent → New to Schaeffler → New to the World
- Market axis: Established → Adjacent → New to Schaeffler → New to the World

Four quadrants (each spanning 2 levels in each direction):
EXPLOIT    — Established/Adjacent tech + Established/Adjacent market → Product Development
EXTEND     — Established/Adjacent tech + New to Schaeffler/World market → Product Development
RADICAL    — New to Schaeffler/World tech + Established/Adjacent market → Innovation pipeline
DISRUPTIVE — New to Schaeffler/World tech + New to Schaeffler/World market → Innovation pipeline

CRITICAL CLASSIFICATION RULES:
1. The IDEA DESCRIPTION is your PRIMARY and most authoritative signal. Base your classification on what the idea actually IS and what market it serves.
2. The three Q&A answers are SUPPLEMENTARY CONTEXT only — they help resolve genuine ambiguity in the idea description, but they must NEVER override a clear signal in the idea itself.
3. Q2 ("Is the target customer new?") is particularly dangerous — Schaeffler serves automotive OEMs, industrial machinery, rail, aerospace, EV drivetrains, energy, etc. An idea for any of these is EXISTING MARKET even if the user says "new customer" due to unfamiliarity with Schaeffler's portfolio. Do not flip to DISRUPTIVE just because the user answered Q2 as "new customer".
4. If the idea technology is clearly novel/breakthrough → lean RADICAL or DISRUPTIVE regardless of Q&A.
5. If both the idea AND the answers strongly point to existing tech + existing market → EXPLOIT. If mixed → EXTEND.
6. RADICAL vs DISRUPTIVE distinction: RADICAL targets markets Schaeffler already serves (automotive, industrial, rail, energy, EV drivetrains). DISRUPTIVE targets markets entirely outside Schaeffler's current scope (e.g. consumer electronics, healthcare devices, retail).
7. The proceed field MUST follow this strict rule: EXPLOIT and EXTEND → proceed:false. RADICAL and DISRUPTIVE → proceed:true. This is mandatory.

For EXPLOIT/EXTEND: name the relevant Schaeffler product division:
E-Mobility / Powertrain & Chassis / Vehicle Lifetime Solutions / Bearings & Industrial Solutions

For RADICAL/DISRUPTIVE: also assign:
- innovation_cluster: most relevant of Schaeffler's 8 clusters from Lau et al. 2023:
  Energy Solutions, Material Solutions, Mobility Solutions, E-Drive Solutions,
  Robotics Solutions, Digital Solutions, Advanced Manufacturing, New Production Concepts
- trend_alignment: 1-2 of Schaeffler's 5 strategic trends (Lau 2023):
  Sustainability & Climate Change, New Mobility & Electrification,
  Autonomous Production, Data Economy & Digitalization, Demographic Change
- product_family: most relevant of Schaeffler's 8 Motion Product Families (Enders 2026):
  Guide Motion, Transmit Motion, Control Motion, Generate Motion,
  Power Motion, Drive Motion, Energize Motion, Sustain Motion
- project_type: FIP (Research & Innovation Project — high uncertainty, long horizon)
  or VEP (Advanced Development Project — technology partially validated)
- innovation_model: Integrated (SHARE network — internal) or Accelerator (open innovation — external partners)
- technology_level: exact level on the axis (Established/Adjacent/New to Schaeffler/New to the World)
- market_level: exact level on the axis (Established/Adjacent/New to Schaeffler/New to the World)

Return ONLY valid JSON:
{
  "quadrant":"RADICAL/DISRUPTIVE/EXPLOIT/EXTEND",
  "confidence":"High/Medium/Low",
  "technology_level":"one of the 4 axis levels",
  "market_level":"one of the 4 axis levels",
  "technology_novelty":"one sentence",
  "market_position":"one sentence",
  "reasoning":"2-3 sentences explaining the classification primarily from the idea description. If Q&A answers conflicted with the idea, explain why you prioritised the idea.",
  "proceed":true/false,
  "schaeffler_division":"division name or empty string",
  "redirect_message":"one sentence if EXPLOIT/EXTEND else empty string",
  "innovation_cluster":"cluster name or empty string",
  "trend_alignment":["trend 1","trend 2"],
  "product_family":"Motion family name or empty string",
  "project_type":"FIP or VEP or empty string",
  "innovation_model":"Integrated or Accelerator or empty string"
}"""
            try:
                raw = call_claude(system,
                    f"IDEA DESCRIPTION (primary classification signal):\n{st.session_state.s1_idea}\n\n"
                    f"SUPPLEMENTARY Q&A (use only to resolve genuine ambiguity — do not override the idea):\n"
                    f"Q: {q[0]} A: {a[0]}\n"
                    f"Q: {q[1]} A: {a[1]}\n"
                    f"Q: {q[2]} A: {a[2]}")
                raw_clean = raw.strip().replace("```json","").replace("```","").strip()
                fb = raw_clean.find("{"); lb = raw_clean.rfind("}") + 1
                if fb >= 0: raw_clean = raw_clean[fb:lb]
                classification = json.loads(raw_clean)
                # Hard-code proceed based on quadrant — never trust Claude's value here
                # EXPLOIT and EXTEND must always redirect; RADICAL and DISRUPTIVE always proceed
                q_result = classification.get("quadrant","").upper()
                classification["proceed"] = q_result in ("RADICAL", "DISRUPTIVE")
                st.session_state.s1_classification = classification
            except Exception as e:
                st.error(f"Classification error: {e}")
                st.stop()
        st.rerun()

    # Step 3 — Show result
    if st.session_state.s1_step == 3 and st.session_state.s1_classification:
        c = st.session_state.s1_classification
        quadrant = c.get("quadrant","")
        proceed  = c.get("proceed", False)

        st.markdown("---")
        st.subheader(T("s1_result"))
        st.info(f"**Your idea:** {st.session_state.s1_idea}")

        if not proceed:
            division = c.get("schaeffler_division","Product Development")
            st.warning(f"**{quadrant}** — {c.get('redirect_message','')}")
            st.markdown(f"→ Suggested home: **{division}**")
            st.markdown(f"""
<div style="background:#1a2d45;border-radius:8px;padding:14px 18px;margin-top:12px;border-left:3px solid #f59e0b;">
<div style="color:#f59e0b;font-size:11px;font-weight:600;letter-spacing:1px;margin-bottom:6px;">WHY THIS IDEA DOESN'T ENTER THE INNOVATION PIPELINE</div>
<div style="color:#e2e8f0;font-size:13px;">
<b>EXPLOIT</b> and <b>EXTEND</b> ideas use established or adjacent technology — they belong in Schaeffler's Product Development divisions, not the Innovation Pipeline, because the core technology risk has already been resolved.<br><br>
The Innovation Pipeline (Stages 02–06) is reserved for <b>RADICAL</b> (breakthrough tech, existing market) and <b>DISRUPTIVE</b> (breakthrough tech, new market) ideas where the technology itself is genuinely novel and unproven.
</div>
<div style="color:#94a3b8;font-size:12px;margin-top:8px;">Technology level: <b>{c.get('technology_level','')}</b> · Market level: <b>{c.get('market_level','')}</b></div>
</div>
""", unsafe_allow_html=True)
        else:
            emoji = "🔬" if quadrant == "RADICAL" else "🚀"
            st.success(f"{emoji} **{quadrant}** — {c.get('reasoning','')}")
            st.caption(f"Confidence: {c.get('confidence','')}")

        tech_score, market_score = get_dot_position(quadrant, c.get("confidence","Medium"))
        st.plotly_chart(ansoff_chart(quadrant, tech_score, market_score), use_container_width=True)

        # ── 4-level axis labels ───────────────────────────────
        if proceed:
            col_t, col_m = st.columns(2)
            col_t.markdown(f'<div style="background:#1a2d45;border-radius:6px;padding:10px 14px;margin:4px 0;"><div style="color:#94a3b8;font-size:11px;letter-spacing:1px;">TECHNOLOGY LEVEL</div><div style="color:#60a5fa;font-size:14px;font-weight:600;">{c.get("technology_level","")}</div></div>', unsafe_allow_html=True)
            col_m.markdown(f'<div style="background:#1a2d45;border-radius:6px;padding:10px 14px;margin:4px 0;"><div style="color:#94a3b8;font-size:11px;letter-spacing:1px;">MARKET LEVEL</div><div style="color:#60a5fa;font-size:14px;font-weight:600;">{c.get("market_level","")}</div></div>', unsafe_allow_html=True)
            col_a, col_b, col_c = st.columns(3)
            col_a.markdown(f'<div style="background:#1a2d45;border-radius:6px;padding:8px 12px;margin:4px 0;"><div style="color:#94a3b8;font-size:10px;letter-spacing:1px;">INNOVATION CLUSTER</div><div style="color:#e2e8f0;font-size:12px;font-weight:600;margin-top:2px;">{c.get("innovation_cluster","")}</div></div>', unsafe_allow_html=True)
            col_b.markdown(f'<div style="background:#1a2d45;border-radius:6px;padding:8px 12px;margin:4px 0;"><div style="color:#94a3b8;font-size:10px;letter-spacing:1px;">PRODUCT FAMILY</div><div style="color:#e2e8f0;font-size:12px;font-weight:600;margin-top:2px;">{c.get("product_family","")}</div></div>', unsafe_allow_html=True)
            col_c.markdown(f'<div style="background:#1a2d45;border-radius:6px;padding:8px 12px;margin:4px 0;"><div style="color:#94a3b8;font-size:10px;letter-spacing:1px;">TREND ALIGNMENT</div><div style="color:#e2e8f0;font-size:12px;margin-top:2px;">{", ".join(c.get("trend_alignment",[]))}</div></div>', unsafe_allow_html=True)
            pipeline = c.get("pipeline_route","")
            if pipeline:
                route_col = "#22c55e" if "Integrated" in pipeline else "#60a5fa"
                st.markdown(f'<div style="background:#0f1e35;border:1px solid {route_col}44;border-radius:6px;padding:8px 14px;margin:8px 0;"><span style="color:{route_col};font-size:12px;font-weight:600;">🔀 Pipeline route: {pipeline}</span></div>', unsafe_allow_html=True)

        col1, col2, col3 = st.columns(3)
        col1.metric("Technology novelty", f"{tech_score} / 10")
        col2.metric("Market novelty",     f"{market_score} / 10")
        col3.metric("Confidence",          c.get("confidence",""))

        # Continue button
        st.markdown("---")
        if proceed:
            st.success(T("s1_qualifies"))
            btn_col1, btn_col2 = st.columns(2)
            with btn_col1:
                if st.button(T("s1_continue"), type="primary", key="s1_continue"):
                    st.session_state.active_stage = 2
                    st.rerun()
            with btn_col2:
                if st.button(T("s1_full_run"), type="secondary", key="s1_full_run"):
                    idea_fa  = st.session_state.s1_idea
                    s1c_fa   = st.session_state.s1_classification
                    quad_fa  = s1c_fa.get("quadrant","RADICAL")
                    with st.spinner("Running full pipeline — Stage 02: Market Intelligence..."):
                        run_stage2(idea_fa, quad_fa, s1c_fa)
                    with st.spinner("Stage 03: Patent Intelligence..."):
                        run_stage3(idea_fa, quad_fa, s1c_fa)
                    with st.spinner("Stage 04: Technical Feasibility..."):
                        run_stage4(idea_fa, quad_fa, s1c_fa)
                    with st.spinner("Stage 05: Organisational Readiness..."):
                        run_stage5(idea_fa, quad_fa, s1c_fa)
                    with st.spinner("Stage 06: Scoring & Synthesis..."):
                        run_stage6_synthesis(idea_fa, quad_fa, s1c_fa)
                    st.session_state.active_stage = 6
                    st.rerun()

        # Post-result chat
        st.markdown("---")
        st.subheader(T("s1_chat_header"))
        for msg in st.session_state.s1_chat:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        user_q = st.chat_input(T("s1_chat_ph"))
        if user_q:
            st.session_state.s1_chat.append({"role":"user","content":user_q})
            with st.chat_message("user"):
                st.markdown(user_q)
            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    context = f"""You are a senior Schaeffler innovation expert discussing a classification result.
Idea: {st.session_state.s1_idea}
Classification: {quadrant} (confidence: {c.get('confidence','')})
Reasoning: {c.get('reasoning','')}
Tech novelty: {c.get('technology_novelty','')}
Market position: {c.get('market_position','')}
{"Division: " + c.get('schaeffler_division','') if not proceed else "Qualifies for Innovation pipeline."}
Be specific and concise — 2-4 sentences. Reference Schaeffler's context (electrification, Vitesco, E-Mobility) where relevant."""
                    history = [{"role":m["role"],"content":m["content"]} for m in st.session_state.s1_chat]
                    reply = call_claude_chat(context, history)
                    st.markdown(reply)
                    st.session_state.s1_chat.append({"role":"assistant","content":reply})

        if st.button(T("s1_startover"), key="s1_startover"):
            for k in ["s1_step","s1_idea","s1_questions","s1_answers","s1_classification","s1_chat","s1_similar_ideas","user_name","user_position","user_dept"]:
                st.session_state[k] = defaults.get(k, "" if k in ("user_name","user_position","user_dept") else defaults.get(k))
            st.rerun()

# ════════════════════════════════════════════════════════════
# STAGE 02 — MARKET INTELLIGENCE
# ════════════════════════════════════════════════════════════
elif st.session_state.active_stage == 2:
    st.markdown(f"## {T('s2_title')}")
    st.markdown(f"""<div style="background:#1a2d45;border-radius:6px;padding:14px 18px;margin-bottom:16px;border-left:4px solid #2E75B6;">
<div style="color:#60a5fa;font-size:11px;letter-spacing:1px;font-weight:600;">{T('s1_what_label')}</div>
<div style="color:#e2e8f0;font-size:13px;margin-top:6px;">{T('s2_what')}</div>
<div style="color:#94a3b8;font-size:12px;margin-top:8px;">{T('s2_you_get')}</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    idea     = st.session_state.s1_idea
    s1c      = st.session_state.s1_classification
    quadrant = s1c.get("quadrant","RADICAL")

    # Intro
    if st.session_state.s2_step == "intro":
        st.info(f"**Idea:** {idea}")
        st.markdown(f"**Quadrant:** {quadrant} · {s1c.get('technology_novelty','')}")
        if st.button(T("s2_run_btn"), type="primary"):
            st.session_state.s2_step = "running"
            st.rerun()

    # Running
    elif st.session_state.s2_step == "running":
        progress = st.progress(0)
        status   = st.empty()

        status.markdown("🔍 Gathering market data...")
        progress.progress(15)
        search_results = []
        if TAVILY_KEY:
            for q in [f"{idea} market size 2024", f"{idea} competitors", f"{idea} market growth"]:
                search_results.extend(tavily_search(q))
        web_ctx = ""
        if search_results:
            web_ctx = "\n\nWeb results:\n" + "\n".join(
                f"- [{r['title']}]({r['url']}): {r['content'][:300]}" for r in search_results[:6])

        status.markdown("📊 Analysing market size and growth...")
        progress.progress(35)
        system_market = """You are a senior market analyst. Analyse the market for this innovation idea.
RULES:
- Use only the most credible sources: McKinsey Global Institute, Gartner, Frost & Sullivan, BloombergNEF, IEA, Roland Berger, Statista, MarketsandMarkets, Grand View Research, Allied Market Research, Mordor Intelligence, Fortune Business Insights, IDC, Wood Mackenzie, S&P Global.
- For each market figure, provide structured source objects. Each source MUST have its own entry in the sources array — never combine two sources into one object.
- For URLs: provide the EXACT deep-link URL to the specific report page. If you do not know the exact URL, omit the url field entirely.
- market_size_current = most recent available year (2024 or 2025). market_size_forecast = 5-7 year projection. cagr = compound annual growth rate for that period.
- market_score: integer 1-10. Rubric: 9-10=large fast-growing (>$10bn, >15% CAGR); 7-8=strong ($2-10bn, 8-15%); 5-6=moderate; 3-4=niche; 1-2=tiny or declining.
Return ONLY valid JSON with NO extra text, NO comments inside the JSON:
{"market_name":"string",
"market_size_current":{"value":"$X.XB","year":"2024","sources":[{"org":"OrgName","title":"Report Title Year","year":"2024","url":"https://url-or-omit-field"}]},
"market_size_forecast":{"value":"$X.XB","year":"2030","sources":[{"org":"OrgName","title":"Report Title Year","year":"2024","url":"https://url-or-omit-field"}]},
"cagr":{"value":"X%","period":"2024-2030","sources":[{"org":"OrgName","title":"Report Title Year","year":"2024","url":"https://url-or-omit-field"}]},
"growth_drivers":["driver 1","driver 2","driver 3"],"market_maturity":"Emerging/Growing/Mature/Declining",
"geographic_focus":"string","market_score":7,"market_score_rationale":"2 sentences"}"""
        try:
            raw = call_claude(system_market, f"Idea: {idea}\nQuadrant: {quadrant}\nTech: {s1c.get('technology_novelty','')}{web_ctx}", max_tokens=1400)
            market = _parse_json(raw)
        except Exception:
            market = {"market_name":"N/A","market_size_current":{"value":"N/A","year":"2024","sources":[]},"market_size_forecast":{"value":"N/A","year":"2030","sources":[]},"cagr":{"value":"N/A","period":"","sources":[]},"growth_drivers":[],"market_maturity":"N/A","geographic_focus":"N/A","market_score":5,"market_score_rationale":""}

        status.markdown("🏢 Mapping competitive landscape...")
        progress.progress(55)
        system_comp = """You are a competitive intelligence analyst. Identify key competitors for this idea.
Every company must have a source. Return ONLY valid JSON with NO inline comments or annotations:
{"competitors":[{"name":"string","type":"Incumbent/Startup/Research","relevance":"one sentence","source":"Source: X, Y"}],
"competitive_intensity":"Low/Medium/High/Very High","white_space":"one sentence","schaeffler_advantage":"one sentence",
"competition_score":7,"competition_score_rationale":"2 sentences"}
Scoring guide (do NOT include in the JSON): 9-10=very open/few players; 7-8=some room; 5-6=moderate; 3-4=crowded; 1-2=saturated."""
        try:
            raw = call_claude(system_comp, f"Idea: {idea}\nMarket: {market.get('market_name','')}\nQuadrant: {quadrant}{web_ctx}")
            comp = _parse_json(raw)
        except Exception:
            comp = {"competitors":[],"competitive_intensity":"N/A","white_space":"N/A",
                    "schaeffler_advantage":"N/A","competition_score":5,"competition_score_rationale":""}

        status.markdown("🎯 Scoring Schaeffler sector clusters...")
        progress.progress(75)
        system_sectors = """You are a Schaeffler strategist. Score fit against Schaeffler's 10 sector clusters (0-10 each).
Clusters: Passenger Cars, Commercial Vehicles, Industrial Machinery, Rail, Aerospace, Two-Wheelers, Construction & Agriculture, Medical Equipment, Conventional Energy, Renewable Energy.
0-2=No relevance, 3-4=Low, 5-6=Moderate, 7-8=High, 9-10=Primary target.
Return ONLY valid JSON with NO inline comments:
{"sector_scores":{"Passenger Cars":{"score":5,"rationale":"one sentence"},"Commercial Vehicles":{"score":5,"rationale":"one sentence"},"Industrial Machinery":{"score":5,"rationale":"one sentence"},"Rail":{"score":5,"rationale":"one sentence"},"Aerospace":{"score":5,"rationale":"one sentence"},"Two-Wheelers":{"score":5,"rationale":"one sentence"},"Construction & Agriculture":{"score":5,"rationale":"one sentence"},"Medical Equipment":{"score":5,"rationale":"one sentence"},"Conventional Energy":{"score":5,"rationale":"one sentence"},"Renewable Energy":{"score":5,"rationale":"one sentence"}},
"primary_sectors":["sector name 1","sector name 2"],"sector_fit_score":7,"sector_fit_rationale":"2 sentences"}"""
        try:
            raw = call_claude(system_sectors, f"Idea: {idea}\nQuadrant: {quadrant}")
            sectors = _parse_json(raw)
        except Exception:
            sectors = {"sector_scores":{},"primary_sectors":[],"sector_fit_score":5,"sector_fit_rationale":""}

        # Final weighted score
        ms  = float(market.get("market_score",5))
        ss  = float(sectors.get("sector_fit_score",5))
        cs  = float(comp.get("competition_score",5))
        final = round(ms*0.40 + ss*0.35 + cs*0.25, 1)

        st.session_state.s2_data = {
            "market": market, "comp": comp, "sectors": sectors,
            "weights": {"Market Attractiveness":(ms,0.40),"Sector Fit":(ss,0.35),"Competition Opportunity":(cs,0.25)},
            "final_score": final, "web_results": search_results
        }
        progress.progress(100)
        status.markdown("✓ Complete.")
        time.sleep(0.5)
        st.session_state.s2_step = "done"
        st.rerun()

    # Results
    elif st.session_state.s2_step == "done":
        d       = st.session_state.s2_data
        market  = d["market"]; comp = d["comp"]; sectors = d["sectors"]
        weights = d["weights"]; final = d["final_score"]

        # ── Score banner ──────────────────────────────────────
        score_col = "#22c55e" if final>=7 else "#f59e0b" if final>=4 else "#ef4444"
        banner = f"""<div style="background:#0f1e35;border-radius:8px;padding:20px 24px;margin-bottom:20px;border:1px solid #2a4a70;"><div style="color:{WHITE};font-size:11px;letter-spacing:1.5px;opacity:0.5;margin-bottom:4px;">MARKET INTELLIGENCE SCORE</div><div style="color:{score_col};font-size:44px;font-weight:700;line-height:1;">{final}<span style="font-size:18px;color:#94a3b8;"> / 10</span></div><div style="color:{WHITE};font-size:13px;margin-top:6px;opacity:0.8;">{market.get('market_name','')}</div></div>"""
        st.markdown(banner, unsafe_allow_html=True)

        # ── Score breakdown ───────────────────────────────────
        cols = st.columns(3)
        for i,(label,(score,weight)) in enumerate(weights.items()):
            cols[i].metric(label, f"{score:.1f}/10", f"{int(weight*100)}% weight")
        st.markdown("---")

        # ── Market size ───────────────────────────────────────
        st.markdown("#### 📊 Market")

        import re as _re, urllib.parse as _up

        # Known org → stable publications page (fallback when no deep URL provided)
        _ORG_PAGES = {
            "mckinsey":             "https://www.mckinsey.com/mgi/research",
            "mckinsey global":      "https://www.mckinsey.com/mgi/research",
            "gartner":              "https://www.gartner.com/en/research/publications",
            "bloomberg":            "https://www.bloomberg.com/professional/insights/",
            "bloombergnef":         "https://about.bnef.com/insights/",
            "bnef":                 "https://about.bnef.com/insights/",
            "iea":                  "https://www.iea.org/reports",
            "statista":             "https://www.statista.com/markets/",
            "roland berger":        "https://www.rolandberger.com/en/Insights/Publications/",
            "frost & sullivan":     "https://store.frost.com/reports.html",
            "frost":                "https://store.frost.com/reports.html",
            "deloitte":             "https://www2.deloitte.com/global/en/insights.html",
            "pwc":                  "https://www.pwc.com/gx/en/industries/",
            "ihs markit":           "https://ihsmarkit.com/research-analysis/",
            "s&p global":           "https://www.spglobal.com/marketintelligence/en/news-insights/research",
            "wood mackenzie":       "https://www.woodmac.com/reports/",
            "mordor":               "https://www.mordorintelligence.com/industry-reports",
            "mordor intelligence":  "https://www.mordorintelligence.com/industry-reports",
            "grand view":           "https://www.grandviewresearch.com/industry-analysis",
            "grand view research":  "https://www.grandviewresearch.com/industry-analysis",
            "allied market":        "https://www.alliedmarketresearch.com/market-research-report",
            "allied market research": "https://www.alliedmarketresearch.com/market-research-report",
            "marketsandmarkets":    "https://www.marketsandmarkets.com/Market-Reports/",
            "fortune business":     "https://www.fortunebusinessinsights.com/reports",
            "fortune business insights": "https://www.fortunebusinessinsights.com/reports",
            "idc":                  "https://www.idc.com/research/viewtoc",
            "precedence":           "https://www.precedenceresearch.com/",
            "technavio":            "https://www.technavio.com/report-store",
            "ibisworld":            "https://www.ibisworld.com/global/",
            "euromonitor":          "https://www.euromonitor.com/reports",
        }

        def _resolve_url(src_obj):
            """
            Test Claude's URL with a HEAD request.
            Returns (url, is_direct) where is_direct=True means the link lands on the actual page.
            Falls back to a targeted Google Search if the URL fails or is missing.
            """
            import urllib.parse as _up2
            raw_url = src_obj.get("url", "").strip()

            # Build the Google Search fallback first — always works
            q_parts = [src_obj.get("org",""), src_obj.get("title",""), src_obj.get("year","")]
            q = " ".join(p for p in q_parts if p).strip()
            google_url = f"https://www.google.com/search?q={_up2.quote(q)}"

            if not raw_url or not raw_url.startswith("http"):
                # No URL provided — try org page first, then Google
                org_lower = src_obj.get("org", "").lower()
                for org_key, org_page in _ORG_PAGES.items():
                    if org_key in org_lower:
                        return org_page, False
                return google_url, False

            # Check the URL has a real path (not just homepage)
            try:
                from urllib.parse import urlparse as _ulp
                _parsed = _ulp(raw_url)
                path = _parsed.path.strip("/")
                if not path or len(path) <= 3:
                    return google_url, False
            except Exception:
                return google_url, False

            # Live HEAD check — 4 second timeout
            try:
                resp = requests.head(raw_url, timeout=4, allow_redirects=True,
                                     headers={"User-Agent": "Mozilla/5.0"})
                if resp.status_code < 400:
                    return raw_url, True
            except Exception:
                pass

            return google_url, False

        def _pill_label(src_obj):
            """Short display label for a source pill."""
            org = src_obj.get("org", "")
            year = src_obj.get("year", "")
            title = src_obj.get("title", "")
            if org and year:
                return f"{org}, {year}"
            elif org:
                return org
            elif title:
                return title[:35] + ("…" if len(title) > 35 else "")
            return "Source"

        def render_pills_structured(sources_list):
            """Return HTML for individual pill <a> tags from a list of source dicts.
            🔗 = direct link verified live. 🔍 = Google Search (guaranteed to open)."""
            if not sources_list:
                return '<span style="color:#4a6fa5;font-size:10px;">no source available</span>'
            html = ""
            for src in sources_list:
                url, is_direct = _resolve_url(src)
                lbl = _pill_label(src)
                icon = "🔗" if is_direct else "🔍"
                html += (
                    f'<a href="{url}" target="_blank" '
                    f'style="display:inline-block;background:#1e3a5f;color:#93c5fd;'
                    f'font-size:10px;padding:3px 10px;border-radius:12px;margin:3px 2px 0;'
                    f'text-decoration:none;border:1px solid #2a4a70;'
                    f'white-space:nowrap;line-height:1.5;">{icon} {lbl}</a>'
                )
            return html

        def _compat_field(field_val, fallback_key=None):
            """Handle both new structured format and legacy string format gracefully."""
            if isinstance(field_val, dict):
                return field_val
            # Legacy string format — wrap it
            raw = str(field_val) if field_val else "N/A"
            val = raw.split("[Source:")[0].strip().split("(")[0].split(";")[0].strip()
            return {"value": val, "year": "", "period": "", "sources": []}

        cur_obj  = _compat_field(market.get("market_size_current") or market.get("market_size_2024"))
        fore_obj = _compat_field(market.get("market_size_forecast") or market.get("market_size_2030"))
        cagr_obj = _compat_field(market.get("cagr"))

        # ── Big-number cards ──────────────────────────────────
        mc1, mc2, mc3 = st.columns(3)
        cards = [
            (mc1, "CURRENT MARKET SIZE",  cur_obj["value"],  cur_obj.get("year",""),         cur_obj.get("sources",[])),
            (mc2, "FORECAST MARKET SIZE", fore_obj["value"], fore_obj.get("year",""),         fore_obj.get("sources",[])),
            (mc3, "CAGR",                 cagr_obj["value"], cagr_obj.get("period",""),       cagr_obj.get("sources",[])),
        ]
        for col, top_label, num_val, sub_label, srcs in cards:
            pills_html = render_pills_structured(srcs)
            col.markdown(f"""
<div style="background:#1a2d45;border-radius:8px;padding:18px 14px 14px;text-align:center;min-height:140px;">
  <div style="color:#94a3b8;font-size:10px;letter-spacing:1.5px;font-weight:600;margin-bottom:6px;">{top_label}</div>
  <div style="color:#60a5fa;font-size:36px;font-weight:700;line-height:1.1;">{num_val}</div>
  <div style="color:#64748b;font-size:11px;margin-top:4px;margin-bottom:8px;">{sub_label}</div>
  <div style="margin-top:8px;line-height:2.0;">{pills_html}</div>
</div>""", unsafe_allow_html=True)

        st.markdown("<div style='margin-top:4px'></div>", unsafe_allow_html=True)
        st.caption("🔗 = direct link verified · 🔍 = Google Search to find the report (both always open)")

        col_l, col_r = st.columns(2)
        col_l.markdown(f"**Maturity** · {market.get('market_maturity','')}")
        col_r.markdown(f"**Geography** · {market.get('geographic_focus','')}")
        if market.get("growth_drivers"):
            drivers = market["growth_drivers"]
            show_n   = min(3, len(drivers))
            total_n  = len(drivers)
            drivers_html = "".join(
                f'<div style="display:flex;gap:10px;margin:6px 0;color:#e2e8f0;font-size:13px;'
                f'"><span style="color:#60a5fa;flex-shrink:0;">›</span><span>{drv}</span></div>'
                for drv in drivers[:3]
            )
            extra_html = '<div style="color:#64748b;font-size:11px;margin-top:6px;">Full list in the downloaded report.</div>' if total_n > 3 else ""
            st.markdown(
                '<div style="background:#1a2d45;border-radius:6px;padding:12px 16px;margin:8px 0 4px 0;">' +
                f'<div style="color:#60a5fa;font-size:11px;letter-spacing:1px;font-weight:600;margin-bottom:8px;">GROWTH DRIVERS &middot; TOP {show_n} OF {total_n}</div>' +
                drivers_html + extra_html +
                '</div>',
                unsafe_allow_html=True
            )
        st.markdown("---")

        # ── Sector cluster chart ──────────────────────────────
        st.markdown("#### 🎯 Schaeffler Sector Fit")
        sector_scores = sectors.get("sector_scores",{})
        primary = sectors.get("primary_sectors",[])
        if sector_scores:
            names  = list(sector_scores.keys())
            vals   = [sector_scores[s]["score"] for s in names]
            colors = [BLUE if s in primary else DIM for s in names]
            fig_bar = go.Figure(go.Bar(
                x=vals, y=names, orientation="h", marker_color=colors,
                text=[f"{v}/10" for v in vals], textposition="outside",
                textfont=dict(color=WHITE,size=11),
                hovertemplate="<b>%{y}</b><br>%{x}/10<extra></extra>"
            ))
            fig_bar.update_layout(
                plot_bgcolor=BG, paper_bgcolor=BG, height=340,
                xaxis=dict(range=[0,12],showgrid=False,zeroline=False,
                           tickfont=dict(color=WHITE),title_font=dict(color=DIM)),
                yaxis=dict(showgrid=False,zeroline=False,tickfont=dict(color=WHITE,size=11)),
                margin=dict(l=10,r=60,t=10,b=30), font=dict(color=WHITE)
            )
            st.plotly_chart(fig_bar, use_container_width=True)
            st.caption(f"Primary sectors: {', '.join(primary)}")
        st.markdown("---")

        # ── Competitive landscape — summary only ──────────────
        st.markdown("#### 🏢 Competitive Landscape")
        cc1,cc2,cc3 = st.columns(3)
        cc1.metric("Intensity", comp.get("competitive_intensity",""))
        cc2.metric("Competition Score", f"{comp.get('competition_score',0)}/10")
        cc3.metric("Key Players", len(comp.get("competitors",[])))
        st.markdown(f"**White space** · {comp.get('white_space','')}")
        st.markdown(f"**Schaeffler edge** · {comp.get('schaeffler_advantage','')}")
        st.caption("Full competitor list and detailed analysis in the downloadable report.")


        # ── Download report ──────────────────────────────────
        st.markdown("---")
        _one_click_dl(
            T("dl_market"),
            lambda: generate_market_report(idea, quadrant, s1c, market, comp, sectors, weights, final),
            f"Schaeffler_Market_Intelligence_{datetime.now().strftime('%Y%m%d')}.docx"
        )

        # ── Chat ──────────────────────────────────────────────
        st.markdown("---")
        st.subheader(T("s2_chat_header"))
        for msg in st.session_state.s2_chat:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        user_q = st.chat_input(T("s2_chat_ph"))
        if user_q:
            st.session_state.s2_chat.append({"role":"user","content":user_q})
            with st.chat_message("user"):
                st.markdown(user_q)
            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    ctx = f"""You are a senior market analyst discussing market intelligence results for a Schaeffler innovation idea.
Idea: {idea} | Quadrant: {quadrant}
Market: {market.get('market_name','')} | Size 2024: {_mval(market.get('market_size_current') or market.get('market_size_2024',''))} | CAGR: {_mval(market.get('cagr',''))}
Competitive intensity: {comp.get('competitive_intensity','')} | White space: {comp.get('white_space','')}
Primary sectors (Schaeffler's 10 clusters): {', '.join(primary)} | Final score: {final}/10
Be specific, cite sources where possible, 3-4 sentences max."""
                    history = [{"role":m["role"],"content":m["content"]} for m in st.session_state.s2_chat]
                    reply = call_claude_chat(ctx, history)
                    st.markdown(reply)
                    st.session_state.s2_chat.append({"role":"assistant","content":reply})

        # ── Continue ──────────────────────────────────────────
        st.markdown("---")
        st.success(T("s2_success").format(score=final))
        if st.button(T("s2_continue"), type="primary", key="s2_continue"):
            st.session_state.active_stage = 3
            st.rerun()

        if st.button(T("s2_rerun"), key="s2_rerun"):
            st.session_state.s2_step = "intro"
            st.session_state.s2_data = {}
            st.session_state.s2_chat = []
            st.session_state.s2_report_buf = None
            st.rerun()





# ════════════════════════════════════════════════════════════
# STAGE 03 — PATENT INTELLIGENCE
# ════════════════════════════════════════════════════════════
elif st.session_state.active_stage == 3:
    st.markdown(f"## {T('s3_title')}")
    st.markdown(f"""<div style="background:#1a2d45;border-radius:6px;padding:14px 18px;margin-bottom:16px;border-left:4px solid #2E75B6;">
<div style="color:#60a5fa;font-size:11px;letter-spacing:1px;font-weight:600;">{T('s1_what_label')}</div>
<div style="color:#e2e8f0;font-size:13px;margin-top:6px;">{T('s3_what')}</div>
<div style="color:#94a3b8;font-size:12px;margin-top:8px;">{T('s3_you_get')}</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    idea     = st.session_state.s1_idea
    s1c      = st.session_state.s1_classification
    quadrant = s1c.get("quadrant", "RADICAL")

    # ── Session state for stage 03 ────────────────────────────
    for k, v in {"s3_step":"intro","s3_data":{},"s3_chat":[]}.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # ── Intro ─────────────────────────────────────────────────
    if st.session_state.s3_step == "intro":
        st.info(f"**Idea:** {idea}")
        st.markdown(f"**Quadrant:** {quadrant}")
        if st.button(T("s3_run_btn"), type="primary"):
            st.session_state.s3_step = "running"
            st.rerun()

    # ── Running ───────────────────────────────────────────────
    elif st.session_state.s3_step == "running":
        progress = st.progress(0)
        status   = st.empty()

        status.markdown("🔍 Analysing external patent landscape...")
        progress.progress(20)

        system_landscape = """You are a patent intelligence analyst specialising in industrial technology.
Analyse the external patent landscape for this innovation idea.

RULES:
- Focus on COMPANIES filing patents, not individual patents
- Classify each company as: Competitor / Customer / Research Institution / Patent Troll / Adjacent Player
- Every claim must have [Source: org, year] where possible
- Be specific about technology sub-areas

Return ONLY valid JSON:
{
  "technology_keywords": ["3-5 key patent search terms for this idea"],
  "landscape_summary": "2-3 sentences on overall patent activity in this space",
  "activity_level": "Low / Moderate / High / Very High",
  "filing_trend": "Increasing / Stable / Decreasing",
  "filing_trend_rationale": "one sentence",
  "patent_landscape_score": integer 1-10 (9-10=very open; 7-8=some activity/clear gaps; 5-6=moderate density; 3-4=dense; 1-2=saturated),
  "key_filers": [
    {
      "company": "company name",
      "type": "Competitor / Customer / Research Institution / Adjacent Player",
      "focus": "one sentence on what they are patenting",
      "threat_level": "Low / Medium / High",
      "schaeffler_relationship": "Direct competitor / Potential customer / Partner / Unknown",
      "source": "Source: X, Year"
    }
  ],
  "white_spaces": ["white space opportunity 1", "white space opportunity 2", "white space opportunity 3"],
  "patent_landscape_score": integer 1-10 for landscape openness. Use: 9-10=very few filings, open territory; 7-8=some activity but clear gaps; 5-6=moderate filing density; 3-4=dense filing landscape; 1-2=saturated with patents by large incumbents
}"""

        try:
            raw = call_claude(system_landscape,
                f"Idea: {idea}\nQuadrant: {quadrant}\nTech novelty: {s1c.get('technology_novelty','')}", max_tokens=2000)
            raw_clean = raw.strip().replace("```json","").replace("```","").strip()
            fb = raw_clean.find("{"); lb = raw_clean.rfind("}") + 1
            if fb >= 0: raw_clean = raw_clean[fb:lb]
            landscape = json.loads(raw_clean)
        except Exception as e:
            landscape = {"technology_keywords":[],"landscape_summary":"Analysis unavailable.",
                        "activity_level":"N/A","filing_trend":"N/A","filing_trend_rationale":"",
                        "key_filers":[],"white_spaces":[],"patent_landscape_score":5}

        progress.progress(50)
        status.markdown("🏢 Mapping filers onto Schaeffler's Ansoff matrix...")

        system_ansoff = """You are a Schaeffler Group patent strategist.
Map patent filing companies onto Schaeffler's modified Ansoff matrix based on where their patents sit.

The matrix axes (MUST match Schaeffler's Stage 1 Innovation Framework):
- X axis: Technology Dimension (0=Existing/Established Technology → 10=New to the World)
- Y axis: Market Dimension (0=Existing/Established Market → 10=New to the World)

Quadrant positions (identical to Schaeffler's quadrant classifier):
- EXPLOIT  (bottom-left):  existing tech  + existing market  — x_score 0–5,  y_score 0–5
- EXTEND   (top-left):     existing tech  + new market       — x_score 0–5,  y_score 5–10
- RADICAL  (top-right):    new tech       + new market       — x_score 5–10, y_score 5–10
- DISRUPT  (bottom-right): new tech       + existing market  — x_score 5–10, y_score 0–5

For each company, assign:
- matrix_position: which quadrant their patent activity sits in
- x_score: 0-10 (0=existing technology, 10=new to the world technology)
- y_score: 0-10 (0=existing market, 10=new to the world market)

Also map where SCHAEFFLER's own known IP sits relative to this idea.

Return ONLY valid JSON:
{
  "filer_positions": [
    {
      "company": "company name",
      "matrix_position": "EXPLOIT/EXTEND/RADICAL/DISRUPT",
      "x_score": 0-10,
      "y_score": 0-10,
      "rationale": "one sentence"
    }
  ],
  "schaeffler_position": {
    "matrix_position": "EXPLOIT/EXTEND/RADICAL/DISRUPT",
    "x_score": 0-10,
    "y_score": 0-10,
    "existing_ip": "one sentence on what Schaeffler already has in this space",
    "gap": "one sentence on the IP gap this idea addresses"
  },
  "idea_position": {
    "x_score": 0-10,
    "y_score": 0-10
  },
  "novelty_signal": "Strong / Moderate / Weak",
  "novelty_rationale": "one sentence",
  "ip_risk": "Low / Medium / High",
  "ip_risk_rationale": "one sentence"
}"""

        # Pass FULL filer objects (not just names) so Claude can map every one
        key_filers = landscape.get("key_filers", [])
        filers_full_context = json.dumps([
            {"company": f.get("company",""), "type": f.get("type",""), "focus": f.get("focus",""), "threat_level": f.get("threat_level","")}
            for f in key_filers
        ])
        try:
            raw2 = call_claude(system_ansoff,
                f"Idea: {idea}\nQuadrant: {quadrant}\nTech keywords: {landscape.get('technology_keywords','')}\n"
                f"IMPORTANT: You MUST map ALL {len(key_filers)} filers listed below. Do not skip any.\n"
                f"Key filers (map every single one): {filers_full_context}",
                max_tokens=max(1800, len(key_filers) * 200 + 800))
            raw2_clean = raw2.strip().replace("```json","").replace("```","").strip()
            fb2 = raw2_clean.find("{"); lb2 = raw2_clean.rfind("}") + 1
            if fb2 >= 0: raw2_clean = raw2_clean[fb2:lb2]
            ansoff_data = json.loads(raw2_clean)
        except Exception as e:
            ansoff_data = {"filer_positions":[],"schaeffler_position":{"matrix_position":"EXPLOIT","x_score":2,"y_score":2,"existing_ip":"N/A","gap":"N/A"},
                          "idea_position":{"x_score":7,"y_score":7},"novelty_signal":"Moderate","novelty_rationale":"","ip_risk":"Medium","ip_risk_rationale":""}

        # ── Guarantee every key_filer appears in filer_positions ──────────
        # Build a lookup of which companies already have positions
        positioned_companies = {fp.get("company","").lower() for fp in ansoff_data.get("filer_positions", [])}
        # Quadrant → default score ranges for auto-placement fallback
        # Quadrant → default score ranges (X=Technology, Y=Market — matches Stage 1 convention)
        # EXPLOIT=bottom-left(low x,low y), EXTEND=top-left(low x,high y),
        # RADICAL=top-right(high x,high y), DISRUPT=bottom-right(high x,low y)
        type_defaults = {
            "Competitor":          ("EXPLOIT", 3.0, 3.0),   # established tech + established market
            "Customer":            ("EXTEND",  2.5, 6.5),   # established tech + new market
            "Research Institution":("RADICAL", 7.0, 7.5),  # new tech + new market
            "Adjacent Player":     ("DISRUPT", 6.5, 3.5),  # new tech + established market
            "Patent Troll":        ("EXPLOIT", 2.0, 2.0),  # established tech + established market
        }
        import random
        for i, f in enumerate(key_filers):
            name = f.get("company","")
            if name.lower() not in positioned_companies and name:
                quad, bx, by = type_defaults.get(f.get("type","Adjacent Player"), ("EXPLOIT", 4.0, 4.0))
                # Small deterministic nudge so overlapping filers spread out
                nudge_x = ((i * 0.7) % 2.0) - 1.0
                nudge_y = ((i * 1.1) % 2.0) - 1.0
                ansoff_data.setdefault("filer_positions", []).append({
                    "company": name,
                    "type": f.get("type","Adjacent Player"),
                    "matrix_position": quad,
                    "x_score": round(min(9.5, max(0.5, bx + nudge_x)), 1),
                    "y_score": round(min(9.5, max(0.5, by + nudge_y)), 1),
                    "rationale": f.get("focus","Auto-placed based on filer type")
                })

        progress.progress(80)
        status.markdown("📊 Calculating patent intelligence score...")

        # Patent score: landscape openness + novelty signal + ip risk
        landscape_score = float(landscape.get("patent_landscape_score", 5))
        novelty_map = {"Strong":9,"Moderate":6,"Weak":3}
        ip_risk_map  = {"Low":8,"Medium":5,"High":2}
        novelty_score = novelty_map.get(ansoff_data.get("novelty_signal","Moderate"), 6)
        ip_score      = ip_risk_map.get(ansoff_data.get("ip_risk","Medium"), 5)
        final_patent  = round(landscape_score*0.40 + novelty_score*0.35 + ip_score*0.25, 1)

        st.session_state.s3_data = {
            "landscape": landscape,
            "ansoff_data": ansoff_data,
            "novelty_score": novelty_score,
            "ip_score": ip_score,
            "landscape_score": landscape_score,
            "final_score": final_patent
        }

        progress.progress(100)
        status.markdown("✓ Complete.")
        time.sleep(0.5)
        st.session_state.s3_step = "done"
        st.rerun()

    # ── Results ───────────────────────────────────────────────
    elif st.session_state.s3_step == "done":
        d            = st.session_state.s3_data
        landscape    = d["landscape"]
        ansoff_data  = d["ansoff_data"]
        final        = d["final_score"]

        # Score banner
        score_col = "#22c55e" if final>=7 else "#f59e0b" if final>=4 else "#ef4444"
        st.markdown(f"""
<div style="background:#0f1e35;border-radius:8px;padding:20px 24px;margin-bottom:20px;border:1px solid #2a4a70;">
  <div style="color:#94a3b8;font-size:9px;letter-spacing:2px;font-weight:700;font-family:Arial,sans-serif;margin-bottom:4px;">PATENT INTELLIGENCE SCORE</div>
  <div style="color:{score_col};font-size:42px;font-weight:700;line-height:1;">{final}<span style="font-size:18px;color:#94a3b8;"> / 10</span></div>
  <div style="color:{WHITE};font-size:13px;margin-top:6px;opacity:0.8;">{landscape.get("landscape_summary","")}</div>
</div>
""", unsafe_allow_html=True)

        # Score breakdown
        col1, col2, col3 = st.columns(3)
        col1.metric("Landscape Openness", f"{d['landscape_score']:.1f} / 10", "40% weight")
        col2.metric("Novelty Signal",      f"{d['novelty_score']:.1f} / 10",  "35% weight")
        col3.metric("IP Risk",             f"{d['ip_score']:.1f} / 10",       "25% weight")
        st.markdown("---")

        # ── Patent activity overview ──────────────────────────
        st.markdown("#### 📋 Patent Activity")
        c1, c2 = st.columns(2)
        c1.metric("Filing Activity", landscape.get("activity_level",""))
        c2.metric("Trend", landscape.get("filing_trend",""))
        st.caption(landscape.get("filing_trend_rationale",""))

        if landscape.get("technology_keywords"):
            kws = "  ·  ".join([f"`{k}`" for k in landscape["technology_keywords"]])
            st.markdown(f"**Search terms:** {kws}")
        st.markdown("---")

        # ── Ansoff matrix with all filers plotted ─────────────
        st.markdown("#### 🗺️ Patent Position Map — Schaeffler Ansoff Matrix")

        # ── Pattern commentary — generated from the filer data ──
        filer_positions = ansoff_data.get("filer_positions", [])
        schaeffler_pos  = ansoff_data.get("schaeffler_position", {})
        idea_pos        = ansoff_data.get("idea_position", {"x_score":7,"y_score":7})

        # Enrich filer_positions with 'type' from key_filers where missing
        key_filers_lookup = {f.get("company","").lower(): f for f in landscape.get("key_filers",[])}
        for fp in filer_positions:
            if not fp.get("type"):
                kf = key_filers_lookup.get(fp.get("company","").lower(), {})
                fp["type"] = kf.get("type", "Adjacent Player")

        # Nudge overlapping points apart so labels don't stack
        used_positions = []
        for fp in filer_positions:
            x, y = fp.get("x_score", 5), fp.get("y_score", 5)
            attempts = 0
            while any(abs(x-ux) < 0.6 and abs(y-uy) < 0.6 for ux,uy in used_positions) and attempts < 8:
                x = round(min(9.5, max(0.5, x + 0.4)), 1)
                y = round(min(9.5, max(0.5, y + 0.3)), 1)
                attempts += 1
            fp["x_score"] = x
            fp["y_score"] = y
            used_positions.append((x, y))

        # Compute pattern stats for commentary
        if filer_positions:
            idea_x = idea_pos.get("x_score", 7)
            idea_y = idea_pos.get("y_score", 7)
            # Quadrant distribution
            q_counts = {"EXPLOIT":0,"EXTEND":0,"RADICAL":0,"DISRUPT":0}
            close_filers = []
            for fp in filer_positions:
                q_counts[fp.get("matrix_position","EXPLOIT")] = q_counts.get(fp.get("matrix_position","EXPLOIT"),0) + 1
                # Close = within 2 units on both axes
                dx = abs(fp.get("x_score",5) - idea_x)
                dy = abs(fp.get("y_score",5) - idea_y)
                if dx <= 2.5 and dy <= 2.5:
                    close_filers.append(fp.get("company",""))

            dominant_q  = max(q_counts, key=q_counts.get)
            dominant_n  = q_counts[dominant_q]
            total_filers = len(filer_positions)
            competitors_count = sum(1 for fp in filer_positions if fp.get("type")=="Competitor")
            research_count    = sum(1 for fp in filer_positions if fp.get("type")=="Research Institution")

            # Build pointer bullets
            pointers = []
            if close_filers:
                pointers.append(f"**{len(close_filers)} filer{'s' if len(close_filers)>1 else ''} sitting close to your idea** — {', '.join(close_filers[:3])}{'...' if len(close_filers)>3 else ''}. Check IP risk before external disclosure.")
            else:
                pointers.append(f"**No filers plotted close to your idea's position** — the immediate IP zone appears uncontested.")
            if dominant_n > total_filers * 0.5 and dominant_q != quadrant:
                pointers.append(f"**Most activity is in the {dominant_q} quadrant** ({dominant_n} of {total_filers} filers) — your idea targets a less contested zone.")
            if competitors_count >= 3:
                pointers.append(f"**{competitors_count} direct competitors** identified — conduct freedom-to-operate analysis before committing R&D budget.")
            elif competitors_count == 0:
                pointers.append("**No direct competitors mapped** — validate this through a formal patent search before relying on it.")
            if research_count >= 2:
                pointers.append(f"**{research_count} research institutions** active in this space — potential co-development or licensing partners.")
            sch_x = schaeffler_pos.get("x_score", 2)
            sch_y = schaeffler_pos.get("y_score", 2)
            gap_x = abs(idea_x - sch_x)
            gap_y = abs(idea_y - sch_y)
            if gap_x > 3 or gap_y > 3:
                pointers.append(f"**Schaeffler's existing IP sits far from this idea** — significant new IP investment likely required to protect the concept.")
            else:
                pointers.append(f"**Schaeffler's existing IP is adjacent to this idea** — existing patents may provide partial protection or a foundation to build from.")

            # Render commentary box — convert **markdown bold** to <b> for HTML context
            import re as _re
            def md_bold_to_html(text):
                return _re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
            bullets_html = "".join(f'<div style="margin:5px 0;color:#e2e8f0;font-size:13px;">› {md_bold_to_html(p)}</div>' for p in pointers)
            st.markdown(f"""
<div style="background:#1a2d45;border-left:3px solid #2E75B6;border-radius:4px;padding:12px 16px;margin-bottom:14px;">
<div style="color:#60a5fa;font-size:11px;letter-spacing:1px;font-weight:600;margin-bottom:8px;">PATTERN NOTES</div>
{bullets_html}
</div>
""", unsafe_allow_html=True)

        st.caption("Each point = a company's patent filing position. Your idea shown in green. Schaeffler's existing IP shown in orange.")

        fig = go.Figure()

        # Quadrant shading — X=Technology, Y=Market (matches Stage 1 Ansoff convention)
        # EXPLOIT=bottom-left, EXTEND=top-left, RADICAL=top-right, DISRUPT=bottom-right
        q_fills = [
            dict(x=[0,5,5,0],   y=[0,0,5,5],   name="EXPLOIT", fill="#1a2d45", lx=2.5,ly=2.5),
            dict(x=[0,5,5,0],   y=[5,5,10,10], name="EXTEND",  fill="#1e3a5f", lx=2.5,ly=7.5),
            dict(x=[5,10,10,5], y=[5,5,10,10], name="RADICAL", fill="#1F3864", lx=7.5,ly=7.5),
            dict(x=[5,10,10,5], y=[0,0,5,5],   name="DISRUPT", fill="#0d2137", lx=7.5,ly=2.5),
        ]
        for q in q_fills:
            fig.add_trace(go.Scatter(
                x=q["x"]+[q["x"][0]], y=q["y"]+[q["y"][0]],
                fill="toself", fillcolor=q["fill"],
                line=dict(color="#2a4a70",width=1),
                mode="lines", showlegend=False, hoverinfo="skip"
            ))
            fig.add_annotation(x=q["lx"],y=q["ly"],text=f"<b>{q['name']}</b>",
                showarrow=False, font=dict(size=12,color="#4a6fa5"))

        # Grid lines
        fig.add_shape(type="line",x0=5,x1=5,y0=0,y1=10,line=dict(color="#2a4a70",width=1.5,dash="dot"))
        fig.add_shape(type="line",x0=0,x1=10,y0=5,y1=5,line=dict(color="#2a4a70",width=1.5,dash="dot"))

        # Competitor/filer points
        type_colours = {
            "Competitor":         "#ef4444",
            "Customer":           "#60a5fa",
            "Research Institution":"#a78bfa",
            "Adjacent Player":    "#f59e0b",
            "Patent Troll":       "#6b7280",
        }
        for fp in filer_positions:
            col = type_colours.get(fp.get("type","Adjacent Player"), "#6b7280")
            fig.add_trace(go.Scatter(
                x=[fp.get("x_score",5)], y=[fp.get("y_score",5)],
                mode="markers+text",
                marker=dict(size=12, color=col, line=dict(color="white",width=1.5)),
                text=[f"  {fp.get('company','')}"],
                textposition="middle right",
                textfont=dict(size=10, color="#e2e8f0"),
                name=fp.get("type",""),
                hovertemplate=f"<b>{fp.get('company','')}</b><br>{fp.get('type','')}<br>{fp.get('rationale','')}<extra></extra>",
                showlegend=False
            ))

        # Schaeffler existing IP
        if schaeffler_pos:
            fig.add_trace(go.Scatter(
                x=[schaeffler_pos.get("x_score",2)], y=[schaeffler_pos.get("y_score",2)],
                mode="markers+text",
                marker=dict(size=16, color="#f97316", symbol="diamond",
                           line=dict(color="white",width=2)),
                text=["  Schaeffler IP"],
                textposition="middle right",
                textfont=dict(size=11,color="#f97316",family="Arial Bold"),
                showlegend=False,
                hovertemplate=f"<b>Schaeffler existing IP</b><br>{schaeffler_pos.get('existing_ip','')}<extra></extra>"
            ))

        # This idea
        fig.add_trace(go.Scatter(
            x=[idea_pos.get("x_score",7)], y=[idea_pos.get("y_score",7)],
            mode="markers+text",
            marker=dict(size=18, color="#22c55e", symbol="star",
                       line=dict(color="white",width=2)),
            text=["  Your idea"],
            textposition="middle right",
            textfont=dict(size=12,color="#22c55e",family="Arial Bold"),
            showlegend=False,
            hovertemplate="<b>Your innovation idea</b><extra></extra>"
        ))

        # Axis labels
        for ann in [
            dict(x=2.5,y=-0.9,text="Established Tech",angle=0),
            dict(x=7.5,y=-0.9,text="New Tech",angle=0),
            dict(x=-1.1,y=2.5,text="Established Market",angle=-90),
            dict(x=-1.1,y=7.5,text="New Market",angle=-90),
        ]:
            fig.add_annotation(x=ann["x"],y=ann["y"],text=ann["text"],
                showarrow=False,font=dict(size=10,color="#e2e8f0"),textangle=ann["angle"])

        fig.update_layout(
            plot_bgcolor=BG, paper_bgcolor=BG, height=460,
            xaxis=dict(range=[-1.5,12],showticklabels=False,showgrid=False,zeroline=False,
                      title="Technology Dimension →",title_font=dict(size=11,color=DIM)),
            yaxis=dict(range=[-1.5,11],showticklabels=False,showgrid=False,zeroline=False,
                      title="Market Dimension →",title_font=dict(size=11,color=DIM)),
            margin=dict(l=70,r=20,t=20,b=55), font=dict(color=WHITE)
        )
        st.plotly_chart(fig, use_container_width=True)

        # Legend
        legend_items = [("🔴","Competitor"),("🔵","Customer"),("🟣","Research Institution"),("🟡","Adjacent Player"),("🟠","Schaeffler existing IP"),("🟢","Your idea")]
        st.markdown("  ".join([f"{e} {l}" for e,l in legend_items]))
        st.markdown("---")

        # ── Key filers summary ────────────────────────────────
        st.markdown("#### 🏢 Key Patent Filers")
        filers = landscape.get("key_filers",[])
        if filers:
            top_filers = sorted(filers, key=lambda x: {"High":0,"Medium":1,"Low":2}.get(x.get("threat_level","Medium"),1))[:5]
            if len(filers) > 5:
                st.caption(f"Showing top 5 of {len(filers)} filers by threat level. Full list in the downloaded report.")
            for fi in top_filers:
                threat_col = {"High":"#ef4444","Medium":"#f59e0b","Low":"#22c55e"}.get(fi.get("threat_level","Medium"),"#6b7280")
                type_col   = type_colours.get(fi.get("type","Adjacent Player"),"#6b7280")
                st.markdown(f"""
<div style="background:#1a2d45;border-radius:6px;padding:10px 14px;margin:5px 0;display:flex;align-items:center;gap:12px;">
  <div style="min-width:130px;color:{WHITE};font-weight:600;font-size:13px;">{fi.get('company','')}</div>
  <div style="background:{type_col}22;color:{type_col};font-size:11px;padding:2px 8px;border-radius:10px;min-width:120px;text-align:center;">{fi.get('type','')}</div>
  <div style="background:{threat_col}22;color:{threat_col};font-size:11px;padding:2px 8px;border-radius:10px;min-width:80px;text-align:center;">⚡ {fi.get('threat_level','')} threat</div>
  <div style="color:#94a3b8;font-size:12px;flex:1;">{fi.get('focus','')}</div>
</div>
""", unsafe_allow_html=True)

        st.markdown("---")

        # ── Schaeffler IP position ────────────────────────────
        st.markdown("#### 🏭 Schaeffler IP Position")
        sc1, sc2, sc3 = st.columns(3)
        sc1.metric("Novelty Signal", ansoff_data.get("novelty_signal",""))
        sc2.metric("IP Risk",        ansoff_data.get("ip_risk",""))
        sc3.metric("IP Score",       f"{d['ip_score']:.1f} / 10")
        st.caption(f"Existing Schaeffler IP: {schaeffler_pos.get('existing_ip','')}")
        st.caption(f"IP gap this idea addresses: {schaeffler_pos.get('gap','')}")

        # White spaces
        white_spaces = landscape.get("white_spaces",[])
        if white_spaces:
            st.markdown("**IP white spaces identified:**")
            for ws in white_spaces[:3]:
                st.markdown(f"- {ws}")
            if len(white_spaces) > 3:
                st.caption(f"+{len(white_spaces)-3} more in the downloaded report.")

        # ── Download report ───────────────────────────────────
        st.markdown("---")
        _one_click_dl(
            T("dl_patent"),
            lambda: generate_patent_report(idea, quadrant, s1c, landscape, ansoff_data, d),
            f"Schaeffler_Patent_Intelligence_{datetime.now().strftime('%Y%m%d')}.docx"
        )

        # ── Chat ──────────────────────────────────────────────
        st.markdown("---")
        st.subheader(T("s3_chat_header"))
        for msg in st.session_state.s3_chat:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        user_q = st.chat_input(T("s3_chat_ph"))
        if user_q:
            st.session_state.s3_chat.append({"role":"user","content":user_q})
            with st.chat_message("user"):
                st.markdown(user_q)
            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    ctx = f"""You are a Schaeffler patent intelligence expert.
Idea: {idea} | Quadrant: {quadrant}
Patent activity: {landscape.get('activity_level','')} | Trend: {landscape.get('filing_trend','')}
Novelty signal: {ansoff_data.get('novelty_signal','')} | IP risk: {ansoff_data.get('ip_risk','')}
White spaces: {', '.join(landscape.get('white_spaces',[]))}
Schaeffler IP gap: {schaeffler_pos.get('gap','')}
Be specific and concise — 2-4 sentences."""
                    history = [{"role":m["role"],"content":m["content"]} for m in st.session_state.s3_chat]
                    reply = call_claude_chat(ctx, history)
                    st.markdown(reply)
                    st.session_state.s3_chat.append({"role":"assistant","content":reply})

        # ── Continue ──────────────────────────────────────────
        st.markdown("---")
        st.success(T("s3_success").format(score=final))
        if st.button(T("s3_continue"), type="primary", key="s3_continue"):
            st.session_state.active_stage = 4
            st.rerun()

        if st.button(T("s3_rerun"), key="s3_rerun"):
            st.session_state.s3_step = "intro"
            st.session_state.s3_data = {}
            st.session_state.s3_chat = []
            st.session_state.s3_report_buf = None
            st.rerun()



# ════════════════════════════════════════════════════════════
# STAGE 04 — TECHNICAL FEASIBILITY
# ════════════════════════════════════════════════════════════
elif st.session_state.active_stage == 4:
    st.markdown(f"## {T('s4_title')}")
    st.markdown(f"""<div style="background:#1a2d45;border-radius:6px;padding:14px 18px;margin-bottom:16px;border-left:4px solid #2E75B6;">
<div style="color:#60a5fa;font-size:11px;letter-spacing:1px;font-weight:600;">{T('s1_what_label')}</div>
<div style="color:#e2e8f0;font-size:13px;margin-top:6px;">{T('s4_what')}</div>
<div style="color:#94a3b8;font-size:12px;margin-top:8px;">{T('s4_you_get')}</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    idea     = st.session_state.s1_idea
    s1c      = st.session_state.s1_classification
    quadrant = s1c.get("quadrant", "RADICAL")

    for k, v in {"s4_step":"intro","s4_data":{},"s4_chat":[]}.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # ── Intro ─────────────────────────────────────────────────
    if st.session_state.s4_step == "intro":
        st.info(f"**Idea:** {idea}")
        st.markdown(f"**Quadrant:** {quadrant}")
        if st.button(T("s4_run_btn"), type="primary"):
            st.session_state.s4_step = "running"
            st.rerun()

    # ── Running ───────────────────────────────────────────────
    elif st.session_state.s4_step == "running":
        progress = st.progress(0)
        status   = st.empty()

        # ── Step A: Existence check & evidence ───────────────
        status.markdown("🔬 Checking for technology demonstrations...")
        progress.progress(20)

        system_existence = """You are a technology intelligence analyst specialising in industrial and automotive R&D.
Check whether the core technology behind this innovation idea has been demonstrated anywhere.

Look for evidence in: academic research, university labs, startup products, government programmes, 
industry pilots, defence/aerospace applications (often ahead of industrial), and adjacent industries.

RULES:
- Every evidence item MUST include a source citation [Source: org/journal/publication, year]
- Only cite sources you are confident exist — do not fabricate
- Be specific about what was demonstrated vs what is still theoretical
- Rate confidence in each evidence item: High / Medium / Low

Return ONLY valid JSON:
{
  "technology_core": "one sentence describing the core technology mechanism",
  "existence_verdict": "Demonstrated / Partially Demonstrated / Research Stage / Theoretical",
  "existence_summary": "2-3 sentences on the state of the art",
  "evidence": [
    {
      "type": "Academic Paper / Startup / Pilot / Government Programme / Industry Deployment",
      "title": "name of the paper, company, or programme",
      "description": "one sentence on what was demonstrated",
      "source": "Source: org/journal, year",
      "confidence": "High / Medium / Low",
      "relevance": "Direct / Adjacent / Analogous"
    }
  ],
  "technology_gaps": ["gap 1 between current state and full deployment", "gap 2", "gap 3"],
  "time_to_readiness": "estimated years to production readiness",
  "keywords": ["6-10 key technical terms from this domain for keyword map"]
}"""

        try:
            raw = call_claude(system_existence,
                f"Idea: {idea}\nQuadrant: {quadrant}\nTech: {s1c.get('technology_novelty','')}", max_tokens=2000)
            # Strip everything before first { and after last }
            raw_clean = raw.strip()
            raw_clean = raw_clean.replace("```json","").replace("```","").strip()
            first_brace = raw_clean.find("{")
            last_brace  = raw_clean.rfind("}") + 1
            if first_brace >= 0:
                raw_clean = raw_clean[first_brace:last_brace]
            existence = json.loads(raw_clean)
        except Exception as e:
            st.warning(f"Evidence parsing issue: {e} — using fallback")
            existence = {"technology_core":"N/A","existence_verdict":"Research Stage","existence_summary":"",
                        "evidence":[],"technology_gaps":[],"time_to_readiness":"Estimated 5–8 years (fallback — re-run for accurate estimate)","keywords":[]}

        progress.progress(50)

        # ── Step B: Schaeffler-adapted TRL assessment ─────────
        status.markdown("📊 Assessing technology readiness level...")
        progress.progress(65)

        system_trl = """You are a Schaeffler R&D director assessing technology maturity.
Use the Schaeffler-adapted TRL framework (modified from NASA TRL for industrial/automotive context):

TRL 1 — Basic principles observed (theoretical concept only)
TRL 2 — Technology concept formulated (application identified, no testing)
TRL 3 — Experimental proof of concept (lab demonstration, key functions validated)
TRL 4 — Technology validated in lab (component/subsystem tested in controlled environment)
TRL 5 — Technology validated in relevant environment (prototype tested in industrial-like conditions)
TRL 6 — Technology demonstrated in relevant environment (system prototype demonstrated)
TRL 7 — System prototype demonstrated in operational environment (field trial or pilot)
TRL 8 — System complete and qualified (full production design, limited production run)
TRL 9 — Actual system proven in operational environment (commercial deployment at scale)

For Schaeffler's innovation pipeline, ideas typically enter at TRL 3-5.
TRL 1-2 = too early for Innovation; TRL 6+ = should be in Product Development.

Return ONLY valid JSON:
{
  "trl_level": 1-9,
  "trl_label": "TRL X — label from framework above",
  "trl_rationale": "2-3 sentences justifying this TRL rating based on evidence",
  "schaeffler_entry_readiness": "Too Early / Ready for Innovation / Ready for Product Development",
  "entry_rationale": "one sentence",
  "key_technical_risks": [
    {"risk": "technical risk description", "severity": "High/Medium/Low", "mitigation": "one sentence"},
    {"risk": "technical risk description", "severity": "High/Medium/Low", "mitigation": "one sentence"},
    {"risk": "technical risk description", "severity": "High/Medium/Low", "mitigation": "one sentence"}
  ],
  "analogous_schaeffler_technologies": "one sentence on which of Schaeffler's 8 Motion Product Families (Guide Motion/Transmit Motion/Control Motion/Generate Motion/Power Motion/Drive Motion/Energize Motion/Sustain Motion) this technology is closest to",
  "trl_score": integer 1-10 mapped directly from TRL level. TRL1=1, TRL2=2, TRL3=3.5, TRL4=5, TRL5=6, TRL6=7, TRL7=8, TRL8=9, TRL9=10
}"""

        try:
            raw2 = call_claude(system_trl,
                f"Idea: {idea}\nExistence verdict: {existence.get('existence_verdict','')}\nEvidence: {json.dumps(existence.get('evidence',[])[:3])}\nGaps: {existence.get('technology_gaps',[])}",
                max_tokens=1500)
            raw2_clean = raw2.strip().replace("```json","").replace("```","").strip()
            first_brace = raw2_clean.find("{")
            last_brace  = raw2_clean.rfind("}") + 1
            if first_brace >= 0:
                raw2_clean = raw2_clean[first_brace:last_brace]
            trl = json.loads(raw2_clean)
        except Exception as e:
            st.warning(f"TRL parsing issue: {e} — using fallback")
            trl = {"trl_level":3,"trl_label":"TRL 3 — Experimental proof of concept",
                  "trl_rationale":"","schaeffler_entry_readiness":"Ready for Innovation",
                  "entry_rationale":"","key_technical_risks":[],"analogous_schaeffler_technologies":"","trl_score":5}

        progress.progress(85)
        status.markdown("📐 Calculating feasibility score...")

        # Feasibility score: TRL score (50%) + existence quality (30%) + risk profile (20%)
        trl_score = float(trl.get("trl_score", 5))
        existence_map = {"Demonstrated":9,"Partially Demonstrated":6,"Research Stage":3,"Theoretical":1}
        existence_score = existence_map.get(existence.get("existence_verdict","Research Stage"), 5)
        risk_scores = [{"High":2,"Medium":5,"Low":8}.get(r.get("severity","Medium"),5)
                      for r in trl.get("key_technical_risks",[])]
        risk_score = sum(risk_scores)/len(risk_scores) if risk_scores else 5.0
        final_feasibility = round(trl_score*0.50 + existence_score*0.30 + risk_score*0.20, 1)

        st.session_state.s4_data = {
            "existence": existence,
            "trl": trl,
            "trl_score": trl_score,
            "existence_score": existence_score,
            "risk_score": round(risk_score,1),
            "final_score": final_feasibility
        }

        progress.progress(100)
        status.markdown("✓ Complete.")
        time.sleep(0.5)
        st.session_state.s4_step = "done"
        st.rerun()

    # ── Results ───────────────────────────────────────────────
    elif st.session_state.s4_step == "done":
        d         = st.session_state.s4_data
        existence = d["existence"]
        trl       = d["trl"]
        final     = d["final_score"]

        # ── Score banner ──────────────────────────────────────
        score_col = "#22c55e" if final>=7 else "#f59e0b" if final>=4 else "#ef4444"
        trl_level = trl.get("trl_level", 3)
        trl_pct   = int((trl_level / 9) * 100)
        entry_col = {"Too Early":"#ef4444","Ready for Innovation":"#22c55e",
                     "Ready for Product Development":"#60a5fa"}.get(
                     trl.get("schaeffler_entry_readiness",""), "#f59e0b")

        st.markdown(f"""
<div style="background:#0f1e35;border-radius:8px;padding:20px 24px;margin-bottom:20px;border:1px solid #2a4a70;">
  <div style="color:#94a3b8;font-size:9px;letter-spacing:2px;font-weight:700;font-family:Arial,sans-serif;margin-bottom:4px;">TECHNICAL FEASIBILITY SCORE</div>
  <div style="display:flex;align-items:flex-end;gap:24px;">
    <div style="color:{score_col};font-size:42px;font-weight:700;line-height:1;">{final}<span style="font-size:18px;color:#94a3b8;"> / 10</span></div>
    <div>
      <div style="color:{entry_col};font-size:14px;font-weight:600;">{trl.get("schaeffler_entry_readiness","")}</div>
      <div style="color:{WHITE};font-size:12px;opacity:0.7;">{existence.get("existence_verdict","")}</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

        # Score breakdown
        col1, col2, col3 = st.columns(3)
        col1.metric("TRL Score",        f"{d['trl_score']:.1f} / 10", "50% weight")
        col2.metric("Existence Quality", f"{d['existence_score']:.1f} / 10", "30% weight")
        col3.metric("Risk Profile",      f"{d['risk_score']:.1f} / 10", "20% weight")
        st.markdown("---")

        # ── TRL gauge ─────────────────────────────────────────
        st.markdown("#### 🎯 Technology Readiness Level")
        trl_colours = {
            1:"#ef4444",2:"#ef4444",3:"#f97316",
            4:"#f59e0b",5:"#eab308",6:"#84cc16",
            7:"#22c55e",8:"#10b981",9:"#06b6d4"
        }
        trl_col = trl_colours.get(trl_level,"#f59e0b")

        # TRL bar
        st.markdown(f"""
<div style="margin:12px 0;">
  <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
    <span style="color:{WHITE};font-weight:600;font-size:15px;">{trl.get("trl_label","")}</span>
    <span style="color:{trl_col};font-weight:700;font-size:18px;">TRL {trl_level}/9</span>
  </div>
  <div style="background:#1a2d45;border-radius:6px;height:14px;overflow:hidden;">
    <div style="background:{trl_col};height:100%;width:{trl_pct}%;border-radius:6px;transition:width 0.3s;"></div>
  </div>
  <div style="display:flex;justify-content:space-between;margin-top:4px;">
    <span style="color:#4a6fa5;font-size:10px;">TRL 1 — Theoretical</span>
    <span style="color:#4a6fa5;font-size:10px;">TRL 9 — Commercial</span>
  </div>
</div>
""", unsafe_allow_html=True)

        st.caption(trl.get("trl_rationale",""))
        st.caption(f"Entry rationale: {trl.get('entry_rationale','')}")
        if trl.get("analogous_schaeffler_technologies") or trl.get("analogous_schaeffler_tech"):
            st.caption(f"Schaeffler analogous experience: {trl.get('analogous_schaeffler_technologies') or trl.get('analogous_schaeffler_tech','')}")
        st.markdown("---")

        # ── TRL scale reference ───────────────────────────────
        with st.expander("  Schaeffler-adapted TRL scale reference"):
            trl_descriptions = [
                (1,"Basic principles observed","Theoretical concept only, no testing"),
                (2,"Technology concept formulated","Application identified, no experimental testing"),
                (3,"Experimental proof of concept","Lab demonstration, key functions validated"),
                (4,"Technology validated in lab","Component tested in controlled environment"),
                (5,"Validated in relevant environment","Prototype tested in industrial-like conditions"),
                (6,"Demonstrated in relevant environment","System prototype demonstrated"),
                (7,"Prototype in operational environment","Field trial or industrial pilot"),
                (8,"System complete and qualified","Full production design, limited production run"),
                (9,"Proven in operational environment","Commercial deployment at scale"),
            ]
            for lvl, label, desc in trl_descriptions:
                is_current = (lvl == trl_level)
                bg_col = "#1F3864" if is_current else "#1a2d45"
                text_col_inner = "#60a5fa" if is_current else "#94a3b8"
                st.markdown(f"""
<div style="background:{bg_col};border-radius:4px;padding:6px 12px;margin:3px 0;display:flex;gap:12px;align-items:center;">
  <div style="color:{trl_colours.get(lvl,'#f59e0b')};font-weight:700;min-width:40px;">TRL {lvl}</div>
  <div style="color:{WHITE if is_current else '#e2e8f0'};font-size:13px;min-width:220px;">{"▶ " if is_current else ""}{label}</div>
  <div style="color:{text_col_inner};font-size:12px;">{desc}</div>
</div>
""", unsafe_allow_html=True)

        st.markdown("---")

        # ── Evidence ──────────────────────────────────────────
        st.markdown("#### 📚 Existing Evidence")
        st.caption(existence.get("existence_summary",""))

        evidence_items = existence.get("evidence",[])
        if evidence_items:
            type_icons = {
                "Academic Paper":"📄","Startup":"🚀","Pilot":"🔧",
                "Government Programme":"🏛️","Industry Deployment":"🏭"
            }
            rel_cols = {"Direct":"#22c55e","Adjacent":"#60a5fa","Analogous":"#f59e0b"}
            conf_cols = {"High":"#22c55e","Medium":"#f59e0b","Low":"#ef4444"}
            # Sort: Direct > Adjacent > Analogous, then High > Medium > Low confidence
            rel_order  = {"Direct":0,"Adjacent":1,"Analogous":2}
            conf_order = {"High":0,"Medium":1,"Low":2}
            sorted_ev = sorted(evidence_items,
                key=lambda x: (rel_order.get(x.get("relevance","Adjacent"),1),
                               conf_order.get(x.get("confidence","Medium"),1)))
            top_ev = sorted_ev[:5]
            if len(evidence_items) > 5:
                st.caption(f"Showing top 5 of {len(evidence_items)} evidence items by relevance. Full list in the downloaded report.")
            for ev in top_ev:
                icon = type_icons.get(ev.get("type",""), "📌")
                rel_col  = rel_cols.get(ev.get("relevance","Adjacent"), "#60a5fa")
                conf_col = conf_cols.get(ev.get("confidence","Medium"), "#f59e0b")
                st.markdown(f"""
<div style="background:#1a2d45;border-radius:6px;padding:10px 14px;margin:5px 0;">
  <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px;">
    <span style="font-size:14px;">{icon}</span>
    <span style="color:{WHITE};font-weight:600;font-size:13px;">{ev.get("title","")}</span>
    <span style="background:{rel_col}22;color:{rel_col};font-size:10px;padding:2px 7px;border-radius:8px;">{ev.get("relevance","")}</span>
    <span style="background:{conf_col}22;color:{conf_col};font-size:10px;padding:2px 7px;border-radius:8px;">{ev.get("confidence","")} confidence</span>
  </div>
  <div style="color:#cbd5e1;font-size:12px;">{ev.get("description","")} <span style="color:#4a6fa5;">{ev.get("source","")}</span></div>
</div>
""", unsafe_allow_html=True)
        st.markdown("---")

        # ── Keyword map ───────────────────────────────────────
        st.markdown("#### 🗝️ Technology Keyword Map")
        st.caption("Key technical terms from this domain — indicative of academic and research weight.")
        keywords = existence.get("keywords", [])
        if keywords:
            # Size keywords by estimated relevance (first = most relevant)
            sizes = [28, 24, 22, 20, 18, 16, 15, 14, 13, 12]
            colours_kw = ["#60a5fa","#34d399","#a78bfa","#f59e0b","#f472b6",
                          "#60a5fa","#34d399","#a78bfa","#f59e0b","#f472b6"]
            badges = ""
            for i, kw in enumerate(keywords[:10]):
                sz  = sizes[i] if i < len(sizes) else 12
                col = colours_kw[i % len(colours_kw)]
                badges += f'<span style="background:{col}22;color:{col};font-size:{sz}px;padding:6px 14px;border-radius:20px;margin:4px;display:inline-block;font-weight:600;">{kw}</span>'
            st.markdown(f'<div style="background:#0f1e35;border-radius:8px;padding:16px;line-height:2.2;border:1px solid #2a4a70;">{badges}</div>',
                        unsafe_allow_html=True)
        st.markdown("---")

        # ── Technical risks ───────────────────────────────────
        st.markdown("#### ⚠️ Key Technical Risks")
        risks = trl.get("key_technical_risks", [])
        if risks:
            top_risks = sorted(risks, key=lambda x: {"High":0,"Medium":1,"Low":2}.get(x.get("severity","Medium"),1))[:5]
            if len(risks) > 5:
                st.caption(f"Showing top 5 of {len(risks)} risks by severity. Full list in the downloaded report.")
            for risk in top_risks:
                sev_col = {"High":"#ef4444","Medium":"#f59e0b","Low":"#22c55e"}.get(risk.get("severity","Medium"),"#f59e0b")
                st.markdown(f"""
<div style="background:#1a2d45;border-radius:6px;padding:10px 14px;margin:5px 0;border-left:3px solid {sev_col};">
  <div style="display:flex;align-items:center;gap:8px;margin-bottom:3px;">
    <span style="background:{sev_col}22;color:{sev_col};font-size:10px;padding:2px 8px;border-radius:8px;">{risk.get("severity","")} severity</span>
    <span style="color:{WHITE};font-size:13px;">{risk.get("risk","")}</span>
  </div>
  <div style="color:#94a3b8;font-size:12px;">↳ Mitigation: {risk.get("mitigation","")}</div>
</div>
""", unsafe_allow_html=True)

        if existence.get("technology_gaps"):
            gaps = existence["technology_gaps"]
            st.markdown("**Technology gaps to bridge:**")
            for gap in gaps[:5]:
                st.markdown(f"- {gap}")
            if len(gaps) > 5:
                st.caption(f"+{len(gaps)-5} more in the downloaded report.")

        st.caption(f"Estimated time to production readiness: **{existence.get('time_to_readiness','')}**")

        # ── Download report ───────────────────────────────────
        st.markdown("---")
        _one_click_dl(
            T("dl_feasib"),
            lambda: generate_feasibility_report(idea, quadrant, s1c, existence, trl, d),
            f"Schaeffler_Technical_Feasibility_{datetime.now().strftime('%Y%m%d')}.docx"
        )

        # ── Chat ──────────────────────────────────────────────
        st.markdown("---")
        st.subheader(T("s4_chat_header"))
        for msg in st.session_state.s4_chat:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        user_q = st.chat_input(T("s4_chat_ph"))
        if user_q:
            st.session_state.s4_chat.append({"role":"user","content":user_q})
            with st.chat_message("user"):
                st.markdown(user_q)
            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    ctx = f"""You are a Schaeffler R&D director discussing technical feasibility results.
Idea: {idea} | Quadrant: {quadrant}
TRL: {trl.get("trl_level","")} — {trl.get("trl_label","")}
Existence: {existence.get("existence_verdict","")}
Entry readiness: {trl.get("schaeffler_entry_readiness","")}
Time to readiness: {existence.get("time_to_readiness","")}
Key gaps: {existence.get("technology_gaps",[])}
Be specific, reference evidence where relevant, keep to 3-4 sentences."""
                    history = [{"role":m["role"],"content":m["content"]} for m in st.session_state.s4_chat]
                    reply = call_claude_chat(ctx, history)
                    st.markdown(reply)
                    st.session_state.s4_chat.append({"role":"assistant","content":reply})

        # ── Continue ──────────────────────────────────────────
        st.markdown("---")
        st.success(T("s4_success").format(score=final))
        if st.button(T("s4_continue"), type="primary", key="s4_continue"):
            st.session_state.active_stage = 5
            st.rerun()

        if st.button(T("s4_rerun"), key="s4_rerun"):
            st.session_state.s4_step = "intro"
            st.session_state.s4_data = {}
            st.session_state.s4_chat = []
            st.session_state.s4_report_buf = None
            st.rerun()



# ════════════════════════════════════════════════════════════
# STAGE 05 — ORGANISATIONAL READINESS
# ════════════════════════════════════════════════════════════
elif st.session_state.active_stage == 5:
    st.markdown(f"## {T('s5_title')}")
    st.markdown(f"""<div style="background:#1a2d45;border-radius:6px;padding:14px 18px;margin-bottom:16px;border-left:4px solid #2E75B6;">
<div style="color:#60a5fa;font-size:11px;letter-spacing:1px;font-weight:600;">{T('s1_what_label')}</div>
<div style="color:#e2e8f0;font-size:13px;margin-top:6px;">{T('s5_what')}</div>
<div style="color:#94a3b8;font-size:12px;margin-top:8px;">{T('s5_you_get')}</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    idea     = st.session_state.s1_idea
    s1c      = st.session_state.s1_classification
    quadrant = s1c.get("quadrant", "RADICAL")

    for k, v in {"s5_step":"intro","s5_data":{},"s5_chat":[]}.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # ── Intro ─────────────────────────────────────────────────
    if st.session_state.s5_step == "intro":
        st.info(f"**Idea:** {idea}")
        st.markdown(f"**Quadrant:** {quadrant}")
        # Pull context from previous stages for enriched prompt
        s3_landscape = st.session_state.get("s3_data",{}).get("landscape",{})
        s4_existence = st.session_state.get("s4_data",{}).get("existence",{})
        if st.button(T("s5_run_btn"), type="primary"):
            st.session_state.s5_step = "running"
            st.rerun()

    # ── Running ───────────────────────────────────────────────
    elif st.session_state.s5_step == "running":
        progress = st.progress(0)
        status   = st.empty()

        # Pull cross-stage context
        s3_landscape = st.session_state.get("s3_data",{}).get("landscape",{})
        s3_ansoff    = st.session_state.get("s3_data",{}).get("ansoff_data",{})
        s4_existence = st.session_state.get("s4_data",{}).get("existence",{})
        s4_trl       = st.session_state.get("s4_data",{}).get("trl",{})
        prior_filers = [f.get("company","") for f in s3_landscape.get("key_filers",[])]
        prior_evidence_sources = [e.get("source","") for e in s4_existence.get("evidence",[])]
        trl_level    = s4_trl.get("trl_level", 3)
        innovation_cluster = s1c.get("innovation_cluster","")
        product_family     = s1c.get("product_family","")
        trend_alignment    = s1c.get("trend_alignment",[])

        status.markdown("🏭 Assessing Schaeffler competency and asset fit...")
        progress.progress(25)

        system_readiness = f"""You are a senior Schaeffler Group innovation strategist assessing internal organisational readiness.

Schaeffler's P³ formula: Performance = Portfolio × People × Process
- Portfolio: Does this idea fit Schaeffler's strategic portfolio and innovation clusters?
- People: Does Schaeffler have the human skills, competencies, and teams to develop this?
- Process: Does Schaeffler have the processes, infrastructure, and assets to execute?

Schaeffler context:
- Innovation cluster for this idea: {innovation_cluster}
- Product family: {product_family}
- Strategic trend alignment: {', '.join(trend_alignment)}
- Current TRL of the idea's technology: {trl_level}
- Known patent filers in this space (from Stage 03): {', '.join(prior_filers[:6])}
- Evidence sources found (from Stage 04): {', '.join(prior_evidence_sources[:5])}

Schaeffler's known competency domains (use these to assess fit):
Precision motion systems, rolling and plain bearings, mechatronics, power electronics (via Vitesco merger), 
tribology and lubrication, automotive drivetrains (ICE and EV), industrial automation, embedded sensors,
ASPICE/ISO 26262 automotive software processes, OEM Tier 1 supply chain, 41 R&D centres globally.

Return ONLY valid JSON:
{{
  "p3_portfolio": {{
    "score": 0-10,
    "rationale": "2 sentences on strategic portfolio fit using Schaeffler's innovation clusters and trends",
    "cluster_fit": "one sentence on fit to the assigned innovation cluster",
    "strengths": ["strength 1", "strength 2"],
    "gaps": ["gap 1"]
  }},
  "p3_people": {{
    "score": 0-10,
    "rationale": "2 sentences on human capital and competency readiness",
    "matched_competencies": ["competency 1", "competency 2", "competency 3"],
    "competency_gap": "the single most critical missing competency",
    "sourcing_route": "Hire / Acquire / Partner / Upskill — one sentence on how to close the gap"
  }},
  "p3_process": {{
    "score": 0-10,
    "rationale": "2 sentences on process, infrastructure, and asset readiness",
    "applicable_assets": ["asset or process 1", "asset or process 2"],
    "investment_required": "one sentence on what needs to be built or acquired",
    "time_to_close": "estimated months or years"
  }},
  "partnership_candidates": [
    {{
      "name": "organisation name",
      "type": "Startup / University / Customer / Supplier / Research Institute",
      "rationale": "one sentence — why them and how they fill a specific gap",
      "route": "Co-develop / Acquire / License / JDA"
    }}
  ],
  "org_gaps": [
    {{
      "gap": "gap name",
      "severity": "High / Medium / Low",
      "closure_route": "one sentence on fastest route to close",
      "timeline": "estimated months"
    }}
  ],
  "build_or_partner": {{
    "recommendation": "Build internally / Co-develop / Acquire / License",
    "rationale": "2-3 sentences justifying the recommendation",
    "time_to_trl6_internal": "estimated timeline if built internally",
    "time_to_trl6_partner": "estimated timeline with external partnership"
  }},
  "org_readiness_score": 0-10
}}"""

        try:
            raw = call_claude(system_readiness,
                f"Innovation idea: {idea}\nQuadrant: {quadrant}\nTRL level: {trl_level}",
                max_tokens=2500)
            raw_clean = raw.strip().replace("```json","").replace("```","").strip()
            fb = raw_clean.find("{"); lb = raw_clean.rfind("}") + 1
            if fb >= 0: raw_clean = raw_clean[fb:lb]
            org_data = json.loads(raw_clean)
        except Exception as e:
            org_data = {
                "p3_portfolio":{"score":5,"rationale":"Assessment unavailable.","cluster_fit":"N/A","strengths":[],"gaps":[]},
                "p3_people":{"score":5,"rationale":"Assessment unavailable.","matched_competencies":[],"competency_gap":"N/A","sourcing_route":"N/A"},
                "p3_process":{"score":5,"rationale":"Assessment unavailable.","applicable_assets":[],"investment_required":"N/A","time_to_close":"N/A"},
                "partnership_candidates":[],
                "org_gaps":[],
                "build_or_partner":{"recommendation":"Co-develop","rationale":"N/A","time_to_trl6_internal":"N/A","time_to_trl6_partner":"N/A"},
                "org_readiness_score":5
            }

        progress.progress(75)
        status.markdown("🔍 Identifying partnership candidates...")

        # Weighted score: Portfolio 35%, People 40%, Process 25%
        p_portfolio = float(org_data.get("p3_portfolio",{}).get("score",5))
        p_people    = float(org_data.get("p3_people",{}).get("score",5))
        p_process   = float(org_data.get("p3_process",{}).get("score",5))
        final_org   = round(p_portfolio*0.35 + p_people*0.40 + p_process*0.25, 1)

        st.session_state.s5_data = {
            "org_data": org_data,
            "p_portfolio": p_portfolio,
            "p_people": p_people,
            "p_process": p_process,
            "final_score": final_org
        }
        progress.progress(100)
        status.markdown("✓ Complete.")
        time.sleep(0.5)
        st.session_state.s5_step = "done"
        st.rerun()

    # ── Results ───────────────────────────────────────────────
    elif st.session_state.s5_step == "done":
        d        = st.session_state.s5_data
        org_data = d["org_data"]
        final    = d["final_score"]

        # ── Score banner ──────────────────────────────────────
        score_col = "#22c55e" if final>=7 else "#f59e0b" if final>=4 else "#ef4444"
        bop_rec = org_data.get("build_or_partner",{}).get("recommendation","Co-develop")
        st.markdown(f"""
<div style="background:#0f1e35;border-radius:8px;padding:20px 24px;margin-bottom:20px;border:1px solid #2a4a70;">
  <div style="color:#94a3b8;font-size:9px;letter-spacing:2px;font-weight:700;font-family:Arial,sans-serif;margin-bottom:4px;">ORGANISATIONAL READINESS SCORE</div>
  <div style="color:{score_col};font-size:44px;font-weight:700;line-height:1;">{final}<span style="font-size:18px;color:#94a3b8;"> / 10</span></div>
  <div style="color:{WHITE};font-size:13px;margin-top:6px;opacity:0.8;">Build strategy: <b>{bop_rec}</b></div>
</div>
""", unsafe_allow_html=True)

        # ── P³ score cards ────────────────────────────────────
        st.markdown("#### P³ Assessment — Portfolio × People × Process")
        st.caption("Schaeffler's own innovation performance formula applied to this idea's organisational readiness")
        col1, col2, col3 = st.columns(3)
        col1.metric("Portfolio fit",  f"{d['p_portfolio']:.1f}/10", "35% weight")
        col2.metric("People (competency)", f"{d['p_people']:.1f}/10",  "40% weight")
        col3.metric("Process (assets)",   f"{d['p_process']:.1f}/10", "25% weight")
        st.markdown("---")

        # ── Portfolio dimension ───────────────────────────────
        port = org_data.get("p3_portfolio",{})
        st.markdown("#### Portfolio — Strategic fit")
        st.markdown(f"<div style='background:#1a2d45;border-radius:6px;padding:12px 16px;margin-bottom:10px;'><div style='color:#94a3b8;font-size:11px;'>RATIONALE</div><div style='color:#e2e8f0;font-size:13px;margin-top:4px;'>{port.get('rationale','')}</div><div style='color:#94a3b8;font-size:11px;margin-top:8px;'>CLUSTER FIT</div><div style='color:#60a5fa;font-size:13px;margin-top:4px;'>{port.get('cluster_fit','')}</div></div>", unsafe_allow_html=True)
        if port.get("strengths"):
            cols_ps = st.columns(len(port["strengths"][:3]))
            for i, s in enumerate(port["strengths"][:3]):
                cols_ps[i].markdown(f'<div style="background:#0a2010;border:1px solid #60a5fa33;border-radius:6px;padding:8px 12px;font-size:12px;color:#22c55e;">✓ {s}</div>', unsafe_allow_html=True)
        if port.get("gaps"):
            for g in port["gaps"][:2]:
                st.markdown(f'<div style="background:#1a0f0f;border:1px solid #ef444433;border-radius:6px;padding:8px 12px;margin-top:6px;font-size:12px;color:#ef4444;">✗ Gap: {g}</div>', unsafe_allow_html=True)
        st.markdown("---")

        # ── People dimension ──────────────────────────────────
        peop = org_data.get("p3_people",{})
        st.markdown("#### People — Competency & skills")
        st.markdown(f"<div style='background:#1a2d45;border-radius:6px;padding:12px 16px;margin-bottom:10px;'><div style='color:#94a3b8;font-size:11px;'>RATIONALE</div><div style='color:#e2e8f0;font-size:13px;margin-top:4px;'>{peop.get('rationale','')}</div></div>", unsafe_allow_html=True)
        if peop.get("matched_competencies"):
            st.markdown("**Matched Schaeffler competencies**")
            comp_cols = st.columns(min(3,len(peop["matched_competencies"])))
            for i, c in enumerate(peop["matched_competencies"][:3]):
                comp_cols[i].markdown(f'<div style="background:#0a2419;border:1px solid #60a5fa33;border-radius:4px;padding:8px;font-size:12px;color:#60a5fa;text-align:center;">{c}</div>', unsafe_allow_html=True)
        if peop.get("competency_gap"):
            st.markdown(f'<div style="background:#1a1020;border:1px solid #a78bfa44;border-radius:6px;padding:10px 14px;margin-top:8px;font-size:13px;color:#a78bfa;">⚠️ Critical gap: {peop["competency_gap"]}<br><span style="color:#94a3b8;font-size:12px;">Closure route: {peop.get("sourcing_route","")}</span></div>', unsafe_allow_html=True)
        st.markdown("---")

        # ── Process dimension ─────────────────────────────────
        proc = org_data.get("p3_process",{})
        st.markdown("#### Process — Infrastructure & assets")
        st.markdown(f"<div style='background:#1a2d45;border-radius:6px;padding:12px 16px;margin-bottom:10px;'><div style='color:#94a3b8;font-size:11px;'>RATIONALE</div><div style='color:#e2e8f0;font-size:13px;margin-top:4px;'>{proc.get('rationale','')}</div></div>", unsafe_allow_html=True)
        if proc.get("applicable_assets"):
            st.markdown("**Applicable assets / processes**")
            for a in proc["applicable_assets"][:3]:
                st.markdown(f"- {a}")
        if proc.get("investment_required"):
            st.markdown(f'<div style="background:#1a1a0f;border:1px solid #f59e0b44;border-radius:6px;padding:10px 14px;margin-top:8px;font-size:13px;color:#f59e0b;">🔧 Investment required: {proc["investment_required"]}<br><span style="color:#94a3b8;font-size:12px;">Estimated time to close: {proc.get("time_to_close","")}</span></div>', unsafe_allow_html=True)
        st.markdown("---")

        # ── Partnership candidates ────────────────────────────
        partners = org_data.get("partnership_candidates",[])
        if partners:
            st.markdown("#### Partnership candidates")
            route_cols = {"Co-develop":"#60a5fa","Acquire":"#f59e0b","License":"#a78bfa","JDA":"#22c55e"}
            for p in partners[:4]:
                rc = route_cols.get(p.get("route","Co-develop"),"#60a5fa")
                st.markdown(f"""
<div style="background:#1a2d45;border-radius:6px;padding:10px 14px;margin:5px 0;display:flex;align-items:flex-start;gap:12px;">
  <div style="flex:1;">
    <div style="color:{WHITE};font-weight:600;font-size:13px;">{p.get('name','')}</div>
    <div style="color:#94a3b8;font-size:12px;margin-top:3px;">{p.get('type','')} · {p.get('rationale','')}</div>
  </div>
  <div style="background:{rc}22;color:{rc};font-size:11px;padding:3px 10px;border-radius:10px;white-space:nowrap;">{p.get('route','')}</div>
</div>""", unsafe_allow_html=True)
            st.markdown("---")

        # ── Org gaps register ─────────────────────────────────
        gaps = org_data.get("org_gaps",[])
        if gaps:
            st.markdown("#### Organisational gaps")
            sev_col = {"High":"#ef4444","Medium":"#f59e0b","Low":"#22c55e"}
            for g in gaps[:4]:
                sc = sev_col.get(g.get("severity","Medium"),"#f59e0b")
                st.markdown(f"""
<div style="background:#1a2d45;border-radius:6px;padding:10px 14px;margin:5px 0;display:flex;align-items:flex-start;gap:12px;">
  <div style="flex:1;">
    <div style="color:{WHITE};font-weight:600;font-size:13px;">{g.get('gap','')}</div>
    <div style="color:#94a3b8;font-size:12px;margin-top:3px;">{g.get('closure_route','')} · Est. {g.get('timeline','')}</div>
  </div>
  <div style="background:{sc}22;color:{sc};font-size:11px;padding:3px 10px;border-radius:10px;white-space:nowrap;">{g.get('severity','')} severity</div>
</div>""", unsafe_allow_html=True)
            st.markdown("---")

        # ── Build or partner ──────────────────────────────────
        bop = org_data.get("build_or_partner",{})
        st.markdown("#### Build or partner?")
        bop_c1, bop_c2 = st.columns(2)
        bop_c1.markdown(f'<div style="background:#1a2d45;border-radius:8px;padding:14px 16px;"><div style="color:#94a3b8;font-size:11px;margin-bottom:4px;">RECOMMENDATION</div><div style="color:#60a5fa;font-size:16px;font-weight:600;">{bop.get("recommendation","")}</div><div style="color:#e2e8f0;font-size:12px;margin-top:6px;">{bop.get("rationale","")}</div></div>', unsafe_allow_html=True)
        bop_c2.markdown(f'<div style="background:#1a2d45;border-radius:8px;padding:14px 16px;"><div style="color:#94a3b8;font-size:11px;margin-bottom:4px;">TIME TO TRL 6</div><div style="color:#e2e8f0;font-size:13px;"><b>Internal:</b> {bop.get("time_to_trl6_internal","")}<br><b>With partner:</b> {bop.get("time_to_trl6_partner","")}</div></div>', unsafe_allow_html=True)

        # ── Chat ──────────────────────────────────────────────
        st.markdown("---")
        st.subheader(T("s5_chat_header"))
        for msg in st.session_state.s5_chat:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        user_q = st.chat_input(T("s5_chat_ph"))
        if user_q:
            st.session_state.s5_chat.append({"role":"user","content":user_q})
            with st.chat_message("user"):
                st.markdown(user_q)
            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    ctx = f"""You are a Schaeffler innovation readiness expert discussing P³ organisational assessment.
Idea: {idea} | Quadrant: {quadrant}
P³ scores — Portfolio: {d['p_portfolio']}/10 | People: {d['p_people']}/10 | Process: {d['p_process']}/10
Build recommendation: {bop.get('recommendation','')}
Critical gap: {org_data.get('p3_people',{}).get('competency_gap','')}
Be specific to Schaeffler's context (Vitesco integration, E-Mobility shift, OEM relationships). 3-4 sentences."""
                    history = [{"role":m["role"],"content":m["content"]} for m in st.session_state.s5_chat]
                    reply = call_claude_chat(ctx, history)
                    st.markdown(reply)
                    st.session_state.s5_chat.append({"role":"assistant","content":reply})

        # ── Continue ──────────────────────────────────────────
        st.markdown("---")
        _one_click_dl(
            T("dl_org"),
            lambda: generate_org_report(idea, quadrant, s1c, d),
            f"Schaeffler_Org_Readiness_{datetime.now().strftime('%Y%m%d')}.docx"
        )
        st.markdown("---")
        st.success(T("s5_success").format(score=final))
        if st.button(T("s5_continue"), type="primary", key="s5_continue"):
            st.session_state.active_stage = 6
            st.rerun()

        if st.button(T("s5_rerun"), key="s5_rerun"):
            st.session_state.s5_step = "intro"
            st.session_state.s5_data = {}
            st.session_state.s5_chat = []
            st.session_state.s5_report_buf = None
            st.rerun()


# ════════════════════════════════════════════════════════════
# STAGE 06 — SCORING & SYNTHESIS
# ════════════════════════════════════════════════════════════
elif st.session_state.active_stage == 6:
    st.markdown(f"## {T('s6_title')}")
    st.markdown(f"""<div style="background:#1a2d45;border-radius:6px;padding:14px 18px;margin-bottom:16px;border-left:4px solid #2E75B6;">
<div style="color:#60a5fa;font-size:11px;letter-spacing:1px;font-weight:600;">{T('s1_what_label')}</div>
<div style="color:#e2e8f0;font-size:13px;margin-top:6px;">{T('s6_what')}</div>
<div style="color:#94a3b8;font-size:12px;margin-top:8px;">{T('s6_you_get')}</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    idea     = st.session_state.s1_idea
    s1c      = st.session_state.s1_classification
    quadrant = s1c.get("quadrant","RADICAL")

    for k, v in {"s6_step":"intro","s6_data":{},"s6_chat":[]}.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # ── Check all stages are complete ─────────────────────────
    s2_done = bool(st.session_state.get("s2_data"))
    s3_done = bool(st.session_state.get("s3_data"))
    s4_done = bool(st.session_state.get("s4_data"))
    s5_done = bool(st.session_state.get("s5_data"))

    if not (s2_done and s3_done and s4_done and s5_done):
        st.warning("Complete Stages 02, 03, 04, and 05 first before running the final synthesis.")
        missing = []
        if not s2_done: missing.append("Stage 02: Market Intelligence")
        if not s3_done: missing.append("Stage 03: Patent Intelligence")
        if not s4_done: missing.append("Stage 04: Technical Feasibility")
        if not s5_done: missing.append("Stage 05: Organisational Readiness")
        for m in missing:
            st.markdown(f"- ⬜ {m}")
        if st.button("← Back", key="s6_back2"):
            st.session_state.active_stage = 5
            st.rerun()
        st.stop()

    # Pull scores from previous stages
    market_score      = st.session_state.s2_data.get("final_score", 5.0)
    patent_score      = st.session_state.s3_data.get("final_score", 5.0)
    feasibility_score = st.session_state.s4_data.get("final_score", 5.0)
    org_score         = st.session_state.s5_data.get("final_score", 5.0)

    # ── Intro: show weights + let user adjust ─────────────────
    if st.session_state.s6_step == "intro":
        st.info(f"**Idea:** {idea}")
        st.markdown(f"**Quadrant:** {quadrant}")
        st.markdown("---")

        st.markdown("#### Scores from previous stages")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Market Intelligence",    f"{market_score} / 10")
        c2.metric("Patent Intelligence",    f"{patent_score} / 10")
        c3.metric("Technical Feasibility",  f"{feasibility_score} / 10")
        c4.metric("Org Readiness",          f"{org_score} / 10")

        st.markdown("---")
        st.markdown("#### Innovation Potential Index — Scoring Weights")
        st.caption("Default weights set for first iteration. Adjust and refine with Johannes Enders.")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            w_market = st.slider("Market Intelligence", 0, 100, 35, 5, key="w_market")
        with col2:
            w_patent = st.slider("Patent Intelligence", 0, 100, 25, 5, key="w_patent")
        with col3:
            w_feasibility = st.slider("Technical Feasibility", 0, 100, 25, 5, key="w_feasibility")
        with col4:
            w_org = st.slider("Org Readiness", 0, 100, 15, 5, key="w_org")

        total_weight = w_market + w_patent + w_feasibility + w_org
        if total_weight != 100:
            st.warning(f"Weights must add up to 100. Current total: {total_weight}. Adjust the sliders.")
        else:
            st.success(f"✓ Weights sum to 100")
            if st.button("Run Final Synthesis →", type="primary"):
                st.session_state.s6_weights = {"market": w_market, "patent": w_patent, "feasibility": w_feasibility, "org": w_org}
                st.session_state.s6_step = "running"
                st.rerun()

    # ── Running ───────────────────────────────────────────────
    elif st.session_state.s6_step == "running":
        progress = st.progress(0)
        status   = st.empty()

        weights = st.session_state.get("s6_weights", {"market":35,"patent":25,"feasibility":25,"org":15})

        status.markdown("📐 Calculating Innovation Potential Index...")
        progress.progress(20)

        wm = weights["market"] / 100
        wp = weights["patent"] / 100
        wf = weights["feasibility"] / 100
        wo = weights.get("org", 15) / 100
        ipi = round(market_score * wm + patent_score * wp + feasibility_score * wf + org_score * wo, 1)

        status.markdown("🧠 Writing narrative synthesis...")
        progress.progress(45)

        s2d = st.session_state.s2_data
        s3d = st.session_state.s3_data
        s4d = st.session_state.s4_data
        s5d = st.session_state.s5_data
        org_d = s5d.get("org_data", {})

        system_synthesis = """You are a senior Schaeffler innovation strategist writing a final investment recommendation.
Synthesise all four stages of analysis into a clear, actionable assessment.

Be direct, specific, and honest — including about weaknesses. 
Reference Schaeffler's strategic context (electrification transition, Vitesco merger, E-Mobility growth).

Return ONLY valid JSON:
{
  "headline": "one punchy sentence summarising the overall verdict",
  "recommendation": "PROCEED / PROCEED WITH CONDITIONS / DEFER / REJECT",
  "recommendation_rationale": "2-3 sentences on why this recommendation",
  "strongest_signals": ["top positive signal 1", "top positive signal 2", "top positive signal 3"],
  "key_concerns": ["main concern 1", "main concern 2", "main concern 3"],
  "conditions": ["condition 1 if applicable", "condition 2"],
  "strategic_fit": "2-3 sentences on how this fits Schaeffler's current strategic priorities",
  "next_steps": ["concrete next step 1", "concrete next step 2", "concrete next step 3"],
  "narrative": "5-6 paragraph narrative synthesis covering the full picture — market opportunity, IP position, technical readiness, and strategic recommendation"
}"""

        synthesis_context = f"""
Idea: {idea}
Quadrant: {quadrant}
Innovation Potential Index: {ipi}/10

Innovation Cluster: {st.session_state.s1_classification.get('innovation_cluster','')}
Product Family: {st.session_state.s1_classification.get('product_family','')}
Pipeline Route: {st.session_state.s1_classification.get('pipeline_route','')}
Trend Alignment: {', '.join(st.session_state.s1_classification.get('trend_alignment',[]))}

Stage 02 — Market Intelligence ({weights['market']}% weight): {market_score}/10
- Market: {s2d.get('market',{}).get('market_name','')}
- Size: {_mval(s2d.get('market',{}).get('market_size_current') or s2d.get('market',{}).get('market_size_2024',''))}
- CAGR: {s2d.get('market',{}).get('cagr','')}
- Maturity: {s2d.get('market',{}).get('market_maturity','')}
- Primary sectors: {', '.join(s2d.get('sectors',{}).get('primary_sectors',[]))}
- Competition: {s2d.get('comp',{}).get('competitive_intensity','')}
- White space: {s2d.get('comp',{}).get('white_space','')}

Stage 03 — Patent Intelligence ({weights['patent']}% weight): {patent_score}/10
- Activity: {s3d.get('landscape',{}).get('activity_level','')}
- Trend: {s3d.get('landscape',{}).get('filing_trend','')}
- Novelty signal: {s3d.get('ansoff_data',{}).get('novelty_signal','')}
- IP risk: {s3d.get('ansoff_data',{}).get('ip_risk','')}
- Schaeffler IP gap: {s3d.get('ansoff_data',{}).get('schaeffler_position',{}).get('gap','')}

Stage 04 — Technical Feasibility ({weights['feasibility']}% weight): {feasibility_score}/10
- TRL: {s4d.get('trl',{}).get('trl_level','')} — {s4d.get('trl',{}).get('trl_label','')}
- Existence: {s4d.get('existence',{}).get('existence_verdict','')}
- Entry readiness: {s4d.get('trl',{}).get('schaeffler_entry_readiness','')}
- Time to readiness: {s4d.get('existence',{}).get('time_to_readiness','')}

Stage 05 — Organisational Readiness ({weights.get('org',15)}% weight): {org_score}/10
- P³ Portfolio: {s5d.get('p_portfolio',5)}/10
- P³ People: {s5d.get('p_people',5)}/10
- P³ Process: {s5d.get('p_process',5)}/10
- Critical competency gap: {org_d.get('p3_people',{}).get('competency_gap','')}
- Build strategy: {org_d.get('build_or_partner',{}).get('recommendation','')}
- Time to TRL6 with partner: {org_d.get('build_or_partner',{}).get('time_to_trl6_partner','')}
"""

        # ── Call 1: structured fields ─────────────────────────
        system_structured = """You are a senior Schaeffler innovation strategist.
Return ONLY valid JSON with exactly these fields — no markdown, no extra text, no trailing commas:
{
  "headline": "one direct sentence summarising the overall verdict on this idea",
  "recommendation": "PROCEED or PROCEED WITH CONDITIONS or DEFER or REJECT",
  "recommendation_rationale": "2-3 sentences explaining why this recommendation",
  "strongest_signals": ["specific positive signal from the data 1", "specific positive signal 2", "specific positive signal 3"],
  "key_concerns": ["specific concern from the data 1", "specific concern 2", "specific concern 3"],
  "conditions": ["specific condition to meet before proceeding 1", "condition 2"],
  "strategic_fit": "2-3 sentences on how this fits Schaeffler strategy — reference electrification transition, Vitesco merger, E-Mobility growth, and the specific product family and innovation cluster",
  "risks": ["specific risk 1 with mitigation approach", "specific risk 2 with mitigation approach", "specific risk 3 with mitigation approach"],
  "next_steps": ["concrete Schaeffler-specific action step 1", "concrete action step 2", "concrete action step 3", "concrete action step 4"]
}"""
        raw1 = call_claude(system_structured, synthesis_context, max_tokens=1000)
        raw1_clean = raw1.strip().replace("```json","").replace("```","").strip()
        fb = raw1_clean.find("{"); lb = raw1_clean.rfind("}") + 1
        if fb >= 0: raw1_clean = raw1_clean[fb:lb]
        try:
            synthesis_structured = json.loads(raw1_clean)
        except:
            synthesis_structured = {
                "headline": f"IPI score of {ipi}/10 — {'strong' if ipi>=7 else 'moderate' if ipi>=4 else 'weak'} opportunity in {s2d.get('market',{}).get('market_name','this market')}.",
                "recommendation": "PROCEED WITH CONDITIONS" if ipi >= 5 else "DEFER",
                "recommendation_rationale": f"The idea scores {ipi}/10 across market, patent, and feasibility dimensions. Market opportunity is {'strong' if market_score>=7 else 'moderate'}, IP position is {'favourable' if patent_score>=7 else 'mixed'}, and technical readiness is {'high' if feasibility_score>=7 else 'still developing'}.",
                "strongest_signals": [
                    f"Market Intelligence score of {market_score}/10 — {s2d.get('market',{}).get('market_name','')} shows growth",
                    f"Primary sector fit: {', '.join(s2d.get('sectors',{}).get('primary_sectors',[])[:2])}",
                    f"Patent novelty signal: {s3d.get('ansoff_data',{}).get('novelty_signal','Moderate')}"
                ],
                "key_concerns": [
                    f"Technical feasibility at TRL {s4d.get('trl',{}).get('trl_level',3)} — {s4d.get('trl',{}).get('schaeffler_entry_readiness','')}",
                    f"IP risk level: {s3d.get('ansoff_data',{}).get('ip_risk','Medium')}",
                    f"Competitive intensity: {s2d.get('comp',{}).get('competitive_intensity','')}"
                ],
                "conditions": [
                    "Confirm IP freedom-to-operate before committing R&D budget",
                    "Validate technical feasibility with internal engineering team"
                ],
                "strategic_fit": f"This idea aligns with Schaeffler's {quadrant} innovation quadrant and targets {', '.join(s2d.get('sectors',{}).get('primary_sectors',[])[:2])} — sectors central to Schaeffler's post-Vitesco portfolio. The electrification transition creates urgency for exactly this type of innovation investment.",
                "risks": [
                    f"IP risk: {s3d.get('ansoff_data',{}).get('ip_risk','Medium')} — conduct freedom-to-operate analysis before R&D commitment",
                    f"Technical maturity: TRL {s4d.get('trl',{}).get('trl_level',3)} — further development required before production readiness",
                    "Market timing risk — validate demand with target customers before scaling investment"
                ],
                "next_steps": [
                    "Commission internal engineering feasibility review within 30 days",
                    "Conduct freedom-to-operate IP analysis with Schaeffler patent team",
                    f"Identify pilot customer in {s2d.get('sectors',{}).get('primary_sectors',['target sector'])[0]} for co-development conversation",
                    "Present to Innovation steering committee with this assessment as supporting material"
                ]
            }

        # ── Call 2: narrative separately as plain text ────────
        system_narrative = f"""You are a senior Schaeffler innovation strategist. Write a 5-paragraph narrative synthesis for this innovation assessment. Write in flowing prose — no bullet points, no headers. Be specific about the market opportunity, IP landscape, technical maturity, and strategic recommendation. Reference Schaeffler's context (electrification, Vitesco merger, E-Mobility growth, OEM relationships)."""
        raw2 = call_claude(system_narrative, synthesis_context + f"\n\nRecommendation: {synthesis_structured.get('recommendation','')}\nIPI: {ipi}/10", max_tokens=800)
        narrative_text = raw2.strip().replace("```","").strip()

        synthesis = {**synthesis_structured, "narrative": narrative_text}

        progress.progress(80)
        status.markdown("✓ Complete.")
        time.sleep(0.5)

        st.session_state.s6_data = {
            "ipi": ipi,
            "weights": weights,
            "synthesis": synthesis,
            "scores": {
                "market": market_score,
                "patent": patent_score,
                "feasibility": feasibility_score,
                "org": org_score
            }
        }
        st.session_state.s6_step = "done"
        st.rerun()

    # ── Results ───────────────────────────────────────────────
    elif st.session_state.s6_step == "done":
        d         = st.session_state.s6_data
        ipi       = d.get("ipi", 0)
        weights   = d.get("weights", {"market":35,"patent":25,"feasibility":25,"org":15})
        synthesis = d.get("synthesis", {})
        scores    = d.get("scores", {"market":5,"patent":5,"feasibility":5,"org":5})

        rec = synthesis.get("recommendation","PROCEED WITH CONDITIONS")
        rec_colours = {
            "PROCEED":               "#22c55e",
            "PROCEED WITH CONDITIONS":"#f59e0b",
            "DEFER":                 "#f97316",
            "REJECT":                "#ef4444"
        }
        rec_col = rec_colours.get(rec, "#f59e0b")
        ipi_col = "#22c55e" if ipi>=7 else "#f59e0b" if ipi>=4 else "#ef4444"

        # ── Auto-save to Google Sheets (once per session) ─────
        if not st.session_state.get("_s6_saved_to_sheets"):
            s2d_sv = st.session_state.get("s2_data", {})
            s3d_sv = st.session_state.get("s3_data", {})
            s4d_sv = st.session_state.get("s4_data", {})
            s5d_sv = st.session_state.get("s5_data", {})
            qa_pairs = []
            for q, a in zip(st.session_state.get("s1_questions",[]), st.session_state.get("s1_answers",[])):
                qa_pairs.append(f"Q: {q} / A: {a}")
            row = {
                "Date":                  datetime.now().strftime("%Y-%m-%d %H:%M"),
                "Submitter Name":        st.session_state.get("user_name",""),
                "Position":              st.session_state.get("user_position",""),
                "Department":            st.session_state.get("user_dept",""),
                "Full Idea Description": st.session_state.get("s1_idea",""),
                "Clarifying Q&A":        " | ".join(qa_pairs),
                "Quadrant":              quadrant,
                "Innovation Cluster":    s1c.get("innovation_cluster",""),
                "Product Family":        s1c.get("product_family",""),
                "Market Score":          str(scores.get("market","")),
                "Patent Score":          str(scores.get("patent","")),
                "Feasibility Score":     str(scores.get("feasibility","")),
                "Org Readiness Score":   str(scores.get("org","")),
                "IPI Score":             str(ipi),
                "Recommendation":        rec,
                "Key Concerns":          " | ".join(synthesis.get("key_concerns",[])[:3]),
                "Next Steps":            " | ".join(synthesis.get("next_steps",[])[:4]),
                "Market Name":           s2d_sv.get("market",{}).get("market_name",""),
                "Market Size 2024":      _mval(s2d_sv.get("market",{}).get("market_size_current") or s2d_sv.get("market",{}).get("market_size_2024","")),
                "CAGR":                  s2d_sv.get("market",{}).get("cagr",""),
                "TRL Level":             str(s4d_sv.get("trl",{}).get("trl_level","")),
                "Build Strategy":        s5d_sv.get("org_data",{}).get("build_or_partner",{}).get("recommendation",""),
            }
            saved = save_idea_to_sheets(row)
            st.session_state["_s6_saved_to_sheets"] = True
            if saved:
                st.success("✅ Idea automatically saved to the Innovation Ideas Log.")

        # ── IPI banner ────────────────────────────────────────
        st.markdown(f"""
<div style="background:#0f1e35;border-radius:8px;padding:20px 24px;margin-bottom:20px;border:1px solid #2a4a70;">
  <div style="display:flex;justify-content:space-between;align-items:flex-start;">
    <div>
      <div style="color:#94a3b8;font-size:9px;letter-spacing:2px;font-weight:700;font-family:Arial,sans-serif;margin-bottom:4px;">INNOVATION POTENTIAL INDEX</div>
      <div style="color:{ipi_col};font-size:48px;font-weight:700;line-height:1;">{ipi}<span style="font-size:20px;color:#94a3b8;"> / 10</span></div>
      <div style="color:{WHITE};font-size:13px;margin-top:6px;opacity:0.8;">{synthesis.get('headline','')}</div>
    </div>
    <div style="text-align:right;">
      <div style="color:{WHITE};font-size:11px;opacity:0.5;margin-bottom:4px;">RECOMMENDATION</div>
      <div style="background:{rec_col}22;color:{rec_col};font-size:16px;font-weight:700;padding:8px 16px;border-radius:6px;border:1px solid {rec_col}44;">{rec}</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

        # ── Radar chart ───────────────────────────────────────
        st.markdown("#### 📡 Innovation Potential Radar")

        categories   = ["Market Intelligence", "Patent Intelligence", "Technical Feasibility", "Org Readiness"]
        values       = [scores["market"], scores["patent"], scores["feasibility"], scores.get("org", 5)]
        values_close = values + [values[0]]
        cats_close   = categories + [categories[0]]

        fig_radar = go.Figure()
        fig_radar.add_trace(go.Scatterpolar(
            r=values_close, theta=cats_close,
            fill="toself",
            fillcolor=f"rgba(96,165,250,0.15)",
            line=dict(color=BLUE, width=2),
            marker=dict(size=8, color=BLUE),
            name="Score"
        ))
        # Add benchmark line at 7
        fig_radar.add_trace(go.Scatterpolar(
            r=[7,7,7,7,7], theta=cats_close,
            line=dict(color="#4a6fa5", width=1, dash="dot"),
            marker=dict(size=0),
            fill=None,
            name="Target (7/10)",
            showlegend=True
        ))
        fig_radar.update_layout(
            polar=dict(
                bgcolor=BG,
                radialaxis=dict(
                    visible=True, range=[0,10],
                    tickfont=dict(color=WHITE, size=10),
                    gridcolor="#2a4a70", linecolor="#2a4a70",
                    tickvals=[2,4,6,8,10]
                ),
                angularaxis=dict(
                    tickfont=dict(color=WHITE, size=12),
                    gridcolor="#2a4a70", linecolor="#2a4a70"
                )
            ),
            paper_bgcolor=BG,
            plot_bgcolor=BG,
            legend=dict(font=dict(color=WHITE), bgcolor=BG),
            height=400,
            margin=dict(l=60,r=60,t=40,b=40)
        )
        st.plotly_chart(fig_radar, use_container_width=True)

        # Score + weight breakdown
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Market Intelligence",   f"{scores['market']} / 10",   f"{weights['market']}% weight")
        col2.metric("Patent Intelligence",   f"{scores['patent']} / 10",   f"{weights['patent']}% weight")
        col3.metric("Technical Feasibility", f"{scores['feasibility']} / 10", f"{weights['feasibility']}% weight")
        col4.metric("Org Readiness",         f"{scores.get('org',5)} / 10", f"{weights.get('org',15)}% weight")
        st.markdown("---")

        # ── Recommendation ────────────────────────────────────
        st.markdown("#### 🎯 Recommendation")
        st.markdown(f"""
<div style="background:{rec_col}11;border:1px solid {rec_col}44;border-radius:8px;padding:16px 20px;margin-bottom:16px;">
  <div style="color:{rec_col};font-size:16px;font-weight:700;margin-bottom:6px;">{rec}</div>
  <div style="color:{WHITE};font-size:13px;">{synthesis.get('recommendation_rationale','')}</div>
</div>
""", unsafe_allow_html=True)

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Strongest signals**")
            for s in synthesis.get("strongest_signals",[])[:3]:
                st.markdown(f"✅ {s}")
        with col_b:
            st.markdown("**Key concerns**")
            for c in synthesis.get("key_concerns",[])[:3]:
                st.markdown(f"⚠️ {c}")

        if synthesis.get("conditions"):
            st.markdown("**Conditions / requirements:**")
            for cond in synthesis.get("conditions",[])[:3]:
                st.markdown(f"→ {cond}")
        st.markdown("---")

        # ── Strategic fit ─────────────────────────────────────
        st.markdown("#### 🏭 Strategic Fit for Schaeffler")
        st.markdown(synthesis.get("strategic_fit",""))
        st.markdown("---")

        # ── Next steps ────────────────────────────────────────
        st.markdown("#### 📋 Recommended Next Steps")
        for i, step in enumerate(synthesis.get("next_steps",[])[:5], 1):
            st.markdown(f"**{i}.** {step}")
        st.markdown("---")

        # ── Full narrative (expandable) ───────────────────────
        with st.expander("  📖 Read full narrative synthesis"):
            st.markdown(synthesis.get("narrative",""))

        # ── Visual mockup generation ──────────────────────────
        st.markdown("---")
        st.markdown("#### 🎨 Solution Visualisation")
        st.caption("Generate an AI image showing how this solution could look in a real-world context.")

        if "s5_mockup_image" not in st.session_state:
            st.session_state.s6_mockup_image = None

        if st.session_state.s6_mockup_image is None:
            if st.button(T("s6_image_btn"), type="secondary", key="s6_image"):
                with st.spinner("Generating image — this takes 10–20 seconds..."):
                    try:
                        # Step 1: Claude writes a precise image prompt
                        img_prompt_raw = call_claude(
                            "You write precise image generation prompts. Return ONLY the prompt, no quotes, no explanation, no preamble. Maximum 50 words.",
                            f"Write a photorealistic product visualisation prompt for: {idea}. Show the technology deployed in an industrial/automotive setting. Professional engineering photography style, high detail."
                        )
                        img_prompt = img_prompt_raw.strip().replace('"','').replace("'","")[:200]

                        # Step 2: Generate image via Pollinations.ai (free, no API key)
                        import urllib.parse
                        encoded_prompt = urllib.parse.quote(img_prompt)
                        image_url = f"https://image.pollinations.ai/prompt/{encoded_prompt}?width=960&height=640&nologo=true&seed=42&model=flux"

                        img_response = None
                        for attempt in range(3):
                            try:
                                img_response = requests.get(image_url, timeout=90)
                                if img_response.status_code == 200 and len(img_response.content) > 5000:
                                    break
                            except requests.exceptions.Timeout:
                                if attempt < 2:
                                    time.sleep(5)
                                    continue
                                else:
                                    raise

                        if img_response and img_response.status_code == 200 and len(img_response.content) > 5000:
                            st.session_state.s6_mockup_image = img_response.content
                            st.session_state.s6_mockup_prompt_used = img_prompt
                            st.rerun()
                        else:
                            st.error("Image generation timed out. Try again — Pollinations.ai can be slow on first request.")
                    except Exception as e:
                        st.error(f"Image generation error: {e}")
        else:
            st.image(st.session_state.s6_mockup_image, use_container_width=True)
            st.caption(f"Prompt used: {st.session_state.get('s5_mockup_prompt_used','')}")
            if st.button(T("s6_image_redo"), key="s6_image_redo"):
                st.session_state.s6_mockup_image = None
                st.session_state.s6_mockup_prompt_used = None
                st.rerun()

        # ── Master report download ────────────────────────────
        st.markdown("---")
        _s2d = st.session_state.s2_data
        _s3d = st.session_state.s3_data
        _s4d = st.session_state.s4_data
        _s5d = st.session_state.s5_data
        _one_click_dl(
            T("dl_master"),
            lambda: generate_master_report(idea, quadrant, s1c, _s2d, _s3d, _s4d, _s5d, d),
            f"Schaeffler_Innovation_Assessment_{datetime.now().strftime('%Y%m%d')}.docx"
        )
        st.caption(T("dl_caption"))

        # ── Chat ──────────────────────────────────────────────
        st.markdown("---")
        st.subheader(T("s6_chat_header"))
        for msg in st.session_state.s6_chat:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        user_q = st.chat_input(T("s6_chat_ph"))
        if user_q:
            st.session_state.s6_chat.append({"role":"user","content":user_q})
            with st.chat_message("user"):
                st.markdown(user_q)
            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    ctx = f"""You are a senior Schaeffler innovation strategist discussing the final assessment.
Idea: {idea} | Quadrant: {quadrant}
IPI Score: {ipi}/10 | Recommendation: {rec}
Market: {scores['market']}/10 | Patent: {scores['patent']}/10 | Feasibility: {scores['feasibility']}/10 | Org: {scores.get('org',5)}/10
Headline: {synthesis.get('headline','')}
Strongest signals: {synthesis.get('strongest_signals',[])}
Key concerns: {synthesis.get('key_concerns',[])}
Strategic fit: {synthesis.get('strategic_fit','')}
Be direct and specific. Reference Schaeffler's context where relevant. 3-4 sentences."""
                    history = [{"role":m["role"],"content":m["content"]} for m in st.session_state.s6_chat]
                    reply = call_claude_chat(ctx, history)
                    st.markdown(reply)
                    st.session_state.s6_chat.append({"role":"assistant","content":reply})

        # ── Re-run with different weights ─────────────────────
        st.markdown("---")
        if st.button(T("s6_rerun"), key="s6_rerun"):
            st.session_state.s6_step = "intro"
            st.session_state.s6_data = {}
            st.session_state.s6_chat = []
            st.session_state.s6_report_buf = None
            st.rerun()
