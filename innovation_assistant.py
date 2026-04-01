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

# Load API key — Streamlit secrets take priority, fallback to hardcoded for local dev
TAVILY_KEY = ""  # optional

st.set_page_config(page_title="Schaeffler Innovation Assistant", page_icon="⚙️", layout="centered")

# ── API key — secrets for deployment, hardcoded fallback for local ────────────
try:
    ANTHROPIC_KEY = st.secrets["ANTHROPIC_API_KEY"]
except Exception:
    ANTHROPIC_KEY = "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"

# ── Scroll to top on every page load ─────────────────────────
# ── Styling ───────────────────────────────────────────────────
st.markdown("""
<style>
.source-tag {
    background:#1e3a5f; color:#93c5fd;
    font-size:11px; padding:2px 8px;
    border-radius:3px; margin-left:6px;
    font-family: monospace;
}
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────
BG    = "#0f1e35"
BLUE  = "#60a5fa"
DIM   = "#4a6fa5"
WHITE = "#e2e8f0"
NAVY  = "#1F3864"

# ── Helpers ───────────────────────────────────────────────────
def call_claude(system, user, max_tokens=2000):
    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)
    for attempt in range(3):
        try:
            msg = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=max_tokens,
                system=system,
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
    for attempt in range(3):
        try:
            msg = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=max_tokens,
                system=system,
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


# ── Session state ─────────────────────────────────────────────
defaults = {
    "active_stage": 1,
    # Stage 01
    "s1_step": 1,
    "s1_idea": "",
    "s1_questions": [],
    "s1_answers": [],
    "s1_classification": {},
    "s1_chat": [],
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
    st.markdown("### ⚙️ Schaeffler Innovation Assistant")
    st.markdown("---")
    st.markdown("**Stage progress**")
    stages = [
        (1, "01 · Quadrant Classifier"),
        (2, "02 · Market Intelligence"),
        (3, "03 · Patent Intelligence"),
        (4, "04 · Technical Feasibility"),
        (5, "05 · Scoring & Synthesis"),
    ]
    # Determine which stages have been completed
    completed = set()
    if st.session_state.get("s1_classification"): completed.add(1)
    if st.session_state.get("s2_data"):           completed.add(2)
    if st.session_state.get("s3_data"):           completed.add(3)
    if st.session_state.get("s4_data"):           completed.add(4)
    if st.session_state.get("s5_data"):           completed.add(5)

    for num, label in stages:
        active = st.session_state.active_stage
        if num == active:
            st.markdown(f"🔵 **{label}** ← here")
        elif num in completed:
            if st.button(f"✅ {label}", key=f"nav_{num}", use_container_width=True):
                st.session_state.active_stage = num
                st.rerun()
        else:
            st.markdown(f"⬜ {label}")

    if st.session_state.s1_idea:
        st.markdown("---")
        st.caption(f"**Idea:** {st.session_state.s1_idea[:60]}...")
        if st.session_state.s1_classification:
            st.caption(f"**Quadrant:** {st.session_state.s1_classification.get('quadrant','')}")

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
    for y_pos, y_label in [(1.25,"Established"),(3.75,"Adjacent"),(6.25,"New to Schaeffler"),(8.75,"New to the World")]:
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
        margin=dict(l=90,r=30,t=50,b=70), font=dict(color=text_col)
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



def generate_master_report(idea, quadrant, s1c, s2d, s3d, s4d, s5d):
    """Generate the full master Innovation Assessment Word report covering all 4 stages."""
    ipi       = s5d["ipi"]
    weights   = s5d["weights"]
    synthesis = s5d["synthesis"]
    scores    = s5d["scores"]
    market    = s2d.get("market",{})
    comp      = s2d.get("comp",{})
    sectors   = s2d.get("sectors",{})
    landscape = s3d.get("landscape",{})
    ansoff_d  = s3d.get("ansoff_data",{})
    existence = s4d.get("existence",{})
    trl       = s4d.get("trl",{})

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

    # ── Header ────────────────────────────────────────────────
    t=doc.add_table(rows=1,cols=1); t.style="Table Grid"; c=t.cell(0,0); set_bg(c,"1F3864")
    p=c.paragraphs[0]; p.paragraph_format.space_before=Pt(12); p.paragraph_format.space_after=Pt(2)
    r=p.add_run("INNOVATION ASSESSMENT REPORT"); r.bold=True; r.font.size=Pt(11); r.font.color.rgb=WHITE
    p2=c.add_paragraph(); p2.paragraph_format.space_before=Pt(0); p2.paragraph_format.space_after=Pt(12)
    r2=p2.add_run("Schaeffler AI Innovation Research Assistant  ·  Full Pipeline Assessment  ·  Stages 01–05")
    r2.font.size=Pt(9); r2.font.color.rgb=RGBColor(0x93,0xC5,0xFD)
    doc.add_paragraph()

    # ── Title ─────────────────────────────────────────────────
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2)
    r=p.add_run(market.get("market_name", idea[:80])); r.bold=True; r.font.size=Pt(20); r.font.color.rgb=NAVY
    p2=doc.add_paragraph()
    rec_text = synthesis.get("recommendation","")
    r2=p2.add_run(f"IPI Score: {ipi}/10  ·  {rec_text}  ·  Quadrant: {quadrant}  ·  {datetime.now().strftime('%d %B %Y')}")
    r2.font.size=Pt(10); r2.italic=True; r2.font.color.rgb=GREY

    # ── Idea box ──────────────────────────────────────────────
    doc.add_paragraph()
    tb=doc.add_table(rows=1,cols=2); tb.style="Table Grid"
    c1=tb.cell(0,0); c2=tb.cell(0,1); set_bg(c1,"1F3864"); set_bg(c2,"EAF1FB"); c1.width=Inches(0.12)
    c1.paragraphs[0].add_run("")
    rp=c2.paragraphs[0]; rp.paragraph_format.space_before=Pt(8); rp.paragraph_format.space_after=Pt(2)
    rb=rp.add_run("Innovation Idea"); rb.bold=True; rb.font.size=Pt(9); rb.font.color.rgb=NAVY
    rp2=c2.add_paragraph(); rp2.paragraph_format.space_before=Pt(0); rp2.paragraph_format.space_after=Pt(8)
    ri=rp2.add_run(idea); ri.font.size=Pt(10); ri.italic=True
    doc.add_paragraph()

    # ── IPI Score table ───────────────────────────────────────
    h1(doc,"Innovation Potential Index")
    ipi_tbl=doc.add_table(rows=1,cols=4); ipi_tbl.style="Table Grid"
    for i,h in enumerate(["Stage","Score","Weight","Weighted"]):
        c=ipi_tbl.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
        r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
    stage_rows=[
        ("02 · Market Intelligence",   f"{scores['market']:.1f}/10",     f"{weights['market']}%",     f"{scores['market']*weights['market']/100:.1f}"),
        ("03 · Patent Intelligence",   f"{scores['patent']:.1f}/10",     f"{weights['patent']}%",     f"{scores['patent']*weights['patent']/100:.1f}"),
        ("04 · Technical Feasibility", f"{scores['feasibility']:.1f}/10",f"{weights['feasibility']}%",f"{scores['feasibility']*weights['feasibility']/100:.1f}"),
    ]
    for i,(stage,score,wt,weighted) in enumerate(stage_rows):
        row=ipi_tbl.add_row(); fill="EAF1FB" if i%2==0 else "FFFFFF"
        for c in row.cells: set_bg(c,fill)
        for j,val in enumerate([stage,score,wt,weighted]):
            r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(10); r.bold=(j==0)
    fr=ipi_tbl.add_row()
    for c in fr.cells: set_bg(c,"1F3864")
    r=fr.cells[0].paragraphs[0].add_run("INNOVATION POTENTIAL INDEX"); r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
    r2=fr.cells[3].paragraphs[0].add_run(f"{ipi}/10"); r2.bold=True; r2.font.size=Pt(12); r2.font.color.rgb=LBLUE
    doc.add_paragraph()

    # ── Recommendation ────────────────────────────────────────
    h1(doc,"Recommendation & Synthesis")
    kv(doc,"Recommendation",synthesis.get("recommendation",""))
    kv(doc,"Rationale",synthesis.get("recommendation_rationale",""))
    doc.add_paragraph()
    if synthesis.get("strongest_signals"):
        h2(doc,"Strongest Signals")
        for s in synthesis["strongest_signals"]: bul(doc,f"✓ {s}")
    if synthesis.get("key_concerns"):
        h2(doc,"Key Concerns")
        for c in synthesis["key_concerns"]: bul(doc,f"⚠ {c}")
    if synthesis.get("conditions"):
        h2(doc,"Conditions")
        for c in synthesis["conditions"]: bul(doc,f"→ {c}")
    doc.add_paragraph()
    kv(doc,"Strategic fit",synthesis.get("strategic_fit",""))
    doc.add_paragraph()
    body(doc,synthesis.get("narrative",""))

    # ── Next steps ────────────────────────────────────────────
    if synthesis.get("next_steps"):
        h2(doc,"Recommended Next Steps")
        for i,step in enumerate(synthesis["next_steps"],1):
            p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
            r=p.add_run(f"{i}. {step}"); r.font.size=Pt(10.5)

    # ── Stage 02 summary ──────────────────────────────────────
    h1(doc,"Stage 02 · Market Intelligence Summary")
    kv(doc,"Score",f"{scores['market']}/10")
    kv(doc,"Market",market.get("market_name",""))
    kv(doc,"Size (2024)",market.get("market_size_2024",""))
    kv(doc,"Projected (2030)",market.get("market_size_2030",""))
    kv(doc,"CAGR",market.get("cagr",""))
    kv(doc,"Maturity",market.get("market_maturity",""))
    kv(doc,"Primary sectors",", ".join(sectors.get("primary_sectors",[])))
    kv(doc,"Competitive intensity",comp.get("competitive_intensity",""))
    kv(doc,"White space",comp.get("white_space",""))
    kv(doc,"Schaeffler advantage",comp.get("schaeffler_advantage",""))
    if comp.get("competitors"):
        doc.add_paragraph()
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
        doc.add_paragraph()
        h2(doc,"Sector Cluster Scores")
        st2=doc.add_table(rows=1,cols=3); st2.style="Table Grid"
        for i,h in enumerate(["Sector","Score","Rationale"]):
            c=st2.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        primary=sectors.get("primary_sectors",[])
        for idx,(sec,data) in enumerate(sectors["sector_scores"].items()):
            row=st2.add_row(); fill="EAF5EA" if sec in primary else ("EAF1FB" if idx%2==0 else "FFFFFF")
            for c in row.cells: set_bg(c,fill)
            r0=row.cells[0].paragraphs[0].add_run(sec); r0.font.size=Pt(10); r0.bold=(sec in primary)
            r1=row.cells[1].paragraphs[0].add_run(f"{data.get('score',0)}/10"); r1.font.size=Pt(10); r1.bold=True
            r2=row.cells[2].paragraphs[0].add_run(data.get("rationale","")); r2.font.size=Pt(9.5)

    # ── Stage 03 summary ──────────────────────────────────────
    h1(doc,"Stage 03 · Patent Intelligence Summary")
    kv(doc,"Score",f"{scores['patent']}/10")
    kv(doc,"Filing activity",landscape.get("activity_level",""))
    kv(doc,"Trend",landscape.get("filing_trend",""))
    kv(doc,"Novelty signal",ansoff_d.get("novelty_signal",""))
    kv(doc,"IP risk",ansoff_d.get("ip_risk",""))
    sp=ansoff_d.get("schaeffler_position",{})
    kv(doc,"Schaeffler existing IP",sp.get("existing_ip",""))
    kv(doc,"IP gap addressed",sp.get("gap",""))
    if landscape.get("white_spaces"):
        doc.add_paragraph()
        h2(doc,"IP White Spaces")
        for ws in landscape["white_spaces"]: bul(doc,ws)
    if landscape.get("key_filers"):
        doc.add_paragraph()
        h2(doc,"Key Patent Filers")
        ft=doc.add_table(rows=1,cols=4); ft.style="Table Grid"
        for i,h in enumerate(["Company","Type","Threat","Focus"]):
            c=ft.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,fi in enumerate(landscape["key_filers"]):
            row=ft.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            for c in row.cells: set_bg(c,fill)
            for j,val in enumerate([fi.get("company",""),fi.get("type",""),fi.get("threat_level",""),fi.get("focus","")+" "+fi.get("source","")]):
                r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.bold=(j==0)

    # ── Stage 04 summary ──────────────────────────────────────
    h1(doc,"Stage 04 · Technical Feasibility Summary")
    kv(doc,"Score",f"{scores['feasibility']}/10")
    kv(doc,"TRL",trl.get("trl_label",""))
    kv(doc,"Existence",existence.get("existence_verdict",""))
    kv(doc,"Entry readiness",trl.get("schaeffler_entry_readiness",""))
    kv(doc,"Time to readiness",existence.get("time_to_readiness",""))
    body(doc,trl.get("trl_rationale",""))
    if existence.get("evidence"):
        doc.add_paragraph()
        h2(doc,"Evidence")
        ev_tbl=doc.add_table(rows=1,cols=3); ev_tbl.style="Table Grid"
        for i,h in enumerate(["Type","Title","Source"]):
            c=ev_tbl.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,ev in enumerate(existence["evidence"]):
            row=ev_tbl.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            for c in row.cells: set_bg(c,fill)
            for j,val in enumerate([ev.get("type",""),ev.get("title","")+" — "+ev.get("description",""),ev.get("source","")]):
                r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.bold=(j==1)
    if trl.get("key_technical_risks"):
        doc.add_paragraph()
        h2(doc,"Technical Risks")
        rt=doc.add_table(rows=1,cols=3); rt.style="Table Grid"
        for i,h in enumerate(["Risk","Severity","Mitigation"]):
            c=rt.cell(0,i); set_bg(c,"1F3864"); r=c.paragraphs[0].add_run(h)
            r.bold=True; r.font.size=Pt(10); r.font.color.rgb=WHITE
        for idx,risk in enumerate(trl["key_technical_risks"]):
            row=rt.add_row(); fill="EAF1FB" if idx%2==0 else "FFFFFF"
            for c in row.cells: set_bg(c,fill)
            for j,val in enumerate([risk.get("risk",""),risk.get("severity",""),risk.get("mitigation","")]):
                r=row.cells[j].paragraphs[0].add_run(val); r.font.size=Pt(9.5); r.bold=(j==0)

    # ── Footer ────────────────────────────────────────────────
    doc.add_paragraph()
    ft=doc.add_table(rows=1,cols=1); ft.style="Table Grid"; fc=ft.cell(0,0); set_bg(fc,"1F3864")
    fp=fc.paragraphs[0]; fp.paragraph_format.space_before=Pt(6); fp.paragraph_format.space_after=Pt(6)
    fr=fp.add_run(f"Schaeffler AI Innovation Research Assistant  ·  Full Innovation Assessment  ·  {datetime.now().strftime('%d %B %Y')}  ·  Capstone Project — Arpan Chowdhury, EBS Universität")
    fr.font.size=Pt(8); fr.font.color.rgb=RGBColor(0x93,0xC5,0xFD)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf


def generate_feasibility_report(idea, quadrant, s1c, existence, trl, scores):
    """Generate a Technical Feasibility Word report."""
    extended = call_claude(
        """Write a professional technical feasibility report for Schaeffler Group. Return ONLY valid JSON:
{
  "executive_summary": "3-4 paragraph summary of the technical feasibility assessment",
  "technology_analysis": "3-4 paragraphs on the core technology, state of the art, and maturity",
  "schaeffler_readiness": "2-3 paragraphs on Schaeffler-specific readiness and capability fit",
  "development_pathway": "2-3 paragraphs on suggested development pathway from current TRL to deployment",
  "risks": ["technical risk 1 with mitigation", "risk 2", "risk 3"],
  "recommendations": ["rec 1", "rec 2", "rec 3"]
}""",
        f"Idea: {idea}\nTRL: {trl.get('trl_level','')}\nExistence: {existence.get('existence_verdict','')}\nEntry readiness: {trl.get('schaeffler_entry_readiness','')}\nTime to readiness: {existence.get('time_to_readiness','')}\nGaps: {existence.get('technology_gaps',[])}\nScore: {scores['final_score']}/10",
        max_tokens=1500
    )
    try:
        ext = json.loads(extended.strip().replace("```json","").replace("```","").strip())
    except:
        ext = {"executive_summary":"See data below.","technology_analysis":"","schaeffler_readiness":"","development_pathway":"","risks":[],"recommendations":[]}

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
    extended = call_claude(
        """Write a professional patent intelligence report for Schaeffler Group. Return ONLY valid JSON:
{
  "executive_summary": "3-4 paragraph summary of the patent landscape and IP opportunity",
  "landscape_analysis": "3-4 paragraphs on filing activity, trends, and key players",
  "ip_strategy": "2-3 paragraphs on recommended IP strategy for Schaeffler",
  "risks": ["IP risk 1 with mitigation", "IP risk 2 with mitigation", "IP risk 3"],
  "recommendations": ["rec 1", "rec 2", "rec 3"]
}""",
        f"Idea: {idea}\nQuadrant: {quadrant}\nActivity: {landscape.get('activity_level','')}\nTrend: {landscape.get('filing_trend','')}\nNovelty: {ansoff_data.get('novelty_signal','')}\nIP risk: {ansoff_data.get('ip_risk','')}\nWhite spaces: {landscape.get('white_spaces',[])}\nScore: {scores['final_score']}/10",
        max_tokens=1500
    )
    try:
        ext = json.loads(extended.strip().replace("```json","").replace("```","").strip())
    except:
        ext = {"executive_summary":"See data below.","landscape_analysis":"","ip_strategy":"","risks":[],"recommendations":[]}

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
    # Get Claude to write extended analysis
    extended = call_claude(
        """Write a professional market intelligence report for Schaeffler Group. Return ONLY valid JSON:
{
  "executive_summary": "3-4 paragraph executive summary",
  "market_deep_dive": "3-4 paragraphs on market dynamics and trends",
  "competitive_analysis": "3-4 paragraphs on competitive landscape",
  "schaeffler_fit": "2-3 paragraphs on strategic fit for Schaeffler",
  "risks": ["risk 1 with mitigation", "risk 2 with mitigation", "risk 3 with mitigation"],
  "recommendations": ["rec 1", "rec 2", "rec 3"]
}""",
        f"Idea: {idea}\nQuadrant: {quadrant}\nMarket: {market.get('market_name','')}\nSize 2024: {market.get('market_size_2024','')}\nCAGR: {market.get('cagr','')}\nCompetition: {comp.get('competitive_intensity','')}\nPrimary sectors: {', '.join(sectors.get('primary_sectors',[]))}\nFinal score: {final_score}/10",
        max_tokens=2000
    )
    try:
        ext = json.loads(extended.strip().replace("```json","").replace("```","").strip())
    except:
        ext = {"executive_summary":"See data below.","market_deep_dive":"","competitive_analysis":"","schaeffler_fit":"","risks":[],"recommendations":[]}

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
    kv(doc,"Market size (2024)",market.get("market_size_2024","N/A"))
    kv(doc,"Projected (2030)",market.get("market_size_2030","N/A"))
    kv(doc,"CAGR",market.get("cagr","N/A"))
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
    risks_list = synthesis.get("risks", [])
    if not risks_list:
        risks_list = [
            f"IP risk level: {ansoff_d.get('ip_risk','Medium')} — conduct freedom-to-operate analysis before committing R&D budget",
            f"Technical readiness at TRL {trl.get('trl_level',3)} — further development required before production readiness",
            "Market timing risk — validate demand with target customers before scaling investment"
        ]
    for risk in risks_list: bul(doc, risk)
    h1(doc,"Recommendations")
    recs_list = synthesis.get("next_steps", [])
    if not recs_list:
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


if st.session_state.active_stage == 1:
    st.markdown("## Stage 01 · Quadrant Classifier")
    st.markdown("""<div style="background:#1a2d45;border-radius:8px;padding:14px 18px;margin-bottom:16px;border-left:4px solid #2E75B6;">
<div style="color:#60a5fa;font-size:11px;letter-spacing:1px;font-weight:600;">WHAT THIS STAGE DOES</div>
<div style="color:#e2e8f0;font-size:13px;margin-top:6px;">Maps your idea onto Schaeffler's modified Ansoff matrix — Exploit, Extend, Radical, or Disrupt. Ideas in Radical and Disrupt proceed through the full pipeline. Others are redirected to the right Schaeffler product division.</div>
<div style="color:#94a3b8;font-size:12px;margin-top:8px;"><b style="color:#e2e8f0;">You get:</b> Quadrant classification · Schaeffler Motion product family fit · Strategic trend alignment · Innovation pathway (Start-Up Mode vs Innovation Factory)</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    # Step 1 — Input
    if st.session_state.s1_step == 1:
        st.subheader("Step 1 — Describe your idea")
        idea = st.text_area("What is your innovation idea?", height=150,
            placeholder="e.g. A self-lubricating bearing system that uses micro-reservoirs embedded within the bearing material to release lubricant automatically based on temperature and load sensing...")

        if st.button("Submit idea", type="primary"):
            if not idea.strip():
                st.warning("Please enter your idea first.")
            elif len(idea.split()) < 15:
                st.warning("A bit brief — can you add more detail? What does it do, and for whom?")
            else:
                with st.spinner("Checking your idea..."):
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
                    st.session_state.s1_step = 2
                    st.rerun()

    # Step 2 — Generate questions
    if st.session_state.s1_step == 2 and not st.session_state.s1_questions:
        with st.spinner("Preparing follow-up questions..."):
            try:
                raw = call_claude(
                    "You are an innovation manager at Schaeffler. Ask exactly 3 sharp follow-up questions to classify an idea as EXPLOIT, EXTEND, RADICAL, or DISRUPT. Cover: technology novelty, target customer, core problem. Return ONLY a JSON array of 3 strings.",
                    st.session_state.s1_idea
                )
                st.session_state.s1_questions = json.loads(raw.strip().replace("```json","").replace("```","").strip())
            except:
                st.session_state.s1_questions = [
                    "Does this technology already exist anywhere, or is it genuinely new?",
                    "Who is the primary customer — existing Schaeffler clients or a new market?",
                    "What specific problem does it solve and for whom?"
                ]
        st.rerun()

    # Step 2 — Show questions
    if st.session_state.s1_step == 2 and st.session_state.s1_questions:
        st.markdown("---")
        st.subheader("Step 2 — A few follow-up questions")
        st.info(f"**Your idea:** {st.session_state.s1_idea}")

        q = st.session_state.s1_questions
        st.markdown(f"**1. {q[0]}**")
        a1 = st.text_area("", key="a1", height=80)
        st.markdown(f"**2. {q[1]}**")
        a2 = st.text_area("", key="a2", height=80)
        st.markdown(f"**3. {q[2]}**")
        a3 = st.text_area("", key="a3", height=80)

        if st.button("Classify my idea →", type="primary"):
            if not a1.strip() or not a2.strip() or not a3.strip():
                st.warning("Please answer all three questions.")
            else:
                st.session_state.s1_answers = [a1, a2, a3]
                st.session_state.s1_step = 3
                st.rerun()

    # Step 3 — Classify
    if st.session_state.s1_step == 3 and not st.session_state.s1_classification:
        with st.spinner("Classifying your idea..."):
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
  "reasoning":"2-3 sentences",
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
                raw = call_claude(system, f"Idea: {st.session_state.s1_idea}\nQ: {q[0]} A: {a[0]}\nQ: {q[1]} A: {a[1]}\nQ: {q[2]} A: {a[2]}")
                st.session_state.s1_classification = json.loads(raw.strip().replace("```json","").replace("```","").strip())
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
        st.subheader("Result")
        st.info(f"**Your idea:** {st.session_state.s1_idea}")

        if not proceed:
            division = c.get("schaeffler_division","Product Development")
            st.warning(f"**{quadrant}** · {c.get('redirect_message','')}")
            st.markdown(f"→ Suggested home: **{division}**")
        else:
            emoji = "🔬" if quadrant == "RADICAL" else "🚀"
            st.success(f"{emoji} **{quadrant}** — {c.get('reasoning','')}")
            st.caption(f"Tech axis: {c.get('tech_axis_level','')}  ·  Market axis: {c.get('market_axis_level','')}")

        # ── Enriched classification details ──────────────────
        col_a, col_b = st.columns(2)
        with col_a:
            if c.get("trend_alignment"):
                trends = " · ".join(c["trend_alignment"])
                st.markdown(f"**Trend alignment:** {trends}")
            if c.get("innovation_cluster"):
                st.markdown(f"**Innovation cluster:** {c.get('innovation_cluster','')}")
        with col_b:
            if c.get("product_family"):
                st.markdown(f"**Product family:** {c.get('product_family','')}")
            if c.get("pipeline_route"):
                route_col = "#22c55e" if "Innovation" in c.get("pipeline_route","") else "#60a5fa"
                st.markdown(f'**Pipeline route:** <span style="color:{route_col};font-weight:600;">{c.get("pipeline_route","")}</span>', unsafe_allow_html=True)

            # Product family, trends, pathway
            col_pf, col_tr, col_pw = st.columns(3)
            with col_pf:
                pf = c.get("product_family","")
                if pf:
                    st.markdown(f"""<div style="background:#1a2d45;border-radius:6px;padding:10px 12px;">
<div style="color:#4a6fa5;font-size:10px;letter-spacing:1px;">PRODUCT FAMILY</div>
<div style="color:#e2e8f0;font-size:13px;font-weight:600;margin-top:4px;">{pf}</div>
</div>""", unsafe_allow_html=True)
            with col_tr:
                trends = c.get("trend_alignment",[])
                if trends:
                    trend_badges = " ".join([f'<span style="background:#1F3864;color:#60a5fa;font-size:10px;padding:2px 7px;border-radius:8px;margin:2px;display:inline-block;">{t}</span>' for t in trends])
                    st.markdown(f"""<div style="background:#1a2d45;border-radius:6px;padding:10px 12px;">
<div style="color:#4a6fa5;font-size:10px;letter-spacing:1px;">TREND ALIGNMENT</div>
<div style="margin-top:4px;">{trend_badges}</div>
</div>""", unsafe_allow_html=True)
            with col_pw:
                pw = c.get("innovation_pathway","")
                pw_r = c.get("pathway_rationale","")
                if pw:
                    pw_col = "#22c55e" if pw == "Start-Up Mode" else "#60a5fa" if pw == "Innovation Factory" else "#f59e0b"
                    st.markdown(f"""<div style="background:#1a2d45;border-radius:6px;padding:10px 12px;">
<div style="color:#4a6fa5;font-size:10px;letter-spacing:1px;">PATHWAY</div>
<div style="color:{pw_col};font-size:13px;font-weight:600;margin-top:4px;">{pw}</div>
</div>""", unsafe_allow_html=True)
                    st.caption(pw_r)

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
            st.success("✓ This idea qualifies for the full Innovation pipeline.")
            if st.button("Continue to Stage 02: Market Intelligence →", type="primary", key="s1_continue"):
                st.session_state.active_stage = 2
                st.rerun()

        # Post-result chat
        st.markdown("---")
        st.subheader("💬 Questions about this classification?")
        for msg in st.session_state.s1_chat:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        user_q = st.chat_input("Ask about the classification...")
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

        if st.button("← Start over", key="s1_startover"):
            for k in ["s1_step","s1_idea","s1_questions","s1_answers","s1_classification","s1_chat"]:
                st.session_state[k] = defaults[k]
            st.rerun()

# ════════════════════════════════════════════════════════════
# STAGE 02 — MARKET INTELLIGENCE
# ════════════════════════════════════════════════════════════
elif st.session_state.active_stage == 2:
    st.markdown("## Stage 02 · Market Intelligence")
    st.markdown("""<div style="background:#1a2d45;border-radius:8px;padding:14px 18px;margin-bottom:16px;border-left:4px solid #2E75B6;">
<div style="color:#60a5fa;font-size:11px;letter-spacing:1px;font-weight:600;">WHAT THIS STAGE DOES</div>
<div style="color:#e2e8f0;font-size:13px;margin-top:6px;">Analyses the commercial opportunity behind your idea — how big the market is, how fast it is growing, who the competitors are, and how well the idea fits across Schaeffler's 10 customer sector clusters.</div>
<div style="color:#94a3b8;font-size:12px;margin-top:8px;"><b style="color:#e2e8f0;">You get:</b> Market size & CAGR with sources · Sector cluster fit chart · Competitor landscape · Market Intelligence Score (0–10)</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    idea     = st.session_state.s1_idea
    s1c      = st.session_state.s1_classification
    quadrant = s1c.get("quadrant","RADICAL")

    # Intro
    if st.session_state.s2_step == "intro":
        st.info(f"**Idea:** {idea}")
        st.markdown(f"**Quadrant:** {quadrant} · {s1c.get('technology_novelty','')}")
        if st.button("Run Market Intelligence →", type="primary"):
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
RULES: Every data point MUST include [Source: Org, Year]. Use only credible sources (McKinsey, Gartner, Frost & Sullivan, BloombergNEF, IEA, Roland Berger, Statista, industry associations). If uncertain, give a range.
Return ONLY valid JSON:
{"market_name":"string","market_size_2024":"value [Source: X, Y]","market_size_2030":"value [Source: X, Y]",
"cagr":"% [Source: X, Y]","growth_drivers":["driver with source x3"],"market_maturity":"Emerging/Growing/Mature/Declining",
"geographic_focus":"string","market_score":integer 1-10 (9-10=large fast-growing >$10bn/>15%CAGR; 7-8=strong $2-10bn/8-15%; 5-6=moderate; 3-4=niche; 1-2=tiny or declining),"market_score_rationale":"2 sentences"}"""
        try:
            raw = call_claude(system_market, f"Idea: {idea}\nQuadrant: {quadrant}\nTech: {s1c.get('technology_novelty','')}{web_ctx}")
            market = json.loads(raw.strip().replace("```json","").replace("```","").strip())
        except:
            market = {"market_name":"N/A","market_size_2024":"N/A","market_size_2030":"N/A","cagr":"N/A",
                      "growth_drivers":[],"market_maturity":"N/A","geographic_focus":"N/A","market_score":5,"market_score_rationale":""}

        status.markdown("🏢 Mapping competitive landscape...")
        progress.progress(55)
        system_comp = """You are a competitive intelligence analyst. Identify key competitors for this idea.
Every company must have a source. Return ONLY valid JSON:
{"competitors":[{"name":"string","type":"Incumbent/Startup/Research","relevance":"one sentence","source":"Source: X, Y"}],
"competitive_intensity":"Low/Medium/High/Very High","white_space":"one sentence","schaeffler_advantage":"one sentence",
"competition_score":integer 1-10 openness (9-10=very open/few players; 7-8=some room; 5-6=moderate; 3-4=crowded; 1-2=saturated),"competition_score_rationale":"2 sentences"}"""
        try:
            raw = call_claude(system_comp, f"Idea: {idea}\nMarket: {market.get('market_name','')}\nQuadrant: {quadrant}{web_ctx}")
            comp = json.loads(raw.strip().replace("```json","").replace("```","").strip())
        except:
            comp = {"competitors":[],"competitive_intensity":"N/A","white_space":"N/A",
                    "schaeffler_advantage":"N/A","competition_score":5,"competition_score_rationale":""}

        status.markdown("🎯 Scoring Schaeffler sector clusters...")
        progress.progress(75)
        system_sectors = """You are a Schaeffler strategist. Score fit against Schaeffler's 10 sector clusters (0-10 each).
Clusters: Passenger Cars, Commercial Vehicles, Industrial Machinery, Rail, Aerospace, Two-Wheelers, Construction & Agriculture, Medical Equipment, Conventional Energy, Renewable Energy.
0-2=No relevance, 3-4=Low, 5-6=Moderate, 7-8=High, 9-10=Primary target.
Return ONLY valid JSON:
{"sector_scores":{"Passenger Cars":{"score":0-10,"rationale":"one sentence"},"Commercial Vehicles":{"score":0-10,"rationale":"one sentence"},"Industrial Machinery":{"score":0-10,"rationale":"one sentence"},"Rail":{"score":0-10,"rationale":"one sentence"},"Aerospace":{"score":0-10,"rationale":"one sentence"},"Two-Wheelers":{"score":0-10,"rationale":"one sentence"},"Construction & Agriculture":{"score":0-10,"rationale":"one sentence"},"Medical Equipment":{"score":0-10,"rationale":"one sentence"},"Conventional Energy":{"score":0-10,"rationale":"one sentence"},"Renewable Energy":{"score":0-10,"rationale":"one sentence"}},
"primary_sectors":["top 2-3 sector names"],"sector_fit_score":0-10,"sector_fit_rationale":"2 sentences"}
Sector fit score rubric: Average the top 3 sector scores. If primary sector scores 9-10 = sector_fit 9-10; if 7-8 = 7-8; etc."""
        try:
            raw = call_claude(system_sectors, f"Idea: {idea}\nQuadrant: {quadrant}")
            sectors = json.loads(raw.strip().replace("```json","").replace("```","").strip())
        except:
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
        banner = f"""<div style="background:{BG};border-radius:8px;padding:20px 24px;margin-bottom:20px;border:1px solid {DIM};"><div style="color:{WHITE};font-size:11px;letter-spacing:1.5px;opacity:0.5;margin-bottom:4px;">MARKET INTELLIGENCE SCORE</div><div style="color:{score_col};font-size:44px;font-weight:700;line-height:1;">{final}<span style="font-size:18px;color:{DIM};"> / 10</span></div><div style="color:{WHITE};font-size:13px;margin-top:6px;opacity:0.8;">{market.get('market_name','')}</div></div>"""
        st.markdown(banner, unsafe_allow_html=True)

        # ── Score breakdown ───────────────────────────────────
        cols = st.columns(3)
        for i,(label,(score,weight)) in enumerate(weights.items()):
            cols[i].metric(label, f"{score:.1f}/10", f"{int(weight*100)}% weight")
        st.markdown("---")

        # ── Market size ───────────────────────────────────────
        st.markdown("#### 📊 Market")
        c1,c2,c3 = st.columns(3)
        c1.metric("Size (2024)", market.get("market_size_2024","N/A").split("[")[0].strip())
        c2.metric("Size (2030)", market.get("market_size_2030","N/A").split("[")[0].strip())
        c3.metric("CAGR", market.get("cagr","N/A").split("[")[0].strip())
        srcs = []
        for field in ["market_size_2024","market_size_2030","cagr"]:
            val = market.get(field,"")
            if "[Source:" in val:
                srcs.append(val[val.find("[Source:")+8:val.find("]",val.find("[Source:"))])
        if srcs:
            st.caption("Sources: " + "  ·  ".join(dict.fromkeys(srcs)))
        col_l,col_r = st.columns(2)
        col_l.markdown(f"**Maturity** · {market.get('market_maturity','')}")
        col_r.markdown(f"**Geography** · {market.get('geographic_focus','')}")
        if market.get("growth_drivers"):
            with st.expander(f"Growth drivers (top {min(3,len(market['growth_drivers']))} of {len(market['growth_drivers'])})"):
                for drv in market["growth_drivers"][:3]:
                    st.markdown(f"- {drv}")
                if len(market["growth_drivers"]) > 3:
                    st.caption("Full list in the downloaded report.")
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
        if "s2_report_buf" not in st.session_state:
            st.session_state.s2_report_buf = None

        if st.session_state.s2_report_buf is None:
            if st.button("⬇️ Download Market Intelligence Report", type="primary"):
                with st.spinner("Generating your report — this takes about 20 seconds..."):
                    try:
                        st.session_state.s2_report_buf = generate_market_report(
                            idea, quadrant, s1c, market, comp, sectors, weights, final
                        )
                        st.rerun()
                    except Exception as e:
                        st.error(f"Report generation error: {e}")
        else:
            st.download_button(
                label="⬇️ Download Market Intelligence Report",
                data=st.session_state.s2_report_buf,
                file_name=f"Schaeffler_Market_Intelligence_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )

        # ── Chat ──────────────────────────────────────────────
        st.markdown("---")
        st.subheader("💬 Questions about the market analysis?")
        for msg in st.session_state.s2_chat:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        user_q = st.chat_input("Ask about the market, sectors, competitors...")
        if user_q:
            st.session_state.s2_chat.append({"role":"user","content":user_q})
            with st.chat_message("user"):
                st.markdown(user_q)
            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    ctx = f"""You are a senior market analyst discussing market intelligence results for a Schaeffler innovation idea.
Idea: {idea} | Quadrant: {quadrant}
Market: {market.get('market_name','')} | Size 2024: {market.get('market_size_2024','')} | CAGR: {market.get('cagr','')}
Competitive intensity: {comp.get('competitive_intensity','')} | White space: {comp.get('white_space','')}
Primary sectors (Schaeffler's 10 clusters): {', '.join(primary)} | Final score: {final}/10
Be specific, cite sources where possible, 3-4 sentences max."""
                    history = [{"role":m["role"],"content":m["content"]} for m in st.session_state.s2_chat]
                    reply = call_claude_chat(ctx, history)
                    st.markdown(reply)
                    st.session_state.s2_chat.append({"role":"assistant","content":reply})

        # ── Continue ──────────────────────────────────────────
        st.markdown("---")
        st.success(f"✓ Market Intelligence complete. Final score: **{final}/10**")
        if st.button("Continue to Stage 03: Patent Intelligence →", type="primary", key="s2_continue"):
            st.session_state.active_stage = 3
            st.rerun()

        if st.button("← Re-run analysis", key="s2_rerun"):
            st.session_state.s2_step = "intro"
            st.session_state.s2_data = {}
            st.session_state.s2_chat = []
            st.session_state.s2_report_buf = None
            st.rerun()





# ════════════════════════════════════════════════════════════
# STAGE 03 — PATENT INTELLIGENCE
# ════════════════════════════════════════════════════════════
elif st.session_state.active_stage == 3:
    st.markdown("## Stage 03 · Patent Intelligence")
    st.markdown("""<div style="background:#1a2d45;border-radius:8px;padding:14px 18px;margin-bottom:16px;border-left:4px solid #2E75B6;">
<div style="color:#60a5fa;font-size:11px;letter-spacing:1px;font-weight:600;">WHAT THIS STAGE DOES</div>
<div style="color:#e2e8f0;font-size:13px;margin-top:6px;">Maps the patent landscape for your idea's core technology — who is filing, whether they are competitors or potential customers, where the IP white spaces are, and how Schaeffler's existing patent portfolio relates to the idea.</div>
<div style="color:#94a3b8;font-size:12px;margin-top:8px;"><b style="color:#e2e8f0;">You get:</b> Patent Ansoff map with all key filers plotted · IP white spaces · Schaeffler IP gap analysis · Patent Intelligence Score (0–10)</div>
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
        if st.button("Run Patent Intelligence →", type="primary"):
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
                f"Idea: {idea}\nQuadrant: {quadrant}\nTech novelty: {s1c.get('technology_novelty','')}")
            landscape = json.loads(raw.strip().replace("```json","").replace("```","").strip())
        except Exception as e:
            landscape = {"technology_keywords":[],"landscape_summary":"Analysis unavailable.",
                        "activity_level":"N/A","filing_trend":"N/A","filing_trend_rationale":"",
                        "key_filers":[],"white_spaces":[],"patent_landscape_score":5}

        progress.progress(50)
        status.markdown("🏢 Mapping filers onto Schaeffler's Ansoff matrix...")

        system_ansoff = """You are a Schaeffler Group patent strategist.
Map patent filing companies onto Schaeffler's modified Ansoff matrix based on where their patents sit.

The matrix axes:
- X axis: Market Dimension (Existing Market → New Market)
- Y axis: Technology Dimension (Existing Technology → New Technology)

Quadrants:
- EXPLOIT (existing tech, existing market): incremental improvements, defensive filings
- EXTEND (existing tech, new market): technology transfer to new applications
- RADICAL (new tech, existing market): breakthrough technology for known customers
- DISRUPT (new tech, new market): entirely new technology for new markets

For each company, assign:
- matrix_position: which quadrant their patent activity sits in
- x_score: 0-10 (0=existing market, 10=new market)
- y_score: 0-10 (0=existing tech, 10=new tech)

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

        filers_context = json.dumps([f.get("company","") for f in landscape.get("key_filers",[])])
        try:
            raw2 = call_claude(system_ansoff,
                f"Idea: {idea}\nQuadrant: {quadrant}\nKey filers: {filers_context}\nTech keywords: {landscape.get('technology_keywords','')}")
            ansoff_data = json.loads(raw2.strip().replace("```json","").replace("```","").strip())
        except Exception as e:
            ansoff_data = {"filer_positions":[],"schaeffler_position":{"matrix_position":"EXPLOIT","x_score":2,"y_score":2,"existing_ip":"N/A","gap":"N/A"},
                          "idea_position":{"x_score":7,"y_score":7},"novelty_signal":"Moderate","novelty_rationale":"","ip_risk":"Medium","ip_risk_rationale":""}

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
<div style="background:{BG};border-radius:8px;padding:20px 24px;margin-bottom:20px;border:1px solid {DIM};">
  <div style="color:{WHITE};font-size:11px;letter-spacing:1px;opacity:0.5;margin-bottom:4px;">PATENT INTELLIGENCE SCORE</div>
  <div style="color:{score_col};font-size:42px;font-weight:700;line-height:1;">{final}<span style="font-size:18px;color:{DIM};"> / 10</span></div>
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
        st.caption("Each point = a company's patent filing position. Your idea shown in green. Schaeffler's existing IP shown in orange.")

        filer_positions = ansoff_data.get("filer_positions", [])
        schaeffler_pos  = ansoff_data.get("schaeffler_position", {})
        idea_pos        = ansoff_data.get("idea_position", {"x_score":7,"y_score":7})

        fig = go.Figure()

        # Quadrant shading
        q_fills = [
            dict(x=[0,5,5,0],   y=[0,0,5,5],   name="EXPLOIT", fill="#1a2d45", lx=2.5,ly=2.5),
            dict(x=[5,10,10,5], y=[0,0,5,5],   name="EXTEND",  fill="#1e3a5f", lx=7.5,ly=2.5),
            dict(x=[0,5,5,0],   y=[5,5,10,10], name="RADICAL", fill="#1F3864", lx=2.5,ly=7.5),
            dict(x=[5,10,10,5], y=[5,5,10,10], name="DISRUPT", fill="#0d2137", lx=7.5,ly=7.5),
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
        if "s3_report_buf" not in st.session_state:
            st.session_state.s3_report_buf = None

        if st.session_state.s3_report_buf is None:
            if st.button("⬇️ Download Patent Intelligence Report", type="primary"):
                with st.spinner("Generating report..."):
                    try:
                        st.session_state.s3_report_buf = generate_patent_report(
                            idea, quadrant, s1c, landscape, ansoff_data, d
                        )
                        st.rerun()
                    except Exception as e:
                        st.error(f"Report error: {e}")
        else:
            st.download_button(
                label="⬇️ Download Patent Intelligence Report",
                data=st.session_state.s3_report_buf,
                file_name=f"Schaeffler_Patent_Intelligence_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )

        # ── Chat ──────────────────────────────────────────────
        st.markdown("---")
        st.subheader("💬 Questions about the patent landscape?")
        for msg in st.session_state.s3_chat:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        user_q = st.chat_input("Ask about patents, IP position, competitors...")
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
        st.success(f"✓ Patent Intelligence complete. Score: **{final}/10**")
        if st.button("Continue to Stage 04: Technical Feasibility →", type="primary", key="s3_continue"):
            st.session_state.active_stage = 4
            st.rerun()

        if st.button("← Re-run analysis", key="s3_rerun"):
            st.session_state.s3_step = "intro"
            st.session_state.s3_data = {}
            st.session_state.s3_chat = []
            st.session_state.s3_report_buf = None
            st.rerun()



# ════════════════════════════════════════════════════════════
# STAGE 04 — TECHNICAL FEASIBILITY
# ════════════════════════════════════════════════════════════
elif st.session_state.active_stage == 4:
    st.markdown("## Stage 04 · Technical Feasibility")
    st.markdown("""<div style="background:#1a2d45;border-radius:8px;padding:14px 18px;margin-bottom:16px;border-left:4px solid #2E75B6;">
<div style="color:#60a5fa;font-size:11px;letter-spacing:1px;font-weight:600;">WHAT THIS STAGE DOES</div>
<div style="color:#e2e8f0;font-size:13px;margin-top:6px;">Checks whether the core technology has actually been demonstrated anywhere — in labs, startups, pilots, or adjacent industries. Rates maturity using a Schaeffler-adapted version of NASA's TRL scale (1–9) and identifies the key technical risks to address.</div>
<div style="color:#94a3b8;font-size:12px;margin-top:8px;"><b style="color:#e2e8f0;">You get:</b> TRL rating with rationale · Evidence from research & industry · Technology keyword map · Risk register · Feasibility Score (0–10)</div>
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
        if st.button("Run Technical Feasibility Analysis →", type="primary"):
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
                        "evidence":[],"technology_gaps":[],"time_to_readiness":"Unknown","keywords":[]}

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
<div style="background:{BG};border-radius:8px;padding:20px 24px;margin-bottom:20px;border:1px solid {DIM};">
  <div style="color:{WHITE};font-size:11px;letter-spacing:1px;opacity:0.5;margin-bottom:4px;">TECHNICAL FEASIBILITY SCORE</div>
  <div style="display:flex;align-items:flex-end;gap:24px;">
    <div style="color:{score_col};font-size:42px;font-weight:700;line-height:1;">{final}<span style="font-size:18px;color:{DIM};"> / 10</span></div>
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
        if trl.get("analogous_schaeffler_technologies"):
            st.caption(f"Schaeffler analogous experience: {trl.get('analogous_schaeffler_technologies','')}")
        st.markdown("---")

        # ── TRL scale reference ───────────────────────────────
        with st.expander("Schaeffler-adapted TRL scale reference"):
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
            st.markdown(f'<div style="background:{BG};border-radius:8px;padding:16px;line-height:2.2;border:1px solid {DIM};">{badges}</div>',
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
        if "s4_report_buf" not in st.session_state:
            st.session_state.s4_report_buf = None

        if st.session_state.s4_report_buf is None:
            if st.button("⬇️ Download Technical Feasibility Report", type="primary"):
                with st.spinner("Generating report..."):
                    try:
                        st.session_state.s4_report_buf = generate_feasibility_report(
                            idea, quadrant, s1c, existence, trl, d
                        )
                        st.rerun()
                    except Exception as e:
                        st.error(f"Report error: {e}")
        else:
            st.download_button(
                label="⬇️ Download Technical Feasibility Report",
                data=st.session_state.s4_report_buf,
                file_name=f"Schaeffler_Technical_Feasibility_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )

        # ── Chat ──────────────────────────────────────────────
        st.markdown("---")
        st.subheader("💬 Questions about technical feasibility?")
        for msg in st.session_state.s4_chat:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        user_q = st.chat_input("Ask about TRL, evidence, risks, or readiness...")
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
        st.success(f"✓ Technical Feasibility complete. Score: **{final}/10**")
        if st.button("Continue to Stage 05: Scoring & Synthesis →", type="primary", key="s4_continue"):
            st.session_state.active_stage = 5
            st.rerun()

        if st.button("← Re-run analysis", key="s4_rerun"):
            st.session_state.s4_step = "intro"
            st.session_state.s4_data = {}
            st.session_state.s4_chat = []
            st.session_state.s4_report_buf = None
            st.rerun()



# ════════════════════════════════════════════════════════════
# STAGE 05 — SCORING & SYNTHESIS
# ════════════════════════════════════════════════════════════
elif st.session_state.active_stage == 5:
    st.markdown("## Stage 05 · Scoring & Synthesis")
    st.markdown("""<div style="background:#1a2d45;border-radius:8px;padding:14px 18px;margin-bottom:16px;border-left:4px solid #2E75B6;">
<div style="color:#60a5fa;font-size:11px;letter-spacing:1px;font-weight:600;">WHAT THIS STAGE DOES</div>
<div style="color:#e2e8f0;font-size:13px;margin-top:6px;">Combines scores from Stages 02, 03, and 04 into a single Innovation Potential Index (IPI). You set the weights. The assistant generates a final recommendation, strategic synthesis, and a downloadable master report covering the full pipeline analysis.</div>
<div style="color:#94a3b8;font-size:12px;margin-top:8px;"><b style="color:#e2e8f0;">You get:</b> Weighted IPI score · Radar chart · PROCEED / DEFER / REJECT recommendation · Strongest signals & concerns · Next steps · Full Innovation Assessment Report</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    idea     = st.session_state.s1_idea
    s1c      = st.session_state.s1_classification
    quadrant = s1c.get("quadrant","RADICAL")

    for k, v in {"s5_step":"intro","s5_data":{},"s5_chat":[]}.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # ── Check all stages are complete ─────────────────────────
    s2_done = bool(st.session_state.get("s2_data"))
    s3_done = bool(st.session_state.get("s3_data"))
    s4_done = bool(st.session_state.get("s4_data"))

    if not (s2_done and s3_done and s4_done):
        st.warning("Complete Stages 02, 03, and 04 first before running the final synthesis.")
        missing = []
        if not s2_done: missing.append("Stage 02: Market Intelligence")
        if not s3_done: missing.append("Stage 03: Patent Intelligence")
        if not s4_done: missing.append("Stage 04: Technical Feasibility")
        for m in missing:
            st.markdown(f"- ⬜ {m}")
        if st.button("← Back", key="s5_back2"):
            st.session_state.active_stage = 4
            st.rerun()
        st.stop()

    # Pull scores from previous stages
    market_score    = st.session_state.s2_data.get("final_score", 5.0)
    patent_score    = st.session_state.s3_data.get("final_score", 5.0)
    feasibility_score = st.session_state.s4_data.get("final_score", 5.0)

    # ── Intro: show weights + let user adjust ─────────────────
    if st.session_state.s5_step == "intro":
        st.info(f"**Idea:** {idea}")
        st.markdown(f"**Quadrant:** {quadrant}")
        st.markdown("---")

        st.markdown("#### Scores from previous stages")
        c1, c2, c3 = st.columns(3)
        c1.metric("Market Intelligence",    f"{market_score} / 10")
        c2.metric("Patent Intelligence",    f"{patent_score} / 10")
        c3.metric("Technical Feasibility",  f"{feasibility_score} / 10")

        st.markdown("---")
        st.markdown("#### Innovation Potential Index — Scoring Weights")
        st.caption("Default weights set for first iteration. Adjust and refine with Johannes Enders.")

        col1, col2, col3 = st.columns(3)
        with col1:
            w_market = st.slider("Market Intelligence", 0, 100, 40, 5, key="w_market")
        with col2:
            w_patent = st.slider("Patent Intelligence", 0, 100, 30, 5, key="w_patent")
        with col3:
            w_feasibility = st.slider("Technical Feasibility", 0, 100, 30, 5, key="w_feasibility")

        total_weight = w_market + w_patent + w_feasibility
        if total_weight != 100:
            st.warning(f"Weights must add up to 100. Current total: {total_weight}. Adjust the sliders.")
        else:
            st.success(f"✓ Weights sum to 100")
            if st.button("Run Final Synthesis →", type="primary"):
                st.session_state.s5_weights = {"market": w_market, "patent": w_patent, "feasibility": w_feasibility}
                st.session_state.s5_step = "running"
                st.rerun()

    # ── Running ───────────────────────────────────────────────
    elif st.session_state.s5_step == "running":
        progress = st.progress(0)
        status   = st.empty()

        weights = st.session_state.get("s5_weights", {"market":40,"patent":30,"feasibility":30})

        status.markdown("📐 Calculating Innovation Potential Index...")
        progress.progress(20)

        wm = weights["market"] / 100
        wp = weights["patent"] / 100
        wf = weights["feasibility"] / 100
        ipi = round(market_score * wm + patent_score * wp + feasibility_score * wf, 1)

        status.markdown("🧠 Writing narrative synthesis...")
        progress.progress(45)

        s2d = st.session_state.s2_data
        s3d = st.session_state.s3_data
        s4d = st.session_state.s4_data

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
- Size: {s2d.get('market',{}).get('market_size_2024','')}
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

        st.session_state.s5_data = {
            "ipi": ipi,
            "weights": weights,
            "synthesis": synthesis,
            "scores": {
                "market": market_score,
                "patent": patent_score,
                "feasibility": feasibility_score
            }
        }
        st.session_state.s5_step = "done"
        st.rerun()

    # ── Results ───────────────────────────────────────────────
    elif st.session_state.s5_step == "done":
        d         = st.session_state.s5_data
        ipi       = d.get("ipi", 0)
        weights   = d.get("weights", {"market":40,"patent":30,"feasibility":30})
        synthesis = d.get("synthesis", {})
        scores    = d.get("scores", {"market":5,"patent":5,"feasibility":5})

        rec = synthesis.get("recommendation","PROCEED WITH CONDITIONS")
        rec_colours = {
            "PROCEED":               "#22c55e",
            "PROCEED WITH CONDITIONS":"#f59e0b",
            "DEFER":                 "#f97316",
            "REJECT":                "#ef4444"
        }
        rec_col = rec_colours.get(rec, "#f59e0b")
        ipi_col = "#22c55e" if ipi>=7 else "#f59e0b" if ipi>=4 else "#ef4444"

        # ── IPI banner ────────────────────────────────────────
        st.markdown(f"""
<div style="background:{BG};border-radius:8px;padding:20px 24px;margin-bottom:20px;border:1px solid {DIM};">
  <div style="display:flex;justify-content:space-between;align-items:flex-start;">
    <div>
      <div style="color:{WHITE};font-size:11px;letter-spacing:1px;opacity:0.5;margin-bottom:4px;">INNOVATION POTENTIAL INDEX</div>
      <div style="color:{ipi_col};font-size:48px;font-weight:700;line-height:1;">{ipi}<span style="font-size:20px;color:{DIM};"> / 10</span></div>
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

        categories   = ["Market Intelligence", "Patent Intelligence", "Technical Feasibility"]
        values       = [scores["market"], scores["patent"], scores["feasibility"]]
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
            r=[7,7,7,7], theta=cats_close,
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
        col4.metric("**IPI Score**",          f"**{ipi} / 10**")
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
        with st.expander("📖 Read full narrative synthesis"):
            st.markdown(synthesis.get("narrative",""))

        # ── Visual mockup generation ──────────────────────────
        st.markdown("---")
        st.markdown("#### 🎨 Solution Visualisation")
        st.caption("Generate an AI image showing how this solution could look in a real-world context.")

        if "s5_mockup_image" not in st.session_state:
            st.session_state.s5_mockup_image = None

        if st.session_state.s5_mockup_image is None:
            if st.button("🖼️ Generate Solution Image", type="secondary", key="s5_image"):
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
                            st.session_state.s5_mockup_image = img_response.content
                            st.session_state.s5_mockup_prompt_used = img_prompt
                            st.rerun()
                        else:
                            st.error("Image generation timed out. Try again — Pollinations.ai can be slow on first request.")
                    except Exception as e:
                        st.error(f"Image generation error: {e}")
        else:
            st.image(st.session_state.s5_mockup_image, use_container_width=True)
            st.caption(f"Prompt used: {st.session_state.get('s5_mockup_prompt_used','')}")
            if st.button("🔄 Generate different image", key="s5_image_redo"):
                st.session_state.s5_mockup_image = None
                st.session_state.s5_mockup_prompt_used = None
                st.rerun()

        # ── Master report download ────────────────────────────
        st.markdown("---")
        if "s5_report_buf" not in st.session_state:
            st.session_state.s5_report_buf = None

        if st.session_state.s5_report_buf is None:
            if st.button("⬇️ Download Full Innovation Assessment Report", type="primary"):
                with st.spinner("Generating master report — this covers all 4 stages..."):
                    try:
                        st.session_state.s5_report_buf = generate_master_report(
                            idea, quadrant, s1c,
                            st.session_state.s2_data,
                            st.session_state.s3_data,
                            st.session_state.s4_data,
                            d
                        )
                        st.rerun()
                    except Exception as e:
                        st.error(f"Report error: {e}")
        else:
            st.download_button(
                label="⬇️ Download Full Innovation Assessment Report",
                data=st.session_state.s5_report_buf,
                file_name=f"Schaeffler_Innovation_Assessment_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )
            st.caption("Covers all 4 stages: Market Intelligence · Patent Intelligence · Technical Feasibility · Innovation Potential Index")

        # ── Chat ──────────────────────────────────────────────
        st.markdown("---")
        st.subheader("💬 Questions about the overall assessment?")
        for msg in st.session_state.s5_chat:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        user_q = st.chat_input("Ask about the IPI score, recommendation, or next steps...")
        if user_q:
            st.session_state.s5_chat.append({"role":"user","content":user_q})
            with st.chat_message("user"):
                st.markdown(user_q)
            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    ctx = f"""You are a senior Schaeffler innovation strategist discussing the final assessment.
Idea: {idea} | Quadrant: {quadrant}
IPI Score: {ipi}/10 | Recommendation: {rec}
Market: {scores['market']}/10 | Patent: {scores['patent']}/10 | Feasibility: {scores['feasibility']}/10
Headline: {synthesis.get('headline','')}
Strongest signals: {synthesis.get('strongest_signals',[])}
Key concerns: {synthesis.get('key_concerns',[])}
Strategic fit: {synthesis.get('strategic_fit','')}
Be direct and specific. Reference Schaeffler's context where relevant. 3-4 sentences."""
                    history = [{"role":m["role"],"content":m["content"]} for m in st.session_state.s5_chat]
                    reply = call_claude_chat(ctx, history)
                    st.markdown(reply)
                    st.session_state.s5_chat.append({"role":"assistant","content":reply})

        # ── Re-run with different weights ─────────────────────
        st.markdown("---")
        if st.button("← Adjust weights and re-run", key="s5_rerun"):
            st.session_state.s5_step = "intro"
            st.session_state.s5_data = {}
            st.session_state.s5_chat = []
            st.session_state.s5_report_buf = None
            st.rerun()


