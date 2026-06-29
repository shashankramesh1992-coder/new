"""
Content Brief: How Niching Your Social Media Agency Makes You More Money
SocialPilot — same design template as previous briefs
"""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

BLUE      = "1A56DB"
WHITE     = "FFFFFF"
NEAR_BLACK= "141413"
DARK_GREY = "3D3D3A"
LIGHT_BLUE= "EEF2FF"
ALT_WHITE = "FFFFFF"

def hex_to_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar= OxmlElement("w:tcMar")
    for side, val in [("top",top),("bottom",bottom),("left",left),("right",right)]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"),    str(val))
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)

def set_row_height(row, height_twips):
    trPr = row._tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), str(height_twips))
    trPr.append(trHeight)

def no_space(para):
    pPr = para._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"),  "0")
    pPr.append(spacing)

def section_banner(doc, text):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.rows[0].cells[0]
    set_cell_bg(cell, BLUE)
    set_cell_margins(cell, top=120, bottom=120, left=160, right=160)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    no_space(p)
    run = p.add_run(text)
    run.bold = True
    run.font.color.rgb = hex_to_rgb(WHITE)
    run.font.size = Pt(11)
    run.font.name = "Calibri"
    doc.add_paragraph()

def sub_banner(doc, text):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.rows[0].cells[0]
    set_cell_bg(cell, BLUE)
    set_cell_margins(cell, top=80, bottom=80, left=160, right=160)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    no_space(p)
    run = p.add_run(text)
    run.bold = True
    run.font.color.rgb = hex_to_rgb(WHITE)
    run.font.size = Pt(10)
    run.font.name = "Calibri"
    doc.add_paragraph()

def body_para(doc, text, bold=False, size=9, color=NEAR_BLACK, indent=0):
    p = doc.add_paragraph()
    no_space(p)
    pPr = p._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"),  "60")
    pPr.append(spacing)
    if indent:
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str(indent))
        pPr.append(ind)
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = hex_to_rgb(color)
    run.font.name = "Calibri"
    return p

def label_value_para(doc, label, value, size=9):
    p = doc.add_paragraph()
    no_space(p)
    pPr = p._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"),  "60")
    pPr.append(spacing)
    r1 = p.add_run(label + ": ")
    r1.bold = True
    r1.font.size = Pt(size)
    r1.font.color.rgb = hex_to_rgb(DARK_GREY)
    r1.font.name = "Calibri"
    r2 = p.add_run(value)
    r2.bold = False
    r2.font.size = Pt(size)
    r2.font.color.rgb = hex_to_rgb(NEAR_BLACK)
    r2.font.name = "Calibri"

def italic_note(doc, text, size=8.5):
    p = doc.add_paragraph()
    no_space(p)
    pPr = p._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"),  "80")
    pPr.append(spacing)
    run = p.add_run(text)
    run.italic = True
    run.font.size = Pt(size)
    run.font.color.rgb = hex_to_rgb(DARK_GREY)
    run.font.name = "Calibri"

def bullet(doc, text, bold_prefix=None, size=9, indent=240, hanging=240):
    p = doc.add_paragraph()
    no_space(p)
    pPr = p._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"),  "60")
    pPr.append(spacing)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"),    str(indent))
    ind.set(qn("w:hanging"), str(hanging))
    pPr.append(ind)
    num = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl"); ilvl.set(qn("w:val"), "0"); num.append(ilvl)
    numId= OxmlElement("w:numId"); numId.set(qn("w:val"), "1"); num.append(numId)
    pPr.append(num)
    if bold_prefix:
        r1 = p.add_run(bold_prefix + " ")
        r1.bold = True
        r1.font.size = Pt(size)
        r1.font.color.rgb = hex_to_rgb(NEAR_BLACK)
        r1.font.name = "Calibri"
    r2 = p.add_run(text)
    r2.font.size = Pt(size)
    r2.font.color.rgb = hex_to_rgb(NEAR_BLACK)
    r2.font.name = "Calibri"

def checkbox(doc, text, size=9):
    p = doc.add_paragraph()
    no_space(p)
    pPr = p._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"),  "60")
    pPr.append(spacing)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"),    "240")
    ind.set(qn("w:hanging"), "240")
    pPr.append(ind)
    run = p.add_run("☐  " + text)
    run.font.size = Pt(size)
    run.font.color.rgb = hex_to_rgb(NEAR_BLACK)
    run.font.name = "Calibri"

def kv_table(doc, rows_data):
    tbl = doc.add_table(rows=len(rows_data), cols=2)
    tbl.style = "Table Grid"
    col_widths = [Inches(1.8), Inches(4.7)]
    for i, (label, value) in enumerate(rows_data):
        row  = tbl.rows[i]
        bg   = LIGHT_BLUE if i % 2 == 0 else ALT_WHITE
        c0, c1 = row.cells[0], row.cells[1]
        set_cell_bg(c0, bg); set_cell_bg(c1, bg)
        set_cell_margins(c0); set_cell_margins(c1)
        p0 = c0.paragraphs[0]; no_space(p0)
        r0 = p0.add_run(label)
        r0.bold = True; r0.font.size = Pt(9)
        r0.font.color.rgb = hex_to_rgb(DARK_GREY)
        r0.font.name = "Calibri"
        p1 = c1.paragraphs[0]; no_space(p1)
        r1 = p1.add_run(value)
        r1.font.size = Pt(9)
        r1.font.color.rgb = hex_to_rgb(NEAR_BLACK)
        r1.font.name = "Calibri"
    doc.add_paragraph()

def data_table(doc, headers, rows_data):
    col_count = len(headers)
    tbl = doc.add_table(rows=1+len(rows_data), cols=col_count)
    tbl.style = "Table Grid"
    hdr_row = tbl.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        set_cell_bg(cell, BLUE)
        set_cell_margins(cell)
        p = cell.paragraphs[0]; no_space(p)
        run = p.add_run(h)
        run.bold = True; run.font.size = Pt(9)
        run.font.color.rgb = hex_to_rgb(WHITE)
        run.font.name = "Calibri"
    for ri, row_vals in enumerate(rows_data):
        bg = LIGHT_BLUE if ri % 2 == 0 else ALT_WHITE
        row = tbl.rows[ri+1]
        for ci, val in enumerate(row_vals):
            cell = row.cells[ci]
            set_cell_bg(cell, bg)
            set_cell_margins(cell)
            p = cell.paragraphs[0]; no_space(p)
            run = p.add_run(val)
            run.font.size = Pt(9)
            run.font.color.rgb = hex_to_rgb(NEAR_BLACK)
            run.font.name = "Calibri"
    doc.add_paragraph()

def pain_point_table(doc, pain_points):
    tbl = doc.add_table(rows=1+len(pain_points), cols=3)
    tbl.style = "Table Grid"
    headers = ["Pain Point", "What They Feel", "What They Need"]
    hdr_row = tbl.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        set_cell_bg(cell, BLUE)
        set_cell_margins(cell)
        p = cell.paragraphs[0]; no_space(p)
        run = p.add_run(h)
        run.bold = True; run.font.size = Pt(9)
        run.font.color.rgb = hex_to_rgb(WHITE)
        run.font.name = "Calibri"
    for ri, (pp, feel, need) in enumerate(pain_points):
        bg = LIGHT_BLUE if ri % 2 == 0 else ALT_WHITE
        row = tbl.rows[ri+1]
        for ci, val in enumerate([pp, feel, need]):
            cell = row.cells[ci]
            set_cell_bg(cell, bg)
            set_cell_margins(cell)
            p = cell.paragraphs[0]; no_space(p)
            run = p.add_run(val)
            run.font.size = Pt(9)
            run.font.color.rgb = hex_to_rgb(NEAR_BLACK)
            run.font.name = "Calibri"
    doc.add_paragraph()

def heading2(doc, text):
    p = doc.add_paragraph()
    no_space(p)
    pPr = p._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "120")
    spacing.set(qn("w:after"),  "60")
    pPr.append(spacing)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = hex_to_rgb(DARK_GREY)
    run.font.name = "Calibri"

def gap_block(doc, gap, opportunity, angle):
    tbl = doc.add_table(rows=3, cols=2)
    tbl.style = "Table Grid"
    labels = ["Content Gap", "Opportunity", "Our Angle"]
    values = [gap, opportunity, angle]
    for i, (label, value) in enumerate(zip(labels, values)):
        bg = LIGHT_BLUE if i % 2 == 0 else ALT_WHITE
        row = tbl.rows[i]
        set_cell_bg(row.cells[0], bg); set_cell_bg(row.cells[1], bg)
        set_cell_margins(row.cells[0]); set_cell_margins(row.cells[1])
        p0 = row.cells[0].paragraphs[0]; no_space(p0)
        r0 = p0.add_run(label); r0.bold = True
        r0.font.size = Pt(9); r0.font.color.rgb = hex_to_rgb(DARK_GREY)
        r0.font.name = "Calibri"
        p1 = row.cells[1].paragraphs[0]; no_space(p1)
        r1 = p1.add_run(value)
        r1.font.size = Pt(9); r1.font.color.rgb = hex_to_rgb(NEAR_BLACK)
        r1.font.name = "Calibri"
    doc.add_paragraph()

def claim_block(doc, claim_num, claim, source, url=""):
    tbl = doc.add_table(rows=3, cols=2)
    tbl.style = "Table Grid"
    rows_data = [
        ("Claim", claim),
        ("Source", source),
        ("URL", url),
    ]
    for i, (label, value) in enumerate(rows_data):
        bg = LIGHT_BLUE if i % 2 == 0 else ALT_WHITE
        row = tbl.rows[i]
        set_cell_bg(row.cells[0], bg); set_cell_bg(row.cells[1], bg)
        set_cell_margins(row.cells[0]); set_cell_margins(row.cells[1])
        p0 = row.cells[0].paragraphs[0]; no_space(p0)
        r0 = p0.add_run(label); r0.bold = True
        r0.font.size = Pt(9); r0.font.color.rgb = hex_to_rgb(DARK_GREY)
        r0.font.name = "Calibri"
        p1 = row.cells[1].paragraphs[0]; no_space(p1)
        r1 = p1.add_run(value)
        r1.font.size = Pt(9); r1.font.color.rgb = hex_to_rgb(NEAR_BLACK)
        r1.font.name = "Calibri"
    doc.add_paragraph()

def stat_bullet(doc, stat, source, size=9):
    p = doc.add_paragraph()
    no_space(p)
    pPr = p._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"),  "80")
    pPr.append(spacing)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"),    "240")
    ind.set(qn("w:hanging"), "240")
    pPr.append(ind)
    r1 = p.add_run("▸  ")
    r1.font.size = Pt(size); r1.font.color.rgb = hex_to_rgb(BLUE)
    r1.font.name = "Calibri"
    r2 = p.add_run(stat + " ")
    r2.bold = True; r2.font.size = Pt(size)
    r2.font.color.rgb = hex_to_rgb(NEAR_BLACK)
    r2.font.name = "Calibri"
    r3 = p.add_run(f"({source})")
    r3.italic = True; r3.font.size = Pt(size - 0.5)
    r3.font.color.rgb = hex_to_rgb(DARK_GREY)
    r3.font.name = "Calibri"

# ── BUILD ──────────────────────────────────────────────────────────────────────

doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin    = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin   = Inches(0.85)
    section.right_margin  = Inches(0.85)

# ── TITLE ──────────────────────────────────────────────────────────────────────
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
no_space(title_p)
pPr = title_p._p.get_or_add_pPr()
sp  = OxmlElement("w:spacing"); sp.set(qn("w:after"), "160"); pPr.append(sp)
r   = title_p.add_run("CONTENT BRIEF: How Niching Your Social Media Agency Makes You More Money")
r.bold = True; r.font.size = Pt(15)
r.font.color.rgb = hex_to_rgb(BLUE)
r.font.name = "Calibri"

# ── 1. BASICS ──────────────────────────────────────────────────────────────────
section_banner(doc, "1. BASICS")
kv_table(doc, [
    ("Topic / Working Title", "How Niching Your Social Media Agency Makes You More Money"),
    ("Content Type",          "Long-form Educational Blog Post"),
    ("Funnel Stage",          "Top of Funnel (ToFU) — Awareness & Consideration"),
    ("Primary Goal",          "Educate agency owners on the financial & operational case for specialisation; drive organic traffic from agency decision-makers researching growth strategy"),
    ("Target Word Count",     "3,000–3,600 words"),
    ("Author Brief",          "Written for a knowledgeable content writer; no fluff — data-led narrative with specific benchmarks, named experts, and actionable frameworks"),
    ("Publish Destination",   "SocialPilot Blog"),
    ("Assigned Writer",       "[TBD]"),
    ("Due Date",              "[TBD]"),
])

# ── 2. SEO TARGETS ─────────────────────────────────────────────────────────────
section_banner(doc, "2. SEO TARGETS")

heading2(doc, "Primary Keyword")
data_table(doc,
    ["Keyword", "Est. Monthly Volume", "Difficulty", "Intent"],
    [
        ("niche social media agency", "800–1,200", "Medium", "Informational / Commercial"),
        ("social media agency niche", "600–900",   "Medium", "Informational"),
    ]
)

heading2(doc, "Secondary Keywords")
sec_kws = [
    ("how to niche down as an agency",             "300–500",   "Low–Medium", "Informational"),
    ("social media agency specialization",         "400–600",   "Medium",     "Informational"),
    ("most profitable social media agency niche",  "200–400",   "Low",        "Informational / Commercial"),
    ("vertical social media agency",               "150–300",   "Low",        "Informational"),
    ("agency niche selection framework",           "100–200",   "Low",        "Informational"),
    ("niche agency vs generalist agency",          "200–350",   "Low–Medium", "Informational"),
    ("how to grow a social media agency",          "500–800",   "Medium",     "Informational"),
    ("social media agency profit margin",          "150–250",   "Low",        "Informational"),
    ("best niche for social media agency",         "250–400",   "Low",        "Commercial"),
    ("agency specialization revenue",             "100–200",   "Low",        "Informational"),
    ("social media agency pricing strategy",      "300–500",   "Medium",     "Informational / Commercial"),
    ("how to position a social media agency",     "150–300",   "Low",        "Informational"),
]
data_table(doc,
    ["Keyword", "Est. Monthly Volume", "Difficulty", "Intent"],
    sec_kws
)

heading2(doc, "People Also Ask / PAA Questions to Answer")
paa = [
    "How do I choose a niche for my social media agency?",
    "Is it better to be a generalist or specialist agency?",
    "What are the most profitable niches for a social media agency?",
    "How much more do niche agencies charge than generalists?",
    "How do niche agencies get clients?",
    "What is the average profit margin for a social media agency?",
    "How do you grow a social media agency fast?",
    "What industries pay the most for social media management?",
    "How do I differentiate my social media agency?",
    "What is vertical specialisation in a marketing agency?",
    "Can a small agency compete against large agencies by niching?",
    "What are the risks of niching your agency?",
]
for q in paa:
    bullet(doc, q)
doc.add_paragraph()

heading2(doc, "Meta Title (≤ 60 chars)")
body_para(doc, "How Niching Your Social Media Agency Makes You More Money")
heading2(doc, "Meta Description (≤ 160 chars)")
body_para(doc, "Niche agencies charge 20–50% more, close 40% more proposals, and retain clients longer. Here's the data-backed case for specialising your social media agency.")
doc.add_paragraph()

# ── 3. AUDIENCE ────────────────────────────────────────────────────────────────
section_banner(doc, "3. TARGET AUDIENCE & ICP")

heading2(doc, "Primary Reader")
kv_table(doc, [
    ("Role",                "Founder / Owner of a social media or digital marketing agency"),
    ("Agency Size",         "2–50 employees; $200K–$5M annual revenue"),
    ("Stage",               "Growing but plateauing — frustrated with inconsistent deal flow and price pressure from larger agencies"),
    ("Pain State",          "Juggling too many client types, underpriced, losing pitches to specialists who 'get' the client's industry"),
    ("Goal",                "Scale revenue, increase profit margins, reduce sales cycle length, win better-fit clients"),
    ("Platform Context",    "Likely uses SocialPilot or similar tools; reads agency growth content; follows Jason Swenk, Karl Sakas, Blair Enns"),
])

heading2(doc, "Secondary Reader")
body_para(doc, "Freelance social media manager transitioning to agency model and evaluating whether to go broad or narrow from the start.")
doc.add_paragraph()

heading2(doc, "Reader Mindset")
bullet(doc, "Aware that 'doing everything for everyone' is exhausting but afraid that niching means turning away money")
bullet(doc, "Has probably heard 'niche down' advice but hasn't seen a data-backed case for why")
bullet(doc, "Wants a practical framework — not just inspiration — to choose and validate a niche")
bullet(doc, "Concerned about market concentration risk ('what if the industry tanks?')")
doc.add_paragraph()

# ── 4. SERP & COMPETITIVE ANALYSIS ────────────────────────────────────────────
section_banner(doc, "4. SERP & COMPETITIVE ANALYSIS")

heading2(doc, "Current SERP Landscape")
body_para(doc, "The SERP for 'niche social media agency' and related queries is dominated by listicle-style posts ('best niches for agencies'), shallow how-to posts from agency tool vendors, and generic business advice articles. Most results lack primary benchmark data or expert citations. There is a clear gap for a comprehensive, data-led piece that combines financial benchmarks, named case studies, and an actionable niche-selection framework.")
doc.add_paragraph()

heading2(doc, "Top Competing Content Pieces")
data_table(doc,
    ["Source", "Title / Angle", "Weakness / Gap"],
    [
        ("Agency Analytics",      "'How to Pick Your Agency Niche' — framework-focused",               "Lacks revenue/margin benchmarks; no expert citations"),
        ("AgencyAnalytics Blog",  "'Why Niche Agencies Excel: 16 Experts' — quote-heavy",             "Heavy on opinion; light on verified statistics"),
        ("Phantom Leads AI",      "'Should Your Agency Niche Down? The Data Says Yes'",               "Good data but thin content; no operational efficiency angle"),
        ("ALM Corp Blog",         "'Most Profitable Niches for Digital Marketing Agencies 2026'",     "Niche listing only; no how-to framework or case studies"),
        ("InMotion Hosting Blog", "'Generalist vs. Niche Agency: Which Model Grows Faster'",         "Balanced but shallow; no actionable selection process"),
    ]
)

heading2(doc, "Content Gap & Our Angle")
gap_block(doc,
    gap="No authoritative piece combines: (a) financial benchmarks from primary reports, (b) expert voices (Blair Enns, Karl Sakas, Jason Swenk, David C. Baker), (c) named case studies, AND (d) an actionable niche-selection framework in a single article.",
    opportunity="SocialPilot can own the definitive data-backed guide to agency niching — bridging the 'why' (financial case) with the 'how' (framework, validation, risk mitigation) in a way no competitor currently does.",
    angle="Lead with the financial case (margins, rates, retention), validate with named case studies and expert quotes, and close with a practical 3-step niche-selection framework that agency owners can apply this week."
)

# ── 5. LLM CONSIDERATIONS ──────────────────────────────────────────────────────
section_banner(doc, "5. LLM & AI SEARCH CONSIDERATIONS")
body_para(doc, "With AI-assisted search growing, this article must be structured to be extractable by LLMs answering queries like 'how do I choose a niche for my social media agency?' or 'are niche agencies more profitable?'. Prioritise:", bold=False)
doc.add_paragraph()
bullet(doc, "Clear H2/H3 hierarchy with question-format headers where appropriate (e.g. 'Why Do Niche Agencies Charge More?')")
bullet(doc, "Data tables for benchmark comparisons (niche vs. generalist margins, pricing premiums by vertical)")
bullet(doc, "Concise definition blocks for key terms (vertical specialisation, service specialisation, niche agency)")
bullet(doc, "Quotable expert sentences attributed to named individuals (Blair Enns, Jason Swenk, Karl Sakas, David C. Baker)")
bullet(doc, "Schema-friendly structure: FAQ section at end with 4–6 questions and direct answers")
doc.add_paragraph()

# ── 6. KEY REQUIREMENTS ────────────────────────────────────────────────────────
section_banner(doc, "6. KEY REQUIREMENTS & MANDATORY INCLUSIONS")

req_data = [
    ("REQ 1", "Financial Benchmarks Table",
     "Include a comparison table: Niche Agency vs. Generalist Agency across Gross Margin, Net Margin, Average Retainer Size, Client Retention Rate, and Revenue Per Employee. Source: Agency Analytics 2024, Predictable Profits 2025, Promethean Research 2025."),
    ("REQ 2", "Pricing Premium by Vertical",
     "Include a table showing specialist rate premiums by niche: Healthcare (25–32% margins), Legal ($10K–$50K+ monthly), B2B SaaS (28–35% margins), E-commerce, Fitness/Wellness. Cite ALM Corp 2026 data."),
    ("REQ 3", "Named Case Study: Rankings.io",
     "Feature Rankings.io growing from $2M to $12M revenue in 4 years through niche specialisation (Karl Sakas case study). This is the strongest real-world proof point available."),
    ("REQ 4", "Expert Quotes (Attributed & Named)",
     "Include direct quotes or paraphrased positions from: Blair Enns (Win Without Pitching) on pricing power, Karl Sakas on revenue-per-employee targets ($250K+ for specialists vs. $180K+ for generalists), Jason Swenk on quarterly client pattern review, David C. Baker on expertise emerging from past work."),
    ("REQ 5", "Niche Selection Framework",
     "Present a 3-part framework — Expertise × Market Demand × Passion — with Karl Sakas' Goldilocks market-sizing rule: 2,000–10,000 prospects, 10–200 competitors. Include a validation step (quiet pilot before full commitment)."),
    ("REQ 6", "Referral Network & CAC Reduction Data",
     "State that referrals account for 34.2% of agency client discovery (Alex Berman data), and that niche agencies reduce CAC by 30–50% vs. generalists through industry-specific word-of-mouth networks."),
    ("REQ 7", "Risks Section (Honest & Balanced)",
     "Cover market concentration risk, recession vulnerability, and scaling constraints. Include 'When NOT to niche' guidance: insufficient expertise, small geographic markets, or personal burnout risk."),
    ("REQ 8", "Service Reduction ≠ Fewer Clients",
     "Cite Promethean Research 2025 finding: agencies that REDUCED services posted 30% net margins in 2025 vs. 10% for agencies that EXPANDED services. Counter the instinct that offering more services grows revenue."),
    ("REQ 9", "SocialPilot Natural Integration",
     "Tie in SocialPilot's white-label platform and multi-client management tools as infrastructure that enables niche agencies to systemise delivery across 25–100+ clients in a single vertical — without custom development. Natural, not forced."),
]

for req_id, req_title, req_desc in req_data:
    heading2(doc, f"{req_id}: {req_title}")
    body_para(doc, req_desc)
    doc.add_paragraph()

# ── 7. CITABLE CLAIMS ─────────────────────────────────────────────────────────
section_banner(doc, "7. CITABLE CLAIMS (Use Verbatim or Closely Paraphrased)")

claims = [
    ("Claim 1",
     "Agencies that reduced their service offerings in 2025 averaged 30% net profit margins, compared to just 10% for agencies that expanded their service mix.",
     "Promethean Research 2025 Digital Agency Industry Report",
     "https://prometheanresearch.com/2025-digital-agency-industry-report/"),
    ("Claim 2",
     "Niche agencies charge 2–3x more than generalists, close 40% more proposals, and grow 3x faster in their first three years.",
     "Phantom Leads AI — Should Your Agency Niche Down? The Data Says Yes",
     "https://phantomleads.ai/blog/niche-down-or-stay-general"),
    ("Claim 3",
     "Top-performing 8-figure agencies retain 92% of clients annually and maintain 25–32% profit margins — compared to 78% retention and 18–22% margins for 7-figure agencies.",
     "Predictable Profits 2025 Agency Growth Benchmark (300+ agencies surveyed)",
     "https://predictableprofits.com/2025-agency-growth-benchmark-key-metrics-from-300-7-8-figure-agencies/"),
    ("Claim 4",
     "Referrals account for 34.2% of how clients discover agencies — making it the single most effective acquisition channel — and niche agencies reduce client acquisition cost by 30–50% through industry-specific word-of-mouth.",
     "Alex Berman / Digital Marketing Agency Niche research; Talkable referral marketing benchmarks",
     "https://alexberman.com/digital-marketing-agency-niche"),
    ("Claim 5",
     "Karl Sakas' target benchmark: specialist agencies should achieve $250,000+ revenue per employee, vs. $180,000+ for generalist agencies.",
     "Sakas & Company — Agency Specialization Research",
     "https://sakasandcompany.com/agency-specialization/"),
    ("Claim 6",
     "Rankings.io grew from $2 million to $12 million in revenue in four years while simultaneously reducing client workload — by committing to a single niche.",
     "Sakas & Company Case Study — Rankings.io",
     "https://sakasandcompany.com/case-studies/rankings-io-agency-growth/"),
]
for num, claim, source, url in claims:
    heading2(doc, num)
    claim_block(doc, num, claim, source, url)

# ── 8. PAIN POINT ANALYSIS ────────────────────────────────────────────────────
section_banner(doc, "8. PAIN POINT ANALYSIS")
pain_points = [
    ("Racing to the bottom on price",
     "Constantly losing deals to cheaper competitors or agencies that 'specialise' in the prospect's industry",
     "A credible, data-backed reason why specialisation commands premium pricing — not just positioning theory"),
    ("Unpredictable revenue & pipeline",
     "No clear ideal client profile means marketing to everyone and converting no one; proposal win rates are low",
     "Proof that a defined niche shortens sales cycles, improves win rates, and generates referral-driven inbound"),
    ("Operational chaos at scale",
     "Every new client is different; no reusable systems; team constantly reinventing the wheel",
     "Concrete evidence that vertical focus enables SOPs, faster onboarding, and higher utilisation rates"),
    ("Fear of leaving money on the table",
     "Turning away clients outside the niche feels like lost revenue — the 'what if we're missing out?' anxiety",
     "Reframing: the revenue lost from rejected clients is recovered through higher rates, lower CAC, and better retention in the niche"),
    ("Not knowing which niche to pick",
     "Overwhelmed by options; afraid of picking the wrong vertical and being stuck with it",
     "A practical, low-risk validation framework — test quietly before fully committing"),
    ("Competing against larger agencies",
     "Feels impossible to win against full-service agencies with bigger teams and brand recognition",
     "Data showing that specialists consistently beat generalists in niche-specific RFPs, even against much larger competitors"),
]
pain_point_table(doc, pain_points)

# ── 9. KEY STATS FOR WRITER ───────────────────────────────────────────────────
section_banner(doc, "9. KEY STATS FOR WRITER")

stat_bullet(doc,
    "Niche agencies achieve gross margins of 40–75% vs. ~20% for generalist agencies.",
    "Predictable Profits 2025 / Agency Analytics 2024")
stat_bullet(doc,
    "Agencies that reduced services in 2025 averaged 30% net margins vs. 10% for those that expanded.",
    "Promethean Research 2025 Digital Agency Industry Report")
stat_bullet(doc,
    "Specialist agencies command 20–50% higher billing rates than generalist counterparts.",
    "Agency Analytics 2024 Benchmarks Report")
stat_bullet(doc,
    "Niche agencies close 40% more proposals and grow 3x faster in their first 3 years.",
    "Phantom Leads AI, 2025")
stat_bullet(doc,
    "8-figure agencies retain 92% of clients annually vs. 78% for 7-figure agencies.",
    "Predictable Profits 2025 (300+ agencies surveyed)")
stat_bullet(doc,
    "Referrals drive 34.2% of new agency client discovery — the top acquisition channel.",
    "Alex Berman, Digital Marketing Agency Niche research")
stat_bullet(doc,
    "Niche specialisation reduces client acquisition cost by 30–50% through industry word-of-mouth.",
    "ALM Corp, Most Profitable Niches 2026")
stat_bullet(doc,
    "84% of digital agencies now self-identify as specialists, signalling broad-based generalist decline.",
    "Phantom Leads AI, 2025")
stat_bullet(doc,
    "Rankings.io grew from $2M to $12M in 4 years through niche specialisation, per Karl Sakas case study.",
    "Sakas & Company")
stat_bullet(doc,
    "Legal marketing specialists charge $10,000–$50,000+ monthly for comprehensive programmes.",
    "ALM Corp, Most Profitable Niches 2026")
stat_bullet(doc,
    "Senior specialist consultants earn $300–$500/hour vs. $150–$200 for generalist consultants.",
    "ConsultFees.com 2026 Benchmarks")
stat_bullet(doc,
    "Karl Sakas' target: $250,000+ revenue per employee for specialist agencies vs. $180,000+ for generalists.",
    "Sakas & Company — Agency Specialization")
doc.add_paragraph()

# ── 10. KEY SOURCES FOR WRITER ────────────────────────────────────────────────
section_banner(doc, "10. KEY SOURCES FOR WRITER")
body_para(doc, "The writer must read and cite the following sources. Do not fabricate statistics — only use data that can be verified at these URLs.", bold=True)
doc.add_paragraph()

sources = [
    ("Predictable Profits 2025 Agency Growth Benchmark",
     "https://predictableprofits.com/2025-agency-growth-benchmark-key-metrics-from-300-7-8-figure-agencies/"),
    ("Promethean Research 2025 Digital Agency Industry Report",
     "https://prometheanresearch.com/2025-digital-agency-industry-report/"),
    ("Agency Analytics 2024 Marketing Agency Benchmarks Report",
     "https://agencyanalytics.com/company/newsroom/2024-benchmarks-report-press-release"),
    ("Sakas & Company — Rankings.io Case Study ($2M to $12M)",
     "https://sakasandcompany.com/case-studies/rankings-io-agency-growth/"),
    ("Sakas & Company — Agency Specialization Research",
     "https://sakasandcompany.com/agency-specialization/"),
    ("ALM Corp — Most Profitable Niches for Digital Marketing Agencies 2026",
     "https://almcorp.com/blog/most-profitable-niches-digital-marketing-agencies-2026/"),
    ("Phantom Leads AI — Should Your Agency Niche Down? The Data Says Yes",
     "https://phantomleads.ai/blog/niche-down-or-stay-general"),
    ("Agency Analytics — How to Pick Your Agency Niche",
     "https://agencyanalytics.com/blog/how-to-pick-a-niche"),
    ("Alex Berman — Digital Marketing Agency Niche (Referral Data)",
     "https://alexberman.com/digital-marketing-agency-niche"),
    ("ConsultFees.com — Hourly Consulting Rate Benchmarks 2026",
     "https://consultfees.com/blog/hourly-consulting-rate"),
]
data_table(doc,
    ["Source", "URL"],
    sources
)

# ── 11. ARTICLE STRUCTURE / WRITER CHECKLIST ──────────────────────────────────
section_banner(doc, "11. ARTICLE STRUCTURE & WRITER CHECKLIST")

checklist_sections = [
    ("INTRO (200–300 words)",
     [
         "Open with a scenario: agency owner losing a pitch to a smaller specialist firm that charged 40% more",
         "State the core paradox: 'Serving fewer clients — and fewer types of clients — is the counterintuitive path to a bigger, more profitable agency'",
         "Promise: financial benchmarks, a named case study, expert-backed framework",
         "Include primary keyword naturally in first 100 words",
     ]),
    ("SECTION 1: The Financial Case for Niching (500–600 words)",
     [
         "Lead with Promethean Research 2025 service-reduction finding (30% vs. 10% margins)",
         "Include niche vs. generalist comparison table: margins, rates, retention, revenue per employee",
         "Cite Agency Analytics 2024 and Predictable Profits 2025 benchmarks",
         "Cover the specialist pricing premium: 20–50% higher rates in healthcare, legal, B2B SaaS",
         "Mention 8-figure vs. 7-figure agency retention gap (92% vs. 78%)",
     ]),
    ("SECTION 2: Why Niche Agencies Win More Business (400–500 words)",
     [
         "Cover the referral network advantage: 34.2% of clients find agencies via referrals",
         "Explain how industry-specific word-of-mouth creates compounding growth",
         "Address the shorter sales cycle: specialists pre-qualify; prospects self-select",
         "Close deal rate: niche agencies close 40% more proposals",
         "Quote Blair Enns on the three benefits of specialisation (sales advantage, pricing premium, client power)",
     ]),
    ("SECTION 3: The Operational Efficiency Dividend (300–400 words)",
     [
         "Explain how vertical focus enables reusable SOPs, content templates, reporting dashboards",
         "Cover faster client onboarding when the agency already knows the industry",
         "Mention SocialPilot's white-label platform as infrastructure for systemised multi-client delivery",
         "Reference Databox agency efficiency finding: 92% of high-margin agencies use PM and time-tracking tools",
         "Quote Jason Swenk: identify patterns in profitable clients quarterly, then focus marketing on similar prospects",
     ]),
    ("SECTION 4: Real-World Proof — The Rankings.io Story (300–400 words)",
     [
         "Feature Rankings.io: SEO agency that niched to personal injury law firms",
         "Grew from $2M to $12M in 4 years (Karl Sakas case study)",
         "Explain the mechanics: premium pricing, referral network, operational systemisation",
         "Briefly mention 2–3 other named niche agencies: Cardinal Digital Marketing (healthcare), Gripped.io (B2B SaaS), CSTMR (fintech)",
         "Show range of verticals to help readers identify with at least one",
     ]),
    ("SECTION 5: How to Choose Your Niche — The 3-Part Framework (400–500 words)",
     [
         "Framework: Expertise × Market Demand × Passion — need all three",
         "Expertise: What have your 3–5 most successful clients had in common? Start there",
         "Market Demand: Use Karl Sakas' Goldilocks rule — 2,000–10,000 prospects, 10–200 competitors",
         "Passion: A motivated generalist outperforms a burned-out specialist — don't ignore this",
         "Quote David C. Baker: 'Expertise is not invented but emerges from what you've already done'",
         "Include validation step: narrow pipeline quietly for 90 days before full commitment",
     ]),
    ("SECTION 6: The Risks Are Real — Here's How to Manage Them (250–300 words)",
     [
         "Acknowledge market concentration risk honestly — don't dismiss it",
         "Cover recession vulnerability (single-vertical dependency)",
         "When NOT to niche: insufficient portfolio, very small local market, personal burnout risk",
         "Mitigation: choose verticals with regulatory complexity (creates barriers to entry for competitors)",
         "Mitigation: validate market depth before committing (2,000+ prospects minimum)",
     ]),
    ("SECTION 7: The Most Profitable Niches Right Now (300–400 words)",
     [
         "Healthcare / medical: 25–32% margins; pharma digital spend accelerating",
         "Legal: $10K–$50K+ monthly retainers; law firms in major markets spending $360K+/month",
         "B2B SaaS: 28–35% margins; AI and automation driving explosive growth",
         "E-commerce / DTC: US market projected 53% growth by 2027",
         "Fitness & wellness: $8.5 trillion global market",
         "Home services: underserved by specialists; strong referral networks",
         "Frame as a starting point for validation, not a guaranteed gold rush",
     ]),
    ("CONCLUSION & CTA (200–250 words)",
     [
         "Restate the core case: niching is not about turning away money — it's about charging more for the money you keep",
         "Summarise: higher margins, lower CAC, better retention, shorter sales cycles",
         "Call to action: Download SocialPilot's agency growth resources / start a free trial of SocialPilot for multi-client management",
         "FAQ section (4–5 questions with concise answers for schema markup)",
     ]),
]

for section_title, checks in checklist_sections:
    sub_banner(doc, section_title)
    for item in checks:
        checkbox(doc, item)
    doc.add_paragraph()

# ── 12. INTERNAL LINKS ────────────────────────────────────────────────────────
section_banner(doc, "12. INTERNAL LINKS (Suggested)")
links = [
    ("How to Get Clients for a Social Media Agency",     "SocialPilot blog — agency client acquisition"),
    ("How to Price Your Social Media Management Services", "SocialPilot blog — agency pricing"),
    ("Social Media Management for Agencies",             "SocialPilot solutions page — agencies"),
    ("White Label Social Media Management",              "SocialPilot white-label product page"),
    ("Social Media Marketing Agency Software",           "SocialPilot features for agencies"),
    ("How to Scale a Social Media Agency",               "SocialPilot blog — agency growth"),
    ("Social Media Reporting for Agencies",              "SocialPilot blog — client reporting"),
]
data_table(doc,
    ["Suggested Anchor Text", "Target Page / Context"],
    links
)

# ── SAVE ───────────────────────────────────────────────────────────────────────
out = "/home/user/new/CONTENT_BRIEF__SocialPilot__Niche_Agency_v1.docx"
doc.save(out)
print(f"Saved: {out}")
