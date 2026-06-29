"""
Content Brief v2: How Niching Your Social Media Agency Makes You More Money
SocialPilot — deep research rebuild
"""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BLUE       = "1A56DB"
WHITE      = "FFFFFF"
NEAR_BLACK = "141413"
DARK_GREY  = "3D3D3A"
LIGHT_BLUE = "EEF2FF"
ALT_WHITE  = "FFFFFF"

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

def body_para(doc, text, bold=False, size=9, color=NEAR_BLACK, after=60):
    p = doc.add_paragraph()
    no_space(p)
    pPr = p._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"),  str(after))
    pPr.append(spacing)
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
    r1.bold = True; r1.font.size = Pt(size)
    r1.font.color.rgb = hex_to_rgb(DARK_GREY)
    r1.font.name = "Calibri"
    r2 = p.add_run(value)
    r2.font.size = Pt(size)
    r2.font.color.rgb = hex_to_rgb(NEAR_BLACK)
    r2.font.name = "Calibri"

def bullet(doc, text, bold_prefix=None, size=9):
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
    num = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl"); ilvl.set(qn("w:val"), "0"); num.append(ilvl)
    numId= OxmlElement("w:numId"); numId.set(qn("w:val"), "1"); num.append(numId)
    pPr.append(num)
    if bold_prefix:
        r1 = p.add_run(bold_prefix + " ")
        r1.bold = True; r1.font.size = Pt(size)
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
    for i, (label, value) in enumerate(rows_data):
        bg = LIGHT_BLUE if i % 2 == 0 else ALT_WHITE
        c0, c1 = tbl.rows[i].cells[0], tbl.rows[i].cells[1]
        set_cell_bg(c0, bg); set_cell_bg(c1, bg)
        set_cell_margins(c0); set_cell_margins(c1)
        p0 = c0.paragraphs[0]; no_space(p0)
        r0 = p0.add_run(label); r0.bold = True
        r0.font.size = Pt(9); r0.font.color.rgb = hex_to_rgb(DARK_GREY); r0.font.name = "Calibri"
        p1 = c1.paragraphs[0]; no_space(p1)
        r1 = p1.add_run(value)
        r1.font.size = Pt(9); r1.font.color.rgb = hex_to_rgb(NEAR_BLACK); r1.font.name = "Calibri"
    doc.add_paragraph()

def data_table(doc, headers, rows_data):
    col_count = len(headers)
    tbl = doc.add_table(rows=1+len(rows_data), cols=col_count)
    tbl.style = "Table Grid"
    hdr_row = tbl.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        set_cell_bg(cell, BLUE); set_cell_margins(cell)
        p = cell.paragraphs[0]; no_space(p)
        run = p.add_run(h); run.bold = True
        run.font.size = Pt(9); run.font.color.rgb = hex_to_rgb(WHITE); run.font.name = "Calibri"
    for ri, row_vals in enumerate(rows_data):
        bg = LIGHT_BLUE if ri % 2 == 0 else ALT_WHITE
        row = tbl.rows[ri+1]
        for ci, val in enumerate(row_vals):
            cell = row.cells[ci]
            set_cell_bg(cell, bg); set_cell_margins(cell)
            p = cell.paragraphs[0]; no_space(p)
            run = p.add_run(val)
            run.font.size = Pt(9); run.font.color.rgb = hex_to_rgb(NEAR_BLACK); run.font.name = "Calibri"
    doc.add_paragraph()

def pain_point_table(doc, pain_points):
    tbl = doc.add_table(rows=1+len(pain_points), cols=3)
    tbl.style = "Table Grid"
    headers = ["Pain Point", "What They Feel", "What They Need"]
    hdr_row = tbl.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        set_cell_bg(cell, BLUE); set_cell_margins(cell)
        p = cell.paragraphs[0]; no_space(p)
        run = p.add_run(h); run.bold = True
        run.font.size = Pt(9); run.font.color.rgb = hex_to_rgb(WHITE); run.font.name = "Calibri"
    for ri, (pp, feel, need) in enumerate(pain_points):
        bg = LIGHT_BLUE if ri % 2 == 0 else ALT_WHITE
        row = tbl.rows[ri+1]
        for ci, val in enumerate([pp, feel, need]):
            cell = row.cells[ci]
            set_cell_bg(cell, bg); set_cell_margins(cell)
            p = cell.paragraphs[0]; no_space(p)
            run = p.add_run(val)
            run.font.size = Pt(9); run.font.color.rgb = hex_to_rgb(NEAR_BLACK); run.font.name = "Calibri"
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
    run.bold = True; run.font.size = Pt(10)
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
        r0.font.size = Pt(9); r0.font.color.rgb = hex_to_rgb(DARK_GREY); r0.font.name = "Calibri"
        p1 = row.cells[1].paragraphs[0]; no_space(p1)
        r1 = p1.add_run(value)
        r1.font.size = Pt(9); r1.font.color.rgb = hex_to_rgb(NEAR_BLACK); r1.font.name = "Calibri"
    doc.add_paragraph()

def claim_block(doc, num, claim, source, url=""):
    rows_data = [("Claim", claim), ("Source", source), ("URL", url)]
    tbl = doc.add_table(rows=3, cols=2)
    tbl.style = "Table Grid"
    for i, (label, value) in enumerate(rows_data):
        bg = LIGHT_BLUE if i % 2 == 0 else ALT_WHITE
        row = tbl.rows[i]
        set_cell_bg(row.cells[0], bg); set_cell_bg(row.cells[1], bg)
        set_cell_margins(row.cells[0]); set_cell_margins(row.cells[1])
        p0 = row.cells[0].paragraphs[0]; no_space(p0)
        r0 = p0.add_run(label); r0.bold = True
        r0.font.size = Pt(9); r0.font.color.rgb = hex_to_rgb(DARK_GREY); r0.font.name = "Calibri"
        p1 = row.cells[1].paragraphs[0]; no_space(p1)
        r1 = p1.add_run(value)
        r1.font.size = Pt(9); r1.font.color.rgb = hex_to_rgb(NEAR_BLACK); r1.font.name = "Calibri"
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
    r1.font.size = Pt(size); r1.font.color.rgb = hex_to_rgb(BLUE); r1.font.name = "Calibri"
    r2 = p.add_run(stat + " ")
    r2.bold = True; r2.font.size = Pt(size)
    r2.font.color.rgb = hex_to_rgb(NEAR_BLACK); r2.font.name = "Calibri"
    r3 = p.add_run(f"({source})")
    r3.italic = True; r3.font.size = Pt(size - 0.5)
    r3.font.color.rgb = hex_to_rgb(DARK_GREY); r3.font.name = "Calibri"

# ═══════════════════════════════════════════════════════════════════════════════
# BUILD DOCUMENT
# ═══════════════════════════════════════════════════════════════════════════════
doc = Document()

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
sp = OxmlElement("w:spacing"); sp.set(qn("w:after"), "160"); pPr.append(sp)
r  = title_p.add_run("CONTENT BRIEF: How Niching Your Social Media Agency Makes You More Money")
r.bold = True; r.font.size = Pt(15)
r.font.color.rgb = hex_to_rgb(BLUE); r.font.name = "Calibri"

# ── 1. BASICS ──────────────────────────────────────────────────────────────────
section_banner(doc, "1. BASICS")
kv_table(doc, [
    ("Topic / Working Title", "How Niching Your Social Media Agency Makes You More Money"),
    ("Content Type",          "Long-form Educational Blog Post"),
    ("Funnel Stage",          "Top of Funnel (ToFU) — Awareness & Consideration"),
    ("Primary Goal",          "Convince agency owners (with hard data) that specialisation is the primary lever for higher margins, better clients, and faster growth — and give them a framework to choose and validate a niche"),
    ("Target Word Count",     "3,200–3,800 words"),
    ("Author Note",           "Data-led narrative throughout. Every stat must cite a named source. No generic advice — specific benchmarks, named agencies, named expert quotes only."),
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
        ("niche social media agency",  "800–1,200", "Medium", "Informational / Commercial"),
        ("social media agency niche",  "600–900",   "Medium", "Informational"),
    ]
)

heading2(doc, "Secondary Keywords")
data_table(doc,
    ["Keyword", "Est. Monthly Volume", "Difficulty", "Intent"],
    [
        ("how to niche down as an agency",            "300–500", "Low–Medium", "Informational"),
        ("social media agency specialization",        "400–600", "Medium",     "Informational"),
        ("most profitable social media agency niche", "200–400", "Low",        "Informational / Commercial"),
        ("vertical social media agency",              "150–300", "Low",        "Informational"),
        ("niche agency vs generalist agency",         "200–350", "Low–Medium", "Informational"),
        ("best niche for social media agency",        "250–400", "Low",        "Commercial"),
        ("how to grow a social media agency",         "500–800", "Medium",     "Informational"),
        ("social media agency profit margin",         "150–250", "Low",        "Informational"),
        ("agency niche selection framework",          "100–200", "Low",        "Informational"),
        ("social media agency pricing strategy",      "300–500", "Medium",     "Informational / Commercial"),
        ("how to position a social media agency",     "150–300", "Low",        "Informational"),
        ("agency specialization revenue",             "100–200", "Low",        "Informational"),
    ]
)

heading2(doc, "People Also Ask / PAA Questions to Answer in Article")
paa_qs = [
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
for q in paa_qs:
    bullet(doc, q)
doc.add_paragraph()

heading2(doc, "Meta Title (≤ 60 chars)")
body_para(doc, "How Niching Your Social Media Agency Makes You More Money")
heading2(doc, "Meta Description (≤ 160 chars)")
body_para(doc, "84% of agencies now identify as specialists — and they earn 25–40% margins vs. 15–20% for generalists. Here's the data and a framework to niche your agency.")
doc.add_paragraph()

# ── 3. AUDIENCE ────────────────────────────────────────────────────────────────
section_banner(doc, "3. TARGET AUDIENCE & ICP")

heading2(doc, "Primary Reader")
kv_table(doc, [
    ("Role",             "Founder / Owner of a social media or digital marketing agency"),
    ("Agency Size",      "2–50 employees; $200K–$5M annual revenue"),
    ("Stage",            "Growing but plateauing — frustrated by inconsistent deal flow, price pressure, and losing pitches to specialists who already know the prospect's industry"),
    ("Pain State",       "Serves too many client types; can't systematise delivery; underpriced compared to what specialist agencies in their space charge"),
    ("Goal",             "Scale revenue, increase margins, reduce sales cycle length, win better-fit clients who don't negotiate on price"),
    ("Reads / Follows",  "Jason Swenk, Karl Sakas, Blair Enns, Corey Quinn; agency podcasts; SocialPilot blog"),
])

heading2(doc, "Secondary Reader")
body_para(doc, "Freelance social media manager transitioning to an agency model, evaluating whether to go broad or narrow from day one.")
doc.add_paragraph()

heading2(doc, "Reader Mindset & Objections")
bullet(doc, "Knows 'niche down' is common advice but hasn't seen the actual financial case — needs proof, not inspiration")
bullet(doc, "Fears that niching = turning away revenue in the short term")
bullet(doc, "Wants a practical validation framework, not a theoretical positioning exercise")
bullet(doc, "Worried about market concentration risk: 'What if my niche tanks?'")
bullet(doc, "Has probably served a handful of clients in one vertical already but hasn't made it intentional")
doc.add_paragraph()

# ── 4. SERP & COMPETITIVE ANALYSIS ────────────────────────────────────────────
section_banner(doc, "4. SERP & COMPETITIVE ANALYSIS")

heading2(doc, "Current SERP Landscape")
body_para(doc, "The SERP for 'niche social media agency' is dominated by shallow listicle posts ('best niches to target'), vendor blog posts from agency tools, and generic how-to guides. Almost none cite primary benchmark reports or real case studies with specific revenue figures. There is clear white space for a single authoritative, data-first piece that combines financial benchmarks, named agency case studies, and an actionable niche-selection framework.")
doc.add_paragraph()

heading2(doc, "Top Competing Content Pieces")
data_table(doc,
    ["Source", "Angle", "Gap / Weakness"],
    [
        ("Agency Analytics Blog",    "'How to Pick Your Agency Niche' — framework-focused",                   "No revenue or margin benchmarks; no case studies with figures"),
        ("Agency Analytics Blog",    "'Why Niche Agencies Excel: 16 Experts' — quote compilation",           "Opinion-heavy; no primary data; no Promethean/Predictable Profits sourcing"),
        ("Phantom Leads AI",         "'Should Your Agency Niche Down? The Data Says Yes'",                   "Strong data points but thin content; no operational or vertical pricing angle"),
        ("ALM Corp Blog",            "'Most Profitable Niches for Digital Marketing Agencies 2026'",         "Niche listing only; no niche-selection framework or expert attribution"),
        ("InMotion Hosting Blog",    "'Generalist vs. Niche Agency: Which Model Grows Faster'",             "Balanced but shallow; no actionable steps or case study proof"),
        ("Seven Figure Agency Blog", "Agency niching and growth framework",                                   "Strong case studies but vendor-biased; no financial benchmarks from independent reports"),
    ]
)

heading2(doc, "Content Gap & SocialPilot Angle")
gap_block(doc,
    gap="No single piece combines: (a) independent financial benchmarks from Promethean Research, Predictable Profits, and Agency Analytics, (b) named case studies with specific revenue figures (Rankings.io, Seven Figure Agency's plumbing/HVAC agency), (c) expert quotes from Blair Enns, Karl Sakas, Corey Quinn, Jason Swenk, and David C. Baker, AND (d) a practical 3-step niche-selection and validation framework — in one article.",
    opportunity="SocialPilot can own the definitive data-backed guide to agency niching. The 84% specialist statistic (Promethean 2025) is a powerful hook: niching is no longer a growth hack — it's the new baseline, and the 16% still operating as generalists are being left behind.",
    angle="Lead with the paradigm shift (84% now specialists), build the financial case with hard benchmarks, validate with named case studies, give the practical framework, and close with niche-specific market sizing data. SocialPilot white-label positioned naturally as the infrastructure that enables scale."
)

# ── 5. LLM CONSIDERATIONS ──────────────────────────────────────────────────────
section_banner(doc, "5. LLM & AI SEARCH CONSIDERATIONS")
body_para(doc, "Structure this article to be extractable by AI-assisted search (ChatGPT, Gemini, Perplexity) answering queries like 'are niche agencies more profitable?' or 'how to choose a niche for my social media agency?'")
doc.add_paragraph()
bullet(doc, "Use question-format H2s and H3s where possible (e.g. 'Why Do Niche Agencies Earn More?', 'How Do You Choose the Right Niche?')")
bullet(doc, "Include data comparison tables: niche vs. generalist side-by-side across margin, retention, win rate, revenue per employee, and CAC")
bullet(doc, "Provide quotable, attribution-ready sentences: e.g. 'According to Promethean Research's 2025 survey of 1,452 agency leaders, 84% now identify as specialists'")
bullet(doc, "Include a concise FAQ section (5–6 Q&A pairs) at the end for schema markup eligibility")
bullet(doc, "Define key terms inline: vertical specialisation, service specialisation, category of one, agency of record (AOR)")
doc.add_paragraph()

# ── 6. KEY REQUIREMENTS ────────────────────────────────────────────────────────
section_banner(doc, "6. KEY REQUIREMENTS & MANDATORY INCLUSIONS")

reqs = [
    ("REQ 1", "The 84% Paradigm Shift Statistic — Lead With It",
     "Promethean Research's 2025 survey of 1,452 agency leaders found that 84% now identify as specialists. This is the hook. Frame it as: niching is no longer a growth tactic — it's the new baseline. The 16% still operating as generalists are competing against a market that has already moved."),

    ("REQ 2", "Niche vs. Generalist Benchmark Comparison Table",
     "Include a side-by-side table across: Gross Margin (niche: 40–75%; generalist: ~20%), Net Profit Margin (specialists: 25–40%; generalists: 15–20%), Annual Client Retention (8-figure niche agencies: 92%; 7-figure generalists: 78%), Client Annual Churn (niche: 15–20%; generalist: 25–35%), Proposal Close Rate (niche: 40–55%; generalist: 25–30%), Revenue per Employee (specialist target: $250K+; generalist target: $180K+). Sources: Promethean Research 2025, Predictable Profits 2025, Agency Analytics 2024, Phantom Leads AI 2025."),

    ("REQ 3", "The Service Reduction Finding (Promethean 2025)",
     "Agencies that REDUCED their service mix in 2025 posted 30% net margins vs. 10% for agencies that expanded services. Additionally, agencies making NO changes grew only 1.1% in 2024, vs. 8–9.7% for those that repositioned or refocused. Cite Promethean Research 2025 Digital Agency Industry Report (1,452 leaders surveyed)."),

    ("REQ 4", "Rankings.io Case Study — Full Story",
     "Chris Dreyer founded Rankings.io as a generalist law firm SEO agency. After coaching from Karl Sakas, he doubled down on personal injury law — which already represented ~66% of gross revenue. Result: grew from $2M to $12M in 4 years (6X), then to $30M (15X). Now 7 consecutive years on the Inc 5000 list. Raised minimum pricing 5X. Source: Sakas & Company case study + Starter Story interview."),

    ("REQ 5", "Seven Figure Agency Case Study — Plumbing/HVAC Agency",
     "An agency featured by Seven Figure Agency billed $425,000/month ($5.1M/year) by focusing exclusively on plumbing and HVAC companies — after previously serving 'everyone from lawyers to restaurants.' Hit seven figures within two years of niching. Source: Seven Figure Agency (who has tracked 133 agencies reaching $1M+ revenue — all had a defined niche)."),

    ("REQ 6", "Expert Voices — Named, Attributed, Specific",
     "REQUIRED quotes/positions from: (1) Blair Enns (Win Without Pitching): specialisation's three core benefits — sales advantage, pricing premium, client power; goal is to reduce or eliminate competition. (2) Karl Sakas: specialist agencies target $250K+ revenue/employee; Goldilocks market size = 2,000–10,000 prospects, 10–200 competitors. (3) David C. Baker: vertically-oriented firms are usually the Agency of Record relationship driver, so they get paid the most; expertise is not invented but emerges from what you've already done. (4) Corey Quinn: scaled Scorpion from $20M to $150M on deep specialisation; clients using his Focus Finder framework grow 93% faster. (5) Jason Swenk: review your most profitable clients quarterly, identify patterns, focus marketing on similar prospects; charge more so you can afford the right people."),

    ("REQ 7", "Niche-Specific Pricing Table by Vertical",
     "Include a verified pricing table: Healthcare ($3K–$15K/month; HIPAA compliance premium); Legal/Personal Injury ($12K–$25K+/month retainer + $5K–$50K ad spend); B2B SaaS by stage (Seed: $3K–$8K; Series A: $8K–$15K; Series B+: $15K–$75K+/month); E-commerce/DTC (performance-focused; $5K–$20K+/month); Restaurant/Hospitality ($1.5K–$7K+/month; more commoditised); Real Estate ($500–$3K/month; lower-margin vertical). Source: Clutch, Healthcare Success, CreeksideMarketing, GrowthSpree."),

    ("REQ 8", "Niche Selection Framework (3-Part + Validation Step)",
     "Framework: Expertise × Market Demand × Passion. All three required. Expertise: look at your 3–5 most profitable clients — what do they have in common? That's your starting point. Market Demand: Karl Sakas' Goldilocks rule — 2,000–10,000 prospects, 10–200 competitors. Passion: a motivated generalist outperforms a burned-out specialist. Quote David C. Baker: expertise emerges from what you've already done. Validation step: narrow your pipeline quietly for 90 days, sharpen case studies to spotlight target vertical, test new positioning before full commitment."),

    ("REQ 9", "Risks Section — Honest and Balanced",
     "Cover: market concentration risk (single-vertical dependency if industry contracts); recession vulnerability (hospitality in 2020 as example); scaling constraints if niche is too narrow. When NOT to niche: insufficient portfolio in target vertical, very small local market, personal burnout risk. Mitigation: choose verticals with regulatory complexity (HIPAA, legal compliance) as natural barriers to competition; validate market depth (2,000+ prospects) before committing."),

    ("REQ 10", "SocialPilot White-Label as Operational Enabler",
     "Tie in naturally: niche agencies need to deliver consistently and efficiently across multiple clients in the same vertical. SocialPilot's white-label platform (branded dashboards, automated reporting, multi-account management) enables niche agencies to systemise delivery without building proprietary tech. Reference the Elevation Brands case study: used SocialPilot's white-label and scheduling to scale client volume and substantially increase profit margins. Keep it editorial, not promotional."),
]
for req_id, req_title, req_desc in reqs:
    heading2(doc, f"{req_id}: {req_title}")
    body_para(doc, req_desc)
    doc.add_paragraph()

# ── 7. CITABLE CLAIMS ─────────────────────────────────────────────────────────
section_banner(doc, "7. CITABLE CLAIMS (Use Verbatim or Closely Paraphrased)")

claims = [
    ("Claim 1",
     "84% of digital agencies now identify as specialists — signalling that generalist positioning is no longer a viable default in 2025.",
     "Promethean Research 2025 Digital Agency Industry Report (1,452 agency leaders surveyed)",
     "https://prometheanresearch.com/2025-digital-agency-industry-report/"),
    ("Claim 2",
     "Agencies that reduced their service offerings in 2025 averaged 30% net profit margins, compared to just 10% for agencies that expanded their service mix — and agencies making no changes at all grew only 1.1%.",
     "Promethean Research 2025 Digital Agency Industry Report",
     "https://prometheanresearch.com/2025-digital-agency-industry-report/"),
    ("Claim 3",
     "Niche agencies close 40% more proposals, charge 2–3x more than generalists, and grow 3x faster in their first three years.",
     "Phantom Leads AI — Should Your Agency Niche Down? The Data Says Yes (2025)",
     "https://phantomleads.ai/blog/niche-down-or-stay-general"),
    ("Claim 4",
     "Top-performing 8-figure agencies retain 92% of clients annually and earn 25–32% net profit margins — compared to 78% retention and 18–22% margins for 7-figure agencies.",
     "Predictable Profits 2025 Agency Growth Benchmark (300+ agencies surveyed)",
     "https://predictableprofits.com/2025-agency-growth-benchmark-key-metrics-from-300-7-8-figure-agencies/"),
    ("Claim 5",
     "Rankings.io grew from $2 million to $12 million in four years — and subsequently to $30 million — after committing to personal injury law as its sole niche, earning seven consecutive years on the Inc 5000 list.",
     "Sakas & Company Case Study — Rankings.io; Starter Story",
     "https://sakasandcompany.com/case-studies/rankings-io-agency-growth/"),
    ("Claim 6",
     "Of 133 digital marketing agencies that reached seven-figure revenue tracked by Seven Figure Agency, every single one had a defined niche. One agency billing $425,000/month focused exclusively on plumbing and HVAC companies — after previously serving everyone from lawyers to restaurants.",
     "Seven Figure Agency — Agency Niche Growth Framework (2025)",
     "https://sevenfigureagency.com/niche-agency-growth/"),
    ("Claim 7",
     "Specialist agencies should target $250,000+ in revenue per employee, versus $180,000+ for generalist agencies — a 40% productivity gap driven by higher billing rates and operational efficiency.",
     "Karl Sakas / Sakas & Company — Agency Specialization Research; confirmed by Promethean Research 2025",
     "https://sakasandcompany.com/agency-specialization/"),
    ("Claim 8",
     "Average agency-client tenure has more than doubled since 2016 — from 3.2 years to approximately 7 years — reflecting the compounding retention advantage of agencies that develop deep client relationships through specialisation.",
     "ANA/4As 2025 Study (as cited in Focus Digital Agency Churn Report)",
     "https://focus-digital.co/average-marketing-agency-churn/"),
]
for num, claim, source, url in claims:
    heading2(doc, num)
    claim_block(doc, num, claim, source, url)

# ── 8. PAIN POINT ANALYSIS ────────────────────────────────────────────────────
section_banner(doc, "8. PAIN POINT ANALYSIS")
pain_point_table(doc, [
    ("Racing to the bottom on price",
     "Constantly losing bids to specialist agencies that charge more but win because the prospect trusts their vertical expertise",
     "Hard data showing exactly how much more specialists charge — and why clients pay it willingly"),
    ("Unpredictable pipeline and low close rates",
     "Marketing to everyone means no clear ICP; proposals feel generic; close rates hover around 25–30%",
     "Proof that niche agencies close 40–55% of proposals through faster trust-building and relevant case studies"),
    ("Operational chaos with each new client",
     "Every client feels like starting from scratch; no reusable systems; team constantly context-switching",
     "Evidence that vertical focus enables reusable SOPs, faster onboarding, and higher utilisation — directly lifting margins"),
    ("Fear of leaving money on the table",
     "Saying no to clients outside the niche feels like burning cash, especially in slow months",
     "The math: higher rates + lower CAC + better retention = more revenue than a full generalist pipeline"),
    ("Not knowing which niche to pick — and being stuck with it",
     "Overwhelmed by choices; worried about picking the wrong vertical and being trapped",
     "A low-risk validation framework: test quietly for 90 days before full commitment; expertise already exists in your client history"),
    ("Losing to bigger agencies",
     "Full-service agencies with 200 staff, established brands, and big portfolios seem unbeatable in pitches",
     "Data showing specialists consistently outperform generalists in niche-specific RFPs — even against much larger competitors — because they speak the prospect's language from the first call"),
])

# ── 9. KEY STATS FOR WRITER ───────────────────────────────────────────────────
section_banner(doc, "9. KEY STATS FOR WRITER")

stat_bullet(doc,
    "84% of digital agencies now identify as specialists — up from a minority position a decade ago.",
    "Promethean Research 2025 (1,452 agency leaders surveyed)")
stat_bullet(doc,
    "Agencies that reduced services in 2025 averaged 30% net margins vs. 10% for those that expanded; agencies making no changes grew only 1.1%.",
    "Promethean Research 2025 Digital Agency Industry Report")
stat_bullet(doc,
    "Specialist agencies: 25–40% profit margins. Generalist agencies: 15–20% profit margins.",
    "Swydo Agency Benchmarks / Agency Analytics 2024")
stat_bullet(doc,
    "Niche agencies close 40% more proposals, charge 2–3x more, and grow 3x faster in their first 3 years.",
    "Phantom Leads AI 2025")
stat_bullet(doc,
    "8-figure agencies retain 92% of clients annually vs. 78% for 7-figure agencies.",
    "Predictable Profits 2025 (300+ agencies)")
stat_bullet(doc,
    "Niche client annual churn: 15–20%. Generalist client annual churn: 25–35%.",
    "Phantom Leads AI / Focus Digital Agency Churn Report 2025")
stat_bullet(doc,
    "Average agency-client tenure has doubled since 2016 — from 3.2 years to ~7 years.",
    "ANA/4As 2025 Study")
stat_bullet(doc,
    "Specialist agencies target $250K+ revenue per employee vs. $180K+ for generalists — a 40% gap.",
    "Karl Sakas / Sakas & Company; Promethean Research 2025")
stat_bullet(doc,
    "Rankings.io: $2M → $12M (6X) → $30M (15X) by niching exclusively to personal injury law. 7 consecutive Inc 5000 appearances.",
    "Sakas & Company Case Study; Starter Story")
stat_bullet(doc,
    "133 agencies tracked by Seven Figure Agency all had a niche. One plumbing/HVAC-only agency bills $425K/month.",
    "Seven Figure Agency 2025")
stat_bullet(doc,
    "Corey Quinn scaled Scorpion from $20M to $150M on deep specialisation. Clients using his framework grow 93% faster than generalists.",
    "Corey Quinn / coreyquinn.com 2025")
stat_bullet(doc,
    "Legal marketing specialist agencies charge $12K–$25K+/month for personal injury clients, plus $5K–$50K/month in ad spend management.",
    "CreeksideMarketing / Dashing Digital 2024–2025")
stat_bullet(doc,
    "B2B SaaS agency retainers: $3K–$8K/month (seed stage), $8K–$15K/month (Series A), $15K–$75K+/month (Series B+).",
    "GrowthSpree / SaaSHero 2025–2026")
stat_bullet(doc,
    "US healthcare digital marketing spend: $24.77 billion in 2025, growing 13.3% YoY.",
    "IMARC Group / Statista 2025")
stat_bullet(doc,
    "US legal services advertising: $2.5+ billion across 26.9 million ads in 2024; most firms spend only 2–5% of revenue on marketing (under-invested).",
    "Hinge Marketing 2024 High Growth Study / LEXGRO 2026")
stat_bullet(doc,
    "Global e-commerce ad spend: $271 billion projected for 2025; 74% of e-commerce companies increasing ad budgets.",
    "MarketingLTB / Venture Media 2025")
stat_bullet(doc,
    "Only 22% of agencies raised their rates in 2024 — down from nearly half the prior year — reflecting the pricing pressure facing generalists who cannot justify premiums.",
    "Promethean Research 2025 Digital Agency Industry Report")
doc.add_paragraph()

# ── 10. NICHE VERTICAL MARKET SNAPSHOT ────────────────────────────────────────
section_banner(doc, "10. NICHE VERTICAL MARKET SNAPSHOT")
body_para(doc, "Use this table to help the writer frame each niche accurately. Do not rank these as 'best' — frame as a starting point for validation.", bold=True, after=100)
data_table(doc,
    ["Vertical", "Typical Monthly Retainer", "Market Size / Spend (US)", "Key Barrier to Entry", "Notes for Writer"],
    [
        ("Healthcare / Medical",   "$3,000–$15,000/month",        "$24.77B digital spend (2025, +13.3% YoY)",  "HIPAA compliance knowledge",           "High moat, high retention — clients don't switch easily"),
        ("Legal / Personal Injury","$12,000–$25,000+/month",      "$2.5B+ total legal ad spend (2024)",         "Legal marketing regulations, compliance","Rankings.io case study lives here — personal injury is peak-value sub-niche"),
        ("B2B SaaS",               "$8,000–$75,000+/month (stage-based)", "Rapid growth; AI/automation driving spend", "Deep product & funnel knowledge",    "Seed to Series B pricing ladder creates upsell path"),
        ("E-commerce / DTC",       "$5,000–$20,000+/month",       "$271B global ecom ad spend (2025)",          "Attribution & ROAS expertise",         "High spend but volatile; performance-linked pricing preferred"),
        ("Restaurant / Hospitality","$1,500–$7,000+/month",       "63% of restaurants increasing digital spend","Visual content & local SEO expertise", "Competitive and price-sensitive — need volume or premium positioning"),
        ("Real Estate",            "$500–$3,000/month",           "Moderate; seasonal spend patterns",           "Local market knowledge",               "Lower retainer floor — focus on volume or adjacent premium niches (luxury, commercial)"),
        ("Fitness & Wellness",     "$2,000–$8,000/month",         "$8.5T global wellness market",               "Community & influencer expertise",      "Strong content opportunities; high consumer engagement"),
        ("Home Services (HVAC, Plumbing)", "$2,000–$10,000/month","Fragmented, growing fast",                   "Local SEO & GMB expertise",            "Seven Figure Agency's $425K/month billing example lives here"),
    ]
)

# ── 11. ARTICLE STRUCTURE / WRITER CHECKLIST ──────────────────────────────────
section_banner(doc, "11. ARTICLE STRUCTURE & WRITER CHECKLIST")

checklist_sections = [
    ("INTRO (200–300 words)",
     [
         "Open with the paradigm-shift stat: 84% of agencies now identify as specialists (Promethean 2025). Niching is no longer the bold move — staying generalist is.",
         "State the core paradox: agencies that serve fewer types of clients earn more per client, close more proposals, and retain clients longer.",
         "Name what the article will deliver: financial benchmarks, two named case studies with specific revenue figures, and a practical niche-selection framework.",
         "Include primary keyword naturally within first 100 words.",
         "Do NOT open with a generic 'In today's competitive landscape...' sentence.",
     ]),
    ("SECTION 1: The Financial Case — What the Data Actually Shows (500–600 words)",
     [
         "Open with the Promethean service-reduction stat (30% vs. 10% margins; 1.1% growth for agencies making no changes).",
         "Present the niche vs. generalist comparison table (REQ 2) — this is the centrepiece of the financial case.",
         "Cover margin compression facing generalists: only 22% of agencies raised rates in 2024, down from nearly half the prior year (Promethean 2025).",
         "Cite the revenue-per-employee gap: $250K+ for specialists vs. $180K+ for generalists (Karl Sakas).",
         "Cover client churn differential: niche agencies 15–20% annual churn vs. 25–35% for generalists.",
         "Close with the ANA/4As stat: average agency-client tenure has doubled from 3.2 to ~7 years — retention compounds.",
     ]),
    ("SECTION 2: Why Niche Agencies Win More Business and Charge More (400–500 words)",
     [
         "Open with the close rate data: niche agencies close 40–55% of proposals vs. 25–30% for generalists.",
         "Explain the mechanism: the specialist enters the first sales call already knowing the prospect's industry, competitors, and challenges. No education phase = faster trust = shorter cycle.",
         "Quote Blair Enns on specialisation's three benefits: sales advantage, pricing premium, client relationship power; goal is to reduce or eliminate competition.",
         "Cover referral network advantage: niche agencies build industry-specific word-of-mouth that compounds; prospects in the same industry know each other.",
         "Include the vertical pricing table (REQ 7) to show concrete premium by niche.",
         "Quote Corey Quinn: scaled Scorpion from $20M to $150M on specialisation; his clients grow 93% faster.",
     ]),
    ("SECTION 3: The Operational Dividend — How Specialisation Multiplies Efficiency (300–400 words)",
     [
         "Cover SOPs, content templates, reporting dashboards that become reusable across clients in the same vertical.",
         "Faster onboarding: when agency already knows the industry, onboarding time drops significantly.",
         "Quote Jason Swenk: 'Review your most profitable clients quarterly, identify patterns, focus marketing on similar prospects. Charge more so you can afford the right people.'",
         "Quote David C. Baker: vertically-oriented firms are usually the Agency of Record relationship driver — they get paid the most.",
         "Mention SocialPilot white-label naturally: Elevation Brands case study — used SocialPilot to scale client volume and 'substantially increase profit margins' without building custom tech.",
     ]),
    ("SECTION 4: Real Proof — Two Agencies That Did It (400–500 words)",
     [
         "Case Study 1: Rankings.io. Full narrative (REQ 4): $2M → $12M → $30M. The personal injury specialisation decision. 5X pricing increase. 7 Inc 5000 appearances. Karl Sakas coaching role.",
         "Case Study 2: The Plumbing/HVAC Agency. $425K/month billing. Went from serving 'everyone from lawyers to restaurants' to HVAC-only. Hit seven figures within two years. Source: Seven Figure Agency (tracked 133 agencies — all had a niche).",
         "Brief mention of 2–3 other named specialists: Cardinal Digital Marketing (healthcare), Gripped.io (B2B SaaS), CSTMR (fintech) — shows range of verticals.",
         "Tie together: both case studies followed the same pattern — identified where they already had wins, doubled down, raised prices, built referral networks, systemised delivery.",
     ]),
    ("SECTION 5: How to Choose Your Niche — The 3-Part Framework + Validation (450–550 words)",
     [
         "Framework: Expertise × Market Demand × Passion — all three required. Missing one creates fragile specialisation.",
         "Expertise step: look at your 3–5 most successful/profitable clients — what vertical, size, or problem do they share? That overlap is your starting point. Quote David C. Baker: expertise is not invented; it emerges from what you've already done.",
         "Market Demand step: apply Karl Sakas' Goldilocks rule — 2,000–10,000 prospects, 10–200 competitors. Use: vertical-specific directories, LinkedIn Sales Navigator, Google Keyword Planner for search demand.",
         "Passion step: honest self-assessment. A motivated generalist outperforms a burned-out specialist. If thinking about this niche for 5 years makes you want to quit, listen to that.",
         "Validation step (CRITICAL): do not change your website, fire existing clients, or rebrand. For 90 days: narrow your active pipeline to the target niche, update case studies to spotlight target-vertical wins, adjust sales talk track. If it works, commit. If not, iterate.",
         "Include 'When NOT to niche' guidance: insufficient past experience in the target vertical; local market too small to sustain the niche; personal burnout risk from the subject matter.",
     ]),
    ("SECTION 6: The Risks Are Real — Here's How to Manage Them (250–300 words)",
     [
         "Acknowledge market concentration risk honestly — do not dismiss it. If a vertical contracts (hospitality in 2020, startup SaaS in 2023 funding winter), revenue is at risk.",
         "Mitigation: choose verticals with regulatory complexity (HIPAA, legal advertising rules) — compliance knowledge creates competitive barriers that protect the agency.",
         "Mitigation: validate market depth before committing. 2,000+ addressable prospects is minimum for a sustainable niche at the agency scale.",
         "Scaling constraint: very narrow niches (e.g. 'marketing for artisan cheese makers') have a ceiling — research the total addressable market before committing.",
         "Honest caveat: a motivated generalist outperforms a burned-out specialist. Passion matters.",
     ]),
    ("SECTION 7: The Most Profitable Niches Right Now — Market Data by Vertical (300–400 words)",
     [
         "Use the Niche Vertical Market Snapshot table from the brief (Section 10) as the data foundation.",
         "Frame as: 'Here's what the market data shows — but the right niche for YOUR agency is the intersection of this data and your existing expertise.' Do not make this a guaranteed gold rush list.",
         "Highlight three top-tier opportunities with strongest data: Healthcare ($24.77B market, 13.3% YoY growth, compliance moat), Legal (personal injury — Rankings.io proof, $12K–$25K+ retainers, $2.5B+ market), B2B SaaS ($8K–$75K/month by growth stage, AI-driven expansion).",
         "Cover the underdog opportunity: Home Services (HVAC, plumbing, roofing) — fragmented, growing, underserved by specialists, as proven by the $425K/month agency.",
     ]),
    ("CONCLUSION & CTA (200–250 words)",
     [
         "Restate the core argument: the data is unambiguous — specialists earn more, retain clients longer, close more deals, and build businesses that sell for more.",
         "Return to the 84% stat: if 84% of agencies are already specialists, staying generalist means competing with the bottom 16% on price while specialists take the best clients.",
         "Practical next step: identify your best 3 clients right now. What do they have in common? Start there.",
         "CTA: link to SocialPilot's agency resources / white-label solution for agencies scaling a specialised practice.",
         "FAQ section (5–6 Q&As) for schema markup. Questions drawn from the PAA list in Section 2 of this brief.",
     ]),
]

for section_title, checks in checklist_sections:
    sub_banner(doc, section_title)
    for item in checks:
        checkbox(doc, item)
    doc.add_paragraph()

# ── 12. KEY SOURCES FOR WRITER ────────────────────────────────────────────────
section_banner(doc, "12. KEY SOURCES FOR WRITER")
body_para(doc, "The writer must read and cite from the sources below. Do not fabricate or paraphrase statistics without verifying them at the source URL.", bold=True, after=100)
data_table(doc,
    ["Source", "Key Data Point", "URL"],
    [
        ("Promethean Research 2025 Digital Agency Industry Report",
         "84% specialists; 30% vs. 10% margins; 1.1% growth for no-change agencies; 22% raised rates",
         "https://prometheanresearch.com/2025-digital-agency-industry-report/"),
        ("Predictable Profits 2025 Agency Growth Benchmark (300+ agencies)",
         "92% vs. 78% retention; 25–32% vs. 18–22% margins by revenue tier",
         "https://predictableprofits.com/2025-agency-growth-benchmark-key-metrics-from-300-7-8-figure-agencies/"),
        ("Agency Analytics 2024 Marketing Agency Benchmarks Report",
         "Specialist billing rate premiums; retainer prevalence (43% of packages)",
         "https://agencyanalytics.com/company/newsroom/2024-benchmarks-report-press-release"),
        ("Phantom Leads AI — Should Your Agency Niche Down?",
         "40% more proposals closed; 2–3x higher rates; 3x faster growth in year 1–3",
         "https://phantomleads.ai/blog/niche-down-or-stay-general"),
        ("Sakas & Company — Rankings.io Case Study",
         "Full $2M → $12M → $30M revenue progression; pricing 5X after niching",
         "https://sakasandcompany.com/case-studies/rankings-io-agency-growth/"),
        ("Sakas & Company — Agency Specialization",
         "$250K+ revenue/employee target; Goldilocks market sizing (2K–10K prospects)",
         "https://sakasandcompany.com/agency-specialization/"),
        ("Seven Figure Agency — Niche Agency Growth",
         "133 agencies tracked; all had a niche; $425K/month HVAC agency",
         "https://sevenfigureagency.com/niche-agency-growth/"),
        ("Corey Quinn — Deep Specialization Accelerator",
         "Scorpion scaled $20M to $150M on specialisation; 93% faster growth",
         "https://www.coreyquinn.com/about"),
        ("Focus Digital — Average Marketing Agency Churn Report",
         "Niche 15–20% churn vs. generalist 25–35%; ANA/4As tenure doubling",
         "https://focus-digital.co/average-marketing-agency-churn/"),
        ("IMARC Group — US Healthcare Digital Marketing Market",
         "$24.77B in 2025; 13.3% YoY growth; online sub-segment at 12.44% CAGR",
         "https://www.imarcgroup.com/us-healthcare-advertising-market"),
        ("GrowthSpree — B2B SaaS Marketing Agency Pricing 2025–2026",
         "$3K–$8K (seed); $8K–$15K (Series A); $15K–$75K+ (Series B+)",
         "https://www.growthspreeofficial.com/blogs/b2b-saas-marketing-agency-pricing-2026-what-youll-actually-pay"),
        ("LEXGRO / Hinge Marketing — Legal Marketing Spend",
         "$2.5B+ total legal ad spend; 2–5% of revenue typical (underinvestment opportunity)",
         "https://lexgro.com/insights/law-firm-marketing-spend-2026/"),
        ("SocialPilot — Elevation Brands Case Study",
         "White-label + scheduling led to scaled client volume and substantially higher profit margins",
         "https://www.socialpilot.co/case-study/elevation-brands"),
    ]
)

# ── 13. INTERNAL LINKS ────────────────────────────────────────────────────────
section_banner(doc, "13. INTERNAL LINKS (Suggested)")
data_table(doc,
    ["Suggested Anchor Text", "URL", "Placement Suggestion"],
    [
        ("How to get clients for your digital marketing agency",
         "https://www.socialpilot.co/agency-success-guide/how-to-get-clients-for-digital-marketing",
         "Section 2 (Win Rates) — when discussing how niche agencies generate inbound referrals"),
        ("How to grow a digital marketing agency",
         "https://www.socialpilot.co/agency-success-guide/how-to-grow-a-digital-marketing-agency",
         "Intro or Conclusion — when framing niching as a growth lever"),
        ("Performance-based retainers for social media agencies",
         "https://www.socialpilot.co/blog/performance-based-retainer-social-media-agency",
         "Section 1 (Financial Case) — when covering pricing models that reward specialist results"),
        ("Social media management cost: how much to charge",
         "https://www.socialpilot.co/blog/social-media-management-cost",
         "Section 1 (Financial Case) or niche pricing table — when referencing rate benchmarks"),
        ("Client retention: challenges, strategies and more",
         "https://www.socialpilot.co/agency-success-guide/client-retention",
         "Section 1 (Financial Case) — when covering the retention advantage of niche agencies"),
        ("White-label social media reporting for agencies",
         "https://www.socialpilot.co/blog/white-label-social-media-reporting",
         "Section 3 (Operational Dividend) — when discussing systemised delivery and reporting"),
        ("How to upsell social media clients",
         "https://www.socialpilot.co/blog/upsell-social-media-clients",
         "Conclusion — once a niche is established, upselling deeper services is the growth lever"),
        ("Client onboarding guide for agencies",
         "https://www.socialpilot.co/agency-success-guide/guide-to-client-onboarding",
         "Section 3 (Operational Dividend) — when covering faster onboarding in a specialised niche"),
        ("SocialPilot for agencies",
         "https://www.socialpilot.co/solutions/agencies",
         "Section 3 or CTA — main agency solutions page"),
        ("Agency Success Guide hub",
         "https://www.socialpilot.co/agency-success-guide",
         "Conclusion — signpost readers to broader agency growth resources"),
    ]
)

# ── SAVE ───────────────────────────────────────────────────────────────────────
out = "/home/user/new/CONTENT_BRIEF__SocialPilot__Niche_Agency_v2.docx"
doc.save(out)
print(f"Saved: {out}")
