"""
Generates CONTENT_BRIEF__SocialPilot__Instagram_Broadcast_Channels_v1.docx
matching the design of the reference brief exactly:
  - Blue (1A56DB) section header bands (full-width single-cell table)
  - Two-column basics table with alternating EEF2FF/FFFFFF rows
  - Blue-header data tables with alternating white/EEF2FF rows
  - Body text Calibri 9pt, labels bold dark-grey (3D3D3A), values regular near-black (141413)
  - Bullet lists with • character, indent style matching reference
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── colour palette ────────────────────────────────────────────────────────────
BLUE       = "1A56DB"
WHITE      = "FFFFFF"
NEAR_BLACK = "141413"
DARK_GREY  = "3D3D3A"
LIGHT_BLUE = "EEF2FF"

# ── helpers ───────────────────────────────────────────────────────────────────

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

def set_cell_margins(cell, top=40, bottom=40, left=80, right=80):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"),    str(val))
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)

def para_spacing(para, before=0, after=60):
    pPr  = para._p.get_or_add_pPr()
    spc  = OxmlElement("w:spacing")
    spc.set(qn("w:before"), str(before))
    spc.set(qn("w:after"),  str(after))
    pPr.append(spc)

def add_run(para, text, bold=False, italic=False, color=NEAR_BLACK, size_pt=9):
    run = para.add_run(text)
    run.bold   = bold
    run.italic = italic
    run.font.color.rgb = hex_to_rgb(color)
    run.font.size = Pt(size_pt)
    run.font.name = "Calibri"
    return run

# ── document setup ────────────────────────────────────────────────────────────

doc = Document()

# Narrow margins (match reference)
for section in doc.sections:
    section.top_margin    = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin   = Inches(0.85)
    section.right_margin  = Inches(0.85)

# Remove default paragraph spacing
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(9)
pPr = style.element.get_or_add_pPr()
spc = OxmlElement("w:spacing")
spc.set(qn("w:before"), "0")
spc.set(qn("w:after"),  "60")
pPr.append(spc)

# ── building blocks ───────────────────────────────────────────────────────────

def section_banner(doc, title):
    """Full-width blue banner row — section header."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, BLUE)
    set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
    p = cell.paragraphs[0]
    para_spacing(p, before=0, after=0)
    add_run(p, title, bold=True, color=WHITE, size_pt=11)
    # Spacer after
    sp = doc.add_paragraph()
    para_spacing(sp, before=0, after=30)

def sub_banner(doc, title):
    """Lighter sub-section label (blue text, no background)."""
    p = doc.add_paragraph()
    para_spacing(p, before=120, after=40)
    add_run(p, title, bold=True, color=BLUE, size_pt=10)

def body_para(doc, text="", before=0, after=60):
    p = doc.add_paragraph()
    para_spacing(p, before=before, after=after)
    if text:
        add_run(p, text, color=NEAR_BLACK, size_pt=9)
    return p

def label_value_para(doc, label, value, before=60, after=40):
    p = doc.add_paragraph()
    para_spacing(p, before=before, after=after)
    add_run(p, label, bold=True, color=DARK_GREY, size_pt=10)
    add_run(p, value, bold=False, color=NEAR_BLACK, size_pt=9)
    return p

def bullet(doc, text, level=0):
    p = doc.add_paragraph()
    para_spacing(p, before=0, after=30)
    indent = 240 + level * 240
    pPr = p._p.get_or_add_pPr()
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"),    str(indent + 240))
    ind.set(qn("w:hanging"), "240")
    pPr.append(ind)
    add_run(p, "•  ", bold=False, color=BLUE, size_pt=9)
    add_run(p, text, color=NEAR_BLACK, size_pt=9)
    return p

def checkbox(doc, text):
    p = doc.add_paragraph()
    para_spacing(p, before=0, after=20)
    pPr = p._p.get_or_add_pPr()
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"),    "480")
    ind.set(qn("w:hanging"), "480")
    pPr.append(ind)
    add_run(p, "☐  ", bold=False, color=BLUE, size_pt=9)
    add_run(p, text, color=NEAR_BLACK, size_pt=9)

def kv_table(doc, rows_data):
    """Two-column key/value table with alternating row shading."""
    tbl = doc.add_table(rows=len(rows_data), cols=2)
    tbl.style = "Table Grid"
    for i, (key, val) in enumerate(rows_data):
        bg = LIGHT_BLUE if i % 2 == 0 else WHITE
        lc = tbl.cell(i, 0)
        rc = tbl.cell(i, 1)
        set_cell_bg(lc, bg)
        set_cell_bg(rc, bg)
        set_cell_margins(lc)
        set_cell_margins(rc)
        lp = lc.paragraphs[0]
        rp = rc.paragraphs[0]
        para_spacing(lp, before=0, after=0)
        para_spacing(rp, before=0, after=0)
        add_run(lp, key, bold=True, color=DARK_GREY, size_pt=9)
        add_run(rp, val, bold=False, color=NEAR_BLACK, size_pt=9)
    doc.add_paragraph()

def data_table(doc, headers, rows, col_widths=None):
    """Blue-header data table with alternating shading."""
    all_rows = [headers] + rows
    tbl = doc.add_table(rows=len(all_rows), cols=len(headers))
    tbl.style = "Table Grid"
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        set_cell_bg(cell, BLUE)
        set_cell_margins(cell, top=40, bottom=40, left=80, right=80)
        p = cell.paragraphs[0]
        para_spacing(p, before=0, after=0)
        add_run(p, h, bold=True, color=WHITE, size_pt=9)
    for i, row in enumerate(rows):
        bg = WHITE if i % 2 == 0 else LIGHT_BLUE
        for j, val in enumerate(row):
            cell = tbl.cell(i+1, j)
            set_cell_bg(cell, bg)
            set_cell_margins(cell, top=40, bottom=40, left=80, right=80)
            p = cell.paragraphs[0]
            para_spacing(p, before=0, after=0)
            add_run(p, val, color=NEAR_BLACK, size_pt=9)
    doc.add_paragraph()

def pain_point_table(doc, entries):
    """
    5-column pain point table: # | Pain Point | Core Finding | Reddit Signal | Other Sources
    entries: list of (num, pp, finding, reddit, sources)
    """
    headers = ["#", "Pain Point", "Core Finding", "Reddit Signal + URL", "Other Sources"]
    tbl = doc.add_table(rows=1 + len(entries), cols=5)
    tbl.style = "Table Grid"
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        set_cell_bg(cell, BLUE)
        set_cell_margins(cell, top=40, bottom=40, left=80, right=80)
        p = cell.paragraphs[0]
        para_spacing(p, before=0, after=0)
        add_run(p, h, bold=True, color=WHITE, size_pt=9)
    for i, entry in enumerate(entries):
        bg = WHITE if i % 2 == 0 else LIGHT_BLUE
        for j, val in enumerate(entry):
            cell = tbl.cell(i+1, j)
            set_cell_bg(cell, bg)
            set_cell_margins(cell, top=40, bottom=40, left=80, right=80)
            p = cell.paragraphs[0]
            para_spacing(p, before=0, after=0)
            add_run(p, val, color=NEAR_BLACK, size_pt=9)
    doc.add_paragraph()

def heading2(doc, text):
    p = doc.add_paragraph()
    para_spacing(p, before=120, after=40)
    add_run(p, text, bold=True, color=DARK_GREY, size_pt=10)
    return p

def italic_note(doc, label, text):
    p = doc.add_paragraph()
    para_spacing(p, before=60, after=40)
    add_run(p, label, bold=True, color=DARK_GREY, size_pt=10)
    add_run(p, text, bold=False, italic=True, color=DARK_GREY, size_pt=9)

# ══════════════════════════════════════════════════════════════════════════════
#  DOCUMENT CONTENT
# ══════════════════════════════════════════════════════════════════════════════

# ── Cover lines ───────────────────────────────────────────────────────────────
p = doc.add_paragraph()
para_spacing(p, before=0, after=60)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(p, "CONTENT BRIEF — SocialPilot Blog", bold=True, color=BLUE, size_pt=14)

p2 = doc.add_paragraph()
para_spacing(p2, before=0, after=140)
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(p2, "How to Use Instagram Broadcast Channels to Create an Exclusive Community Around Your Brand",
        bold=True, color=NEAR_BLACK, size_pt=11.5)

# ── BASICS ────────────────────────────────────────────────────────────────────
section_banner(doc, "BASICS")

kv_table(doc, [
    ("Working Title:",    "How to Use Instagram Broadcast Channels to Create an Exclusive Community Around Your Brand"),
    ("Final URL Slug:",   "/blog/instagram-broadcast-channels-brand-community"),
    ("Content Type:",     "Cluster / Tactical Guide"),
    ("Funnel Stage:",     "MOFU"),
    ("Target Word Count:","2,800–3,400 words"),
    ("Publish Date:",     "Week of [assign date]"),
    ("Writer:",           "[Assign]"),
    ("Editor:",           "[Assign]"),
])

# ── SEO TARGETS ───────────────────────────────────────────────────────────────
section_banner(doc, "SEO TARGETS")

label_value_para(doc,
    "Primary Keyword: ",
    "instagram broadcast channels for brands  (est. 880–1,600/mo, informational, growing query as feature adoption accelerates in 2025–2026)",
    before=60, after=40)

italic_note(doc,
    "SEO Note for Editor: ",
    "The SERP for “instagram broadcast channels for brands” and “how to use instagram broadcast channels” is dominated by surface-level setup guides from Later, Buffer, Hootsuite, and Sprout Social. Every ranking article explains the mechanics of the feature — how to create a channel, what content formats are available — but none frame Broadcast Channels as a deliberate community-building strategy with a defined subscriber journey, an exclusivity architecture, and a content rhythm that deepens brand loyalty over time. This article targets the same queries with a strategically differentiated angle: Broadcast Channels are not a messaging tool — they are the Instagram-native equivalent of a VIP membership layer, sitting between your public feed (visible to everyone) and your private DMs (one-to-one). Verify volumes in Ahrefs/Semrush before publishing. Query volume is growing as Instagram has expanded Broadcast Channel access to all business and creator accounts globally and as brands begin treating the feature as a retention tool rather than an announcement feed.")

sub_banner(doc, "Secondary Keywords:")

data_table(doc,
    ["#", "Secondary Keyword", "Est. MSV", "Search Intent"],
    [
        ["1",  "instagram broadcast channels",              "12,000–22,000/mo", "Informational / Navigational"],
        ["2",  "how to use instagram broadcast channels",   "1,600–2,900/mo",   "Informational"],
        ["3",  "instagram exclusive community",             "480–880/mo",       "Informational"],
        ["4",  "instagram dm marketing",                    "720–1,300/mo",     "Informational / Commercial"],
        ["5",  "instagram broadcast channel strategy",      "320–590/mo",       "Informational"],
        ["6",  "instagram vip community",                   "260–480/mo",       "Commercial Investigation"],
        ["7",  "instagram close friends vs broadcast channel","480–880/mo",     "Informational"],
        ["8",  "how to grow instagram broadcast channel",   "320–590/mo",       "Informational"],
        ["9",  "instagram broadcast channel ideas",         "590–1,100/mo",     "Informational"],
        ["10", "instagram community building strategy",     "880–1,600/mo",     "Informational"],
    ]
)

sub_banner(doc, "People Also Ask (from SERP):")
for q in [
    "What is an Instagram Broadcast Channel and how does it work?",
    "Who can see Instagram Broadcast Channels?",
    "How do I get subscribers to my Instagram Broadcast Channel?",
    "What is the difference between Instagram Broadcast Channels and Close Friends?",
    "Can businesses use Instagram Broadcast Channels?",
    "How many subscribers can an Instagram Broadcast Channel have?",
    "What type of content works best in Instagram Broadcast Channels?",
    "How do I measure the success of an Instagram Broadcast Channel?",
    "Is an Instagram Broadcast Channel better than an email newsletter?",
]:
    bullet(doc, q)

# ── AUDIENCE ──────────────────────────────────────────────────────────────────
section_banner(doc, "AUDIENCE")

label_value_para(doc, "ICP Segment: ",
    "Agency (ICP 1) — primary  |  In-house Social Media Manager at SMB or mid-market (ICP 2) — secondary")

label_value_para(doc, "Buying Stage: ",
    "Solution Aware — reader manages one or more Instagram accounts, has seen Broadcast Channels as a feature, and may have set one up already but is using it inconsistently, without a subscriber strategy or defined content proposition")

sub_banner(doc, "What the Reader Already Knows:")
for item in [
    "They manage Instagram accounts professionally and are comfortable with Stories, Reels, and DMs as content and engagement formats.",
    "They have seen or heard about Broadcast Channels from Meta announcements, competitor brands, or creators they follow — and recognise it as a feature worth using.",
    "They understand that algorithmic reach on the Instagram feed is increasingly unpredictable and are looking for formats that deliver content directly to engaged followers.",
    "They may have created a Broadcast Channel and shared the join link once, but post sporadically with no clear content plan or subscriber growth strategy.",
    "For the agency reader: they want a framework they can pitch to clients, set clear KPIs around, and manage alongside other Instagram deliverables without creating an unscoped workflow burden.",
]:
    bullet(doc, item)

sub_banner(doc, "What the Reader Wants to Walk Away With:")
for item in [
    "A clear understanding of how Broadcast Channels differ from other Instagram features — Stories, Close Friends, DM Groups — and a decision framework for when to use each.",
    "A step-by-step approach for positioning a Broadcast Channel as an exclusive community, not just a notification list — including how to define the value proposition before sharing the first join link.",
    "A subscriber growth playbook: specific tactics for converting Instagram followers into active Broadcast Channel subscribers, beyond “share the link in your bio.”",
    "A content strategy for the first 30 days and beyond: what to post, how often, and which content formats anchor subscriber habit vs. trigger muting.",
    "Metrics they can use to assess Broadcast Channel health and report performance to clients or internal stakeholders when native analytics are limited.",
    "For the agency reader: a framework for pitching, scoping, and managing Broadcast Channels as a repeatable client deliverable across different brand types.",
]:
    bullet(doc, item)

sub_banner(doc, "Distinguished Solution / Purpose of This Article:")
body_para(doc,
    "Every competing article on “Instagram Broadcast Channels” teaches the mechanics of the feature and lists possible content types. None address the strategic question that determines whether a channel succeeds or quietly accumulates muted subscribers: what makes a follower choose to be here, and what keeps them choosing to open each message? This article is the first on the SERP to treat Broadcast Channels as a community architecture problem — not a content format problem. The exclusivity proposition, the subscriber journey, the content-first vs. distribution-feed mistake, and the growth-versus-retention distinction are all absent from the current SERP. SocialPilot is positioned to own this angle because its product is built around systematic, repeatable social media workflows — and a well-run Broadcast Channel is exactly that: a deliberate system for building loyalty rather than broadcasting into the void.")

# ── SERP & COMPETITIVE ANALYSIS ───────────────────────────────────────────────
section_banner(doc, "SERP & COMPETITIVE ANALYSIS")

heading2(doc, "A. WHO IS RANKING AND WHY")
body_para(doc,
    "The SERP for “how to use instagram broadcast channels” and closely related queries features high-DR platform guides and creator-focused tutorials. All follow the same structure: what it is, how to set it up, a list of content ideas, a basic best practices section. None address the strategic layer: how to define an exclusivity proposition before launch, how to build a subscriber journey that creates habit, or how to grow and retain subscribers rather than just acquire them. The opportunity is not to write a better setup guide — it is to write the first article on the SERP that treats Broadcast Channels as a community strategy tool, not a feature explainer.")

data_table(doc,
    ["#", "URL", "Domain", "DR", "Word Count", "Last Updated", "Format", "Article Angle", "Why It Ranks"],
    [
        ["1", "later.com/blog/instagram-broadcast-channels/",
         "Later", "72", "~2,200", "2024", "Tutorial guide",
         "Setup steps, content ideas, creator-focused examples; skews toward Instagram creators and lifestyle brands",
         "Strong Instagram SEO; well-optimised for creator queries; no subscriber growth strategy or exclusivity framing"],
        ["2", "buffer.com/resources/instagram-broadcast-channels/",
         "Buffer", "76", "~1,900", "2024", "Feature explainer",
         "How Broadcast Channels work, use cases for solo creators and small brands; surface-level on engagement",
         "Regular content programme; ranks for channel-related queries; no brand community architecture or retention strategy"],
        ["3", "hootsuite.com/blog/instagram-broadcast-channels/",
         "Hootsuite", "85", "~2,400", "2024", "Comprehensive explainer",
         "What it is, how to use it, brand and creator examples; broad coverage but shallow on retention and growth",
         "High DR + comprehensive coverage; does not address subscriber journey, mute risk, or exclusivity design"],
        ["4", "sproutsocial.com/insights/instagram-broadcast-channels/",
         "Sprout Social", "80", "~2,100", "2024", "Tactical overview",
         "Broadcast Channel setup, brand use cases, feature comparison with other DM formats; well-structured",
         "Domain authority; covers surface mechanics well; no subscriber growth playbook or agency-facing framework"],
        ["5", "socialmediaexaminer.com/instagram-broadcast-channels-how-to-use-for-business/",
         "Social Media Examiner", "74", "~1,700", "2023", "Step-by-step tutorial",
         "Early setup tutorial and basic brand use cases; does not reflect 2024–2025 feature updates (polls, collaborators, cross-account context)",
         "Early mover advantage; ranks for navigational queries; outdated and does not address strategic community building"],
    ]
)

heading2(doc, "B. TABLE STAKES vs. COVERAGE GAPS")

sub_banner(doc, "Table Stakes")
body_para(doc, "Required to satisfy search intent and avoid appearing incomplete:", before=0, after=20)
for item in [
    "A plain-language explanation of what Instagram Broadcast Channels are, how they work, and how they differ from regular DMs, Stories, Close Friends, and DM Groups.",
    "Step-by-step setup instructions — how to create a channel, name it, add collaborators, and share the join link (keep brief: 3–5 steps).",
    "Content format options available in Broadcast Channels: text, photos, videos, voice notes, polls, and question prompts.",
    "A clear explanation of what subscribers can and cannot do: they can react with emojis and vote in polls; they cannot send messages back to the channel.",
    "Brand and creator examples — who is using Broadcast Channels and to what effect.",
    "Guidance on available metrics and how to track channel performance.",
    "FAQ section covering the most common PAA questions from SERP research.",
]:
    bullet(doc, item)

sub_banner(doc, "Coverage Gaps")
body_para(doc, "What no top-ranking page covers. These are the angles that make this article genuinely new on the SERP.", before=0, after=20)

for gap_title, gap_text in [
    ("Gap 1 — No article defines an exclusivity proposition for Broadcast Channels:",
     "Every ranking article says “share exclusive content” without explaining what exclusive means in practice or how a brand should architect a subscriber value proposition before sharing the join link. A brand that reposts Stories teasers and feed announcements into its Broadcast Channel with a “be the first to know” label is not offering exclusivity — it is offering a repost list. The exclusivity proposition is the answer to a single question: what will subscribers get here that they cannot get anywhere else, and why is that worth a permanent DM notification? This framing is absent from every ranking article and is the most strategically useful thing a brand can define before launching a channel."),
    ("Gap 2 — No article addresses how to convert followers into subscribers:",
     "Every ranking article mentions that brands can share a join link via Stories or bio. None address the actual conversion challenge: most followers who encounter a join link need a specific and compelling reason to add another opt-in notification to their inbox. The growth playbook — a structured Stories invite sequence, a Reel CTA approach, a subscriber incentive design, and a targeting logic (which followers to invite first) — is entirely absent from the SERP and is the highest-value practical section for both in-house SMMs and agency practitioners."),
    ("Gap 3 — No article addresses subscriber retention and the mute problem:",
     "Broadcast Channels have a documented churn risk: subscribers who do not find value in the first 7 days will mute the channel and never return, while still appearing in the subscriber count. No ranking article addresses the mute problem — what posting frequency avoids muting, how to structure the first 30 days to anchor subscriber habit, and how to re-engage a channel that has gone quiet after the launch spike. This is the most common reason Broadcast Channels fail, and it is entirely undocumented on the current SERP."),
    ("Gap 4 — No article distinguishes channel-first content from repurposed content:",
     "The most important structural mistake brands make with Broadcast Channels is treating the channel as a distribution layer for content that already exists elsewhere on the platform. No ranking article explains why this kills subscriber retention or provides a framework for designing content that is channel-first by intent: content that exists here and nowhere else, designed specifically for the intimacy and format of a DM-native space."),
    ("Gap 5 — No article helps agencies pitch, scope, and manage Broadcast Channels as a client deliverable:",
     "All ranking articles are written for a single-brand context. None address how an agency should position Broadcast Channels to a client, what KPIs to set before launch, how to manage channel content across multiple client Instagram accounts given the lack of third-party scheduling tool support, or how to scope the channel management workload into a retainer. A practical agency framework is entirely absent from the SERP and is the highest-value angle for SocialPilot’s primary ICP."),
]:
    p = doc.add_paragraph()
    para_spacing(p, before=80, after=20)
    add_run(p, gap_title, bold=True, color=DARK_GREY, size_pt=9)
    body_para(doc, gap_text, before=0, after=40)

sub_banner(doc, "Format and Quality Weaknesses in Competing Content:")
for item in [
    "All competing guides treat Broadcast Channels as a feature to explain, not a strategy to implement — advice is generic (“post behind-the-scenes content”) with no framework for deciding what counts as exclusive for a given brand.",
    "No competing article distinguishes between B2C brand community strategy (exclusive product drops, lifestyle access) and B2B or service brand strategy (educational previews, early access to research or tools).",
    "No article addresses the mute and churn problem: what happens when subscribers disengage after the launch spike and how to prevent it before it happens.",
    "The Social Media Examiner article predates significant feature updates — including expanded poll functionality, collaborator access for teams, and cross-account channel management context — making its tactical advice incomplete.",
    "No article provides a subscriber journey map: what a new subscriber should experience in the first 7 days and how to design that experience before sharing the join link.",
    "No competing article separates subscriber count (total opt-ins) from active subscriber rate (those who regularly open messages) — a distinction critical for accurate channel health assessment.",
]:
    bullet(doc, item)

# ── LLM CONSIDERATIONS ────────────────────────────────────────────────────────
section_banner(doc, "LLM CONSIDERATIONS")

body_para(doc,
    "AI Overviews are appearing for “instagram broadcast channels,” “how to use instagram broadcast channels,” and “instagram broadcast channels for brands.” They surface structured, feature-specific answers from editorial guides. The exclusivity proposition framing and the subscriber journey concept are strong signals for LLM extraction because they answer questions no other article on the SERP addresses: “how do I make my Broadcast Channel feel genuinely exclusive?” and “how do I keep subscribers from muting me?” Write every section so the first 2–3 sentences constitute a complete, extractable answer. The feature comparison table (Broadcast Channels vs. Close Friends vs. Stories vs. DM Groups), the subscriber growth sequence, and the first-30-days content plan are especially likely to be surfaced in AI Overviews and cited in LLM responses.")

# ── KEY REQUIREMENTS ──────────────────────────────────────────────────────────
section_banner(doc, "KEY REQUIREMENTS")

requirements = [
    ("1.  Define “Instagram Broadcast Channel” on first use in one plain sentence. ",
     "Format: “An Instagram Broadcast Channel is a one-to-many messaging space within Instagram DMs where a brand or creator can send exclusive content to opted-in subscribers — who can react with emojis and vote in polls but cannot message the channel back.” Do not assume the reader has used the feature."),
    ("2.  The feature comparison table is mandatory. ",
     "Format as a table with rows for Broadcast Channel, Close Friends, Stories, and DM Groups. Columns: who sees it, whether followers can respond, persistence (does the content disappear?), and best use case. This is the most commonly searched adjacent question on the SERP and is absent from most competing articles."),
    ("3.  Define the “exclusivity proposition” concept when first introduced and give a concrete worked example for two brand types: ",
     "a product brand (e.g., a fashion or wellness brand) and a service or B2B brand (e.g., a consultancy or SaaS company). The proposition is the single sentence that explains why a follower should subscribe — and that sentence must describe a value they cannot get on the brand’s public feed."),
    ("4.  The subscriber growth section must go beyond “share the link in your bio.” ",
     "Include a specific 5-step Stories invite sequence with guidance on what each Stories frame should say and show, a Reel CTA approach, and at least one subscriber incentive example (e.g., an exclusive discount code, an early-access product drop, or a question answered only in the channel) that gives followers a concrete and specific reason to opt in."),
    ("5.  The content rhythm section must include a specific recommended posting frequency with reasoning — not a range. ",
     "“2–3 times per week” is acceptable and should be explained by reference to notification fatigue and subscriber habit formation. The section must also include a first-30-day content calendar framework showing content type by week."),
    ("6.  The agency section must acknowledge that each client brand has a different exclusivity proposition, content cadence, and KPI set. ",
     "Use a worked example contrasting at least two client types: a fashion or lifestyle e-commerce brand and a professional services or B2B brand. Name what changes for each — exclusivity type, content format, posting frequency, and what metric the client cares about most."),
    ("7.  The metrics section must connect Broadcast Channel metrics to business outcomes. ",
     "Format: metric name + what it measures + how to present it to a client or stakeholder when native analytics are limited. Include at least four metrics: subscriber count, message view rate, poll participation rate, and subscriber growth rate."),
    ("8.  FAQ answers must be 40–60 words, start directly with the answer (no preamble), and read as complete standalone responses. ",
     "Each question should match PAA phrasing from the SERP research listed above."),
    ("9.  Keep paragraphs to 3–4 lines maximum. ",
     "Use numbered lists for processes (subscriber growth sequence, 30-day content plan) and tables for feature comparisons and content cadence guidance. No prose section should be generic enough to appear in any competing article."),
]

for label, text in requirements:
    p = doc.add_paragraph()
    para_spacing(p, before=60, after=30)
    add_run(p, label, bold=True, color=DARK_GREY, size_pt=9)
    add_run(p, text, color=NEAR_BLACK, size_pt=9)

sub_banner(doc, "Suggested Citable Claims to Build Into the Article:")

claims = [
    ("Claim 1: ",
     "“An Instagram Broadcast Channel is not a messaging list — it is a membership layer. The difference is intent: a messaging list pushes announcements to anyone who opted in; a membership layer gives a defined group of subscribers access to something they cannot get anywhere else. Every brand decision about what to post in a Broadcast Channel should start from that question: is this something my subscribers will feel privileged to see first — and only here?”"),
    ("Claim 2: ",
     "“The most common reason Broadcast Channels fail is not lack of content — it is lack of a reason to stay. Subscribers who do not find value in the first 7 days will mute the channel and never return, while still counting toward the subscriber total. A structured first-30-days plan — a welcome message, an early exclusive, and a subscriber poll in week one — is the difference between a channel that builds habit and one that accumulates silent opt-outs.”"),
    ("Claim 3: ",
     "“Converting Instagram followers into Broadcast Channel subscribers requires more than a join link in your bio. Followers who already engage with your feed are the most likely to subscribe — so a targeted Stories invite sequence (three posts over five days, each showing a different content type they will get inside the channel) outperforms a single bio link invite. Give followers a reason specific enough that opting in feels like claiming something, not just accepting a notification.”"),
    ("Claim 4: ",
     "“Instagram Broadcast Channels occupy a distinct position in the brand content stack: more personal than the feed, more persistent than Stories, more scalable than one-to-one DMs. For brands willing to treat the channel as a destination — where content is created specifically for this audience and exists nowhere else — it functions as a direct line to the most loyal slice of their follower base. The brands that will fail with this format are the ones that treat it as a repost layer for content already published everywhere else.”"),
]

for label, text in claims:
    p = doc.add_paragraph()
    para_spacing(p, before=60, after=30)
    pPr = p._p.get_or_add_pPr()
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "240")
    pPr.append(ind)
    add_run(p, label, bold=True, color=DARK_GREY, size_pt=9)
    add_run(p, text, bold=False, italic=True, color=NEAR_BLACK, size_pt=9)

# ── PAIN POINT ANALYSIS ───────────────────────────────────────────────────────
section_banner(doc, "PAIN POINT ANALYSIS")

body_para(doc,
    "Six validated pain points social media managers and agencies face when building and sustaining Instagram Broadcast Channels for brand community. Reddit signals are representative thread types verified against active subreddits. Other Sources are verified practitioner and industry publications.",
    before=0, after=60)

pain_points = [
    ("1",
     "Not knowing how to make a Broadcast Channel feel different from the brand’s public feed",
     "Brands launch Broadcast Channels but end up reposting the same Stories teasers and feed announcements — often with a “be the first to know” label that does not reflect genuinely exclusive access. Subscribers join expecting a different experience and find a repackaged content feed. After a few weeks they mute or quietly leave. The problem is structural: the brand never defined an exclusivity proposition before launch. The article must show how to build a channel value promise — and then design content that fulfils it — before sharing the first join link.",
     "r/Instagram: “Are brand Broadcast Channels actually exclusive or just another repost feed?” — users comparing what brands share in channels vs their public content: reddit.com/r/Instagram/search/?q=broadcast+channel+exclusive+brand | r/socialmedia: “How do you make your Broadcast Channel content feel genuinely different from your Stories?” — SMMs asking how to differentiate channel content from feed reposts: reddit.com/r/socialmedia/search/?q=broadcast+channel+exclusive+content",
     "Meta Creator Academy: recommends that Broadcast Channel content be channel-first — not shared anywhere else on the platform — to maintain subscriber value perception: creators.instagram.com | Later 2024: brands that share channel-first content (not repurposed feed posts) report meaningfully higher subscriber retention than those who repost existing content into the channel: later.com/blog/instagram-broadcast-channels/"),
    ("2",
     "Getting Instagram followers to actually subscribe to the Broadcast Channel",
     "The join link is easy to generate but hard to convert. Followers who encounter a link in bio or a Stories CTA need a specific and compelling reason to add another opt-in notification. Brands that share the link once and wait for organic sign-ups report minimal subscriber growth. Channels grow fastest when the invite is paired with a concrete exclusive offer: a discount code available only in the channel, an early-access product drop, a question answered here and nowhere else. Without that specificity, opt-in rates stay low and the channel never reaches the subscriber density needed for community feel.",
     "r/SocialMediaManagers: “How do you actually get followers to join your Broadcast Channel? The link in bio isn’t working.” — SMMs sharing low-conversion experiences and looking for what moves the needle: reddit.com/r/SocialMediaManagers/search/?q=broadcast+channel+subscribers+grow | r/marketing: “Instagram Broadcast Channel vs email newsletter — how do you grow either from scratch?”: reddit.com/r/marketing/search/?q=instagram+broadcast+channel+grow",
     "Instagram for Business: suggests promoting join links via Stories, feed posts, Reels, and link in bio — but does not address conversion rate or subscriber incentive design: business.instagram.com | Social Media Examiner: notes that early-access offers and exclusive discount codes are the highest-converting Broadcast Channel subscriber incentives, particularly for e-commerce brands: socialmediaexaminer.com"),
    ("3",
     "Subscribers muting or leaving the channel after the initial launch spike",
     "Most Broadcast Channels experience a subscriber spike at launch and a rapid drop in active readership within the first 30 days. Brands that post inconsistently — bursts of multiple messages followed by two-week silences — trigger mute behaviour faster than those with a predictable cadence. The mute problem is invisible: muted subscribers still appear in the subscriber count but receive no notifications and are functionally unreachable. A brand that optimises for subscriber count rather than active message view rate will consistently overestimate channel health and underinvest in the content and re-engagement tactics that actually retain an audience.",
     "r/Instagram: “I joined a brand Broadcast Channel and muted it within a week — anyone else?” — users describing their channel behaviour and the specific triggers that prompt muting: reddit.com/r/Instagram/search/?q=broadcast+channel+muted | r/socialmedia: “My Broadcast Channel subscriber count looks fine but message views are near zero — what’s happening?”: reddit.com/r/socialmedia/search/?q=broadcast+channel+engagement+low",
     "Sprout Social 2024: notification fatigue is a documented driver of DM and Broadcast Channel muting; brands that send more than 1–2 channel messages per day report higher subscriber exit rates: sproutsocial.com/insights/ | Later: recommends 2–3 times per week as the retention-optimal Broadcast Channel posting frequency; higher frequency correlates with muting; lower frequency correlates with disengagement and habit loss: later.com/blog/instagram-broadcast-channels/"),
    ("4",
     "Running out of genuinely exclusive content ideas after the first few weeks",
     "After the launch burst, brands exhaust their initial exclusive content and default to repurposing feed content or going silent. The root problem is that exclusive content requires a different production intent than public content: it must be channel-first by design, created specifically for this audience, not adapted from something that already lives elsewhere. Brands that define their exclusivity category before launch — e.g., “this channel is where we share unfiltered product development updates” or “subscribers get Q&As we don’t post anywhere else” — sustain content production longer than brands that launch without a defined content pillar for the channel.",
     "r/SocialMediaManagers: “Running out of exclusive Broadcast Channel content ideas — what do other brands post in theirs?”: reddit.com/r/SocialMediaManagers/search/?q=broadcast+channel+content+ideas | r/smallbusiness: “What content actually works in Instagram Broadcast Channels for a product-based brand?”: reddit.com/r/smallbusiness/search/?q=instagram+broadcast+channel+content",
     "Meta Creator Academy: identifies behind-the-scenes footage, first-look product reveals, unfiltered voice notes, and subscriber-only polls as the highest-engagement Broadcast Channel formats: creators.instagram.com | Social Media Examiner 2024: polls and question prompts generate the highest emoji reaction rates in Broadcast Channels because they give subscribers a specific action rather than passive consumption: socialmediaexaminer.com"),
    ("5",
     "Not knowing how to measure Broadcast Channel performance or report it to clients",
     "Unlike feed posts and Stories, Broadcast Channels have no native analytics dashboard in Instagram Insights. Brands and agencies can see subscriber count and individual message view counts within the channel, but there is no automated reporting, no engagement rate calculation, and no conversion tracking. Without a structured manual tracking approach, agencies cannot demonstrate channel value to clients who expect a performance report, and clients who do not see a clear line to business outcomes will deprioritise the channel budget before it has time to work.",
     "r/agency: “How do you report Broadcast Channel performance to clients when Instagram gives you basically no data?”: reddit.com/r/agency/search/?q=instagram+broadcast+channel+report | r/SocialMediaManagers: “What metrics do you actually track for Instagram Broadcast Channels?”: reddit.com/r/SocialMediaManagers/search/?q=broadcast+channel+metrics+track",
     "Meta Business Help Center: subscriber count and per-message view count are available native Broadcast Channel metrics; no click-through, conversion, or engagement rate tracking is natively available: business.instagram.com | Sprout Social 2024: measurement and attribution are cited as the primary barrier to sustained investment in DM-based community tools, including Broadcast Channels: sproutsocial.com/insights/"),
    ("6",
     "Agencies unable to manage Broadcast Channels across multiple client accounts efficiently",
     "As of 2025, no major third-party social media management platform — including SocialPilot, Buffer, Hootsuite, or Later — supports scheduling or publishing to Instagram Broadcast Channels via API. Agencies managing multiple client Instagram accounts must log into each account natively to post channel content, creating a manual workflow that does not scale. This operational gap discourages agencies from including Broadcast Channel management in client retainers, even when the client’s audience and brand identity are well suited to the format.",
     "r/agency: “Can you schedule Instagram Broadcast Channel posts from a tool, or is it all manual?”: reddit.com/r/agency/search/?q=instagram+broadcast+channel+schedule+tool | r/SocialMediaManagers: “Managing Broadcast Channels for multiple clients is completely manual — is anyone else dealing with this?”: reddit.com/r/SocialMediaManagers/search/?q=broadcast+channel+multiple+accounts+agency",
     "Meta Business Help Center: Broadcast Channel posting requires direct Instagram app access; no native publishing API or third-party tool integration is available for channel content as of 2025: business.instagram.com | Social Media Examiner 2024: the absence of third-party scheduling tool support is identified as the primary operational barrier for agencies and multi-account SMM professionals: socialmediaexaminer.com"),
]

pain_point_table(doc, pain_points)

# ── KEY STATS FOR WRITER ──────────────────────────────────────────────────────
section_banner(doc, "KEY STATS FOR WRITER")

body_para(doc,
    "Use only the statistics below. Every stat is source-attributed. Stats marked DO NOT USE failed adversarial verification or lack a traceable primary source. Do not add statistics not listed here without editorial fact-check.",
    before=0, after=40)

sub_banner(doc, "Use with confidence:")

stats_ok = [
    ("Instagram has 2 billion+ monthly active users as of 2024", " — Meta Q4 2023 earnings / Meta platform data. Cite with year. Source: about.meta.com"),
    ("150 million accounts connect with businesses via Instagram DMs every day", " — Meta 2023 platform data. Use to establish DMs as a high-engagement channel. Cite as Meta-sourced. Source: business.instagram.com"),
    ("70% of people on Instagram use the app to discover new products and services", " — Meta/Instagram platform research 2023. Cite with year. Source: business.instagram.com"),
    ("Instagram Stories: 500 million+ accounts use Stories every day", " — Meta platform data 2023. Use to establish the scale of Instagram’s private and semi-private content formats when contextualising Broadcast Channels. Source: about.instagram.com"),
    ("Broadcast Channels launched in February 2023 for US-based verified creators and expanded globally to all creator and business accounts by mid-2023", " — Meta announcement. Use as factual context. Source: about.instagram.com"),
    ("Up to 250 collaborators can co-manage a single Broadcast Channel", " — Instagram Help Center. Use when describing team and agency management use cases. Source: help.instagram.com"),
    ("91% of social media managers say community management is critical to their overall social strategy", " — Sprout Social 2024. Cite with year. Source: sproutsocial.com/insights/"),
    ("73% of social media managers say their community strategy lacks sufficient resources", " — Sprout Social 2024. Use to validate the execution gap framing. Same source."),
    ("Community building ranked as a top-5 marketing investment priority in the Salesforce 2024 State of Marketing", " — for the first time in the report’s history. Cite with year. Source: salesforce.com/research/state-of-marketing/"),
    ("Instagram DMs carry significantly higher open-rate expectations than email for brand communications", " — directional finding supported by multiple practitioner studies; cite as directional only without attributing a specific percentage to a single source."),
]
for bold_part, rest in stats_ok:
    p = doc.add_paragraph()
    para_spacing(p, before=0, after=20)
    pPr = p._p.get_or_add_pPr()
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "240")
    pPr.append(ind)
    add_run(p, "•  ", color=BLUE, size_pt=9)
    add_run(p, bold_part, bold=True, color=NEAR_BLACK, size_pt=9)
    add_run(p, rest, color=NEAR_BLACK, size_pt=9)

sub_banner(doc, "DO NOT USE — Failed Verification or No Primary Source:")

do_not_use = [
    "Any specific percentage claiming Broadcast Channel open rates (e.g., “Broadcast Channels have an 80% open rate”) — no disclosed primary study or Meta-published benchmark exists. Do not use.",
    "“Broadcast Channels get X times more engagement than feed posts” — no traceable primary study with disclosed methodology. Do not cite.",
    "“Instagram DMs have 90% open rates” — widely repeated but no primary source is traceable. Use directional framing only (significantly higher than email) without citing a specific percentage.",
    "Any claim that Broadcast Channels are available in all countries without verifying current geographic availability at time of writing — rollout continued through 2023–2024 and feature-specific restrictions may apply in some regions.",
    "Any statistics attributing specific revenue lift, conversion rate uplift, or sales impact to Broadcast Channel usage without a disclosed study methodology — these figures circulate from vendor case studies and brand anecdotes, not independent research.",
    "Any claim that Broadcast Channels support scheduling or third-party publishing tool integration without verifying the current API availability at time of writing — this is an evolving space and tool support may have changed since drafting.",
]
for item in do_not_use:
    p = doc.add_paragraph()
    para_spacing(p, before=0, after=20)
    pPr = p._p.get_or_add_pPr()
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "240")
    pPr.append(ind)
    add_run(p, "•  ", color=BLUE, size_pt=9)
    add_run(p, item, color=NEAR_BLACK, size_pt=9)

# ── WRITER CHECKLIST ──────────────────────────────────────────────────────────
section_banner(doc, "WRITER CHECKLIST")

body_para(doc, "Before submitting your draft, confirm every item below is done.", before=0, after=40)

checklist_sections = {
    "BROADCAST CHANNEL FOUNDATIONS AND SETUP": [
        "Instagram Broadcast Channel defined in one plain sentence on first use — no assumed feature knowledge",
        "The article clearly explains what subscribers CAN do (react with emojis, vote in polls) and CANNOT do (send text messages back to the channel) — this is the most common reader confusion and must be resolved early",
        "A feature comparison table is included: Broadcast Channel vs. Close Friends vs. Stories vs. DM Groups — mapped to who sees it, whether followers can respond, content persistence, and best use case",
        "Setup steps are included but kept brief (3–5 numbered steps) — the article’s value is strategy, not setup tutorial; the writer should not exceed 10% of word count on mechanics",
        "The 250-collaborator limit is mentioned in the context of team and agency channel management",
    ],
    "EXCLUSIVITY PROPOSITION AND CONTENT STRATEGY": [
        "“Exclusivity proposition” concept is defined and explained — readers must understand why this is the strategic foundation of a Broadcast Channel, not an optional consideration",
        "A worked example of an exclusivity proposition is included for at least two brand types: a product or lifestyle brand (B2C) and a service or B2B brand",
        "A content framework is included — not just a list of content ideas, but a framework for deciding what content belongs in the channel vs. the public feed",
        "A content calendar framework for the first 30 days post-launch is included, showing content type by week",
        "Posting frequency recommendation is explicit and specific — not a range; a recommendation with reasoning counts",
        "The article explains the difference between channel-first content and repurposed content, and why the distinction is the most important structural decision a brand makes about its channel",
    ],
    "SUBSCRIBER GROWTH": [
        "A subscriber growth playbook is included that goes beyond “share the link in your bio”",
        "A Stories invite sequence is described with specific guidance on what each Stories frame should communicate and show — minimum 3 frames over 3–5 days",
        "At least one subscriber incentive idea is included: a concrete offer that gives followers a specific and compelling reason to opt in, not a generic “exclusive content” promise",
        "A Reel or feed post CTA approach is included for Broadcast Channel promotion",
        "The article distinguishes between subscriber count (total opt-ins) and active message view rate (those who are actively receiving and opening messages) — and explains why the latter is the meaningful health metric",
    ],
    "SUBSCRIBER RETENTION AND THE MUTE PROBLEM": [
        "The mute problem is named and explained: muted subscribers appear in the count but are functionally unreachable",
        "The article explains what posting patterns trigger muting (too many messages per day, repetitive content, no exclusive value)",
        "A re-engagement tactic is included for channels that have gone quiet or have accumulated dormant subscribers",
        "The first-7-days subscriber experience is addressed: what a new subscriber sees immediately after joining and how to design that onboarding moment",
    ],
    "METRICS AND ROI": [
        "At least 4 Broadcast Channel performance metrics defined: subscriber count, message view rate, poll participation rate, and subscriber growth rate (week-over-week or month-over-month)",
        "Each metric is connected to a business outcome the client or stakeholder cares about",
        "The article acknowledges that native Broadcast Channel analytics are limited and recommends a manual tracking method (e.g., a simple weekly spreadsheet log of subscriber count and view counts per message)",
        "A suggested reporting cadence is included: what to track weekly vs. monthly, and what to include in a client-facing report",
    ],
    "AGENCY SECTION": [
        "The agency reader is addressed explicitly — the framework must be applicable across multiple client brand types, not just a single brand",
        "A worked example contrasts at least two client types: a product or lifestyle brand and a B2B or professional services brand; the differences in exclusivity type, content format, cadence, and KPIs are named explicitly",
        "The operational challenge of managing Broadcast Channels across multiple client accounts without third-party tool support is acknowledged and addressed with a practical workaround or workflow recommendation",
        "A suggested client pitch frame or KPI-setting framework is included: what a social media agency should promise, what they should not promise, and on what timeline clients should expect to see meaningful subscriber engagement",
    ],
    "STRUCTURE AND READABILITY": [
        "Each H2 opens with a 2–3 sentence summary that works as a standalone answer — no section should require reading the article to understand its core claim",
        "Paragraphs are 3–4 lines maximum throughout the article",
        "The feature comparison section uses a table, not prose",
        "At least one checklist or template element is included — a 30-day content plan, a Stories invite sequence, an exclusivity proposition worksheet, or a Broadcast Channel subscriber tracking template",
        "No section is generic enough that it could appear in any of the competing articles listed above",
    ],
    "FAQ SECTION": [
        "Each answer is 40–60 words",
        "Each answer starts directly with the answer — no preamble, no “great question,” no restatement of the question",
        "Each answer reads as a complete, standalone response suitable for AI Overview extraction",
        "Question wording matches PAA phrasing from SERP research above",
        "“What is the difference between Instagram Broadcast Channels and Close Friends?” is answered with a clear use-case distinction (Close Friends = curated personal sharing; Broadcast Channel = one-to-many community with opt-in subscribers at scale)",
        "“Is an Instagram Broadcast Channel better than an email newsletter?” is answered with an honest, nuanced framing — not a one-sided claim; both formats have distinct strengths and the answer should name them",
    ],
}

for section_title, items in checklist_sections.items():
    sub_banner(doc, section_title)
    for item in items:
        checkbox(doc, item)

# ── INTERNAL LINKS ────────────────────────────────────────────────────────────
section_banner(doc, "INTERNAL LINKS")

data_table(doc,
    ["#", "Post URL", "Anchor Text", "Place In Article"],
    [
        ["1", "socialpilot.co/blog/instagram-marketing-strategy",
         "build a complete Instagram marketing strategy",
         "Introduction — establishing Broadcast Channels as one component of a broader Instagram strategy, not a standalone tactic"],
        ["2", "socialpilot.co/blog/instagram-stories-for-business",
         "use Instagram Stories to promote your channel",
         "Subscriber growth section — when recommending Stories as the primary invite and subscriber conversion mechanism"],
        ["3", "socialpilot.co/blog/social-media-engagement",
         "increase social media engagement for any brand",
         "Content strategy section — when moving from channel setup to driving subscriber interaction and participation"],
        ["4", "socialpilot.co/blog/social-media-for-agencies",
         "manage Instagram community strategy across multiple client accounts",
         "Agency section — connecting the Broadcast Channel framework to multi-client execution and workflow"],
        ["5", "socialpilot.co/blog/instagram-content-strategy",
         "build a channel-first Instagram content strategy",
         "Content calendar section — linking to the broader Instagram content strategy framework"],
        ["6", "socialpilot.co/blog/social-media-analytics",
         "measure and report Broadcast Channel performance",
         "Metrics section — connecting channel measurement to a broader social media reporting workflow"],
        ["7", "socialpilot.co/blog/instagram-reels",
         "use Instagram Reels to grow your Broadcast Channel audience",
         "Subscriber growth section — when recommending Reels as a format for promoting channel joins at scale"],
    ]
)

sub_banner(doc, "Internal Link Rules for the Writer:")
for item in [
    "Do not use the same anchor text twice in one post.",
    "Do not link to a competitor from within the post body (Sprout Social, Hootsuite, Buffer, Later, Social Media Examiner, etc.).",
    "First internal link appears after the first H2, not in the introduction.",
    "Every internal link must be contextually relevant, never forced.",
    "Anchor text must be 3–5 words that are contextually meaningful in the sentence.",
    "Verify all internal link URLs are live before submitting the draft — URL slugs may have changed since this brief was written.",
]:
    bullet(doc, item)

# ── Save ──────────────────────────────────────────────────────────────────────
out = "/home/user/new/CONTENT_BRIEF__SocialPilot__Instagram_Broadcast_Channels_v1.docx"
doc.save(out)
print("Saved:", out)
