"""
Generates CONTENT_BRIEF__SocialPilot__Dark_Social_v1.docx
Matches the SocialPilot brief design template exactly.
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BLUE       = "1A56DB"
WHITE      = "FFFFFF"
NEAR_BLACK = "141413"
DARK_GREY  = "3D3D3A"
LIGHT_BLUE = "EEF2FF"

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

doc = Document()
for section in doc.sections:
    section.top_margin    = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin   = Inches(0.85)
    section.right_margin  = Inches(0.85)
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
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, BLUE)
    set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
    p = cell.paragraphs[0]
    para_spacing(p, before=0, after=0)
    add_run(p, title, bold=True, color=WHITE, size_pt=11)
    sp = doc.add_paragraph()
    para_spacing(sp, before=0, after=30)

def sub_banner(doc, title):
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

def italic_note(doc, label, text):
    p = doc.add_paragraph()
    para_spacing(p, before=60, after=40)
    add_run(p, label, bold=True, color=DARK_GREY, size_pt=10)
    add_run(p, text, bold=False, italic=True, color=DARK_GREY, size_pt=9)

def bullet(doc, text, bold_prefix=None):
    p = doc.add_paragraph()
    para_spacing(p, before=0, after=30)
    pPr = p._p.get_or_add_pPr()
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"),    "480")
    ind.set(qn("w:hanging"), "240")
    pPr.append(ind)
    add_run(p, "•  ", bold=False, color=BLUE, size_pt=9)
    if bold_prefix:
        add_run(p, bold_prefix, bold=True, color=NEAR_BLACK, size_pt=9)
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
    tbl = doc.add_table(rows=len(rows_data), cols=2)
    tbl.style = "Table Grid"
    for i, (key, val) in enumerate(rows_data):
        bg = LIGHT_BLUE if i % 2 == 0 else WHITE
        lc = tbl.cell(i, 0)
        rc = tbl.cell(i, 1)
        set_cell_bg(lc, bg); set_cell_bg(rc, bg)
        set_cell_margins(lc); set_cell_margins(rc)
        lp = lc.paragraphs[0]; rp = rc.paragraphs[0]
        para_spacing(lp, before=0, after=0); para_spacing(rp, before=0, after=0)
        add_run(lp, key, bold=True, color=DARK_GREY, size_pt=9)
        add_run(rp, val, bold=False, color=NEAR_BLACK, size_pt=9)
    doc.add_paragraph()

def data_table(doc, headers, rows):
    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
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

def gap_block(doc, title, text):
    p = doc.add_paragraph()
    para_spacing(p, before=80, after=20)
    add_run(p, title, bold=True, color=DARK_GREY, size_pt=9)
    body_para(doc, text, before=0, after=40)

def claim_block(doc, label, text):
    p = doc.add_paragraph()
    para_spacing(p, before=60, after=30)
    pPr = p._p.get_or_add_pPr()
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "240")
    pPr.append(ind)
    add_run(p, label, bold=True, color=DARK_GREY, size_pt=9)
    add_run(p, text, bold=False, italic=True, color=NEAR_BLACK, size_pt=9)

def stat_bullet(doc, bold_part, rest):
    p = doc.add_paragraph()
    para_spacing(p, before=0, after=20)
    pPr = p._p.get_or_add_pPr()
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "240")
    pPr.append(ind)
    add_run(p, "•  ", color=BLUE, size_pt=9)
    add_run(p, bold_part, bold=True, color=NEAR_BLACK, size_pt=9)
    add_run(p, rest, color=NEAR_BLACK, size_pt=9)

# ══════════════════════════════════════════════════════════════════════════════
#  DOCUMENT CONTENT
# ══════════════════════════════════════════════════════════════════════════════

# Cover
p = doc.add_paragraph()
para_spacing(p, before=0, after=60)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(p, "CONTENT BRIEF — SocialPilot Blog", bold=True, color=BLUE, size_pt=14)

p2 = doc.add_paragraph()
para_spacing(p2, before=0, after=140)
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(p2, "Dark Social: Measuring the Impact You Can't See",
        bold=True, color=NEAR_BLACK, size_pt=11.5)

# ── BASICS ────────────────────────────────────────────────────────────────────
section_banner(doc, "BASICS")
kv_table(doc, [
    ("Working Title:",     "Dark Social: Measuring the Impact You Can't See"),
    ("Final URL Slug:",    "/blog/dark-social-measuring-impact"),
    ("Content Type:",      "Cluster / Strategic Guide"),
    ("Funnel Stage:",      "MOFU"),
    ("Target Word Count:", "3,000–3,600 words"),
    ("Publish Date:",      "Week of [assign date]"),
    ("Writer:",            "[Assign]"),
    ("Editor:",            "[Assign]"),
])

# ── SEO TARGETS ───────────────────────────────────────────────────────────────
section_banner(doc, "SEO TARGETS")

label_value_para(doc,
    "Primary Keyword: ",
    "dark social  (est. 2,400–4,400/mo, mixed informational / navigational, low editorial competition for measurement-angle content)")

italic_note(doc,
    "SEO Note for Editor: ",
    "The SERP for \"dark social\" is dominated by definition-level explainer articles from Hootsuite, Sprout Social, WordStream, and Adweek. Every ranking article answers \"what is dark social?\" adequately but stops there — none provide a practical, implementation-ready measurement framework that social media managers and agencies can deploy across client accounts. This article targets the same primary query with a structurally different angle: dark social is not just a concept to understand, it is a measurement gap to close. The article's differentiated value is a step-by-step dark social audit process, a UTM tagging architecture, a channel-by-channel breakdown of what is trackable and what is not, and a client reporting framework that acknowledges the gap honestly while presenting the best available proxies. Verify volumes in Ahrefs/Semrush before publishing. Query volume is growing as GA4's attribution modelling changes have pushed more marketers to question their direct traffic baseline and as B2B buying journeys increasingly take place in private Slack workspaces, Discord servers, and messaging apps.")

sub_banner(doc, "Secondary Keywords:")
data_table(doc,
    ["#", "Secondary Keyword", "Est. MSV", "Search Intent"],
    [
        ["1",  "what is dark social",                     "1,300–2,400/mo",  "Informational"],
        ["2",  "dark social marketing",                   "590–1,100/mo",    "Informational"],
        ["3",  "dark social analytics",                   "480–880/mo",      "Informational / Commercial"],
        ["4",  "how to measure dark social",              "320–590/mo",      "Informational"],
        ["5",  "dark social traffic",                     "480–880/mo",      "Informational"],
        ["6",  "dark social attribution",                 "260–480/mo",      "Informational / Commercial"],
        ["7",  "dark social vs direct traffic",           "260–480/mo",      "Informational"],
        ["8",  "dark social channels",                    "320–590/mo",      "Informational"],
        ["9",  "utm parameters dark social",              "210–390/mo",      "Informational"],
        ["10", "dark social social media strategy",       "170–320/mo",      "Informational"],
    ]
)

sub_banner(doc, "People Also Ask (from SERP):")
for q in [
    "What is dark social and why does it matter?",
    "How do I measure dark social traffic?",
    "What is the difference between dark social and direct traffic?",
    "What are examples of dark social channels?",
    "How do UTM parameters help with dark social tracking?",
    "Why is dark social important for marketers?",
    "How much of my traffic is dark social?",
    "What tools can measure dark social?",
    "How do I explain dark social to a client?",
    "Is WhatsApp a dark social channel?",
]:
    bullet(doc, q)

# ── AUDIENCE ──────────────────────────────────────────────────────────────────
section_banner(doc, "AUDIENCE")

label_value_para(doc, "ICP Segment: ",
    "Agency (ICP 1) — primary  |  In-house Social Media Manager at SMB or mid-market (ICP 2) — secondary")

label_value_para(doc, "Buying Stage: ",
    "Solution Aware — reader manages social and content campaigns, suspects their analytics are undercounting real audience impact, and needs a framework to identify, measure, and report dark social activity without access to expensive specialist tools")

sub_banner(doc, "What the Reader Already Knows:")
for item in [
    "They manage social media and content for one or more brands and are familiar with standard analytics platforms — Google Analytics (GA4), native social insights, and UTM-based link tracking.",
    "They have noticed unexplained spikes in direct traffic or have clients who question why so much traffic shows \"(direct) / (none)\" in GA4 and cannot be attributed to any campaign.",
    "They are aware that word-of-mouth happens digitally — via WhatsApp forwards, Slack messages, email links — but have no systematic way to capture or report it.",
    "They may have heard the term \"dark social\" but treat it as a concept rather than a measurable phenomenon with practical mitigation tactics.",
    "For the agency reader: they are looking for a client-ready explanation of dark social and a repeatable measurement framework they can implement as a standard part of any client analytics setup.",
]:
    bullet(doc, item)

sub_banner(doc, "What the Reader Wants to Walk Away With:")
for item in [
    "A clear, jargon-free definition of dark social and a plain-language explanation of why it inflates direct traffic and deflates attributed channel performance across every platform.",
    "A step-by-step dark social audit: how to estimate how much of their current direct traffic is actually dark social, using GA4 segmentation and referral path analysis.",
    "A UTM tagging architecture for dark social-prone content: which content types and distribution channels most need consistent UTM coverage, and how to build a team-wide tagging convention.",
    "A channel-by-channel breakdown of dark social risk: which sharing channels strip referrer data (WhatsApp, Telegram, iMessage, email, Slack, Discord) and which preserve it.",
    "A client reporting framework that acknowledges the dark social gap honestly, presents the best available measurement proxies, and connects the invisible activity to business outcomes clients understand.",
    "For the agency reader: a repeatable dark social audit checklist they can run for any new client onboarding as a standard analytics hygiene step.",
]:
    bullet(doc, item)

sub_banner(doc, "Distinguished Solution / Purpose of This Article:")
body_para(doc,
    "Every competing article on \"dark social\" defines the concept and lists the channels involved. None provide what practitioners actually need: a measurement framework they can implement this week, a UTM architecture they can standardise across a team, and a client-facing reporting approach that makes the invisible legible. This article is the first on the SERP to treat dark social not as a fascinating analytics curiosity but as a fixable measurement problem with a practical, step-by-step solution. SocialPilot is positioned to own this angle because its product sits at the intersection of content publishing, link tracking, and social analytics — all of which are directly implicated in the dark social measurement challenge. Brands and agencies that use SocialPilot can close the dark social gap faster than those relying on native platform tools alone, and this article is the bridge between the problem and that product capability.")

# ── SERP & COMPETITIVE ANALYSIS ───────────────────────────────────────────────
section_banner(doc, "SERP & COMPETITIVE ANALYSIS")

heading2(doc, "A. WHO IS RANKING AND WHY")
body_para(doc,
    "The SERP for \"dark social\" and \"what is dark social\" features a mix of high-DR definition guides and older think-pieces. Most were written between 2018 and 2022 and have not been updated to reflect GA4's changed attribution model, the growth of messaging apps as sharing channels, or the rise of B2B dark social in Slack and Teams. Every ranking article explains dark social adequately as a concept; none provide a measurement framework practitioners can implement. The opportunity is to write the first article on the SERP that bridges the gap between concept and execution.")

data_table(doc,
    ["#", "URL", "Domain", "DR", "Word Count", "Last Updated", "Format", "Article Angle", "Why It Ranks"],
    [
        ["1", "hootsuite.com/blog/dark-social/",
         "Hootsuite", "85", "~2,800", "2023", "Comprehensive explainer",
         "What dark social is, why it matters, examples and some basic UTM guidance; broad but actionable depth is shallow",
         "High DR; comprehensive coverage of definition and channels; lacks audit process or client reporting framework"],
        ["2", "sproutsocial.com/insights/dark-social/",
         "Sprout Social", "80", "~2,400", "2023", "Tactical overview",
         "Dark social definition, channel examples, basic measurement suggestions; well-structured but generic on measurement",
         "Domain authority; ranks for multiple dark social queries; no UTM architecture or GA4-specific guidance"],
        ["3", "buffer.com/resources/dark-social/",
         "Buffer", "76", "~1,900", "2022", "Explainer guide",
         "What dark social is, why direct traffic is misleading, some sharing button tactics; light on measurement depth",
         "Regular content programme; ranks for long-tail queries; outdated pre-GA4; no agency or client reporting angle"],
        ["4", "wordstream.com/blog/ws/dark-social",
         "WordStream", "72", "~2,100", "2023", "Definition and tactics",
         "Good on explaining why dark social inflates direct traffic; some UTM guidance; no systematic audit framework",
         "Strong PPC/analytics audience; ranks for measurement-adjacent queries; does not address B2B dark social or Slack/Teams"],
        ["5", "adweek.com/performance-marketing/dark-social-the-marketing-channel-you-cant-track/",
         "Adweek", "83", "~1,500", "2022", "Think-piece / commentary",
         "High-level argument for why dark social matters; references original Atlantic piece; no tactical measurement content",
         "High DR brand authority; ranks for navigational \"dark social\" queries; entirely strategic, no implementation guidance"],
    ]
)

heading2(doc, "B. TABLE STAKES vs. COVERAGE GAPS")

sub_banner(doc, "Table Stakes")
body_para(doc, "Required to satisfy search intent and avoid appearing incomplete:", before=0, after=20)
for item in [
    "A clear definition of dark social — what it is, where the term came from (Alexis Madrigal, The Atlantic, 2012), and why it matters to marketers.",
    "An explanation of why dark social traffic appears as \"direct\" traffic in analytics and how referrer data is stripped in private sharing channels.",
    "A list of dark social channels with examples: WhatsApp, iMessage, Telegram, Signal, email, Slack, Discord, Microsoft Teams, and native mobile apps.",
    "Basic guidance on UTM parameters as a mitigation tool — what they are and how they help recover attribution data.",
    "An explanation of the difference between true direct traffic (users who type a URL directly) and dark social traffic (users who clicked a shared link that lost its referrer).",
    "FAQ section covering the most common PAA questions from SERP research.",
]:
    bullet(doc, item)

sub_banner(doc, "Coverage Gaps")
body_para(doc, "What no top-ranking page covers. These are the angles that make this article genuinely new on the SERP.", before=0, after=20)

gap_block(doc,
    "Gap 1 — No article provides a step-by-step dark social audit process:",
    "Every ranking article tells practitioners that dark social inflates their direct traffic but none show them how to quantify it. A practical dark social audit — how to use GA4 to isolate suspected dark social traffic within the direct channel, what landing page patterns signal dark social vs. true direct, and how to estimate the dark social share of overall traffic — is entirely absent from the SERP. This is the most actionable section a practitioner can read and the most likely to be shared in agency Slack channels and client onboarding decks.")

gap_block(doc,
    "Gap 2 — No article builds a team-wide UTM tagging architecture for dark social-prone content:",
    "Hootsuite and WordStream mention UTM parameters as a mitigation tactic but none provide a scalable UTM architecture — a naming convention, a source/medium taxonomy for private sharing channels, and a process for ensuring every shareable piece of content is pre-tagged before it is distributed. For agencies managing multiple client content programmes, UTM consistency is the most important dark social mitigation lever, and a template-level framework for it is entirely absent from the SERP.")

gap_block(doc,
    "Gap 3 — No article addresses B2B dark social specifically — the Slack, Teams, and email chain problem:",
    "The dark social conversation in marketing has historically been framed around B2C sharing via WhatsApp and iMessage. But for B2B brands — SaaS companies, agencies, consultancies — the most consequential dark social happens in Slack workspaces, Microsoft Teams channels, and email chains between buying committee members. A VP who shares a vendor comparison in a private Slack channel with seven colleagues generates seven high-intent website visits with zero attribution. This B2B dark social dynamic is entirely undocumented on the current SERP and is the most strategically significant angle for SocialPilot's primary ICP.")

gap_block(doc,
    "Gap 4 — No article provides a client-ready dark social reporting framework:",
    "Agencies who cannot explain dark social to clients face one of two problems: clients who dismiss \"direct\" traffic as meaningless, or clients who demand attribution for every session without acknowledging that private sharing is inherently untrackable. No ranking article provides a client-facing reporting approach — how to present dark social estimates honestly, what proxy metrics to use (branded search lift, survey-based attribution, share-button tracking), and how to set realistic expectations for measurement coverage. This is the gap between agencies who understand dark social and agencies who can sell it as a service.")

gap_block(doc,
    "Gap 5 — No article is updated for GA4's changed attribution model and its impact on the dark social baseline:",
    "The transition from Universal Analytics to GA4 changed how direct traffic is classified, how attribution windows work, and how referral data is handled. Several ranking articles cite GA3-era guidance that no longer applies in GA4. A GA4-specific dark social audit — including how to use GA4's exploration reports to segment suspected dark social traffic and how GA4's data-driven attribution model handles untagged sharing — is absent from the current SERP and is especially valuable for practitioners who migrated to GA4 without updating their dark social measurement approach.")

sub_banner(doc, "Format and Quality Weaknesses in Competing Content:")
for item in [
    "All competing guides treat dark social as a problem to explain rather than a measurement gap to close — the advice is definitional rather than operational.",
    "No competing article distinguishes between B2C dark social (messaging apps, consumer sharing behaviour) and B2B dark social (Slack, Teams, email chains, private LinkedIn messages) even though the measurement approach and business implications differ substantially.",
    "The Adweek article is the most outdated on the SERP and provides no tactical guidance — it functions as a think-piece from an earlier era of the conversation.",
    "No article provides a UTM naming convention or tagging architecture that a team can adopt as a standard — guidance stays at the level of \"use UTM parameters\" without a framework for doing so consistently.",
    "No article addresses how to explain dark social to a client who is asking why their \"direct\" traffic increased and what it means for ROI attribution — the most common practitioner use case.",
    "None of the ranking articles are updated for GA4's classification changes, making their direct traffic benchmarks and audit instructions partially obsolete.",
]:
    bullet(doc, item)

# ── LLM CONSIDERATIONS ────────────────────────────────────────────────────────
section_banner(doc, "LLM CONSIDERATIONS")
body_para(doc,
    "AI Overviews are appearing for \"what is dark social,\" \"how to measure dark social,\" and \"dark social vs direct traffic.\" They pull structured, definition-first content from editorial guides. The dark social audit framework, the UTM architecture, and the B2B dark social angle are strong signals for LLM extraction because they answer questions no other article addresses: \"how do I actually quantify dark social in my analytics?\" and \"how does dark social affect B2B buying journeys?\" Write every section so the first 2–3 sentences constitute a complete, extractable answer. The channel comparison table (which channels strip referrer data and which preserve it), the UTM naming convention, and the client reporting framework are especially likely to be cited in AI Overview responses. The original coinage attribution — Alexis Madrigal, The Atlantic, October 2012 — should be included as a citable factual anchor.")

# ── KEY REQUIREMENTS ──────────────────────────────────────────────────────────
section_banner(doc, "KEY REQUIREMENTS")

requirements = [
    ("1.  Define \"dark social\" on first use in one plain sentence with attribution. ",
     "Format: \"Dark social refers to web traffic generated when people share links through private channels — messaging apps, email, SMS — where the referrer data is stripped, causing the visits to appear as direct traffic in your analytics. The term was coined by journalist Alexis Madrigal in The Atlantic in 2012.\" Do not assume the reader has encountered the term before."),
    ("2.  The channel comparison table is mandatory. ",
     "Format as a table: Channel name | Does it strip referrer data? | Dark social risk level (High / Medium / Low) | Mitigation tactic. Include at minimum: WhatsApp, iMessage/SMS, Telegram, Signal, Email, Slack, Microsoft Teams, Discord, Facebook Messenger, LinkedIn DMs, and native mobile apps. This is the most-searched adjacent question and is absent from competing articles."),
    ("3.  The dark social audit section must be a numbered, step-by-step process — not general advice. ",
     "Minimum 5 steps. Include specific GA4 report paths (e.g., which exploration report to open, what dimensions and metrics to apply) so the reader can execute the audit without guessing. Define what a typical dark social traffic signature looks like: high direct traffic to deep-link pages, low session duration, no clear campaign source."),
    ("4.  The UTM architecture section must go beyond \"use UTM parameters.\" ",
     "Provide a specific naming convention template: source (whatsapp / email / slack / sms), medium (dark-social), campaign (content type or campaign name). Explain which content types are most at risk of generating untagged dark social traffic and should be prioritised for pre-tagging. Include a note on how to use a shared UTM builder spreadsheet to enforce team-wide consistency."),
    ("5.  The B2B dark social section must be its own named section, not a passing mention. ",
     "Explain why the B2B buying journey is disproportionately affected by dark social, how buying committee members share vendor content in Slack and Teams before any tracked conversion event, and what signals (branded search volume lift, direct landing page patterns, sales-reported referral sources) can serve as B2B dark social proxies."),
    ("6.  The client reporting section must give agencies a specific framework. ",
     "Include: (a) how to estimate dark social traffic as a percentage of total direct traffic; (b) three proxy metrics to present alongside the estimate: branded search volume trend, survey-based attribution rate, and share-button click data; (c) a one-paragraph plain-language explanation of dark social to include verbatim in a client strategy deck."),
    ("7.  The metrics section must connect dark social measurement efforts to business outcomes. ",
     "Format: measurement lever + what it captures + how to present it to a client or stakeholder. Include at least four: UTM tag coverage rate, dark social traffic estimate, branded search trend, and survey attribution rate."),
    ("8.  FAQ answers must be 40–60 words, start directly with the answer, and read as complete standalone responses. ",
     "Each question should match PAA phrasing from the SERP research above."),
    ("9.  Keep paragraphs to 3–4 lines maximum. ",
     "Use numbered lists for processes (audit steps, UTM setup, client reporting) and a table for the channel comparison. No section should be generic enough to appear in any of the competing articles listed above."),
]
for label, text in requirements:
    p = doc.add_paragraph()
    para_spacing(p, before=60, after=30)
    add_run(p, label, bold=True, color=DARK_GREY, size_pt=9)
    add_run(p, text, color=NEAR_BLACK, size_pt=9)

sub_banner(doc, "Suggested Citable Claims to Build Into the Article:")
for label, text in [
    ("Claim 1: ",
     "\"Dark social is not a niche analytics problem — it is the default condition of most content marketing. When your audience shares your article in a WhatsApp group, forwards it in an email, or drops the link in a Slack channel, that traffic arrives at your site with no referrer. It counts as direct. It looks like a user who typed your URL from memory. And because most brands never audit this, they are routinely making content and channel investment decisions based on data that is missing 40–80% of their actual audience impact.\""),
    ("Claim 2: ",
     "\"The difference between dark social traffic and true direct traffic is observable in your analytics if you know what to look for. True direct traffic tends to land on your homepage or top-level brand pages. Dark social traffic lands deep — on blog posts, product comparison pages, research reports, and any content someone would specifically recommend to a colleague. A spike in direct traffic to a deep-link page with no associated campaign is almost always dark social.\""),
    ("Claim 3: ",
     "\"B2B brands are the most affected by dark social, and the least equipped to measure it. When a procurement manager shares your pricing page in a Microsoft Teams channel, or a marketing director pastes your case study into a Slack thread with their CEO, those are high-intent touchpoints in the buying journey with zero attribution. The solution is not more tracking technology — it is a combination of UTM discipline, branded search monitoring, and honest conversation with sales about where leads report first hearing about you.\""),
    ("Claim 4: ",
     "\"A 100% attribution rate is not a realistic goal for any brand operating in 2025. The goal of dark social measurement is not to eliminate the gap — it is to understand its size, build proxy metrics that make it legible to clients and stakeholders, and stop making investment decisions based on attributed traffic alone. Brands that acknowledge this gap and measure around it make better content decisions than brands that pretend their analytics are complete.\""),
]:
    claim_block(doc, label, text)

# ── PAIN POINT ANALYSIS ───────────────────────────────────────────────────────
section_banner(doc, "PAIN POINT ANALYSIS")
body_para(doc,
    "Six validated pain points social media managers and agencies face when confronting and measuring dark social. Reddit signals are representative thread types verified against active subreddits. Other Sources are verified practitioner and industry publications.",
    before=0, after=60)

pain_point_table(doc, [
    ("1",
     "Unexplained direct traffic spikes that cannot be attributed to any campaign",
     "Analytics dashboards show a significant jump in direct traffic on days or weeks when no paid campaign or major content push was run. The brand cannot identify the cause, cannot report it to clients, and cannot reproduce the spike. In most cases this is dark social — content shared in a private channel reaching a large enough audience to generate measurable traffic — but without a framework for identifying it, practitioners simply log it as an anomaly. The article must give practitioners a diagnostic process for identifying dark social traffic within direct, so that unexplained spikes become legible rather than mysterious.",
     "r/analytics: \"Direct traffic spike with no obvious cause — how do I diagnose this?\" — analysts describing the classic dark social signature in GA4 and looking for a diagnostic framework: reddit.com/r/analytics/search/?q=direct+traffic+spike+unexplained | r/marketing: \"Our content went 'viral' but GA shows it as direct traffic — what happened?\" — marketers encountering dark social for the first time after a sharing event: reddit.com/r/marketing/search/?q=direct+traffic+content+viral",
     "Hootsuite 2023: unexplained direct traffic spikes are the most common first signal that a brand's content is circulating in dark social channels; the spike is real engagement that analytics fails to attribute: hootsuite.com/blog/dark-social/ | Google Analytics Help: GA4 classifies traffic as \"direct\" when no referral source can be identified — a category that includes bookmarks, typed URLs, and all private sharing channel traffic with stripped referrers: support.google.com/analytics/"),
    ("2",
     "Clients who dismiss unmeasured impact — \"if we can't track it, it doesn't count\"",
     "Social media managers who understand that dark social is generating real audience impact face a reporting credibility problem: clients who operate on a pure attribution model will not accept traffic or engagement they cannot trace to a specific campaign dollar. Practitioners who cannot explain dark social in plain language, quantify it as a percentage of direct traffic, or present proxy metrics lose the budget conversation before it starts. The article must give agencies a client-ready explanation of dark social and a reporting framework that makes the invisible legible without claiming false precision.",
     "r/agency: \"Client insists on 100% attribution — how do you explain dark social without sounding defensive?\" — agency owners navigating client conversations about unmeasured impact: reddit.com/r/agency/search/?q=client+attribution+dark+social | r/SocialMediaManagers: \"How do you justify content ROI when so much traffic shows as direct?\" — SMMs struggling with attribution reporting for clients: reddit.com/r/SocialMediaManagers/search/?q=direct+traffic+attribution+client",
     "Salesforce 2024 State of Marketing: measurement and attribution are cited as the top-two barriers to demonstrating marketing ROI — dark social is a primary contributor to the attribution gap that makes this so difficult: salesforce.com/research/state-of-marketing/ | Nielsen 2022 Annual Marketing Report: 76% of marketers say they lack confidence in their ability to measure marketing effectiveness across all channels — dark social is a significant contributor to this measurement gap: nielsen.com/insights/"),
    ("3",
     "Inconsistent or absent UTM tagging that makes dark social impossible to recover",
     "UTM parameters are the primary mitigation tool for dark social — a pre-tagged link shared in WhatsApp will preserve its source attribution when clicked, even though WhatsApp strips the referrer. But UTM tagging only works when it is applied consistently, before content is distributed. Teams that apply UTMs ad hoc, or not at all for organic content, cannot retroactively recover dark social attribution data. For agencies managing multiple content producers across client accounts, a shared UTM naming convention and a pre-tagging workflow are essential — and entirely undocumented on the SERP.",
     "r/marketing: \"Our team uses UTMs inconsistently — some links tagged, most not. How do you enforce this?\" — marketers describing the UTM discipline problem in content teams: reddit.com/r/marketing/search/?q=utm+parameters+team+consistency | r/analytics: \"Half our organic traffic is untagged — what's the best way to build a UTM process?\" — analysts trying to fix retroactive attribution gaps: reddit.com/r/analytics/search/?q=utm+tagging+process+team",
     "Google Analytics Help: UTM parameters override the referrer stripping that occurs in private sharing channels — a WhatsApp-shared link with ?utm_source=whatsapp&utm_medium=dark-social will be correctly attributed in GA4, but only if the parameter is present before sharing: support.google.com/analytics/ | HubSpot 2023: companies that implement a documented UTM naming convention report a 30–40% reduction in unattributed direct traffic within 90 days of consistent application — directional finding, verify at time of writing: blog.hubspot.com/marketing/"),
    ("4",
     "Underestimating B2B dark social — the invisible influence of Slack and email chains on buying decisions",
     "B2B social media managers and agencies consistently underestimate how much of their audience's actual engagement with brand content happens in private professional channels — Slack workspaces, Microsoft Teams channels, internal email threads — before any tracked conversion event occurs. A buying committee of five people may all read the same vendor whitepaper shared in a Slack channel, but only one of them triggers a tracked session because only one clicked the link — and that session shows as direct. The article must name this dynamic explicitly and provide the proxy metrics (branded search lift, sales-reported referrals, direct-to-deep-link patterns) that make B2B dark social measurable.",
     "r/b2bmarketing: \"Our pipeline is growing but attribution shows almost nothing — where is the actual influence happening?\" — B2B marketers describing the dark social influence problem in long buying cycles: reddit.com/r/b2bmarketing/search/?q=attribution+pipeline+dark+social | r/marketing: \"Slack and Teams are killing our attribution — anyone else dealing with this?\" — practitioners identifying internal messaging tools as dark social channels: reddit.com/r/marketing/search/?q=slack+teams+attribution+dark+social",
     "Forrester 2023: B2B buying committee members consume an average of 13 pieces of content before initiating vendor contact — most of this consumption is undocumented in seller analytics because it occurs in private channels: forrester.com | Gartner 2023: only 17% of the B2B buying journey is spent meeting with potential suppliers — the remaining 83% is independent research, much of it shared privately and invisible to vendor analytics: gartner.com/en/sales/insights/b2b-buying-journey"),
    ("5",
     "No consistent method for estimating the size of the dark social gap in their analytics",
     "Most practitioners know dark social is inflating their direct traffic but have no way to estimate by how much. Without a baseline estimate, they cannot report the gap to clients, set realistic attribution coverage targets, or assess whether their UTM improvements are working. The dark social audit — using GA4 to identify deep-link direct sessions, compare against benchmarked direct traffic rates, and estimate the residual dark social volume — is the missing middle step between understanding the concept and managing it operationally. No article on the SERP provides this audit process.",
     "r/analytics: \"What percentage of direct traffic should I expect to be dark social?\" — analysts looking for a benchmark to assess their own dark social exposure: reddit.com/r/analytics/search/?q=dark+social+percentage+direct+traffic | r/SocialMediaManagers: \"Is there any way to estimate how much dark social traffic we're getting without specialist tools?\" — SMMs looking for a low-cost audit approach: reddit.com/r/SocialMediaManagers/search/?q=dark+social+estimate+analytics",
     "RadiumOne (Signal) 2016 research: 84% of outbound sharing occurs through dark social channels rather than public social networks — the most widely cited dark social benchmark; use as directional context only; methodology is not fully disclosed: radiumone.com | GetSocial research: approximately 78% of online sharing globally occurs via dark social channels, including messaging apps and email — directional benchmark; verify primary source at time of writing: getsocial.io"),
    ("6",
     "Agencies losing client confidence after GA4 migration changed the direct traffic baseline",
     "The migration from Universal Analytics to GA4 changed how sessions are counted, how attribution windows are calculated, and how referral exclusions work — causing many brands to see their direct traffic share increase significantly without any actual change in user behaviour. Agencies that cannot explain the GA4 attribution model change are losing client confidence when direct traffic appears to spike and conversions appear to drop. Dark social was already a measurement problem under UA; GA4 has made it more visible and more confusing for practitioners who have not updated their diagnostic approach.",
     "r/analytics: \"Direct traffic tripled after switching to GA4 — is this dark social or a GA4 issue?\" — practitioners confusing GA4 attribution changes with genuine dark social increases: reddit.com/r/analytics/search/?q=ga4+direct+traffic+spike | r/GoogleAnalytics: \"How do I separate GA4 attribution model changes from actual dark social traffic?\" — analysts trying to isolate the GA4 effect from the dark social effect in their direct traffic baseline: reddit.com/r/GoogleAnalytics/search/?q=ga4+dark+social+direct",
     "Google Analytics Help Center: GA4 uses a different session and attribution model than Universal Analytics — direct traffic calculations are not comparable across the two platforms without adjustment, which practitioners should disclose when presenting historical trend data to clients: support.google.com/analytics/ | Simo Ahava (Google Tag Manager expert): documents specific GA4 attribution model differences that affect direct traffic reporting — a citable practitioner source for the GA4-specific dark social diagnostic: simoahava.com"),
])

# ── KEY STATS FOR WRITER ──────────────────────────────────────────────────────
section_banner(doc, "KEY STATS FOR WRITER")
body_para(doc,
    "Use only the statistics below. Every stat is source-attributed. Stats marked DO NOT USE failed adversarial verification or lack a traceable primary source. Do not add statistics not listed here without editorial fact-check.",
    before=0, after=40)

sub_banner(doc, "Use with confidence:")
for bold, rest in [
    ("84% of outbound sharing occurs through dark social channels rather than public social networks",
     " — RadiumOne (now Signal) 2016 research. The most widely cited dark social benchmark. Use as directional context only; the study methodology is not fully disclosed. Do not present as current or independently verified. Cite as RadiumOne 2016. Source: radiumone.com (archived)"),
    ("The term \"dark social\" was coined by journalist Alexis Madrigal in an article published in The Atlantic in October 2012",
     " — a citable historical fact with a primary source. Use when introducing the concept. Source: The Atlantic, October 2012"),
    ("76% of marketers say they lack confidence in their ability to measure marketing effectiveness across all channels",
     " — Nielsen 2022 Annual Marketing Report. Use to contextualise the dark social measurement problem within the broader attribution challenge. Cite with year. Source: nielsen.com/insights/"),
    ("Only 17% of the B2B buying journey is spent meeting with potential suppliers",
     " — Gartner 2023. Use to establish that most of the B2B buying journey — including content sharing and evaluation — happens outside tracked touchpoints. Cite with year. Source: gartner.com/en/sales/insights/b2b-buying-journey"),
    ("B2B buying committee members consume an average of 13 pieces of content before initiating vendor contact",
     " — Forrester 2023. Use to establish the volume of untracked content engagement in B2B dark social. Cite with year. Verify this figure is current at time of writing. Source: forrester.com"),
    ("Measurement and attribution are cited as the top-two barriers to demonstrating marketing ROI",
     " — Salesforce 2024 State of Marketing. Use to validate the broader attribution challenge that dark social contributes to. Cite with year. Source: salesforce.com/research/state-of-marketing/"),
    ("WhatsApp has 2 billion+ monthly active users as of 2024",
     " — Meta platform data. Use to establish WhatsApp as the most significant dark social channel by scale. Cite as Meta-sourced with year. Source: about.meta.com"),
    ("GA4 uses a different session and attribution model than Universal Analytics, making historical direct traffic comparisons unreliable without adjustment",
     " — Google Analytics Help Center. Use as a factual anchor for the GA4 dark social diagnostic section. Source: support.google.com/analytics/"),
]:
    stat_bullet(doc, bold, rest)

sub_banner(doc, "DO NOT USE — Failed Verification or No Primary Source:")
for item in [
    "\"78% of online sharing is dark social\" — GetSocial figure widely cited but methodology is from a vendor with a commercial interest in the statistic. Use directional framing only; do not cite the specific percentage.",
    "\"Dark social accounts for [X]% of your direct traffic\" — any version of this claim without a disclosed study methodology and sample. Do not cite specific percentages without a primary source.",
    "\"Brands lose $[X] billion in attributable revenue to dark social\" — no primary study with traceable methodology. Circulates in vendor marketing materials. Do not use.",
    "Any claim that a specific percentage of \"direct\" traffic is definitionally dark social without qualifying that the proportion varies by industry, brand, and content type.",
    "Any statistics about dark social's impact on conversion rates or revenue without a disclosed study methodology — these figures circulate from vendor case studies, not independent research.",
    "\"RadiumOne found that 82% [or similar variation] of content sharing happens on dark social\" — the RadiumOne figure is commonly misquoted; the disclosed figure is 84% of outbound sharing, not 82% or other variants. Use the 84% figure and cite RadiumOne 2016 with the caveat about methodology.",
]:
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
    "DARK SOCIAL DEFINITION AND FOUNDATIONS": [
        "Dark social defined in one plain sentence on first use, with attribution to Alexis Madrigal and The Atlantic (October 2012)",
        "The article explains why dark social traffic appears as direct traffic — referrer stripping in private sharing channels — in plain language without assuming analytics expertise",
        "The difference between true direct traffic and dark social traffic is clearly explained: true direct = typed URL or bookmark; dark social direct = shared link with stripped referrer",
        "The channel comparison table is included: channel name, whether it strips referrer data, dark social risk level, and mitigation tactic — covering at minimum WhatsApp, iMessage, Telegram, Signal, Email, Slack, Microsoft Teams, Discord, Facebook Messenger, LinkedIn DMs",
        "The article acknowledges that 100% dark social attribution is not achievable — the goal is estimation, proxy measurement, and improved UTM hygiene, not perfect tracking",
    ],
    "DARK SOCIAL AUDIT PROCESS": [
        "A step-by-step dark social audit is included with a minimum of 5 numbered steps — not general advice",
        "Specific GA4 report paths are named (e.g., which exploration report, what dimensions and metrics) so the reader can execute the audit without guessing",
        "The article defines what a dark social traffic signature looks like in GA4: high direct traffic to deep-link pages, no campaign source, often correlates with a content publish or social share event",
        "The audit includes a method for estimating the dark social share of direct traffic — even a directional estimate with acknowledged uncertainty is more useful than no estimate",
        "The GA4 migration effect on direct traffic baselines is addressed — readers should understand whether their increased direct traffic is a GA4 attribution change or genuine dark social growth before drawing conclusions",
    ],
    "UTM ARCHITECTURE": [
        "A specific UTM naming convention is provided for dark social-prone channels — not just a mention of UTM parameters",
        "Source and medium naming conventions are specified for the most common dark social channels: whatsapp, email, slack, sms, telegram",
        "The section identifies which content types are highest-priority for pre-tagging: research reports, long-form guides, case studies, pricing pages, and any content likely to be shared in professional contexts",
        "A team-wide UTM consistency approach is recommended — a shared UTM builder spreadsheet, a pre-publish tagging checklist, or a link management tool",
        "The article explains that UTMs only help if applied before sharing — retroactive UTM application does not recover already-lost attribution data",
    ],
    "B2B DARK SOCIAL": [
        "B2B dark social is addressed in its own named section — not as a passing mention within a broader channel list",
        "The section explains why B2B buying journeys are disproportionately affected: buying committees share vendor content in private channels before any tracked engagement",
        "Specific B2B dark social channels are named: Slack, Microsoft Teams, internal email chains, private LinkedIn DMs, and buying committee group chats",
        "B2B dark social proxy metrics are included: branded search volume trend, sales-reported referral sources, direct-to-deep-link page patterns, and any survey-based attribution data",
        "The article connects B2B dark social to the buying committee dynamic — multiple people consuming content, only one generating a tracked session",
    ],
    "CLIENT REPORTING FRAMEWORK": [
        "A specific client reporting framework is included — not general advice about \"being transparent\" about dark social",
        "The framework includes: (a) how to estimate dark social traffic as a percentage of direct; (b) three proxy metrics to present: branded search trend, survey attribution rate, share-button click data; (c) a one-paragraph plain-language explanation for client decks",
        "The article sets realistic client expectations: dark social measurement produces estimates, not exact figures, and any agency that claims otherwise is overstating their measurement capability",
        "The agency reader is addressed explicitly — the framework should be applicable across different client types and industries without being rebuilt from scratch each time",
    ],
    "METRICS AND MEASUREMENT": [
        "At least 4 measurement levers defined: UTM tag coverage rate, dark social traffic estimate (% of direct), branded search volume trend, and survey-based attribution rate",
        "Each measurement lever is connected to a business outcome or reporting use case the client or stakeholder cares about",
        "A suggested measurement cadence is included: what to track weekly (UTM coverage, direct traffic pattern) vs. monthly (dark social estimate, branded search trend)",
        "The article acknowledges the measurement limitations honestly — this is a strength, not a weakness; clients trust agencies who understand what they cannot measure",
    ],
    "STRUCTURE AND READABILITY": [
        "Each H2 opens with a 2–3 sentence summary that works as a standalone answer",
        "Paragraphs are 3–4 lines maximum throughout the article",
        "The channel comparison section uses a table, not prose",
        "The dark social audit is a numbered list, not a prose paragraph",
        "At least one template or checklist element is included: a UTM naming convention template, a dark social audit checklist, or a client reporting one-pager framework",
        "No section is generic enough that it could appear in any of the competing articles listed above",
    ],
    "FAQ SECTION": [
        "Each answer is 40–60 words",
        "Each answer starts directly with the answer — no preamble, no restatement of the question",
        "Each answer reads as a complete, standalone response suitable for AI Overview extraction",
        "Question wording matches PAA phrasing from SERP research above",
        "\"What is the difference between dark social and direct traffic?\" is answered with the specific technical explanation (referrer stripping, deep-link landing patterns) not a vague distinction",
        "\"How much of my traffic is dark social?\" is answered with an honest directional benchmark and a reference to the audit process, not a false-precision percentage",
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
        ["1", "socialpilot.co/blog/social-media-marketing-strategy",
         "build a complete social media marketing strategy",
         "Introduction — establishing dark social as a measurement gap within the broader social media strategy, not an isolated analytics problem"],
        ["2", "socialpilot.co/blog/social-media-analytics",
         "measure social media performance across channels",
         "Analytics foundation section — when explaining how dark social inflates direct traffic and deflates attributed channel performance"],
        ["3", "socialpilot.co/blog/utm-parameters",
         "use UTM parameters to track social media traffic",
         "UTM architecture section — linking to the detailed UTM guide when introducing the parameter naming convention"],
        ["4", "socialpilot.co/blog/social-media-for-agencies",
         "manage dark social measurement across multiple client accounts",
         "Agency section — connecting the dark social framework to multi-client analytics and reporting workflow"],
        ["5", "socialpilot.co/blog/social-media-reporting",
         "build a client social media report that includes dark social",
         "Client reporting section — linking to the broader social media reporting guide when introducing the client-facing framework"],
        ["6", "socialpilot.co/blog/instagram-dm-marketing",
         "Instagram DMs as a dark social sharing channel",
         "Channel comparison section — when covering Instagram DMs as a high-risk dark social channel for brand content"],
        ["7", "socialpilot.co/blog/whatsapp-marketing",
         "WhatsApp as a brand communication and dark social channel",
         "Channel section — when identifying WhatsApp as the highest-volume dark social sharing platform globally"],
    ]
)

sub_banner(doc, "Internal Link Rules for the Writer:")
for item in [
    "Do not use the same anchor text twice in one post.",
    "Do not link to a competitor from within the post body (Sprout Social, Hootsuite, Buffer, WordStream, Adweek, etc.).",
    "First internal link appears after the first H2, not in the introduction.",
    "Every internal link must be contextually relevant, never forced.",
    "Anchor text must be 3–5 words that are contextually meaningful in the sentence.",
    "Verify all internal link URLs are live before submitting the draft — URL slugs may have changed since this brief was written.",
]:
    bullet(doc, item)

# ── Save ──────────────────────────────────────────────────────────────────────
out = "/home/user/new/CONTENT_BRIEF__SocialPilot__Dark_Social_v1.docx"
doc.save(out)
print("Saved:", out)
