from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

for section in doc.sections:
    section.top_margin    = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin   = Inches(0.8)
    section.right_margin  = Inches(0.8)

BLACK   = RGBColor(0x14, 0x14, 0x13)
DARK    = RGBColor(0x3D, 0x3D, 0x3A)
BLUE    = RGBColor(0x1A, 0x56, 0xDB)
HEAD_BG = "1A56DB"
HEAD_FG = RGBColor(0xFF, 0xFF, 0xFF)
ROW_BG  = "EEF2FF"
ALT_BG  = "FAFAFA"

def set_cell_bg(cell, hex_color):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    for s in tcPr.findall(qn('w:shd')): tcPr.remove(s)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def cell_para(cell, text, bold=False, size=9, color=BLACK, italic=False, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = cell.paragraphs[0]; p.clear(); p.alignment = align
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text); run.bold = bold; run.italic = italic
    run.font.size = Pt(size); run.font.color.rgb = color
    return p

def add_run(para, text, bold=False, size=10, color=BLACK, italic=False):
    run = para.add_run(text); run.bold = bold; run.italic = italic
    run.font.size = Pt(size); run.font.color.rgb = color
    return run

def body_para(text='', bold=False, size=10, color=BLACK, space_before=4, space_after=4, indent_pt=0):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(space_before); p.paragraph_format.space_after = Pt(space_after)
    if indent_pt: p.paragraph_format.left_indent = Pt(indent_pt)
    if text:
        run = p.add_run(text); run.bold = bold; run.font.size = Pt(size); run.font.color.rgb = color
    return p

def section_header(text):
    tbl = doc.add_table(rows=1, cols=1); tbl.style = 'Table Grid'
    cell = tbl.rows[0].cells[0]; set_cell_bg(cell, HEAD_BG)
    p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text); run.bold = True; run.font.size = Pt(11); run.font.color.rgb = HEAD_FG
    body_para(space_before=2, space_after=2)

def sub_header(text):
    p = body_para(space_before=8, space_after=3)
    run = p.add_run(text); run.bold = True; run.font.size = Pt(10); run.font.color.rgb = BLUE
    return p

def bullet(text, size=9.5):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(1); p.paragraph_format.space_after = Pt(1)
    run = p.add_run(text); run.font.size = Pt(size); run.font.color.rgb = BLACK
    return p

def make_table(headers, rows, col_widths, header_size=9, row_size=8.5):
    tbl = doc.add_table(rows=1+len(rows), cols=len(headers)); tbl.style = 'Table Grid'
    hrow = tbl.rows[0]
    for i, (h, w) in enumerate(zip(headers, col_widths)):
        c = hrow.cells[i]; c.width = w
        set_cell_bg(c, HEAD_BG); cell_para(c, h, bold=True, size=header_size, color=HEAD_FG)
    for ri, rd in enumerate(rows):
        row = tbl.rows[ri+1]
        bg = 'FFFFFF' if ri % 2 == 0 else ALT_BG
        for ci, (val, w) in enumerate(zip(rd, col_widths)):
            c = row.cells[ci]; c.width = w
            set_cell_bg(c, bg); cell_para(c, val, size=row_size)
    return tbl

# --- TITLE ---
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(3)
r = p.add_run('CONTENT BRIEF -- SocialPilot Blog')
r.bold = True; r.font.size = Pt(14); r.font.color.rgb = BLUE

p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_before = Pt(0); p2.paragraph_format.space_after = Pt(10)
r2 = p2.add_run('How Niching Your Social Media Agency Makes You More Money')
r2.bold = True; r2.font.size = Pt(11.5); r2.font.color.rgb = BLACK

# --- BASICS ---
section_header('BASICS')
basics = [
    ('Working Title:',     'How Niching Your Social Media Agency Makes You More Money'),
    ('Final URL Slug:',    '/blog/social-media-agency-niche'),
    ('Content Type:',      'Cluster / Strategic Guide'),
    ('Funnel Stage:',      'MOFU / BOFU'),
    ('Target Word Count:', '3,000-3,600 words'),
    ('Publish Date:',      'Week of [assign date]'),
    ('Writer:',            '[Assign]'),
    ('Editor:',            '[Assign]'),
]
tbl = doc.add_table(rows=len(basics), cols=2); tbl.style = 'Table Grid'
for i, (label, value) in enumerate(basics):
    c0, c1 = tbl.rows[i].cells
    bg = ROW_BG if i % 2 == 0 else 'FFFFFF'
    set_cell_bg(c0, bg); set_cell_bg(c1, bg)
    cell_para(c0, label, bold=True, size=9, color=DARK)
    cell_para(c1, value, size=9)

p_note = body_para(space_before=5, space_after=2)
add_run(p_note, 'Editorial note: ', bold=True, size=9, color=DARK)
add_run(p_note, 'Deep research adversarially verified every quantified claim about niche agency financial advantages circulating across agency-focused content. Every specific figure (23% higher margins, 60% pricing premiums, 64% CAC reduction) failed verification at 0-3 votes -- the sources either do not contain the stated data or cite secondary blogs with no primary research behind them. The financial case for niching is universally asserted but currently unsupported by verifiable primary data. This article must NOT repeat those unverified figures. Instead, it must build the financial argument from structural logic, the single verified Hinge Marketing correlation, and first-person practitioner evidence from r/agency. This approach is genuinely more credible than every competing article and is the editorial angle.', size=9, italic=True, color=DARK)
body_para(space_before=3, space_after=3)

# --- SEO TARGETS ---
section_header('SEO TARGETS')
p = body_para(space_before=5, space_after=2)
add_run(p, 'Primary Keyword: ', bold=True, size=10)
add_run(p, 'how to niche your agency  (est. 480-880/mo, informational, low-competition SERP dominated by strategic advice without verified financial data)', size=10)

p = body_para(space_before=3, space_after=2)
add_run(p, 'SEO Note for Editor: ', bold=True, size=10, color=DARK)
add_run(p, 'The SERP for agency niching queries is dominated by tactical how-to guides and aspirational financial claims that do not hold up to scrutiny. The gap this article fills: the financial case built from structural logic and practitioner evidence rather than recycled statistics. The working title ("makes you more money") is deliberately direct and differentiating. Verify keyword volumes in Ahrefs/Semrush before publishing. This topic is growing in search volume as more agency owners enter the market and face positioning decisions.', size=9, italic=True, color=DARK)

sub_header('Secondary Keywords:')
make_table(
    ['#', 'Secondary Keyword', 'Est. MSV', 'Search Intent'],
    [
        ('1',  'niche marketing agency',                       '720-1,300/mo',  'Commercial Investigation'),
        ('2',  'social media agency niche',                    '320-590/mo',    'Informational / Commercial'),
        ('3',  'should I niche my agency',                     '200-480/mo',    'Informational'),
        ('4',  'agency niching strategy',                      '200-480/mo',    'Informational'),
        ('5',  'niche agency vs generalist',                   '200-480/mo',    'Informational'),
        ('6',  'how to choose an agency niche',                '200-480/mo',    'Informational'),
        ('7',  'how to find your agency niche',                '200-480/mo',    'Informational'),
        ('8',  'specialize social media agency',               '200-480/mo',    'Informational'),
        ('9',  'social media agency positioning',              '200-480/mo',    'Commercial Investigation'),
        ('10', 'how to transition from generalist to niche',   '200-480/mo',    'Informational'),
    ],
    [Inches(0.28), Inches(2.65), Inches(1.05), Inches(2.82)]
)

p3 = body_para(space_before=8, space_after=2)
add_run(p3, 'People Also Ask (from SERP):', bold=True, size=10)
paas = [
    'Should I niche down my agency?',
    'What are the best niches for a social media agency?',
    'How do I choose a niche for my marketing agency?',
    'Is it better to be a generalist or specialist agency?',
    'How do I transition from a generalist to a niche agency?',
    'What niches are most profitable for social media agencies?',
    'How do I validate my agency niche before committing?',
    'Can a generalist agency compete with niche agencies?',
    'How long does it take to establish a niche agency?',
]
for paa in paas: bullet(paa)
body_para(space_before=3, space_after=3)

# --- AUDIENCE ---
section_header('AUDIENCE')
p = body_para(space_before=5, space_after=2)
add_run(p, 'ICP Segment: ', bold=True, size=10)
add_run(p, 'Agency (ICP 1) -- primary; specifically agency owners and founders, not hired social media managers', size=10)

p = body_para(space_before=3, space_after=2)
add_run(p, 'Buying Stage: ', bold=True, size=10)
add_run(p, 'Problem Aware to Solution Aware -- the reader is running a generalist agency, experiencing at least one of: low margins, price competition, exhausting sales cycles, high client churn, or difficulty standing out. They are aware niching is an option but are held back by specific fears.', size=10)

p = body_para(space_before=5, space_after=2)
add_run(p, 'What the Reader Already Knows:', bold=True, size=10)
bullet('They run or are growing a social media agency and currently take clients from multiple verticals')
bullet('They have heard the advice to "niche down" and are either skeptical, scared, or unsure how to do it')
bullet('They have at least one vertical where they feel more confident, get better results, or receive more referrals')
bullet('They understand what a retainer, a pitch, and a client churn rate are -- no need to define these')

p = body_para(space_before=5, space_after=2)
add_run(p, 'What the Reader Wants to Walk Away With:', bold=True, size=10)
bullet('A clear, evidence-based answer to "will niching actually make me more money?" -- not hype, not unverified statistics')
bullet('A framework for identifying the right niche from their existing client base rather than guessing')
bullet('A practical transition plan that does not require firing all their current clients immediately')
bullet('Confidence that the financial upside is real even without specific benchmark figures -- and an understanding of why the structural logic holds')
bullet('Permission to claim a niche even before they feel like an undisputed expert in it')

p = body_para(space_before=5, space_after=2)
add_run(p, 'Distinguished Solution / Purpose of This Article:', bold=True, size=10)
p2 = body_para(space_before=2, space_after=4)
add_run(p2, 'Every competing article either repeats unverified statistics about niche agency financial advantages or gives generic strategic advice without explaining the mechanism. This article does neither. It builds the financial case from first principles: why premium pricing is structurally available to niche agencies (expertise signal reduces price resistance), why client acquisition costs fall (narrower targeting, stronger referral density, clearer positioning), and why margins improve (less scope creep, fewer non-ICP clients that drain the team). The only verified primary research cited is the Hinge Marketing High Growth Study correlation -- which says more than any fabricated percentage because it comes from 50,000+ professional services firms over two decades. The rest of the argument is built from structural logic and real practitioner testimony from r/agency threads that the research workflow verified and sourced.', size=10)
body_para(space_before=3, space_after=3)

# --- SERP & COMPETITIVE ANALYSIS ---
section_header('SERP & COMPETITIVE ANALYSIS')

sub_header('A. WHO IS RANKING AND WHY')
p = body_para(space_before=3, space_after=4)
add_run(p, 'The SERP for "how to niche your agency" and related queries is dominated by Agency Analytics and Databox -- two agency software companies with high DR and regular content programs. Both publish strategic advice articles with aspirational financial claims that failed adversarial verification. The content gap is not tactical coverage but credibility: no article on the SERP makes the financial case with honest sourcing. That is the position this article takes.', size=9, italic=True, color=DARK)

make_table(
    ['#', 'URL', 'Domain', 'DR', 'Word Count', 'Last Updated', 'Format', 'Article Angle', 'Why It Ranks'],
    [
        ('1',
         'agencyanalytics.com/blog/niche-agency',
         'Agency Analytics', '74', '~3,200', '2025',
         'Strategic guide',
         'Pros, cons, and how-to for agency niching; covers niching dimensions and transition steps; includes unverified financial figures',
         'High DR + comprehensive coverage; frequently linked in agency communities; the financial figures cited are unverifiable secondary statistics'),
        ('2',
         'agencyanalytics.com/blog/how-to-find-a-niche',
         'Agency Analytics', '74', '~2,800', '2025',
         'Step-by-step tactical guide',
         'Niche-finding process from existing client audit through market validation; strong on process, weak on financial rationale',
         'Same domain authority; ranks for long-tail "how to find niche" queries; good tactical content but no honest sourcing on financial claims'),
        ('3',
         'databox.com/agencies-choose-niche',
         'Databox', '80', '~2,600', '2024',
         'Survey-based article',
         'Headline promises data ("here\'s what the data says"); actual data is unverified secondary statistics recycled from other blogs; no disclosed methodology',
         'High DR + credibility signalling from headline; ranks for commercial investigation queries; misleading data presentation is a quality weakness'),
        ('4',
         'doubleyourfreelancing.com/the-fear-of-niching/',
         'Double Your Freelancing', '58', '~2,400', '2023',
         'Psychology/motivation article',
         'Addresses the emotional barriers to niching with practitioner case studies; strong on fear analysis, does not make a financial argument',
         'Strong on-page signals; well-linked in freelancing communities; the fear framing is useful but the financial mechanism is absent'),
        ('5',
         'wordstream.com/blog/niche-agency',
         'WordStream', '82', '~1,800', '2024',
         'Tactical overview',
         'How to build niche agency positioning; covers service packaging and messaging; no financial benchmarks',
         'Brand authority; ranks broadly on marketing agency queries; thinnest content of the top results'),
    ],
    [Inches(0.28), Inches(1.65), Inches(0.65), Inches(0.32), Inches(0.68), Inches(0.75), Inches(0.85), Inches(1.48), Inches(1.34)]
)

sub_header('B. TABLE STAKES vs. COVERAGE GAPS')

p = body_para(space_before=4, space_after=2)
add_run(p, 'Table Stakes', bold=True, size=10)
p2 = body_para(space_before=1, space_after=3)
add_run(p2, 'Required to satisfy search intent and avoid appearing incomplete.', size=9, italic=True, color=DARK)
table_stakes = [
    'A clear definition of what "niching" means in the context of a social media agency',
    'The five documented approaches to niching: by vertical, by platform, by content format, by business stage, by geography',
    'A framework for identifying the right niche from existing client data',
    'How to validate a niche before fully committing',
    'How to manage the transition from generalist to niche without losing all current revenue',
    'A section addressing the most common fears: turning away clients, niche too small, imposter syndrome',
    'FAQ section covering PAA questions above',
]
for ts in table_stakes: bullet(ts)

p = body_para(space_before=8, space_after=2)
add_run(p, 'Coverage Gaps', bold=True, size=10)
p2 = body_para(space_before=1, space_after=3)
add_run(p2, 'What no top-ranking page covers. These are the angles that make this article genuinely new on the SERP.', size=9, italic=True, color=DARK)

gaps = [
    ('Gap 1 -- No article makes the financial case honestly:',
     'Every ranking article asserting financial benefits of niching uses unverified secondary statistics that did not survive adversarial fact-checking. The agency content ecosystem recycles figures (23% higher EBITDA, 60% pricing premiums, 64% CAC reduction) that have no traceable primary source. This article is the first on the SERP to acknowledge this directly and make the financial argument from structural logic instead: premium positioning reduces price resistance, narrower targeting reduces CAC, deeper expertise reduces scope creep, and stronger referral density from a defined community compounds over time. The argument is more persuasive because it is honest about what is verified and what is structural logic.'),
    ('Gap 2 -- No article addresses the revenue dip during transition:',
     'Every tactical guide describes niching as a positive strategic move without acknowledging what multiple r/agency practitioners confirm: the transition period involves a genuine dip in revenue as generalist clients churn or get wound down and niche pipeline builds. Practitioners who try to niche without financial resilience for a 3-6 month transition often revert to generalism. This article addresses the transition honestly -- including what financial preparation the agency owner needs before committing -- which is both more useful and more trustworthy than articles that skip this part.'),
    ('Gap 3 -- No article explains the mechanism behind the financial benefits:',
     'Competing articles assert that niche agencies earn more without explaining why. This article explains the three mechanisms: (1) expertise signal -- when a prospect believes you have done exactly their type of work before, price resistance drops because the perceived risk of switching is lower; (2) referral density -- a defined ICP community refers within itself, meaning each client is a marketing investment in a concentrated network rather than a one-off transaction; (3) scope creep reduction -- niche agencies have more precisely scoped retainers because they have seen the same client problems repeatedly and know what is and is not in scope.'),
    ('Gap 4 -- No article is written specifically for social media agency owners:',
     'Most niching articles target general marketing agencies or freelancers. The specific dynamics of social media agency niching -- by platform (LinkedIn specialist, TikTok agency), by content type (video-first agency, UGC agency), and by vertical with platform implications (SaaS brands on LinkedIn, e-commerce brands on Instagram) -- are not addressed. SocialPilot can own the social-media-specific framing while competitors write for a generic agency audience.'),
]
for title, desc in gaps:
    p = body_para(space_before=5, space_after=1)
    add_run(p, title, bold=True, size=9.5)
    p2 = body_para(space_before=1, space_after=3, indent_pt=14)
    add_run(p2, desc, size=9, color=DARK)

p = body_para(space_before=8, space_after=2)
add_run(p, 'Format and Quality Weaknesses in Competing Content:', bold=True, size=10)
weaknesses = [
    'All competing articles cite financial figures (margin differentials, pricing premiums, win-rate improvements) that failed adversarial verification -- no traceable primary source exists for any of them',
    'The Databox article headlines as data-driven but the underlying data is secondary blog recycling with no disclosed methodology',
    'No competing article addresses the transition revenue dip -- the most common reason agency owners fail to execute a niching decision',
    'No competing article is written specifically for social media agency owners; the platform-specific niching options (LinkedIn specialist, TikTok-only, UGC agency) are undocumented on the SERP',
    'The fear analysis in Doubleyourfreelancing is strong but does not progress to a practical decision framework for choosing the niche',
    'No competing article provides a client audit template or scoring model for identifying which existing vertical is the strongest niche candidate',
]
for w in weaknesses: bullet(w)
body_para(space_before=3, space_after=3)

# --- LLM CONSIDERATIONS ---
section_header('LLM CONSIDERATIONS')
p = body_para(space_before=4, space_after=4)
add_run(p, 'AI Overviews appear for "should I niche my agency" and "how to niche your agency." They currently pull from Agency Analytics and Databox -- both of which contain unverified statistics. An article that explicitly acknowledges the lack of verified financial data and builds the case from structural logic and Hinge Marketing research will stand out as more trustworthy to LLMs trained to weight source credibility. Write every H2 to open with a complete, standalone answer. The mechanism-based financial argument (expertise signal, referral density, scope creep reduction) is especially likely to be cited because it is specific and original.', size=10)

sub_header('KEY REQUIREMENTS')
requirements = [
    '1.  Do not use any of the following figures anywhere in the article: 23% higher EBITDA, 27-60% pricing premium, 64% CAC reduction, 20% to 35% win-rate improvement, 72% of high-growth agencies are niche. All failed adversarial verification. Do not use these even as "some sources claim" -- that framing still spreads unverified data.',
    '2.  The only verified primary research is the Hinge Marketing High Growth Study correlation: "Year after year, we see data correlating specialization with greater growth and profitability" (drawn from 50,000+ professional services firms over two decades). Cite verbatim with source. Acknowledge it applies to professional services broadly, not social media agencies specifically.',
    '3.  The three financial mechanisms must each be explained in one dedicated paragraph with a concrete example: (a) expertise signal and pricing, (b) referral density within a defined ICP community, (c) scope creep reduction from repeated exposure to the same client problem type.',
    '4.  The transition revenue dip must be addressed honestly. Practitioners on r/agency confirm it is real: 3-6 months of lower revenue is typical before niche pipeline recovers. The article must help the reader prepare for this rather than pretend it does not happen.',
    '5.  The client audit section must include a concrete scoring approach: which existing clients have highest retention, generate most referrals, are most enjoyable to serve, and produce the best results. The intersection of those four criteria is the niche signal.',
    '6.  FAQ answers must be 40-60 words, start directly with the answer, and read as complete standalone responses.',
    '7.  Keep paragraphs to 3-4 lines maximum throughout.',
    '8.  Write 3-4 sentences a social media agency owner could include verbatim in their own internal positioning document or pitch deck. These are the citable claims LLMs will surface.',
]
for r in requirements: bullet(r)

sub_header('Suggested Citable Claims to Build Into the Article:')
claims = [
    'Claim 1: "Hinge Marketing\'s High Growth Study, drawn from more than 50,000 professional services firms over two decades, finds a consistent correlation between specialization and greater growth and profitability. The study does not quantify the premium -- but the direction of the relationship has held across every annual iteration of the research. No equivalent primary study exists specifically for social media agencies, but the structural logic is the same."',
    'Claim 2: "A niche agency does not charge more because it decides to. It charges more because its ICP perceives less risk in hiring it. When a restaurant chain shortlists social media agencies, the agency that has worked exclusively with F&B brands for three years presents a categorically lower risk profile than a generalist that has worked with a restaurant once. Reduced perceived risk is how premium pricing becomes available -- not through positioning statements, but through proof."',
    'Claim 3: "The transition from generalist to niche agency is a 6-18 month process, not a pivot. The approach that works: decline new non-niche clients, honour existing out-of-niche contracts through their current term, and invest all new-business marketing into the chosen vertical. Revenue dips in the first 3-6 months as generalist clients churn. Agencies that try to niche without financial resilience for this period typically revert to generalism before the niche pipeline builds. This is the part most niching advice skips."',
    'Claim 4: "The most reliable signal for niche selection is not passion or market size research -- it is existing client data. The vertical that generates your highest retention, your most unsolicited referrals, your cleanest delivery, and your most satisfied clients is the one you are already good at. Most agency owners discover their niche by looking backward at what they have already proven, not forward at what they might become."',
]
for c in claims: bullet(c)
body_para(space_before=3, space_after=3)

# --- PAIN POINT ANALYSIS ---
section_header('PAIN POINT ANALYSIS')
p = body_para(space_before=4, space_after=4)
add_run(p, 'Six validated pain points agency owners face when considering or executing a niching decision. Reddit signals are verified threads from r/agency and r/Entrepreneur confirmed via Google site search. Other Sources are verified practitioner and industry publications.', size=9, italic=True, color=DARK)

pp_headers = ['#', 'Pain Point', 'Core Finding', 'Reddit Signal + URL', 'Other Sources']
pp_rows = [
    ('1',
     'Fear of turning away paying clients -- agency owners worry niching means immediate revenue loss',
     'The most frequently cited barrier to niching is not strategic uncertainty but financial fear: the prospect of declining a paying client feels immediately costly even when the long-term case for niching is accepted. Agency owners in active growth mode, or those carrying fixed costs that require consistent revenue, find the idea of saying no to inbound work psychologically difficult even when the client is outside the target niche. The article must address this fear directly and give the reader a practical framework for when it is safe to start declining non-niche work -- tied to pipeline coverage ratio and financial runway, not just positioning conviction.',
     'r/agency: "Should I niche down my agency?" -- agency owners debating whether the revenue risk of niching is worth it; multiple responses confirm the fear of turning away clients as the primary barrier: reddit.com/r/agency/comments/1gf77o3/should_i_niche_down_my_agency/ | r/agency: "Those who have niched down their agency, was it worth it?" -- post-niching practitioners confirming the transition dip is real but the long-term outcome is positive: reddit.com/r/agency/comments/1cwt5xg/those_who_have_niched_down_their_agency_was_it/',
     'Doubleyourfreelancing.com: structured analysis of niching fear psychology with practitioner case studies; documents the fear of revenue loss as the single most common barrier: doubleyourfreelancing.com/the-fear-of-niching/ | Hinge Marketing High Growth Study: specialization correlates with greater growth and profitability across professional services firms -- the long-term financial case for accepting short-term revenue discipline: hingemarketing.com/blog/story/to-specialize-or-not-to-specialize'),

    ('2',
     'Not knowing which niche to choose -- paralysis between following passion, following the money, or following existing clients',
     'Agency owners who accept the case for niching often stall on the choice itself. Common patterns: choosing a niche based on personal interest without validating whether there is sufficient client demand; choosing based on perceived market size without evaluating whether they can actually win in that vertical; or overthinking the decision entirely and deferring it indefinitely. The research consistently shows the most reliable niche signal is existing client data -- not passion or market research. The article must give the reader a concrete audit they can run on their current client base to surface the niche they have already been building without realising it.',
     'r/agency: "Share your experience about niching down your agency" -- practitioners describing their niche selection process; multiple accounts confirm existing client patterns as the clearest signal: reddit.com/r/agency/comments/1la56x4/share_your_experience_about_niching_down_your_agency/ | r/Entrepreneur: "Scared to niche down my freelance business, I don\'t [know how to choose]" -- freelancers and agency owners describing the decision paralysis: reddit.com/r/Entrepreneur/comments/1b0pq1r/scared_to_niche_down_my_freelance_business_i_dont/',
     'Agency Analytics: "How to Find Your Agency Niche" -- documents the existing client audit as the recommended first step; confirms retention, referral rate, and delivery quality as the key scoring criteria: agencyanalytics.com/blog/how-to-find-a-niche | Hinge Marketing: specialization selection guidance recommends choosing based on demonstrated competence and existing market position, not aspirational positioning: hingemarketing.com/blog/story/to-specialize-or-not-to-specialize'),

    ('3',
     'Managing the transition -- how to wind down generalist clients without a revenue cliff',
     'Even agency owners who have chosen their niche and committed to it face a practical operational problem: what to do with the existing out-of-niche clients who are currently paying the bills. Abruptly firing non-niche clients destroys revenue before the niche pipeline is built. But continuing to take non-niche clients undermines positioning and keeps the agency stuck in generalism. The solution documented by practitioners is a phased transition: honour existing contracts through their current term, decline renewals outside the niche, and reinvest all new-business effort into the chosen vertical. The article must give the reader a concrete transition timeline and decision framework, not just abstract advice to "gradually transition."',
     'r/agency: "Those who have niched down their agency, was it worth it?" -- practitioners describing how they managed the transition; multiple accounts confirm the phased approach and document the 3-6 month revenue dip before recovery: reddit.com/r/agency/comments/1cwt5xg/those_who_have_niched_down_their_agency_was_it/ | r/agency: "Should I niche down my agency?" -- current agency owners asking specifically about how to handle existing out-of-niche clients: reddit.com/r/agency/comments/1gf77o3/should_i_niche_down_my_agency/',
     'Agency Analytics: operational transition guidance documents the phased model as standard practice: agencyanalytics.com/blog/niche-agency | Doubleyourfreelancing.com: case studies of agencies that transitioned successfully; confirms the 6-18 month timeline and the importance of financial runway before committing: doubleyourfreelancing.com/the-fear-of-niching/'),

    ('4',
     'Imposter syndrome -- agency owners feel they need to already be the leading expert before claiming a niche',
     'A significant subset of agency owners resist niching not because of financial fear but because of identity uncertainty: they do not feel qualified to position themselves as specialists before they have achieved recognised expertise. This is a cognitive distortion that the article must address directly. The niche claim does not require being the industry\'s foremost authority -- it requires being more focused than generalist competitors, which is immediately true the moment the agency stops pitching to everyone. Practitioners on r/Entrepreneur confirm this fear is common and that most agency owners discover their claimed expertise becomes real through the process of serving niche clients exclusively, not before it.',
     'r/Entrepreneur: "Scared to niche down my freelance business, I don\'t [feel qualified enough]" -- practitioners describing the imposter syndrome barrier to niching; thread confirms the fear is common and largely dissolves after committing: reddit.com/r/Entrepreneur/comments/1b0pq1r/scared_to_niche_down_my_freelance_business_i_dont/ | r/agency: "Share your experience about niching down your agency" -- post-niching practitioners confirming they did not feel like experts before niching but became experts faster after narrowing focus: reddit.com/r/agency/comments/1la56x4/share_your_experience_about_niching_down_your_agency/',
     'Doubleyourfreelancing.com: documents imposter syndrome as the second most common niching barrier after revenue fear; addresses it with practitioner case studies: doubleyourfreelancing.com/the-fear-of-niching/ | Hinge Marketing: specialization builds perceived expertise faster than generalism because the market has fewer reference points to compare against; the expert signal is partly a function of narrowness, not just depth: hingemarketing.com/blog/story/to-specialize-or-not-to-specialize'),

    ('5',
     'Worry that the chosen niche is too small to sustain the agency long-term',
     'Agency owners who have accepted the case for niching and identified a candidate vertical often stall on market size: is there enough demand in this niche to fill a full client roster and grow? The concern is legitimate but usually miscalibrated. Most social media agencies serving 10-20 clients can fill their entire capacity from a national (or even regional) niche vertical without needing a market of millions. The article must give the reader a concrete market sizing check: a niche that can sustain a 10-client roster at target retainer value is large enough. The risk of choosing too narrow a niche (hyper-local, single obsoleting platform, micro-segment) is real but is different from the generalised worry about niche size.',
     'r/agency: "Should I niche down my agency?" -- thread where market size concern is explicitly raised by multiple commenters; responses distinguish between genuinely too-small niches and unfounded size anxiety: reddit.com/r/agency/comments/1gf77o3/should_i_niche_down_my_agency/ | r/agency: "Share your experience about niching down your agency" -- practitioners confirming that niche market size concern was unfounded in most cases; the limiting factor was not market size but pipeline execution: reddit.com/r/agency/comments/1la56x4/share_your_experience_about_niching_down_your_agency/',
     'Agency Analytics: niche selection guide documents the market size validation step; recommends confirming that enough potential clients exist in the vertical before committing: agencyanalytics.com/blog/how-to-find-a-niche | Hinge Marketing: notes that "going too narrow can constrain growth" as a documented specialization risk; the risk is real for hyper-narrow niches but not for most industry vertical choices: hingemarketing.com/blog/story/to-specialize-or-not-to-specialize'),

    ('6',
     'Difficulty charging premium rates in a niche when clients can still find generalist alternatives at lower prices',
     'Agency owners who niche and attempt to raise prices often find that clients push back by citing cheaper generalist alternatives. The premium pricing case fails when the agency cannot articulate why niche expertise is worth the difference -- and falls back on positioning statements ("we specialise in your industry") without proof. The article must explain the mechanism through which niche expertise commands premium pricing: it is not the claim but the proof -- case studies in the specific vertical, a roster of recognisable names in the niche, and the ability to speak to client problems without being briefed. Premium pricing is available when the prospect believes the risk of hiring a generalist is higher than the price differential of hiring the specialist.',
     'r/agency: "Share your experience about niching down your agency" -- practitioners discussing pricing power post-niching; accounts confirm that premium pricing required demonstrating niche proof, not just claiming specialisation: reddit.com/r/agency/comments/1la56x4/share_your_experience_about_niching_down_your_agency/ | r/agency: "Those who have niched down their agency, was it worth it?" -- post-niching practitioners reporting that pricing improved after building a niche case study portfolio, not immediately upon declaring a niche: reddit.com/r/agency/comments/1cwt5xg/those_who_have_niched_down_their_agency_was_it/',
     'Hinge Marketing: specialization reduces perceived risk for the buyer; reduced perceived risk is the structural mechanism through which premium pricing becomes available -- the premium is not claimed, it is earned through proof: hingemarketing.com/blog/story/to-specialize-or-not-to-specialize | Agency Analytics: niche agency guide documents case studies and vertical-specific social proof as prerequisites for premium positioning, not outputs of it: agencyanalytics.com/blog/niche-agency'),
]

pp_col_widths = [Inches(0.28), Inches(1.10), Inches(1.60), Inches(1.90), Inches(1.92)]
make_table(pp_headers, pp_rows, pp_col_widths, header_size=9, row_size=8.5)
body_para(space_before=3, space_after=3)

# --- KEY STATS FOR WRITER ---
section_header('KEY STATS FOR WRITER')
p = body_para(space_before=4, space_after=2)
add_run(p, 'IMPORTANT: Deep adversarial research found that every specific quantified claim about niche agency financial advantages circulating in agency content failed verification. None have traceable primary sources. The stats section below reflects this honestly. Do not add statistics not listed here without editorial fact-check.', size=9.5, italic=True, color=DARK)

sub_header('Use with confidence:')
verified_stats = [
    'Hinge Marketing High Growth Study: "Year after year, we see data correlating specialization with greater growth and profitability." Drawn from 50,000+ professional services firms over two decades. Cite verbatim. Acknowledge this applies to professional services broadly, not social media agencies specifically. Source: hingemarketing.com/blog/story/to-specialize-or-not-to-specialize',
    'Agency Analytics documents five niching dimensions used by social media agencies: by industry vertical, by platform specialty, by content format, by business stage, and by geographic market. Cite as industry documentation, not a statistic. Source: agencyanalytics.com/blog/niche-agency',
    'The gradual transition model (6-18 months, phased wind-down of generalist clients) is consistently confirmed by practitioners in r/agency and documented by Agency Analytics and Doubleyourfreelancing as the standard approach. Cite as practitioner consensus.',
    'r/agency practitioner accounts confirm a 3-6 month revenue dip during the transition period before niche pipeline recovers. This is first-person practitioner testimony, not a study. Cite as such: "practitioners on r/agency consistently report..." Source: reddit.com/r/agency/comments/1cwt5xg/those_who_have_niched_down_their_agency_was_it/',
    'Hinge Marketing documents "going too narrow can constrain growth" as a real specialization risk. Use to balance the article. Same source as above.',
    'SocialPilot product context (verify with product team): SocialPilot serves X agencies managing Y+ client accounts. Use a real figure from the product team if available -- do not estimate.',
]
for s in verified_stats: bullet(s)

sub_header('DO NOT USE -- Failed Adversarial Verification (0-3 votes, no traceable primary source):')
bad_stats = [
    '"Niche agencies achieve 23% EBITDA margins vs 14% for generalists" -- source does not contain this data. Do not use.',
    '"Specialised agencies command 27-60% pricing premiums" -- no primary study found. Do not use.',
    '"Niche agencies reduce CAC from $11,800 to $4,200" -- no primary study found. Do not use.',
    '"Niche agencies improve win rates from 20% to 35%" -- no primary study found. Do not use.',
    '"72% of high-growth agencies are niche-focused" -- Agency Analytics article does not contain this figure. Do not use.',
    '"Specialised agencies achieve 60% higher average project value" -- no primary study found. Do not use.',
    'Any other specific percentage or dollar figure comparing niche vs generalist agency performance that does not appear in the Key Stats section above -- treat all such figures as unverified until a primary source is confirmed at time of writing.',
]
for s in bad_stats: bullet(s)
body_para(space_before=3, space_after=3)

# --- WRITER CHECKLIST ---
section_header('WRITER CHECKLIST')
p = body_para(space_before=4, space_after=4)
add_run(p, 'Before submitting your draft, confirm every item below is done.', size=10, italic=True)

checklist = [
    ('DATA INTEGRITY (CRITICAL)', [
        '[ ] No unverified financial figure appears anywhere in the article -- not even as "some sources claim X%" (see DO NOT USE list above)',
        '[ ] The Hinge Marketing correlation is the only cited primary research on financial benefits; it is cited verbatim and its scope (professional services broadly) is disclosed',
        '[ ] The three financial mechanisms (expertise signal, referral density, scope creep reduction) are each explained with a concrete example',
        '[ ] The article acknowledges directly that specific quantified data on niche agency financial performance does not currently exist as verified primary research',
    ]),
    ('NICHING FRAMEWORK', [
        '[ ] All five niching dimensions are covered: vertical, platform, content format, business stage, geography -- with social-media-specific examples for each',
        '[ ] The existing client audit framework includes a concrete scoring approach: retention rate, referral volume, delivery quality, client satisfaction -- not vague advice to "look at your best clients"',
        '[ ] A market size validation check is included: how to confirm a niche is large enough to sustain the agency\'s target roster',
        '[ ] The article distinguishes between genuinely too-narrow niches (single city, one obsoleting platform) and the unfounded niche-too-small anxiety most agency owners experience',
    ]),
    ('TRANSITION GUIDANCE', [
        '[ ] The 3-6 month revenue dip is addressed honestly -- with language that prepares the reader rather than minimises the risk',
        '[ ] A concrete transition timeline is provided: what to do in months 1-3, 3-6, and 6-18',
        '[ ] The decision trigger for when to start declining non-niche clients is defined (e.g., tied to pipeline coverage or financial runway)',
        '[ ] Imposter syndrome is addressed explicitly and the reframe is included: niche expertise is built through focused delivery, not claimed before it exists',
    ]),
    ('STRUCTURE AND READABILITY', [
        '[ ] Each H2 opens with a 2-3 sentence standalone answer',
        '[ ] Paragraphs are 3-4 lines maximum throughout',
        '[ ] The client audit section uses a table or structured scoring format, not prose',
        '[ ] The niching dimensions section uses a table mapping niche type to social-media-specific examples',
        '[ ] No section reads like a list of advice that could appear in any of the competing articles above',
    ]),
    ('FAQ SECTION', [
        '[ ] Each answer is 40-60 words',
        '[ ] Each answer starts directly with the answer, no preamble',
        '[ ] "What niches are most profitable?" is answered with a framework for evaluating profitability, not a list of hot niches',
        '[ ] "Can a generalist compete with niche agencies?" is answered honestly: yes, on price; no, on expertise positioning and referral density',
        '[ ] "How long does it take?" is answered with the honest 6-18 month practitioner-confirmed timeline',
    ]),
]
for section_title, items in checklist:
    p = body_para(space_before=7, space_after=2)
    add_run(p, section_title, bold=True, size=10, color=BLUE)
    for item in items: bullet(item)
body_para(space_before=3, space_after=3)

# --- INTERNAL LINKS ---
section_header('INTERNAL LINKS')
il_rows = [
    ('1', 'socialpilot.co/blog/social-media-for-agencies',
     'manage social media across multiple client accounts',
     'Introduction -- establishing the agency context before introducing niching'),
    ('2', 'socialpilot.co/blog/ideal-customer-profile',
     'define your ideal customer profile',
     'Niching strategy section -- when introducing the ICP-to-niche connection; first internal link after first H2'),
    ('3', 'socialpilot.co/blog/social-media-marketing-strategy',
     'build a social media strategy for your niche vertical',
     'Niche positioning section -- connecting niche choice to content strategy execution'),
    ('4', 'socialpilot.co/blog/social-media-proposal',
     'write a social media proposal for a niche client',
     'Sales and pitching section -- connecting niche positioning to the new-business process'),
    ('5', 'socialpilot.co/blog/social-media-agency-pricing',
     'set pricing for your social media agency',
     'Premium pricing section -- directly relevant to the pricing mechanism argument'),
    ('6', 'socialpilot.co/blog/client-onboarding',
     'onboard new clients into your agency workflow',
     'Transition section -- connecting the niche decision to the operational changes it requires'),
    ('7', 'socialpilot.co/blog/agency-reporting',
     'report social media results to niche clients',
     'Conclusion -- connecting specialisation to the reporting advantage niche agencies have'),
]
make_table(
    ['#', 'Post URL', 'Anchor Text', 'Place In Article'],
    il_rows,
    [Inches(0.28), Inches(2.1), Inches(2.1), Inches(2.32)]
)

p = body_para(space_before=8, space_after=2)
add_run(p, 'Internal Link Rules for the Writer:', bold=True, size=10)
il_rules = [
    'Do not use the same anchor text twice in one post',
    'Do not link to a competitor from within the post body (Agency Analytics, Databox, Hootsuite, Sprout Social, etc.)',
    'First internal link appears after the first H2, not in the introduction',
    'Every internal link must be contextually relevant, never forced',
    'Anchor text must be 3-5 words that are contextually meaningful',
    'Verify all internal link URLs are live before submitting the draft -- URL slugs may have changed',
]
for r in il_rules: bullet(r)
body_para(space_before=3, space_after=3)

# --- SAVE ---
out = '/home/user/new/CONTENT BRIEF - SocialPilot - Agency Niching v1.docx'
doc.save(out)
print(f'SAVED: {out}')
