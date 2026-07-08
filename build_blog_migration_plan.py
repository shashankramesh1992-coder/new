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
RED     = RGBColor(0xB4, 0x23, 0x18)
HEAD_BG = "1A56DB"
HEAD_FG = RGBColor(0xFF, 0xFF, 0xFF)
ALT_BG  = "FAFAFA"
FLAG_BG = "FDECEC"

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

def body_para(text='', bold=False, size=10, color=BLACK, space_before=4, space_after=4, indent_pt=0, italic=False):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(space_before); p.paragraph_format.space_after = Pt(space_after)
    if indent_pt: p.paragraph_format.left_indent = Pt(indent_pt)
    if text:
        run = p.add_run(text); run.bold = bold; run.italic = italic; run.font.size = Pt(size); run.font.color.rgb = color
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

def bullet(text, size=9.5, bold_lead=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(1); p.paragraph_format.space_after = Pt(1)
    if bold_lead:
        add_run(p, bold_lead, bold=True, size=size)
        add_run(p, text, size=size)
    else:
        run = p.add_run(text); run.font.size = Pt(size); run.font.color.rgb = BLACK
    return p

def numbered(text, size=9.5):
    p = doc.add_paragraph(style='List Number')
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text); run.font.size = Pt(size); run.font.color.rgb = BLACK
    return p

def make_table(headers, rows, col_widths, header_size=9, row_size=8.5, bold_last_row=False):
    tbl = doc.add_table(rows=1+len(rows), cols=len(headers)); tbl.style = 'Table Grid'
    hrow = tbl.rows[0]
    for i, (h, w) in enumerate(zip(headers, col_widths)):
        c = hrow.cells[i]; c.width = w
        set_cell_bg(c, HEAD_BG); cell_para(c, h, bold=True, size=header_size, color=HEAD_FG)
    for ri, rd in enumerate(rows):
        row = tbl.rows[ri+1]
        is_last = bold_last_row and ri == len(rows) - 1
        bg = 'FFFFFF' if ri % 2 == 0 else ALT_BG
        if is_last: bg = 'E5E9F5'
        for ci, (val, w) in enumerate(zip(rd, col_widths)):
            c = row.cells[ci]; c.width = w
            set_cell_bg(c, bg); cell_para(c, val, size=row_size, bold=is_last)
    return tbl

def flag_box(title, lines):
    tbl = doc.add_table(rows=1, cols=1); tbl.style = 'Table Grid'
    cell = tbl.rows[0].cells[0]; set_cell_bg(cell, FLAG_BG)
    p = cell.paragraphs[0]; p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(2)
    run = p.add_run(title); run.bold = True; run.font.size = Pt(10); run.font.color.rgb = RED
    for line in lines:
        p2 = cell.add_paragraph(); p2.paragraph_format.space_before = Pt(1); p2.paragraph_format.space_after = Pt(3)
        r2 = p2.add_run('•  ' + line); r2.font.size = Pt(9.5); r2.font.color.rgb = BLACK
    body_para(space_before=2, space_after=2)

# ============================================================
# TITLE
# ============================================================
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(3)
r = p.add_run('BLOG MIGRATION PLAN')
r.bold = True; r.font.size = Pt(16); r.font.color.rgb = BLUE

p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_before = Pt(0); p2.paragraph_format.space_after = Pt(10)
r2 = p2.add_run('SocialPilot Blog Restructure — /blog → /insights, /strategy, /compare')
r2.bold = True; r2.font.size = Pt(11.5); r2.font.color.rgb = BLACK

p3 = doc.add_paragraph(); p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
p3.paragraph_format.space_before = Pt(0); p3.paragraph_format.space_after = Pt(12)
r3 = p3.add_run('Based on: Blog URL Mapping (Final) — Complete Blogs List + URL Mapping + Redirect Plan tabs, cross-checked')
r3.italic = True; r3.font.size = Pt(9); r3.font.color.rgb = DARK

# ============================================================
# TL;DR
# ============================================================
section_header('TL;DR')
bullet('529 legacy /blog URLs were audited. 283 articles are being retained (including all /blog/tag/* pages and 30 articles reclassified out of removal per the new /insights-vs-/strategy split below), 246 are being removed / pruned.', bold_lead='Full audit: ')
bullet('Retained content moves into three new top-level sections — /insights, /strategy, /compare — plus a smaller /blog for the handful of posts (and all tag pages) that don’t fit those buckets.', bold_lead='New IA: ')
bullet('240 articles have a clean, verified, 1-to-1 old → new URL redirect ready to hand to dev (223 from the original mapping + 17 newly reclassified). 43 more (blog homepage, all 28 /blog/tag/* pages, 1 glossary post, + 13 reclassified articles that didn’t clearly fit /insights or /strategy) keep their current URL — no redirect needed.', bold_lead='Redirects: ')
bullet('The 246 removed URLs carry only ~7.5K sessions/12mo combined, vs. 278.9K for retained content — a ~2.7% share. Most are either already-dead legacy aliases or thin, low-traffic posts.', bold_lead='Risk: ')
bullet('All 170 removed URLs that lacked a defined disposition now have one: 113 get a redirect to a genuinely relevant, verified-live page elsewhere on socialpilot.co; 57 are confirmed 410s with no suitable equivalent. See Section 4 and the appendix workbook.', bold_lead='Redirect research done: ')
bullet('Effective today, 30 URLs previously slated for a 410 are being reclassified as retained instead: content that\'s observational/analytical about the market or platform ("what\'s happening") goes to /insights, content that\'s actionable how-to/tactics/planning ("what we should do about it") goes to /strategy — and anything that doesn\'t clearly fit either stays in /blog rather than being forced into a bucket, capped at 13 holdbacks. 17 of the 30 clearly (or plausibly) fit; the other 13 — the lowest-traffic, most purely agency-internal-process articles — hold back in /blog. See Section 1a.', bold_lead='Policy update: ')
bullet('1,727 internal links currently point at pages being removed and need repointing regardless of the redirect/410 call — see Section 6.', bold_lead='Open item: ')

# ============================================================
# 1. THE NEW STRUCTURE
# ============================================================
section_header('1. THE NEW STRUCTURE')
body_para('Everything being retained lands in one of four buckets:', size=9.5, space_after=6)
make_table(
    ['New Section', 'Path', 'Articles', 'Sessions (12mo)'],
    [
        ('Insights', '/insights/*', '66', '125,025'),
        ('Strategy', '/strategy/*', '133', '131,422'),
        ('Compare', '/compare/*', '5', '1,401'),
        ('Stays in /blog (incl. all 28 tag pages + 13 reclassified holdbacks)', '/blog/*', '79', '21,093'),
        ('Total retained', '—', '283', '278,941'),
    ],
    [Inches(1.7), Inches(1.4), Inches(1.1), Inches(1.72)],
    bold_last_row=True,
)
body_para('Insights = trend/news commentary ("what’s happening"). Strategy = how-to / execution content (including a 13-post Brand & Audience Strategy sub-cluster). Compare = the 5 competitor-pricing pages. /blog retains generic commentary that doesn’t fit the other three, plus every /blog/tag/* archive page — tag pages are being kept, not pruned.', size=8.5, space_before=6)
body_para('Note: these are the only 4 destinations in this migration — there is no new "/industry" section. Section 4 sends a small number of REMOVED articles (3 of 246) to /industry/* pages that already exist on the site today, unrelated to this restructure, because they happen to be the closest live equivalent for that specific vertical topic.', size=8.5, italic=True, space_before=2)

# ============================================================
# 1a. INSIGHTS VS STRATEGY RECLASSIFICATION
# ============================================================
section_header('1a. POLICY UPDATE — 30 ARTICLES RECLASSIFIED FROM REMOVAL')
body_para('Effective today: content goes to /insights if it\'s observational/analytical about the market or platform ("what\'s happening" — data, trends, algorithms, stats). Content goes to /strategy if it\'s actionable — how-to guides, tactics, planning ("what we should do about it"). Anything that doesn\'t clearly fall under one of those two holds back in /blog rather than being forced into a bucket — capped at 13 holdbacks, so borderline cases get the benefit of the doubt before landing in /blog. 30 articles previously marked for a 410 were re-reviewed against this rule and are all being retained. Full list with rationale for every one of the 30: BLOG_RECLASSIFIED_TO_INSIGHTS_STRATEGY__v3.xlsx.', size=9.5, space_after=6)
make_table(
    ['Destination', 'Count', 'Sessions (12mo)', 'Example'],
    [
        ('/insights', '5', '160', 'instagram-shadow-banned (a specific platform algorithm/moderation phenomenon)'),
        ('/strategy', '12', '535', 'podcast-promotion (an actionable how-to tactic)'),
        ('Held back in /blog (doesn\'t clearly fit either)', '13', '59', 'scale-social-media-approval-process (agency business operations, not social-platform/market data or a marketing execution tactic)'),
        ('Total re-reviewed', '30', '754', ''),
    ],
    [Inches(1.85), Inches(0.55), Inches(1.0), Inches(2.82)],
    row_size=8,
    bold_last_row=True,
)
body_para('The cutoff: articles with more than 8 sessions/12mo got the benefit of the doubt on a plausible secondary fit (e.g. social-media-manager-skills reads as a checklist for hiring/skill-building, so it moved to /strategy). The 13 held back are all <=8 sessions/12mo and almost entirely the "agency operations" cluster — client approval workflows, team structure, agency scaling pain points. That content is about running an agency\'s internal business, not the social media market/platform (/insights) or social-media-marketing execution tactics (/strategy), so it stays in /blog rather than being stretched to fit. The 17 that do move keep their existing slug — only the path prefix changes (e.g. /blog/instagram-shadow-banned → /insights/instagram-shadow-banned), the same pattern as every other redirect in this migration.', size=8.5, space_before=6, italic=True)

# ============================================================
# 2. WHAT'S MOVING
# ============================================================
section_header('2. WHAT’S MOVING (REDIRECTS)')
bullet('The URL Mapping tab has a full, verified old → new mapping for all 223 articles: no duplicate destinations, no duplicate sources, one hop each. Add the 17 reclassified articles (Section 1a) on top of this — same pattern, new destination — for 240 total.', bold_lead='Ready to implement: ')
bullet('43 retained URLs need no redirect — the /blog homepage, all 28 /blog/tag/* archive pages, 1 glossary post, and the 13 reclassified articles that didn’t clearly fit /insights or /strategy (Section 1a) all stay exactly where they are.', bold_lead='No change needed: ')
bullet('Hand the URL Mapping tab + the reclassified-articles workbook directly to dev/SEO as the combined 301 redirect map — Current URL → New URL, one row per rule.', bold_lead='Action: ')

# ============================================================
# 3. WHAT'S BEING REMOVED
# ============================================================
section_header('3. WHAT’S BEING REMOVED')
body_para('246 legacy URLs are being retired — this excludes all /blog/tag/* pages (kept as-is, Section 2) and the 30 articles reclassified out of removal (Section 1a: 17 into /insights or /strategy, 13 held back in /blog). Breaking the 246 down by what they actually are today:', size=9.5, space_after=6)
make_table(
    ['Group', 'Count', 'What it is', 'Action needed'],
    [
        ('Already-redirected aliases → a retained article', '76', 'Old URL already 301s into a page we’re keeping', 'Re-point redirect to that article’s new URL'),
        ('Already-redirected aliases → another removed page', '25', 'Old URL already 301s into a page also being cut', 'Retire together, no separate traffic impact'),
        ('Dead-end aliases (target not in this audit)', '17', 'Old redirect chain, unclear/off-site target', 'Spot-check, then 410'),
        ('Live, standalone thin/low-traffic pages', '128', 'Avg. 40 sessions/yr, max 163', 'Resolved — redirect or 410 assigned, see Section 4'),
    ],
    [Inches(2.0), Inches(0.6), Inches(2.1), Inches(2.22)],
    row_size=8,
)
body_para('Combined traffic exposure across all 246 removed URLs: 7,493 sessions/12mo, vs. 278,941 for retained content — a ~2.7% share. This is a low-risk cut.', size=9.5, space_before=8, bold=True)

# ============================================================
# 4. WHERE REMOVED ARTICLES GO (RESEARCH COMPLETE)
# ============================================================
section_header('4. WHERE DO REMOVED ARTICLES GO? (RESEARCH COMPLETE)')
body_para('The source sheet only specified a 410 for 6 URLs, leaving 170 removed URLs (the 128 live articles above, plus 42 legacy aliases whose old redirect target is also being cut) with no defined disposition — /blog/tag/* pages and the 30 reclassified articles (Section 1a) are excluded from this count since they are being retained, not removed. Each of the 170 was checked against the retained articles’ new URLs AND the live socialpilot.co domain — /features, /industry, /tools, /compare — for a genuinely relevant, real, verified-live destination before deciding.', size=9.5, space_after=6)
make_table(
    ['Outcome', 'Count', 'Sessions (12mo)', 'What it means'],
    [
        ('Redirect — high confidence', '55', '—', 'Clear topical match to a retained page (e.g. how-to-sell-on-pinterest → how-to-make-money-on-pinterest)'),
        ('Redirect — medium/low confidence', '58', '—', 'Reasonable but broader match (e.g. a niche how-to folding into a general topic hub)'),
        ('410 Gone', '57', '1,870', 'No equivalent content exists anywhere on the domain — safe to retire outright'),
        ('Total resolved', '170', '4,786 (redirects)', ''),
    ],
    [Inches(2.0), Inches(0.7), Inches(1.15), Inches(2.07)],
    row_size=8,
)
body_para('Full URL-by-URL detail — old URL, action, redirect target, confidence, and rationale — is in the companion workbook: BLOG_REDIRECT_PLAN__removed_urls__v3.xlsx (covers all 246 removed URLs, including the 76 already-redirected aliases above)', size=9.5, space_before=8, bold=True)
flag_box('Still needs sign-off before launch', [
    'The 58 "medium/low confidence" redirects are judgment calls (e.g. a dentists article folding into a general doctors/healthcare page) — SEO/content lead should spot-check these before they go live.',
    '8 of the 113 new redirects point to real, live pages found outside the original mapping tab: 5 go to the /compare or /insights section hub (still within this migration’s structure), and 3 go to pre-existing /industry/* landing pages (photography, NGO, restaurant marketing agency) that have nothing to do with this restructure — confirm all 8 are the intended landing spots, not just an acceptable fallback.',
    'The 30 articles re-reviewed in Section 1a involved judgment calls too — both which 17 fit /insights or /strategy (some only plausibly, not clearly), and which 13 didn\'t (e.g. "10 mistakes to avoid" reads as either analytical or actionable, which is why it held back) — worth a quick pass by whoever owns the /insights-vs-/strategy split to confirm before launch.',
    '1,727 internal links across the site currently point at pages being removed. These need to be found and repointed to the new redirect targets (or removed, for pages going to a 410) — a redirect alone does not fix a stale internal link. See Section 6.',
])

# ============================================================
# 5. DATA INTEGRITY VERIFICATION
# ============================================================
section_header('5. DATA INTEGRITY VERIFICATION')
body_para('Before finalizing, the three source tabs (Complete Blogs List, URL Mapping, Redirect Plan) were cross-checked programmatically against each other. All checks passed:', size=9.5, space_after=6)
bullet('Zero duplicate URLs across all three tabs.')
bullet('Every URL in URL Mapping (223 rows) is marked "To Be Retained" in Complete Blogs List — and vice versa, every "To Be Retained" URL not in URL Mapping is one of the 30 stay-as-is pages (Section 2).')
bullet('Every URL in Redirect Plan (then 276 rows) is marked "To be Removed" in Complete Blogs List, with zero overlap against the URL Mapping set.')
bullet('Every redirect target is a verified real URL — either the New URL for a retained article, or one of the 5 live domain pages found outside the mapping tab (Section 4). Zero blank targets, zero self-redirects, and zero redirect chains (no target that is itself another removed URL).')
bullet('All 76 "already-aliased" rows independently re-verified against Complete Blogs List’s redirect-chain data and URL Mapping’s new-URL column — exact match on all 76.')
bullet('223 (URL Mapping) + 276 (Redirect Plan) + 30 (stay-as-is) = 529 — reconciled exactly with the full blog audit at that point.')
bullet('Since then, the 30 re-reviewed articles (Section 1a) were pulled out of Redirect Plan and independently checked: all 30 confirmed present exactly once in Complete Blogs List, zero duplicates against each other or against the URL Mapping / Redirect Plan sets, and the 17 that get a new URL follow the same slug-preserving pattern as the rest of the migration. 240 (URL Mapping + 17 reclassified) + 246 (Redirect Plan) + 43 (stay-as-is, incl. the 13 held back in /blog) = 529 — still reconciles.')
flag_box('One source-data note (not a plan error)', [
    'Complete Blogs List’s Status column still shows the 22 corrected tag pages, and now also the 30 reclassified articles, as "To be Removed" — it was never updated after either correction. The Redirect Plan, the reclassified-articles workbook, and this document all correctly override it; anyone pulling counts directly from that Status column should be aware of both overrides.',
])

# ============================================================
# 6. EXECUTION CHECKLIST
# ============================================================
section_header('6. EXECUTION CHECKLIST')
numbered('SEO/content lead spot-checks the 58 medium/low-confidence redirects and the 30 reclassified /insights-vs-/strategy-vs-/blog calls (Sections 1a and 4).')
numbered('Implement all 223 verified 301 redirects from the URL Mapping tab, plus the 17 from the reclassified-articles workbook (240 total retained-article redirects).')
numbered('Implement the 189 redirects + 57 410s for removed URLs from the redirect-plan workbook — do NOT touch /blog/tag/* pages or the 13 reclassified articles held back in /blog, they are being kept as-is.')
numbered('Audit and repoint the 1,727 internal links currently pointing at removed URLs.')
numbered('Update the XML sitemap: add new /insights, /strategy, /compare URLs; remove pruned URLs (tag pages stay in the sitemap).')
numbered('Update main nav, footer, and related-posts modules to reflect the 3 new sections.')
numbered('Submit the updated sitemap in Google Search Console and spot-check a sample of redirects live.')
numbered('Monitor 404s, redirect chains, and organic traffic for all migrated/redirected URLs at 7 / 30 / 60 / 90 days post-launch.')

# ============================================================
# 7. SUGGESTED TIMELINE
# ============================================================
section_header('7. SUGGESTED TIMELINE')
make_table(
    ['Week', 'Milestone'],
    [
        ('Week 1', 'SEO sign-off on the removed-URL redirect map (Section 4); QA full redirect set'),
        ('Week 2', 'Implement redirects + 410s in staging; repoint internal links'),
        ('Week 3', 'Launch; submit sitemap to Google Search Console'),
        ('Weeks 4–12', 'Monitor rankings/traffic; fix any redirect chains or stray 404s'),
    ],
    [Inches(1.1), Inches(4.92)],
)

# ============================================================
# 7. SUCCESS METRICS
# ============================================================
section_header('8. SUCCESS METRICS')
bullet('No net loss in organic sessions to migrated content at the 60/90-day mark (baseline: 278,941 sessions/12mo across the 283 retained articles).')
bullet('Zero 404s reachable via internal links post-launch.')
bullet('All 240 retained-article redirects resolve in a single hop — no chains.')
bullet('Search Console shows the new /insights, /strategy, /compare URLs indexed within 30 days.')

# --- SAVE ---
out = '/home/user/new/BLOG_MIGRATION_PLAN__SocialPilot_v1.docx'
doc.save(out)
print(f'SAVED: {out}')
