"""Generate the daily trading-intelligence .docx brief."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-07-03"
BIAS = "SELECTIVE-BULL"
ENTRY_PERMISSION = "YELLOW"

OUT_PATH = f"trading-briefs/2026/07-July/Trading_Brief_{DATE_ISO}_{BIAS}.docx"


def shade_cell(cell, hex_fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tc_pr.append(shd)


def add_heading(doc, text, level=1, color=None):
    h = doc.add_heading(text, level=level)
    if color:
        for run in h.runs:
            run.font.color.rgb = color
    return h


def add_para(doc, text, bold=False, italic=False, size=11, color=None, align=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    return p


def add_bullets(doc, items):
    for item in items:
        doc.add_paragraph(item, style="List Bullet")


def add_kv_table(doc, rows, header=None, shade_header=True):
    ncols = len(rows[0])
    table = doc.add_table(rows=(1 if header else 0) + len(rows), cols=ncols)
    table.style = "Light Grid Accent 1"
    r_idx = 0
    if header:
        hdr = table.rows[0]
        for i, txt in enumerate(header):
            cell = hdr.cells[i]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(txt)
            run.bold = True
            run.font.size = Pt(11)
            if shade_header:
                shade_cell(cell, "1F3864")
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        r_idx = 1
    for i, row in enumerate(rows):
        trow = table.rows[r_idx + i]
        for j, val in enumerate(row):
            cell = trow.cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            run.font.size = Pt(11)
            if j == 0:
                run.bold = True
    return table


def badge_line(doc, label, value, fill_hex, text_color=RGBColor(0xFF, 0xFF, 0xFF)):
    table = doc.add_table(rows=1, cols=1)
    table.autofit = True
    cell = table.rows[0].cells[0]
    cell.text = ""
    shade_cell(cell, fill_hex)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{label}: {value}")
    run.bold = True
    run.font.size = Pt(13)
    run.font.color.rgb = text_color


# ============== BUILD DOCUMENT ==============
doc = Document()

# Tighten page margins
for section in doc.sections:
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

# ---------- HEADER ----------
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run("ELITE DAILY TRADING INTELLIGENCE BRIEF")
run.bold = True
run.font.size = Pt(20)
run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub_run = sub.add_run("Indian F&O Day Trader — Nifty + Stocks")
sub_run.italic = True
sub_run.font.size = Pt(12)

date_p = doc.add_paragraph()
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
dr = date_p.add_run("Date: Friday, 3 July 2026  |  Generated ~08:15 IST")
dr.bold = True
dr.font.size = Pt(11)

# Bias + permission badges
badge_line(doc, "TODAY'S BIAS", BIAS, "6AA84F")               # green-ish for cautious bull
badge_line(doc, "ENTRY PERMISSION", ENTRY_PERMISSION, "F1C232")  # amber for YELLOW
badge_line(doc, "CRUDE RULE MODE", "AGGRESSIVE BULL (₹6,505 MCX)", "38761D")

doc.add_paragraph()

# ---------- SECTION 0 — POSTMORTEM / COLD RESTART ----------
add_heading(doc, "Section 0 — Postmortem: Cold Restart", level=1, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "Brief accuracy score: N/A. The last brief in this repo is dated 23 April 2026 (71 days / ~48 sessions ago), "
    "and trade_log.txt has zero recorded entries in that entire window. This is not a continuation — it is a "
    "cold restart of the routine.", bold=True, size=11)
add_bullets(doc, [
    "The routine lapsed for 71 days with no brief and no logged signals — that gap is the biggest failure to flag, bigger than any single bad call would have been.",
    "Adjustment applied today: stock count trimmed to 5 (not 7), every grade held one notch below what catalysts alone would justify, and a 60-minute (not 30-minute) confirmation wait before any new position.",
    "Normal grading resumes only after a few honestly-logged sessions rebuild an actual track record.",
])

# ---------- SECTION 1 — MACRO SNAPSHOT ----------
doc.add_paragraph()
add_heading(doc, "Section 1 — Macro Snapshot", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Metric", "Reading", "Interpretation"],
    rows=[
        ["MCX Crude (₹/bbl)", "~6,505", "AGGRESSIVE BULL (<7,500); no boundary risk"],
        ["WTI / Brent", "$68.4 / $70.6-70.7", "Normal ~$2.2 spread — crisis premium gone (was $9 in April)"],
        ["USD/INR", "~95.08 (range 94.90-95.33 across sources)", "Recheck live; input to MCX proxy"],
        ["India VIX", "12.28 (-7.25%)", "Lowest reading in this repo's history — full-size band, but flagged as complacent"],
        ["Nifty 50 (2 Jul close)", "24,175.70 (+0.71%); H 24,194.55 / L 24,058.80", "Near recent highs"],
        ["Sensex (2 Jul close)", "77,502 (+0.75%)", "Fresh highs"],
        ["Gift Nifty", "~24,355 vs futures 24,270.1", "Implied gap-up ~+179 pts — sources conflict on timestamp, recheck at 9:00"],
        ["Dow / S&P 500 / Nasdaq-100 (2 Jul)", "52,900.07 record / 7,483.24 flat / -1.61%", "Mixed: record Dow, chip-led Nasdaq selloff"],
        ["US markets today", "CLOSED", "4 Jul holiday observed (4 Jul falls on Saturday) — no fresh cross-check session"],
        ["Asian markets (stale: 30 Jun-1 Jul)", "Nikkei +0.59%, Hang Seng ~22.9k, Kospi +1.0%, Shanghai +0.44%", "Broadly firm, 2-3 sessions old"],
        ["FII / DII (2 Jul, provisional)", "-₹311.8 cr / +₹1,784.4 cr", "DII absorption > FII selling — floor signal"],
        ["Bank Nifty (2 Jul close)", "~57,542.90 (range 57,456.65-58,011.95)", "Near highs, lot size 30"],
        ["Nifty F&O ban list", "NIL as of 30 Jun (last confirmed)", "Live ban file returned 403 on fetch — recheck at open"],
        ["Nifty weekly expiry", "Shifted to Tuesday since 1 Sep 2025; next 7 Jul", "Today is NOT an expiry day"],
    ])

# ---------- SECTION 2 — CRUDE RULE APPLIED ----------
doc.add_paragraph()
add_heading(doc, "Section 2 — Crude Rule Applied", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "MODE: AGGRESSIVE BULL — full regime flip vs the last logged brief.", bold=True, size=12)
add_bullets(doc, [
    "MCX crude sits at ~₹6,505 — well inside the <₹7,500 AGGRESSIVE BULL band, roughly ₹1,000 clear of the nearest boundary. No flip risk today.",
    "This is a collapse from the ₹8,390 reading in the 23 April brief (WTI was ~$89 then vs $68.4 now) following a ceasefire MoU signed 17 June, which has now held through more than the routine's own 3-5 session validation bar — on the price action, if not on the underlying politics.",
    "Caveat: search sources reference unverified reports of an Iran leadership crisis and funeral processions (7-9 July) — flagged as a live tail-risk headline, not a confirmed fact. If real, it could reverse this regime intraday.",
    "US markets are closed today for the observed 4 July holiday, removing a session's worth of cross-check on Thursday's Nasdaq-100 semiconductor rout.",
])

add_para(doc, "Crude-aligned sector bias — REVERSED vs the April brief:", bold=True, size=11)
add_bullets(doc, [
    "Bullish now: BPCL, HPCL, IOC (OMC marketing-margin tailwind), IndiGo & aviation (fuel-cost relief) — the opposite call from April.",
    "Bearish/pressured now: ONGC, Oil India (lower per-barrel realizations as crude falls).",
])

# ---------- SECTION 3 — INDIAN MARKET INTERNALS ----------
doc.add_paragraph()
add_heading(doc, "Section 3 — Indian Market Internals", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "Nifty Technical Levels", bold=True, size=12)
add_kv_table(doc,
    header=["Level", "Value", "Meaning"],
    rows=[
        ["Previous Close", "24,175.70", "+0.71% on 2 Jul"],
        ["Expected Open", "~24,270-24,355", "Gap-up per Gift Nifty (data-quality flagged, recheck live)"],
        ["Support S1", "24,058.80", "2 Jul intraday low"],
        ["Support S2", "24,000", "Psychological + max-pain zone + heaviest put OI"],
        ["Resistance R1", "24,194.55", "2 Jul intraday high"],
        ["Resistance R2", "24,500", "Heaviest call OI — 'call wall'"],
    ])

doc.add_paragraph()
add_para(doc, "Volatility & Flows", bold=True, size=12)
add_bullets(doc, [
    "India VIX: 12.28 (-7.25%) — lowest reading in this repo's history. Nominally full-size, but treated as HALF-SIZE today given the cold-restart and unverified-headline caveats.",
    "FII (2 Jul): Net SELL ₹311.8 cr — mild.",
    "DII (2 Jul): Net BUY ₹1,784.4 cr — absorbing FII selling comfortably; floor signal.",
])

doc.add_paragraph()
add_para(doc, "Options Positioning (7 Jul weekly expiry — Tuesday, not today)", bold=True, size=12)
add_bullets(doc, [
    "Max Pain (7 Jul): ~23,987.5 — gravitational pull level, close to psychological 24,000.",
    "Highest Put OI: concentrated near 24,000 — support/floor zone.",
    "Highest Call OI: concentrated near 24,500 — ceiling/resistance ('call wall').",
    "Interpretation: range-bound drift 24,000-24,500 into next Tuesday's expiry is the base case absent a fresh headline.",
])

# ---------- SECTION 4 — STOCK OPPORTUNITIES ----------
doc.add_paragraph()
add_heading(doc, "Section 4 — Stock Opportunities", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc,
    "Five setups today (trimmed from the usual 5-7 under the cold-restart rule). All CMPs are the latest confirmed "
    "web-sourced prints (dated per stock), NOT a live broker feed — reconfirm at market open before sizing.",
    italic=True, size=10)


def stock_card(doc, rows):
    t = doc.add_table(rows=len(rows), cols=2)
    t.style = "Light Grid Accent 1"
    for i, (k, v) in enumerate(rows):
        lc = t.rows[i].cells[0]
        rc = t.rows[i].cells[1]
        lc.text = ""
        rc.text = ""
        p1 = lc.paragraphs[0]
        r1 = p1.add_run(k)
        r1.bold = True
        r1.font.size = Pt(11)
        shade_cell(lc, "D9E1F2")
        p2 = rc.paragraphs[0]
        r2 = p2.add_run(v)
        r2.font.size = Pt(11)


add_para(doc, "STOCK 1 — MARUTI SUZUKI CALL — Grade B+ (downgraded from A)", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot / Sector", "₹14,359 (2 Jul) / 50 / Auto"),
    ("Catalyst", "June sales +19% YoY (200,390 units, released 1 Jul); cheap crude is a secondary demand tailwind"),
    ("Option Setup", "ATM CE, July monthly (avoid 7 Jul weekly — too close, low liquidity for a fresh position)"),
    ("Entry Trigger", "15-min close above 2 Jul high AND Nifty holding above 24,150 AND VIX under 14 AND 60-min post-open wait"),
    ("Targets / SL", "T1 +3%, T2 +5-6% / SL: 15-min close below 2 Jul low"),
    ("R:R / Time Stop", ">=2:1 / exit if untriggered by 11:00 AM"),
])

doc.add_paragraph()
add_para(doc, "STOCK 2 — TATA MOTORS (TMPV) CALL — Grade B+ (downgraded from A)", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot / Sector", "₹346 (2 Jul, passenger-vehicle entity post Oct-25 demerger) / 800 / Auto"),
    ("Catalyst", "PV volumes +69% YoY, EVs +183%, Harrier EV demand ~2x capacity"),
    ("Option Setup", "ATM CE, July monthly"),
    ("Entry Trigger", "15-min close above 2 Jul high AND Maruti setup also confirming AND 60-min wait"),
    ("Targets / SL", "T1 +4%, T2 +7% / SL: 15-min close below 2 Jul low"),
    ("R:R / Time Stop", ">=2:1 / 11:00 AM"),
])

doc.add_paragraph()
add_para(doc, "STOCK 3 — BPCL CALL — Grade B+ (crude-regime-flip trade)", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot / Sector", "₹310.4 (2 Jul) / 1,975 / Energy (OMC)"),
    ("Catalyst", "Direct beneficiary of the crude collapse — marketing-margin tailwind that did NOT exist in the April brief"),
    ("Option Setup", "ATM CE, July monthly"),
    ("Entry Trigger", "MCX crude confirmed under ₹7,000 at open AND 15-min close above 2 Jul high AND 60-min wait"),
    ("Targets / SL", "T1 +3%, T2 +5% / SL: 15-min close below 2 Jul low, or crude reclaims ₹7,000"),
    ("R:R", ">=2:1"),
])

doc.add_paragraph()
add_para(doc, "STOCK 4 — ONGC PUT — Grade B (counter-directional crude-headwind trade)", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot / Sector", "₹233.98-235.96 (2 Jul, sources not fully reconciled) / 2,250 / Energy (Upstream E&P)"),
    ("Catalyst", "Mirror image of BPCL — the same crude collapse pressures per-barrel realizations"),
    ("Option Setup", "ATM PE, July monthly"),
    ("Entry Trigger", "MCX crude stays under ₹7,000 AND 15-min close below 2 Jul low AND 60-min wait; do not pair naively with Stock 3"),
    ("Targets / SL", "T1 -3%, T2 -5% / SL: 15-min close above 2 Jul high, or crude spikes above ₹7,500"),
    ("R:R", ">=2:1"),
])

doc.add_paragraph()
add_para(doc, "STOCK 5 — BANK NIFTY CALL — Grade B (conditional, macro-alignment diversifier)", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot / Sector", "~57,542.90 (2 Jul, range 57,456.65-58,011.95) / 30 / Banking/Index"),
    ("Catalyst", "DII absorption (+₹1,784 cr) and index proximity to highs — passive macro alignment, not catalyst-driven"),
    ("Option Setup", "ATM CE, July monthly (skip 7 Jul weekly)"),
    ("Entry Trigger", "Nifty confirms gap-up hold above 24,150 for 60 minutes AND VIX stays under 14"),
    ("Targets / SL", "T1 +2%, T2 +3.5% / SL: 15-min close below 57,456"),
    ("R:R", ">=2:1"),
])

doc.add_paragraph()
add_para(doc, "Macro hedge (flagged, not a full trade card)", bold=True, size=11)
add_bullets(doc, [
    "Given the unverified Iran-leadership headline risk, a defense hedge (HAL/BEL) or gold exposure belongs here.",
    "No verified live CMP/lot data was obtained for HAL/BEL this session — flagged as a watch item, not invented. Check live quotes before adding a small hedge leg.",
])

doc.add_paragraph()
add_para(doc, "Stocks to actively avoid today", bold=True, size=11)
add_bullets(doc, [
    "TCS / broad IT — CMP unreconciled across sources (₹2,068 vs ₹2,045.50); no catalyst until 9 Jul results; US markets closed today means no confirmation on Thursday's Nasdaq-100 semiconductor rout (-1.61%). Wait for Monday's US session.",
    "Hyundai Motor India (HYUNDAI) — newly F&O-listed (Apr 2026, thin derivatives history), hit by a supplier fire this week; too illiquid/new a contract to trade with confidence today.",
    "Any 7 Jul weekly options for a fresh position — 2 sessions from expiry; monthly (July) contracts only.",
])

# ---------- SECTION 5 — DOMESTIC MACRO ----------
doc.add_paragraph()
add_heading(doc, "Section 5 — Domestic Macro & India-Specific Triggers", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Driver", "Status"],
    rows=[
        ["RBI Repo Rate", "5.25%, neutral stance; next MPC 3-5 August 2026"],
        ["India-US Trade Talks", "Unresolved as of late June; hard tariff-snapback deadline 24 July 2026"],
        ["Monsoon (IMD)", "Below-normal July flagged after a very weak June (~40% deficit); recovering since 29 Jun"],
        ["SEBI", "Proposed MTF (margin trading) tightening — relevant for broking/NBFC names"],
        ["Earnings Today", "None major scheduled for 3 July"],
        ["Earnings This Month", "TCS board 9 Jul (Q1 FY27 first major IT print); HCL Tech 13 Jul; Infosys 22-23 Jul"],
        ["June Auto Sales (released)", "Maruti 200,390 units (+19%); Tata Motors PV +69%; M&M highest-ever domestic PV (+28%); Hyundai hit by supplier fire"],
        ["US Fed", "Held rates 3.50-3.75% on 17 Jun; next FOMC 28-29 Jul (no fresh dot plot that meeting)"],
        ["China PMI (June)", "50.3, beat expectations, tech/AI-export led"],
        ["Domestic Macro Verdict", "Supportive on rates + DII flow; headwinds from tariff deadline + weak monsoon, both largely unpriced"],
    ])

# ---------- SECTION 6 — FINAL VERDICT ----------
doc.add_paragraph()
add_heading(doc, "Section 6 — Final Verdict", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_kv_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["Crude Rule Mode", "AGGRESSIVE BULL (₹6,505 MCX) — no boundary risk"],
        ["Market Bias", "Cautiously bullish, selective — not a blanket buy-the-gap day"],
        ["VIX Sizing Rule", "HALF-SIZE applied (VIX 12.28 nominally full-size, but overridden by entry-permission caveats)"],
        ["Key Support", "24,058.80 -> 24,000"],
        ["Key Resistance", "24,194.55 -> 24,500"],
        ["Entry Permission", "YELLOW — selective, 60-min confirmation, half-size"],
        ["Top Risk 1", "Unverified Iran leadership/succession claim; funeral-period headlines 7-9 July"],
        ["Top Risk 2", "US markets closed today — no cross-check on Thursday's Nasdaq semiconductor rout"],
        ["Top Risk 3", "No validated track record after a 71-day routine gap"],
    ])

doc.add_paragraph()
add_para(doc, "Reason for YELLOW (not GREEN):", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "Every headline metric today looks clean — low VIX, DII buying every dip, Nifty/Sensex at highs, crude collapsed "
    "into an aggressive-bull band, Gift Nifty pointing to a gap-up. But the routine has no track record after a "
    "71-day gap, a genuinely enormous and unconfirmed geopolitical claim is circulating, and US markets are shut "
    "today removing a full session of cross-check on a real Thursday-night tech selloff. Clean-looking days built "
    "on unverified inputs are exactly when this routine's Layer 2/3 checks are supposed to slow things down.",
    size=11)

doc.add_paragraph()
add_para(doc, "ONE-LINE SUMMARY", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "\"Cold restart after a 71-day brief gap: crude has crashed to Rs6,505 MCX (AGGRESSIVE BULL, reversing April's "
    "upstream-bullish call to OMC/auto-bullish, upstream-bearish), Nifty/Sensex sit at fresh highs on DII support, "
    "but an unverified Iran-leadership headline and a US market holiday leave real risk unconfirmed -- trade "
    "selectively, half-size, 60-min confirmation, watch Rs7,000 MCX as the invalidation line.\"",
    italic=True, size=12)

# ---------- FOOTER ----------
doc.add_paragraph()
doc.add_paragraph()
foot = doc.add_paragraph()
foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = foot.add_run("Generated by Daily Trading Routine — for personal educational use only. Not financial advice.")
fr.italic = True
fr.font.size = Pt(9)
fr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

doc.save(OUT_PATH)
print(f"Wrote {OUT_PATH}")
