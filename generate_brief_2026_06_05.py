"""Generate the daily trading-intelligence .docx brief — 5 June 2026."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-06-05"
BIAS = "BEAR-CAUTION"
ENTRY_PERMISSION = "YELLOW"

OUT_PATH = f"trading-briefs/2026/06-June/Trading_Brief_{DATE_ISO}_{BIAS}.docx"


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


# ============== BUILD DOCUMENT ==============
doc = Document()

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
dr = date_p.add_run("Date: Friday, 5 June 2026  |  Generated 08:45 IST  |  Capital: ₹2,00,000")
dr.bold = True
dr.font.size = Pt(11)

# Bias badges — RED for BEAR, YELLOW for CAUTION entry
badge_line(doc, "TODAY'S BIAS", "BEAR-CAUTION", "CC0000")
badge_line(doc, "ENTRY PERMISSION", "YELLOW — NO TRADES BEFORE 10:30 AM (RBI MPC at 10:00 AM)", "E69138")
badge_line(doc, "CRUDE RULE MODE", "BEAR (MCX ₹9,234) — CALLS BLOCKED except ONGC / OIL India / MRPL", "CC0000")
badge_line(doc, "VIX SIZING RULE", "HALF SIZE (India VIX 16.28 — in 15–20 band)", "6FA8DC")

doc.add_paragraph()

# ========================================================
# SECTION 0 — POSTMORTEM
# ========================================================
add_heading(doc, "Section 0 — Yesterday's Postmortem", level=1, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc, "Brief Accuracy Score: N/A — Session 1 (Inaugural run of this brief format)", bold=True, size=12)
add_bullets(doc, [
    "No prior brief exists in this conversation to audit. This is the baseline session.",
    "Postmortem framework begins in earnest from Session 2 (Monday, 8 June 2026).",
    "Today's methodology adjustment: None — fresh baseline. All rules applied at full strength.",
    "NOTE: The most recent brief in archive was 23 April 2026 (WAIT day, different context). "
    "That brief is too far in the past (43 days) to score against today's conditions.",
])

doc.add_paragraph()

# ========================================================
# SECTION 1 — MACRO SNAPSHOT
# ========================================================
add_heading(doc, "Section 1 — Macro Snapshot (Multi-Source Verified)", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Metric", "Reading", "Signal"],
    rows=[
        ["WTI Crude (NYMEX)", "~$93.86/bbl (range $93–$95.84 across sources)", "BEAR crude — primary driver"],
        ["Brent Crude (context only)", "~$96.97/bbl (down 0.86% June 4)", "DO NOT use for MCX proxy"],
        ["Brent–WTI Spread", "~$3.10 (narrowed from $7–10 during Iran spike)", "Spread compression — monitor"],
        ["MCX Crude ACTUAL", "₹9,234/bbl (June 4 close)", "BEAR regime ACTIVE (9,000–10,000)"],
        ["WTI × USDINR proxy", "₹8,984/bbl ($93.86 × 95.72)", "CAUTION zone — DISCREPANCY noted ⚠"],
        ["REGIME VERDICT", "MCX ACTUAL used as operative — BEAR", "Calls BLOCKED except ONGC/OIL/MRPL"],
        ["USDINR", "95.72 (June 4 close)", "Near record lows — FII caution"],
        ["India VIX", "~16.28 (June 4 close, range 15.53–16.39)", "HALF SIZE rule applies (15–20 band)"],
        ["Nifty 50", "23,416.55 (June 4 close, +0.05%)", "Flat; no directional conviction"],
        ["Nifty 50 intraday", "Opened ~23,290 (-0.53%), recovered to close flat", "Buyers emerged at 23,250–23,300"],
        ["Sensex", "74,360.01 (+0.02%)", "Flat — mirrors Nifty"],
        ["GIFT Nifty (last read)", "~23,372.5 (pre-June 5 data)", "~44 pts below close → slight neg open"],
        ["Nifty Max Pain (June)", "~23,396", "Gravity zone — very close to close"],
        ["US Dow Jones", "51,561.93 (+1.73%) — ALL-TIME RECORD", "Positive macro backdrop"],
        ["S&P 500", "7,584.31 (+0.41%)", "Positive"],
        ["Nasdaq", "26,830.96 (-0.09%)", "Slight tech drag — Indian IT headwind"],
        ["Nikkei 225 (June 3)", "~68,402 (+2.5% on June 3)", "Asian bull — context only"],
        ["Hang Seng (June 3)", "Down ~1.62% in last hour", "Weak EM Asia"],
        ["FII Flow (June 4)", "NET SOLD ₹4,447 Cr (cash) + sold 264,568 Nifty futures", "BEARISH — consistent selling"],
        ["DII Flow (June 3)", "NET BOUGHT ₹5,741 Cr", "Absorbing FII; floor present"],
        ["FII cumulative trend", "Consistent net sellers in 2026", "Structural ceiling on rallies"],
        ["Monsoon (IMD)", "90% of LPA — below normal; onset June 4 (3 days late)", "Slow-burn bearish — food inflation"],
        ["RBI Repo Rate", "5.25% (current); DECISION TODAY 10:00 AM IST", "THE KEY EVENT of the day"],
    ])

doc.add_paragraph()

# ========================================================
# SECTION 2 — CRUDE REGIME
# ========================================================
add_heading(doc, "Section 2 — Crude Regime Analysis", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "MODE: BEAR (MCX ₹9,234) — Active", bold=True, size=13, color=RGBColor(0xCC, 0x00, 0x00))
add_bullets(doc, [
    "WTI NYMEX: ~$93.86/bbl | USDINR: 95.72 | WTI proxy = ₹8,984 (CAUTION zone)",
    "MCX ACTUAL: ₹9,234/bbl — operative number used → BEAR regime confirmed",
    "DISCREPANCY: WTI proxy (₹8,984) says CAUTION; MCX actual (₹9,234) says BEAR. "
    "MCX actual takes precedence. Gap of ₹250 likely due to roll costs + quality premium.",
    "BOUNDARY FLAG: WTI proxy is only ₹16 above the 9,000 CAUTION/BEAR line. "
    "If Iran deal confirmed + crude falls, regime could flip to CAUTION within ONE session.",
    "Crude direction: Brent fell 0.86% on June 4. WTI range $93–$95. Stabilising but elevated.",
    "Iran MOU (unconfirmed) = potential crude downside catalyst IF formally signed by Trump.",
    "Strait of Hormuz: Contested. Iran threatening closure while simultaneously discussing MOU.",
])

doc.add_paragraph()
add_para(doc, "Crude Regime Impact by Sector:", bold=True, size=11)
add_kv_table(doc,
    header=["Sector", "Crude Regime Impact", "Direction"],
    rows=[
        ["ONGC / OIL India (upstream)", "BULLISH — every $1 crude = large earnings uplift", "CALL (exception)"],
        ["MRPL (upstream refiner/ONGC sub)", "BULLISH — explicit BEAR regime exception", "CALL (exception)"],
        ["BPCL / HPCL / IOC (downstream)", "BEARISH — crude is input cost; margin squeeze", "PUT or AVOID"],
        ["IndiGo / Aviation", "BEARISH — ATF price pressure; Q4 loss reported", "PUT"],
        ["Industrials / Defence (BEL, HAL)", "NEUTRAL-POSITIVE — non-crude correlated", "CALL (macro hedge)"],
        ["FMCG (HUL, Nestle)", "BEARISH slow-burn — monsoon fail + input costs", "AVOID CALLS"],
        ["IT (TCS, Infosys)", "NEUTRAL-NEGATIVE — Nasdaq drag; USD/INR mixed", "AVOID today"],
        ["Banking (SBI, ICICI)", "BINARY — RBI MPC dependent; no pre-position", "WAIT post-10 AM"],
    ])

doc.add_paragraph()

# ========================================================
# SECTION 3 — NEWS & CATALYSTS
# ========================================================
add_heading(doc, "Section 3 — News & Catalyst Digest", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "CAT A — Geopolitical (Highest Priority)", bold=True, size=12, color=RGBColor(0xCC, 0x00, 0x00))
add_bullets(doc, [
    "Iran MOU status: Preliminary 60-day ceasefire extension MOU reportedly agreed — NEEDS Trump's final approval. "
    "Treat as UNCONFIRMED until formally announced. Rule: 3–5 sessions validation required.",
    "Simultaneously: Iran accusing US of ceasefire violations and threatening to close Strait of Hormuz again. "
    "This is a DUAL-SIGNAL DIVERGENCE. Media running the optimistic headline. Smart money aware of collapse risk.",
    "Iran ceasefire MoUs have collapsed within 48 hours TWICE in 2026. Do not trade the peace narrative.",
    "Strait of Hormuz: PARTIALLY contested. Not fully open. Not fully closed. Risk of sudden disruption.",
    "Israel-Lebanon: Israel Defence Minister confirmed continued strikes on Lebanon.",
    "Trump social media: Watch for any formal Iran deal announcement. Only CONFIRMED signing = actionable.",
    "US-Iran exchange of strikes noted in recent days — conflict has NOT fully stopped despite MOU language.",
])

doc.add_paragraph()
add_para(doc, "CAT B — Domestic India", bold=True, size=12, color=RGBColor(0x1F, 0x38, 0x64))
add_bullets(doc, [
    "RBI MPC DECISION TODAY at 10:00 AM IST — Governor press conference at 12:00 PM IST.",
    "Consensus: HOLD at 5.25%, NEUTRAL stance. High crude + weak rupee + monsoon risk = no cut room.",
    "Scenario A (Hold + Dovish): Nifty tests 23,550–23,600. Banks rally. PSU banks outperform.",
    "Scenario B (Hold + Neutral): Nifty range-bound 23,350–23,500. No strong directional move.",
    "Scenario C (Hold + Hawkish signal on inflation): Nifty drops to 23,150–23,200. Banks fall hard.",
    "Monsoon: IMD confirms below-normal (90% LPA). Onset reached Kerala June 4 (3 days delayed). "
    "El Niño conditions developing. Food inflation risk building — supports RBI hawkish lean.",
    "SEBI/RBI no other major regulatory announcements identified.",
])

doc.add_paragraph()
add_para(doc, "CAT C — Corporate Catalysts", bold=True, size=12, color=RGBColor(0x1F, 0x38, 0x64))
add_bullets(doc, [
    "IndiGo Q4 FY26: Reported ₹2,536 Cr net LOSS. Analysts remain 'constructive long-term' but near-term weak.",
    "ATF government cap: Temporary intervention — ATF rose from ₹60.5/L to ₹142/L (May 2026). Cap is political, not structural.",
    "ONGC CMP ~₹264–265 (June 2–3 data). Multiple brokerages: Motilal Oswal NEUTRAL (TP ₹265), "
    "Emkay ADD (TP ₹315), CLSA HIGH CONVICTION OUTPERFORM (TP ₹405 on Brent crude thesis).",
    "BEL: CMP ~₹410. Record order book ₹75,000 Cr. Revenue growth 20–25%. Nuvama top pick (March 2026).",
    "F&O Ban list (June 2): AMBER, KAYNES were in ban. VERIFY today's ban list at 9:05 AM before ANY trade.",
    "No major IPO listings flagged today.",
])

doc.add_paragraph()
add_para(doc, "CAT D — Global Macro", bold=True, size=12, color=RGBColor(0x1F, 0x38, 0x64))
add_bullets(doc, [
    "US Dow at RECORD 51,561.93 (+1.73%). Key nuance: rotation FROM tech TO non-tech (cyclicals).",
    "Nasdaq SLIGHTLY RED (-0.09%): Chip/growth selling. DIRECTLY bearish for Indian IT sector — TCS, Infosys.",
    "This Dow record ≠ straightforward EM rally. FII are still SELLING India into global risk-on.",
    "China PMI / Asian macro: Hang Seng weak. Kospi mixed. No strong Asia-led tailwind.",
    "US Fed: No major event today. Watch for any Fed governor speak on tape re: inflation.",
    "Gold: No specific data retrieved — expect elevated given geopolitical uncertainty.",
])

doc.add_paragraph()
add_para(doc, "CAT E — F&O Specific", bold=True, size=12, color=RGBColor(0x1F, 0x38, 0x64))
add_bullets(doc, [
    "NSE Monthly expiry: Last Tuesday of June = ~June 30, 2026 (VERIFY exact date on NSE).",
    "NOTE: A prior brief noted NSE moved weekly expiry to Tuesday. Monthly schedule may also have changed.",
    "Max Pain (June): ~23,396 — very close to current price 23,417. Gravitational pull is STRONG near expiry.",
    "OI build-up: Call OI heavy at 23,500 CE (ceiling). Put OI concentrated at 23,000 (floor).",
    "PCR: Neutral-to-slightly bearish. Call writers have upper hand near 23,500.",
    "Smart money positioning: Range-bound 23,000–23,500 appears likely unless RBI shock.",
    "F&O Ban: AMBER, KAYNES (June 2 data) — reverify for June 5. Do NOT trade banned securities.",
])

doc.add_paragraph()

# ========================================================
# SECTION 4 — LAYER 1/2/3 ANALYSIS
# ========================================================
add_heading(doc, "Section 4 — Contrarian Check (Layer 1 / 2 / 3)", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "LAYER 1 — Surface Consensus (What retail is saying):", bold=True, size=12)
add_para(doc,
    "\"Dow hit an all-time record, Iran is getting a peace deal, RBI will hold and maybe signal a cut later. "
    "Markets should gap up and rally. Buy financials and ride the global risk-on. It's a bull day.\"",
    italic=True, size=11)

doc.add_paragraph()
add_para(doc, "LAYER 2 — What the Consensus is Likely WRONG About:", bold=True, size=12, color=RGBColor(0xCC, 0x00, 0x00))
add_bullets(doc, [
    "The Iran MOU is simultaneously the most optimistic and most fragile development. "
    "Iran is ALSO accusing the US of ceasefire violations today — two contradictory signals. "
    "Media leads with the MOU. Smart money remembers the two prior MOU collapses in 2026.",
    "GIFT Nifty suggests a NEGATIVE open (~-44 pts), NOT a gap-up. The 'Dow record = Indian gap-up' "
    "narrative is being contradicted by the pre-open futures data.",
    "FII sold ₹4,447 Cr on June 4 while the Dow was rising. FII are rotating OUT of India during global risk-on. "
    "This is structurally bearish and consensus is ignoring it.",
    "Retail stops are clustered above 23,500 CE. If RBI gives even a slightly hawkish signal, "
    "smart money will push into those stops — selling 23,500 CEs aggressively.",
    "Nasdaq is SLIGHTLY RED — chip/tech rotation. Indian IT (TCS, Infosys) will NOT benefit from "
    "the Dow record. These are different indices with different compositions.",
])

doc.add_paragraph()
add_para(doc, "LAYER 3 — What Everyone is Underestimating (Contrarian):", bold=True, size=12, color=RGBColor(0x1F, 0x38, 0x64))
add_bullets(doc, [
    "MONSOON SLOW-BURN: IMD confirmed below-normal (90% LPA), delayed onset, El Niño developing. "
    "Markets are 100% focused on Iran + RBI. Nobody is pricing the cumulative food inflation risk "
    "that a bad monsoon builds over June–August. FMCG stocks look 'defensive' — they are NOT "
    "if monsoon fails. This trade builds over weeks, not today, but the positioning starts now.",
    "BRENT–WTI SPREAD COMPRESSION: The spread dropped from $7–10 to ~$3.10. This is unusual. "
    "If Iran deal IS confirmed, Brent falls FASTER than WTI (Brent has more Iran-risk premium). "
    "This would rapidly improve OMC downstream margins (BPCL/HPCL/IOC). "
    "An unexpected short-squeeze in downstream OMCs is possible if crude falls sharply.",
    "NASDAQ DRAG ON IT: TCS and Infosys follow Nasdaq, not Dow. With Nasdaq marginally red and "
    "FY27 guidance already soft (Infosys: 1.5–3.5% revenue growth), IT stocks face "
    "double pressure today. Most analysts are focused on banking for the RBI play — "
    "the IT underperformance is flying under the radar.",
    "MCX REGIME BOUNDARY: The WTI proxy (₹8,984) is ALREADY in CAUTION zone — only ₹16 from "
    "the 9,000 line. If crude falls ₹300–400 today on Iran news, the ENTIRE strategy framework "
    "flips: BEAR → CAUTION, more calls become available. Watch this number actively.",
])

doc.add_paragraph()
add_kv_table(doc,
    header=["Scenario", "Trigger", "Action"],
    rows=[
        ["BULLISH FLIP", "Trump formally confirms Iran deal + MCX crude drops below ₹8,500",
         "Regime shifts CAUTION → more calls valid. Watch aviation (IndiGo) for reversal."],
        ["BEARISH FLIP", "Iran denies MOU / Hormuz closure news + RBI hawkish",
         "Crude spikes. STRONG BEAR. All longs stopped. Widen puts. Quarter-size only."],
        ["BASE CASE", "RBI holds, neutral tone. Iran ambiguous. Crude ₹9,000–9,300.",
         "BEAR regime holds. PUT-heavy portfolio. Range 23,300–23,500."],
    ])

doc.add_paragraph()
add_para(doc, "ENTRY PERMISSION: YELLOW — No trades before 10:30 AM IST (post-RBI decision)", bold=True,
         size=12, color=RGBColor(0xE6, 0x91, 0x38))

doc.add_paragraph()

# ========================================================
# SECTION 5 — NIFTY STRUCTURE
# ========================================================
add_heading(doc, "Section 5 — Nifty / Bank Nifty Structure", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Level", "Value", "Meaning"],
    rows=[
        ["Strong Resistance", "23,750", "Multiple tops; heavy supply zone"],
        ["Resistance R2", "23,600–23,750", "Secondary supply; RBI dovish scenario target"],
        ["Resistance R1", "23,500", "Immediate ceiling; OI-heavy call strike"],
        ["Previous Close", "23,416.55", "Sitting in no-man's land between S/R"],
        ["Max Pain (June)", "~23,396", "Gravity magnet — extremely close to close"],
        ["Support S1", "23,300–23,250", "First meaningful buyer zone"],
        ["Support S2", "23,000", "Critical floor — break = acceleration to 22,700"],
        ["200 DMA", "~23,000 area", "Approximate; provides long-term support"],
        ["Bank Nifty Resistance", "53,800–54,000", "Immediate supply; RBI event binary"],
        ["Bank Nifty Support", "52,800–53,000", "Key floor; break = 52,500"],
    ])

doc.add_paragraph()
add_para(doc, "Expected Nifty Range Today:", bold=True, size=11)
add_bullets(doc, [
    "Pre-RBI (9:15–10:00 AM): Tight range 23,350–23,450. Low volume, max pain gravity.",
    "Post-RBI DOVISH (Hold + rate cut signal): 23,500–23,600. Banks lead.",
    "Post-RBI NEUTRAL (Hold + neutral stance): 23,300–23,500. Choppy range.",
    "Post-RBI HAWKISH (Hold + inflation warning): 23,150–23,300. Banks fall. Puts activated.",
    "THE ONE NUMBER: ₹9,000 MCX crude — this, not Nifty levels, determines today's framework.",
])

doc.add_paragraph()

# ========================================================
# SECTION 6 — STOCK SETUPS
# ========================================================
add_heading(doc, "Section 6 — Stock Setups", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc,
    "BEAR crude regime: CALL positions BLOCKED except ONGC / OIL India / MRPL. "
    "VIX 16.28 = HALF SIZE. No entries before 10:30 AM (post-RBI). Verify F&O ban list by 9:05 AM.",
    italic=True, bold=False, size=10, color=RGBColor(0xCC, 0x00, 0x00))

doc.add_paragraph()

# --- STOCK 1: MRPL ---
add_para(doc, "STOCK 1 — MRPL (CALL) — Grade: A+", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP", "~₹150–155 (estimate from May 29 data ~₹148.70; VERIFY at pre-open)"),
    ("Lot Size", "~2,500–3,750 shares (VERIFY on NSE — CRITICAL for capital check)"),
    ("Sector", "Energy / Upstream Refining | F&O: Yes"),
    ("Capital Check", "IF lot=2,500: 2,500×152×3%=₹11,400 ✓ PASSES | IF lot=3,750: ₹17,100 ✗ FAILS. "
     "DO NOT TRADE if lot×CMP×3% > ₹15,000."),
    ("Why A+", "BEAR crude regime EXPLICIT EXCEPTION. ONGC subsidiary aligned to upstream crude cycle. "
     "MCX at ₹9,234 supports upstream valuation. DII likely holding PSU complex. "
     "No negative corporate catalyst identified. Multiple crude-sector brokerages constructive."),
    ("Why NOT A++", "Capital check uncertainty (lot size unconfirmed). Refinery margin dynamics complex "
     "when crude is input cost. Fundamentally upstream producers (ONGC/OIL) are cleaner crude plays."),
    ("Option Setup", "MRPL JUN 155 CE or 160 CE | Monthly expiry (verify: ~June 25 or June 30)"),
    ("Entry Trigger — ALL 4 required", "1) India VIX ≤ 18 at entry time | "
     "2) MRPL holds above ₹148 on 15-min close basis | "
     "3) MCX Crude ≥ ₹9,000 (regime not slipping to CAUTION) | "
     "4) Time: After 10:30 AM (post-RBI decision clarity)"),
    ("Target 1 (40% book)", "₹165 | Exit 40% of position"),
    ("Target 2 (40% book)", "₹172 | Exit 40% of position"),
    ("Remaining 20%", "Trail with 15-min SL; let run if crude stays above ₹9,000"),
    ("Stop Loss", "₹143 — 15-min CLOSE below this level = exit FULL position immediately"),
    ("R:R", "Risk ~₹7–9 | Reward ₹12–20 | Ratio ~2.2:1 ✓"),
    ("Time Stop", "Exit by 2:30 PM if T1 not hit"),
    ("Invalidation", "MCX crude drops below ₹8,800 (regime flip incoming) OR F&O ban list includes MRPL"),
])

doc.add_paragraph()

# --- STOCK 2: IndiGo ---
add_para(doc, "STOCK 2 — IndiGo / InterGlobe Aviation (PUT) — Grade: A", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP", "~₹2,900–₹3,200 (down ~11–25% in 2026; VERIFY at pre-open)"),
    ("Lot Size", "~150 shares (estimate; VERIFY on NSE)"),
    ("Sector", "Aviation | F&O: Yes"),
    ("Capital Check", "IF CMP=₹3,000, lot=150: 150×3,000×3%=₹13,500 ✓ PASSES | "
     "IF CMP=₹3,500: ₹15,750 ✗ FAILS. Verify before entry."),
    ("Why A", "BEAR crude = ATF (aviation turbine fuel) cost pressure. Q4 FY26 LOSS of ₹2,536 Cr reported. "
     "ATF rose from ₹60.5/L to ₹142/L; government cap is political intervention, not structural. "
     "Morgan Stanley flagged near-term margin pressure. FII selling aviation as discretionary sector. "
     "Nasdaq slightly negative = IndiGo's premium growth valuation compresses with risk-off."),
    ("What could go wrong", "If Iran deal formally confirmed + crude drops sharply, IndiGo could rally 5–8% "
     "on ATF cost reversal. Monitor MCX crude continuously. If crude drops below ₹8,500 = exit put."),
    ("Option Setup", "IndiGo JUN PE | Strike: ~3% below CMP (ATM or slight OTM put) | Monthly expiry"),
    ("Entry Trigger — ALL 4 required", "1) IndiGo opens below previous close OR fails to hold above ₹3,000 "
     "in first 30 min | 2) MCX Crude stays ≥ ₹9,000 | "
     "3) RBI decision: HOLD with neutral or hawkish tone (not a surprise cut) | "
     "4) Time: After 10:30 AM"),
    ("Target 1 (40% book)", "₹200 below entry CMP (e.g., if entry ~₹3,000, T1 = ₹2,800)"),
    ("Target 2 (40% book)", "₹350 below entry CMP (e.g., T2 = ₹2,650)"),
    ("Stop Loss", "15-min CLOSE above entry + ₹150 = exit FULL position"),
    ("R:R", "~2.3:1 estimated"),
    ("Time Stop", "2:30 PM if T1 not hit"),
    ("Invalidation", "MCX crude drops below ₹8,500 (regime flip) OR Iran deal formally confirmed"),
])

doc.add_paragraph()

# --- STOCK 3: BEL ---
add_para(doc, "STOCK 3 — BEL / Bharat Electronics (CALL) — Grade: A  ⚠ WATCH-ONLY (Capital Constraint)", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP", "~₹410 (June 3–4 data; VERIFY)"),
    ("Lot Size", "1,425 shares (confirmed from NSE data)"),
    ("Sector", "Defence Electronics — MACRO HEDGE | F&O: Yes"),
    ("Capital Check", "1,425 × 410 × 3% = ₹17,527 — EXCEEDS ₹15,000 limit ✗ FAILS"),
    ("Cannot trade unless", "CMP drops to ≤ ₹350 (1,425 × 350 × 3% = ₹14,963 ✓) "
     "OR NSE reduces lot size. PRESENTED AS WATCH-ONLY."),
    ("Why MACRO HEDGE (A grade thesis)", "Geopolitical instability (Iran/Hormuz) = continued "
     "defence budget spending. Record order book ₹75,000 Cr — 3–4 year revenue visibility. "
     "20–25% revenue growth trajectory. Non-crude correlated (electronics, not fuel). "
     "Nuvama top pick. Provides portfolio hedge against Iran ESCALATION scenario."),
    ("Action today", "OBSERVE only. Do not trade. If CMP pulls back to ₹345–350, reassess. "
     "This is the setup to have on radar for Session 2 onward."),
    ("If capital constraint lifts", "Entry: BEL JUN 420 CE | T1: ₹435 | T2: ₹455 | SL: ₹395 close | R:R ~2.5:1"),
])

doc.add_paragraph()

# --- STOCK 4: NTPC ---
add_para(doc, "STOCK 4 — NTPC (PUT) — Grade: B+", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP", "~₹360–380 (estimate; VERIFY at open)"),
    ("Lot Size", "~1,125 shares (estimate; VERIFY on NSE)"),
    ("Sector", "Power / Utilities | F&O: Yes"),
    ("Capital Check", "IF CMP=₹370, lot=1,125: 1,125×370×3%=₹12,488 ✓ PASSES"),
    ("Why B+", "Power sector faces higher fuel costs from elevated crude (gas-linked inputs). "
     "NTPC's regulated returns cap upside; FII broadly selling PSU utilities. "
     "Below-normal monsoon = lower hydro complement = higher thermal load = more fuel costs. "
     "No near-term positive catalyst identified."),
    ("Why not higher grade", "Single-thesis setup (no catalyst + technical = B+). "
     "NTPC moves slowly — needs patience. Conditional on RBI neutral/hawkish tone."),
    ("Option Setup", "NTPC JUN 360 PE or 350 PE | Monthly expiry"),
    ("Entry Trigger — ALL 4 required", "1) NTPC fails to hold above ₹365 in first 30 min | "
     "2) RBI: HOLD with neutral or hawkish language | "
     "3) India VIX < 20 (not spiking) | "
     "4) Broader market weak: Nifty below 23,350 post-RBI"),
    ("Target 1 (40% book)", "₹352"),
    ("Target 2 (40% book)", "₹340"),
    ("Stop Loss", "₹382 — 15-min close above = exit"),
    ("R:R", "~2.2:1"),
    ("Time Stop", "1:30 PM (utilities move slowly — don't hold a short utility into afternoon)"),
])

doc.add_paragraph()

# --- STOCK 5: SBI ---
add_para(doc, "STOCK 5 — SBI (WATCH ONLY — Do NOT Pre-Position)", bold=True, size=13, color=RGBColor(0x99, 0x00, 0x00))
stock_card(doc, [
    ("CMP", "~₹956 (June 2 data; verify at open)"),
    ("Lot Size", "750 shares (confirmed)"),
    ("Capital Check", "750 × 956 × 3% = ₹21,510 — FAILS ₹15K limit at current price"),
    ("Why watch-only", "RBI MPC TODAY = complete binary outcome. Direction 100% depends on RBI tone. "
     "Capital check FAILS regardless. Pre-positioning is pure coin-flip with no edge."),
    ("Post-RBI playbook (10:30 AM onward)", "If RBI DOVISH → Watch SBI above ₹970 sustained = CALL bias | "
     "If RBI HAWKISH → Watch SBI below ₹940 sustained = PUT bias"),
    ("Capital note", "SBI requires capital base >₹2,50,000 or lot size reduction to trade at this price. "
     "Document this constraint for future sessions."),
])

doc.add_paragraph()

# ========================================================
# SECTION 7 — SECTOR DIVERSIFICATION + AVOID LIST
# ========================================================
add_heading(doc, "Section 7 — Sector Coverage & Avoid List", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "Sector Diversification Summary:", bold=True, size=12)
add_kv_table(doc,
    header=["Sector", "Stock", "Direction", "Grade", "Status"],
    rows=[
        ["Energy / Upstream", "MRPL", "CALL", "A+", "LIVE (verify lot/capital)"],
        ["Aviation", "IndiGo", "PUT", "A", "LIVE (verify lot/capital)"],
        ["Defence Electronics", "BEL", "CALL", "A", "WATCH-ONLY (capital fail)"],
        ["Power / Utilities", "NTPC", "PUT", "B+", "LIVE (verify lot/capital, conditional)"],
        ["Banking / PSU", "SBI", "Watch", "—", "WATCH-ONLY (capital fail + RBI binary)"],
    ])

doc.add_paragraph()
add_kv_table(doc,
    header=["Requirement", "Status"],
    rows=[
        ["Min 4 sectors", "5 sectors covered ✓"],
        ["Max 2 per sector", "1 per sector ✓"],
        ["CALL + PUT mix", "2 CALLs + 2 PUTs ✓"],
        ["Macro hedge (defence)", "BEL (watch-only but thesis documented) ✓"],
        ["No more than 7 stocks", "5 total ✓"],
    ])

doc.add_paragraph()
add_para(doc, "STOCKS TO ACTIVELY AVOID TODAY:", bold=True, size=12, color=RGBColor(0xCC, 0x00, 0x00))
add_kv_table(doc,
    header=["Stock / Sector", "Reason"],
    rows=[
        ["ONGC", "Capital FAILS (2,250 × 265 × 3% = ₹17,888 > ₹15K). Despite crude exception, cannot trade."],
        ["OIL India", "Capital check needed — likely fails at current price/lot. Verify before considering."],
        ["BPCL / HPCL / IOC", "Downstream OMCs — NOT in BEAR regime exceptions. High crude = margin squeeze. Avoid calls."],
        ["TCS / Infosys", "Nasdaq mildly red. IT follows Nasdaq not Dow. FY27 guidance soft. No catalyst."],
        ["AMBER", "Was in F&O BAN June 2. Verify June 5 status. AVOID until confirmed off ban."],
        ["KAYNES", "Was in F&O BAN June 2. Verify June 5 status. AVOID until confirmed off ban."],
        ["All weekly options (expiry today)", "June 5 is Friday. Weekly expiry = ABSOLUTE PROHIBITION."],
        ["All FMCG CALLS (HUL, Nestle, etc.)", "Below-normal monsoon building bearish case. Not a safe haven."],
        ["IndiGo CALLS", "Crude bearish = ATF cost pressure. Only PUT setup is valid for IndiGo."],
        ["Any banking stock pre-10 AM", "RBI MPC binary — no pre-positioning regardless of directional view."],
    ])

doc.add_paragraph()

# ========================================================
# SECTION 8 — FINAL VERDICT
# ========================================================
add_heading(doc, "Section 8 — Final Verdict", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_kv_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["Today's Bias", "CAUTIOUSLY BEARISH to NEUTRAL"],
        ["Confidence", "MEDIUM (RBI creates genuine two-way uncertainty)"],
        ["Crude Mode", "BEAR — MCX ₹9,234 | Boundary watch: WTI proxy only ₹16 from CAUTION line"],
        ["VIX Sizing", "HALF SIZE (VIX 16.28 — confirm at open; if >20, downgrade to QUARTER)"],
        ["Entry Permission", "YELLOW — NO entries before 10:30 AM IST (post-RBI)"],
        ["Expected Nifty Range", "23,200–23,600 (RBI-outcome dependent)"],
        ["Key Event 1", "RBI MPC at 10:00 AM — THE pivot event of the session"],
        ["Key Event 2", "Iran MOU formal confirmation OR breakdown (could arrive any time)"],
        ["Key Level", "₹9,000 MCX Crude — regime flip level to monitor all session"],
        ["Max Pain gravity", "23,396 — Nifty may drift here if no strong directional event"],
    ])

doc.add_paragraph()
add_para(doc, "BULL CASE:", bold=True, size=11, color=RGBColor(0x00, 0x80, 0x00))
add_para(doc,
    "RBI holds + gives DOVISH signal (acknowledges growth support) → Banks rally → Nifty tests 23,550–23,600. "
    "SIMULTANEOUSLY, Trump formally confirms Iran deal → Crude drops to ₹8,500–8,800 → Regime flips to CAUTION "
    "→ More call setups open → IndiGo PUT exits → Aviation calls become valid. A coordinated positive outcome.",
    size=11)

doc.add_paragraph()
add_para(doc, "BEAR CASE:", bold=True, size=11, color=RGBColor(0xCC, 0x00, 0x00))
add_para(doc,
    "RBI holds + gives HAWKISH signal (crude + monsoon inflation risk cited) → Banks fall → Nifty breaks 23,300 "
    "→ Tests 23,150–23,200. SIMULTANEOUSLY, Iran denies MOU / threatens Hormuz → Crude spikes → "
    "STRONG BEAR territory → All longs stopped → Only crude-exception CALLs (MRPL) survive → "
    "VIX spikes above 20 → Downgrade to QUARTER SIZE.",
    size=11)

doc.add_paragraph()
add_para(doc, "MOST LIKELY SCENARIO:", bold=True, size=12, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "RBI HOLD + NEUTRAL stance. Nifty oscillates in 23,300–23,500 band post-announcement. "
    "Iran situation remains ambiguous (no formal deal, no formal breakdown). MCX crude ₹9,000–9,300. "
    "FII continues selling; DII absorbs. BEAR regime stays active. Volume remains below-average "
    "as market awaits next major catalyst. Day ends near max pain (23,396).",
    size=11)

doc.add_paragraph()

add_para(doc, "THE ONE NUMBER THAT MATTERS:", bold=True, size=14, color=RGBColor(0xCC, 0x00, 0x00),
         align=WD_ALIGN_PARAGRAPH.CENTER)
add_para(doc,
    "₹9,000 — MCX Crude. Drop below = regime flips CAUTION, strategy opens up. "
    "Rise above ₹9,500 = STRONG BEAR, quarter-size only.",
    bold=True, size=13, color=RGBColor(0xCC, 0x00, 0x00), align=WD_ALIGN_PARAGRAPH.CENTER)

doc.add_paragraph()

# ========================================================
# SECTION 9 — PAPER TRADING NOTE
# ========================================================
add_heading(doc, "Section 9 — Paper Trading Log Note", level=1, color=RGBColor(0x1F, 0x38, 0x64))
add_bullets(doc, [
    "System likely sees: Cautious slightly-negative open (GIFT Nifty -44 pts). Tight pre-RBI range 9:15–10:00 AM.",
    "Any signal firing 9:15–10:00 AM: FLAG AS NOISE. Log it but do not execute. RBI creates binary outcome.",
    "Post-10:30 AM signals: VALID. First confirmed signal after 10:30 AM = first meaningful data point of Session 1.",
    "Watch MRPL and IndiGo specifically — log whether signal fires and in which direction. "
    "This validates the regime identification.",
    "MCX crude watch: Log ₹9,234 (open proxy). Note if it breaks ₹9,000 or rises to ₹9,500 at any point.",
    "F&O ban check at 9:05 AM: If MRPL or IndiGo in ban, log 'SETUP INVALIDATED — BAN' and observe only.",
    "Zero-trade count: This is Session 1. Paper system's zero-trade run counter starts fresh today.",
    "Mandate: Do NOT reduce confirmation gate requirements regardless of zero-trade frustration. "
    "Data accumulates. Week 2 backtest begins Monday June 8.",
])

doc.add_paragraph()

# ========================================================
# ONE-LINE SUMMARY
# ========================================================
add_heading(doc, "One-Line Summary", level=1, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "BEAR crude (MCX ₹9,234) forces PUT-heavy portfolio with only MRPL CALL as the regime exception; "
    "RBI MPC at 10 AM demands zero entries before 10:30 AM; Iran MOU is unconfirmed and historically "
    "fragile — trade the ambiguity not the peace; IndiGo PUT (ATF pressure, Q4 loss) and NTPC PUT "
    "(utility/fuel cost squeeze) are the core setups; BEL is the macro hedge to build toward when "
    "capital allows; ₹9,000 MCX crude is the pivot number that determines the ENTIRE day's framework; "
    "VIX 16.28 = half-size only; Dow record is a mirage for India — FII are still selling.",
    italic=True, bold=True, size=11)

doc.add_paragraph()

# FOOTER
foot = doc.add_paragraph()
foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = foot.add_run(
    "Generated by Daily Trading Routine — 5 June 2026 — for personal educational use only. "
    "Not financial advice. Verify all prices and lot sizes at market open."
)
fr.italic = True
fr.font.size = Pt(9)
fr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

doc.save(OUT_PATH)
print(f"Wrote {OUT_PATH}")
