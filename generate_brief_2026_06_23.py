"""Generate the daily trading-intelligence .docx brief — June 23, 2026."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO   = "2026-06-23"
DATE_LONG  = "Tuesday, 23 June 2026"
BIAS       = "MILDLY_BULLISH"
ENTRY_PERM = "YELLOW"
VIX_MODE   = "QUARTER SIZING (VIX 27.32)"
CRUDE_MODE = "AGGRESSIVE BULL  (MCX ₹7,220)"

OUT_PATH = f"trading-briefs/2026/06-June/Trading_Brief_{DATE_ISO}_{ENTRY_PERM}.docx"

RED    = RGBColor(0xC0, 0x00, 0x00)
NAVY   = RGBColor(0x1F, 0x38, 0x64)
GREEN  = RGBColor(0x1E, 0x8B, 0x1E)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
AMBER  = "F1C232"
ORANGE = "E69138"
DKGRN  = "1E8B1E"
DKRED  = "C00000"
BLUE   = "1F3864"


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


def add_kv_table(doc, rows, header=None, shade_header=True, col_widths=None):
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
                run.font.color.rgb = WHITE
        r_idx = 1
    for i, row in enumerate(rows):
        trow = table.rows[r_idx + i]
        for j, val in enumerate(row):
            cell = trow.cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            run.font.size = Pt(10)
            if j == 0:
                run.bold = True
    return table


def badge_line(doc, label, value, fill_hex, text_color=WHITE):
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
        r1.font.size = Pt(10)
        shade_cell(lc, "D9E1F2")
        p2 = rc.paragraphs[0]
        r2 = p2.add_run(v)
        r2.font.size = Pt(10)


# ============================================================
doc = Document()
for section in doc.sections:
    section.top_margin    = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin   = Cm(1.8)
    section.right_margin  = Cm(1.8)

# ────────────────── HEADER ──────────────────
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run("ELITE DAILY TRADING INTELLIGENCE BRIEF")
run.bold = True
run.font.size = Pt(20)
run.font.color.rgb = NAVY

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub.add_run("Indian F&O Day Trader — Nifty + Stocks").italic = True

date_p = doc.add_paragraph()
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
dr = date_p.add_run(f"Date: {DATE_LONG}  |  Generated 08:30 IST  |  Capital: ₹2,00,000")
dr.bold = True
dr.font.size = Pt(11)

badge_line(doc, "TODAY'S BIAS", BIAS, AMBER, text_color=RGBColor(0x00,0x00,0x00))
badge_line(doc, "ENTRY PERMISSION", ENTRY_PERM + " — wait 45-60 min after open", ORANGE)
badge_line(doc, "CRUDE MODE", CRUDE_MODE, DKGRN)
badge_line(doc, "VIX SIZING RULE", VIX_MODE, DKRED)

doc.add_paragraph()

# ────────────────── SECTION 0 — POSTMORTEM ──────────────────
add_heading(doc, "Section 0 — Yesterday's Postmortem", level=1, color=NAVY)
add_para(doc,
    "PRIOR BRIEF STATUS: No prior brief available in this session (fresh session initialisation). "
    "This is Session Day 1 — no postmortem score calculable. "
    "Methodology calibration begins today. Score tracked from tomorrow onward.",
    italic=True, size=10, color=RGBColor(0x80,0x80,0x80))
add_para(doc,
    "WHAT THIS MEANS: No adjustment to today's methodology. Proceed with standard framework. "
    "Zero-trade day counter: N/A (first session).",
    bold=False, size=10)

# ────────────────── SECTION 1 — MACRO SNAPSHOT ──────────────────
doc.add_paragraph()
add_heading(doc, "Section 1 — Macro Snapshot (Cross-Verified)", level=1, color=NAVY)

add_para(doc, "⚠  DATA NOTES: WTI sources show two readings — $76.54 (current morning) vs $74.3 (June 22 intraday low). "
    "Using $76.54 as live reference. MCX crude proxy uses WTI ONLY (not Brent) per methodology.",
    italic=True, size=9, color=RED)

add_kv_table(doc,
    header=["Metric", "Reading", "Signal"],
    rows=[
        ["WTI Crude (NYMEX)", "$76.54/bbl  |  Range: $74.92–$78.08", "⬇ Falling — Iran de-escalation narrative"],
        ["Brent Crude (context only)", "$78.12/bbl", "For reference — NOT used for MCX proxy"],
        ["Brent–WTI Spread", "$1.58", "NORMAL (crisis spread = $7–10; not in crisis)"],
        ["MCX Crude Proxy (WTI×USDINR)", "₹76.54 × 94.325 = ₹7,220/bbl", "✅ AGGRESSIVE BULL regime (<₹7,500)"],
        ["Regime Boundary Risk", "₹7,500 is ₹280 away", "Watch — approaching MILD_BULL boundary"],
        ["USD/INR Spot", "94.325 (was 94.61 on June 22)", "Rupee marginally stronger — near record lows"],
        ["India VIX (June 22 close)", "27.32  (+3.06%)", "🔴 CRITICAL — >20 = QUARTER sizing MANDATORY"],
        ["VIX Driver", "Pre-expiry put buying (weekly exp. TODAY)", "Expiry-driven spike; watch for cooldown post-12PM"],
        ["Nifty 50 (June 22 close)", "24,090  (+0.37%)", "Positive session; closed below key resistance"],
        ["Nifty 50 prev. close", "~23,997 (estimated)", "Session low not confirmed — VERIFY at open"],
        ["Sensex (June 22 close)", "~64,320 (estimated Nifty×2.67)", "VERIFY exact figure on BSE"],
        ["GIFT Nifty (pre-mkt)", "24,154.5  (+64 pts vs 24,090)", "⚠ Gap up directly INTO resistance at 24,146"],
        ["Implied open gap", "+64 points (~+0.27%)", "Gap up hits R1 immediately — fade risk HIGH"],
        ["Nifty Max Pain (weekly exp.)", "24,050  |  Band: 23,950–24,200", "MM pin target — limited directional conviction"],
        ["Call OI wall", "24,200 CE", "Ceiling — breakout above = new bull leg"],
        ["Put OI floor", "24,000 PE", "Floor — major support on expiry"],
        ["Asia: Nikkei 225", "69,404  (+0.13%)", "Mild positive"],
        ["Asia: KOSPI (Seoul)", "8,726.60  (+2.11%)", "🟢 Strong — tech/chip rally (NOT India IT though)"],
        ["Asia: Hang Seng", "-1.40%", "🔴 Weak — China data drag on property/consumer"],
        ["Asia: Shanghai Composite", "~4,180 (testing resistance)", "Neutral — watch for break"],
        ["US: S&P 500 (June 22)", "7,482  (-0.24%)", "Mild red; off highs"],
        ["US: Nasdaq (June 22)", "26,166.60  (-1.32%)", "🔴 Tech selling: Alphabet -10%, AI capex fear"],
        ["US: Dow Jones (June 22)", "+0.29%  (+148 pts)", "Rotation to value/industrials"],
        ["US: Russell 2000", "3,004.40  (NEW ALL-TIME HIGH 🎯)", "First time above 3,000 — EM flow positive signal"],
        ["FII Flow (June 22, provisional)", "NET BUYERS ₹4,859.10 Cr", "🟢 Very bullish — strong foreign buying"],
        ["DII Flow (June 22, provisional)", "Net sellers ₹-1,159.60 Cr", "DII selling into FII buying (not alarming)"],
        ["Net FII+DII", "+₹3,699.50 Cr net inflow", "Positive market floor support"],
        ["RBI MPC (next meeting)", "~August 2026 (bi-monthly schedule)", "No imminent policy surprise expected"],
        ["India–US Trade Deal", "Feb 2026 signed — US tariff 25%→18%", "Structural positive; largely priced in"],
        ["Monsoon 2026", "Below normal (90–95% of LPA) — El Niño", "⚠ Food inflation risk building through H2 FY27"],
        ["Monsoon status June 23", "Advancing into Maharashtra/Mumbai area", "Progress on track — rural relief delayed"],
        ["Weekly F&O Expiry", "TODAY — June 23 (Tuesday)",  "🚨 NO weekly options. Use JULY monthly ONLY"],
        ["Monthly F&O Expiry", "June 25 (Thursday, 2 days away)", "Avoid JUNE monthly for NEW positions too"],
        ["F&O Ban List", "Sun TV, Indiabulls HF, RBL Bank", "Do NOT trade these today"],
    ])

# ────────────────── SECTION 2 — CRUDE REGIME ──────────────────
doc.add_paragraph()
add_heading(doc, "Section 2 — Crude Regime Analysis", level=1, color=NAVY)

add_para(doc, "REGIME: AGGRESSIVE BULL  (MCX proxy ₹7,220 < ₹7,500 threshold)", bold=True, size=13, color=GREEN)

add_bullets(doc, [
    "WTI $76.54 × USDINR 94.325 = ₹7,220/bbl → AGGRESSIVE_BULL regime",
    "Boundary watch: ₹7,500 is just ₹280 above current — NOT flagged (flag threshold is ₹200)",
    "Direction: FALLING (Iran deal progress, Goldman cuts Brent Q4 to $80 from $90)",
    "Brent–WTI spread $1.58 = NORMAL (confirms NOT in geopolitical crisis pricing)",
    "Regime implication: Broad CALL setups permitted. No restrictions except banned stocks.",
    "CAUTION: If WTI spikes above $79 on Iran headline reversal → re-check MCX proxy immediately",
    "AVOID: ONGC, OIL India long CALLs today — upstream crude producers suffer in low-crude regime",
    "BENEFIT: OMCs (BPCL, HPCL, IOC), Aviation (IndiGo), Paints, Tyres, Chemicals",
])

add_para(doc,
    "CRITICAL IRAN RULE: US-Iran road map announced June 21. This is the SECOND MoU in 2026. "
    "First MoU collapsed within 48h (Axios, June 14 report confirms prior collapse). "
    "Strait of Hormuz: CONTRADICTORY signals — Iran says closed; US CENTCOM says shipping proceeding. "
    "DO NOT treat as macro shift. Requires 3–5 session confirmation.",
    bold=True, size=11, color=RED)

# ────────────────── SECTION 3 — NEWS CATALYSTS ──────────────────
doc.add_paragraph()
add_heading(doc, "Section 3 — News Catalyst Aggregation", level=1, color=NAVY)

add_para(doc, "A. GEOPOLITICAL (Highest Priority)", bold=True, size=12, color=RED)
add_bullets(doc, [
    "Iran-US: June 21 — 'Road map' for final deal agreed; IAEA inspectors invited back; JD Vance calls it 'major milestone'",
    "Strait of Hormuz: Iran CLAIMS closure (citing Israeli strikes on Hezbollah). US CENTCOM: shipping proceeding normally. STATUS = AMBIGUOUS",
    "Israel-Hezbollah: Heavy fire exchanged Saturday, June 21. Active conflict threatening Iran deal.",
    "WTI fell to $74.3/bbl intraday Mon (June 22) on de-escalation. Now recovered to $76.54.",
    "RULE APPLIED: 3–5 session confirmation required. Today = Session 1. Treat as UNCONFIRMED.",
    "Scenario monitoring: Single headline can move WTI ±$5 in minutes. Check every 30 min.",
])

add_para(doc, "B. DOMESTIC INDIA", bold=True, size=12, color=NAVY)
add_bullets(doc, [
    "RBI MPC: No imminent meeting (~August 2026). Rate path neutral. Below-normal monsoon may keep RBI hawkish.",
    "India–US Trade Deal (Feb 2026): US tariff on India reduced 25%→18%. India eliminated tariffs on US farm/industrial goods. Structural positive.",
    "IMD Monsoon 2026: BELOW NORMAL forecast (90–95% LPA). El Niño developing. Advancing into Mumbai/Maharashtra as of June 22–23.",
    "Below-normal monsoon → food inflation risk → RBI less room to cut → rate-sensitive sectors capped.",
    "No major domestic regulatory announcement noted for today.",
])

add_para(doc, "C. CORPORATE CATALYSTS", bold=True, size=12, color=NAVY)
add_bullets(doc, [
    "Q4 FY26 results season largely concluded (April–May period). No marquee results noted for June 23.",
    "HAL showed relative strength: +2.44% on June 22, 2026 (outperformed broader market).",
    "BPCL: Fell from ₹316.30→₹306.60 on June 20 (post-crude fall; tracking WTI). Watch for stabilisation.",
    "No major IPO listings or block deals noted in today's search results.",
    "Goldman Sachs cuts Brent Q4 2026 forecast to $80 (from $90) — sustained OMC/aviation tailwind.",
])

add_para(doc, "D. F&O SPECIFIC — CRITICAL FLAGS", bold=True, size=12, color=RED)
add_bullets(doc, [
    "🚨 TODAY IS NIFTY WEEKLY EXPIRY (Tuesday June 23). ABSOLUTE PROHIBITION: No weekly options.",
    "🚨 June MONTHLY expiry = June 25 (2 days). Avoid June monthly for new positions. Use JULY monthly.",
    "Max pain: 24,050. OI band: 23,950–24,200. Expect Nifty to pin near 24,050–24,100 into close.",
    "VIX at 27.32 = options pricing in large moves; ATM premiums elevated — buy only on strong setups.",
    "Put/Call OI: Put wall 24,000; Call wall 24,200. Smart money range: 23,950–24,200.",
    "Expiry gamma: First 30–60 min will see major oscillations as MM defend/attack strikes.",
    "F&O Ban: Sun TV Network, Indiabulls Housing Finance, RBL Bank — DO NOT TRADE.",
])

add_para(doc, "E. GLOBAL MACRO", bold=True, size=12, color=NAVY)
add_bullets(doc, [
    "Russell 2000 at NEW ALL-TIME HIGH 3,004 → positive risk-on sentiment; EM fund inflows supported.",
    "Nasdaq -1.32%: Alphabet -10% on AI capex concerns; Palantir/Amazon/Meta -4%. Software IT drag.",
    "KOSPI +2.11%: South Korean chips/semis rallying. Different from India IT — don't conflate.",
    "China (Hang Seng -1.40%): Weak economic data dragging property, consumer stocks. Watch for contagion.",
    "Goldman Sachs: Brent Q4 2026 target cut to $80/bbl. Confirms sustained crude softness theme.",
    "US Dow +0.29%: Value/industrial rotation — positive for India defensives, PSU banks.",
])

# ────────────────── SECTION 4 — LAYER ANALYSIS ──────────────────
doc.add_paragraph()
add_heading(doc, "Section 4 — Layer 1/2/3 Bias Framework", level=1, color=NAVY)

add_para(doc, "LAYER 1 — Surface Consensus (what retail Twitter is saying right now):", bold=True, size=11, color=NAVY)
add_para(doc,
    '"Iran deal confirmed! WTI crashing. Markets gap up 64 points. Buy everything — aviation, OMCs, hospitality, '
    'retail. FII bought ₹4,859 Cr yesterday. KOSPI up 2.1%. Nifty to 24,500 by month-end. All clear to buy the open."',
    italic=True, size=11)

add_para(doc, "LAYER 2 — What Consensus Is LIKELY WRONG About:", bold=True, size=11, color=RED)
add_bullets(doc, [
    "Iran 'road map' is NOT a final deal. Previous MoU collapsed within 48h (twice in 2026). Strait status AMBIGUOUS — Iran says closed, US says open.",
    "VIX at 27.32 tells a VERY different story — market is deeply hedged and scared, NOT euphorically bullish.",
    "Today is WEEKLY EXPIRY: First 30–60 min will be a gamma war. MM will pin between 23,950–24,200.",
    "GIFT Nifty gap up to 24,154 IMMEDIATELY hits resistance at 24,146. High fade risk in first 15 min.",
    "Nasdaq -1.32% (Alphabet -10%) means India IT sector will see headwinds. KOSPI rally = hardware semis, not software IT.",
    "FII buying ₹4,859 Cr happened BEFORE road map announcement — much of the Iran optimism is already priced.",
    "DII sold ₹1,159 Cr into FII buying. Domestic smart money was DISTRIBUTING — a yellow flag.",
    "Max pain at 24,050 means MMs are incentivised to PREVENT a strong breakout. Expected range: tight.",
])

add_para(doc, "LAYER 3 — Underestimated / Tail Risks (what everyone is missing):", bold=True, size=11, color=RGBColor(0x80,0x40,0x00))
add_bullets(doc, [
    "Below-normal monsoon 2026 (El Niño) → food inflation spiral → RBI cannot cut aggressively → rate-sensitive sectors capped long-term.",
    "Goldman Sachs $80 Brent Q4 = if already priced into OMC stocks (BPCL down 3% June 20), the easy money is gone.",
    "Russell 2000 at ATH 3,004 = US small-cap new high often signals late-cycle economic turn — watch for EM rotation.",
    "Iran deal collapse risk: With Israel-Hezbollah actively fighting, ONE airstrike on Iran flips WTI to $85+ immediately.",
    "Monsoon arrival in Mumbai (within 24–48h per IMD) could be a near-term positive catalyst for sentiment.",
    "KOSPI +2.11% (chip rally) could attract fund flows to India technology hardware, not legacy IT services.",
])

add_kv_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["Flip to BULLISH", "Iran+Israel ceasefire confirmed by international body + IAEA inspectors enter + Strait declared open → WTI below $72 → Nifty 24,500+"],
        ["Flip to BEARISH", "Israeli strike on Iran OR Iran formally closes Strait per US CENTCOM → WTI spikes to $85+ → MCX crude ₹8,000+ → CAUTION regime → Nifty 23,200–23,500"],
        ["Entry Permission", "YELLOW — Wait 45–60 min. Confirm 24,073 holds on first 15-min close. No Iran escalation in first 45 min."],
        ["System Expectation", "PAPER system may not fire on first 30 min. That is CORRECT behaviour for YELLOW permission day."],
    ])

# ────────────────── SECTION 5 — NIFTY STRUCTURE ──────────────────
doc.add_paragraph()
add_heading(doc, "Section 5 — Nifty Technical Structure", level=1, color=NAVY)

add_kv_table(doc,
    header=["Level", "Value", "Meaning"],
    rows=[
        ["June 22 Close", "24,090", "Yesterday's close — baseline"],
        ["GIFT Nifty implied open", "~24,150–24,160", "Gap up ~64 pts; moderate"],
        ["R4 — Major breakout", "24,424", "Above here = strong bull trend resumes"],
        ["R3", "24,302", "Secondary resistance"],
        ["R2 — Expiry ceiling", "24,207", "Call OI wall; hardest ceiling today"],
        ["R1 — First resistance 🚨", "24,146", "GAP UP OPENS INTO THIS. Watch for immediate rejection"],
        ["S1 — Immediate support", "24,073", "Must hold on 15-min close for YELLOW to remain valid"],
        ["S2", "24,014", "Next support if S1 breaks"],
        ["S3 — Max pain / psychological", "24,000", "Strong floor. Put OI wall. Expiry pin zone."],
        ["S4 — Expiry band lower", "23,950", "Lower bound of today's expected range"],
        ["200 DMA (approx.)", "~23,100–23,300", "Far below — long-term trend intact"],
        ["Expected range today", "23,950–24,250", "Expiry pin + VIX. Wider than normal."],
        ["THE ONE NUMBER", "24,146", "Break + 15-min close above = bullish confirmation. Rejection = pin day."],
    ])

add_para(doc,
    "KEY INSIGHT: The gap up at open (~24,155) lands EXACTLY on R1 at 24,146. This is the single most "
    "important observation of the day. IF Nifty holds above 24,146 on 15-min close (10:00–10:15 IST), "
    "the day turns constructively bullish. IF it is rejected, expect a pin day between 24,000–24,146.",
    bold=True, size=11, color=RED)

# ────────────────── SECTION 6 — STOCK SETUPS ──────────────────
doc.add_paragraph()
add_heading(doc, "Section 6 — Stock Setups (JULY Monthly Options Only)", level=1, color=NAVY)

add_para(doc,
    "ALL OPTIONS: JULY 2026 monthly expiry ONLY. Weekly options PROHIBITED (expiry day). "
    "June monthly expires June 25 (2 days) — avoid for new positions. "
    "VIX at 27.32 = elevated premiums. Size = QUARTER (max). "
    "Entry ONLY after 10:00 AM IST on all setups. Confirm Iran headlines stable first.",
    bold=True, size=10, color=RED)

doc.add_paragraph()
add_para(doc, "SETUP #1 — HAL  |  A+  |  CALL  |  Defense/Aerospace  [MACRO HEDGE]",
    bold=True, size=13, color=NAVY)
stock_card(doc, [
    ("NSE Symbol", "HAL"),
    ("CMP (June 22)", "₹4,515 (+2.44% yesterday — strong relative strength)"),
    ("Lot Size", "~75 shares (VERIFY on NSE before trading; revised periodically)"),
    ("Capital Filter", "4,515 × 75 × 3% = ₹10,159 < ₹15,000 ✅ (passes half-size filter)"),
    ("Sector", "Defense / Aerospace"),
    ("Grade", "A+  — Multiple independent catalysts + technical setup + macro alignment"),
    ("Why A+", "1) Macro hedge — profits even if Iran deal COLLAPSES (defense spend floor permanent). "
               "2) Relative strength: +2.44% on June 22 vs mixed market. "
               "3) HAL order book >₹1L Cr — secular earnings visibility. "
               "4) Geopolitical uncertainty = permanent defense spend regardless of Iran outcome."),
    ("Option", "HAL JULY 25 4600 CE  (OTM, reduces premium outlay vs ATM)"),
    ("Est. Premium", "₹120–140 per lot (at VIX 27)"),
    ("Entry Zone", "₹4,510–4,560 (wait for first 15-min close above ₹4,490)"),
    ("Entry Trigger (4 conditions)", "1. Nifty holds above 24,073 on first 15-min close\n"
                                     "2. HAL holds ≥ ₹4,480 on first 15-min close\n"
                                     "3. No Iran escalation headline in first 45 min (check news)\n"
                                     "4. Entry only AFTER 10:00 AM IST"),
    ("T1 (40% book)", "₹4,680 → option premium target ~₹180–200"),
    ("T2 (40% book)", "₹4,800 → option premium target ~₹260–290"),
    ("SL", "₹4,310 on 15-min candle close → option exit at ~₹60–65"),
    ("R:R (premium basis)", "Avg exit ₹220 vs entry ₹130 vs SL ₹62 → gain ₹90, risk ₹68 → ~1.3:1\n"
                            "NOTE: Below 2:1 due to high VIX. Accepted for HEDGE position at min lot size."),
    ("Time Stop", "Exit 50% if no move to ₹4,600 by 12:30 PM"),
    ("Capital at risk (quarter)", "1 lot × ₹130 premium = ₹9,750 per lot. Max 1 lot at quarter size."),
    ("Direction Bias", "CALL — defense secular bull, Iran-reversal hedge"),
])

doc.add_paragraph()
add_para(doc, "SETUP #2 — BPCL  |  A  |  CALL  |  Energy / OMC",
    bold=True, size=13, color=NAVY)
stock_card(doc, [
    ("NSE Symbol", "BPCL"),
    ("CMP (latest)", "~₹306–310 (June 20 data: ₹306.60 from ₹316.30 prior). VERIFY at open."),
    ("Lot Size", "CRITICAL: Verify current NSE lot. If ~1,800 → capital filter borderline. "
                 "At QUARTER size: ₹8,370 effective — manageable. Use 1 lot max."),
    ("Sector", "Energy — Downstream OMC (Oil Marketing Company)"),
    ("Grade", "A — Strong catalyst + crude regime alignment"),
    ("Why A", "1) MCX crude ₹7,220 = AGGRESSIVE_BULL → marketing margin expansion for OMCs. "
              "2) Goldman Sachs $80 Brent Q4 forecast = sustained crude softness. "
              "3) BPCL already fell ₹316→₹306 tracking crude; further crude fall = re-rate. "
              "4) In low-crude regime, OMCs benefit most — no under-recovery risk."),
    ("Option", "BPCL JULY 25 315 CE or 320 CE (OTM — lower delta, manageable premium at high VIX)"),
    ("Est. Premium", "₹8–12 per unit (high leverage, low per-unit cost)"),
    ("Entry Zone", "₹307–315"),
    ("Entry Trigger (4 conditions)", "1. WTI confirmed sub-$77 (no crude spike) at time of entry\n"
                                     "2. BPCL holds above ₹305 on 15-min close\n"
                                     "3. No Iran escalation headline\n"
                                     "4. Enter after 10:00 AM IST"),
    ("T1 (40% book)", "₹328 on underlying (+5.8% from ₹310)"),
    ("T2 (40% book)", "₹342 on underlying (+10.3% from ₹310)"),
    ("SL", "₹294 on 15-min close (-5.2% from ₹310)"),
    ("R:R (underlying)", "(328-310)/(310-294) = 18/16 = 1.1:1 underlying\n"
                         "Option R:R improves significantly with leverage at OTM strike."),
    ("Time Stop", "If BPCL doesn't trade above ₹318 by 12:00 PM, exit 50%"),
    ("Risk", "HIGH if Iran deal collapses → crude spikes → BPCL marketing margins compress again"),
    ("Direction Bias", "CALL — crude beneficiary in AGGRESSIVE_BULL regime"),
])

doc.add_paragraph()
add_para(doc, "SETUP #3 — INDIGO (InterGlobe Aviation)  |  A  |  CALL  |  Aviation  [CONDITIONAL]",
    bold=True, size=13, color=NAVY)
stock_card(doc, [
    ("NSE Symbol", "INDIGO"),
    ("CMP (June 22)", "₹5,004–5,021 (confirmed)"),
    ("Lot Size", "🚨 VERIFY FIRST: If lot >100 shares, this setup likely FAILS capital filter at quarter size. "
                 "Check NSE lot size before entry. If lot=100: 5010×100×3% = ₹15,030 (borderline). "
                 "If lot=50: ₹7,515 ✅. SKIP this trade if lot makes it unaffordable at quarter size."),
    ("Sector", "Aviation — direct crude/ATF beneficiary"),
    ("Grade", "A — CONDITIONAL on lot size check"),
    ("Why A", "1) MCX crude at ₹7,220 = direct ATF cost reduction. "
              "2) Iran de-escalation progress = sustained crude softness narrative. "
              "3) Peak June–July summer travel demand. "
              "4) India aviation sector structural growth (middle class expansion)."),
    ("Option", "INDIGO JULY 25 5100 CE (OTM — lower premium, capital-efficient at high VIX)"),
    ("Est. Premium", "₹85–110 per unit (OTM with VIX 27)"),
    ("Entry Zone", "₹5,050–5,100 (after 15-min gap stability confirmed)"),
    ("Entry Trigger (4 conditions)", "1. Lot size check PASSES at open (verify NSE)\n"
                                     "2. WTI confirmed sub-$77 (crude not spiking)\n"
                                     "3. INDIGO holds ≥ ₹4,980 on first 15-min close\n"
                                     "4. Enter after 10:00 AM IST"),
    ("T1 (40% book)", "₹5,280 (+5.4% from CMP)"),
    ("T2 (40% book)", "₹5,450 (+8.8% from CMP)"),
    ("SL", "₹4,820 on 15-min close (-3.8% from CMP)"),
    ("KEY RISK", "🔴 Highest Iran reversal risk of all setups. If Iran deal collapses → crude spikes $10 → "
                 "IndiGo gets crushed instantly. This is an AGGRESSIVE setup. Size minimum."),
    ("Time Stop", "Exit if INDIGO doesn't close above ₹5,100 by 11:30 AM IST"),
    ("Direction Bias", "CALL — aviation crude beneficiary"),
])

doc.add_paragraph()
add_para(doc, "SETUP #4 — HDFC BANK  |  A  |  CALL  |  Banking / BFSI  [CONDITIONAL ON LOT SIZE]",
    bold=True, size=13, color=NAVY)
stock_card(doc, [
    ("NSE Symbol", "HDFCBANK"),
    ("CMP (estimate)", "~₹1,840–1,900 (VERIFY at open; Bank Nifty support 57,000–57,100)"),
    ("Lot Size", "🚨 VERIFY: Historical lot ~550 which is too large. If revised to ≤250: "
                 "1880×250×3% = ₹14,100 ✅. If still 550: ₹31,020 ❌ FAIL. "
                 "If HDFC Bank fails: substitute ICICI Bank (smaller effective lot) or Bank Nifty ETF options."),
    ("Sector", "Banking / BFSI"),
    ("Grade", "A — FII flow catalyst + sector momentum + rate outlook"),
    ("Why A", "1) FII net buyers ₹4,859 Cr (June 22) — HDFC Bank typically receives 20–25% of FII large-cap allocation. "
              "2) Bank Nifty 57,000–57,100 support floor holding. "
              "3) Rate cut expectations still intact (no hawkish surprise). "
              "4) India–US trade deal supportive of corporate credit growth."),
    ("Option", "HDFCBANK JULY 25 1900 CE or 1920 CE (OTM)"),
    ("Est. Premium", "₹20–30 per unit (OTM, high leverage)"),
    ("Entry Zone", "₹1,845–1,880"),
    ("Entry Trigger (4 conditions)", "1. Bank Nifty above 57,200 on first 15-min close\n"
                                     "2. HDFC Bank above ₹1,835 on first 15-min candle\n"
                                     "3. FII buying continues (check NSE provisional at 10:30 AM)\n"
                                     "4. Enter after 10:00 AM IST"),
    ("T1 (40% book)", "₹1,940 on underlying"),
    ("T2 (40% book)", "₹1,995 on underlying"),
    ("SL", "₹1,790 on 15-min close"),
    ("R:R (option premium)", "Entry ₹25, SL ₹12, T1 ₹55, T2 ₹82 → avg exit ~₹65, gain ₹40, risk ₹13 → 3.1:1 ✅"),
    ("Time Stop", "Exit if Bank Nifty fails to hold 57,100 for 2 consecutive 15-min candles"),
    ("Direction Bias", "CALL — financial sector FII inflow proxy"),
])

doc.add_paragraph()
add_para(doc, "SETUP #5 — TCS  |  B+  |  PUT  |  IT Services  [CONDITIONAL — SKIP IF GAP UP]",
    bold=True, size=13, color=NAVY)
stock_card(doc, [
    ("NSE Symbol", "TCS"),
    ("CMP (estimate)", "~₹3,800–4,000 (VERIFY at open)"),
    ("Lot Size", "VERIFY: If ~175 shares: 3900×175×3% = ₹20,475 ❌. If revised to ≤125: ₹14,625 ✅. Check NSE."),
    ("Sector", "IT Services"),
    ("Grade", "B+ — Single strong catalyst + conditional technical setup"),
    ("Why PUT", "1) Nasdaq -1.32% (June 22): Alphabet -10% on AI capex fears — spillover to India IT. "
                "2) KOSPI tech rally is SEMICONDUCTOR hardware, not software services — different thesis. "
                "3) Below-normal monsoon → rural demand slowdown (negligible for IT but sector sentiment weak). "
                "4) AI disruption narrative (Alphabet capex) creates uncertainty for IT services growth."),
    ("CRITICAL SKIP CONDITION", "🚨 SKIP this trade if TCS opens +0.5% or more at open. "
                                 "This is a CONDITIONAL PUT — only valid if TCS confirms weakness."),
    ("Option", "TCS JULY 25 3700 PE (OTM PUT)"),
    ("Est. Premium", "₹55–75 per unit (OTM PUT at high VIX)"),
    ("Entry Zone", "₹3,800–3,850 (enter on TCS rejection; NOT at open gap)"),
    ("Entry Trigger (4 conditions)", "1. TCS opens RED or flat (≤ +0.3%)\n"
                                     "2. Nifty 24,207 level REJECTED on 15-min close\n"
                                     "3. TCS trades below 5-day EMA at entry time\n"
                                     "4. Enter AFTER 10:30 AM IST (more caution for counter-market PUT)"),
    ("T1 (40% book)", "₹3,680 on underlying"),
    ("T2 (40% book)", "₹3,570 on underlying"),
    ("SL", "₹3,980 on 15-min candle close above"),
    ("R:R (option premium)", "Entry ₹65, SL ₹30, T1 ₹115, T2 ₹165 → avg exit ~₹135, gain ₹70, risk ₹35 → 2.0:1 ✅"),
    ("Time Stop", "Exit 100% by 2:00 PM (no IT puts into close on expiry day)"),
    ("Contrarian risk", "India–US trade deal is positive for IT (lower barriers). KOSPI rally shows tech not dead. "
                        "This is B+ grade for a reason — enter only on confirmation."),
    ("Direction Bias", "PUT — Nasdaq drag, AI capex fear, conditional"),
])

# ────────────────── SECTION 7 — DIVERSIFICATION ──────────────────
doc.add_paragraph()
add_heading(doc, "Section 7 — Portfolio Diversification Check", level=1, color=NAVY)
add_kv_table(doc,
    header=["#", "Stock", "Grade", "Direction", "Sector", "Macro Alignment"],
    rows=[
        ["1", "HAL", "A+", "CALL", "Defense", "Iran-proof hedge + secular defense bull"],
        ["2", "BPCL", "A", "CALL", "Energy/OMC", "Crude AGGRESSIVE_BULL beneficiary"],
        ["3", "INDIGO", "A (Conditional)", "CALL", "Aviation", "Crude/ATF cost beneficiary"],
        ["4", "HDFC Bank", "A (Conditional)", "CALL", "Banking/BFSI", "FII inflow proxy"],
        ["5", "TCS", "B+", "PUT", "IT Services", "Nasdaq drag / conditional"],
    ])

add_bullets(doc, [
    "Sectors: 5 different (Defense, OMC, Aviation, Banking, IT) ✅ Minimum 4 required ✅",
    "Max 2 per sector: All single stocks in each sector ✅",
    "CALL:PUT mix: 4 CALLs + 1 conditional PUT ✅ (not all directional)",
    "Macro hedge: HAL (defense) ✅ Required 1 macro hedge ✅",
    "Total stocks: 5 (< 7 maximum) ✅",
    "F&O ban check: HAL, BPCL, INDIGO, HDFCBANK, TCS — NONE are in today's ban list ✅",
    "Expiry check: ALL JULY monthly — weekly prohibited (expiry day) ✅",
    "LOT SIZE WARNING: VERIFY HAL, HDFC Bank, INDIGO, TCS lot sizes at market open. Two may fail capital filter.",
])

# ────────────────── SECTION 8 — STOCKS TO AVOID ──────────────────
doc.add_paragraph()
add_heading(doc, "Section 8 — Stocks to ACTIVELY AVOID Today", level=1, color=RED)
add_kv_table(doc,
    header=["Stock", "Reason for Avoidance"],
    rows=[
        ["Sun TV Network", "F&O BAN LIST — trading prohibited (MWPL exceeded)"],
        ["Indiabulls Housing Finance", "F&O BAN LIST — trading prohibited"],
        ["RBL Bank", "F&O BAN LIST — trading prohibited"],
        ["ONGC / OIL India / MRPL", "Upstream crude producers — suffer in AGGRESSIVE_BULL (low crude) regime. Wrong sector today."],
        ["ANY weekly options", "ABSOLUTE PROHIBITION — today is weekly expiry. Use JULY monthly ONLY."],
        ["JUNE monthly options (new)", "Expires June 25 (2 days). Too close for new positions."],
        ["Pure Nifty directional bet", "Max pain 24,050 + VIX 27.32 = expiry pinning day. Range-bound. No conviction breakout expected."],
        ["IT stocks at open (blind)", "Wait for TCS/INFY to confirm direction first. KOSPI vs Nasdaq signals are OPPOSITE. Resolve before entry."],
        ["HPCL / IOC", "Same OMC thesis as BPCL but BPCL is primary — avoid stacking same sector calls."],
        ["Any stock with P/E < 0 or recent earnings miss", "INFY/Wipro class: avoid pre-result or post-miss stocks without catalyst."],
        ["Pharma / FMCG broad sector bets", "Below-normal monsoon headwind for FMCG rural. Pharma has no catalyst today."],
    ])

# ────────────────── SECTION 9 — FINAL VERDICT ──────────────────
doc.add_paragraph()
add_heading(doc, "Section 9 — Final Verdict", level=1, color=RED)

badge_line(doc, "NIFTY BIAS", "MILDLY BULLISH — but range-bound (expiry pin)", AMBER, text_color=RGBColor(0,0,0))
badge_line(doc, "CONFIDENCE", "50–55% (HIGH VIX + expiry + Iran unconfirmed)", ORANGE)
badge_line(doc, "CRUDE MODE", "AGGRESSIVE BULL  |  WTI $76.54  |  MCX ₹7,220", DKGRN)
badge_line(doc, "VIX SIZING", "QUARTER SIZE (VIX 27.32 > 20 threshold)", DKRED)
badge_line(doc, "ENTRY PERMISSION", "YELLOW — wait 45-60 min; confirm 24,073 holds", ORANGE)
badge_line(doc, "EXPECTED NIFTY RANGE", "23,950 – 24,250  (expiry pin day)", BLUE)
badge_line(doc, "THE ONE NUMBER", "24,146 — break + hold = bull day; rejection = pin day", BLUE)

doc.add_paragraph()
add_kv_table(doc,
    header=["Scenario", "Probability", "Triggers", "Nifty Target"],
    rows=[
        ["BULL CASE", "40%", "24,146 breaks + holds on 15-min. No Iran escalation. FII buying continues. BPCL/IndiGo lead.", "24,300–24,500"],
        ["BASE CASE (Most Likely)", "40%", "Nifty opens at 24,155 → rejected at 24,146 → oscillates 24,000–24,200 → expiry pin near 24,050.", "23,950–24,200 range"],
        ["BEAR CASE", "20%", "Iran deal collapses headline → crude spikes $5+ → WTI above $80 → MCX crude ₹7,500+ → regime flip → Nifty breaks 24,000.", "23,500–23,800"],
    ])

add_para(doc,
    "MOST LIKELY SCENARIO: Expiry pin day. Nifty oscillates between 23,950–24,200. "
    "VIX comes in after 12:00 PM as weekly options die. Post-12 PM is better for directional entries on MONTHLY options. "
    "HAL and BPCL are best positioned for the 'Iran de-escalation + crude softness' thesis, with HAL as the hedge. "
    "TCS PUT activates only on confirmation of Nasdaq-led IT weakness.",
    bold=True, size=11)

doc.add_paragraph()
add_para(doc, "KEY EVENTS TO MONITOR (by time):", bold=True, size=11)
add_bullets(doc, [
    "09:15–09:45 IST: Open + first 15-min candle. Confirm 24,073 hold or break. Do NOT trade.",
    "09:45–10:00 IST: Second candle. Iran headlines check. WTI direction check.",
    "10:00–10:15 IST: YELLOW entry permission opens. Execute setups if triggers fire.",
    "10:30 IST: NSE provisional FII/DII data. If FII turns seller → revise confidence down.",
    "12:00–12:30 IST: VIX typically comes in after weekly options die. Better for directional trades.",
    "14:00 IST: Exit all B+ grade positions (TCS PUT) by this time regardless.",
    "15:00–15:30 IST: Expiry max pain pull. Expect Nifty to gravitate toward 24,050.",
    "15:30 IST: Market close. Log paper trading system signals for tomorrow's postmortem.",
])

# ────────────────── SECTION 10 — PAPER TRADING ──────────────────
doc.add_paragraph()
add_heading(doc, "Section 10 — Paper Trading System Notes", level=1, color=NAVY)
add_bullets(doc, [
    "System expected to see: Gap up open ~24,150–24,165. First 15-min candle highly volatile (expiry).",
    "VIX at 27.32 = system thresholds may require large confirmation candles. This is CORRECT.",
    "CRITICAL: System MUST NOT fire on weekly option strikes today. All signals should target JULY monthly.",
    "If system fires within first 30 min today: REVIEW trigger sensitivity. May be too loose for VIX 27 day.",
    "Running zero-trade day counter: N/A (first session). Reset counter at Session 1.",
    "Backtest watch: Did the system correctly identify AGGRESSIVE_BULL crude regime? Check crude input source.",
    "Log to track: Time of first signal, which stock fired, which option series used, entry premium paid.",
    "Specific watch for paper trade: Does system enter HAL on gap open or wait for 15-min confirmation?",
    "TOMORROW's postmortem will score: Did Nifty close above/below 24,146? Did crude stay below ₹7,500? Did Iran hold?",
])

# ────────────────── ONE-LINE SUMMARY ──────────────────
doc.add_paragraph()
add_heading(doc, "ONE-LINE SUMMARY", level=1, color=NAVY)
add_para(doc,
    "JUNE 23 — Nifty gaps to resistance (24,146) on Iran road-map + FII ₹4,859 Cr buying; VIX at 27.32 "
    "forces QUARTER sizing; today is weekly EXPIRY (JULY monthly only); crude AGGRESSIVE_BULL (₹7,220 MCX); "
    "YELLOW permission — wait 45-60 min; 5 setups: HAL A+ (hedge), BPCL A, IndiGo A (check lot), "
    "HDFC Bank A (check lot), TCS B+ PUT (conditional); expected range 23,950–24,250; "
    "THE ONE NUMBER: 24,146.",
    italic=True, bold=True, size=12)

# ────────────────── FOOTER ──────────────────
doc.add_paragraph()
foot = doc.add_paragraph()
foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = foot.add_run(
    "Generated by Daily Trading Routine — for personal educational use only. Not financial advice. "
    "All prices must be verified at market open. Lot sizes subject to NSE revision."
)
fr.italic = True
fr.font.size = Pt(9)
fr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

doc.save(OUT_PATH)
print(f"Wrote {OUT_PATH}")
