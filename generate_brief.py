"""Generate the daily trading-intelligence .docx brief."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-05-21"
BIAS = "CAUTIOUS-BULL"
CONFIDENCE = "MEDIUM"
ENTRY_PERMISSION = "YELLOW"

OUT_PATH = f"trading-briefs/2026/05-May/Trading_Brief_{DATE_ISO}_{BIAS}.docx"


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


def add_kv_table(doc, rows, col_widths=(1.4, 1.6, 2.2)):
    table = doc.add_table(rows=1 + len(rows), cols=3)
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text = "Metric", "Value", "Signal"
    for c in hdr:
        for p in c.paragraphs:
            for r in p.runs:
                r.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade_cell(c, "1F4E79")
    for i, (k, v, s) in enumerate(rows, 1):
        row = table.rows[i].cells
        row[0].text, row[1].text, row[2].text = k, v, s
    for i, w in enumerate(col_widths):
        for r in table.rows:
            r.cells[i].width = Inches(w)


def setup_stock_block(doc, data):
    add_para(doc, "=" * 60, size=9, color=RGBColor(0x70, 0x70, 0x70))
    title = f"{data['name']} (NSE: {data['ticker']}) | {data['side']} | Grade {data['grade']}"
    add_para(doc, title, bold=True, size=13, color=RGBColor(0x1F, 0x4E, 0x79))
    add_para(doc, f"CMP: Rs {data['cmp']} | Lot: {data['lot']} | Crude aligned: {data['crude_aligned']}", size=10)
    add_para(doc, f"CATALYST: {data['catalyst']}", italic=True, size=11)
    add_para(doc, "TECHNICAL:", bold=True, size=11)
    add_bullets(doc, [
        f"Daily trend: {data['daily']}",
        f"Above 200 DMA: {data['dma200']}",
        f"Weekly trend: {data['weekly']}",
        f"Support Rs {data['support']} | Resistance Rs {data['resistance']}",
    ])
    add_para(doc, "STRUCTURAL:", bold=True, size=11)
    add_bullets(doc, [
        f"OI direction: {data['oi']}",
        f"Delivery %: {data['delivery']}",
        f"Block deal: {data['block']} | Ban list: {data['ban']}",
    ])
    add_para(doc, "OPTION SETUP:", bold=True, size=11)
    add_bullets(doc, [
        f"Strike: Rs {data['strike']} [{data['moneyness']}] | Expiry: {data['expiry']}",
        f"THETA WARNING (<=3 DTE): {data['theta_warn']}",
    ])
    add_para(doc, "ENTRY (ALL 4 must fire):", bold=True, size=11)
    add_bullets(doc, [
        f"SuperTrend {data['st']} on 15-min confirmed",
        f"RSI {data['rsi']} 50 on 15-min",
        f"StochRSI cross {data['srsi']} from extreme",
        "Volume > 1.5x 20-period avg",
        f"Confirmation candle: 15-min close {data['conf_side']} Rs {data['conf_level']}",
    ])
    add_para(doc, f"T1 (book 40%): Rs {data['t1']}  |  T2 (book 40%): Rs {data['t2']}", bold=True, size=11)
    add_para(doc, f"SL: Rs {data['sl']} + {data['sl_cond']}", bold=True, size=11)
    add_para(doc, f"R:R: {data['rr']}  |  Time stop: Exit by 13:00 IST", size=11)
    add_para(doc, f"GRADE REASON: {data['reason']}", italic=True, size=10, color=RGBColor(0x70, 0x70, 0x70))


def main():
    doc = Document()

    # Title
    title = doc.add_heading("ELITE DAILY TRADING BRIEF", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run(f"{DATE_ISO}  |  Bias: {BIAS}  |  Confidence: {CONFIDENCE}")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    mr = meta.add_run("Generated: 03:15 IST  |  Weekly expiry: 5 days  |  Monthly expiry: 5 days  |  Expiry week: NO")
    mr.italic = True
    mr.font.size = Pt(10)

    # Headline box
    box = doc.add_table(rows=1, cols=1)
    box.style = "Light Shading Accent 1"
    cell = box.rows[0].cells[0]
    shade_cell(cell, "FFF2CC")
    cell.text = ""
    p1 = cell.paragraphs[0]
    p1.add_run("BIAS: CAUTIOUS BULL (Contrarian Override Active)").bold = True
    cell.add_paragraph("ENTRY PERMISSION: YELLOW - gap-up + crude crash. Wait for 15-min structure.")
    cell.add_paragraph("CRUDE MODE: MILD BULL  |  Direction: FALLING HARD (-5% O/N)")
    cell.add_paragraph("VIX SIZING: HALF (India VIX 19.42 - between 15-20)")
    cell.add_paragraph("EXPIRY ALERT: Weekly Tue 26-May (5 DTE). Not expiry week. Monthly = 5 DTE.")

    # ===== Section 1: Macro Snapshot =====
    add_heading(doc, "Section 1 - Macro Snapshot", level=1)
    add_kv_table(doc, [
        ("MCX Crude (proxy)", "~Rs 8,352/bbl", "MILD BULL (just fell from CAUTION)"),
        ("Brent / WTI", "$105.02 / $98.26", "Crashed -5%+ overnight"),
        ("GIFT Nifty gap", "+199 pts (23,858 vs 23,659)", "Bullish gap-up"),
        ("S&P 500", "+1.0% (May 20)", "Risk ON"),
        ("Nasdaq", "+1.2% (May 20)", "Risk ON"),
        ("US VIX", "Easing (calm post Iran)", "Calm"),
        ("India VIX", "19.42 (+0.75%)", "HALF SIZE (rule)"),
        ("Nifty close", "Rs 23,659.00 (+0.17%)", "Below 200 DMA zone"),
        ("FII cash (18-May)", "+Rs 2,813 Cr", "Buying (MTD -21,842)"),
        ("DII cash (18-May)", "+Rs 2,682 Cr", "Buying support"),
        ("USD-INR", "~Rs 85 (est.)", "Stable"),
        ("INFY ADR", "$12.72", "Flat - neutral gap"),
        ("HDB ADR", "$24.55", "Slight positive bias"),
    ])
    add_para(doc, "Geopolitical: Trump says US-Iran deal in 'final stages'. First supertanker crossed Hormuz overnight after 12-week blockade. Citi warned market underpricing tail risk - Brent could spike to $120 if talks collapse.", italic=True, size=10)
    add_para(doc, "Today's intraday headline risk: HIGH (binary on Iran headlines - both ways).", bold=True, color=RGBColor(0xC0, 0x00, 0x00), size=11)

    # ===== Section 2: Crude + Structural =====
    add_heading(doc, "Section 2 - Crude Rule & Structural Read", level=1)
    add_para(doc, "CRUDE MODE: MILD BULL  |  DIRECTION: FALLING HARD  |  FLIP LEVEL: Rs 7,500 (aggressive bull) / Rs 8,500 (back to caution)", bold=True)
    add_para(doc, "Avoid (now flipped to OPPORTUNITY): historically the avoid list - OMCs (BPCL/HPCL/IOC), aviation, paints, tyres, fertilisers - has become today's BUY list as crude collapses. Crude producers (ONGC, Oil India) flip to PUT candidates.", size=11)
    add_para(doc, "Exception calls: Defence names (BEL - results today) retain independent bid regardless of crude.", size=11)
    add_para(doc, "Structural read:", bold=True)
    add_bullets(doc, [
        "Nifty vs max pain: DATA_UNAVAILABLE for live PCR/max pain. Use 23,500 PE / 23,800 CE as estimated OI walls.",
        "Weekly expiry now Tuesday (NSE rule change) - 26-May. 5 DTE. Monthly = 29-May. Both 5 DTE; weekly + monthly OK.",
        "Futures basis: index futures typically tracking spot; Bank Nifty leadership will decide whether crude-crash gap holds.",
        "Expiry-week dynamics: NOT in effect (next expiry Tue). PCR squeeze risk LOW for today's session specifically.",
    ])
    add_para(doc, "Bank Nifty leadership check: HDB ADR firm + global risk-on = banks should lead the gap. If Bank Nifty lags Nifty after 9:30, gap-up is suspect and faders win.", bold=True)

    # ===== Section 3: Contrarian Check =====
    add_heading(doc, "Section 3 - Contrarian Check (Layer 3)", level=1)
    add_bullets(doc, [
        "Q1. CONSENSUS today: Every TV analyst will scream 'crude crash = BUY paints, aviation, OMCs'. Retail will chase BPCL/INDIGO/ASIANPAINT calls at the open.",
        "Q2. THE TRAP: Iran deal headline is single-source (Trump tweet). One contradictory headline from Tehran reverses crude in minutes. Long gap-ups in those sectors get knifed at 10:30-11:00 if oil bounces.",
        "Q3. RETAIL STOPS: Long stops cluster below 23,600 (yesterday's low ~23,610) and below today's gap fill 23,659. Short-side stops above 23,900 (round + GIFT high).",
        "Q4. FLIP TRIGGER: A single Iran statement rejecting US terms -> crude rips back $5+, Nifty fills gap to 23,659 and breaks. Also: USD-INR spike >Rs 85.50 reverses FII bid.",
    ])

    add_heading(doc, "Contrarian Override Logic - Applied", level=2)
    add_kv_table(doc, [
        ("A. GIFT gap-up > +100 pts?", "YES (+199)", "Override condition met"),
        ("B. Crude falling > 2% O/N?", "YES (-5%+)", "Override condition met"),
        ("C. Contrarian probability > 40%?", "YES (single-headline driver)", "Confidence drag"),
        ("D. Expiry week + PCR < 0.8?", "NO (5 DTE)", "Squeeze risk LOW"),
    ], col_widths=(2.3, 1.8, 1.6))
    add_para(doc, "VERDICT: A + B + C -> CAUTIOUS BULL, permission YELLOW. 'Consensus bearish-crude trade is the trap today - but only if Iran headline holds.'", bold=True, color=RGBColor(0xB0, 0x60, 0x00))

    # ===== Section 4: 5 setups =====
    add_heading(doc, "Section 4 - Five F&O Setups", level=1)

    stocks = [
        {
            "name": "BPCL", "ticker": "BPCL", "side": "CALL", "grade": "A",
            "cmp": "312.90", "lot": "1,800", "crude_aligned": "YES",
            "catalyst": "Q4 results today (BPCL on calendar). Crude crash -5% massively expands OMC marketing margins; auto-fuel pricing freeze becomes profitable.",
            "daily": "DOWN turning UP (-21% in 6M; bottom-fishing setup)", "dma200": "NO (below - recovery trade)",
            "weekly": "DOWN", "support": "308", "resistance": "320 / 325",
            "oi": "Watch for LONG BUILD on gap-up confirmation", "delivery": "RISING expected on results day",
            "block": "NO", "ban": "NO (not in ban list)",
            "strike": "320 CE", "moneyness": "OTM", "expiry": "26-May (weekly) - 5 DTE",
            "theta_warn": "NO (5 DTE > 3)",
            "st": "GREEN", "rsi": "above", "srsi": "up", "conf_side": "above", "conf_level": "316",
            "t1": "324", "t2": "330", "sl": "311", "sl_cond": "ST flip RED on 15-min",
            "rr": "2.2:1",
            "reason": "Direct crude-crash beneficiary + Q4 catalyst today + nearby resistance gives clean 2:1.",
        },
        {
            "name": "InterGlobe Aviation (IndiGo)", "ticker": "INDIGO", "side": "CALL", "grade": "A",
            "cmp": "4,230.30", "lot": "300", "crude_aligned": "YES",
            "catalyst": "ATF (jet fuel) is ~40% of opex. Crude -5% overnight = direct EBITDA tailwind. Brokerage downgrades last week (target cut to Rs 5,200) created oversold base.",
            "daily": "DOWN (-23% in 6M)", "dma200": "NO (recovery setup)",
            "weekly": "DOWN turning", "support": "4,180", "resistance": "4,310 / 4,400",
            "oi": "Short cover candidate if 4,250 breaks", "delivery": "STABLE",
            "block": "NO", "ban": "NO",
            "strike": "4,300 CE", "moneyness": "OTM", "expiry": "26-May (weekly)",
            "theta_warn": "NO",
            "st": "GREEN", "rsi": "above", "srsi": "up", "conf_side": "above", "conf_level": "4,260",
            "t1": "4,350", "t2": "4,420", "sl": "4,195", "sl_cond": "15-min close below 4,200",
            "rr": "2.4:1",
            "reason": "Cleanest crude-crash beneficiary, oversold daily, strong 4,180 support gives tight SL.",
        },
        {
            "name": "Asian Paints", "ticker": "ASIANPAINT", "side": "CALL", "grade": "A",
            "cmp": "2,514.40", "lot": "200", "crude_aligned": "YES",
            "catalyst": "Crude derivatives = ~30% of paint COGS. -5% crude flows straight to gross margin in 2-3 weeks. Stock at 52-wk low zone (Rs 2,115 low; Rs 2,985 high) - value buyers active.",
            "daily": "SIDEWAYS basing", "dma200": "NO (but holding well above 52w low)",
            "weekly": "BASING", "support": "2,490", "resistance": "2,545 / 2,580",
            "oi": "LONG BUILD on Wednesday's tape", "delivery": "RISING",
            "block": "NO", "ban": "NO",
            "strike": "2,540 CE", "moneyness": "OTM", "expiry": "29-May (monthly)",
            "theta_warn": "NO",
            "st": "GREEN", "rsi": "above", "srsi": "up", "conf_side": "above", "conf_level": "2,525",
            "t1": "2,560", "t2": "2,585", "sl": "2,498", "sl_cond": "15-min close below 2,500",
            "rr": "2.3:1",
            "reason": "Highest beta to crude move in the basket; tight technical structure; lot 200 = capital-efficient.",
        },
        {
            "name": "HDFC Bank", "ticker": "HDFCBANK", "side": "CALL", "grade": "B+",
            "cmp": "767.50", "lot": "550", "crude_aligned": "Neutral",
            "catalyst": "HDB ADR firm overnight + global risk-on. Bank Nifty leadership trade - if Nifty gap-up holds, HDFC Bank leads the move. No earnings, but largest weight.",
            "daily": "DOWN (-21% in 1Y - washed out)", "dma200": "NO (well below)",
            "weekly": "BASING near 52w low (726)", "support": "760", "resistance": "775 / 785",
            "oi": "Watch first 15-min for LONG BUILD on gap", "delivery": "STABLE",
            "block": "NO", "ban": "NO",
            "strike": "780 CE", "moneyness": "OTM", "expiry": "29-May (monthly)",
            "theta_warn": "NO",
            "st": "GREEN", "rsi": "above", "srsi": "up", "conf_side": "above", "conf_level": "770",
            "t1": "778", "t2": "786", "sl": "762", "sl_cond": "15-min close below 762",
            "rr": "2.1:1",
            "reason": "B+ not A - no stock-specific catalyst; rides Nifty/Bank Nifty leadership. Lot 550 = good capital efficiency.",
        },
        {
            "name": "ONGC", "ticker": "ONGC", "side": "PUT", "grade": "B",
            "cmp": "297.45", "lot": "1,500", "crude_aligned": "YES (inverse)",
            "catalyst": "Crude producer - every $1 fall = realisation hit. Crude -5% = direct top-line headwind. Stock had run +25% YTD; profit-taking expected.",
            "daily": "UP overextended (+25% 1Y)", "dma200": "YES (above - but profit-take)",
            "weekly": "TOPPY near 52w high (307)", "support": "292 / 288", "resistance": "302",
            "oi": "Watch for SHORT BUILD", "delivery": "FALLING expected",
            "block": "NO", "ban": "NO",
            "strike": "295 PE", "moneyness": "ATM/OTM", "expiry": "29-May (monthly)",
            "theta_warn": "NO",
            "st": "RED", "rsi": "below", "srsi": "down", "conf_side": "below", "conf_level": "295",
            "t1": "291", "t2": "287", "sl": "299", "sl_cond": "15-min close above 300 + ST flip GREEN",
            "rr": "2.0:1",
            "reason": "B grade - Iran-deal headline risk cuts both ways. If talks collapse, ONGC rips to upside. Use as hedge, smaller size.",
        },
    ]
    for s in stocks:
        setup_stock_block(doc, s)

    # ===== Section 5: Final Verdict =====
    add_heading(doc, "Section 5 - Final Verdict", level=1)
    verdict = doc.add_table(rows=5, cols=2)
    verdict.style = "Light Grid Accent 1"
    verdict_rows = [
        ("TODAY'S BIAS", "CAUTIOUS BULL (override applied)"),
        ("CONFIDENCE", "MEDIUM"),
        ("CRUDE MODE", "MILD BULL - falling hard"),
        ("VIX SIZING", "HALF (mandatory, VIX 19.42)"),
        ("ENTRY PERMISSION", "YELLOW - wait for 15-min structure"),
    ]
    for i, (k, v) in enumerate(verdict_rows):
        verdict.rows[i].cells[0].text = k
        verdict.rows[i].cells[1].text = v
        for p in verdict.rows[i].cells[0].paragraphs:
            for r in p.runs:
                r.bold = True

    add_para(doc, "THE BULL CASE:", bold=True)
    add_para(doc, "Crude crashed -5% on Iran deal - direct margin expansion for paints, OMCs, aviation. GIFT Nifty +199 pts, S&P +1%, HDB ADR firm = clean global risk-on tape. FII + DII both bought May 18. Nifty defending 23,600.", size=11)
    add_para(doc, "THE BEAR CASE:", bold=True)
    add_para(doc, "Iran headline is single-source (Trump). Any Tehran rebuttal sends crude +$5 in minutes - gap fill 23,659 -> break = 23,500/23,400. India VIX still 19.4 (elevated). Bank Nifty has been a chronic laggard. FII MTD net SOLD Rs 21,842 Cr.", size=11)

    add_para(doc, "FLIP TRIGGER:", bold=True, color=RGBColor(0xC0, 0x00, 0x00))
    add_para(doc, "If Iran rejects US terms OR crude futures rally back above $103 (WTI) intraday -> bias flips to BEAR, all CALL setups invalid, ONGC PUT wins.", size=11)

    add_para(doc, "NIFTY KEY LEVELS:", bold=True)
    add_para(doc, "S2: 23,450  |  S1: 23,560  |  CRITICAL: 23,659 (yesterday's close / gap fill)  |  R1: 23,860  |  R2: 23,980", size=11)

    add_para(doc, "TOP 3 RANKED:", bold=True)
    add_bullets(doc, [
        "#1  INDIGO  - Grade A - cleanest crude-crash beneficiary, oversold base - entry above Rs 4,260",
        "#2  BPCL    - Grade A - OMC margin expansion + Q4 results catalyst today - entry above Rs 316",
        "#3  ASIANPAINT - Grade A - paint COGS relief + tight base - entry above Rs 2,525",
    ])

    add_para(doc, "ONE RISK THAT RUINS EVERYTHING:", bold=True, color=RGBColor(0xC0, 0x00, 0x00))
    add_para(doc, "A Tehran statement rejecting the US 14-point proposal before 11:00 IST. Crude rips +$5, gap fills instantly, all CALL trades stop out at SL within 30 mins. This is binary headline risk - keep tight time-stops, never average down.", size=11)

    # ===== One-line summary =====
    add_heading(doc, "One-Line Summary (read at 9:10 AM)", level=2)
    summary = (
        "Today is CAUTIOUS BULL because crude crashed -5% on Iran deal headline and GIFT Nifty gapped +199. "
        "Crude Rs 8,352 = MILD BULL, falling. Watch INDIGO Call (Rs 4,300) and BPCL Call (Rs 320). "
        "Key risk: Tehran rebuttal flips crude back up. Size HALF (VIX 19.4). "
        "Flips if WTI back above $103 or Nifty breaks 23,560."
    )
    p = doc.add_paragraph()
    r = p.add_run(summary)
    r.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    # Footer disclaimer
    add_para(doc, "Disclaimer: Educational/analytical brief. Not investment advice. Verify all live prices and OI on NSE before entry. Some data fields marked DATA_UNAVAILABLE due to source limitations at brief generation time (03:15 IST).", italic=True, size=8, color=RGBColor(0x70, 0x70, 0x70))

    import os
    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    doc.save(OUT_PATH)
    print(f"Saved: {OUT_PATH}")


if __name__ == "__main__":
    main()
