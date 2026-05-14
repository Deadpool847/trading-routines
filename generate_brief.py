"""Generate the daily trading-intelligence .docx brief."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-05-14"
BIAS = "RANGE"
ENTRY_PERMISSION = "YELLOW"
GEN_TIME = "08:55 IST"

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


def add_kv_table(doc, rows, col_widths=(1.6, 2.0, 2.4)):
    table = doc.add_table(rows=1, cols=3)
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    hdr[0].text = "Metric"
    hdr[1].text = "Value"
    hdr[2].text = "Signal"
    for c in hdr:
        for p in c.paragraphs:
            for r in p.runs:
                r.bold = True
        shade_cell(c, "1F4E78")
        for p in c.paragraphs:
            for r in p.runs:
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for metric, value, signal in rows:
        row = table.add_row().cells
        row[0].text = metric
        row[1].text = value
        row[2].text = signal
    for i, w in enumerate(col_widths):
        for cell in table.columns[i].cells:
            cell.width = Inches(w)
    return table


def add_setup_block(doc, stock):
    add_para(
        doc,
        "══════════════════════════════════════",
        bold=True,
        color=RGBColor(0x1F, 0x4E, 0x78),
    )
    add_para(
        doc,
        f"{stock['name']} — NSE: {stock['ticker']} | {stock['side']} | Grade {stock['grade']}",
        bold=True,
        size=13,
        color=RGBColor(0xC0, 0x00, 0x00) if stock['side'] == "PUT" else RGBColor(0x00, 0x70, 0x30),
    )
    add_para(
        doc,
        f"CMP: ₹{stock['cmp']} | Lot: {stock['lot']} | Crude aligned: {stock['crude_aligned']}",
        bold=True,
    )

    add_para(doc, f"CATALYST: {stock['catalyst']}")

    add_para(doc, "TECHNICAL:", bold=True)
    add_bullets(
        doc,
        [
            f"Daily trend: {stock['daily_trend']}",
            f"Above 200 DMA: {stock['above_200dma']}",
            f"Weekly trend: {stock['weekly_trend']}",
            f"Support ₹{stock['support']} | Resistance ₹{stock['resistance']}",
        ],
    )

    add_para(doc, "STRUCTURAL:", bold=True)
    add_bullets(
        doc,
        [
            f"OI direction: {stock['oi_direction']}",
            f"Delivery %: {stock['delivery']}",
            f"Block deal: {stock['block_deal']} | Ban list: {stock['ban_list']}",
        ],
    )

    add_para(doc, "OPTION SETUP:", bold=True)
    add_bullets(
        doc,
        [
            f"Strike: ₹{stock['strike']} [{stock['moneyness']}] | Expiry: {stock['expiry']}",
            f"THETA WARNING if expiry <= 3 days: {stock['theta_warn']}",
        ],
    )

    add_para(doc, "ENTRY (ALL 4 must fire):", bold=True)
    add_bullets(
        doc,
        [
            f"ST {stock['st']} on 15-min confirmed",
            f"RSI {stock['rsi_cond']} 50 on 15-min",
            f"StochRSI cross {stock['stoch']} from extreme",
            "Volume > 1.5x 20-period avg",
            f"Confirmation candle: 15-min close {stock['cand_dir']} ₹{stock['cand_level']}",
        ],
    )

    add_para(
        doc,
        f"T1 (book 40%): ₹{stock['t1']} | T2 (book 40%): ₹{stock['t2']}",
        bold=True,
    )
    add_para(doc, f"SL: ₹{stock['sl']} + {stock['sl_cond']}", bold=True)
    add_para(doc, f"R:R: {stock['rr']}", bold=True)
    add_para(doc, "Time stop: Exit by 13:00 IST", italic=True)
    add_para(doc, f"GRADE REASON: {stock['grade_reason']}", italic=True)
    add_para(doc, "")


def main():
    doc = Document()

    title = doc.add_heading(
        f"ELITE DAILY TRADING BRIEF — {DATE_ISO} — {BIAS}", level=0
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_para(
        doc,
        f"Generated: {GEN_TIME} | Weekly expiry: 5 days (Tue 19-May) | "
        f"Monthly expiry: 12 days (Tue 26-May index / Thu 28-May stocks) | Expiry week: NO",
        italic=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )

    cbox = doc.add_table(rows=5, cols=1)
    cbox.style = "Light List Accent 5"
    cbox.rows[0].cells[0].text = "BIAS: RANGE (cautious-bull lean)"
    cbox.rows[1].cells[0].text = "ENTRY PERMISSION: YELLOW"
    cbox.rows[2].cells[0].text = "CRUDE MODE: CAUTION (Rs 8,500-9,000) | Direction: ELEVATED / sideways-up"
    cbox.rows[3].cells[0].text = "VIX SIZING: HALF SIZE (India VIX 19.42, mandatory 15-20 band)"
    cbox.rows[4].cells[0].text = "EXPIRY ALERT: NO (next NSE weekly Tue 19-May)"
    for r in cbox.rows:
        for c in r.cells:
            for p in c.paragraphs:
                for run in p.runs:
                    run.bold = True
            shade_cell(c, "FFF2CC")

    add_heading(doc, "SECTION 1 — MACRO SNAPSHOT", level=1)

    rows = [
        ("MCX Crude (proxy)", "~ Rs 8,740 / bbl", "CAUTION — half size"),
        ("Brent / WTI", "$107.0 / $102.21", "ELEVATED (Hormuz war premium)"),
        ("GIFT Nifty gap", "+171 pts (23,584 vs 23,412.60)", "Gap-up — bullish open"),
        ("S&P 500 (12-May close)", "7,444.25 (+0.58%)", "Risk-on; record close"),
        ("Nasdaq (12-May close)", "26,402.34 (+1.20%)", "Tech-led record"),
        ("Dow (13-May)", "Slipped on hot PPI", "Mixed risk-off undertone"),
        ("US VIX", "DATA_UNAVAILABLE", "Recent calm but PPI shock"),
        ("India VIX", "19.42 (+0.75%)", "HALF-SIZE mandatory (15-20)"),
        ("Nifty close (13-May)", "23,412.60 (+33.05, +0.14%)", "Recovered intraday low"),
        ("FII cash (13-May)", "-Rs 4,520.91 Cr", "SELL (5th day of net sell)"),
        ("DII cash (13-May)", "+Rs 5,523.51 Cr", "BUY — absorbing FII supply"),
        ("USD-INR (est.)", "~ Rs 85.5 (war-driven weakness)", "Weak rupee — IT mixed"),
        ("INFY ADR", "$11.69 vs prev $12.05 = -2.99%", "INFY likely opens DOWN"),
        ("WIT ADR", "$1.84 (-1.87%)", "Wipro opens DOWN"),
        ("IBN ADR", "$25.75 (-0.10%)", "ICICI Bank flat"),
        ("HDB ADR", "~ $23.79 (recent -3.3%)", "HDFC Bank flat-to-down"),
    ]
    add_kv_table(doc, rows)

    add_para(doc, "")
    add_para(
        doc,
        "Geopolitical one-liner: Strait of Hormuz near-shut since 28-Feb US/Israel-Iran war; "
        "Pakistan-mediated talks ongoing; ~14.5 mb/d supply gap; ~2,000 vessels stranded.",
        italic=True,
    )
    add_para(
        doc,
        "Today's intraday headline risk: HIGH (any Iran/talks headline = whippy crude & defensives).",
        bold=True,
        color=RGBColor(0xC0, 0x00, 0x00),
    )

    add_heading(doc, "SECTION 2 — CRUDE RULE + STRUCTURAL READ", level=1)

    add_para(doc, "CRUDE_MODE: CAUTION (band Rs 8,500-9,000) — half size mandatory", bold=True)
    add_para(doc, "CRUDE_DIRECTION: ELEVATED; sideways-up bias; war premium intact", bold=True)
    add_para(
        doc,
        "CRUDE_FLIP_LEVEL: Rs 9,000 -> STRONG BEAR (puts only). "
        "Downside flip Rs 7,500 -> MILD BULL (would need Hormuz reopening headline).",
        bold=True,
    )
    add_para(doc, "AVOID: BPCL, IOC, paints (APNT/BRGR), aviation (INDIGO), tyres (APOLLOTYRE/MRF), fertilisers (CHAMBLFERT)", bold=True)
    add_para(doc, "EXCEPTION_CALLS: ONGC, OIL INDIA, MRPL (upstream oil); defence (BEL, HAL) is war-narrative ally", bold=True)

    add_para(doc, "")
    add_para(doc, "Structural read (Nifty internals):", bold=True)
    add_bullets(
        doc,
        [
            "Nifty 23,412.60 vs estimated Max Pain ~23,400-23,500 -> magnet near current spot (range bias).",
            "Highest OI Call resistance (est.): 23,500 / 23,600 — ceiling for the day.",
            "Highest OI Put support (est.): 23,300 / 23,200 — floor.",
            "PCR: DATA_UNAVAILABLE specifically; inferred 0.85-0.95 from FII sell + heavy DII buy + flat close.",
            "Futures basis: usually slight premium when DIIs buying; flag at open if discount appears = warning.",
            "OI change reading: Tuesday's weekly expired calm; fresh OI build expected at 23,500CE / 23,300PE on open.",
        ],
    )
    add_para(
        doc,
        "Bank Nifty leadership check: Banking ADRs (IBN flat, HDB slightly weak) -> "
        "Bank Nifty likely LAGGING. Any Nifty rally not confirmed by Bank Nifty is SUSPECT — fade.",
        bold=True,
    )

    add_heading(doc, "SECTION 3 — CONTRARIAN CHECK (Layer 3)", level=1)
    add_para(
        doc,
        "Q1. CONSENSUS: Retail and TV analysts will say 'gap-up follow-through to 23,600, "
        "buy on dips, defence-and-PSU rally'. Most will chase opening 15-min strength in BEL / "
        "Adani names and shun IT.",
        bold=True,
    )
    add_para(
        doc,
        "Q2. THE TRAP: Gap-ups against a 5-day FII sell streak and elevated VIX historically reverse "
        "within the first 90 minutes. If Bank Nifty fails to confirm and crude ticks up on any Hormuz "
        "headline, the gap fills and 23,300 PE writers get squeezed — but heavy DII bid prevents a true "
        "break. Net: chop both ways, theta-rich pin to max pain. Option sellers profit; option buyers bleed.",
        bold=True,
    )
    add_para(
        doc,
        "Q3. RETAIL STOPS: Long stops clustered below 23,350 (gap-fill) and 23,300 (Put support). "
        "Short stops above 23,550-23,600 (Call wall). Both sides will get tagged before any directional move.",
        bold=True,
    )
    add_para(
        doc,
        "Q4. FLIP TRIGGER: A single Reuters/CNBC flash on (i) Hormuz reopening = explosive UP day, "
        "fade all OMC puts, chase IT calls; or (ii) Iran retaliation = crude +5% = SHARP DOWN, IT and "
        "OMCs collapse together. Watch the wires till 12:30 PM.",
        bold=True,
    )

    add_para(doc, "")
    add_heading(doc, "Contrarian Override Logic Applied", level=2)
    add_bullets(
        doc,
        [
            "A: GIFT Nifty gap-up > +100 pts? YES (+171).",
            "B: Crude falling > 2% overnight? NO (crude flat-to-elevated).",
            "C: Contrarian scenario probability > 40%? YES (gap-fill scenario plausible).",
            "D: Expiry week + PCR < 0.8? NO (5 days to weekly, not expiry week).",
        ],
    )
    add_para(
        doc,
        "Result: A YES + B NO -> NO full override. C YES alone keeps confidence at MEDIUM, "
        "not HIGH. Final bias = RANGE with mild bull lean; permission YELLOW.",
        bold=True,
        color=RGBColor(0xC0, 0x55, 0x00),
    )

    add_heading(doc, "SECTION 4 — 5 INDIVIDUAL F&O SETUPS", level=1)

    setups = [
        {
            "name": "ONGC",
            "ticker": "ONGC",
            "side": "CALL",
            "grade": "A+",
            "cmp": "297.15",
            "lot": "2,250",
            "crude_aligned": "YES (upstream beneficiary)",
            "catalyst": "Brent $107 / WTI $102 — realisation tailwind; 71% of India's crude output; "
                        "near 52-wk high Rs 307.50; royalty cut to 8% boosts margin.",
            "daily_trend": "UP",
            "above_200dma": "YES",
            "weekly_trend": "UP (+20.69% YoY, +18.06% 6m)",
            "support": "290 / 285",
            "resistance": "307.50 (52-wk high) / 315",
            "oi_direction": "LONG BUILD (price up + OI up trend last 3 sessions)",
            "delivery": "RISING",
            "block_deal": "NO",
            "ban_list": "NO",
            "strike": "300",
            "moneyness": "ATM",
            "expiry": "28-May-2026 (monthly)",
            "theta_warn": "NO (14 days to expiry)",
            "st": "GREEN",
            "rsi_cond": "ABOVE",
            "stoch": "UP",
            "cand_dir": "ABOVE",
            "cand_level": "299.50",
            "t1": "305.00",
            "t2": "310.00",
            "sl": "294.50",
            "sl_cond": "ST flip to RED on 15m",
            "rr": "2.4:1",
            "grade_reason": "Best directional alignment — crude tailwind + chart breakout zone + ADR sector clue.",
        },
        {
            "name": "Bharat Electronics",
            "ticker": "BEL",
            "side": "CALL",
            "grade": "A",
            "cmp": "428.10",
            "lot": "2,850",
            "crude_aligned": "Neutral (war-narrative beneficiary)",
            "catalyst": "Fresh Rs 1,251 Cr GBMES + Rs 569 Cr avionics orders; order book Rs 73,015 Cr; "
                        "war premium intact while Hormuz unresolved.",
            "daily_trend": "UP",
            "above_200dma": "YES",
            "weekly_trend": "UP (+10% YTD, +41% YoY)",
            "support": "418 / 410",
            "resistance": "445 / 473.45 (52-wk high)",
            "oi_direction": "LONG BUILD continuing",
            "delivery": "STABLE",
            "block_deal": "NO",
            "ban_list": "NO",
            "strike": "430",
            "moneyness": "ATM/slight-OTM",
            "expiry": "28-May-2026 (monthly)",
            "theta_warn": "NO",
            "st": "GREEN",
            "rsi_cond": "ABOVE",
            "stoch": "UP",
            "cand_dir": "ABOVE",
            "cand_level": "431.50",
            "t1": "438.00",
            "t2": "445.00",
            "sl": "423.50",
            "sl_cond": "15m close below 425 & ST RED",
            "rr": "2.2:1",
            "grade_reason": "Cleanest war-narrative momentum stock; valuation high (PE 51) so size half.",
        },
        {
            "name": "Infosys",
            "ticker": "INFY",
            "side": "PUT",
            "grade": "A",
            "cmp": "1,129.50",
            "lot": "400",
            "crude_aligned": "Aligned (weak rupee from crude is complex; discretionary IT spend collapsing globally)",
            "catalyst": "ADR INFY at $11.69 vs $12.05 = -2.99% gap-down; stock near 52-wk low Rs 1,123; "
                        "RSI 38.82 already weak; in yesterday's losers; client IT budget cuts amid US PPI shock.",
            "daily_trend": "DOWN",
            "above_200dma": "NO (well below)",
            "weekly_trend": "DOWN",
            "support": "1,123 (52-wk low) / 1,100",
            "resistance": "1,140 / 1,160",
            "oi_direction": "SHORT BUILD (price down + OI up)",
            "delivery": "STABLE",
            "block_deal": "NO",
            "ban_list": "NO",
            "strike": "1,120",
            "moneyness": "ATM after gap",
            "expiry": "28-May-2026 (monthly)",
            "theta_warn": "NO",
            "st": "RED",
            "rsi_cond": "BELOW",
            "stoch": "DOWN",
            "cand_dir": "BELOW",
            "cand_level": "1,120.00",
            "t1": "1,108.00",
            "t2": "1,098.00",
            "sl": "1,134.00",
            "sl_cond": "15m close above 1,134 & ST GREEN",
            "rr": "2.0:1",
            "grade_reason": "ADR-pre-confirmed direction; risk = 52-wk low magnet bounce, hence tight SL.",
        },
        {
            "name": "Asian Paints",
            "ticker": "ASIANPAINT",
            "side": "PUT",
            "grade": "A-",
            "cmp": "2,583.00",
            "lot": "200",
            "crude_aligned": "YES (crude-derivative inputs = 55% of RM)",
            "catalyst": "Yesterday's +3.17% spike into resistance 2,560-2,565; 2nd price hike effective "
                        "5-May = margin admission, not strength; paint stocks tumbled 5% on prior crude rally; "
                        "valuation concerns flagged.",
            "daily_trend": "Choppy UP into resistance",
            "above_200dma": "YES (above 5/20/50/100/200 DMA)",
            "weekly_trend": "Sideways-up; faded prior peaks",
            "support": "2,460 / 2,380 / 2,300",
            "resistance": "2,600 (psychological + heavy call OI 6,910 contracts)",
            "oi_direction": "Call writing at 2,600 (resistance defended)",
            "delivery": "STABLE",
            "block_deal": "NO",
            "ban_list": "NO",
            "strike": "2,560",
            "moneyness": "Slight-ITM after open",
            "expiry": "28-May-2026 (monthly)",
            "theta_warn": "NO",
            "st": "RED",
            "rsi_cond": "BELOW",
            "stoch": "DOWN (from >70)",
            "cand_dir": "BELOW",
            "cand_level": "2,570.00",
            "t1": "2,530.00",
            "t2": "2,495.00",
            "sl": "2,605.00",
            "sl_cond": "15m close above 2,605 (call wall break)",
            "rr": "2.1:1",
            "grade_reason": "Classic fade-the-spike into call wall with crude margin headwind; lot of 200 = half-size friendly.",
        },
        {
            "name": "HPCL",
            "ticker": "HINDPETRO",
            "side": "PUT",
            "grade": "B+",
            "cmp": "377.55",
            "lot": "2,700",
            "crude_aligned": "YES (OMC = crude victim, downstream)",
            "catalyst": "Yesterday's +3% post-Q4 beat = sell-the-news candidate; YTD -24% vs Nifty -10% "
                        "shows structural weakness; GRM $14.27 looks peak quarter, mean-reverts as crude stays "
                        "high. Dividend already priced.",
            "daily_trend": "Bounce within downtrend",
            "above_200dma": "NO",
            "weekly_trend": "DOWN (-22.14% 6m)",
            "support": "365 / 358",
            "resistance": "385 / 395",
            "oi_direction": "SHORT BUILD on bounce (price up + OI flat-to-up)",
            "delivery": "FALLING (event-driven bounce only)",
            "block_deal": "NO",
            "ban_list": "NO (verify at 8:45 AM)",
            "strike": "375",
            "moneyness": "ATM",
            "expiry": "28-May-2026 (monthly)",
            "theta_warn": "NO",
            "st": "RED",
            "rsi_cond": "BELOW",
            "stoch": "DOWN",
            "cand_dir": "BELOW",
            "cand_level": "374.00",
            "t1": "367.00",
            "t2": "360.00",
            "sl": "382.50",
            "sl_cond": "15m close above 382.50 (post-result gap fill)",
            "rr": "2.0:1",
            "grade_reason": "B+ not A because Q4 beat creates short-term support; crude thesis durable but timing soft.",
        },
    ]
    for s in setups:
        add_setup_block(doc, s)

    add_heading(doc, "SECTION 5 — FINAL VERDICT", level=1)

    verdict_rows = [
        ("TODAY'S BIAS", "RANGE (cautious-bull lean)"),
        ("CONFIDENCE", "MEDIUM"),
        ("CRUDE MODE", "CAUTION (Rs 8,500-9,000 band; live verified)"),
        ("VIX SIZING", "HALF SIZE (India VIX 19.42)"),
        ("ENTRY PERMISSION", "YELLOW"),
    ]
    vt = doc.add_table(rows=0, cols=2)
    vt.style = "Light Grid Accent 1"
    for k, v in verdict_rows:
        row = vt.add_row().cells
        row[0].text = k
        row[1].text = v
        for run in row[0].paragraphs[0].runs:
            run.bold = True

    add_para(doc, "")
    add_para(doc, "THE BULL CASE:", bold=True, color=RGBColor(0x00, 0x70, 0x30))
    add_para(
        doc,
        "GIFT Nifty +171, US indices at records (May 12), DIIs net +Rs 5,523 Cr absorbing FII supply, "
        "war-narrative names (ONGC, BEL, defence) rallying on real order flow.",
    )

    add_para(doc, "THE BEAR CASE:", bold=True, color=RGBColor(0xC0, 0x00, 0x00))
    add_para(
        doc,
        "FII net sellers 5 consecutive days, India VIX 19.42 elevated, crude $107 sticky, INFY ADR "
        "-3% pulls IT down, US PPI hotter than expected, Hormuz risk asymmetric to upside in crude.",
    )

    add_para(doc, "FLIP TRIGGER:", bold=True)
    add_para(
        doc,
        "If Nifty 15-min closes below 23,300 with Bank Nifty confirming -> bias flips to BEAR, "
        "exit all longs, add puts. If crude pulls back >2% on a Hormuz reopening headline AND Nifty "
        "breaks 23,580 -> bias flips to BULL, exit OMC/paint puts immediately.",
    )

    add_para(doc, "")
    add_para(doc, "NIFTY KEY LEVELS:", bold=True)
    add_para(
        doc,
        "S2 23,200 | S1 23,300 | CRITICAL 23,400 (pivot/max-pain) | R1 23,500 | R2 23,600",
        bold=True,
    )

    add_para(doc, "")
    add_para(doc, "TOP 3 RANKED:", bold=True)
    add_bullets(
        doc,
        [
            "#1 ONGC — A+ — Upstream crude beneficiary near 52-wk high — entry ABOVE Rs 299.50 (300 CE).",
            "#2 BEL — A — War-narrative defence momentum + fresh orders — entry ABOVE Rs 431.50 (430 CE).",
            "#3 INFY — A — ADR-confirmed gap-down setup near 52-wk low — entry BELOW Rs 1,120 (1120 PE).",
        ],
    )

    add_para(doc, "")
    add_para(doc, "ONE RISK THAT RUINS EVERYTHING TODAY:", bold=True, color=RGBColor(0xC0, 0x00, 0x00))
    add_para(
        doc,
        "Surprise Hormuz reopening / ceasefire headline during market hours -> crude collapses 5-8%, "
        "ONGC/BEL gap down, INFY and paints rip higher in seconds. All 5 setups invalidated simultaneously. "
        "Mitigate: half-size + 13:00 IST time-stop + ready to flip with the headline.",
    )

    add_heading(doc, "ONE-LINE SUMMARY (read at 9:10 AM)", level=1)
    add_para(
        doc,
        "\"Today is RANGE with mild-bull lean because gap-up +171 meets crude-elevated headwind and "
        "FII selling. Crude ~ Rs 8,740 = CAUTION, sideways-up. Watch ONGC CALL and INFY PUT. Key risk: "
        "Hormuz headline whipsaw. Size HALF. Flips if Nifty 15m-closes below 23,300 (-> BEAR) or above "
        "23,580 on crude crash (-> BULL).\"",
        italic=True,
        bold=True,
    )

    doc.save(OUT_PATH)
    print(f"Saved: {OUT_PATH}")


if __name__ == "__main__":
    main()
