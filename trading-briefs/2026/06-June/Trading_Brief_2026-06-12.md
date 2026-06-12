# TODAY — Friday, June 12, 2026

---

## YESTERDAY'S POSTMORTEM

> **STATUS: CANNOT RUN — First brief executed in this session. No prior brief found in repository (trade_log.txt empty; April 23 docx not machine-readable). Establishing fresh methodology baseline from today.**

- Brief accuracy score: **N/A (baseline session)**
- What worked: N/A
- What missed: N/A
- Today's methodology adjustment: None — running full routine clean.

**ACTION REQUIRED FOR OPERATOR:** Please log today's outcome in `trading-briefs/trade_log.txt` using this format:
```
DATE | STOCK | DIRECTION | RESULT (HIT T1 / HIT T2 / SL / TIME-STOP) | PNL_PTS
```
This enables postmortem from tomorrow's brief onward.

---

## TODAY'S MACRO FORCES (5)

| # | Force | Reading | Trading Implication |
|---|-------|---------|---------------------|
| 1 | **Trump suspends Iran strikes** | WTI fell ~4% to $86.65 (lowest since April). Crude regime SHIFTS overnight from CAUTION to MILD_BULL. US markets soared 1.8–2.5%. | **BULLISH for India** — gap-up open, OMC/aviation/paints benefit. But this is ONE headline from ONE person. Fragile. |
| 2 | **US markets surge** | Dow +923 pts (+1.85%), S&P +1.8%, Nasdaq +2.5%. Tech led: Micron +11%, AMD +8%, Lam Research +12.7%, Intel +10%. SpaceX IPO (largest in history, $75B at $1.75T valuation) begins trading on Nasdaq today. | **BULLISH for Indian IT** — TCS/Infy dip from June 11 will recover. Tech euphoria global. |
| 3 | **IAEA finds Iran NON-COMPLIANT** | IAEA announced June 12 (TODAY, before India market open) that Iran is non-compliant with nuclear obligations — **first finding in 20 years**. Iran defence minister: "all US bases in range if talks collapse." | **BEARISH surprise risk** — This is a morning headline. Markets may not have priced it. If crude bounces on this, regime flips back to CAUTION. |
| 4 | **RBI held rates at 5.25%** (June 5 decision) | Third consecutive hold. Hawkish CPI projection RAISED to 5.1% (from 4.6%). GDP lowered to 6.6%. Stance: Neutral. | **NEUTRAL-HAWKISH** — No rate-cut tailwind for banks. IT/consumer positions not hurt. Watch for continued RBI commentary. |
| 5 | **India-US Trade Deal** — first tranche by mid-July | 99% of deal agreed. US tariff on Indian goods at 18% (down from 50%). First tranche target: mid-July 2026. US team was in New Delhi June 2-4. | **SLOW BULLISH** — Structural multi-week tailwind for IT exports and manufacturing. Not a day-trade catalyst alone, but confirms macro backdrop. |

---

## MACRO SNAPSHOT TABLE

| Metric | Reading | Source | Signal |
|--------|---------|--------|--------|
| WTI Crude (NYMEX) | **$86.65/bbl** | CNBC/TradingEconomics | MILD_BULL trigger |
| Brent Crude (context only — NOT for MCX proxy) | **$89.14/bbl** | TradingEconomics | Brent premium $2.49 over WTI |
| USDINR spot | **95.26** | TradingEconomics (June 11 close) | Near record INR weakness |
| MCX Crude (India close, June 11) | **₹8,512/bbl** (–1.42%) | univest.in/MCX | Was AT CAUTION/MILD_BULL boundary |
| MCX Crude proxy (WTI × USDINR, today) | **~₹8,254/bbl** | Computed: 86.65 × 95.26 | MILD_BULL (but boundary risk — see below) |
| India VIX (June 11 close) | **15.63** | NSE India | **HALF size** (15–20 range) |
| Nifty 50 (June 11 close) | **23,161.60** (–53.35, –0.23%) | NSE India | Bearish close; intraday low 23,072 |
| Nifty June Futures (June 11 close) | **23,147.10** | NSE India | 24.1-pt premium to spot |
| Sensex (June 11 close) | **73,832** (–~150 pts) | BSE India | Bearish close |
| Bank Nifty (June 11 close) | **55,176.75** | NSE India | Bullish candle, high 55,600 |
| GIFT Nifty (June 12 @ 7:06 AM IST) | **23,354** (–48 from prev GIFT close) | NSE IFSC/equitypandit | Implied gap: **+207 pts** vs Nifty futures close |
| Nikkei 225 | **66,204** (+3.09%) | Bloomberg | Strong Asian tailwind |
| Hang Seng | **24,501** (+1.04%) | Bloomberg | Asian support |
| Dow Jones | **50,841** (+923, +1.85%) | TheStreet | Strong US close |
| S&P 500 | +1.8% | TheStreet | US bull continues |
| Nasdaq | +2.5% | TheStreet | Tech-led rally |
| FII Cash (June 11, provisional) | **–₹1,987 Cr** (net selling) | NSE India | Still selling in cash |
| FII Index Futures | **+₹780 Cr** (net buying) | NSE India | Hedged positioning — partial reversal |
| DII Cash (June 11) | **+₹4,225 Cr** (net buying) | NSE India | DII absorbing FII outflow (+₹2,238 Cr net positive) |

> **SOURCE DISCREPANCY FLAG:** One source cited WTI at $91.56 (intraday high, before Trump statement). The $86.65 figure reflects the post-announcement level. Using $86.65 as it is the live/closing reference. Brent at $89.14 is also post-announcement. Do NOT use Brent for MCX proxy — always WTI.

---

## CRUDE REGIME

```
WTI (current):     $86.65/bbl
USDINR:            95.26
MCX PROXY:         $86.65 × 95.26 = ₹8,254/bbl
MCX ACTUAL CLOSE (June 11):  ₹8,512/bbl  ← was sitting AT the regime boundary

Regime (implied today):  MILD_BULL  (₹7,500–8,500 range)
Regime (yesterday close): CAUTION    (₹8,500–9,000 range)
→ OVERNIGHT REGIME SHIFT: CAUTION → MILD_BULL
```

**Direction:** FALLING sharply (–4% intraday, –1.42% at India MCX close)

> ⚠️ **BOUNDARY RISK — CRITICAL FLAG:**
> MCX proxy today ₹8,254 is only ₹246 below the CAUTION boundary (₹8,500).
> WTI only needs to rise from $86.65 → $89.25 (a mere +3%) to flip regime back to CAUTION.
> Given fresh IAEA non-compliance finding (June 12 morning), this risk is LIVE.
> **Watch MCX crude price throughout the session. If MCX crosses ₹8,400 and holds, begin reducing CALL positions. If MCX crosses ₹8,500 on 15-min close — EXIT ALL CALLs immediately.**

---

## F&O BAN LIST — June 12, 2026

**Banned stocks (no new positions):**
- Aditya Birla Fashion
- CDSL (Central Depository Services)
- Chambal Fertilizers
- Hindustan Copper
- IEX (Indian Energy Exchange)
- IREN (Indian Renewable Energy)
- RBL Bank
- Titagarh Rail Systems

None of today's recommended stocks appear on this list. ✓

---

## WEEKLY / MONTHLY EXPIRY CALENDAR

| Expiry | Date | Notes |
|--------|------|-------|
| Nifty weekly | **June 11 (YESTERDAY)** | Already expired. New cycle started today. |
| Nifty weekly (next) | June 18 (Thu) | Only 6 days — HIGH theta burn. **AVOID for new entries.** |
| Nifty monthly | **June 25 (Thu)** | Use this for ALL new positions today. |
| Bank Nifty monthly | June 25 (Thu) | Only monthly expiry (no weeklies since Oct 2024). |
| July monthly | July 30 (Thu) | For longer-dated trades if needed. |

> **RULE:** Use June 25 or July monthly options for all setups today. Do NOT use June 18 weeklies under any circumstance.

---

## NIFTY STRUCTURE

| Level | Meaning |
|-------|---------|
| **23,550** | Key resistance — round number + prior weekly high zone |
| **23,500** | Psychological resistance — bull target today |
| **~23,370** | Expected open (GIFT implied: +207 pts over futures close of 23,147) |
| **23,327** | June 11 intraday HIGH — Nifty opens ABOVE this today |
| **23,250** | ⭐ **THE ONE NUMBER — Gap-fill zone / Intraday pivotal support** |
| 23,161 | June 11 spot close — support if gap partially fills |
| 23,104 | June 11 gap-down open level |
| 23,072 | June 11 intraday LOW |
| **23,000** | STRONG SUPPORT — psychological floor |
| 22,800 | Bear case target (if regime flips and Iran deal collapses) |

**Today opens ABOVE the June 11 intraday high (23,327) for the first time this week — no overhead supply from yesterday's session at the open.**

---

## CONTRARIAN CHECK

**Layer 1 — Consensus narrative:**
"Massive gap-up today! Trump cancelled Iran strikes, crude crashed, US markets up 2.5%, SpaceX IPO euphoria. Buy IT, buy OMCs, buy everything. Nifty tests 23,500 today. Rally is ON — get in before you miss it."

**Layer 2 — What consensus is likely WRONG about:**
- The IAEA just declared Iran non-compliant (June 12 morning, BEFORE Indian market opens). This is a material risk the consensus is not pricing because it happened after the bullish headlines.
- Trump's "suspension" is not a peace deal. It is one man's mood on one morning. Iran-US deals have collapsed within 48 hours twice in 2026. This pattern is DOCUMENTED.
- The crude regime "shift" is ₹258 above the CAUTION boundary. One negative headline reverses it in minutes.
- VIX at 15.63 is NOT complacent — market is still pricing meaningful risk. Half-size rule is non-negotiable.
- FII are STILL net sellers in cash (–₹1,987 Cr). The buying is DII-driven. FII cash selling continuing = foreign money is not yet convinced.
- RBI's hawkish CPI projection (5.1%) means financial sector has a headwind. The gap-up in banks may not sustain.

**Layer 3 — What everyone is underestimating:**
- **Weekend black swan risk:** Trump says Iran deal signing "could be this weekend in Europe." If the deal fails over the weekend → crude spikes Monday morning → every CALL position opened today faces a gap-down Monday. Friday sessions with weekend geopolitical events MUST have a PUT hedge.
- **SpaceX IPO intraday volatility:** World's largest IPO ever starts trading on Nasdaq TODAY during Indian market hours (US pre-market 4–9 AM EST = 1:30–6:30 PM IST). US markets could see wild intraday swings from retail scramble → could affect GIFT Nifty late-session and US Nasdaq futures.
- **AI disruption structural headwind on IT:** June 11 IT sell-off was partially from "AI fears roiling IT stocks" (distinct from geopolitics). This is a SLOW-BURNING structural concern. TCS recovery today is tactical, not structural.
- **Monsoon risk:** No search result specifically mentioned this today, but India IMD forecast is below-normal monsoon. With CPI already hawkish (5.1% projection), food inflation from weak monsoon could push RBI into an unexpected hawkish turn next meeting.

**Flip triggers:**
- BULLISH → BEARISH: Any Iran breakdown headline | MCX crude above ₹8,500 on 15-min close | Nifty breaks 23,250 post-first-hour
- BEARISH → BULLISH: Iran deal formally signed (removes weekend risk) | FII turns net buyer in first-hour provisional data | MCX crude drops below ₹8,000

**Entry permission: 🟡 YELLOW**
Gap-up opens, but IAEA risk + regime boundary + Friday weekend risk = caution warranted.
**Wait 30 min minimum (A+ TCS), 45 min (A and B+ stocks) before first entry.**
Do NOT chase the gap. If Nifty is above 23,350 at 9:45 AM without any Iran news, then GREEN conditions apply.

---

## STOCK SETUPS

### #1 — TCS CALL — Grade A+
**CMP:** ~₹3,780 (verify at open) | **Lot:** 75 | **Sector:** IT/Technology
**Capital test:** 75 × ₹3,780 × 3% = **₹8,505 ✓**

**Why A+:**
Multiple independent catalysts all converging:
1. Nasdaq +2.5% — TCS fell ~1% on June 11 on geopolitical selling; US tech recovery = direct dip-reversal
2. Fresh brokerage targets published June 9 (Motilal Oswal, ICICI Sec, others)
3. SpaceX IPO = global tech sector euphoria today; positive sentiment contagion to Indian IT
4. India-US trade deal first tranche by mid-July = direct IT export pipeline confirmation
5. Technical: June 11 sell-off was on EXTERNAL fear, not fundamental; dip is buyable
**Single headwind to monitor:** AI disruption concern is ongoing (structural) — not a day-specific risk but watch for any analyst downgrade comments.

**Option setup:** TCS 3800 CE June 25 Monthly
**Entry trigger (all 4 required):**
1. Nifty holds ABOVE **23,250** (gap-fill zone support) after 30 minutes (by 9:45 AM IST)
2. TCS 15-min candle closes ABOVE **₹3,820**
3. MCX crude at or below **₹8,350** (MILD_BULL regime holding)
4. Nifty IT sector index showing positive breadth (> 60% IT stocks green)

| Level | Price |
|-------|-------|
| Entry zone | ₹3,800–3,820 |
| T1 (40% book) | ₹4,050 |
| T2 (40% book) | ₹4,250 |
| Trailing (remaining 20%) | ₹4,400+ |
| **SL (15-min close)** | **₹3,680** |
| **R:R** | **250:120 = 2.08:1 ✓** |
| Time stop | **3:10 PM IST — NO weekend carry** |

Grade: **A+** | Direction: **CALL** | Sector: IT

---

### #2 — HAL CALL — Grade A *(Mandatory Macro Hedge)*
**CMP:** ₹4,188.10 (confirmed) | **Lot:** 75 | **Sector:** Defence/Aerospace
**Capital test:** 75 × ₹4,188 × 3% = **₹9,423 ✓**

**Why A (Macro Hedge):**
HAL is today's required macro hedge. Geopolitical uncertainty does not disappear because Iran talks are "progressing" — they have collapsed twice this year. HAL benefits regardless:
- If Iran deal succeeds → India defence budget is already locked in (+12–15% YoY FY27). Budget allocation doesn't depend on current Iran status.
- If Iran deal collapses → crude spikes, geopolitical fear rises, defence stocks get a direct bid.
- HAL order book exceeds ₹1 lakh crore (₹48,000 Cr Tejas Mk-1A alone). HALE/MALE drone program = new revenue stream.
- June 11 sold off on "IT and defence sell-off" alongside broad market — this is a BUY THE DIP.
- Post-Operation Sindoor (earlier 2026), defence stocks outperformed Nifty 50 by 8% vs 3% in April.

**Option setup:** HAL 4200 CE June 25 Monthly
**Entry trigger (all 4 required):**
1. Nifty holds above **23,300** after first 30 min
2. HAL 15-min candle closes ABOVE **₹4,230**
3. No negative Iran escalation headline in first 45 min (no new strikes announced)
4. Defence sector breadth positive: BEL > ₹405 or DRDO-linked stocks in green

| Level | Price |
|-------|-------|
| Entry zone | ₹4,200–4,240 |
| T1 (40% book) | ₹4,450 |
| T2 (40% book) | ₹4,700 |
| Trailing (remaining 20%) | ₹4,900+ |
| **SL (15-min close)** | **₹4,030** |
| **R:R** | **262:158 ≈ 1.66:1** ⚠️ |
| Time stop | **3:10 PM IST — NO weekend carry** |

> ⚠️ R:R note: This setup is 1.66:1, technically below the 2:1 rule. Acceptable **only because HAL is the mandatory macro hedge**. If you want cleaner R:R: wait for a dip to ₹4,050 → T1 ₹4,350, SL ₹3,950 → R:R 3:1. Entry at ₹4,050 may not come if gap-up is sustained.

Grade: **A** | Direction: **CALL** | Sector: Defence

---

### #3 — BPCL CALL — Grade A
**CMP:** ₹285.05 (confirmed) | **Lot:** ~1,800 ⚠️ *VERIFY before trading* | **Sector:** Energy/OMC
**Capital test:** 1,800 × ₹285 × 3% = **₹15,390** — *₹390 over limit*
> ⚠️ **LOT SIZE FLAG:** If NSE revised BPCL lot to 1,500 → capital test = ₹12,825 ✓. Verify exact lot size from NSE website or broker before entering. Do NOT enter if lot size confirms ≥1,800 without operator approval.

**Why A:**
WTI fell 4% overnight (from ~$91 to $86.65). MCX crude proxy fell from ₹8,512 (Indian close) to implied ₹8,254 = ~₹258 lower.
- BPCL buys crude at market rates. Government capped retail pump prices (petrol/diesel).
- Lower crude + stable pump prices = **improved marketing margin directly**.
- This is the most DIRECT and MECHANICAL beneficiary of the overnight crude fall.
- OMC sector was one of the weakest June 11 (sold off with broad market). Buy the dip.
- Catalysts: Crude regime SHIFT + government margin protection framework.

**Option setup:** BPCL 290 CE June 25 Monthly
**Entry trigger (all 4 required):**
1. MCX crude (June 12 open) stays BELOW **₹8,300** (critical — if it opens higher, abort BPCL)
2. BPCL 15-min candle closes ABOVE **₹292** (confirms gap-up holds)
3. OMC sector breadth: HPCL also in green (sector confirmation, not just BPCL-specific)
4. No Iran escalation news reversing crude direction in first hour

| Level | Price |
|-------|-------|
| Entry zone | ₹287–292 |
| T1 (40% book) | ₹312 |
| T2 (40% book) | ₹330 |
| Trailing (remaining 20%) | ₹345+ |
| **SL (15-min close)** | **₹275** |
| **R:R** | **25:12 = 2.08:1 ✓** |
| Time stop | **2:30 PM IST** (crude-sensitive — exit early on Friday) |

Grade: **A** | Direction: **CALL** | Sector: Energy/OMC

---

### #4 — TATAMOTORS CALL — Grade B+
**CMP:** ~₹650 *(VERIFY — estimate only)* | **Lot:** ~750 *(VERIFY)* | **Sector:** Automobile
**Capital test (estimated):** 750 × ₹650 × 3% = **₹14,625 ✓**

**Why B+:**
Single primary catalyst + secondary:
1. **Crude fall → fuel cost reduction** → consumer sentiment for ICE autos improves. Lower running costs = higher discretionary auto purchase intent.
2. **India-US trade deal first tranche by mid-July** → Tata Motors JLR (UK operations) and EV export pipeline benefit from reduced tariff friction.
3. Tata.ev brand momentum (India's largest EV player).
Grade B+ because: only 1-2 catalysts (not 4). Requires clean macro confirmation.

**Option setup:** TATAMOTORS 660 CE June 25 Monthly
**Entry trigger (all 4 required — stricter for B+):**
1. Nifty holds above **23,300** after **45 min** (stricter wait for B+)
2. TATAMOTORS 15-min candle closes ABOVE **₹665**
3. MCX crude below **₹8,300** (auto thesis needs crude directionally down)
4. Auto sector breadth: Maruti or M&M also positive

| Level | Price |
|-------|-------|
| Entry zone | ₹655–668 |
| T1 (40% book) | ₹705 |
| T2 (40% book) | ₹745 |
| **SL (15-min close)** | **₹630** |
| **R:R** | **50:25 = 2.0:1 ✓** |
| Time stop | **2:00 PM IST** (B+ grade = earlier exit on Friday) |

> **Skip if entry triggers don't fire by 10:45 AM IST. This is a B+ — don't chase.**

Grade: **B+** | Direction: **CALL** | Sector: Automobile

---

### #5 — ASIANPAINT CALL — Grade B+
**CMP:** ~₹2,280 *(VERIFY — estimate only)* | **Lot:** ~200 *(VERIFY)* | **Sector:** Consumer Goods/Paints
**Capital test (estimated):** 200 × ₹2,280 × 3% = **₹13,680 ✓**

**Why B+:**
Crude-derivative raw material benefit:
- TiO2 (titanium dioxide), VAM, and petrochemical-based binders are key paint inputs.
- Crude fall of 4% → input cost improvement signals → margin expansion narrative.
- Asian Paints has strong brand premium pricing power (partial pass-through in reverse).
Grade B+ (not A) because: AI Birla Opus competition is ongoing structural headwind; needs MCX crude clearly below ₹8,200 for strong conviction (not just ₹8,254 borderline).

**Option setup:** ASIANPAINT 2300 CE June 25 Monthly
**Entry trigger (all 4 required — strictest for B+):**
1. MCX crude holds BELOW **₹8,200** (stricter crude threshold than BPCL)
2. ASIANPAINT 15-min candle closes ABOVE **₹2,320**
3. Nifty holds above **23,300**
4. No negative competitive news (Birla Opus pricing announcement, market share data)

| Level | Price |
|-------|-------|
| Entry zone | ₹2,290–2,325 |
| T1 (40% book) | ₹2,470 |
| T2 (40% book) | ₹2,600 |
| **SL (15-min close)** | **₹2,200** |
| **R:R** | **180:90 = 2.0:1 ✓** |
| Time stop | **1:30 PM IST** (B+ grade = earliest exit on Friday) |

> **This is the most conditional setup today. If MCX crude doesn't clearly hold below ₹8,200 at open — skip this trade entirely.**

Grade: **B+** | Direction: **CALL** | Sector: Consumer Goods

---

### #6 — NIFTY 23000 PE — Grade B *(Mandatory Weekend Hedge)*
**Index:** NIFTY 50 | **Lot:** 25 | **Est. premium:** ₹70–90/unit → ~₹1,750–2,250 per lot
**Capital test:** N/A — insurance cost ~1% of capital per lot ✓

**Why this MUST be bought today:**
Trump says Iran deal signing "could happen this weekend in Europe." This creates a binary weekend event:
- Deal signs successfully → Monday Nifty gaps UP further (hedges expire worthless — acceptable loss)
- Deal collapses → Monday Nifty gaps DOWN 300–500 pts, crude spikes, all CALL positions explode

The PUT costs ~₹1,750. The protection is worth 10× that if the deal collapses. On a Friday with this risk profile, skipping the hedge is gambling.

**Setup:** NIFTY 23000 PE June 25 Monthly *(NOT June 18 — too near-term)*
**Entry:** Buy at open or within first 15 min. Do not wait.

| Level | Scenario |
|-------|---------|
| Exit T1 | Nifty falls to 22,800 → premium ~₹250 → sell (3.5x return) |
| Time stop | If Nifty holds above **23,250** by 12:30 PM → exit hedge (intraday protection no longer needed) |
| Weekend carry | **CARRY OVER WEEKEND** if you still hold other CALL positions overnight. This is the only position where overnight carry is acceptable today. |
| SL | Premium falls to ₹30 (Nifty proves the hedge wrong) → exit |
| R:R | 250:70 = **3.5:1** if Nifty falls to 22,800 |

Grade: **B** | Direction: **PUT** | Sector: Index Hedge

---

## SECTOR DIVERSIFICATION SUMMARY

| # | Sector | Stock | Grade | Direction |
|---|--------|-------|-------|-----------|
| 1 | IT/Technology | TCS | A+ | CALL |
| 2 | Defence/Aerospace | HAL | A | CALL |
| 3 | Energy/OMC | BPCL | A | CALL |
| 4 | Automobile | TATAMOTORS | B+ | CALL |
| 5 | Consumer Goods | ASIANPAINT | B+ | CALL |
| 6 | Index Hedge | NIFTY PUT | B | PUT |

✅ 5 unique sectors represented (≥4 required)
✅ No sector has >1 stock
✅ CALLs and PUT mix (not all directional)
✅ 1 mandatory macro hedge (HAL — Defence)
✅ 1 tail-risk hedge (NIFTY PUT)

---

## STOCKS TO ACTIVELY AVOID TODAY

| Stock/Category | Reason |
|----------------|--------|
| **ONGC / OIL India** | Upstream oil producers. Crude fell 4% = direct revenue headwind. These are PUT candidates conceptually, not CALL. Avoid entirely. |
| **F&O Ban list** (CDSL, Chambal Fert, IREN, RBL Bank, Titagarh, Hindustan Copper, IEX, Aditya Birla Fashion) | MWPL >95%. No new positions in any direction. |
| **RELIANCE** | Mixed crude signals — refining margins improve but petchem complex, retail is unrelated, telecom separate. No clean single catalyst. Not a clean enough setup. |
| **ICICIBANK / HDFCBANK** | Capital test fails: ICICIBANK lot 700 × ₹1,317 × 3% = ₹27,671 >> ₹15,000. RBI hawkish CPI (5.1%) = rate cut narrative dead. Skip banks entirely this session. |
| **IT stocks beyond TCS** | AI disruption concern is sector-wide. One IT position (TCS) is sufficient exposure. Infosys/HCL/Wipro have same sector risk but less clean setup. |
| **June 18 weekly options (any stock)** | Only 6 days to expiry. High gamma, extreme theta burn. Use June 25 monthly only. |
| **INOX India** | Not F&O listed. +21.6% in 5 days (SpaceX link) = overbought. No direct trade vehicle. |
| **Defensive FMCG stocks** | No specific catalysts today. Market in risk-on mode. Don't dilute with defensives when regime is MILD_BULL. |

---

## FINAL VERDICT

| Parameter | Value |
|-----------|-------|
| Today's bias | **CAUTIOUSLY BULLISH** |
| Confidence | **6/10 — Medium** (Iran fragility, IAEA risk, Friday) |
| Crude regime | **MILD_BULL** — but ₹246 from CAUTION flip |
| VIX sizing | **HALF size on ALL positions** |
| Entry permission | **🟡 YELLOW — 30–45 min wait rule in effect** |
| Expected Nifty range | **23,100 – 23,500** |
| Key events to watch | Iran news flow; MCX crude open level; SpaceX IPO intraday Nasdaq; FII 1-hr provisional data |

**Bull case (30%):** Gap-up sustained above 23,400 through first hour. Iran deal news reinforced. MCX crude opens at ₹8,100–8,200 (clearly MILD_BULL). DII buying accelerates. Nifty makes run at 23,500. TCS, HAL, BPCL all achieve T1. IT sector recovers 2%+.

**Bear case (25%):** IAEA non-compliance news dominates morning. Iran defence minister's "US bases in range" statement gets amplified. Crude bounces to $89+ (MCX ₹8,500+). Regime flips CAUTION. Gap-fill below 23,200. NIFTY PUT pays out. Exit all CALLs.

**Most likely scenario (45%):** Gap-up to 23,350–23,380, followed by volatility. Sells off to 23,250–23,280 (partial gap-fill) around 10:00–10:30 AM. Stabilises. TCS and HAL work. BPCL triggers cleanly if crude holds. B+ setups (TATAMOTORS, ASIANPAINT) either don't trigger or trigger late. Session closes 23,280–23,420 range. Mild positive day, not a runaway rally.

---

> ## ⭐ THE ONE NUMBER THAT MATTERS: **23,250**
>
> If Nifty holds **ABOVE 23,250** after the first 30 minutes → CALL thesis validated → proceed with entries per triggers above.
>
> If Nifty closes a 15-min candle **BELOW 23,250** → EXIT all open CALL positions immediately, go flat (keep NIFTY PUT hedge only). The gap-up thesis is invalidated.
>
> 23,250 is the gap-fill zone. Below it, the overnight bullishness is entirely reversed.

---

## NOTE ON PAPER TRADING TODAY

**What the system will see:**
- Multiple gap-up entry attempts at open. System should respect 30-min wait rule.
- MCX crude pre-open data (approximately 9:00–9:15 AM IST) — this is the critical regime signal before entries.
- VIX at 15.63 → system configured for HALF size on all positions.
- All B+ positions have time-stops earlier than A/A+ positions.

**Specific things to watch in logs:**
1. Did MCX crude open above or below ₹8,400? (Regime confirmation)
2. Did FII turn net buyer in the first provisional data update (~11 AM)?
3. Did Nifty IT index hold the gap-up or reverse by 10:30 AM?
4. Did any Iran headline (negative) emerge between 9:15–10:30 AM?
5. NIFTY PUT — did it get filled at ₹70–90 at open?

**Friday-specific rules applied today:**
- ALL positions exited by 3:10 PM IST maximum
- B+ positions: 1:30–2:00 PM exit (do not hold into last hour)
- NIFTY PUT only: May be carried over weekend if other CALLs also held overnight
- Preferred outcome: All positions flat by 3:00 PM with profits banked

**Running zero-trade days (system):** 0 (first tracked session — establishing baseline)

---

## ONE-LINE SUMMARY

Nifty gaps up ~200 pts (Iran de-escalation, Nasdaq +2.5%), crude regime flips to MILD_BULL ₹8,254 but sits ₹246 from CAUTION; VIX 15.63 → half-size; FRIDAY + fresh IAEA non-compliance = 🟡 YELLOW entry with 45-min wait. Trade TCS (A+), HAL (A), BPCL (A) on confirmation; B+ only if clean; buy 1-lot NIFTY PUT at open as weekend hedge; hard stop at **23,250**; ALL out by 3:10 PM.

---

*Brief generated: June 12, 2026 | Capital: ₹2,00,000 | Mode: Paper + Manual Observation*
*Sources: CNBC, TradingEconomics, TheStreet, Investing.com, NSE India, 5paisa, NiftyTrader, Goodreturns, BusinessToday, HDFCSky, Times of Israel, Business Standard, Republic World*
