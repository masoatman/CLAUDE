// Builds the iShares ETF deep-research deck.
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const OUT = process.argv[2] || "iShares_ETF_Deep_Research.pptx";

// Palette: ink navy dominant, teal secondary, amber accent
const C = {
  ink: "0B1F33", ink2: "16324F", teal: "0E8C8C", tealLt: "D5EEEE",
  amber: "E8A33D", red: "C0392B", bg: "FFFFFF", mist: "F1F5F8",
  text: "1D2733", muted: "5E6B78", white: "FFFFFF", line: "D6DEE6",
};
const H = "Cambria", B = "Calibri";

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.title = "iShares ETFs: Deep Research";
  const W = 13.333;

  const title = (s, t, sub) => {
    s.addText(t, { x: 0.6, y: 0.4, w: W - 1.2, h: 0.8, fontFace: H, fontSize: 32, bold: true, color: C.ink, margin: 0, isTextBox: true });
    if (sub) s.addText(sub, { x: 0.6, y: 1.15, w: W - 1.2, h: 0.45, fontFace: B, fontSize: 15, italic: true, color: C.teal, margin: 0, isTextBox: true });
  };
  const foot = (s, src, n) => {
    s.addText(src, { x: 0.6, y: 6.95, w: W - 2.2, h: 0.35, fontFace: B, fontSize: 9, color: C.muted, margin: 0, isTextBox: true });
    s.addText(String(n), { x: W - 1.1, y: 6.95, w: 0.5, h: 0.35, fontFace: B, fontSize: 10, color: C.muted, align: "right", margin: 0, isTextBox: true });
  };
  const card = (s, x, y, w, h, fill = C.mist) =>
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.12 });
  const iconDot = async (s, Comp, x, y, d = 0.6, bg = C.teal) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg } });
    s.addImage({ data: await icon(Comp, "FFFFFF"), x: x + d * 0.22, y: y + d * 0.22, w: d * 0.56, h: d * 0.56 });
  };
  const chartBase = () => ({
    catAxisLabelColor: C.muted, valAxisLabelColor: C.muted, catAxisLabelFontFace: B, valAxisLabelFontFace: B,
    catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, valGridLine: { color: "E3E8ED", size: 0.5 },
    catGridLine: { style: "none" }, showValue: true, dataLabelFontFace: B, dataLabelFontSize: 11,
    dataLabelColor: C.text, titleFontFace: H, titleColor: C.ink, titleFontSize: 14,
  });

  // 1. Title
  {
    const s = pres.addSlide(); s.background = { color: C.ink };
    s.addImage({ data: await icon(fa.FaLayerGroup, C.teal), x: 0.8, y: 1.2, w: 0.9, h: 0.9 });
    s.addText("iShares ETFs", { x: 0.8, y: 2.3, w: 11, h: 1.1, fontFace: H, fontSize: 54, bold: true, color: C.white, margin: 0, isTextBox: true });
    s.addText("Deep research: franchise economics, product architecture, competitive position and portfolio use",
      { x: 0.8, y: 3.4, w: 10.5, h: 0.9, fontFace: B, fontSize: 20, color: "C9D6E3", margin: 0, isTextBox: true });
    s.addText("Institutional equity & fund research  |  September 2026", { x: 0.8, y: 5.6, w: 10, h: 0.4, fontFace: B, fontSize: 14, color: C.amber, margin: 0, isTextBox: true });
    s.addText("Independent analysis. Not produced by or affiliated with BlackRock or iShares. Not investment advice.",
      { x: 0.8, y: 6.2, w: 11, h: 0.4, fontFace: B, fontSize: 11, italic: true, color: "8FA3B8", margin: 0, isTextBox: true });
    s.addNotes("Deck built from BlackRock Q1/Q2 2026 results coverage, iShares H1 2026 flows commentary, and ETF data aggregators (Sep 2026). Fund-level AUM figures are approximate and should be verified before external use.");
  }

  // 2. Executive summary
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Executive summary", "The toll road of passive investing is still widening, but no longer uncontested");
    card(s, 0.6, 1.9, 3.6, 4.8, C.ink);
    s.addText("$6T+", { x: 0.8, y: 2.3, w: 3.2, h: 1.2, fontFace: H, fontSize: 66, bold: true, color: C.amber, margin: 0, isTextBox: true });
    s.addText("iShares global AUM, crossed in Q2 2026, roughly doubling in three years", { x: 0.8, y: 3.6, w: 3.2, h: 1.0, fontFace: B, fontSize: 15, color: C.white, margin: 0, isTextBox: true });
    s.addText("$178B", { x: 0.8, y: 4.8, w: 3.2, h: 0.8, fontFace: H, fontSize: 40, bold: true, color: C.white, margin: 0, isTextBox: true });
    s.addText("BlackRock ETF net inflows in Q2 2026 alone", { x: 0.8, y: 5.6, w: 3.2, h: 0.7, fontFace: B, fontSize: 14, color: "C9D6E3", margin: 0, isTextBox: true });

    const pts = [
      [fa.FaChartLine, "Scale compounding", "Record flows (Q1 $132B, Q2 $178B) and market beta lifted BlackRock AUM to $15.3T; iShares is the engine."],
      [fa.FaBalanceScale, "Two-tier pricing works", "Core funds at 3–9 bps win allocators; legacy 'Precision' funds keep premium fees because traders pay for liquidity."],
      [fa.FaChessKnight, "Share pressure is real", "Vanguard has overtaken iShares in US ETF assets; VOO ($1T) now outsizes IVV ($818B)."],
      [fa.FaRocket, "Next legs: fixed income, active, digital, international", "Bond ETFs, active ETFs (>$70B 12-month flows), IBIT, and a $1.5T European UCITS franchise diversify growth."],
    ];
    let y = 1.95;
    for (const [ic, h, b] of pts) {
      await iconDot(s, ic, 4.6, y + 0.05, 0.6);
      s.addText(h, { x: 5.4, y, w: 7.3, h: 0.4, fontFace: B, fontSize: 17, bold: true, color: C.ink, margin: 0, isTextBox: true });
      s.addText(b, { x: 5.4, y: y + 0.42, w: 7.3, h: 0.7, fontFace: B, fontSize: 14, color: C.text, margin: 0, isTextBox: true });
      y += 1.2;
    }
    foot(s, "Sources: Markets Media; CNBC (Apr 14, 2026); Investing.com Q2 2026 slides; ETFdb.", 2);
  }

  // 3. Franchise at a glance
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "The franchise at a glance", "iShares is the largest single profit pool inside the world's largest asset manager");
    const stats = [
      ["$15.34T", "BlackRock total AUM, Q2 2026 (vs $12.53T a year earlier)"],
      ["$6T+", "iShares ETF AUM globally, Q2 2026"],
      ["$1.5T", "iShares Europe AUM; ~$80B raised YTD by mid-2026"],
      ["$100B+", "Locally domiciled iShares AUM in Asia Pacific"],
      ["45.9%", "BlackRock operating margin peak cited in Q2 2026 results"],
      ["14%", "Organic growth rate cited for ETFs, Q2 2026"],
    ];
    stats.forEach(([n, l], i) => {
      const col = i % 3, row = Math.floor(i / 3);
      const x = 0.6 + col * 4.1, y = 1.95 + row * 2.45;
      card(s, x, y, 3.8, 2.15, row === 0 ? C.tealLt : C.mist);
      s.addText(n, { x: x + 0.3, y: y + 0.25, w: 3.2, h: 0.95, fontFace: H, fontSize: 44, bold: true, color: row === 0 ? C.teal : C.ink, margin: 0, isTextBox: true });
      s.addText(l, { x: x + 0.3, y: y + 1.2, w: 3.2, h: 0.8, fontFace: B, fontSize: 14, color: C.text, margin: 0, isTextBox: true });
    });
    foot(s, "Sources: Global Business Outlook; Markets Media; Investing.com Q2 2026 slides. Figures as reported in press coverage of BlackRock results.", 3);
  }

  // 4. Growth trajectory
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Growth trajectory: flows plus beta", "BlackRock AUM grew ~22% year on year; ETF inflows accelerated quarter over quarter");
    s.addChart(pres.charts.BAR, [{ name: "BlackRock AUM ($T)", labels: ["Q2 2025", "Q1 2026", "Q2 2026"], values: [12.53, 13.89, 15.34] }], {
      x: 0.6, y: 1.8, w: 6.0, h: 4.9, barDir: "col", ...chartBase(), chartColors: [C.ink], showLegend: false,
      showTitle: true, title: "BlackRock total AUM ($ trillions)", dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00", valAxisMinVal: 0, barGapWidthPct: 60,
    });
    s.addChart(pres.charts.BAR, [{ name: "ETF net inflows ($B)", labels: ["Q1 2026", "Q2 2026"], values: [132, 178] }], {
      x: 6.9, y: 1.8, w: 3.4, h: 4.9, barDir: "col", ...chartBase(), chartColors: [C.teal], showLegend: false,
      showTitle: true, title: "ETF net inflows ($B)", dataLabelPosition: "outEnd", valAxisMinVal: 0, barGapWidthPct: 50,
    });
    card(s, 10.6, 1.9, 2.15, 4.7, C.mist);
    s.addText([
      { text: "Read-across", options: { bold: true, fontSize: 15, color: C.ink, breakLine: true } },
      { text: "Flow acceleration during a rising market is the best case for a fee-on-AUM model: revenue scales with both.", options: { fontSize: 13, color: C.text, breakLine: true } },
      { text: " ", options: { fontSize: 8, breakLine: true } },
      { text: "Flip side: much of the AUM gain is market beta, which reverses in a drawdown.", options: { fontSize: 13, color: C.red } },
    ], { x: 10.8, y: 2.05, w: 1.8, h: 4.4, fontFace: B, valign: "top", margin: 0, isTextBox: true });
    foot(s, "Sources: Global Business Outlook; BigGo Finance (Q1 2026 call); Investing.com Q2 2026 slides.", 4);
  }

  // 5. Product architecture
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Product architecture: four engines", "A shelf built to serve every buyer, from retail savers to hedge-fund traders");
    const q = [
      [fa.FaCubes, "Core", "Low-cost building blocks for long-term allocators and model portfolios.", "IVV · IEFA · IEMG · IJH · IJR · AGG", C.ink],
      [fa.FaCrosshairs, "Precision / tactical", "Deep-liquidity legacy funds used by institutions to trade, hedge and lend.", "EFA · EEM · IWM · sector & country funds", C.teal],
      [fa.FaUniversity, "Fixed income", "Category leader in bond ETFs, including defined-maturity iBonds.", "AGG · LQD · HYG · TLT · TIP · SGOV · iBonds", C.ink2],
      [fa.FaBitcoin, "Active, thematic & digital", "Newest growth vectors with higher fees per dollar.", "IBIT · ETHA · active ETFs · factor (QUAL, USMV)", C.amber],
    ];
    for (let i = 0; i < 4; i++) {
      const [ic, h, d, t, col] = q[i];
      const x = 0.6 + (i % 2) * 6.15, y = 1.9 + Math.floor(i / 2) * 2.5;
      card(s, x, y, 5.9, 2.25);
      await iconDot(s, ic, x + 0.3, y + 0.3, 0.7, col);
      s.addText(h, { x: x + 1.2, y: y + 0.3, w: 4.4, h: 0.5, fontFace: H, fontSize: 20, bold: true, color: C.ink, margin: 0, isTextBox: true });
      s.addText(d, { x: x + 1.2, y: y + 0.85, w: 4.4, h: 0.75, fontFace: B, fontSize: 14, color: C.text, margin: 0, isTextBox: true });
      s.addText(t, { x: x + 1.2, y: y + 1.6, w: 4.4, h: 0.45, fontFace: B, fontSize: 13, bold: true, color: C.teal, margin: 0, isTextBox: true });
    }
    foot(s, "Tickers are illustrative members of each segment, not a complete list.", 5);
  }

  // 6. Flagship funds table
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Flagship funds", "Ten tickers carry a large share of US-listed iShares assets");
    const hdr = ["Ticker", "Fund", "Exposure", "Expense ratio", "Role in a portfolio"].map(t => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.ink } } }));
    const rows = [
      ["IVV", "Core S&P 500", "US large cap", "0.03%", "Core US equity; ~$818B AUM"],
      ["IEFA", "Core MSCI EAFE", "Developed ex-US", "0.07%", "Core international; ~$191B AUM"],
      ["IEMG", "Core MSCI Emerging Markets", "EM equity", "0.09%", "Core EM sleeve"],
      ["AGG", "Core US Aggregate Bond", "US IG bonds", "0.03%", "Core fixed income ballast"],
      ["IJH / IJR", "Core S&P Mid-Cap / Small-Cap", "US mid & small", "0.05% / 0.06%", "Size diversification"],
      ["IWF / IWM", "Russell 1000 Growth / Russell 2000", "Style & small cap", "0.18% / 0.19%", "Tactical tilts, hedging"],
      ["LQD / HYG", "iBoxx IG / HY Corporate", "Credit", "0.14% / 0.49%", "Liquid credit beta"],
      ["TLT", "20+ Year Treasury", "Long duration", "0.15%", "Rates hedge, duration bets"],
      ["SGOV", "0–3 Month Treasury", "T-bills", "0.09%", "Cash management"],
      ["IBIT", "Bitcoin Trust", "Spot bitcoin", "0.25%", "Digital asset exposure"],
    ];
    const body = rows.map((r, i) => r.map((c, j) => ({ text: c, options: { fill: { color: i % 2 ? C.white : C.mist }, bold: j === 0, color: j === 0 ? C.teal : C.text } })));
    s.addTable([hdr, ...body], {
      x: 0.6, y: 1.85, w: W - 1.2, colW: [1.5, 3.4, 2.1, 1.8, 3.333], fontFace: B, fontSize: 13,
      border: { type: "solid", pt: 0.5, color: C.line }, rowH: 0.43, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1],
    });
    foot(s, "Expense ratios per fund prospectuses (verify current figures on ishares.com). AUM: IVV and IEFA per FinanceCharts, Sep 18, 2026; other AUM figures omitted where sources conflicted.", 6);
  }

  // 7. Fee strategy
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Fee strategy: the two-tier playbook", "Same exposure, very different price, because the buyers pay for different things");
    s.addChart(pres.charts.BAR, [
      { name: "Legacy / Precision", labels: ["Developed ex-US", "Emerging markets", "US small cap"], values: [0.32, 0.70, 0.19] },
      { name: "Core", labels: ["Developed ex-US", "Emerging markets", "US small cap"], values: [0.07, 0.09, 0.06] },
    ], {
      x: 0.6, y: 1.8, w: 7.4, h: 4.9, barDir: "col", ...chartBase(), chartColors: [C.ink, C.teal],
      showLegend: true, legendPos: "b", legendFontFace: B, legendFontSize: 12, showTitle: true,
      title: "Expense ratio (%): EFA vs IEFA · EEM vs IEMG · IWM vs IJR", dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00", valAxisMinVal: 0, barGapWidthPct: 70,
    });
    const pts = [
      ["Why legacy funds keep their fees", "Tight spreads, deep options markets and securities lending make EFA, EEM and IWM the tools traders use. Liquidity is the product; the fee is secondary."],
      ["Why Core undercuts", "Core was launched to defend share against Vanguard and Schwab among advisors and retail buyers, who compare funds on expense ratio."],
      ["The analyst's take", "Price discrimination protects revenue per dollar of AUM while still winning the low-cost buyer. Watch the mix shift toward Core."],
    ];
    let y = 1.9;
    for (const [h, b] of pts) {
      card(s, 8.3, y, 4.45, 1.5);
      s.addText(h, { x: 8.5, y: y + 0.12, w: 4.05, h: 0.4, fontFace: B, fontSize: 15, bold: true, color: C.ink, margin: 0, isTextBox: true });
      s.addText(b, { x: 8.5, y: y + 0.52, w: 4.05, h: 0.9, fontFace: B, fontSize: 12, color: C.text, margin: 0, isTextBox: true });
      y += 1.65;
    }
    foot(s, "Expense ratios per fund prospectuses; IWM (Russell 2000) and IJR (S&P 600) track different small-cap indices. EEM shown at ~0.70%.", 7);
  }

  // 8. Competitive landscape
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Competitive landscape", "Vanguard has taken the US crown; iShares keeps its lead in breadth, bonds and global reach");
    s.addChart(pres.charts.BAR, [{ name: "AUM ($B)", labels: ["VOO (Vanguard)", "IVV (iShares)", "SPY (SPDR)"], values: [1000, 818, 777] }], {
      x: 0.6, y: 1.8, w: 5.8, h: 4.9, barDir: "bar", ...chartBase(), chartColors: [C.muted, C.teal, C.muted],
      showLegend: false, showTitle: true, title: "S&P 500 ETF AUM, Sep 2026 ($B)", dataLabelPosition: "outEnd", valAxisMinVal: 0, valAxisMaxVal: 1200, barGapWidthPct: 50,
    });
    const moats = [
      [fa.FaThList, "Breadth", "The widest shelf in the industry across asset classes and regions."],
      [fa.FaUniversity, "Bond ETF leadership", "First-mover scale in fixed-income ETFs and the defined-maturity iBonds range."],
      [fa.FaExchangeAlt, "Trading liquidity", "Institutions use EFA, EEM, IWM, HYG and LQD to hedge and express views."],
      [fa.FaGlobeEurope, "Global distribution", "$1.5T in Europe, where iShares leads UCITS ETFs; $100B+ in APAC."],
      [fa.FaNetworkWired, "Aladdin & models", "Model portfolios and tech distribution put iShares funds on the default path."],
    ];
    let y = 1.85;
    for (const [ic, h, b] of moats) {
      await iconDot(s, ic, 6.8, y + 0.05, 0.55, C.ink);
      s.addText(h, { x: 7.55, y, w: 5.2, h: 0.36, fontFace: B, fontSize: 15, bold: true, color: C.ink, margin: 0, isTextBox: true });
      s.addText(b, { x: 7.55, y: y + 0.36, w: 5.2, h: 0.55, fontFace: B, fontSize: 13, color: C.text, margin: 0, isTextBox: true });
      y += 0.98;
    }
    foot(s, "Sources: ETFdb / ETF Trends ('Vanguard Overtakes iShares'); Markets Media; FinanceCharts (Sep 2026).", 8);
  }

  // 9. Growth vectors
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Where the next $6T comes from", "Four growth vectors, ranked by our conviction");
    const v = [
      ["1", "Fixed income ETFs", "Bonds are still mostly held and traded outside ETFs. Portfolio trading and model portfolios keep moving bond exposure into ETF wrappers.", "High"],
      ["2", "International / UCITS", "European ETF adoption lags the US by years; $80B raised YTD on a $1.5T base is roughly 11% annualised organic growth.", "High"],
      ["3", "Active ETFs", "More than $70B of net inflows in 12 months, with BlackRock leading industry active flows in 2026. Higher fees help revenue mix.", "Medium-high"],
      ["4", "Digital assets", "IBIT became the dominant spot bitcoin ETF after its 2024 launch. Fees are high, but flows follow crypto sentiment.", "Medium"],
    ];
    v.forEach(([n, h, b, conv], i) => {
      const y = 1.9 + i * 1.22;
      s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.1, w: 0.75, h: 0.75, fill: { color: i < 2 ? C.teal : C.ink2 }, line: { color: i < 2 ? C.teal : C.ink2 } });
      s.addText(n, { x: 0.6, y: y + 0.1, w: 0.75, h: 0.75, fontFace: H, fontSize: 24, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(h, { x: 1.65, y: y + 0.05, w: 3.2, h: 0.85, fontFace: H, fontSize: 17, bold: true, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
      s.addText(b, { x: 5.05, y, w: 5.75, h: 1.0, fontFace: B, fontSize: 13, color: C.text, valign: "middle", margin: 0, isTextBox: true });
      card(s, 11.0, y + 0.2, 1.75, 0.6, i < 2 ? C.tealLt : C.mist);
      s.addText(conv, { x: 11.0, y: y + 0.2, w: 1.75, h: 0.6, fontFace: B, fontSize: 13, bold: true, color: i < 2 ? C.teal : C.ink, align: "center", valign: "middle", margin: 0, isTextBox: true });
    });
    s.addText("Conviction", { x: 11.0, y: 1.55, w: 1.75, h: 0.3, fontFace: B, fontSize: 11, color: C.muted, align: "center", margin: 0, isTextBox: true });
    foot(s, "Sources: Markets Media; ETF Express (Apr 2026); iShares H1 2026 ETF market trends. Conviction ranking is analyst judgement.", 9);
  }

  // 10. Risk matrix
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Key risks", "What could break the compounding story");
    const risks = [
      ["Fee compression", "Core funds already sit at 3–9 bps. Another price war would compress revenue faster than AUM grows.", "High", C.red],
      ["US share loss", "Vanguard leads US ETF AUM and VOO has pulled ahead of IVV. Structural share drift is hard to reverse.", "High", C.red],
      ["Market beta", "AUM, and so revenue, falls with markets. A 20% equity drawdown hits fees immediately.", "Medium", C.amber],
      ["Bond ETF stress", "In a crisis, bond ETFs can trade at discounts to NAV (as in March 2020), inviting regulatory scrutiny.", "Medium", C.amber],
      ["Crypto & reputation", "IBIT's size links flows and headlines to bitcoin volatility.", "Medium", C.amber],
      ["Political / ESG scrutiny", "Stewardship and ESG positions have drawn pushback from US states and European regulators.", "Low-med", C.teal],
    ];
    risks.forEach(([h, b, lvl, col], i) => {
      const x = 0.6 + (i % 2) * 6.15, y = 1.85 + Math.floor(i / 2) * 1.6;
      card(s, x, y, 5.9, 1.4);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 4.55, y: y + 0.2, w: 1.15, h: 0.4, fill: { color: col }, line: { color: col }, rectRadius: 0.2 });
      s.addText(lvl, { x: x + 4.55, y: y + 0.2, w: 1.15, h: 0.4, fontFace: B, fontSize: 12, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(h, { x: x + 0.25, y: y + 0.15, w: 4.2, h: 0.45, fontFace: B, fontSize: 16, bold: true, color: C.ink, margin: 0, isTextBox: true });
      s.addText(b, { x: x + 0.25, y: y + 0.62, w: 5.4, h: 0.7, fontFace: B, fontSize: 12.5, color: C.text, margin: 0, isTextBox: true });
    });
    foot(s, "Severity ratings are analyst judgement.", 10);
  }

  // 11. Illustrative portfolio
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Putting it to work: an all-iShares 60/40", "A diversified core portfolio for about 6 bps a year");
    const alloc = [["IVV", 36], ["IEFA", 14], ["IEMG", 6], ["IJR", 4], ["AGG", 25], ["TIP", 5], ["SGOV", 5], ["IAU", 5]];
    s.addChart(pres.charts.DOUGHNUT, [{ name: "Weight", labels: alloc.map(a => a[0]), values: alloc.map(a => a[1]) }], {
      x: 0.6, y: 1.8, w: 5.4, h: 4.9, holeSize: 55, showPercent: false, showValue: true, showLabel: false,
      chartColors: [C.ink, C.ink2, "3C5A78", "6F8BA6", C.teal, "4DB3B3", "9ED3D3", C.amber],
      dataLabelColor: C.white, dataLabelFontSize: 12, dataLabelFontFace: B, showLegend: true, legendPos: "r", legendFontFace: B, legendFontSize: 12,
    });
    const hdr = ["Sleeve", "Funds", "Weight", "Fee"].map(t => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.ink } } }));
    const rows = [
      ["US equity", "IVV 36% · IJR 4%", "40%", "0.03–0.06%"],
      ["International equity", "IEFA 14% · IEMG 6%", "20%", "0.07–0.09%"],
      ["Core bonds", "AGG", "25%", "0.03%"],
      ["Inflation hedge", "TIP", "5%", "0.18%"],
      ["Cash / T-bills", "SGOV", "5%", "0.09%"],
      ["Real asset", "IAU (gold)", "5%", "0.25%"],
      ["Blended", "8 funds", "100%", "≈0.06%"],
    ];
    const body = rows.map((r, i) => r.map((c, j) => ({ text: c, options: { fill: { color: i === 6 ? C.tealLt : (i % 2 ? C.white : C.mist) }, bold: i === 6 || j === 0, color: C.text } })));
    s.addTable([hdr, ...body], { x: 6.4, y: 1.9, w: 6.35, colW: [1.9, 2.2, 1.0, 1.25], fontFace: B, fontSize: 12.5, rowH: 0.45, border: { type: "solid", pt: 0.5, color: C.line }, valign: "middle", margin: [0.03, 0.08, 0.03, 0.08] });
    s.addText("Illustrative only: not a recommendation. Rebalance annually or at ±5% drift. Swap IVV for a lower-turnover total-market fund if you want to capture small caps more fully.",
      { x: 6.4, y: 5.7, w: 6.35, h: 0.9, fontFace: B, fontSize: 12, italic: true, color: C.muted, margin: 0, isTextBox: true });
    foot(s, "Blended fee = weighted average of prospectus expense ratios (≈6.2 bps). Excludes trading costs and taxes.", 11);
  }

  // 12. Analyst verdict
  {
    const s = pres.addSlide(); s.background = { color: C.ink };
    s.addText("Analyst verdict", { x: 0.6, y: 0.4, w: 8, h: 0.8, fontFace: H, fontSize: 32, bold: true, color: C.white, margin: 0, isTextBox: true });
    s.addText("As a product shelf: best-in-class. As a franchise: dominant, but no longer uncontested.",
      { x: 0.6, y: 1.15, w: 12, h: 0.5, fontFace: B, fontSize: 16, italic: true, color: C.amber, margin: 0, isTextBox: true });
    const cols = [
      ["Scorecard", [["Scale & liquidity", "A"], ["Cost (Core range)", "A"], ["Breadth & innovation", "A"], ["US share momentum", "B−"], ["Fee resilience", "B"]]],
    ];
    card(s, 0.6, 1.95, 3.9, 4.75, C.ink2);
    s.addText("Scorecard", { x: 0.85, y: 2.1, w: 3.4, h: 0.45, fontFace: H, fontSize: 18, bold: true, color: C.white, margin: 0, isTextBox: true });
    cols[0][1].forEach(([k, g], i) => {
      const y = 2.7 + i * 0.75;
      s.addText(k, { x: 0.85, y, w: 2.6, h: 0.55, fontFace: B, fontSize: 14, color: "C9D6E3", valign: "middle", margin: 0, isTextBox: true });
      s.addText(g, { x: 3.5, y, w: 0.75, h: 0.55, fontFace: H, fontSize: 22, bold: true, color: g.startsWith("A") ? "5FD0C7" : C.amber, align: "right", valign: "middle", margin: 0, isTextBox: true });
    });
    const box = (x, h, items) => {
      card(s, x, 1.95, 3.9, 4.75, C.ink2);
      s.addText(h, { x: x + 0.25, y: 2.1, w: 3.4, h: 0.45, fontFace: H, fontSize: 18, bold: true, color: C.white, margin: 0, isTextBox: true });
      s.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1 } })),
        { x: x + 0.25, y: 2.7, w: 3.45, h: 3.8, fontFace: B, fontSize: 13.5, color: "E1E8EF", paraSpaceAfter: 8, valign: "top", margin: 0, isTextBox: true });
    };
    box(4.72, "Signposts to watch", [
      "Quarterly US ETF share against Vanguard",
      "Core versus Precision revenue mix and average fee rate",
      "Fixed income and active ETF flow share",
      "European UCITS organic growth (above 10%?)",
      "IBIT flows through a crypto drawdown",
    ]);
    box(8.84, "What would change our view", [
      "A new fee cut on Core funds below 3 bps",
      "Two or more quarters of net US ETF outflows",
      "A bond ETF liquidity event that prompts regulation",
      "Loss of Europe leadership to a low-cost entrant",
    ]);
    foot(s, "Grades are analyst judgement. Not investment advice.", 12);
  }

  // 13. Sources
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Sources & methodology");
    const src = [
      "Markets Media: \"BlackRock iShares Surpasses $6 Trillion in Assets\" (2026)",
      "CNBC: \"BlackRock quarterly profit rises on active ETFs and performance fees\" (Apr 14, 2026)",
      "Investing.com: \"BlackRock Q2 2026 slides: record flows drive 45.9% margin peak\"",
      "Global Business Outlook: \"BlackRock assets hit record USD 15 trillion amid strong ETF inflows\"",
      "ETF Express: \"iShares by BlackRock reports record net inflows in Q1\" (Apr 14, 2026)",
      "BigGo Finance: BlackRock Q1 2026 earnings call coverage",
      "ETFdb / ETF Trends: \"Vanguard Overtakes iShares as Largest ETF Provider\"",
      "FinanceCharts: Top iShares ETFs / Biggest ETFs (data as of Sep 18, 2026)",
      "iShares: \"H1 2026 ETF & ETP Market Trends\"; fund prospectuses for expense ratios",
    ];
    s.addText(src.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < src.length - 1 } })),
      { x: 0.6, y: 1.4, w: 7.6, h: 5.3, fontFace: B, fontSize: 13, color: C.text, paraSpaceAfter: 6, valign: "top", margin: 0, isTextBox: true });
    card(s, 8.6, 1.5, 4.15, 5.1, C.mist);
    s.addText([
      { text: "Methodology & caveats", options: { bold: true, fontSize: 16, color: C.ink, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Figures are from press coverage of BlackRock results and third-party ETF data as of September 2026. They were not independently audited.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Fund AUM changes daily; confirm on ishares.com before use.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Independent research, not affiliated with BlackRock. Not investment advice." },
    ], { x: 8.85, y: 1.7, w: 3.7, h: 4.7, fontFace: B, fontSize: 13, color: C.text, valign: "top", margin: 0, isTextBox: true });
    foot(s, "", 13);
  }

  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT);
})();
