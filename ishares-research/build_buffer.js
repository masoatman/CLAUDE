// Builds the buffer ETF + alts liquidity deck. Figures come from buffer_liquidity_model.py.
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const OUT = process.argv[2] || "Buffer_ETFs_and_Alts_Liquidity.pptx";

const C = {
  ink: "0B1F33", ink2: "16324F", teal: "0E8C8C", tealLt: "D5EEEE",
  amber: "E8A33D", red: "C0392B", redLt: "F6DEDB", bg: "FFFFFF", mist: "F1F5F8",
  text: "1D2733", muted: "5E6B78", white: "FFFFFF", line: "D6DEE6", grey: "8A97A5",
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
  pres.layout = "LAYOUT_WIDE";
  pres.title = "Buffer ETFs and Alts Liquidity";
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
    catGridLine: { style: "none" }, dataLabelFontFace: B, dataLabelFontSize: 10,
    dataLabelColor: C.text, titleFontFace: H, titleColor: C.ink, titleFontSize: 14,
    legendFontFace: B, legendFontSize: 12,
  });
  const cellTable = (s, hdr, rows, opts, hl = () => false) => {
    const h = hdr.map(t => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.ink } } }));
    const body = rows.map((r, i) => r.map((c, j) => ({ text: c, options: {
      fill: { color: hl(i) ? C.tealLt : (i % 2 ? C.white : C.mist) }, bold: j === 0 || hl(i), color: C.text } })));
    s.addTable([h, ...body], { fontFace: B, border: { type: "solid", pt: 0.5, color: C.line }, valign: "middle",
      margin: [0.03, 0.1, 0.03, 0.1], ...opts });
  };

  // 1. Title
  {
    const s = pres.addSlide(); s.background = { color: C.ink };
    s.addImage({ data: await icon(fa.FaShieldAlt, C.teal), x: 0.8, y: 1.2, w: 0.85, h: 0.85 });
    s.addText("Buffer ETFs & alts: liquidity in a downturn", { x: 0.8, y: 2.2, w: 11.8, h: 1.1, fontFace: H, fontSize: 40, bold: true, color: C.white, margin: 0, isTextBox: true });
    s.addText("How STEN and SMAX behave, and what adding alternatives does to your ability to raise cash when equities fall",
      { x: 0.8, y: 3.5, w: 10.8, h: 0.9, fontFace: B, fontSize: 20, color: "C9D6E3", margin: 0, isTextBox: true });
    s.addText("Portfolio strategy review  |  September 2026", { x: 0.8, y: 5.6, w: 10, h: 0.4, fontFace: B, fontSize: 14, color: C.amber, margin: 0, isTextBox: true });
    s.addText("Independent analysis. Not produced by or affiliated with BlackRock or iShares. Not investment advice.",
      { x: 0.8, y: 6.2, w: 11, h: 0.4, fontFace: B, fontSize: 11, italic: true, color: "8FA3B8", margin: 0, isTextBox: true });
  }

  // 2. Executive summary
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Executive summary", "Alts rarely hurt through losses. They hurt by being unsellable when you need cash");
    card(s, 0.6, 1.9, 3.6, 4.8, C.ink);
    s.addText("6.5%", { x: 0.8, y: 2.25, w: 3.2, h: 1.2, fontFace: H, fontSize: 66, bold: true, color: C.amber, margin: 0, isTextBox: true });
    s.addText("Liquid reserve left in a 60/40 with 25% semi-liquid alts, after a 30% crash, a 10% withdrawal and one rebalance",
      { x: 0.8, y: 3.5, w: 3.2, h: 1.3, fontFace: B, fontSize: 14, color: C.white, margin: 0, isTextBox: true });
    s.addText("vs 40%", { x: 0.8, y: 4.95, w: 3.2, h: 0.8, fontFace: H, fontSize: 40, bold: true, color: C.white, margin: 0, isTextBox: true });
    s.addText("in the same portfolio with no alts", { x: 0.8, y: 5.75, w: 3.2, h: 0.6, fontFace: B, fontSize: 14, color: "C9D6E3", margin: 0, isTextBox: true });
    const pts = [
      [fa.FaShieldAlt, "SMAX works as a liquidity reserve", "About 100% downside buffer; worth roughly −2% to −3% if sold during a 20–30% selloff. It trades daily."],
      [fa.FaChartLine, "STEN is still equity", "Buffers the first 10% only. It is worth about −22% if sold mid-period in a 30% crash. Use it to replace IVV, not bonds."],
      [fa.FaLock, "Semi-liquid alts shrink your reserve", "In Q2 2026, investors asked to redeem 12.4% of NAV and only 38% was paid. Every point moved out of bonds is a point you can't sell."],
      [fa.FaCalendarCheck, "Buy at the Oct 1 reset", "Both funds reset this week. Higher T-bill yields point to a higher SMAX cap (model: about 9% before fees)."],
    ];
    let y = 1.95;
    for (const [ic, h, b] of pts) {
      await iconDot(s, ic, 4.6, y + 0.05, 0.6);
      s.addText(h, { x: 5.4, y, w: 7.3, h: 0.4, fontFace: B, fontSize: 17, bold: true, color: C.ink, margin: 0, isTextBox: true });
      s.addText(b, { x: 5.4, y: y + 0.42, w: 7.3, h: 0.7, fontFace: B, fontSize: 14, color: C.text, margin: 0, isTextBox: true });
      y += 1.2;
    }
    foot(s, "Sources: iShares fund pages and SEC filings; Robert A. Stanger & Co. via Ferrante Capital; model in buffer_liquidity_model.py.", 2);
  }

  // 3. How the funds work
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Two funds, two different jobs", "Both hold IVV plus options and reset every October 1. The protection only applies if held for the full year");
    const funds = [
      ["SMAX", "iShares Large Cap Max Buffer Sep ETF", C.teal, [
        ["Downside buffer", "About 100% of S&P losses"],
        ["Cap, period ending Sep 30, 2026", "7.80% gross · 7.30% net"],
        ["Fee", "0.50%"],
        ["Worst case, held to end", "About −0.5% (the fee)"],
        ["Behaves like", "A 1-year Treasury plus capped S&P upside"],
        ["Job in the portfolio", "Liquidity reserve / partial bond substitute"],
      ]],
      ["STEN", "iShares Large Cap 10% Target Buffer Sep ETF", C.ink, [
        ["Downside buffer", "First 10% of S&P losses"],
        ["Cap, period ending Sep 30, 2026", "17.63% gross · 17.13% net"],
        ["Fee", "0.50%"],
        ["Worst case, held to end", "S&P loss minus 10% (−30% → about −20%)"],
        ["Behaves like", "S&P exposure with a smaller loss in a mild fall"],
        ["Job in the portfolio", "Lower-drawdown equity substitute"],
      ]],
    ];
    funds.forEach(([tk, name, col, rows], i) => {
      const x = 0.6 + i * 6.15;
      card(s, x, 1.9, 5.9, 4.8);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.3, y: 2.1, w: 1.5, h: 0.6, fill: { color: col }, line: { color: col }, rectRadius: 0.1 });
      s.addText(tk, { x: x + 0.3, y: 2.1, w: 1.5, h: 0.6, fontFace: H, fontSize: 22, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(name, { x: x + 2.0, y: 2.1, w: 3.7, h: 0.6, fontFace: B, fontSize: 13, italic: true, color: C.muted, valign: "middle", margin: 0, isTextBox: true });
      rows.forEach(([k, v], j) => {
        const y = 2.95 + j * 0.6;
        s.addText(k, { x: x + 0.3, y, w: 2.3, h: 0.55, fontFace: B, fontSize: 12.5, color: C.muted, valign: "middle", margin: 0, isTextBox: true });
        s.addText(v, { x: x + 2.65, y, w: 3.05, h: 0.55, fontFace: B, fontSize: 13.5, bold: j === 5, color: j === 5 ? col : C.text, valign: "middle", margin: 0, isTextBox: true });
      });
    });
    foot(s, "Both track IVV's price return only: most dividends (~1.2%/yr) fund the options. SMAX distribution yield 0.76% (Jun 2026). Sources: iShares; SEC Form 497 (2025).", 3);
  }

  // 4. End-of-period payoff
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Payoff if held for the full outcome period", "SMAX trades almost all upside for near-certainty; STEN keeps more upside but still takes big losses");
    const xs = []; for (let r = -40; r <= 30; r += 5) xs.push(r);
    const smax = xs.map(r => +(Math.min(Math.max(r, 0), 7.8) - 0.5).toFixed(2));
    const sten = xs.map(r => +((r >= 0 ? Math.min(r, 17.63) : r >= -10 ? 0 : r + 10) - 0.5).toFixed(2));
    const labels = xs.map(r => (r > 0 ? "+" : "") + r + "%");
    s.addChart(pres.charts.LINE, [
      { name: "S&P 500 price return", labels, values: xs },
      { name: "STEN (net of fee)", labels, values: sten },
      { name: "SMAX (net of fee)", labels, values: smax },
    ], {
      x: 0.6, y: 1.8, w: 8.2, h: 4.95, ...chartBase(), chartColors: [C.grey, C.ink, C.teal], lineSize: 2.5, lineDataSymbol: "none",
      showLegend: true, legendPos: "b", showTitle: true, title: "Fund return vs S&P 500 price return over the year (%)",
      catAxisTitle: "S&P 500 price return over the outcome period", showCatAxisTitle: true, catAxisTitleFontSize: 11, catAxisTitleColor: C.muted,
      valAxisMinVal: -40, valAxisMaxVal: 30, valAxisMajorUnit: 10, valAxisLabelFormatCode: "0",
    });
    const pts = [
      ["S&P −30%", "SMAX −0.5%  ·  STEN −20.5%"],
      ["S&P −8%", "SMAX −0.5%  ·  STEN −0.5%"],
      ["S&P +25%", "SMAX +7.3%  ·  STEN +17.1%"],
    ];
    s.addText("Reading the chart", { x: 9.2, y: 1.95, w: 3.55, h: 0.4, fontFace: H, fontSize: 17, bold: true, color: C.ink, margin: 0, isTextBox: true });
    pts.forEach(([k, v], i) => {
      const y = 2.5 + i * 1.05;
      card(s, 9.2, y, 3.55, 0.9);
      s.addText(k, { x: 9.4, y: y + 0.08, w: 3.2, h: 0.35, fontFace: B, fontSize: 14, bold: true, color: C.teal, margin: 0, isTextBox: true });
      s.addText(v, { x: 9.4, y: y + 0.45, w: 3.2, h: 0.35, fontFace: B, fontSize: 13, color: C.text, margin: 0, isTextBox: true });
    });
    s.addText("Caps shown are for the period ending Sep 30, 2026. New caps are set at the Oct 1 reset.",
      { x: 9.2, y: 5.75, w: 3.55, h: 0.8, fontFace: B, fontSize: 12, italic: true, color: C.muted, margin: 0, isTextBox: true });
    foot(s, "Contractual payoff before dividends, net of the 0.50% fee. Source: fund terms per iShares.", 4);
  }

  // 5. Mid-period sale
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "What you get if you must sell mid-period", "This is what matters for liquidity: you realise today's value, not the end-of-period protection");
    const labels = ["Month 3: −20%,\nvol up, rates down", "Month 6: −30%,\ncrash conditions", "Month 6: −20%,\nrates +1.5pt (2022)", "Month 6: +10%,\ncalm market"];
    s.addChart(pres.charts.BAR, [
      { name: "S&P 500", labels, values: [-20, -30, -20, 10] },
      { name: "STEN", labels, values: [-14.1, -21.7, -13.1, 8.4] },
      { name: "SMAX", labels, values: [-2.3, -1.8, -3.2, 3.6] },
    ], {
      x: 0.6, y: 1.8, w: 8.4, h: 4.95, barDir: "col", ...chartBase(), chartColors: [C.grey, C.ink, C.teal],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", showLegend: true, legendPos: "b",
      showTitle: true, title: "Price return if sold at that point (%)", valAxisMinVal: -35, valAxisMaxVal: 15, valAxisMajorUnit: 5, barGapWidthPct: 60,
    });
    const pts = [
      [fa.FaCheckCircle, C.teal, "SMAX holds value", "Its put is deep in the money, so it barely moves in a selloff. It is mainly sensitive to interest rates."],
      [fa.FaExclamationTriangle, C.amber, "STEN falls about 2/3 as much", "The 10% buffer only fully applies at expiry. Mid-period, time value blunts it."],
      [fa.FaClock, C.ink, "Buy at the reset", "Buying mid-period gives a different buffer and cap from the headline figures."],
    ];
    let y = 1.95;
    for (const [ic, col, h, b] of pts) {
      await iconDot(s, ic, 9.3, y, 0.55, col);
      s.addText(h, { x: 10.0, y, w: 2.75, h: 0.4, fontFace: B, fontSize: 15, bold: true, color: C.ink, margin: 0, isTextBox: true });
      s.addText(b, { x: 10.0, y: y + 0.42, w: 2.75, h: 1.05, fontFace: B, fontSize: 12.5, color: C.text, margin: 0, isTextBox: true });
      y += 1.6;
    }
    foot(s, "Model: Black-Scholes with a simple volatility skew, calibrated to the Oct 2025 caps; after pro-rata fee. Estimates, not quotes.", 5);
  }

  // 6. Reset timing
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "The October 1 reset: timing and new caps", "Rates are higher than a year ago, which should lift SMAX's cap");
    const stats = [
      ["Sep 30", "Current outcome period ends (Wednesday)", C.tealLt, C.teal],
      ["4.51%", "12-month T-bill yield, Sep 25, 2026 (about 3.8% at the last reset)", C.mist, C.ink],
      ["~9%", "Model estimate of SMAX's new cap before fees (was 7.80%)", C.tealLt, C.teal],
      ["Higher", "STEN's new cap direction. Very sensitive to volatility: wait for the published figure", C.mist, C.ink],
    ];
    stats.forEach(([n, l, bg, fg], i) => {
      const x = 0.6 + (i % 2) * 3.35, y = 1.9 + Math.floor(i / 2) * 2.45;
      card(s, x, y, 3.1, 2.2, bg);
      s.addText(n, { x: x + 0.25, y: y + 0.2, w: 2.6, h: 0.9, fontFace: H, fontSize: 40, bold: true, color: fg, margin: 0, isTextBox: true });
      s.addText(l, { x: x + 0.25, y: y + 1.1, w: 2.6, h: 0.95, fontFace: B, fontSize: 13, color: C.text, margin: 0, isTextBox: true });
    });
    card(s, 7.6, 1.9, 5.15, 4.65, C.ink);
    s.addText("Why rates drive SMAX's cap", { x: 7.9, y: 2.1, w: 4.6, h: 0.5, fontFace: H, fontSize: 18, bold: true, color: C.white, margin: 0, isTextBox: true });
    s.addText([
      { text: "A 100% buffer is economically close to a 1-year zero-coupon bond plus a call spread on the S&P.", options: { bullet: true, breakLine: true } },
      { text: "The higher the 1-year yield, the more upside the fund can buy, so the cap rises.", options: { bullet: true, breakLine: true } },
      { text: "Volatility matters much less for SMAX. For STEN it dominates, because STEN is mostly equity.", options: { bullet: true, breakLine: true } },
      { text: "Action: check the new caps on ishares.com on Oct 1 and buy at or close to the reset.", options: { bullet: true, bold: true, color: C.amber } },
    ], { x: 7.9, y: 2.75, w: 4.6, h: 3.6, fontFace: B, fontSize: 14, color: "E1E8EF", paraSpaceAfter: 10, valign: "top", margin: 0, isTextBox: true });
    foot(s, "Sources: Forbes Treasury rates (Sep 25, 2026); BlackRock 'Higher volatility, higher caps'. Cap estimates from buffer_liquidity_model.py.", 6);
  }

  // 7. Liquidity stress test
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Stress test: the liquid reserve after a crash", "30% equity crash, then a 10% withdrawal, then a rebalance back to 60% equity");
    const labels = ["Base 60/40", "+15% semi-liquid alts", "+25% semi-liquid alts", "Buffer sleeve\n(SMAX 10, STEN 10)"];
    s.addChart(pres.charts.BAR, [
      { name: "At the trough", labels, values: [49.6, 31.4, 19.1, 48.8] },
      { name: "After withdrawal + rebalance", labels, values: [40.0, 20.1, 6.5, 40.0] },
    ], {
      x: 0.6, y: 1.8, w: 8.2, h: 4.95, barDir: "col", ...chartBase(), chartColors: ["9ED3D3", C.teal],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", showLegend: true, legendPos: "b",
      showTitle: true, title: "Liquid defensive assets (% of portfolio)", valAxisMinVal: 0, valAxisMaxVal: 60, barGapWidthPct: 60,
    });
    card(s, 9.1, 1.9, 3.65, 2.3, C.redLt);
    s.addText("6.5%", { x: 9.35, y: 2.05, w: 3.2, h: 0.85, fontFace: H, fontSize: 40, bold: true, color: C.red, margin: 0, isTextBox: true });
    s.addText("left at 25% alts. One more bad quarter and you're selling equities at the lows.", { x: 9.35, y: 2.95, w: 3.2, h: 1.1, fontFace: B, fontSize: 13, color: C.text, margin: 0, isTextBox: true });
    card(s, 9.1, 4.4, 3.65, 2.15, C.tealLt);
    s.addText("40%", { x: 9.35, y: 4.55, w: 3.2, h: 0.85, fontFace: H, fontSize: 40, bold: true, color: C.teal, margin: 0, isTextBox: true });
    s.addText("left with the buffer sleeve instead: SMAX counts as liquid.", { x: 9.35, y: 5.45, w: 3.2, h: 0.9, fontFace: B, fontSize: 13, color: C.text, margin: 0, isTextBox: true });
    foot(s, "Liquid defensive = AGG, TIP, SGOV, IAU, SMAX. Alts funded from AGG (and TIP at 25%), marked with a lag (−4%). Illustrative portfolio from slide 11 of the iShares deck.", 7);
  }

  // 8. Denominator effect + evidence
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "The trap: alts grow as everything else shrinks", "Lagged marks and quarterly exits mean the alt weight rises just when you need to sell");
    s.addChart(pres.charts.BAR, [
      { name: "Target weight", labels: ["15% alts", "25% alts"], values: [15, 25] },
      { name: "After crash + withdrawal", labels: ["15% alts", "25% alts"], values: [19.9, 33.5] },
    ], {
      x: 0.6, y: 1.8, w: 5.6, h: 4.95, barDir: "col", ...chartBase(), chartColors: ["9FB2C5", C.ink],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", showLegend: true, legendPos: "b",
      showTitle: true, title: "Alt weight (% of portfolio)", valAxisMinVal: 0, valAxisMaxVal: 40, barGapWidthPct: 50,
    });
    s.addText("2026: what 'semi-liquid' looked like in practice", { x: 6.6, y: 1.9, w: 6.15, h: 0.45, fontFace: H, fontSize: 17, bold: true, color: C.ink, margin: 0, isTextBox: true });
    const ev = [
      ["$20.8B", "requested from the largest semi-liquid private credit funds in Q1 2026; managers met barely half"],
      ["12.4%", "of NAV requested in Q2 2026, the highest on record; only 38% was fulfilled"],
      ["5%/qtr", "typical tender cap. Interval funds must honour it; non-traded BDCs can gate entirely"],
    ];
    ev.forEach(([n, l], i) => {
      const y = 2.5 + i * 1.4;
      card(s, 6.6, y, 6.15, 1.2);
      s.addText(n, { x: 6.8, y, w: 1.9, h: 1.2, fontFace: H, fontSize: 28, bold: true, color: C.red, valign: "middle", margin: 0, isTextBox: true });
      s.addText(l, { x: 8.75, y, w: 3.85, h: 1.2, fontFace: B, fontSize: 13, color: C.text, valign: "middle", margin: 0, isTextBox: true });
    });
    foot(s, "Sources: Bloomberg (Mar 26, 2026); Robert A. Stanger & Co. via Ferrante Capital; ReadTheDrift; CRS IN12674.", 8);
  }

  // 9. Two regimes
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Two kinds of bear market, two different reserves", "With the 10-year at 5.18% and the Fed leaning to hikes, a 2022-style bear is a live risk");
    cellTable(s, ["Holding", "Deflationary crash (2008/2020)", "Inflationary bear (2022)", "Verdict as a reserve"], [
      ["S&P 500", "−30%", "−20%", "Not a reserve"],
      ["AGG (core bonds)", "+4%", "−10%", "Good in a crash, poor in inflation"],
      ["SGOV (T-bills)", "+2%", "+2%", "Best in both"],
      ["SMAX (mid-period)", "−1.8%", "−3.2%", "Solid in both; beats AGG in 2022"],
      ["STEN (mid-period)", "−21.7%", "−13.1%", "Equity, not a reserve"],
      ["Semi-liquid alts (reported)", "−4%", "−2%", "Looks stable, can't be sold"],
    ], { x: 0.6, y: 1.85, w: 8.3, colW: [2.4, 2.1, 1.8, 2.0], fontSize: 13, rowH: 0.62 }, i => i === 2);
    card(s, 9.2, 1.9, 3.55, 4.65, C.ink);
    s.addText("Portfolio return at the trough", { x: 9.45, y: 2.05, w: 3.1, h: 0.7, fontFace: H, fontSize: 16, bold: true, color: C.white, margin: 0, isTextBox: true });
    const pr = [["Base", "−16.6%", "−14.9%"], ["+15% alts", "−17.8%", "−13.7%"], ["+25% alts", "−18.4%", "−13.0%"], ["Buffer", "−16.3%", "−13.5%"]];
    s.addText("Crash    2022", { x: 10.85, y: 2.8, w: 1.75, h: 0.35, fontFace: B, fontSize: 11, color: "8FA3B8", align: "right", margin: 0, isTextBox: true });
    pr.forEach(([k, a, b], i) => {
      const y = 3.2 + i * 0.5;
      s.addText(k, { x: 9.45, y, w: 1.5, h: 0.45, fontFace: B, fontSize: 13, color: "E1E8EF", margin: 0, valign: "middle", isTextBox: true });
      s.addText(`${a}   ${b}`, { x: 10.85, y, w: 1.75, h: 0.45, fontFace: B, fontSize: 13, bold: true, color: C.white, align: "right", valign: "middle", margin: 0, isTextBox: true });
    });
    s.addText("In 2022, alts look better only because their marks lag the market. That is not protection.",
      { x: 9.45, y: 5.3, w: 3.1, h: 1.1, fontFace: B, fontSize: 12.5, italic: true, color: C.amber, margin: 0, isTextBox: true });
    foot(s, "Scenario returns are stress assumptions at a month-6 trough; SMAX/STEN from the option model. Source for rates: Forbes (Sep 25, 2026).", 9);
  }

  // 10. Liquidity waterfall
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "The liquidity waterfall: what to sell, in order", "Fund withdrawals and rebalancing from the top down; never count on the bottom tier");
    const tiers = [
      ["1", "T-bills", "SGOV", "Next day, stable value in any regime", C.teal],
      ["2", "Max buffer", "SMAX", "Next day, small mark-down in a selloff", "3AA6A6"],
      ["3", "Core bonds & real assets", "AGG · TIP · IAU", "Next day; value depends on the regime", C.ink2],
      ["4", "Equities & 10% buffer", "IVV · IEFA · IEMG · STEN", "Next day, but selling means locking in losses", C.grey],
      ["5", "Semi-liquid alts", "Private credit · BDCs", "Quarterly, pro-rated or gated. Treat as unavailable", C.red],
    ];
    tiers.forEach(([n, h, tk, d, col], i) => {
      const y = 1.85 + i * 0.98, x = 0.6 + i * 0.35;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 12.15 - i * 0.35, h: 0.82, fill: { color: i === 4 ? C.redLt : C.mist }, line: { color: i === 4 ? C.redLt : C.mist }, rectRadius: 0.1 });
      s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y + 0.11, w: 0.6, h: 0.6, fill: { color: col }, line: { color: col } });
      s.addText(n, { x: x + 0.15, y: y + 0.11, w: 0.6, h: 0.6, fontFace: H, fontSize: 20, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(h, { x: x + 1.0, y, w: 3.3, h: 0.82, fontFace: B, fontSize: 16, bold: true, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
      s.addText(tk, { x: x + 4.3, y, w: 3.5, h: 0.82, fontFace: B, fontSize: 14, bold: true, color: C.teal, valign: "middle", margin: 0, isTextBox: true });
      s.addText(d, { x: x + 7.8, y, w: 4.2 - i * 0.35, h: 0.82, fontFace: B, fontSize: 13, color: C.text, valign: "middle", margin: 0, isTextBox: true });
    });
    foot(s, "Tickers are examples from the illustrative portfolio.", 10);
  }

  // 11. Recommendations
  {
    const s = pres.addSlide(); s.background = { color: C.ink };
    s.addText("Recommendations", { x: 0.6, y: 0.4, w: 8, h: 0.8, fontFace: H, fontSize: 32, bold: true, color: C.white, margin: 0, isTextBox: true });
    s.addText("Size alts by the liquidity you must keep, not by the return you hope to add",
      { x: 0.6, y: 1.15, w: 12, h: 0.5, fontFace: B, fontSize: 16, italic: true, color: C.amber, margin: 0, isTextBox: true });
    const recs = [
      ["Set a liquidity floor first", "After a 30% crash and a year of withdrawals, keep at least 20% in liquid defensive assets. That caps semi-liquid alts at about 10–15%, funded from AGG."],
      ["Don't count alts as liquidity", "Sell in this order: SGOV → SMAX → AGG/TIP → equities. Assume alts are unavailable for about three years."],
      ["Use SMAX to split the bond sleeve", "For example 10% SMAX + 15% AGG. SMAX covers the 2022-style case; AGG still gains in a deflationary crash."],
      ["Use STEN against equities, not bonds", "Swap part of IVV for STEN if you want smaller losses in a moderate fall. It is not dry powder."],
      ["Liquid alts for money you may need", "Daily-dealing alts (liquid-alt ETFs or funds, gold such as IAU) for the accessible part; quarterly vehicles only for 3+ year money."],
    ];
    recs.forEach(([h, b], i) => {
      const col = i < 3 ? 0 : 1, row = i < 3 ? i : i - 3;
      const x = 0.6 + col * 6.15, y = 1.95 + row * 1.6;
      card(s, x, y, 5.9, 1.4, C.ink2);
      s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: y + 0.25, w: 0.55, h: 0.55, fill: { color: C.teal }, line: { color: C.teal } });
      s.addText(String(i + 1), { x: x + 0.25, y: y + 0.25, w: 0.55, h: 0.55, fontFace: H, fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(h, { x: x + 1.0, y: y + 0.15, w: 4.7, h: 0.4, fontFace: B, fontSize: 15, bold: true, color: C.white, margin: 0, isTextBox: true });
      s.addText(b, { x: x + 1.0, y: y + 0.55, w: 4.7, h: 0.8, fontFace: B, fontSize: 12, color: "C9D6E3", margin: 0, isTextBox: true });
    });
    foot(s, "Analysis on the illustrative portfolio, not personal advice.", 11);
  }

  // 12. Methodology & sources
  {
    const s = pres.addSlide(); s.background = { color: C.bg };
    title(s, "Methodology & sources");
    const src = [
      "iShares product pages and fact sheets: SMAX, STEN (caps, fees, structure)",
      "SEC: iShares Trust Form 497 (2025), STEN outcome-period cap",
      "BlackRock: 'Expanding your options for downside protection'; 'Higher volatility, higher caps'",
      "Bloomberg: 'Redemption Requests Surge, Leaving Billions Locked in Private Credit Funds' (Mar 26, 2026)",
      "Ferrante Capital / Robert A. Stanger & Co.: Q2 2026 repurchase data",
      "WealthManagement.com; ReadTheDrift; Congressional Research Service IN12674",
      "Forbes Treasury rates (Sep 25, 2026); Bondsavvy, Sep 2026 Fed dot plot",
    ];
    s.addText(src.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < src.length - 1 } })),
      { x: 0.6, y: 1.4, w: 7.6, h: 5.3, fontFace: B, fontSize: 13, color: C.text, paraSpaceAfter: 8, valign: "top", margin: 0, isTextBox: true });
    card(s, 8.6, 1.5, 4.15, 5.1, C.mist);
    s.addText([
      { text: "Model assumptions", options: { bold: true, fontSize: 16, color: C.ink, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Option values: Black-Scholes, 1.2% dividend yield, simple volatility skew, calibrated to the Oct 2025 caps. Directional, not quotes.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Stress test: illustrative 60/40 from the iShares deck; alts marked with a lag; 38% pro-rata fill.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Code: ishares-research/buffer_liquidity_model.py. Not investment advice." },
    ], { x: 8.85, y: 1.7, w: 3.7, h: 4.7, fontFace: B, fontSize: 13, color: C.text, valign: "top", margin: 0, isTextBox: true });
    foot(s, "", 12);
  }

  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT);
})();
