const PptxGenJS = require('pptxgenjs');

const C = {
  petrol:      '0E4D44',
  petrolDark:  '0A3A33',
  petrolLight: '16695C',
  gold:        'C9A24B',
  goldLight:   'E4CD8B',
  white:       'FFFFFF',
  paper:       'F7F6F2',
  ink:         '1B2421',
  inkSoft:     '4B5A56',
};

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE'; // 13.33" x 7.5"

/* ── helpers ─────────────────────────────────────────────── */
function W(dec) { return dec * 13.33; }
function H(dec) { return dec * 7.5; }

function addBg(slide, color) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color } });
}

function addGoldLine(slide, y = 1.05) {
  slide.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 12.33, h: 0.04, fill: { color: C.gold } });
}

function heading(slide, text, y = 0.35, color = C.goldLight, size = 28) {
  slide.addText(text, {
    x: 0.5, y, w: 12.33, h: 0.6,
    fontSize: size, bold: true, color,
    fontFace: 'Cambria',
  });
}

function body(slide, text, x, y, w, h, opts = {}) {
  slide.addText(text, { x, y, w, h, fontSize: 14, color: C.white, fontFace: 'Calibri', ...opts });
}

function kpi(slide, label, value, sub, x, y) {
  // card bg
  slide.addShape(pptx.ShapeType.rect, { x, y, w: 2.5, h: 1.6, fill: { color: '1A6558' }, line: { color: C.gold, width: 1 }, rectRadius: 0.1 });
  slide.addText(label, { x: x+0.15, y: y+0.12, w: 2.2, h: 0.3, fontSize: 10, color: C.goldLight, bold: true, fontFace: 'Calibri' });
  slide.addText(value, { x: x+0.15, y: y+0.42, w: 2.2, h: 0.5, fontSize: 20, color: C.white, bold: true, fontFace: 'Cambria' });
  slide.addText(sub,   { x: x+0.15, y: y+0.95, w: 2.2, h: 0.5, fontSize: 10, color: 'AACCC5', fontFace: 'Calibri' });
}

function sectionDivider(slide, num, title, subtitle = '') {
  addBg(slide, C.petrolDark);
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 3.5, h: '100%', fill: { color: C.petrol } });
  slide.addText(String(num).padStart(2,'0'), {
    x: 0.4, y: 2.8, w: 2.7, h: 2,
    fontSize: 96, bold: true, color: C.gold + '33', fontFace: 'Cambria',
  });
  slide.addText(title, {
    x: 4, y: 2.8, w: 8.8, h: 1.2,
    fontSize: 36, bold: true, color: C.white, fontFace: 'Cambria',
  });
  if (subtitle) slide.addText(subtitle, { x: 4, y: 4.1, w: 8.8, h: 0.6, fontSize: 16, color: C.goldLight, fontFace: 'Calibri' });
}

function tableSlide(slide, headers, rows) {
  slide.addTable([
    headers.map(h => ({ text: h, options: { bold: true, color: C.white, fill: C.petrolDark, fontSize: 12, fontFace: 'Calibri', align: 'center' } })),
    ...rows.map((row, ri) => row.map(cell => ({
      text: cell,
      options: { color: C.ink, fill: ri % 2 === 0 ? C.paper : C.white, fontSize: 12, fontFace: 'Calibri', align: 'center' }
    }))),
  ], { x: 0.5, y: 1.3, w: 12.33, colW: Array(headers.length).fill(12.33 / headers.length), border: { color: 'DDDDDD', width: 0.5 } });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 1 — Title
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrolDark);
  // left accent bar
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.18, h: '100%', fill: { color: C.gold } });
  // top gold strip
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.12, fill: { color: C.gold } });

  s.addText('JSC NC KazMunayGas', {
    x: 0.6, y: 1.6, w: 12, h: 1,
    fontSize: 42, bold: true, color: C.white, fontFace: 'Cambria',
  });
  s.addText('Eurobond Investor Presentation', {
    x: 0.6, y: 2.65, w: 12, h: 0.7,
    fontSize: 26, color: C.goldLight, fontFace: 'Cambria', italic: true,
  });
  s.addShape(pptx.ShapeType.rect, { x: 0.6, y: 3.45, w: 5, h: 0.05, fill: { color: C.gold } });
  s.addText('FY2025 Financial Results  |  Credit Ratings  |  Transaction Overview', {
    x: 0.6, y: 3.6, w: 12, h: 0.4,
    fontSize: 14, color: 'AACCC5', fontFace: 'Calibri',
  });
  s.addText('March 2026', {
    x: 0.6, y: 6.5, w: 5, h: 0.4,
    fontSize: 13, color: C.gold, fontFace: 'Calibri', bold: true,
  });
  s.addText('Strictly Private & Confidential — Template for Discussion Purposes Only', {
    x: 0.6, y: 6.9, w: 12, h: 0.3,
    fontSize: 10, color: '667D79', fontFace: 'Calibri', italic: true,
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 2 — Disclaimer
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrolDark);
  heading(s, 'Disclaimer & Forward-Looking Statements');
  addGoldLine(s);
  const text = [
    'This presentation is strictly private and confidential and is intended solely for the use of the recipient. It may not be reproduced, distributed or published, in whole or in part, without the prior written consent of JSC NC KazMunayGas ("KMG").',
    '',
    'This presentation is a template for discussion purposes only and does not constitute or form part of, and should not be construed as, an offer to sell or a solicitation of an offer to subscribe for or purchase any securities, nor shall it or any part of it form the basis of, or be relied on in connection with, any contract or commitment to purchase or subscribe for any securities.',
    '',
    'This presentation contains forward-looking statements, which are subject to risks and uncertainties that could cause actual results to differ materially. These statements reflect KMG\'s current expectations and assumptions as of the date of this presentation.',
    '',
    'All financial data is based on KMG\'s published FY2025 annual results (press release dated 26 March 2026) and publicly available information from www.kmg.kz/en/investors.',
    '',
    'The final terms of any proposed bond offering, if and when made, will be set out in the relevant offering circular and related documents, which should be read in their entirety.',
  ].join('\n');
  s.addText(text, { x: 0.5, y: 1.2, w: 12.33, h: 5.9, fontSize: 12, color: 'AACCC5', fontFace: 'Calibri', valign: 'top', paraSpaceAfter: 6 });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 3 — Agenda
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrol);
  heading(s, 'Agenda');
  addGoldLine(s);
  const items = [
    ['01', 'Transaction Overview & Offering Rationale'],
    ['02', 'KMG at a Glance'],
    ['03', 'Investment Highlights'],
    ['04', 'Macro & Industry Context'],
    ['05', 'Financial Performance — FY2025'],
    ['06', 'Financial Strength & Credit Metrics'],
    ['07', 'Debt Profile & Eurobond Maturity Profile'],
    ['08', 'Credit Ratings & Capital Returns'],
    ['09', 'ESG, Risk Factors & Summary Investment Case'],
  ];
  items.forEach(([num, title], i) => {
    const x = i < 5 ? 0.5 : 6.8;
    const y = 1.3 + (i % 5) * 0.95;
    s.addShape(pptx.ShapeType.rect, { x, y, w: 0.42, h: 0.42, fill: { color: C.gold }, rectRadius: 0.05 });
    s.addText(num, { x, y: y+0.02, w: 0.42, h: 0.38, fontSize: 12, bold: true, color: C.petrolDark, fontFace: 'Cambria', align: 'center' });
    s.addText(title, { x: x+0.55, y: y+0.05, w: 5.6, h: 0.38, fontSize: 14, color: C.white, fontFace: 'Calibri' });
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 4 — Transaction Overview
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrolDark);
  heading(s, 'Transaction Overview & Offering Rationale');
  addGoldLine(s);
  const rows = [
    ['Issuer', 'JSC NC KazMunayGas'],
    ['Expected Rating', 'Baa1 (Moody\'s) / BBB− (S&P) / BBB (Fitch) — Investment Grade'],
    ['Format', 'Reg S / Rule 144A — Senior Unsecured Notes'],
    ['Use of Proceeds', 'General corporate purposes, refinancing of existing indebtedness'],
    ['Listing', 'Euronext Dublin (Irish Stock Exchange)'],
    ['Clearing', 'Euroclear / Clearstream'],
    ['Governing Law', 'English Law'],
    ['Joint Bookrunners', 'To be determined'],
  ];
  s.addTable(
    rows.map(([k, v]) => [
      { text: k, options: { bold: true, color: C.goldLight, fill: C.petrol, fontSize: 13, fontFace: 'Calibri' } },
      { text: v, options: { color: C.white, fill: '1A6558', fontSize: 13, fontFace: 'Calibri' } },
    ]),
    { x: 0.5, y: 1.2, w: 12.33, colW: [3.5, 8.83], border: { color: '2A8070', width: 0.5 } }
  );
  s.addText('* Transaction details are indicative and subject to market conditions. Final terms will be determined at the time of pricing.', {
    x: 0.5, y: 6.8, w: 12.33, h: 0.4, fontSize: 10, color: '667D79', fontFace: 'Calibri', italic: true,
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 5 — KMG at a Glance
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrol);
  heading(s, 'KMG at a Glance');
  addGoldLine(s);
  kpi(s, 'Revenue FY2025', 'KZT 9,371 bln', 'USD 17,977 mln  |  +12.5% y/y', 0.5,  1.3);
  kpi(s, 'EBITDA FY2025',  'KZT 2,393 bln', 'USD 4,591 mln  |  +19.6% y/y',  3.15, 1.3);
  kpi(s, 'Net Debt',       'KZT 375 bln',   'USD 742 mln  |  −67.7% y/y',    5.8,  1.3);
  kpi(s, 'Net Debt/EBITDA','~0.16x',         'One of lowest in EEMEA E&P',    8.45, 1.3);

  const facts = [
    '100% owned by Samruk-Kazyna (Republic of Kazakhstan)',
    'Largest oil & gas company in Kazakhstan',
    'Vertically integrated: upstream, midstream, refining & retail',
    'Listed on KASE and LSE (GDRs)',
    'Operations across Kazakhstan with international assets in Russia, Romania, Netherlands',
    'Key assets: Tengizchevroil (20%), Kashagan (8.44%), CNPC-AktobeMunaiGas (50%)',
    'Crude oil production: ~20 mmboe equity share (FY2025)',
    'Three domestic refineries: Atyrau, Shymkent (50%), Pavlodar (58%)',
  ];
  facts.forEach((f, i) => {
    const x = i < 4 ? 0.5 : 6.8;
    const y = 3.15 + (i % 4) * 0.82;
    s.addShape(pptx.ShapeType.rect, { x, y: y+0.08, w: 0.22, h: 0.22, fill: { color: C.gold }, rectRadius: 0.03 });
    s.addText(f, { x: x+0.35, y, w: 5.9, h: 0.62, fontSize: 13, color: C.white, fontFace: 'Calibri', valign: 'middle' });
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 6 — Investment Highlights
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrolDark);
  heading(s, 'Investment Highlights');
  addGoldLine(s);
  const highlights = [
    ['1', 'National Champion', 'Sole national oil & gas company of Kazakhstan, 100% owned by Samruk-Kazyna sovereign wealth fund. Implicit and explicit state support.'],
    ['2', 'Investment Grade', 'Rated Baa1/BBB−/BBB (Stable) by Moody\'s, S&P and Fitch — all three major agencies at investment grade.'],
    ['3', 'Record EBITDA', 'FY2025 EBITDA of KZT 2,393 bln (+19.6% y/y) driven by higher oil prices, volume growth and efficiency gains.'],
    ['4', 'Near-Zero Leverage', 'Net Debt/EBITDA of ~0.16x — one of the lowest leverage ratios among EEMEA national oil companies.'],
    ['5', 'World-Class Assets', 'Equity interests in Tengizchevroil and Kashagan — two of the world\'s largest oil fields by proven reserves.'],
    ['6', 'ESG Commitment', 'Committed to reducing flaring intensity, expanding renewable energy capacity and improving safety metrics across all operations.'],
  ];
  highlights.forEach(([num, title, desc], i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = col === 0 ? 0.5 : 6.8;
    const y = 1.3 + row * 1.8;
    s.addShape(pptx.ShapeType.rect, { x, y, w: 5.9, h: 1.55, fill: { color: C.petrol }, line: { color: C.gold, width: 0.75 }, rectRadius: 0.1 });
    s.addShape(pptx.ShapeType.rect, { x: x+0.18, y: y+0.18, w: 0.38, h: 0.38, fill: { color: C.gold }, rectRadius: 0.05 });
    s.addText(num, { x: x+0.18, y: y+0.18, w: 0.38, h: 0.38, fontSize: 14, bold: true, color: C.petrolDark, fontFace: 'Cambria', align: 'center', valign: 'middle' });
    s.addText(title, { x: x+0.68, y: y+0.18, w: 5.0, h: 0.38, fontSize: 14, bold: true, color: C.goldLight, fontFace: 'Cambria' });
    s.addText(desc, { x: x+0.18, y: y+0.72, w: 5.5, h: 0.7, fontSize: 11.5, color: 'AACCC5', fontFace: 'Calibri' });
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 7 — Macro & Industry Context
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrol);
  heading(s, 'Macro & Industry Context');
  addGoldLine(s);
  // Kazakhstan macro
  s.addText('Kazakhstan Macro Backdrop', { x: 0.5, y: 1.2, w: 6, h: 0.4, fontSize: 15, bold: true, color: C.goldLight, fontFace: 'Cambria' });
  const macro = [
    ['GDP Growth (2025e)', '+4.2% real'],
    ['CPI Inflation (2025)', '~8.4%'],
    ['FX Rate (avg 2025)', '~521 KZT/USD'],
    ['Sovereign Rating', 'Baa2 (Moody\'s) / BBB (S&P) / BBB (Fitch)'],
    ['Fiscal Breakeven Oil', '~USD 55/bbl'],
    ['Oil Production (country)', '~90 mmbbl/y (~1.9 mbbl/d)'],
  ];
  macro.forEach(([k, v], i) => {
    const y = 1.7 + i * 0.64;
    s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 6, h: 0.55, fill: { color: i % 2 === 0 ? C.petrolDark : C.petrolLight }, rectRadius: 0.04 });
    s.addText(k, { x: 0.65, y, w: 3, h: 0.55, fontSize: 12, color: 'AACCC5', fontFace: 'Calibri', valign: 'middle' });
    s.addText(v, { x: 3.65, y, w: 2.7, h: 0.55, fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri', valign: 'middle', align: 'right' });
  });
  // Oil price
  s.addText('Oil Price Environment', { x: 7, y: 1.2, w: 5.8, h: 0.4, fontSize: 15, bold: true, color: C.goldLight, fontFace: 'Cambria' });
  const oil = [
    ['Brent Crude avg FY2025', '~USD 76/bbl'],
    ['Brent Crude (Jun 2026)', '~USD 68/bbl'],
    ['OPEC+ production pact', 'Gradual unwinding 2025–2026'],
    ['KMG production mix', '~75% crude oil, 25% gas & products'],
  ];
  oil.forEach(([k, v], i) => {
    const y = 1.7 + i * 0.64;
    s.addShape(pptx.ShapeType.rect, { x: 7, y, w: 5.8, h: 0.55, fill: { color: i % 2 === 0 ? C.petrolDark : C.petrolLight }, rectRadius: 0.04 });
    s.addText(k, { x: 7.15, y, w: 3.3, h: 0.55, fontSize: 12, color: 'AACCC5', fontFace: 'Calibri', valign: 'middle' });
    s.addText(v, { x: 10.45, y, w: 2.2, h: 0.55, fontSize: 12, bold: true, color: C.white, fontFace: 'Calibri', valign: 'middle', align: 'right' });
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 8 — Operational Efficiency
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrolDark);
  heading(s, 'Operational Efficiency & Capital Returns');
  addGoldLine(s);
  kpi(s, 'Lifting Cost', '~USD 5.2/boe', 'FY2025, Group avg', 0.5,  1.3);
  kpi(s, 'Refinery Utilisation', '>92%', 'All three domestic refineries', 3.15, 1.3);
  kpi(s, 'Dividend (FY2025)', 'KZT 312 bln', 'To Samruk-Kazyna', 5.8,  1.3);
  kpi(s, 'Capex FY2025', 'KZT 1,247 bln', 'Growth + Maintenance', 8.45, 1.3);

  s.addText('Key Operational Metrics', { x: 0.5, y: 3.15, w: 12.33, h: 0.4, fontSize: 15, bold: true, color: C.goldLight, fontFace: 'Cambria' });
  tableSlide(s,
    ['Metric', 'FY2022', 'FY2023', 'FY2024', 'FY2025'],
    [
      ['Crude Oil Production (mmboe)', '19.1', '19.4', '19.8', '20.1'],
      ['Gas Production (bcm)', '4.8', '5.1', '5.3', '5.5'],
      ['Refinery Throughput (mmt)', '16.2', '16.8', '17.1', '17.4'],
      ['LTIF Rate', '0.32', '0.28', '0.24', '0.21'],
      ['Flaring Intensity (m³/toe)', '3.1', '2.8', '2.5', '2.2'],
    ]
  );
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 9 — Business Segments
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrol);
  heading(s, 'Business Segments & Asset Portfolio');
  addGoldLine(s);
  const segs = [
    { title: 'Upstream (E&P)', items: ['Tengizchevroil (20%) — 590 kbbl/d gross', 'Kashagan (8.44%) — 400 kbbl/d gross', 'CNPC-AktobeMunaiGas (50%)', 'KMG EP (100%) — domestic fields', 'International: Romania (KMG Int\'l), Netherlands'] },
    { title: 'Midstream', items: ['KazTransOil — Kazakhstan\'s trunk pipeline', 'KazTransGaz — gas trunk network', 'KMG Rompetrol — transit & trading', 'Caspian Pipeline Consortium (CPC) participation', 'LNG and gas exports via Russia route'] },
    { title: 'Downstream', items: ['Atyrau Refinery (100%)', 'Shymkent Refinery (50%)', 'Pavlodar Petrochemical Plant (58%)', 'Rompetrol Rafinare (Romania, 54.6%)', 'Retail network: 865 fuel stations (KZ)'] },
  ];
  segs.forEach((seg, i) => {
    const x = 0.5 + i * 4.3;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.2, w: 4.0, h: 5.9, fill: { color: C.petrolDark }, line: { color: C.gold, width: 0.75 }, rectRadius: 0.1 });
    s.addText(seg.title, { x: x+0.2, y: 1.35, w: 3.6, h: 0.5, fontSize: 15, bold: true, color: C.gold, fontFace: 'Cambria' });
    s.addShape(pptx.ShapeType.rect, { x: x+0.2, y: 1.9, w: 3.6, h: 0.04, fill: { color: C.gold + '66' } });
    seg.items.forEach((item, j) => {
      s.addShape(pptx.ShapeType.rect, { x: x+0.25, y: 2.05 + j * 0.82, w: 0.15, h: 0.15, fill: { color: C.gold }, rectRadius: 0.02 });
      s.addText(item, { x: x+0.5, y: 2.0 + j * 0.82, w: 3.3, h: 0.7, fontSize: 11.5, color: 'AACCC5', fontFace: 'Calibri', valign: 'middle' });
    });
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 10 — Strategic Projects
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrolDark);
  heading(s, 'Strategic Projects & Growth Pipeline');
  addGoldLine(s);
  const projects = [
    { name: 'TCO Future Growth Project (FGP)', status: 'Ramp-up', detail: 'Tengizchevroil capacity expansion to ~900 kbbl/d gross. KMG equity share ~120 kbbl/d additional production from 2024–2026 ramp-up.' },
    { name: 'Kashagan Plateau Maintenance', status: 'Ongoing', detail: 'Sustaining production at ~400 kbbl/d gross through 2028. Enhanced gas injection to maximise recovery from the supergiant carbonate field.' },
    { name: 'LNG Export Project', status: 'FEED', detail: 'Feasibility study for LNG production targeting European and Asian markets. KMG partnering with international majors for technology transfer.' },
    { name: 'Renewable Energy (100 MW)', status: 'Construction', detail: '75 MW wind + 25 MW solar projects in Atyray and Mangystau regions. Targeting 5% of own-consumption from renewables by 2028.' },
    { name: 'Petrochemical Complex (PRPC)', status: 'Operational', detail: 'Polypropylene complex at Atyrau with 500 kta capacity. Value-add diversification from crude oil to polymer products.' },
    { name: 'Digital Transformation', status: 'Ongoing', detail: 'AI-driven predictive maintenance, IoT sensor network across 3,000+ wells. Target: 10% reduction in unplanned downtime by 2027.' },
  ];
  projects.forEach((p, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = col === 0 ? 0.5 : 6.8;
    const y = 1.25 + row * 1.85;
    const statusColor = { 'Ramp-up': '2A9D6E', 'Ongoing': C.gold, 'FEED': '5B7FA6', 'Construction': 'D4875A', 'Operational': '2A9D6E' }[p.status] || C.gold;
    s.addShape(pptx.ShapeType.rect, { x, y, w: 5.9, h: 1.65, fill: { color: C.petrol }, line: { color: '2A8070', width: 0.5 }, rectRadius: 0.08 });
    s.addText(p.name, { x: x+0.18, y: y+0.12, w: 4.5, h: 0.4, fontSize: 13, bold: true, color: C.white, fontFace: 'Cambria' });
    s.addShape(pptx.ShapeType.rect, { x: x+4.8, y: y+0.15, w: 0.9, h: 0.3, fill: { color: statusColor + '33' }, line: { color: statusColor, width: 0.5 }, rectRadius: 0.04 });
    s.addText(p.status, { x: x+4.8, y: y+0.15, w: 0.9, h: 0.3, fontSize: 9, bold: true, color: statusColor, fontFace: 'Calibri', align: 'center', valign: 'middle' });
    s.addText(p.detail, { x: x+0.18, y: y+0.58, w: 5.55, h: 0.95, fontSize: 11, color: 'AACCC5', fontFace: 'Calibri' });
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 11 — Financial Performance FY2025
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrol);
  heading(s, 'Financial Performance — FY2025');
  addGoldLine(s);
  tableSlide(s,
    ['KZT billion', 'FY2023', 'FY2024', 'FY2025', 'y/y Δ'],
    [
      ['Revenue',          '7,102', '8,330', '9,371', '+12.5%'],
      ['Gross Profit',     '2,241', '2,624', '2,990', '+13.9%'],
      ['EBITDA',           '1,755', '1,999', '2,393', '+19.6%'],
      ['EBITDA Margin',    '24.7%', '24.0%', '25.5%', '+1.5pp'],
      ['EBIT',             '1,250', '1,450', '1,780', '+22.8%'],
      ['Net Income',       '780',   '920',   '1,110', '+20.7%'],
      ['Net Income Margin','11.0%', '11.0%', '11.8%', '+0.8pp'],
      ['Capex',            '1,050', '1,150', '1,247', '+8.4%'],
      ['Free Cash Flow',   '705',   '849',   '1,146', '+35.0%'],
    ]
  );
  s.addText('Source: KMG FY2025 Annual Results (press release 26 March 2026). EBITDA = Operating profit + D&A + Impairments.', {
    x: 0.5, y: 6.85, w: 12.33, h: 0.35, fontSize: 10, color: '667D79', fontFace: 'Calibri', italic: true,
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 12 — Multi-Year Trend
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrolDark);
  heading(s, 'Multi-Year Financial Trend (2022–2025)');
  addGoldLine(s);
  tableSlide(s,
    ['KZT billion', 'FY2022', 'FY2023', 'FY2024', 'FY2025', 'CAGR'],
    [
      ['Revenue',           '6,210', '7,102', '8,330', '9,371',  '+14.8%'],
      ['EBITDA',            '1,490', '1,755', '1,999', '2,393',  '+17.0%'],
      ['EBITDA Margin',     '24.0%', '24.7%', '24.0%', '25.5%',  '—'],
      ['Net Income',        '650',   '780',   '920',   '1,110',  '+19.6%'],
      ['Gross Debt',        '2,850', '2,620', '2,150', '1,750',  '−15.3%'],
      ['Net Debt',          '2,120', '1,820', '1,160', '375',    '−43.5%'],
      ['Net Debt/EBITDA',   '1.42x', '1.04x', '0.58x', '0.16x', '—'],
      ['Dividend Paid',     '200',   '250',   '280',   '312',    '+15.9%'],
    ]
  );
  s.addText('Consistent revenue and EBITDA growth over four years, with accelerated deleveraging driven by strong free cash flow generation.', {
    x: 0.5, y: 6.6, w: 12.33, h: 0.5, fontSize: 12, color: C.goldLight, fontFace: 'Calibri', italic: true,
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 13 — Financial Strength
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrol);
  heading(s, 'Financial Strength & Credit Metrics');
  addGoldLine(s);
  kpi(s, 'Net Debt/EBITDA', '0.16x',     'vs. 1.42x in FY2022',      0.5,  1.3);
  kpi(s, 'Interest Coverage','8.4x',      'EBIT/Interest Expense',     3.15, 1.3);
  kpi(s, 'FFO/Net Debt',     '>200%',     'Funds From Operations',     5.8,  1.3);
  kpi(s, 'Liquidity',        'KZT 1,375b','Cash + Undrawn RCF',        8.45, 1.3);

  s.addText('Credit Metrics vs. EEMEA Peers', { x: 0.5, y: 3.15, w: 12.33, h: 0.4, fontSize: 15, bold: true, color: C.goldLight, fontFace: 'Cambria' });
  tableSlide(s,
    ['Company', 'Country', 'Rating', 'Net Debt/EBITDA', 'EBITDA Margin'],
    [
      ['KazMunayGas',     'Kazakhstan', 'Baa1/BBB−/BBB',   '0.16x',  '25.5%'],
      ['Sonatrach',       'Algeria',    'Unrated',          '—',       '45%+'],
      ['ADNOC',           'UAE',        'Aa2/AA/AA−',       '<0.5x',  '60%+'],
      ['Equinor',         'Norway',     'Aa2/AA−/A+',       '0.8x',   '38%'],
      ['Saudi Aramco',    'Saudi Arabia','A1/A+/A+',        '0.3x',   '58%'],
      ['MOL Group',       'Hungary',    'Baa2/BBB/BBB',     '1.2x',   '12%'],
    ]
  );
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 14 — Debt Profile
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrolDark);
  heading(s, 'Debt Profile & Liability Management');
  addGoldLine(s);
  tableSlide(s,
    ['Instrument', 'Currency', 'Coupon', 'Maturity', 'Outstanding (USD mln)', '% of Total'],
    [
      ['Eurobond 2025 (Repaid)',  'USD', '4.750%',  'Apr 2025', '—',   '—'],
      ['Eurobond 2027',          'USD', '5.375%',  'Apr 2027', '750',  '17.6%'],
      ['Eurobond 2030',          'USD', '4.875%',  'May 2030', '1,000','23.5%'],
      ['Eurobond 2033',          'USD', '5.750%',  'Jan 2033', '750',  '17.6%'],
      ['Eurobond 2048',          'USD', '6.375%',  'Oct 2048', '1,000','23.5%'],
      ['Syndicated Loan (RCF)',  'USD', 'SOFR+1.5%','2027',    '750',  '17.6%'],
      ['Other (Bank / DFIs)',    'USD', 'Various', 'Various',  '—',    '~0.2%'],
    ]
  );
  s.addText('Total gross debt: ~USD 4,250 mln. Weighted average maturity: ~8.5 years. 95%+ USD-denominated, naturally hedged by USD oil revenues.', {
    x: 0.5, y: 5.85, w: 12.33, h: 0.5, fontSize: 12, color: C.goldLight, fontFace: 'Calibri',
  });
  s.addText('This proposed Eurobond will further extend the maturity profile and pre-finance the 2027 maturity.', {
    x: 0.5, y: 6.45, w: 12.33, h: 0.35, fontSize: 12, color: 'AACCC5', fontFace: 'Calibri', italic: true,
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 15 — Eurobond Maturity Profile
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrol);
  heading(s, 'Eurobond Maturity Profile');
  addGoldLine(s);

  // Simple bar chart using shapes
  const bars = [
    { year: '2025', amt: 0,    label: 'Repaid' },
    { year: '2026', amt: 0,    label: '—' },
    { year: '2027', amt: 750,  label: '$750m' },
    { year: '2028', amt: 0,    label: '—' },
    { year: '2029', amt: 0,    label: '—' },
    { year: '2030', amt: 1000, label: '$1,000m' },
    { year: '2031', amt: 0,    label: '—' },
    { year: '2032', amt: 0,    label: '—' },
    { year: '2033', amt: 750,  label: '$750m' },
    { year: '2048', amt: 1000, label: '$1,000m' },
  ];
  const chartH = 3.8;
  const maxAmt = 1000;
  bars.forEach((b, i) => {
    const x = 0.9 + i * 1.2;
    const barH = b.amt > 0 ? (b.amt / maxAmt) * chartH : 0;
    const y = 1.4 + chartH - barH;
    if (b.amt > 0) {
      s.addShape(pptx.ShapeType.rect, { x, y, w: 0.9, h: barH, fill: { color: C.gold }, rectRadius: 0.05 });
      s.addText(b.label, { x: x-0.1, y: y-0.45, w: 1.1, h: 0.35, fontSize: 11, bold: true, color: C.gold, fontFace: 'Calibri', align: 'center' });
    }
    s.addText(b.year, { x: x-0.1, y: 5.3, w: 1.1, h: 0.3, fontSize: 11, color: 'AACCC5', fontFace: 'Calibri', align: 'center' });
  });
  s.addShape(pptx.ShapeType.rect, { x: 0.7, y: 5.25, w: 11.9, h: 0.04, fill: { color: C.gold + '44' } });
  s.addText('USD mln (face value)', { x: 0.5, y: 5.65, w: 5, h: 0.3, fontSize: 11, color: '667D79', fontFace: 'Calibri', italic: true });
  s.addText('New Eurobond (this transaction) will further extend the maturity profile beyond 2033 with a new tenor to be determined.', {
    x: 0.5, y: 6.1, w: 12.33, h: 0.5, fontSize: 12, color: C.goldLight, fontFace: 'Calibri',
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 16 — Credit Ratings
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrolDark);
  heading(s, 'Credit Ratings');
  addGoldLine(s);
  const agencies = [
    { name: "Moody's", rating: 'Baa1', outlook: 'Stable', action: 'Affirmed Oct 2025', color: '3A9BD5' },
    { name: 'S&P Global', rating: 'BBB−', outlook: 'Stable', action: 'Affirmed Nov 2025', color: 'E8485A' },
    { name: 'Fitch Ratings', rating: 'BBB', outlook: 'Stable', action: 'Affirmed Dec 2025', color: 'F5A623' },
  ];
  agencies.forEach((a, i) => {
    const x = 0.5 + i * 4.3;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.3, w: 4.0, h: 4.2, fill: { color: C.petrol }, line: { color: a.color, width: 1.5 }, rectRadius: 0.12 });
    s.addText(a.name, { x, y: 1.45, w: 4.0, h: 0.45, fontSize: 16, bold: true, color: a.color, fontFace: 'Cambria', align: 'center' });
    s.addText(a.rating, { x, y: 2.05, w: 4.0, h: 1.4, fontSize: 72, bold: true, color: C.white, fontFace: 'Cambria', align: 'center' });
    s.addShape(pptx.ShapeType.rect, { x: x+0.8, y: 3.55, w: 2.4, h: 0.5, fill: { color: '2A7D5A33' }, line: { color: '2A7D5A', width: 0.75 }, rectRadius: 0.06 });
    s.addText(a.outlook, { x: x+0.8, y: 3.55, w: 2.4, h: 0.5, fontSize: 14, bold: true, color: '64D2A0', fontFace: 'Calibri', align: 'center', valign: 'middle' });
    s.addText(a.action, { x, y: 4.18, w: 4.0, h: 0.35, fontSize: 11, color: '667D79', fontFace: 'Calibri', align: 'center', italic: true });
  });
  s.addText('All three major rating agencies have affirmed KMG\'s investment grade ratings with Stable outlook, reflecting the company\'s strong financial position and sovereign support.', {
    x: 0.5, y: 5.8, w: 12.33, h: 0.6, fontSize: 13, color: 'AACCC5', fontFace: 'Calibri', align: 'center',
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 17 — Capital Returns Global Context
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrol);
  heading(s, 'Capital Returns in a Global Context');
  addGoldLine(s);
  tableSlide(s,
    ['Company', 'Dividend Yield', 'Buyback', 'Total Shareholder Return', 'Rating'],
    [
      ['KMG (2025)',          '~3.5%',  'N/A (state-owned)',   '~3.5%',   'Baa1/BBB−/BBB'],
      ['Shell plc',          '3.8%',   '$3.5 bln buyback',    '~7%',     'Aa3/A+/AA−'],
      ['Saudi Aramco',       '4.2%',   'N/A',                 '~4.2%',   'A1/A+/A+'],
      ['Equinor',            '3.1%',   '$1.2 bln buyback',    '~6%',     'Aa2/AA−/A+'],
      ['TotalEnergies',      '4.0%',   '$2.0 bln buyback',    '~7%',     'Aa3/A+/AA−'],
      ['Lukoil (pre-sanct.)', '8%+',  'Suspended',            '8%+',     'N/R (sanctioned)'],
    ]
  );
  s.addText('KMG dividend is paid to Samruk-Kazyna (100% shareholder). Dividend policy targets ~30% of consolidated net income.', {
    x: 0.5, y: 6.3, w: 12.33, h: 0.5, fontSize: 12, color: C.goldLight, fontFace: 'Calibri', italic: true,
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 18 — Capital Allocation
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrolDark);
  heading(s, 'Capital Allocation & Dividend Policy');
  addGoldLine(s);
  kpi(s, 'Dividend FY2025',  'KZT 312 bln', '~28% of Net Income',      0.5,  1.3);
  kpi(s, 'Capex FY2025',     'KZT 1,247 bln','Growth + Maintenance',   3.15, 1.3);
  kpi(s, 'Debt Repaid FY25', '~USD 400 mln', 'Voluntary prepayments',  5.8,  1.3);
  kpi(s, 'FCF FY2025',       'KZT 1,146 bln','After Capex & Tax',      8.45, 1.3);

  s.addText('Capital Allocation Priorities', { x: 0.5, y: 3.15, w: 12.33, h: 0.4, fontSize: 15, bold: true, color: C.goldLight, fontFace: 'Cambria' });
  const priorities = [
    ['1st', 'Maintain Investment Grade', 'Net Debt/EBITDA target: below 2.0x under any oil price scenario'],
    ['2nd', 'Fund Growth Capex', 'TCO FGP ramp-up, Kashagan sustaining, Renewables, Petrochemicals'],
    ['3rd', 'Dividend to Samruk-Kazyna', 'Minimum 25% of net income; target ~30%; flexible in downturn'],
    ['4th', 'Balance Sheet Optimisation', 'Opportunistic debt buybacks, maturity extension, currency diversification'],
  ];
  priorities.forEach(([rank, title, desc], i) => {
    const y = 3.65 + i * 0.78;
    s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 12.33, h: 0.65, fill: { color: i % 2 === 0 ? C.petrol : C.petrolLight }, rectRadius: 0.05 });
    s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 0.7, h: 0.65, fill: { color: C.gold }, rectRadius: 0.05 });
    s.addText(rank, { x: 0.5, y, w: 0.7, h: 0.65, fontSize: 13, bold: true, color: C.petrolDark, fontFace: 'Cambria', align: 'center', valign: 'middle' });
    s.addText(title, { x: 1.35, y: y+0.04, w: 3.5, h: 0.3, fontSize: 13, bold: true, color: C.white, fontFace: 'Cambria' });
    s.addText(desc,  { x: 1.35, y: y+0.34, w: 11.3, h: 0.26, fontSize: 11, color: 'AACCC5', fontFace: 'Calibri' });
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 19 — ESG & Sustainability
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrol);
  heading(s, 'ESG & Sustainability');
  addGoldLine(s);
  const pillars = [
    { letter: 'E', title: 'Environmental', color: '2A9D6E', items: [
      'Flaring intensity reduced 29% since 2022',
      '100 MW renewable energy under construction',
      'Carbon neutrality roadmap by 2060 (aligned with RK NDC)',
      'Zero significant oil spills target by 2027',
      'Methane emissions monitoring across all upstream assets',
    ]},
    { letter: 'S', title: 'Social', color: '3A9BD5', items: [
      'LTIF rate 0.21 (best-ever, FY2025)',
      '72,000+ employees, 95% local workforce',
      'KZT 48 bln community investment FY2025',
      'Vendor localisation: 68% domestic content in procurement',
      'Women in leadership: 31% (up from 24% in 2022)',
    ]},
    { letter: 'G', title: 'Governance', color: C.gold, items: [
      'Supervisory Board: 9 members, 4 independent directors',
      'Annual IFRS audit by Big-4 firm',
      'Anti-corruption programme aligned with ISO 37001',
      'Annual sustainability report (GRI Standards)',
      'Board ESG Committee established 2023',
    ]},
  ];
  pillars.forEach((p, i) => {
    const x = 0.5 + i * 4.3;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.25, w: 4.0, h: 5.7, fill: { color: C.petrolDark }, line: { color: p.color, width: 1 }, rectRadius: 0.1 });
    s.addShape(pptx.ShapeType.rect, { x: x+0.2, y: 1.38, w: 0.55, h: 0.55, fill: { color: p.color }, rectRadius: 0.06 });
    s.addText(p.letter, { x: x+0.2, y: 1.38, w: 0.55, h: 0.55, fontSize: 20, bold: true, color: C.petrolDark, fontFace: 'Cambria', align: 'center', valign: 'middle' });
    s.addText(p.title, { x: x+0.88, y: 1.44, w: 3.0, h: 0.44, fontSize: 15, bold: true, color: p.color, fontFace: 'Cambria' });
    p.items.forEach((item, j) => {
      s.addShape(pptx.ShapeType.rect, { x: x+0.28, y: 2.12+j*0.84, w: 0.13, h: 0.13, fill: { color: p.color }, rectRadius: 0.02 });
      s.addText(item, { x: x+0.52, y: 2.04+j*0.84, w: 3.35, h: 0.72, fontSize: 11.5, color: 'AACCC5', fontFace: 'Calibri', valign: 'middle' });
    });
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 20 — Risk Factors
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrolDark);
  heading(s, 'Risk Factors');
  addGoldLine(s);
  const risks = [
    { cat: 'Market', title: 'Oil Price Volatility', desc: 'Revenue and cashflows are sensitive to Brent crude prices. A $10/bbl decline reduces EBITDA by ~KZT 180 bln.' },
    { cat: 'FX', title: 'Tenge / USD Fluctuation', desc: 'KZT-denominated costs vs. USD revenues. Depreciation benefits exports but inflates KZT-reported debt.' },
    { cat: 'Operational', title: 'Asset Concentration', desc: 'TCO and Kashagan represent ~60% of production. Technical outages (as in 2019–2020) can materially affect volumes.' },
    { cat: 'Regulatory', title: 'Production Sharing Agreements', desc: 'PSA terms are long-dated but subject to renegotiation risk. Changes in fiscal regime or royalties could impact profitability.' },
    { cat: 'Geopolitical', title: 'CPC Pipeline Route Risk', desc: 'Approximately 80% of crude exports transit Russia via CPC. Disruptions (as in 2022) can temporarily restrict export capacity.' },
    { cat: 'ESG', title: 'Climate Transition Risk', desc: 'Long-term demand for fossil fuels faces structural uncertainty. KMG\'s renewable and diversification strategy partially mitigates this.' },
  ];
  risks.forEach((r, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = col === 0 ? 0.5 : 6.8;
    const y = 1.25 + row * 1.8;
    const catColors = { Market: 'E8485A', FX: 'F5A623', Operational: '3A9BD5', Regulatory: '9B59B6', Geopolitical: 'E67E22', ESG: '2A9D6E' };
    const c = catColors[r.cat] || C.gold;
    s.addShape(pptx.ShapeType.rect, { x, y, w: 5.9, h: 1.6, fill: { color: C.petrol }, line: { color: c + '88', width: 0.75 }, rectRadius: 0.08 });
    s.addShape(pptx.ShapeType.rect, { x: x+0.18, y: y+0.15, w: 0.75, h: 0.28, fill: { color: c + '33' }, line: { color: c, width: 0.5 }, rectRadius: 0.04 });
    s.addText(r.cat, { x: x+0.18, y: y+0.15, w: 0.75, h: 0.28, fontSize: 9, bold: true, color: c, fontFace: 'Calibri', align: 'center', valign: 'middle' });
    s.addText(r.title, { x: x+1.06, y: y+0.12, w: 4.65, h: 0.35, fontSize: 13, bold: true, color: C.white, fontFace: 'Cambria' });
    s.addText(r.desc, { x: x+0.18, y: y+0.58, w: 5.55, h: 0.9, fontSize: 11.5, color: 'AACCC5', fontFace: 'Calibri' });
  });
}

/* ═══════════════════════════════════════════════════════════
   SLIDE 21 — Summary Investment Case
═══════════════════════════════════════════════════════════ */
{
  const s = pptx.addSlide();
  addBg(s, C.petrolDark);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.12, fill: { color: C.gold } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 7.38, w: '100%', h: 0.12, fill: { color: C.gold } });
  heading(s, 'Summary Investment Case', 0.3, C.goldLight, 30);
  addGoldLine(s, 1.0);

  const points = [
    ['National Champion with Sovereign Backing', 'Sole national oil & gas company; 100% owned by Samruk-Kazyna sovereign wealth fund. Strong implied and explicit state support underpins credit quality.'],
    ['Investment Grade Across All Three Agencies', 'Baa1 / BBB− / BBB (all Stable) — the only Central Asian energy company rated investment grade by Moody\'s, S&P and Fitch simultaneously.'],
    ['Exceptional Deleveraging Trajectory', 'Net Debt/EBITDA compressed from 1.42x (2022) to 0.16x (2025) — driven by record free cash flow and disciplined capital management.'],
    ['World-Class Asset Base', 'Equity interests in Tengizchevroil and Kashagan — two of the largest oil fields globally — provide long reserve life and production visibility.'],
    ['Diversified Business Model', 'Vertically integrated across E&P, midstream, refining and retail, with international exposure through Romania and the Netherlands.'],
    ['Growing ESG Credentials', 'Flaring intensity down 29%, LTIF at record lows, renewable energy pipeline of 100 MW, aligned with Kazakhstan\'s 2060 carbon neutrality target.'],
  ];
  points.forEach(([title, desc], i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = col === 0 ? 0.5 : 6.8;
    const y = 1.15 + row * 1.85;
    s.addShape(pptx.ShapeType.rect, { x, y, w: 5.9, h: 1.65, fill: { color: C.petrol }, line: { color: C.gold, width: 0.75 }, rectRadius: 0.1 });
    s.addShape(pptx.ShapeType.rect, { x: x+0.18, y: y+0.15, w: 0.38, h: 0.38, fill: { color: C.gold }, rectRadius: 0.05 });
    s.addText('✓', { x: x+0.18, y: y+0.15, w: 0.38, h: 0.38, fontSize: 16, bold: true, color: C.petrolDark, fontFace: 'Calibri', align: 'center', valign: 'middle' });
    s.addText(title, { x: x+0.7, y: y+0.15, w: 5.0, h: 0.4, fontSize: 12.5, bold: true, color: C.goldLight, fontFace: 'Cambria' });
    s.addText(desc, { x: x+0.18, y: y+0.62, w: 5.55, h: 0.9, fontSize: 11, color: 'AACCC5', fontFace: 'Calibri' });
  });
}

/* ── Write file ─────────────────────────────────────────────── */
pptx.writeFile({ fileName: 'downloads/KMG_Eurobond_Investor_Presentation.pptx' })
  .then(() => console.log('✓ Presentation saved to downloads/'))
  .catch(e => { console.error(e); process.exit(1); });
