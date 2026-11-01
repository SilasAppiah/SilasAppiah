// Builds the October 2026 Digital Marketing Report deck in the Cohort 49 deck style.
// Run: node deck/build_october.js  (from the repo root)
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const fa = require("react-icons/fa6");

const SKILL = "/root/.claude/skills/synced/078f54ce-71b1-48f9-9a94-f0cdff29f035_3a34c141-1a04-4bcc-bbd9-25ec14e5f233/pptx";
const OUT = path.join(__dirname, "..", "reports", "October-2026-Digital-Marketing-Report.pptx");
const ASSETS = path.join(__dirname, "assets");

const HEX = {
  navy: "171365", ink: "1D1B4F", muted: "5B6180", pale: "EEF2FA", pale2: "D6E0F2",
  red: "ED2047", redDark: "B01035", gold: "E2B872", blue: "3864A9", pink: "FBDCE3", white: "FFFFFF",
};
const THEME = {
  name: "HRCC Cohort 49",
  headFontFace: "General Sans",
  bodyFontFace: "Geist",
  colors: {
    dk1: HEX.ink, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.pale,
    accent1: HEX.red, accent2: HEX.blue, accent3: HEX.gold, accent4: HEX.muted,
    accent5: HEX.pale2, accent6: HEX.pink, hlink: HEX.blue, folHlink: HEX.muted,
  },
};
const NUM = "Bebas Neue Regular";
const DATE_LINE = "Silas Atuahene Appiah  |  Digital Marketing Officer  |  November 2026";

// ---------- image helpers ----------
const cache = {};
async function png(svg, key) {
  if (cache[key]) return cache[key];
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return (cache[key] = "image/png;base64," + buf.toString("base64"));
}
const PX = 200;
async function card(w, h, kind = "grad", r = 0.12) {
  const W = Math.round(w * PX), H = Math.round(h * PX), R = Math.round(r * PX);
  const fills = {
    grad: ["1A1A6B", "3864A9"], navy: [HEX.navy, HEX.navy], pale: [HEX.pale, HEX.pale],
    pale2: [HEX.pale2, HEX.pale2], red: [HEX.red, HEX.redDark],
  }[kind];
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
    <defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
    <stop offset="0" stop-color="#${fills[0]}"/><stop offset="1" stop-color="#${fills[1]}"/></linearGradient></defs>
    <rect width="${W}" height="${H}" rx="${R}" ry="${R}" fill="url(#g)"/></svg>`;
  return png(svg, `card-${kind}-${W}-${H}-${R}`);
}
async function icon(name, color = "FFFFFF", circle = null) {
  const svgIcon = ReactDOMServer.renderToStaticMarkup(
    React.createElement(fa[name], { color: "#" + color, size: circle ? 150 : 256 })
  );
  if (!circle) return png(svgIcon, `icon-${name}-${color}`);
  const inner = Buffer.from(svgIcon).toString("base64");
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="256" height="256">
    <circle cx="128" cy="128" r="128" fill="#${circle}"/>
    <image x="53" y="53" width="150" height="150" xlink:href="data:image/svg+xml;base64,${inner}"/></svg>`;
  return png(svg, `icon-${name}-${color}-${circle}`);
}
const img = (file) => "image/png;base64," + fs.readFileSync(path.join(ASSETS, file)).toString("base64");

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Silas Atuahene Appiah";
  pres.company = "HR Certification Centre";
  pres.title = "Digital Marketing Report: October 2026";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  const header = (dark) => [
    { text: { text: "HR CERTIFICATION CENTRE", options: { x: 0.5, y: 0.3, w: 4, h: 0.3, fontSize: 9, bold: true, charSpacing: 2, color: dark ? C.background1 : C.text2, fontFace: THEME.headFontFace, margin: 0 } } },
    { text: { text: "October 2026 report", options: { x: 6.5, y: 0.3, w: 3, h: 0.3, fontSize: 9, align: "right", color: dark ? C.accent5 : C.accent4, margin: 0 } } },
  ];
  pres.defineSlideMaster({
    title: "HRCC Content Titled",
    background: { color: HEX.white },
    objects: [
      ...header(false),
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.68, w: 9, h: 0.65, fontSize: 26, bold: true, color: C.text2, fontFace: THEME.headFontFace, margin: 0, valign: "middle", align: "left" }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.5, y: 1.33, w: 9, h: 0.35, fontSize: 13, color: C.accent4, margin: 0, valign: "top", align: "left" }, text: "" } },
    ],
    slideNumber: { x: 9.1, y: 5.25, w: 0.4, h: 0.25, fontSize: 8, color: HEX.muted, align: "right" },
  });
  pres.defineSlideMaster({
    title: "HRCC Dark",
    background: { data: img("bg_hero.png") },
    objects: [header(true)[0]],
  });

  const T = (slide, text, opts) => slide.addText(text, { isTextBox: true, margin: 0, valign: "top", ...opts });
  const content = (sectionTitle, title, sub) => {
    const s = pres.addSlide({ masterName: "HRCC Content Titled", sectionTitle });
    s.addText(title, { placeholder: "title" });
    s.addText(sub, { placeholder: "body" });
    return s;
  };
  const bullets = (items, o = {}) =>
    items.map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < items.length - 1, paraSpaceAfter: 5, ...o } }));
  const watermark = (s, text, o) =>
    T(s, text, { fontFace: NUM, fontSize: 280, color: C.background1, transparency: 86, align: "right", valign: "middle", objectName: "Watermark", ...o });
  const chartText = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt" };

  // ===== 1. Title =====
  pres.addSection({ title: "Opening" });
  let s = pres.addSlide({ masterName: "HRCC Dark", sectionTitle: "Opening" });
  watermark(s, "OCT", { x: 3.6, y: -0.2, w: 6.6, h: 5.6 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.25, w: 2.5, h: 0.34, rectRadius: 0.17, fill: { color: C.background1, transparency: 100 }, line: { color: C.background1, width: 1 }, objectName: "Pill" });
  T(s, "Monthly report", { x: 0.5, y: 1.25, w: 2.5, h: 0.34, fontSize: 11, color: C.background1, align: "center", valign: "middle" });
  T(s, "Digital Marketing\nReport: October 2026", { x: 0.5, y: 1.8, w: 7, h: 1.45, fontSize: 36, bold: true, color: C.background1, fontFace: THEME.headFontFace, valign: "middle" });
  T(s, "Cohort 49 ads, the HR Summit push,\nand what we change in November.", { x: 0.5, y: 3.35, w: 6, h: 0.7, fontSize: 15, color: C.accent5 });
  T(s, DATE_LINE, { x: 0.5, y: 4.95, w: 7, h: 0.3, fontSize: 10, color: C.accent5 });
  s.addNotes("This is the digital marketing report for October. I'll cover what the ads delivered for Cohort 49, how the organic content went, the other work done this month, and what I plan to change in November before intake closes on 29 November.");

  // ===== 2. Summary =====
  s = content("Opening", "October in four numbers", "Cohort 49 ads are bringing in enrolments at about $29 each");
  const hero = [
    ["34", "Cohort 49 enrolments", "so far, and more than 90% came through our ads"],
    ["$888", "Cohort 49 ad spend", "in October across Meta and LinkedIn"],
    ["$29", "Per enrolment", "October spend against the 31 or more ad-sourced enrolments"],
    ["1,603", "Leads and chats", "214 LinkedIn leads, 146 Meta leads and 1,243 WhatsApp and Messenger chats"],
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.5 + i * 2.3;
    s.addImage({ data: await card(2.1, 2.5, i === 2 ? "red" : "grad"), x, y: 1.9, w: 2.1, h: 2.5, objectName: `Hero card ${i + 1}` });
    T(s, hero[i][1], { x: x + 0.2, y: 2.05, w: 1.75, h: 0.3, fontSize: 12, color: i === 2 ? C.background1 : C.accent5 });
    T(s, hero[i][0], { x: x + 0.2, y: 2.35, w: 1.75, h: 0.95, fontFace: NUM, fontSize: 44, color: C.background1, valign: "middle" });
    T(s, hero[i][2], { x: x + 0.2, y: 3.35, w: 1.75, h: 0.95, fontSize: 10.5, color: C.background1 });
  }
  T(s, "The 34 enrolments are worth at least GHC 265,200 in fees (if all were aPHRi at GHC 7,800). November's job is turning the leads we already have into enrolments.", { x: 0.5, y: 4.6, w: 9, h: 0.5, fontSize: 11, color: C.accent4 });
  s.addNotes("The headline: 34 people have enrolled for Cohort 49 so far, and more than 90% of them came through the ads. We spent $888 on Cohort 49 ads in October, which is about $29 per enrolment. The Meta campaign started on 27 September, so a few days of September spend aren't in that figure, and the true number is a little higher. At the lowest fee, aPHRi at GHC 7,800, those 34 enrolments are worth at least GHC 265,200. If asked: the $29 uses October spend only, set against at least 31 ad-sourced enrolments.");

  // ===== 3. Paid overview =====
  pres.addSection({ title: "Paid advertising" });
  const SP = "Paid advertising";
  s = content(SP, "Where the ad money went", "Three campaigns, $1,055.89 in total. Figures cover October only");
  s.addChart(pres.charts.BAR, [{ name: "Spend ($)", labels: ["Cohort 49, Meta", "Cohort 49, LinkedIn", "HR Summit, Meta"], values: [552.38, 335.82, 167.69] }], {
    x: 0.5, y: 1.85, w: 4.4, h: 3.2, barDir: "bar", chartColors: [HEX.navy, HEX.blue, HEX.red],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "$#,##0", dataLabelColor: HEX.ink, dataLabelFontSize: 11,
    catAxisLabelColor: HEX.muted, catAxisLabelFontSize: 11, catAxisOrientation: "maxMin",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMaxVal: 700,
    showLegend: false, showTitle: true, title: "Spend by campaign", titleColor: HEX.navy, titleFontSize: 12, ...chartText,
    objectName: "Spend chart",
  });
  const camp = [
    ["FaMeta", "Cohort 49, Meta", "1,243 chats and 146 leads. About $0.44 per conversation"],
    ["FaLinkedin", "Cohort 49, LinkedIn (14 to 31 Oct)", "214 leads at $1.57 each, from 53,844 impressions"],
    ["FaBullhorn", "HR Summit awareness, Meta", "768,946 people reached for $0.22 per 1,000"],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.9 + i * 1.07;
    s.addImage({ data: await card(4.45, 0.92, "pale"), x: 5.05, y, w: 4.45, h: 0.92, objectName: `Campaign card ${i + 1}` });
    s.addImage({ data: await icon(camp[i][0], HEX.white, i === 2 ? HEX.red : HEX.navy), x: 5.22, y: y + 0.22, w: 0.48, h: 0.48 });
    T(s, camp[i][1], { x: 5.85, y: y + 0.12, w: 3.5, h: 0.3, fontSize: 12, bold: true, color: C.text2, fontFace: THEME.headFontFace });
    T(s, camp[i][2], { x: 5.85, y: y + 0.42, w: 3.5, h: 0.42, fontSize: 10.5, color: C.accent4 });
  }
  s.addNotes("Total ad spend in October was $1,055.89. Most of it, $888, went to Cohort 49: $552 on Meta and $336 on LinkedIn. The LinkedIn campaign only ran from 14 October, so that's 18 days. The other $168 was the HR Summit awareness ad, which reached almost 769,000 people at 22 cents per thousand, the cheapest reach of the month.");

  // ===== 4. LinkedIn =====
  s = content(SP, "LinkedIn brought in 214 leads in 18 days", "Cohort 49 lead ad, 14 to 31 October, $335.82");
  s.addImage({ data: await card(3.7, 3.2), x: 0.5, y: 1.9, w: 3.7, h: 3.2, objectName: "LinkedIn stats card" });
  const li = [["214", "leads"], ["$1.57", "per lead"], ["2.06%", "click-through rate"], ["23%", "finished the form"]];
  for (let i = 0; i < 4; i++) {
    const x = 0.7 + (i % 2) * 1.75, y = 2.1 + Math.floor(i / 2) * 1.4;
    T(s, li[i][0], { x, y, w: 1.7, h: 0.8, fontFace: NUM, fontSize: 34, color: i === 1 ? C.accent3 : C.background1, valign: "middle" });
    T(s, li[i][1], { x, y: y + 0.82, w: 1.7, h: 0.35, fontSize: 11, color: C.accent5 });
  }
  s.addChart(pres.charts.BAR, [{ name: "Leads", labels: ["Education", "Administrative", "Customer Success", "Sales", "Operations", "Human Resources", "Business Dev."], values: [19, 15, 14, 14, 11, 9, 7] }], {
    x: 4.45, y: 1.85, w: 5.05, h: 2.95, barDir: "bar", chartColors: [HEX.navy, HEX.navy, HEX.navy, HEX.navy, HEX.navy, HEX.red, HEX.navy],
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: HEX.ink, dataLabelFontSize: 10,
    catAxisLabelColor: HEX.muted, catAxisLabelFontSize: 10, catAxisOrientation: "maxMin",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMaxVal: 22,
    showLegend: false, showTitle: true, title: "Leads by job function", titleColor: HEX.navy, titleFontSize: 12, ...chartText,
    objectName: "Job function chart",
  });
  T(s, "Most leads came from outside HR, which suits aPHRi. Functions with fewer than 3 leads are not shown.", { x: 4.55, y: 4.8, w: 4.95, h: 0.35, fontSize: 9.5, color: C.accent4 });
  s.addNotes("LinkedIn brought in 214 leads in 18 days, about 12 a day, at $1.57 per lead. That's close to the $1.35 our June Cohort 47 Meta ad achieved, and LinkedIn is usually more expensive. The surprise is who responded. Education, Administrative, Customer Success, Sales and Operations all brought in more leads than HR. That suits aPHRi, which needs no HR experience. About half of the 214 leads come from job functions too small for LinkedIn to report on their own.");

  // ===== 5. HR drop-off =====
  s = content(SP, "HR professionals open the form, then stop", "Share of people who finished the LinkedIn lead form after opening it");
  s.addChart(pres.charts.BAR, [{ name: "Completion rate", labels: ["Sales", "Education", "Accounting", "Business Dev.", "Administrative", "Operations", "Customer Success", "Human Resources"], values: [48.28, 46.34, 41.67, 30.43, 28.3, 27.5, 25.93, 14.06] }], {
    x: 0.5, y: 1.85, w: 5.4, h: 3.3, barDir: "bar", chartColors: [HEX.navy, HEX.navy, HEX.navy, HEX.navy, HEX.navy, HEX.navy, HEX.navy, HEX.red],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0\"%\"", dataLabelColor: HEX.ink, dataLabelFontSize: 10,
    catAxisLabelColor: HEX.muted, catAxisLabelFontSize: 10, catAxisOrientation: "maxMin",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMaxVal: 58,
    showLegend: false, ...chartText, objectName: "Completion rate chart",
  });
  s.addImage({ data: await card(3.4, 3.2, "red"), x: 6.1, y: 1.9, w: 3.4, h: 3.2, objectName: "HR callout card" });
  T(s, "Human Resources", { x: 6.35, y: 2.08, w: 2.9, h: 0.3, fontSize: 12, color: C.background1, bold: true });
  T(s, "64", { x: 6.35, y: 2.4, w: 1.2, h: 0.8, fontFace: NUM, fontSize: 52, color: C.background1, valign: "middle" });
  T(s, "form opens,\nthe most of any group", { x: 7.5, y: 2.5, w: 1.85, h: 0.6, fontSize: 10.5, color: C.background1 });
  T(s, "9", { x: 6.35, y: 3.2, w: 1.2, h: 0.8, fontFace: NUM, fontSize: 52, color: C.background1, valign: "middle" });
  T(s, "finished it\n(14%)", { x: 7.5, y: 3.3, w: 1.85, h: 0.6, fontSize: 10.5, color: C.background1 });
  T(s, "Our PHRi and SPHRi audience is interested. In November I'll test a shorter form and an HR-specific message.", { x: 6.35, y: 4.1, w: 2.95, h: 0.85, fontSize: 10.5, color: C.background1 });
  s.addNotes("This is the main thing I learned from LinkedIn. HR professionals opened the lead form more than anyone else, 64 times, but only 9 finished it. That's 14%, against 46 to 48% for Education and Sales. These are the people we most want for PHRi and SPHRi. My guess is the form asks too much or the message is pitched at beginners. In November I'll test a shorter form with a PHRi and SPHRi message aimed at them.");

  // ===== 6. Meta =====
  s = content(SP, "Meta started 1,243 conversations", "Cohort 49 Meta ads, 1 to 31 October, $552.38");
  const funnel = [
    ["313,744", "people reached, each about twice"],
    ["1,243", "WhatsApp and Messenger conversations started"],
    ["1,106", "of them with new contacts"],
    ["146", "Meta leads"],
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.5 + i * 0.3, w = 5.3 - i * 0.3, y = 1.9 + i * 0.8;
    s.addImage({ data: await card(w, 0.66, i === 3 ? "red" : "grad", 0.08), x, y, w, h: 0.66, objectName: `Funnel bar ${i + 1}` });
    T(s, funnel[i][0], { x: x + 0.2, y, w: 1.75, h: 0.66, fontFace: NUM, fontSize: 28, color: C.background1, valign: "middle" });
    T(s, funnel[i][1], { x: x + 2.0, y, w: w - 2.15, h: 0.66, fontSize: 11, color: C.background1, valign: "middle" });
  }
  s.addImage({ data: await card(3.45, 3.1, "pale"), x: 6.05, y: 1.9, w: 3.45, h: 3.1, objectName: "Meta cost card" });
  T(s, "$0.44", { x: 6.3, y: 2.05, w: 1.6, h: 0.7, fontFace: NUM, fontSize: 32, color: C.text2, valign: "middle" });
  T(s, "per conversation", { x: 7.95, y: 2.25, w: 1.6, h: 0.35, fontSize: 11, color: C.accent4 });
  T(s, "$0.82", { x: 6.3, y: 2.75, w: 1.6, h: 0.7, fontFace: NUM, fontSize: 32, color: C.text2, valign: "middle" });
  T(s, "per 1,000 impressions", { x: 7.95, y: 2.85, w: 1.45, h: 0.5, h: 0.35, fontSize: 11, color: C.accent4 });
  T(s, bullets([
    "Only 8 Apply clicks and 3 website registrations: people enrol through chats with Sales",
    "Since 27 September: 1,461 chats and 176 leads",
  ]), { x: 6.3, y: 3.6, w: 3.0, h: 1.3, fontSize: 10.5, color: C.text1 });
  s.addNotes("Meta works differently. It's mainly a conversation ad. In October it reached about 314,000 people and started 1,243 WhatsApp and Messenger chats, at about 44 cents each, and produced 146 leads. Very few people went through the website: 8 Apply clicks and 3 registrations. So most enrolments happen in chats with Grace and the sales team. Counting from when the campaign started on 27 September, it's 1,461 chats and 176 leads. If asked about the 'orders' and 'purchases' Meta shows: those don't fit a lead campaign and look like a tracking setup issue, which I'm fixing.");

  // ===== 7. Organic =====
  pres.addSection({ title: "Organic and other work" });
  const SO = "Organic and other work";
  s = content(SO, "181 posts, well past target on the main pages", "Posts in October against the monthly target. TikTok and the blog fell short");
  s.addChart(pres.charts.BAR, [
    { name: "Monthly target", labels: ["LinkedIn", "Facebook", "Instagram", "TikTok", "YouTube", "Blog"], values: [15, 25, 12, 12, 3, 4] },
    { name: "October", labels: ["LinkedIn", "Facebook", "Instagram", "TikTok", "YouTube", "Blog"], values: [67, 56, 53, 4, 1, 0] },
  ], {
    x: 0.5, y: 1.85, w: 5.3, h: 3.3, barDir: "col", barGrouping: "clustered", chartColors: [HEX.pale2, HEX.navy],
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: HEX.ink, dataLabelFontSize: 9,
    catAxisLabelColor: HEX.muted, catAxisLabelFontSize: 10,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMaxVal: 75,
    showLegend: true, legendPos: "t", legendFontSize: 10, legendColor: HEX.muted, ...chartText,
    objectName: "Posts vs target chart",
  });
  s.addImage({ data: await card(3.5, 3.2, "pale"), x: 6.0, y: 1.9, w: 3.5, h: 3.2, objectName: "Themes card" });
  T(s, "What we posted about", { x: 6.22, y: 2.05, w: 3.1, h: 0.3, fontSize: 13, bold: true, color: C.text2, fontFace: THEME.headFontFace });
  T(s, bullets([
    "HR Summit: about 30 pieces, speakers, partners and the Guess the Speaker challenge",
    "Cohort 49: about 10 pieces, including the Get Ahead series",
    "Consulting: HR & the Law, PDP, AI and HR Strategy masterclasses",
    "Mentorship, HR Fix LIVE and staff celebrations",
  ]), { x: 6.22, y: 2.42, w: 3.1, h: 2.0, fontSize: 10.5, color: C.text1 });
  T(s, "Messaging: 5 of 6 SMS and 2 of 3 email broadcasts", { x: 6.22, y: 4.5, w: 3.1, h: 0.45, fontSize: 10, color: C.accent4 });
  s.addNotes("On organic content, we published 181 posts, counting each platform separately. LinkedIn, Facebook and Instagram were all well above target. TikTok and the blog fell behind because the Summit took a lot of content time. The Summit was the biggest theme with about 30 pieces. Cohort 49 had about 10. We also sent 5 of 6 planned SMS broadcasts and 2 of 3 emails.");

  // ===== 8. Other work =====
  s = content(SO, "Beyond the feed", "Partnerships, Summit preparation and support work in October");
  const work = [
    ["FaHandshake", "Sunny FM", "Media partnership secured"],
    ["FaFileSignature", "New partners", "Proposal to Altimate Advertising, arrangement sent to MTN MoMo"],
    ["FaBuilding", "Visits", "Mövenpick and Ghana Maritime Authority, plus GhIPSS HR team with a Summit invite"],
    ["FaToolbox", "Summit toolkit", "Partner toolkit, panellist flyers with Raymond, invitation card samples"],
    ["FaCommentSms", "Birthday SMS", "Automated birthday messages set up"],
    ["FaVideo", "Video and events", "Two HR Fix videos edited, AI Masterclass supported"],
  ];
  for (let i = 0; i < 6; i++) {
    const x = 0.5 + (i % 3) * 3.07, y = 1.9 + Math.floor(i / 3) * 1.6;
    s.addImage({ data: await card(2.85, 1.45, "pale"), x, y, w: 2.85, h: 1.45, objectName: `Work card ${i + 1}` });
    s.addImage({ data: await icon(work[i][0], HEX.white, i % 2 ? HEX.red : HEX.navy), x: x + 0.2, y: y + 0.2, w: 0.45, h: 0.45 });
    T(s, work[i][1], { x: x + 0.8, y: y + 0.2, w: 1.9, h: 0.45, fontSize: 13, bold: true, color: C.text2, valign: "middle", fontFace: THEME.headFontFace });
    T(s, work[i][2], { x: x + 0.2, y: y + 0.75, w: 2.5, h: 0.62, fontSize: 10.5, color: C.accent4 });
  }
  s.addNotes("Outside the posts and ads: we secured Sunny FM as a media partner, sent proposals to Altimate Advertising and MTN MoMo, visited Mövenpick and the Ghana Maritime Authority, and hosted the GhIPSS HR team. For the Summit I built the partner toolkit, worked on panellist flyers with Raymond and produced invitation card samples. I also set up automatic birthday SMS messages and edited two HR Fix videos.");

  // ===== 9. Challenges =====
  pres.addSection({ title: "Next steps" });
  const SN = "Next steps";
  s = content(SN, "What held us back", "Four issues to fix before intake closes on 29 November");
  const ch = [
    ["FaChartLine", "Tracking gaps", "We can't yet link each enrolment to the ad or post behind it. Meta also shows 'orders' and 'purchases' that don't fit a lead campaign."],
    ["FaUserTie", "HR professionals drop off", "They open the LinkedIn form the most but finish it the least (14%)."],
    ["FaComments", "Lead volume for Sales", "1,243 chats in a month is a lot for one sales line open 8am to 5pm. Reply speed affects enrolments."],
    ["FaLayerGroup", "Workload", "The Summit took much of the month's content time, so TikTok and the blog fell behind."],
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.5 + (i % 2) * 4.6, y = 1.9 + Math.floor(i / 2) * 1.62;
    s.addImage({ data: await card(4.4, 1.47, "pale"), x, y, w: 4.4, h: 1.47, objectName: `Challenge card ${i + 1}` });
    s.addImage({ data: await icon(ch[i][0], HEX.white, HEX.red), x: x + 0.2, y: y + 0.2, w: 0.45, h: 0.45 });
    T(s, ch[i][1], { x: x + 0.8, y: y + 0.2, w: 3.4, h: 0.45, fontSize: 13, bold: true, color: C.text2, valign: "middle", fontFace: THEME.headFontFace });
    T(s, ch[i][2], { x: x + 0.2, y: y + 0.75, w: 4.0, h: 0.65, fontSize: 10.5, color: C.accent4 });
  }
  s.addNotes("Four things held us back. First, tracking: we can't yet say which ad or post brought in each enrolment, and Meta is reporting orders and purchases that don't fit a lead campaign, which points to a setup problem. Second, HR professionals dropping off at the LinkedIn form. Third, volume: over 1,200 chats is a lot for one sales line, and how fast those chats are answered affects enrolments. Fourth, workload: the Summit took a lot of content time.");

  // ===== 10. November plan =====
  s = content(SN, "Plan for November", "Focus: turn the leads we have into enrolments before 29 November");
  const plan = [
    ["Keep Cohort 49 ads running", "Final push in the last two weeks before intake closes"],
    ["Test an HR-focused LinkedIn ad", "Shorter form and a PHRi or SPHRi message"],
    ["Fix tracking", "Correct Meta conversion events and add UTM links to every campaign"],
    ["Run both Q4 innovations in full", "Straight from the Classroom every weekend, ref-coded WhatsApp links for Sales"],
    ["Weekly check with Grace", "How many leads and chats became applications and enrolments"],
    ["Get TikTok back on track", "Short clips from the classroom interviews"],
  ];
  for (let i = 0; i < 6; i++) {
    const x = 0.5 + (i % 2) * 4.6, y = 1.85 + Math.floor(i / 2) * 1.08;
    T(s, String(i + 1).padStart(2, "0"), { x, y, w: 0.85, h: 0.8, fontFace: NUM, fontSize: 34, color: C.accent1, valign: "middle" });
    T(s, plan[i][0], { x: x + 0.9, y: y + 0.08, w: 3.5, h: 0.32, fontSize: 13, bold: true, color: C.text2, fontFace: THEME.headFontFace });
    T(s, plan[i][1], { x: x + 0.9, y: y + 0.42, w: 3.5, h: 0.4, fontSize: 10.5, color: C.accent4 });
    if (i < 4) s.addShape(pres.shapes.LINE, { x, y: y + 0.97, w: 4.3, h: 0, line: { color: C.accent5, width: 0.75 } });
  }
  s.addNotes("For November, everything points at one goal: turning the leads we have into enrolments before intake closes on 29 November. Keep the ads running with a final push. Test a LinkedIn ad aimed at HR professionals. Fix the tracking. Run both Q4 innovations in full. Set up a weekly check with Grace on what's converting. And get TikTok back on track using the classroom clips.");

  // ===== 11. Support needed =====
  s = content(SN, "What I need", "Two things that let us report cost per enrolment by channel");
  const asks = [
    ["A weekly lead-source count from Sales", "How many chats and leads came from Meta, LinkedIn and our pages, and how many enrolled."],
    ["Help checking the Meta pixel", "Whoever manages the website, to fix the conversion events so the numbers are trustworthy."],
  ];
  for (let i = 0; i < 2; i++) {
    const y = 1.9 + i * 1.08;
    s.addImage({ data: await card(0.92, 0.92), x: 0.5, y, w: 0.92, h: 0.92, objectName: `Ask number ${i + 1}` });
    T(s, `${i + 1}`, { x: 0.5, y, w: 0.92, h: 0.92, fontFace: NUM, fontSize: 40, color: C.background1, align: "center", valign: "middle" });
    s.addImage({ data: await card(7.93, 0.92, "pale"), x: 1.57, y, w: 7.93, h: 0.92, objectName: `Ask card ${i + 1}` });
    T(s, asks[i][0], { x: 1.8, y: y + 0.12, w: 7.5, h: 0.3, fontSize: 13, bold: true, color: C.text2, fontFace: THEME.headFontFace });
    T(s, asks[i][1], { x: 1.8, y: y + 0.44, w: 7.5, h: 0.42, fontSize: 11, color: C.accent4 });
  }
  const dates = [["29 NOV", "Cohort 49 intake closes"], ["5 DEC", "Cohort 49 starts"]];
  for (let i = 0; i < 2; i++) {
    const x = 0.5 + i * 4.55;
    s.addImage({ data: await card(4.45, 0.82, "red"), x, y: 4.18, w: 4.45, h: 0.82, objectName: `Date card ${i + 1}` });
    s.addImage({ data: await icon("FaCalendarDays", HEX.white), x: x + 0.22, y: 4.4, w: 0.38, h: 0.38 });
    T(s, dates[i][0], { x: x + 0.75, y: 4.18, w: 1.65, h: 0.82, fontFace: NUM, fontSize: 26, color: C.background1, valign: "middle" });
    T(s, dates[i][1], { x: x + 2.45, y: 4.18, w: 1.9, h: 0.82, fontSize: 12, color: C.background1, valign: "middle" });
  }
  s.addNotes("I need two things. A weekly lead-source count from Sales, so we can see which channel each enrolment came from and report cost per enrolment properly. And help from whoever manages the website to fix the Meta conversion events. With both in place, the November report will show exactly what each cedi of ad spend brings in.");

  // ===== 12. Closing =====
  s = pres.addSlide({ masterName: "HRCC Dark", sectionTitle: SN });
  watermark(s, "34", { x: 4.6, y: -0.2, w: 5.6, h: 5.6 });
  T(s, "Thank you", { x: 0.5, y: 1.75, w: 6.5, h: 1.0, fontSize: 44, bold: true, color: C.background1, fontFace: THEME.headFontFace, valign: "middle" });
  T(s, "34 enrolled so far. Next stop: turning\nOctober's leads into Cohort 49 seats.", { x: 0.5, y: 2.95, w: 6.2, h: 0.75, fontSize: 15, color: C.accent5 });
  T(s, DATE_LINE, { x: 0.5, y: 4.95, w: 7, h: 0.3, fontSize: 10, color: C.accent5 });
  s.addNotes("Thank you. I'm happy to take questions.");

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  await pres.writeFile({ fileName: OUT });
  const { applyTheme } = require(path.join(SKILL, "scripts/apply_theme.js"));
  await applyTheme(OUT, THEME);
  console.log("Wrote", OUT);
}

build().catch((e) => { console.error(e); process.exit(1); });
