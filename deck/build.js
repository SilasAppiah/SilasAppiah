// Builds the Q4 Operational Innovation deck in the Cohort 49 deck style.
// Run: node deck/build.js  (from the repo root)
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const fa = require("react-icons/fa6");

const SKILL = "/root/.claude/skills/synced/078f54ce-71b1-48f9-9a94-f0cdff29f035_3a34c141-1a04-4bcc-bbd9-25ec14e5f233/pptx";
const OUT = path.join(__dirname, "..", "Q4_Innovations_Silas_Appiah.pptx");
const ASSETS = path.join(__dirname, "assets");

// Brand palette from the Cohort 49 Digital Marketing Plan deck
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

// ---------- image helpers ----------
const cache = {};
async function png(svg, key) {
  if (cache[key]) return cache[key];
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return (cache[key] = "image/png;base64," + buf.toString("base64"));
}
const PX = 200; // pixels per inch for generated art
async function card(w, h, kind = "grad", r = 0.12) {
  const W = Math.round(w * PX), H = Math.round(h * PX), R = Math.round(r * PX);
  const fills = {
    grad: ["1A1A6B", "3864A9"], navy: [HEX.navy, HEX.navy], pale: [HEX.pale, HEX.pale],
    pale2: [HEX.pale2, HEX.pale2], pink: [HEX.pink, HEX.pink], red: [HEX.red, HEX.redDark],
    redfade: [HEX.red, "1A1A6B"],
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

// ---------- deck ----------
async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Silas Atuahene Appiah";
  pres.company = "HR Certification Centre";
  pres.title = "Q4 Operational Innovations: Digital Marketing";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  const header = (dark) => [
    { text: { text: "HR CERTIFICATION CENTRE", options: { x: 0.5, y: 0.3, w: 4, h: 0.3, fontSize: 9, bold: true, charSpacing: 2, color: dark ? C.background1 : C.text2, fontFace: THEME.headFontFace, margin: 0 } } },
    { text: { text: "Q4 innovations", options: { x: 7.5, y: 0.3, w: 2, h: 0.3, fontSize: 9, align: "right", color: dark ? C.accent5 : C.accent4, margin: 0 } } },
  ];

  pres.defineSlideMaster({
    title: "HRCC Content",
    background: { color: HEX.white },
    objects: header(false),
    slideNumber: { x: 9.1, y: 5.25, w: 0.4, h: 0.25, fontSize: 8, color: HEX.muted, align: "right" },
    margin: [0.5, 0.5, 0.5, 0.5],
  });
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
    T(s, text, { fontFace: NUM, fontSize: 300, color: C.background1, transparency: 86, align: "right", valign: "middle", objectName: "Watermark", ...o });

  // ===== 1. Title =====
  pres.addSection({ title: "Opening" });
  let s = pres.addSlide({ masterName: "HRCC Dark", sectionTitle: "Opening" });
  watermark(s, "Q4", { x: 4.6, y: -0.2, w: 5.6, h: 5.6 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.25, w: 2.5, h: 0.34, rectRadius: 0.17, fill: { color: C.background1, transparency: 100 }, line: { color: C.background1, width: 1 }, objectName: "Pill" });
  T(s, "Q4 Operational Innovation", { x: 0.5, y: 1.25, w: 2.5, h: 0.34, fontSize: 11, color: C.background1, align: "center", valign: "middle" });
  T(s, "Two better ways\nof working", { x: 0.5, y: 1.8, w: 6.5, h: 1.45, fontSize: 38, bold: true, color: C.background1, fontFace: THEME.headFontFace, valign: "middle" });
  T(s, "Capture the teaching, share it as value,\nand track who it brings in.", { x: 0.5, y: 3.35, w: 6, h: 0.7, fontSize: 15, color: C.accent5 });
  T(s, "Silas Atuahene Appiah  |  Digital Marketing Officer  |  5 October 2026", { x: 0.5, y: 4.95, w: 7, h: 0.3, fontSize: 10, color: C.accent5 });
  s.addNotes("Good morning. My two innovations are about doing the work differently, not doing more of it. One changes where our content comes from. The other changes what happens when that content gets someone interested. Both use tools HRCC already has.");

  // ===== 2. Contents =====
  s = pres.addSlide({ masterName: "HRCC Content", sectionTitle: "Opening" });
  s.addImage({ data: img("bg_panel.png"), x: 0, y: 0, w: 3.3, h: 5.625, objectName: "Side panel" });
  T(s, "HR CERTIFICATION CENTRE", { x: 0.5, y: 0.3, w: 2.6, h: 0.3, fontSize: 9, bold: true, charSpacing: 2, color: C.background1, fontFace: THEME.headFontFace });
  watermark(s, "Q4", { x: -0.3, y: 0.4, w: 3.4, h: 3.4, fontSize: 200, align: "left" });
  T(s, "What I'll\ncover", { x: 0.5, y: 3.7, w: 2.6, h: 1.2, fontSize: 30, bold: true, color: C.background1, fontFace: THEME.headFontFace, valign: "bottom" });
  const agenda = [
    ["01", "Where we are", "The plan is set. The gap is how the work gets done."],
    ["02", "Classroom to Content", "Turning every module into a week of value posts."],
    ["03", "Social Enquiry Routing", "Instant answers, and a tagged handover to Sales."],
    ["04", "Impact and support", "One loop, the numbers to watch, and three asks."],
  ];
  agenda.forEach(([n, h, d], i) => {
    const y = 0.75 + i * 1.12;
    T(s, n, { x: 3.85, y, w: 0.9, h: 0.75, fontFace: NUM, fontSize: 44, color: C.accent1, valign: "middle" });
    T(s, h, { x: 4.85, y: y + 0.08, w: 4.6, h: 0.32, fontSize: 16, bold: true, color: C.text2, fontFace: THEME.headFontFace });
    T(s, d, { x: 4.85, y: y + 0.42, w: 4.6, h: 0.3, fontSize: 12, color: C.accent4 });
    if (i < 3) s.addShape(pres.shapes.LINE, { x: 3.85, y: y + 1.0, w: 5.65, h: 0, line: { color: C.accent5, width: 0.75 } });
  });
  s.addNotes("Four parts: where we are, the two innovations, then the combined impact and what I need to make them work.");

  // ===== 3. Context =====
  s = content("Opening", "The plan is set. The gap is how the work gets done", "Cohort 49 intake closes on 29 November, and the plan rests on steady content and fast follow-up");
  const hero = [
    ["80%", "Social media's share", "of HRCC's marketing, carrying both HRCC and HRCC Consulting on the same pages"],
    ["3", "Modules every weekend", "of expert teaching, Saturdays in person and Sundays online, that our audience rarely sees"],
    ["1", "Marketing officer", "selling for every department, with a creative team stretched across many requests"],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.5 + i * 3.1;
    s.addImage({ data: await card(2.8, 2.55), x, y: 1.95, w: 2.8, h: 2.55, objectName: `Hero card ${i + 1}` });
    T(s, hero[i][1], { x: x + 0.25, y: 2.12, w: 2.3, h: 0.3, fontSize: 13, color: C.accent5 });
    T(s, hero[i][0], { x: x + 0.25, y: 2.42, w: 2.3, h: 1.0, fontFace: NUM, fontSize: 66, color: C.background1, valign: "middle" });
    T(s, hero[i][2], { x: x + 0.25, y: 3.5, w: 2.3, h: 0.85, fontSize: 11.5, color: C.background1 });
  }
  T(s, "The Cohort 49 plan already names facilitator content and tagged leads. These two innovations are the routines that make them happen every week, without extra people.", { x: 0.5, y: 4.68, w: 9, h: 0.45, fontSize: 11, color: C.accent4 });
  s.addNotes("The Cohort 49 plan says what we should post and how leads should reach Sales. What it does not solve is how one marketing officer delivers that every week while the creative team is stretched. So my innovations are routines, not new campaigns. If anyone asks whether this is already in the plan: the plan names the content, these innovations are how it actually gets made and followed up.");

  // ===== Divider helper =====
  const divider = (section, n, title, line) => {
    const d = pres.addSlide({ masterName: "HRCC Dark", sectionTitle: section });
    watermark(d, n, { x: 4.4, y: -0.2, w: 5.8, h: 5.8 });
    T(d, `Innovation ${n}`, { x: 0.5, y: 1.75, w: 5, h: 0.4, fontSize: 14, color: C.accent3, bold: true });
    T(d, title, { x: 0.5, y: 2.15, w: 6.5, h: 0.9, fontSize: 38, bold: true, color: C.background1, fontFace: THEME.headFontFace, valign: "middle" });
    T(d, line, { x: 0.5, y: 3.15, w: 5.6, h: 0.75, fontSize: 15, color: C.accent5 });
    return d;
  };

  // ===== 4. Divider 1 =====
  pres.addSection({ title: "Innovation 1: Classroom to Content" });
  const S1 = "Innovation 1: Classroom to Content";
  s = divider(S1, "01", "Classroom to Content", "One 3-minute interview after every module becomes a week of value posts.");
  s.addNotes("First innovation: Classroom to Content. A series I'm calling Straight from the Classroom.");

  // ===== 5. I1 gap and idea =====
  s = content(S1, "From the classroom to the feed in 48 hours", "1. The innovation  ·  2. The current gap");
  s.addImage({ data: await card(4.3, 2.95, "pale"), x: 0.5, y: 1.95, w: 4.3, h: 2.95, objectName: "Gap card" });
  T(s, "The gap today", { x: 0.75, y: 2.12, w: 3.8, h: 0.35, fontSize: 15, bold: true, color: C.text2, fontFace: THEME.headFontFace });
  T(s, bullets([
    "Content waits on a creative team serving every department",
    "The Wednesday shoot runs long and pulls staff away from their work",
    "The feed mixes HRCC and Consulting, so it feels salesy",
    "Planned facilitator content has no simple weekly routine behind it",
  ]), { x: 0.75, y: 2.55, w: 3.85, h: 2.2, fontSize: 12, color: C.text1 });
  s.addImage({ data: await card(4.5, 2.95), x: 5.0, y: 1.95, w: 4.5, h: 2.95, objectName: "Innovation card" });
  T(s, "The innovation", { x: 5.25, y: 2.12, w: 4, h: 0.3, fontSize: 13, color: C.accent5 });
  T(s, "After every module, a 3-minute interview with the facilitator using three fixed questions. AI turns the transcript into a week of posts for the series Straight from the Classroom.", { x: 5.25, y: 2.45, w: 4.0, h: 1.2, fontSize: 13, color: C.background1 });
  T(s, "1", { x: 5.25, y: 3.75, w: 0.4, h: 0.9, fontFace: NUM, fontSize: 60, color: C.background1, valign: "middle" });
  T(s, "interview", { x: 5.7, y: 4.08, w: 1.0, h: 0.3, fontSize: 11, color: C.accent5 });
  s.addImage({ data: await icon("FaArrowRight", HEX.gold), x: 6.75, y: 4.06, w: 0.32, h: 0.32 });
  T(s, "4+", { x: 7.25, y: 3.75, w: 0.75, h: 0.9, fontFace: NUM, fontSize: 60, color: C.accent3, valign: "middle" });
  T(s, "posts across\nour channels", { x: 8.05, y: 3.98, w: 1.3, h: 0.5, fontSize: 11, color: C.accent5 });
  s.addNotes("The gap: content depends on a stretched creative team, the Wednesday shoot takes too long, and the feed feels salesy. Meanwhile three modules of expert teaching happen every weekend. The innovation: a 3-minute interview after each module, three fixed questions, and AI turns it into four or more posts. Analogy: a cooking show filmed in a real kitchen. The cooking happens anyway. We just bring a camera and ask three questions.");

  // ===== 6. I1 how it works =====
  s = content(S1, "One weekly rhythm, about an hour of my time", "4. Action plan");
  const steps = [
    ["FaVideo", "Saturday", "In person", "Phone on a tripod with a clip-on mic. Three minutes after class. A programme officer films if I can't attend."],
    ["FaLaptop", "Sunday", "Online", "The facilitator stays on the call after participants leave. I join from home and record."],
    ["FaWandMagicSparkles", "Monday", "Process", "CapCut adds captions. AI drafts the caption, LinkedIn post and quote cards. I edit and approve."],
    ["FaCalendarCheck", "Tue to Sun", "Publish", "Meta Business Suite for Facebook and Instagram, Buffer for LinkedIn. Each post ends with a keyword prompt."],
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.5 + i * 2.3;
    s.addImage({ data: await card(2.1, 2.3), x, y: 1.9, w: 2.1, h: 2.3, objectName: `Step ${i + 1}` });
    s.addImage({ data: await icon(steps[i][0], HEX.white, HEX.red), x: x + 0.2, y: 2.07, w: 0.42, h: 0.42 });
    T(s, steps[i][1], { x: x + 0.72, y: 2.05, w: 1.3, h: 0.27, fontSize: 13, bold: true, color: C.background1, fontFace: THEME.headFontFace });
    T(s, steps[i][2], { x: x + 0.72, y: 2.3, w: 1.3, h: 0.22, fontSize: 10.5, color: C.accent3 });
    T(s, steps[i][3], { x: x + 0.2, y: 2.68, w: 1.72, h: 1.45, fontSize: 11, color: C.background1 });
  }
  s.addImage({ data: await card(9, 0.82, "pale"), x: 0.5, y: 4.38, w: 9, h: 0.82, objectName: "Questions strip" });
  T(s, "The same three questions", { x: 0.75, y: 4.47, w: 3, h: 0.28, fontSize: 12, bold: true, color: C.text2 });
  const qs = ["One idea to use on Monday?", "A mistake people often make?", "A class question that stuck?"];
  qs.forEach((q, i) => {
    T(s, `${i + 1}`, { x: 0.75 + i * 2.95, y: 4.76, w: 0.3, h: 0.3, fontFace: NUM, fontSize: 20, color: C.accent1, valign: "middle" });
    T(s, q, { x: 1.05 + i * 2.95, y: 4.76, w: 2.5, h: 0.3, fontSize: 11, color: C.text1, valign: "middle" });
  });
  s.addNotes("Saturday in person: phone, tripod, clip-on mic, three minutes. Sunday online: the facilitator stays on after participants leave, so no participants are filmed. Monday I process all three in about an hour with captions and AI drafts. Tuesday to Sunday the posts go out on schedule, each ending with a keyword like 'Comment PHRI', which feeds my second innovation. The Wednesday shoot then only covers announcements and promotions, so it gets shorter.");

  // ===== 7. I1 objectives and KPIs =====
  s = content(S1, "What success looks like by 31 December", "3. Objectives  ·  7. Performance indicators");
  const tiles1 = [["80%", "of modules captured"], ["4+", "posts from every interview"], ["48h", "from class to first post"], ["50%", "of weekly posts made without the creative team"]];
  for (let i = 0; i < 4; i++) {
    const x = 0.5 + (i % 2) * 2.2, y = 1.9 + Math.floor(i / 2) * 1.6;
    s.addImage({ data: await card(2.05, 1.45), x, y, w: 2.05, h: 1.45, objectName: `Target ${i + 1}` });
    T(s, tiles1[i][0], { x: x + 0.2, y: y + 0.1, w: 1.7, h: 0.75, fontFace: NUM, fontSize: 48, color: C.background1, valign: "middle" });
    T(s, tiles1[i][1], { x: x + 0.2, y: y + 0.85, w: 1.7, h: 0.5, fontSize: 11, color: C.accent5 });
  }
  const th = (t) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.navy } } });
  const kpiTable = (rows) => [[th("Indicator"), th("Baseline"), th("Q4 target")], ...rows.map((r, i) => r.map((c) => ({ text: c, options: { fill: { color: i % 2 ? HEX.white : HEX.pale } } })))];
  s.addTable(kpiTable([
    ["Modules captured", "0%", "80% or more"],
    ["Posts per interview", "None yet", "4 or more"],
    ["Class to first post", "No routine", "48 hours"],
    ["Posts without creative team", "Week 1 count", "50% or more"],
    ["Wednesday shoot length", "Time in week 1", "Half of today"],
    ["Saves and shares per post", "Week 1 average", "Beat promo posts"],
  ]), { x: 4.95, y: 1.9, w: 4.55, colW: [2.0, 1.25, 1.3], fontSize: 10.5, color: HEX.ink, fontFace: "Geist", rowH: 0.4, border: { type: "solid", pt: 0.5, color: HEX.pale2 }, valign: "middle", margin: [0.03, 0.08, 0.03, 0.08] });
  T(s, "Baselines are taken in the first week of October so the December figures can be compared fairly.", { x: 4.95, y: 4.9, w: 4.55, h: 0.3, fontSize: 10, color: C.accent4 });
  s.addNotes("Four targets: capture at least 80% of modules, get four or more posts from each, publish within 48 hours of class, and make half of our weekly posts without the creative team. I'll measure the Wednesday shoot and engagement this week so the December comparison is fair.");

  // ===== timeline helper =====
  const timeline = async (slide, months, supportTitle, support) => {
    const labels = ["OCT", "NOV", "DEC"];
    for (let i = 0; i < 3; i++) {
      const x = 0.5 + i * 2.05;
      slide.addImage({ data: await card(1.9, 3.2), x, y: 1.9, w: 1.9, h: 3.2, objectName: `${labels[i]} card` });
      T(slide, labels[i], { x: x + 0.2, y: 2.0, w: 1.5, h: 0.7, fontFace: NUM, fontSize: 44, color: C.background1, valign: "middle" });
      T(slide, months[i][0], { x: x + 0.2, y: 2.72, w: 1.6, h: 0.27, fontSize: 11, bold: true, color: C.accent3 });
      T(slide, months[i][1], { x: x + 0.2, y: 3.05, w: 1.6, h: 1.95, fontSize: 11, color: C.background1 });
    }
    slide.addImage({ data: await card(2.75, 3.2, "pale"), x: 6.75, y: 1.9, w: 2.75, h: 3.2, objectName: "Support card" });
    slide.addImage({ data: await icon("FaHandshake", HEX.white, HEX.navy), x: 6.95, y: 2.08, w: 0.4, h: 0.4 });
    T(slide, supportTitle, { x: 7.45, y: 2.1, w: 2.0, h: 0.36, fontSize: 13, bold: true, color: C.text2, valign: "middle", fontFace: THEME.headFontFace });
    T(slide, bullets(support, { paraSpaceAfter: 6 }), { x: 6.95, y: 2.62, w: 2.4, h: 2.4, fontSize: 11, color: C.text1 });
  };

  // ===== 8. I1 timeline and resources =====
  s = content(S1, "October to December, step by step", "5. Resources and support  ·  6. Timeline");
  await timeline(s, [
    ["Set up and pilot", "Agree the format with facilitators and get consent. Buy the kit, build the templates, take baselines. Pilot on two weekends."],
    ["Full weekly run", "Three interviews a week, four or more posts each, all scheduled a week ahead. Weekly check on the numbers."],
    ["Review and refine", "Keep the formats that work. Build a highlights set for the next intake. Write a one-page playbook."],
  ], "Support I need", [
    "Your endorsement of the series",
    "Written consent from each facilitator",
    "A programme officer to film on Saturdays when I can't",
    "One tripod and one clip-on mic",
    "Free tools we have: CapCut, an AI assistant, Meta Business Suite, Buffer, Figma",
  ]);
  s.addNotes("October is setup and a two-weekend pilot. November is the full weekly run. December is review, a highlights set for the next intake, and a playbook. The only cost is a tripod and a clip-on mic. Everything else is free or already in use.");

  // ===== 9. I1 impact and sustainability =====
  s = content(S1, "Expected Q4 impact", "8. Expected impact  ·  9. Sustainability");
  const bars1 = [
    ["24", "interviews before Cohort 49 intake closes (8 weekends, 3 modules)"],
    ["96+", "value posts from real HRCC experts"],
    ["50%", "less time in the Wednesday shoot"],
    ["0", "extra staff or ad spend needed"],
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.5 + i * 0.3, w = 5.3 - i * 0.3, y = 1.9 + i * 0.8;
    s.addImage({ data: await card(w, 0.66, "red", 0.08), x, y, w, h: 0.66, objectName: `Impact bar ${i + 1}` });
    T(s, bars1[i][0], { x: x + 0.2, y, w: 1.05, h: 0.66, fontFace: NUM, fontSize: 34, color: C.background1, valign: "middle" });
    T(s, bars1[i][1], { x: x + 1.3, y, w: w - 1.45, h: 0.66, fontSize: 11, color: C.background1, valign: "middle" });
  }
  s.addImage({ data: await card(3.45, 3.1), x: 6.05, y: 1.9, w: 3.45, h: 3.1, objectName: "Sustainability card" });
  s.addImage({ data: await icon("FaArrowsRotate", HEX.white, HEX.red), x: 6.28, y: 2.08, w: 0.4, h: 0.4 });
  T(s, "Why it lasts", { x: 6.78, y: 2.1, w: 2.5, h: 0.36, fontSize: 13, bold: true, color: C.background1, valign: "middle", fontFace: THEME.headFontFace });
  T(s, bullets([
    "Every cohort has modules, so the supply never runs out",
    "Three questions, two templates and a playbook mean anyone can run it",
    "Clips are filed by programme for future intakes",
    "Consulting trainings can run as their own lane, keeping the brands apart",
  ], { paraSpaceAfter: 6 }), { x: 6.28, y: 2.62, w: 3.05, h: 2.3, fontSize: 11, color: C.background1 });
  T(s, "Assumes classes run every weekend from 10 October to 29 November. Confirm against the class calendar.", { x: 0.5, y: 5.08, w: 5.4, h: 0.3, fontSize: 9, color: C.accent4 });
  s.addNotes("Between 10 October and the 29 November intake close there are eight teaching weekends. Three modules each gives 24 interviews and close to a hundred posts, made from teaching that happens anyway. The feed becomes more expert and less salesy, and the Wednesday shoot can be cut by half. Likely question: are we giving away paid content? No. One takeaway per module, not the lesson. It's the trailer, not the movie.");

  // ===== 10. Divider 2 =====
  const S2 = "Innovation 2: Social Enquiry Routing";
  pres.addSection({ title: S2 });
  s = divider(S2, "02", "Social Enquiry Routing", "Instant answers on social, and a tagged handover to Sales for everyone ready to enrol.");
  s.addNotes("Second innovation: what happens when that content gets someone interested.");

  // ===== 11. I2 gap and idea =====
  s = content(S2, "Every enquiry answered, every handover counted", "1. The innovation  ·  2. The current gap");
  s.addImage({ data: await card(4.3, 2.95, "pale"), x: 0.5, y: 1.95, w: 4.3, h: 2.95, objectName: "Gap card" });
  T(s, "The gap today", { x: 0.75, y: 2.12, w: 3.8, h: 0.35, fontSize: 15, bold: true, color: C.text2, fontFace: THEME.headFontFace });
  T(s, bullets([
    "I answer every DM and comment by hand, often the same questions",
    "The handover to Sales is a bare WhatsApp number",
    "We can't tell how many reach Sales, or which programme they want",
    "Lead forms and creators are tagged, but our own DMs are not",
  ]), { x: 0.75, y: 2.55, w: 3.85, h: 2.2, fontSize: 12, color: C.text1 });
  s.addImage({ data: await card(4.5, 2.95), x: 5.0, y: 1.95, w: 4.5, h: 2.95, objectName: "Innovation card" });
  T(s, "The innovation", { x: 5.25, y: 2.12, w: 4, h: 0.3, fontSize: 13, color: C.accent5 });
  T(s, "Meta Business Suite gives the first answer from an AI-drafted FAQ library. Anyone ready to enrol gets a WhatsApp link for their programme with a reference code, and Sales adds one label.", { x: 5.25, y: 2.45, w: 4.0, h: 1.3, fontSize: 13, color: C.background1 });
  s.addImage({ data: await card(4.0, 0.62, "pale2", 0.1), x: 5.25, y: 4.05, w: 4.0, h: 0.62, objectName: "Message preview" });
  s.addImage({ data: await icon("FaWhatsapp", HEX.navy), x: 5.4, y: 4.2, w: 0.32, h: 0.32 });
  T(s, "\"Hi, I'm interested in PHRi (ref: SM-PHRI)\"", { x: 5.85, y: 4.05, w: 3.3, h: 0.62, fontSize: 11.5, color: C.text2, valign: "middle", italic: true });
  s.addNotes("Today I answer each DM by hand, and when someone is ready I send the Sales number. After that we lose sight of them. The plan already tags lead forms and creators. Our own DMs are the untagged route. The innovation: automatic first answers from a FAQ library, then a WhatsApp link per programme that pre-fills a message with a reference code. Analogy: a receptionist who hands over a referral slip instead of pointing down the hall. The slip says who they are and why they came, and slips can be counted.");

  // ===== 12. I2 how it works =====
  s = content(S2, "Five steps from comment to Sales", "4. Action plan");
  const flow = [
    ["FaComments", "Comment or DM", "Someone comments a keyword like PHRI, or sends a message."],
    ["FaBolt", "Instant reply", "Meta Business Suite answers at once from the FAQ library."],
    ["FaUserCheck", "I take over", "A person replies within 15 minutes in working hours."],
    ["FaWhatsapp", "Tagged link", "Ready to enrol? They tap the programme link and press send."],
    ["FaTag", "Sales labels", "Sales sees the ref code, adds 'Social lead', and we count on Friday."],
  ];
  for (let i = 0; i < 5; i++) {
    const x = 0.5 + i * 1.835;
    s.addImage({ data: await card(1.66, 2.15), x, y: 1.9, w: 1.66, h: 2.15, objectName: `Flow step ${i + 1}` });
    T(s, `0${i + 1}`, { x: x + 0.15, y: 1.98, w: 0.7, h: 0.5, fontFace: NUM, fontSize: 26, color: C.accent3, valign: "middle" });
    s.addImage({ data: await icon(flow[i][0], HEX.white, HEX.red), x: x + 1.08, y: 2.04, w: 0.4, h: 0.4 });
    T(s, flow[i][1], { x: x + 0.15, y: 2.55, w: 1.4, h: 0.28, fontSize: 12, bold: true, color: C.background1, fontFace: THEME.headFontFace });
    T(s, flow[i][2], { x: x + 0.15, y: 2.88, w: 1.4, h: 1.1, fontSize: 10.5, color: C.accent5 });
  }
  s.addImage({ data: await card(4.4, 1.0, "pale"), x: 0.5, y: 4.2, w: 4.4, h: 1.0, objectName: "Setup note" });
  T(s, "Set up once", { x: 0.7, y: 4.3, w: 4, h: 0.26, fontSize: 12, bold: true, color: C.text2 });
  T(s, "Top 15 questions answered with AI, facts confirmed by Sales. Instant reply, FAQs, saved replies and keywords switched on.", { x: 0.7, y: 4.58, w: 4.05, h: 0.55, fontSize: 10.5, color: C.text1 });
  s.addImage({ data: await card(4.4, 1.0, "pale"), x: 5.1, y: 4.2, w: 4.4, h: 1.0, objectName: "UTM note" });
  T(s, "Website links tracked too", { x: 5.3, y: 4.3, w: 4, h: 0.26, fontSize: 12, bold: true, color: C.text2 });
  T(s, "Every website and portal link carries a UTM tag, read in Google Analytics, so all three routes to HRCC are counted.", { x: 5.3, y: 4.58, w: 4.05, h: 0.55, fontSize: 10.5, color: C.text1 });
  s.addNotes("Five steps. Someone comments a keyword or sends a DM. Meta Business Suite replies instantly. I follow up personally within 15 minutes in working hours. When they're ready, they tap the link for their programme and WhatsApp opens with the message already written, including the code. Sales sees the code, adds one label, and we count on Fridays. Website links get UTM tags so that route is counted too. Likely question: isn't WhatsApp Sales' job? Yes, and it stays theirs. My part ends at a clean, tagged handover.");

  // ===== 13. I2 objectives and KPIs =====
  s = content(S2, "What success looks like by 31 December", "3. Objectives  ·  7. Performance indicators");
  const tiles2 = [["15 min", "personal reply in working hours"], ["60%", "of routine questions answered automatically"], ["100%", "of handovers through tagged links"], ["Weekly", "count of social leads shared with Sales"]];
  for (let i = 0; i < 4; i++) {
    const x = 0.5 + (i % 2) * 2.2, y = 1.9 + Math.floor(i / 2) * 1.6;
    s.addImage({ data: await card(2.05, 1.45), x, y, w: 2.05, h: 1.45, objectName: `Target ${i + 1}` });
    T(s, tiles2[i][0], { x: x + 0.2, y: y + 0.1, w: 1.7, h: 0.75, fontFace: NUM, fontSize: 44, color: C.background1, valign: "middle" });
    T(s, tiles2[i][1], { x: x + 0.2, y: y + 0.85, w: 1.7, h: 0.5, fontSize: 11, color: C.accent5 });
  }
  s.addTable(kpiTable([
    ["Personal reply time", "20-DM sample", "15 min or less"],
    ["Answered automatically", "0%", "60% or more"],
    ["Tagged handovers to Sales", "Not tracked", "Counted weekly"],
    ["Social leads that enrol", "Not tracked", "Monthly report"],
    ["Hours on repeat replies", "Week 1 log", "Down a third"],
    ["Keyword-triggered DMs", "0", "Tracked per post"],
  ]), { x: 4.95, y: 1.9, w: 4.55, colW: [2.0, 1.25, 1.3], fontSize: 10.5, color: HEX.ink, fontFace: "Geist", rowH: 0.4, border: { type: "solid", pt: 0.5, color: HEX.pale2 }, valign: "middle", margin: [0.03, 0.08, 0.03, 0.08] });
  T(s, "Enrolment figures come from Sales' 'Social lead' label, so the count is a minimum.", { x: 4.95, y: 4.9, w: 4.55, h: 0.3, fontSize: 10, color: C.accent4 });
  s.addNotes("The targets: a personal reply within 15 minutes, 60% of routine questions handled automatically, every handover tagged, and a weekly count shared with Sales. Some people will delete the pre-filled text, so the count is a minimum. Today the count is zero.");

  // ===== 14. I2 timeline and resources =====
  s = content(S2, "October to December, step by step", "5. Resources and support  ·  6. Timeline");
  await timeline(s, [
    ["Build and switch on", "Time 20 recent DMs for a baseline. Write the FAQ library, switch on automations, create the links, brief Sales."],
    ["Live on every post", "Running on all Cohort 49 posts and the classroom series. Friday count with Sales. Fix weak answers."],
    ["Report and review", "Compare with baseline. Report social leads and enrolments. Playbook ready for the next intake."],
  ], "Support I need", [
    "Your approval of the reply standard and the Sales label step",
    "Sales: one label per tagged chat and a Friday count",
    "Sales to confirm FAQ facts: fees, dates, requirements",
    "Tools we already use: Meta Business Suite, WhatsApp Business, Google Analytics",
  ]);
  s.addNotes("October: baselines, the FAQ library, automations and links, and a short briefing with Sales. November: live on every post. December: report and playbook. No new cost. The only ask of Sales is one label and a Friday count.");

  // ===== 15. I2 impact and sustainability =====
  s = content(S2, "Expected Q4 impact", "8. Expected impact  ·  9. Sustainability");
  const bars2 = [
    ["1st", "time HRCC can count enrolments that began as a social conversation"],
    ["15", "minutes or less to a personal reply, from no set standard today"],
    ["60%", "of repeat questions answered without my time"],
    ["0", "new tools or cost: Meta Business Suite and WhatsApp Business"],
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.5 + i * 0.3, w = 5.3 - i * 0.3, y = 1.9 + i * 0.8;
    s.addImage({ data: await card(w, 0.66, "red", 0.08), x, y, w, h: 0.66, objectName: `Impact bar ${i + 1}` });
    T(s, bars2[i][0], { x: x + 0.2, y, w: 1.05, h: 0.66, fontFace: NUM, fontSize: 34, color: C.background1, valign: "middle" });
    T(s, bars2[i][1], { x: x + 1.3, y, w: w - 1.45, h: 0.66, fontSize: 11, color: C.background1, valign: "middle" });
  }
  s.addImage({ data: await card(3.45, 3.1), x: 6.05, y: 1.9, w: 3.45, h: 3.1, objectName: "Sustainability card" });
  s.addImage({ data: await icon("FaArrowsRotate", HEX.white, HEX.red), x: 6.28, y: 2.08, w: 0.4, h: 0.4 });
  T(s, "Why it lasts", { x: 6.78, y: 2.1, w: 2.5, h: 0.36, fontSize: 13, bold: true, color: C.background1, valign: "middle", fontFace: THEME.headFontFace });
  T(s, bullets([
    "Answers, keywords and links live in Meta Business Suite, so a backup can run them",
    "The weekly count goes into my monthly report",
    "New programmes, Consulting and the next intake just need new links and answers",
  ], { paraSpaceAfter: 6 }), { x: 6.28, y: 2.62, w: 3.05, h: 2.3, fontSize: 11, color: C.background1 });
  s.addNotes("The biggest change is visibility. For the first time we can say how many enrolments started as a conversation on our pages. Prospects get faster answers, Sales gets warmer conversations, and I get time back. Likely question: will it feel robotic? Only the first reply is automatic, and a person takes over straight after.");

  // ===== 16. How they connect =====
  pres.addSection({ title: "Close" });
  s = content("Close", "One loop, two innovations", "Each step feeds the next, and every step can be counted");
  const cx = 3.0, cy = 3.5, rx = 1.95, ry = 1.25;
  s.addShape(pres.shapes.OVAL, { x: cx - rx, y: cy - ry, w: rx * 2, h: ry * 2, fill: { color: HEX.white, transparency: 100 }, line: { color: C.accent5, width: 2, dashType: "dash" }, objectName: "Loop ring" });
  const nodes = [
    ["Facilitator interview", "grad"], ["Week of value posts", "grad"], ["Keyword comment or DM", "red"],
    ["Instant reply and FAQ", "red"], ["Tagged WhatsApp to Sales", "red"], ["Questions feed next interviews", "grad"],
  ];
  for (let i = 0; i < 6; i++) {
    const a = -Math.PI / 2 + (i * 2 * Math.PI) / 6;
    const nx = cx + rx * Math.cos(a) - 0.85, ny = cy + ry * Math.sin(a) - 0.27;
    s.addImage({ data: await card(1.7, 0.54, nodes[i][1], 0.1), x: nx, y: ny, w: 1.7, h: 0.54, objectName: `Loop node ${i + 1}` });
    T(s, nodes[i][0], { x: nx + 0.08, y: ny, w: 1.54, h: 0.54, fontSize: 10, color: C.background1, align: "center", valign: "middle", bold: true });
  }
  T(s, "ONE\nLOOP", { x: cx - 0.8, y: cy - 0.45, w: 1.6, h: 0.9, fontFace: NUM, fontSize: 30, color: C.text2, align: "center", valign: "middle", lineSpacingMultiple: 0.85 });
  s.addImage({ data: await card(3.45, 3.25, "navy"), x: 6.05, y: 1.9, w: 3.45, h: 3.25, objectName: "Story card" });
  T(s, "Answer in public.\nSell in private.\nTrack the handover.", { x: 6.3, y: 2.08, w: 3.0, h: 1.25, fontSize: 17, bold: true, color: C.background1, fontFace: THEME.headFontFace });
  s.addImage({ data: await card(0.26, 0.26, "grad", 0.05), x: 6.3, y: 3.6, w: 0.26, h: 0.26 });
  T(s, "Classroom to Content: makes the posts", { x: 6.68, y: 3.55, w: 2.7, h: 0.36, fontSize: 11, color: C.accent5, valign: "middle" });
  s.addImage({ data: await card(0.26, 0.26, "red", 0.05), x: 6.3, y: 4.05, w: 0.26, h: 0.26 });
  T(s, "Social Enquiry Routing: turns interest into tracked leads", { x: 6.68, y: 3.98, w: 2.7, h: 0.42, fontSize: 11, color: C.accent5, valign: "middle" });
  T(s, "Both use tools HRCC already has, with AI doing the drafting and a person approving every word.", { x: 6.3, y: 4.5, w: 3.0, h: 0.55, fontSize: 10, color: C.accent3 });
  s.addNotes("The two innovations form one loop. The classroom series gives people something useful. Each post invites a keyword comment. The comment triggers an instant answer, ready buyers go to Sales with a code, and the questions people ask become next weekend's interview questions. In one line: answer in public, sell in private, and track the handover.");

  // ===== 17. Asks =====
  s = content("Close", "What I need from you", "Three decisions this week keep both innovations on their dates");
  const asks = [
    ["Endorse the Straight from the Classroom series", "So facilitators and programme officers know it's a priority, and a programme officer can film on Saturdays when I can't."],
    ["Approve the social reply standard", "Instant first reply, a personal reply within 15 minutes in working hours, and Sales adding a 'Social lead' label with a Friday count."],
    ["Agree that Sales confirms the FAQ facts", "Fees, dates and requirements checked once in October, then whenever they change."],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.9 + i * 1.08;
    s.addImage({ data: await card(0.92, 0.92), x: 0.5, y, w: 0.92, h: 0.92, objectName: `Ask number ${i + 1}` });
    T(s, `${i + 1}`, { x: 0.5, y, w: 0.92, h: 0.92, fontFace: NUM, fontSize: 40, color: C.background1, align: "center", valign: "middle" });
    s.addImage({ data: await card(7.93, 0.92, "pale"), x: 1.57, y, w: 7.93, h: 0.92, objectName: `Ask card ${i + 1}` });
    T(s, asks[i][0], { x: 1.8, y: y + 0.12, w: 7.5, h: 0.3, fontSize: 13, bold: true, color: C.text2, fontFace: THEME.headFontFace });
    T(s, asks[i][1], { x: 1.8, y: y + 0.44, w: 7.5, h: 0.42, fontSize: 11, color: C.accent4 });
  }
  s.addNotes("Three decisions. Endorse the classroom series. Approve the reply standard and the one Sales label. And agree that Sales confirms the FAQ facts once. With those, both innovations start this week.");

  // ===== 18. Closing =====
  s = pres.addSlide({ masterName: "HRCC Dark", sectionTitle: "Close" });
  watermark(s, "Q4", { x: 4.6, y: -0.2, w: 5.6, h: 5.6 });
  T(s, "Better,\nnot more.", { x: 0.5, y: 1.55, w: 6.5, h: 1.6, fontSize: 44, bold: true, color: C.background1, fontFace: THEME.headFontFace, valign: "middle" });
  T(s, "Two weekly routines, tools we already have,\nand numbers we can check in December.", { x: 0.5, y: 3.3, w: 6.2, h: 0.75, fontSize: 15, color: C.accent5 });
  T(s, "Silas Atuahene Appiah  |  Digital Marketing Officer  |  5 October 2026", { x: 0.5, y: 4.95, w: 7, h: 0.3, fontSize: 10, color: C.accent5 });
  s.addNotes("The brief asked for a better way of working, not simply more work. These two routines do that, and in December the numbers will show whether they worked. Thank you. I'm happy to take questions.");

  await pres.writeFile({ fileName: OUT });
  const { applyTheme } = require(path.join(SKILL, "scripts/apply_theme.js"));
  await applyTheme(OUT, THEME);
  console.log("Wrote", OUT);
}

build().catch((e) => { console.error(e); process.exit(1); });
