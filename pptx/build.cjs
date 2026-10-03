/* Builds pptx/AI-Agents.pptx: the PowerPoint twin of index.html.
   Same content, same black/white/gray system, native PowerPoint motion:
   a fade transition on every slide, then each element fades in on its own
   timing (no clicks needed).

   Run:  npm install && npm run pptx
   Layout is authored on the web deck's 1920×1080 grid: 144 px = 1 inch, 2 px = 1 pt. */

const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const JSZip = require("jszip");

const ROOT = path.join(__dirname, "..");
const OUT = path.join(__dirname, "AI-Agents.pptx");

global.window = {};
require(path.join(ROOT, "assets/icons.js"));
const ICONS = global.window.ICONS;

// --------------------------------------------------------------------------- tokens
const C = {
  black: "0A0A0A", white: "FFFFFF", paper: "FAFAFA",
  g100: "F0F0F0", g200: "E3E3E3", g300: "CFCFCF", g400: "A8A8A8",
  g500: "8A8A8A", g600: "636363", g800: "2A2A2A", dim: "808080",
};
const SANS = "Arial";
const SERIF = "Georgia";
const px = (v) => v / 144;           // px on the 1920 grid → inches
const pt = (v) => v / 2;             // px font size → points

// --------------------------------------------------------------------------- helpers
let animSeq = 0;
const nameFor = (delay, label) => (delay == null ? label || undefined : `a${delay}_${++animSeq}`);

function text(slide, runs, o) {
  const opts = {
    x: px(o.x), y: px(o.y), w: px(o.w), h: px(o.h),
    fontFace: o.font || SANS, fontSize: pt(o.size || 30), color: o.color || C.black,
    italic: !!o.italic, bold: !!o.bold, align: o.align || "left", valign: o.valign || "top",
    margin: 0, isTextBox: true, fit: "none",
    lineSpacingMultiple: o.lh || 1.1, charSpacing: o.cs || 0,
    objectName: nameFor(o.delay, o.label),
  };
  if (o.para) opts.paraSpaceAfter = o.para;
  slide.addText(runs, opts);
}
function label(slide, str, o) {
  text(slide, str.toUpperCase(), { size: 20, cs: 1.6, color: C.g600, h: 30, ...o });
}
function rect(slide, o) {
  const shape = o.radius ? "roundRect" : "rect";
  slide.addShape(shape, {
    x: px(o.x), y: px(o.y), w: px(o.w), h: px(o.h),
    fill: o.fill ? { color: o.fill } : { type: "none" },
    line: o.line ? { color: o.line, width: o.lw || 1 } : { type: "none" },
    rectRadius: o.radius ? px(o.radius) : undefined,
    objectName: nameFor(o.delay, o.label),
  });
}
function ellipse(slide, o) {
  slide.addShape("ellipse", {
    x: px(o.x), y: px(o.y), w: px(o.w), h: px(o.h),
    fill: o.fill ? { color: o.fill } : { type: "none" },
    line: o.line ? { color: o.line, width: o.lw || 1 } : { type: "none" },
    objectName: nameFor(o.delay, o.label),
  });
}
function line(slide, x1, y1, x2, y2, o = {}) {
  slide.addShape("line", {
    x: px(Math.min(x1, x2)), y: px(Math.min(y1, y2)),
    w: px(Math.max(Math.abs(x2 - x1), 0.01)), h: px(Math.max(Math.abs(y2 - y1), 0.01)),
    line: { color: o.color || C.g400, width: o.lw || 1 },
    objectName: nameFor(o.delay, o.label),
  });
}

const iconCache = {};
async function iconPng(name, color) {
  const key = name + color;
  if (!iconCache[key]) {
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="256" height="256" fill="none" stroke="#${color}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">${ICONS[name]}</svg>`;
    const buf = await sharp(Buffer.from(svg)).png().toBuffer();
    iconCache[key] = "image/png;base64," + buf.toString("base64");
  }
  return iconCache[key];
}
async function icon(slide, name, o) {
  slide.addImage({ data: await iconPng(name, o.color || C.black), x: px(o.x), y: px(o.y), w: px(o.s), h: px(o.s), objectName: nameFor(o.delay, o.label) });
}
async function shot(slide, file, o) {
  const buf = await sharp(path.join(ROOT, "assets/img", file)).png().toBuffer();
  slide.addImage({ data: "image/png;base64," + buf.toString("base64"), x: px(o.x), y: px(o.y), w: px(o.w), h: px(o.h), rounding: false, objectName: nameFor(o.delay, o.label) });
}

const TOTAL = 14;
function chrome(slide, n, section, dark) {
  if (n !== 1 && n !== TOTAL) {
    text(slide, section.toUpperCase(), { x: 144, y: 1006, w: 900, h: 28, size: 18, cs: 1.2, color: dark ? C.dim : C.g400, label: "footer-section" });
    text(slide, [
      { text: String(n).padStart(2, "0"), options: { color: dark ? C.white : C.black } },
      { text: " / " + TOTAL, options: { color: dark ? C.dim : C.g400 } },
    ], { x: 1476, y: 1006, w: 300, h: 28, size: 18, cs: 1.2, align: "right", label: "footer-count" });
  }
  rect(slide, { x: 0, y: 1076, w: 1920 * (n / TOTAL), h: 4, fill: dark ? C.white : C.black, label: "progress" });
}
function base(pres, dark) {
  const s = pres.addSlide();
  s.background = { color: dark ? C.black : C.paper };
  return s;
}
// Two-part headline: sans line(s) + Georgia italic accent
function headline(slide, a, b, o = {}) {
  const size = o.size || 96, x = o.x || 144, y = o.y || 128, w = o.w || 1632;
  if (o.inline) {
    text(slide, [
      { text: a + " ", options: { fontFace: SANS } },
      { text: b, options: { fontFace: SERIF, italic: true, color: o.accentColor } },
    ], { x, y, w, h: size * 1.25, size, color: o.color, delay: 0, cs: -2 });
    return;
  }
  text(slide, a, { x, y, w, h: size * 1.15, size, color: o.color, delay: 0, cs: -2 });
  text(slide, b, { x, y: y + size * 1.02, w, h: size * 1.25, size: size * 1.04, font: SERIF, italic: true, color: o.color, delay: 120 });
}

// --------------------------------------------------------------------------- deck
async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 × 7.5 in = 1920 × 1080 on our grid
  pres.title = "AI Agents: a simple guide";
  pres.subject = "From chatting with AI to working with AI";
  pres.theme = { headFontFace: SANS, bodyFontFace: SANS };

  // 01 · Cover ------------------------------------------------------------
  {
    const s = base(pres, true);
    label(s, "A simple guide", { x: 144, y: 112, w: 600, color: C.g500, delay: 0 });
    label(s, "For the HR team", { x: 1176, y: 112, w: 600, color: C.g500, align: "right", delay: 0 });
    text(s, "AI", { x: 136, y: 200, w: 1200, h: 320, size: 300, color: C.white, cs: -6, lh: 0.9, delay: 150 });
    text(s, "Agents", { x: 136, y: 470, w: 1300, h: 360, size: 316, font: SERIF, italic: true, color: C.white, lh: 0.9, delay: 350 });
    text(s, [
      { text: "From chatting with AI", options: { breakLine: true } },
      { text: "to " },
      { text: "working", options: { fontFace: SERIF, italic: true, color: C.white } },
      { text: " with AI." },
    ], { x: 144, y: 832, w: 900, h: 140, size: 48, color: C.g500, lh: 1.25, delay: 700 });
    text(s, "Press → to begin".toUpperCase(), { x: 1376, y: 930, w: 400, h: 30, size: 20, cs: 1.6, color: C.dim, align: "right", delay: 1100 });
    chrome(s, 1, "AI Agents", true);
    s.addNotes("Today I want to explain AI agents in simple terms, show what I have built with them, and suggest how we could use them in HR.");
  }

  // 02 · What is an AI agent ------------------------------------------------
  {
    const s = base(pres, false);
    headline(s, "An AI that", "gets things done", { y: 110 });
    text(s, [
      { text: "Normal AI " , options: { color: C.g600 } },
      { text: "answers your question.", options: { color: C.black, bold: true } },
    ], { x: 144, y: 380, w: 760, h: 60, size: 34, delay: 400 });
    text(s, [
      { text: "An AI agent ", options: { color: C.g600 } },
      { text: "takes a goal", options: { color: C.black, bold: true } },
      { text: ", works on it step by step, uses tools, and checks its own result.", options: { color: C.g600 } },
    ], { x: 144, y: 456, w: 760, h: 110, size: 34, lh: 1.3, delay: 520 });
    line(s, 144, 640, 904, 640, { color: C.g200, delay: 700 });
    text(s, "Think of it as an AI employee, not an AI search box.", { x: 144, y: 676, w: 760, h: 140, size: 52, font: SERIF, italic: true, lh: 1.1, delay: 760 });
    // the loop
    const cx = 1430, cy = 560, r = 250;
    ellipse(s, { x: cx - r, y: cy - r, w: r * 2, h: r * 2, line: C.black, lw: 1.5, delay: 400 });
    const nodes = [["Plan", cx, cy - r], ["Act", cx + r, cy], ["Check", cx, cy + r], ["Improve", cx - r, cy]];
    nodes.forEach(([t, x, y], k) => {
      ellipse(s, { x: x - 72, y: y - 72, w: 144, h: 144, fill: C.paper, line: C.black, lw: 1.5, delay: 600 + k * 200 });
      text(s, t, { x: x - 72, y: y - 72, w: 144, h: 144, size: 30, bold: true, align: "center", valign: "middle", delay: 600 + k * 200 });
    });
    ellipse(s, { x: cx - 105, y: cy - 105, w: 210, h: 210, fill: C.black, delay: 1500 });
    text(s, "Your goal", { x: cx - 105, y: cy - 105, w: 210, h: 210, size: 46, font: SERIF, italic: true, color: C.white, align: "center", valign: "middle", lh: 0.95, delay: 1500 });
    chrome(s, 2, "What is an AI agent", false);
    s.addNotes("A chatbot answers and stops. An agent gets a goal and loops: plan, act with tools, check the result, improve, until the job is done.");
  }

  // 03 · Chat AI vs AI agent ----------------------------------------------
  {
    const s = base(pres, false);
    text(s, [
      { text: "Chat AI " }, { text: "vs", options: { fontFace: SERIF, italic: true, color: C.g400 } }, { text: " AI Agent" },
    ], { x: 144, y: 120, w: 1632, h: 120, size: 96, cs: -2, delay: 0 });
    const top = 290, h = 556;
    rect(s, { x: 144, y: top, w: 832, h, fill: C.white, line: C.g200, radius: 24, delay: 200 });
    rect(s, { x: 944, y: top, w: 832, h, fill: C.black, radius: 24, delay: 350 });
    const col = (x, dark, kicker, head, sub, steps, d0) => {
      label(s, kicker, { x: x + 64, y: top + 60, w: 700, color: dark ? C.g500 : C.g600, delay: d0 });
      text(s, head, { x: x + 64, y: top + 96, w: 720, h: 70, size: 52, bold: true, color: dark ? C.white : C.black, delay: d0 });
      text(s, sub, { x: x + 64, y: top + 170, w: 720, h: 40, size: 26, color: dark ? C.g500 : C.g600, delay: d0 });
      steps.forEach((st, k) => {
        const y = top + 240 + k * 100;
        line(s, x + 64, y, x + 768, y, { color: dark ? C.g800 : C.g200, delay: d0 + 200 + k * 200 });
        text(s, "0" + (k + 1), { x: x + 64, y: y + 34, w: 60, h: 30, size: 20, bold: true, color: dark ? C.dim : C.g400, delay: d0 + 200 + k * 200 });
        text(s, st, { x: x + 136, y: y + 24, w: 620, h: 52, size: 36, color: dark ? C.white : (k === 2 ? C.g600 : C.black), delay: d0 + 200 + k * 200 });
      });
    };
    col(144, false, "LLM / Chat AI", "The brain", "ChatGPT · Claude · Gemini", ["You ask a question", "It gives an answer", "It stops and waits"], 300);
    col(944, true, "AI Agent", "Brain + hands + a schedule", "Hermes · Claude Code · Codex", ["You give it a goal", "It plans and does the work", "It reports back when done"], 450);
    text(s, [
      { text: "LLM", options: { bold: true } }, { text: "  =  ", options: { color: C.g400 } },
      { text: "the brain.", options: { fontFace: SERIF, italic: true } },
      { text: "          Agent", options: { bold: true } }, { text: "  =  ", options: { color: C.g400 } },
      { text: "brain + hands + a schedule.", options: { fontFace: SERIF, italic: true } },
    ], { x: 144, y: 900, w: 1632, h: 60, size: 38, delay: 1300 });
    chrome(s, 3, "Chat AI vs AI agent", false);
    s.addNotes("ChatGPT, Claude and Gemini are the brains. An agent like Hermes, Claude Code or Codex wraps a brain with tools and a schedule, so it can actually do the work.");
  }

  // 04 · What agents can do -------------------------------------------------
  {
    const s = base(pres, false);
    headline(s, "What an agent can do", "that chat can’t", { y: 110 });
    line(s, 144, 400, 1776, 400, { color: C.black, delay: 200 });
    const caps = [
      ["clock", "Works on a schedule", "Runs tasks at set times, without being asked"],
      ["wrench", "Uses tools", "Browses the web, edits files, runs apps"],
      ["cpu", "Thinks in steps", "Breaks a goal into steps and decides what’s next"],
      ["user", "Has a persona", "A fixed role, tone and set of rules"],
      ["box", "Learns skills", "Install new abilities, like apps on a phone"],
      ["share-2", "Works in a team", "Hands work to other agents in a workflow"],
    ];
    line(s, 144, 686, 1776, 686, { color: C.g200, delay: 300 });
    line(s, 688, 400, 688, 972, { color: C.g200, delay: 300 });
    line(s, 1232, 400, 1232, 972, { color: C.g200, delay: 300 });
    for (let k = 0; k < caps.length; k++) {
      const [ic, h, b] = caps[k];
      const x = 144 + (k % 3) * 544 + (k % 3 ? 48 : 0), y = 440 + Math.floor(k / 3) * 286, d = 350 + k * 110;
      await icon(s, ic, { x, y, s: 56, delay: d });
      text(s, h, { x, y: y + 84, w: 480, h: 56, size: 40, bold: true, delay: d });
      text(s, b, { x, y: y + 150, w: 440, h: 90, size: 28, color: C.g600, lh: 1.35, delay: d });
    }
    chrome(s, 4, "What agents can do", false);
    s.addNotes("These six things separate an agent from a chatbot: schedules, tools, step-by-step thinking, a persona, installable skills, and teamwork with other agents.");
  }

  // 05 · Building blocks ------------------------------------------------------
  {
    const s = base(pres, false);
    headline(s, "What every agent", "is made of", { y: 110 });
    rect(s, { x: 144, y: 812, w: 1632, h: 158, fill: C.black, radius: 24, delay: 200 });
    await icon(s, "cpu", { x: 196, y: 863, s: 56, color: C.white, delay: 300 });
    text(s, "Model", { x: 276, y: 858, w: 240, h: 70, size: 56, bold: true, color: C.white, delay: 300 });
    text(s, "THE BRAIN UNDERNEATH", { x: 476, y: 882, w: 300, h: 30, size: 19, cs: 1.6, color: C.g500, delay: 300 });
    text(s, [
      { text: "Picked per job: ", options: { color: C.g500 } },
      { text: "a small, fast model", options: { color: C.white, bold: true } },
      { text: " for simple tasks, ", options: { color: C.g500 } },
      { text: "a stronger one", options: { color: C.white, bold: true } },
      { text: " for judgment and creative work.", options: { color: C.g500 } },
    ], { x: 800, y: 850, w: 930, h: 90, size: 28, lh: 1.35, valign: "middle", delay: 400 });
    const blocks = [["heart", "Soul", "Who it is", "Personality, role and rules, written in one file"],
      ["zap", "Skills", "What it can do", "Step-by-step playbooks and tools it can use"],
      ["database", "Memory", "What it remembers", "Notes about you, your work and past tasks"]];
    for (let k = 0; k < 3; k++) {
      const [ic, h, kick, b] = blocks[k];
      const x = 144 + k * 551, w = 530, y = 390, d = 650 + k * 140;
      rect(s, { x, y, w, h: 400, fill: C.white, line: C.g300, radius: 24, delay: d });
      await icon(s, ic, { x: x + 48, y: y + 44, s: 56, delay: d });
      label(s, kick, { x: x + 200, y: y + 56, w: w - 248, align: "right", delay: d });
      text(s, h, { x: x + 48, y: y + 196, w: w - 96, h: 76, size: 56, bold: true, delay: d });
      text(s, b, { x: x + 48, y: y + 286, w: w - 96, h: 90, size: 28, color: C.g600, lh: 1.35, delay: d });
    }
    chrome(s, 5, "The building blocks", false);
    s.addNotes("An agent is a model plus three things. The model is the brain, and you choose it per job: cheap and fast for simple work, stronger for judgment. Soul is its character and job description, Skills are its playbooks, Memory is its experience with you.");
  }

  // 06 · What I am building -------------------------------------------------
  {
    const s = base(pres, true);
    text(s, "My own", { x: 144, y: 300, w: 900, h: 140, size: 128, color: C.white, cs: -3, delay: 0 });
    text(s, "AI team", { x: 144, y: 430, w: 900, h: 160, size: 136, font: SERIF, italic: true, color: C.white, delay: 150 });
    text(s, [
      { text: "8 agents", options: { color: C.white, bold: true } },
      { text: ", running 24/7 on one small server. I manage them from Discord.", options: { color: C.g500 } },
    ], { x: 144, y: 640, w: 820, h: 170, size: 44, lh: 1.3, delay: 500 });
    text(s, "8", { x: 1140, y: 40, w: 680, h: 960, size: 1000, font: SERIF, italic: true, color: C.white, align: "center", valign: "middle", lh: 1, delay: 250 });
    chrome(s, 6, "What I am building", true);
    s.addNotes("This is what I have been building: a small team of AI agents that run all day on one cheap server. I talk to them through Discord, like a team chat.");
  }

  // 07 · Team structure ------------------------------------------------------
  {
    const s = base(pres, false);
    headline(s, "How my team", "is structured", { inline: true, y: 120 });
    rect(s, { x: 905, y: 268, w: 110, h: 62, fill: C.paper, line: C.black, lw: 1.5, radius: 31, delay: 100 });
    text(s, "Me", { x: 905, y: 268, w: 110, h: 62, size: 28, bold: true, align: "center", valign: "middle", delay: 100 });
    line(s, 960, 330, 960, 384, { delay: 300 });
    rect(s, { x: 844, y: 384, w: 232, h: 96, fill: C.black, radius: 20, delay: 300 });
    text(s, [{ text: "CEO", options: { breakLine: true } }, { text: "RUDI", options: { fontSize: 9, color: C.g500, charSpacing: 1 } }],
      { x: 844, y: 384, w: 232, h: 96, size: 40, bold: true, color: C.white, align: "center", valign: "middle", lh: 0.95, delay: 300 });
    text(s, "Receives my tasks,", { x: 470, y: 418, w: 350, h: 32, size: 22, color: C.g600, align: "right", delay: 600 });
    text(s, "delegates, reports back", { x: 1100, y: 418, w: 350, h: 32, size: 22, color: C.g600, delay: 600 });
    const crew = [["search", "Researcher", "Edward", "Daily news and markets"], ["pen-tool", "Designer", "Leonardo", "Visuals and UI/UX"],
      ["code", "Engineer", "Astra", "Builds and fixes the system"], ["trending-up", "Trading analyst", "Kimi", "Reviews my trading journal"],
      ["activity", "Health coach", "Alexa", "Routine and nutrition check-ins"], ["message-circle", "Marketing", "Biti", "Content and captions"],
      ["map", "Travel planner", "Qunce", "Trips and itineraries"]];
    const step = (1632 + 16) / 7, busY = 600;
    line(s, 960, 480, 960, busY, { delay: 650 });
    line(s, 144 + 36, busY, 144 + 6 * step + 36, busY, { delay: 750 });
    for (let k = 0; k < 7; k++) {
      const [ic, role, who, what] = crew[k];
      const x = 144 + k * step, d = 1000 + Math.abs(k - 3) * 90;
      line(s, x + 36, busY, x + 36, 704, { delay: 850 });
      ellipse(s, { x, y: 704, w: 72, h: 72, fill: C.paper, line: C.black, lw: 1.5, delay: d });
      await icon(s, ic, { x: x + 20, y: 724, s: 32, delay: d });
      text(s, role, { x, y: 800, w: step - 16, h: 40, size: 28, bold: true, delay: d });
      text(s, who, { x, y: 838, w: step - 16, h: 36, size: 26, font: SERIF, italic: true, color: C.g600, delay: d });
      text(s, what, { x, y: 880, w: step - 24, h: 70, size: 21, color: C.g600, lh: 1.3, delay: d });
    }
    chrome(s, 7, "My AI team", false);
    s.addNotes("I give tasks to my CEO agent. It breaks them down, hands them to the right specialist, then reports back to me. Each specialist has its own role and personality.");
  }

  // 08 · The office ------------------------------------------------------------
  {
    const s = base(pres, false);
    headline(s, "Their office,", "live", { y: 220, size: 100, w: 580 });
    text(s, [
      { text: "A live dashboard built by my engineer agent. Every agent has a desk, so I can see at a glance ", options: { color: C.g600 } },
      { text: "who is working", options: { bold: true } },
      { text: " and who is on standby.", options: { color: C.g600 } },
    ], { x: 144, y: 470, w: 560, h: 210, size: 30, lh: 1.42, delay: 300 });
    const leg = [["Working", C.black, null], ["Queued", C.g400, null], ["Standby", null, C.g400]];
    leg.forEach(([t, fill, ln], k) => {
      ellipse(s, { x: 144, y: 738 + k * 52, w: 14, h: 14, fill, line: ln, lw: 1.5, delay: 500 + k * 80 });
      text(s, t, { x: 180, y: 728 + k * 52, w: 300, h: 36, size: 26, bold: true, delay: 500 + k * 80 });
    });
    await shot(s, "office.webp", { x: 804, y: 230, w: 972, h: 972 * 838 / 1311, delay: 200 });
    chrome(s, 8, "The office", false);
    s.addNotes("This is a live office view my engineer agent built for the team. Each character is an agent; when one is working on a task you can see it move to its desk.");
  }

  // 09 · How work flows ----------------------------------------------------------
  {
    const s = base(pres, false);
    headline(s, "One request,", "a whole team", { inline: true, y: 120 });
    label(s, "Me", { x: 144, y: 345, w: 80, delay: 200 });
    rect(s, { x: 210, y: 300, w: 1300, h: 116, fill: C.black, radius: 40, delay: 200 });
    text(s, "“Make a visual and a caption for tomorrow’s post”", { x: 250, y: 300, w: 1240, h: 116, size: 44, font: SERIF, italic: true, color: C.white, valign: "middle", delay: 200 });
    const steps = [["compass", "CEO", "Plans and delegates"], ["pen-tool", "Designer", "Designs the visual"], ["message-circle", "Marketing", "Writes the caption"],
      ["square-check", "CEO", "Checks it all together"], ["thumbs-up", "Me", "Approves before posting"]];
    const colW = 1632 / 5;
    line(s, 144 + 36, 536, 144 + 4 * colW + 36, 536, { color: C.black, lw: 1.5, delay: 600 });
    for (let k = 0; k < 5; k++) {
      const [ic, who, what] = steps[k];
      const x = 144 + k * colW, d = 700 + k * 380, last = k === 4;
      ellipse(s, { x, y: 500, w: 72, h: 72, fill: last ? C.black : C.paper, line: C.black, lw: 1.5, delay: d });
      await icon(s, ic, { x: x + 20, y: 520, s: 32, color: last ? C.white : C.black, delay: d });
      text(s, who, { x, y: 612, w: colW - 32, h: 50, size: 36, bold: true, delay: d });
      text(s, what, { x, y: 670, w: colW - 40, h: 80, size: 28, color: C.g600, lh: 1.3, delay: d });
    }
    rect(s, { x: 144, y: 862, w: 1632, h: 104, fill: C.g100, radius: 20, delay: 2600 });
    await icon(s, "clock", { x: 184, y: 894, s: 40, color: C.g600, delay: 2600 });
    text(s, [
      { text: "On autopilot every day: ", options: { bold: true, color: C.black } },
      { text: "the researcher scans the news  →  the CEO sends me a daily report. ", options: { color: C.g600 } },
      { text: "No prompt needed.", options: { bold: true, color: C.black } },
    ], { x: 252, y: 862, w: 1490, h: 104, size: 28, valign: "middle", delay: 2600 });
    chrome(s, 9, "How work flows", false);
    s.addNotes("One message from me becomes a small project: the CEO plans it, the designer makes the visual, marketing writes the caption, the CEO checks it, and I approve. Some work also runs on its own every day.");
  }

  // 10 · Real output -------------------------------------------------------------
  {
    const s = base(pres, false);
    headline(s, "Real output,", "every morning", { inline: true, y: 120 });
    const w = 792;
    const figs = [["Researcher", "Edward", "07:00 news brief", "news-brief.webp", 817 / 1204], ["CEO", "Rudi", "Daily report on everyone’s work", "daily-report.webp", 858 / 1228]];
    for (let k = 0; k < 2; k++) {
      const [role, who, kick, file, ar] = figs[k];
      const x = 144 + k * (w + 48), d = 200 + k * 200;
      text(s, [{ text: role + " · ", options: {} }, { text: who, options: { fontFace: SERIF, italic: true, color: C.g600 } }], { x, y: 300, w: 420, h: 44, size: 30, bold: false, delay: d });
      label(s, kick, { x: x + 300, y: 308, w: w - 300, align: "right", delay: d });
      await shot(s, file, { x, y: 362, w, h: w * ar, delay: d });
    }
    chrome(s, 10, "Agents at work", false);
    s.addNotes("Left: the researcher posts a news brief at 7 AM with the day's key events. Right: the CEO sends me one daily report summarising what every agent did.");
  }

  // 11 · Notes to journal ----------------------------------------------------------
  {
    const s = base(pres, false);
    headline(s, "From my notes", "to a clean journal", { inline: true, y: 120 });
    label(s, "1 · I send my notes + chart", { x: 144, y: 340, w: 640, delay: 200 });
    await shot(s, "trading-notes.webp", { x: 144, y: 384, w: 640, h: 640 * 564 / 1227, delay: 200 });
    await icon(s, "arrow-right", { x: 806, y: 520, s: 56, delay: 700 });
    label(s, "2 · The agent reviews and logs it", { x: 884, y: 340, w: 892, delay: 900 });
    await shot(s, "trading-journal.webp", { x: 884, y: 384, w: 892, h: 892 * 827 / 1789, delay: 900 });
    text(s, [{ text: "Trading analyst:  ", options: { bold: true, color: C.black } }, { text: "turns a casual chat message into a structured journal entry, with a review of the decision.", options: { color: C.g600 } }],
      { x: 144, y: 880, w: 1632, h: 50, size: 30, delay: 1300 });
    chrome(s, 11, "Agents at work", false);
    s.addNotes("I just chat with the trading analyst like a person: what I saw and what I did, plus a chart. It turns that into a clean journal entry in Notion, with the plan, the reason, and an honest review of the decision.");
  }

  // 12 · Dojima ---------------------------------------------------------------------
  {
    const s = base(pres, true);
    text(s, "Dojima", { x: 140, y: 100, w: 600, h: 190, size: 160, color: C.white, cs: -6, delay: 0 });
    text(s, [
      { text: "A trading-practice web app I built with Claude Code, ", options: { color: C.g500 } },
      { text: "without writing the code myself.", options: { color: C.white, bold: true } },
    ], { x: 144, y: 320, w: 540, h: 180, size: 34, lh: 1.3, delay: 300 });
    const layers = [["monitor", "Frontend", "What you see: charts, journal, themes"], ["server", "Backend", "The engine: orders, replay, login"], ["database", "Database", "The memory: accounts and trades"]];
    line(s, 144, 618, 704, 618, { color: C.g800, delay: 600 });
    for (let k = 0; k < 3; k++) {
      const [ic, h, b] = layers[k];
      const y = 618 + k * 118, d = 700 + k * 120;
      await icon(s, ic, { x: 144, y: y + 30, s: 40, color: C.white, delay: d });
      text(s, h, { x: 218, y: y + 22, w: 480, h: 40, size: 30, bold: true, color: C.white, delay: d });
      text(s, b, { x: 218, y: y + 62, w: 480, h: 36, size: 23, color: C.g500, delay: d });
      line(s, 144, y + 118, 704, y + 118, { color: C.g800, delay: d });
    }
    await shot(s, "dojima.webp", { x: 784, y: 270, w: 992, h: 992 * 906 / 1919, delay: 250 });
    text(s, [
      { text: "I ", options: { color: C.g500 } }, { text: "describe it", options: { color: C.white, bold: true } },
      { text: "   →   Claude Code ", options: { color: C.g500 } }, { text: "builds it", options: { color: C.white, bold: true } },
      { text: "   →   it ", options: { color: C.g500 } }, { text: "goes live", options: { color: C.white, bold: true } },
    ], { x: 784, y: 790, w: 992, h: 50, size: 30, delay: 1300 });
    chrome(s, 12, "Built with an agent", true);
    s.addNotes("Dojima is a full trading-practice app with paper money. I am not a programmer: I describe what I want, Claude Code (an AI agent) writes the code and publishes it. It has a frontend you see, a backend that runs the logic, and a database that stores the data.");
  }

  // 13 · AI agents for HR -----------------------------------------------------------------
  {
    const s = base(pres, false);
    headline(s, "An AI team", "for HR", { y: 110, size: 104, w: 560 });
    text(s, "One orchestrator that understands HR, with a specialist for each job.", { x: 144, y: 380, w: 500, h: 170, size: 34, color: C.g600, lh: 1.3, delay: 300 });
    await icon(s, "square-check", { x: 144, y: 934, s: 28, color: C.g600, delay: 1600 });
    text(s, "HR still reviews and approves everything.", { x: 188, y: 930, w: 640, h: 36, size: 24, color: C.g600, delay: 1600 });
    // map
    const hx = 760, headY = 560;
    rect(s, { x: hx, y: 300, w: 170, h: 64, fill: C.paper, line: C.black, lw: 1.5, radius: 32, delay: 200 });
    text(s, "HR team", { x: hx, y: 300, w: 170, h: 64, size: 28, bold: true, align: "center", valign: "middle", delay: 200 });
    line(s, hx + 40, 364, hx + 40, headY - 64, { delay: 350 });
    rect(s, { x: hx, y: headY - 64, w: 220, h: 128, fill: C.black, radius: 20, delay: 450 });
    text(s, [{ text: "Head of HR", options: { breakLine: true } }, { text: "agent", options: { fontFace: SERIF, italic: true, bold: false, color: C.g500 } }],
      { x: hx + 26, y: headY - 64, w: 190, h: 128, size: 32, bold: true, color: C.white, valign: "middle", lh: 1.1, delay: 450 });
    const specs = [["file-text", "Job descriptions", "Drafts JDs from a short brief"], ["user-check", "Recruitment", "Screening notes and interview kits"],
      ["user-plus", "Onboarding", "Checklists and first-week plans"], ["chart-no-axes-column", "HR reports", "Weekly and monthly summaries"],
      ["book-open", "Policy Q&A", "Answers staff questions from the handbook"], ["award", "Training", "Learning plans and materials"]];
    const busX = hx + 280, rowH = 98, y0 = headY - (rowH * 5) / 2;
    line(s, hx + 220, headY, busX, headY, { delay: 650 });
    line(s, busX, y0, busX, y0 + rowH * 5, { delay: 700 });
    for (let k = 0; k < 6; k++) {
      const [ic, h, b] = specs[k];
      const cy = y0 + k * rowH, x = busX + 60, d = 900 + k * 80;
      line(s, busX, cy, x, cy, { delay: 800 });
      ellipse(s, { x, y: cy - 28, w: 56, h: 56, fill: C.paper, line: C.black, lw: 1.5, delay: d });
      await icon(s, ic, { x: x + 14, y: cy - 14, s: 28, delay: d });
      text(s, h, { x: x + 84, y: cy - 22, w: 280, h: 44, size: 32, bold: true, valign: "middle", delay: d });
      text(s, b, { x: x + 380, y: cy - 34, w: 340, h: 68, size: 24, color: C.g600, valign: "middle", lh: 1.25, delay: d });
    }
    chrome(s, 13, "AI agents for HR", false);
    s.addNotes("Same idea as my team, applied to HR. One Head of HR agent receives requests from the HR team and delegates to specialists: job descriptions, recruitment, onboarding, reports, policy questions and training. HR still reviews and approves everything.");
  }

  // 14 · Close -------------------------------------------------------------------------
  {
    const s = base(pres, true);
    label(s, "Thank you", { x: 144, y: 112, w: 600, color: C.g500, delay: 0 });
    label(s, "AI Agents · A simple guide", { x: 1176, y: 112, w: 600, color: C.g500, align: "right", delay: 0 });
    text(s, "Start", { x: 136, y: 250, w: 1200, h: 260, size: 240, color: C.white, cs: -5, lh: 0.9, delay: 150 });
    text(s, "small.", { x: 136, y: 470, w: 1200, h: 290, size: 252, font: SERIF, italic: true, color: C.white, lh: 0.9, delay: 350 });
    ["One agent.", "One task.", "One week."].forEach((t, k) =>
      text(s, t, { x: 144 + k * 420, y: 876, w: 440, h: 100, size: 64, font: SERIF, italic: true, color: C.white, delay: 700 + k * 300 }));
    text(s, "Questions?", { x: 1426, y: 885, w: 350, h: 70, size: 56, bold: true, color: C.white, align: "right", delay: 1900 });
    chrome(s, 14, "Thank you", true);
    s.addNotes("My suggestion: start small with one agent and one task, see it working within a week, then grow from there.");
  }

  // ------------------------------------------------------------- motion pass
  const buf = await pres.write({ outputType: "nodebuffer" });
  const zip = await JSZip.loadAsync(buf);
  const slideFiles = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
  for (const f of slideFiles) {
    let xml = await zip.file(f).async("string");
    const targets = [...xml.matchAll(/<p:nv(Sp|Pic)Pr>\s*<p:cNvPr id="(\d+)" name="a(\d+)_\d+"/g)]
      .map((m) => ({ id: m[2], delay: +m[3], sp: m[1] === "Sp" }));
    xml = xml.replace("</p:sld>", transitionXml() + timingXml(targets) + "</p:sld>");
    zip.file(f, xml);
  }
  fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log("Wrote", path.relative(ROOT, OUT));
}

function transitionXml() {
  return '<p:transition spd="slow"><p:fade/></p:transition>';
}

// Every tagged element fades in automatically after the transition, at its own delay.
function timingXml(targets) {
  if (!targets.length) return "";
  let id = 4;
  const effects = targets.map(({ id: spid, delay }) => {
    const a = ++id, b = ++id, c = ++id;
    return `<p:par><p:cTn id="${a}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="withEffect"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst>` +
      `<p:set><p:cBhvr><p:cTn id="${b}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
      `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${c}" dur="700"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animEffect>` +
      `</p:childTnLst></p:cTn></p:par>`;
  }).join("");
  const bld = targets.filter((t) => t.sp).map(({ id: spid }) => `<p:bldP spid="${spid}" grpId="0" animBg="1"/>`).join("");
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>` +
    `<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${effects}</p:childTnLst></p:cTn></p:par>` +
    `</p:childTnLst></p:cTn></p:par>` +
    `</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>` +
    `</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>${bld}</p:bldLst></p:timing>`;
}

build().catch((e) => { console.error(e); process.exit(1); });
