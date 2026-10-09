// Builds the final hackathon report as a Word document: docs/Final_Report_SOC_Agentic_Triage.docx
// Run: node tools/build_report.js   (needs the npm package "docx")
// Text marked [[TO FILL: ...]] is highlighted yellow in the output so the team can find and complete it.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType,
  BorderStyle, LevelFormat, HeadingLevel, Footer, Header, PageNumber, PageBreak, ImageRun, TableOfContents,
  PageOrientation, VerticalAlign, TabStopType,
} = require('docx');

const ROOT = path.join(__dirname, '..');
const OUT = path.join(ROOT, 'docs', 'Final_Report_SOC_Agentic_Triage.docx');
const INK = '1C2430', GREY = '5B6573', RULE = 'C9CED6', HEAD = 'EEF0F3', ACCENT = '2F6F4E', FONT = 'Calibri';
const W = 9026; // A4 portrait content width (1" margins), DXA
const WL = 13958; // A4 landscape content width

// ---------- inline markup: **bold**, [[TO FILL: ...]] highlighted, `code`
function runs(text, base = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\[\[[^\]]+\]\]|`[^`]+`)/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), font: FONT, size: 22, color: INK, ...base }));
    const t = m[0];
    if (t.startsWith('**')) out.push(new TextRun({ text: t.slice(2, -2), font: FONT, size: 22, color: INK, ...base, bold: true }));
    else if (t.startsWith('[[')) out.push(new TextRun({ text: '[' + t.slice(2, -2) + ']', font: FONT, size: 22, color: INK, ...base, highlight: 'yellow' }));
    else out.push(new TextRun({ text: t.slice(1, -1), font: 'Consolas', size: (base.size || 22) - 2, color: INK }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), font: FONT, size: 22, color: INK, ...base }));
  return out;
}
const P = (t, o = {}) => new Paragraph({ spacing: { after: 140, line: 288 }, alignment: o.align || AlignmentType.JUSTIFIED, children: runs(t, o.run) });
const HEADINGS = [];
const H1 = (t) => (HEADINGS.push([1, t]), new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, spacing: { before: 0, after: 200 }, children: [new TextRun({ text: t, font: FONT, size: 34, bold: true, color: INK })] }));
const H2 = (t) => (HEADINGS.push([2, t]), new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, spacing: { before: 260, after: 120 }, children: [new TextRun({ text: t, font: FONT, size: 26, bold: true, color: INK })] }));
const H3 = (t) => new Paragraph({ keepNext: true, spacing: { before: 180, after: 80 }, children: [new TextRun({ text: t, font: FONT, size: 22, bold: true, color: ACCENT })] });
const B = (t, lvl = 0) => new Paragraph({ numbering: { reference: 'bul', level: lvl }, spacing: { after: 70, line: 276 }, children: runs(t) });
const N = (t, ref = 'num') => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 70, line: 276 }, children: runs(t) });
const SP = () => new Paragraph({ spacing: { after: 60 }, children: [] });
const NOTE = (t) => new Paragraph({ spacing: { before: 60, after: 160, line: 276 }, shading: { type: ShadingType.CLEAR, fill: 'F4F5F7', color: 'auto' },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color: ACCENT, space: 8 } }, indent: { left: 160 }, children: runs(t, { size: 20 }) });

const bd = { style: BorderStyle.SINGLE, size: 4, color: RULE };
const borders = { top: bd, bottom: bd, left: bd, right: bd };
function T(widths, rows, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const fs = o.size || 19;
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map((cells, ri) => new TableRow({
      tableHeader: ri === 0, cantSplit: true,
      children: cells.map((c, ci) => new TableCell({
        width: { size: widths[ci], type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.TOP,
        shading: ri === 0 ? { type: ShadingType.CLEAR, fill: HEAD, color: 'auto' } : undefined,
        margins: { top: 70, bottom: 70, left: 100, right: 100 },
        children: String(c).split('\n').map((line) => new Paragraph({ spacing: { after: 20, line: 252 }, children: runs(line, { size: fs, bold: ri === 0 ? true : undefined }) })),
      })),
    })),
  });
}
let figNo = 0;
function FIG(file, widthIn, caption) {
  const p = path.join(ROOT, 'docs', 'diagrams', file);
  const dims = { 'fig0_causal_loop.png': [2609, 2005], 'fig1_as_is.png': [5705, 756], 'fig1_as_is_print.png': [1507, 2550], 'fig2_to_be.png': [2561, 2709], 'fig3_architecture.png': [2189, 1526] }[file];
  const w = Math.round(widthIn * 96), h = Math.round(w * dims[1] / dims[0]);
  figNo += 1;
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120, after: 60 }, children: [new ImageRun({ type: 'png', data: fs.readFileSync(p), transformation: { width: w, height: h } })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: runs(caption, { size: 18, color: GREY, italics: true }) }),
  ];
}


// ---- screenshots: docs/screenshots/FigNN.(jpg|png); a missing file becomes a yellow placeholder
function imgSize(buf) {
  if (buf[0] === 0x89) return [buf.readUInt32BE(16), buf.readUInt32BE(20)];
  let i = 2;
  while (i < buf.length) {
    const marker = buf[i + 1], len = buf.readUInt16BE(i + 2);
    if (marker >= 0xC0 && marker <= 0xCF && ![0xC4, 0xC8, 0xCC].includes(marker)) return [buf.readUInt16BE(i + 7), buf.readUInt16BE(i + 5)];
    i += 2 + len;
  }
  throw new Error('cannot read image size');
}

const SHOT_DIR = path.join(ROOT, 'docs', 'screenshots');
const shotBase = (k) => { const mm = String(k).match(/^(\d+)(.*)$/); if (!mm) return k; return 'Fig' + mm[1].padStart(2, '0') + mm[2]; };
const DIAG = { cld: 'fig0_causal_loop_r.png', asis: 'fig1_as_is_print_r.png', tobe: 'fig2_to_be_r.png', arch: 'fig3_architecture_r.png' };
const shotFile = (k) => DIAG[k] ? path.join(ROOT, 'docs', 'diagrams', DIAG[k]) : ['.png', '.jpg', '.jpeg'].map((e) => path.join(SHOT_DIR, shotBase(k) + e)).find((f) => fs.existsSync(f));
const SHOT_ORDER = ['cld', '4a', '4b', '5a', '5b', '5c', 'asis', 'tobe', 'arch', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30'];
const FIGNO = {}; { let n = 1; for (const k of SHOT_ORDER) if (shotFile(k)) FIGNO[k] = n++; }
const F = (k) => FIGNO[String(k)];
const figs = (...ks) => { const ns = ks.map(String).filter((k) => FIGNO[k]).map((k) => FIGNO[k]); return ns.length === 0 ? '—' : (ns.length === 1 ? 'Fig ' + ns[0] : 'Figs ' + ns.join(', ')); };
function SHOT(n, caption, widthIn = 6.2, labelNo = null) {
  const file = shotFile(n);
  if (!file) return [];
  const label = `Figure ${labelNo || F(n)}. ${caption}`;
  const buf = fs.readFileSync(file);
  const [iw, ih] = imgSize(buf);
  let w = Math.round(widthIn * 96), h = Math.round(w * ih / iw);
  const maxH = Math.round(7.4 * 96);
  if (h > maxH) { h = maxH; w = Math.round(h * iw / ih); }
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120, after: 60 }, children: [new ImageRun({ type: file.endsWith('.png') ? 'png' : 'jpg', data: buf, transformation: { width: w, height: h } })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: runs(label, { size: 18, color: GREY, italics: true }) }),
  ];
}

// ======================================================================= content
const ROLES = [
  ['25030241063', 'Aviral Bajpei', 'Solution architecture, n8n workflows and guardrails, governance mapping, final report'],
  ['25030241073', 'Jatin Balecha', 'Local deployment, console test runs, model benchmark, screenshots'],
  ['25030241085', 'Prakhar Aryan', 'n8n Cloud deployment, expert interview coordination, demonstration video'],
  ['25030241070', 'Garvit Saraswat', 'AS-IS and TO-BE process mapping, test cases and synthetic data'],
  ['25030241061', 'Arijit Paul', 'Problem research, STOIC prompting, market research, presentation'],
];
const CTR = (t, o = {}) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: o.after ?? 100, before: o.before ?? 0 }, children: runs(t, o.run) });
const lh = fs.readFileSync(path.join(ROOT, 'docs', 'screenshots', 'scit_letterhead.png'));
const nb = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: nb, bottom: nb, left: nb, right: nb, insideHorizontal: nb, insideVertical: nb };
const nameTable = new Table({
  width: { size: 5200, type: WidthType.DXA }, columnWidths: [2800, 2400], alignment: AlignmentType.CENTER, borders: noBorders,
  rows: ROLES.map(([prn, name]) => new TableRow({ children: [
    new TableCell({ width: { size: 2800, type: WidthType.DXA }, borders: noBorders, margins: { top: 30, bottom: 30, left: 80, right: 160 },
      children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: name, font: FONT, size: 24, bold: true, color: INK })] })] }),
    new TableCell({ width: { size: 2400, type: WidthType.DXA }, borders: noBorders, margins: { top: 30, bottom: 30, left: 160, right: 80 },
      children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: prn, font: FONT, size: 24, color: GREY })] })] }),
  ] })),
});
const RULE_P = (after = 280) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after }, indent: { left: 3600, right: 3600 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 1 } }, children: [] });
const cover = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 }, children: [new ImageRun({ type: 'png', data: lh, transformation: { width: Math.round(6.0 * 96), height: Math.round(6.0 * 96 * 284 / 1107) } })] }),
  CTR('AGENTIC AI DIGITAL TRANSFORMATION HACKATHON', { run: { size: 20, bold: true, color: ACCENT, characterSpacing: 40 }, after: 60 }),
  CTR('Cybersecurity Challenge', { run: { size: 22, color: GREY }, after: 480 }),
  CTR('Governed Agentic AI for SOC Alert Triage', { run: { size: 52, bold: true, color: INK }, after: 40 }),
  CTR('and Incident Investigation', { run: { size: 52, bold: true, color: INK }, after: 240 }),
  RULE_P(240),
  CTR('Final Report', { run: { size: 26, color: GREY }, after: 900 }),
  CTR('SUBMITTED TO', { run: { size: 18, bold: true, color: ACCENT, characterSpacing: 30 }, after: 60 }),
  CTR('Dr. Anshul Saxena', { run: { size: 28, bold: true, color: INK }, after: 560 }),
  CTR('SUBMITTED BY', { run: { size: 18, bold: true, color: ACCENT, characterSpacing: 30 }, after: 120 }),
  nameTable,
  CTR('MBA (IT Business Management) · Batch 2025–27', { run: { size: 22, color: INK }, before: 900, after: 40 }),
  CTR('Symbiosis Centre for Information Technology, Pune', { run: { size: 22, color: GREY }, after: 40 }),
  CTR('October 2026', { run: { size: 22, color: GREY }, after: 0 }),
];
const details = [
  new Paragraph({ pageBreakBefore: true, spacing: { after: 240 }, children: [new TextRun({ text: 'Project team and resources', font: FONT, size: 34, bold: true, color: INK })] }),
  new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: 'Project team', font: FONT, size: 26, bold: true, color: INK })] }),
  T([1600, 2100, 5326], [['PRN', 'Name', 'Role in project'], ...ROLES], { size: 20 }),
  SP(),
  new Paragraph({ spacing: { before: 240, after: 120 }, children: [new TextRun({ text: 'Project resources', font: FONT, size: 26, bold: true, color: INK })] }),
  T([2300, 6726], [
    ['Resource', 'Link'],
    ['Project website', 'https://aviralbajpei.github.io/soc-triage-website/'],
    ['Videos and documents', 'https://drive.google.com/drive/folders/1RpVaA3kFUzE1MyHtbhU2kA-EvTd-V9jx?usp=sharing'],
    ['Live console', '[[Live console link: to be added (password shared separately)]]'],
  ], { size: 20 }),
  SP(),
  new Paragraph({ spacing: { before: 240, after: 120 }, children: [new TextRun({ text: 'Submission details', font: FONT, size: 26, bold: true, color: INK })] }),
  T([2300, 6726], [
    ['Item', 'Detail'],
    ['Programme', 'MBA (IT Business Management), Batch 2025–27'],
    ['Institute', 'Symbiosis Centre for Information Technology, Pune'],
    ['Submitted to', 'Dr. Anshul Saxena'],
    ['Challenge', 'Agentic AI Digital Transformation Hackathon: Cybersecurity (SOC alert triage and incident investigation)'],
    ['Date', 'October 2026'],
  ], { size: 20 }),
];

let TOC_PAGES = {};
try { TOC_PAGES = JSON.parse(fs.readFileSync(process.env.SOC_TOC_PAGES || '', 'utf8')); } catch (e) { /* first pass */ }
function buildToc() {
  const out = [new Paragraph({ pageBreakBefore: true, spacing: { after: 240 }, children: [new TextRun({ text: 'Contents', font: FONT, size: 34, bold: true, color: INK })] })];
  for (const [lvl, t] of HEADINGS) {
    const pg = String(TOC_PAGES[t] ?? '00');
    out.push(new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 9026, leader: 'dot' }], indent: { left: lvl === 1 ? 0 : 360 },
      spacing: { before: lvl === 1 ? 120 : 0, after: 40 },
      children: [new TextRun({ text: t, font: FONT, size: lvl === 1 ? 22 : 20, bold: lvl === 1, color: INK }), new TextRun({ text: '\t' + pg, font: FONT, size: lvl === 1 ? 22 : 20, bold: lvl === 1, color: INK })] }));
  }
  return out;
}

const s1 = [
  H1('1. Executive summary'),
  P('Security Operations Centres (SOCs) receive far more alerts than their analysts can investigate. Most of the time spent on each alert goes into gathering context from separate tools, and decisions vary between analysts and shifts. Approvals for containment often happen by phone or chat with little record.'),
  P('We redesigned first-line alert triage as a **governed, multi-agent workflow**:'),
  B('**Five AI agents** (Alert Correlation, Threat Intelligence, Asset Context, Risk/Prioritisation, Incident Summary) assemble and assess the evidence for each alert, using open-weight language models.'),
  B('**Deterministic code**, not a model, validates input, scores risk, applies guardrails and decides the route.'),
  B('**Human analysts** approve every consequential action under a role matrix enforced on the server: Tier 1 up to MEDIUM, Tier 2 up to HIGH, SOC Lead for CRITICAL.'),
  B('**Every step is logged** with actor, model, prompt version and latency, so each decision can be reconstructed.'),
  P('The working proof of concept connects a Streamlit console to seven n8n workflows. It correctly handles normal, ambiguous and high-risk alerts, and stops prompt-injection and malformed alerts before any model sees them. It runs on a laptop with a one-click setup, can be shared through a password-protected link, and has been deployed to n8n Cloud.'),
  P('Every risk we identified is mapped to a guardrail, a governance requirement (NIST AI RMF 1.0, ISO/IEC 42001, ISO/IEC 27001:2022, DPDP Act 2023) and the exact place it is enforced in the prototype.'),
  P('We compared **three** open-weight models (GPT-OSS 20B, GPT-OSS 120B and Qwen 3.8 27B) on identical cases. Both GPT-OSS models reached the correct final risk class in every scored run, against 50% for Qwen 3.8 27B, and GPT-OSS 120B matched the expected route most often (83%), so it is our recommended model for a pilot. An industry expert, **Jatin Modi (Director, Commercial IQ)**, reviewed the design in a video interview; his findings are summarised in section 8.'),
];

const s2 = [
  H1('2. The problem'),
  P('A survey of 282 security leaders found that the median SOC handles about **960 alerts a day**, that roughly **40% are never investigated**, and that a full investigation takes about **70 minutes** on average. 61% of teams admitted ignoring alerts that later proved critical [R1]. In the 2025 SANS Detection and Response Survey, 73% of teams named false positives as their biggest detection challenge [R2].'),
  T([3000, 3013, 3013], [
    ['~960', '~40%', '~70 min'],
    ['alerts a day in the median SOC', 'of alerts never investigated', 'to fully investigate one alert'],
  ], { size: 22 }),
  SP(),
  P('**Problem statement:** how might a SOC triage every alert consistently and quickly, with evidence-backed recommendations, while keeping humans in control of every high-impact decision and leaving a complete audit trail?'),
];

const s3 = [
  H1('3. Stage 1: Structured thinking, prompt engineering and research'),
  H2('3.1 Critical thinking'),
  P('We separated claims from evidence and listed the assumptions behind each one, with how we would test it.'),
  T([2300, 2500, 2200, 2026], [
    ['Claim', 'Evidence', 'Assumption', 'How we test it'],
    ['Triage is the main bottleneck, not detection', 'Survey data on uninvestigated alerts and time per alert; internship experience of L1 monitoring on Wazuh', 'The same holds for mid-size Indian SOCs and MSSPs', 'Expert interview, Q1 and Q3'],
    ['Most triage time goes on gathering context, not judgement', 'Analysts pivot across SIEM, EDR, TI portals and CMDB for each alert', 'Context sources can be queried programmatically', 'PoC tool lookups; expert Q2 and Q4'],
    ['AI can do the gathering and first assessment', 'LLMs summarise and reason over structured evidence', 'Small open-weight models produce valid structured output often enough', 'Benchmark: JSON compliance, invented IDs'],
    ['Humans must control response actions', 'Containment errors cause outages or tip off attackers', 'Approval friction is acceptable if fast and well informed', 'HITL design; expert feedback'],
    ['An LLM recommendation can be trusted', '**Not assumed.** Models hallucinate and can be manipulated by text inside alerts', 'Code, not the model, must decide routing and scoring', 'Guardrail tests: injection, invented IDs'],
  ]),
  H3('Five Whys: why are critical alerts missed?'),
  N('Analysts do not reach them, or close them too quickly.', 'why'),
  N('The queue is longer than the team\'s capacity, so alerts are skimmed.', 'why'),
  N('Each alert takes long because context sits in five or more separate tools.', 'why'),
  N('There is no single case record that assembles evidence automatically.', 'why'),
  N('Tools were bought separately and triage is organised around people, not evidence.', 'why'),
  P('**Root cause:** evidence assembly is manual and fragmented. That is the step to automate, while judgement on high-impact actions stays with humans.'),
  P('We also checked our thinking against **automation bias** (analysts rubber-stamping AI advice), **solutionism** (we screened non-AI options too), **availability bias** (most alerts are benign, so the design must close benign alerts safely, not only escalate) and **adversarial thinking** (attackers can write into logs, so alert text is treated as data, never as instructions).'),
  H2('3.2 Systems thinking'),
  P('The causal loop diagram shows why the problem keeps getting worse, and where an agentic system intervenes.'),
  B('**R1, fatigue spiral (reinforcing):** alert volume → workload → fatigue → attrition → less experience → shallower investigation → more missed detections.'),
  B('**R2, noise spiral (reinforcing):** shallow investigation → poor feedback to detection tuning → more false positives → more volume.'),
  B('**B1, agentic triage (balancing):** automated evidence assembly lowers workload and raises consistency.'),
  B('**B2, automation-bias risk (balancing):** agentic triage can create over-trust, countered by HITL controls and showing the evidence behind each recommendation.'),
  ...SHOT('cld', 'Causal loop view of SOC triage (systems thinking)', 5.4),
  P('**Unintended consequences designed against:** deskilling of junior analysts (the console explains its reasoning and cites evidence), attackers adapting through prompt injection, and dependence on one model vendor (models are swappable and open-weight).'),
  H2('3.3 STOIC prompt engineering'),
  P('STOIC structures a prompt as **S**ituation, **T**ask, **O**bjective, **I**nputs and **C**onstraints. We refined our brainstorming prompt in three versions.'),
  T([1300, 4900, 2826], [
    ['Version', 'Prompt', 'Result'],
    ['v1', '"Give me ideas to use AI in a SOC."', 'Generic ideas ("AI chatbot", "automate everything"); no context, constraints or way to compare'],
    ['v2 (STOIC)', '**Situation:** MBA team designing a PoC for a mid-size SOC (~1,000 alerts a day, small L1/L2 team, Wazuh-type SIEM). **Task:** 8 distinct agentic-AI ideas for triage and investigation. **Objective:** most reduce time-to-triage without removing human control. **Inputs:** current process, our tools (n8n, Streamlit, open-weight LLMs), the brief\'s HITL and guardrail requirements. **Constraints:** buildable within the hackathon; synthetic data; no autonomous containment; open-weight models; one line per idea with its main risk.', 'Eight concrete ideas with risks; the AI picked ideas 1 and 2 combined, but the ideas were hard to compare (Figures ' + F('4a') + '–' + F('4b') + ')'],
    ['v3 (refined STOIC)', 'v2 plus: score each idea for Usefulness, Realism and Feasibility (1–5) with a one-line justification; name the human decision point; flag assumption-dependent scores; return a table and a five-line recommendation.', 'A comparable screening table and a five-line recommendation (Figures ' + F('5a') + '–' + F('5c') + '; section 3.4)'],
  ]),
  SP(),
  P('Each agent prompt in the prototype follows the same structure: a shared situation block (one component of a governed SOC, synthetic data, advises but never acts), the task, a stated objective, the case evidence delimited as data, and constraints (cite only supplied evidence IDs, return one JSON object matching the schema). Prompts are versioned (v1.0) and the version is recorded on every case.'),
  P('The v2 and v3 prompts were run in a general-purpose AI assistant. Figures ' + F('4a') + ' and ' + F('4b') + ' show the v2 prompt and its answer; Figures ' + F('5a') + ' to ' + F('5c') + ' show the refined v3 prompt, the scored screening table and the recommendation that section 3.4 summarises.'),
  ...SHOT('4a', 'STOIC v2 prompt and the first five of the eight ideas, each with its main risk', 6.0),
  ...SHOT('4b', 'v2 result: the AI picks ideas 1 and 2 combined (enrich the alert, then propose a verdict with cited evidence)', 5.4),
  ...SHOT('5a', 'Refined STOIC v3 prompt and the scored screening table (ideas 1 to 4)', 5.6),
  ...SHOT('5b', 'v3 screening table (ideas 5 to 8) and the assumptions behind the flagged scores', 5.6),
  ...SHOT('5c', 'v3 five-line recommendation: build the auto-enrichment evidence card, with a suggested-verdict slice of idea 2', 6.0),
  H2('3.4 Idea generation and screening'),
  P('The table reproduces the AI\'s v3 scores (1–5; † marks a score the AI flagged as depending on an assumption) and the main risk it gave for each idea in v2. The weighted column is the team\'s own: usefulness 40%, realism 30%, feasibility 30%.'),
  T([3300, 900, 900, 900, 1000, 2026], [
    ['Idea (AI output)', 'Useful', 'Realistic', 'Feasible', 'Weighted', 'Main risk'],
    ['**1. Auto-enrichment evidence card**', '5†', '5', '4†', '**4.7**', 'Stale or wrong enrichment data is over-trusted'],
    ['6. Escalation packager', '3', '5', '5', '4.2', 'Saves time only after the decision'],
    ['**2. Verdict and priority scorer**', '4†', '3', '4', '3.7', 'Overconfident wrong verdicts bury real incidents'],
    ['3. Alert clustering', '4†', '4', '3†', '3.7', 'Over-merging hides a distinct attack'],
    ['7. Runbook-grounded recommender', '3', '4', '3†', '3.3', 'Retrieves the wrong playbook or drifts beyond it'],
    ['8. Feedback and drift monitor', '2', '4', '4', '3.2', 'Sparse feedback; no direct time saving'],
    ['4. Investigation planner (read-only queries)', '4', '3', '2†', '3.1', 'Hallucinated queries give a false narrative'],
    ['5. ATT&CK mapper and chain stitcher', '2', '3', '3', '2.6', 'Wrong mappings give false confidence'],
  ]),
  SP(),
  P('**What the AI recommended.** In v2 it picked ideas 1 and 2 combined; in v3 it recommended idea 1 as the core, with a suggested verdict (a thin slice of idea 2) as a low-risk stretch, because enrichment is the biggest per-alert time sink and the analyst still makes every decision. It also proposed the guardrails to show: a confidence floor that forces human review, mandatory evidence citations, an allow-list and a full audit log.'),
  P('**What the team chose.** We built ideas 1 and 2 together as **multi-agent triage with human approval**: three context agents (correlation, threat intel, asset context) produce the evidence card (idea 1), and the risk agent proposes a verdict and action citing evidence IDs (idea 2). Code, not the model, scores and routes the alert, and every outcome needs a human decision. We kept light versions of idea 6 (the summary agent writes the escalation note for the Lead) and idea 8 (the audit log records each AI recommendation against the analyst\'s final decision), and adopted all four suggested guardrails. Clustering (idea 3) and the investigation planner (idea 4) are left as future work. We overrode the AI on one point: it suggested letting the model choose which enrichment tools to call; we fixed the lookups in code so that every case gathers the same evidence and can be audited.'),
  H2('3.5 Market research'),
  T([2300, 2700, 2300, 1726], [
    ['Category', 'Examples', 'What they do', 'Gap for our context'],
    ['AI assistants in security platforms', 'Microsoft Security Copilot, Google Security Operations, CrowdStrike Charlotte AI', 'Natural-language investigation and summaries inside one vendor\'s stack', 'Vendor lock-in; proprietary models'],
    ['SOAR and security automation', 'Splunk SOAR, Cortex XSOAR, Tines, Torq', 'Playbook-driven enrichment and response', 'Brittle on ambiguity; costly'],
    ['Autonomous AI SOC analysts', 'Dropzone AI, Prophet Security and others', 'Agentic investigation of every alert', 'Closed logic; cloud processing; cost'],
    ['Open-source building blocks', 'Wazuh, TheHive, Shuffle, MISP', 'Detection, cases, automation, TI', 'No agentic reasoning layer'],
  ]),
  SP(),
  P('**Positioning:** an open, governed reference design for mid-size SOCs and MSSPs, built from open-weight models and an open orchestration tool, with explicit HITL tiers and a guardrail-to-framework mapping that a GRC team can audit. This matters in India\'s regulated sectors (DPDP Act, CERT-In directions, RBI expectations). Product names are category examples from public vendor material.'),
  H2('3.6 Technology choices'),
  T([1900, 2700, 2000, 2426], [
    ['Decision', 'Options considered', 'Choice', 'Reason'],
    ['Orchestration', 'n8n, LangGraph, CrewAI, Power Automate', '**n8n**', 'Named in the brief; visual workflows; webhooks, branching, error workflows, data tables'],
    ['UI', 'Streamlit, Lovable, Replit, Flask', '**Streamlit**', 'Python; quick case views and charts'],
    ['Models', 'Qwen, Llama, Mistral, Gemma, GPT-OSS', '**GPT-OSS 20B and 120B, Qwen 3.8 27B** (Groq); Ollama for local', 'Open-weight; free hosted tier, no GPU; same prompts for all'],
    ['Structured output', 'Free text, JSON mode, function calling', '**JSON mode + strict schema in code + one retry**', 'Measurable compliance; failures go to humans'],
    ['Guardrails', 'LLM-as-judge, libraries, code', '**Deterministic code outside the model**', 'Predictable, testable, auditable'],
    ['Data store', 'Google Sheets, Postgres, n8n Data Tables', '**n8n Data Tables**', 'No credentials; built into n8n'],
  ]),
  H2('3.7 Opportunity justification'),
  P('First-line triage is **high-volume and repetitive** (strong automation value), **evidence-heavy** (well suited to AI summarising structured context), has **clear decision points** (natural places for HITL) and is **measurable** (triage time, coverage, false positives, escalation accuracy). Specialised agents with narrow inputs and outputs are easier to validate, swap and audit than one large prompt, and keeping the orchestrator deterministic means routing never depends on one model\'s free-form judgement.'),
  P('Value hypotheses for expert validation (not measured claims): a structured first assessment for close to 100% of alerts; evidence-gathering time cut from tens of minutes to minutes; consistent dispositions across analysts and shifts; a complete decision trail for auditors.'),
];

const s4 = [
  H1('4. Stage 2: Process mapping and architecture'),
  H2('4.1 AS-IS process'),
  P('The current process (Figure ' + F('asis') + ') is manual and spread across several consoles and teams.'),
  T([1700, 2900, 2100, 2326], [
    ['Actor', 'Activity', 'System', 'Bottleneck or exception'],
    ['Security tools', 'Raise alert into a shared queue', 'SIEM / EDR / IDS', 'Volume above capacity'],
    ['Tier 1 analyst', 'Pivot across consoles; look up IPs, hashes, domains; find asset owner and criticality', 'SIEM, EDR, firewall, TI portals, CMDB sheet, email', 'Swivel-chair work; CMDB gaps; conflicting TI'],
    ['Tier 1 analyst', '**Decide:** benign → close with free-text note; unsure → escalation ticket', 'Ticketing', 'Inconsistent between shifts; context lost in **handoff**'],
    ['Tier 2 / IR', 'Re-investigate; **decide** whether it is an incident', 'Same consoles again', 'Duplicate work'],
    ['SOC Lead / IR Manager', '**Approve** containment', 'Phone or chat', 'Weak audit trail; out-of-hours delay'],
    ['IT / asset owner', 'Isolate host, reset credentials, block IOC; update ticket', 'EDR, AD, firewall', 'Manual; errors cause outages'],
  ]),
  ...SHOT('asis', 'AS-IS manual and semi-manual SOC alert triage, with pain points', 4.4),
];
const s4b = [
  H2('4.2 TO-BE governed agentic process'),
  P('In the redesigned process (Figure ' + F('tobe') + ') every activity is marked **A** (autonomous), **H** (human-in-the-loop) or **E** (escalation). Routing is decided by code, never by a model.'),
  ...SHOT('tobe', 'TO-BE governed agentic triage process (A autonomous, H human-in-the-loop, E escalation)', 5.2),
  T([600, 3500, 2200, 800, 1926], [
    ['#', 'Activity', 'Performed by', 'Mode', 'Output'],
    ['1', 'Validate input schema and sizes', 'Orchestrator (code)', 'A', 'Pass, or exception'],
    ['2', 'Screen for prompt injection; mask PII', 'Orchestrator (code)', 'A', 'Clean, masked alert'],
    ['3', 'Look up threat intel, CMDB, identity, alert history', 'Orchestrator tools', 'A', 'Evidence set with IDs'],
    ['4', 'Correlate related events', 'Alert Correlation Agent', 'A', 'Related evidence, pattern, gaps'],
    ['5', 'Interpret threat intelligence', 'Threat Intelligence Agent', 'A', 'Verdict, conflicts, confidence'],
    ['6', 'Assess asset and identity impact', 'Asset Context Agent', 'A', 'Criticality, privilege'],
    ['7', 'Propose classification and one action', 'Risk/Prioritisation Agent', 'A', 'Proposal with evidence IDs'],
    ['8', 'Score F1–F7, apply overrides R1–R4', 'Scoring engine (code)', 'A', 'Score 0–100, class, flags'],
    ['9', 'Check evidence IDs and policy rules', 'Guardrails (code)', 'A → E', 'Pass, or exception'],
    ['10', 'Write and check incident summary', 'Incident Summary Agent + guardrail', 'A → E', 'Analyst-ready summary'],
    ['11', 'Confirm closure (LOW, benign, confidence ≥ 0.80)', 'Tier 1 analyst', 'H', 'Closed (simulated)'],
    ['12', 'Review (MEDIUM or low confidence)', 'Tier 1 / Tier 2', 'H', 'Approve, modify or reject'],
    ['13', 'Notify IR lead; approve containment (HIGH / CRITICAL)', 'Tier 2 or SOC Lead', 'E → H', 'Decision; IR notified'],
    ['14', 'Re-escalate if unapproved after 10 minutes', 'SLA watcher', 'E', 'SOC Lead alerted'],
    ['15', 'Handle exceptions manually', 'Analyst', 'E → H', 'Manual triage'],
    ['16', 'Execute approved allow-listed action; record everything', 'Orchestrator', 'A', 'Simulated action; audit row'],
  ]),
  H2('4.3 Agent roles'),
  T([1900, 2700, 2600, 1826], [
    ['Agent', 'Receives', 'Returns (strict JSON)', 'Cannot'],
    ['Alert Correlation', 'Alert, related history, entities', 'Related evidence IDs, pattern, missing evidence', 'Invent events'],
    ['Threat Intelligence', 'TI records for this case\'s indicators', 'Verdict per indicator, conflicts', 'Use outside knowledge'],
    ['Asset Context', 'CMDB and identity records', 'Criticality, sensitivity, privilege, gaps', 'Treat a missing asset as low risk'],
    ['Risk/Prioritisation', 'The three outputs above + alert', 'Behaviour level, proposed class, one action, reasons', 'Set the final score or route'],
    ['Incident Summary', 'Everything above + code\'s score and route', 'Summary; options marked requires_approval', 'Claim an action was taken'],
  ]),
  H2('4.4 Human-in-the-loop and escalation matrix'),
  T([2700, 1700, 1800, 2826], [
    ['Case class (from code)', 'Route', 'Who decides', 'Escalation'],
    ['LOW, benign, confidence ≥ 0.80', 'Closure confirmation', 'Tier 1', 'None'],
    ['LOW without benign explanation; any MEDIUM', 'Analyst review', 'Tier 1 (up to MEDIUM)', 'Tier 2 if unsure'],
    ['HIGH', 'Escalate', 'Tier 2 or Lead', 'IR Lead notified; SLA re-escalation'],
    ['CRITICAL', 'Escalate', 'SOC Lead only', 'IR Lead notified; SLA re-escalation'],
    ['Any exception', 'Exception queue', 'Tier 2 (HIGH/CRITICAL/unknown), else Tier 1', 'Manual; no AI recommendation to approve'],
  ]),
  SP(),
  P('The governing class for a decision is the higher of the original and any modified class, so an analyst cannot lower a case\'s class in order to approve it themselves.'),
  H2('4.5 Technical and governance architecture'),
  ...SHOT('arch', 'Technical and governance architecture in six layers', 6.2),
  T([2300, 3700, 3026], [
    ['Layer', 'Components', 'Governance role'],
    ['1 UI / UX', 'Streamlit console: queue, submit, case detail, decision form, audit log, model comparison', 'Evidence before recommendation; reason codes; password gate'],
    ['2 Orchestration', 'n8n WF-1 triage, SW-1 agent caller, WF-2 decision, WF-3 read API, WF-4 errors, WF-5 SLA', 'Deterministic routing; every path ends in a human queue'],
    ['3 Guardrails', 'Validation, injection screen, masking, schema + retry, evidence-ID check, policy rules, allow-list, role matrix', 'Controls the model cannot override'],
    ['4 Agents / LLMs', 'Five agents, versioned prompts; GPT-OSS 20B/120B, Qwen 3.8 27B (Groq) or local Ollama models', 'Swappable models; same prompts for fair comparison'],
    ['5 Tools / data', 'TI, CMDB, identity, alert history; soc_cases and soc_audit tables', 'Evidence IDs make claims traceable; synthetic data'],
    ['6 Oversight', 'Approval tiers, SLA, audit trail, framework mapping, benchmark, expert validation', 'Accountability and continuous assurance'],
  ]),
];

const s5 = [
  H1('5. Stage 3: Integrated working solution'),
  H2('5.1 What was built'),
  T([2600, 6426], [
    ['Component', 'What it does'],
    ['WF-1 Triage (n8n)', 'Webhook from the console; validation, injection and PII screen, tool lookups, five agent calls, scoring, guardrails, routing, case and audit records'],
    ['SW-1 LLM Agent (n8n)', 'Calls the model, validates the JSON against a strict schema, one corrective retry; waits and resends on provider rate limits; never throws'],
    ['WF-2 Analyst Decision (n8n)', 'Approve, modify or reject; server-side role matrix; allow-list re-check; simulated action with ticket number'],
    ['WF-3 Read API / WF-4 Error Handler / WF-5 SLA Watcher', 'Read cases and audit for the console; catch crashes into the manual queue; re-escalate stale escalations'],
    ['Streamlit console', 'Alert queue, submit alert, case detail (score breakdown, flags, agent findings), decision form, audit log with CSV export, model comparison, guardrails page'],
    ['Data', 'Five synthetic alerts; threat intel, CMDB, identity and alert-history lookups; expected outcomes in ground_truth.json'],
    ['Benchmark', 'Runs every model on the same alerts and repeats, scores against ground truth, refuses mock models, writes a blind rubric sheet for explanation quality'],
  ]),
  SP(),
  P('**Model routing is never left to the model.** A weaker model shows up as lower confidence, more reviews or more exceptions, not as a wrong closure.'),
  ...SHOT(6, 'n8n workflow list: the seven SOC workflows (WF-5 SLA Watcher is optional and left unpublished)', 5.6),
  H2('5.2 Brief requirements and where they are shown'),
  T([3000, 3800, 2226], [
    ['Requirement', 'Implementation', 'Evidence'],
    ['UI triggers the workflow', 'Console → POST /soc-triage', figs(7, 10)],
    ['Multiple agents and orchestration', 'Five agents through SW-1, orchestrated by WF-1', figs(6, 8, 9)],
    ['Conditional routing', 'Code-computed route: closure, review, escalate, exception', figs(8, 17)],
    ['Tool / data integration', 'TI, CMDB, identity, history lookups; data tables', figs(11)],
    ['Human-in-the-loop', 'WF-2 approve / modify / reject with role matrix', figs(12, 16, 17, 18, 19)],
    ['Escalation', 'IR notification; SLA re-escalation', figs(13, 17)],
    ['Exception handling', 'Validation, injection, schema retry, guardrails, outage, error workflow', figs(14, 20, 21, 22, 23, 24)],
    ['≥ 3 test cases', 'Normal, ambiguous, high-risk + injection + malformed', 'Section 5.3'],
    ['≥ 3 open-weight models compared', 'Benchmark on identical cases', 'Section 6'],
  ]),
  SP(),
  H3('Console triggers the workflow'),
  P('The analyst submits an alert in the Streamlit console, which posts it to the n8n triage webhook and shows the case status as soon as the workflow returns (Figure ' + F(7) + ').'),
  ...SHOT(7, 'Console: the analyst submits ALRT-1001, which starts the n8n triage workflow; the case returns as awaiting closure confirmation', 5.8),
  H3('Orchestration and agents'),
  P('WF-1 orchestrates the whole case: validation and screening, tool lookups, the five agent calls, scoring, guardrails and routing (Figure ' + F(8) + '). Each agent call goes through the reusable SW-1 workflow, which validates the model\'s JSON and retries once (Figure ' + F(9) + '). Every run is recorded in n8n\'s execution history (Figure ' + F(10) + ').'),
  ...SHOT(8, 'n8n WF-1 Triage: validation, screening, tool lookups, five agent calls, scoring, guardrails and code-based routing', 6.3),
  ...SHOT(9, 'n8n SW-1 LLM Agent: model call, strict JSON validation and one corrective retry; every call is audited', 6.3),
  ...SHOT(10, 'n8n WF-1 Executions view: every triage run is recorded with its status and duration', 6.0),
  H3('Tool and data integration'),
  P('Lookups against threat intelligence, the asset register, the identity directory and alert history feed the agents; every case and audit event is stored in n8n Data Tables (Figure ' + F(11) + ').'),
  ...SHOT(11, 'n8n Data Table soc_cases: one record per case with status, route and priority', 6.0),
  H3('Human-in-the-loop'),
  P('WF-2 receives every approve, modify or reject decision, derives the analyst\'s role on the server, applies the role matrix and the action allow-list, and audits denials (Figure ' + F(12) + ').'),
  ...SHOT(12, 'n8n WF-2 Analyst Decision: server-side role matrix and allow-list check', 6.3),
  H3('Escalation'),
  P('CRITICAL and HIGH cases are escalated with a simulated notification to the incident response lead. WF-5 checks every two minutes and re-escalates any escalation left unapproved for more than 10 minutes (Figure ' + F(13) + ').'),
  ...SHOT(13, 'n8n WF-5 SLA Watcher: re-escalates unapproved escalations after 10 minutes', 6.0),
  H3('Exception handling'),
  P('Invalid input and prompt-injection attempts never reach a model: Validate Input and the Injection & PII Screen route them straight to the Exception Handler and the manual queue (Figure ' + F(14) + ').'),
  ...SHOT(14, 'n8n WF-1 close-up: Validate Input and the Injection & PII Screen route bad input to the Exception Handler', 6.3),
  H2('5.3 Test cases and results'),
  P('The pipeline was verified end to end through n8n using a test model that returns known answers. This proves the workflows, scoring, guardrails and approvals; model quality is measured separately in section 6.'),
  T([1300, 1600, 2700, 3426], [
    ['Alert', 'Type', 'Expected', 'Result'],
    ['ALRT-1001', 'Normal', 'LOW → closure confirmation', 'Score 12, LOW, awaiting closure confirmation ✔'],
    ['ALRT-2001', 'Ambiguous', 'MEDIUM → analyst review', 'Score 49, MEDIUM, flags UNKNOWN_ASSET and TI_CONFLICT ✔'],
    ['ALRT-3001', 'High risk', 'CRITICAL → escalation', 'Score 97, CRITICAL, escalated pending approval; IR notified (simulated) ✔'],
    ['ALRT-4001', 'Prompt injection', 'Stopped before any model call', 'Exception queue at injection screen ✔'],
    ['ALRT-5001', 'Malformed', 'Stopped at validation', 'Exception queue at input validation ✔'],
  ]),
  H3('Normal case: ALRT-1001'),
  P('An authorised weekly vulnerability scan triggers a port-scan alert. Every lookup is clean, the score is 12 (LOW) and the case waits for a Tier 1 analyst to confirm closure (Figure ' + F(15) + ').'),
  ...SHOT(15, 'Case detail, ALRT-1001 (normal): LOW, score 12, no flags; closed after Tier 1 confirmation (simulated)', 6.0),
  H3('Ambiguous case: ALRT-2001'),
  P('An out-of-baseline VPN login is followed by a large HR file-share read. The asset is missing from the register and threat intel conflicts, so the case is flagged, scored 49 (MEDIUM) and sent to analyst review instead of being closed (Figure ' + F(16) + ').'),
  ...SHOT(16, 'Case detail, ALRT-2001 (ambiguous): MEDIUM, score 49, flags UNKNOWN_ASSET and TI_CONFLICT; the Tier 1 analyst modified the action to MONITOR (simulated)', 6.0),
  H3('High-risk case: ALRT-3001'),
  P('A multi-stage chain on a critical finance database server ends in command-and-control traffic. Overrides R1 and R2 lift the case to CRITICAL (97); a Tier 1 approval is refused by the server and the SOC Lead approves a simulated isolation (Figure ' + F(17) + ').'),
  ...SHOT(17, 'Case detail, ALRT-3001 (high risk): CRITICAL, score 97, overrides R1 and R2; approved by the SOC Lead as a simulated action', 6.0),
  P('The two attack cases, ALRT-4001 (prompt injection) and ALRT-5001 (malformed input), are stopped before any model call; their results appear in the self-test output (Figure ' + F(24) + ').'),
  SP(),
  SP(),
  T([4600, 4426], [
    ['Control test', 'Result'],
    ['Invented evidence ID (LOG-9999-99)', 'Caught by output guardrail; case sent to exception queue ✔'],
    ['Broken JSON from every agent on first try', 'All five recovered on the one retry ✔'],
    ['Provider rate limit (HTTP 429) on every first call', 'Waited and resent; recovered in 94 s ✔'],
    ['Model unreachable / API key missing', 'Case sent safely to the manual queue with the cause recorded ✔'],
    ['Decision rules (12 cases)', '12/12 passed: Tier 1 refused on CRITICAL; downgrade-to-approve refused; off-list action refused; second decision refused; Lead approval simulated and labelled ✔'],
    ['Unit tests of the core logic', '46/46 passed ✔'],
  ]),
  SP(),
  H3('Guardrail and control evidence'),
  P('When a test model cites evidence that does not exist, the output guardrail rejects the recommendation and hands the case to a human (Figure ' + F(22) + '). The audit log records every agent call, guardrail exception and human decision (Figure ' + F(23) + '), and the self-test confirms all five alerts and all twelve decision rules (Figure ' + F(24) + ').'),
  ...SHOT(22, 'ALRT-3001 run with a test model that invents evidence: the output guardrail catches LOG-9999-99, so no AI recommendation is trusted and a human must handle the case', 6.0),
  ...SHOT(23, 'Audit log: every agent call, guardrail exception and human decision with actor, model, prompt version and latency (note the invented ID LOG-9999-99 caught and a denied second decision); CSV export', 6.3),
  ...SHOT(24, 'SELF_TEST.bat: all five test alerts (including the injection and malformed cases), 12/12 human-decision rules (including Tier 1 refused on CRITICAL) and the invented-evidence guardrail demo passed', 6.3),
  H2('5.4 Deployment'),
  P('The system runs in three ways, all free:'),
  T([2300, 3800, 2926], [
    ['Mode', 'How', 'Use'],
    ['Laptop', 'Node.js 24 and Python; 1_SETUP.bat once, 2_START.bat each session; models via Groq\'s free API', 'Development, rehearsal, benchmark'],
    ['Shared from laptop', '3_SHARE_ONLINE.bat opens a password-protected public link through a Cloudflare tunnel; n8n stays private', 'Live evaluation session'],
    ['Cloud', 'n8n Cloud hosts the workflows (Groq key held in an encrypted n8n credential); Streamlit Community Cloud hosts the console behind a password', 'Always-on demo for the evaluation window'],
  ]),
  SP(),
  P('The project website is published on GitHub Pages. The n8n Cloud deployment uses a 14-day trial with 1,000 executions and a 180-second limit per execution; the cloud build shortens the pacing between agent calls so each case finishes in well under a minute, and the benchmark is run on the laptop to preserve the quota.'),
  ...SHOT(25, 'n8n Cloud: the workflows imported and published for the always-on demo', 5.6),
];

const s6 = [
  H1('6. Stage 3: Model comparison'),
  P('**Method.** Each model received the same alerts, prompts, evidence and output schema at temperature 0 with a fixed seed, through the same n8n workflow, and every run was scored automatically against the expected outcomes in `ground_truth.json`. The alerts that never reach a model (injection and malformed input) count only towards routing and escalation. Provider throttling is recorded separately and does not count against a model.'),
  P('**Sample.** The benchmark completed 6 scored runs for each GPT-OSS model and 20 for Qwen 3.8 27B. The sample is small and unequal, so the results below are indicative rather than statistically conclusive.'),
  T([3600, 1808, 1808, 1810], [
    ['Metric', 'GPT-OSS 20B', 'GPT-OSS 120B', 'Qwen 3.8 27B'],
    ['Runs scored', '6', '6', '20'],
    ['Final risk class: exact / within one level', '**100% / 100%**', '**100% / 100%**', '50% / 50%'],
    ['Model\'s own proposed class correct', '**75%**', '50%', '50%'],
    ['Incident category correct', '**75%**', '25%', '50%'],
    ['Recommended action correct', '**75%**', '50%', '50%'],
    ['Route matches the expected route', '67%', '**83%**', '60%'],
    ['Escalation decision correct', '**83%**', '**83%**', 'See Annex E'],
  ]),
  SP(),
  P('Values are taken from `benchmark/results/summary_by_model.csv` (Figure ' + F(26) + '). The same file records, per model, false positives and negatives, invented evidence IDs caught by the guardrail, JSON validity on the first attempt, retries, rate-limited calls, latency and task completion; raw per-run data is in `runs.csv` (Annex E). Explanation quality has not yet been scored; the blind rubric sheet is generated for that step.'),
  ...SHOT(26, 'Console Model comparison page: results read only from benchmark/results'),
  H2('6.1 Analysis'),
  P('**Accuracy.** Both GPT-OSS models ended with the correct final risk class in every scored run; Qwen 3.8 27B did so in half. Qwen\'s within-one rate equals its exact rate, so its misses are not near-misses on the risk scale. They are more likely runs that ended without a usable model result; the cause of each one is recorded in `runs.csv`.'),
  P('**What the models proposed themselves.** The final class is computed in code from the evidence, so the model\'s own suggestion is scored separately. GPT-OSS 20B agreed with the expected answer most often (75% for class, category and action), ahead of GPT-OSS 120B (50%, 25% and 50%) and Qwen (50%). Where a model under- or over-called a case, the code scoring and override rules still produced the correct final class. This is the property the design relies on: the model advises, code decides.'),
  P('**Routing, escalation and safety.** GPT-OSS 120B matched the expected route most often (83%), against 67% for GPT-OSS 20B and 60% for Qwen. Both GPT-OSS models made the correct escalate-or-not decision in 5 of 6 runs. Every route, including closure, ends with a human decision, so a routing miss costs analyst time rather than causing an unreviewed closure or action.'),
  P('**Structured output, guardrails and speed.** JSON validity, retries and invented-ID catches for each model are recorded in the benchmark files (Annex E); the guardrails sent any invalid or unsupported output to the exception queue, never to an analyst as a trusted recommendation. Latency in this run is dominated by the deliberate pause between agent calls that keeps the workflow inside the provider\'s free tier, so it does not separate the models.'),
  P('**Recommendation.** We recommend **GPT-OSS 120B** as the primary model for a pilot: correct final class in every run and the best routing. **GPT-OSS 20B** is the fallback for high alert volumes, with the same final accuracy and the best agreement in its own proposals. Qwen 3.8 27B is not recommended until its incomplete runs are explained. Before a pilot, the benchmark should be repeated with an equal number of runs per model (5 alerts × 3 repeats) and blind explanation scoring.'),
];

const RISKS = [
  ['AIR-01 Personal data exposed to the model', 'Mask emails and phone numbers before any model call', 'NIST MEASURE 2.10; ISO 42001 A.7; ISO 27001 A.8.11, A.5.34; DPDP Act', 'Injection & PII screen', 'Partial'],
  ['AIR-02 Confidential data leaves the organisation', 'Synthetic data only; explicit endpoints; key never in files', 'NIST GOVERN 6.1; ISO 42001 A.10; ISO 27001 A.5.19–5.23, A.8.12', 'Model config; start script; n8n credential', 'Implemented (hosted mode a stated deviation)'],
  ['AIR-03 Invented evidence', 'Every cited ID must exist in the case evidence', 'NIST MEASURE 2.5, 2.9; ISO 42001 A.6.2.4', 'Output and summary guardrails', 'Implemented'],
  ['AIR-04 Overstated conclusions', 'Containment needs malicious TI or multi-stage; closure needs benign explanation; no claimed actions', 'NIST MEASURE 2.6, MANAGE 1.3; ISO 42001 A.9.4', 'Policy and summary guardrails', 'Implemented'],
  ['AIR-05 Prompt injection', 'Screen before any model call; text treated as data; code-driven routing', 'NIST MEASURE 2.7; ISO 27001 A.8.28, A.8.26', 'Injection screen; shared prompt', 'Implemented (known patterns)'],
  ['AIR-06 Malformed or oversized input', 'Required fields, formats, size limits', 'NIST MEASURE 2.5; ISO 42001 A.7.4', 'Validate Input', 'Implemented'],
  ['AIR-07 Malformed model output', 'Strict schema; one corrective retry; human queue on failure', 'NIST MANAGE 2.4; ISO 42001 A.6.2.6', 'SW-1 parse and validate', 'Implemented'],
  ['AIR-08 Unsafe autonomous action', 'No action without approval; allow-list; simulated', 'NIST GOVERN 3.2, MAP 3.5; ISO 42001 A.9.2; ISO 27001 A.5.26', 'WF-2 decision workflow', 'Implemented'],
  ['AIR-09 Unauthorised approval', 'Role from server; class caps authority; no downgrade to self-approve', 'NIST GOVERN 2.1; ISO 42001 A.3.2; ISO 27001 A.5.15, A.8.2, A.5.3', 'Role-matrix check', 'Partial (no SSO in PoC)'],
  ['AIR-10 Miscalibrated thresholds', 'Versioned thresholds; hard overrides; low confidence forces review', 'NIST MAP 1.1, MEASURE 1.1; ISO 42001 A.5.2', 'Scoring engine', 'Partial (expert sign-off)'],
  ['AIR-11 Automation bias', 'Evidence before recommendation; reason code on reject', 'NIST MAP 3.5, MEASURE 2.8; ISO 42001 A.8.2', 'Case detail page', 'Implemented'],
  ['AIR-12 Unfair prioritisation', 'Unknown asset or user scored above medium, not as safe', 'NIST MEASURE 2.11; ISO 42001 A.5.2', 'Asset lookup ordering; flags', 'Partial'],
  ['AIR-13 Missing audit trail', 'Every step logged with actor, model, prompt version, latency', 'NIST MANAGE 4.1; ISO 42001 A.6.2.8; ISO 27001 A.8.15', 'soc_audit table; CSV export', 'Partial (not tamper-proof)'],
  ['AIR-14 Pipeline failure or outage', 'Every failure ends in the manual queue; timeouts; error workflow', 'NIST MANAGE 2.4; ISO 27001 A.5.24', 'Exception handler; WF-4', 'Implemented'],
  ['AIR-15 Escalation not picked up', 'Re-escalate after 10 minutes', 'NIST MANAGE 4.3; ISO 27001 A.5.25', 'WF-5 SLA watcher', 'Implemented, untested'],
  ['AIR-16 Uncontrolled model or prompt change', 'Approved model list; versions on every case; re-benchmark', 'NIST GOVERN 1.4; ISO 27001 A.8.9, A.8.32', 'Validate Input; benchmark', 'Implemented (sign-off planned)'],
];
const s7 = [
  H1('7. Stage 4: Guardrails and governance'),
  P('Each risk is traced as **Risk → Guardrail → Governance requirement → Implementation in the PoC**, against **NIST AI RMF 1.0** (primary), **ISO/IEC 42001:2023** Annex A and **ISO/IEC 27001:2022** Annex A, with **ISO/IEC 27701** and India\'s **DPDP Act 2023** for personal data. The full register with detailed status notes is in Annex D.'),
  ...SHOT(27, 'Console Guardrails and scope page', 5.6),
  H2('7.1 Design principles'),
  N('**The model advises; code decides.** Thresholds, overrides, routing and permitted actions live in code and a versioned config file.', 'pr'),
  N('**Nothing reaches a model unscreened.** Validation and an injection and PII screen run first.', 'pr'),
  N('**Every claim must point to evidence it was given.** Every cited ID is checked.', 'pr'),
  N('**Fail closed.** Every failure sends the case to a human queue with the reason recorded.', 'pr'),
  N('**Humans own consequential decisions**, with authority matched to risk.', 'pr'),
  N('**Every step is logged.**', 'pr'),
  H2('7.2 Risk register'),
  T([2050, 2150, 2050, 1500, 1276], [['Risk', 'Guardrail', 'Governance requirement', 'Where enforced', 'Status'], ...RISKS], { size: 16 }),
  SP(),
  NOTE('Clause numbers are given so the mapping can be audited. Each reference was cross-checked against the subcategory and control titles in NIST AI 100-1 (AI RMF 1.0), ISO/IEC 42001:2023 Annex A and ISO/IEC 27001:2022 Annex A; where a risk touches several controls, the most specific one is listed first.'),
  H2('7.3 Thresholds and confidence rules'),
  T([1800, 1600, 5626], [
    ['Band', 'Score', 'Default route (before overrides)'],
    ['LOW', '0–29', 'Closure confirmation only if confidence ≥ 0.80 and behaviour is benign-explained; otherwise review'],
    ['MEDIUM', '30–59', 'Analyst review'],
    ['HIGH', '60–79', 'Escalate, pending approval'],
    ['CRITICAL', '80–100', 'Escalate, pending approval; IR Lead notified (simulated)'],
  ]),
  SP(),
  P('Overrides set a minimum class regardless of score: **R1** critical asset + malicious TI (confidence ≥ 0.70) → at least HIGH; **R2** privileged user with suspicious or multi-stage behaviour on a HIGH/CRITICAL asset → at least HIGH; **R3** asset missing from the CMDB → at least MEDIUM; **R4** unknown or conflicting TI on a High/Critical alert → at least MEDIUM. Confidence loses 0.10 per completeness flag; below 0.60 a case cannot be routed to closure. Where model and rules disagree, the rules win and the analyst sees a MODEL_RULE_DISAGREEMENT flag.'),
  P('These values are the team\'s design judgement and need expert sign-off and re-calibration on real alerts before live use. Proposed pilot targets: zero false negatives on CRITICAL, fewer than 10% of LOW cases escalated, and zero invented evidence IDs.'),
  H2('7.4 Governance operating model (proposed)'),
  T([4000, 2600, 2426], [
    ['Activity', 'Owner', 'Frequency'],
    ['Approve models, prompts and config for release', 'AI system owner + SOC Lead', 'Every change'],
    ['Re-run the benchmark before release', 'SOC engineering', 'Every change'],
    ['Review overrides, rejections and exceptions', 'SOC Lead', 'Weekly'],
    ['Calibrate thresholds and weights', 'SOC Lead + GRC', 'Quarterly or after a missed incident'],
    ['Audit-log integrity and access review', 'GRC / internal audit', 'Quarterly'],
    ['Update AI impact assessment', 'AI system owner + DPO', 'Annually or on scope change'],
  ]),
];

const s8 = [
  H1('8. Stage 5: Industry expert validation'),
  P('The expert profile was approved by faculty before the interview. The full video and transcript are in the shared folder (Annexes A and B).'),
  T([2600, 6426], [
    ['Item', 'Detail'],
    ['Expert', 'Jatin Modi'],
    ['Designation and organisation', 'Director, Commercial IQ'],
    ['Level', 'Director'],
    ['Faculty approval', '30 September 2026'],
    ['Interview date and mode', '9 October 2026; video call'],
  ]),
  SP(),
  P('The interview covered the five mandatory questions, then a walkthrough of the AS-IS and TO-BE processes, the architecture, the live prototype (normal, ambiguous and high-risk cases), the human approval controls and the guardrails, followed by validation questions on feasibility, risks, usefulness, thresholds and improvements.'),
  ...SHOT(28, 'Expert interview (video call, shown with consent)', 5.0),
  H2('8.1 Responses to the mandatory questions'),
  T([3600, 5426], [
    ['Question', 'Summary of the expert\'s answer'],
    ['Q1. What makes SOC alert triage and investigation complex? What evidence, tools, rules and judgement are involved?', '[[TO FILL]]'],
    ['Q2. Which people, systems, tools, data sources or teams does an analyst depend on, and where do these cause delays?', '[[TO FILL]]'],
    ['Q3. How important is speed, and which stages are most time-sensitive?', '[[TO FILL]]'],
    ['Q4. What types and volume of information are examined, and what goes wrong when it is incomplete or fragmented?', '[[TO FILL]]'],
    ['Q5. Which unusual, ambiguous or high-risk alerts are hard to handle with standard rules, and how are they escalated?', '[[TO FILL]]'],
  ]),
  H2('8.2 Validation findings and our response'),
  T([1900, 3900, 3226], [
    ['Theme', 'Expert view', 'Our response / change made'],
    ['Feasibility', '[[TO FILL]]', '[[TO FILL]]'],
    ['Risks', '[[TO FILL]]', '[[TO FILL]]'],
    ['Usefulness', '[[TO FILL]]', '[[TO FILL]]'],
    ['Approval tiers', '[[TO FILL]]', '[[TO FILL]]'],
    ['Thresholds and ground truth', '[[TO FILL]]', '[[TO FILL]]'],
    ['Improvements', '[[TO FILL]]', '[[TO FILL]]'],
  ]),
];

const s9 = [
  H1('9. Stage 6: Website, demonstration and deployment'),
  T([2600, 6426], [
    ['Deliverable', 'Where'],
    ['Project website', 'https://aviralbajpei.github.io/soc-triage-website/ (GitHub Pages; problem, process diagrams, architecture, results, guardrails, expert findings, video)'],
    ['Demonstration video (10–15 min, all members)', 'Shared Drive folder (link on the Project team and resources page)'],
    ['Live console', '[[Live console link: to be added]] (password shared with evaluators)'],
    ['Source code and documents', 'Project repository: workflows, console, prompts, data, benchmark, tests, Stage 1–6 documents'],
  ]),
  SP(),
  P('The video follows a twelve-segment script: problem and research, AS-IS and TO-BE processes, architecture, the three live cases (including a refused Tier 1 approval), the injection and invented-evidence guardrails, the model comparison, the guardrail register, the expert findings, and the limitations, with every member presenting.'),
  ...SHOT(29, 'Project website on GitHub Pages (aviralbajpei.github.io/soc-triage-website)'),
  ...SHOT(30, 'Live console on Streamlit Community Cloud'),
];

const s10 = [
  H1('10. Limitations and recommendations'),
  H2('10.1 Limitations'),
  B('**Small test set:** five synthetic alerts show the controls work, not accuracy at scale or statistical fairness.'),
  B('**Injection screening** is pattern-based and could be paraphrased around; routing in code limits the damage.'),
  B('**Usernames and hostnames** reach the model unmasked; only direct identifiers are masked.'),
  B('**Analyst identity** comes from a selector, not real sign-in.'),
  B('**The audit table** is append-only by design, not tamper-proof.'),
  B('**Thresholds** reflect the team\'s design judgement until expert sign-off.'),
  B('**Hosted inference** sends masked, synthetic data to a third party; in the cloud deployment the n8n webhook addresses are reachable on the internet (not published; synthetic data only).'),
  B('**All response actions are simulated.**'),
  H2('10.2 Recommendations'),
  N('Run a shadow-mode pilot on real alerts, with expert-set thresholds and local or contracted inference.', 'rec'),
  N('Integrate with a live SIEM (for example Wazuh) and single sign-on with MFA for analyst identity.', 'rec'),
  N('Forward audit events to the SIEM or write-once storage.', 'rec'),
  N('Add a classifier-based injection screen and pseudonymise usernames and hostnames.', 'rec'),
  N('Feed analyst decisions back into detection tuning, and re-run the benchmark before every model or prompt change.', 'rec'),
  H2('References'),
  P('[R1] The Hacker News, "The State of AI in the SOC 2025: Insights from Recent Study", September 2025 (survey of 282 security leaders).', { run: { size: 20 } }),
  P('[R2] SANS 2025 Detection and Response Survey, as reported by LMNTRIX, "Alert Fatigue and the SOC Noise Crisis", 2026.', { run: { size: 20 } }),
  P('NIST AI 100-1, Artificial Intelligence Risk Management Framework (AI RMF 1.0), 2023. ISO/IEC 42001:2023, Artificial intelligence management system. ISO/IEC 27001:2022, Information security management systems. ISO/IEC 27701, Privacy information management. Digital Personal Data Protection Act, 2023 (India).', { run: { size: 20 } }),
  H2('Annexes'),
  T([900, 8126], [
    ['Annex', 'Content'],
    ['A', 'Expert interview video (shared Drive folder)'],
    ['B', 'Full interview transcript'],
    ['C', 'Expert profile, evidence of designation and faculty approval'],
    ['D', 'Stage 1–4 working documents, including the full guardrail register (project docs folder)'],
    ['E', 'Benchmark raw data (benchmark/results/runs.csv)'],
    ['F', 'Deployment guides: laptop setup, sharing online, n8n Cloud and Streamlit hosting'],
  ]),
];

// ======================================================================= document
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
  new TextRun({ text: 'Governed Agentic AI for SOC Alert Triage  ·  ', font: FONT, size: 16, color: GREY }),
  new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY })] })] });
const portrait = { page: { margin: { top: 1440, bottom: 1300, left: 1440, right: 1440 } } };
const landscape = { page: { size: { orientation: PageOrientation.LANDSCAPE }, margin: { top: 1200, bottom: 1200, left: 1440, right: 1440 } } };
const numCfg = (ref) => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 360 } } } }] });

const doc = new Document({
  creator: 'Aviral Bajpei, Jatin Balecha, Prakhar Aryan, Garvit Saraswat, Arijit Paul',
  title: 'Governed Agentic AI for SOC Alert Triage and Incident Investigation',
  styles: {
    default: { document: { run: { font: FONT, size: 22, color: INK } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 34, bold: true, color: INK }, paragraph: { outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 26, bold: true, color: INK }, paragraph: { outlineLevel: 1 } },
    ],
  },
  numbering: { config: [
    { reference: 'bul', levels: [
      { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 280 } } } },
      { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 840, hanging: 280 } } } }] },
    numCfg('num'), numCfg('why'), numCfg('pr'), numCfg('rec'),
  ] },
  sections: [
    { properties: portrait, children: [...cover, ...details] },
    { properties: portrait, footers: { default: footer }, children: [...buildToc(), ...s1, ...s2, ...s3, ...s4, ...s4b, ...s5, ...s6, ...s7, ...s8, ...s9, ...s10] },
  ],
});
if (process.env.SOC_HEADINGS_OUT) fs.writeFileSync(process.env.SOC_HEADINGS_OUT, JSON.stringify(HEADINGS));
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(OUT, b); console.log('wrote', OUT, Math.round(b.length / 1024) + ' KB'); });
