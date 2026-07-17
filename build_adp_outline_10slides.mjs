import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath, pathToFileURL } from "node:url";

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const RUNTIME_ROOT = process.env.CODEX_RUNTIME_ROOT ?? path.join(
  os.homedir(),
  ".cache/codex-runtimes/codex-primary-runtime/dependencies",
);
const artifactToolUrl = pathToFileURL(path.join(
  RUNTIME_ROOT,
  "node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs",
)).href;
const { Presentation, PresentationFile } = await import(artifactToolUrl);
const WORK = path.join(ROOT, "work/presentations/adp_outline_10slides");
const TMP = path.join(WORK, "tmp");
const PREVIEW = path.join(TMP, "preview");
const LAYOUT = path.join(TMP, "layout");
const ASSET = path.join(TMP, "assets");
const QA = path.join(TMP, "qa");
const OUT = path.join(ROOT, "outputs");
const FINAL_PPTX = path.join(OUT, "Value_Iteration_ADP_导师汇报最终版.pptx");

const requireFromRuntime = createRequire(
  path.join(RUNTIME_ROOT, "node/node_modules/.pnpm/node_modules/package.json"),
);
const requireFromProject = createRequire(path.join(ROOT, "package.json"));
const { chromium } = requireFromRuntime("playwright-core");
const katex = requireFromProject("katex");
const katexCss = await fs.readFile(
  requireFromProject.resolve("katex/dist/katex.min.css"),
  "utf8",
);

const C = {
  bg: "#F5F8FC",
  ink: "#172033",
  muted: "#5D6D83",
  navy: "#103052",
  blue: "#2F80ED",
  cyan: "#22AFC3",
  green: "#159A6A",
  amber: "#B77900",
  softBlue: "#E9F3FF",
  softGreen: "#EAF9F2",
  softAmber: "#FFF7E6",
  white: "#FFFFFF",
  line: "#DCE5EF",
};

const formulas = {
  system: String.raw`x_{k+1}=F(x_k,u_k)`,
  cost: String.raw`J(x_0,u_0)=\sum_{k=0}^{\infty}U(x_k,u_k)`,
  policy: String.raw`v_i(x)=\underset{u}{\arg\min}\{U(x,u)+V_i(F(x,u))\}`,
  value: String.raw`V_{i+1}(x)=U(x,v_i(x))+V_i(F(x,v_i(x)))`,
  limit: String.raw`\lim_{i\to\infty}V_i(x)=J^*(x)`,
  conv: String.raw`|V_{i+1}(x)-V_i(x)|\leq\varepsilon`,
  adm: String.raw`V_{i+1}(x)-V_i(x)<U(x,v_i(x))`,
};

async function renderFormulaImages() {
  await fs.mkdir(ASSET, { recursive: true });
  const browser = await chromium.launch({
    headless: true,
    executablePath: "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
  });
  const page = await browser.newPage({
    viewport: { width: 1300, height: 260 },
    deviceScaleFactor: 3,
  });
  const result = {};
  for (const [key, source] of Object.entries(formulas)) {
    const html = katex.renderToString(source, {
      displayMode: true,
      throwOnError: false,
      output: "htmlAndMathml",
    });
    await page.setContent(`<!doctype html>
      <html>
        <head>
          <style>
            ${katexCss}
            .katex-display { margin: 0; }
            .katex { font-size: 1em; color: #172033; }
          </style>
        </head>
        <body style="margin:0;background:transparent;">
          <div id="eq" style="display:inline-block;padding:16px 24px;font-size:40px;line-height:1.2;">
            ${html}
          </div>
        </body>
      </html>`);
    const file = path.join(ASSET, `${key}.png`);
    await page.locator("#eq").screenshot({ path: file, omitBackground: true });
    result[key] = await fs.readFile(file);
  }
  await browser.close();
  return result;
}

const formulaImages = await renderFormulaImages();
const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });

function addText(slide, value, pos, style = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    position: pos,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = value;
  shape.text.style = {
    typeface: style.typeface ?? "Aptos",
    fontSize: style.fontSize ?? 24,
    color: style.color ?? C.ink,
    bold: style.bold ?? false,
    alignment: style.alignment ?? "left",
  };
  return shape;
}

function addBox(slide, pos, fill = C.white, line = C.line, radius = 10) {
  return slide.shapes.add({
    geometry: "roundRect",
    position: pos,
    fill,
    line: { style: "solid", fill: line, width: 1 },
    borderRadius: radius,
    shadow: "shadow-sm",
  });
}

function addTitle(slide, section, title, n) {
  addText(slide, "学术汇报", { left: 72, top: 42, width: 520, height: 24 }, {
    fontSize: 12,
    color: C.blue,
    bold: true,
  });
  addText(slide, title, { left: 72, top: 78, width: 1000, height: 56 }, {
    typeface: "Aptos Display",
    fontSize: 35,
    color: C.navy,
    bold: true,
  });
  slide.shapes.add({
    geometry: "line",
    position: { left: 72, top: 150, width: 1136, height: 0 },
    fill: "none",
    line: { style: "solid", fill: C.line, width: 1 },
  });
  addText(slide, "Wei, Liu & Lin, IEEE Transactions on Cybernetics, 2016", { left: 72, top: 672, width: 620, height: 20 }, {
    fontSize: 10,
    color: C.muted,
  });
  addText(slide, String(n).padStart(2, "0"), { left: 1150, top: 666, width: 58, height: 22 }, {
    fontSize: 12,
    color: C.navy,
    bold: true,
    alignment: "right",
  });
}

function addNote(slide, head, body, pos, color = C.blue, fill = C.white) {
  addBox(slide, pos, fill, C.line, 10);
  slide.shapes.add({
    geometry: "rect",
    position: { left: pos.left, top: pos.top, width: 8, height: pos.height },
    fill: color,
    line: { style: "solid", fill: "none", width: 0 },
  });
  addText(slide, head, { left: pos.left + 28, top: pos.top + 20, width: pos.width - 56, height: 34 }, {
    fontSize: 24,
    color,
    bold: true,
  });
  addText(slide, body, { left: pos.left + 28, top: pos.top + 66, width: pos.width - 56, height: pos.height - 82 }, {
    fontSize: 21,
    color: C.ink,
  });
}

function addFormula(slide, key, pos, fill = C.white) {
  addBox(slide, pos, fill, C.blue, 8);
  slide.images.add({
    blob: formulaImages[key],
    contentType: "image/png",
    alt: `Formula ${key}`,
    fit: "contain",
    position: {
      left: pos.left + 26,
      top: pos.top + 12,
      width: pos.width - 52,
      height: pos.height - 24,
    },
  });
}

function addBullets(slide, items, x, y, w, fs = 24, gap = 62) {
  for (const [i, item] of items.entries()) {
    const yy = y + i * gap;
    slide.shapes.add({
      geometry: "ellipse",
      position: { left: x, top: yy + 12, width: 9, height: 9 },
      fill: C.cyan,
      line: { style: "solid", fill: "none", width: 0 },
    });
    addText(slide, item, { left: x + 26, top: yy, width: w - 26, height: gap - 6 }, {
      fontSize: fs,
      color: C.ink,
    });
  }
}

function addArrow(slide, x, y) {
  slide.shapes.add({
    geometry: "rightArrow",
    position: { left: x, top: y, width: 44, height: 26 },
    fill: C.cyan,
    line: { style: "solid", fill: "none", width: 0 },
  });
}

function setNotes(slide, notes) {
  slide.speakerNotes.textFrame.setText(notes);
  slide.speakerNotes.setVisible(true);
}

// 1
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  s.shapes.add({
    geometry: "rect",
    position: { left: 0, top: 0, width: 470, height: 720 },
    fill: C.navy,
    line: { style: "solid", fill: "none", width: 0 },
  });
  addText(s, "学术汇报", { left: 72, top: 70, width: 180, height: 28 }, {
    fontSize: 18,
    color: "#BBD7FF",
    bold: true,
  });
  addText(s, "Value\nIteration\nADP", { left: 72, top: 156, width: 330, height: 190 }, {
    typeface: "Aptos Display",
    fontSize: 56,
    color: C.white,
    bold: true,
  });
  addText(s, "离散时间非线性系统的最优控制", { left: 72, top: 410, width: 340, height: 34 }, {
    fontSize: 21,
    color: "#DCEBFF",
    bold: true,
  });
  addText(s, "一句话总结", { left: 610, top: 150, width: 220, height: 34 }, {
    fontSize: 26,
    color: C.blue,
    bold: true,
  });
  addText(s, "这篇论文关注 value iteration ADP 在有限迭代停止后，得到的控制律是否稳定、是否可用。", { left: 610, top: 210, width: 520, height: 150 }, {
    fontSize: 34,
    color: C.navy,
    bold: true,
  });
  addNote(s, "汇报主线", "问题建模 -> 传统 VI 不足 -> 算法更新 -> 收敛性 -> admissibility -> 仿真与理解", { left: 610, top: 455, width: 500, height: 130 }, C.green, C.softGreen);
  setNotes(s, "老师好，我汇报的论文是 Value Iteration Adaptive Dynamic Programming for Optimal Control of Discrete-Time Nonlinear Systems。我的一句话理解是：这篇论文不是只讨论值函数最终收敛，而是关注 value iteration ADP 在有限迭代停止后，控制律是否稳定、是否真正能用。");
}

// 2
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  addTitle(s, "RESEARCH PROBLEM", "研究问题是离散时间非线性系统的最优控制", 2);
  addNote(s, "系统模型", "状态由当前状态和控制输入共同决定。", { left: 90, top: 220, width: 390, height: 150 }, C.blue);
  addFormula(s, "system", { left: 540, top: 224, width: 600, height: 86 });
  addNote(s, "性能指标", "无限时域无折扣累积代价，目标是最小化长期控制性能。", { left: 90, top: 430, width: 390, height: 150 }, C.green, C.softGreen);
  addFormula(s, "cost", { left: 540, top: 430, width: 600, height: 96 });
  setNotes(s, "论文研究的是离散时间非线性系统。系统状态满足 x_{k+1}=F(x_k,u_k)，性能指标是从当前时刻到无穷远的效用函数累加。控制目标是找到一个反馈控制律，使系统稳定，同时最小化长期代价。");
}

// 3
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  addTitle(s, "WHY THIS PAPER", "传统 value iteration 的问题在有限停止时暴露出来", 3);
  addBullets(s, [
    "初始值函数常取 V₀(x)=0，条件比较特殊。",
    "理论最优依赖 i→∞，但实际系统必须有限步停止。",
    "有限迭代得到的 vᵢ(x) 未必稳定，也未必 admissible。",
  ], 120, 220, 620, 25, 78);
  addNote(s, "本文切入点", "关键问题不是“最终会不会收敛”，而是“现在停下来能不能用”。", { left: 810, top: 250, width: 300, height: 220 }, C.amber, C.softAmber);
  setNotes(s, "作者指出传统 value iteration 有三个问题：零初始条件不够一般，理论上需要无限迭代，有限步得到的控制律不一定稳定。这篇论文最核心的切入点是第三点，即算法停止时控制律是否 admissible。");
}

// 4
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  addTitle(s, "ALGORITHM FLOW", "本文算法的核心是交替更新控制律和值函数", 4);
  const labels = ["给定初始值函数", "求迭代控制律", "更新值函数", "检查终止准则"];
  labels.forEach((label, i) => {
    const x = 118 + i * 285;
    addBox(s, { left: x, top: 265, width: 205, height: 92 }, i === 3 ? C.navy : C.white, i === 3 ? C.navy : C.line, 12);
    addText(s, label, { left: x + 18, top: 294, width: 169, height: 34 }, {
      fontSize: 22,
      color: i === 3 ? C.white : C.navy,
      bold: true,
      alignment: "center",
    });
    if (i < labels.length - 1) addArrow(s, x + 220, 298);
  });
  addNote(s, "直观解释", "当前值函数估计未来代价；控制律根据这个估计选择输入；再用新控制律更新长期代价估计。", { left: 180, top: 455, width: 920, height: 130 }, C.blue, C.white);
  setNotes(s, "算法从一个正半定初始值函数开始。每一轮先用当前值函数求控制律，再用得到的控制律更新值函数。也就是值函数指导控制律，控制律反过来更新值函数。");
}

// 5
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  addTitle(s, "KEY FORMULAS", "两个更新公式说明 value iteration ADP 怎样迭代", 5);
  addFormula(s, "policy", { left: 120, top: 220, width: 1040, height: 100 });
  addText(s, "控制律更新：选取当前代价和未来代价之和最小的输入。", { left: 150, top: 340, width: 980, height: 32 }, {
    fontSize: 23,
    color: C.muted,
    alignment: "center",
  });
  addFormula(s, "value", { left: 120, top: 430, width: 1040, height: 100 }, C.softGreen);
  addText(s, "值函数更新：用当前控制律把一步代价和下一状态价值累加。", { left: 150, top: 550, width: 980, height: 32 }, {
    fontSize: 23,
    color: C.muted,
    alignment: "center",
  });
  setNotes(s, "这里是两个核心公式。第一个公式用当前值函数求控制律，考虑当前效用和下一状态的未来代价。第二个公式用得到的控制律更新值函数。汇报时不需要逐项推导，重点讲它们是交替更新关系。");
}

// 6
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  addTitle(s, "THEORETICAL RESULT", "理论结果放宽了 value iteration 的初始条件", 6);
  addFormula(s, "limit", { left: 110, top: 230, width: 520, height: 96 });
  addBullets(s, [
    "不要求初始值函数必须为零。",
    "只要 V₀(x) 是正半定，Vᵢ(x) 最终收敛到 J*(x)。",
    "不同初始函数会改变收敛路径，但不改变最终目标。",
  ], 110, 370, 560, 24, 64);
  addNote(s, "收敛路径", "可能单调递增、单调递减，也可能非单调。论文重点是证明这些路径最终都指向最优值函数。", { left: 760, top: 245, width: 350, height: 260 }, C.green, C.softGreen);
  setNotes(s, "论文证明初始值函数不必局限于零。只要初始值函数是正半定的，迭代值函数最终收敛到最优值函数。不同初始值函数会带来不同的收敛路径，但最终目标一致。");
}

// 7
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  addTitle(s, "ADMISSIBILITY", "新终止准则解决有限迭代控制律是否可用的问题", 7);
  addNote(s, "传统收敛准则", "只能说明值函数变化很小，不能直接保证控制律稳定。", { left: 90, top: 220, width: 430, height: 150 }, C.amber, C.softAmber);
  addFormula(s, "conv", { left: 620, top: 240, width: 500, height: 82 });
  addNote(s, "本文补充条件", "满足该条件后，可证明当前迭代控制律 vᵢ(x) 是 admissible。", { left: 90, top: 430, width: 430, height: 150 }, C.green, C.softGreen);
  addFormula(s, "adm", { left: 620, top: 450, width: 500, height: 82 }, C.softGreen);
  setNotes(s, "这是我认为论文最重要的部分。传统算法只看值函数是否收敛，但这不保证当前控制律稳定。本文加入 admissibility 终止准则，满足后可以证明当前控制律是 admissible 的，也就是稳定且性能指标有限。");
}

// 8
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  addTitle(s, "SIMULATION STUDIES", "仿真实验验证的重点是终止准则的必要性", 8);
  addNote(s, "离散化倒立摆", "不同正半定初始值函数下，值函数都能收敛；满足双终止准则后控制律有效。", { left: 100, top: 230, width: 500, height: 160 }, C.blue);
  addNote(s, "离散化扭转摆", "只满足传统收敛准则时，控制律可能仍不稳定；admissibility 条件能排除这种情况。", { left: 100, top: 430, width: 500, height: 160 }, C.green, C.softGreen);
  addBox(s, { left: 760, top: 270, width: 320, height: 240 }, C.navy, C.navy, 14);
  addText(s, "仿真的结论", { left: 810, top: 320, width: 220, height: 34 }, {
    fontSize: 26,
    color: C.white,
    bold: true,
    alignment: "center",
  });
  addText(s, "不是只看曲线收敛，\n而是验证停下来之后\n控制律仍然可用。", { left: 800, top: 390, width: 240, height: 95 }, {
    fontSize: 24,
    color: C.white,
    bold: true,
    alignment: "center",
  });
  setNotes(s, "仿真部分用了两个例子：离散化倒立摆和扭转摆。它们验证了两点：不同初始函数下值函数可以收敛；更重要的是，只看传统收敛准则可能得到不稳定控制律，所以 admissibility 终止准则是必要的。");
}

// 9
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  addTitle(s, "MY UNDERSTANDING", "我的理解是论文把 VI 从理论收敛推进到有限步可用", 9);
  addBullets(s, [
    "传统分析更关注 i→∞ 时的最优性。",
    "实际系统必须在有限步停止，因此要判断当前控制律是否可用。",
    "本文的 admissibility 分析补上了这个实际应用缺口。",
  ], 130, 230, 690, 25, 78);
  addNote(s, "一句话理解", "这篇论文的价值不只是提出迭代公式，而是回答“什么时候可以安全停止”。", { left: 830, top: 270, width: 290, height: 230 }, C.blue, C.softBlue);
  setNotes(s, "我的理解是，这篇论文的核心贡献不是简单提出一个新的 value iteration 更新公式，而是把 value iteration ADP 从理论上的最终收敛推进到有限迭代停止后控制律仍然可用。");
}

// 10
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  addTitle(s, "NEXT STEP", "后续可以先复现倒立摆例子来验证理解", 10);
  addBullets(s, [
    "先实现系统模型、效用函数和 value iteration 更新。",
    "对比只用收敛准则与加入 admissibility 准则的停止结果。",
    "观察系统状态轨迹是否稳定，而不只看值函数曲线。",
    "如果时间允许，再补充神经网络近似实现。",
  ], 130, 215, 700, 24, 72);
  addNote(s, "向老师请教的问题", "复现时应优先关注理论判据，还是优先关注神经网络实现细节？", { left: 850, top: 280, width: 280, height: 210 }, C.green, C.softGreen);
  setNotes(s, "后续如果继续深入，我建议先复现倒立摆例子。重点不是单纯画出值函数收敛曲线，而是对比只用传统收敛准则和加入 admissibility 准则时，系统状态轨迹是否稳定。");
}

await fs.mkdir(PREVIEW, { recursive: true });
await fs.mkdir(LAYOUT, { recursive: true });
await fs.mkdir(QA, { recursive: true });
await fs.mkdir(OUT, { recursive: true });

await fs.writeFile(
  path.join(TMP, "source-notes.txt"),
  `Source: user-provided paper "Value Iteration Adaptive Dynamic Programming for Optimal Control of Discrete-Time Nonlinear Systems", Wei, Liu and Lin, IEEE Transactions on Cybernetics, 2016.
Deck purpose: support a short advisor-facing reading report focused on the paper's logic.
Central takeaway: the paper extends value iteration ADP from asymptotic convergence analysis toward finite-iteration usable control laws through admissibility termination criteria.
`,
  "utf8",
);
await fs.writeFile(
  path.join(TMP, "slide-plan.txt"),
  `10-slide structure requested by user:
1 title and one-sentence summary
2 research problem
3 traditional value iteration limitations
4 algorithm core flow
5 key formulas
6 theoretical convergence result
7 admissibility and new termination criteria
8 simulation studies
9 personal understanding
10 future reproduction plan
`,
  "utf8",
);

for (const [index, slide] of deck.slides.items.entries()) {
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  const png = await deck.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(path.join(PREVIEW, `${stem}.png`), new Uint8Array(await png.arrayBuffer()));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(path.join(LAYOUT, `${stem}.layout.json`), await layout.text(), "utf8");
}

const montage = await deck.export({ format: "webp", montage: true, scale: 1 });
await fs.writeFile(path.join(PREVIEW, "deck-montage.webp"), new Uint8Array(await montage.arrayBuffer()));

const pptx = await PresentationFile.exportPptx(deck);
await pptx.save(FINAL_PPTX);

await fs.writeFile(
  path.join(QA, "visual-qa.txt"),
  `PPTX exists: ${FINAL_PPTX}
Slide count: 10
Notes: speaker notes included on every slide
Formula rendering: KaTeX -> PNG -> embedded in PPTX. Iterative symbols use V_0(x), V_i(x), V_{i+1}(x), and v_i(x).
Previews rendered: yes
Remaining caveat: formulas are embedded images for visual fidelity, not editable PowerPoint equation objects.
`,
  "utf8",
);

console.log(`PPTX: ${FINAL_PPTX}`);
