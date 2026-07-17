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
const WORK = path.join(ROOT, "work/presentations/adp_report_v2");
const TMP = path.join(WORK, "tmp");
const PREVIEW = path.join(TMP, "preview");
const LAYOUT = path.join(TMP, "layout");
const QA = path.join(TMP, "qa");
const ASSET = path.join(TMP, "assets");
const OUT = path.join(ROOT, "outputs");
const FINAL_PPTX = path.join(OUT, "ADP_Value_Iteration_report_math_v3.pptx");
const SCRIPT_TXT = path.join(OUT, "ADP_report_speaker_script_v2.txt");

const W = 1280;
const H = 720;
const C = {
  bg: "#F5F8FC",
  ink: "#172033",
  muted: "#607086",
  navy: "#103052",
  blue: "#2F80ED",
  cyan: "#22AFC3",
  green: "#159A6A",
  amber: "#F4B942",
  red: "#D85555",
  line: "#DCE5EF",
  white: "#FFFFFF",
  softBlue: "#E9F3FF",
  softGreen: "#EAF9F2",
  softAmber: "#FFF7E6",
};

const paper = "Wei, Liu & Lin, IEEE TCYB 2016";
const requireFromRuntime = createRequire(
  path.join(RUNTIME_ROOT, "node/node_modules/.pnpm/node_modules/package.json"),
);
const { chromium } = requireFromRuntime("playwright-core");
const requireFromProject = createRequire(path.join(ROOT, "package.json"));
const katex = requireFromProject("katex");
const katexCss = await fs.readFile(
  requireFromProject.resolve("katex/dist/katex.min.css"),
  "utf8",
);

const math = {
  system: String.raw`x_{k+1}=F(x_k,u_k)`,
  cost: String.raw`J(x_0,u_0)=\sum_{k=0}^{\infty}U(x_k,u_k)`,
  policy: String.raw`v_i(x_k)=\underset{u_k}{\arg\min}\{U(x_k,u_k)+V_i(F(x_k,u_k))\}`,
  value: String.raw`V_{i+1}(x_k)=U(x_k,v_i(x_k))+V_i(F(x_k,v_i(x_k)))`,
  limit: String.raw`\lim_{i\to\infty}V_i(x_k)=J^*(x_k)`,
  convergence: String.raw`|V_{i+1}(x_k)-V_i(x_k)|\leq\varepsilon`,
  admissible: String.raw`V_{i+1}(x_k)-V_i(x_k)<U(x_k,v_i(x_k))`,
};

async function renderMathImages() {
  await fs.mkdir(ASSET, { recursive: true });
  const browser = await chromium.launch({
    headless: true,
    executablePath: "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
  });
  const page = await browser.newPage({
    viewport: { width: 1200, height: 240 },
    deviceScaleFactor: 3,
  });
  const out = {};
  for (const [key, latexSource] of Object.entries(math)) {
    const rendered = katex.renderToString(latexSource, {
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
          <div id="eq" style="display:inline-block;padding:16px 24px;font-size:38px;line-height:1.2;">
            ${rendered}
          </div>
        </body>
      </html>`);
    const file = path.join(ASSET, `${key}.png`);
    await page.locator("#eq").screenshot({ path: file, omitBackground: true });
    out[key] = await fs.readFile(file);
  }
  await browser.close();
  return out;
}

const formulaImages = await renderMathImages();
const slides = [
  {
    title: "Value Iteration ADP",
    notes:
      "各位老师好，我汇报的论文是 Wei、Liu 和 Lin 发表在 IEEE Transactions on Cybernetics 2016 年的文章。我的汇报重点是：对于离散时间非线性系统，value iteration ADP 如何在有限迭代后得到一个真正可用的控制律。",
  },
  {
    title: "研究对象：离散时间非线性系统的无限时域最优控制",
    notes:
      "这篇论文研究离散时间非线性系统。系统状态由 x_{k+1}=F(x_k,u_k) 描述，性能指标是从当前时刻到无穷远的效用函数累加。目标是找到反馈控制律，使系统稳定，同时最小化无限时域性能指标。",
  },
  {
    title: "传统 value iteration 的问题不只是不够快",
    notes:
      "作者指出传统 value iteration 有三个实际问题：初始值函数通常取零，理论上要无限迭代，有限迭代得到的控制律不一定 admissible。第三点是本文最重要的切入点，因为实际控制系统必须有限步停止。",
  },
  {
    title: "算法主线：控制律和值函数交替更新",
    notes:
      "算法从任意正半定初始值函数开始。每一轮先通过最小化当前效用加下一状态值函数得到控制律，再用这个控制律更新值函数。也就是用值函数改进控制律，再用控制律更新值函数。",
  },
  {
    title: "收敛性：初始值函数可以更一般",
    notes:
      "论文证明，只要初始值函数是正半定，迭代值函数最终收敛到最优值函数。不同初始值函数会改变收敛路径，可能单调递增、单调递减，也可能非单调，但最终都逼近 J star。",
  },
  {
    title: "关键点：数值收敛不等于控制律可用",
    notes:
      "传统终止准则只看两次值函数之差是否足够小。本文强调，这只能说明值函数变化小，不能保证当前控制律稳定。因此需要再检查 admissibility 判据。",
  },
  {
    title: "Algorithm 1：把停止条件分成两个分支",
    notes:
      "Algorithm 1 中，如果 V1 小于等于 V0，进入 Block 1，只用收敛准则。否则进入 Block 2，需要同时满足收敛准则和 admissibility 准则。这个设计让有限迭代停止后的控制律更有理论保证。",
  },
  {
    title: "实现方式：critic 评估，action 改进",
    notes:
      "论文用两个神经网络实现算法。critic 网络近似值函数，action 网络近似控制律。这个结构对应 ADP 里的评价和改进两个角色。",
  },
  {
    title: "Part IV 仿真：验证的是“可停止、可使用”",
    notes:
      "仿真部分包括离散化倒立摆和扭转摆。作者不仅展示不同初始函数下值函数收敛，还展示只满足收敛准则时控制律可能不稳定；加入 admissibility 准则后，有限迭代结果才更可靠。",
  },
  {
    title: "汇报结论与后续复现",
    notes:
      "我的理解是，这篇论文的贡献不是简单提出新的 value iteration 公式，而是把 value iteration 推向实际可停止、停止后控制律可用。后续代码复现可以从倒立摆例子开始，对比单一收敛准则和双终止准则。",
  },
];

function box(slide, pos, fill = C.white, line = C.line, radius = 10) {
  return slide.shapes.add({
    geometry: "roundRect",
    position: pos,
    fill,
    line: { style: "solid", fill: line, width: 1 },
    borderRadius: radius,
    shadow: "shadow-sm",
  });
}

function text(slide, value, pos, style = {}) {
  const s = slide.shapes.add({
    geometry: "textbox",
    position: pos,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  s.text = value;
  s.text.style = {
    typeface: style.typeface ?? "Aptos",
    fontSize: style.fontSize ?? 22,
    color: style.color ?? C.ink,
    bold: style.bold ?? false,
    alignment: style.alignment ?? "left",
  };
  return s;
}

function title(slide, value, kicker, n) {
  text(slide, kicker, { left: 72, top: 42, width: 520, height: 24 }, {
    fontSize: 12,
    color: C.blue,
    bold: true,
  });
  text(slide, value, { left: 72, top: 76, width: 990, height: 56 }, {
    typeface: "Aptos Display",
    fontSize: 33,
    color: C.navy,
    bold: true,
  });
  slide.shapes.add({
    geometry: "line",
    position: { left: 72, top: 148, width: 1136, height: 0 },
    fill: "none",
    line: { style: "solid", fill: C.line, width: 1 },
  });
  text(slide, paper, { left: 72, top: 672, width: 500, height: 20 }, {
    fontSize: 10,
    color: C.muted,
  });
  text(slide, String(n).padStart(2, "0"), { left: 1150, top: 665, width: 58, height: 24 }, {
    fontSize: 12,
    color: C.navy,
    bold: true,
    alignment: "right",
  });
}

function latex(slide, key, pos, fill = "#FFFFFF") {
  box(slide, pos, fill, C.blue, 8);
  slide.images.add({
    blob: formulaImages[key],
    contentType: "image/png",
    alt: `Formula ${key}`,
    fit: "contain",
    position: {
      left: pos.left + 28,
      top: pos.top + 13,
      width: pos.width - 56,
      height: pos.height - 26,
    },
  });
}

function note(slide, head, body, pos, color = C.blue, fill = C.white) {
  box(slide, pos, fill, C.line, 10);
  slide.shapes.add({
    geometry: "rect",
    position: { left: pos.left, top: pos.top, width: 8, height: pos.height },
    fill: color,
    line: { style: "solid", fill: "none", width: 0 },
  });
  text(slide, head, { left: pos.left + 26, top: pos.top + 18, width: pos.width - 48, height: 28 }, {
    fontSize: 20,
    color,
    bold: true,
  });
  text(slide, body, { left: pos.left + 26, top: pos.top + 58, width: pos.width - 48, height: pos.height - 72 }, {
    fontSize: 19,
    color: C.ink,
  });
}

function bullets(slide, items, x, y, w, fs = 22, gap = 58) {
  items.forEach((item, i) => {
    const yy = y + i * gap;
    slide.shapes.add({
      geometry: "ellipse",
      position: { left: x, top: yy + 9, width: 9, height: 9 },
      fill: C.cyan,
      line: { style: "solid", fill: "none", width: 0 },
    });
    text(slide, item, { left: x + 26, top: yy, width: w - 26, height: gap - 6 }, {
      fontSize: fs,
      color: C.ink,
    });
  });
}

function chip(slide, value, x, y, w, fill = C.softBlue, color = C.navy) {
  box(slide, { left: x, top: y, width: w, height: 40 }, fill, "none", 20);
  text(slide, value, { left: x + 12, top: y + 9, width: w - 24, height: 22 }, {
    fontSize: 15,
    color,
    bold: true,
    alignment: "center",
  });
}

function line(slide, x1, y1, x2, y2, color = C.blue, width = 3) {
  slide.shapes.add({
    geometry: "line",
    position: { left: x1, top: y1, width: x2 - x1, height: y2 - y1 },
    fill: "none",
    line: { style: "solid", fill: color, width },
  });
}

function curve(slide, x, y, mode, label) {
  box(slide, { left: x, top: y, width: 280, height: 175 }, C.white, C.line, 10);
  text(slide, label, { left: x + 22, top: y + 18, width: 210, height: 24 }, {
    fontSize: 18,
    color: C.navy,
    bold: true,
  });
  line(slide, x + 44, y + 132, x + 238, y + 132, C.line, 1);
  line(slide, x + 44, y + 58, x + 44, y + 132, C.line, 1);
  const pts =
    mode === "up"
      ? [[0, 64], [36, 54], [72, 43], [108, 33], [144, 24], [180, 16]]
      : mode === "down"
        ? [[0, 16], [36, 31], [72, 48], [108, 61], [144, 70], [180, 75]]
        : [[0, 56], [36, 25], [72, 59], [108, 38], [144, 47], [180, 39]];
  for (let i = 0; i < pts.length - 1; i++) {
    line(slide, x + 48 + pts[i][0], y + 54 + pts[i][1], x + 48 + pts[i + 1][0], y + 54 + pts[i + 1][1]);
  }
}

const deck = Presentation.create({ slideSize: { width: W, height: H } });

// 1
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  s.shapes.add({ geometry: "rect", position: { left: 0, top: 0, width: 470, height: H }, fill: C.navy, line: { style: "solid", fill: "none", width: 0 } });
  text(s, "论文汇报", { left: 72, top: 68, width: 180, height: 28 }, { fontSize: 17, color: "#BBD7FF", bold: true });
  text(s, "Value\nIteration\nADP", { left: 72, top: 158, width: 330, height: 190 }, {
    typeface: "Aptos Display",
    fontSize: 56,
    color: C.white,
    bold: true,
  });
  text(s, "离散时间非线性系统的最优控制", { left: 72, top: 408, width: 350, height: 34 }, {
    fontSize: 21,
    color: "#DCEBFF",
    bold: true,
  });
  text(s, "Qinglai Wei, Derong Liu, Hanquan Lin\nIEEE Transactions on Cybernetics, 2016", { left: 72, top: 572, width: 360, height: 55 }, {
    fontSize: 16,
    color: "#DCEBFF",
  });
  s.shapes.add({ geometry: "ellipse", position: { left: 730, top: 120, width: 320, height: 320 }, fill: C.softBlue, line: { style: "solid", fill: "none", width: 0 } });
  s.shapes.add({ geometry: "ellipse", position: { left: 818, top: 208, width: 144, height: 144 }, fill: C.blue, line: { style: "solid", fill: "none", width: 0 } });
  text(s, "VI\nADP", { left: 818, top: 244, width: 144, height: 70 }, { fontSize: 31, color: C.white, bold: true, alignment: "center" });
  chip(s, "convergence", 640, 520, 170);
  chip(s, "termination", 840, 520, 170);
  chip(s, "admissibility", 1040, 520, 170);
  s.speakerNotes.textFrame.setText(slides[0].notes);
  s.speakerNotes.setVisible(true);
}

// 2
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  title(s, slides[1].title, "PROBLEM FORMULATION", 2);
  note(s, "系统模型", "状态由当前状态和控制输入共同决定。", { left: 92, top: 215, width: 388, height: 150 }, C.blue);
  latex(s, "system", { left: 540, top: 220, width: 600, height: 82 });
  note(s, "性能指标", "无限时域无折扣累积代价，目标是最小化长期控制性能。", { left: 92, top: 420, width: 388, height: 150 }, C.green);
  latex(s, "cost", { left: 540, top: 424, width: 600, height: 92 });
  text(s, "汇报时一句话：求一个稳定系统并最小化长期代价的反馈控制律。", { left: 540, top: 552, width: 600, height: 34 }, { fontSize: 21, color: C.navy, bold: true, alignment: "center" });
  s.speakerNotes.textFrame.setText(slides[1].notes);
  s.speakerNotes.setVisible(true);
}

// 3
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  title(s, slides[2].title, "MOTIVATION", 3);
  bullets(s, [
    "初始值函数常取 V_0(x)=0，条件偏特殊。",
    "理论最优通常依赖 i\\to\\infty，实际系统不能无限迭代。",
    "有限步得到的 v_i(x) 未必稳定，也未必 admissible。",
  ], 110, 205, 550, 24, 76);
  box(s, { left: 760, top: 208, width: 360, height: 280 }, C.softAmber, "#F1D08A", 12);
  text(s, "本文真正关心的问题", { left: 802, top: 255, width: 280, height: 36 }, { fontSize: 24, color: "#8A5A00", bold: true, alignment: "center" });
  text(s, "算法在有限步停止时，\n控制律到底能不能用？", { left: 802, top: 330, width: 280, height: 90 }, { fontSize: 26, color: C.navy, bold: true, alignment: "center" });
  text(s, "这是后面 admissibility criterion 的动机。", { left: 745, top: 548, width: 400, height: 32 }, { fontSize: 19, color: C.muted, alignment: "center" });
  s.speakerNotes.textFrame.setText(slides[2].notes);
  s.speakerNotes.setVisible(true);
}

// 4
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  title(s, slides[3].title, "VALUE ITERATION ADP", 4);
  latex(s, "policy", { left: 112, top: 202, width: 1056, height: 86 });
  latex(s, "value", { left: 112, top: 332, width: 1056, height: 86 }, C.softGreen);
  const y = 515;
  ["初始化 V_0", "求 v_i", "更新 V_{i+1}", "检查停止"].forEach((v, i) => {
    const x = 168 + i * 250;
    box(s, { left: x, top: y, width: 175, height: 58 }, i === 3 ? C.navy : C.white, i === 3 ? C.navy : C.line, 12);
    text(s, v, { left: x + 8, top: y + 17, width: 159, height: 24 }, { fontSize: 18, color: i === 3 ? C.white : C.navy, bold: true, alignment: "center" });
    if (i < 3) s.shapes.add({ geometry: "rightArrow", position: { left: x + 185, top: y + 18, width: 42, height: 24 }, fill: C.cyan, line: { style: "solid", fill: "none", width: 0 } });
  });
  s.speakerNotes.textFrame.setText(slides[3].notes);
  s.speakerNotes.setVisible(true);
}

// 5
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  title(s, slides[4].title, "CONVERGENCE", 5);
  text(s, "核心结论", { left: 110, top: 195, width: 240, height: 32 }, { fontSize: 24, color: C.navy, bold: true });
  latex(s, "limit", { left: 110, top: 244, width: 500, height: 86 });
  text(s, "不要求 V_0(x) 必须为零；只要初始函数为正半定，最终收敛到最优值函数。", { left: 110, top: 365, width: 500, height: 88 }, { fontSize: 23, color: C.ink });
  curve(s, 710, 206, "up", "单调递增");
  curve(s, 710, 430, "down", "单调递减");
  curve(s, 1010, 318, "zig", "非单调");
  s.speakerNotes.textFrame.setText(slides[4].notes);
  s.speakerNotes.setVisible(true);
}

// 6
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  title(s, slides[5].title, "TERMINATION CRITERIA", 6);
  note(s, "只看收敛", "传统准则说明值函数变化很小，但不能保证控制律稳定。", { left: 92, top: 220, width: 458, height: 170 }, "#9A6A00", C.softAmber);
  latex(s, "convergence", { left: 620, top: 244, width: 500, height: 82 });
  note(s, "还要看 admissibility", "满足该不等式时，可证明当前迭代控制律 v_i(x) 是 admissible。", { left: 92, top: 430, width: 458, height: 170 }, C.green, C.softGreen);
  latex(s, "admissible", { left: 620, top: 454, width: 500, height: 82 }, C.softGreen);
  s.speakerNotes.textFrame.setText(slides[5].notes);
  s.speakerNotes.setVisible(true);
}

// 7
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  title(s, slides[6].title, "ALGORITHM 1", 7);
  box(s, { left: 100, top: 212, width: 230, height: 88 }, C.white, C.line, 12);
  text(s, "初始化\nV_0(x)=\\Phi(x)", { left: 118, top: 230, width: 194, height: 52 }, { fontSize: 19, color: C.navy, bold: true, alignment: "center" });
  s.shapes.add({ geometry: "rightArrow", position: { left: 350, top: 244, width: 48, height: 26 }, fill: C.cyan, line: { style: "solid", fill: "none", width: 0 } });
  box(s, { left: 420, top: 212, width: 230, height: 88 }, C.white, C.line, 12);
  text(s, "计算\nv_0(x), V_1(x)", { left: 438, top: 230, width: 194, height: 52 }, { fontSize: 19, color: C.navy, bold: true, alignment: "center" });
  s.shapes.add({ geometry: "rightArrow", position: { left: 670, top: 244, width: 48, height: 26 }, fill: C.cyan, line: { style: "solid", fill: "none", width: 0 } });
  box(s, { left: 740, top: 212, width: 340, height: 88 }, C.navy, C.navy, 12);
  text(s, "V_1(x)\\leq V_0(x) ?", { left: 760, top: 238, width: 300, height: 32 }, { fontSize: 23, color: C.white, bold: true, alignment: "center" });
  note(s, "Block 1", "若成立：值函数单调不增，只检查收敛准则。", { left: 150, top: 410, width: 420, height: 130 }, C.blue);
  note(s, "Block 2", "若不成立：收敛准则与 admissibility 准则都必须满足。", { left: 710, top: 410, width: 420, height: 130 }, C.green, C.softGreen);
  s.speakerNotes.textFrame.setText(slides[6].notes);
  s.speakerNotes.setVisible(true);
}

// 8
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  title(s, slides[7].title, "NEURAL NETWORK IMPLEMENTATION", 8);
  note(s, "Critic network", "近似 V_i(x)：负责评价当前状态之后的长期性能。", { left: 120, top: 240, width: 390, height: 190 }, C.blue);
  s.shapes.add({ geometry: "rightArrow", position: { left: 558, top: 315, width: 70, height: 42 }, fill: C.cyan, line: { style: "solid", fill: "none", width: 0 } });
  note(s, "Action network", "近似 v_i(x)：根据当前值函数寻找控制输入。", { left: 690, top: 240, width: 390, height: 190 }, C.green, C.softGreen);
  text(s, "汇报时不必展开网络权重；讲清楚 critic = 评价，action = 改进。", { left: 190, top: 535, width: 900, height: 36 }, { fontSize: 24, color: C.navy, bold: true, alignment: "center" });
  s.speakerNotes.textFrame.setText(slides[7].notes);
  s.speakerNotes.setVisible(true);
}

// 9
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  title(s, slides[8].title, "SIMULATION STUDIES", 9);
  note(s, "Example 1: discretized inverted pendulum", "不同正半定初始值函数下，迭代值函数收敛；25 次迭代后满足两个终止条件。", { left: 94, top: 210, width: 500, height: 160 }, C.blue);
  note(s, "Example 2: discretized torsional pendulum", "进一步展示：只满足收敛准则时，控制律仍可能不稳定。", { left: 94, top: 420, width: 500, height: 160 }, C.green, C.softGreen);
  box(s, { left: 720, top: 226, width: 360, height: 300 }, C.navy, C.navy, 14);
  text(s, "Part IV 的重点", { left: 768, top: 278, width: 264, height: 34 }, { fontSize: 25, color: C.white, bold: true, alignment: "center" });
  text(s, "不是只证明曲线收敛，\n而是证明算法停止后\n控制律仍然可用。", { left: 768, top: 350, width: 264, height: 110 }, { fontSize: 24, color: C.white, bold: true, alignment: "center" });
  s.speakerNotes.textFrame.setText(slides[8].notes);
  s.speakerNotes.setVisible(true);
}

// 10
{
  const s = deck.slides.add();
  s.background.fill = C.bg;
  title(s, slides[9].title, "TAKEAWAYS", 10);
  bullets(s, [
    "允许任意正半定初始值函数 V_0(x)。",
    "证明 V_i(x)\\to J^*(x)，并分析不同收敛路径。",
    "建立 finite-iteration 控制律的 admissibility 判据。",
    "提出“收敛 + admissibility”的双终止准则。",
  ], 120, 195, 610, 23, 70);
  box(s, { left: 800, top: 220, width: 320, height: 245 }, C.softBlue, "#BFD8F5", 14);
  text(s, "后续代码复现", { left: 840, top: 274, width: 240, height: 34 }, { fontSize: 24, color: C.navy, bold: true, alignment: "center" });
  text(s, "先从倒立摆例子开始，\n对比单一收敛准则\n和双终止准则。", { left: 840, top: 342, width: 240, height: 92 }, { fontSize: 22, color: C.ink, alignment: "center" });
  s.speakerNotes.textFrame.setText(slides[9].notes);
  s.speakerNotes.setVisible(true);
}

await fs.mkdir(PREVIEW, { recursive: true });
await fs.mkdir(LAYOUT, { recursive: true });
await fs.mkdir(QA, { recursive: true });
await fs.mkdir(OUT, { recursive: true });

await fs.writeFile(
  path.join(TMP, "source-notes.txt"),
  `Source: user-provided PDF, "Value Iteration Adaptive Dynamic Programming for Optimal Control of Discrete-Time Nonlinear Systems", Wei, Liu and Lin, IEEE Transactions on Cybernetics, 2016.
Slides 2-7 use Section II and III equations, Theorem 4, Theorem 9, and Algorithm 1.
Slide 8 summarizes the critic/action neural-network implementation.
Slide 9 summarizes Section IV simulation studies.
All displayed equations are written in LaTeX-style notation for readability.
`,
  "utf8",
);
await fs.writeFile(
  path.join(TMP, "slide-plan.txt"),
  `Create-mode deck, v2 layout.
Design: two-column academic layout, fewer cards, larger whitespace, LaTeX-style formula boxes.
Palette: navy dominant with blue/cyan and green accents.
Fonts: Aptos Display, Aptos, Cambria Math for formula boxes.
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

const speakerScript = `《Value Iteration ADP》论文汇报演讲稿 v2

第 1 页：开场
各位老师好，我汇报的论文是 Wei、Liu 和 Lin 发表在 IEEE Transactions on Cybernetics 2016 年的文章。我的汇报重点是：对于离散时间非线性系统，value iteration ADP 如何在有限迭代后得到一个真正可用的控制律。

第 2 页：研究对象
这篇论文研究离散时间非线性系统。系统状态由 x_{k+1}=F(x_k,u_k) 描述，性能指标是从当前时刻到无穷远的效用函数累加。目标是找到反馈控制律，使系统稳定，同时最小化无限时域性能指标。

第 3 页：研究动机
传统 value iteration 有三个实际问题：初始值函数常取零，理论上要无限迭代，有限迭代得到的控制律不一定 admissible。第三点是本文最重要的切入点，因为实际控制系统必须有限步停止。

第 4 页：算法主线
算法从任意正半定初始值函数开始。每一轮先通过最小化当前效用加下一状态值函数得到控制律，再用这个控制律更新值函数。也就是用值函数改进控制律，再用控制律更新值函数。

第 5 页：收敛性
论文证明，只要初始值函数是正半定，迭代值函数最终收敛到最优值函数。不同初始值函数会改变收敛路径，可能单调递增、单调递减，也可能非单调，但最终都逼近 J star。

第 6 页：终止准则
传统终止准则只看两次值函数之差是否足够小。本文强调，这只能说明值函数变化小，不能保证当前控制律稳定。因此需要再检查 admissibility 判据。

第 7 页：Algorithm 1
Algorithm 1 中，如果 V1 小于等于 V0，进入 Block 1，只用收敛准则。否则进入 Block 2，需要同时满足收敛准则和 admissibility 准则。这个设计让有限迭代停止后的控制律更有理论保证。

第 8 页：神经网络实现
论文用两个神经网络实现算法。critic 网络近似值函数，action 网络近似控制律。这个结构对应 ADP 里的评价和改进两个角色。

第 9 页：仿真实验
仿真部分包括离散化倒立摆和扭转摆。作者不仅展示不同初始函数下值函数收敛，还展示只满足收敛准则时控制律可能不稳定；加入 admissibility 准则后，有限迭代结果才更可靠。

第 10 页：总结
我的理解是，这篇论文的贡献不是简单提出新的 value iteration 公式，而是把 value iteration 推向实际可停止、停止后控制律可用。后续代码复现可以从倒立摆例子开始，对比单一收敛准则和双终止准则。
`;
await fs.writeFile(SCRIPT_TXT, speakerScript, "utf8");

await fs.writeFile(
  path.join(QA, "visual-qa.txt"),
  `PPTX generated: ${FINAL_PPTX}
Slide count: 10
Notes count: 10
All slides rendered to preview PNGs.
Visual direction: improved spacing, fewer dense cards, LaTeX-style formula notation.
Remaining caveat: formulas are displayed as LaTeX-style text, not compiled equation objects.
`,
  "utf8",
);

console.log(`PPTX: ${FINAL_PPTX}`);
console.log(`Script: ${SCRIPT_TXT}`);
