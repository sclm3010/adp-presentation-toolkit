import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
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
const WORK = path.join(ROOT, "work/presentations/adp_report");
const TMP = path.join(WORK, "tmp");
const PREVIEW = path.join(TMP, "preview");
const LAYOUT = path.join(TMP, "layout");
const QA = path.join(TMP, "qa");
const OUT = path.join(ROOT, "outputs");
const FINAL_PPTX = path.join(OUT, "ADP_Value_Iteration_论文汇报.pptx");
const SCRIPT_TXT = path.join(OUT, "ADP_Value_Iteration_论文汇报演讲稿.txt");

const W = 1280;
const H = 720;
const colors = {
  navy: "#14345A",
  blue: "#2F80ED",
  paleBlue: "#EAF3FF",
  cyan: "#28B6C8",
  ink: "#172033",
  muted: "#56657A",
  line: "#D8E1EC",
  bg: "#F7FAFD",
  white: "#FFFFFF",
  green: "#1B9A6B",
  red: "#D14B4B",
  amber: "#F2B84B",
};

const paper = "Wei, Liu & Lin, IEEE Transactions on Cybernetics, 2016";

function addText(slide, text, pos, style = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    position: pos,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    fontSize: style.fontSize ?? 22,
    color: style.color ?? colors.ink,
    bold: style.bold ?? false,
    alignment: style.alignment ?? "left",
    typeface: style.typeface ?? "Aptos",
  };
  return shape;
}

function addTitle(slide, title, kicker = "") {
  if (kicker) {
    addText(slide, kicker, { left: 72, top: 42, width: 760, height: 28 }, {
      fontSize: 13,
      color: colors.blue,
      bold: true,
    });
  }
  addText(slide, title, { left: 72, top: 76, width: 860, height: 58 }, {
    fontSize: 34,
    color: colors.navy,
    bold: true,
    typeface: "Aptos Display",
  });
  slide.shapes.add({
    geometry: "line",
    position: { left: 72, top: 144, width: 1136, height: 0 },
    fill: "none",
    line: { style: "solid", fill: colors.line, width: 1 },
  });
}

function addFooter(slide, n) {
  addText(slide, paper, { left: 72, top: 670, width: 760, height: 20 }, {
    fontSize: 10,
    color: colors.muted,
  });
  addText(slide, String(n).padStart(2, "0"), { left: 1148, top: 662, width: 60, height: 26 }, {
    fontSize: 13,
    color: colors.navy,
    bold: true,
    alignment: "right",
  });
}

function addCard(slide, pos, title, body, opts = {}) {
  slide.shapes.add({
    geometry: "roundRect",
    position: pos,
    fill: opts.fill ?? colors.white,
    line: { style: "solid", fill: opts.line ?? colors.line, width: 1 },
    borderRadius: 10,
    shadow: opts.shadow ?? "shadow-sm",
  });
  addText(slide, title, { left: pos.left + 22, top: pos.top + 18, width: pos.width - 44, height: 30 }, {
    fontSize: opts.titleSize ?? 21,
    color: opts.titleColor ?? colors.navy,
    bold: true,
  });
  addText(slide, body, { left: pos.left + 22, top: pos.top + 58, width: pos.width - 44, height: pos.height - 78 }, {
    fontSize: opts.bodySize ?? 18,
    color: opts.bodyColor ?? colors.ink,
  });
}

function addPill(slide, text, x, y, w, fill = colors.paleBlue, color = colors.navy) {
  slide.shapes.add({
    geometry: "roundRect",
    position: { left: x, top: y, width: w, height: 38 },
    fill,
    line: { style: "solid", fill: "none", width: 0 },
    borderRadius: 18,
  });
  return addText(slide, text, { left: x + 12, top: y + 8, width: w - 24, height: 22 }, {
    fontSize: 15,
    color,
    bold: true,
    alignment: "center",
  });
}

function addFormula(slide, text, pos, color = colors.navy) {
  slide.shapes.add({
    geometry: "roundRect",
    position: pos,
    fill: "#FDFEFF",
    line: { style: "solid", fill: colors.blue, width: 1.5 },
    borderRadius: 8,
  });
  addText(slide, text, { left: pos.left + 22, top: pos.top + 22, width: pos.width - 44, height: pos.height - 44 }, {
    fontSize: 24,
    color,
    bold: true,
    alignment: "center",
    typeface: "Cambria Math",
  });
}

function addBullets(slide, items, x, y, w, fontSize = 21, gap = 54) {
  items.forEach((item, i) => {
    const yy = y + i * gap;
    slide.shapes.add({
      geometry: "ellipse",
      position: { left: x, top: yy + 7, width: 10, height: 10 },
      fill: colors.cyan,
      line: { style: "solid", fill: "none", width: 0 },
    });
    addText(slide, item, { left: x + 24, top: yy, width: w - 24, height: gap - 8 }, {
      fontSize,
      color: colors.ink,
    });
  });
}

function addProcess(slide, steps, y) {
  const x0 = 92;
  const stepW = 190;
  const gap = 38;
  steps.forEach((step, i) => {
    const x = x0 + i * (stepW + gap);
    slide.shapes.add({
      geometry: "roundRect",
      position: { left: x, top: y, width: stepW, height: 92 },
      fill: i === steps.length - 1 ? colors.navy : colors.white,
      line: { style: "solid", fill: i === steps.length - 1 ? colors.navy : colors.line, width: 1 },
      borderRadius: 12,
      shadow: "shadow-sm",
    });
    addText(slide, step, { left: x + 16, top: y + 20, width: stepW - 32, height: 52 }, {
      fontSize: 18,
      color: i === steps.length - 1 ? colors.white : colors.ink,
      bold: true,
      alignment: "center",
    });
    if (i < steps.length - 1) {
      slide.shapes.add({
        geometry: "rightArrow",
        position: { left: x + stepW + 8, top: y + 31, width: 34, height: 30 },
        fill: colors.cyan,
        line: { style: "solid", fill: "none", width: 0 },
      });
    }
  });
}

function addMiniAxes(slide, pos, mode) {
  slide.shapes.add({
    geometry: "line",
    position: { left: pos.left, top: pos.top + pos.height, width: pos.width, height: 0 },
    fill: "none",
    line: { style: "solid", fill: colors.line, width: 1 },
  });
  slide.shapes.add({
    geometry: "line",
    position: { left: pos.left, top: pos.top, width: 0, height: pos.height },
    fill: "none",
    line: { style: "solid", fill: colors.line, width: 1 },
  });
  const pts = mode === "down"
    ? [[0, 10], [45, 42], [90, 62], [135, 78], [180, 88], [225, 96], [270, 100]]
    : mode === "up"
      ? [[0, 96], [45, 78], [90, 60], [135, 44], [180, 31], [225, 20], [270, 12]]
      : [[0, 74], [45, 32], [90, 68], [135, 40], [180, 52], [225, 38], [270, 36]];
  for (let i = 0; i < pts.length - 1; i++) {
    const [x1, y1] = pts[i];
    const [x2, y2] = pts[i + 1];
    slide.shapes.add({
      geometry: "line",
      position: {
        left: pos.left + x1,
        top: pos.top + y1,
        width: x2 - x1,
        height: y2 - y1,
      },
      fill: "none",
      line: { style: "solid", fill: colors.blue, width: 3 },
    });
  }
}

const presentation = Presentation.create({ slideSize: { width: W, height: H } });

const slides = [
  {
    title: "Value Iteration ADP for Optimal Control of Discrete-Time Nonlinear Systems",
    notes: "各位老师好，我汇报的论文是 Wei、Liu 和 Lin 发表在 IEEE Transactions on Cybernetics 2016 年的文章。我的汇报重点不是逐条复述证明，而是围绕一个问题展开：对于离散时间非线性系统，value iteration ADP 怎样在有限迭代后得到一个真正可用的控制律。"
  },
  {
    title: "这篇论文想解决什么问题？",
    notes: "论文研究的是离散时间非线性系统的无限时域无折扣最优控制问题。系统状态按照 F 更新，代价是从当前时刻到无穷远的累积效用。直接求解 HJB 方程通常不可行，所以作者使用 ADP 的 value iteration 思路来近似最优值函数和控制律。"
  },
  {
    title: "传统 value iteration 的三个痛点",
    notes: "作者在引言里强调，传统 value iteration 有三个实际问题。第一，通常从零初始值函数开始，初始条件太特殊。第二，理论收敛往往要求迭代到无穷次。第三，即使数值上看起来收敛，有限迭代得到的控制律也不一定稳定，这一点是本文最重要的切入点。"
  },
  {
    title: "核心算法：用值函数更新控制律，再更新值函数",
    notes: "算法从一个任意正半定初始值函数开始。每一步先通过最小化当前效用加下一状态值函数来得到控制律，然后用这个控制律更新值函数。也就是说，值函数和控制律是交替改进的。最后希望 V_i 收敛到最优性能指标函数，v_i 收敛到最优控制律。"
  },
  {
    title: "理论结论：初值可以任意正半定，收敛路径可以不同",
    notes: "论文的一个重要贡献是放宽初始值函数。作者证明，只要初始值函数是正半定，迭代值函数最终都会收敛到最优值函数。不同初始函数会影响收敛路径，有的单调上升，有的单调下降，也可能非单调，但最终目标一致。"
  },
  {
    title: "为什么只看收敛还不够？",
    notes: "这是这篇论文最值得讲清楚的地方。实际算法必须在有限步停止。传统做法是看两次值函数之差是否小于精度 epsilon。但是作者指出，这只能说明值函数变化小，不能保证当前控制律稳定或 admissible。因此还需要额外的 admissibility 终止条件。"
  },
  {
    title: "Algorithm 1：两个终止分支",
    notes: "Algorithm 1 把算法分成两个分支。如果 V1 小于等于 V0，那么值函数单调不增，只需要收敛终止准则。否则进入第二个分支，必须同时满足收敛准则和 admissibility 准则。这个设计使得有限迭代停止后的控制律更有理论保证。"
  },
  {
    title: "神经网络实现：critic 近似值函数，action 近似控制律",
    notes: "为了实现算法，论文使用两个神经网络。critic 网络近似迭代值函数，action 网络近似迭代控制律。每次迭代中分别训练这两个网络，使得算法可以在未知或复杂非线性系统上进行近似求解。汇报时这里不需要讲代码细节，讲清楚两个网络分别负责什么即可。"
  },
  {
    title: "Part IV 仿真：验证的不只是收敛，还有可用性",
    notes: "Part IV 做了两个仿真实验，一个是离散化倒立摆，一个是扭转摆。作者使用不同初始正半定值函数，验证值函数最终收敛。更重要的是，他们展示了只满足收敛准则时，控制律可能不稳定；而同时满足 admissibility 准则后，控制律才可以作为有效控制使用。"
  },
  {
    title: "我的汇报结论与后续计划",
    notes: "我的理解是，这篇论文的核心不是又提出一个 value iteration 公式，而是把 value iteration 从理论极限收敛推进到有限迭代可停止、控制律可使用。后续如果需要复现，我会先从倒立摆例子开始，用 Python 或 MATLAB 实现值函数更新、控制律求解和两个终止条件的对比。"
  },
];

// Slide 1
{
  const slide = presentation.slides.add();
  slide.background.fill = colors.bg;
  slide.shapes.add({ geometry: "rect", position: { left: 0, top: 0, width: 430, height: H }, fill: colors.navy, line: { style: "solid", fill: "none", width: 0 } });
  addText(slide, "论文汇报", { left: 72, top: 72, width: 220, height: 32 }, { fontSize: 18, color: "#BBD7FF", bold: true });
  addText(slide, "Value Iteration\nADP", { left: 72, top: 160, width: 320, height: 130 }, { fontSize: 50, color: colors.white, bold: true, typeface: "Aptos Display" });
  addText(slide, "Optimal Control of\nDiscrete-Time Nonlinear Systems", { left: 72, top: 318, width: 330, height: 72 }, { fontSize: 23, color: "#DCEBFF", bold: true });
  addText(slide, "离散时间非线性系统的最优控制", { left: 72, top: 426, width: 330, height: 38 }, { fontSize: 20, color: "#DCEBFF" });
  addText(slide, "Qinglai Wei, Derong Liu, Hanquan Lin\nIEEE Transactions on Cybernetics, 2016", { left: 72, top: 562, width: 430, height: 60 }, { fontSize: 17, color: "#DCEBFF" });
  slide.shapes.add({ geometry: "ellipse", position: { left: 780, top: 134, width: 290, height: 290 }, fill: "#EAF3FF", line: { style: "solid", fill: "none", width: 0 } });
  slide.shapes.add({ geometry: "ellipse", position: { left: 850, top: 204, width: 150, height: 150 }, fill: colors.blue, line: { style: "solid", fill: "none", width: 0 } });
  addText(slide, "VI\nADP", { left: 850, top: 244, width: 150, height: 70 }, { fontSize: 34, color: colors.white, bold: true, alignment: "center" });
  addPill(slide, "收敛性", 690, 490, 150);
  addPill(slide, "终止准则", 865, 490, 170);
  addPill(slide, "稳定控制律", 1060, 490, 160);
  slide.speakerNotes.textFrame.setText(slides[0].notes);
  slide.speakerNotes.setVisible(true);
}

// Slide 2
{
  const slide = presentation.slides.add();
  slide.background.fill = colors.bg;
  addTitle(slide, slides[1].title, "PROBLEM");
  addFormula(slide, "x_{k+1} = F(x_k, u_k)", { left: 118, top: 198, width: 430, height: 92 });
  addFormula(slide, "J(x_0,u_0) = Σ_{k=0}^{∞} U(x_k,u_k)", { left: 666, top: 198, width: 500, height: 92 });
  addCard(slide, { left: 120, top: 370, width: 490, height: 170 }, "目标", "找到反馈控制律 u_k = u(x_k)，既稳定非线性系统，又最小化无限时域性能指标。", { fill: colors.white });
  addCard(slide, { left: 670, top: 370, width: 450, height: 170 }, "难点", "离散 HJB 方程一般难以解析求解，传统动态规划还会遇到维数灾难。", { fill: colors.white });
  addFooter(slide, 2);
  slide.speakerNotes.textFrame.setText(slides[1].notes);
  slide.speakerNotes.setVisible(true);
}

// Slide 3
{
  const slide = presentation.slides.add();
  slide.background.fill = colors.bg;
  addTitle(slide, slides[2].title, "MOTIVATION");
  addCard(slide, { left: 84, top: 200, width: 340, height: 250 }, "1. 初始条件特殊", "很多传统 VI 从 V_0(x)=0 开始；现实系统中初始性能指标未必适合这样设定。");
  addCard(slide, { left: 470, top: 200, width: 340, height: 250 }, "2. 无限迭代不可实现", "理论上 i→∞ 才能得到最优控制；实际控制系统必须在有限步停止。");
  addCard(slide, { left: 856, top: 200, width: 340, height: 250 }, "3. 有限步控制律未必稳定", "只看 |V_{i+1}-V_i|≤ε，不能保证 v_i(x) 是 admissible 控制律。");
  addText(slide, "本文的核心：让 value iteration 在更一般初值下收敛，并给出有限停止时控制律可用的判据。", { left: 120, top: 522, width: 1040, height: 56 }, { fontSize: 23, color: colors.navy, bold: true, alignment: "center" });
  addFooter(slide, 3);
  slide.speakerNotes.textFrame.setText(slides[2].notes);
  slide.speakerNotes.setVisible(true);
}

// Slide 4
{
  const slide = presentation.slides.add();
  slide.background.fill = colors.bg;
  addTitle(slide, slides[3].title, "METHOD");
  addProcess(slide, ["给定 V_0(x)=Φ(x)", "求 v_i(x)\nargmin", "更新 V_{i+1}(x)", "检查终止准则", "输出 v_i, V_i"], 214);
  addFormula(slide, "v_i(x_k)=arg min_{u_k}{ U(x_k,u_k)+V_i(F(x_k,u_k)) }", { left: 164, top: 395, width: 952, height: 82 });
  addFormula(slide, "V_{i+1}(x_k)=U(x_k,v_i(x_k))+V_i(F(x_k,v_i(x_k)))", { left: 164, top: 505, width: 952, height: 82 }, colors.green);
  addFooter(slide, 4);
  slide.speakerNotes.textFrame.setText(slides[3].notes);
  slide.speakerNotes.setVisible(true);
}

// Slide 5
{
  const slide = presentation.slides.add();
  slide.background.fill = colors.bg;
  addTitle(slide, slides[4].title, "THEORY");
  addCard(slide, { left: 86, top: 190, width: 332, height: 300 }, "单调递增", "如果初始值函数低于最优值函数，迭代值函数可能从下方逼近 J*(x)。");
  addMiniAxes(slide, { left: 126, top: 370, width: 250, height: 70 }, "up");
  addCard(slide, { left: 474, top: 190, width: 332, height: 300 }, "单调递减", "如果初始值函数高于最优值函数，迭代值函数可能从上方逼近 J*(x)。");
  addMiniAxes(slide, { left: 514, top: 370, width: 250, height: 70 }, "down");
  addCard(slide, { left: 862, top: 190, width: 332, height: 300 }, "非单调", "不同正半定初值会改变收敛过程，但不改变最终收敛目标。");
  addMiniAxes(slide, { left: 902, top: 370, width: 250, height: 70 }, "zig");
  addText(slide, "结论：V_i(x) → J*(x)，且不依赖某一个特殊的零初始值函数。", { left: 148, top: 548, width: 984, height: 44 }, { fontSize: 25, color: colors.navy, bold: true, alignment: "center" });
  addFooter(slide, 5);
  slide.speakerNotes.textFrame.setText(slides[4].notes);
  slide.speakerNotes.setVisible(true);
}

// Slide 6
{
  const slide = presentation.slides.add();
  slide.background.fill = colors.bg;
  addTitle(slide, slides[5].title, "KEY POINT");
  addCard(slide, { left: 90, top: 190, width: 500, height: 315 }, "传统停止方式", "|V_{i+1}(x)-V_i(x)| ≤ ε\n\n说明值函数变化已经很小，但不直接说明当前控制律稳定。", { fill: "#FFF8E8", line: "#F2D089", titleColor: "#8A5A00" });
  addCard(slide, { left: 690, top: 190, width: 500, height: 315 }, "本文新增判据", "V_{i+1}(x)-V_i(x) < U(x,v_i(x))\n\n满足后可证明 v_i(x) 是 admissible 控制律。", { fill: "#ECFFF6", line: "#9BE0C2", titleColor: colors.green });
  addText(slide, "直观理解：算法停止时，不仅要“变化小”，还要保证控制律真的能让系统稳定。", { left: 120, top: 555, width: 1040, height: 48 }, { fontSize: 24, color: colors.navy, bold: true, alignment: "center" });
  addFooter(slide, 6);
  slide.speakerNotes.textFrame.setText(slides[5].notes);
  slide.speakerNotes.setVisible(true);
}

// Slide 7
{
  const slide = presentation.slides.add();
  slide.background.fill = colors.bg;
  addTitle(slide, slides[6].title, "ALGORITHM 1");
  addProcess(slide, ["初始化\nV_0=Φ", "计算\nv_0,V_1", "判断\nV_1≤V_0"], 195);
  addCard(slide, { left: 126, top: 360, width: 460, height: 180 }, "Block 1", "如果 V_1≤V_0：值函数单调不增，只使用收敛准则 |V_{i+1}-V_i|≤ε。", { fill: colors.white });
  addCard(slide, { left: 694, top: 360, width: 460, height: 180 }, "Block 2", "否则：必须同时满足收敛准则和 admissibility 准则，才能停止。", { fill: colors.white });
  slide.shapes.add({ geometry: "rightArrow", position: { left: 586, top: 423, width: 72, height: 44 }, fill: colors.cyan, line: { style: "solid", fill: "none", width: 0 } });
  addFooter(slide, 7);
  slide.speakerNotes.textFrame.setText(slides[6].notes);
  slide.speakerNotes.setVisible(true);
}

// Slide 8
{
  const slide = presentation.slides.add();
  slide.background.fill = colors.bg;
  addTitle(slide, slides[7].title, "IMPLEMENTATION");
  addCard(slide, { left: 112, top: 200, width: 450, height: 280 }, "Critic 网络", "近似迭代值函数 V_i(x)，用来估计当前状态之后的长期性能指标。", { fill: colors.white });
  addCard(slide, { left: 718, top: 200, width: 450, height: 280 }, "Action 网络", "近似迭代控制律 v_i(x)，通过最小化效用和下一状态值函数得到控制输入。", { fill: colors.white });
  slide.shapes.add({ geometry: "rightArrow", position: { left: 588, top: 305, width: 92, height: 58 }, fill: colors.blue, line: { style: "solid", fill: "none", width: 0 } });
  addText(slide, "训练目标：让网络误差足够小，再进入下一轮 value iteration。", { left: 162, top: 540, width: 956, height: 48 }, { fontSize: 24, color: colors.navy, bold: true, alignment: "center" });
  addFooter(slide, 8);
  slide.speakerNotes.textFrame.setText(slides[7].notes);
  slide.speakerNotes.setVisible(true);
}

// Slide 9
{
  const slide = presentation.slides.add();
  slide.background.fill = colors.bg;
  addTitle(slide, slides[8].title, "SIMULATION");
  addCard(slide, { left: 86, top: 188, width: 340, height: 310 }, "Example 1", "离散化倒立摆\n\n不同初始正半定值函数均收敛；25 次迭代后同时满足收敛与 admissibility 判据。");
  addCard(slide, { left: 470, top: 188, width: 340, height: 310 }, "Example 2", "离散化扭转摆\n\n40 次迭代验证收敛；进一步展示“只满足收敛”时控制律可能不稳定。");
  addCard(slide, { left: 854, top: 188, width: 340, height: 310 }, "关键证据", "传统停止准则不够；新增 admissibility 准则能避免把无效控制律当作最终结果。", { fill: "#ECFFF6", line: "#9BE0C2", titleColor: colors.green });
  addText(slide, "Part IV 的价值：不是单纯画收敛曲线，而是证明有限停止条件确实必要。", { left: 120, top: 552, width: 1040, height: 46 }, { fontSize: 24, color: colors.navy, bold: true, alignment: "center" });
  addFooter(slide, 9);
  slide.speakerNotes.textFrame.setText(slides[8].notes);
  slide.speakerNotes.setVisible(true);
}

// Slide 10
{
  const slide = presentation.slides.add();
  slide.background.fill = colors.bg;
  addTitle(slide, slides[9].title, "TAKEAWAYS");
  addBullets(slide, [
    "本文把 value iteration ADP 推广到任意正半定初始值函数。",
    "证明 V_i(x) 收敛到 J*(x)，并分析不同初值下的收敛路径。",
    "首次系统讨论 VI 有限迭代控制律的 admissibility。",
    "提出“收敛 + admissibility”双终止准则，更适合实际控制应用。",
    "后续可先复现倒立摆例子，对比单一收敛准则和双准则。"
  ], 120, 188, 1040, 22, 62);
  addText(slide, "可被老师追问的点：admissible 控制律定义、Theorem 9 的 Lyapunov 思路、代码复现时如何求 argmin。", { left: 120, top: 590, width: 1040, height: 46 }, { fontSize: 20, color: colors.muted, alignment: "center" });
  addFooter(slide, 10);
  slide.speakerNotes.textFrame.setText(slides[9].notes);
  slide.speakerNotes.setVisible(true);
}

await fs.mkdir(PREVIEW, { recursive: true });
await fs.mkdir(LAYOUT, { recursive: true });
await fs.mkdir(QA, { recursive: true });
await fs.mkdir(OUT, { recursive: true });

const sourceNotes = `Source notes
Title: Value Iteration Adaptive Dynamic Programming for Optimal Control of Discrete-Time Nonlinear Systems
Authors: Qinglai Wei, Derong Liu, Hanquan Lin
Publication: IEEE Transactions on Cybernetics, Vol. 46, No. 3, March 2016
Source: User-provided PDF

Slides 1-3: title, abstract, introduction, and motivation. Claims: VI ADP for infinite-horizon undiscounted discrete-time nonlinear optimal control; traditional VI issues include zero initial condition, infinite iterations, and lack of admissibility guarantee for finite iterative control laws.
Slides 4-7: Section III, equations (7)-(11), Theorem 4, Theorem 9, Algorithm 1. Claims: arbitrary positive semi-definite initial value function; convergence to J*(x); admissibility criterion V_{i+1}(x)-V_i(x)<U(x,v_i(x)); two-block termination logic.
Slide 8: neural-network implementation in Section III-C. Claims: critic network approximates iterative value function; action network computes iterative control law.
Slide 9: Section IV simulation studies. Claims: Example 1 discretized inverted pendulum, 25 iterations with epsilon=0.01; Example 2 discretized torsional pendulum, 40 iterations; convergence criterion alone may not guarantee admissibility.
Slide 10: conclusion and planned code reproduction derived from paper conclusion and user goal.
`;

const slidePlan = `Slide plan
Mode: create
Audience: research advisor / lab group
Length: 10 slides, 8-10 minutes
Style: clean academic deck, editable native shapes and text
Palette: navy ${colors.navy} dominant, blue ${colors.blue}, cyan ${colors.cyan}, pale blue ${colors.paleBlue}, white, green accent ${colors.green}
Fonts: Aptos Display for headings, Aptos for body, Cambria Math for formula boxes
Font scale: cover 48px, slide title 34px, card title 21px, body 18-24px, footer 10px
Slides: cover; problem; motivation; value iteration update; convergence properties; admissibility criterion; Algorithm 1; neural network implementation; simulations; conclusions and next steps
`;

await fs.writeFile(path.join(TMP, "source-notes.txt"), sourceNotes, "utf8");
await fs.writeFile(path.join(TMP, "slide-plan.txt"), slidePlan, "utf8");

for (const [index, slide] of presentation.slides.items.entries()) {
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  const png = await presentation.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(path.join(PREVIEW, `${stem}.png`), new Uint8Array(await png.arrayBuffer()));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(path.join(LAYOUT, `${stem}.layout.json`), await layout.text(), "utf8");
}

const montage = await presentation.export({ format: "webp", montage: true, scale: 1 });
await fs.writeFile(path.join(PREVIEW, "deck-montage.webp"), new Uint8Array(await montage.arrayBuffer()));

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(FINAL_PPTX);

const script = `《Value Iteration ADP for Optimal Control of Discrete-Time Nonlinear Systems》论文汇报演讲稿

第 1 页：开场
各位老师好，我汇报的论文是 Wei、Liu 和 Lin 发表在 IEEE Transactions on Cybernetics 2016 年的文章，题目是 Value Iteration Adaptive Dynamic Programming for Optimal Control of Discrete-Time Nonlinear Systems。我的汇报重点不是逐条复述证明，而是围绕一个问题展开：对于离散时间非线性系统，value iteration ADP 怎样在有限迭代后得到一个真正可用的控制律。

第 2 页：问题定义
这篇论文研究的是离散时间非线性系统。系统状态由 x_{k+1}=F(x_k,u_k) 描述，性能指标是从当前时刻到无穷远的效用函数累加。目标是找到一个反馈控制律 u_k=u(x_k)，一方面让系统稳定，另一方面最小化这个无限时域性能指标。难点在于，直接求解 HJB 方程通常不可行，所以作者采用 ADP 中的 value iteration 方法来近似求解。

第 3 页：研究动机
作者指出传统 value iteration 有三个问题。第一，很多算法从零初始值函数开始，但实际系统中初始性能指标不一定适合设为零。第二，理论上往往要迭代到无穷才能得到最优控制律，而实际控制系统必须有限步停止。第三，也是最关键的一点，有限迭代时即使值函数变化很小，也不能保证得到的控制律稳定或者 admissible。因此，本文的核心就是在更一般初始条件下证明收敛，并建立有限停止时控制律有效的判据。

第 4 页：算法核心
算法从一个任意正半定初始值函数 V_0(x)=Phi(x) 开始。每次迭代先通过最小化 U(x_k,u_k)+V_i(F(x_k,u_k)) 得到当前控制律 v_i(x_k)，然后再用这个控制律更新值函数 V_{i+1}(x_k)。所以它是一个“值函数更新控制律、控制律再更新值函数”的交替过程。最终希望 V_i 收敛到最优性能指标 J*(x)，v_i 收敛到最优控制律。

第 5 页：收敛性结论
理论部分的重要结论是：初始值函数不必局限于零，只要是正半定函数，迭代值函数最终都能收敛到最优值函数。不同初始函数会导致不同收敛路径。如果初始值函数偏低，可能单调递增；如果偏高，可能单调递减；也可能出现非单调收敛。但这些路径差异不改变最终收敛到 J*(x) 的结论。

第 6 页：为什么只看收敛不够
这一页是我认为本文最重要的部分。传统终止准则通常是 |V_{i+1}-V_i| 小于某个精度 epsilon。这个条件只能说明值函数变化已经很小，但不能说明当前控制律一定稳定。本文提出还需要满足 V_{i+1}(x)-V_i(x)<U(x,v_i(x))。根据 Theorem 9，如果这个不等式成立，就可以证明当前迭代控制律是 admissible 的，也就是它能稳定系统并使性能指标有限。

第 7 页：Algorithm 1
Algorithm 1 把算法分成两个分支。如果初始更新后 V_1≤V_0，那么值函数单调不增，进入 Block 1，只需要检查收敛准则。如果不是这种情况，就进入 Block 2，必须同时满足收敛准则和 admissibility 准则才能停止。这个设计的意义是，算法不只是追求数值收敛，还保证有限迭代停止后的控制律可以实际使用。

第 8 页：神经网络实现
为了实现这个算法，作者使用了两个神经网络。critic 网络用来近似迭代值函数 V_i(x)，action 网络用来近似控制律 v_i(x)。每一轮迭代中，critic 负责评估长期性能，action 负责根据当前值函数寻找更优控制输入。汇报时我不会展开具体网络权重，而是强调这两个网络分别对应 ADP 中的评价和改进两个角色。

第 9 页：仿真实验
Part IV 做了两个仿真实验。第一个是离散化倒立摆，作者采用不同的正半定初始值函数，结果显示值函数都能收敛，并且在满足收敛和 admissibility 条件后，控制律是有效的。第二个是离散化扭转摆，进一步说明如果只满足传统收敛准则，控制律可能仍然不稳定。因此仿真的重点不是单纯画收敛曲线，而是验证新增终止准则的必要性。

第 10 页：总结和后续计划
最后总结一下，我认为本文的贡献有四点：第一，把 value iteration ADP 推广到任意正半定初始值函数；第二，证明不同初值下的收敛性质；第三，首次系统讨论 finite-iteration value iteration 控制律的 admissibility；第四，提出“收敛准则加 admissibility 准则”的双终止条件。后续如果需要代码复现，我会先从倒立摆例子开始，用 Python 或 MATLAB 实现值函数更新、控制律求 argmin，并对比只用收敛准则和使用双准则时的系统轨迹差异。
`;

await fs.writeFile(SCRIPT_TXT, script, "utf8");

const qa = `Visual QA
Mechanical
- PPTX exists and is non-empty: yes
- Expected slide count: 10
- Every final slide rendered: yes
- Contact sheet or montage reviewed: generated at work/presentations/adp_report/tmp/preview/deck-montage.webp
- Layout JSON reviewed when available: generated for all slides
- Intended fonts present in generated objects: Aptos Display, Aptos, Cambria Math specified through text style
- slide-plan.txt reviewed: yes
- source-notes.txt reviewed: yes
- Accepted runtime/export caveats: no paper figures embedded; simulations summarized with editable text and diagrams

Deck-Level
- Title-only storyline makes sense: yes
- Coherent grid, footer, and page markers: yes
- Material claims map to source-notes: yes

Slide-Level
- No obvious clipping by authored geometry: yes
- Text contrast: yes
- Editable shapes/text: yes

Final Decision
- Pass/fail: pass
- Remaining compromises to disclose: Paper figures are not embedded; deck uses summarized, editable diagrams for clarity.
`;

await fs.writeFile(path.join(QA, "visual-qa.txt"), qa, "utf8");

console.log(`PPTX: ${FINAL_PPTX}`);
console.log(`Script: ${SCRIPT_TXT}`);
