import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const workspaceDir = "/home/izu/Projects/gradient-descent";
const skillDir = "/home/izu/.codex/plugins/cache/openai-primary-runtime/presentations/26.923.10815/skills/presentations";
const buildDir = path.join(workspaceDir, "tmp/pptx-build");
const candidatePath = path.join(buildDir, "reader-facing-v4-final-candidate.pptx");
const outputPath = path.join(workspaceDir, "presentation/gradient-descent-seminar-reader-facing-v4-final.pptx");

const { resolvePresentationFont, makeNativeBulletParagraphs, finalizePresentation } = await import(
  pathToFileURL(path.join(skillDir, "container_tools/artifact_tool_utils.mjs")).href,
);

await fs.mkdir(buildDir, { recursive: true });
await fs.mkdir(path.dirname(outputPath), { recursive: true });
const font = resolvePresentationFont({ availableFonts: ["Arial"] });

const W = 1280;
const H = 720;
const TOTAL_SLIDES = 28;
const C = {
  ink: "#17243A",
  muted: "#5D6B80",
  blue: "#3158C9",
  blueLight: "#EDF2FF",
  teal: "#13877F",
  tealLight: "#E9F6F3",
  amber: "#A9610B",
  amberLight: "#FFF4E2",
  red: "#B53D4A",
  redLight: "#FCECEF",
  line: "#DCE3ED",
  pale: "#F4F7FB",
  paper: "#FBFCFE",
  navy: "#14223B",
  white: "#FFFFFF",
  purple: "#6652A3",
};

const sectionColors = {
  Foundations: C.blue,
  Geometry: C.blue,
  Stability: C.amber,
  Optimizers: C.teal,
  Application: C.purple,
  Summary: C.blue,
};

const presentation = Presentation.create({ slideSize: { width: W, height: H } });
presentation.theme.colorScheme = {
  name: "Reader Facing Seminar",
  themeColors: {
    accent1: C.blue, accent2: C.teal, accent3: C.amber, accent4: C.red,
    accent5: C.purple, accent6: "#4C7A34", bg1: C.white, bg2: C.pale,
    tx1: C.ink, tx2: C.muted, dk1: "#000000", dk2: C.ink,
    lt1: C.white, lt2: C.line, hlink: C.blue, folHlink: C.purple,
  },
};

const pngCache = new Map();
async function pngFigure(name) {
  if (!pngCache.has(name)) pngCache.set(name, await fs.readFile(path.join(workspaceDir, "figures/png", name)));
  return pngCache.get(name);
}

const svgCache = new Map();
async function svgFigure(name) {
  if (!svgCache.has(name)) svgCache.set(name, await fs.readFile(path.join(workspaceDir, "figures/svg", name), "utf8"));
  return svgCache.get(name);
}

const gifCache = new Map();
async function gifFigure(name) {
  if (!gifCache.has(name)) gifCache.set(name, await fs.readFile(path.join(workspaceDir, "figures/gif", name)));
  return gifCache.get(name);
}

function shape(slide, geometry, position, options = {}) {
  return slide.shapes.add({
    geometry,
    position,
    fill: options.fill ?? "none",
    line: options.line ?? { fill: "none", width: 0 },
    borderRadius: options.borderRadius,
    shadow: options.shadow,
  });
}

function text(slide, value, position, style = {}) {
  const s = shape(slide, "textbox", position, { fill: "none", line: { fill: "none", width: 0 } });
  s.text = value;
  s.text.style = {
    typeface: font,
    fontSize: style.fontSize ?? 26,
    bold: style.bold ?? false,
    italic: style.italic ?? false,
    color: style.color ?? C.ink,
    alignment: style.alignment ?? "left",
    verticalAlignment: style.verticalAlignment ?? "top",
    autoFit: style.autoFit ?? "shrinkText",
    wrap: "square",
    lineSpacing: style.lineSpacing ?? 1.14,
    insets: style.insets ?? { top: 0, right: 0, bottom: 0, left: 0 },
  };
  return s;
}

function bullets(slide, items, position, style = {}) {
  const s = shape(slide, "textbox", position, { fill: "none", line: { fill: "none", width: 0 } });
  s.text = makeNativeBulletParagraphs(items, {
    marginLeftPoints: style.marginLeftPoints ?? 22,
    hangingPoints: style.hangingPoints ?? 11,
    spaceAfterPoints: style.spaceAfterPoints ?? 14,
  });
  s.text.style = {
    typeface: font,
    fontSize: style.fontSize ?? 26,
    color: style.color ?? C.ink,
    autoFit: "shrinkText",
    wrap: "square",
    lineSpacing: style.lineSpacing ?? 1.14,
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  };
  return s;
}

function addPng(slide, bytes, position, alt) {
  return slide.images.add({ blob: bytes, contentType: "image/png", alt, fit: "contain", position });
}

function addSvg(slide, source, position, alt) {
  return slide.images.add({ svg: source, alt, fit: "contain", position });
}

function addGif(slide, bytes, position, alt) {
  return slide.images.add({ blob: bytes, contentType: "image/gif", alt, fit: "contain", position });
}

function divider(slide, top = 116, color = C.line) {
  shape(slide, "line", { left: 64, top, width: 1152, height: 0 }, {
    line: { style: "solid", fill: color, width: 1.5 },
  });
}

function footer(slide, number, section) {
  const accent = sectionColors[section] ?? C.blue;
  shape(slide, "line", { left: 64, top: 672, width: 1152, height: 0 }, {
    line: { style: "solid", fill: C.line, width: 1.5 },
  });
  shape(slide, "line", { left: 64, top: 672, width: 1152 * number / TOTAL_SLIDES, height: 0 }, {
    line: { style: "solid", fill: accent, width: 3 },
  });
  text(slide, `${String(number).padStart(2, "0")} / ${TOTAL_SLIDES}`, { left: 1120, top: 680, width: 96, height: 18 }, {
    fontSize: 12, bold: true, color: accent, alignment: "right",
  });
}

function header(slide, title, section) {
  const accent = sectionColors[section] ?? C.blue;
  text(slide, section.toUpperCase(), { left: 64, top: 28, width: 700, height: 18 }, {
    fontSize: 13, bold: true, color: accent,
  });
  const titleSize = title.length > 50 ? 38 : title.length > 38 ? 41 : 44;
  text(slide, title, { left: 64, top: 51, width: 1152, height: 52 }, {
    fontSize: titleSize, bold: true, color: C.ink,
  });
  divider(slide, 116, accent);
}

function note(slide, lines) {
  slide.speakerNotes.textFrame.setText(lines);
}

function newSlide(title, section, number, notes) {
  const slide = presentation.slides.add();
  slide.background.fill = C.paper;
  header(slide, title, section);
  footer(slide, number, section);
  if (notes) note(slide, notes);
  return slide;
}

function callout(slide, label, body, position, colors = { fill: C.blueLight, accent: C.blue }) {
  shape(slide, "line", { left: position.left, top: position.top + 4, width: 0, height: Math.max(24, position.height - 8) }, {
    line: { style: "solid", fill: colors.accent, width: 3 },
  });
  text(slide, label.toUpperCase(), { left: position.left + 16, top: position.top + 8, width: position.width - 20, height: 18 }, {
    fontSize: 13, bold: true, color: colors.accent,
  });
  text(slide, body, { left: position.left + 16, top: position.top + 34, width: position.width - 20, height: position.height - 40 }, {
    fontSize: 24, color: C.ink, verticalAlignment: "middle",
  });
}

function equation(slide, body, position, color = C.ink, size = 32) {
  shape(slide, "line", { left: position.left, top: position.top + 10, width: 0, height: Math.max(24, position.height - 20) }, {
    line: { style: "solid", fill: color, width: 2.5 },
  });
  const s = shape(slide, "textbox", { left: position.left + 20, top: position.top, width: position.width - 20, height: position.height });
  s.text = body;
  s.text.style = {
    typeface: font, fontSize: Math.max(size, 30), color,
    alignment: "left", verticalAlignment: "middle", autoFit: "shrinkText",
    insets: { top: 8, right: 4, bottom: 8, left: 4 },
  };
  return s;
}

// 1. Cover
{
  const slide = presentation.slides.add();
  slide.background.fill = C.navy;
  shape(slide, "rect", { left: 0, top: 0, width: 12, height: H }, { fill: C.teal, line: { fill: "none", width: 0 } });
  text(slide, "GRADIENT DESCENT", { left: 88, top: 92, width: 500, height: 24 }, { fontSize: 15, bold: true, color: "#9EB5F8" });
  text(slide, "Gradient Descent", { left: 84, top: 164, width: 800, height: 86 }, { fontSize: 64, bold: true, color: C.white, lineSpacing: 0.96 });
  text(slide, "and optimizer variants", { left: 88, top: 260, width: 790, height: 60 }, { fontSize: 40, bold: true, color: "#AFC0F5" });
  text(slide, "How local slope becomes a parameter update", { left: 88, top: 366, width: 690, height: 42 }, { fontSize: 26, color: "#D1D9E8" });
  equation(slide, "θₜ₊₁ = θₜ − η ∇f(θₜ)", { left: 88, top: 476, width: 650, height: 76 }, "#B7CBFF", 38);
  addPng(slide, await pngFigure("seminar-web-qr.png"), { left: 932, top: 160, width: 236, height: 236 }, "QR code for the web companion");
  text(slide, "WEB COMPANION", { left: 914, top: 420, width: 272, height: 22 }, { fontSize: 14, bold: true, color: "#9EB5F8", alignment: "center" });
  text(slide, "https://1zuki.github.io/gradient-descent/", { left: 900, top: 454, width: 300, height: 30 }, { fontSize: 16, color: C.white, alignment: "center", verticalAlignment: "middle" });
  text(slide, "SOURCE CODE", { left: 914, top: 520, width: 272, height: 22 }, { fontSize: 14, bold: true, color: "#9EB5F8", alignment: "center" });
  text(slide, "github.com/1zuki/numpy-cnn", { left: 900, top: 552, width: 300, height: 30 }, { fontSize: 16, color: C.white, alignment: "center", verticalAlignment: "middle" });
  text(slide, "A visual lecture in local optimization", { left: 88, top: 626, width: 520, height: 24 }, { fontSize: 17, color: "#AAB6C9" });
  text(slide, `01 / ${TOTAL_SLIDES}`, { left: 1120, top: 680, width: 96, height: 20 }, { fontSize: 13, bold: true, color: "#AFC0F5", alignment: "right" });
  note(slide, [
    "Opening prompt: The deck follows one question: how does local slope become a parameter update?",
    "Source: PROJECT_PLAN.md and docs/TALK_SCRIPT.md.",
  ]);
}

// 2. Concrete motivation
{
  const slide = newSlide("A model can be a line", "Foundations", 2, [
    "Start with a familiar problem: fit a line to observations before introducing abstract parameters.",
    "Ask: what can change in the line? The slope w and intercept b.",
    "Source: PROJECT_PLAN.md, sections 1 and 7.",
  ]);
  equation(slide, "ŷ = wx + b", { left: 86, top: 164, width: 380, height: 76 }, C.blue, 38);
  equation(slide, "L(w,b) = 1/n Σᵢ (wxᵢ + b − yᵢ)²", { left: 86, top: 286, width: 640, height: 82 }, C.teal, 31);
  text(slide, "Three observations", { left: 86, top: 416, width: 300, height: 32 }, { fontSize: 28, bold: true });
  text(slide, "(x, y) = (1, 2), (2, 3), (3, 5)", { left: 86, top: 466, width: 540, height: 42 }, { fontSize: 31, color: C.ink });
  callout(slide, "Why minimize?", "Different values of w and b produce different predictions. The squared error gives us one number to make smaller.", { left: 760, top: 218, width: 370, height: 206 }, { fill: C.tealLight, accent: C.teal });
  text(slide, "The loss turns “fit the data” into an optimization problem.", { left: 760, top: 500, width: 390, height: 72 }, { fontSize: 28, bold: true, color: C.blue });
}

// 3. Parameters become a vector
{
  const slide = newSlide("Two parameters make a two-dimensional search", "Foundations", 3, [
    "Connect the line-fitting example to the notation used later.",
    "Explain that θ is pronounced “theta” and stores all adjustable parameters.",
    "Source: docs/MATH_NOTES.md and PROJECT_PLAN.md.",
  ]);
  equation(slide, "θ = (w, b)", { left: 86, top: 172, width: 360, height: 80 }, C.blue, 38);
  equation(slide, "f(θ) = L(w,b)", { left: 86, top: 304, width: 420, height: 80 }, C.teal, 36);
  bullets(slide, [
    "One point in parameter space is one candidate line.",
    "Changing w tilts the line.",
    "Changing b shifts the line vertically.",
  ], { left: 86, top: 444, width: 540, height: 150 }, { fontSize: 28, spaceAfterPoints: 10 });
  callout(slide, "The abstract picture", "Training searches through parameter space for a point with a lower loss.", { left: 760, top: 248, width: 370, height: 178 }, { fill: C.blueLight, accent: C.blue });
}

// 4. Partial derivatives
{
  const slide = newSlide("A partial derivative holds the other coordinate fixed", "Foundations", 4, [
    "Work through the meaning of holding a coordinate fixed, rather than treating partial derivatives as a notation jump.",
    "Ask students to differentiate x² + 4 with respect to x before revealing 2x.",
    "Source: docs/MATH_NOTES.md.",
  ]);
  equation(slide, "f(x,y) = x² + y²", { left: 86, top: 164, width: 480, height: 76 }, C.blue, 36);
  equation(slide, "hold y = 2:  f(x,2) = x² + 4\n∂f/∂x = 2x", { left: 86, top: 286, width: 570, height: 126 }, C.teal, 31);
  equation(slide, "hold x = 1:  f(1,y) = 1 + y²\n∂f/∂y = 2y", { left: 86, top: 452, width: 570, height: 126 }, C.amber, 31);
  callout(slide, "Meaning", "Each partial derivative asks how the loss changes when one parameter moves and the others stay fixed.", { left: 760, top: 264, width: 370, height: 188 }, { fill: C.tealLight, accent: C.teal });
}

// 5. Gradient field
{
  const slide = newSlide("The gradient lives in parameter space", "Geometry", 5, [
    "Pause at (1,2) and ask for the uphill and downhill directions before showing the values.",
    "Clarify that the arrows are vectors in the x-y parameter plane. They are not physical objects sliding on the 3D surface.",
    "For this quadratic only, the gradient magnitude grows with distance from the minimum.",
    "Source figure: figures/svg/01_gradient_field_sphere.svg.",
  ]);
  addSvg(slide, await svgFigure("01_gradient_field_sphere.svg"), { left: 680, top: 140, width: 486, height: 500 }, "Gradient vector field for x squared plus y squared");
  equation(slide, "at (1,2):  ∇f = (2,4)\ndownhill:  −∇f = (−2,−4)", { left: 86, top: 180, width: 540, height: 132 }, C.blue, 31);
  bullets(slide, [
    "The gradient points toward local increase.",
    "The negative gradient points toward local decrease.",
    "It is perpendicular to the contour through the current point.",
  ], { left: 86, top: 390, width: 530, height: 190 }, { fontSize: 27, spaceAfterPoints: 10 });
}

// 6. Surface and direction
{
  const slide = newSlide("Why the negative gradient decreases the function", "Foundations", 6, [
    "State the missing assumption explicitly: the direction vector has unit length.",
    "Derive the extremum using the dot-product identity and Cauchy–Schwarz.",
    "Qualify the Taylor statement: nonzero gradient, positive and sufficiently small η.",
    "Source: docs/MATH_NOTES.md, directional derivative and Taylor sections.",
  ]);
  addSvg(slide, await svgFigure("02_surface_sphere_3d.svg"), { left: 760, top: 150, width: 390, height: 390 }, "Three-dimensional bowl for x squared plus y squared");
  equation(slide, "||u|| = 1\nDᵤf = ∇f · u = ||∇f|| cos φ", { left: 86, top: 164, width: 610, height: 126 }, C.blue, 31);
  equation(slide, "largest increase: φ = 0\nlargest decrease: φ = π", { left: 86, top: 332, width: 560, height: 108 }, C.red, 30);
  equation(slide, "f(θ + δ) ≈ f(θ) + ∇f(θ)ᵀδ\nδ = −η∇f(θ)  ⇒  local decrease", { left: 86, top: 486, width: 630, height: 126 }, C.teal, 28);
}

// 7. Update and animation
{
  const slide = newSlide("Gradient descent repeats one update", "Foundations", 7, ["Source: PROJECT_PLAN.md, sections 1 and 7.3; animated source: figures/gif/gradient_descent.gif."]);
  equation(slide, "θₜ₊₁ = θₜ − η∇f(θₜ)", { left: 86, top: 194, width: 570, height: 94 }, C.blue, 42);
  text(slide, "The gradient supplies a direction. The step length is η||∇f(θₜ)||.", { left: 86, top: 344, width: 590, height: 92 }, { fontSize: 31, bold: true });
  callout(slide, "Animated trajectory", "The orange arrow points uphill. The blue step follows its negative.", { left: 86, top: 484, width: 540, height: 108 }, { fill: C.blueLight, accent: C.blue });
  addGif(slide, await gifFigure("gradient_descent.gif"), { left: 700, top: 148, width: 470, height: 404 }, "Animated gradient descent trajectory on a bowl");
}

// 8. First numeric step
{
  const slide = newSlide("One numerical step", "Foundations", 8, [
    "Pause: Before revealing the final line, compute the update from θ₀ = (1, 2), ∇f = (2, 4), and η = 0.1.",
    "Source: docs/MATH_NOTES.md and docs/QUESTIONS_FOR_CLASS.md.",
  ]);
  text(slide, "Start at θ₀ = (1, 2)", { left: 86, top: 166, width: 470, height: 44 }, { fontSize: 34, bold: true });
  equation(slide, "∇f(1, 2) = (2, 4)", { left: 86, top: 254, width: 500, height: 76 }, C.blue, 34);
  equation(slide, "η = 0.1", { left: 86, top: 368, width: 260, height: 72 }, C.amber, 32);
  equation(slide, "θ₁ = (1, 2) − 0.1(2, 4)\n   = (0.8, 1.6)", { left: 86, top: 478, width: 570, height: 112 }, C.teal, 32);
  callout(slide, "Loss", "f(1, 2) = 5\nf(0.8, 1.6) = 3.2", { left: 810, top: 252, width: 300, height: 170 }, { fill: C.tealLight, accent: C.teal });
  text(slide, "The update moved downhill.", { left: 770, top: 488, width: 370, height: 48 }, { fontSize: 32, bold: true, color: C.teal, alignment: "center" });
}

// 9. Learning rate
{
  const slide = newSlide("The learning rate controls stability", "Stability", 9, [
    "Prediction prompt: Which trajectory should be stable, and what changes when η becomes too large?",
    "Source figure: figures/svg/04_learning_rate_ill_conditioned.svg.",
  ]);
  addSvg(slide, await svgFigure("04_learning_rate_ill_conditioned.svg"), { left: 650, top: 146, width: 520, height: 480 }, "Learning-rate regimes on an ill-conditioned quadratic");
  text(slide, "Same gradient information. Different step sizes.", { left: 86, top: 176, width: 520, height: 52 }, { fontSize: 34, bold: true });
  bullets(slide, [
    "Small η: stable but slow.",
    "Moderate η: steady progress.",
    "Large η: overshoot or divergence.",
  ], { left: 86, top: 294, width: 500, height: 190 }, { fontSize: 29 });
  callout(slide, "Interpretation", "A correct gradient can still produce a bad update when the step is too large.", { left: 86, top: 510, width: 500, height: 104 }, { fill: C.amberLight, accent: C.amber });
}

// 10. Ill-conditioned curvature
{
  const slide = newSlide("Curvature predicts the next step", "Stability", 10, [
    "Prediction checkpoint: Ask students to choose the next point before revealing it.",
    "Starting at (2,1) with η = 0.03, ∇f = (4,50), so θ₁ = (1.88,−0.5).",
    "Then derive |1−50η| < 1 and intersect it with the x-coordinate condition.",
    "Source: docs/MATH_NOTES.md, ill-conditioned quadratic section; exercise adapted from the seminar review.",
  ]);
  equation(slide, "f(x, y) = x² + 25y²", { left: 86, top: 164, width: 480, height: 76 }, C.blue, 34);
  equation(slide, "start: (2,1)\n∇f = (4,50),   η = 0.03", { left: 86, top: 282, width: 510, height: 126 }, C.teal, 29);
  equation(slide, "θ₁ = (2,1) − 0.03(4,50)\n   = (1.88, −0.5)", { left: 86, top: 450, width: 590, height: 126 }, C.ink, 29);
  callout(slide, "Stability", "|1 − 50η| < 1  ⇒  0 < η < 0.04\nThe x-coordinate also imposes |1 − 2η| < 1. The intersection is controlled by the sharper curvature.", { left: 730, top: 238, width: 420, height: 224 }, { fill: C.redLight, accent: C.red });
  callout(slide, "One minimum", "This surface is convex. Zigzags come from unequal curvature, not local-minimum traps.", { left: 730, top: 510, width: 420, height: 106 }, { fill: C.amberLight, accent: C.amber });
}

// 11. Nonconvex
{
  const slide = newSlide("Nonconvex landscapes have multiple basins", "Stability", 11, ["Source figure: figures/svg/06_nonconvex_local_minima.svg; docs/MATH_NOTES.md."]);
  addSvg(slide, await svgFigure("06_nonconvex_local_minima.svg"), { left: 692, top: 142, width: 470, height: 500 }, "Nonconvex landscape with a global and a higher local minimum");
  equation(slide, "f(x, y) = (x² − 1)² + 0.2x + y²", { left: 86, top: 176, width: 570, height: 78 }, C.blue, 27);
  bullets(slide, [
    "Different starting points can enter different basins.",
    "A local minimum can have a higher value than another minimum.",
    "The gradient describes local change, not the full global landscape.",
  ], { left: 86, top: 310, width: 560, height: 220 }, { fontSize: 27 });
}

// 12. Saddle
{
  const slide = newSlide("A zero gradient does not guarantee a minimum", "Stability", 12, [
    "Demonstrate the two paths through the origin before naming the point a saddle.",
    "Along y=0, f(x,0)=x² increases away from 0. Along x=0, f(0,y)=−y² decreases away from 0.",
    "Source figure: figures/svg/07_saddle_point.svg; docs/MATH_NOTES.md.",
  ]);
  addSvg(slide, await svgFigure("07_saddle_point.svg"), { left: 692, top: 142, width: 470, height: 500 }, "Saddle point with zero gradient at the center");
  equation(slide, "f(x, y) = x² − y²", { left: 86, top: 182, width: 430, height: 76 }, C.red, 34);
  equation(slide, "f(x,0)=x²  increases\nf(0,y)=−y²  decreases", { left: 86, top: 302, width: 540, height: 120 }, C.blue, 31);
  callout(slide, "At (0, 0)", "∇f = 0, but nearby points have both larger and smaller values. That is a saddle point.", { left: 86, top: 472, width: 540, height: 136 }, { fill: C.redLight, accent: C.red });
}

// 13. Gradient estimation and notation
{
  const slide = newSlide("The gradient can be estimated", "Optimizers", 13, [
    "Define every symbol before using it in a new update rule.",
    "Explain that vector squaring and division are element-wise in these formulas.",
    "Source: PROJECT_PLAN.md, section 8; docs/OPTIMIZER_CHEATSHEET.md.",
  ]);
  equation(slide, "gₜ = gradient estimate at step t\nĝₜ ≈ ∇L(θₜ)  for a mini-batch", { left: 86, top: 166, width: 600, height: 126 }, C.blue, 29);
  equation(slide, "vₜ = state / accumulated direction\nβ = memory coefficient,   ε = small stabilizer", { left: 86, top: 344, width: 630, height: 126 }, C.teal, 28);
  callout(slide, "Full batch vs mini-batch", "A full batch uses every example. A mini-batch uses an estimate that is cheaper and noisier.", { left: 800, top: 224, width: 340, height: 166 }, { fill: C.amberLight, accent: C.amber });
  text(slide, "g², √v, and g/(√v+ε) operate coordinate by coordinate.", { left: 800, top: 462, width: 340, height: 78 }, { fontSize: 26, bold: true, color: C.ink });
}

// 14. Momentum
{
  const slide = newSlide("Momentum accumulates consistent directions", "Optimizers", 14, [
    "Work the two updates aloud before showing the animation.",
    "Use β=0.5, η=0.1, v₀=(0,0), g₁=(2,4), and g₂=(2,−4). The x-component accumulates while the y-component cancels.",
    "Source: docs/OPTIMIZER_CHEATSHEET.md; animated source: figures/gif/momentum.gif.",
  ]);
  equation(slide, "vₜ = βvₜ₋₁ + gₜ\nθₜ₊₁ = θₜ − ηvₜ", { left: 86, top: 166, width: 510, height: 126 }, C.blue, 31);
  equation(slide, "v₁ = (2,4)\nθ₁ = θ₀ − 0.1(2,4)", { left: 86, top: 338, width: 480, height: 112 }, C.teal, 29);
  equation(slide, "v₂ = 0.5(2,4) + (2,−4)\n   = (3,−2)", { left: 86, top: 492, width: 570, height: 112 }, C.amber, 28);
  addGif(slide, await gifFigure("momentum.gif"), { left: 690, top: 148, width: 480, height: 404 }, "Animated momentum trajectory on an ill-conditioned surface");
}

// 15. Nesterov
{
  const slide = newSlide("Nesterov evaluates a look-ahead gradient", "Optimizers", 15, [
    "Emphasize the state update before interpreting the look-ahead position.",
    "Source: docs/OPTIMIZER_CHEATSHEET.md; animated source: figures/gif/nesterov.gif.",
  ]);
  equation(slide, "look-ahead: θ̃ₜ = θₜ − ηβvₜ₋₁\ngₜ = ∇f(θ̃ₜ)\nvₜ = βvₜ₋₁ + gₜ", { left: 86, top: 160, width: 620, height: 170 }, C.teal, 28);
  equation(slide, "θₜ₊₁ = θₜ − ηvₜ", { left: 86, top: 360, width: 500, height: 76 }, C.ink, 32);
  text(slide, "The gradient is evaluated near the position momentum is about to reach.", { left: 86, top: 464, width: 570, height: 70 }, { fontSize: 29, bold: true });
  callout(slide, "Look ahead", "Preview first, measure the slope there, then update the actual parameters.", { left: 86, top: 548, width: 570, height: 64 }, { fill: C.tealLight, accent: C.teal });
  addGif(slide, await gifFigure("nesterov.gif"), { left: 700, top: 148, width: 470, height: 404 }, "Animated Nesterov trajectory on an ill-conditioned surface");
}

// 16. AdaGrad
{
  const slide = newSlide("AdaGrad scales coordinates by their history", "Optimizers", 16, ["Source: docs/OPTIMIZER_CHEATSHEET.md; animated source: figures/gif/adagrad.gif."]);
  equation(slide, "Gₜ = Gₜ₋₁ + gₜ²\nθₜ₊₁ = θₜ − η gₜ/(√Gₜ + ε)", { left: 86, top: 184, width: 620, height: 136 }, C.blue, 34);
  text(slide, "Coordinates with large accumulated gradients receive smaller effective steps.", { left: 86, top: 388, width: 610, height: 92 }, { fontSize: 31, bold: true });
  callout(slide, "Tradeoff", "The accumulator only grows, so the effective step can shrink indefinitely.", { left: 86, top: 520, width: 610, height: 94 }, { fill: C.blueLight, accent: C.blue });
  addGif(slide, await gifFigure("adagrad.gif"), { left: 748, top: 150, width: 420, height: 370 }, "Animated AdaGrad trajectory on an ill-conditioned surface");
}

// 17. RMSProp
{
  const slide = newSlide("RMSProp keeps recent gradient history", "Optimizers", 17, ["Source: docs/OPTIMIZER_CHEATSHEET.md; animated source: figures/gif/rmsprop.gif."]);
  equation(slide, "sₜ = ρsₜ₋₁ + (1−ρ)gₜ²\nθₜ₊₁ = θₜ − η gₜ/(√sₜ + ε)", { left: 86, top: 184, width: 650, height: 136 }, C.teal, 34);
  text(slide, "An exponential moving average emphasizes recent gradient magnitudes.", { left: 86, top: 388, width: 630, height: 92 }, { fontSize: 31, bold: true });
  callout(slide, "Contrast", "RMSProp replaces AdaGrad's cumulative sum with a moving average.", { left: 86, top: 520, width: 610, height: 94 }, { fill: C.tealLight, accent: C.teal });
  addGif(slide, await gifFigure("rmsprop.gif"), { left: 748, top: 150, width: 420, height: 370 }, "Animated RMSProp trajectory on an ill-conditioned surface");
}

// 18. Adam
{
  const slide = newSlide("Adam combines memory with scale", "Optimizers", 18, [
    "Connect Adam to the previous slides: m remembers direction, v remembers squared magnitude.",
    "Derive the first-step bias correction: m₁=(1−β₁)g₁, so m̂₁=m₁/(1−β₁)=g₁.",
    "Source: docs/OPTIMIZER_CHEATSHEET.md; Kingma and Ba, arXiv:1412.6980; animated source: figures/gif/adam.gif.",
  ]);
  equation(slide, "mₜ = β₁mₜ₋₁ + (1−β₁)gₜ\nvₜ = β₂vₜ₋₁ + (1−β₂)gₜ²", { left: 86, top: 160, width: 620, height: 132 }, C.blue, 29);
  equation(slide, "m̂ₜ = mₜ/(1−β₁ᵗ)\nv̂ₜ = vₜ/(1−β₂ᵗ)", { left: 86, top: 330, width: 500, height: 112 }, C.teal, 30);
  equation(slide, "at t=1:  m₁=(1−β₁)g₁\n          m̂₁=g₁", { left: 86, top: 480, width: 500, height: 112 }, C.amber, 28);
  addGif(slide, await gifFigure("adam.gif"), { left: 748, top: 156, width: 420, height: 370 }, "Animated Adam trajectory on an ill-conditioned surface");
}

// 19. AdamW
{
  const slide = newSlide("AdamW decouples weight decay", "Optimizers", 19, [
    "Emphasize that the decay term acts on the original parameters, not on the already-updated parameters.",
    "Use the combined rule shown on the slide. This is the conventional decoupled form.",
    "Source: docs/OPTIMIZER_CHEATSHEET.md; Loshchilov and Hutter, arXiv:1711.05101; animated source: figures/gif/adamw.gif.",
  ]);
  equation(slide, "uₜ = Adam update direction", { left: 86, top: 180, width: 560, height: 76 }, C.blue, 31);
  equation(slide, "θₜ₊₁ = (1−ηλ)θₜ − ηuₜ", { left: 86, top: 312, width: 620, height: 82 }, C.teal, 34);
  bullets(slide, [
    "The decay term shrinks θₜ directly.",
    "The adaptive direction uₜ remains separate.",
    "λ controls the strength of the shrinkage.",
  ], { left: 86, top: 456, width: 590, height: 150 }, { fontSize: 27, spaceAfterPoints: 10 });
  addGif(slide, await gifFigure("adamw.gif"), { left: 760, top: 156, width: 400, height: 370 }, "Animated AdamW trajectory on an ill-conditioned surface");
}

// 20. Family map
{
  const slide = newSlide("Optimizer choices change different parts of the loop", "Optimizers", 20, [
    "Do not present this as a single family tree. Separate sampling, stateful update rules, regularization, and alternative methods.",
    "SAM wraps an update with a sharpness-aware perturbation. L-BFGS uses curvature approximations. Lion uses sign-based momentum.",
    "Source: docs/OPTIMIZER_CHEATSHEET.md and docs/SOURCES.md.",
  ]);
  const groups = [
    ["GRADIENT SOURCE", "full batch\nmini-batch", C.blueLight, C.blue, 86, 210, 260],
    ["STATEFUL UPDATE", "SGD\nMomentum\nNesterov\nAdaGrad / RMSProp\nAdam / AdamW", C.tealLight, C.teal, 450, 174, 286],
    ["REGULARIZATION", "weight decay\n(dropout is model-side)", C.amberLight, C.amber, 846, 210, 300],
    ["ALTERNATIVE METHODS", "SAM · L-BFGS · Lion", C.redLight, C.red, 450, 480, 286],
  ];
  for (const [label, body, fill, accent, left, top, width] of groups) {
    const boxHeight = label === "STATEFUL UPDATE" ? 260 : 150;
    const box = shape(slide, "roundRect", { left, top, width, height: boxHeight }, { fill, line: { style: "solid", fill: accent, width: 1.5 }, borderRadius: 14 });
    box.text = `${label}\n\n${body}`;
    box.text.style = { typeface: font, fontSize: label === "STATEFUL UPDATE" ? 24 : 25, bold: true, color: accent, alignment: "center", verticalAlignment: "middle", autoFit: "shrinkText", insets: { top: 14, right: 16, bottom: 14, left: 16 } };
  }
  text(slide, "These choices can be combined, but they answer different questions.", { left: 240, top: 622, width: 800, height: 28 }, { fontSize: 22, color: C.muted, alignment: "center" });
}

// 21. Comparison
{
  const slide = newSlide("Compare mechanisms before comparing every optimizer", "Optimizers", 21, [
    "Use animations in pairs: gradient descent vs momentum, AdaGrad vs RMSProp, then Adam vs AdamW.",
    "Treat the combined plot as a summary, not a ranking. State objective, starting point, steps, learning rates, and decay coefficient.",
    "The deterministic curves are gradient descent unless stochastic gradient noise is explicitly introduced.",
    "Source figure: figures/svg/05_optimizer_trajectories_ill_conditioned.svg; PROJECT_PLAN.md, sections 5 and 9.",
  ]);
  addSvg(slide, await svgFigure("05_optimizer_trajectories_ill_conditioned.svg"), { left: 72, top: 146, width: 790, height: 450 }, "Optimizer trajectories on the same ill-conditioned objective");
  bullets(slide, [
    "Same objective and starting point",
    "Method-specific learning rates",
    "Fixed step budget",
    "Mechanism evidence, not a universal ranking",
  ], { left: 900, top: 198, width: 270, height: 220 }, { fontSize: 23, spaceAfterPoints: 11 });
  callout(slide, "Teaching sequence", "Show one pair, pause, predict the next step, then reveal the trajectory.", { left: 900, top: 470, width: 270, height: 116 }, { fill: C.tealLight, accent: C.teal });
}

// 22. CNN architecture
{
  const slide = newSlide("The same update loop appears in a NumPy CNN", "Application", 22, [
    "Connect the architecture diagram to the loop: forward pass, loss, backpropagation, and optimizer update.",
    "Source figure: figures/svg/09_numpy_cnn_architecture.svg; docs/CNN_NOTES_FROM_UPLOADED_CODE.md.",
  ]);
  addSvg(slide, await svgFigure("09_numpy_cnn_architecture.svg"), { left: 56, top: 144, width: 1168, height: 484 }, "NumPy CNN architecture from image input to logits");
  text(slide, "Input → feature learning → classification", { left: 92, top: 630, width: 700, height: 28 }, { fontSize: 22, color: C.muted });
}

// 23. CNN propagation
{
  const slide = newSlide("Forward pass, loss, backpropagation, update", "Application", 23, [
    "Point out the boundary between gradient computation and parameter update.",
    "Source: docs/TALK_SCRIPT.md and animated source: figures/gif/cnn_forward_backward.gif.",
  ]);
  addGif(slide, await gifFigure("cnn_forward_backward.gif"), { left: 92, top: 146, width: 1096, height: 484 }, "Animated forward pass, loss, backpropagation, and parameter update in the NumPy CNN");
  text(slide, "Backpropagation computes gradients. The optimizer transforms them into parameter updates.", { left: 92, top: 630, width: 1000, height: 28 }, { fontSize: 22, color: C.muted });
}

// 24. Evaluation
{
  const slide = newSlide("A fair comparison needs controlled conditions", "Application", 24, [
    "Comparison prompt: If an optimizer wins on one seed or one learning rate, what evidence is still missing?",
    "Source: docs/EXPERIMENT_PLAN.md and PROJECT_PLAN.md, sections 9 and 10.",
  ]);
  bullets(slide, [
    "Hold the architecture, data split, initialization, batch order, and step budget constant.",
    "Use validation data during development and reserve the test set for final evaluation.",
    "Compare multiple random seeds and report variability with the mean.",
    "Report accuracy, loss, runtime, and the comparison conditions together.",
  ], { left: 92, top: 170, width: 690, height: 350 }, { fontSize: 28 });
  callout(slide, "Interpretation", "A faster path on one two-dimensional surface is evidence about a mechanism, not a universal performance claim.", { left: 830, top: 238, width: 300, height: 210 }, { fill: C.amberLight, accent: C.amber });
}

// 25. CNN gradient flow
{
  const slide = newSlide("Softmax cross-entropy starts backpropagation", "Application", 25, [
    "Show one concrete derivative connecting the CNN's output to its parameter gradients.",
    "For the true class y, the logit derivative is p_k − 1{k=y}. The final dense layer receives this vector through the chain rule.",
    "This slide explains the mechanism. It does not claim validated MNIST benchmark results.",
    "Source: docs/CNN_NOTES_FROM_UPLOADED_CODE.md and standard softmax cross-entropy derivative.",
  ]);
  equation(slide, "p_k = e^{z_k}/Σⱼe^{zⱼ}\nL = −log p_y", { left: 86, top: 160, width: 540, height: 126 }, C.blue, 31);
  equation(slide, "∂L/∂z_k = p_k − 1{k=y}", { left: 86, top: 338, width: 610, height: 82 }, C.teal, 34);
  equation(slide, "∂L/∂W = (∂L/∂z)hᵀ\nW ← W − η ∂L/∂W", { left: 86, top: 472, width: 610, height: 126 }, C.amber, 29);
  callout(slide, "The connection", "The loss creates a gradient at the logits. Backpropagation carries it into the weights. The optimizer changes those weights.", { left: 790, top: 246, width: 340, height: 210 }, { fill: C.tealLight, accent: C.teal });
}

// 26. Whole loop
{
  const slide = newSlide("Training is a cycle", "Application", 26, [
    "Trace the arrows in order and emphasize the return to updated parameters.",
    "The cycle is the central mechanism: compute a loss, differentiate it, update parameters, and repeat.",
    "Source: PROJECT_PLAN.md, sections 2 and 10.",
  ]);
  const boxes = [
    ["FORWARD PASS", "predictions\nfrom θₜ", 82, 214, 258, 142, C.blueLight, C.blue],
    ["LOSS", "measure\nerror", 470, 214, 258, 142, C.tealLight, C.teal],
    ["BACKPROPAGATION", "compute\n∇θL", 858, 214, 300, 142, C.amberLight, C.amber],
    ["OPTIMIZER", "use η and\nstate", 858, 466, 300, 126, C.redLight, C.red],
    ["UPDATED PARAMETERS", "θₜ₊₁", 470, 466, 258, 126, C.blueLight, C.blue],
  ];
  for (const [label, body, left, top, width, height, fill, accent] of boxes) {
    const box = shape(slide, "roundRect", { left, top, width, height }, { fill, line: { style: "solid", fill: accent, width: 1.5 }, borderRadius: 14 });
    box.text = `${label}\n\n${body}`;
    box.text.style = { typeface: font, fontSize: 25, bold: true, color: accent, alignment: "center", verticalAlignment: "middle", autoFit: "shrinkText", insets: { top: 12, right: 14, bottom: 12, left: 14 } };
  }
  shape(slide, "line", { left: 340, top: 285, width: 130, height: 0 }, { line: { style: "solid", fill: C.muted, width: 3, endArrowType: "triangle" } });
  shape(slide, "line", { left: 728, top: 285, width: 130, height: 0 }, { line: { style: "solid", fill: C.muted, width: 3, endArrowType: "triangle" } });
  shape(slide, "line", { left: 1008, top: 356, width: 0, height: 110 }, { line: { style: "solid", fill: C.muted, width: 3, endArrowType: "triangle" } });
  shape(slide, "line", { left: 728, top: 529, width: 130, height: 0 }, { line: { style: "solid", fill: C.muted, width: 3, endArrowType: "triangle" } });
  shape(slide, "line", { left: 340, top: 529, width: 130, height: 0 }, { line: { style: "solid", fill: C.muted, width: 3, beginArrowType: "triangle" } });
  shape(slide, "line", { left: 211, top: 356, width: 0, height: 110 }, { line: { style: "solid", fill: C.blue, width: 3, beginArrowType: "triangle" } });
  text(slide, "repeat", { left: 142, top: 386, width: 70, height: 28 }, { fontSize: 20, bold: true, color: C.blue, alignment: "center" });
}

// 27. Summary
{
  const slide = newSlide("TL;DR", "Summary", 27, [
    "Closing checkpoint: ask students to explain where the gradient comes from and what the optimizer changes.",
    "Source: PROJECT_PLAN.md, sections 1, 3, 8, and 10; discussion prompts: docs/QUESTIONS_FOR_CLASS.md.",
  ]);
  bullets(slide, [
    "A loss turns model behavior into a number we can minimize.",
    "The gradient combines coordinate-wise slopes and points toward local increase.",
    "The learning rate and curvature determine whether steps converge, oscillate, or diverge.",
    "Momentum remembers direction. Adaptive methods rescale coordinates. Adam combines both ideas.",
    "Backpropagation computes gradients. The optimizer updates parameters.",
    "A fair CNN comparison needs controlled conditions and validated measurements.",
  ], { left: 94, top: 164, width: 1050, height: 386 }, { fontSize: 28, spaceAfterPoints: 14 });
  callout(slide, "Core update", "θₜ₊₁ = θₜ − η∇f(θₜ)", { left: 220, top: 570, width: 840, height: 78 }, { fill: C.blueLight, accent: C.blue });
}

// 28. References
{
  const slide = newSlide("References", "Summary", 28, [
    "These sources support the optimizer terminology, primary methods, and MNIST dataset reference.",
    "Source: docs/SOURCES.md.",
  ]);
  text(slide, "Primary methods", { left: 94, top: 154, width: 420, height: 34 }, { fontSize: 28, bold: true, color: C.blue });
  bullets(slide, [
    "Kingma and Ba, “Adam: A Method for Stochastic Optimization” — arXiv:1412.6980",
    "Duchi et al., “Adaptive Subgradient Methods for Online Learning and Stochastic Optimization” — JMLR 2011",
    "Tieleman and Hinton, “RMSProp” — lecture notes, 2012",
    "Loshchilov and Hutter, “Decoupled Weight Decay Regularization” — arXiv:1711.05101",
    "Chen et al., “Symbolic Discovery of Optimization Algorithms” — arXiv:2302.06675",
  ], { left: 96, top: 204, width: 1060, height: 248 }, { fontSize: 20, spaceAfterPoints: 8 });
  text(slide, "Additional sources", { left: 94, top: 478, width: 420, height: 34 }, { fontSize: 28, bold: true, color: C.teal });
  text(slide, "Foret et al., “Sharpness-Aware Minimization” — arXiv:2010.01412\nPyTorch optimizer documentation — docs.pytorch.org/docs/stable/optim.html\nMNIST archive — storage.googleapis.com/tensorflow/tf-keras-datasets/mnist.npz\nSource code — github.com/1zuki/numpy-cnn", { left: 96, top: 526, width: 1040, height: 112 }, { fontSize: 20, color: C.ink });
}

const candidate = await (await PresentationFile.exportPptx(presentation)).save(candidatePath);
const requirements = {
  explicitTotalSlideCount: TOTAL_SLIDES,
  requiredNativeTableOwnerSlides: [],
  requiredNativeChartOwnerSlides: [],
};
const finalizerDir = path.join(buildDir, "finalizer");
await fs.mkdir(finalizerDir, { recursive: true });
const expectedSlideSizeEmu = "12192000,6858000";
const result = await finalizePresentation({
  ...requirements,
  workspaceDir,
  candidatePath,
  finalPath: outputPath,
  pythonExecutable: "/home/izu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3.12",
  integrityValidatorPath: path.join(skillDir, "container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(skillDir, "container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs: ["--expected-slide-size-emu", expectedSlideSizeEmu, "--validate-bullet-geometry", "--validate-heading-fit"],
  fontPolicy: { basis: "design", families: ["Arial"] },
  verifyArtifactToolImport: true,
  receiptPath: path.join(finalizerDir, "gradient-descent-seminar-reader-facing-v4-final.validation.json"),
});
console.log(JSON.stringify({ outputPath, candidatePath, slideCount: presentation.slides.items?.length ?? TOTAL_SLIDES, result }, null, 2));
