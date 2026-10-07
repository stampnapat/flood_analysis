// Build the presentation deck: node slides/build_deck.js
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const ROOT = path.resolve(__dirname, "..");
const FIG = (n) => path.join(ROOT, "figures", n + ".png");
const R = JSON.parse(fs.readFileSync(path.join(ROOT, "data/processed/results.json"), "utf8"));
const LOG = JSON.parse(fs.readFileSync(path.join(ROOT, "data/processed/data_log.json"), "utf8"));
const SIZES = JSON.parse(fs.readFileSync(path.join(ROOT, "figures/sizes.json"), "utf8"));
const OUT = path.join(ROOT, "slides", "Flood_Clustering_Thailand.pptx");

const THEME = {
  name: "Flood Analytics",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B2631", lt1: "FFFFFF", dk2: "0B3C5D", lt2: "EAF2F8",
    accent1: "1D7874", accent2: "D7191C", accent3: "FDAE61",
    accent4: "2C7BB6", accent5: "7B3294", accent6: "5D6D7E",
    hlink: "1D7874", folHlink: "7B3294",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Spatial Clustering of Flood Impact in Thailand 2019-2024";
const C = pres.SchemeColor;

// ---------------- layouts ----------------
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: HEX.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.2, w: 8.8, h: 1.6, fontSize: 36, bold: true, color: C.background1, valign: "bottom", fontFace: THEME.headFontFace }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 3.0, w: 8.8, h: 1.9, fontSize: 18, color: C.background2, valign: "top" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.lt1 },
  margin: [0.5, 0.5, 0.5, 0.5],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.25, w: 9.0, h: 0.75, fontSize: 28, bold: true, color: C.text2, valign: "middle", fontFace: THEME.headFontFace }, text: "" } },
    { text: { text: "Flood impact clustering · Thailand 2019–2024", options: { x: 0.5, y: 5.25, w: 6, h: 0.25, fontSize: 9, color: C.accent6 } } },
  ],
  slideNumber: { x: 9.0, y: 5.25, w: 0.5, h: 0.25, fontSize: 9, color: C.accent6, align: "right" },
});
pres.defineSlideMaster({
  title: "SECTION",
  background: { color: HEX.accent1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.9, w: 8.8, h: 1.0, fontSize: 36, bold: true, color: C.background1, fontFace: THEME.headFontFace }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 2.95, w: 8.8, h: 0.8, fontSize: 18, color: C.background1 }, text: "" } },
  ],
});

// ---------------- helpers ----------------
function img(slide, name, x, y, w, h, opts = {}) {
  const [pw, ph] = SIZES[name];
  const r = pw / ph;
  let iw = w, ih = w / r;
  if (ih > h) { ih = h; iw = h * r; }
  slide.addImage({ path: FIG(name), x: x + (w - iw) / 2, y: y + (h - ih) / 2, w: iw, h: ih, altText: opts.alt || name, objectName: "fig_" + name });
}
function text(slide, t, x, y, w, h, o = {}) {
  slide.addText(t, Object.assign({ x, y, w, h, isTextBox: true, fontSize: 14, color: C.text1, valign: "top", margin: 0.05 }, o));
}
function bullets(slide, items, x, y, w, h, o = {}) {
  const arr = items.map((it, i) => {
    const runOpts = { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 6 };
    if (Array.isArray(it)) {
      return it.map((r, j) => ({ text: r.text, options: Object.assign({}, r.options || {}, j === 0 ? { bullet: true } : {}, j === it.length - 1 ? { breakLine: i < items.length - 1, paraSpaceAfter: 6 } : {}) }));
    }
    return [{ text: it, options: runOpts }];
  }).flat();
  slide.addText(arr, Object.assign({ x, y, w, h, isTextBox: true, fontSize: 14, color: C.text1, valign: "top", margin: 0.05 }, o));
}
function card(slide, x, y, w, h, fill = HEX.lt2) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill } });
}
function stat(slide, big, label, x, y, w, color = C.accent1) {
  text(slide, big, x, y, w, 0.65, { fontSize: 34, bold: true, color, fontFace: THEME.headFontFace, align: "left", margin: 0 });
  text(slide, label, x, y + 0.65, w, 0.55, { fontSize: 11, color: C.accent6, margin: 0 });
}
const B = (t) => ({ text: t, options: { bold: true } });
const N = (t) => ({ text: t });
let sec = null;
function add(layout, section) {
  if (section && section !== sec) { pres.addSection({ title: section }); sec = section; }
  return pres.addSlide({ masterName: layout, sectionTitle: sec });
}
const f2 = (v) => Number(v).toFixed(2);
const f3 = (v) => Number(v).toFixed(3);
const ev = Object.fromEntries(R.internal_eval.map((r) => [r.index, r]));
const ext = R.external_eval;
const prof = R.kmeans_profile;
const nm = R.kmeans_names;
const kmK = R.kmeans_k_used;

// ================= 1. Title =================
let s = add("TITLE_DARK", "Introduction");
s.addText("Where Floods Hurt vs. Where Water Recurs: Spatial Clustering of Flood Impact in Thailand (2019–2024)", { placeholder: "title" });
s.addText([
  { text: "Group Assignment: Analytics on Complex Data", options: { breakLine: true } },
  { text: "Presented by: [Member 1] · [Member 2] · [Member 3]", options: { breakLine: true } },
  { text: "Faculty of Engineering, Kasetsart University", options: {} },
], { placeholder: "body" });
s.addNotes("สวัสดีครับ หัวข้อของกลุ่มเราคือการจัดกลุ่มจังหวัดตามรูปแบบผลกระทบจากน้ำท่วมในประเทศไทยช่วงปี 2019–2024 และหา hotspot เชิงพื้นที่ โดยใช้ข้อมูลจาก 4 หน่วยงาน คำถามหลักคือ จังหวัดที่ได้รับความเสียหายจากน้ำท่วมมาก เป็นจังหวัดเดียวกับที่น้ำท่วมซ้ำซากหรือไม่");

// ================= 2. Complex data & application =================
s = add("CONTENT", "Complex Data");
s.addText("Complex data: 4 agencies, 6 schemas, 1 map", { placeholder: "title" });
const srcs = [
  ["DDPM (ปภ.)", "Flood disaster statistics 2019–2024: affected people, houses, farmland, roads … (xlsx)"],
  ["HII / ThaiWater", "Monthly areal rainfall per province + 30-yr normal (csv)"],
  ["DOPA (กรมการปกครอง)", "Registered population per province per year (HTML tables saved as .xls)"],
  ["Recurrent-flood map (LDD)", "Land flooded ≥4 times in 10 yrs, per sub-district – used as validation reference"],
];
srcs.forEach(([h, d], i) => {
  const x = 0.5 + (i % 2) * 4.6, y = 1.15 + Math.floor(i / 2) * 1.35;
  card(s, x, y, 4.4, 1.2);
  text(s, h, x + 0.15, y + 0.1, 4.1, 0.35, { fontSize: 15, bold: true, color: C.text2 });
  text(s, d, x + 0.15, y + 0.45, 4.1, 0.7, { fontSize: 12 });
});
text(s, [B("Why complex? "), N("multi-source & multi-format · spatial (polygons + neighbours) · temporal (6 yrs) · heterogeneous granularity (province / district-event / village) · ~49 heavily skewed, zero-inflated attributes")],
  0.5, 3.95, 9.0, 0.6, { fontSize: 13 });
text(s, [B("Application: "), N("flood-risk zoning for DDPM budget & preparedness – which provinces need which type of intervention?")],
  0.5, 4.55, 9.0, 0.5, { fontSize: 13, color: C.accent1 });
s.addNotes("ข้อมูลของเราเป็น complex data เพราะมาจาก 4 หน่วยงาน คนละรูปแบบไฟล์ คนละความละเอียด: ปภ. ให้สถิติความเสียหาย, สถาบันสารสนเทศทรัพยากรน้ำ (ThaiWater) ให้ปริมาณฝนรายเดือนรายจังหวัด, กรมการปกครองให้จำนวนประชากร ซึ่งไฟล์ .xls จริงๆ เป็นตาราง HTML, และแผนที่น้ำท่วมซ้ำซากที่เราใช้เป็นตัวอ้างอิงตอนประเมินผล เป็นข้อมูลเชิงพื้นที่และเวลา และมี attribute ที่เบ้มาก การประยุกต์ใช้คือช่วยจัดลำดับพื้นที่เสี่ยงให้หน่วยงานวางแผนงบประมาณและการเตรียมรับมือ");

// ================= 3. Data complexity in numbers =================
s = add("CONTENT", "Complex Data");
s.addText("The raw data is messy by design", { placeholder: "title" });
stat(s, "44,044", "raw DDPM rows (province, district-event and village level)", 0.5, 1.2, 2.1);
stat(s, "44–66", "columns per year – a different schema every year", 2.75, 1.2, 2.1);
stat(s, "144", "embedded subtotal / grand-total rows to remove", 5.0, 1.2, 2.1, C.accent2);
stat(s, "462", "clean province × year records (77 × 6)", 7.25, 1.2, 2.1);
const rows3 = [["Year", "Granularity", "Header", "Damage scheme"]].concat(
  [["2019", "province", "Eng + Thai rows", "49 detailed + THB value"],
   ["2020–2022", "district-event", "1–2 merged rows", "56–60 detailed items"],
   ["2023", "province", "3 merged rows", "13 coarse categories"],
   ["2024", "village (34,881 rows)", "English", "13 coarse categories"]]);
s.addTable(rows3.map((r, i) => r.map((c) => ({ text: c, options: i === 0 ? { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 } } : {} }))),
  { x: 0.5, y: 2.65, w: 9.0, colW: [1.3, 2.3, 2.2, 3.2], fontSize: 12, fontFace: THEME.bodyFontFace, color: HEX.dk1, border: { type: "solid", color: "D5DBDB", pt: 0.75 }, rowH: 0.36, objectName: "schema_table" });
text(s, "Other traps: province names written 3 ways (“สมุทรปราการ”, “จังหวัดสมุทรปราการ”, “จ.สมุทรปราการ”), merged cells, “-” and blanks for zero, 2026 rainfall rows with swapped columns.",
  0.5, 4.55, 9.0, 0.6, { fontSize: 12, color: C.accent6 });
s.addNotes("ตัวเลขนี้แสดงความซับซ้อนของข้อมูล: แถวดิบของ ปภ. รวม 44,044 แถว แต่ละปีมีคอลัมน์ 44 ถึง 66 คอลัมน์และ schema ไม่เหมือนกันเลย ปี 2019 เป็นระดับจังหวัด ปี 2020–2022 เป็นระดับอำเภอต่อเหตุการณ์ ปี 2024 เป็นระดับหมู่บ้านเกือบ 35,000 แถว และปี 2023 เปลี่ยนเป็นหมวดความเสียหายแบบหยาบ 13 หมวด นอกจากนี้ยังมีแถวยอดรวม 144 แถวแทรกอยู่ในข้อมูล ถ้าไม่ลบจะนับซ้ำ และชื่อจังหวัดเขียนได้หลายแบบ");

// ================= 4. EDA – time & space =================
s = add("CONTENT", "Exploration");
s.addText("Exploration: impact varies by year and by region", { placeholder: "title" });
const yrs = ["2019", "2020", "2021", "2022", "2023", "2024"];
s.addChart(pres.charts.BAR, [{ name: "Affected people (million)", labels: yrs, values: [1.95, 3.58, 2.52, 8.17, 1.13, 3.75] }], {
  x: 0.4, y: 1.05, w: 3.6, h: 3.4, barDir: "col", chartColors: [HEX.accent4],
  showTitle: true, title: "Affected people per year (million)", titleFontSize: 12, titleColor: HEX.dk2, titleFontFace: "+mn-lt",
  showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 10, dataLabelColor: HEX.dk1, dataLabelFormatCode: "0.0", dataLabelFontFace: "+mn-lt",
  catAxisLabelColor: HEX.accent6, valAxisLabelColor: HEX.accent6, catAxisLabelFontSize: 10, valAxisLabelFontSize: 9, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  valGridLine: { color: "E5E8E8", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false, valAxisHidden: true,
});
img(s, "eda_maps", 4.1, 1.0, 5.5, 3.5);
text(s, [B("2022 "), N("(storm Noru) affected 8.2 M people – 7× the 2023 total. Impact per capita is highest in the "), B("north and deep south"), N(", not in the rain-richest east/south-west.")],
  0.5, 4.55, 9.0, 0.6, { fontSize: 13 });
s.addNotes("จากการสำรวจข้อมูล ผลกระทบต่างกันมากในแต่ละปี ปี 2022 จากพายุโนรูมีผู้ได้รับผลกระทบ 8.2 ล้านคน มากกว่าปี 2023 ถึง 7 เท่า แผนที่ด้านขวาแสดงว่าผู้ได้รับผลกระทบต่อประชากรพันคนสูงในภาคเหนือและใต้สุด ขณะที่ฝนมากที่สุดอยู่ภาคตะวันออกและฝั่งอันดามัน แปลว่าฝนอย่างเดียวอธิบายผลกระทบไม่ได้ เกือบทุกจังหวัดมีการประกาศภัยทุกปี");

// ================= 5. EDA – skew =================
s = add("CONTENT", "Exploration");
s.addText("Exploration: skewed, zero-inflated data", { placeholder: "title" });
stat(s, String(R.eda_raw_attributes), "damage attributes in the 2019 file", 0.5, 1.15, 2.0);
stat(s, f2(R.eda_median_skew).replace(/0$/, ""), "median skewness (>1 is already 'highly skewed')", 0.5, 2.3, 2.0, C.accent2);
stat(s, Math.round(R.eda_median_zero_share * 100) + "%", "median share of zeros per attribute", 0.5, 3.45, 2.0, C.accent2);
img(s, "eda_skew", 2.7, 1.1, 6.9, 2.7);
text(s, "→ Raw values cannot go into a distance-based model: a few provinces would dominate every distance. We need population/area normalisation + log transform. Most impact features are also positively correlated (Spearman up to 0.65) → PCA.",
  2.8, 4.0, 6.8, 1.0, { fontSize: 13 });
s.addNotes("attribute ความเสียหาย 48 ตัวในไฟล์ปี 2019 มีค่าความเบ้มัธยฐาน 5.3 และมีค่าศูนย์เฉลี่ยถึง 82% ของจังหวัด ถ้าเอาค่าดิบไปทำ clustering ที่ใช้ระยะทาง จังหวัดไม่กี่จังหวัดจะครอบงำผลทั้งหมด จึงต้องทำ normalisation ด้วยประชากร/พื้นที่ และ log transform และ feature หลายตัวมีความสัมพันธ์กัน จึงใช้ PCA ต่อ");

// ================= 6. Preprocessing – integration pipeline =================
s = add("CONTENT", "Pre-processing");
s.addText("Pre-processing 1: integrating four sources", { placeholder: "title" });
const steps = [
  ["Parse", "6 DDPM layouts, HTML-in-.xls, merged-cell CSV"],
  ["Clean", "drop 144 subtotal rows; '-', blanks → 0; strip zero-width chars"],
  ["Harmonise", "44–66 columns → 13 common attributes by keyword mapping"],
  ["Match", "normalise province names → 77 official codes (0 unmatched)"],
  ["Join", "province × year grid with rain + population; absent year = no flood"],
];
steps.forEach(([h, d], i) => {
  const x = 0.5 + i * 1.84;
  s.addShape(pres.shapes.OVAL, { x: x + 0.6, y: 1.2, w: 0.5, h: 0.5, fill: { color: HEX.accent1 }, line: { color: HEX.accent1 } });
  text(s, String(i + 1), x + 0.6, 1.2, 0.5, 0.5, { fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
  if (i < steps.length - 1) s.addShape(pres.shapes.LINE, { x: x + 1.15, y: 1.45, w: 1.24, h: 0, line: { color: "B3B6B7", width: 1.25, endArrowType: "triangle" } });
  text(s, h, x, 1.8, 1.7, 0.35, { fontSize: 15, bold: true, color: C.text2, align: "center" });
  text(s, d, x, 2.15, 1.7, 1.2, { fontSize: 11.5, align: "center" });
});
card(s, 0.5, 3.45, 9.0, 1.6);
text(s, "Harmonised attributes (sum of the matching raw columns per year)", 0.7, 3.55, 8.6, 0.35, { fontSize: 13, bold: true, color: C.text2 });
text(s, "affected people · affected households · evacuated · deaths · injured/missing · housing · farmland (rai) · livestock · fishery · roads & bridges · public sites · districts · villages",
  0.7, 3.95, 8.6, 0.55, { fontSize: 12 });
text(s, "Check: yearly totals reproduce the official 'รวมทั้งหมด' rows exactly (e.g. 2020: 3,576,314 people) before imputation.",
  0.7, 4.5, 8.6, 0.45, { fontSize: 11.5, color: C.accent6, italic: true });
s.addNotes("ขั้นตอนรวมข้อมูลมี 5 ขั้น: อ่านไฟล์ 6 รูปแบบ, ทำความสะอาดโดยลบแถวยอดรวม 144 แถว และแปลง '-' กับช่องว่างเป็นศูนย์, harmonise คอลัมน์ 44–66 คอลัมน์ให้เหลือ attribute ร่วมกัน, จับคู่ชื่อจังหวัดเป็นรหัสจังหวัด 77 รหัสโดยไม่มีชื่อที่จับคู่ไม่ได้, แล้ว join กับฝนและประชากรด้วยจังหวัดและปี เราตรวจว่ายอดรวมรายปีตรงกับแถวยอดรวมทางการของ ปภ. ก่อนเติมค่าที่หายไป");

// ================= 7. Preprocessing – missing, normalisation, PCA =================
s = add("CONTENT", "Pre-processing");
s.addText("Pre-processing 2: missing values, scaling, PCA", { placeholder: "title" });
bullets(s, [
  [B("Zeros that mean missing: "), N(`households > 0 but people = 0 (e.g. Roi Et 2019: 47,822 households, 0 people) → people = households × ${f2(LOG.median_people_per_household)} (median ratio); ${LOG.imputed_people_from_households} cells`)],
  [B("Unreliable features dropped: "), N("evacuees (2023 schema records only 20 nationwide), livestock (one-off poultry spikes)")],
  [B("Normalisation: "), N("per 1,000 / 100,000 residents or per km² – removes the big-province effect")],
  [B("log1p + z-score: "), N(`median skewness ${f2(R.skew_raw_median)} → ${f2(R.skew_log_median)}`)],
  [B("PCA: "), N(`9 features → ${R.pca_n_components} PCs keep ${Math.round(R.pca_cum_var * 100)}% of variance; PC1 (${Math.round(R.pca_pc1_var * 100)}%) loads on every impact feature → used as the Flood Severity Index`)],
], 0.5, 1.1, 4.4, 4.0, { fontSize: 12.5 });
img(s, "pca", 5.0, 1.15, 4.6, 3.6);
s.addNotes("ต่อมาคือการจัดการค่าที่หายไป ข้อมูลมีศูนย์ที่จริงๆ คือค่าหาย เช่น ร้อยเอ็ดปี 2019 มีครัวเรือน 47,822 แต่จำนวนคนเป็นศูนย์ เราเติมด้วยจำนวนครัวเรือนคูณค่ามัธยฐานคนต่อครัวเรือน 2.74 นอกจากนี้ตัด feature ที่ไม่น่าเชื่อถือ เช่น ผู้อพยพ ที่ schema ปี 2023 บันทึกได้แค่ 20 คนทั้งประเทศ แล้ว normalise ต่อประชากรหรือพื้นที่ ทำ log1p และ z-score ทำให้ความเบ้ลดเหลือใกล้ศูนย์ สุดท้ายใช้ PCA เหลือ 5 องค์ประกอบ ครอบคลุม 85% ของความแปรปรวน และ PC1 ใช้เป็นดัชนีความรุนแรง");

// ================= 8. Technique – clustering =================
s = add("CONTENT", "Analytics Technique");
s.addText("Technique 1: K-Means vs HDBSCAN", { placeholder: "title" });
[[0.5, "K-Means (covered in class)", HEX.lt2, [
  "Partitions provinces into k spherical groups around centroids",
  "k chosen by elbow + silhouette + stability over 20 seeds",
  "Every province must join a cluster – outliers distort centroids",
]], [5.1, "HDBSCAN (not covered in class)", "FDEBD0", [
  "Hierarchical density-based clustering: builds a minimum spanning tree on mutual-reachability distance, keeps the most stable density clusters",
  "No k needed; only min_cluster_size",
  "Labels low-density points as noise → natural outlier detector",
]]].forEach(([x, h, fill, items]) => {
  card(s, x, 1.1, 4.4, 3.0, fill);
  text(s, h, x + 0.2, 1.2, 4.0, 0.4, { fontSize: 16, bold: true, color: C.text2 });
  bullets(s, items, x + 0.2, 1.65, 4.0, 2.4, { fontSize: 13 });
});
text(s, "Input for both: the 5 PCA scores of 77 provinces. Question: are there natural flood-impact types, and which provinces are atypical?",
  0.5, 4.3, 9.0, 0.7, { fontSize: 13, color: C.accent1 });
s.addNotes("เทคนิคแรกคือ clustering เปรียบเทียบ K-Means ที่เรียนในห้อง กับ HDBSCAN ที่ไม่ได้เรียน HDBSCAN เป็น clustering แบบอิงความหนาแน่นเชิงลำดับชั้น ไม่ต้องกำหนดจำนวน k ล่วงหน้า และติดป้าย noise ให้จุดที่อยู่ในบริเวณความหนาแน่นต่ำ จึงใช้หา outlier ได้ ส่วน K-Means เราเลือก k ด้วย elbow, silhouette และความเสถียรเมื่อเปลี่ยน random seed 20 ครั้ง ทั้งสองวิธีใช้ PCA score 5 มิติของ 77 จังหวัดเป็น input");

// ================= 9. Technique – Moran / LISA =================
s = add("CONTENT", "Analytics Technique");
s.addText("Technique 2: Moran’s I and LISA hotspots", { placeholder: "title" });
card(s, 0.5, 1.1, 4.4, 1.55);
text(s, "Global Moran’s I (not covered in class)", 0.7, 1.15, 4.0, 0.4, { fontSize: 15, bold: true, color: C.text2 });
text(s, "I = (n / ΣΣwᵢⱼ) · ΣΣ wᵢⱼ zᵢ zⱼ / Σ zᵢ²", 0.7, 1.55, 4.0, 0.4, { fontSize: 15, fontFace: "Cambria", italic: true });
text(s, "Do neighbouring provinces have similar severity? I > 0 = clustered, ≈ 0 = random.", 0.7, 1.95, 4.0, 0.65, { fontSize: 12 });
card(s, 0.5, 2.8, 4.4, 1.55);
text(s, "LISA – Local Moran’s Iᵢ", 0.7, 2.85, 4.0, 0.4, { fontSize: 15, bold: true, color: C.text2 });
text(s, "One statistic per province: HH = high severity surrounded by high (hotspot), LL = coldspot, HL/LH = spatial outliers. Significance by 9,999 conditional permutations, p < 0.05.",
  0.7, 3.25, 4.0, 1.05, { fontSize: 12 });
text(s, "Spatial weights: Queen contiguity (shared border), row-standardised; mean 4.6 neighbours per province.", 0.5, 4.5, 4.4, 0.6, { fontSize: 11.5, color: C.accent6 });
// quadrant diagram
const qx = 5.4, qy = 1.15, qs = 1.75;
const quads = [["LH", "ABD9E9", 0, 0, C.text1], ["HH  hotspot", HEX.accent2, 1, 0, C.background1], ["LL  coldspot", HEX.accent4, 0, 1, C.background1], ["HL", HEX.accent3, 1, 1, C.text1]];
quads.forEach(([t, col, cx, cy, tc]) => {
  s.addShape(pres.shapes.RECTANGLE, { x: qx + cx * qs, y: qy + cy * qs, w: qs - 0.05, h: qs - 0.05, fill: { color: col }, line: { color: col } });
  text(s, t, qx + cx * qs, qy + cy * qs, qs - 0.05, qs - 0.05, { fontSize: 15, bold: true, align: "center", valign: "middle", color: tc });
});
text(s, "own severity (z) →", qx, qy + 2 * qs + 0.02, 2 * qs, 0.3, { fontSize: 11, align: "center", color: C.accent6 });
text(s, "neighbours’ average (spatial lag) →", qx + 2 * qs + 0.05, qy, 0.6, 2 * qs, { fontSize: 11, color: C.accent6, vert: "vert270", align: "center", valign: "middle" });
s.addNotes("เทคนิคที่สองคือสถิติ spatial autocorrelation ซึ่งไม่ได้เรียนในห้อง Global Moran's I วัดว่าจังหวัดที่ติดกันมีความรุนแรงคล้ายกันหรือไม่ ถ้ามากกว่าศูนย์แปลว่าจับกลุ่มเชิงพื้นที่ ส่วน LISA หรือ Local Moran คำนวณค่าให้แต่ละจังหวัด แบ่งเป็น HH คือ hotspot ที่จังหวัดนั้นและเพื่อนบ้านรุนแรงสูง, LL คือ coldspot, และ HL/LH คือ outlier เชิงพื้นที่ ทดสอบนัยสำคัญด้วย permutation 9,999 ครั้ง น้ำหนักเชิงพื้นที่ใช้ Queen contiguity คือจังหวัดที่มีเขตแดนติดกัน");

// ================= 10. Results – clusters =================
s = add("CONTENT", "Analytics Technique");
s.addText(`Result: ${kmK} flood-impact types (K-Means)`, { placeholder: "title" });
img(s, "clusters_map", 0.3, 1.0, 4.2, 4.2);
const pp = Object.fromEntries(prof.map((r) => [r.km, r]));
const order = [2, 3, 1, 0].filter((k) => pp[k]);
const col = { 0: HEX.accent4, 1: "ABD9E9", 2: HEX.accent3, 3: HEX.accent2 };
const desc = {
  0: "Bangkok ring & Mekong NE – floods declared rarely; low damage per capita",
  1: "East, lower south, parts of NE – moderate, frequent but small",
  2: "North uplands & deep south – most people affected (70/1k/yr), casualties, public sites",
  3: "NE & lower-north basins – 10 rai farmland/km²/yr, fisheries, roads",
};
order.forEach((k, i) => {
  const y = 1.05 + i * 1.03;
  card(s, 4.7, y, 4.85, 0.93, HEX.lt2);
  s.addShape(pres.shapes.OVAL, { x: 4.82, y: y + 0.12, w: 0.32, h: 0.32, fill: { color: col[k] }, line: { color: col[k] } });
  text(s, `K${k} ${nm[k]} (${pp[k].n_provinces})`, 5.22, y + 0.06, 4.25, 0.36, { fontSize: 13.5, bold: true, color: C.text2 });
  text(s, desc[k], 5.22, y + 0.42, 4.25, 0.5, { fontSize: 11.5 });
});
s.addNotes("ผลจาก K-Means ได้ 4 กลุ่ม: K2 น้ำป่าไหลหลาก ผลกระทบต่อคนสูง อยู่ภาคเหนือตอนบนและใต้สุด ผู้ได้รับผลกระทบเฉลี่ย 70 คนต่อพันคนต่อปี; K3 น้ำท่วมพื้นที่เกษตรในลุ่มน้ำ อยู่อีสานและภาคเหนือตอนล่าง พื้นที่เกษตรเสียหาย 10 ไร่ต่อตารางกิโลเมตรต่อปี; K1 ผลกระทบปานกลาง; และ K0 ผลกระทบที่รายงานต่ำ คือกรุงเทพฯ-ปริมณฑลและริมโขง ชื่อกลุ่มตั้งจากโปรไฟล์ค่ามัธยฐานของแต่ละกลุ่ม");

// ================= 11. Results – HDBSCAN =================
s = add("CONTENT", "Analytics Technique");
s.addText("Result: HDBSCAN finds a continuum + outliers", { placeholder: "title" });
img(s, "clusters_pca", 0.4, 1.05, 5.6, 2.6);
bullets(s, [
  [B(`${R.hdbscan_n_clusters} density clusters, ${R.hdbscan_noise} noise `), N(`(min_cluster_size = ${R.hdbscan_params.min_cluster_size}, best silhouette with ≤25% noise)`)],
  [B("One dense core of 71 provinces: "), N("impact varies smoothly – no density gaps between 'types'")],
  [B("Noise = "), N(R.hdbscan_noise_provinces.join(", ") + " – protected metro / extreme lowland provinces")],
], 6.2, 1.05, 3.4, 3.9, { fontSize: 12.5 });
text(s, "Take-away: K-Means gives a usable typology of a continuous gradient; HDBSCAN tells us the gradient has no natural breaks and flags atypical provinces.",
  0.5, 3.85, 5.5, 1.1, { fontSize: 13, color: C.accent1 });
s.addNotes("HDBSCAN ให้ผลต่างจาก K-Means: เจอแกนหนาแน่นกลุ่มเดียว 71 จังหวัด กลุ่มเล็กอีก 1 กลุ่ม และ noise 3 จังหวัด คือ กรุงเทพฯ อยุธยา และสิงห์บุรี แปลว่าผลกระทบน้ำท่วมเปลี่ยนแบบต่อเนื่อง ไม่มีช่องว่างความหนาแน่นแบ่งกลุ่มตามธรรมชาติ K-Means จึงเป็นการแบ่งช่วงของ gradient ให้ใช้งานได้ ส่วน HDBSCAN มีประโยชน์ในการชี้จังหวัดที่ผิดปกติ");

// ================= 12. Results – LISA =================
s = add("CONTENT", "Analytics Technique");
s.addText("Result: impact is spatially clustered", { placeholder: "title" });
img(s, "lisa", 0.3, 1.0, 5.6, 3.55);
stat(s, `I = ${f2(R.moran_I)}`, `Global Moran’s I of severity index (z = ${f2(R.moran_z)}, p < 0.001)`, 6.1, 1.0, 3.5, C.accent2);
const my = R.moran_by_year;
s.addChart(pres.charts.LINE, [{ name: "Moran's I", labels: Object.keys(my), values: Object.values(my).map((v) => v.I) }], {
  x: 6.0, y: 2.2, w: 3.6, h: 1.8, chartColors: [HEX.accent2], lineSize: 2, lineDataSymbolSize: 6,
  showTitle: true, title: "Moran’s I per year (all p < 0.01)", titleFontSize: 11, titleColor: HEX.dk2, titleFontFace: "+mn-lt",
  showValue: true, dataLabelPosition: "t", dataLabelFontSize: 9, dataLabelFormatCode: "0.00", dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
  catAxisLabelColor: HEX.accent6, valAxisLabelColor: HEX.accent6, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  valAxisMinVal: 0, valAxisMaxVal: 0.6, valGridLine: { color: "E5E8E8", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false, valAxisHidden: true,
});
text(s, [B(`${R.lisa_hotspots.length} hotspots: `), N(R.lisa_hotspots.join(", ")), B(`  ·  ${R.lisa_coldspots.length} coldspots: `), N("Bangkok metro, Mekong NE, west coast")],
  0.5, 4.6, 9.0, 0.6, { fontSize: 11.5 });
s.addNotes("Global Moran's I ของดัชนีความรุนแรงเท่ากับ 0.44 มีนัยสำคัญสูงมาก แปลว่าจังหวัดที่ติดกันมีความรุนแรงคล้ายกัน และเป็นแบบนี้ทุกปีตั้งแต่ 2019 ถึง 2024 ค่า I อยู่ระหว่าง 0.26 ถึง 0.42 LISA พบ hotspot 9 จังหวัด กระจุกตัวในภาคเหนือตอนบนและตอนล่าง เช่น เชียงราย ลำปาง แพร่ ตาก สุโขทัย และ coldspot 12 จังหวัด คือกรุงเทพฯ-ปริมณฑล ริมโขงอีสาน และชายฝั่งตะวันตก");

// ================= 13. Demo =================
s = add("CONTENT", "Demo");
s.addText("Demo: Python notebook pipeline", { placeholder: "title" });
const code = [
  "py = build_dataset.main()   # 462 province-years",
  "Z = StandardScaler().fit_transform(np.log1p(F))",
  "P = PCA(n_components=5).fit_transform(Z)",
  "",
  "km = KMeans(4, n_init=50).fit(P)",
  "hd = HDBSCAN(min_cluster_size=3, min_samples=2).fit(P)",
  "",
  "w  = Queen.from_dataframe(provinces); w.transform = 'r'",
  "mi   = Moran(severity, w, permutations=9999)",
  "lisa = Moran_Local(severity, w, permutations=9999)",
];
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.1, w: 5.6, h: 3.4, rectRadius: 0.08, fill: { color: "1B2631" }, line: { color: "1B2631" } });
s.addText(code.map((l, i) => ({ text: l || " ", options: { breakLine: i < code.length - 1 } })),
  { x: 0.65, y: 1.2, w: 5.35, h: 3.2, isTextBox: true, fontFace: "Courier New", fontSize: 11, color: "D5F5E3", valign: "top", margin: 0.05 });
bullets(s, [
  [B("src/build_dataset.py "), N("– parse, clean, harmonise, join")],
  [B("flood_analysis.ipynb "), N("– EDA → preprocessing → clustering → Moran/LISA → evaluation")],
  [B("Libraries: "), N("pandas, geopandas, scikit-learn 1.9 (KMeans, HDBSCAN, PCA), PySAL libpysal + esda")],
  [B("Reproducible: "), N("raw files + one command; all figures and numbers regenerate")],
], 6.3, 1.1, 3.3, 3.5, { fontSize: 12.5 });
text(s, "Live demo: run the notebook top-to-bottom (~1 min) and show the cluster + LISA maps.", 0.5, 4.65, 9.0, 0.45, { fontSize: 12.5, color: C.accent1 });
s.addNotes("ส่วน demo: โค้ดแบ่งเป็นสองส่วน ไฟล์ build_dataset.py สำหรับอ่านและรวมข้อมูล และ notebook สำหรับการวิเคราะห์ทั้งหมด ใช้ pandas, geopandas, scikit-learn สำหรับ K-Means, HDBSCAN, PCA และ PySAL สำหรับ Moran's I กับ LISA ในวิดีโอให้รัน notebook ตั้งแต่ต้นจนจบ ใช้เวลาประมาณ 1 นาที แล้วโชว์แผนที่ cluster และ LISA");

// ================= 14. Evaluation – internal =================
s = add("CONTENT", "Evaluation");
s.addText("Evaluation 1: internal cluster quality", { placeholder: "title" });
const kt = R.kmeans_table;
const head = ["k", "Silhouette", "Davies–Bouldin", "Smallest cluster", "Seed stability (ARI)"];
const ktRows = [head].concat(kt.map((r) => [String(r.k), f3(r.silhouette), f3(r.davies_bouldin), String(r.smallest_cluster), f3(r.seed_stability_ari)]));
s.addTable(ktRows.map((r, i) => r.map((c) => ({ text: c, options: i === 0 ? { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 } } : (r[0] === String(kmK) ? { bold: true, fill: { color: "D6EAF8" } } : {}) }))),
  { x: 0.5, y: 1.1, w: 5.2, colW: [0.5, 1.0, 1.15, 1.15, 1.4], fontSize: 11, fontFace: THEME.bodyFontFace, color: HEX.dk1, border: { type: "solid", color: "D5DBDB", pt: 0.75 }, rowH: 0.3, align: "center", objectName: "k_table" });
text(s, "Rule: best silhouette among k ≥ 3 with smallest cluster ≥ 3 and seed-ARI ≥ 0.75 → k = 4 (k = 2 is only 'more vs less flooded').",
  0.5, 3.65, 5.2, 0.7, { fontSize: 11.5, color: C.accent6 });
const km = ev["K-Means"], hd = ev["HDBSCAN (non-noise)"];
card(s, 6.0, 1.1, 3.6, 3.25);
text(s, "Final models", 6.2, 1.15, 3.2, 0.35, { fontSize: 14, bold: true, color: C.text2 });
const ft = [["", "K-Means", "HDBSCAN"], ["clusters", String(km.clusters), String(hd.clusters)], ["noise", "0", String(hd.noise)],
  ["silhouette", f3(km.silhouette), f3(hd.silhouette)], ["Davies–B. (lower)", f3(km.davies_bouldin), f3(hd.davies_bouldin)], ["Calinski–H.", f2(km.calinski_harabasz), f2(hd.calinski_harabasz)]];
s.addTable(ft.map((r, i) => r.map((c, j) => ({ text: c, options: i === 0 || j === 0 ? { bold: true } : {} }))),
  { x: 6.1, y: 1.5, w: 3.4, colW: [1.5, 0.95, 0.95], fontSize: 11, fontFace: THEME.bodyFontFace, color: HEX.dk1, border: { type: "none" }, rowH: 0.3, objectName: "final_table" });
text(s, "HDBSCAN’s higher silhouette comes from one 71-province cluster – tidy but uninformative.", 6.2, 3.55, 3.2, 0.75, { fontSize: 11.5, italic: true });
text(s, "Silhouette ≈ 0.2 = weak structure: provinces form a gradient, not well-separated groups.", 0.5, 4.45, 9.0, 0.5, { fontSize: 13, color: C.accent1 });
s.addNotes("การประเมินภายใน: ตารางซ้ายคือการเลือก k silhouette สูงสุดที่ k=2 แต่เป็นแค่การแบ่งท่วมมากกับท่วมน้อย k=3 ถึง 8 มี silhouette ใกล้กันมาก ประมาณ 0.2 เราจึงใช้เกณฑ์ความเสถียรเพิ่ม เลือก k=4 ซึ่ง ARI ระหว่าง seed เท่ากับ 0.80 ด้านขวา HDBSCAN มี silhouette สูงกว่า 0.47 แต่เพราะมีกลุ่มใหญ่กลุ่มเดียว 71 จังหวัด ค่า silhouette ประมาณ 0.2 บอกว่าโครงสร้างกลุ่มอ่อน ข้อมูลเป็น gradient");

// ================= 15. Evaluation – external =================
s = add("CONTENT", "Evaluation");
s.addText("Evaluation 2: agreement with recurrence map", { placeholder: "title" });
img(s, "mismatch", 0.3, 0.95, 5.2, 2.75);
const exRows = [["Flagged provinces vs reference top-25%", "n", "F1", "random F1", "p"]].concat(
  ext.map((r) => [r.method.replace("K-Means ", "").replace(" – ", " "), String(r.flagged), f2(r.f1), f2(r.random_f1), r.perm_p < 0.01 ? f3(r.perm_p) : f2(r.perm_p)]));
s.addTable(exRows.map((r, i) => r.map((c) => ({ text: c, options: i === 0 ? { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 } } : (r[0].startsWith("K3") ? { bold: true, fill: { color: "FADBD8" } } : {}) }))),
  { x: 5.6, y: 1.0, w: 4.0, colW: [2.0, 0.35, 0.45, 0.65, 0.55], fontSize: 9.5, fontFace: THEME.bodyFontFace, color: HEX.dk1, border: { type: "solid", color: "D5DBDB", pt: 0.5 }, rowH: 0.3, objectName: "ext_table" });
bullets(s, [
  [B(`Severity vs recurrent share: ρ = ${f2(R.spearman_severity_ref.rho)} `), N("(no overall relation)")],
  [B("Only the agricultural-basin cluster K3 matches "), N(`(F1 ${f2(ext[0].f1)} vs ${f2(ext[0].random_f1)} random, p = ${f3(ext[0].perm_p)}); Kruskal–Wallis across clusters p = ${f3(R.kruskal_km_p)}`)],
  [B("Recurrent-flood hotspots = lower Chao Phraya "), N("(Bangkok, Pathum Thani, Ayutthaya …) where reported impact is LOW")],
], 0.5, 3.8, 9.1, 1.4, { fontSize: 12 });
s.addNotes("การประเมินภายนอกเทียบกับแผนที่น้ำท่วมซ้ำซาก ดัชนีความรุนแรงโดยรวมไม่สัมพันธ์กับสัดส่วนพื้นที่น้ำท่วมซ้ำซาก ρ เท่ากับ 0.03 แต่กลุ่ม K3 น้ำท่วมพื้นที่เกษตรสอดคล้องกับแผนที่อย่างมีนัยสำคัญ F1 0.51 เทียบกับสุ่มที่ 0.30 p เท่ากับ 0.007 ส่วน hotspot น้ำป่าภาคเหนือไม่ตรงเลย เพราะเป็นภัยคนละแบบ LISA ของแผนที่อ้างอิงพบ hotspot ที่ลุ่มเจ้าพระยาตอนล่าง ซึ่งเป็นจังหวัดที่รายงานผลกระทบต่ำ แผนที่ขวาสุดแสดงการจับคู่: ม่วงคือสูงทั้งคู่, แดงคือผลกระทบอย่างเดียว, น้ำเงินคือท่วมซ้ำซากอย่างเดียว");

// ================= 16. Findings =================
s = add("CONTENT", "Conclusion");
s.addText("Key findings", { placeholder: "title" });
const finds = [
  ["1", "Two different flood problems", "Flash floods in the northern uplands & deep south hurt people; basin floods in the NE & lower north hurt farmland."],
  ["2", "Impact is spatially clustered", `Moran’s I = ${f2(R.moran_I)} (p < 0.001), stable every year 2019–2024; ${R.lisa_hotspots.length} northern hotspots.`],
  ["3", "Where water recurs ≠ where damage is reported", "The lower Chao Phraya floods most often but reports little impact – protection works or under-reporting; only the farmland cluster matches the recurrence map."],
  ["4", "Data preparation was most of the work", "Without removing 144 subtotal rows and harmonising 6 schemas, every total would be double-counted."],
];
finds.forEach(([n, h, d], i) => {
  const x = 0.5 + (i % 2) * 4.6, y = 1.1 + Math.floor(i / 2) * 2.0;
  card(s, x, y, 4.4, 1.85);
  s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y + 0.15, w: 0.45, h: 0.45, fill: { color: HEX.accent1 }, line: { color: HEX.accent1 } });
  text(s, n, x + 0.15, y + 0.15, 0.45, 0.45, { fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
  text(s, h, x + 0.75, y + 0.15, 3.55, 0.5, { fontSize: 14, bold: true, color: C.text2, valign: "middle" });
  text(s, d, x + 0.2, y + 0.72, 4.05, 1.1, { fontSize: 12 });
});
s.addNotes("ข้อค้นพบหลัก 4 ข้อ: หนึ่ง ประเทศไทยมีปัญหาน้ำท่วมสองแบบ คือน้ำป่าไหลหลากที่กระทบคนในภาคเหนือและใต้สุด และน้ำท่วมลุ่มน้ำที่กระทบพื้นที่เกษตร สอง ผลกระทบจับกลุ่มเชิงพื้นที่อย่างมีนัยสำคัญทุกปี สาม พื้นที่ท่วมซ้ำซากไม่ใช่พื้นที่ที่รายงานความเสียหายมาก ลุ่มเจ้าพระยาตอนล่างท่วมบ่อยที่สุดแต่รายงานผลกระทบน้อย อาจเพราะมีระบบป้องกันหรือรายงานไม่ครบ สี่ งานส่วนใหญ่คือการเตรียมข้อมูล");

// ================= 17. Applications =================
s = add("CONTENT", "Conclusion");
s.addText("Possible applications", { placeholder: "title" });
const apps = [
  ["Targeted preparedness", "K2 flash-flood provinces: early-warning & evacuation; K3 basin provinces: crop insurance, drainage, farmer compensation"],
  ["Budget allocation", "Use the severity index + LISA hotspots to prioritise DDPM disaster-relief budgets by region rather than province by province"],
  ["Data-quality audit", "'Recurrence-only' provinces and HDBSCAN outliers flag possible under-reporting to check with local offices"],
  ["Monitoring", "Re-run yearly as new DDPM data arrives; track whether hotspots move (Moran’s I per year)"],
];
apps.forEach(([h, d], i) => {
  const y = 1.1 + i * 1.0;
  card(s, 0.5, y, 9.0, 0.88);
  text(s, h, 0.7, y + 0.08, 2.6, 0.72, { fontSize: 14, bold: true, color: C.text2, valign: "middle" });
  text(s, d, 3.3, y + 0.08, 6.05, 0.72, { fontSize: 12.5, valign: "middle" });
});
s.addNotes("การนำไปใช้: หนึ่ง เตรียมรับมือให้ตรงประเภท จังหวัดน้ำป่าเน้นเตือนภัยล่วงหน้าและอพยพ จังหวัดลุ่มน้ำเน้นประกันพืชผลและการระบายน้ำ สอง จัดสรรงบประมาณตาม hotspot ระดับภูมิภาค สาม ตรวจคุณภาพข้อมูล จังหวัดที่ท่วมซ้ำซากแต่รายงานผลกระทบต่ำอาจมีการรายงานไม่ครบ สี่ รันซ้ำทุกปีเพื่อติดตามว่า hotspot เคลื่อนที่หรือไม่");

// ================= 18. References =================
s = add("CONTENT", "Conclusion");
s.addText("References and data sources", { placeholder: "title" });
bullets(s, [
  "DDPM – สถิติรายปีการเกิดอุทกภัย 2562–2567, catalog.disaster.go.th/dataset/dpm-gd027",
  "HII / ThaiWater – ปริมาณฝนเชิงพื้นที่รายเดือนรายจังหวัด และค่าเฉลี่ย 30 ปี, data.go.th/dataset/spatial-rain",
  "DOPA – สถิติจำนวนประชากร 2562–2567, data.go.th/dataset/dopa-star",
  "Land Development Department – สถิติพื้นที่น้ำท่วมซ้ำซากรายตำบล, data.go.th/dataset/ldd_21_04",
  "Province boundaries – github.com/chingchai/OpenGISData-Thailand",
  "Campello, Moulavi & Sander (2013). Density-based clustering based on hierarchical density estimates. PAKDD.",
  "Anselin, L. (1995). Local Indicators of Spatial Association – LISA. Geographical Analysis 27(2).",
  "Moran, P.A.P. (1950). Notes on continuous stochastic phenomena. Biometrika 37.",
  "Rey & Anselin (2007). PySAL: A Python library of spatial analytical methods. scikit-learn 1.9 documentation.",
], 0.5, 1.05, 9.0, 4.1, { fontSize: 11.5 });
s.addNotes("รายการอ้างอิงแหล่งข้อมูลและเอกสารวิชาการของเทคนิคที่ใช้");

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
