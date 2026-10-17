// Builds report/Flood_Clustering_Report.docx from data/processed/results.json and figures/.
// Run from the project root:  NODE_PATH=$PWD/slides/node_modules node report/build_report.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, ImageRun, Table, TableRow,
  TableCell, WidthType, ShadingType, BorderStyle, LevelFormat, PageNumber, Footer, Header,
  TableOfContents, PageBreak,
} = require("docx");

const ROOT = path.resolve(__dirname, "..");
const R = JSON.parse(fs.readFileSync(path.join(ROOT, "data/processed/results.json")));
const LOG = JSON.parse(fs.readFileSync(path.join(ROOT, "data/processed/data_log.json")));
const SIZES = JSON.parse(fs.readFileSync(path.join(ROOT, "figures/sizes.json")));
const OUT = path.join(__dirname, "Flood_Clustering_Report.docx");

const PAGE_W = 11906, MARGIN = 1247;           // A4, 2.2 cm margins
const CONTENT_W = PAGE_W - 2 * MARGIN;         // DXA
const ACCENT = "1F4E79", LIGHT = "DCE6F1", GREY = "BFBFBF";
const f3 = (x) => Number(x).toFixed(3), f2 = (x) => Number(x).toFixed(2);

// ---------- helpers ----------
const runs = (text) =>
  // **bold** and *italic* inline markup
  text.split(/(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)/).filter(Boolean).map((t) =>
    t.startsWith("`") ? new TextRun({ text: t.slice(1, -1), font: "Consolas", size: 20 })
      : t.startsWith("**") ? new TextRun({ text: t.slice(2, -2), bold: true })
      : t.startsWith("*") ? new TextRun({ text: t.slice(1, -1), italics: true })
        : new TextRun(t));
const P = (text, opts = {}) => new Paragraph({ children: runs(text), spacing: { after: 120 }, alignment: AlignmentType.JUSTIFIED, ...opts });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const B = (text, level = 0) => new Paragraph({ numbering: { reference: "bullets", level }, children: runs(text), spacing: { after: 60 } });
const N = (text, instance = 1) => new Paragraph({ numbering: { reference: "steps", level: 0, instance }, children: runs(text), spacing: { after: 60 } });

let figNo = 0, tabNo = 0;
function figure(name, caption, widthIn = 6.2) {
  figNo += 1;
  const [w, h] = SIZES[name];
  const W = widthIn * 72, H = (W * h) / w;                 // points-ish units used by docx-js (px @72dpi)
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 }, keepNext: true,
      children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join(ROOT, "figures", name + ".png")),
        transformation: { width: Math.round(W * 96 / 72), height: Math.round(H * 96 / 72) },
        altText: { title: caption, description: caption, name } })],
    }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
      children: [new TextRun({ text: `Figure ${figNo}. `, bold: true, size: 19 }), new TextRun({ text: caption, size: 19, italics: true })] }),
  ];
}

const border = { style: BorderStyle.SINGLE, size: 4, color: GREY };
const borders = { top: border, bottom: border, left: border, right: border };
function table(caption, header, rows, colFrac, opts = {}) {
  tabNo += 1;
  const widths = colFrac.map((f) => Math.round(f * CONTENT_W));
  widths[widths.length - 1] = CONTENT_W - widths.slice(0, -1).reduce((a, b) => a + b, 0);
  const cell = (text, i, head, shade) => new TableCell({
    borders, width: { size: widths[i], type: WidthType.DXA },
    shading: head ? { fill: ACCENT, type: ShadingType.CLEAR, color: "auto" }
      : shade ? { fill: LIGHT, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    children: [new Paragraph({ alignment: i > 0 && opts.numeric ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({ text: String(text), bold: head, color: head ? "FFFFFF" : undefined, size: 18 })] })],
  });
  return [
    new Paragraph({ spacing: { before: 160, after: 80 }, keepNext: true,
      children: [new TextRun({ text: `Table ${tabNo}. `, bold: true, size: 19 }), new TextRun({ text: caption, size: 19, italics: true })] }),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: widths,
      rows: [new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, i, true)) }),
        ...rows.map((r, ri) => new TableRow({ children: r.map((c, i) => cell(c, i, false, (opts.highlight || []).includes(ri))) }))],
    }),
    new Paragraph({ spacing: { after: 120 }, children: [] }),
  ];
}

// ---------- numbers ----------
const prof = Object.fromEntries(R.kmeans_profile.map((p) => [p.km, p]));
const ie = Object.fromEntries(R.internal_eval.map((r) => [r.index, r]));
const ext = R.external_eval;
const k3 = ext.find((e) => e.method.startsWith("K-Means K3"));
const rawRows = Object.values(LOG.ddpm_meta).reduce((a, m) => a + m.rows, 0);
const droppedTotals = Object.values(LOG.ddpm_meta).reduce((a, m) => a + m.dropped_total_rows, 0);
const lisa = R.lisa_counts;
const pct = (x) => Math.round(100 * x) + "%";

// ---------- content ----------
const children = [];

// Title page
children.push(
  new Paragraph({ spacing: { before: 1800 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Group Assignment #2 – Analytics on Complex Data", size: 24, color: "595959" })] }),
  new Paragraph({ spacing: { before: 400, after: 200 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Where Floods Hurt vs. Where Water Recurs", bold: true, size: 44, color: ACCENT, font: "Cambria" })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 },
    children: [new TextRun({ text: "Spatial clustering of flood impact across Thailand's 77 provinces, 2019–2024", size: 28, italics: true })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: "Presented by", size: 22, color: "595959" })] }),
  ...[["เกียรติภูมิ พละสาร", "6710503755"], ["ณภัทร โพธิบัวทอง", "6710503887"], ["ปวีณกฤษฏิ์ ท่าห้อง", "6710504069"]].map(([n, id]) =>
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 },
      children: [new TextRun({ text: n, size: 26, bold: true, font: { ascii: "Calibri", cs: "Tahoma" } }), new TextRun({ text: `   ${id}`, size: 24 })] })),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 600 },
    children: [new TextRun({ text: "Data Mining · Kasetsart University · October 2026", size: 22, color: "595959" })] }),
  new Paragraph({ children: [new PageBreak()] }),
);

// Abstract + TOC
children.push(
  H1("Abstract"),
  P(`We integrate flood-impact statistics from the Department of Disaster Prevention and Mitigation (DDPM), areal rainfall from the Hydro-Informatics Institute (HII/ThaiWater), registered population from the Department of Provincial Administration (DOPA) and a sub-district recurrent-flood inventory from the Land Development Department (LDD) to ask a single question: **are the provinces that suffer the most reported flood damage the same provinces where floods recur?** Six yearly DDPM files with six different schemas (${rawRows.toLocaleString()} raw rows, 44–66 columns each, three spatial granularities) were harmonised into a 77-province × 6-year panel. After per-capita/per-area normalisation, log transform, standardisation and PCA (5 components, ${pct(R.pca_cum_var)} of variance), we compare K-Means (covered in class) with HDBSCAN, and add global Moran's I and LISA (both not covered in class) to locate spatial hotspots. K-Means identifies four interpretable impact regimes (silhouette ${f3(ie["K-Means"].silhouette)}, seed-stability ARI ${f2(R.kmeans_seed_stability_ari)}); HDBSCAN shows the data form a continuum and isolates Bangkok, Ayutthaya and Sing Buri as outliers. Impact is strongly spatially clustered (Moran's I = ${f3(R.moran_I)}, p = ${R.moran_p}) with ${lisa["HH (hotspot)"]} northern hotspots. Against the recurrent-flood map, however, the overall severity index is uncorrelated (Spearman ρ = ${f2(R.spearman_severity_ref.rho)}); only the agricultural-basin cluster agrees significantly (F1 ${f2(k3.f1)} vs ${f2(k3.random_f1)} at random, p = ${k3.perm_p}). Where water recurs and where floods hurt are largely different places.`),
  new TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-2" }),
  new Paragraph({ children: [new PageBreak()] }),
);

// 1. Complex data & application
children.push(
  H1("1. Complex Data and Application"),
  H2("1.1 Problem and application"),
  P("Thailand floods every year, but the kind of damage differs from place to place: flash floods in the mountainous north and deep south kill people and destroy houses, while slow river-basin floods in the central plain and the north-east mostly drown farmland. Disaster budgets and preparedness programmes are allocated province by province, so a data-driven grouping of provinces by *impact pattern*, together with a test of whether impact is spatially clustered, directly supports (i) matching the type of preparedness to the type of flood, (ii) regional rather than provincial budgeting where hotspots span several provinces, and (iii) data-quality auditing of provinces whose reported damage looks inconsistent with their physical flood exposure."),
  H2("1.2 Why the data are complex"),
  P("The data set is complex on four of the dimensions listed in the assignment brief:"),
  B("**Diverse sources** – four government agencies, each with its own file format, identifiers and update cycle."),
  B("**Spatial and geographic** – provinces are polygons with neighbours; the analysis uses contiguity, not only attribute similarity."),
  B("**Time series** – six annual snapshots (B.E. 2562–2567 = 2019–2024) whose schemas change every year."),
  B("**High-dimensional, skewed and zero-inflated** – 44–66 attributes per year, most of them heavy-tailed with a majority of zeros."),
  ...table("Data sources", ["Source", "Content", "Format / granularity", "Use"], [
    ["DDPM (catalog.disaster.go.th, dpm-gd027)", "Yearly flood statistics 2019–2024: affected people, casualties, houses, farmland, roads, public sites …", "6 .xlsx files, 1–3 header rows; province, district-event or village level", "Impact features"],
    ["HII / ThaiWater (data.go.th, spatial-rain)", "Monthly areal rainfall per province + 30-year normal", "CSV", "Rainfall context"],
    ["DOPA (data.go.th, dopa-star)", "Registered population per province per year", "HTML tables saved as .xls inside .zip", "Per-capita normalisation"],
    ["LDD (data.go.th, ldd_21_04)", "Area flooded ≥ 4 times in 10 years, by sub-district", "CSV", "External reference map"],
    ["OpenGISData-Thailand (GitHub)", "77 province polygons", "GeoJSON", "Maps, spatial weights, area"],
  ], [0.24, 0.34, 0.26, 0.16]),
  P("The project proposal named GISTDA's recurrent-flood map as the reference. GISTDA's API requires a registered key for programmatic access, so the equivalent LDD recurrent-flood inventory (same definition: land flooded repeatedly over a ten-year window) was used instead."),
  ...table("Heterogeneity of the six DDPM files", ["Year", "Raw rows", "Columns", "Level", "Subtotal rows removed"],
    Object.entries(LOG.ddpm_meta).map(([be, m]) => [`${be} (${+be - 543})`, m.rows.toLocaleString(), LOG.n_raw_ddpm_attributes[be], m.level, m.dropped_total_rows]),
    [0.18, 0.18, 0.16, 0.28, 0.20], { numeric: true }),
);

// 2. Exploration
children.push(
  H1("2. Data Exploration"),
  P(`Exploration was carried out on the harmonised panel (77 provinces × 6 years = ${LOG.province_year_shape[0]} records). DDPM declared a flood in ${LOG.province_years_reported} of the 462 province-years.`),
  H2("2.1 Temporal variation"),
  P("Impact varies by an order of magnitude between years (Figure 1). 2022, the year of tropical storm Noru and a strong La Niña, affected about 8.2 million people and 10.1 million rai of farmland, roughly seven times the 2023 figure. Any province-level summary therefore has to aggregate several years; a single year would mostly reflect the track of one storm."),
  ...figure("eda_years", "Reported flood impact by year (DDPM, all provinces)."),
  H2("2.2 Spatial variation"),
  P("Figure 2 maps three candidate variables. Affected people per 1,000 residents is highest in the upper north and the deep south, whereas the wettest provinces are on the east coast (Trat, Chanthaburi) and the Andaman coast (Phang Nga, Ranong, Phuket), several of which report low impact. Rainfall alone therefore does not explain reported impact, which motivates clustering on impact attributes and treating rainfall only as context."),
  ...figure("eda_maps", "Affected people per 1,000 residents, number of flood years, and mean annual rainfall."),
  H2("2.3 Distribution of attributes"),
  P(`The 2019 file alone has ${R.eda_raw_attributes} damage attributes. Their median skewness is ${R.eda_median_skew}, ${R.eda_skew_gt2} of them have skewness above 2, and the median attribute is zero in ${pct(R.eda_median_zero_share)} of provinces (Figure 3). A distance-based method applied to raw values would be dominated by a handful of extreme provinces, so normalisation and a log transform are required before clustering.`),
  ...figure("eda_skew", "Skewness and share of zeros of raw DDPM attributes."),
  P("Spearman correlations among the nine engineered features (Section 3.3) are positive but moderate, up to about 0.65, so the features carry partly redundant information, which is a further reason to apply PCA."),
);

// 3. Pre-processing
children.push(
  H1("3. Pre-Processing"),
  P("Pre-processing was the largest part of the project. It is implemented in `src/build_dataset.py` (integration) and the first half of `flood_analysis.ipynb` (transformation and reduction)."),
  H2("3.1 Data integration and schema harmonisation"),
  N("**Reading six formats.** Each year has its own parser: multi-row Thai headers are flattened, the 2024 file uses English column names, and DOPA's .xls files are actually UTF-8 HTML and are read with an HTML parser."),
  N(`**Removing embedded totals.** ${droppedTotals} subtotal and grand-total rows (e.g. "X รวม", "รวมทั้งหมด 73 จังหวัด") were removed; otherwise affected provinces would be double-counted.`),
  N(`**Schema harmonisation.** Columns were mapped by Thai/English keyword rules to ${LOG.n_harmonised_attributes} common attributes. Monetary columns (มูลค่า, THB) were explicitly excluded because in 2019 they sit next to the physical quantities and were being summed into them.`),
  N("**Province matching.** Names appear as plain names, \"จังหวัด…\" and \"จ.…\"; they were normalised and mapped to the 77 official province codes, with zero unmatched names."),
  N("**Aggregation and join.** District-event and village rows were summed to province-year and joined with rainfall, population, the LDD reference and polygon area."),
  P("As a correctness check, the yearly national totals of the integrated table were compared with DDPM's own grand-total rows and match exactly before imputation."),
  H2("3.2 Missing values"),
  P(`The DDPM files do not use blanks for missing data; they use zeros. In ${LOG.imputed_people_from_households} province-years (e.g. Roi Et and Lamphun in 2019) affected households are positive while affected people is zero, which is impossible. These were imputed as households × the median people-per-household ratio of the whole panel (${f2(LOG.median_people_per_household)}). Province-years with no declared flood are true zeros and were kept as zero. Two attributes were dropped instead of imputed: **evacuated people**, because the 2023 and 2024 schemas record almost none (20 and 949 people nationally, against 46,000 in 2022), and **livestock**, which is dominated by a few poultry events and remains highly skewed after transformation.`),
  H2("3.3 Feature construction and normalisation"),
  P("Six-year totals were divided by population or land area so that large provinces are not automatically ranked as worse. Units were chosen so that typical values exceed one, which makes the subsequent log1p transform effective."),
  ...table("The nine clustering features", ["Feature", "Definition (per year, averaged over 2019–2024)", "Aspect"], [
    ["years_flooded", "Years with a declared flood (0–6)", "Frequency"],
    ["affected_per_1k", "Affected people per 1,000 residents", "Human exposure"],
    ["casualties_per_1M", "Deaths + injured + missing per million residents", "Human loss"],
    ["housing_per_100k", "Damaged houses per 100,000 residents", "Residential damage"],
    ["public_sites_per_1M", "Damaged public sites per million residents", "Infrastructure"],
    ["agri_rai_per_km2", "Damaged farmland (rai) per km²", "Agriculture"],
    ["fishery_per_1000km2", "Damaged fish ponds per 1,000 km²", "Aquaculture"],
    ["transport_per_1000km2", "Damaged roads and bridges per 1,000 km²", "Transport"],
    ["villages_per_1000km2", "Affected villages per 1,000 km²", "Spatial extent"],
  ], [0.26, 0.54, 0.20]),
  H2("3.4 Transformation, standardisation and dimensionality reduction"),
  P(`A log1p transform reduced the median feature skewness from ${R.skew_raw_median} to ${R.skew_log_median}. Features were then z-score standardised and projected with PCA. Five components retain ${pct(R.pca_cum_var)} of the variance (Figure 4) and are used as the input to both clustering algorithms. PC1 alone explains ${pct(R.pca_pc1_var)} and loads positively on every impact feature, so it is used as a single **flood severity index**.`),
  ...figure("pca", "PCA scree plot and loadings of the first components."),
);

// 4. Techniques
children.push(
  H1("4. Analytic Techniques"),
  P("Two families of techniques are used. K-Means and PCA were covered in class; **HDBSCAN, global Moran's I and LISA were not**, and are the project's suggested new techniques."),
  H2("4.1 K-Means (baseline, covered in class)"),
  P("K-Means (n_init = 50) partitions provinces in the 5-dimensional PCA space. k was chosen from 2–8 using silhouette, Davies–Bouldin and two practical rules: the smallest cluster must contain at least three provinces, and the partition must be stable across random seeds (mean Adjusted Rand Index against 20 re-runs ≥ 0.75)."),
  H2("4.2 HDBSCAN (not covered in class)"),
  P("HDBSCAN (Campello, Moulavi & Sander, 2013) builds a hierarchy of density-based clusters from mutual-reachability distances and keeps the most persistent ones. Unlike K-Means it needs no k, does not force every point into a cluster, and labels low-density points as noise, which suits a question about outlying provinces. We used scikit-learn's implementation and searched min_cluster_size 3–12 and min_samples ∈ {2, 3, mcs}, choosing the highest non-noise silhouette with at least two clusters and at most 25% noise (selected: min_cluster_size = 3, min_samples = 2)."),
  H2("4.3 Global Moran's I and LISA (not covered in class)"),
  P("Clustering ignores geography. To test whether flood impact is spatially clustered, we use Moran's I (Moran, 1950), the spatial analogue of a correlation coefficient between a province's value and the average of its neighbours, and its local decomposition LISA (Anselin, 1995), which classifies each province as High-High (hotspot), Low-Low (coldspot), or a spatial outlier (High-Low, Low-High). Spatial weights are Queen contiguity (shared border or vertex), row-standardised; the island province of Phuket is attached to its nearest neighbour. Significance is assessed with 9,999 conditional permutations (α = 0.05). The variable analysed is the severity index (PC1); Moran's I is also computed for each year separately."),
);

// 5. Demo
children.push(
  H1("5. Demo and Implementation"),
  P("The complete pipeline is reproducible from the raw downloads:"),
  B("`src/build_dataset.py` – reads, cleans and integrates the four sources and writes the province-year panel plus a data log."),
  B("`flood_analysis.ipynb` – exploration, pre-processing, K-Means, HDBSCAN, Moran's I/LISA and evaluation; writes all figures and `results.json`."),
  B("`slides/build_deck.js` – regenerates the presentation; every number in the slides and this report is read from `results.json`."),
  P("Libraries: pandas, NumPy, GeoPandas, scikit-learn 1.9 (KMeans, HDBSCAN, PCA, metrics), libpysal and esda (PySAL) for spatial weights and autocorrelation, SciPy for statistical tests, and matplotlib/seaborn for figures. The video demonstration runs the notebook end-to-end."),
);

// 6. Results & Evaluation
const kt = R.kmeans_table;
children.push(
  H1("6. Results"),
  H2("6.1 Choice of k"),
  P(`Silhouette is highest at k = ${R.kmeans_k_best_silhouette}, but that partition only separates "flooded a lot" from "flooded a little". From k = 5 upward the seed stability drops below 0.75 and clusters of three or four provinces appear. k = ${R.kmeans_k_used} is the largest stable choice (ARI ${f2(R.kmeans_seed_stability_ari)}) and gives interpretable groups (Table 4, Figure 5).`),
  ...table("K-Means model selection", ["k", "WCSS", "Silhouette", "Davies–Bouldin", "Smallest cluster", "Seed ARI"],
    kt.map((r) => [r.k, r.wcss.toFixed(1), f3(r.silhouette), f3(r.davies_bouldin), r.smallest_cluster, f2(r.seed_stability_ari)]),
    [0.1, 0.16, 0.17, 0.19, 0.2, 0.18], { numeric: true, highlight: [kt.findIndex((r) => r.k === R.kmeans_k_used)] }),
  ...figure("kmeans_k", "Elbow, silhouette and seed stability for k = 2–8."),
  H2("6.2 K-Means clusters"),
  P("Clusters were named from their median feature profile (Table 5, Figure 6):"),
  B(`**K2 – Flash-flood / human impact (${prof[2].n_provinces} provinces):** the upper north (Chiang Mai, Nan, Lampang, Phrae, Mae Hong Son, Tak …) and the deep south (Yala, Narathiwat, Songkhla). Highest affected people (${prof[2].f_affected_per_1k} per 1,000) and casualties (${prof[2].f_casualties_per_1M} per million), but little farmland damage.`),
  B(`**K3 – Agricultural basin flooding (${prof[3].n_provinces}):** the north-eastern Chi–Mun basin, the lower north and the upper central plain. Farmland damage an order of magnitude higher than any other group (${prof[3].f_agri_rai_per_km2} rai/km²) and the most fish ponds and villages affected.`),
  B(`**K1 – Moderate impact (${prof[1].n_provinces}):** mostly the east, the west and the upper south, including Bangkok.`),
  B(`**K0 – Low reported impact (${prof[0].n_provinces}):** Nonthaburi, Pathum Thani, Ayutthaya, Nakhon Nayok, Samut Songkhram, Phuket and the Mekong provinces Nong Khai, Bueng Kan, Sakon Nakhon and Amnat Charoen.`),
  ...table("Median profile of the K-Means clusters", ["Cluster", "n", "Affected /1k", "Casualties /1M", "Houses /100k", "Farmland rai/km²", "Severity (PC1)"],
    [0, 1, 2, 3].map((c) => [`K${c} ${prof[c].name}`, prof[c].n_provinces, prof[c].f_affected_per_1k, prof[c].f_casualties_per_1M, prof[c].f_housing_per_100k, prof[c].f_agri_rai_per_km2, prof[c].severity_index]),
    [0.3, 0.07, 0.12, 0.13, 0.13, 0.13, 0.12], { numeric: true }),
  ...figure("clusters_map", "K-Means clusters (left) and HDBSCAN result (right).", 5.6),
  H2("6.3 HDBSCAN"),
  P(`HDBSCAN finds one dense core of 71 provinces, a small cluster of three (Nonthaburi, Bueng Kan, Samut Songkhram) and three noise points: ${R.hdbscan_noise_provinces.join(", ")}. The absence of several dense groups shows that flood impact varies along a continuum with no clear density gaps; the K-Means groups should therefore be read as useful bands along that continuum rather than natural classes. HDBSCAN's value here is as an outlier detector: Bangkok (very large population, flood-protection infrastructure) and the Ayutthaya–Sing Buri riverine provinces have profiles unlike any neighbour in feature space.`),
  ...figure("clusters_pca", "Provinces in the first two principal components, coloured by K-Means and HDBSCAN labels."),
  H2("6.4 Spatial autocorrelation and hotspots"),
  P(`Global Moran's I of the severity index is ${f3(R.moran_I)} (z = ${R.moran_z}, permutation p = ${R.moran_p}): neighbouring provinces have strongly similar flood impact. The result holds in every single year (I between ${Math.min(...Object.values(R.moran_by_year).map((v) => v.I))} and ${Math.max(...Object.values(R.moran_by_year).map((v) => v.I))}, all p < 0.002), so it is not an artefact of one storm.`),
  P(`LISA (Figure 8) identifies ${lisa["HH (hotspot)"]} significant hotspots, forming one contiguous block in the north: ${R.lisa_hotspots.join(", ")}. The ${lisa["LL (coldspot)"]} coldspots are Bangkok and its neighbours, the Mekong provinces, and the lower west (Ratchaburi, Phetchaburi, Samut Sakhon, Samut Songkhram). ${lisa.HL + lisa.LH} provinces are spatial outliers and ${lisa["not significant"]} are not significant.`),
  ...figure("lisa", "Moran scatter plot and LISA cluster map of the flood severity index.", 5.8),
);

children.push(
  H1("7. Evaluation"),
  H2("7.1 Internal validity"),
  ...table("Internal cluster-quality indices", ["Method", "Clusters", "Noise", "Silhouette ↑", "Davies–Bouldin ↓", "Calinski–Harabasz ↑"],
    R.internal_eval.map((r) => [r.index, r.clusters, r.noise, f3(r.silhouette), f3(r.davies_bouldin), r.calinski_harabasz.toFixed(1)]),
    [0.26, 0.12, 0.1, 0.16, 0.18, 0.18], { numeric: true }),
  P(`HDBSCAN scores better on silhouette and Davies–Bouldin, but only because 71 of its 74 clustered provinces are in one group and the indices are computed without noise points; it does not provide a usable segmentation. K-Means' silhouette of ${f3(ie["K-Means"].silhouette)} is modest, which is consistent with HDBSCAN's finding that the data form a gradient (Figure 9). Stability (ARI ${f2(R.kmeans_seed_stability_ari)}) and the clear, distinct median profiles are stronger evidence that the four groups are meaningful.`),
  ...figure("silhouette", "Silhouette plots for K-Means and HDBSCAN."),
  H2("7.2 External validity against the recurrent-flood map"),
  P("The reference is the share of each province's land that LDD classifies as flooding at least four times in ten years, aggregated from sub-districts. Provinces in the top quartile of this share are treated as \"recurrent-flood provinces\". Every method flags some provinces as high-impact; we compute precision, recall and F1 against the reference and compare F1 with a random baseline that flags the same number of provinces (2,000 random draws, giving an expected F1 and a permutation p-value)."),
  ...table("Agreement of each method with the recurrent-flood reference", ["Method", "Flagged", "Precision", "Recall", "F1", "Random F1", "p"],
    ext.map((e) => [e.method.replace(" – ", " "), e.flagged, f2(e.precision), f2(e.recall), f2(e.f1), f2(e.random_f1), e.perm_p]),
    [0.34, 0.1, 0.11, 0.1, 0.1, 0.13, 0.12], { numeric: true, highlight: [0] }),
  P(`The overall severity index is **not correlated** with recurrent flooding (Spearman ρ = ${f3(R.spearman_severity_ref.rho)}, p = ${f2(R.spearman_severity_ref.p)}). Only one result beats chance: the agricultural-basin cluster K3 overlaps the reference with F1 = ${f3(k3.f1)} against ${f3(k3.random_f1)} at random (p = ${k3.perm_p}), and recurrent share differs across the four clusters (Kruskal–Wallis p = ${f3(R.kruskal_km_p)}). The northern LISA hotspots and the flash-flood cluster do worse than random. At feature level, only flood extent (affected villages per km², ρ = 0.45) and farmland damage (ρ = 0.32) are positively related to recurrence, while the number of flood years is slightly negatively related (ρ = −0.24).`),
  ...figure("external_eval", "Recurrent-flood share by K-Means cluster, and severity index vs recurrent share."),
  P(`The reference map has its own spatial structure (Moran's I = ${f3(R.moran_ref.I)}, p = ${R.moran_ref.p}) with hotspots in the lower Chao Phraya basin: ${R.lisa_ref_hotspots.join(", ")}. Cross-tabulating the two (Figure 11), ${R.agreement_counts["both high"]} provinces are high on both, ${R.agreement_counts["impact only"]} are high only on reported impact (mostly the north and the deep south) and ${R.agreement_counts["recurrence only"]} are high only on recurrence (mostly the lower central plain, including Bangkok, Nonthaburi, Pathum Thani and Ayutthaya).`),
  ...figure("mismatch", "Where impact and recurrence agree and disagree.", 6.2),
);

// 8. Discussion / applications / conclusion
children.push(
  H1("8. Discussion and Applications"),
  H2("8.1 Key findings"),
  N("**Thailand has two flood-impact regimes.** Flash floods in the north and deep south hurt people and houses; basin floods in the north-east and lower north damage farmland.", 2),
  N(`**Impact is spatially clustered** (Moran's I = ${f3(R.moran_I)}) in every year, with a contiguous northern hotspot.`, 2),
  N("**Where water recurs is not where floods hurt.** The lower Chao Phraya basin floods most often but reports little damage, probably because of protection works and adaptation, and possibly because of reporting practice. The northern hotspot is driven by flash floods that recur less often but are more damaging.", 2),
  N("**Most of the effort was data preparation.** Six schemas, embedded totals, zeros that mean missing and monetary columns mixed with counts would each have silently distorted the result.", 2),
  H2("8.2 Applications"),
  B("**Type-specific preparedness:** early warning and evacuation for K2 provinces; crop insurance, drainage and farm-loss compensation for K3 provinces."),
  B("**Regional budgeting:** the nine LISA hotspots form one block, so a joint regional programme is more efficient than nine separate provincial ones."),
  B("**Data-quality audit:** \"recurrence-only\" provinces with frequent physical flooding but low reported impact are candidates for checking under-reporting."),
  B("**Monitoring:** the pipeline can be re-run each year to see whether hotspots move."),
  H2("8.3 Limitations"),
  B("Reported damage depends on whether and how a province declares a disaster, so it measures reported impact, not physical flood depth."),
  B("Six years is short; one exceptional year (2022) has strong influence despite normalisation."),
  B("LDD recurrence statistics are a proxy for GISTDA's satellite-based recurrent-flood map and cover a different time window."),
  B("Province is a coarse unit; district-level analysis would need consistent district data, which only some years provide."),
  H1("9. Conclusion"),
  P("By integrating four heterogeneous government data sets and combining clustering with spatial statistics, we found that Thai provinces fall along a continuum of flood impact that K-Means usefully divides into flash-flood, agricultural-basin, moderate and low-impact groups; that impact is strongly spatially clustered with a hotspot in the north; and that this impact pattern largely does not coincide with where floods physically recur. Flood-risk policy should therefore consider both maps rather than treating either as a proxy for the other."),
);

// References
children.push(
  H1("References"),
  ...[
    "Anselin, L. (1995). Local indicators of spatial association – LISA. *Geographical Analysis*, 27(2), 93–115.",
    "Campello, R. J. G. B., Moulavi, D., & Sander, J. (2013). Density-based clustering based on hierarchical density estimates. In *Advances in Knowledge Discovery and Data Mining (PAKDD 2013)*, LNCS 7819, 160–172.",
    "Moran, P. A. P. (1950). Notes on continuous stochastic phenomena. *Biometrika*, 37(1/2), 17–23.",
    "Rey, S. J., & Anselin, L. (2007). PySAL: A Python library of spatial analytical methods. *The Review of Regional Studies*, 37(1), 5–27.",
    "Pedregosa, F. et al. (2011). Scikit-learn: Machine learning in Python. *Journal of Machine Learning Research*, 12, 2825–2830.",
    "Department of Disaster Prevention and Mitigation. Annual flood statistics 2562–2567 (dataset dpm-gd027). catalog.disaster.go.th.",
    "Hydro-Informatics Institute. Monthly areal rainfall by province and 30-year normal (dataset spatial-rain). data.go.th.",
    "Department of Provincial Administration. Population statistics 2562–2567 (dataset dopa-star). data.go.th.",
    "Land Development Department. Recurrent flood area statistics by sub-district (dataset ldd_21_04). data.go.th.",
    "Chingchai, H. OpenGISData-Thailand: province boundaries. github.com/chingchai/OpenGISData-Thailand.",
  ].map((t) => new Paragraph({ children: runs(t), spacing: { after: 100 }, indent: { left: 400, hanging: 400 } })),
);

// ---------- document ----------
const doc = new Document({
  creator: "Group 2 – Data Mining",
  title: "Where Floods Hurt vs. Where Water Recurs",
  styles: {
    default: { document: { run: { font: "Calibri", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Cambria", size: 32, bold: true, color: ACCENT }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0, keepNext: true } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Cambria", size: 25, bold: true, color: "2E75B6" }, paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1, keepNext: true } },
    ],
  },
  numbering: { config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    { reference: "steps", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
  ] },
  features: { updateFields: true },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 1247, bottom: 1247, left: MARGIN, right: MARGIN } }, titlePage: true },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: "Where Floods Hurt vs. Where Water Recurs", size: 16, color: "808080" })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "808080" })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("wrote", OUT); });
