// Génère Presentation_Economie.pptx (sujets 1, 2 et 3). Lancer : node generer_powerpoint.js
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(path.join(
  process.env.USERPROFILE,
  ".claude/skills/synced/5c85afb7-3a6e-4a11-9b64-e650b58d7607_a2f95966-c5d5-4e73-9081-d761e4029ec0/pptx/scripts/apply_theme.js"
));

const OUT = process.env.PPTX_OUT || path.join(__dirname, "Presentation_Economie.pptx");

const THEME = {
  name: "Pauvrete France",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1A1A1A", lt1: "FFFFFF", dk2: "1F2A44", lt2: "F1F3F7",
    accent1: "2A6FC9", // gauche
    accent2: "C8541F", // libéral
    accent3: "12805A", // milieu
    accent4: "C9C8C2", // neutre
    accent5: "5A5955", // texte secondaire
    accent6: "CADCFC", // bleu glacier (texte sur fond sombre)
    hlink: "2A6FC9", folHlink: "6250D6",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 in
pres.title = "Trois débats économiques";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

// ---------------------------------------------------------------------------
// Layouts
// ---------------------------------------------------------------------------
pres.defineSlideMaster({
  title: "TITRE",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.6, w: 11.7, h: 1.3,
      fontFace: THEME.headFontFace, fontSize: 48, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 3.05, w: 11.7, h: 0.8,
      fontFace: THEME.bodyFontFace, fontSize: 24, color: C.accent6, valign: "top", margin: 0 }, text: "" } },
  ],
});

pres.defineSlideMaster({
  title: "CONTENU",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.72, w: 12.1, h: 0.8,
      fontFace: THEME.headFontFace, fontSize: 34, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.3, y: 6.95, w: 0.5, h: 0.3, fontSize: 11, color: C.accent5, align: "right" },
});

// ---------------------------------------------------------------------------
// Aides
// ---------------------------------------------------------------------------
const FONT_CHART = "+mn-lt";
// Virgule décimale quelle que soit la langue de PowerPoint : les valeurs sont saisies x10 (137 = 13,7 %).
const DEC1 = '0","0" %"';

function tag(slide, text, color) {
  slide.addText(text, {
    x: 0.6, y: 0.28, w: Math.max(1.6, 0.11 * text.length + 0.5), h: 0.34,
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.17, fill: { color },
    color: C.background1, fontSize: 12, bold: true, align: "center", valign: "middle",
    charSpacing: 1, margin: 0, isTextBox: true, objectName: "Étiquette",
  });
}

function source(slide, text) {
  slide.addText("Source : " + text, {
    x: 0.6, y: 6.95, w: 11.4, h: 0.3, fontSize: 10, color: C.accent5, margin: 0,
    isTextBox: true, objectName: "Source",
  });
}

function card(slide, x, y, w, h, name) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: C.background2 },
    objectName: name,
  });
}

function bullets(slide, items, opts) {
  slide.addText(
    items.map((it, i) => {
      const runs = Array.isArray(it) ? it : [{ text: it }];
      return runs.map((r, j) => ({
        text: r.text,
        options: {
          bold: !!r.bold, color: r.color || C.text1,
          ...(j === 0 ? { bullet: { indent: 18 }, paraSpaceAfter: 10 } : {}),
          ...(j === runs.length - 1 && i < items.length - 1 ? { breakLine: true } : {}),
        },
      }));
    }).flat(),
    { fontSize: 17, valign: "top", margin: 0, isTextBox: true, ...opts }
  );
}

function auteur(slide, text, y = 1.72) {
  slide.addText(text, {
    x: 0.6, y, w: 12.1, h: 0.36, fontSize: 15, italic: true, color: C.accent5, margin: 0,
    isTextBox: true, objectName: "Auteur",
  });
}

const chartBase = {
  catAxisLabelColor: HEX.accent5, valAxisLabelColor: HEX.accent5,
  catAxisLabelFontFace: FONT_CHART, valAxisLabelFontFace: FONT_CHART,
  dataLabelFontFace: FONT_CHART, titleFontFace: FONT_CHART,
  catAxisLabelFontSize: 14, dataLabelFontSize: 16, dataLabelColor: HEX.dk1, dataLabelFontBold: true,
  valGridLine: { style: "none" }, catGridLine: { style: "none" },
  valAxisHidden: true, catAxisLineShow: true, catAxisLineColor: HEX.accent4,
  showLegend: false, showValue: true, dataLabelPosition: "outEnd",
};

// ---------------------------------------------------------------------------
// 1. Titre
// ---------------------------------------------------------------------------
// ===========================================================================
// TABLEAU DE BORD DE LA CONJONCTURE (données au 19 octobre 2026)
// ===========================================================================
const RED = "B3261E", AMBER = "B26A00", GREEN = "2E7D32";
pres.addSection({ title: "Tableau de bord" });

// --- Diapo 1 : le carré magique, comparé aux références des économistes ---
let s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Tableau de bord" });
tag(s, "TABLEAU DE BORD", C.text2);
s.addText("Le carré magique de l'économie française", { placeholder: "title" });
const quad = [
  { t: "Croissance du PIB", st: "Sous son potentiel", col: RED, big: "0,4 %", sub: "prévue en 2026, pour une croissance potentielle de 1,2 %",
    labels: ["Potentiel", "2025", "2026 (prév.)"], values: [12, 8, 4], fmt: DEC1, max: 15 },
  { t: "Plein emploi (chômage)", st: "Au-dessus du chômage structurel", col: AMBER, big: "8,3 %", sub: "au 2e trimestre 2026, pour un chômage structurel estimé à 7,8 %",
    labels: ["Structurel", "T2 2025", "T2 2026"], values: [78, 76, 83], fmt: DEC1, max: 100 },
  { t: "Stabilité des prix (inflation)", st: "Au-dessus de la cible", col: RED, big: "3,1 %", sub: "sur un an en septembre 2026, pour une cible de la BCE de 2 %",
    labels: ["Cible BCE", "Sept. 2025", "Sept. 2026"], values: [20, 12, 31], fmt: DEC1, max: 36 },
  { t: "Équilibre extérieur (balance courante)", st: "Proche de l'équilibre", col: GREEN, big: "− 10 Md€", sub: "de déficit courant sur 12 mois (fin août 2026), contre − 17 Md€ un an plus tôt",
    labels: ["Équilibre", "1 an avant", "12 mois à août 26"], values: [0, 17, 10], fmt: '[=0]"0";"− "0', max: 22 },
];
quad.forEach((q, i) => {
  const x = 0.6 + (i % 2) * 6.15, y = 1.7 + Math.floor(i / 2) * 2.6;
  card(s, x, y, 5.95, 2.45, "Objectif " + (i + 1));
  s.addText(q.t, { x: x + 0.25, y: y + 0.15, w: 2.7, h: 0.4, fontSize: 14, bold: true, color: C.text2, margin: 0, isTextBox: true });
  s.addText("● " + q.st, { x: x + 0.25, y: y + 0.55, w: 2.7, h: 0.35, fontSize: 12, bold: true, color: q.col, margin: 0, isTextBox: true });
  s.addText(q.big, { x: x + 0.25, y: y + 0.95, w: 2.7, h: 0.7, fontSize: 30, bold: true, color: C.text1, fontFace: THEME.headFontFace, margin: 0, valign: "bottom", isTextBox: true });
  s.addText(q.sub, { x: x + 0.25, y: y + 1.7, w: 2.7, h: 0.65, fontSize: 11, color: C.accent5, margin: 0, valign: "top", isTextBox: true });
  s.addChart(pres.charts.BAR, [{ name: q.t, labels: q.labels, values: q.values }], {
    x: x + 3.0, y: y + 0.1, w: 2.85, h: 2.25, barDir: "col", barGapWidthPct: 45,
    chartColors: [HEX.accent3, HEX.accent4, HEX.dk2],
    valAxisMinVal: 0, valAxisMaxVal: q.max, dataLabelFormatCode: q.fmt,
    ...chartBase, catAxisLabelFontSize: 9, dataLabelFontSize: 12,
    objectName: "Mini-graphique " + q.t,
  });
});
source(s, "Insee ; Banque de France (prévision PIB, chômage structurel 2025, balance courante) ; HCFP (croissance potentielle, avis 2026-5) ; BCE. Barre verte : la référence.");
s.addNotes(
  "Pour juger la situation, je compare chaque objectif du carré magique de Kaldor à sa référence, comme le font les économistes. " +
  "Croissance : la France ne devrait croître que de 0,4 % en 2026, alors que sa croissance potentielle, ce qu'elle peut tenir durablement, est estimée à 1,2 % par le Haut Conseil des finances publiques. " +
  "Emploi : le chômage, à 8,3 %, dépasse le chômage structurel, estimé à 7,8 % par la Banque de France, ce qui suggère une part conjoncturelle. Le niveau structurel reste une estimation, avec une marge d'incertitude. " +
  "Prix : l'inflation, à 3,1 %, est au-dessus de la cible de 2 % de la Banque centrale européenne. " +
  "Seul point positif : la balance courante est presque à l'équilibre, avec un déficit d'environ 10 milliards sur un an, en amélioration."
);

// --- Diapo 2 : le tableau détaillé ---
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Tableau de bord" });
tag(s, "TABLEAU DE BORD", C.text2);
s.addText("Les indicateurs et leurs références", { placeholder: "title" });
const H = (t, al = "center") => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, valign: "middle", align: al } });
const B = (t) => ({ text: t, options: { bold: true, color: C.text2, fill: { color: C.background2 }, colspan: 6, valign: "middle" } });
const L = (t) => ({ text: t, options: { bold: true, color: C.text1, valign: "middle" } });
const V = (t, o = {}) => ({ text: t, options: { color: C.text1, valign: "middle", align: "center", ...o } });
const N = (t) => V(t, { bold: true });
const SIG = (col, t) => ({ text: "● " + t, options: { color: col, bold: true, valign: "middle", align: "center" } });
s.addTable([
  [H("Indicateur", "left"), H("Aujourd'hui"), H("Il y a 1 an"), H("Référence"), H("Prévision"), H("Situation")],
  [B("Les objectifs du carré magique")],
  [L("Croissance du PIB"), N("0,0 % (T2 2026)"), V("0,8 % (2025)"), V("potentiel : 1,2 %"), V("0,4 % (2026) ; 1,0 % (2027, gvt)"), SIG(RED, "sous le potentiel")],
  [L("Chômage"), N("8,3 % (T2)"), V("7,6 %"), V("structurel : 7,8 %"), V("8,6 % fin 2026"), SIG(AMBER, "au-dessus")],
  [L("Inflation (sur un an)"), N("3,1 % (sept.)"), V("1,2 %"), V("cible BCE : 2 %"), V("2,9 % fin 2026"), SIG(RED, "au-dessus")],
  [L("Balance courante (12 mois)"), N("− 10,3 Md€"), V("− 17,2 Md€"), V("équilibre : 0"), V("—"), SIG(GREEN, "proche")],
  [B("Les finances publiques")],
  [L("Déficit public (% du PIB)"), N("5,4 % (prévu)"), V("5,1 %"), V("règle UE : 3 %"), V("5,0 % (2027)"), SIG(RED, "très au-dessus")],
  [L("Dette publique (% du PIB)"), N("119,3 % (prévu)"), V("115,7 %"), V("règle UE : 60 %"), V("121,7 % (2027)"), SIG(RED, "le double")],
  [L("Déficit primaire (hors intérêts)"), N("2,8 pts de PIB (2026)"), V("—"), V("≈ 0 pour stabiliser la dette"), V("2,1 pts (2027)"), SIG(RED, "la dette continue de monter")],
  [L("Charge d'intérêts de la dette"), N("79,2 Md€ (2026)"), V("—"), V("—"), V("91,2 Md€ (2027)"), SIG(RED, "+ 12 Md€")],
  [L("Note de la France (Moody's)"), N("Aa3, perspective négative"), V("—"), V("—"), V("Décision le 23 oct."), SIG(AMBER, "sous surveillance")],
  [B("Les ménages et les marchés")],
  [L("Confiance des ménages"), N("86 (sept.)"), V("88"), V("moyenne : 100"), V("—"), SIG(RED, "basse")],
  [L("Taux d'emprunt à 10 ans (écart avec l'Allemagne)"), N("4,80 % (137 pb)"), V("3,30 % (≈ 76 pb)"), V("Allemagne : ≈ 3,43 %"), V("—"), SIG(RED, "prime de risque élevée")],
], {
  x: 0.6, y: 1.6, w: 12.1, colW: [3.0, 1.95, 1.5, 1.8, 2.25, 1.6],
  rowH: [0.38, 0.28, 0.34, 0.34, 0.34, 0.34, 0.28, 0.34, 0.34, 0.34, 0.34, 0.34, 0.28, 0.34, 0.34],
  fontSize: 11, fontFace: THEME.bodyFontFace, border: { type: "solid", pt: 1, color: HEX.accent4 }, margin: 0.04,
  objectName: "Tableau de bord détaillé",
});
source(s, "Insee ; Banque de France ; HCFP (avis 2026-5) ; PLF 2027 ; AFP (taux et écart du 16/10/2026, Moody's) ; BNP Paribas (taux 2025)");
s.addNotes(
  "Le tableau détaillé compare chaque indicateur à sa référence. " +
  "Les finances publiques sont le point le plus préoccupant : le déficit, à 5,4 % du PIB, dépasse largement la règle européenne de 3 %, et la dette, près de 120 %, est le double de la règle de 60 %. " +
  "Surtout, même sans compter les intérêts, l'État dépense plus qu'il ne gagne : c'est le déficit primaire, 2,8 points de PIB en 2026 selon le Haut Conseil des finances publiques. Tant qu'il ne revient pas vers zéro, la dette continue de monter. " +
  "Les intérêts passeraient de 79 à 91 milliards d'euros l'an prochain, et Moody's réexamine la note de la France le 23 octobre. " +
  "Enfin, les marchés demandent à la France 1,37 point de plus qu'à l'Allemagne pour lui prêter sur dix ans, presque le double d'il y a un an."
);

// --- Diapo 3 : la France face à ses voisins ---
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Tableau de bord" });
tag(s, "TABLEAU DE BORD", C.text2);
s.addText("La France face à ses voisins", { placeholder: "title" });
const FR = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, valign: "middle", align: "center" } });
const FRL = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, valign: "middle" } });
s.addTable([
  [H("", "left"), H("Croissance\n2025"), H("Chômage\naoût 2026"), H("Inflation (IPCH)\nsept. 2026"), H("Déficit\n2025"), H("Dette\n2025")],
  [FRL("France"), FR("0,8 %"), FR("8,2 %"), FR("3,4 %"), FR("5,1 %"), FR("115,6 %")],
  [L("Zone euro"), V("1,5 %"), V("6,4 %"), V("3,8 %"), V("2,9 %"), V("87,8 %")],
  [L("Allemagne"), V("0,3 %"), V("4,0 %"), V("3,3 %"), V("2,7 %"), V("63,5 %")],
  [L("Italie"), V("0,7 %"), V("6,2 %"), V("4,2 %"), V("3,1 %"), V("137,1 %")],
  [L("Espagne"), V("2,8 %"), V("10,0 %"), V("5,0 %"), V("2,4 %"), V("100,7 %")],
  [L("Position de la France"), SIG(RED, "sous la moyenne"), SIG(RED, "au-dessus"), SIG(GREEN, "en dessous"), SIG(RED, "le plus élevé"), SIG(RED, "2e plus élevée")],
], {
  x: 0.6, y: 1.7, w: 12.1, colW: [2.6, 1.9, 1.9, 1.9, 1.9, 1.9],
  rowH: [0.6, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5],
  fontSize: 14, fontFace: THEME.bodyFontFace, border: { type: "solid", pt: 1, color: HEX.accent4 }, margin: 0.05,
  objectName: "Comparaison européenne",
});
card(s, 0.6, 5.55, 12.1, 1.15, "À retenir");
s.addText([
  { text: "À retenir : ", options: { bold: true, color: C.text2 } },
  { text: "l'inflation touche toute la zone euro, et la France est légèrement en dessous de la moyenne. Son déficit est le plus élevé des cinq et sa dette la deuxième : ces données vont dans le sens d'une difficulté propre à la France sur les finances publiques, sans permettre de le démontrer (5 pays, une seule année, sans test statistique).", options: { color: C.text1 } },
], { x: 0.85, y: 5.6, w: 11.6, h: 1.05, fontSize: 15, valign: "middle", margin: 0, isTextBox: true });
source(s, "Eurostat (déficit et dette 2025, 22/04/2026 ; inflation, 16/10/2026 ; chômage, 01/10/2026) ; croissance 2025 : Eurostat (zone euro), Destatis, Istat, INE.");
s.addNotes(
  "Pour savoir si ces difficultés sont propres à la France, comparons-la avec ses voisins. " +
  "L'inflation touche toute la zone euro : 3,8 % en moyenne, et la France est légèrement en dessous, à 3,4 % selon l'indice européen. " +
  "Le chômage français, à 8,2 %, est au-dessus de la moyenne de la zone euro, mais loin de l'Espagne. " +
  "Sur les finances publiques, en revanche, la France se distingue nettement : son déficit, 5,1 % du PIB, est le plus élevé des cinq, et sa dette la deuxième après l'Italie. " +
  "Ces chiffres sont compatibles avec l'hypothèse d'un problème propre à la France, mais avec cinq pays et une seule année, on ne peut pas trancher. " +
  "Ils sont aussi cohérents avec le fait que les marchés demandent à la France un taux plus élevé qu'à l'Allemagne. Attention : l'inflation de cette diapo utilise l'indice européen, l'IPCH, c'est pourquoi elle diffère des 3,1 % de l'Insee."
);

pres.addSection({ title: "Pauvreté · Introduction" });
s = pres.addSlide({ masterName: "TITRE", sectionTitle: "Pauvreté · Introduction" });
s.addText("La pauvreté en France", { placeholder: "title" });
s.addText("Choix politique, échec de l'État ou phénomène complexe ?", { placeholder: "body" });
[
  ["Gauche · L'Essentiel de l'Éco", C.accent1],
  ["Milieu · OFCE", C.accent3],
  ["Libéral · IREF", C.accent2],
].forEach(([t, col], i) => {
  s.addText(t, {
    x: 0.8 + i * 3.75, y: 5.0, w: 3.5, h: 0.6, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.3,
    fill: { color: col }, color: C.background1, fontSize: 16, bold: true, align: "center", valign: "middle",
    margin: 0, isTextBox: true, objectName: "Point de vue " + (i + 1),
  });
});
s.addText("Trois lectures d'un même chiffre de l'Insee", {
  x: 0.8, y: 4.45, w: 11.7, h: 0.4, fontSize: 16, color: C.accent6, margin: 0, isTextBox: true,
});
s.addNotes("Présentation du sujet : la hausse de la pauvreté en France, et trois façons opposées de l'interpréter.");

// ---------------------------------------------------------------------------
// 2. Accroche
// ---------------------------------------------------------------------------
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Pauvreté · Introduction" });
tag(s, "INTRODUCTION", C.text2);
s.addText("Un record depuis 1996", { placeholder: "title" });
s.addChart(pres.charts.BAR, [{ name: "Taux de pauvreté", labels: ["2017", "2022", "2023"], values: [138, 144, 154] }], {
  x: 0.6, y: 1.75, w: 7.3, h: 4.95, barDir: "col", barGapWidthPct: 60,
  chartColors: [HEX.accent4, HEX.accent4, HEX.dk2],
  valAxisMinVal: 0, valAxisMaxVal: 180, dataLabelFormatCode: DEC1,
  showTitle: true, title: "Taux de pauvreté en France (en % de la population)", titleFontSize: 14, titleColor: HEX.accent5,
  ...chartBase,
  objectName: "Graphique taux de pauvreté",
});
[
  ["15,4 %", "de la population sous le seuil de pauvreté en 2023", C.text2, 54],
  ["9,8 millions", "de personnes pauvres", C.text2, 40],
  ["+ 650 000", "pauvres en un an (2022 → 2023)", C.accent2, 40],
].forEach(([big, small, col, size], i) => {
  const y = 1.85 + i * 1.65;
  s.addText(big, { x: 8.5, y, w: 4.2, h: 0.85, fontSize: size, bold: true, color: col,
    fontFace: THEME.headFontFace, margin: 0, valign: "bottom", isTextBox: true });
  s.addText(small, { x: 8.5, y: y + 0.88, w: 4.2, h: 0.6, fontSize: 15, color: C.accent5,
    margin: 0, valign: "top", isTextBox: true });
});
source(s, "Insee (juillet 2025) ; L'Essentiel de l'Éco (08/07/2025)");
s.addNotes(
  "En juillet 2025, l'Insee a publié un chiffre qui a fait beaucoup de bruit : 15,4 % des Français vivent sous le seuil de pauvreté en 2023. " +
  "C'est le niveau le plus élevé depuis le début des séries comparables, en 1996. Cela représente 9,8 millions de personnes, et 650 000 de plus en un seul an. " +
  "Mais ce chiffre ne se lit pas de la même façon selon qui le commente. Je vais vous présenter trois lectures : deux opposées, et une plus nuancée."
);

// ---------------------------------------------------------------------------
// 3. Constat
// ---------------------------------------------------------------------------
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Pauvreté · Introduction" });
tag(s, "LE CONSTAT", C.text2);
s.addText("Qui est pauvre en France ?", { placeholder: "title" });
card(s, 0.6, 1.75, 4.4, 4.95, "Carte définition");
s.addText("Définition de l'Insee", { x: 0.9, y: 2.0, w: 3.8, h: 0.4, fontSize: 17, bold: true, color: C.text2, margin: 0, isTextBox: true });
s.addText("Pauvre = moins de", { x: 0.9, y: 2.55, w: 3.8, h: 0.4, fontSize: 16, color: C.text1, margin: 0, isTextBox: true });
s.addText("60 %", { x: 0.9, y: 2.95, w: 3.8, h: 1.0, fontSize: 60, bold: true, color: C.text2, fontFace: THEME.headFontFace, margin: 0, isTextBox: true });
s.addText("du niveau de vie médian", { x: 0.9, y: 3.95, w: 3.8, h: 0.4, fontSize: 16, color: C.text1, margin: 0, isTextBox: true });
s.addText(
  "Le médian coupe la population en deux moitiés. La pauvreté est donc relative : elle dépend du revenu des autres.",
  { x: 0.9, y: 4.6, w: 3.8, h: 1.8, fontSize: 15, color: C.accent5, margin: 0, valign: "top", isTextBox: true }
);
s.addChart(pres.charts.BAR, [{
  name: "Taux de pauvreté",
  labels: ["Ensemble de la population", "Indépendants", "Familles monoparentales", "Chômeurs", "Autres inactifs (dont étudiants)"],
  values: [154, 192, 343, 361, 373],
}], {
  x: 5.4, y: 1.75, w: 7.3, h: 4.95, barDir: "bar", barGapWidthPct: 55,
  chartColors: [HEX.accent4, HEX.dk2, HEX.dk2, HEX.dk2, HEX.dk2],
  valAxisMinVal: 0, valAxisMaxVal: 450, dataLabelFormatCode: DEC1,
  showTitle: true, title: "Taux de pauvreté par catégorie en 2023 (%)", titleFontSize: 14, titleColor: HEX.accent5,
  ...chartBase,
  objectName: "Graphique catégories",
});
source(s, "Insee (juillet 2025), chiffres cités par L'Essentiel de l'Éco (08/07/2025)");
s.addNotes(
  "D'abord, une définition. Pour l'Insee, on est pauvre quand on vit avec moins de 60 % du niveau de vie médian. " +
  "Le niveau médian, c'est celui qui coupe la population en deux moitiés. C'est donc une pauvreté relative : elle dépend du revenu des autres. Ce point sera important pour la suite. " +
  "Ensuite, la pauvreté ne touche pas tout le monde de la même façon : plus d'un tiers des chômeurs, plus d'un tiers des familles monoparentales, " +
  "et même près d'un indépendant sur cinq. Avoir un travail ne protège donc pas toujours."
);

// ---------------------------------------------------------------------------
// 4. Point de vue 1 : gauche
// ---------------------------------------------------------------------------
pres.addSection({ title: "Pauvreté · Trois points de vue" });
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Pauvreté · Trois points de vue" });
tag(s, "POINT DE VUE 1 · GAUCHE", C.accent1);
s.addText("« Le résultat de choix politiques »", { placeholder: "title" });
auteur(s, "Paul Crauchet, L'Essentiel de l'Éco, 8 juillet 2025");
bullets(s, [
  [{ text: "Depuis 2017 : ", bold: true }, { text: "13,8 % → 15,4 %, soit 1,2 million de pauvres en plus" }],
  [{ text: "Aides anti-inflation ", bold: true }, { text: "de 2022 non reconduites" }],
  [{ text: "Aides au logement ", bold: true }, { text: "revalorisées moins vite que les prix" }],
  [{ text: "Conclusion : ", bold: true }, { text: "un « projet » politique, pas un accident" }],
], { x: 0.6, y: 2.3, w: 5.6, h: 4.3 });

// Gini
card(s, 6.7, 1.75, 6.0, 2.35, "Carte Gini");
s.addText("Les inégalités se creusent : indice de Gini", { x: 7.0, y: 1.92, w: 5.5, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true });
const gx = (v) => 7.1 + ((v - 0.25) / 0.07) * 5.2;
s.addShape(pres.shapes.RECTANGLE, { x: 7.1, y: 3.27, w: 5.2, h: 0.06, fill: { color: C.accent4 }, line: { color: C.accent4 } });
s.addShape(pres.shapes.LINE, { x: gx(0.281) + 0.18, y: 3.3, w: gx(0.297) - gx(0.281) - 0.36, h: 0,
  line: { color: C.accent2, width: 3, endArrowType: "triangle" } });
[[0.281, "2017 · 0,281", C.accent1], [0.297, "2023 · 0,297", C.accent2]].forEach(([v, lab, col]) => {
  s.addShape(pres.shapes.OVAL, { x: gx(v) - 0.17, y: 3.13, w: 0.34, h: 0.34, fill: { color: col }, line: { color: C.background1, width: 2 } });
  const left = v < 0.29;
  s.addText(lab, { x: left ? gx(v) - 2.2 : gx(v) - 0.2, y: 2.55, w: 2.4, h: 0.45, fontSize: 18, bold: true, color: C.text1,
    align: left ? "right" : "left", margin: 0, isTextBox: true });
});
s.addText("0 = égalité parfaite, 1 = une seule personne a tout (échelle zoomée)", { x: 7.0, y: 3.62, w: 5.5, h: 0.3, fontSize: 12, color: C.accent5, margin: 0, isTextBox: true });

// Sous-estimation
card(s, 6.7, 4.35, 6.0, 2.35, "Carte sous-estimation");
s.addText("Un chiffre officiel sous-estimé ?", { x: 7.0, y: 4.52, w: 5.5, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true });
const sc = 4.3 / 12;
s.addText("9,8 M comptés", { x: 7.0, y: 5.1, w: 9.8 * sc, h: 0.6, fill: { color: C.accent1 }, color: C.background1,
  fontSize: 15, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
s.addText("+ 2,1", { x: 7.0 + 9.8 * sc + 0.03, y: 5.1, w: 2.1 * sc, h: 0.6, fill: { color: C.accent2 }, color: C.background1,
  fontSize: 14, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
s.addText("≈ 12 M", { x: 11.45, y: 5.1, w: 1.1, h: 0.6, fontSize: 20, bold: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
s.addText("Non comptés : sans-abri, foyers, étudiants précaires, outre-mer", { x: 7.0, y: 5.9, w: 5.5, h: 0.55, fontSize: 13, color: C.accent5, margin: 0, valign: "top", isTextBox: true });
source(s, "Paul Crauchet, L'Essentiel de l'Éco, 08/07/2025 ; Insee (Gini). Les 2,1 M hors champ : estimation Insee pour 2021.");
s.addNotes(
  "Pour Paul Crauchet, ce n'est pas un accident, c'est un « projet ». Depuis 2017, le taux de pauvreté est passé de 13,8 % à 15,4 %, soit 1,2 million de pauvres en plus. " +
  "Ses arguments : les aides exceptionnelles contre l'inflation de 2022 n'ont pas été reconduites ; les aides au logement ont augmenté moins vite que les prix ; " +
  "pendant ce temps, les inégalités se creusent : l'indice de Gini passe de 0,281 à 0,297. Plus il est proche de 1, plus la société est inégalitaire. " +
  "Il ajoute que le chiffre officiel sous-estime la réalité : les sans-abri, les étudiants précaires et l'outre-mer ne sont pas comptés. En les ajoutant, on approcherait 12 millions de pauvres. " +
  "En résumé : l'État a fait le choix de moins protéger les plus modestes."
);

// ---------------------------------------------------------------------------
// 5. Point de vue 2 : libéral (pauvreté relative)
// ---------------------------------------------------------------------------
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Pauvreté · Trois points de vue" });
tag(s, "POINT DE VUE 2 · LIBÉRAL", C.accent2);
s.addText("« La pauvreté relative est trompeuse »", { placeholder: "title" });
auteur(s, "Élodie Messéant, IREF (repris de L'Opinion), 26 août 2025");
bullets(s, [
  [{ text: "Le seuil suit le revenu médian : ", bold: true }, { text: "si tout le monde s'enrichit, le seuil monte aussi" }],
  [{ text: "Un pauvre d'aujourd'hui ", bold: true }, { text: "vit bien mieux qu'un pauvre d'il y a deux siècles" }],
  [{ text: "Le bon indicateur : ", bold: true }, { text: "la pauvreté extrême, qui recule fortement grâce à la liberté économique et à la mondialisation" }],
  [{ text: "Conclusion : ", bold: true }, { text: "le chiffre mesure surtout les inégalités, pas la misère" }],
], { x: 0.6, y: 2.3, w: 5.4, h: 4.3 });

s.addChart(pres.charts.BAR, [{ name: "Pauvreté extrême", labels: ["1990", "2015"], values: [36, 10] }], {
  x: 6.4, y: 1.75, w: 6.3, h: 4.95, barDir: "col", barGapWidthPct: 70,
  chartColors: [HEX.accent4, HEX.accent2],
  valAxisMinVal: 0, valAxisMaxVal: 42, dataLabelFormatCode: '0" %"',
  showTitle: true, title: "Population mondiale vivant avec moins de 1,90 $ par jour (%)", titleFontSize: 14, titleColor: HEX.accent5,
  ...chartBase, dataLabelFontSize: 22,
  objectName: "Graphique pauvreté extrême",
});
source(s, "IREF (26/08/2025) ; graphique : Banque mondiale, Poverty and Shared Prosperity 2018 (seuil de 1,90 $ par jour)");
s.addNotes(
  "À l'opposé, Élodie Messéant, de l'IREF, un think tank libéral, conteste d'abord la mesure. Comme le seuil est relatif, un pauvre d'aujourd'hui vit beaucoup mieux qu'un pauvre d'il y a deux siècles, " +
  "mais il reste « pauvre » dans les statistiques. Selon elle, ce chiffre mesure surtout les inégalités, pas la misère. " +
  "Pour elle, le bon indicateur est la pauvreté extrême : dans le monde, elle est passée de 36 % de la population en 1990 à 10 % en 2015, selon la Banque mondiale, grâce à la liberté économique et à la mondialisation, et non à la redistribution."
);
/* ancien exemple fictif supprimé
card(s, 6.4, 1.75, 6.3, 4.95, "Carte exemple fictif");
s.addText("Exemple fictif : un ménage modeste dans deux pays", { x: 6.7, y: 1.9, w: 5.8, h: 0.4, fontSize: 15, bold: true, color: C.text2, margin: 0, isTextBox: true });
s.addShape(pres.shapes.RECTANGLE, { x: 6.7, y: 2.5, w: 0.45, h: 0.06, fill: { color: C.accent2 }, line: { color: C.accent2 } });
s.addText("seuil de pauvreté du pays", { x: 7.25, y: 2.38, w: 3.5, h: 0.3, fontSize: 13, color: C.accent2, bold: true, margin: 0, isTextBox: true });
const base = 5.75, k = 0.0017;
const eur = (n) => n.toLocaleString("fr-FR").replace(/[  ]/g, " ") + " €";
[
  { x: 7.3, rev: 500, seuil: 420, verdict: "Pas pauvre", lab: "Pays A : pauvre mais égalitaire" },
  { x: 10.1, rev: 1200, seuil: 1500, verdict: "« Pauvre »", lab: "Pays B : riche mais inégalitaire" },
].forEach((p) => {
  const bw = 1.5;
  s.addText(eur(p.rev), {
    x: p.x, y: base - p.rev * k, w: bw, h: p.rev * k, fill: { color: C.accent1 }, color: C.background1,
    fontSize: 18, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  const ys = base - p.seuil * k;
  s.addShape(pres.shapes.RECTANGLE, { x: p.x - 0.3, y: ys - 0.03, w: bw + 0.6, h: 0.06, fill: { color: C.accent2 }, line: { color: C.accent2 } });
  const top = base - Math.max(p.rev, p.seuil) * k;
  s.addText(p.verdict, { x: p.x - 0.3, y: top - 0.5, w: bw + 0.6, h: 0.4, fontSize: 18, bold: true, color: C.text1, align: "center", margin: 0, isTextBox: true });
  s.addText([
    { text: p.lab, options: { breakLine: true } },
    { text: "seuil : " + eur(p.seuil), options: { color: C.accent2, bold: true } },
  ], { x: p.x - 0.65, y: base + 0.1, w: bw + 1.3, h: 0.7, fontSize: 13, color: C.text1, align: "center", margin: 0, isTextBox: true });
});
s.addShape(pres.shapes.LINE, { x: 6.8, y: base, w: 5.5, h: 0, line: { color: C.accent4, width: 1 } });
source(s, "chiffres inventés pour illustrer l'argument de l'IREF. Seuil = 60 % du revenu médian du pays.");
s.addNotes(
  "À l'opposé, Élodie Messéant, de l'IREF, un think tank libéral, conteste d'abord la mesure. Comme le seuil est relatif, un pauvre d'aujourd'hui vit beaucoup mieux qu'un pauvre d'il y a deux siècles, " +
  "mais il reste « pauvre » dans les statistiques. Selon elle, ce chiffre mesure surtout les inégalités, pas la misère. " +
  "Exemple fictif à droite : le ménage du pays B gagne 2,4 fois plus que celui du pays A, mais c'est lui qui est compté pauvre. Les chiffres sont inventés, c'est juste pour illustrer l'argument."
);
*/

// ---------------------------------------------------------------------------
// 6. Point de vue 2 : libéral (redistribution)
// ---------------------------------------------------------------------------
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Pauvreté · Trois points de vue" });
tag(s, "POINT DE VUE 2 · LIBÉRAL", C.accent2);
s.addText("« La France redistribue déjà énormément »", { placeholder: "title" });
s.addChart(pres.charts.BAR, [{ name: "Écart de revenus", labels: ["Avant impôts et aides", "Après impôts et aides"], values: [18, 3] }], {
  x: 0.6, y: 1.75, w: 6.3, h: 4.95, barDir: "col", barGapWidthPct: 70,
  chartColors: [HEX.accent4, HEX.accent2],
  valAxisMinVal: 0, valAxisMaxVal: 21, dataLabelFormatCode: '"× "0',
  showTitle: true, title: "Les plus aisés gagnent combien de fois plus que les plus modestes ?", titleFontSize: 14, titleColor: HEX.accent5,
  ...chartBase, dataLabelFontSize: 22,
  objectName: "Graphique redistribution",
});
s.addText("45,6 %", { x: 7.4, y: 1.8, w: 5.3, h: 0.95, fontSize: 54, bold: true, color: C.accent2, fontFace: THEME.headFontFace, margin: 0, valign: "bottom", isTextBox: true });
s.addText("de prélèvements obligatoires en France, contre environ 40 % dans l'UE", { x: 7.4, y: 2.8, w: 5.3, h: 0.7, fontSize: 15, color: C.accent5, margin: 0, valign: "top", isTextBox: true });
bullets(s, [
  [{ text: "Effets pervers : ", bold: true }, { text: "la redistribution crée de la dépendance et décourage le travail" }],
  [{ text: "Solution : ", bold: true }, { text: "liberté économique, emploi, investissement privé" }],
  [{ text: "Exemple cité : ", bold: true }, { text: "l'Argentine de Javier Milei" }],
], { x: 7.4, y: 3.85, w: 5.3, h: 2.8 });
source(s, "Élodie Messéant, « Hausse du taux de pauvreté : le socialisme ne fonctionne pas », IREF, 26/08/2025");
s.addNotes(
  "Elle rappelle ensuite que la France redistribue déjà énormément : 45,6 % de prélèvements obligatoires, contre environ 40 % dans l'Union européenne. " +
  "Avant impôts et aides, les plus aisés gagnent 18 fois plus que les plus modestes ; après, l'écart tombe à 3 fois. " +
  "Pour elle, plus de redistribution n'est pas la solution : elle crée de la dépendance et décourage le travail. Il faut au contraire plus de liberté économique, d'emploi et d'investissement. " +
  "Elle cite l'exemple de l'Argentine de Javier Milei. En résumé : le problème, ce n'est pas le manque d'État, c'est le trop d'État."
);

// ---------------------------------------------------------------------------
// 7. Point de vue 3 : milieu (OFCE)
// ---------------------------------------------------------------------------
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Pauvreté · Trois points de vue" });
tag(s, "POINT DE VUE 3 · MILIEU", C.accent3);
s.addText("« Il faut séparer les causes »", { placeholder: "title" });
auteur(s, "Pierre Madec, OFCE – Sciences Po, 27 mars 2026");
const causes = [
  ["Le seuil", "monte avec le revenu médian"],
  ["La démographie", "ex. plus de familles monoparentales"],
  ["Les aides sociales", "sont-elles assez efficaces ?"],
  ["Les revenus du travail", "emploi, salaires, précarité"],
];
const cw = 2.85, gap = 0.233, cy = 2.3, ch = 1.85;
const target = { x: 6.665, y: 4.8 };
causes.forEach(([t, d], i) => {
  const x = 0.6 + i * (cw + gap);
  card(s, x, cy, cw, ch, "Cause " + (i + 1));
  s.addText(String(i + 1), { x: x + 0.22, y: cy + 0.22, w: 0.46, h: 0.46, shape: pres.shapes.OVAL, fill: { color: C.accent3 },
    color: C.background1, fontSize: 16, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addText(t, { x: x + 0.22, y: cy + 0.8, w: cw - 0.44, h: 0.4, fontSize: 17, bold: true, color: C.text2, margin: 0, isTextBox: true });
  s.addText(d, { x: x + 0.22, y: cy + 1.2, w: cw - 0.44, h: 0.55, fontSize: 14, color: C.accent5, margin: 0, valign: "top", isTextBox: true });
  const sx = x + cw / 2, sy = cy + ch + 0.05;
  const dx = target.x - sx;
  s.addShape(pres.shapes.LINE, {
    x: Math.min(sx, target.x), y: sy, w: Math.max(Math.abs(dx), 0.01), h: target.y - sy - 0.05,
    flipH: dx < 0, line: { color: C.accent5, width: 1.5, endArrowType: "triangle" },
  });
});
s.addText("Taux de pauvreté : 15,4 %", { x: 4.165, y: target.y, w: 5.0, h: 0.85, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1,
  fill: { color: C.accent3 }, color: C.background1, fontSize: 22, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true,
  objectName: "Résultat" });
s.addText([
  { text: "+ regarder le « reste à vivre » : ", options: { bold: true, color: C.text2 } },
  { text: "ce qui reste une fois le loyer, l'énergie et les autres dépenses obligatoires payés", options: { color: C.accent5 } },
], { x: 0.6, y: 6.0, w: 12.1, h: 0.5, fontSize: 15, italic: true, align: "center", margin: 0, isTextBox: true });
source(s, "Pierre Madec, « Pauvreté monétaire en France : les causes d'un décrochage inquiétant », OFCE, 27/03/2026");
s.addNotes(
  "Entre les deux, l'économiste Pierre Madec part d'un paradoxe : la France a l'une des protections sociales les plus développées d'Europe, et pourtant la pauvreté augmente. " +
  "Plutôt que de chercher un coupable, il décompose la hausse en quatre effets : le seuil, qui monte avec le revenu médian ; la démographie, par exemple plus de familles monoparentales ; " +
  "l'efficacité des aides sociales ; et les revenus du travail. Il propose aussi de regarder le « reste à vivre », c'est-à-dire ce qu'il reste une fois le loyer, l'énergie et les autres dépenses obligatoires payés. " +
  "En résumé : plusieurs causes se mélangent, donc il n'y a pas de solution unique."
);

// ---------------------------------------------------------------------------
// 8. Synthèse
// ---------------------------------------------------------------------------
pres.addSection({ title: "Pauvreté · Conclusion" });
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Pauvreté · Conclusion" });
tag(s, "CONCLUSION", C.text2);
s.addText("Trois lectures d'un même chiffre", { placeholder: "title" });
const head = (t, col) => ({ text: t, options: { fill: { color: col }, color: C.background1, bold: true, align: "center", valign: "middle" } });
const rowLab = (t) => ({ text: t, options: { bold: true, color: C.text2, valign: "middle" } });
const cell = (t) => ({ text: t, options: { color: C.text1, valign: "middle" } });
s.addTable([
  [{ text: "", options: {} }, head("Gauche (L'Essentiel de l'Éco)", C.accent1), head("Milieu (OFCE)", C.accent3), head("Libéral (IREF)", C.accent2)],
  [rowLab("Diagnostic"), cell("Des choix politiques depuis 2017"), cell("Plusieurs causes mêlées"), cell("Trop d'État")],
  [rowLab("Les aides"), cell("Coupées ou mal revalorisées"), cell("Efficacité à évaluer"), cell("Créent dépendance, découragent le travail")],
  [rowLab("Solution"), cell("Protéger davantage les plus modestes"), cell("Analyser avant d'agir"), cell("Liberté économique, emploi")],
  [rowLab("La mesure"), cell("Sous-estime (≈ 12 millions)"), cell("Compléter par le « reste à vivre »"), cell("Exagère : mesure les inégalités")],
], {
  x: 0.6, y: 1.75, w: 12.1, colW: [1.9, 3.4, 3.4, 3.4], rowH: [0.6, 0.72, 0.72, 0.72, 0.72],
  fontSize: 15, fontFace: THEME.bodyFontFace, border: { type: "solid", pt: 1, color: HEX.accent4 }, margin: 0.08,
  objectName: "Tableau synthèse",
});
s.addText("Pour faire baisser la pauvreté : plus d'aides, des aides différentes, ou moins d'aides ?", {
  x: 0.6, y: 5.65, w: 12.1, h: 1.0, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, fill: { color: C.text2 },
  color: C.background1, fontSize: 21, bold: true, align: "center", valign: "middle", margin: 0.1, isTextBox: true,
  objectName: "Question",
});
s.addNotes(
  "Les trois textes partent du même chiffre, mais ils s'opposent sur deux points : la mesure — pour l'IREF, le chiffre exagère la pauvreté ; pour L'Essentiel de l'Éco, il la sous-estime — " +
  "et la cause : des choix politiques, trop d'État, ou plusieurs facteurs mêlés. " +
  "Je vous laisse avec cette question : pour faire baisser la pauvreté, faut-il plus d'aides, des aides différentes, ou moins d'aides ?"
);

// ---------------------------------------------------------------------------
// 9. Sources
// ---------------------------------------------------------------------------
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Pauvreté · Conclusion" });
tag(s, "SOURCES", C.text2);
s.addText("Sources", { placeholder: "title" });
const srcs = [
  ["Paul Crauchet", "La face cachée du modèle Macron révélée par l'Insee", "L'Essentiel de l'Éco, 8 juillet 2025",
    "https://lessentieldeleco.fr/2656-la-face-cachee-du-modele-macron-revelee-par-linsee/", C.accent1],
  ["Élodie Messéant", "Hausse du taux de pauvreté : le socialisme ne fonctionne pas", "IREF / L'Opinion, 26 août 2025",
    "https://fr.irefeurope.org/hausse-taux-pauvrete-france/", C.accent2],
  ["Pierre Madec", "Pauvreté monétaire en France : les causes d'un décrochage inquiétant", "OFCE – Sciences Po, 27 mars 2026",
    "https://www.ofce.fr/blog2024/fr/2026/20260327_PIM/index.html", C.accent3],
  ["Insee", "L'essentiel sur… la pauvreté", "Chiffres de référence",
    "https://www.insee.fr/fr/statistiques/5759045", C.text2],
];
srcs.forEach(([who, t, where, url, col], i) => {
  const y = 1.8 + i * 1.22;
  s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.12, w: 0.4, h: 0.4, fill: { color: col }, line: { color: col } });
  s.addText([
    { text: who + " — ", options: { bold: true, color: C.text1 } },
    { text: "« " + t + " »", options: { color: C.text1, hyperlink: { url } } },
    { text: where, options: { color: C.accent5, fontSize: 14, breakLine: false } },
  ].map((r, j) => (j === 1 ? { ...r, options: { ...r.options, breakLine: true } } : r)), {
    x: 1.25, y, w: 11.4, h: 0.95, fontSize: 17, valign: "top", margin: 0, isTextBox: true,
  });
});

// ===========================================================================
// SUJET 2 : LA TAXE ZUCMAN
// ===========================================================================
const SRC_LCP = "R. Marchal, « Gabriel Zucman dénonce un projet de budget qui protège les milliardaires », LCP, 02/10/2026";
const SRC_FONDAPOL = "Victor Fouquet, « L'impasse de la taxe Zucman », Fondapol, juin 2026";
const SRC_OFCE2 = "G. Allègre et X. Timbeau, « Pourquoi et comment taxer les plus riches ? », OFCE, avril 2026";

// Barre horizontale dessinée (valeur affichée à droite)
function hbar(slide, { x, y, w, h, val, max, color, label, valueText, labelW = 2.4 }) {
  slide.addText(label, { x, y, w: labelW, h, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  const bw = (w - labelW - 1.3) * (val / max);
  slide.addShape(pres.shapes.RECTANGLE, { x: x + labelW + 0.1, y: y + h * 0.15, w: bw, h: h * 0.7, fill: { color }, line: { color } });
  slide.addText(valueText, { x: x + labelW + 0.2 + bw, y, w: 1.3, h, fontSize: 18, bold: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
}

// Barre verticale dessinée (valeur affichée au-dessus, libellé dessous)
function vbar(slide, { x, base, w, hMax, val, max, color, valueText, label }) {
  const h = hMax * (val / max);
  slide.addShape(pres.shapes.RECTANGLE, { x, y: base - h, w, h, fill: { color }, line: { color } });
  slide.addText(valueText, { x: x - 0.4, y: base - h - 0.5, w: w + 0.8, h: 0.45, fontSize: 22, bold: true, color: C.text1, align: "center", margin: 0, isTextBox: true });
  slide.addText(label, { x: x - 0.5, y: base + 0.08, w: w + 1.0, h: 0.6, fontSize: 13, color: C.text1, align: "center", margin: 0, isTextBox: true });
}

function statCard(slide, { x, y, w, h, big, small, color }) {
  card(slide, x, y, w, h, "Chiffre clé");
  slide.addText(big, { x: x + 0.3, y: y + 0.2, w: w - 0.6, h: 0.8, fontSize: 40, bold: true, color, fontFace: THEME.headFontFace, margin: 0, valign: "bottom", isTextBox: true });
  slide.addText(small, { x: x + 0.3, y: y + 1.05, w: w - 0.6, h: h - 1.2, fontSize: 14, color: C.accent5, margin: 0, valign: "top", isTextBox: true });
}

// --- Titre du sujet 2 ---
pres.addSection({ title: "Taxe Zucman · Introduction" });
s = pres.addSlide({ masterName: "TITRE", sectionTitle: "Taxe Zucman · Introduction" });
s.addText("La taxe Zucman", { placeholder: "title" });
s.addText("Faut-il taxer davantage les ultra-riches ?", { placeholder: "body" });
s.addText("Trois positions sur l'impôt minimum de 2 % sur les très grandes fortunes", {
  x: 0.8, y: 4.45, w: 11.7, h: 0.4, fontSize: 16, color: C.accent6, margin: 0, isTextBox: true,
});
[
  ["Pour · Zucman, Oxfam", C.accent1],
  ["Milieu · OFCE", C.accent3],
  ["Contre · Fondapol, IFRAP", C.accent2],
].forEach(([t, col], i) => {
  s.addText(t, {
    x: 0.8 + i * 3.75, y: 5.0, w: 3.5, h: 0.6, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.3,
    fill: { color: col }, color: C.background1, fontSize: 16, bold: true, align: "center", valign: "middle",
    margin: 0, isTextBox: true, objectName: "Position " + (i + 1),
  });
});
s.addNotes("Deuxième sujet : la taxe Zucman, et trois positions sur la façon de taxer les très grandes fortunes.");

// --- Qu'est-ce que la taxe Zucman ? ---
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Taxe Zucman · Introduction" });
tag(s, "INTRODUCTION", C.text2);
s.addText("Qu'est-ce que la taxe Zucman ?", { placeholder: "title" });
[
  ["2 %", "impôt minimum sur le patrimoine net"],
  ["100 M€", "seuil de patrimoine à partir duquel la taxe s'applique"],
  ["≈ 1 800", "foyers fiscaux concernés en France"],
].forEach(([big, small], i) => {
  statCard(s, { x: 0.6 + i * 4.1, y: 1.75, w: 3.9, h: 2.1, big, small, color: C.text2 });
});
card(s, 0.6, 4.1, 12.1, 1.55, "Carte calcul");
s.addText("Comment ça marche ?", { x: 0.9, y: 4.25, w: 4, h: 0.35, fontSize: 15, bold: true, color: C.text2, margin: 0, isTextBox: true });
const etapes = ["2 % du patrimoine", "impôts déjà payés", "complément à payer (si positif)"];
etapes.forEach((t, i) => {
  s.addText(t, { x: 0.9 + i * 4.05, y: 4.75, w: 3.35, h: 0.65, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08,
    fill: { color: C.background1 }, line: { color: C.accent4 }, color: C.text1, fontSize: 16, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
  if (i < 2) s.addText(i === 0 ? "−" : "=", { x: 4.25 + i * 4.05, y: 4.75, w: 0.7, h: 0.65, fontSize: 28, bold: true, color: C.text2, align: "center", valign: "middle", margin: 0, isTextBox: true });
});
s.addText([
  { text: "Pourquoi ? ", options: { bold: true, color: C.text2 } },
  { text: "Selon Zucman, à cause de l'optimisation fiscale (revenus laissés dans des holdings), les plus riches paient proportionnellement moins d'impôts.", options: { color: C.text1 } },
], { x: 0.6, y: 5.85, w: 12.1, h: 0.5, fontSize: 16, margin: 0, isTextBox: true });
s.addText([
  { text: "16 octobre 2026 : ", options: { bold: true, color: C.text2 } },
  { text: "l'Assemblée nationale rejette la taxe pendant l'examen du budget 2027.", options: { color: C.text1 } },
], { x: 0.6, y: 6.38, w: 12.1, h: 0.4, fontSize: 16, margin: 0, isTextBox: true });
source(s, "France 24 / AFP (16/10/2026) ; LCP (02/10/2026)");
s.addNotes(
  "Le 16 octobre 2026, en plein débat sur le budget 2027, l'Assemblée nationale a rejeté la taxe Zucman. Elle porte le nom de l'économiste Gabriel Zucman, qui l'a proposée. " +
  "Le principe est simple : un impôt minimum de 2 % sur les patrimoines de plus de 100 millions d'euros. Cela concerne environ 1 800 foyers. " +
  "On calcule 2 % de leur patrimoine, on retire les impôts qu'ils ont déjà payés, et s'il manque quelque chose, ils paient la différence. " +
  "Pourquoi cette taxe ? Selon Zucman, les plus riches paient proportionnellement moins d'impôts à cause de l'optimisation fiscale : leurs revenus restent dans des sociétés, des holdings, et ne sont presque pas taxés. La taxe sert de filet de sécurité. " +
  "C'est devenu l'un des débats les plus vifs avant la présidentielle de 2027. Je vais vous présenter trois positions : pour, contre, et entre les deux."
);

// --- Pour ---
pres.addSection({ title: "Taxe Zucman · Trois positions" });
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Taxe Zucman · Trois positions" });
tag(s, "POSITION 1 · POUR", C.accent1);
s.addText("« Les milliardaires paient moins que les autres »", { placeholder: "title" });
auteur(s, "Gabriel Zucman (LCP) et Oxfam (Reporterre), octobre 2026");
bullets(s, [
  [{ text: "« Un problème de recettes », ", bold: true }, { text: "pas de dépenses" }],
  [{ text: "Les milliardaires ", bold: true }, { text: "paient moins que les classes moyennes ; les grandes entreprises moins que les PME" }],
  [{ text: "Le budget 2027 ", bold: true }, { text: "fait payer les retraités (≈ 5 Md€) plutôt que les grandes fortunes" }],
  [{ text: "Oxfam ", bold: true }, { text: "attaque l'État en justice pour obtenir la taxe" }],
  [{ text: "Soutenue par 7 prix Nobel ", bold: true }, { text: "d'économie" }],
], { x: 0.6, y: 2.3, w: 5.6, h: 4.3 });
card(s, 6.7, 1.75, 6.0, 2.75, "Carte taux Oxfam");
s.addText("Part des revenus payée en impôts, selon Oxfam", { x: 7.0, y: 1.92, w: 5.5, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true });
hbar(s, { x: 7.0, y: 2.5, w: 5.5, h: 0.75, val: 25, max: 50, color: C.accent1, label: "Les 0,0002 % les plus riches", valueText: "≈ 25 %" });
hbar(s, { x: 7.0, y: 3.35, w: 5.5, h: 0.75, val: 50, max: 50, color: C.accent4, label: "La plupart des Français", valueText: "> 50 %" });
statCard(s, { x: 6.7, y: 4.75, w: 6.0, h: 1.95, big: "20 Md€ / an", small: "le rendement annoncé par Gabriel Zucman", color: C.accent1 });
source(s, SRC_LCP + " ; Reporterre (13/10/2026), calculs d'Oxfam");
s.addNotes(
  "Pour Gabriel Zucman, la France n'a pas un problème de dépenses, elle a « un problème de recettes ». Selon lui, les milliardaires paient proportionnellement moins d'impôts que les classes moyennes, " +
  "et les grandes entreprises moins que les PME. L'ONG Oxfam va dans le même sens : selon ses calculs, les plus riches paient environ 25 % de leurs revenus en impôts, contre plus de 50 % pour la plupart des Français. " +
  "Oxfam a même engagé une action en justice contre l'État, et la taxe est soutenue par sept prix Nobel d'économie. Zucman estime que sa taxe rapporterait 20 milliards d'euros par an. Pour lui, c'est plus juste que de faire payer les retraités, à qui le budget 2027 retire environ 5 milliards. " +
  "En résumé : c'est une question de justice fiscale. Attention : les chiffres d'Oxfam sont leurs propres calculs, la méthode n'est pas publiée dans l'article."
);

// --- Contre (1/2) ---
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Taxe Zucman · Trois positions" });
tag(s, "POSITION 2 · CONTRE", C.accent2);
s.addText("« Un diagnostic faux »", { placeholder: "title" });
auteur(s, "Victor Fouquet, Fondapol (think tank libéral), juin 2026");
bullets(s, [
  [{ text: "Le calcul de Zucman est biaisé : ", bold: true }, { text: "il compte les aides sociales d'un côté du calcul mais pas de l'autre" }],
  [{ text: "Le système est déjà très redistributif ", bold: true }, { text: "grâce aux impôts et aux aides" }],
  [{ text: "Les dividendes sont déjà taxés ", bold: true }, { text: "jusqu'à 60,8 % avec l'impôt sur les sociétés" }],
], { x: 0.6, y: 2.3, w: 5.4, h: 4.3 });
card(s, 6.4, 1.75, 6.3, 4.95, "Carte redistribution");
s.addText("Écart de niveau de vie entre les 10 % les plus aisés et les 10 % les plus modestes", { x: 6.7, y: 1.9, w: 5.7, h: 0.65, fontSize: 15, bold: true, color: C.text2, margin: 0, isTextBox: true });
vbar(s, { x: 7.35, base: 5.75, w: 1.5, hMax: 2.5, val: 26, max: 26, color: C.accent4, valueText: "× 26", label: "Avant impôts\net aides" });
vbar(s, { x: 10.15, base: 5.75, w: 1.5, hMax: 2.5, val: 3.5, max: 26, color: C.accent2, valueText: "× 3,5", label: "Après impôts\net aides" });
s.addShape(pres.shapes.LINE, { x: 6.8, y: 5.75, w: 5.5, h: 0, line: { color: C.accent4, width: 1 } });
source(s, SRC_FONDAPOL + " (données Insee, redistribution élargie)");
s.addNotes(
  "Victor Fouquet, pour la Fondapol, un think tank libéral, conteste d'abord le diagnostic. Le système français serait déjà très redistributif : " +
  "grâce aux impôts et aux aides, l'écart de niveau de vie entre les 10 % les plus aisés et les 10 % les plus modestes passe de 26 à 3,5. " +
  "Et les dividendes sont déjà taxés jusqu'à 60,8 % si l'on compte l'impôt sur les sociétés."
);

// --- Contre (2/2) ---
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Taxe Zucman · Trois positions" });
tag(s, "POSITION 2 · CONTRE", C.accent2);
s.addText("« Peu de recettes, beaucoup de risques »", { placeholder: "title" });
card(s, 0.6, 1.75, 6.6, 2.6, "Carte rendement");
s.addText("Combien rapporterait la taxe chaque année ?", { x: 0.9, y: 1.92, w: 6.0, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true });
hbar(s, { x: 0.9, y: 2.42, w: 6.1, h: 0.6, val: 20, max: 20, color: C.accent1, label: "Gabriel Zucman", valueText: "20 Md€", labelW: 2.6 });
hbar(s, { x: 0.9, y: 3.02, w: 6.1, h: 0.6, val: 5, max: 20, color: C.accent4, label: "7 économistes (Le Monde)", valueText: "5 Md€", labelW: 2.6 });
hbar(s, { x: 0.9, y: 3.62, w: 6.1, h: 0.6, val: 2.5, max: 20, color: C.accent2, label: "Sénat (IS déduit)", valueText: "2 à 3 Md€", labelW: 2.6 });
statCard(s, { x: 0.6, y: 4.6, w: 6.6, h: 2.1, big: "÷ 2", small: "le rendement de la taxe si seulement les 3 plus gros contribuables quittent la France (Fondapol)", color: C.accent2 });
card(s, 7.5, 1.75, 5.2, 4.95, "Carte constitution");
s.addText("Un risque constitutionnel", { x: 7.8, y: 1.92, w: 4.6, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true });
s.addText("Taux sur la fortune admis par le Conseil constitutionnel sans plafond", { x: 7.8, y: 2.5, w: 4.6, h: 0.6, fontSize: 14, color: C.accent5, margin: 0, isTextBox: true });
s.addText("0,5 %", { x: 7.8, y: 3.05, w: 4.6, h: 0.8, fontSize: 40, bold: true, color: C.text2, fontFace: THEME.headFontFace, margin: 0, isTextBox: true });
s.addText("Taux de la taxe Zucman, sans plafond", { x: 7.8, y: 4.05, w: 4.6, h: 0.4, fontSize: 14, color: C.accent5, margin: 0, isTextBox: true });
s.addText("2 %", { x: 7.8, y: 4.4, w: 4.6, h: 0.8, fontSize: 40, bold: true, color: C.accent2, fontFace: THEME.headFontFace, margin: 0, isTextBox: true });
s.addText("Selon la Fondapol et l'IFRAP, la taxe serait donc probablement censurée. Ce n'est pas une décision du Conseil.", { x: 7.8, y: 5.4, w: 4.6, h: 1.1, fontSize: 13, color: C.accent5, margin: 0, valign: "top", isTextBox: true });
source(s, "Fondapol (juin 2026) ; IFRAP (sept. 2025) ; tribune de 7 économistes dont P. Aghion et X. Jaravel, Le Monde (2025)");
s.addNotes(
  "Il conteste ensuite le rendement. Il suffirait que les trois plus gros contribuables partent pour diviser les recettes par deux. " +
  "Sept économistes, dont Philippe Aghion et Xavier Jaravel, estiment dans une tribune au Monde que les plus riches trouveraient de nouvelles façons d'optimiser : la taxe rapporterait 5 milliards, pas 20. " +
  "Et selon un rapport du Sénat cité par l'IFRAP, la taxe ne rapporterait que 2 à 3 milliards si l'on déduit l'impôt sur les sociétés déjà payé. On est très loin des 20 milliards. " +
  "Enfin, la taxe serait inconstitutionnelle : le Conseil constitutionnel n'a jamais accepté un impôt sur la fortune de 2 % sans plafond. " +
  "En résumé : la taxe ferait fuir les riches et l'investissement, pour peu de recettes."
);

// --- Milieu (OFCE) ---
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Taxe Zucman · Trois positions" });
tag(s, "POSITION 3 · MILIEU", C.accent3);
s.addText("« Bon diagnostic, mauvais outil »", { placeholder: "title" });
auteur(s, "Guillaume Allègre et Xavier Timbeau, OFCE – Sciences Po, avril 2026");
bullets(s, [
  [{ text: "Oui, ", bold: true }, { text: "les très riches sont proportionnellement moins taxés" }],
  [{ text: "Mais un patrimoine est difficile à évaluer : ", bold: true }, { text: "une fortune « sur le papier » n'est pas de l'argent disponible (exemple de Mistral AI)" }],
  [{ text: "Le vrai risque : ", bold: true }, { text: "le départ des entrepreneurs, pas de l'argent" }],
], { x: 0.6, y: 2.3, w: 5.6, h: 2.9 });
s.addText([
  { text: "Leur alternative : ", options: { bold: true, color: C.background1, breakLine: true } },
  { text: "taxer les plus-values quand elles sont réalisées (vente, donation, héritage)", options: { color: C.background1 } },
], { x: 0.6, y: 5.3, w: 5.6, h: 1.35, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, fill: { color: C.accent3 },
  fontSize: 16, valign: "middle", margin: 14, isTextBox: true, objectName: "Alternative" });
s.addChart(pres.charts.BAR, [{ name: "Effet", labels: ["Chiffre d'affaires", "Valeur ajoutée", "Masse salariale"], values: [-15, -24, -31] }], {
  x: 6.7, y: 1.75, w: 6.0, h: 4.95, barDir: "col", barGapWidthPct: 60,
  chartColors: [HEX.accent3], valAxisMinVal: -35, valAxisMaxVal: 0, dataLabelFormatCode: '0" %"',
  showTitle: true, title: "Entreprise dont un gros actionnaire quitte la France : évolution après 5 ans", titleFontSize: 14, titleColor: HEX.accent5,
  ...chartBase, catAxisLabelPos: "high",
  objectName: "Graphique départ actionnaire",
});
source(s, SRC_OFCE2 + " (étude Bach et al., 2023)");
s.addNotes(
  "Entre les deux, deux économistes de l'OFCE donnent en partie raison à Zucman : oui, les très riches sont proportionnellement moins taxés. " +
  "Mais pour eux, taxer le patrimoine est le mauvais outil. Une entreprise non cotée est très difficile à évaluer. Et une fortune « sur le papier » n'est pas de l'argent disponible : " +
  "les fondateurs de Mistral AI verraient leur impôt presque doubler après une simple réévaluation de leur entreprise, sans gagner un euro de plus. " +
  "Le vrai risque, selon eux, n'est pas la fuite de l'argent mais le départ des entrepreneurs. Quand un gros actionnaire quitte la France, sa société perd environ 31 % de sa masse salariale en cinq ans. " +
  "Leur alternative : taxer les plus-values au moment où elles sont réalisées, lors d'une vente, d'une donation ou d'un héritage. En résumé : il faut bien taxer davantage les ultra-riches, mais autrement."
);

// --- Synthèse ---
pres.addSection({ title: "Taxe Zucman · Conclusion" });
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Taxe Zucman · Conclusion" });
tag(s, "CONCLUSION", C.text2);
s.addText("Trois réponses à une même question", { placeholder: "title" });
s.addTable([
  [{ text: "", options: {} }, head("Pour (Zucman, Oxfam)", C.accent1), head("Milieu (OFCE)", C.accent3), head("Contre (Fondapol, IFRAP)", C.accent2)],
  [rowLab("Diagnostic"), cell("Les riches paient moins que les autres"), cell("Vrai, les très riches sont moins taxés"), cell("Faux : le système est déjà très redistributif")],
  [rowLab("Rendement"), cell("≈ 20 Md€ par an"), cell("Incertain"), cell("2 à 3 Md€, et fragile")],
  [rowLab("Risque"), cell("Faible, c'est une question de justice"), cell("Départ des entrepreneurs"), cell("Exil, moins d'investissement, censure")],
  [rowLab("Optimisation"), cell("La cause de l'injustice : la taxe la corrige"), cell("Reporter l'impôt est trop facile : taxer à la vente, au don, à l'héritage"), cell("Problème réel des holdings : « holding pour tous »")],
  [rowLab("Solution"), cell("Impôt minimum de 2 %"), cell("Taxer les plus-values réalisées"), cell("Réformer la fiscalité du capital sans taxer le patrimoine")],
], {
  x: 0.6, y: 1.75, w: 12.1, colW: [1.9, 3.4, 3.4, 3.4], rowH: [0.55, 0.62, 0.62, 0.62, 0.62, 0.62],
  fontSize: 14, fontFace: THEME.bodyFontFace, border: { type: "solid", pt: 1, color: HEX.accent4 }, margin: 0.08,
  objectName: "Tableau synthèse Zucman",
});
s.addText("La justice fiscale vaut-elle le risque de faire partir les plus riches ?", {
  x: 0.6, y: 5.75, w: 12.1, h: 0.9, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, fill: { color: C.text2 },
  color: C.background1, fontSize: 21, bold: true, align: "center", valign: "middle", margin: 0.1, isTextBox: true,
  objectName: "Question",
});
s.addNotes(
  "Point commun étonnant : les trois camps reconnaissent le problème de l'optimisation fiscale via les holdings. Ils s'opposent sur la solution. " +
  "Les trois textes s'opposent sur trois questions : le diagnostic — les riches paient-ils vraiment moins que les autres ? — le rendement — 20, 5 ou 2 à 3 milliards ? — " +
  "et l'outil — taxer le patrimoine ou les plus-values ? Je vous laisse avec cette question : la justice fiscale vaut-elle le risque de faire partir les plus riches ?"
);

// --- Sources ---
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Taxe Zucman · Conclusion" });
tag(s, "SOURCES", C.text2);
s.addText("Sources", { placeholder: "title" });
[
  ["Raphaël Marchal", "Gabriel Zucman dénonce un projet de budget qui « protège les milliardaires »", "LCP, 2 octobre 2026",
    "https://lcp.fr/actualites/gabriel-zucman-denonce-un-projet-de-budget-qui-protege-les-milliardaires-442793", C.accent1],
  ["Reporterre", "Impôts favorables aux riches : Oxfam attaque l'État en justice", "13 octobre 2026",
    "https://reporterre.net/Impots-favorables-aux-riches-Oxfam-attaque-l-Etat-en-justice", C.accent1],
  ["Victor Fouquet", "L'impasse de la « taxe Zucman »", "Fondapol, juin 2026",
    "https://www.fondapol.org/etude/limpasse-de-la-taxe-zucman-pour-une-reforme-fiscale-equitable-et-efficace/", C.accent2],
  ["Fondation IFRAP", "Fiscalité des riches : le mirage des milliards de recettes", "Septembre 2025",
    "https://www.ifrap.org/la-revue/fiscalite-des-riches-le-mirage-des-milliards-eu-de-recettes", C.accent2],
  ["G. Allègre et X. Timbeau", "Pourquoi et comment taxer les plus riches ?", "OFCE – Sciences Po, avril 2026",
    "https://www.ofce.fr/wp/2026/6/v0/OFCE_SL.html", C.accent3],
].forEach(([who, t, where, url, col], i) => {
  const y = 1.75 + i * 0.98;
  s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.1, w: 0.36, h: 0.36, fill: { color: col }, line: { color: col } });
  s.addText([
    { text: who + " — ", options: { bold: true, color: C.text1 } },
    { text: t, options: { color: C.text1, hyperlink: { url }, breakLine: true } },
    { text: where, options: { color: C.accent5, fontSize: 14 } },
  ], { x: 1.2, y, w: 11.5, h: 0.85, fontSize: 16, valign: "top", margin: 0, isTextBox: true });
});

// ===========================================================================
// SUJET 3 : RETRAITES ET CAPITALISATION
// ===========================================================================
const SRC_FONDAPOL3 = "Martinot, Gravier et Muret, « Vers un système de retraite mixte répartition-capitalisation », Fondapol, nov. 2025";
const SRC_WEIL = "Eric Weil, La Grande Conversation (Terra Nova), mai 2025";
const SRC_TL = "D. Thesmar et A. Landier, interview, L'Express, 08/09/2026";

// --- Titre ---
pres.addSection({ title: "Retraites · Introduction" });
s = pres.addSlide({ masterName: "TITRE", sectionTitle: "Retraites · Introduction" });
s.addText("Les retraites", { placeholder: "title" });
s.addText("Faut-il introduire de la capitalisation ?", { placeholder: "body" });
s.addText("Trois positions sur l'avenir de notre système de retraite", {
  x: 0.8, y: 4.45, w: 11.7, h: 0.4, fontSize: 16, color: C.accent6, margin: 0, isTextBox: true,
});
[
  ["Pour · Fondapol", C.accent2],
  ["Milieu · Thesmar et Landier", C.accent3],
  ["Contre · CGT", C.accent1],
].forEach(([t, col], i) => {
  s.addText(t, {
    x: 0.8 + i * 3.75, y: 5.0, w: 3.5, h: 0.6, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.3,
    fill: { color: col }, color: C.background1, fontSize: 16, bold: true, align: "center", valign: "middle",
    margin: 0, isTextBox: true, objectName: "Position " + (i + 1),
  });
});
s.addNotes("Troisième sujet : les retraites, et la question de la capitalisation, très présente dans la campagne présidentielle 2027.");

// --- Répartition ou capitalisation ? ---
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Retraites · Introduction" });
tag(s, "INTRODUCTION", C.text2);
s.addText("Répartition ou capitalisation ?", { placeholder: "title" });
[
  ["Répartition (système actuel)", "Les cotisations des actifs d'aujourd'hui paient directement les pensions des retraités d'aujourd'hui. C'est la solidarité entre générations."],
  ["Capitalisation", "Chacun épargne pour sa propre retraite. L'argent est placé, souvent en Bourse, puis récupéré au moment de partir."],
].forEach(([t, d], i) => {
  const y = 1.75 + i * 2.55;
  card(s, 0.6, y, 5.6, 2.35, "Carte " + t);
  s.addText(t, { x: 0.9, y: y + 0.25, w: 5.0, h: 0.45, fontSize: 19, bold: true, color: C.text2, margin: 0, isTextBox: true });
  s.addText(d, { x: 0.9, y: y + 0.85, w: 5.0, h: 1.35, fontSize: 16, color: C.text1, margin: 0, valign: "top", isTextBox: true });
});
card(s, 6.6, 1.75, 6.1, 4.95, "Carte démographie");
s.addText("Nombre d'actifs (20-64 ans) pour un retraité (65 ans et plus)", { x: 6.9, y: 1.9, w: 5.5, h: 0.65, fontSize: 15, bold: true, color: C.text2, margin: 0, isTextBox: true });
[[1990, 4, "> 4"], ["Aujourd'hui", 2.5, "2,5"], ["Après 2040", 2, "< 2"]].forEach(([lab, v, txt], i) => {
  vbar(s, { x: 7.25 + i * 1.8, base: 5.75, w: 1.2, hMax: 2.5, val: v, max: 4, color: i === 0 ? C.accent4 : C.text2, valueText: txt, label: String(lab) });
});
s.addShape(pres.shapes.LINE, { x: 6.9, y: 5.75, w: 5.5, h: 0, line: { color: C.accent4, width: 1 } });
source(s, SRC_WEIL);
s.addNotes(
  "Aujourd'hui en France, les retraites fonctionnent presque entièrement par répartition : les cotisations des actifs d'aujourd'hui paient directement les pensions des retraités d'aujourd'hui. " +
  "C'est la solidarité entre générations. Dans la capitalisation, au contraire, chacun épargne pour sa propre retraite : l'argent est placé, souvent en Bourse, et on le récupère au moment de partir. " +
  "Le problème, c'est la démographie : on comptait plus de 4 actifs pour un retraité en 1990, on en compte 2,5 aujourd'hui, et ce sera moins de 2 après 2040."
);

// --- La campagne 2027 ---
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Retraites · Introduction" });
tag(s, "INTRODUCTION", C.text2);
s.addText("La capitalisation dans la campagne 2027", { placeholder: "title" });
[
  ["Édouard Philippe", "Horizons", "15 %", "de capitalisation dans le système"],
  ["Gabriel Attal", "Renaissance", "1 000 €", "versés par l'État à chaque nouveau-né sur un compte d'épargne"],
  ["Philippe Brun", "PS", "15 Md€ / an", "pour un fonds de pension public géré par les syndicats"],
].forEach(([who, party, big, small], i) => {
  const x = 0.6 + i * 4.1;
  card(s, x, 1.75, 3.9, 3.0, "Candidat " + who);
  s.addText([
    { text: who, options: { bold: true, color: C.text1, breakLine: true } },
    { text: party, options: { color: C.accent5, fontSize: 14 } },
  ], { x: x + 0.3, y: 1.95, w: 3.3, h: 0.75, fontSize: 18, margin: 0, isTextBox: true });
  s.addText(big, { x: x + 0.3, y: 2.75, w: 3.3, h: 0.8, fontSize: 36, bold: true, color: C.text2, fontFace: THEME.headFontFace, margin: 0, valign: "bottom", isTextBox: true });
  s.addText(small, { x: x + 0.3, y: 3.6, w: 3.3, h: 1.0, fontSize: 14, color: C.accent5, margin: 0, valign: "top", isTextBox: true });
});
s.addText("Aussi pour : Bruno Retailleau, David Lisnard, Xavier Bertrand ; Marine Le Pen pour une part « volontaire »", {
  x: 0.6, y: 4.9, w: 12.1, h: 0.4, fontSize: 15, color: C.text1, margin: 0, isTextBox: true,
});
card(s, 0.6, 5.5, 12.1, 1.15, "Carte part répartition");
s.addText([
  { text: "98 % ", options: { bold: true, color: C.text2, fontSize: 28 } },
  { text: "des retraites françaises sont payées par répartition, contre environ ", options: { color: C.text1 } },
  { text: "83 % ", options: { bold: true, color: C.text2, fontSize: 28 } },
  { text: "en moyenne dans l'OCDE", options: { color: C.text1 } },
], { x: 0.9, y: 5.6, w: 11.5, h: 0.95, fontSize: 17, valign: "middle", margin: 0, isTextBox: true });
source(s, "Journal du Net, 21/08/2026 ; " + SRC_WEIL);
s.addNotes(
  "La France paie 98 % de ses retraites par répartition, contre environ 83 % en moyenne dans l'OCDE. Mais la capitalisation est revenue dans la campagne : " +
  "Édouard Philippe veut 15 % de capitalisation, Gabriel Attal propose de verser 1 000 euros à chaque nouveau-né sur un compte d'épargne, " +
  "et même le socialiste Philippe Brun propose un fonds de pension public de 15 milliards par an."
);

// --- Pour (Fondapol) ---
pres.addSection({ title: "Retraites · Trois positions" });
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Retraites · Trois positions" });
tag(s, "POSITION 1 · POUR", C.accent2);
s.addText("« Un quart des retraites en capitalisation »", { placeholder: "title" });
auteur(s, "Martinot, Gravier et Muret, Fondapol, novembre 2025");
bullets(s, [
  [{ text: "Un étage obligatoire : ", bold: true }, { text: "25 % des pensions du privé" }],
  [{ text: "Un meilleur rendement ", bold: true }, { text: "que la répartition" }],
  [{ text: "Un patrimoine ", bold: true }, { text: "pour chaque salarié" }],
  [{ text: "Face aux krachs : ", bold: true }, { text: "fonds testé avec une crise tous les 12 ans" }],
  [{ text: "En version collective : ", bold: true }, { text: "on lisse les rendements entre générations et on sécurise l'épargne avant le départ" }],
], { x: 0.6, y: 2.3, w: 5.6, h: 4.3 });
card(s, 6.7, 1.75, 6.0, 2.75, "Carte rendements");
s.addText("Rendement par an au-dessus de l'inflation", { x: 7.0, y: 1.92, w: 5.5, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true });
hbar(s, { x: 7.0, y: 2.42, w: 5.5, h: 0.6, val: 1, max: 5.8, color: C.accent4, label: "Répartition", valueText: "≈ 1 %" });
hbar(s, { x: 7.0, y: 3.02, w: 5.5, h: 0.6, val: 4, max: 5.8, color: C.accent2, label: "Hypothèse Fondapol", valueText: "4 %" });
hbar(s, { x: 7.0, y: 3.62, w: 5.5, h: 0.6, val: 5.8, max: 5.8, color: C.accent2, label: "Fonds suédois AP7", valueText: "5,8 %" });
statCard(s, { x: 6.7, y: 4.75, w: 6.0, h: 1.95, big: "− 6 points", small: "de cotisations retraite à long terme, selon les simulations de la Fondapol", color: C.accent2 });
source(s, SRC_FONDAPOL3 + " ; répartition : Thesmar et Landier (2026)");
s.addNotes(
  "La Fondapol, un think tank libéral, propose un étage de capitalisation obligatoire qui paierait à terme 25 % des pensions du privé. " +
  "Son premier argument est le rendement. La répartition rapporte environ 1 % par an au-dessus de l'inflation, parce qu'elle dépend de la hausse des salaires. " +
  "L'étude retient 4 % pour la capitalisation, et le fonds de pension public suédois a fait encore mieux : 5,8 % par an entre 2000 et 2023. " +
  "Résultat selon eux : à long terme, les cotisations pourraient baisser de plus de 6 points, et chaque salarié aurait un vrai patrimoine. " +
  "Ils ont même testé le système avec une crise boursière tous les 12 ans : le fonds tiendrait. " +
  "Et avec une capitalisation collective, gérée par un fonds national, on peut se protéger des krachs : on garde une partie des rendements pour lisser entre les générations, " +
  "et on sécurise l'épargne à l'approche de la retraite. En résumé : la capitalisation rendrait les retraites plus sûres et moins chères."
);

// --- Contre (CGT) ---
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Retraites · Trois positions" });
tag(s, "POSITION 2 · CONTRE", C.accent1);
s.addText("« Une régression sociale »", { placeholder: "title" });
auteur(s, "CGT, janvier 2025 (et Eric Weil, La Grande Conversation, mai 2025)");
bullets(s, [
  [{ text: "Fin de la solidarité : ", bold: true }, { text: "chômage, maladie ou maternité ne rapportent plus de droits" }],
  [{ text: "Les inégalités de salaire ", bold: true }, { text: "se retrouvent à la retraite" }],
  [{ text: "Le risque boursier : ", bold: true }, { text: "aux Pays-Bas, la crise de 2008 a fait fondre l'épargne retraite" }],
  [{ text: "Payer deux fois : ", bold: true }, { text: "les pensions actuelles et sa propre épargne, une « génération sacrifiée »" }],
], { x: 0.6, y: 2.3, w: 5.6, h: 4.3 });
statCard(s, { x: 6.7, y: 1.75, w: 6.0, h: 2.35, big: "300 à 400 Md€", small: "coût cumulé de la transition pour seulement 10 à 15 % de capitalisation (environ 20 Md€ de plus par an)", color: C.accent1 });
statCard(s, { x: 6.7, y: 4.35, w: 6.0, h: 2.35, big: "30 à 40 ans", small: "avant que la capitalisation ne produise des gains nets", color: C.accent1 });
source(s, "CGT, « La capitalisation n'est pas une solution : c'est une régression sociale », 31/01/2025 ; chiffres : " + SRC_WEIL);
s.addNotes(
  "Pour la CGT, c'est tout l'inverse. D'abord, la capitalisation casse la solidarité : les périodes de chômage, de maladie ou de maternité ne rapportent plus de droits, " +
  "et les inégalités de salaire se retrouvent à la retraite. Ensuite, il y a le risque boursier. Aux Pays-Bas, la crise de 2008 a fait fondre l'épargne retraite et a obligé à réformer le système. " +
  "Enfin, il faut payer deux fois pendant la transition : les pensions des retraités actuels, et en plus l'épargne des actifs. " +
  "L'économiste Eric Weil chiffre ce coût à 300 à 400 milliards d'euros pour seulement 10 à 15 % de capitalisation, avec des gains qui n'apparaîtraient qu'après 30 à 40 ans. " +
  "En résumé : plus de risques et moins de solidarité, pour un coût énorme. Attention : l'article de la CGT ne donne pas de chiffres, ceux de la diapo viennent d'Eric Weil."
);

// --- Qui paierait la transition ? ---
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Retraites · Trois positions" });
tag(s, "LE CŒUR DU DÉBAT", C.text2);
s.addText("Qui paierait la transition ?", { placeholder: "title" });
card(s, 0.6, 1.75, 4.6, 4.95, "Carte génération sacrifiée");
s.addText("« Une génération un peu sacrifiée »", { x: 0.9, y: 2.0, w: 4.0, h: 0.9, fontSize: 22, bold: true, color: C.text2, fontFace: THEME.headFontFace, margin: 0, valign: "top", isTextBox: true });
s.addText(
  "Pendant des décennies, les actifs devraient payer les pensions des retraités actuels ET remplir le nouveau fonds de capitalisation, jusqu'à ce que les engagements de la répartition diminuent.",
  { x: 0.9, y: 3.05, w: 4.0, h: 2.2, fontSize: 16, color: C.text1, margin: 0, valign: "top", isTextBox: true });
s.addText("C'est votre génération qui serait concernée.", { x: 0.9, y: 5.6, w: 4.0, h: 0.8, fontSize: 16, bold: true, color: C.text2, margin: 0, valign: "top", isTextBox: true });
s.addText("Les pistes pour financer le départ, et leurs limites", { x: 5.6, y: 1.75, w: 7.1, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true });
[
  ["≈ 91 Md€", "Prendre les réserves de l'Agirc-Arrco", "Mais elles servent à payer les retraites en cas de crise financière"],
  ["≈ 80 Md€", "Vendre des participations de l'État dans des entreprises", "Mais l'État perd leurs dividendes, et ça ne suffit pas"],
  ["Pensions", "Les faire augmenter moins vite que l'inflation", "Ce sont alors les retraités actuels qui paient l'effort"],
].forEach(([big, piste, limite], i) => {
  const y = 2.3 + i * 1.5;
  card(s, 5.6, y, 7.1, 1.35, "Piste " + (i + 1));
  s.addText(big, { x: 5.85, y: y + 0.15, w: 1.9, h: 1.05, fontSize: 24, bold: true, color: C.text2, fontFace: THEME.headFontFace, valign: "middle", margin: 0, isTextBox: true });
  s.addText([
    { text: piste, options: { bold: true, color: C.text1, breakLine: true } },
    { text: limite, options: { color: C.accent5, fontSize: 14 } },
  ], { x: 7.85, y: y + 0.12, w: 4.65, h: 1.1, fontSize: 15, valign: "middle", margin: 0, isTextBox: true });
});
source(s, "Agirc-Arrco (réserves au 31/12/2025) ; vidéo YouTube « Retraite par capitalisation : ce qu'il faut savoir » (80 Md€ : non vérifié) ; Fondapol (nov. 2025)");
s.addNotes(
  "Tout le monde, ou presque, est d'accord sur un point : le plus difficile, c'est la transition. Pendant des décennies, les actifs devraient payer deux fois : " +
  "les pensions des retraités actuels, et en plus l'épargne du nouveau fonds. Il y aurait donc une génération un peu sacrifiée, et ce serait la nôtre. " +
  "Pour financer le départ, plusieurs pistes existent, mais aucune n'est gratuite : prendre les réserves de l'Agirc-Arrco, environ 91 milliards fin 2025, qui servent pourtant en cas de crise ; " +
  "vendre des participations de l'État, environ 80 milliards, mais l'État perdrait leurs dividendes ; ou faire augmenter les pensions moins vite que l'inflation, ce qui fait payer les retraités actuels. " +
  "La Fondapol propose d'ailleurs un mélange de ces pistes."
);

// --- Milieu (Thesmar et Landier) ---
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Retraites · Trois positions" });
tag(s, "POSITION 3 · MILIEU", C.accent3);
s.addText("« Pas une recette miracle »", { placeholder: "title" });
auteur(s, "D. Thesmar (MIT) et A. Landier (HEC), L'Express, sept. 2026");
bullets(s, [
  [{ text: "Oui à un complément ", bold: true }, { text: "de capitalisation, sur le modèle suédois" }],
  [{ text: "Faux argument : ", bold: true }, { text: "le rendement de la Bourse ne suffira pas à « boucher le trou »" }],
  [{ text: "Les pensions déjà promises ", bold: true }, { text: "doivent continuer à être payées" }],
  [{ text: "Risques : ", bold: true }, { text: "frais de gestion élevés, épargne détournée par l'État ou des lobbies" }],
], { x: 0.6, y: 2.3, w: 5.6, h: 3.0 });
s.addText([
  { text: "Seuls deux leviers : ", options: { bold: true, color: C.background1, breakLine: true } },
  { text: "l'âge de départ ou le niveau des pensions, avec un effort des générations nées entre 1950 et 1970", options: { color: C.background1 } },
], { x: 0.6, y: 5.3, w: 5.6, h: 1.35, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, fill: { color: C.accent3 },
  fontSize: 16, valign: "middle", margin: 14, isTextBox: true, objectName: "Leviers" });
card(s, 6.7, 1.75, 6.0, 4.95, "Carte déficit");
s.addText("Le déficit public, en points de PIB", { x: 7.0, y: 1.92, w: 5.5, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true });
vbar(s, { x: 7.6, base: 5.75, w: 1.5, hMax: 2.8, val: 5, max: 5, color: C.accent4, valueText: "≈ 5", label: "Déficit public\ntotal" });
vbar(s, { x: 10.3, base: 5.75, w: 1.5, hMax: 2.8, val: 3, max: 5, color: C.accent3, valueText: "≈ 3", label: "dont le déficit\ndes retraites" });
s.addShape(pres.shapes.LINE, { x: 7.0, y: 5.75, w: 5.5, h: 0, line: { color: C.accent4, width: 1 } });
source(s, SRC_TL);
s.addNotes(
  "Entre les deux, les économistes David Thesmar et Augustin Landier sont favorables à un complément de capitalisation, sur le modèle suédois. " +
  "Mais ils dénoncent un faux argument : croire que le meilleur rendement de la Bourse suffira à « boucher le trou » des retraites. " +
  "Le déficit des retraites représente environ 3 points de PIB, sur un déficit public total d'environ 5 points. Et la capitalisation ne change rien aux pensions déjà promises, qu'il faut continuer à payer. " +
  "Pour eux, il ne reste que deux leviers : l'âge de départ ou le niveau des pensions. Ils proposent que les générations nées entre 1950 et 1970 fassent un effort. " +
  "En résumé : la capitalisation est un bon complément, pas une solution au déficit."
);

// --- Synthèse ---
pres.addSection({ title: "Retraites · Conclusion" });
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Retraites · Conclusion" });
tag(s, "CONCLUSION", C.text2);
s.addText("Trois visions de nos retraites", { placeholder: "title" });
s.addTable([
  [{ text: "", options: {} }, head("Pour (Fondapol)", C.accent2), head("Milieu (Thesmar et Landier)", C.accent3), head("Contre (CGT)", C.accent1)],
  [rowLab("Le risque"), cell("Maîtrisé : le fonds résiste aux crises"), cell("Réel : frais, détournement de l'épargne"), cell("Trop grand : krach de 2008")],
  [rowLab("La transition"), cell("Faisable avec un capital de départ"), cell("Coûteuse : il faut payer les pensions promises"), cell("Payer deux fois (300 à 400 Md€)")],
  [rowLab("La solidarité"), cell("Un patrimoine pour chaque salarié"), cell("Effort des générations 1950-1970"), cell("La répartition protège les carrières hachées")],
  [rowLab("Solution"), cell("25 % de capitalisation obligatoire"), cell("Petit complément + âge ou pensions"), cell("Rester en répartition")],
], {
  x: 0.6, y: 1.75, w: 12.1, colW: [1.9, 3.4, 3.4, 3.4], rowH: [0.6, 0.72, 0.72, 0.72, 0.72],
  fontSize: 15, fontFace: THEME.bodyFontFace, border: { type: "solid", pt: 1, color: HEX.accent4 }, margin: 0.08,
  objectName: "Tableau synthèse retraites",
});
s.addText("Pour votre retraite : solidarité entre générations ou votre propre épargne ?", {
  x: 0.6, y: 5.65, w: 12.1, h: 1.0, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, fill: { color: C.text2 },
  color: C.background1, fontSize: 21, bold: true, align: "center", valign: "middle", margin: 0.1, isTextBox: true,
  objectName: "Question",
});
s.addNotes(
  "Les trois positions s'opposent sur trois points : le risque — la Bourse est-elle trop dangereuse pour nos retraites ? — la transition — qui paie pendant le passage d'un système à l'autre ? — " +
  "et la solidarité — entre générations, ou chacun pour soi ? Et vous, qui êtes la génération concernée : pour votre retraite, préférez-vous la solidarité entre générations ou votre propre épargne ?"
);

// --- Sources ---
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Retraites · Conclusion" });
tag(s, "SOURCES", C.text2);
s.addText("Sources", { placeholder: "title" });
[
  ["Martinot, Gravier et Muret", "Vers un système de retraite mixte répartition-capitalisation", "Fondapol, novembre 2025",
    "https://www.fondapol.org/etude/vers-un-systeme-de-retraite-mixte-repartition-capitalisation/", C.accent2],
  ["CGT", "La capitalisation n'est pas une solution : c'est une régression sociale", "31 janvier 2025",
    "https://www.cgt.fr/actualites/retraites/la-capitalisation-nest-pas-une-solution-cest-une-regression-sociale-58", C.accent1],
  ["Eric Weil", "Retraites par capitalisation : est-ce utile ? Est-ce faisable ?", "La Grande Conversation (Terra Nova), mai 2025",
    "https://www.lagrandeconversation.com/debat/societe/retraites-par-capitalisation-est-ce-utile-est-ce-faisable/", C.accent1],
  ["D. Thesmar et A. Landier", "« La capitalisation n'est pas la recette miracle pour les retraites »", "L'Express, 8 septembre 2026",
    "https://intelli.news/2026/09/08/la-capitalisation-nest-pas-la-recette-miracle-pour-les-retraites-la-mise-au-point-des-economistes-thesmar-et-landier/", C.accent3],
  ["Adrien Bar Hiyé", "La retraite par capitalisation a le vent en poupe dans la campagne présidentielle 2027", "Journal du Net, 21 août 2026",
    "https://www.journaldunet.com/business/action-publique/1553921-propositions-de-capitalisation-pour-la-retraite-dans-la-campagne-presidentielle-2027/", C.text2],
  ["Vidéo YouTube", "Retraite par capitalisation : ce qu'il faut savoir", "Chaîne non identifiée",
    "https://www.youtube.com/watch?v=Oestb_xnNPc", C.text2],
].forEach(([who, t, where, url, col], i) => {
  const y = 1.75 + i * 0.84;
  s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.1, w: 0.36, h: 0.36, fill: { color: col }, line: { color: col } });
  s.addText([
    { text: who + " — ", options: { bold: true, color: C.text1 } },
    { text: t, options: { color: C.text1, hyperlink: { url }, breakLine: true } },
    { text: where, options: { color: C.accent5, fontSize: 14 } },
  ], { x: 1.2, y, w: 11.5, h: 0.85, fontSize: 16, valign: "top", margin: 0, isTextBox: true });
});

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("OK", OUT);
})();
