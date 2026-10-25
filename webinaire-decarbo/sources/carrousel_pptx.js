// Génère le carrousel PowerPoint du webinar Axys x Decarbo'Solution (A4 portrait, charte Axys).
// Usage : node carrousel_pptx.js <dossier de sortie>
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = path.join(process.argv[2] || ".", "Carrousel_Webinaire_Axys_Decarbo.pptx");
const DIR = __dirname;
const FOND = path.join(DIR, "fond_degrade.png");
const LOGO = path.join(DIR, "logo_axys.png");
const PHOTO = path.join(DIR, "photo_maeva.png");

const THEME = {
  name: "Axys",
  headFontFace: "Montserrat",
  bodyFontFace: "Montserrat",
  colors: {
    dk1: "404148", lt1: "FFFFFF", dk2: "004C5D", lt2: "F2F2F2",
    accent1: "008A8C", accent2: "02AEB1", accent3: "FF2148",
    accent4: "004C5D", accent5: "9CE6E7", accent6: "404148",
    hlink: "02AEB1", folHlink: "008A8C",
  },
};

const pres = new pptxgen();
pres.defineLayout({ name: "A4_PORTRAIT", width: 8.27, height: 11.69 });
pres.layout = "A4_PORTRAIT";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "Maëva Signor";
pres.company = "Axys Consultants";
pres.title = "Données fournisseurs : collectez moins, exploitez mieux";

const C = pres.SchemeColor;
const W = 8.27, M = 0.67, CW = W - 2 * M;
const LOGO_RATIO = 231 / 58;

// ---------- Dispositions ----------
const kicker = (y) => ({ placeholder: { options: { name: "kicker", type: "body", x: M, y, w: CW, h: 0.32, fontSize: 11.5, bold: true, color: C.accent5, charSpacing: 2, margin: 0, valign: "top" }, text: "SURTITRE" } });
const titre = (y, h, size) => ({ placeholder: { options: { name: "title", type: "title", x: M, y, w: CW, h, fontSize: size, bold: true, color: C.background1, align: "left", margin: 0, valign: "top", lineSpacingMultiple: 0.95 }, text: "TITRE-MESSAGE" } });

pres.defineSlideMaster({
  title: "COUVERTURE",
  background: { path: FOND },
  objects: [
    { image: { path: LOGO, x: (W - 0.5 * LOGO_RATIO) / 2, y: 1.0, w: 0.5 * LOGO_RATIO, h: 0.5 } },
    kicker(3.4),
    titre(3.75, 2.3, 40),
  ],
});

const pied = [
  { line: { x: M, y: 11.1, w: 2.75, h: 0, line: { color: C.background1, width: 0.75, transparency: 45 } } },
  { image: { path: LOGO, x: (W - 0.27 * LOGO_RATIO) / 2, y: 10.965, w: 0.27 * LOGO_RATIO, h: 0.27 } },
  { line: { x: W - M - 2.75 - 0.55, y: 11.1, w: 2.75, h: 0, line: { color: C.background1, width: 0.75, transparency: 45 } } },
];
pres.defineSlideMaster({
  title: "CONTENU",
  background: { path: FOND },
  objects: [...pied, kicker(1.0), titre(1.35, 1.6, 31)],
  slideNumber: { x: W - M - 0.45, y: 10.97, w: 0.45, h: 0.26, fontSize: 9, bold: true, color: C.background1, align: "right", margin: 0 },
});

pres.defineSlideMaster({
  title: "FINAL",
  background: { path: FOND },
  objects: [
    { image: { path: LOGO, x: 2.55, y: 10.86, w: 0.34 * LOGO_RATIO, h: 0.34 } },
    kicker(1.0), titre(1.35, 1.2, 31),
  ],
});

// ---------- Helpers ----------
const carte = (fill = { color: C.background1, transparency: 88 }) => ({ shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, fill, line: { type: "none" } });
const blanc = (o = {}) => ({ color: C.background1, ...o });

pres.addSection({ title: "Carrousel" });

// ---------- Page 1 : accroche ----------
let s = pres.addSlide({ masterName: "COUVERTURE", sectionTitle: "Carrousel" });
s.addText("WEBINAR · AXYS × DECARBO'SOLUTION", { placeholder: "kicker" });
s.addText([
  { text: "DONNÉES FOURNISSEURS", options: { fontSize: 25, color: C.accent5, breakLine: true } },
  { text: "COLLECTEZ MOINS,", options: { breakLine: true } },
  { text: "EXPLOITEZ MIEUX" },
], { placeholder: "title" });
s.addText([
  { text: "Combien de questionnaires durabilité envoyez-vous à vos fournisseurs chaque année ?", options: { breakLine: true } },
  { text: "Et combien de réponses sauriez-vous " },
  { text: "prouver", options: { bold: true } },
  { text: " ?" },
], { objectName: "Accroche", isTextBox: true, x: M, y: 6.25, w: CW, h: 1.25, fontSize: 16, color: C.background1, margin: 0, valign: "top", lineSpacingMultiple: 1.2 });
s.addText("Mardi 24 novembre 2026 · 11h · en ligne", { objectName: "Date", ...carte({ color: C.background1 }), rectRadius: 0.3, x: M, y: 7.85, w: 4.9, h: 0.62, fontSize: 16, bold: true, color: C.text2, align: "center", valign: "middle", margin: 0 });
s.addText("45 min d'échange + 15 min de questions", { objectName: "Format", isTextBox: true, x: M, y: 8.7, w: CW, h: 0.35, fontSize: 13.5, bold: true, color: C.background1, margin: 0 });
s.addText("›››", { objectName: "Chevrons", isTextBox: true, x: W - M - 0.8, y: 10.85, w: 0.8, h: 0.45, fontSize: 24, color: C.background1, align: "right", margin: 0 });
s.addNotes("Page d'accroche : le titre du webinar, la question d'ouverture et la date.");

// ---------- Page 2 : le constat ----------
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Carrousel" });
s.addText("LE CONSTAT", { placeholder: "kicker" });
s.addText("VOS DONNÉES FOURNISSEURS EXISTENT. SONT-ELLES UTILISABLES ?", { placeholder: "title" });
const constats = [
  ["Dispersées", "Outil achats, fichiers Excel, mails, plateformes, équipe RSE : la même information vit à plusieurs endroits."],
  ["Hétérogènes", "Chaque questionnaire a son format et son niveau de détail. Les réponses ne se comparent pas."],
  ["Peu réutilisables", "Une demande, un usage. Puis on recommence, et le fournisseur aussi."],
];
constats.forEach(([t, d], i) => {
  const y = 3.35 + i * 1.5;
  s.addText(String(i + 1), { objectName: `Numéro ${i + 1}`, shape: pres.shapes.OVAL, x: M, y: y + 0.24, w: 0.82, h: 0.82, fill: { color: C.background1, transparency: 80 }, line: { type: "none" }, fontSize: 22, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
  s.addText([
    { text: t.toUpperCase(), options: { fontSize: 14, bold: true, breakLine: true, paraSpaceAfter: 4 } },
    { text: d, options: { fontSize: 13 } },
  ], { objectName: `Constat ${i + 1}`, ...carte(), x: M + 1.05, y, w: CW - 1.05, h: 1.3, color: C.background1, valign: "middle", margin: [16, 12, 12, 12], lineSpacingMultiple: 1.1 });
});
s.addText([
  { text: "Le VS, nouveau format du VSME, pose un cadre commun sur " },
  { text: "ce qu'il faut demander", options: { bold: true, color: C.accent1 } },
  { text: ".", options: { breakLine: true } },
  { text: "Reste à savoir comment le prouver, puis le réutiliser." },
], { objectName: "Message clé", ...carte({ color: C.background1 }), x: M, y: 8.05, w: CW, h: 1.35, fontSize: 15, bold: true, color: C.text2, valign: "middle", margin: [18, 14, 18, 14], lineSpacingMultiple: 1.15 });
s.addNotes("Constat porté par Axys. Ne pas citer de chiffre de volume de questionnaires (non sourcé).");

// ---------- Page 3 : l'enjeu ----------
s = pres.addSlide({ masterName: "CONTENU", sectionTitle: "Carrousel" });
s.addText("L'ENJEU", { placeholder: "kicker" });
s.addText("DEMAIN, UN AUDITEUR REGARDERA TROIS CHOSES", { placeholder: "title" });
const cw3 = (CW - 2 * 0.17) / 3;
["Ce que vous collectez", "Comment c'est connecté", "Comment c'est contrôlé"].forEach((t, i) => {
  s.addText([
    { text: "?", options: { fontSize: 34, bold: true, color: C.accent5, breakLine: true } },
    { text: t, options: { fontSize: 14, bold: true } },
  ], { objectName: `Regard auditeur ${i + 1}`, ...carte(), x: M + i * (cw3 + 0.17), y: 3.15, w: cw3, h: 1.95, color: C.background1, align: "center", valign: "middle", margin: 10 });
});
s.addText("Et une même donnée peut servir plusieurs usages :", { objectName: "Usages titre", isTextBox: true, x: M, y: 5.5, w: CW, h: 0.35, fontSize: 14, bold: true, color: C.background1, margin: 0 });
const chips = [
  [["CSRD · Scope 3", 1.75], ["Audit", 0.95], ["Pilotage fournisseurs", 2.55]],
  [["Coût carbone · ETS, CBAM", 2.95], ["Appels d'offres", 1.85]],
];
chips.forEach((row, r) => {
  let x = M;
  row.forEach(([t, w]) => {
    s.addText(t, { objectName: `Usage ${t}`, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.24, x, y: 6.0 + r * 0.62, w, h: 0.48, fill: { color: C.background1, transparency: 100 }, line: { color: C.background1, width: 1.25 }, fontSize: 12.5, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
    x += w + 0.15;
  });
});
s.addText([
  { text: "Repère : ", options: { bold: true } },
  { text: "l'ISSA 5000 devient la référence internationale de l'assurance durabilité en décembre 2026 (périodes ouvertes à partir du 15 décembre). Ce que l'auditeur attendra de vos données se précise." },
], { objectName: "Repère ISSA 5000", ...carte({ color: C.text2, transparency: 40 }), x: M, y: 7.55, w: CW, h: 1.35, fontSize: 13, color: C.background1, valign: "middle", margin: [18, 12, 18, 12], lineSpacingMultiple: 1.15 });
s.addText("Source : IAASB, ISSA 5000, date d'application (synthèse ICAEW).", { objectName: "Source", isTextBox: true, x: M, y: 10.45, w: CW, h: 0.25, fontSize: 9, color: C.background1, margin: 0 });
s.addNotes("Formulation ISSA 5000 issue du brief ; à revérifier sur la source avant diffusion. Ne pas écrire « s'impose en France ».");

// ---------- Page 4 : rendez-vous ----------
s = pres.addSlide({ masterName: "FINAL", sectionTitle: "Carrousel" });
s.addText("LE 24 NOVEMBRE À 11H", { placeholder: "kicker" });
s.addText([{ text: "TROIS QUESTIONS,", options: { breakLine: true } }, { text: "DEUX VOIX, UNE HEURE" }], { placeholder: "title" });
["Pourquoi une donnée calculée ne suffit plus", "Ce qui rend une donnée carbone authentifiée, tracée, reproductible", "Comment la réutiliser dans vos outils achats"].forEach((t, i) => {
  const y = 2.95 + i * 0.98;
  s.addText(`0${i + 1}`, { objectName: `Question ${i + 1} numéro`, ...carte({ color: C.background1, transparency: 80 }), x: M, y, w: 0.85, h: 0.82, fontSize: 22, bold: true, color: C.accent5, align: "center", valign: "middle", margin: 0 });
  s.addText(t, { objectName: `Question ${i + 1}`, ...carte(), x: M + 1.0, y, w: CW - 1.0, h: 0.82, fontSize: 14.5, bold: true, color: C.background1, valign: "middle", margin: [16, 8, 12, 8] });
});
s.addText("Sondage en direct · Checklist des 5 premiers pas remise aux inscrits", { objectName: "Bonus", isTextBox: true, x: M, y: 5.95, w: CW, h: 0.35, fontSize: 12, color: C.background1, margin: 0 });
s.addText("Je m'inscris →", { objectName: "Bouton inscription", ...carte({ color: C.background1 }), rectRadius: 0.3, x: M, y: 6.55, w: 2.6, h: 0.62, fontSize: 15, bold: true, color: C.text2, align: "center", valign: "middle", margin: 0 });
// Intervenants
s.addImage({ path: PHOTO, objectName: "Photo Maëva Signor", x: M, y: 7.75, w: 0.9, h: 0.9, rounding: true, altText: "Photo de Maëva Signor" });
s.addText([
  { text: "Maëva Signor", options: { fontSize: 14, bold: true, breakLine: true } },
  { text: "Senior Manager Achats", options: { fontSize: 11, breakLine: true } },
  { text: "Axys", options: { fontSize: 11 } },
], { objectName: "Intervenant Axys", isTextBox: true, x: M + 1.05, y: 7.75, w: 2.3, h: 0.9, color: C.background1, valign: "middle", margin: 0 });
s.addText("Photo", { objectName: "Photo Decarbo (à remplacer)", shape: pres.shapes.OVAL, x: 4.2, y: 7.75, w: 0.9, h: 0.9, fill: { color: C.background1, transparency: 80 }, line: { type: "none" }, fontSize: 10, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
s.addText([
  { text: "Philippe Mangeard", options: { fontSize: 14, bold: true, breakLine: true } },
  { text: "[Fonction à confirmer]", options: { fontSize: 11, breakLine: true } },
  { text: "Decarbo'Solution", options: { fontSize: 11 } },
], { objectName: "Intervenant Decarbo", isTextBox: true, x: 5.25, y: 7.75, w: 2.4, h: 0.9, color: C.background1, valign: "middle", margin: 0 });
// Co-branding
s.addText("×", { objectName: "Croix co-branding", isTextBox: true, x: 4.0, y: 10.83, w: 0.35, h: 0.4, fontSize: 18, color: C.background1, align: "center", valign: "middle", margin: 0 });
s.addText("Logo Decarbo'Solution", { objectName: "Logo Decarbo (à remplacer)", shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, x: 4.5, y: 10.8, w: 2.0, h: 0.46, fill: { color: C.background1, transparency: 100 }, line: { color: C.background1, width: 1, dashType: "dash" }, fontSize: 10.5, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
s.addNotes("Remplacer : lien du bouton « Je m'inscris », photo et fonction de l'intervenant Decarbo'Solution, logo Decarbo'Solution.");

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("OK", OUT);
})();
