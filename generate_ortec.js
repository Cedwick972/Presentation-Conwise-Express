const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Icon imports
const { FaTruck, FaShieldAlt, FaChartLine, FaClock, FaMapMarkerAlt, FaUsers, FaCheckCircle, FaLaptop, FaPhone, FaEnvelope, FaGlobe, FaIndustry, FaRoute, FaCogs, FaHandshake, FaFileAlt, FaStar, FaArrowRight } = require("react-icons/fa");
const { MdDashboard, MdSpeed, MdSecurity, MdLocationOn } = require("react-icons/md");
const { HiLightningBolt } = require("react-icons/hi");

// === HELPERS ===
function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

function imgToBase64(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const mime = ext === ".png" ? "image/png" : "image/jpeg";
  const data = fs.readFileSync(filePath);
  return `${mime};base64,${data.toString("base64")}`;
}

// === COLORS (Conwise brand) ===
const C = {
  navy: "1e4b77",
  blue: "3498db",
  lightBlue: "59b3e6",
  paleBlue: "e3f2fd",
  dark: "0f172a",
  gray: "64748b",
  lightGray: "F0F4F8",
  white: "FFFFFF",
  accent: "2ECC71",    // Green accent for checkmarks / success
  orangeAccent: "E67E22",
};

// Factory functions for reusable objects
const makeShadowCard = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.12 });
const makeShadowSoft = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.08 });

// === ASSET PATHS ===
const BASE = "C:\\Users\\User\\Documents\\claude code\\Conwise express propal";
const LOGO_CW = path.join(BASE, "Logo conwise express.png");
const LOGO_ORTEC = path.join(BASE, "logo_Groupe_Ortec-removebg-preview.png");
const IMG_BANNER = path.join(BASE, "image fond banniere couverture conwise express.png");
const IMG_FLEET = path.join(BASE, "flotte-automobile-entreprise-rapatriement- transfert-de-vehicules (1).jpeg");
const IMG_FLEET_MGR = path.join(BASE, "image gestionnaire de flotte entreprise.png");
const IMG_EDL = path.join(BASE, "image homme realise un etat des lieux.png");
const IMG_COMMENT = path.join(BASE, "comment ça marche convoyage (1).jpg");
const IMG_DASHBOARD = path.join(BASE, "dashboard Fleet manager Groupe Ortec.png");
const IMG_REMISE = path.join(BASE, "mise-en-main-livraison-convoyeur-vehicule (1).jpg");
const IMG_CONVOYEUR = path.join(BASE, "Convoyeur-professionnel-volant-rapatriement-automobile (1).jpg");

async function main() {
  let pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Conwise Express";
  pres.title = "Proposition Groupe ORTEC - Conwise Express";

  const logoData = imgToBase64(LOGO_CW);
  const logoOrtecData = imgToBase64(LOGO_ORTEC);
  const bannerData = imgToBase64(IMG_BANNER);
  const fleetData = imgToBase64(IMG_FLEET);
  const fleetMgrData = imgToBase64(IMG_FLEET_MGR);
  const edlData = imgToBase64(IMG_EDL);
  const commentData = imgToBase64(IMG_COMMENT);
  const dashboardData = imgToBase64(IMG_DASHBOARD);
  const remiseData = imgToBase64(IMG_REMISE);
  const convoyeurData = imgToBase64(IMG_CONVOYEUR);

  // Pre-render icons
  const iconTruck = await iconToBase64Png(FaTruck, "#FFFFFF", 256);
  const iconShield = await iconToBase64Png(FaShieldAlt, "#FFFFFF", 256);
  const iconChart = await iconToBase64Png(FaChartLine, "#FFFFFF", 256);
  const iconClock = await iconToBase64Png(FaClock, "#FFFFFF", 256);
  const iconMap = await iconToBase64Png(FaMapMarkerAlt, "#FFFFFF", 256);
  const iconUsers = await iconToBase64Png(FaUsers, "#FFFFFF", 256);
  const iconCheck = await iconToBase64Png(FaCheckCircle, "#2ECC71", 256);
  const iconCheckW = await iconToBase64Png(FaCheckCircle, "#FFFFFF", 256);
  const iconLaptop = await iconToBase64Png(FaLaptop, "#FFFFFF", 256);
  const iconPhone = await iconToBase64Png(FaPhone, "#FFFFFF", 256);
  const iconEnvelope = await iconToBase64Png(FaEnvelope, "#FFFFFF", 256);
  const iconGlobe = await iconToBase64Png(FaGlobe, "#FFFFFF", 256);
  const iconIndustry = await iconToBase64Png(FaIndustry, "#FFFFFF", 256);
  const iconRoute = await iconToBase64Png(FaRoute, "#FFFFFF", 256);
  const iconCogs = await iconToBase64Png(FaCogs, "#FFFFFF", 256);
  const iconHandshake = await iconToBase64Png(FaHandshake, "#FFFFFF", 256);
  const iconFile = await iconToBase64Png(FaFileAlt, "#FFFFFF", 256);
  const iconStar = await iconToBase64Png(FaStar, "#FFFFFF", 256);
  const iconArrow = await iconToBase64Png(FaArrowRight, "#1e4b77", 256);
  const iconDash = await iconToBase64Png(MdDashboard, "#FFFFFF", 256);
  const iconSpeed = await iconToBase64Png(MdSpeed, "#FFFFFF", 256);
  const iconCheckNavy = await iconToBase64Png(FaCheckCircle, "#1e4b77", 256);
  const iconTruckNavy = await iconToBase64Png(FaTruck, "#1e4b77", 256);
  const iconShieldNavy = await iconToBase64Png(FaShieldAlt, "#1e4b77", 256);
  const iconClockNavy = await iconToBase64Png(FaClock, "#1e4b77", 256);
  const iconChartNavy = await iconToBase64Png(FaChartLine, "#1e4b77", 256);
  const iconMapNavy = await iconToBase64Png(FaMapMarkerAlt, "#1e4b77", 256);
  const iconRouteNavy = await iconToBase64Png(FaRoute, "#1e4b77", 256);
  const iconIndustryNavy = await iconToBase64Png(FaIndustry, "#1e4b77", 256);
  const iconFileNavy = await iconToBase64Png(FaFileAlt, "#1e4b77", 256);
  const iconLaptopNavy = await iconToBase64Png(FaLaptop, "#1e4b77", 256);
  const iconHandshakeBlue = await iconToBase64Png(FaHandshake, "#3498db", 256);
  const iconStarBlue = await iconToBase64Png(FaStar, "#3498db", 256);

  // ============================================================
  // SLIDE 1 - COVER
  // ============================================================
  let s1 = pres.addSlide();
  s1.background = { color: C.dark };

  // Dark overlay image
  s1.addImage({ data: bannerData, x: 0, y: 0, w: 10, h: 5.625, sizing: { type: "cover", w: 10, h: 5.625 }, transparency: 60 });

  // Dark gradient overlay from bottom
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.dark, transparency: 40 } });

  // Blue accent bar at top
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.blue } });

  // Logos side by side
  s1.addImage({ data: logoData, x: 1.0, y: 0.5, w: 2.2, h: 0.75, sizing: { type: "contain", w: 2.2, h: 0.75 } });
  s1.addText("x", { x: 3.35, y: 0.5, w: 0.4, h: 0.75, fontSize: 18, color: C.lightBlue, fontFace: "Calibri", align: "center", valign: "middle" });
  s1.addImage({ data: logoOrtecData, x: 3.8, y: 0.4, w: 2.0, h: 0.95, sizing: { type: "contain", w: 2.0, h: 0.95 } });

  // Main title block
  s1.addText("PROPOSITION DE SERVICES", { x: 1.0, y: 1.8, w: 8, h: 0.5, fontSize: 14, color: C.lightBlue, fontFace: "Calibri", charSpacing: 6, bold: true, margin: 0 });
  s1.addText("Gestion & Convoyage\nde votre Flotte Automobile", { x: 1.0, y: 2.3, w: 8, h: 1.5, fontSize: 36, color: C.white, fontFace: "Georgia", bold: true, lineSpacingMultiple: 1.1, margin: 0 });

  // Subtitle
  s1.addText("Solutions sur-mesure pour le transfert et le rapatriement\nde vos v\u00e9hicules sur l'ensemble du territoire fran\u00e7ais", { x: 1.0, y: 3.85, w: 7, h: 0.85, fontSize: 14, color: C.lightBlue, fontFace: "Calibri", lineSpacingMultiple: 1.3, margin: 0 });

  // Bottom bar
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.05, w: 10, h: 0.575, fill: { color: C.navy, transparency: 30 } });
  s1.addText("www.conwise-express.com  |  contact@conwise-express.com  |  +33 7 44 31 79 16", { x: 0.5, y: 5.05, w: 9, h: 0.575, fontSize: 10, color: C.lightBlue, fontFace: "Calibri", align: "center", valign: "middle" });

  // ============================================================
  // SLIDE 2 - QUI SOMMES-NOUS
  // ============================================================
  let s2 = pres.addSlide();
  s2.background = { color: C.white };

  // Header band
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy } });
  s2.addImage({ data: logoData, x: 0.5, y: 0.15, w: 1.8, h: 0.7, sizing: { type: "contain", w: 1.8, h: 0.7 } });
  s2.addText("QUI SOMMES-NOUS ?", { x: 2.8, y: 0, w: 6.7, h: 1.0, fontSize: 24, color: C.white, fontFace: "Georgia", bold: true, align: "right", valign: "middle", margin: 0 });

  // Left column - text
  s2.addText("Votre partenaire convoyage", { x: 0.6, y: 1.3, w: 5.0, h: 0.5, fontSize: 18, color: C.navy, fontFace: "Georgia", bold: true, margin: 0 });

  s2.addText([
    { text: "Conwise Express", options: { bold: true, color: C.blue } },
    { text: " est un r\u00e9seau de plus de ", options: { color: C.gray } },
    { text: "40 convoyeurs professionnels certifi\u00e9s", options: { bold: true, color: C.navy } },
    { text: " r\u00e9partis sur toute la France, sp\u00e9cialis\u00e9s dans le transfert, le rapatriement et la livraison de v\u00e9hicules pour les entreprises et gestionnaires de flottes.", options: { color: C.gray } },
  ], { x: 0.6, y: 1.85, w: 5.0, h: 1.2, fontSize: 11.5, fontFace: "Calibri", lineSpacingMultiple: 1.4, margin: 0 });

  // Key stats - 2x2 grid
  const stats = [
    { icon: iconTruck, num: "40+", label: "Convoyeurs\ncertifi\u00e9s" },
    { icon: iconMap, num: "100%", label: "Couverture\nnationale" },
    { icon: iconClock, num: "24-48h", label: "D\u00e9lai de\nlivraison" },
    { icon: iconShield, num: "100%", label: "Assurance\ntous risques" },
  ];
  for (let i = 0; i < 4; i++) {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const bx = 0.6 + col * 2.55;
    const by = 3.2 + row * 1.1;

    s2.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: 2.35, h: 0.95, fill: { color: C.paleBlue }, shadow: makeShadowSoft() });
    // Icon circle
    s2.addShape(pres.shapes.OVAL, { x: bx + 0.12, y: by + 0.17, w: 0.6, h: 0.6, fill: { color: C.navy } });
    s2.addImage({ data: stats[i].icon, x: bx + 0.25, y: by + 0.3, w: 0.35, h: 0.35 });
    s2.addText(stats[i].num, { x: bx + 0.85, y: by + 0.08, w: 1.4, h: 0.4, fontSize: 18, color: C.navy, fontFace: "Georgia", bold: true, margin: 0 });
    s2.addText(stats[i].label, { x: bx + 0.85, y: by + 0.48, w: 1.4, h: 0.42, fontSize: 9, color: C.gray, fontFace: "Calibri", lineSpacingMultiple: 1.1, margin: 0 });
  }

  // Right column - image
  s2.addImage({ data: convoyeurData, x: 5.9, y: 1.2, w: 3.7, h: 4.1, sizing: { type: "cover", w: 3.7, h: 4.1 } });

  // ============================================================
  // SLIDE 3 - COMPRENDRE VOS ENJEUX (ORTEC)
  // ============================================================
  let s3 = pres.addSlide();
  s3.background = { color: C.lightGray };

  // Header
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy } });
  s3.addImage({ data: logoOrtecData, x: 0.5, y: 0.1, w: 1.6, h: 0.8, sizing: { type: "contain", w: 1.6, h: 0.8 } });
  s3.addText("VOS ENJEUX FLOTTE", { x: 2.5, y: 0, w: 7.0, h: 1.0, fontSize: 24, color: C.white, fontFace: "Georgia", bold: true, align: "right", valign: "middle", margin: 0 });

  // Intro text
  s3.addText([
    { text: "Groupe ORTEC", options: { bold: true, color: C.navy } },
    { text: " : 12 500+ collaborateurs, 300+ implantations dans 28 pays.\nVotre activit\u00e9 multi-sites en ing\u00e9nierie industrielle g\u00e9n\u00e8re des besoins complexes de transfert de v\u00e9hicules.", options: { color: C.gray } },
  ], { x: 0.6, y: 1.15, w: 8.8, h: 0.7, fontSize: 11.5, fontFace: "Calibri", lineSpacingMultiple: 1.35, margin: 0 });

  // 3 challenge cards
  const challenges = [
    {
      icon: iconRouteNavy, title: "Mobilit\u00e9 inter-agences",
      desc: "Transferts fr\u00e9quents de v\u00e9hicules entre vos agences r\u00e9parties sur tout le territoire et vos services carrosserie"
    },
    {
      icon: iconIndustryNavy, title: "Diversit\u00e9 des sites",
      desc: "Sites industriels (nucl\u00e9aire, d\u00e9fense, chimie, p\u00e9trole) avec des contraintes d'acc\u00e8s et de s\u00e9curit\u00e9 sp\u00e9cifiques"
    },
    {
      icon: iconChartNavy, title: "Optimisation des co\u00fbts",
      desc: "Ma\u00eetriser le budget flotte tout en garantissant la disponibilit\u00e9 des v\u00e9hicules pour vos \u00e9quipes terrain"
    },
  ];

  for (let i = 0; i < 3; i++) {
    const cx = 0.6 + i * 3.1;
    s3.addShape(pres.shapes.RECTANGLE, { x: cx, y: 2.05, w: 2.85, h: 2.1, fill: { color: C.white }, shadow: makeShadowCard() });
    // Left accent bar
    s3.addShape(pres.shapes.RECTANGLE, { x: cx, y: 2.05, w: 0.06, h: 2.1, fill: { color: C.blue } });
    s3.addImage({ data: challenges[i].icon, x: cx + 0.25, y: 2.25, w: 0.4, h: 0.4 });
    s3.addText(challenges[i].title, { x: cx + 0.25, y: 2.7, w: 2.35, h: 0.35, fontSize: 13, color: C.navy, fontFace: "Calibri", bold: true, margin: 0 });
    s3.addText(challenges[i].desc, { x: cx + 0.25, y: 3.05, w: 2.35, h: 0.9, fontSize: 10, color: C.gray, fontFace: "Calibri", lineSpacingMultiple: 1.3, margin: 0 });
  }

  // Bottom section - specific ORTEC needs
  s3.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.35, w: 8.8, h: 1.05, fill: { color: C.navy } });
  s3.addText("Besoins identifi\u00e9s pour ORTEC", { x: 0.9, y: 4.38, w: 4, h: 0.35, fontSize: 12, color: C.lightBlue, fontFace: "Calibri", bold: true, margin: 0 });

  const needs = [
    "Transferts v\u00e9hicules entre agences France",
    "Acheminement vers services carrosserie",
    "Rapatriement apr\u00e8s r\u00e9paration",
    "Gestion centralis\u00e9e des commandes multi-sites",
  ];
  for (let i = 0; i < 4; i++) {
    const nx = 0.9 + (i % 2) * 4.3;
    const ny = 4.75 + Math.floor(i / 2) * 0.28;
    s3.addImage({ data: iconCheckW, x: nx, y: ny, w: 0.18, h: 0.18 });
    s3.addText(needs[i], { x: nx + 0.25, y: ny - 0.02, w: 3.9, h: 0.25, fontSize: 10, color: C.white, fontFace: "Calibri", margin: 0 });
  }

  // ============================================================
  // SLIDE 4 - CAS CLIENT DERICHEBOURG (expertise proof)
  // ============================================================
  let s4 = pres.addSlide();
  s4.background = { color: C.white };

  // Header
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy } });
  s4.addImage({ data: logoData, x: 0.5, y: 0.15, w: 1.8, h: 0.7, sizing: { type: "contain", w: 1.8, h: 0.7 } });
  s4.addText("NOTRE EXPERTISE GRANDS GROUPES", { x: 2.8, y: 0, w: 6.7, h: 1.0, fontSize: 22, color: C.white, fontFace: "Georgia", bold: true, align: "right", valign: "middle", margin: 0 });

  // Badge "Cas client"
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.25, w: 1.6, h: 0.35, fill: { color: C.blue } });
  s4.addText("CAS CLIENT", { x: 0.6, y: 1.25, w: 1.6, h: 0.35, fontSize: 10, color: C.white, fontFace: "Calibri", bold: true, align: "center", valign: "middle", charSpacing: 3 });

  s4.addText("Derichebourg Multiservices", { x: 0.6, y: 1.75, w: 5, h: 0.45, fontSize: 22, color: C.navy, fontFace: "Georgia", bold: true, margin: 0 });

  // Left column - details
  s4.addText([
    { text: "48 000 collaborateurs", options: { bold: true, color: C.navy } },
    { text: " | ", options: { color: C.gray } },
    { text: "150 implantations", options: { bold: true, color: C.navy } },
    { text: " | ", options: { color: C.gray } },
    { text: "8 pays", options: { bold: true, color: C.navy } },
  ], { x: 0.6, y: 2.2, w: 5.0, h: 0.35, fontSize: 11, fontFace: "Calibri", margin: 0 });

  s4.addText("Un profil similaire \u00e0 ORTEC : grand groupe multi-sites, \u00e9quipes mobiles r\u00e9parties sur tout le territoire, flotte cons\u00e9quente n\u00e9cessitant une gestion centralis\u00e9e.", { x: 0.6, y: 2.6, w: 5.0, h: 0.7, fontSize: 11, color: C.gray, fontFace: "Calibri", lineSpacingMultiple: 1.35, margin: 0 });

  // What we do for them
  s4.addText("Ce que nous g\u00e9rons pour Derichebourg :", { x: 0.6, y: 3.35, w: 5.0, h: 0.3, fontSize: 12, color: C.navy, fontFace: "Calibri", bold: true, margin: 0 });

  const derichItems = [
    "Transferts inter-sites sur toute la France",
    "Rapatriement de v\u00e9hicules sinistri\u00e9s ou en maintenance",
    "Livraison de v\u00e9hicules neufs aux agences",
    "Gestion centralis\u00e9e via portail Conwise d\u00e9di\u00e9",
    "Tra\u00e7abilit\u00e9 compl\u00e8te et PV digitalis\u00e9s",
  ];
  for (let i = 0; i < derichItems.length; i++) {
    const iy = 3.72 + i * 0.32;
    s4.addImage({ data: iconCheckNavy, x: 0.7, y: iy, w: 0.2, h: 0.2 });
    s4.addText(derichItems[i], { x: 1.0, y: iy - 0.02, w: 4.5, h: 0.28, fontSize: 10.5, color: C.dark, fontFace: "Calibri", margin: 0 });
  }

  // Right side - image + quote
  s4.addImage({ data: fleetData, x: 5.9, y: 1.25, w: 3.7, h: 2.5, sizing: { type: "cover", w: 3.7, h: 2.5 } });

  // Testimonial box
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.9, y: 3.95, w: 3.7, h: 1.35, fill: { color: C.paleBlue }, shadow: makeShadowSoft() });
  s4.addText("Cette expertise multi-sites est directement transposable \u00e0 votre organisation ORTEC.", { x: 6.1, y: 4.05, w: 3.3, h: 0.65, fontSize: 11, color: C.navy, fontFace: "Georgia", italic: true, lineSpacingMultiple: 1.3, margin: 0 });
  s4.addImage({ data: iconArrow, x: 6.1, y: 4.75, w: 0.2, h: 0.2 });
  s4.addText("M\u00eame envergure, m\u00eames enjeux, solutions \u00e9prouv\u00e9es", { x: 6.4, y: 4.72, w: 3.0, h: 0.3, fontSize: 10, color: C.blue, fontFace: "Calibri", bold: true, margin: 0 });

  // ============================================================
  // SLIDE 5 - NOS SERVICES
  // ============================================================
  let s5 = pres.addSlide();
  s5.background = { color: C.lightGray };

  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy } });
  s5.addImage({ data: logoData, x: 0.5, y: 0.15, w: 1.8, h: 0.7, sizing: { type: "contain", w: 1.8, h: 0.7 } });
  s5.addText("NOS SERVICES POUR ORTEC", { x: 2.8, y: 0, w: 6.7, h: 1.0, fontSize: 24, color: C.white, fontFace: "Georgia", bold: true, align: "right", valign: "middle", margin: 0 });

  // 6 service cards in 3x2 grid
  const services = [
    { icon: iconTruckNavy, title: "Convoyage VL", desc: "Transfert de v\u00e9hicules l\u00e9gers entre vos agences, ateliers et services carrosserie" },
    { icon: iconRouteNavy, title: "Rapatriement", desc: "R\u00e9cup\u00e9ration de v\u00e9hicules sinistri\u00e9s, en panne ou en fin de contrat" },
    { icon: iconFileNavy, title: "\u00c9tat des lieux digital", desc: "PV photo d\u00e9taill\u00e9 au d\u00e9part et \u00e0 l'arriv\u00e9e, sign\u00e9 \u00e9lectroniquement" },
    { icon: iconMapNavy, title: "Suivi GPS temps r\u00e9el", desc: "Tra\u00e7abilit\u00e9 compl\u00e8te de chaque mission sur votre portail d\u00e9di\u00e9" },
    { icon: iconLaptopNavy, title: "Portail Fleet Manager", desc: "Interface de commande et pilotage centralis\u00e9 de toutes vos op\u00e9rations" },
    { icon: iconShieldNavy, title: "Assurance tous risques", desc: "Couverture compl\u00e8te de chaque v\u00e9hicule pendant toute la dur\u00e9e du convoyage" },
  ];

  for (let i = 0; i < 6; i++) {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const sx = 0.5 + col * 3.1;
    const sy = 1.25 + row * 2.05;

    s5.addShape(pres.shapes.RECTANGLE, { x: sx, y: sy, w: 2.85, h: 1.8, fill: { color: C.white }, shadow: makeShadowCard() });
    // Icon circle
    s5.addShape(pres.shapes.OVAL, { x: sx + 0.2, y: sy + 0.2, w: 0.55, h: 0.55, fill: { color: C.paleBlue } });
    s5.addImage({ data: services[i].icon, x: sx + 0.3, y: sy + 0.3, w: 0.35, h: 0.35 });
    s5.addText(services[i].title, { x: sx + 0.9, y: sy + 0.25, w: 1.75, h: 0.35, fontSize: 13, color: C.navy, fontFace: "Calibri", bold: true, valign: "middle", margin: 0 });
    s5.addText(services[i].desc, { x: sx + 0.2, y: sy + 0.85, w: 2.45, h: 0.8, fontSize: 10, color: C.gray, fontFace: "Calibri", lineSpacingMultiple: 1.3, margin: 0 });
  }

  // ============================================================
  // SLIDE 6 - COMMENT CA MARCHE
  // ============================================================
  let s6 = pres.addSlide();
  s6.background = { color: C.white };

  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy } });
  s6.addImage({ data: logoData, x: 0.5, y: 0.15, w: 1.8, h: 0.7, sizing: { type: "contain", w: 1.8, h: 0.7 } });
  s6.addText("COMMENT \u00c7A MARCHE ?", { x: 2.8, y: 0, w: 6.7, h: 1.0, fontSize: 24, color: C.white, fontFace: "Georgia", bold: true, align: "right", valign: "middle", margin: 0 });

  // 4 steps horizontal
  const steps = [
    { num: "01", title: "Commande", desc: "Passez votre commande en ligne via le portail Conwise ou par t\u00e9l\u00e9phone", icon: iconLaptopNavy },
    { num: "02", title: "Attribution", desc: "Un convoyeur certifi\u00e9 est assign\u00e9 sous 2h. Vous \u00eates notifi\u00e9 en temps r\u00e9el", icon: iconTruckNavy },
    { num: "03", title: "Convoyage", desc: "\u00c9tat des lieux digital au d\u00e9part, suivi GPS en direct, livraison sous 24-48h", icon: iconRouteNavy },
    { num: "04", title: "Livraison", desc: "Remise en main, PV de livraison sign\u00e9, facture automatis\u00e9e", icon: iconFileNavy },
  ];

  for (let i = 0; i < 4; i++) {
    const stx = 0.4 + i * 2.4;

    // Step number circle
    s6.addShape(pres.shapes.OVAL, { x: stx + 0.55, y: 1.25, w: 0.7, h: 0.7, fill: { color: C.navy } });
    s6.addText(steps[i].num, { x: stx + 0.55, y: 1.25, w: 0.7, h: 0.7, fontSize: 18, color: C.white, fontFace: "Georgia", bold: true, align: "center", valign: "middle" });

    // Connecting line (except last)
    if (i < 3) {
      s6.addShape(pres.shapes.LINE, { x: stx + 1.35, y: 1.6, w: 1.3, h: 0, line: { color: C.lightBlue, width: 2, dashType: "dash" } });
    }

    s6.addText(steps[i].title, { x: stx, y: 2.1, w: 1.9, h: 0.35, fontSize: 14, color: C.navy, fontFace: "Calibri", bold: true, align: "center", margin: 0 });
    s6.addText(steps[i].desc, { x: stx, y: 2.5, w: 1.9, h: 0.85, fontSize: 9.5, color: C.gray, fontFace: "Calibri", align: "center", lineSpacingMultiple: 1.3, margin: 0 });
  }

  // Bottom - image split
  s6.addImage({ data: edlData, x: 0, y: 3.55, w: 5, h: 2.075, sizing: { type: "cover", w: 5, h: 2.075 } });
  s6.addImage({ data: remiseData, x: 5, y: 3.55, w: 5, h: 2.075, sizing: { type: "cover", w: 5, h: 2.075 } });

  // Overlay labels
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.0, w: 2.2, h: 0.35, fill: { color: C.navy, transparency: 20 } });
  s6.addText("\u00c9tat des lieux digital", { x: 0.3, y: 5.0, w: 2.2, h: 0.35, fontSize: 9, color: C.white, fontFace: "Calibri", bold: true, align: "center", valign: "middle" });
  s6.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 5.0, w: 2.2, h: 0.35, fill: { color: C.navy, transparency: 20 } });
  s6.addText("Remise en main", { x: 5.3, y: 5.0, w: 2.2, h: 0.35, fontSize: 9, color: C.white, fontFace: "Calibri", bold: true, align: "center", valign: "middle" });

  // ============================================================
  // SLIDE 7 - PORTAIL DE GESTION / DASHBOARD
  // ============================================================
  let s7 = pres.addSlide();
  s7.background = { color: C.lightGray };

  s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy } });
  s7.addImage({ data: logoData, x: 0.5, y: 0.15, w: 1.8, h: 0.7, sizing: { type: "contain", w: 1.8, h: 0.7 } });
  s7.addText("VOTRE PORTAIL DE GESTION", { x: 2.8, y: 0, w: 6.7, h: 1.0, fontSize: 24, color: C.white, fontFace: "Georgia", bold: true, align: "right", valign: "middle", margin: 0 });

  s7.addText("Un environnement d\u00e9di\u00e9 Groupe ORTEC sur conwise.app", { x: 0.6, y: 1.15, w: 8.8, h: 0.35, fontSize: 13, color: C.navy, fontFace: "Calibri", bold: true, margin: 0 });

  // Dashboard screenshot
  s7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.6, w: 5.8, h: 3.5, fill: { color: C.white }, shadow: makeShadowCard() });
  s7.addImage({ data: dashboardData, x: 0.6, y: 1.7, w: 5.6, h: 3.3, sizing: { type: "contain", w: 5.6, h: 3.3 } });

  // Right side - features
  s7.addText("Fonctionnalit\u00e9s cl\u00e9s", { x: 6.6, y: 1.6, w: 3.0, h: 0.35, fontSize: 14, color: C.navy, fontFace: "Georgia", bold: true, margin: 0 });

  const dashFeatures = [
    { icon: iconLaptopNavy, text: "Commande en ligne simplifi\u00e9e" },
    { icon: iconChartNavy, text: "Tableau de bord et statistiques" },
    { icon: iconRouteNavy, text: "Suivi en temps r\u00e9el des missions" },
    { icon: iconFileNavy, text: "PV et factures centralis\u00e9s" },
    { icon: iconMapNavy, text: "Gestion de vos adresses favorites" },
    { icon: iconClockNavy, text: "Historique complet des convoyages" },
  ];
  for (let i = 0; i < dashFeatures.length; i++) {
    const fy = 2.1 + i * 0.5;
    s7.addImage({ data: dashFeatures[i].icon, x: 6.6, y: fy, w: 0.28, h: 0.28 });
    s7.addText(dashFeatures[i].text, { x: 7.0, y: fy, w: 2.6, h: 0.28, fontSize: 10.5, color: C.dark, fontFace: "Calibri", valign: "middle", margin: 0 });
  }

  // CTA box
  s7.addShape(pres.shapes.RECTANGLE, { x: 6.6, y: 5.0, w: 3.0, h: 0.4, fill: { color: C.blue } });
  s7.addText("conwise.app  \u2192  Acc\u00e8s portail ORTEC", { x: 6.6, y: 5.0, w: 3.0, h: 0.4, fontSize: 10, color: C.white, fontFace: "Calibri", bold: true, align: "center", valign: "middle" });

  // ============================================================
  // SLIDE 8 - POURQUOI CONWISE (Avantages)
  // ============================================================
  let s8 = pres.addSlide();
  s8.background = { color: C.white };

  s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy } });
  s8.addImage({ data: logoData, x: 0.5, y: 0.15, w: 1.8, h: 0.7, sizing: { type: "contain", w: 1.8, h: 0.7 } });
  s8.addText("POURQUOI CONWISE EXPRESS ?", { x: 2.8, y: 0, w: 6.7, h: 1.0, fontSize: 24, color: C.white, fontFace: "Georgia", bold: true, align: "right", valign: "middle", margin: 0 });

  // 3 big advantage columns
  const advantages = [
    {
      iconBg: C.navy, icon: iconHandshake, title: "Partenariat\ngagnant-gagnant",
      items: ["Groupement d'employeurs structur\u00e9", "Tarifs avantageux et transparents", "Pas de frais cach\u00e9s", "Un interlocuteur d\u00e9di\u00e9"]
    },
    {
      iconBg: C.blue, icon: iconSpeed, title: "Performance\nop\u00e9rationnelle",
      items: ["Livraison sous 24-48h", "R\u00e9activit\u00e9 et flexibilit\u00e9", "Convoyeurs certifi\u00e9s et assur\u00e9s", "Couverture nationale compl\u00e8te"]
    },
    {
      iconBg: C.lightBlue, icon: iconDash, title: "Technologie\n& digitalisation",
      items: ["Portail de gestion personnalis\u00e9", "PV digitaux avec photos", "Suivi GPS temps r\u00e9el", "Facturation automatis\u00e9e"]
    },
  ];

  for (let i = 0; i < 3; i++) {
    const ax = 0.45 + i * 3.15;

    // Card
    s8.addShape(pres.shapes.RECTANGLE, { x: ax, y: 1.2, w: 2.9, h: 4.15, fill: { color: C.lightGray }, shadow: makeShadowCard() });

    // Icon circle at top
    s8.addShape(pres.shapes.OVAL, { x: ax + 0.95, y: 1.4, w: 1.0, h: 1.0, fill: { color: advantages[i].iconBg } });
    s8.addImage({ data: advantages[i].icon, x: ax + 1.15, y: 1.6, w: 0.6, h: 0.6 });

    // Title
    s8.addText(advantages[i].title, { x: ax + 0.15, y: 2.55, w: 2.6, h: 0.65, fontSize: 13, color: C.navy, fontFace: "Georgia", bold: true, align: "center", lineSpacingMultiple: 1.15, margin: 0 });

    // Items
    for (let j = 0; j < advantages[i].items.length; j++) {
      const iy = 3.35 + j * 0.45;
      s8.addImage({ data: iconCheck, x: ax + 0.2, y: iy + 0.02, w: 0.2, h: 0.2 });
      s8.addText(advantages[i].items[j], { x: ax + 0.5, y: iy, w: 2.2, h: 0.4, fontSize: 10, color: C.dark, fontFace: "Calibri", valign: "middle", margin: 0 });
    }
  }

  // ============================================================
  // SLIDE 9 - NOS REFERENCES
  // ============================================================
  let s9 = pres.addSlide();
  s9.background = { color: C.lightGray };

  s9.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy } });
  s9.addImage({ data: logoData, x: 0.5, y: 0.15, w: 1.8, h: 0.7, sizing: { type: "contain", w: 1.8, h: 0.7 } });
  s9.addText("NOS R\u00c9F\u00c9RENCES", { x: 2.8, y: 0, w: 6.7, h: 1.0, fontSize: 24, color: C.white, fontFace: "Georgia", bold: true, align: "right", valign: "middle", margin: 0 });

  s9.addText("Des entreprises de premier plan nous font confiance", { x: 0.6, y: 1.2, w: 8.8, h: 0.4, fontSize: 14, color: C.navy, fontFace: "Calibri", bold: true, align: "center", margin: 0 });

  // Reference cards
  const refs = [
    { name: "Derichebourg", sector: "Multiservices", detail: "48 000 collaborateurs\n150 implantations\nGestion compl\u00e8te flotte France" },
    { name: "Elior", sector: "Restauration collective", detail: "Groupe international\nMultiples sites France\nTransferts v\u00e9hicules de service" },
    { name: "Mondial Relay", sector: "Logistique", detail: "R\u00e9seau national\nFlotte de livraison\nConvoyage r\u00e9gulier" },
    { name: "Peugeot Spoticar", sector: "Distribution automobile", detail: "R\u00e9seau de concessions\nTransferts inter-sites\nLivraison clients" },
  ];

  for (let i = 0; i < 4; i++) {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const rx = 0.6 + col * 4.6;
    const ry = 1.85 + row * 1.65;

    s9.addShape(pres.shapes.RECTANGLE, { x: rx, y: ry, w: 4.2, h: 1.4, fill: { color: C.white }, shadow: makeShadowSoft() });
    s9.addShape(pres.shapes.RECTANGLE, { x: rx, y: ry, w: 0.06, h: 1.4, fill: { color: C.blue } });

    s9.addText(refs[i].name, { x: rx + 0.25, y: ry + 0.1, w: 3.7, h: 0.3, fontSize: 14, color: C.navy, fontFace: "Georgia", bold: true, margin: 0 });

    // Sector badge
    s9.addShape(pres.shapes.RECTANGLE, { x: rx + 0.25, y: ry + 0.42, w: 1.8, h: 0.22, fill: { color: C.paleBlue } });
    s9.addText(refs[i].sector, { x: rx + 0.25, y: ry + 0.42, w: 1.8, h: 0.22, fontSize: 8, color: C.blue, fontFace: "Calibri", bold: true, align: "center", valign: "middle" });

    s9.addText(refs[i].detail, { x: rx + 0.25, y: ry + 0.72, w: 3.7, h: 0.6, fontSize: 9, color: C.gray, fontFace: "Calibri", lineSpacingMultiple: 1.25, margin: 0 });
  }

  // Bottom statement
  s9.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 5.0, w: 8.8, h: 0.45, fill: { color: C.navy } });
  s9.addText("Des solutions \u00e9prouv\u00e9es aupr\u00e8s de grands groupes multi-sites  \u2014  La m\u00eame expertise pour ORTEC", { x: 0.6, y: 5.0, w: 8.8, h: 0.45, fontSize: 11, color: C.white, fontFace: "Calibri", bold: true, align: "center", valign: "middle" });

  // ============================================================
  // SLIDE 10 - PROCHAINES ETAPES / CONTACT
  // ============================================================
  let s10 = pres.addSlide();
  s10.background = { color: C.dark };
  s10.addImage({ data: bannerData, x: 0, y: 0, w: 10, h: 5.625, sizing: { type: "cover", w: 10, h: 5.625 }, transparency: 70 });
  s10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.dark, transparency: 30 } });

  // Blue accent bar at top
  s10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.blue } });

  // Logos
  s10.addImage({ data: logoData, x: 3.0, y: 0.4, w: 1.8, h: 0.65, sizing: { type: "contain", w: 1.8, h: 0.65 } });
  s10.addText("x", { x: 4.9, y: 0.4, w: 0.3, h: 0.65, fontSize: 16, color: C.lightBlue, fontFace: "Calibri", align: "center", valign: "middle" });
  s10.addImage({ data: logoOrtecData, x: 5.3, y: 0.3, w: 1.8, h: 0.85, sizing: { type: "contain", w: 1.8, h: 0.85 } });

  // Main title
  s10.addText("PROCHAINES \u00c9TAPES", { x: 1, y: 1.4, w: 8, h: 0.6, fontSize: 30, color: C.white, fontFace: "Georgia", bold: true, align: "center", margin: 0 });

  // Steps
  const nextSteps = [
    { num: "1", text: "D\u00e9finition de vos besoins et volumes pr\u00e9visionnels" },
    { num: "2", text: "Proposition tarifaire personnalis\u00e9e Groupe ORTEC" },
    { num: "3", text: "Configuration de votre portail Conwise d\u00e9di\u00e9" },
    { num: "4", text: "Phase pilote sur un p\u00e9rim\u00e8tre de test" },
  ];

  for (let i = 0; i < 4; i++) {
    const nsy = 2.2 + i * 0.55;
    s10.addShape(pres.shapes.OVAL, { x: 2.5, y: nsy, w: 0.4, h: 0.4, fill: { color: C.blue } });
    s10.addText(nextSteps[i].num, { x: 2.5, y: nsy, w: 0.4, h: 0.4, fontSize: 14, color: C.white, fontFace: "Georgia", bold: true, align: "center", valign: "middle" });
    s10.addText(nextSteps[i].text, { x: 3.1, y: nsy, w: 5, h: 0.4, fontSize: 13, color: C.white, fontFace: "Calibri", valign: "middle", margin: 0 });
  }

  // Contact card
  s10.addShape(pres.shapes.RECTANGLE, { x: 2.0, y: 4.35, w: 6, h: 1.05, fill: { color: C.navy, transparency: 20 } });
  s10.addText("Contactez-nous", { x: 2.0, y: 4.35, w: 6, h: 0.35, fontSize: 14, color: C.lightBlue, fontFace: "Georgia", bold: true, align: "center", valign: "middle" });

  s10.addText("contact@conwise-express.com   |   +33 7 44 31 79 16   |   www.conwise-express.com", { x: 2.0, y: 4.7, w: 6, h: 0.3, fontSize: 11, color: C.white, fontFace: "Calibri", align: "center", valign: "middle" });

  s10.addText("12 Cours des Merveilles, 95000 Cergy", { x: 2.0, y: 5.0, w: 6, h: 0.25, fontSize: 10, color: C.lightBlue, fontFace: "Calibri", align: "center", valign: "middle" });

  // === WRITE FILE ===
  const outputPath = path.join(BASE, "Proposition Conwise Express x Groupe ORTEC.pptx");
  await pres.writeFile({ fileName: outputPath });
  console.log("Presentation saved to:", outputPath);
}

main().catch(err => { console.error(err); process.exit(1); });
