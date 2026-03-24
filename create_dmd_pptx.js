const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Icon rendering utilities
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

// Load image as base64
function imgToBase64(filePath) {
  const data = fs.readFileSync(filePath);
  const ext = path.extname(filePath).toLowerCase();
  const mime = ext === ".png" ? "image/png" : "image/jpeg";
  return `${mime};base64,${data.toString("base64")}`;
}

async function createPresentation() {
  // Load icons
  const { FaTruck, FaMapMarkerAlt, FaClipboardCheck, FaMobileAlt, FaShieldAlt, FaClock, FaChartLine, FaUsers, FaHandshake, FaPhone, FaEnvelope, FaGlobe, FaCheckCircle, FaStar, FaCogs, FaFileInvoiceDollar, FaRoute, FaWarehouse, FaCarSide, FaTools } = require("react-icons/fa");

  // ============ BRAND COLORS ============
  const C = {
    primary: "1E4B77",    // Bleu Corporate
    secondary: "3498DB",  // Bleu Principal
    accent: "59B3E6",     // Bleu Clair
    light: "E3F2FD",      // Bleu Pastel
    dark: "0F172A",       // Noir Profond
    gray: "64748B",       // Gris Neutre
    white: "FFFFFF",
    offWhite: "F8FAFC",
    lightGray: "E2E8F0",
    darkText: "1E293B",
    success: "16A34A",
    orange: "F97316",
  };

  // Helper: factory functions for reusable options (avoid mutation issues)
  const makeShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.12 });
  const makeShadowLight = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.08 });

  // Load images
  const baseDir = "C:\\Users\\User\\Documents\\claude code\\Conwise express propal";
  const logoConwise = imgToBase64(path.join(baseDir, "Logo conwise express.png"));
  const logoDMD = imgToBase64(path.join(baseDir, "DMD GROUPE AUTO logo.png"));
  const imgBanner = imgToBase64(path.join(baseDir, "image fond banniere couverture conwise express.png"));
  const imgFleet = imgToBase64(path.join(baseDir, "image gestionnaire de flotte entreprise.png"));
  const imgEtatLieux = imgToBase64(path.join(baseDir, "image homme realise un etat des lieux.png"));
  const imgDashboard = imgToBase64(path.join(baseDir, "dashboard  Conwise Fleet manager pour DMD groupe.png"));
  const imgConvoyeur = imgToBase64(path.join(baseDir, "Convoyeur-professionnel-volant-rapatriement-automobile (1).jpg"));
  const imgRemise = imgToBase64(path.join(baseDir, "mise-en-main-livraison-convoyeur-vehicule (1).jpg"));
  const imgFlotte = imgToBase64(path.join(baseDir, "flotte-automobile-entreprise-rapatriement- transfert-de-vehicules (1).jpeg"));

  // Pre-render icons
  const iconTruck = await iconToBase64Png(FaTruck, "#FFFFFF", 256);
  const iconMap = await iconToBase64Png(FaMapMarkerAlt, "#FFFFFF", 256);
  const iconClipboard = await iconToBase64Png(FaClipboardCheck, "#FFFFFF", 256);
  const iconMobile = await iconToBase64Png(FaMobileAlt, "#FFFFFF", 256);
  const iconShield = await iconToBase64Png(FaShieldAlt, "#FFFFFF", 256);
  const iconClock = await iconToBase64Png(FaClock, "#FFFFFF", 256);
  const iconChart = await iconToBase64Png(FaChartLine, "#FFFFFF", 256);
  const iconUsers = await iconToBase64Png(FaUsers, "#FFFFFF", 256);
  const iconHandshake = await iconToBase64Png(FaHandshake, "#FFFFFF", 256);
  const iconCheck = await iconToBase64Png(FaCheckCircle, "#16A34A", 256);
  const iconStar = await iconToBase64Png(FaStar, "#F97316", 256);
  const iconCogs = await iconToBase64Png(FaCogs, "#FFFFFF", 256);
  const iconRoute = await iconToBase64Png(FaRoute, "#FFFFFF", 256);
  const iconCar = await iconToBase64Png(FaCarSide, "#FFFFFF", 256);
  const iconTools = await iconToBase64Png(FaTools, "#FFFFFF", 256);
  const iconCheckW = await iconToBase64Png(FaCheckCircle, "#FFFFFF", 256);
  const iconPhone = await iconToBase64Png(FaPhone, "#FFFFFF", 256);
  const iconEnvelope = await iconToBase64Png(FaEnvelope, "#FFFFFF", 256);
  const iconGlobe = await iconToBase64Png(FaGlobe, "#FFFFFF", 256);

  // Blue-tinted icons for light backgrounds
  const iconTruckBlue = await iconToBase64Png(FaTruck, "#1E4B77", 256);
  const iconMapBlue = await iconToBase64Png(FaMapMarkerAlt, "#1E4B77", 256);
  const iconClipboardBlue = await iconToBase64Png(FaClipboardCheck, "#1E4B77", 256);
  const iconMobileBlue = await iconToBase64Png(FaMobileAlt, "#1E4B77", 256);
  const iconShieldBlue = await iconToBase64Png(FaShieldAlt, "#1E4B77", 256);
  const iconClockBlue = await iconToBase64Png(FaClock, "#1E4B77", 256);
  const iconChartBlue = await iconToBase64Png(FaChartLine, "#1E4B77", 256);
  const iconUsersBlue = await iconToBase64Png(FaUsers, "#1E4B77", 256);
  const iconFileBlue = await iconToBase64Png(FaFileInvoiceDollar, "#1E4B77", 256);
  const iconRouteBlue = await iconToBase64Png(FaRoute, "#1E4B77", 256);
  const iconWarehouse = await iconToBase64Png(FaWarehouse, "#1E4B77", 256);

  // ============ CREATE PRESENTATION ============
  let pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Conwise Express";
  pres.title = "Proposition Conwise Express - Groupe DMD";

  // =============================================
  // SLIDE 1: COVER - Title Slide
  // =============================================
  let slide1 = pres.addSlide();
  slide1.background = { color: C.dark };

  // Background image with overlay
  slide1.addImage({ data: imgBanner, x: 0, y: 0, w: 10, h: 5.625, sizing: { type: "cover", w: 10, h: 5.625 } });
  // Dark overlay
  slide1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.dark, transparency: 30 } });

  // Gradient bottom bar
  slide1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.0, w: 10, h: 1.625, fill: { color: C.primary, transparency: 15 } });

  // Logos side by side
  slide1.addImage({ data: logoConwise, x: 1.5, y: 0.5, w: 2.2, h: 0.9, sizing: { type: "contain", w: 2.2, h: 0.9 } });
  slide1.addText("x", { x: 3.9, y: 0.55, w: 0.5, h: 0.7, fontSize: 28, color: C.accent, fontFace: "Arial", align: "center", valign: "middle" });
  slide1.addImage({ data: logoDMD, x: 4.5, y: 0.45, w: 2.2, h: 1.0, sizing: { type: "contain", w: 2.2, h: 1.0 } });

  // Title
  slide1.addText("Votre partenaire convoyage", {
    x: 0.8, y: 1.8, w: 8.4, h: 0.8,
    fontSize: 38, fontFace: "Calibri", color: C.white, bold: true, align: "left", margin: 0
  });
  slide1.addText("pour le Groupe DMD", {
    x: 0.8, y: 2.5, w: 8.4, h: 0.7,
    fontSize: 34, fontFace: "Calibri", color: C.accent, bold: true, align: "left", margin: 0
  });

  // Subtitle
  slide1.addText("Solution de convoyage et gestion de flotte sur-mesure\npour vos 53 concessions dans le Grand Ouest", {
    x: 0.8, y: 3.3, w: 7, h: 0.8,
    fontSize: 15, fontFace: "Calibri", color: C.white, align: "left", lineSpacingMultiple: 1.3, margin: 0
  });

  // Bottom info bar
  slide1.addText("Mars 2026  |  Proposition commerciale confidentielle", {
    x: 0.8, y: 4.9, w: 8.4, h: 0.4,
    fontSize: 11, fontFace: "Calibri", color: C.accent, align: "left", margin: 0
  });

  // =============================================
  // SLIDE 2: SOMMAIRE
  // =============================================
  let slide2 = pres.addSlide();
  slide2.background = { color: C.offWhite };

  // Left colored band
  slide2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: C.secondary } });

  // Title
  slide2.addText("SOMMAIRE", {
    x: 0.6, y: 0.3, w: 4, h: 0.6,
    fontSize: 30, fontFace: "Calibri", color: C.primary, bold: true, margin: 0
  });

  // Line under title
  slide2.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 0.9, w: 2.0, h: 0.04, fill: { color: C.secondary } });

  const sommaire = [
    { num: "01", title: "Conwise Express en bref", icon: iconTruckBlue },
    { num: "02", title: "Comprendre vos enjeux", icon: iconWarehouse },
    { num: "03", title: "Notre solution pour le Groupe DMD", icon: iconRouteBlue },
    { num: "04", title: "La plateforme conwise.app", icon: iconMobileBlue },
    { num: "05", title: "Nos engagements qualit\u00e9", icon: iconShieldBlue },
    { num: "06", title: "R\u00e9f\u00e9rences & t\u00e9moignages", icon: iconUsersBlue },
    { num: "07", title: "Offre & prochaines \u00e9tapes", icon: iconChartBlue },
  ];

  sommaire.forEach((item, i) => {
    const yPos = 1.3 + i * 0.55;
    // Number circle
    slide2.addShape(pres.shapes.OVAL, { x: 0.6, y: yPos, w: 0.4, h: 0.4, fill: { color: C.primary } });
    slide2.addText(item.num, { x: 0.6, y: yPos, w: 0.4, h: 0.4, fontSize: 12, fontFace: "Calibri", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    // Title
    slide2.addText(item.title, { x: 1.2, y: yPos, w: 5, h: 0.4, fontSize: 16, fontFace: "Calibri", color: C.darkText, align: "left", valign: "middle", margin: 0 });
  });

  // Right image
  slide2.addImage({ data: imgFleet, x: 6.2, y: 0.8, w: 3.5, h: 4.2, sizing: { type: "cover", w: 3.5, h: 4.2 }, rounding: false });
  // Soft overlay
  slide2.addShape(pres.shapes.RECTANGLE, { x: 6.2, y: 0.8, w: 3.5, h: 4.2, fill: { color: C.primary, transparency: 80 } });

  // Logo bottom right
  slide2.addImage({ data: logoConwise, x: 7.8, y: 5.0, w: 1.8, h: 0.5, sizing: { type: "contain", w: 1.8, h: 0.5 } });

  // =============================================
  // SLIDE 3: QUI SOMMES-NOUS
  // =============================================
  let slide3 = pres.addSlide();
  slide3.background = { color: C.white };

  // Top banner
  slide3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  slide3.addText("01  |  CONWISE EXPRESS EN BREF", {
    x: 0.5, y: 0, w: 9, h: 0.9,
    fontSize: 20, fontFace: "Calibri", color: C.white, bold: true, align: "left", valign: "middle", margin: 0
  });

  // Main content - left side
  slide3.addText("Le r\u00e9seau national de\nconvoyage automobile", {
    x: 0.5, y: 1.1, w: 5.5, h: 0.8,
    fontSize: 22, fontFace: "Calibri", color: C.primary, bold: true, margin: 0
  });

  slide3.addText("Conwise Express f\u00e9d\u00e8re un r\u00e9seau de plus de 40 convoyeurs certifi\u00e9s r\u00e9partis sur toute la France, op\u00e9rant sur le territoire national et en Europe.", {
    x: 0.5, y: 1.95, w: 5.5, h: 0.7,
    fontSize: 13, fontFace: "Calibri", color: C.gray, lineSpacingMultiple: 1.3, margin: 0
  });

  // Key stats - 2x2 grid
  const stats = [
    { value: "40+", label: "Convoyeurs\ncertifi\u00e9s", icon: iconUsers, bgColor: C.primary },
    { value: "48h", label: "D\u00e9lai de\nlivraison", icon: iconClock, bgColor: C.secondary },
    { value: "100%", label: "Couverture\nnationale", icon: iconMap, bgColor: C.accent },
    { value: "0%", label: "Commission\nconvoyeurs", icon: iconHandshake, bgColor: C.primary },
  ];

  stats.forEach((stat, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const xPos = 0.5 + col * 2.9;
    const yPos = 2.85 + row * 1.25;

    // Card background
    slide3.addShape(pres.shapes.RECTANGLE, {
      x: xPos, y: yPos, w: 2.7, h: 1.1,
      fill: { color: C.offWhite },
      shadow: makeShadowLight()
    });
    // Icon circle
    slide3.addShape(pres.shapes.OVAL, { x: xPos + 0.15, y: yPos + 0.2, w: 0.55, h: 0.55, fill: { color: stat.bgColor } });
    slide3.addImage({ data: stat.icon, x: xPos + 0.27, y: yPos + 0.32, w: 0.3, h: 0.3 });
    // Value
    slide3.addText(stat.value, { x: xPos + 0.85, y: yPos + 0.1, w: 1.5, h: 0.45, fontSize: 24, fontFace: "Calibri", color: C.primary, bold: true, margin: 0 });
    // Label
    slide3.addText(stat.label, { x: xPos + 0.85, y: yPos + 0.5, w: 1.7, h: 0.5, fontSize: 10, fontFace: "Calibri", color: C.gray, margin: 0 });
  });

  // Right side - image
  slide3.addImage({ data: imgConvoyeur, x: 6.3, y: 1.1, w: 3.3, h: 4.0, sizing: { type: "cover", w: 3.3, h: 4.0 } });

  // Bottom bar
  slide3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.light } });
  slide3.addText("conwise-express.com  |  conwise.app", {
    x: 0.5, y: 5.25, w: 9, h: 0.375,
    fontSize: 10, fontFace: "Calibri", color: C.secondary, align: "center", valign: "middle", margin: 0
  });

  // =============================================
  // SLIDE 4: VOS ENJEUX - Comprendre le Groupe DMD
  // =============================================
  let slide4 = pres.addSlide();
  slide4.background = { color: C.white };

  // Top banner
  slide4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  slide4.addText("02  |  COMPRENDRE VOS ENJEUX", {
    x: 0.5, y: 0, w: 9, h: 0.9,
    fontSize: 20, fontFace: "Calibri", color: C.white, bold: true, align: "left", valign: "middle", margin: 0
  });

  // DMD Profile - Left column
  slide4.addImage({ data: logoDMD, x: 0.5, y: 1.1, w: 1.5, h: 0.7, sizing: { type: "contain", w: 1.5, h: 0.7 } });

  slide4.addText("Groupe DMD \u2014 Acteur majeur de la distribution automobile", {
    x: 2.2, y: 1.15, w: 7, h: 0.5,
    fontSize: 16, fontFace: "Calibri", color: C.primary, bold: true, margin: 0
  });

  // DMD Key facts cards - 3 columns
  const dmdFacts = [
    { value: "53", label: "Concessions", sub: "7 d\u00e9partements", color: C.primary },
    { value: "870", label: "Collaborateurs", sub: "Grand Ouest", color: C.secondary },
    { value: "700M\u20ac", label: "Chiffre d'affaires", sub: "2024", color: C.accent },
  ];

  dmdFacts.forEach((fact, i) => {
    const xPos = 0.5 + i * 3.1;
    slide4.addShape(pres.shapes.RECTANGLE, {
      x: xPos, y: 1.9, w: 2.9, h: 1.1,
      fill: { color: fact.color },
      shadow: makeShadow()
    });
    slide4.addText(fact.value, { x: xPos, y: 1.9, w: 2.9, h: 0.6, fontSize: 28, fontFace: "Calibri", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    slide4.addText(fact.label, { x: xPos, y: 2.4, w: 2.9, h: 0.3, fontSize: 13, fontFace: "Calibri", color: C.white, bold: true, align: "center", margin: 0 });
    slide4.addText(fact.sub, { x: xPos, y: 2.65, w: 2.9, h: 0.25, fontSize: 10, fontFace: "Calibri", color: C.white, align: "center", margin: 0 });
  });

  // Challenges section
  slide4.addText("Vos d\u00e9fis logistiques au quotidien", {
    x: 0.5, y: 3.2, w: 9, h: 0.4,
    fontSize: 16, fontFace: "Calibri", color: C.primary, bold: true, margin: 0
  });

  const challenges = [
    { text: "Transf\u00e9rer des v\u00e9hicules entre 53 sites r\u00e9partis sur 7 d\u00e9partements (Finist\u00e8re \u00e0 Vend\u00e9e)", icon: iconRouteBlue },
    { text: "G\u00e9rer un parc multi-marques : Ford, VW, Audi, Skoda, Opel, Jaguar, Land Rover et 10 autres", icon: iconWarehouse },
    { text: "Optimiser les co\u00fbts et d\u00e9lais de livraison VN/VO entre concessions", icon: iconClockBlue },
    { text: "Assurer la tra\u00e7abilit\u00e9 et les \u00e9tats des lieux de chaque transfert", icon: iconClipboardBlue },
  ];

  challenges.forEach((ch, i) => {
    const yPos = 3.7 + i * 0.45;
    slide4.addImage({ data: ch.icon, x: 0.6, y: yPos + 0.02, w: 0.28, h: 0.28 });
    slide4.addText(ch.text, { x: 1.1, y: yPos, w: 8.4, h: 0.38, fontSize: 12, fontFace: "Calibri", color: C.darkText, valign: "middle", margin: 0 });
  });

  // Bottom bar
  slide4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.light } });
  slide4.addImage({ data: logoConwise, x: 8.0, y: 5.25, w: 1.5, h: 0.375, sizing: { type: "contain", w: 1.5, h: 0.375 } });

  // =============================================
  // SLIDE 5: NOTRE SOLUTION
  // =============================================
  let slide5 = pres.addSlide();
  slide5.background = { color: C.white };

  // Top banner
  slide5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  slide5.addText("03  |  NOTRE SOLUTION POUR LE GROUPE DMD", {
    x: 0.5, y: 0, w: 9, h: 0.9,
    fontSize: 20, fontFace: "Calibri", color: C.white, bold: true, align: "left", valign: "middle", margin: 0
  });

  // Services grid - 2x3
  const services = [
    { title: "Convoyage inter-sites", desc: "Transfert de v\u00e9hicules entre vos 53 concessions dans tout le Grand Ouest", icon: iconTruck, color: C.primary },
    { title: "Livraison client", desc: "Mise en main et livraison de v\u00e9hicules neufs et d'occasion \u00e0 vos clients", icon: iconCar, color: C.secondary },
    { title: "\u00c9tats des lieux digitaux", desc: "PV photo d\u00e9taill\u00e9 au d\u00e9part et \u00e0 l'arriv\u00e9e, consultable en temps r\u00e9el", icon: iconClipboard, color: C.accent },
    { title: "Suivi GPS temps r\u00e9el", desc: "Tra\u00e7abilit\u00e9 compl\u00e8te de chaque mission en cours sur votre tableau de bord", icon: iconMap, color: C.primary },
    { title: "Restitution leasing", desc: "Gestion des retours LLD/LOA pour toutes vos marques", icon: iconRoute, color: C.secondary },
    { title: "Rapatriement", desc: "R\u00e9cup\u00e9ration de v\u00e9hicules en panne, accident ou achat \u00e0 distance", icon: iconTools, color: C.accent },
  ];

  services.forEach((svc, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const xPos = 0.5 + col * 3.1;
    const yPos = 1.15 + row * 2.1;

    // Card
    slide5.addShape(pres.shapes.RECTANGLE, {
      x: xPos, y: yPos, w: 2.85, h: 1.9,
      fill: { color: C.offWhite },
      shadow: makeShadowLight()
    });
    // Icon circle
    slide5.addShape(pres.shapes.OVAL, { x: xPos + 0.95, y: yPos + 0.2, w: 0.7, h: 0.7, fill: { color: svc.color } });
    slide5.addImage({ data: svc.icon, x: xPos + 1.1, y: yPos + 0.35, w: 0.4, h: 0.4 });
    // Title
    slide5.addText(svc.title, { x: xPos + 0.15, y: yPos + 1.0, w: 2.55, h: 0.35, fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, align: "center", margin: 0 });
    // Description
    slide5.addText(svc.desc, { x: xPos + 0.15, y: yPos + 1.3, w: 2.55, h: 0.5, fontSize: 10, fontFace: "Calibri", color: C.gray, align: "center", margin: 0 });
  });

  // Bottom bar
  slide5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.light } });
  slide5.addImage({ data: logoConwise, x: 8.0, y: 5.25, w: 1.5, h: 0.375, sizing: { type: "contain", w: 1.5, h: 0.375 } });

  // =============================================
  // SLIDE 6: COMMENT CA MARCHE - Process
  // =============================================
  let slide6 = pres.addSlide();
  slide6.background = { color: C.offWhite };

  // Top banner
  slide6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  slide6.addText("03  |  COMMENT \u00c7A MARCHE", {
    x: 0.5, y: 0, w: 9, h: 0.9,
    fontSize: 20, fontFace: "Calibri", color: C.white, bold: true, align: "left", valign: "middle", margin: 0
  });

  slide6.addText("Un processus simple en 4 \u00e9tapes", {
    x: 0.5, y: 1.05, w: 9, h: 0.5,
    fontSize: 18, fontFace: "Calibri", color: C.primary, bold: true, margin: 0
  });

  const steps = [
    { num: "1", title: "Commandez", desc: "Passez votre commande depuis votre portail conwise.app en quelques clics", icon: iconMobile },
    { num: "2", title: "Prise en charge", desc: "Un convoyeur certifi\u00e9 r\u00e9alise l'\u00e9tat des lieux photo au d\u00e9part", icon: iconClipboard },
    { num: "3", title: "Convoyage", desc: "Suivi GPS en temps r\u00e9el de votre v\u00e9hicule jusqu'\u00e0 destination", icon: iconTruck },
    { num: "4", title: "Livraison", desc: "\u00c9tat des lieux de livraison, PV digital et confirmation instantan\u00e9e", icon: iconCheckW },
  ];

  steps.forEach((step, i) => {
    const xPos = 0.4 + i * 2.4;

    // Card
    slide6.addShape(pres.shapes.RECTANGLE, {
      x: xPos, y: 1.7, w: 2.15, h: 2.8,
      fill: { color: C.white },
      shadow: makeShadow()
    });

    // Step number circle
    slide6.addShape(pres.shapes.OVAL, { x: xPos + 0.7, y: 1.9, w: 0.65, h: 0.65, fill: { color: C.secondary } });
    slide6.addText(step.num, { x: xPos + 0.7, y: 1.9, w: 0.65, h: 0.65, fontSize: 22, fontFace: "Calibri", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

    // Icon
    slide6.addShape(pres.shapes.OVAL, { x: xPos + 0.6, y: 2.75, w: 0.85, h: 0.85, fill: { color: C.primary } });
    slide6.addImage({ data: step.icon, x: xPos + 0.78, y: 2.93, w: 0.5, h: 0.5 });

    // Title
    slide6.addText(step.title, { x: xPos + 0.1, y: 3.75, w: 1.95, h: 0.35, fontSize: 14, fontFace: "Calibri", color: C.primary, bold: true, align: "center", margin: 0 });

    // Description
    slide6.addText(step.desc, { x: xPos + 0.1, y: 4.05, w: 1.95, h: 0.55, fontSize: 10, fontFace: "Calibri", color: C.gray, align: "center", margin: 0 });

    // Connector arrow (except last)
    if (i < 3) {
      slide6.addText("\u25B6", { x: xPos + 2.15, y: 2.85, w: 0.25, h: 0.4, fontSize: 16, color: C.secondary, align: "center", valign: "middle", margin: 0 });
    }
  });

  // Bottom bar
  slide6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.light } });
  slide6.addImage({ data: logoConwise, x: 8.0, y: 5.25, w: 1.5, h: 0.375, sizing: { type: "contain", w: 1.5, h: 0.375 } });

  // =============================================
  // SLIDE 7: PLATEFORME conwise.app
  // =============================================
  let slide7 = pres.addSlide();
  slide7.background = { color: C.white };

  // Top banner
  slide7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  slide7.addText("04  |  LA PLATEFORME CONWISE.APP", {
    x: 0.5, y: 0, w: 9, h: 0.9,
    fontSize: 20, fontFace: "Calibri", color: C.white, bold: true, align: "left", valign: "middle", margin: 0
  });

  // Left side - features
  slide7.addText("Votre tableau de bord Fleet Manager", {
    x: 0.5, y: 1.1, w: 4.5, h: 0.5,
    fontSize: 17, fontFace: "Calibri", color: C.primary, bold: true, margin: 0
  });

  slide7.addText("Un environnement de gestion complet, personnalis\u00e9 aux couleurs du Groupe DMD", {
    x: 0.5, y: 1.55, w: 4.5, h: 0.4,
    fontSize: 11, fontFace: "Calibri", color: C.gray, margin: 0
  });

  const features = [
    { text: "Commandez vos convoyages en quelques clics", icon: iconMobileBlue },
    { text: "Suivez le statut de chaque mission en temps r\u00e9el", icon: iconRouteBlue },
    { text: "\u00c9tats des lieux photo d\u00e9taill\u00e9s et digitalis\u00e9s", icon: iconClipboardBlue },
    { text: "G\u00e9rez vos adresses favorites (53 concessions)", icon: iconMapBlue },
    { text: "Statistiques et tableau de bord analytique", icon: iconChartBlue },
    { text: "Factures et historique centralis\u00e9s", icon: iconFileBlue },
  ];

  features.forEach((feat, i) => {
    const yPos = 2.1 + i * 0.48;
    slide7.addShape(pres.shapes.OVAL, { x: 0.5, y: yPos, w: 0.32, h: 0.32, fill: { color: C.light } });
    slide7.addImage({ data: feat.icon, x: 0.56, y: yPos + 0.06, w: 0.2, h: 0.2 });
    slide7.addText(feat.text, { x: 1.0, y: yPos, w: 4, h: 0.32, fontSize: 12, fontFace: "Calibri", color: C.darkText, valign: "middle", margin: 0 });
  });

  // Right side - Dashboard screenshot
  slide7.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.1, w: 4.5, h: 3.6,
    fill: { color: C.offWhite },
    shadow: makeShadow()
  });
  slide7.addImage({ data: imgDashboard, x: 5.3, y: 1.2, w: 4.3, h: 3.4, sizing: { type: "contain", w: 4.3, h: 3.4 } });

  // URL callout
  slide7.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 4.85, w: 4.5, h: 0.35,
    fill: { color: C.primary }
  });
  slide7.addText("app.conwise.app  \u2014  Acc\u00e8s portail d\u00e9di\u00e9 Groupe DMD", {
    x: 5.2, y: 4.85, w: 4.5, h: 0.35,
    fontSize: 10, fontFace: "Calibri", color: C.white, align: "center", valign: "middle", margin: 0
  });

  // Bottom bar
  slide7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.light } });
  slide7.addImage({ data: logoConwise, x: 8.0, y: 5.25, w: 1.5, h: 0.375, sizing: { type: "contain", w: 1.5, h: 0.375 } });

  // =============================================
  // SLIDE 8: NOS ENGAGEMENTS QUALITE
  // =============================================
  let slide8 = pres.addSlide();
  slide8.background = { color: C.white };

  // Top banner
  slide8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  slide8.addText("05  |  NOS ENGAGEMENTS QUALIT\u00c9", {
    x: 0.5, y: 0, w: 9, h: 0.9,
    fontSize: 20, fontFace: "Calibri", color: C.white, bold: true, align: "left", valign: "middle", margin: 0
  });

  // Left image
  slide8.addImage({ data: imgEtatLieux, x: 0.3, y: 1.1, w: 3.2, h: 4.0, sizing: { type: "cover", w: 3.2, h: 4.0 } });

  // Right content - engagement cards
  const engagements = [
    { title: "Assurance tous risques", desc: "Chaque v\u00e9hicule transport\u00e9 est couvert par une assurance compl\u00e8te pendant toute la dur\u00e9e du convoyage", icon: iconShield, color: C.primary },
    { title: "Convoyeurs certifi\u00e9s", desc: "Tous nos convoyeurs sont form\u00e9s, v\u00e9rifi\u00e9s et \u00e9valu\u00e9s r\u00e9guli\u00e8rement selon nos standards", icon: iconUsers, color: C.secondary },
    { title: "D\u00e9lai garanti 24-48h", desc: "Livraison sous 24 \u00e0 48h sur l'ensemble du territoire, option express disponible", icon: iconClock, color: C.accent },
    { title: "Groupement d'employeurs", desc: "Mod\u00e8le innovant garantissant tarifs comp\u00e9titifs et qualit\u00e9 standardis\u00e9e sur tout le r\u00e9seau", icon: iconCogs, color: C.primary },
    { title: "Tarification transparente", desc: "Prix fixes sans frais cach\u00e9s, devis en 15 minutes, facturation centralis\u00e9e", icon: iconChart, color: C.secondary },
  ];

  engagements.forEach((eng, i) => {
    const yPos = 1.1 + i * 0.82;
    // Icon circle
    slide8.addShape(pres.shapes.OVAL, { x: 3.8, y: yPos + 0.05, w: 0.5, h: 0.5, fill: { color: eng.color } });
    slide8.addImage({ data: eng.icon, x: 3.92, y: yPos + 0.17, w: 0.26, h: 0.26 });
    // Title
    slide8.addText(eng.title, { x: 4.5, y: yPos, w: 5, h: 0.3, fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, margin: 0 });
    // Description
    slide8.addText(eng.desc, { x: 4.5, y: yPos + 0.3, w: 5, h: 0.4, fontSize: 10, fontFace: "Calibri", color: C.gray, margin: 0 });
  });

  // Bottom bar
  slide8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.light } });
  slide8.addImage({ data: logoConwise, x: 8.0, y: 5.25, w: 1.5, h: 0.375, sizing: { type: "contain", w: 1.5, h: 0.375 } });

  // =============================================
  // SLIDE 9: REFERENCES
  // =============================================
  let slide9 = pres.addSlide();
  slide9.background = { color: C.offWhite };

  // Top banner
  slide9.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  slide9.addText("06  |  ILS NOUS FONT CONFIANCE", {
    x: 0.5, y: 0, w: 9, h: 0.9,
    fontSize: 20, fontFace: "Calibri", color: C.white, bold: true, align: "left", valign: "middle", margin: 0
  });

  slide9.addText("Des acteurs majeurs de l'automobile et de la logistique", {
    x: 0.5, y: 1.1, w: 9, h: 0.4,
    fontSize: 16, fontFace: "Calibri", color: C.primary, bold: true, margin: 0
  });

  // Reference logos/names - cards
  const refs = [
    { name: "Elior", sector: "Restauration collective" },
    { name: "Mondial Relay", sector: "Logistique & livraison" },
    { name: "Auto Marchands", sector: "N\u00e9goce automobile" },
    { name: "Peugeot Spoticar", sector: "Occasions constructeur" },
    { name: "Skoda", sector: "Constructeur automobile" },
    { name: "Segafredo", sector: "Distribution" },
  ];

  refs.forEach((ref, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const xPos = 0.5 + col * 3.1;
    const yPos = 1.7 + row * 1.2;

    slide9.addShape(pres.shapes.RECTANGLE, {
      x: xPos, y: yPos, w: 2.85, h: 1.0,
      fill: { color: C.white },
      shadow: makeShadowLight()
    });
    slide9.addText(ref.name, { x: xPos, y: yPos + 0.1, w: 2.85, h: 0.45, fontSize: 16, fontFace: "Calibri", color: C.primary, bold: true, align: "center", margin: 0 });
    slide9.addText(ref.sector, { x: xPos, y: yPos + 0.55, w: 2.85, h: 0.3, fontSize: 10, fontFace: "Calibri", color: C.gray, align: "center", margin: 0 });
  });

  // Testimonial quote area
  slide9.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.2, w: 9, h: 0.85,
    fill: { color: C.primary }
  });
  slide9.addImage({ data: iconStar, x: 0.8, y: 4.35, w: 0.25, h: 0.25 });
  slide9.addImage({ data: iconStar, x: 1.1, y: 4.35, w: 0.25, h: 0.25 });
  slide9.addImage({ data: iconStar, x: 1.4, y: 4.35, w: 0.25, h: 0.25 });
  slide9.addImage({ data: iconStar, x: 1.7, y: 4.35, w: 0.25, h: 0.25 });
  slide9.addImage({ data: iconStar, x: 2.0, y: 4.35, w: 0.25, h: 0.25 });
  slide9.addText("\"Un service fiable, des convoyeurs professionnels et une plateforme digitale qui nous fait gagner un temps pr\u00e9cieux au quotidien.\"", {
    x: 0.8, y: 4.6, w: 8.4, h: 0.35,
    fontSize: 11, fontFace: "Calibri", color: C.white, italic: true, align: "left", margin: 0
  });

  // Bottom bar
  slide9.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.light } });
  slide9.addImage({ data: logoConwise, x: 8.0, y: 5.25, w: 1.5, h: 0.375, sizing: { type: "contain", w: 1.5, h: 0.375 } });

  // =============================================
  // SLIDE 10: OFFRE & PROCHAINES ETAPES
  // =============================================
  let slide10 = pres.addSlide();
  slide10.background = { color: C.white };

  // Top banner
  slide10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  slide10.addText("07  |  OFFRE & PROCHAINES \u00c9TAPES", {
    x: 0.5, y: 0, w: 9, h: 0.9,
    fontSize: 20, fontFace: "Calibri", color: C.white, bold: true, align: "left", valign: "middle", margin: 0
  });

  // Left - Offer card
  slide10.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.1, w: 4.3, h: 3.9,
    fill: { color: C.offWhite },
    shadow: makeShadow()
  });
  // Offer header
  slide10.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.3, h: 0.7, fill: { color: C.primary } });
  slide10.addText("Offre Fleet Manager", {
    x: 0.5, y: 1.1, w: 4.3, h: 0.7,
    fontSize: 18, fontFace: "Calibri", color: C.white, bold: true, align: "center", valign: "middle", margin: 0
  });
  slide10.addText("Pack sur-mesure Groupe DMD", {
    x: 0.7, y: 1.9, w: 3.9, h: 0.35,
    fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, align: "center", margin: 0
  });

  const offerItems = [
    "Acc\u00e8s portail conwise.app d\u00e9di\u00e9",
    "Int\u00e9gration de vos 53 sites",
    "Tableau de bord & statistiques",
    "\u00c9tats des lieux photo digitalis\u00e9s",
    "Suivi GPS temps r\u00e9el",
    "Facturation centralis\u00e9e mensuelle",
    "Interlocuteur d\u00e9di\u00e9",
  ];

  offerItems.forEach((item, i) => {
    const yPos = 2.35 + i * 0.35;
    slide10.addImage({ data: iconCheck, x: 0.8, y: yPos + 0.03, w: 0.2, h: 0.2 });
    slide10.addText(item, { x: 1.15, y: yPos, w: 3.4, h: 0.28, fontSize: 11, fontFace: "Calibri", color: C.darkText, valign: "middle", margin: 0 });
  });

  // Right - Next steps
  slide10.addText("Prochaines \u00e9tapes", {
    x: 5.3, y: 1.1, w: 4.3, h: 0.5,
    fontSize: 18, fontFace: "Calibri", color: C.primary, bold: true, margin: 0
  });

  const nextSteps = [
    { num: "1", title: "Appel de qualification", desc: "Comprendre vos volumes et besoins sp\u00e9cifiques multi-sites" },
    { num: "2", title: "Proposition tarifaire", desc: "Grille de tarifs d\u00e9di\u00e9e adapt\u00e9e \u00e0 vos flux r\u00e9currents" },
    { num: "3", title: "Mise en place", desc: "Configuration du portail avec vos 53 concessions" },
    { num: "4", title: "Phase pilote", desc: "Test sur un p\u00e9rim\u00e8tre r\u00e9duit avant d\u00e9ploiement complet" },
  ];

  nextSteps.forEach((step, i) => {
    const yPos = 1.75 + i * 0.85;
    // Number
    slide10.addShape(pres.shapes.OVAL, { x: 5.3, y: yPos, w: 0.4, h: 0.4, fill: { color: C.secondary } });
    slide10.addText(step.num, { x: 5.3, y: yPos, w: 0.4, h: 0.4, fontSize: 14, fontFace: "Calibri", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    // Content
    slide10.addText(step.title, { x: 5.9, y: yPos, w: 3.6, h: 0.3, fontSize: 13, fontFace: "Calibri", color: C.primary, bold: true, margin: 0 });
    slide10.addText(step.desc, { x: 5.9, y: yPos + 0.3, w: 3.6, h: 0.35, fontSize: 10, fontFace: "Calibri", color: C.gray, margin: 0 });
  });

  // Bottom bar
  slide10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.light } });
  slide10.addImage({ data: logoConwise, x: 8.0, y: 5.25, w: 1.5, h: 0.375, sizing: { type: "contain", w: 1.5, h: 0.375 } });

  // =============================================
  // SLIDE 11: CONTACT - Closing slide
  // =============================================
  let slide11 = pres.addSlide();
  slide11.background = { color: C.dark };

  // Background image with dark overlay
  slide11.addImage({ data: imgBanner, x: 0, y: 0, w: 10, h: 5.625, sizing: { type: "cover", w: 10, h: 5.625 } });
  slide11.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.dark, transparency: 25 } });

  // Logo
  slide11.addImage({ data: logoConwise, x: 3.0, y: 0.6, w: 4.0, h: 1.2, sizing: { type: "contain", w: 4.0, h: 1.2 } });

  // Title
  slide11.addText("Simplifions ensemble la logistique\ndu Groupe DMD", {
    x: 1, y: 2.0, w: 8, h: 1.0,
    fontSize: 26, fontFace: "Calibri", color: C.white, bold: true, align: "center", lineSpacingMultiple: 1.3, margin: 0
  });

  // Contact info cards
  const contacts = [
    { icon: iconPhone, text: "+33 7 44 31 79 16" },
    { icon: iconEnvelope, text: "contact@conwise-express.com" },
    { icon: iconGlobe, text: "conwise-express.com | conwise.app" },
  ];

  contacts.forEach((c, i) => {
    const xPos = 1.5 + i * 2.5;
    slide11.addShape(pres.shapes.RECTANGLE, {
      x: xPos, y: 3.3, w: 2.3, h: 0.8,
      fill: { color: C.primary, transparency: 40 }
    });
    slide11.addImage({ data: c.icon, x: xPos + 0.15, y: 3.48, w: 0.3, h: 0.3 });
    slide11.addText(c.text, { x: xPos + 0.5, y: 3.3, w: 1.7, h: 0.8, fontSize: 10, fontFace: "Calibri", color: C.white, valign: "middle", margin: 0 });
  });

  // Address
  slide11.addText("Conwise Express  \u2014  12 Cours des Merveilles, 95000 Cergy", {
    x: 1, y: 4.4, w: 8, h: 0.4,
    fontSize: 11, fontFace: "Calibri", color: C.accent, align: "center", margin: 0
  });

  // Bottom tagline
  slide11.addText("Le convoyage automobile, simplifi\u00e9.", {
    x: 1, y: 4.9, w: 8, h: 0.4,
    fontSize: 14, fontFace: "Calibri", color: C.white, italic: true, align: "center", margin: 0
  });

  // ============ SAVE ============
  const outputPath = path.join(baseDir, "Proposition Conwise Express - Groupe DMD.pptx");
  await pres.writeFile({ fileName: outputPath });
  console.log("Presentation saved to: " + outputPath);
}

createPresentation().catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
