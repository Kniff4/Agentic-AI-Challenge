import pptxgen from "pptxgenjs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ── Brand tokens from specs.json ──────────────────────────────────────────────
const BRAND = {
  purple:    "24135F",
  pink:      "D0006F",
  grey:      "B1B3B3",
  violet:    "422990",
  teal:      "2DCCD3",
  black:     "1d1d1d",
  lightGrey: "b2b2b2",
  white:     "FFFFFF",
};

const FONT = {
  primary:  "Arial",   // swap to "Campton" if installed
  fallback: "Arial",
};

const LOGO_WHITE   = path.resolve(__dirname, "../../assets/talentia/logo_primary.png");
const LOGO_PRIMARY = path.resolve(__dirname, "../../assets/talentia/logo_white.png");
const ICON_WHITE = path.resolve(__dirname, "../../services/pptx-service/talentia/logo_icon.png");


// ── Helpers ───────────────────────────────────────────────────────────────────

function mmToInches(mm) {
  return mm / 25.4;
}

/** Small logo watermark in the bottom-right corner */
function addLogoCorner(slide, logoPath) {
  slide.addImage({ path: logoPath, x: "85%", y: "85%", w: 1, h:1 });
}

/** Full-bleed rectangle (used as accent bar or overlay) */
function addRect(slide, color, x, y, w, h) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color },
    line: { color, width: 0 },
  });
}

// ── Slide builders ────────────────────────────────────────────────────────────

/** Slide 1 – Title slide (dark purple background) */
function addTitleSlide(pptx, title, subtitle) {
  const slide = pptx.addSlide();

  // Background
  slide.background = { color: BRAND.purple };

  // Pink accent bar on the left
  addRect(slide, BRAND.pink, 0, 0, 0.18, 7.5);

  // Logo (white version on dark bg)
  slide.addImage({ path: LOGO_WHITE, x: 0.4, y: 0.3, w: 2.1, h: 1 });

  // Title
  slide.addText(title, {
    x: 0.5, y: 2.6, w: 9, h: 1.4,
    fontSize: 40,
    bold: true,
    color: BRAND.white,
    fontFace: FONT.primary,
    align: "left",
  });

  // Subtitle
  slide.addText(subtitle, {
    x: 0.5, y: 4.1, w: 8, h: 0.8,
    fontSize: 18,
    color: BRAND.grey,
    fontFace: FONT.primary,
    align: "left",
  });

  // Date / footer
  slide.addText("April 2026 · Talentia Software", {
    x: 0.5, y: 6.9, w: 9, h: 0.4,
    fontSize: 10,
    color: BRAND.lightGrey,
    fontFace: FONT.primary,
    align: "left",
  });
}

/** Slide 2 – Agenda / bullet list */
function addAgendaSlide(pptx, items) {
  const slide = pptx.addSlide();

  // Top accent bar
  addRect(slide, BRAND.purple, 0, 0, 10, 0.12);

  // Section label
  slide.addText("AGENDA", {
    x: 0.5, y: 0.25, w: 3, h: 0.4,
    fontSize: 9,
    bold: true,
    color: BRAND.pink,
    fontFace: FONT.primary,
    charSpacing: 3,
  });

  // Title
  slide.addText("What we'll cover today", {
    x: 0.5, y: 0.7, w: 8.5, h: 0.8,
    fontSize: 28,
    bold: true,
    color: BRAND.purple,
    fontFace: FONT.primary,
  });

  // Divider line
  addRect(slide, BRAND.grey, 0.5, 1.55, 9, 0.025);

  // Bullet items
  items.forEach((item, i) => {
    const y = 1.8 + i * 0.75;
    // Pink number circle
    addRect(slide, BRAND.pink, 0.5, y + 0.05, 0.38, 0.38);
    slide.addText(String(i + 1), {
      x: 0.5, y: y + 0.05, w: 0.38, h: 0.38,
      fontSize: 12,
      bold: true,
      color: BRAND.white,
      fontFace: FONT.primary,
      align: "center",
      valign: "middle",
    });
    slide.addText(item, {
      x: 1.05, y, w: 8, h: 0.55,
      fontSize: 16,
      color: BRAND.black,
      fontFace: FONT.primary,
      valign: "middle",
    });
  });

  addLogoCorner(slide, ICON_WHITE);
}

/** Slide 3 – Section divider */
function addSectionSlide(pptx, sectionNumber, sectionTitle) {
  const slide = pptx.addSlide();

  // Left half: purple
  addRect(slide, BRAND.purple, 0, 0, 5, 7.5);
  // Right half: teal accent strip
  addRect(slide, BRAND.teal, 5, 0, 0.08, 7.5);
  // Right half: white
  addRect(slide, BRAND.white, 5.08, 0, 4.92, 7.5);

  // Section number (large, faint)
  slide.addText(`0${sectionNumber}`, {
    x: 0.3, y: 2.2, w: 4.5, h: 2.5,
    fontSize: 120,
    bold: true,
    color: BRAND.violet,
    fontFace: FONT.primary,
    transparency: 60,
  });

  // Section title on the white side
  slide.addText(sectionTitle, {
    x: 5.3, y: 2.8, w: 4.3, h: 1.5,
    fontSize: 26,
    bold: true,
    color: BRAND.purple,
    fontFace: FONT.primary,
  });

  // White logo on purple side
  slide.addImage({ path: LOGO_WHITE, x: 0.35, y: 0.28, w: 2.1, h: 1 });
}

/** Slide 4 – Two-column content slide */
function addContentSlide(pptx, title, leftText, rightText) {
  const slide = pptx.addSlide();

  addRect(slide, BRAND.purple, 0, 0, 10, 0.12);

  slide.addText(title, {
    x: 0.5, y: 0.25, w: 9, h: 0.7,
    fontSize: 24,
    bold: true,
    color: BRAND.purple,
    fontFace: FONT.primary,
  });

  addRect(slide, BRAND.pink, 0.5, 1.05, 4.2, 0.04);
  addRect(slide, BRAND.pink, 5.3, 1.05, 4.2, 0.04);

  slide.addText(leftText, {
    x: 0.5, y: 1.2, w: 4.2, h: 5.5,
    fontSize: 13,
    color: BRAND.black,
    fontFace: FONT.primary,
    valign: "top",
    wrap: true,
  });

  slide.addText(rightText, {
    x: 5.3, y: 1.2, w: 4.2, h: 5.5,
    fontSize: 13,
    color: BRAND.black,
    fontFace: FONT.primary,
    valign: "top",
    wrap: true,
  });

  addLogoCorner(slide, ICON_WHITE);
}

/** Slide 5 – Closing slide */
function addClosingSlide(pptx, tagline, contact) {
  const slide = pptx.addSlide();

  slide.background = { color: BRAND.purple };
  // Teal bottom stripe
  addRect(slide, BRAND.teal, 0, 7.1, 10, 0.4);
  // Pink left bar
  addRect(slide, BRAND.pink, 0, 0, 0.18, 7.5);

  slide.addImage({ path: LOGO_WHITE, x: 0.4, y: 0.3, w: 2.1, h: 1 });

  slide.addText(tagline, {
    x: 0.6, y: 2.5, w: 8.8, h: 1.6,
    fontSize: 38,
    bold: true,
    color: BRAND.white,
    fontFace: FONT.primary,
    align: "center",
  });

  slide.addText(contact, {
    x: 0.6, y: 4.4, w: 8.8, h: 0.6,
    fontSize: 14,
    color: BRAND.grey,
    fontFace: FONT.primary,
    align: "center",
  });
}

// ── Build & save ──────────────────────────────────────────────────────────────

async function generate() {
  const pptx = new pptxgen();

  pptx.layout = "LAYOUT_4x3";

  addTitleSlide(
    pptx,
    "Talentia Software\nProduct Overview",
    "Empowering HR & Finance teams worldwide"
  );

  addAgendaSlide(pptx, [
    "Company introduction",
    "Core product suite",
    "Customer success stories",
    "Roadmap & next steps",
  ]);

  addSectionSlide(pptx, 1, "Company\nIntroduction");

  addContentSlide(
    pptx,
    "Who We Are",
    "Talentia Software is a European leader in HR and Finance software solutions, helping organisations manage their most critical processes with efficiency and confidence.\n\n• Founded 2009\n• 500+ employees across 10 countries\n• 3 000+ customers worldwide",
    "Our platform covers the full employee lifecycle — from recruitment and onboarding to payroll and performance management — integrated with powerful financial consolidation and reporting tools.\n\n• Cloud-native SaaS delivery\n• Open API ecosystem\n• Multilingual & multi-currency"
  );

  addSectionSlide(pptx, 2, "Core Product\nSuite");

  addContentSlide(
    pptx,
    "Our Solutions",
    "HCM Suite\n—\n• Talent Acquisition\n• Onboarding\n• Learning & Development\n• Performance & Goals\n• Compensation Planning\n• HR Analytics",
    "Finance Suite\n—\n• Financial Consolidation\n• Budgeting & Forecasting\n• Management Reporting\n• IFRS / Local GAAP Compliance\n• Intercompany Eliminations"
  );

  addClosingSlide(
    pptx,
    "Let's build the future\nof work together.",
    "contact@talentia.com  ·  www.talentia-software.com"
  );

  const outPath = path.resolve(__dirname, "example_presentation.pptx");
  await pptx.writeFile({ fileName: outPath });
  console.log(`Presentation saved → ${outPath}`);
}

generate().catch(console.error);
