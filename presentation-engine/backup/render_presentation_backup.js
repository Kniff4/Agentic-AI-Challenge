const fs = require("fs");
const path = require("path");
const PptxGenJS = require("pptxgenjs");

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function loadJson(filePath) {
  const raw = fs.readFileSync(filePath, "utf8");
  return JSON.parse(raw);
}

function addFooter(slide, ppt, globalStyle) {
  if (globalStyle?.footer_text) {
    slide.addText(globalStyle.footer_text, {
      x: 0.3,
      y: 7.0,
      w: 5.0,
      h: 0.2,
      fontSize: 9,
      color: "666666",
      margin: 0
    });
  }

  if (globalStyle?.show_slide_numbers) {
    slide.addText(`${ppt._slides.length}`, {
      x: 12.7,
      y: 7.0,
      w: 0.3,
      h: 0.2,
      fontSize: 9,
      align: "right",
      color: "666666",
      margin: 0
    });
  }
}

function addLogo(slide, assetsMap, globalStyle) {
  if (!globalStyle?.use_logo) return;
  const logo = assetsMap["company_logo"];
  if (!logo || !logo.file_name) return;
  if (!fs.existsSync(logo.file_name)) return;

  slide.addImage({
    path: logo.file_name,
    x: 11.6,
    y: 0.2,
    w: 1.4,
    h: 0.5
  });
}

function baseSlide(ppt, deck, globalStyle, assetsMap) {
  const slide = ppt.addSlide();
  slide.background = { color: globalStyle?.theme_variant === "dark" ? "1F1F1F" : "FFFFFF" };
  addLogo(slide, assetsMap, globalStyle);
  return slide;
}

function titleStyle(globalStyle) {
  return {
    fontFace: "Aptos",
    bold: true,
    color: globalStyle?.theme_variant === "dark" ? "FFFFFF" : "1F1F1F"
  };
}

function bodyStyle(globalStyle) {
  return {
    fontFace: "Aptos",
    color: globalStyle?.theme_variant === "dark" ? "F2F2F2" : "333333"
  };
}

function renderTitleSlide(slide, s, globalStyle) {
  const t = titleStyle(globalStyle);
  const b = bodyStyle(globalStyle);

  slide.addText(s.title || "", {
    x: 0.8,
    y: 1.3,
    w: 10.8,
    h: 0.8,
    fontSize: 26,
    ...t
  });

  slide.addText(s.content?.subtitle || "", {
    x: 0.8,
    y: 2.2,
    w: 9.5,
    h: 0.5,
    fontSize: 16,
    ...b
  });

  if (s.content?.tagline) {
    slide.addText(s.content.tagline, {
      x: 0.8,
      y: 2.9,
      w: 9.0,
      h: 0.4,
      fontSize: 12,
      italic: true,
      color: "666666"
    });
  }
}

function renderSectionDivider(slide, s, globalStyle) {
  const t = titleStyle(globalStyle);

  slide.addShape("rect", {
    x: 0,
    y: 1.5,
    w: 13.33,
    h: 2.2,
    fill: { color: "EAF2FF" },
    line: { color: "EAF2FF" }
  });

  slide.addText(s.title || "", {
    x: 0.8,
    y: 2.05,
    w: 10.0,
    h: 0.8,
    fontSize: 24,
    ...t
  });

  if (s.content?.subtitle) {
    slide.addText(s.content.subtitle, {
      x: 0.8,
      y: 2.75,
      w: 9.0,
      h: 0.4,
      fontSize: 12,
      color: "4D4D4D"
    });
  }
}

function renderAgenda(slide, s, globalStyle) {
  const t = titleStyle(globalStyle);
  const b = bodyStyle(globalStyle);

  slide.addText(s.title || "Agenda", {
    x: 0.8,
    y: 0.7,
    w: 5.0,
    h: 0.5,
    fontSize: 22,
    ...t
  });

  const items = s.content?.items || [];
  const runs = [];
  items.forEach((item, i) => {
    runs.push({ text: `${i + 1}. ${item}`, options: { breakLine: true } });
  });

  slide.addText(runs, {
    x: 1.1,
    y: 1.6,
    w: 8.5,
    h: 4.5,
    fontSize: 18,
    bullet: false,
    ...b
  });
}

function renderBulletSummary(slide, s, globalStyle) {
  const t = titleStyle(globalStyle);
  const b = bodyStyle(globalStyle);

  slide.addText(s.title || "", {
    x: 0.8,
    y: 0.7,
    w: 9.0,
    h: 0.5,
    fontSize: 22,
    ...t
  });

  if (s.content?.intro) {
    slide.addText(s.content.intro, {
      x: 0.9,
      y: 1.4,
      w: 11.0,
      h: 0.5,
      fontSize: 13,
      ...b
    });
  }

  const bullets = (s.content?.bullets || []).map(text => ({ text, options: { bullet: { indent: 18 } } }));
  slide.addText(bullets, {
    x: 1.0,
    y: 2.0,
    w: 10.8,
    h: 3.8,
    fontSize: 18,
    paraSpaceAfterPt: 10,
    ...b
  });
}

function renderTwoColumn(slide, s, globalStyle) {
  const t = titleStyle(globalStyle);
  const b = bodyStyle(globalStyle);

  slide.addText(s.title || "", {
    x: 0.8,
    y: 0.7,
    w: 10.0,
    h: 0.5,
    fontSize: 22,
    ...t
  });

  slide.addText(s.content?.left_title || "", {
    x: 0.9,
    y: 1.5,
    w: 5.0,
    h: 0.4,
    fontSize: 16,
    bold: true,
    ...b
  });

  slide.addText(s.content?.right_title || "", {
    x: 6.8,
    y: 1.5,
    w: 5.0,
    h: 0.4,
    fontSize: 16,
    bold: true,
    ...b
  });

  const leftItems = (s.content?.left_items || []).map(text => ({ text, options: { bullet: { indent: 18 } } }));
  const rightItems = (s.content?.right_items || []).map(text => ({ text, options: { bullet: { indent: 18 } } }));

  slide.addText(leftItems, {
    x: 0.9,
    y: 2.0,
    w: 5.2,
    h: 4.0,
    fontSize: 16,
    ...b
  });

  slide.addText(rightItems, {
    x: 6.8,
    y: 2.0,
    w: 5.2,
    h: 4.0,
    fontSize: 16,
    ...b
  });
}

function renderComparison(slide, s, globalStyle) {
  const t = titleStyle(globalStyle);

  slide.addText(s.title || "", {
    x: 0.8,
    y: 0.7,
    w: 10.0,
    h: 0.5,
    fontSize: 22,
    ...t
  });

  slide.addShape("roundRect", {
    x: 0.8,
    y: 1.6,
    w: 5.6,
    h: 4.6,
    rectRadius: 0.08,
    fill: { color: "F5F5F5" },
    line: { color: "D9D9D9" }
  });

  slide.addShape("roundRect", {
    x: 6.9,
    y: 1.6,
    w: 5.6,
    h: 4.6,
    rectRadius: 0.08,
    fill: { color: "EAF2FF" },
    line: { color: "B7D1FF" }
  });

  slide.addText(s.content?.left_title || "Left", {
    x: 1.1,
    y: 1.9,
    w: 4.5,
    h: 0.4,
    fontSize: 16,
    bold: true,
    color: "333333"
  });

  slide.addText(s.content?.right_title || "Right", {
    x: 7.2,
    y: 1.9,
    w: 4.5,
    h: 0.4,
    fontSize: 16,
    bold: true,
    color: "1F1F1F"
  });

  const leftPoints = (s.content?.left_points || []).map(text => ({ text, options: { bullet: { indent: 18 } } }));
  const rightPoints = (s.content?.right_points || []).map(text => ({ text, options: { bullet: { indent: 18 } } }));

  slide.addText(leftPoints, {
    x: 1.1,
    y: 2.4,
    w: 4.8,
    h: 3.3,
    fontSize: 15,
    color: "444444"
  });

  slide.addText(rightPoints, {
    x: 7.2,
    y: 2.4,
    w: 4.8,
    h: 3.3,
    fontSize: 15,
    color: "222222"
  });
}

function renderProcessFlow(slide, s, globalStyle) {
  const t = titleStyle(globalStyle);

  slide.addText(s.title || "", {
    x: 0.8,
    y: 0.7,
    w: 10.0,
    h: 0.5,
    fontSize: 22,
    ...t
  });

  const steps = s.content?.steps || [];
  const count = Math.max(steps.length, 1);
  const boxW = 2.1;
  const gap = 0.35;
  const totalW = count * boxW + (count - 1) * gap;
  let startX = (13.33 - totalW) / 2;
  const y = 2.3;

  steps.forEach((step, i) => {
    slide.addShape("roundRect", {
      x: startX,
      y,
      w: boxW,
      h: 1.5,
      rectRadius: 0.06,
      fill: { color: i % 2 === 0 ? "EAF2FF" : "F5F5F5" },
      line: { color: "BFBFBF" }
    });

    slide.addText(step.label || "", {
      x: startX + 0.1,
      y: y + 0.15,
      w: boxW - 0.2,
      h: 0.3,
      align: "center",
      fontSize: 15,
      bold: true,
      color: "1F1F1F"
    });

    slide.addText(step.description || "", {
      x: startX + 0.12,
      y: y + 0.5,
      w: boxW - 0.24,
      h: 0.75,
      align: "center",
      valign: "mid",
      fontSize: 10,
      color: "444444"
    });

    if (i < steps.length - 1) {
      slide.addText("→", {
        x: startX + boxW + 0.06,
        y: y + 0.48,
        w: 0.2,
        h: 0.3,
        fontSize: 20,
        color: "666666",
        align: "center"
      });
    }

    startX += boxW + gap;
  });
}

function renderTimeline(slide, s, globalStyle) {
  const t = titleStyle(globalStyle);

  slide.addText(s.title || "", {
    x: 0.8,
    y: 0.7,
    w: 10.0,
    h: 0.5,
    fontSize: 22,
    ...t
  });

  const events = s.content?.events || [];
  slide.addShape("line", {
    x: 1.1,
    y: 3.0,
    w: 10.8,
    h: 0,
    line: { color: "A6A6A6", pt: 1.5 }
  });

  const count = Math.max(events.length, 1);
  const step = 10.8 / Math.max(count - 1, 1);

  events.forEach((event, i) => {
    const x = 1.1 + i * step;
    slide.addShape("ellipse", {
      x: x - 0.08,
      y: 2.92,
      w: 0.16,
      h: 0.16,
      fill: { color: "2F75B5" },
      line: { color: "2F75B5" }
    });

    slide.addText(event.label || "", {
      x: x - 0.5,
      y: 2.3,
      w: 1.0,
      h: 0.3,
      fontSize: 12,
      bold: true,
      align: "center"
    });

    slide.addText(event.date_or_period || "", {
      x: x - 0.7,
      y: 3.25,
      w: 1.4,
      h: 0.25,
      fontSize: 10,
      bold: true,
      color: "666666",
      align: "center"
    });

    slide.addText(event.description || "", {
      x: x - 0.8,
      y: 3.55,
      w: 1.6,
      h: 0.9,
      fontSize: 10,
      align: "center",
      color: "444444"
    });
  });
}

function renderKpiDashboard(slide, s, globalStyle) {
  const t = titleStyle(globalStyle);

  slide.addText(s.title || "", {
    x: 0.8,
    y: 0.7,
    w: 10.0,
    h: 0.5,
    fontSize: 22,
    ...t
  });

  const kpis = s.content?.kpis || [];
  const count = Math.max(kpis.length, 1);
  const boxW = 3.4;
  const gap = 0.45;
  const totalW = count * boxW + (count - 1) * gap;
  let startX = (13.33 - totalW) / 2;
  const y = 2.1;

  kpis.forEach((kpi) => {
    slide.addShape("roundRect", {
      x: startX,
      y,
      w: boxW,
      h: 2.6,
      rectRadius: 0.08,
      fill: { color: "F7FAFE" },
      line: { color: "D9E8FB" }
    });

    slide.addText(kpi.label || "", {
      x: startX + 0.15,
      y: y + 0.22,
      w: boxW - 0.3,
      h: 0.35,
      fontSize: 14,
      bold: true,
      align: "center"
    });

    slide.addText(kpi.value || "", {
      x: startX + 0.15,
      y: y + 0.85,
      w: boxW - 0.3,
      h: 0.65,
      fontSize: 24,
      bold: true,
      align: "center",
      color: "2F75B5"
    });

    slide.addText(kpi.comment || "", {
      x: startX + 0.15,
      y: y + 1.75,
      w: boxW - 0.3,
      h: 0.5,
      fontSize: 10,
      align: "center",
      color: "555555"
    });

    startX += boxW + gap;
  });
}

function renderImageFocus(slide, s, globalStyle, assetsMap) {
  const t = titleStyle(globalStyle);

  slide.addText(s.title || "", {
    x: 0.8,
    y: 0.5,
    w: 10.0,
    h: 0.5,
    fontSize: 22,
    ...t
  });

  const assetRef = s.content?.asset_ref;
  const asset = assetRef ? assetsMap[assetRef] : null;
  const imgPath = asset?.file_name;

  if (imgPath && fs.existsSync(imgPath)) {
    slide.addImage({
      path: imgPath,
      x: 0.9,
      y: 1.2,
      w: 8.0,
      h: 4.8
    });
  } else {
    slide.addShape("rect", {
      x: 0.9,
      y: 1.2,
      w: 8.0,
      h: 4.8,
      fill: { color: "EFEFEF" },
      line: { color: "D0D0D0" }
    });
    slide.addText("Image placeholder", {
      x: 3.0,
      y: 3.3,
      w: 3.5,
      h: 0.4,
      fontSize: 16,
      color: "777777",
      align: "center"
    });
  }

  slide.addText(s.content?.supporting_text || "", {
    x: 9.3,
    y: 1.8,
    w: 3.1,
    h: 2.2,
    fontSize: 15,
    color: "333333",
    valign: "mid"
  });

  if (s.content?.caption) {
    slide.addText(s.content.caption, {
      x: 0.95,
      y: 6.15,
      w: 7.8,
      h: 0.25,
      fontSize: 9,
      italic: true,
      color: "666666"
    });
  }
}

function renderTable(slide, s, globalStyle) {
  const t = titleStyle(globalStyle);

  slide.addText(s.title || "", {
    x: 0.8,
    y: 0.5,
    w: 10.0,
    h: 0.5,
    fontSize: 22,
    ...t
  });

  const columns = s.content?.columns || [];
  const rows = s.content?.rows || [];
  const data = [];
  if (columns.length) data.push(columns);
  rows.forEach(r => data.push(r));

  slide.addTable(data, {
    x: 0.9,
    y: 1.4,
    w: 11.4,
    h: 4.8,
    border: { type: "solid", pt: 1, color: "D9D9D9" },
    fill: "FFFFFF",
    color: "333333",
    fontSize: 11,
    rowH: 0.5,
    autoFit: true,
    bold: false
  });
}

function renderClosing(slide, s, globalStyle) {
  const t = titleStyle(globalStyle);

  slide.addText(s.title || "Next steps", {
    x: 0.8,
    y: 0.8,
    w: 8.0,
    h: 0.5,
    fontSize: 24,
    ...t
  });

  const points = (s.content?.summary_points || []).map(text => ({ text, options: { bullet: { indent: 18 } } }));
  slide.addText(points, {
    x: 1.0,
    y: 1.8,
    w: 7.0,
    h: 3.5,
    fontSize: 18,
    color: "333333"
  });

  if (s.content?.call_to_action) {
    slide.addShape("roundRect", {
      x: 8.7,
      y: 2.2,
      w: 3.4,
      h: 1.6,
      rectRadius: 0.08,
      fill: { color: "2F75B5" },
      line: { color: "2F75B5" }
    });

    slide.addText(s.content.call_to_action, {
      x: 8.9,
      y: 2.7,
      w: 3.0,
      h: 0.5,
      fontSize: 18,
      bold: true,
      color: "FFFFFF",
      align: "center",
      valign: "mid"
    });
  }
}

function renderSlide(ppt, slideSpec, deck, globalStyle, assetsMap) {
  const slide = baseSlide(ppt, deck, globalStyle, assetsMap);

  switch (slideSpec.type) {
    case "title":
      renderTitleSlide(slide, slideSpec, globalStyle);
      break;
    case "section_divider":
      renderSectionDivider(slide, slideSpec, globalStyle);
      break;
    case "agenda":
      renderAgenda(slide, slideSpec, globalStyle);
      break;
    case "bullet_summary":
      renderBulletSummary(slide, slideSpec, globalStyle);
      break;
    case "two_column":
      renderTwoColumn(slide, slideSpec, globalStyle);
      break;
    case "comparison":
      renderComparison(slide, slideSpec, globalStyle);
      break;
    case "process_flow":
      renderProcessFlow(slide, slideSpec, globalStyle);
      break;
    case "timeline":
      renderTimeline(slide, slideSpec, globalStyle);
      break;
    case "kpi_dashboard":
      renderKpiDashboard(slide, slideSpec, globalStyle);
      break;
    case "image_focus":
      renderImageFocus(slide, slideSpec, globalStyle, assetsMap);
      break;
    case "table":
      renderTable(slide, slideSpec, globalStyle);
      break;
    case "closing":
      renderClosing(slide, slideSpec, globalStyle);
      break;
    default:
      renderBulletSummary(slide, {
        ...slideSpec,
        title: slideSpec.title || "Unsupported slide type",
        content: {
          intro: `Unsupported type: ${slideSpec.type}`,
          bullets: ["Renderer fallback was used."]
        }
      }, globalStyle);
  }

  addFooter(slide, ppt, globalStyle);
}

async function main() {
  const inputJson = process.argv[2];
  const outputPptx = process.argv[3] || "output.pptx";

  assert(inputJson, "Usage: node render_presentation.js <input.json> <output.pptx>");

  const spec = loadJson(inputJson);

  assert(spec.schema_version === "smart_presentation_v1", "schema_version must be smart_presentation_v1");
  assert(Array.isArray(spec.slides), "slides must be an array");

  const ppt = new PptxGenJS();
  ppt.layout = "LAYOUT_WIDE";
  ppt.author = "Smart PPT V1";
  ppt.company = spec.global_style?.brand_name || "Unknown";
  ppt.subject = spec.deck?.purpose || "";
  ppt.title = spec.deck?.title || "Presentation";
  ppt.lang = spec.deck?.language || "en-US";
  ppt.theme = {
    headFontFace: "Aptos",
    bodyFontFace: "Aptos",
    lang: spec.deck?.language || "en-US"
  };

  const assetsMap = {};
  for (const asset of spec.assets || []) {
    assetsMap[asset.id] = asset;
  }

  for (const slideSpec of spec.slides) {
    renderSlide(ppt, slideSpec, spec.deck || {}, spec.global_style || {}, assetsMap);
  }

  await ppt.writeFile({ fileName: outputPptx });
  console.log(`PowerPoint created: ${outputPptx}`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});