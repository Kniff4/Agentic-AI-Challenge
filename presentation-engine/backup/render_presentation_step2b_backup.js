const fs = require("fs");
const PptxGenJS = require("pptxgenjs");

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function loadJson(filePath) {
  const raw = fs.readFileSync(filePath, "utf8");
  return JSON.parse(raw);
}

function getTheme(spec) {
  if (spec.theme) return spec.theme;
  if (spec.global_style) return spec.global_style;
  return {};
}

function getColor(theme, key, fallback) {
  return theme?.color_tokens?.[key] || fallback;
}

function getFont(theme, key, fallback) {
  return theme?.font_tokens?.[key] || fallback;
}

function getBlocksByType(slideSpec, blockType) {
  return (slideSpec.content_blocks || []).filter(b => b.block_type === blockType);
}

function getFirstBlockByRole(slideSpec, role) {
  return (slideSpec.content_blocks || []).find(b => b.role === role);
}

function addFooter(slide, slideNumber, theme) {
  const footerText = theme?.footer_text || "";
  if (footerText) {
    slide.addText(footerText, {
      x: 0.3,
      y: 7.0,
      w: 5.0,
      h: 0.2,
      fontSize: 9,
      color: getColor(theme, "muted_text", "666666"),
      margin: 0
    });
  }

  if (theme?.show_slide_numbers) {
    slide.addText(String(slideNumber), {
      x: 12.5,
      y: 7.0,
      w: 0.4,
      h: 0.2,
      fontSize: 9,
      align: "right",
      color: getColor(theme, "muted_text", "666666"),
      margin: 0
    });
  }
}

function renderAccentBar(slide, slideSpec, theme) {
  const bars = (slideSpec.visual_elements || []).filter(
    el => el.element_type === "accent_bar" && el.position === "top"
  );

  if (bars.length > 0) {
    const bar = bars[0];
    slide.addShape("rect", {
      x: 0,
      y: 0,
      w: 13.33,
      h: 0.12,
      fill: { color: getColor(theme, bar.color_token || "primary", "2F75B5") },
      line: { color: getColor(theme, bar.color_token || "primary", "2F75B5") }
    });
  }
}

function baseSlide(ppt, theme, slideSpec) {
  const slide = ppt.addSlide();

  const bgType = slideSpec?.background?.type || "solid";
  if (bgType === "solid") {
    slide.background = {
      color: getColor(theme, slideSpec?.background?.color_token || "background", "FFFFFF")
    };
  } else {
    slide.background = {
      color: getColor(theme, "background", "FFFFFF")
    };
  }

  renderAccentBar(slide, slideSpec, theme);
  return slide;
}

function titleStyle(theme, slideSpec) {
  return {
    fontFace: getFont(theme, "title_font", "Aptos"),
    bold: true,
    color: getColor(theme, "text", "1F1F1F"),
    align: slideSpec?.style_overrides?.title_alignment || "left"
  };
}

function bodyStyle(theme) {
  return {
    fontFace: getFont(theme, "body_font", "Aptos"),
    color: getColor(theme, "text", "333333")
  };
}

function renderTitleSlide(slide, s, theme) {
  const t = titleStyle(theme, s);
  const b = bodyStyle(theme);

  const titleSize = s?.style_overrides?.title_size || getFont(theme, "title_size", 24);
  const subtitleSize = getFont(theme, "subtitle_size", 16);

  slide.addText(s.title || "", {
    x: 0.8,
    y: 1.2,
    w: 11.0,
    h: 0.8,
    fontSize: titleSize,
    ...t
  });

  if (s.subtitle) {
    slide.addText(s.subtitle, {
      x: 0.8,
      y: 2.15,
      w: 10.0,
      h: 0.45,
      fontSize: subtitleSize,
      ...b
    });
  }

  const taglineBlock = getFirstBlockByRole(s, "tagline");
  if (taglineBlock?.text) {
    slide.addText(taglineBlock.text, {
      x: 0.8,
      y: 2.85,
      w: 9.0,
      h: 0.35,
      fontSize: 12,
      italic: true,
      color: getColor(theme, "muted_text", "666666")
    });
  }
}

function renderBulletSummary(slide, s, theme) {
  const t = titleStyle(theme, s);
  const b = bodyStyle(theme);

  const titleSize = s?.style_overrides?.title_size || 22;
  const bodySize = s?.style_overrides?.body_size || 18;

  slide.addText(s.title || "", {
    x: 0.8,
    y: 0.6,
    w: 10.0,
    h: 0.5,
    fontSize: titleSize,
    ...t
  });

  if (s.subtitle) {
    slide.addText(s.subtitle, {
      x: 0.8,
      y: 1.1,
      w: 9.5,
      h: 0.3,
      fontSize: 11,
      color: getColor(theme, "muted_text", "666666")
    });
  }

  const introBlock = getFirstBlockByRole(s, "intro");
  if (introBlock?.text) {
    slide.addText(introBlock.text, {
      x: 0.9,
      y: 1.55,
      w: 11.0,
      h: 0.4,
      fontSize: 13,
      ...b
    });
  }

  const calloutBlock = getFirstBlockByRole(s, "highlight");
  let bulletY = 2.0;

  if (calloutBlock?.text) {
    const useCard = s?.style_overrides?.surface_style === "card";

    if (useCard) {
      slide.addShape("roundRect", {
        x: 0.9,
        y: 1.8,
        w: 11.2,
        h: 0.8,
        rectRadius: 0.05,
        fill: { color: getColor(theme, "secondary", "EAF2FF") },
        line: { color: getColor(theme, "secondary", "EAF2FF") }
      });
    }

    slide.addText(calloutBlock.text, {
      x: 1.15,
      y: 2.02,
      w: 10.6,
      h: 0.28,
      fontSize: 13,
      bold: true,
      color: getColor(theme, s?.style_overrides?.emphasis_color_token || "primary", "2F75B5"),
      align: "left"
    });

    bulletY = 2.95;
  }

  const bulletBlock =
    getFirstBlockByRole(s, "main_points") ||
    getFirstBlockByRole(s, "supporting_points") ||
    getBlocksByType(s, "bullet_list")[0];

  const bulletItems = (bulletBlock?.items || []).map(text => ({
    text,
    options: { bullet: { indent: 18 } }
  }));

  slide.addText(bulletItems, {
    x: 1.0,
    y: bulletY,
    w: 10.8,
    h: 3.5,
    fontSize: bodySize,
    paraSpaceAfterPt: 10,
    ...b
  });
}

function renderClosing(slide, s, theme) {
  const t = titleStyle(theme, s);
  const bodySize = s?.style_overrides?.body_size || 18;

  slide.addText(s.title || "Next steps", {
    x: 0.8,
    y: 0.8,
    w: 8.0,
    h: 0.5,
    fontSize: s?.style_overrides?.title_size || 24,
    ...t
  });

  if (s.subtitle) {
    slide.addText(s.subtitle, {
      x: 0.8,
      y: 1.28,
      w: 8.0,
      h: 0.25,
      fontSize: 11,
      color: getColor(theme, "muted_text", "666666")
    });
  }

  const bulletBlock =
    getFirstBlockByRole(s, "summary_points") ||
    getBlocksByType(s, "bullet_list")[0];

  const points = (bulletBlock?.items || []).map(text => ({
    text,
    options: { bullet: { indent: 18 } }
  }));

  slide.addText(points, {
    x: 1.0,
    y: 1.9,
    w: 7.0,
    h: 3.5,
    fontSize: bodySize,
    color: getColor(theme, "text", "333333")
  });

  const ctaBlock = getFirstBlockByRole(s, "call_to_action");

  if (ctaBlock?.text) {
    slide.addShape("roundRect", {
      x: 8.7,
      y: 2.2,
      w: 3.4,
      h: 1.6,
      rectRadius: 0.08,
      fill: { color: getColor(theme, s?.style_overrides?.emphasis_color_token || "primary", "2F75B5") },
      line: { color: getColor(theme, s?.style_overrides?.emphasis_color_token || "primary", "2F75B5") }
    });

    slide.addText(ctaBlock.text, {
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

function renderFallbackSlide(slide, s, theme) {
  slide.addText(s.title || "Unsupported slide", {
    x: 0.8,
    y: 0.8,
    w: 10.0,
    h: 0.5,
    fontSize: 22,
    bold: true,
    color: getColor(theme, "text", "1F1F1F")
  });

  slide.addText(`Type not yet implemented: ${s.type}`, {
    x: 0.9,
    y: 1.7,
    w: 10.0,
    h: 0.4,
    fontSize: 14,
    color: getColor(theme, "muted_text", "666666")
  });
}

function renderSlide(ppt, slideSpec, theme, slideNumber) {
  const slide = baseSlide(ppt, theme, slideSpec);

  switch (slideSpec.type) {
    case "title":
      renderTitleSlide(slide, slideSpec, theme);
      break;
    case "bullet_summary":
      renderBulletSummary(slide, slideSpec, theme);
      break;
    case "closing":
      renderClosing(slide, slideSpec, theme);
      break;
    default:
      renderFallbackSlide(slide, slideSpec, theme);
  }

  addFooter(slide, slideNumber, theme);
}

async function main() {
  const inputJson = process.argv[2];
  const outputPptx = process.argv[3] || "output_v2.pptx";

  assert(inputJson, "Usage: node render_presentation.js <input.json> <output.pptx>");

  const spec = loadJson(inputJson);

  assert(
    spec.schema_version === "smart_presentation_v1" ||
      spec.schema_version === "smart_presentation_v2",
    "schema_version must be smart_presentation_v1 or smart_presentation_v2"
  );

  assert(Array.isArray(spec.slides), "slides must be an array");

  const theme = getTheme(spec);

  const ppt = new PptxGenJS();
  ppt.layout = "LAYOUT_WIDE";
  ppt.author = "Smart PPT Renderer";
  ppt.company = theme?.brand_name || "Unknown";
  ppt.subject = spec.deck?.purpose || "";
  ppt.title = spec.deck?.title || "Presentation";
  ppt.lang = spec.deck?.language || "en-US";

  spec.slides.forEach((slideSpec, index) => {
    renderSlide(ppt, slideSpec, theme, index + 1);
  });

  await ppt.writeFile({ fileName: outputPptx });
  console.log(`PowerPoint created: ${outputPptx}`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});