#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const JSZip = require("jszip");
const PptxGenJS = require("pptxgenjs");
const {
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
} = require("./helpers/layout");

function parseArgs(argv) {
  const args = { input: "", output: "" };
  for (let i = 2; i < argv.length; i += 1) {
    const token = argv[i];
    if (token === "--input") {
      args.input = argv[i + 1] || "";
      i += 1;
    } else if (token === "--output") {
      args.output = argv[i + 1] || "";
      i += 1;
    }
  }
  if (!args.input || !args.output) {
    throw new Error("Usage: node generate_deck.js --input <payload.json> --output <out.pptx>");
  }
  return args;
}

function mustReadJson(filePath) {
  const raw = fs.readFileSync(filePath, "utf8").replace(/^\uFEFF/, "");
  return JSON.parse(raw);
}

function toHexColor(rgb) {
  const r = Math.max(0, Math.min(255, Number(rgb[0] || 0)));
  const g = Math.max(0, Math.min(255, Number(rgb[1] || 0)));
  const b = Math.max(0, Math.min(255, Number(rgb[2] || 0)));
  return `${r.toString(16).padStart(2, "0")}${g.toString(16).padStart(2, "0")}${b.toString(16).padStart(2, "0")}`.toUpperCase();
}

function clampColor(value) {
  return Math.max(0, Math.min(255, Math.round(Number(value) || 0)));
}

function normalizeStyleKey(styleRaw) {
  const val = String(styleRaw || "management").toLowerCase();
  if (val.includes("technical") || val.includes("tech")) {
    return "technical";
  }
  // Legacy styles are normalized to the two supported modes.
  if (val.includes("academic") || val.includes("research")) {
    return "technical";
  }
  if (val.includes("creative") || val.includes("modern")) {
    return "management";
  }
  return "management";
}

const STYLE_PRESETS = {
  management: [
    {
      bg: [236, 243, 251],
      header: [27, 63, 101],
      text: [29, 45, 64],
      accent: [52, 122, 191],
    },
    {
      bg: [247, 242, 233],
      header: [88, 64, 37],
      text: [66, 49, 33],
      accent: [190, 123, 56],
    },
    {
      bg: [233, 244, 238],
      header: [26, 82, 67],
      text: [30, 54, 47],
      accent: [67, 150, 118],
    },
  ],
  technical: [
    {
      bg: [229, 236, 247],
      header: [24, 43, 77],
      text: [28, 40, 60],
      accent: [66, 104, 181],
    },
    {
      bg: [235, 244, 250],
      header: [20, 64, 92],
      text: [28, 48, 64],
      accent: [43, 137, 186],
    },
    {
      bg: [237, 241, 252],
      header: [45, 55, 105],
      text: [33, 40, 74],
      accent: [88, 106, 202],
    },
  ],
};

function deriveTheme(seedText, styleRaw) {
  const digest = crypto.createHash("md5").update(String(seedText || "default")).digest();
  const styleKey = normalizeStyleKey(styleRaw);
  const variants = STYLE_PRESETS[styleKey] || STYLE_PRESETS.management;
  const preset = variants[digest[0] % variants.length];

  const jitter = (rgb, offset, amp = 12) => {
    const range = amp * 2 + 1;
    return rgb.map((v, idx) => clampColor(v + ((digest[(offset + idx) % digest.length] % range) - amp)));
  };

  const bg = jitter(preset.bg, 1, 8);
  const header = jitter(preset.header, 5, 10);
  const text = jitter(preset.text, 9, 8);
  const accent = jitter(preset.accent, 13, 12);

  const muted = [
    clampColor(text[0] + 36),
    clampColor(text[1] + 34),
    clampColor(text[2] + 34),
  ];
  const line = [
    clampColor((bg[0] + accent[0]) / 2 + 30),
    clampColor((bg[1] + accent[1]) / 2 + 24),
    clampColor((bg[2] + accent[2]) / 2 + 22),
  ];

  return {
    styleKey,
    bg: toHexColor(bg),
    header: toHexColor(header),
    headerText: "F8FBFF",
    card: "FCFEFF",
    cardAlt: toHexColor([
      clampColor(bg[0] - 8),
      clampColor(bg[1] - 5),
      clampColor(bg[2] - 4),
    ]),
    text: toHexColor(text),
    muted: toHexColor(muted),
    line: toHexColor(line),
    accent: toHexColor(accent),
  };
}
function stripEllipsis(text) {
  return String(text || "")
    .replace(/[ \t]{2,}/g, " ")
    .trim();
}

function truncate(text, limit) {
  const value = stripEllipsis(text);
  if (!Number.isFinite(limit) || limit <= 0) return value;
  if (value.length <= limit) return value;
  return value.slice(0, Math.max(0, limit)).trim();
}

function normalizeTopicTitle(raw) {
  let txt = String(raw || "").trim();
  if (!txt) return "";
  txt = txt.replace(/^\u5927\u7eb2(?:[\uff1a:\-|])?\s*/u, "");
  txt = txt.replace(/^(?:outline|agenda|contents)[\s:\-|]*/i, "");
  txt = txt.replace(/\s+/g, " ").trim();
  return txt;
}

function normalizeTocItem(raw) {
  let txt = String(raw || "").trim();
  if (!txt) return "";
  txt = txt.replace(/^\u5927\u7eb2(?:[\uff1a:\-|])?\s*/u, "");
  txt = txt.replace(/^(?:outline|agenda|contents)[\s:\-|]*/i, "");
  txt = txt.replace(/^\u7b2c\s*\d+\s*\u9875[\uff1a:]\s*/u, "");
  txt = txt.replace(/^\d+\s*[\.\u3001\)\uff09]\s*/u, "");
  return txt.trim();
}

function isCoverLikeTitle(title) {
  const low = String(title || "").toLowerCase();
  return (
    low.includes("\u5c01\u9762") ||
    low.includes("\u6807\u9898\u9875") ||
    low.includes("cover") ||
    low.includes("title")
  );
}

function filterRelevantTocItems(items, topic) {
  const topicKey = normalizeCompareKey(topic);
  const out = [];
  const seen = new Set();

  for (const raw of (Array.isArray(items) ? items : [])) {
    const item = normalizeTocItem(raw);
    if (!item) continue;

    const key = normalizeCompareKey(item);
    if (!key || seen.has(key)) continue;

    if (isTocLikeTitle(item) || isCoverLikeTitle(item)) {
      continue;
    }

    if (topicKey && (key === topicKey || key.startsWith(topicKey) || topicKey.startsWith(key))) {
      continue;
    }

    out.push(item);
    seen.add(key);
    if (out.length >= 10) break;
  }

  return out;
}

function isTocLikeTitle(title) {
  const low = String(title || "").toLowerCase();
  return (
    low.includes("\u5927\u7eb2") ||
    low.includes("\u76ee\u5f55") ||
    low.includes("agenda") ||
    low.includes("contents") ||
    low.includes("toc")
  );
}

function contentSlides(slides) {
  const items = Array.isArray(slides) ? slides : [];
  return items.filter((s) => {
    const title = String((s && s.title) || "");
    const low = title.toLowerCase();
    const slideType = String((s && s.slide_type) || "").toLowerCase();
    if (slideType === "title" || slideType === "toc") return false;
    if (low.includes("cover") || low.includes("agenda")) return false;
    if (isTocLikeTitle(title)) return false;
    return true;
  });
}

function addWarnings(slide, pptx) {
  try {
    warnIfSlideHasOverlaps(slide, pptx);
    warnIfSlideElementsOutOfBounds(slide, pptx);
  } catch (err) {
    console.warn(`[pptx_generator] warning checks failed: ${String(err && err.message ? err.message : err)}`);
  }
}

function addTitleBar(slide, title, theme) {
  slide.addShape("roundRect", {
    x: 0.45,
    y: 0.25,
    w: 12.4,
    h: 0.9,
    radius: 0.08,
    fill: { color: theme.header },
    line: { color: theme.header, pt: 0 },
  });
  slide.addText(truncate(title, 78), {
    x: 0.75,
    y: 0.42,
    w: 11.8,
    h: 0.54,
    fontFace: "Microsoft YaHei",
    bold: true,
    color: theme.headerText,
    fontSize: 24,
    valign: "mid",
  });
}

function addImageIfExists(slide, imagePath, x, y, w, h) {
  if (!imagePath) return;
  if (!fs.existsSync(imagePath)) return;
  slide.addShape("roundRect", {
    x,
    y,
    w,
    h,
    radius: 0.06,
    fill: { color: "F8FBFF" },
    line: { color: "D6E0F0", pt: 1 },
  });
  slide.addImage({ path: imagePath, x: x + 0.08, y: y + 0.08, w: Math.max(0.3, w - 0.16), h: Math.max(0.3, h - 0.16) });
}

function addBulletList(slide, bullets, x, y, w, h, theme, options = {}) {
  const maxTextLength = Number.isFinite(options.maxTextLength) ? Number(options.maxTextLength) : 120;
  const clean = (Array.isArray(bullets) ? bullets : [])
    .map((item) => {
      const value = stripEllipsis(item);
      if (!value) return "";
      if (maxTextLength <= 0) return value;
      return truncate(value, Math.max(16, maxTextLength));
    })
    .filter(Boolean);
  const sourceItems = clean.length > 0 ? clean : [];

  if (sourceItems.length === 0) return;

  const maxItems = Number.isFinite(options.maxItems) ? Math.max(1, Number(options.maxItems)) : 6;
  const minLineHeight = Number.isFinite(options.minLineHeight) ? Math.max(0.36, Number(options.minLineHeight)) : 0.5;
  const usableHeight = Math.max(0.42, h - 0.04);
  const maxItemsByHeight = Math.max(1, Math.floor(usableHeight / minLineHeight));
  const items = sourceItems.slice(0, Math.min(maxItems, maxItemsByHeight));

  let cursorY = y;
  const lineHeight = Math.max(minLineHeight, Math.min(0.95, usableHeight / items.length));
  const fontSize = lineHeight >= 0.82 ? 17 : lineHeight >= 0.66 ? 15 : 13;
  const bulletSize = Math.max(0.12, Math.min(0.2, lineHeight * 0.3));

  for (const item of items) {
    const centerY = cursorY + lineHeight / 2;
    slide.addShape("ellipse", {
      x,
      y: centerY - bulletSize / 2,
      w: bulletSize,
      h: bulletSize,
      fill: { color: theme.accent },
      line: { color: theme.accent, pt: 0 },
    });

    slide.addText(item, {
      x: x + bulletSize + 0.14,
      y: cursorY,
      w: Math.max(0.25, w - bulletSize - 0.18),
      h: lineHeight,
      fontFace: "Microsoft YaHei",
      fontSize,
      color: theme.muted,
      valign: "mid",
      breakLine: true,
      fit: "shrink",
      margin: 0,
    });
    cursorY += lineHeight;
  }
}

function normalizeCompareKey(raw) {
  return String(raw || "")
    .toLowerCase()
    .replace(/[\s\u3000]+/g, "")
    .replace(/[\u3002\uFF0C\uFF1B\uFF1A\u3001,.;:!?\uFF1F\uFF01\-_=+\(\)\[\]{}<>\u300A\u300B\u3010\u3011\"'`]/g, "")
    .trim();
}

function dedupeTextItems(items, limit = 4, maxLen = 220) {
  const out = [];
  const seen = new Set();

  for (const item of (Array.isArray(items) ? items : [])) {
    const txt = truncate(item, maxLen);
    if (!txt) continue;

    const key = normalizeCompareKey(txt);
    if (!key || seen.has(key)) continue;

    out.push(txt);
    seen.add(key);
    if (out.length >= limit) break;
  }

  return out;
}

function isMetaNoteLine(line) {
  const txt = String(line || "").trim();
  if (!txt) return true;
  const low = txt.toLowerCase();
  if (/^(\u89c6\u89c9\u5143\u7d20|\u89c6\u89c9\u7126\u70b9|\u6392\u7248\u5e03\u5c40|\u6f14\u8bb2\u8005\u5907\u6ce8|\u5907\u6ce8|\u8bf4\u660e)[\uff1a:]/u.test(txt)) return true;
  if (/^(visual|layout|speaker|note)[\s:]/i.test(low)) return true;
  return false;
}

function cleanTextItems(items, limit = 4, maxLen = 220) {
  return dedupeTextItems(items, limit, maxLen);
}

function getDetailPoints(data, limit = 8) {
  const source = Array.isArray(data && data.detail_points) && data.detail_points.length > 0
    ? data.detail_points
    : (Array.isArray(data && data.bullets) ? data.bullets : []);
  return cleanTextItems(source, limit, 140);
}

function getTextBlocks(data, limit = 4) {
  const source = Array.isArray(data && data.text_blocks) ? data.text_blocks : [];
  const blocks = cleanTextItems(source, limit, 240);
  if (blocks.length > 0) {
    return blocks;
  }
  const summary = String((data && data.summary_text) || "").trim();
  if (summary) {
    return [truncate(summary, 240)];
  }
  return [];
}

function getSummaryText(data) {
  const summary = String((data && data.summary_text) || "").trim();
  if (summary) return truncate(summary, 240);
  const blocks = getTextBlocks(data, 1);
  if (blocks.length > 0) return blocks[0];
  const points = getDetailPoints(data, 1);
  if (points.length > 0) return points[0];
  return "Overview will be completed with project context.";
}

function padItems(items, target) {
  return cleanTextItems(items, target, 140).slice(0, target);
}

function addTextCard(slide, text, x, y, w, h, theme, options = {}) {
  const fillColor = options.fillColor || theme.card;
  const lineColor = options.lineColor || theme.line;
  const textColor = options.textColor || theme.text;
  const radius = Number.isFinite(options.radius) ? Number(options.radius) : 0.06;

  slide.addShape("roundRect", {
    x,
    y,
    w,
    h,
    radius,
    fill: { color: fillColor },
    line: { color: lineColor, pt: 1 },
  });

  const title = String(options.title || "").trim();
  if (title) {
    slide.addText(title, {
      x: x + 0.2,
      y: y + 0.12,
      w: Math.max(0.3, w - 0.4),
      h: 0.24,
      fontFace: "Microsoft YaHei",
      bold: true,
      color: theme.muted,
      fontSize: 11,
    });
  }

  slide.addText(String(text || ""), {
    x: x + 0.2,
    y: y + (title ? 0.38 : 0.18),
    w: Math.max(0.3, w - 0.4),
    h: Math.max(0.35, h - (title ? 0.5 : 0.3)),
    fontFace: "Microsoft YaHei",
    color: textColor,
    fontSize: options.fontSize || 14,
    valign: "top",
    breakLine: true,
    fit: "shrink",
    margin: 0,
  });
}

// Explicit mapping between content needs and layout families.
const SUMMARY_LAYOUT_RULES = {
  narrative_banner: { variants: [6], requirement: "Single long summary paragraph, no bullet list" },
  narrative_split: { variants: [11], requirement: "Two narrative blocks or top-bottom storytelling" },
  top_bottom_story: { variants: [4, 11], requirement: "One overview paragraph then detailed explanations" },
  left_summary_right_points: { variants: [5], requirement: "Overview + 2-4 supporting points" },
  summary_plus_two: { variants: [3, 5], requirement: "Overview + exactly 2 key points" },
  summary_plus_three: { variants: [3, 7], requirement: "Overview + exactly 3 key points" },
  summary_plus_four: { variants: [4, 8], requirement: "Overview + 4 key points" },
  points_three_columns: { variants: [7], requirement: "Three peer points" },
  points_four_grid: { variants: [8, 10], requirement: "Four peer points" },
  points_five_split: { variants: [0, 1, 2], requirement: "5+ points, list-heavy content" },
  center_hub_four: { variants: [9], requirement: "Central overview and surrounding supporting points" },
  four_quadrant_mixed: { variants: [10], requirement: "Four mixed text blocks" },
  mixed_content: { variants: [0, 1, 2], requirement: "General mixed content" },
  typed_slide: { variants: [0, 1], requirement: "Fallback when typed slides are rendered as summary" },
  auto: { variants: [0, 1, 2], requirement: "Fallback auto selection" },
};

function normalizeSummaryProfile(raw) {
  const val = String(raw || "").trim().toLowerCase();
  if (!val) return "";
  if (SUMMARY_LAYOUT_RULES[val]) return val;
  if (val.includes("summary_plus") && SUMMARY_LAYOUT_RULES[val]) return val;
  return "";
}

function inferSummaryProfile(data) {
  const hinted = normalizeSummaryProfile(data && (data.layout_profile || data.content_format));
  if (hinted) return hinted;

  const points = getDetailPoints(data, 8);
  const blocks = getTextBlocks(data, 4);
  const hasSummary = Boolean(String((data && data.summary_text) || "").trim()) || blocks.length > 0;

  if (points.length === 0) {
    return blocks.length >= 2 ? "narrative_split" : "narrative_banner";
  }
  if (hasSummary && points.length === 2) return "summary_plus_two";
  if (hasSummary && points.length === 3) return "summary_plus_three";
  if (hasSummary && points.length >= 4 && blocks.length >= 2) return "center_hub_four";
  if (hasSummary && points.length >= 4) return "summary_plus_four";
  if (points.length === 3) return "points_three_columns";
  if (points.length === 4) return "points_four_grid";
  if (points.length >= 5) return "points_five_split";
  if (hasSummary && blocks.length >= 2) return "top_bottom_story";
  if (hasSummary) return "left_summary_right_points";
  if (blocks.length >= 3) return "four_quadrant_mixed";
  return "mixed_content";
}

function pickSummaryVariant(topic, data, index, styleKey = "management") {
  const profile = inferSummaryProfile(data);
  const rule = SUMMARY_LAYOUT_RULES[profile] || SUMMARY_LAYOUT_RULES.auto;
  const variants = Array.isArray(rule.variants) && rule.variants.length > 0 ? rule.variants : [0];
  const seed = `${topic}|${styleKey}|summary|${index}|${profile}|${getDetailPoints(data, 8).length}|${getTextBlocks(data, 4).length}`;
  return variants[hashInt(seed) % variants.length];
}

function renderCover(pptx, topic, theme, subtitle, variant = 0) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };

  const titleText = truncate(topic || "Report", 60);
  const subtitleText = String(subtitle || "").trim();

  if (variant === 1) {
    slide.addShape("rect", {
      x: 0,
      y: 0,
      w: 13.333,
      h: 2.2,
      fill: { color: theme.header, transparency: 12 },
      line: { color: theme.header, pt: 0 },
    });
    slide.addShape("roundRect", {
      x: 0.95,
      y: 1.35,
      w: 11.4,
      h: 4.7,
      radius: 0.12,
      fill: { color: theme.card },
      line: { color: theme.line, pt: 1.2 },
    });
    slide.addShape("rect", {
      x: 0.95,
      y: 1.35,
      w: 0.32,
      h: 4.7,
      fill: { color: theme.accent },
      line: { color: theme.accent, pt: 0 },
    });
    slide.addText(titleText, {
      x: 1.45,
      y: 2.0,
      w: 10.6,
      h: 1.7,
      fontFace: "Microsoft YaHei",
      bold: true,
      color: theme.text,
      fontSize: 42,
      valign: "top",
    });
    if (subtitleText) {
      slide.addText(subtitleText, {
        x: 1.45,
        y: 4.15,
        w: 10.6,
        h: 0.7,
        fontFace: "Microsoft YaHei",
        color: theme.muted,
        fontSize: 20,
        valign: "top",
      });
    }
    addWarnings(slide, pptx);
    return;
  }

  if (variant === 2) {
    slide.addShape("roundRect", {
      x: 0.75,
      y: 0.9,
      w: 11.9,
      h: 5.7,
      radius: 0.2,
      fill: { color: theme.cardAlt },
      line: { color: theme.line, pt: 1.2 },
    });
    slide.addShape("line", {
      x: 1.15,
      y: 3.95,
      w: 10.9,
      h: 0,
      line: { color: theme.accent, pt: 2.2 },
    });
    slide.addText(titleText, {
      x: 1.35,
      y: 1.95,
      w: 10.5,
      h: 1.7,
      fontFace: "Microsoft YaHei",
      bold: true,
      color: theme.text,
      fontSize: 40,
      align: "center",
      valign: "mid",
    });
    if (subtitleText) {
      slide.addText(subtitleText, {
        x: 1.35,
        y: 4.28,
        w: 10.5,
        h: 0.6,
        fontFace: "Microsoft YaHei",
        color: theme.muted,
        fontSize: 19,
        align: "center",
        valign: "top",
      });
    }
    addWarnings(slide, pptx);
    return;
  }

  slide.addShape("rect", {
    x: 0,
    y: 0,
    w: 13.333,
    h: 7.5,
    fill: { color: theme.header, transparency: 32 },
    line: { color: theme.header, pt: 0 },
  });

  slide.addShape("roundRect", {
    x: 0.9,
    y: 1.1,
    w: 11.6,
    h: 5.2,
    radius: 0.12,
    fill: { color: theme.cardAlt },
    line: { color: theme.line, pt: 1.2 },
  });

  slide.addText(titleText, {
    x: 1.35,
    y: 2.0,
    w: 10.7,
    h: 1.5,
    fontFace: "Microsoft YaHei",
    bold: true,
    color: theme.text,
    fontSize: 44,
    valign: "top",
  });

  if (subtitleText) {
    slide.addText(subtitleText, {
      x: 1.35,
      y: 3.85,
      w: 10.7,
      h: 0.7,
      fontFace: "Microsoft YaHei",
      color: theme.muted,
      fontSize: 20,
      valign: "top",
    });
  }

  addWarnings(slide, pptx);
}
function renderToc(pptx, topic, outline, theme, variant = 0) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };
  addTitleBar(slide, "\u76ee\u5f55", theme);

  const items = filterRelevantTocItems(outline, topic);

  if (variant === 1) {
    slide.addShape("roundRect", {
      x: 0.8,
      y: 1.45,
      w: 11.8,
      h: 5.75,
      radius: 0.08,
      fill: { color: theme.cardAlt },
      line: { color: theme.line, pt: 1.2 },
    });

    const cols = 2;
    items.forEach((item, idx) => {
      const col = idx % cols;
      const row = Math.floor(idx / cols);
      const x = 1.15 + col * 5.65;
      const y = 1.85 + row * 1.0;

      slide.addShape("roundRect", {
        x,
        y,
        w: 5.1,
        h: 0.82,
        radius: 0.05,
        fill: { color: theme.card },
        line: { color: theme.line, pt: 0.8 },
      });

      slide.addText(`${idx + 1}.`, {
        x: x + 0.2,
        y: y + 0.2,
        w: 0.45,
        h: 0.34,
        fontFace: "Microsoft YaHei",
        bold: true,
        color: theme.accent,
        fontSize: 14,
      });

      slide.addText(truncate(item, 56), {
        x: x + 0.75,
        y: y + 0.18,
        w: 4.2,
        h: 0.42,
        fontFace: "Microsoft YaHei",
        color: theme.text,
        fontSize: 15,
        valign: "mid",
      });
    });

    addWarnings(slide, pptx);
    return;
  }

  if (variant === 2) {
    slide.addShape("line", {
      x: 1.2,
      y: 1.75,
      w: 0,
      h: 5.2,
      line: { color: theme.accent, pt: 2.2 },
    });

    let y = 1.8;
    items.forEach((item, idx) => {
      slide.addShape("ellipse", {
        x: 1.05,
        y: y + 0.13,
        w: 0.3,
        h: 0.3,
        fill: { color: theme.accent },
        line: { color: theme.accent, pt: 0 },
      });
      slide.addText(`${idx + 1}. ${truncate(item, 66)}`, {
        x: 1.55,
        y,
        w: 10.4,
        h: 0.55,
        fontFace: "Microsoft YaHei",
        color: theme.text,
        fontSize: 19,
        valign: "mid",
      });
      y += 0.64;
    });

    addWarnings(slide, pptx);
    return;
  }

  slide.addShape("roundRect", {
    x: 0.8,
    y: 1.45,
    w: 11.8,
    h: 5.7,
    radius: 0.08,
    fill: { color: theme.card },
    line: { color: theme.line, pt: 1.2 },
  });

  let y = 1.8;
  items.forEach((item, idx) => {
    slide.addShape("ellipse", {
      x: 1.15,
      y: y + 0.05,
      w: 0.35,
      h: 0.35,
      fill: { color: theme.accent },
      line: { color: theme.accent, pt: 0 },
    });
    slide.addText(String(idx + 1), {
      x: 1.22,
      y: y + 0.07,
      w: 0.2,
      h: 0.2,
      fontFace: "Microsoft YaHei",
      bold: true,
      color: "FFFFFF",
      fontSize: 10,
      valign: "mid",
      align: "center",
    });
    slide.addText(truncate(item, 68), {
      x: 1.6,
      y,
      w: 10.2,
      h: 0.52,
      fontFace: "Microsoft YaHei",
      color: theme.text,
      fontSize: 19,
      valign: "mid",
    });
    y += 0.63;
  });

  addWarnings(slide, pptx);
}
function renderSummarySlide(pptx, data, theme, variant = 0) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };
  addTitleBar(slide, String(data.title || "Content"), theme);

  const hasImage = Boolean(data.generated_image_path && fs.existsSync(data.generated_image_path));
  const points = dedupeTextItems(getDetailPoints(data, 8), 8, 140);
  const overview = getSummaryText(data);
  const blocks = dedupeTextItems(getTextBlocks(data, 4).filter((x) => normalizeCompareKey(x) !== normalizeCompareKey(overview)), 4, 220);

  if (variant === 1) {
    slide.addShape("roundRect", {
      x: 0.8,
      y: 1.45,
      w: 11.9,
      h: 5.8,
      radius: 0.08,
      fill: { color: theme.card },
      line: { color: theme.line, pt: 1.2 },
    });
    slide.addShape("rect", {
      x: 0.8,
      y: 1.45,
      w: 0.26,
      h: 5.8,
      fill: { color: theme.accent },
      line: { color: theme.accent, pt: 0 },
    });
    if (hasImage) {
      addBulletList(slide, points, 1.25, 1.9, 6.9, 5.0, theme, { maxItems: 6 });
      addImageIfExists(slide, data.generated_image_path, 8.25, 1.95, 4.05, 4.9);
    } else {
      addBulletList(slide, points, 1.25, 1.9, 10.9, 5.0, theme, { maxItems: 6 });
    }
    addWarnings(slide, pptx);
    return;
  }

  if (variant === 2) {
    slide.addShape("roundRect", {
      x: 0.9,
      y: 1.55,
      w: 11.5,
      h: 5.65,
      radius: 0.08,
      fill: { color: theme.cardAlt },
      line: { color: theme.line, pt: 1.2 },
    });
    const cardItems = padItems(points.length > 0 ? points : blocks, 4, "Point");
    cardItems.forEach((item, idx) => {
      const col = idx % 2;
      const row = Math.floor(idx / 2);
      const x = 1.2 + col * 5.6;
      const y = 1.95 + row * 2.35;
      addTextCard(slide, item, x, y, 5.1, 2.05, theme, {
        fillColor: theme.card,
        title: `Key ${idx + 1}`,
        fontSize: 14,
      });
    });
    if (hasImage) {
      addImageIfExists(slide, data.generated_image_path, 8.7, 5.05, 2.95, 2.0);
    }
    addWarnings(slide, pptx);
    return;
  }

  if (variant === 3) {
    addTextCard(slide, overview, 0.9, 1.55, 11.5, 2.0, theme, {
      fillColor: theme.cardAlt,
      title: "Overview",
      fontSize: 15,
    });

    const cols = padItems(points, 3, "Point");
    cols.forEach((item, idx) => {
      const x = 0.95 + idx * 3.85;
      addTextCard(slide, item, x, 3.85, 3.65, 3.35, theme, {
        fillColor: idx % 2 === 0 ? theme.card : theme.cardAlt,
        title: `Detail ${idx + 1}`,
        fontSize: 13,
      });
    });
    addWarnings(slide, pptx);
    return;
  }

  if (variant === 4) {
    addTextCard(slide, overview, 0.9, 1.55, 11.5, 2.25, theme, {
      fillColor: theme.cardAlt,
      title: "Summary",
      fontSize: 15,
    });
    slide.addShape("roundRect", {
      x: 0.9,
      y: 3.95,
      w: 11.5,
      h: 3.25,
      radius: 0.08,
      fill: { color: theme.card },
      line: { color: theme.line, pt: 1.2 },
    });
    addBulletList(slide, points, 1.25, 4.3, 10.8, 2.55, theme, { maxItems: 5, minLineHeight: 0.46 });
    addWarnings(slide, pptx);
    return;
  }

  if (variant === 5) {
    addTextCard(slide, overview, 0.9, 1.55, 5.25, 5.65, theme, {
      fillColor: theme.cardAlt,
      title: "Core Message",
      fontSize: 16,
      fit: "shrink",
    });
    slide.addShape("roundRect", {
      x: 6.35,
      y: 1.55,
      w: 6.05,
      h: 5.65,
      radius: 0.08,
      fill: { color: theme.card },
      line: { color: theme.line, pt: 1.2 },
    });
    addBulletList(slide, points, 6.7, 2.0, 5.35, 4.95, theme, { maxItems: 5 });
    addWarnings(slide, pptx);
    return;
  }

  if (variant === 6) {
    const narrative = [overview].concat(blocks.slice(1)).join("\n\n");
    slide.addShape("roundRect", {
      x: 0.9,
      y: 1.55,
      w: 11.5,
      h: 5.65,
      radius: 0.08,
      fill: { color: theme.card },
      line: { color: theme.line, pt: 1.2 },
    });
    slide.addText(narrative, {
      x: 1.2,
      y: 2.0,
      w: hasImage ? 7.0 : 10.9,
      h: 4.95,
      fontFace: "Microsoft YaHei",
      color: theme.text,
      fontSize: 16,
      valign: "top",
      breakLine: true,
      fit: "shrink",
      margin: 0,
    });
    if (hasImage) {
      addImageIfExists(slide, data.generated_image_path, 8.35, 2.05, 3.85, 4.75);
    }
    addWarnings(slide, pptx);
    return;
  }

  if (variant === 7) {
    const cols = padItems(points, 3, "Point");
    cols.forEach((item, idx) => {
      addTextCard(slide, item, 0.9 + idx * 3.85, 1.75, 3.65, 5.35, theme, {
        fillColor: idx % 2 === 0 ? theme.card : theme.cardAlt,
        title: `Track ${idx + 1}`,
        fontSize: 14,
      });
    });
    addWarnings(slide, pptx);
    return;
  }

  if (variant === 8) {
    const items = padItems(points.length > 0 ? points : blocks, 4, "Point");
    items.forEach((item, idx) => {
      const col = idx % 2;
      const row = Math.floor(idx / 2);
      addTextCard(slide, item, 0.95 + col * 5.75, 1.75 + row * 2.7, 5.45, 2.45, theme, {
        fillColor: row % 2 === 0 ? theme.card : theme.cardAlt,
        title: `Block ${idx + 1}`,
        fontSize: 13,
      });
    });
    addWarnings(slide, pptx);
    return;
  }

  if (variant === 9) {
    const ringItems = padItems(points, 4, "Point");
    addTextCard(slide, overview, 4.05, 3.5, 5.25, 1.72, theme, {
      fillColor: theme.cardAlt,
      title: "Core Summary",
      fontSize: 14,
    });

    const slots = [
      { x: 0.95, y: 1.85 },
      { x: 8.75, y: 1.85 },
      { x: 0.95, y: 5.35 },
      { x: 8.75, y: 5.35 },
    ];
    ringItems.forEach((item, idx) => {
      addTextCard(slide, item, slots[idx].x, slots[idx].y, 3.6, 1.55, theme, {
        fillColor: theme.card,
        fontSize: 12,
      });
    });
    addWarnings(slide, pptx);
    return;
  }

  if (variant === 10) {
    const mixed = padItems(blocks.length > 0 ? blocks : points, 4, "Section");
    mixed.forEach((item, idx) => {
      const col = idx % 2;
      const row = Math.floor(idx / 2);
      addTextCard(slide, item, 0.95 + col * 5.75, 1.75 + row * 2.7, 5.45, 2.45, theme, {
        fillColor: idx % 2 === 0 ? theme.cardAlt : theme.card,
        title: `Section ${idx + 1}`,
        fontSize: 13,
      });
    });
    addWarnings(slide, pptx);
    return;
  }

  if (variant === 11) {
    const topBlocks = padItems(blocks.length > 0 ? blocks : [overview], 2);
    const topSlots = [
      { x: 0.9, y: 1.55, w: 5.6, h: 2.45, fillColor: theme.cardAlt, title: "Context" },
      { x: 6.8, y: 1.55, w: 5.6, h: 2.45, fillColor: theme.card, title: "Focus" },
    ];
    topSlots.slice(0, topBlocks.length).forEach((slot, idx) => {
      addTextCard(slide, topBlocks[idx], slot.x, slot.y, slot.w, slot.h, theme, {
        fillColor: slot.fillColor,
        title: slot.title,
        fontSize: 14,
      });
    });

    slide.addShape("roundRect", {
      x: 0.9,
      y: 4.2,
      w: 11.5,
      h: 3.0,
      radius: 0.08,
      fill: { color: theme.card },
      line: { color: theme.line, pt: 1.2 },
    });
    addBulletList(slide, points, 1.2, 4.55, 10.9, 2.35, theme, { maxItems: 4, minLineHeight: 0.5 });
    addWarnings(slide, pptx);
    return;
  }

  slide.addShape("roundRect", {
    x: 0.85,
    y: 1.5,
    w: 11.6,
    h: 5.75,
    radius: 0.08,
    fill: { color: theme.card },
    line: { color: theme.line, pt: 1.2 },
  });

  if (hasImage) {
    addBulletList(slide, points, 1.2, 1.95, 7.0, 4.9, theme);
    addImageIfExists(slide, data.generated_image_path, 8.35, 1.95, 3.95, 4.8);
  } else {
    addBulletList(slide, points, 1.2, 1.95, 10.9, 4.9, theme);
  }

  addWarnings(slide, pptx);
}

function extractRiskHeadingHint(lines) {
  const stopWords = new Set([
    "\u89c6\u89c9\u5143\u7d20",
    "\u89c6\u89c9\u7126\u70b9",
    "\u6392\u7248\u5e03\u5c40",
    "\u6f14\u8bb2\u8005\u5907\u6ce8",
    "\u5907\u6ce8",
    "\u8bf4\u660e",
    "note",
    "notes",
    "speaker",
    "\u5e94\u5bf9",
    "response",
  ]);
  const counts = new Map();

  for (const item of (Array.isArray(lines) ? lines : [])) {
    const text = String(item || "").trim();
    if (!text) continue;

    const match = text.match(/^([^\s:\uFF1A\uFF0C\u3002,\uFF1B;\uFF08\uFF09()\u3010\u3011\[\]<>]{2,24})[\uFF1A:]/u);
    if (!match) continue;

    const candidate = String(match[1] || "").trim();
    const normalized = candidate.toLowerCase();
    if (!candidate || stopWords.has(normalized)) continue;
    if (/[0-9\uFF10-\uFF19]/u.test(candidate)) continue;

    counts.set(candidate, (counts.get(candidate) || 0) + 1);
  }

  let best = "";
  let bestCount = 0;
  for (const [label, count] of counts.entries()) {
    if (count > bestCount || (count === bestCount && best && label.length < best.length) || (count === bestCount && !best)) {
      best = label;
      bestCount = count;
    }
  }

  if (!best) return "";
  return best;
}

function buildRiskLabelFallback(primaryBlob, secondaryBlob, fullBlob, styleKey = "management") {
  const score = (blob, keywords) => keywords.reduce((total, keyword) => (blob.includes(keyword) ? total + 1 : total), 0);

  const attackWords = ["attack", "adversarial", "poison", "threat", "\u653b\u51fb", "\u6295\u6bd2", "\u6e17\u900f", "\u5165\u4fb5", "\u5a01\u80c1"];
  const defenseWords = ["defense", "protect", "mitigate", "\u9632\u62a4", "\u9632\u5fa1", "\u52a0\u56fa", "\u62e6\u622a", "\u68c0\u6d4b", "\u54cd\u5e94", "\u7f13\u89e3"];
  const problemWords = ["problem", "issue", "gap", "pain", "\u95ee\u9898", "\u75db\u70b9", "\u74f6\u9888", "\u77ed\u677f", "\u7f3a\u9677", "\u98ce\u9669"];
  const improveWords = ["improve", "improvement", "optimize", "plan", "\u6539\u8fdb", "\u4f18\u5316", "\u65b9\u6848", "\u6574\u6539", "\u63aa\u65bd", "\u5efa\u8bae"];
  const mechanismWords = ["method", "framework", "mechanism", "model", "\u65b9\u6cd5", "\u673a\u5236", "\u6846\u67b6", "\u5efa\u6a21", "\u5f62\u5f0f\u5316", "\u6d41\u7a0b"];
  const implementWords = ["implement", "execution", "deploy", "\u5b9e\u65bd", "\u6267\u884c", "\u843d\u5730", "\u63a8\u8fdb", "\u8def\u5f84", "\u52a8\u4f5c"];
  const resultWords = ["result", "impact", "evaluation", "finding", "\u5b9e\u9a8c", "\u7ed3\u679c", "\u8bc4\u4f30", "\u6548\u679c", "\u53d1\u73b0"];
  const insightWords = ["insight", "conclusion", "\u542f\u793a", "\u7ed3\u8bba", "\u610f\u4e49", "\u540e\u7eed", "\u5efa\u8bae"];

  const attackScore =
    score(primaryBlob, attackWords) * 3 +
    score(fullBlob, attackWords) * 2 +
    score(secondaryBlob, defenseWords) * 3 +
    score(fullBlob, defenseWords) * 2;

  const issueScore =
    score(primaryBlob, problemWords) * 3 +
    score(fullBlob, problemWords) * 2 +
    score(secondaryBlob, improveWords) * 3 +
    score(fullBlob, improveWords) * 2;

  const mechanismScore =
    score(primaryBlob, mechanismWords) * 3 +
    score(fullBlob, mechanismWords) * 2 +
    score(secondaryBlob, implementWords) * 3 +
    score(fullBlob, implementWords) * 2;

  const resultScore =
    score(primaryBlob, resultWords) * 3 +
    score(fullBlob, resultWords) * 2 +
    score(secondaryBlob, insightWords) * 3 +
    score(fullBlob, insightWords) * 2;

  const rankedProfiles = [
    {
      score: attackScore,
      labels: {
        left: "\u653b\u51fb\u9762\u8bc6\u522b",
        right: "\u9632\u62a4\u7b56\u7565",
        tags: "\u653b\u51fb\u8def\u5f84",
        actions: "\u9632\u62a4\u52a8\u4f5c",
      },
    },
    {
      score: issueScore,
      labels: {
        left: "\u5173\u952e\u95ee\u9898",
        right: "\u6539\u8fdb\u65b9\u6848",
        tags: "\u95ee\u9898\u5206\u7c7b",
        actions: "\u6539\u8fdb\u52a8\u4f5c",
      },
    },
    {
      score: mechanismScore,
      labels: {
        left: "\u5173\u952e\u673a\u5236",
        right: "\u5b9e\u73b0\u8def\u5f84",
        tags: "\u673a\u5236\u5206\u5c42",
        actions: "\u5b9e\u65bd\u52a8\u4f5c",
      },
    },
    {
      score: resultScore,
      labels: {
        left: "\u6838\u5fc3\u53d1\u73b0",
        right: "\u7ed3\u8bba\u4e0e\u542f\u793a",
        tags: "\u8bc1\u636e\u7ef4\u5ea6",
        actions: "\u540e\u7eed\u52a8\u4f5c",
      },
    },
  ].sort((a, b) => b.score - a.score);

  if (rankedProfiles[0].score > 0) {
    return rankedProfiles[0].labels;
  }

  if (styleKey === "technical") {
    return {
      left: "\u6280\u672f\u8981\u70b9",
      right: "\u5b9e\u65bd\u8def\u5f84",
      tags: "\u8981\u70b9\u5206\u7c7b",
      actions: "\u5b9e\u65bd\u52a8\u4f5c",
    };
  }

  return {
    left: "\u6838\u5fc3\u8bae\u9898",
    right: "\u63a8\u8fdb\u65b9\u6848",
    tags: "\u8bae\u9898\u5206\u7c7b",
    actions: "\u6267\u884c\u52a8\u4f5c",
  };
}

function pickRiskLabelSet(data, splitBullets = null, styleKey = "management") {
  const title = String((data && data.title) || "").toLowerCase();
  const primaryBullets = Array.isArray(splitBullets && splitBullets.primaryBullets)
    ? splitBullets.primaryBullets
    : (Array.isArray(data && data.bullets) ? data.bullets : []);
  const secondaryBullets = Array.isArray(splitBullets && splitBullets.secondaryBullets)
    ? splitBullets.secondaryBullets
    : [];

  const primaryBlob = primaryBullets.join(" ").toLowerCase();
  const secondaryBlob = secondaryBullets.join(" ").toLowerCase();
  const fullBlob = [title]
    .concat(primaryBullets)
    .concat(secondaryBullets)
    .concat(Array.isArray(data && data.evidence) ? data.evidence : [])
    .concat(Array.isArray(data && data.text_blocks) ? data.text_blocks : [])
    .join(" ")
    .toLowerCase();

  const fallback = buildRiskLabelFallback(primaryBlob, secondaryBlob, fullBlob, styleKey);
  const leftHint = extractRiskHeadingHint(primaryBullets);
  const rightHint = extractRiskHeadingHint(secondaryBullets);

  let left = leftHint || fallback.left;
  let right = rightHint || fallback.right;

  if (left && right && left === right) {
    if (!leftHint && fallback.left !== right) {
      left = fallback.left;
    } else if (!rightHint && fallback.right !== left) {
      right = fallback.right;
    }
  }

  return {
    left: left || fallback.left,
    right: right || fallback.right,
    tags: leftHint || fallback.tags,
    actions: rightHint || fallback.actions,
  };
}

function hasStrongRiskSignals(data) {
  const blob = [
    String((data && data.title) || ""),
    ...(Array.isArray(data && data.detail_points) ? data.detail_points : []),
    ...(Array.isArray(data && data.bullets) ? data.bullets : []),
    ...(Array.isArray(data && data.evidence) ? data.evidence : []),
  ].join(" ").toLowerCase();

  const hardRiskWords = [
    "risk", "threat", "vulnerability", "attack", "poison", "security", "mitigation",
    "\u98ce\u9669", "\u5a01\u80c1", "\u6f0f\u6d1e", "\u9690\u60a3", "\u653b\u51fb", "\u6295\u6bd2", "\u9632\u62a4", "\u5bf9\u7b56", "\u5e94\u5bf9", "\u7f13\u89e3",
  ];
  const softProblemWords = ["problem", "issue", "challenge", "\u95ee\u9898", "\u6311\u6218", "\u75db\u70b9", "\u7f3a\u9677"];
  const softActionWords = ["solution", "plan", "action", "improve", "\u6539\u8fdb", "\u4f18\u5316", "\u65b9\u6848", "\u63aa\u65bd", "\u5b9e\u65bd"];

  const hasAny = (words) => words.some((w) => blob.includes(w));
  if (hasAny(hardRiskWords)) return true;
  if (hasAny(softProblemWords) && hasAny(softActionWords)) return true;
  return false;
}

function hasMeaningfulRiskSplit(primaryBullets, secondaryBullets, labels) {
  const primary = Array.isArray(primaryBullets) ? primaryBullets : [];
  const secondary = Array.isArray(secondaryBullets) ? secondaryBullets : [];
  if (primary.length === 0 || secondary.length === 0) return false;

  const left = String(labels && labels.left || "").trim();
  const right = String(labels && labels.right || "").trim();
  const tags = String(labels && labels.tags || "").trim();
  const actions = String(labels && labels.actions || "").trim();
  if ((left && right && left === right) || (tags && actions && tags === actions)) return false;

  const primaryKeys = new Set(primary.map((x) => normalizeCompareKey(x)).filter(Boolean));
  const secondaryKeys = new Set(secondary.map((x) => normalizeCompareKey(x)).filter(Boolean));
  let overlap = 0;
  for (const key of secondaryKeys) {
    if (primaryKeys.has(key)) overlap += 1;
  }
  if (overlap >= Math.max(1, Math.floor(Math.min(primaryKeys.size, secondaryKeys.size) * 0.5))) return false;

  const generatedOnly = secondary.every((x) => String(x || "").trim().startsWith("\u5e94\u5bf9\uFF1A"));
  if (generatedOnly) return false;

  return true;
}

function splitRiskBullets(data) {
  const normalizeKey = (raw) => String(raw || "").toLowerCase().replace(/[\s\-_.,:;!?()\[\]{}<>/\\|]+/g, "").trim();
  const sanitizeLine = (raw) => String(raw || "").trim();
  const hasPrimaryCue = (line) => /(\u95ee\u9898|\u75db\u70b9|\u6311\u6218|\u98ce\u9669|\u5a01\u80c1|\u653b\u51fb|issue|risk|threat|attack)/i.test(line);
  const hasSecondaryCue = (line) => /(\u6539\u8fdb|\u4f18\u5316|\u65b9\u6848|\u63aa\u65bd|\u5b9e\u65bd|\u5e94\u5bf9|\u9632\u62a4|\u5bf9\u7b56|solution|plan|action|mitigation|defense)/i.test(line);
  const hasActionBiasCue = (line) => /(\u65b9\u6848|\u63aa\u65bd|\u5e94\u5bf9|\u9632\u62a4|\u89c4\u907f|\u964d\u4f4e|\u63d0\u5347|\u4f18\u5316|\u4fdd\u969c|solution|plan|action|mitigation|defense)/i.test(line);
  const preferSecondaryWhenMixed = (line) => hasSecondaryCue(line) && hasActionBiasCue(line);

  const sourceBullets = (Array.isArray(data && data.detail_points) && data.detail_points.length > 0
    ? data.detail_points
    : (Array.isArray(data && data.bullets) ? data.bullets : []))
    .map((x) => sanitizeLine(x))
    .filter((x) => x && !isMetaNoteLine(x));

  const evidenceBullets = (Array.isArray(data && data.evidence) ? data.evidence : [])
    .map((x) => sanitizeLine(x))
    .filter((x) => x && !isMetaNoteLine(x));

  const primarySeed = [];
  const secondarySeed = [];

  for (const item of sourceBullets) {
    if (hasSecondaryCue(item) && (!hasPrimaryCue(item) || preferSecondaryWhenMixed(item))) {
      secondarySeed.push(item);
    } else {
      primarySeed.push(item);
    }
  }

  for (const item of evidenceBullets) {
    if (hasSecondaryCue(item) && (!hasPrimaryCue(item) || preferSecondaryWhenMixed(item))) {
      secondarySeed.push(item);
    } else if (hasPrimaryCue(item) && !hasSecondaryCue(item)) {
      primarySeed.push(item);
    } else {
      secondarySeed.push(item);
    }
  }

  const uniquePrimary = [];
  const primarySeen = new Set();
  for (const item of primarySeed) {
    const key = normalizeKey(item);
    if (!key || primarySeen.has(key)) continue;
    primarySeen.add(key);
    uniquePrimary.push(item);
    if (uniquePrimary.length >= 5) break;
  }

  const contentFallback = sanitizeLine(getSummaryText(data));
  const primaryBullets = (uniquePrimary.length > 0 ? uniquePrimary : (contentFallback ? [contentFallback] : [])).slice(0, 5);
  const primaryKeys = new Set(primaryBullets.map((x) => normalizeKey(x)));

  const uniqueSecondary = [];
  const secondarySeen = new Set();
  for (const item of secondarySeed) {
    const key = normalizeKey(item);
    if (!key || primaryKeys.has(key) || secondarySeen.has(key)) continue;
    secondarySeen.add(key);
    uniqueSecondary.push(item);
    if (uniqueSecondary.length >= 5) break;
  }

  const secondaryBullets = uniqueSecondary.slice(0, 5);

  return {
    primaryBullets,
    secondaryBullets,
  };
}

function renderRiskSlide(pptx, data, theme, variant = 0) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };
  addTitleBar(slide, String(data.title || "Risk"), theme);

  const { primaryBullets, secondaryBullets } = splitRiskBullets(data);
  const labels = pickRiskLabelSet(data, { primaryBullets, secondaryBullets }, theme.styleKey || "management");

  if (!hasMeaningfulRiskSplit(primaryBullets, secondaryBullets, labels)) {
    const merged = dedupeTextItems(primaryBullets.concat(secondaryBullets), 6, 140);
    slide.addShape("roundRect", {
      x: 0.85,
      y: 1.5,
      w: 11.6,
      h: 5.75,
      radius: 0.08,
      fill: { color: theme.card },
      line: { color: theme.line, pt: 1.2 },
    });
    addBulletList(slide, merged, 1.2, 1.95, 10.9, 4.95, theme, { maxItems: 6, minLineHeight: 0.48 });
    addWarnings(slide, pptx);
    return;
  }

  if (variant === 1) {
    slide.addShape("roundRect", {
      x: 0.95,
      y: 1.5,
      w: 11.35,
      h: 2.35,
      radius: 0.08,
      fill: { color: theme.cardAlt },
      line: { color: theme.line, pt: 1.2 },
    });
    slide.addShape("roundRect", {
      x: 0.95,
      y: 4.0,
      w: 11.35,
      h: 3.2,
      radius: 0.08,
      fill: { color: theme.card },
      line: { color: theme.line, pt: 1.2 },
    });

    slide.addText(labels.left, {
      x: 1.25,
      y: 1.8,
      w: 5.4,
      h: 0.45,
      fontFace: "Microsoft YaHei",
      bold: true,
      color: theme.text,
      fontSize: 20,
      fit: "shrink",
    });
    addBulletList(slide, primaryBullets.slice(0, 2), 1.25, 2.25, 10.6, 1.35, theme, { maxTextLength: 0 });

    slide.addText(labels.right, {
      x: 1.25,
      y: 4.3,
      w: 5.4,
      h: 0.45,
      fontFace: "Microsoft YaHei",
      bold: true,
      color: theme.text,
      fontSize: 20,
      fit: "shrink",
    });
    addBulletList(slide, secondaryBullets, 1.25, 4.75, 10.6, 2.2, theme, { maxTextLength: 0 });

    addWarnings(slide, pptx);
    return;
  }

  if (variant === 2) {
    slide.addShape("roundRect", {
      x: 0.8,
      y: 1.5,
      w: 3.6,
      h: 5.7,
      radius: 0.08,
      fill: { color: theme.cardAlt },
      line: { color: theme.line, pt: 1.2 },
    });
    slide.addShape("roundRect", {
      x: 4.6,
      y: 1.5,
      w: 7.9,
      h: 5.7,
      radius: 0.08,
      fill: { color: theme.card },
      line: { color: theme.line, pt: 1.2 },
    });

    slide.addText(labels.tags, {
      x: 1.05,
      y: 1.8,
      w: 3.0,
      h: 0.4,
      fontFace: "Microsoft YaHei",
      bold: true,
      color: theme.text,
      fontSize: 16,
      fit: "shrink",
    });

    const tags = primaryBullets.slice(0, 5);
    tags.forEach((item, idx) => {
      const y = 2.35 + idx * 0.95;
      slide.addShape("roundRect", {
        x: 1.0,
        y,
        w: 3.1,
        h: 0.62,
        radius: 0.08,
        fill: { color: theme.card },
        line: { color: theme.line, pt: 0.8 },
      });
      slide.addText(item, {
        x: 1.2,
        y: y + 0.08,
        w: 2.7,
        h: 0.46,
        fontFace: "Microsoft YaHei",
        color: theme.text,
        fontSize: 12,
        align: "center",
        valign: "mid",
        breakLine: true,
        fit: "shrink",
        margin: 0,
      });
    });

    slide.addText(labels.actions, {
      x: 4.95,
      y: 1.8,
      w: 6.9,
      h: 0.4,
      fontFace: "Microsoft YaHei",
      bold: true,
      color: theme.text,
      fontSize: 18,
      fit: "shrink",
    });
    addBulletList(slide, secondaryBullets, 4.95, 2.3, 7.2, 4.6, theme, { maxTextLength: 0 });

    addWarnings(slide, pptx);
    return;
  }

  slide.addShape("roundRect", {
    x: 0.8,
    y: 1.5,
    w: 5.7,
    h: 5.7,
    radius: 0.08,
    fill: { color: theme.card },
    line: { color: theme.line, pt: 1.2 },
  });
  slide.addShape("roundRect", {
    x: 6.8,
    y: 1.5,
    w: 5.7,
    h: 5.7,
    radius: 0.08,
    fill: { color: theme.cardAlt },
    line: { color: theme.line, pt: 1.2 },
  });

  slide.addText(labels.left, {
    x: 1.1,
    y: 1.75,
    w: 4.8,
    h: 0.45,
    fontFace: "Microsoft YaHei",
    bold: true,
    color: theme.text,
    fontSize: 20,
    fit: "shrink",
  });

  slide.addText(labels.right, {
    x: 7.1,
    y: 1.75,
    w: 4.8,
    h: 0.45,
    fontFace: "Microsoft YaHei",
    bold: true,
    color: theme.text,
    fontSize: 20,
    fit: "shrink",
  });

  addBulletList(slide, primaryBullets, 1.1, 2.25, 5.0, 4.6, theme, { maxTextLength: 0 });
  addBulletList(slide, secondaryBullets, 7.1, 2.25, 5.0, 4.6, theme, { maxTextLength: 0 });

  if (data.generated_image_path && fs.existsSync(data.generated_image_path)) {
    addImageIfExists(slide, data.generated_image_path, 4.95, 5.15, 3.45, 1.85);
  }

  addWarnings(slide, pptx);
}

function renderTimelineSlide(pptx, data, theme, variant = 0) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };
  addTitleBar(slide, String(data.title || "Timeline"), theme);

  const points = (Array.isArray(data.bullets) ? data.bullets : []).map((x) => truncate(x, 54)).filter(Boolean);
  const items = points.length > 0 ? points.slice(0, 5) : ["Stage detail TBD"];

  if (variant === 1) {
    slide.addShape("roundRect", {
      x: 0.9,
      y: 1.5,
      w: 11.5,
      h: 5.8,
      radius: 0.08,
      fill: { color: theme.card },
      line: { color: theme.line, pt: 1.2 },
    });

    items.forEach((item, idx) => {
      const y = 1.95 + idx * 1.05;
      slide.addShape("roundRect", {
        x: 1.2,
        y,
        w: 1.4,
        h: 0.7,
        radius: 0.06,
        fill: { color: theme.cardAlt },
        line: { color: theme.line, pt: 1 },
      });
      slide.addText(`T${idx + 1}`, {
        x: 1.62,
        y: y + 0.21,
        w: 0.56,
        h: 0.28,
        fontFace: "Microsoft YaHei",
        bold: true,
        color: theme.accent,
        fontSize: 12,
        align: "center",
      });
      slide.addShape("line", {
        x: 2.75,
        y: y + 0.35,
        w: 0.6,
        h: 0,
        line: { color: theme.line, pt: 1.5 },
      });
      slide.addText(item, {
        x: 3.45,
        y: y + 0.1,
        w: 8.6,
        h: 0.45,
        fontFace: "Microsoft YaHei",
        color: theme.text,
        fontSize: 14,
      });
    });

    addWarnings(slide, pptx);
    return;
  }

  if (variant === 2) {
    slide.addShape("line", {
      x: 1.15,
      y: 3.75,
      w: 10.9,
      h: 0,
      line: { color: theme.line, pt: 2 },
    });

    const n = Math.max(1, items.length);
    for (let idx = 0; idx < n; idx += 1) {
      const x = 1.2 + idx * (10.6 / Math.max(1, n - 1));
      const cardY = idx % 2 === 0 ? 1.95 : 4.2;

      slide.addShape("ellipse", {
        x,
        y: 3.52,
        w: 0.42,
        h: 0.42,
        fill: { color: theme.accent },
        line: { color: theme.accent, pt: 0 },
      });

      slide.addShape("roundRect", {
        x: x - 0.8,
        y: cardY,
        w: 2.0,
        h: 1.18,
        radius: 0.06,
        fill: { color: idx % 2 === 0 ? theme.card : theme.cardAlt },
        line: { color: theme.line, pt: 1 },
      });

      slide.addText(`Stage ${idx + 1}`, {
        x: x - 0.7,
        y: cardY + 0.1,
        w: 1.8,
        h: 0.24,
        fontFace: "Microsoft YaHei",
        bold: true,
        color: theme.muted,
        fontSize: 10,
        align: "center",
      });

      slide.addText(items[idx], {
        x: x - 0.7,
        y: cardY + 0.35,
        w: 1.8,
        h: 0.7,
        fontFace: "Microsoft YaHei",
        color: theme.text,
        fontSize: 11,
        valign: "top",
      });
    }

    addWarnings(slide, pptx);
    return;
  }

  slide.addShape("line", {
    x: 1.1,
    y: 3.65,
    w: 11.0,
    h: 0,
    line: { color: theme.line, pt: 2 },
  });

  const n = items.length;
  for (let idx = 0; idx < n; idx += 1) {
    const x = 1.2 + idx * (10.6 / Math.max(1, n - 1));

    slide.addShape("ellipse", {
      x,
      y: 3.42,
      w: 0.42,
      h: 0.42,
      fill: { color: theme.accent },
      line: { color: theme.accent, pt: 0 },
    });

    slide.addShape("roundRect", {
      x: x - 0.5,
      y: 1.95,
      w: 1.5,
      h: 1.2,
      radius: 0.06,
      fill: { color: idx % 2 === 1 ? theme.cardAlt : theme.card },
      line: { color: theme.line, pt: 1.2 },
    });

    slide.addText(`Stage ${idx + 1}`, {
      x: x - 0.45,
      y: 2.02,
      w: 1.4,
      h: 0.2,
      fontFace: "Microsoft YaHei",
      bold: true,
      color: theme.muted,
      fontSize: 11,
    });

    slide.addText(items[idx], {
      x: x - 0.45,
      y: 2.28,
      w: 1.4,
      h: 0.8,
      fontFace: "Microsoft YaHei",
      color: theme.text,
      fontSize: 12,
      valign: "top",
    });
  }

  if (data.generated_image_path && fs.existsSync(data.generated_image_path)) {
    addImageIfExists(slide, data.generated_image_path, 8.55, 4.35, 3.65, 2.55);
  }

  addWarnings(slide, pptx);
}
function compactChartLabel(raw, idx) {
  let txt = String(raw || "").trim();
  if (!txt) return `\u6307\u6807${idx + 1}`;
  txt = txt.replace(/^(\u89c6\u89c9\u5143\u7d20|\u89c6\u89c9\u7126\u70b9|\u6392\u7248\u5e03\u5c40|\u6f14\u8bb2\u8005\u5907\u6ce8|\u5907\u6ce8|\u8bf4\u660e)[\uff1a:]/u, "");
  txt = txt.replace(/[\uff0c,\u3002\uff1b;].*$/, "");
  txt = txt.replace(/[\s\u3000]+/g, " ").trim();
  if (!txt) return `\u6307\u6807${idx + 1}`;
  if (txt.length > 18) txt = txt.slice(0, 18).trim();
  return txt || `\u6307\u6807${idx + 1}`;
}

function inferChartType(data, unit, labels, values) {
  const blob = [String((data && data.title) || "")]
    .concat(Array.isArray(data && data.bullets) ? data.bullets : [])
    .join(" ")
    .toLowerCase();

  const trendWords = ["\u589e\u957f", "\u589e\u901f", "\u8d8b\u52bf", "\u53d8\u5316", "\u6ce2\u52a8", "\u63d0\u5347", "\u4e0b\u964d", "\u540c\u6bd4", "\u73af\u6bd4", "trend", "growth", "rate"];
  const compareWords = ["\u5bf9\u6bd4", "\u5206\u5e03", "\u7ed3\u6784", "\u6392\u540d", "\u5360\u6bd4", "compare", "distribution", "ranking"];

  const hasTrend = trendWords.some((w) => blob.includes(w));
  const hasCompare = compareWords.some((w) => blob.includes(w));

  if (hasTrend) return "line";
  if (hasCompare) return "bar";
  if (unit === "%" && values.length >= 3) return "line";
  if (labels.length >= 4) return "line";
  return "bar";
}

function deriveAxisTopicSeed(data) {
  const rawTitle = String((data && data.title) || "").trim();
  let seed = rawTitle
    .replace(/\d{4}\s*\u5e74?/g, "")
    .replace(/AI|AIGC|LLM|RAG|GPT/gi, "")
    .replace(/(\u73b0\u72b6|\u8d8b\u52bf|\u5206\u6790|\u62a5\u544a|\u4e13\u9898|\u6c47\u62a5|\u5173\u952e|\u6838\u5fc3|\u6280\u672f|\u9769\u65b0|\u5347\u7ea7|\u843d\u5730|\u5b9e\u8df5|\u65b9\u6848|\u95ee\u9898)/gu, "")
    .replace(/[\s\-_/|:\uFF1A]+/g, "")
    .trim();

  if (!seed) {
    const bullets = Array.isArray(data && data.bullets) ? data.bullets : [];
    const blob = bullets.join(" ");
    if (/(\u91cf\u5316|\u63a8\u7406|\u90e8\u7f72)/u.test(blob)) {
      seed = "\u63a8\u7406\u90e8\u7f72";
    } else if (/(\u8bad\u7ec3|\u5fae\u8c03|lora|qlora|zero)/i.test(blob)) {
      seed = "\u8bad\u7ec3\u4f18\u5316";
    } else if (/(\u591a\u6a21\u6001|\u89c6\u9891|\u56fe\u50cf|\u8bed\u97f3)/u.test(blob)) {
      seed = "\u591a\u6a21\u6001";
    } else if (/(\u56fd\u4ea7|\u7b97\u529b|\u82af\u7247|\u66ff\u4ee3)/u.test(blob)) {
      seed = "\u56fd\u4ea7\u5316";
    }
  }

  if (!seed) seed = "\u4e1a\u52a1";
  if (seed.length > 8) seed = seed.slice(0, 8);
  return seed;
}

function inferChartAxisTitles(data, unit, chartType, labels = []) {
  const title = String((data && data.title) || "");
  const bullets = Array.isArray(data && data.bullets) ? data.bullets : [];
  const blob = [title].concat(bullets).join(" ").toLowerCase();
  const labelBlob = (Array.isArray(labels) ? labels : []).join(" ").toLowerCase();
  const axisContext = `${blob} ${labelBlob}`;
  const seed = deriveAxisTopicSeed(data);

  const isTimelineLike = /(\u5e74|\u6708|\u5b63\u5ea6|q\d|\u9636\u6bb5|\u91cc\u7a0b\u7891|timeline|stage)/i.test(axisContext);
  const hasSceneLike = /(\u573a\u666f|\u884c\u4e1a|\u5ba2\u6237|scenario|use case)/i.test(axisContext);
  const hasTechLike = /(\u91cf\u5316|\u63a8\u7406|\u90e8\u7f72|lora|qlora|zero|\u67b6\u6784|\u8282\u70b9|\u7b56\u7565)/i.test(axisContext);

  let catAxisTitle = `${seed}\u7ef4\u5ea6`;
  if (isTimelineLike) {
    catAxisTitle = "\u65f6\u95f4/\u9636\u6bb5";
  } else if (hasSceneLike) {
    catAxisTitle = "\u5e94\u7528\u573a\u666f";
  } else if (hasTechLike) {
    catAxisTitle = "\u6280\u672f/\u7b56\u7565\u9879";
  }

  let valAxisTitle = `${seed}\u6548\u679c\u503c`;
  if (unit) {
    valAxisTitle = `${seed}\u6307\u6807\uff08${unit}\uff09`;
  } else if (chartType === "line" && /(\u589e\u957f|\u8d8b\u52bf|trend|growth|rate|\u63d0\u5347|\u4e0b\u964d)/i.test(axisContext)) {
    valAxisTitle = "\u53d8\u5316\u5e45\u5ea6";
  } else if (/(\u901f\u5ea6|\u5ef6\u8fdf|\u541e\u5410|latency|throughput|performance|\u6027\u80fd|\u6548\u7387)/i.test(axisContext)) {
    valAxisTitle = "\u6027\u80fd\u6536\u76ca\u503c";
  } else if (/(\u6210\u672c|cost|\u8d44\u6e90|\u663e\u5b58|\u7b97\u529b)/i.test(axisContext)) {
    valAxisTitle = "\u8d44\u6e90\u53d8\u5316\u503c";
  }

  return { catAxisTitle, valAxisTitle };
}

function isLowQualityChartData(data, labels, values, unit) {
  if (labels.length < 2 || values.length < 2 || labels.length !== values.length) return true;
  if (labels.some((x) => isMetaNoteLine(x))) return true;

  const absVals = values.map((v) => Math.abs(Number(v))).filter((v) => Number.isFinite(v) && v > 0);
  if (!unit && labels.length <= 2 && absVals.length >= 2) {
    const maxVal = Math.max(...absVals);
    const minVal = Math.min(...absVals);
    if (maxVal / Math.max(1, minVal) >= 10) return true;
  }

  const allGeneric = labels.every((x) => /^\u6307\u6807\d+$/u.test(x));
  if (allGeneric && labels.length <= 2) return true;

  const title = String((data && data.title) || "");
  if (/\u591a\u6a21\u6001/.test(title) && !unit && values.length <= 2) return true;

  return false;
}

function chartPayload(data) {
  const chart = data && typeof data.chart_data === "object" ? data.chart_data : null;
  if (!chart) return null;

  const rawLabels = Array.isArray(chart.labels) ? chart.labels : [];
  const rawValues = Array.isArray(chart.values) ? chart.values : [];
  const unit = String(chart.unit || "").trim();

  const pairCount = Math.min(rawLabels.length, rawValues.length, 6);
  const pairs = [];
  for (let i = 0; i < pairCount; i += 1) {
    const value = Number(rawValues[i]);
    if (!Number.isFinite(value)) continue;
    const label = compactChartLabel(rawLabels[i], pairs.length);
    if (!label) continue;
    pairs.push([label, value]);
  }

  const labels = [];
  const values = [];
  const seen = new Set();
  for (const [label, value] of pairs) {
    const key = normalizeCompareKey(label);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    labels.push(label);
    values.push(value);
  }

  if (isLowQualityChartData(data, labels, values, unit)) {
    return null;
  }

  const chartType = inferChartType(data, unit, labels, values);
  const { catAxisTitle, valAxisTitle } = inferChartAxisTitles(data, unit, chartType, labels);

  return { labels, values, unit, chartType, catAxisTitle, valAxisTitle };
}

function addDataChart(slide, chartType, labels, values, theme, opts) {
  const base = {
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    showLegend: false,
    catAxisTitle: opts.catAxisTitle,
    valAxisTitle: opts.valAxisTitle,
    chartColors: [theme.accent],
  };

  if (chartType === "line") {
    slide.addChart("line", [{ name: "Metrics", labels, values }], {
      ...base,
      lineSize: 2,
      markerSize: 5,
      catAxisLabelPos: "nextTo",
    });
    return;
  }

  slide.addChart("bar", [{ name: "Metrics", labels, values }], {
    ...base,
    barDir: opts.barDir || "col",
    catAxisLabelPos: "nextTo",
    gapWidthPct: opts.gapWidthPct || 30,
  });
}

function addChartAxisCaptions(slide, theme, opts) {
  const cat = String((opts && opts.catAxisTitle) || "").trim();
  const val = String((opts && opts.valAxisTitle) || "").trim();
  const x = Number((opts && opts.x) || 0);
  const y = Number((opts && opts.y) || 0);
  const w = Number((opts && opts.w) || 0);
  const h = Number((opts && opts.h) || 0);

  if (cat) {
    slide.addText(cat, {
      x: x + 0.45,
      y: y + h + 0.02,
      w: Math.max(0.6, w - 0.9),
      h: 0.24,
      fontFace: "Microsoft YaHei",
      color: theme.muted,
      fontSize: 10,
      align: "center",
      valign: "mid",
    });
  }

  if (val) {
    slide.addText(val, {
      x: x + 0.05,
      y: y + 0.25,
      w: 0.28,
      h: Math.max(1.2, h - 0.5),
      fontFace: "Microsoft YaHei",
      color: theme.muted,
      fontSize: 9,
      align: "center",
      valign: "mid",
      vert: "vert270",
    });
  }
}

function renderDataSlide(pptx, data, theme, chart, variant = 0) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };
  addTitleBar(slide, String(data.title || "Data"), theme);

  const hasImage = Boolean(data.generated_image_path && fs.existsSync(data.generated_image_path));
  const labels = hasImage ? chart.labels.slice(0, 4) : chart.labels;
  const values = hasImage ? chart.values.slice(0, 4) : chart.values;
  const chartType = String(chart.chartType || "bar");
  const catAxisTitle = String(chart.catAxisTitle || "\u6307\u6807\u9879");
  const valAxisTitle = String(chart.valAxisTitle || (chart.unit ? `\u6307\u6807\u503c\uff08${chart.unit}\uff09` : "\u6307\u6807\u503c"));

  if (variant === 1) {
    slide.addShape("roundRect", {
      x: 0.85,
      y: 1.5,
      w: 11.6,
      h: 3.45,
      radius: 0.08,
      fill: { color: theme.cardAlt },
      line: { color: theme.line, pt: 1.2 },
    });
    slide.addShape("roundRect", {
      x: 0.85,
      y: 5.1,
      w: 11.6,
      h: 2.15,
      radius: 0.08,
      fill: { color: theme.card },
      line: { color: theme.line, pt: 1.2 },
    });

    const chartOpts = {
      x: 1.2,
      y: 1.9,
      w: 10.9,
      h: 2.75,
      catAxisTitle,
      valAxisTitle,
      barDir: "col",
      gapWidthPct: 28,
    };
    addDataChart(slide, chartType, labels, values, theme, chartOpts);
    addChartAxisCaptions(slide, theme, chartOpts);

    addBulletList(slide, data.bullets, 1.2, 5.35, 10.6, 1.65, theme, { maxItems: 3, minLineHeight: 0.5 });

    addWarnings(slide, pptx);
    return;
  }

  if (variant === 2) {
    slide.addShape("roundRect", {
      x: 0.8,
      y: 1.5,
      w: 6.3,
      h: 5.7,
      radius: 0.08,
      fill: { color: theme.cardAlt },
      line: { color: theme.line, pt: 1.2 },
    });
    slide.addShape("roundRect", {
      x: 7.35,
      y: 1.5,
      w: 5.0,
      h: 5.7,
      radius: 0.08,
      fill: { color: theme.card },
      line: { color: theme.line, pt: 1.2 },
    });

    const chartOpts = {
      x: 1.15,
      y: 2.0,
      w: 5.6,
      h: 3.9,
      catAxisTitle,
      valAxisTitle,
      barDir: chartType === "line" ? "col" : "bar",
      gapWidthPct: 30,
    };
    addDataChart(slide, chartType, labels, values, theme, chartOpts);
    addChartAxisCaptions(slide, theme, chartOpts);

    addBulletList(slide, dedupeTextItems(data.bullets, 4, 140), 7.75, 2.0, 4.2, 4.8, theme, { maxItems: 4 });

    addWarnings(slide, pptx);
    return;
  }

  slide.addShape("roundRect", {
    x: 0.8,
    y: 1.5,
    w: 5.3,
    h: 5.7,
    radius: 0.08,
    fill: { color: theme.card },
    line: { color: theme.line, pt: 1.2 },
  });
  slide.addShape("roundRect", {
    x: 6.35,
    y: 1.5,
    w: 6.0,
    h: 5.7,
    radius: 0.08,
    fill: { color: theme.cardAlt },
    line: { color: theme.line, pt: 1.2 },
  });

  addBulletList(slide, dedupeTextItems(data.bullets, 5, 140), 1.1, 1.95, 4.7, 4.9, theme, { maxItems: 5 });

  const chartOpts = {
    x: 6.75,
    y: 2.0,
    w: hasImage ? 4.65 : 5.2,
    h: hasImage ? 2.15 : 3.75,
    catAxisTitle,
    valAxisTitle,
    barDir: "col",
    gapWidthPct: 28,
  };
  addDataChart(slide, chartType, labels, values, theme, chartOpts);
  addChartAxisCaptions(slide, theme, chartOpts);

  if (hasImage) {
    addImageIfExists(slide, data.generated_image_path, 8.25, 4.35, 3.9, 2.55);
  }

  addWarnings(slide, pptx);
}

function renderConclusion(pptx, bodySlides, theme, variant = 0) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };
  addTitleBar(slide, "Summary", theme);

  const keyPoints = [];
  bodySlides.slice(0, 5).forEach((item) => {
    const title = truncate(item.title || "Untitled", 42);
    const firstBullet = Array.isArray(item.bullets) && item.bullets.length > 0 ? truncate(item.bullets[0], 54) : "Key takeaway";
    keyPoints.push(`${title}: ${firstBullet}`);
  });

  if (variant === 1) {
    slide.addShape("roundRect", {
      x: 0.85,
      y: 1.5,
      w: 11.6,
      h: 5.75,
      radius: 0.08,
      fill: { color: theme.cardAlt },
      line: { color: theme.line, pt: 1.2 },
    });

    keyPoints.forEach((item, idx) => {
      const col = idx % 2;
      const row = Math.floor(idx / 2);
      const x = 1.2 + col * 5.6;
      const y = 1.95 + row * 1.22;
      slide.addShape("roundRect", {
        x,
        y,
        w: 5.1,
        h: 1.0,
        radius: 0.06,
        fill: { color: theme.card },
        line: { color: theme.line, pt: 1 },
      });
      slide.addText(truncate(item, 76), {
        x: x + 0.28,
        y: y + 0.2,
        w: 4.5,
        h: 0.62,
        fontFace: "Microsoft YaHei",
        color: theme.text,
        fontSize: 13,
      });
    });

    addWarnings(slide, pptx);
    return;
  }

  if (variant === 2) {
    slide.addShape("roundRect", {
      x: 0.95,
      y: 1.55,
      w: 11.35,
      h: 5.65,
      radius: 0.08,
      fill: { color: theme.card },
      line: { color: theme.line, pt: 1.2 },
    });
    slide.addShape("line", {
      x: 1.2,
      y: 2.1,
      w: 10.9,
      h: 0,
      line: { color: theme.accent, pt: 1.8 },
    });
    addBulletList(slide, keyPoints, 1.25, 2.35, 10.7, 4.55, theme);
    addWarnings(slide, pptx);
    return;
  }

  slide.addShape("roundRect", {
    x: 0.9,
    y: 1.5,
    w: 11.5,
    h: 5.75,
    radius: 0.08,
    fill: { color: theme.card },
    line: { color: theme.line, pt: 1.2 },
  });

  addBulletList(slide, keyPoints, 1.25, 2.0, 10.8, 4.9, theme);
  addWarnings(slide, pptx);
}
function hashInt(text) {
  const digest = crypto.createHash("md5").update(String(text || "")).digest();
  return (digest[0] << 8) + digest[1];
}

function pickLayoutVariant(topic, slideType, index, styleKey = "management") {
  const layoutCount = {
    cover: 3,
    toc: 3,
    summary: 12,
    risk: 3,
    timeline: 3,
    status: 3,
    data: 3,
    conclusion: 3,
  };
  const count = layoutCount[String(slideType || "summary").toLowerCase()] || 3;
  return hashInt(`${topic}|${styleKey}|${slideType}|${index}`) % count;
}

async function exportFromScratch(payload, outPath) {
  const slides = Array.isArray(payload.slides) ? payload.slides : [];
  const body = contentSlides(slides);
  const topic = normalizeTopicTitle(String(payload.topic || (body[0] && body[0].title) || "Report"));
  const subtitle = String(payload.subtitle || payload.coverSubtitle || "").trim();
  const outline = Array.isArray(payload.outline) ? payload.outline : body.map((s) => String(s.title || ""));
  const rawTocItems = Array.isArray(payload.tocItems) ? payload.tocItems : body.map((s) => String(s.title || ""));
  const tocItems = filterRelevantTocItems(rawTocItems, topic);
  const style = String(payload.style || "management");
  const themeSeed = String(payload.themeSeed || `${payload.templateId || "default"}|${payload.topic || ""}|${style}`);
  const theme = deriveTheme(themeSeed, style);
  const styleKey = String(theme.styleKey || normalizeStyleKey(style));

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "AIppt";
  pptx.subject = "Generated by pptx-generator workflow";
  pptx.company = "AIppt";
  pptx.title = topic;
  pptx.theme = {
    lang: "zh-CN",
    headFontFace: "Microsoft YaHei",
    bodyFontFace: "Microsoft YaHei",
  };

  const coverVariant = pickLayoutVariant(topic, "cover", 0, styleKey);
  const tocVariant = pickLayoutVariant(topic, "toc", 1, styleKey);
  renderCover(pptx, topic, theme, subtitle, coverVariant);
  renderToc(pptx, topic, tocItems, theme, tocVariant);

  body.forEach((slideData, idx) => {
    const slideType = String(slideData.slide_type || "summary").toLowerCase();
    const summaryVariant = pickSummaryVariant(topic, slideData, idx + 2, styleKey);
    const variant = slideType === "summary" ? summaryVariant : pickLayoutVariant(topic, slideType, idx + 2, styleKey);

    if (slideType === "risk") {
      if (!hasStrongRiskSignals(slideData)) {
        renderSummarySlide(pptx, slideData, theme, summaryVariant);
      } else {
        renderRiskSlide(pptx, slideData, theme, variant);
      }
      return;
    }
    if (slideType === "timeline" || slideType === "status") {
      renderTimelineSlide(pptx, slideData, theme, variant);
      return;
    }
    if (slideType === "data") {
      const chart = chartPayload(slideData);
      if (chart) {
        renderDataSlide(pptx, slideData, theme, chart, variant);
      } else {
        renderSummarySlide(pptx, slideData, theme, summaryVariant);
      }
      return;
    }
    renderSummarySlide(pptx, slideData, theme, summaryVariant);
  });

  const conclusionVariant = pickLayoutVariant(topic, "conclusion", body.length + 2, styleKey);
  renderConclusion(pptx, body, theme, conclusionVariant);

  await pptx.writeFile({ fileName: outPath });
}
function escapeXml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function normalizeReplacementText(text) {
  return String(text || "")
    .replace(/\r\n/g, "\n")
    .replace(/^\s*#{1,6}\s+/gm, "")
    .replace(/<\/?[A-Za-z_][A-Za-z0-9._:-]*(?:\s[^>\n]*)?>/g, "")
    .replace(/&lt;\/?[A-Za-z_][^&]{0,120}&gt;/gi, "")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function replaceTextRunsInXml(xml, replacements) {
  if (!Array.isArray(replacements) || replacements.length === 0) {
    return xml;
  }

  const cleanValues = replacements.map((item) => normalizeReplacementText(item));
  let cursor = 0;
  return xml.replace(/<a:t(?![A-Za-z0-9_:-])([^>]*)>([\s\S]*?)<\/a:t>/g, (match, attrs) => {
    if (cursor >= cleanValues.length) {
      return match;
    }
    const value = escapeXml(cleanValues[cursor]);
    cursor += 1;
    return `<a:t${attrs}>${value}</a:t>`;
  });
}

function buildTemplateReplacement(index, topic, subtitle, outline, bodySlides) {
  if (index === 0) {
    return [topic, subtitle];
  }
  if (index === 1) {
    const toc = ["\u76ee\u5f55"];
    const lines = filterRelevantTocItems(outline, topic);
    lines.forEach((item, i) => toc.push(`${i + 1}. ${item}`));
    return toc;
  }

  const bodyIndex = index - 2;
  if (bodyIndex < 0 || bodyIndex >= bodySlides.length) {
    return [];
  }

  const payload = bodySlides[bodyIndex] || {};
  const replacements = [String(payload.title || "")];
  (Array.isArray(payload.bullets) ? payload.bullets : []).slice(0, 8).forEach((line) => replacements.push(String(line || "")));
  return replacements;
}

async function updatePresentationSlides(zip, keepCount) {
  const presentationPath = "ppt/presentation.xml";
  const relsPath = "ppt/_rels/presentation.xml.rels";

  const presentationFile = zip.file(presentationPath);
  if (!presentationFile) {
    return;
  }

  const presentationXml = await presentationFile.async("string");
  const listMatch = presentationXml.match(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/);
  if (!listMatch) {
    return;
  }

  const slideTags = [...listMatch[0].matchAll(/<p:sldId\b[^>]*\/>/g)].map((m) => m[0]);
  if (slideTags.length <= keepCount) {
    return;
  }

  const keptTags = slideTags.slice(0, keepCount);
  const keptRelIds = new Set(
    keptTags
      .map((tag) => {
        const m = tag.match(/\br:id="([^"]+)"/);
        return m ? m[1] : "";
      })
      .filter(Boolean),
  );

  const nextList = `<p:sldIdLst>${keptTags.join("")}</p:sldIdLst>`;
  const nextPresentationXml = presentationXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, nextList);
  zip.file(presentationPath, nextPresentationXml);

  const relsFile = zip.file(relsPath);
  if (!relsFile) {
    return;
  }

  const relsXml = await relsFile.async("string");
  const nextRelsXml = relsXml.replace(/<Relationship\b[^>]*\/>/g, (tag) => {
    const typeMatch = tag.match(/\bType="([^"]+)"/);
    const idMatch = tag.match(/\bId="([^"]+)"/);
    if (!typeMatch || !idMatch) {
      return tag;
    }
    const isSlideRel = /\/relationships\/slide$/i.test(typeMatch[1]);
    if (!isSlideRel) {
      return tag;
    }
    return keptRelIds.has(idMatch[1]) ? tag : "";
  });
  zip.file(relsPath, nextRelsXml);
}

function sortSlideXmlPaths(paths) {
  return [...paths].sort((a, b) => {
    const ai = Number((a.match(/slide(\d+)\.xml$/) || ["", "0"])[1]);
    const bi = Number((b.match(/slide(\d+)\.xml$/) || ["", "0"])[1]);
    return ai - bi;
  });
}

async function exportByXmlTemplate(payload, outPath) {
  const templatePath = String(payload.templatePptxPath || "");
  if (!templatePath || !fs.existsSync(templatePath)) {
    throw new Error("template path missing or does not exist");
  }

  const slides = Array.isArray(payload.slides) ? payload.slides : [];
  const body = contentSlides(slides);
  const topic = String(payload.topic || (body[0] && body[0].title) || "Report");
  const subtitle = String(payload.subtitle || payload.coverSubtitle || "").trim();
  const outline = Array.isArray(payload.outline) ? payload.outline : body.map((s) => String(s.title || ""));

  const raw = fs.readFileSync(templatePath);
  const zip = await JSZip.loadAsync(raw);
  const slideXmlPaths = sortSlideXmlPaths(Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name)));

  const desiredSlideCount = Math.max(1, 2 + body.length);
  const effectiveSlideCount = Math.min(desiredSlideCount, slideXmlPaths.length);

  for (let i = 0; i < effectiveSlideCount; i += 1) {
    const replacements = buildTemplateReplacement(i, topic, subtitle, outline, body);
    if (replacements.length === 0) {
      continue;
    }
    const xml = await zip.file(slideXmlPaths[i]).async("string");
    const nextXml = replaceTextRunsInXml(xml, replacements);
    zip.file(slideXmlPaths[i], nextXml);
  }

  await updatePresentationSlides(zip, effectiveSlideCount);

  const outBuffer = await zip.generateAsync({ type: "nodebuffer" });
  fs.writeFileSync(outPath, outBuffer);
}

async function main() {
  const args = parseArgs(process.argv);
  const payload = mustReadJson(args.input);

  fs.mkdirSync(path.dirname(args.output), { recursive: true });

  const hasTemplate = Boolean(payload.templatePptxPath && fs.existsSync(String(payload.templatePptxPath)));
  if (hasTemplate) {
    await exportByXmlTemplate(payload, args.output);
    return;
  }

  await exportFromScratch(payload, args.output);
}

main().catch((err) => {
  console.error(`[pptx_generator] ${String(err && err.stack ? err.stack : err)}`);
  process.exit(1);
});










