import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, ExternalHyperlink, ImageRun,
} from "docx";
import { readFileSync, writeFileSync, existsSync } from "fs";
import path from "path";
import { imageSize } from "image-size";

// ─── CLI parsing ──────────────────────────────────────────────────
const args = process.argv.slice(2);
function pickArg(flag) {
  const idx = args.indexOf(flag);
  if (idx !== -1 && idx + 1 < args.length) return args.splice(idx, 2)[1];
  if (idx !== -1) args.splice(idx, 1);
  return null;
}

const configPath = pickArg("--config");
const inputFile = (() => {
  for (const a of args) {
    if (!a.startsWith("--") && !a.endsWith(".js")) return path.resolve(a);
  }
  return null;
})();
const outputFile = (() => {
  let found = false;
  for (const a of args) {
    if (!a.startsWith("--")) {
      if (found) return path.resolve(a);
      found = true;
    }
  }
  return null;
})() || (inputFile ? inputFile.replace(/\.md$/i, ".docx") : null);

if (!inputFile) {
  console.error("用法: node md-to-word.js <输入.md> [输出.docx] [--config 配置.json]");
  console.error("示例: node md-to-word.js 文章.md");
  console.error("      node md-to-word.js 文章.md 输出.docx --config my-config.json");
  process.exit(1);
}

// ─── Config loader ────────────────────────────────────────────────
const DEFAULTS = {
  fonts: {
    body: "Microsoft YaHei", bodySize: 24, code: "Consolas",
    codeSize: 20, heading: "Microsoft YaHei", formula: "Consolas",
  },
  page: { marginTop: 1440, marginBottom: 1440, marginLeft: 1440, marginRight: 1440 },
  colors: {
    link: "0563C1", code: "333333", quote: "666666",
    heading: "1F1F1F", body: "333333", formula: "444444",
    imagePlaceholder: "999999", thematicBreak: "999999", tableHeaderBg: "F2F2F2",
  },
  spacing: {
    paragraphAfter: 120, headingBefore: 360, headingBeforeNested: 240, headingAfter: 160,
    listItemAfter: 80, codeBlockBefore: 80, codeBlockAfter: 80, codeBlockIndent: 400,
    blockquoteBefore: 80, blockquoteAfter: 80, blockquoteIndent: 400,
    thematicBreakBefore: 200, thematicBreakAfter: 200,
    imageBefore: 200, imageAfter: 200, imageCaptionSize: 18,
    tableAfter: 100, tableCellPadding: 40, formulaBefore: 200, formulaAfter: 200,
  },
  image: { maxWidth: 500, maxHeight: 700 },
  list: { indentPerLevel: 200, baseIndent: 400, bulletChar: "•" },
  table: { headerBgColor: "F2F2F2" },
};

function deepMerge(base, overrides) {
  const result = { ...base };
  for (const key of Object.keys(overrides || {})) {
    if (overrides[key] && typeof overrides[key] === "object" && !Array.isArray(overrides[key])) {
      result[key] = deepMerge(base[key] || {}, overrides[key]);
    } else {
      result[key] = overrides[key];
    }
  }
  return result;
}

function loadConfig() {
  const searchPaths = [];
  if (configPath) searchPaths.push(path.resolve(configPath));
  searchPaths.push(path.resolve("word-config.json"));
  searchPaths.push(path.join(path.dirname(inputFile), "word-config.json"));

  for (const p of searchPaths) {
    if (existsSync(p)) {
      try {
        const raw = readFileSync(p, "utf-8");
        const userCfg = JSON.parse(raw);
        console.log(`📋 加载配置: ${p}`);
        return deepMerge(DEFAULTS, userCfg);
      } catch (e) {
        console.warn(`⚠️  配置加载失败 ${p}: ${e.message}，使用默认配置`);
      }
    }
  }
  console.log("📋 使用默认配置（可创建 word-config.json 自定义格式）");
  return { ...DEFAULTS };
}

const cfg = loadConfig();
const content = readFileSync(inputFile, "utf-8");

// Convenience aliases
const BF = cfg.fonts.body;
const BS = cfg.fonts.bodySize;
const CF = cfg.fonts.code;
const CS = cfg.fonts.codeSize;
const HF = cfg.fonts.heading;
const FF = cfg.fonts.formula;
const S = cfg.spacing;
const C = cfg.colors;

// ─── Inline parser ────────────────────────────────────────────────
function parseInline(text) {
  const children = [];
  let remaining = text;
  const regex = /(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|`(.+?)`|\$(.+?)\$|\[(.+?)\]\((.+?)\)|\*(.+?)\*)/;

  while (true) {
    const match = remaining.match(regex);
    if (!match) break;

    const before = remaining.slice(0, match.index);
    if (before) {
      children.push(new TextRun({ text: before, size: BS, font: BF, color: C.body }));
    }

    if (match[2]) {
      children.push(new TextRun({ text: match[2], size: BS, font: BF, bold: true, italics: true }));
    } else if (match[3]) {
      children.push(new TextRun({ text: match[3], size: BS, font: BF, bold: true }));
    } else if (match[4]) {
      children.push(new TextRun({ text: match[4], size: CS, font: CF, color: C.code }));
    } else if (match[5]) {
      children.push(new TextRun({ text: match[5], size: BS, font: FF, italics: true, color: C.formula }));
    } else if (match[6] && match[7]) {
      children.push(
        new ExternalHyperlink({
          children: [new TextRun({ text: match[6], size: BS, font: BF, color: C.link, underline: { type: "single" } })],
          link: match[7],
        })
      );
    } else if (match[8]) {
      children.push(new TextRun({ text: match[8], size: BS, font: BF, italics: true }));
    }

    remaining = remaining.slice(match.index + match[0].length);
  }

  if (remaining) {
    children.push(new TextRun({ text: remaining, size: BS, font: BF, color: C.body }));
  }

  return children.length > 0
    ? children
    : [new TextRun({ text, size: BS, font: BF, color: C.body })];
}

// ─── Table parser ─────────────────────────────────────────────────
function parseTableAlignment(line) {
  const inner = line.trim().replace(/^\||\|$/g, "");
  return inner.split("|").map(s => {
    const cell = s.trim();
    const left = cell.startsWith(":");
    const right = cell.endsWith(":");
    if (left && right) return AlignmentType.CENTER;
    if (right) return AlignmentType.RIGHT;
    return AlignmentType.LEFT;
  });
}

function parseTableRow(line) {
  const inner = line.trim().replace(/^\||\|$/g, "");
  return inner.split("|").map(c => c.trim());
}

function buildTable(headerCells, alignments, dataRows) {
  const rows = [];

  rows.push(new TableRow({
    tableHeader: true,
    children: headerCells.map((cell, i) => {
      const alignment = alignments[i] || AlignmentType.LEFT;
      return new TableCell({
        children: [new Paragraph({
          alignment,
          spacing: { before: S.tableCellPadding, after: S.tableCellPadding },
          children: parseInline(cell),
        })],
        shading: { type: "clear", fill: cfg.table.headerBgColor },
      });
    }),
  }));

  for (const row of dataRows) {
    rows.push(new TableRow({
      children: row.map((cell, i) => {
        const alignment = alignments[i] || AlignmentType.LEFT;
        return new TableCell({
          children: [new Paragraph({
            alignment,
            spacing: { before: S.tableCellPadding, after: S.tableCellPadding },
            children: parseInline(cell),
          })],
        });
      }),
    }));
  }

  return new Table({ rows });
}

// ─── Image handler ────────────────────────────────────────────────
function buildImage(alt, imgPath) {
  const elements = [];
  const placeholderColor = C.imagePlaceholder;

  if (!existsSync(imgPath)) {
    elements.push(new Paragraph({
      children: [new TextRun({ text: `[图片: ${alt}]`, size: BS, font: BF, color: placeholderColor, italics: true })],
    }));
    return elements;
  }

  try {
    const imageBuffer = readFileSync(imgPath);
    const dimensions = imageSize(imageBuffer);
    const maxW = cfg.image.maxWidth;
    const maxH = cfg.image.maxHeight;
    let width = dimensions.width;
    let height = dimensions.height;
    if (width > maxW) {
      height = Math.round((height / width) * maxW);
      width = maxW;
    }
    if (height > maxH) {
      width = Math.round((width / height) * maxH);
      height = maxH;
    }

    elements.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: S.imageBefore, after: alt ? 40 : S.imageAfter },
      children: [
        new ImageRun({
          data: imageBuffer,
          type: dimensions.type,
          transformation: { width, height },
        }),
      ],
    }));

    if (alt) {
      elements.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: S.imageAfter },
        children: [new TextRun({ text: alt, size: cfg.spacing.imageCaptionSize, font: BF, color: "888888", italics: true })],
      }));
    }
  } catch {
    elements.push(new Paragraph({
      children: [new TextRun({ text: `[图片加载失败: ${alt}]`, size: BS, font: BF, color: placeholderColor, italics: true })],
    }));
  }

  return elements;
}

// ─── Main parser ──────────────────────────────────────────────────
function parseMarkdownToDocx(md, mdFilePath) {
  const elements = [];
  const mdLines = md.split("\n");
  let i = 0;
  let olCounter = 0;
  const mdDir = mdFilePath ? path.dirname(path.resolve(mdFilePath)) : process.cwd();

  while (i < mdLines.length) {
    const line = mdLines[i];
    if (line.trim() === "") { i++; continue; }

    // ── Code block ──
    if (line.trimStart().startsWith("```")) {
      const codeLines = [];
      i++;
      while (i < mdLines.length && !mdLines[i].trimStart().startsWith("```")) {
        codeLines.push(mdLines[i]);
        i++;
      }
      i++;
      elements.push(
        new Paragraph({
          spacing: { before: S.codeBlockBefore, after: S.codeBlockAfter },
          indent: { left: S.codeBlockIndent },
          children: [new TextRun({ text: codeLines.join("\n"), size: CS, font: CF, color: C.code })],
        })
      );
      continue;
    }

    // ── Block formula $$...$$ ──
    if (line.trim().startsWith("$$")) {
      olCounter = 0;
      const formulaLines = [];
      i++;
      while (i < mdLines.length && !mdLines[i].trim().startsWith("$$")) {
        formulaLines.push(mdLines[i]);
        i++;
      }
      i++;
      const formulaText = formulaLines.join("\n").trim() || line.replace(/\$\$/g, "").trim();
      if (formulaText) {
        elements.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: S.formulaBefore, after: S.formulaAfter },
            children: [new TextRun({ text: formulaText, size: BS, font: FF, italics: true, color: C.formula })],
          })
        );
      }
      continue;
    }

    // ── Headings ──
    const hMatch = line.match(/^(#{1,6})\s+(.+)$/);
    if (hMatch) {
      olCounter = 0;
      const level = hMatch[1].length;
      const levels = [null,
        HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3,
        HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6,
      ];
      elements.push(
        new Paragraph({
          text: hMatch[2],
          heading: levels[level] || HeadingLevel.HEADING_1,
          spacing: { before: level === 1 ? S.headingBefore : S.headingBeforeNested, after: S.headingAfter },
        })
      );
      i++;
      continue;
    }

    // ── Table ──
    if (line.trimStart().startsWith("|")) {
      olCounter = 0;
      const tableLines = [];
      while (i < mdLines.length && mdLines[i].trimStart().startsWith("|")) {
        tableLines.push(mdLines[i]);
        i++;
      }
      if (tableLines.length >= 2) {
        const alignments = parseTableAlignment(tableLines[1]);
        const headerCells = parseTableRow(tableLines[0]);
        const dataRows = tableLines.slice(2).map(l => parseTableRow(l));
        elements.push(buildTable(headerCells, alignments, dataRows));
        elements.push(new Paragraph({ spacing: { before: S.tableAfter, after: S.tableAfter }, children: [] }));
      }
      continue;
    }

    // ── Unordered list ──
    const ulMatch = line.match(/^(\s*)[-*+]\s+(.+)$/);
    if (ulMatch) {
      olCounter = 0;
      const indent = ulMatch[1].length;
      const text = ulMatch[2];
      elements.push(
        new Paragraph({
          spacing: { after: S.listItemAfter },
          indent: { left: cfg.list.baseIndent + indent * cfg.list.indentPerLevel },
          children: [
            new TextRun({ text: `  ${cfg.list.bulletChar}  `, size: BS, font: BF }),
            ...parseInline(text),
          ],
        })
      );
      i++;
      continue;
    }

    // ── Ordered list ──
    const olMatch = line.match(/^(\s*)\d+[.)]\s+(.+)$/);
    if (olMatch) {
      olCounter++;
      const indent = olMatch[1].length;
      const text = olMatch[2];
      elements.push(
        new Paragraph({
          spacing: { after: S.listItemAfter },
          indent: { left: cfg.list.baseIndent + indent * cfg.list.indentPerLevel },
          children: [
            new TextRun({ text: `  ${olCounter}. `, size: BS, font: BF }),
            ...parseInline(text),
          ],
        })
      );
      i++;
      continue;
    }
    olCounter = 0;

    // ── Thematic break ──
    if (/^[-*_]{3,}\s*$/.test(line.trim())) {
      elements.push(
        new Paragraph({
          spacing: { before: S.thematicBreakBefore, after: S.thematicBreakAfter },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "— • — • —", size: BS, color: C.thematicBreak, font: BF })],
        })
      );
      i++;
      continue;
    }

    // ── Blockquote ──
    if (line.trimStart().startsWith("> ")) {
      olCounter = 0;
      const quoteLines = [];
      while (i < mdLines.length && mdLines[i].trimStart().startsWith("> ")) {
        quoteLines.push(mdLines[i].trimStart().slice(2));
        i++;
      }
      elements.push(
        new Paragraph({
          spacing: { before: S.blockquoteBefore, after: S.blockquoteAfter },
          indent: { left: S.blockquoteIndent },
          children: [new TextRun({ text: quoteLines.join("\n"), size: BS, font: BF, italics: true, color: C.quote })],
        })
      );
      continue;
    }

    // ── Standalone image ──
    const imgMatch = line.match(/^!\[(.*)\]\((.+)\)$/);
    if (imgMatch) {
      olCounter = 0;
      let imgPath = imgMatch[2];
      if (/^https?:\/\//.test(imgPath)) {
        elements.push(new Paragraph({
          children: [new TextRun({ text: `[图片: ${imgMatch[1]}] (${imgPath})`, size: BS, font: BF, color: C.imagePlaceholder, italics: true })],
        }));
      } else {
        if (!path.isAbsolute(imgPath)) {
          imgPath = path.resolve(mdDir, imgPath);
        }
        elements.push(...buildImage(imgMatch[1], imgPath));
      }
      i++;
      continue;
    }

    // ── Regular paragraph ──
    const paraLines = [];
    while (i < mdLines.length && mdLines[i].trim() !== "") {
      paraLines.push(mdLines[i]);
      i++;
    }
    while (i < mdLines.length && mdLines[i].trim() === "") { i++; }

    const paraText = paraLines.join(" ").replace(/\n/g, "");
    if (paraText.trim()) {
      elements.push(
        new Paragraph({
          spacing: { after: S.paragraphAfter },
          children: parseInline(paraText.trim()),
        })
      );
    }
  }

  return elements;
}

// ─── Entry ────────────────────────────────────────────────────────
const elements = parseMarkdownToDocx(content, inputFile);

const doc = new Document({
  sections: [{
    properties: {
      page: {
        margin: {
          top: cfg.page.marginTop,
          bottom: cfg.page.marginBottom,
          left: cfg.page.marginLeft,
          right: cfg.page.marginRight,
        },
      },
    },
    children: elements,
  }],
});

const buffer = await Packer.toBuffer(doc);
writeFileSync(outputFile, buffer);
console.log(`✅ Word 文档已生成: ${outputFile}`);
