import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } from "docx";
import { readFileSync, writeFileSync } from "fs";

const args = process.argv.slice(2);
if (args.length < 1) {
  console.error("用法: node md-to-word.js <输入.md> [输出.docx]");
  process.exit(1);
}

const inputFile = args[0];
const outputFile = args[1] || inputFile.replace(/\.md$/i, ".docx");

const content = readFileSync(inputFile, "utf-8");

const bodyFont = "Microsoft YaHei";
const bodySize = 24; // half-pts = 12pt

function parseInline(text) {
  const children = [];
  let remaining = text;
  const regex = /(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*)/;

  while (true) {
    const match = remaining.match(regex);
    if (!match) break;

    const before = remaining.slice(0, match.index);
    if (before) {
      children.push(new TextRun({ text: before, size: bodySize, font: bodyFont }));
    }

    if (match[2]) {
      children.push(new TextRun({ text: match[2], size: bodySize, font: bodyFont, bold: true, italics: true }));
    } else if (match[3]) {
      children.push(new TextRun({ text: match[3], size: bodySize, font: bodyFont, bold: true }));
    } else if (match[4]) {
      children.push(new TextRun({ text: match[4], size: bodySize, font: bodyFont, italics: true }));
    }

    remaining = remaining.slice(match.index + match[0].length);
  }

  if (remaining) {
    children.push(new TextRun({ text: remaining, size: bodySize, font: bodyFont }));
  }

  return children.length > 0
    ? children
    : [new TextRun({ text, size: bodySize, font: bodyFont })];
}

function parseMarkdownToDocx(md) {
  const paragraphs = [];
  const mdLines = md.split("\n");
  let i = 0;
  let olCounter = 0;

  while (i < mdLines.length) {
    const line = mdLines[i];
    if (line.trim() === "") { i++; continue; }

    // Code block
    if (line.trimStart().startsWith("```")) {
      const codeLines = [];
      i++;
      while (i < mdLines.length && !mdLines[i].trimStart().startsWith("```")) {
        codeLines.push(mdLines[i]);
        i++;
      }
      i++;
      paragraphs.push(
        new Paragraph({
          spacing: { before: 80, after: 80 },
          indent: { left: 400 },
          children: [new TextRun({ text: codeLines.join("\n"), size: 20, font: "Consolas", color: "333333" })],
        })
      );
      continue;
    }

    // Headings
    const hMatch = line.match(/^(#{1,6})\s+(.+)$/);
    if (hMatch) {
      olCounter = 0;
      const level = hMatch[1].length;
      const text = hMatch[2];
      const levels = [null, HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3, HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6];
      paragraphs.push(
        new Paragraph({
          text,
          heading: levels[level] || HeadingLevel.HEADING_1,
          spacing: { before: level === 1 ? 360 : 240, after: 160 },
        })
      );
      i++;
      continue;
    }

    // Unordered list
    const ulMatch = line.match(/^(\s*)[-*+]\s+(.+)$/);
    if (ulMatch) {
      olCounter = 0;
      const indent = ulMatch[1].length;
      const text = ulMatch[2];
      paragraphs.push(
        new Paragraph({
          spacing: { after: 80 },
          indent: { left: 400 + indent * 200 },
          children: [new TextRun({ text: `  •  ${text}`, size: bodySize, font: bodyFont })],
        })
      );
      i++;
      continue;
    }

    // Ordered list
    const olMatch = line.match(/^(\s*)\d+[.)]\s+(.+)$/);
    if (olMatch) {
      olCounter++;
      const indent = olMatch[1].length;
      const text = olMatch[2];
      paragraphs.push(
        new Paragraph({
          spacing: { after: 80 },
          indent: { left: 400 + indent * 200 },
          children: [new TextRun({ text: `  ${olCounter}. ${text}`, size: bodySize, font: bodyFont })],
        })
      );
      i++;
      continue;
    }
    olCounter = 0;

    // Thematic break
    if (/^[-*_]{3,}\s*$/.test(line.trim())) {
      paragraphs.push(
        new Paragraph({
          spacing: { before: 200, after: 200 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "— • — • —", size: bodySize, color: "999999", font: bodyFont })],
        })
      );
      i++;
      continue;
    }

    // Blockquote
    if (line.trimStart().startsWith("> ")) {
      const quoteLines = [];
      while (i < mdLines.length && mdLines[i].trimStart().startsWith("> ")) {
        quoteLines.push(mdLines[i].trimStart().slice(2));
        i++;
      }
      paragraphs.push(
        new Paragraph({
          spacing: { before: 80, after: 80 },
          indent: { left: 400 },
          children: [new TextRun({ text: quoteLines.join("\n"), size: bodySize, font: bodyFont, italics: true, color: "666666" })],
        })
      );
      continue;
    }

    // Regular paragraph
    const paraLines = [];
    while (i < mdLines.length && mdLines[i].trim() !== "") {
      paraLines.push(mdLines[i]);
      i++;
    }
    while (i < mdLines.length && mdLines[i].trim() === "") { i++; }

    const paraText = paraLines.join("").replace(/\n/g, "");
    if (paraText.trim()) {
      paragraphs.push(
        new Paragraph({
          spacing: { after: 120 },
          children: parseInline(paraText.trim()),
        })
      );
    }
  }

  return paragraphs;
}

const children = parseMarkdownToDocx(content);

const doc = new Document({
  sections: [{ properties: {}, children }],
});

const buffer = await Packer.toBuffer(doc);
writeFileSync(outputFile, buffer);
console.log(`✅ Word 文档已生成: ${outputFile}`);
