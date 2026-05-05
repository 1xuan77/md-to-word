#!/usr/bin/env python3
"""
Markdown → Word 转换器
用法: python md-to-word.py <输入.md> [输出.docx] [--config 配置.json]

配置说明：
  1. 直接修改本文件顶部的 CONFIG 字典（最方便）
  2. 或创建 word-config.json，脚本会自动加载
  3. 或运行 --config my-config.json 指定配置
  三者会合并（优先级: --config > word-config.json > 内置默认值）
"""

# ═══════════════════════════════════════════════════════════════
#  🎨 格式配置 —— 按需修改这里的值
# ═══════════════════════════════════════════════════════════════
CONFIG = {
    # ── 字体 ──────────────────────────────────────────────────
    "fonts": {
        "body": "Times New Roman",        # 西文正文字体
        "body_east": "宋体",              # 中文正文字体
        "body_size": 12,                  # 正文字号（pt）
        "code": "Consolas",               # 代码字体
        "code_size": 10,                  # 代码字号（pt）
        "heading": "Times New Roman",     # 西文标题字体
        "heading_east": "黑体",           # 中文标题字体
        "formula": "Consolas",            # 公式字体
    },
    # ── 颜色 ──────────────────────────────────────────────────
    "colors": {
        "link": "0563C1",                 # 超链接（十六进制）
        "code": "333333",                 # 行内代码
        "quote": "666666",                # 引用
        "body": "333333",                 # 正文
        "formula": "444444",              # 公式
        "heading": "1F1F1F",              # 标题
    },
    # ── 间距 ──────────────────────────────────────────────────
    "spacing": {
        "para_after": 6,                  # 段落后（pt）
        "heading_before": 18,             # 一级标题前（pt）
        "heading_before_nested": 12,      # 子标题前（pt）
        "heading_after": 8,               # 标题后（pt）
        "list_item_after": 4,             # 列表项后（pt）
        "code_block_before": 4,           # 代码块前（pt）
        "code_block_after": 4,            # 代码块后（pt）
        "code_block_indent": 0.3,         # 代码块缩进（inch）
        "blockquote_before": 4,           # 引用前（pt）
        "blockquote_after": 4,            # 引用后（pt）
        "blockquote_indent": 0.3,         # 引用缩进（inch）
        "image_before": 6,                # 图片前（pt）
        "image_after": 6,                 # 图片后（pt）
        "table_after": 4,                 # 表格后（pt）
        "formula_before": 6,              # 公式前（pt）
        "formula_after": 6,               # 公式后（pt）
    },
    # ── 页面 ──────────────────────────────────────────────────
    "page": {
        "margin_top": 1.0,                # 上边距（inch）
        "margin_bottom": 1.0,             # 下边距（inch）
        "margin_left": 1.0,               # 左边距（inch）
        "margin_right": 1.0,              # 右边距（inch）
    },
    # ── 图片 ──────────────────────────────────────────────────
    "image": {
        "max_width": 5.0,                 # 最大宽度（inch）
    },
    # ── 列表 ──────────────────────────────────────────────────
    "list": {
        "indent_per_level": 0.2,          # 每级缩进增量（inch）
        "base_indent": 0.3,               # 基础缩进（inch）
        "bullet_char": "•",               # 无序符号
    },
    # ── 代码块 ────────────────────────────────────────────────
    "code_block_size": 9,                 # 代码块字号（pt）
}

# ═══════════════════════════════════════════════════════════════
#  以下为转换引擎，一般不需要修改
# ═══════════════════════════════════════════════════════════════

import sys, os, json, re, struct

# Fix stdout encoding for Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement

# ─── 工具函数 ─────────────────────────────────────────────────────

def hex_to_rgb(hex_str):
    """#RRGGBB → (R, G, B)"""
    h = hex_str.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def pt(val):
    return Pt(val)

def inch(val):
    return Inches(val)

# ─── 配置加载 ─────────────────────────────────────────────────────

def deep_merge(base, override):
    result = dict(base)
    for k, v in (override or {}).items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            result[k] = deep_merge(base[k], v)
        else:
            result[k] = v
    return result

def load_config():
    cfg_path = None
    # --config
    if "--config" in sys.argv:
        idx = sys.argv.index("--config")
        if idx + 1 < len(sys.argv):
            cfg_path = sys.argv.pop(idx + 1)
            sys.argv.pop(idx)

    search = []
    if cfg_path:
        search.append(os.path.abspath(cfg_path))
    search.append(os.path.abspath("word-config.json"))
    search.append(os.path.join(os.path.dirname(os.path.abspath(sys.argv[1])), "word-config.json") if len(sys.argv) > 1 and not sys.argv[1].startswith("-") else "")

    cfg = dict(CONFIG)  # deep copy
    for p in search:
        if p and os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    user = json.load(f)
                cfg = deep_merge(cfg, user)
                print(f"📋 加载配置: {p}")
                return cfg
            except Exception as e:
                print(f"⚠️  配置 {p} 加载失败: {e}")
    print("📋 使用内置配置（可创建 word-config.json 自定义）")
    return cfg

# ─── 行内解析器 ───────────────────────────────────────────────────

INLINE_RE = re.compile(
    r"(\*\*\*(.+?)\*\*\*"         # 1,2: ***bold italic***
    r"|\*\*(.+?)\*\*"             # 3: **bold**
    r"|`(.+?)`"                   # 4: `code`
    r"|\$(.+?)\$"                 # 5: $formula$
    r"|\[(.+?)\]\((.+?)\)"        # 6,7: [link](url)
    r"|\*(.+?)\*)",               # 8: *italic*
)

def parse_inline(text, cfg):
    """Parse inline formatting → list of (text, props_dict) tuples."""
    f = cfg["fonts"]
    c = cfg["colors"]
    parts = []
    pos = 0

    for m in INLINE_RE.finditer(text):
        if m.start() > pos:
            parts.append((text[pos:m.start()], {"font": f["body"], "size": f["body_size"], "color": c["body"]}))

        if m.group(2):   # ***bold italic***
            parts.append((m.group(2), {"bold": True, "italic": True, "font": f["body"], "size": f["body_size"]}))
        elif m.group(3): # **bold**
            parts.append((m.group(3), {"bold": True, "font": f["body"], "size": f["body_size"]}))
        elif m.group(4): # `code`
            parts.append((m.group(4), {"font": f["code"], "size": f["code_size"], "color": c["code"]}))
        elif m.group(5): # $formula$
            parts.append((m.group(5), {"font": f["formula"], "size": f["body_size"], "italic": True, "color": c["formula"]}))
        elif m.group(6): # [link](url)
            parts.append(("LINK:" + m.group(6) + "|" + m.group(7), {"font": f["body"], "size": f["body_size"], "color": c["link"], "underline": True}))
        elif m.group(8): # *italic*
            parts.append((m.group(8), {"italic": True, "font": f["body"], "size": f["body_size"]}))

        pos = m.end()

    if pos < len(text):
        parts.append((text[pos:], {"font": f["body"], "size": f["body_size"], "color": c["body"]}))

    return parts

def add_run(para, text, props, cfg):
    """Add a styled run to paragraph."""
    run = para.add_run(text)
    if props.get("bold"): run.bold = True
    if props.get("italic"): run.italic = True
    if props.get("underline"): run.underline = True
    if props.get("font"):
        set_run_font(run, props["font"], cfg)
    if props.get("size"): run.font.size = Pt(props["size"])
    if props.get("color"):
        try:
            run.font.color.rgb = RGBColor(*hex_to_rgb(props["color"]))
        except:
            pass
    return run

def set_run_font(run, font_name, cfg, font_type="body"):
    """Set font with western+east-Asian support.
    font_type: 'body' or 'heading' — determines which east-Asian font to use."""
    f = cfg["fonts"]
    run.font.name = font_name
    east_key = f"{font_type}_east"
    east_font = f.get(east_key, font_name)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        from docx.oxml import OxmlElement
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:eastAsia"), east_font)

def apply_inline(para, parts, cfg):
    """Apply parsed inline parts to a paragraph."""
    for text, props in parts:
        if text.startswith("LINK:"):
            # Hyperlink
            rest = text[5:]
            link_text, url = rest.split("|", 1)
            add_hyperlink(para, link_text, url, cfg)
        else:
            add_run(para, text, props, cfg)

def add_hyperlink(paragraph, text, url, cfg):
    """Add a clickable hyperlink to paragraph."""
    c = cfg["colors"]
    part = paragraph.part
    r_id = part.relate_to(url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    run_elem = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    # Color
    clr = OxmlElement("w:color")
    clr.set(qn("w:val"), c["link"])
    rPr.append(clr)
    # Underline
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)
    # Font
    rFonts = OxmlElement("w:rFonts")
    f = cfg["fonts"]
    rFonts.set(qn("w:ascii"), f["body"])
    rFonts.set(qn("w:hAnsi"), f["body"])
    rFonts.set(qn("w:eastAsia"), f.get("body_east", f["body"]))
    rPr.append(rFonts)
    # Size
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(f["body_size"] * 2))
    rPr.append(sz)

    run_elem.append(rPr)
    run_elem.text = text
    hyperlink.append(run_elem)
    paragraph._p.append(hyperlink)

# ─── 间距辅助 ─────────────────────────────────────────────────────

def set_space(para, before=0, after=0, line=None):
    """Set paragraph spacing."""
    pf = para.paragraph_format
    if before: pf.space_before = Pt(before)
    if after: pf.space_after = Pt(after)
    if line: pf.line_spacing = Pt(line)

def set_indent(para, left=0):
    """Set paragraph indent in inches."""
    if left:
        para.paragraph_format.left_indent = inch(left)

def set_alignment(para, align):
    para.alignment = align

# ─── 表格解析 ─────────────────────────────────────────────────────

def parse_table_alignment(line):
    """Parse alignment row: |:---|---:|:---:|"""
    inner = line.strip().strip("|")
    aligns = []
    for cell in inner.split("|"):
        cell = cell.strip()
        left = cell.startswith(":")
        right = cell.endswith(":")
        if left and right:
            aligns.append(WD_ALIGN_PARAGRAPH.CENTER)
        elif right:
            aligns.append(WD_ALIGN_PARAGRAPH.RIGHT)
        else:
            aligns.append(WD_ALIGN_PARAGRAPH.LEFT)
    return aligns

def parse_table_row(line):
    """Parse | a | b | c |"""
    inner = line.strip().strip("|")
    return [c.strip() for c in inner.split("|")]

# ─── 图片 ─────────────────────────────────────────────────────────

def get_image_dims(path, max_width_inch):
    """Get (width, height) in inches, constrained by max_width."""
    try:
        from PIL import Image
        with Image.open(path) as img:
            w, h = img.size
    except ImportError:
        # Fallback: read PNG/JPEG headers manually
        w, h = _read_image_dims(path)
        if w is None:
            return max_width_inch, max_width_inch * 0.75  # guess

    aspect = h / w
    target_w = min(w / 96, max_width_inch)  # assume 96 DPI
    return target_w, target_w * aspect

def _read_image_dims(path):
    """Minimal PNG/JPEG dimension reader (no Pillow)."""
    with open(path, "rb") as f:
        sig = f.read(8)
        if sig[:8] == b"\x89PNG\r\n\x1a\n":
            f.read(4)  # chunk length
            f.read(4)  # IHDR
            w = struct.unpack(">I", f.read(4))[0]
            h = struct.unpack(">I", f.read(4))[0]
            return w, h
        elif sig[:2] == b"\xff\xd8":
            f.seek(0)
            while True:
                marker = f.read(2)
                if not marker: break
                if marker[0] != 0xFF: break
                if marker[1] == 0xC0 or marker[1] == 0xC2:
                    f.read(3)
                    h = struct.unpack(">H", f.read(2))[0]
                    w = struct.unpack(">H", f.read(2))[0]
                    return w, h
                else:
                    length = struct.unpack(">H", f.read(2))[0]
                    f.seek(f.tell() + length - 2)
    return None, None

# ─── 主解析器 ─────────────────────────────────────────────────────

def convert_md_to_docx(md_text, md_path, cfg):
    doc = Document()
    f = cfg["fonts"]
    s = cfg["spacing"]
    pg = cfg["page"]
    c = cfg["colors"]

    # ── 页面设置 ──
    section = doc.sections[0]
    section.top_margin = inch(pg["margin_top"])
    section.bottom_margin = inch(pg["margin_bottom"])
    section.left_margin = inch(pg["margin_left"])
    section.right_margin = inch(pg["margin_right"])

    lines = md_text.split("\n")
    i = 0
    ol_counter = 0
    md_dir = os.path.dirname(os.path.abspath(md_path)) if md_path else os.getcwd()

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if stripped == "":
            i += 1
            continue

        # ── 代码块 ──
        if line.lstrip().startswith("```"):
            code = []
            i += 1
            while i < len(lines) and not lines[i].lstrip().startswith("```"):
                code.append(lines[i])
                i += 1
            i += 1
            p = doc.add_paragraph()
            set_space(p, before=s["code_block_before"], after=s["code_block_after"])
            set_indent(p, s["code_block_indent"])
            r = p.add_run("\n".join(code))
            r.font.name = f["code"]
            r.font.size = Pt(cfg["code_block_size"])
            r.font.color.rgb = RGBColor(*hex_to_rgb(c["code"]))
            continue

        # ── 块级公式 $$...$$ ──
        if stripped.startswith("$$"):
            formula_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("$$"):
                formula_lines.append(lines[i])
                i += 1
            i += 1
            formula = "\n".join(formula_lines).strip() or stripped.replace("$$", "").strip()
            if formula:
                p = doc.add_paragraph()
                set_alignment(p, WD_ALIGN_PARAGRAPH.CENTER)
                set_space(p, before=s["formula_before"], after=s["formula_after"])
                r = p.add_run(formula)
                r.font.name = f["formula"]
                r.font.size = Pt(f["body_size"])
                r.italic = True
                r.font.color.rgb = RGBColor(*hex_to_rgb(c["formula"]))
            continue

        # ── 标题 ──
        hm = re.match(r"^(#{1,6})\s+(.+)$", line)
        if hm:
            ol_counter = 0
            level = len(hm.group(1))
            text = hm.group(2)
            p = doc.add_heading(text, level=level)
            space_before = s["heading_before"] if level == 1 else s["heading_before_nested"]
            set_space(p, before=space_before, after=s["heading_after"])
            # Override heading font (western + east-Asian)
            for run in p.runs:
                set_run_font(run, f["heading"], cfg, "heading")
            i += 1
            continue

        # ── 表格 ──
        if line.lstrip().startswith("|"):
            ol_counter = 0
            table_lines = []
            while i < len(lines) and lines[i].lstrip().startswith("|"):
                table_lines.append(lines[i])
                i += 1
            if len(table_lines) >= 2:
                aligns = parse_table_alignment(table_lines[1])
                header = parse_table_row(table_lines[0])
                data = [parse_table_row(r) for r in table_lines[2:]]

                cols = len(header)
                rows = 1 + len(data)
                table = doc.add_table(rows=rows, cols=cols)
                table.autofit = True

                # Header
                for ci, cell_text in enumerate(header):
                    cell = table.rows[0].cells[ci]
                    cell.text = ""
                    p = cell.paragraphs[0]
                    if ci < len(aligns):
                        p.alignment = aligns[ci]
                    r = p.add_run(cell_text)
                    r.bold = True
                    set_run_font(r, f["body"], cfg)
                    r.font.size = Pt(f["body_size"])
                    # Gray background
                    shading = OxmlElement("w:shd")
                    shading.set(qn("w:fill"), "F2F2F2")
                    shading.set(qn("w:val"), "clear")
                    cell._tc.get_or_add_tcPr().append(shading)

                # Data rows
                for ri, row_data in enumerate(data):
                    for ci, cell_text in enumerate(row_data):
                        if ci < cols:
                            cell = table.rows[ri + 1].cells[ci]
                            cell.text = ""
                            p = cell.paragraphs[0]
                            if ci < len(aligns):
                                p.alignment = aligns[ci]
                            parts = parse_inline(cell_text, cfg)
                            apply_inline(p, parts, cfg)

                # Spacing after table
                p = doc.add_paragraph()
                set_space(p, before=s["table_after"], after=s["table_after"])
            continue

        # ── 无序列表 ──
        ulm = re.match(r"^(\s*)[-*+]\s+(.+)$", line)
        if ulm:
            ol_counter = 0
            indent_level = len(ulm.group(1))
            text = ulm.group(2)
            p = doc.add_paragraph()
            set_space(p, after=s["list_item_after"])
            set_indent(p, cfg["list"]["base_indent"] + indent_level * cfg["list"]["indent_per_level"])
            r = p.add_run(f"  {cfg['list']['bullet_char']}  ")
            set_run_font(r, f["body"], cfg)
            r.font.size = Pt(f["body_size"])
            parts = parse_inline(text, cfg)
            apply_inline(p, parts, cfg)
            i += 1
            continue

        # ── 有序列表 ──
        olm = re.match(r"^(\s*)\d+[.)]\s+(.+)$", line)
        if olm:
            ol_counter += 1
            indent_level = len(olm.group(1))
            text = olm.group(2)
            p = doc.add_paragraph()
            set_space(p, after=s["list_item_after"])
            set_indent(p, cfg["list"]["base_indent"] + indent_level * cfg["list"]["indent_per_level"])
            r = p.add_run(f"  {ol_counter}. ")
            set_run_font(r, f["body"], cfg)
            r.font.size = Pt(f["body_size"])
            parts = parse_inline(text, cfg)
            apply_inline(p, parts, cfg)
            i += 1
            continue

        ol_counter = 0

        # ── 分隔线 ──
        if re.match(r"^[-*_]{3,}\s*$", stripped):
            p = doc.add_paragraph()
            set_alignment(p, WD_ALIGN_PARAGRAPH.CENTER)
            set_space(p, before=s.get("heading_before", 6), after=s.get("heading_after", 6))
            r = p.add_run("— • — • —")
            r.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
            r.font.size = Pt(f["body_size"])
            i += 1
            continue

        # ── 引用 ──
        if line.lstrip().startswith("> "):
            ol_counter = 0
            quote_lines = []
            while i < len(lines) and lines[i].lstrip().startswith("> "):
                quote_lines.append(lines[i].lstrip()[2:])
                i += 1
            p = doc.add_paragraph()
            set_space(p, before=s["blockquote_before"], after=s["blockquote_after"])
            set_indent(p, s["blockquote_indent"])
            r = p.add_run("\n".join(quote_lines))
            r.italic = True
            set_run_font(r, f["body"], cfg)
            r.font.size = Pt(f["body_size"])
            r.font.color.rgb = RGBColor(*hex_to_rgb(c["quote"]))
            continue

        # ── 独立图片 ──
        im = re.match(r"^!\[(.*)\]\((.+)\)$", line)
        if im:
            ol_counter = 0
            alt = im.group(1)
            img_path = im.group(2)

            if re.match(r"^https?://", img_path):
                p = doc.add_paragraph()
                r = p.add_run(f"[图片: {alt}] ({img_path})")
                r.italic = True
                r.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
                set_run_font(r, f["body"], cfg)
                r.font.size = Pt(f["body_size"])
            else:
                if not os.path.isabs(img_path):
                    img_path = os.path.join(md_dir, img_path)
                if os.path.exists(img_path):
                    try:
                        w, h = get_image_dims(img_path, cfg["image"]["max_width"])
                        p = doc.add_paragraph()
                        set_alignment(p, WD_ALIGN_PARAGRAPH.CENTER)
                        set_space(p, before=s["image_before"], after=s["image_after"])
                        run = p.add_run()
                        run.add_picture(img_path, width=inch(w))
                        if alt:
                            cap = doc.add_paragraph()
                            set_alignment(cap, WD_ALIGN_PARAGRAPH.CENTER)
                            r = cap.add_run(alt)
                            r.italic = True
                            r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
                            r.font.size = Pt(9)
                    except Exception as e:
                        p = doc.add_paragraph()
                        p.add_run(f"[图片加载失败: {alt}]").italic = True
                else:
                    p = doc.add_paragraph()
                    p.add_run(f"[图片: {alt}]").italic = True
            i += 1
            continue

        # ── 正文段落 ──
        para_lines = []
        while i < len(lines) and lines[i].strip() != "":
            para_lines.append(lines[i])
            i += 1
        while i < len(lines) and lines[i].strip() == "":
            i += 1

        text = " ".join(para_lines).replace("\n", "")
        if text.strip():
            p = doc.add_paragraph()
            set_space(p, after=s["para_after"])
            parts = parse_inline(text.strip(), cfg)
            apply_inline(p, parts, cfg)

    return doc


# ─── 入口 ─────────────────────────────────────────────────────────

def main():
    cfg = load_config()

    # Parse positional args
    positional = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not positional:
        print("用法: python md-to-word.py <输入.md> [输出.docx]")
        print("      python md-to-word.py 文档.md")
        print("      python md-to-word.py 文档.md 输出.docx --config my.json")
        sys.exit(1)

    input_file = os.path.abspath(positional[0])
    output_file = os.path.abspath(positional[1]) if len(positional) > 1 else re.sub(r"\.md$", ".docx", input_file, flags=re.I)

    if not os.path.exists(input_file):
        print(f"❌ 文件不存在: {input_file}")
        sys.exit(1)

    with open(input_file, "r", encoding="utf-8") as f:
        md_content = f.read()

    doc = convert_md_to_docx(md_content, input_file, cfg)
    doc.save(output_file)
    print(f"✅ Word 文档已生成: {output_file}")


if __name__ == "__main__":
    main()
