#!/usr/bin/env python3
"""
mddoc - Markdown 转学术格式 DOCX

用法:
    python md2docx.py <input.md>          # 转换 Markdown 文件
    python md2docx.py <input.md> -o out.docx  # 指定输出路径
    python md2docx.py --text "markdown..."  # 直接转换文本

依赖:
    pip install python-docx Pillow requests mistune
"""

import argparse
import os
import re
import sys
import tempfile
import urllib.parse
from io import BytesIO

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from PIL import Image


# ============================================================
# 字体工具
# ============================================================

def set_run_font(run, cn_font, en_font="Times New Roman", size_pt=10.5, bold=False):
    """设置 run 的中英文字体、字号、加粗"""
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.name = en_font
    run.font.color.rgb = None  # 黑色
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), cn_font)
    rFonts.set(qn('w:ascii'), en_font)
    rFonts.set(qn('w:hAnsi'), en_font)


def add_empty_para(doc):
    """添加五号空行"""
    p = doc.add_paragraph()
    run = p.add_run('')
    set_run_font(run, '宋体', size_pt=10.5)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    return p


# ============================================================
# 图片工具
# ============================================================

def download_image(url):
    """下载图片到临时文件，返回路径。失败返回 None"""
    import requests
    headers = {'User-Agent': 'Mozilla/5.0 (compatible; mddoc-converter/1.0)'}
    try:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        suffix = os.path.splitext(urllib.parse.urlparse(url).path)[1] or '.png'
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(resp.content)
        tmp.close()
        return tmp.name
    except Exception:
        return None


def calc_image_size(img_path):
    """计算图片在 docx 中的显示尺寸：8-12cm 规则，等比缩放"""
    img = Image.open(img_path)
    dpi = img.info.get('dpi', (96, 96))
    dpi_x = dpi[0] if dpi and dpi[0] > 0 else 96
    dpi_y = dpi[1] if dpi and dpi[1] > 0 else 96

    width_cm = img.width / dpi_x * 2.54
    height_cm = img.height / dpi_y * 2.54

    if width_cm > 12:
        target_w = 12.0
    elif width_cm < 8:
        target_w = 8.0
    else:
        target_w = width_cm

    ratio = target_w / width_cm
    return Cm(target_w), Cm(height_cm * ratio)


def add_placeholder_image(doc, text):
    """生成占位图片：灰色矩形 + 文字描述"""
    from PIL import ImageDraw, ImageFont
    img = Image.new('RGB', (600, 200), (220, 220, 220))
    draw = ImageDraw.Draw(img)
    draw.text((10, 80), text, fill=(80, 80, 80))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    img.save(tmp, format='PNG')
    tmp.close()
    return tmp.name


# ============================================================
# 表格工具
# ============================================================

def set_three_line_table(table):
    """三线表：仅顶线粗+底线粗，表头格底线细，数据行之间无线，无竖线"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)

    borders = OxmlElement('w:tblBorders')

    # 顶线：粗 1.5pt
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), '12')
    top.set(qn('w:space'), '0')
    top.set(qn('w:color'), '000000')
    borders.append(top)

    # 底线：粗 1.5pt
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12')
    bottom.set(qn('w:space'), '0')
    bottom.set(qn('w:color'), '000000')
    borders.append(bottom)

    # 所有内部线+竖线全部关闭（不用insideH！数据行之间无线）
    for edge in ('left', 'right', 'insideH', 'insideV'):
        elem = OxmlElement(f'w:{edge}')
        elem.set(qn('w:val'), 'none')
        elem.set(qn('w:sz'), '0')
        elem.set(qn('w:space'), '0')
        elem.set(qn('w:color'), 'auto')
        borders.append(elem)

    tblPr.append(borders)

    # 表头行每格底部单独加细线 0.75pt（cell级border，非table级insideH）
    if table.rows:
        for cell in table.rows[0].cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            cell_bottom = OxmlElement('w:bottom')
            cell_bottom.set(qn('w:val'), 'single')
            cell_bottom.set(qn('w:sz'), '6')
            cell_bottom.set(qn('w:space'), '0')
            cell_bottom.set(qn('w:color'), '000000')
            tcBorders.append(cell_bottom)
            tcPr.append(tcBorders)


def add_page_break_before(paragraph):
    """在段落前插入分页符"""
    pPr = paragraph._element.get_or_add_pPr()
    pb = OxmlElement('w:pageBreakBefore')
    pPr.append(pb)


def set_outline_level(paragraph, level):
    """设置 outline level（XML方式，兼容所有python-docx版本）"""
    pPr = paragraph._element.get_or_add_pPr()
    ol = OxmlElement('w:outlineLvl')
    ol.set(qn('w:val'), str(level))
    pPr.append(ol)


# ============================================================
# Markdown 解析
# ============================================================

def parse_markdown(text):
    """解析 Markdown 文本，返回节点列表。
    每个节点: {'type': 'title'|'heading'|'para'|'image'|'table'|'code'|'list',
               'level': int (仅heading),
               'text': str,
               'children': [...] (仅list),
               'alt': str (仅image),
               'url': str (仅image),
               'header': list (仅table),
               'rows': list (仅table)}
    """
    import mistune

    nodes = []
    lines = text.split('\n')
    i = 0
    heading_count = 0  # 跟踪 # 标题数量，第一个是题目

    while i < len(lines):
        line = lines[i]

        # 代码块
        if line.strip().startswith('```'):
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            i += 1
            nodes.append({'type': 'code', 'text': '\n'.join(code_lines)})
            continue

        # 标题 (# ...)
        m = re.match(r'^(#{1,6})\s+(.*)', line)
        if m:
            level = len(m.group(1))
            title_text = m.group(2).strip()
            heading_count += 1
            if level == 1 and heading_count == 1:
                nodes.append({'type': 'title', 'text': title_text, 'level': 1})
            else:
                nodes.append({'type': 'heading', 'text': title_text, 'level': level})
            i += 1
            continue

        # 图片 (![alt](url))
        m = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)', line.strip())
        if m:
            nodes.append({'type': 'image', 'alt': m.group(1), 'url': m.group(2)})
            i += 1
            continue

        # 表格（检测 | 开头的行）
        if line.strip().startswith('|'):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i].strip())
                i += 1
            # 解析表格：跳过第二行（分隔符行）
            header = [c.strip() for c in table_lines[0].split('|')[1:-1]]
            rows = []
            for tl in table_lines[2:]:
                rows.append([c.strip() for c in tl.split('|')[1:-1]])
            nodes.append({'type': 'table', 'header': header, 'rows': rows})
            continue

        # 空行
        if not line.strip():
            nodes.append({'type': 'empty'})
            i += 1
            continue

        # 无序列表
        m = re.match(r'^(\s*)[-*+]\s+(.*)', line)
        if m:
            list_items = []
            indent = len(m.group(1))
            while i < len(lines):
                lm = re.match(r'^(\s*)[-*+]\s+(.*)', lines[i])
                if not lm:
                    break
                if abs(len(lm.group(1)) - indent) > 2 and list_items:
                    break
                list_items.append(lm.group(2).strip())
                indent = len(lm.group(1))
                i += 1
            nodes.append({'type': 'list', 'children': list_items})
            continue

        # 普通段落
        nodes.append({'type': 'para', 'text': line.strip()})
        i += 1

    return nodes


# ============================================================
# DOCX 生成
# ============================================================

def extract_title_text(title_node):
    """从题目节点提取纯标题文字（去除'第一章'等前缀）"""
    text = title_node['text']
    # 去除中文序号前缀如 "第一章 "、"第1章 "
    text = re.sub(r'^第[一二三四五六七八九十\d]+章\s*', '', text)
    return text.strip()


def generate_docx(nodes, output_path, title_text):
    """根据节点列表生成 docx 文件"""
    doc = Document()

    # --- 默认样式 ---
    style = doc.styles['Normal']
    style.font.size = Pt(10.5)
    style.font.name = 'Times New Roman'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.paragraph_format.line_spacing = 1.3
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)

    # --- 节/页眉设置 ---
    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False
    # 左顶格 xxxxx
    hp_left = header.paragraphs[0]
    hp_left.paragraph_format.space_after = Pt(0)
    run_l = hp_left.add_run('xxxxx')
    set_run_font(run_l, '黑体', size_pt=9)
    # 右顶格 题目
    hp_right = header.add_paragraph()
    hp_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hp_right.paragraph_format.space_before = Pt(0)
    run_r = hp_right.add_run(title_text)
    set_run_font(run_r, '黑体', size_pt=9)

    # --- 页脚/页码设置 ---
    section.footer_distance = Cm(1.1)
    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # "第"
    r1 = fp.add_run('第')
    set_run_font(r1, '宋体', size_pt=10.5)
    # PAGE 字段
    r_page = fp.add_run()
    fc1 = OxmlElement('w:fldChar'); fc1.set(qn('w:fldCharType'), 'begin')
    r_page._r.append(fc1)
    it1 = OxmlElement('w:instrText'); it1.set(qn('xml:space'), 'preserve'); it1.text = ' PAGE '
    r_page._r.append(it1)
    fc2 = OxmlElement('w:fldChar'); fc2.set(qn('w:fldCharType'), 'end')
    r_page._r.append(fc2)
    # "页  共"
    r2 = fp.add_run('页  共')
    set_run_font(r2, '宋体', size_pt=10.5)
    # NUMPAGES 字段
    r_total = fp.add_run()
    fc3 = OxmlElement('w:fldChar'); fc3.set(qn('w:fldCharType'), 'begin')
    r_total._r.append(fc3)
    it2 = OxmlElement('w:instrText'); it2.set(qn('xml:space'), 'preserve'); it2.text = ' NUMPAGES '
    r_total._r.append(it2)
    fc4 = OxmlElement('w:fldChar'); fc4.set(qn('w:fldCharType'), 'end')
    r_total._r.append(fc4)
    # "页"
    r3 = fp.add_run('页')
    set_run_font(r3, '宋体', size_pt=10.5)

    # --- 计数器 ---
    chapter_path = [1]  # 一级标题序号
    fig_counter = {}  # key: chapter index tuple, value: counter
    tab_counter = {}

    def get_chapter_key():
        return tuple(chapter_path[:1])  # 只用一级序号作为章key

    def incr_fig():
        key = get_chapter_key()
        fig_counter[key] = fig_counter.get(key, 0) + 1
        return fig_counter[key]

    def incr_tab():
        key = get_chapter_key()
        tab_counter[key] = tab_counter.get(key, 0) + 1
        return tab_counter[key]

    def make_fig_label():
        return f"图{chapter_path[0]}-{incr_fig()}"

    def make_tab_label():
        return f"表{chapter_path[0]}-{incr_tab()}"

    # --- 遍历节点生成内容 ---
    for node in nodes:
        t = node['type']

        if t == 'title':
            add_empty_para(doc)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(node['text'])
            set_run_font(run, '黑体', size_pt=16)
            add_empty_para(doc)

        elif t == 'heading':
            lv = node['level']
            if lv == 1:
                # 一级标题：新页 + 三号黑体居中
                chapter_path[0] += 1
                chapter_path[1:] = [0]  # reset sub-levels

                add_empty_para(doc)
                p = doc.add_paragraph()
                add_page_break_before(p)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_outline_level(p, 1)
                run = p.add_run(node['text'])
                set_run_font(run, '黑体', size_pt=16)
                add_empty_para(doc)

            elif lv == 2:
                # 二级标题：四号黑体顶格不加粗，上下不空行
                p = doc.add_paragraph()
                set_outline_level(p, 2)
                p.paragraph_format.first_line_indent = Pt(0)
                run = p.add_run(node['text'])
                set_run_font(run, '黑体', size_pt=14, bold=False)

            elif lv >= 3:
                # 三级标题：小四宋体不加粗，首行缩进同正文，上下不空行
                p = doc.add_paragraph()
                set_outline_level(p, min(lv, 3))
                p.paragraph_format.first_line_indent = Pt(21)
                run = p.add_run(node['text'])
                set_run_font(run, '宋体', size_pt=12, bold=False)

        elif t == 'para':
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Pt(21)
            p.paragraph_format.line_spacing = 1.3
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(node['text'])
            set_run_font(run, '宋体', size_pt=10.5)

        elif t == 'image':
            url = node['url']
            alt = node['alt']
            img_path = download_image(url)
            if img_path is None:
                img_path = add_placeholder_image(doc, alt or url)

            w, h = calc_image_size(img_path)

            add_empty_para(doc)
            # 图片居中
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_img = p_img.add_run()
            run_img.add_picture(img_path, width=w)

            # 图题：小五宋体加粗居中
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_cap = p_cap.add_run(f'{make_fig_label()} {alt}')
            set_run_font(run_cap, '宋体', size_pt=9, bold=True)
            add_empty_para(doc)

            # 清理临时图片
            try:
                os.unlink(img_path)
            except Exception:
                pass

        elif t == 'table':
            header_cells = node['header']
            rows = node['rows']
            ncols = len(header_cells)

            # 检查前一个节点是否为表题段落
            # （在 parse 阶段已处理，这里简化：表题由前一个 para 节点承担，此处检测）
            # 实际简化处理：如果有 header text，自动加表题

            # 表题
            add_empty_para(doc)
            p_tab_cap = doc.add_paragraph()
            p_tab_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_tc = p_tab_cap.add_run(f'{make_tab_label()} {header_cells[0] if header_cells else ""}')
            set_run_font(run_tc, '宋体', size_pt=10.5, bold=True)

            # 表格
            table = doc.add_table(rows=len(rows) + 1, cols=ncols)
            table.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_three_line_table(table)

            # 表头行设为重复标题行（跨页自动出现）
            tblHeader = OxmlElement('w:tblHeader')
            trPr = table.rows[0]._tr.get_or_add_trPr()
            trPr.append(tblHeader)

            # 表头行
            for j, cell_text in enumerate(header_cells):
                cell = table.rows[0].cells[j]
                cell.paragraphs[0].clear()
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = cell.paragraphs[0].add_run(cell_text)
                set_run_font(run, '宋体', size_pt=9)
                cell.paragraphs[0].paragraph_format.first_line_indent = Pt(0)

            # 数据行
            for i, row_data in enumerate(rows):
                for j, cell_text in enumerate(row_data):
                    cell = table.rows[i + 1].cells[j]
                    cell.paragraphs[0].clear()
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    run = cell.paragraphs[0].add_run(cell_text)
                    set_run_font(run, '宋体', size_pt=9)
                    cell.paragraphs[0].paragraph_format.first_line_indent = Pt(0)

            add_empty_para(doc)

        elif t == 'code':
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(1)
            run = p.add_run(node['text'])
            set_run_font(run, '宋体', en_font='Courier New', size_pt=10.5)

        elif t == 'list':
            for idx, item in enumerate(node['children'], 1):
                p = doc.add_paragraph()
                p.paragraph_format.first_line_indent = Pt(21)
                p.paragraph_format.line_spacing = 1.3
                run = p.add_run(f'（{idx}）{item}')
                set_run_font(run, '宋体', size_pt=10.5)

        elif t == 'empty':
            pass  # 段落间不空行，空行由标题/图表handler显式添加

    doc.save(output_path)
    return output_path


# ============================================================
# 主入口
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description='mddoc - Markdown 转学术格式 DOCX',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python md2docx.py paper.md                    # 转换文件，输出到同目录
  python md2docx.py paper.md -o output.docx     # 指定输出路径
  python md2docx.py --text "# 标题\\n\\n正文内容" -o out.docx  # 直接转换文本
        """.strip())
    parser.add_argument('input', nargs='?', help='输入 Markdown 文件路径')
    parser.add_argument('-o', '--output', help='输出 DOCX 文件路径')
    parser.add_argument('--text', help='直接传入 Markdown 文本（不需要文件）')
    args = parser.parse_args()

    if not args.input and not args.text:
        parser.print_help()
        sys.exit(1)

    # 读取内容
    if args.text:
        md_content = args.text
        source_dir = os.getcwd()
    else:
        with open(args.input, 'r', encoding='utf-8') as f:
            md_content = f.read()
        source_dir = os.path.dirname(os.path.abspath(args.input))

    # 解析
    nodes = parse_markdown(md_content)

    # 提取题目
    title_node = next((n for n in nodes if n['type'] == 'title'), None)
    if title_node is None:
        print("错误：未找到文档标题（第一个 # 标题）", file=sys.stderr)
        sys.exit(1)
    title_text = extract_title_text(title_node)

    # 确定输出路径
    if args.output:
        output_path = args.output
    else:
        safe_name = re.sub(r'[\\/*?:"<>|]', '', title_text)
        output_path = os.path.join(source_dir, f'{safe_name}.docx')

    # 生成
    generate_docx(nodes, output_path, title_text)
    print(f'已生成: {output_path}')


if __name__ == '__main__':
    main()
