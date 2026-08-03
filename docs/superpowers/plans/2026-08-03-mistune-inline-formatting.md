# Markdown 内联格式化与列表支持 — 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 用 mistune 3.x AST 替换手写 Markdown 解析器，新增分隔线分页、加粗、斜体、加粗斜体、Word 原生列表支持，覆盖 Word 内置标题样式。

**Architecture:** 删除 `parse_markdown()` 和 `split_inline_math()`（~190行），改用 `mistune.create_markdown(renderer='ast', plugins=[table.table, math.math])`。`generate_docx()` 改为 AST walker，新增 `walk_inline()` 递归渲染富文本 runs，新增 `render_list()` 渲染 Word 原生列表。

**Tech Stack:** Python 3, mistune 3.3.3, python-docx, lxml

## Global Constraints

- 公式渲染：`_LatexParser` / `latex_to_omml()` 零改动
- 三线表、页眉页脚、页码、图题/表题编号 不变
- mistune 插件：`table.table` + `math.math`
- 标题样式：生成前覆盖 Heading 1-6 为学术格式
- SKILL.md 同步更新

---

### Task 1: 新增 `_override_builtin_heading_styles()`

**Files:**
- Modify: `skills/mddoc/scripts/md2docx.py` (在 `set_outline_level` 之后插入)

**Interfaces:**
- Produces: `_override_builtin_heading_styles(doc)` — 无返回值，修改 doc 样式

- [ ] **Step 1: 在 `generate_docx` 开头调用，覆盖 Heading 1-6**

在 `set_outline_level()` 函数之后新增：

```python
def _override_builtin_heading_styles(doc):
    """覆盖 Word 内置 Heading 1-6 样式为学术格式，防止默认蓝色/Cambria/加粗干扰"""
    specs = {
        'Heading 1': ('黑体', 16, True, WD_ALIGN_PARAGRAPH.CENTER),
        'Heading 2': ('黑体', 14, False, WD_ALIGN_PARAGRAPH.LEFT),
        'Heading 3': ('宋体', 12, False, WD_ALIGN_PARAGRAPH.LEFT),
        'Heading 4': ('宋体', 10.5, False, WD_ALIGN_PARAGRAPH.LEFT),
        'Heading 5': ('宋体', 10.5, False, WD_ALIGN_PARAGRAPH.LEFT),
        'Heading 6': ('宋体', 10.5, False, WD_ALIGN_PARAGRAPH.LEFT),
    }
    for name, (cn_font, size_pt, bold, align) in specs.items():
        style = doc.styles[name]
        style.font.size = Pt(size_pt)
        style.font.name = 'Times New Roman'
        style.font.bold = bold
        style.font.color.rgb = None  # 黑色
        style.font.italic = False
        style.font.underline = False
        # 中文字体
        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:eastAsia'), cn_font)
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        # 段落间距：1.3倍行距，段前段后0
        pPr = style.element.get_or_add_pPr()
        spacing = pPr.find(qn('w:spacing'))
        if spacing is None:
            spacing = OxmlElement('w:spacing')
            pPr.append(spacing)
        spacing.set(qn('w:before'), '0')
        spacing.set(qn('w:after'), '0')
        spacing.set(qn('w:line'), '312')
        spacing.set(qn('w:lineRule'), 'auto')
```

- [ ] **Step 2: 在 `generate_docx()` 中 `doc = Document()` 后立即调用**

```python
def generate_docx(ast, output_path, title_text=None):
    doc = Document()
    _override_builtin_heading_styles(doc)  # ← 新增
    # ... 后续不变 ...
```

- [ ] **Step 3: 冒烟测试**

```bash
source .venv/bin/activate && python skills/mddoc/scripts/md2docx.py skills/mddoc/evals/test-sample.md -o /tmp/test.docx
```

---

### Task 2: 替换解析器 — `parse_markdown` → mistune

**Files:**
- Modify: `skills/mddoc/scripts/md2docx.py` (替换 `parse_markdown` 函数体)

**Interfaces:**
- Consumes: `latex_to_omml`, `split_inline_math` (即将删除)
- Produces: `parse_markdown(text)` → `list[dict]` (mistune AST)

- [ ] **Step 1: 配置 mistune 解析器**

将 `parse_markdown()` 替换为：

```python
def parse_markdown(text):
    """使用 mistune 解析 Markdown 文本为 AST 节点列表

    参数:
        text: Markdown 原始文本

    返回:
        mistune AST 节点列表，包含 heading/paragraph/list/table/block_math/
        block_code/thematic_break 等类型
    """
    import mistune
    from mistune.plugins import table, math

    md = mistune.create_markdown(
        renderer='ast',
        plugins=[table.table, math.math]
    )
    return md(text)
```

- [ ] **Step 2: 删除 `split_inline_math` 函数（~30行）**

inline 解析由 `walk_inline()` 替代（下一任务）。

- [ ] **Step 3: 验证解析输出**

```bash
source .venv/bin/activate && python -c "
from md2docx import parse_markdown
import json
ast = parse_markdown('**bold** *italic* \$x^2\$')
print(json.dumps(ast, indent=2, ensure_ascii=False))
"
```

---

### Task 3: 新增 `walk_inline()` 内联渲染

**Files:**
- Modify: `skills/mddoc/scripts/md2docx.py` (在 `generate_docx` 之前插入)

**Interfaces:**
- Produces: `walk_inline(paragraph, children, base_font_cn='宋体', base_font_en='Times New Roman', base_size=10.5)` — 递归遍历 inline AST，追加格式化的 run/OMML 到段落

- [ ] **Step 1: 实现 `walk_inline`**

```python
def walk_inline(paragraph, children, base_font_cn='宋体', base_font_en='Times New Roman', base_size=10.5):
    """递归遍历 inline AST children，将格式化文本和 OMML 附加到段落

    参数:
        paragraph: python-docx Paragraph 对象
        children: mistune inline AST 子节点列表
        base_font_cn: 中文字体名
        base_font_en: 英文字体名
        base_size: 基础字号，单位 pt
    """
    for child in children:
        ct = child['type']

        if ct == 'text':
            raw = child.get('raw', '')
            if raw:
                run = paragraph.add_run(raw)
                set_run_font(run, base_font_cn, en_font=base_font_en, size_pt=base_size)

        elif ct == 'strong':
            # 先用临时段落收集内部 runs，再复制到目标段落（含 bold）
            tmp_para = _make_temp_para()
            walk_inline(tmp_para, child.get('children', []), base_font_cn, base_font_en, base_size)
            for r in tmp_para.runs:
                is_italic = r.font.italic
                new_run = paragraph.add_run(r.text)
                set_run_font(new_run, base_font_cn, en_font=base_font_en, size_pt=base_size, bold=True)
                if is_italic:
                    new_run.font.italic = True

        elif ct == 'emphasis':
            tmp_para = _make_temp_para()
            walk_inline(tmp_para, child.get('children', []), base_font_cn, base_font_en, base_size)
            for r in tmp_para.runs:
                is_bold = r.font.bold
                new_run = paragraph.add_run(r.text)
                set_run_font(new_run, base_font_cn, en_font=base_font_en, size_pt=base_size, bold=is_bold)
                new_run.font.italic = True

        elif ct == 'inline_math':
            try:
                omml = latex_to_omml(child['raw'], display=False)
                paragraph._element.append(omml)
            except Exception:
                run = paragraph.add_run(child['raw'])
                set_run_font(run, base_font_cn, en_font=base_font_en, size_pt=base_size)
                run.font.italic = True

        elif ct == 'codespan':
            run = paragraph.add_run(child.get('raw', ''))
            set_run_font(run, base_font_cn, en_font=base_font_en, size_pt=base_size)

        elif ct == 'link':
            tmp_para = _make_temp_para()
            walk_inline(tmp_para, child.get('children', []), base_font_cn, base_font_en, base_size)
            for r in tmp_para.runs:
                new_run = paragraph.add_run(r.text)
                set_run_font(new_run, base_font_cn, en_font=base_font_en, size_pt=base_size)
                # 蓝色下划线（学术格式通常不需要，但保留链接可识别性）
                new_run.font.color.rgb = None  # 黑色
                new_run.font.underline = True

        elif ct == 'image':
            url = child.get('attrs', {}).get('url', child.get('raw', ''))
            alt = child.get('alt', '')
            img_path = download_image(url)
            if img_path is None:
                img_path = add_placeholder_image(doc_ref, alt or url)
            if img_path:
                w, h = calc_image_size(img_path)
                run = paragraph.add_run()
                run.add_picture(img_path, width=w)
                try:
                    os.unlink(img_path)
                except Exception:
                    pass

        elif ct == 'linebreak':
            run = paragraph.add_run()
            run.add_break()
```

- [ ] **Step 2: 实现辅助函数 `_make_temp_para()`**

```python
def _make_temp_para():
    """创建临时 Document + Paragraph，用于收集嵌套 inline 的 runs 序列"""
    from docx import Document as DocCls
    tmp_doc = DocCls()
    return tmp_doc.add_paragraph()
```

- [ ] **Step 3: 验证 walk_inline**

```bash
source .venv/bin/activate && python -c "
from docx import Document
from md2docx import walk_inline
doc = Document()
p = doc.add_paragraph()
walk_inline(p, [
    {'type': 'strong', 'children': [{'type': 'text', 'raw': 'bold'}]},
    {'type': 'text', 'raw': ' and '},
    {'type': 'emphasis', 'children': [{'type': 'text', 'raw': 'italic'}]},
])
assert len(p.runs) == 3
assert p.runs[0].font.bold == True
assert p.runs[2].font.italic == True
print('walk_inline OK')
"
```

---

### Task 4: 重构 `generate_docx()` — AST Walker

**Files:**
- Modify: `skills/mddoc/scripts/md2docx.py` (重写 `generate_docx` 主循环)

**Interfaces:**
- Consumes: `walk_inline()`, `_override_builtin_heading_styles()`, `_make_temp_para()`
- Produces: 完整 DOCX 文件

- [ ] **Step 1: 新增 `render_list()`**

```python
def render_list(doc, node):
    """渲染 Word 原生列表（有序/无序），列表项支持富文本 inline

    参数:
        doc: python-docx Document 对象
        node: mistune list AST 节点
    """
    ordered = node.get('attrs', {}).get('ordered', False) if 'attrs' in node else node.get('ordered', False)

    for item in node.get('children', []):
        if item['type'] != 'list_item':
            continue

        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.3
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

        # 悬挂缩进：项目符号/编号悬出，文本缩进对齐
        p.paragraph_format.left_indent = Cm(0.63)
        p.paragraph_format.first_line_indent = Cm(-0.63)

        # Word 原生列表编号/项目符号
        _set_list_num_pr(p, ordered=ordered, level=0)

        # 渲染列表项内富文本
        for child in item.get('children', []):
            if child['type'] == 'block_text':
                walk_inline(p, child.get('children', []))
            elif child['type'] == 'paragraph':
                walk_inline(p, child.get('children', []))
```

- [ ] **Step 2: 新增 `_set_list_num_pr()` — 列表编号 XML**

```python
def _set_list_num_pr(paragraph, ordered=False, level=0):
    """设置段落的 Word 原生列表编号属性 (w:numPr)

    参数:
        paragraph: python-docx Paragraph 对象
        ordered: True=编号列表(1,2,3...), False=项目符号(•)
        level: 列表嵌套层级（0=顶层）
    """
    # 使用 Word 内置抽象编号定义
    # 有序：numId=1 映射到内置编号；无序：numId=2 映射到内置项目符号
    num_id = 1 if ordered else 2
    pPr = paragraph._element.get_or_add_pPr()
    numPr = OxmlElement('w:numPr')
    ilvl = OxmlElement('w:ilvl')
    ilvl.set(qn('w:val'), str(level))
    numPr.append(ilvl)
    numId = OxmlElement('w:numId')
    numId.set(qn('w:val'), str(num_id))
    numPr.append(numId)
    pPr.append(numPr)
```

- [ ] **Step 3: 新增 `_ensure_list_numbering_defs()` — 在 document 中定义列表编号**

```python
def _ensure_list_numbering_defs(doc):
    """确保文档的 numbering.xml 中有有序和无序列表的抽象编号定义。
    python-docx 默认不含列表定义，需手动在 XML 级别注入。

    参数:
        doc: python-docx Document 对象
    """
    numbering_part = doc.part.numbering_part
    numbering_elem = numbering_part._element

    # 检查是否已存在我们的 numId
    for num in numbering_elem.findall(qn('w:num')):
        nid = num.get(qn('w:numId'))
        if nid in ('1', '2'):
            return  # 已定义

    # 抽象编号 0：有序列表 (1, 2, 3...)
    abstract_num_0 = OxmlElement('w:abstractNum')
    abstract_num_0.set(qn('w:abstractNumId'), '0')
    lvl_ord = OxmlElement('w:lvl')
    lvl_ord.set(qn('w:ilvl'), '0')
    start_ord = OxmlElement('w:start')
    start_ord.set(qn('w:val'), '1')
    lvl_ord.append(start_ord)
    numFmt_ord = OxmlElement('w:numFmt')
    numFmt_ord.set(qn('w:val'), 'decimal')
    lvl_ord.append(numFmt_ord)
    lvlText_ord = OxmlElement('w:lvlText')
    lvlText_ord.set(qn('w:val'), '%1.')
    lvl_ord.append(lvlText_ord)
    lvlJc_ord = OxmlElement('w:lvlJc')
    lvlJc_ord.set(qn('w:val'), 'left')
    lvl_ord.append(lvlJc_ord)
    # 缩进：悬挂 0.63cm
    pPr_ord = OxmlElement('w:pPr')
    ind_ord = OxmlElement('w:ind')
    ind_ord.set(qn('w:left'), '360')   # 0.63cm in twips
    ind_ord.set(qn('w:hanging'), '360')
    pPr_ord.append(ind_ord)
    lvl_ord.append(pPr_ord)
    abstract_num_0.append(lvl_ord)

    # 抽象编号 1：无序列表 (• bullet)
    abstract_num_1 = OxmlElement('w:abstractNum')
    abstract_num_1.set(qn('w:abstractNumId'), '1')
    lvl_unord = OxmlElement('w:lvl')
    lvl_unord.set(qn('w:ilvl'), '0')
    start_unord = OxmlElement('w:start')
    start_unord.set(qn('w:val'), '1')
    lvl_unord.append(start_unord)
    numFmt_unord = OxmlElement('w:numFmt')
    numFmt_unord.set(qn('w:val'), 'bullet')
    lvl_unord.append(numFmt_unord)
    lvlText_unord = OxmlElement('w:lvlText')
    lvlText_unord.set(qn('w:val'), '•')  # • (U+2022)
    lvl_unord.append(lvlText_unord)
    lvlJc_unord = OxmlElement('w:lvlJc')
    lvlJc_unord.set(qn('w:val'), 'left')
    lvl_unord.append(lvlJc_unord)
    pPr_unord = OxmlElement('w:pPr')
    ind_unord = OxmlElement('w:ind')
    ind_unord.set(qn('w:left'), '360')
    ind_unord.set(qn('w:hanging'), '360')
    pPr_unord.append(ind_unord)
    lvl_unord.append(pPr_unord)
    abstract_num_1.append(lvl_unord)

    numbering_elem.append(abstract_num_0)
    numbering_elem.append(abstract_num_1)

    # num 实例 1：有序列表
    num_ord = OxmlElement('w:num')
    num_ord.set(qn('w:numId'), '1')
    absRef_ord = OxmlElement('w:abstractNumId')
    absRef_ord.set(qn('w:val'), '0')
    num_ord.append(absRef_ord)
    numbering_elem.append(num_ord)

    # num 实例 2：无序列表
    num_unord = OxmlElement('w:num')
    num_unord.set(qn('w:numId'), '2')
    absRef_unord = OxmlElement('w:abstractNumId')
    absRef_unord.set(qn('w:val'), '1')
    num_unord.append(absRef_unord)
    numbering_elem.append(num_unord)
```

- [ ] **Step 4: 重构主循环 — 按 mistune AST type 分发**

将现有 `for node in nodes:` 循环替换为 mistune AST walker：

```python
def generate_docx(ast, output_path, title_text=None):
    doc = Document()
    _override_builtin_heading_styles(doc)

    # --- 默认样式（不变）---
    # ... Normal style, margins, header, footer ...

    # 列表编号定义（在生成列表前注入）
    _ensure_list_numbering_defs(doc)

    # 标题计数器
    first_heading = True
    heading_count = 0
    chapter_path = [1]
    has_chapter = False
    fig_counter, tab_counter, eq_counter = {}, {}, {}
    prev_node = None  # 表题检测

    for node in ast:
        t = node['type']

        if t == 'heading':
            heading_count += 1
            level = node.get('attrs', {}).get('level', 1)
            text = _extract_ast_text(node.get('children', []))

            if level == 1 and heading_count == 1:
                # 题目
                _render_title(doc, text)
            else:
                if level == 1:
                    has_chapter = True
                    chapter_path[0] += 1

                if level == 1:
                    add_empty_para(doc)
                    p = doc.add_paragraph()
                    add_page_break_before(p)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_outline_level(p, 1)
                    run = p.add_run(text)
                    set_run_font(run, '黑体', size_pt=16)
                    add_empty_para(doc)
                elif level == 2:
                    p = doc.add_paragraph()
                    set_outline_level(p, 2)
                    p.paragraph_format.first_line_indent = Pt(0)
                    run = p.add_run(text)
                    set_run_font(run, '黑体', size_pt=14, bold=False)
                elif level == 3:
                    p = doc.add_paragraph()
                    set_outline_level(p, 3)
                    p.paragraph_format.first_line_indent = Pt(21)
                    run = p.add_run(text)
                    set_run_font(run, '宋体', size_pt=12, bold=False)
                else:  # level >= 4
                    p = doc.add_paragraph()
                    set_outline_level(p, level)
                    p.paragraph_format.first_line_indent = Pt(21)
                    p.paragraph_format.line_spacing = 1.3
                    run = p.add_run(text)
                    set_run_font(run, '宋体', size_pt=10.5, bold=False)

        elif t == 'paragraph':
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Pt(21)
            p.paragraph_format.line_spacing = 1.3
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            walk_inline(p, node.get('children', []))
            prev_node = node

        elif t == 'thematic_break':
            p = doc.add_paragraph()
            run = p.add_run('')
            set_run_font(run, '宋体', size_pt=10.5)
            add_page_break_before(p)

        elif t == 'list':
            render_list(doc, node)

        elif t == 'block_math':
            # 行间公式（复用现有逻辑）
            _render_display_math(doc, node['raw'], chapter_path, eq_counter, has_chapter)

        elif t == 'block_code':
            add_empty_para(doc)
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(1)
            pPr = p._element.get_or_add_pPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), 'D9D9D9')
            pPr.append(shd)
            run = p.add_run(node.get('raw', ''))
            set_run_font(run, '宋体', en_font='Times New Roman', size_pt=10.5)
            add_empty_para(doc)

        elif t == 'table':
            caption = _extract_table_caption_from_prev(prev_node)
            _render_table(doc, node, caption)  # 复用现有三线表逻辑
            prev_node = None  # 表题已消费

        elif t == 'blank_line':
            pass

    doc.save(output_path)
    return output_path
```

- [ ] **Step 5: 辅助函数**

```python
def _extract_ast_text(children):
    """从 mistune inline AST children 提取纯文本"""
    parts = []
    for c in children:
        if c['type'] == 'text':
            parts.append(c.get('raw', ''))
        elif 'children' in c:
            parts.append(_extract_ast_text(c['children']))
    return ''.join(parts)


def _extract_table_caption_from_prev(prev_node):
    """从表格前一个段落节点提取表题文本"""
    if prev_node and prev_node['type'] == 'paragraph':
        text = _extract_ast_text(prev_node.get('children', []))
        if '表' in text:
            return text.strip()
    return ''
```

- [ ] **Step 6: 提取 `_render_title` 和 `_render_table` 为独立函数（从原 generate_docx 中拆出）**

`_render_title`：复用原题目渲染逻辑（16pt 黑体居中，上下各空一行）

`_render_table`：复用原三线表渲染逻辑，适配 mistune table AST：
```python
def _render_table(doc, node, caption):
    """渲染三线表，适配 mistune table AST

    参数:
        doc: Document 对象
        node: mistune table AST 节点
        caption: 表题文本（可为空）
    """
    children = node.get('children', [])
    # 提取 table_head 和 table_body
    head_rows = []
    body_rows = []
    for c in children:
        if c['type'] == 'table_head':
            for row in c.get('children', []):
                head_rows.append(row)
        elif c['type'] == 'table_body':
            for row in c.get('children', []):
                body_rows.append(row)

    # ... 三线表渲染（同现有逻辑）...
```

- [ ] **Step 7: 提取 `_render_display_math` 为独立函数**

复用原行间公式渲染逻辑（上下空行、OMLL 居中、编号右对齐）。

---

### Task 5: 更新 SKILL.md

**Files:**
- Modify: `skills/mddoc/SKILL.md`

- [ ] **Step 1: 更新列表代码示例**

删除手动 `（1）` `（2）` 序号和 `•` 符号的做法，改为 Word 原生列表 + `walk_inline` 风格。

- [ ] **Step 2: 更新 inline 格式示例**

新增 bold/italic/inline_math 的 `walk_inline` 用法。

- [ ] **Step 3: 新增标题样式覆盖说明**

添加上 `_override_builtin_heading_styles()` 的用法。

---

### Task 6: 冒烟测试 + 验证

**Files:**
- 测试: `skills/mddoc/evals/test-sample.md` (输入)
- 输出: `/tmp/test.docx` (验证)

- [ ] **Step 1: 创建富文本测试输入**

```bash
source .venv/bin/activate && python skills/mddoc/scripts/md2docx.py skills/mddoc/evals/test-sample.md -o /tmp/test.docx
```

- [ ] **Step 2: 验证 docx 内容**

```bash
source .venv/bin/activate && python -c "
from docx import Document
doc = Document('/tmp/test.docx')
for p in doc.paragraphs:
    ol = p._element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}outlineLvl')
    lvl = ol.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if ol is not None else '-'
    text = p.text[:100] if p.text else '(empty)'
    print(f'[{lvl}] {text}')
"
```

- [ ] **Step 3: 验证要点**
  - outline_level 正确
  - 图片嵌入、页眉、页码 正确
  - 无回归错误

- [ ] **Step 4: 提交**

```bash
git add skills/mddoc/scripts/md2docx.py skills/mddoc/SKILL.md
git commit -m "feat: mistune AST parser, inline formatting, native lists, heading style override"
```
