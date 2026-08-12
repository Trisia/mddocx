# 表格单元格内联格式渲染 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 表格单元格支持斜体、加粗、行内 LaTeX 公式（含嵌套组合如 `**$x$**`），并修复正文 `walk_inline` 嵌套公式丢失的同一缺陷；单元格垂直居中、公式字号跟随 9pt。

**Architecture:** 把 `walk_inline` 从"临时段落拷贝 runs"重构为"格式状态递归"（bold/italic/underline 下传，同时作用于文本 run 和 OMML `m:sty`）。表格 `_render_table` 改存 children AST 并调用 `walk_inline`，新增 `_apply_omml_style`（OMML 上样式）和 `_set_para_mark_size`（段落标记字号，供公式继承）。SKILL.md 同步，Agent 生成代码一致。

**Tech Stack:** Python 3、mistune 3.x（AST）、python-docx、lxml、OMML XML。

**设计文档:** `docs/superpowers/specs/2026-08-12-table-cell-inline-formatting-design.md`

## Global Constraints

- 所有命令先 `cd /home/kkk/project/mddocx && source .venv/bin/activate`
- 禁止函数内 import；所有 import 在文件头部完成（`md2docx.py:14-31` 区域追加）
- 公开函数写注释：说明、参数、返回值；字号参数注明单位 pt
- 避免过度封装，不新增只调用一次的无意义包装函数
- 禁止 `git push`（需用户明确同意）
- 修改格式规范时 `SKILL.md` 与 `md2docx.py` 必须同步
- 本计划不 bump 版本号

---

### Task 1: 重构 `walk_inline` 为格式状态递归 + 新增 `_apply_omml_style`

**Files:**
- Modify: `skills/mddoc/scripts/md2docx.py:1375-1458`（`_make_temp_para` + `walk_inline`）

**Interfaces:**
- Consumes: `latex_to_omml`（`md2docx.py:1038`）、`set_run_font`（`md2docx.py:1072`）、`_m_elem`/`_m_qn`（`md2docx.py:124/41`）、`NSM` 命名空间常量
- Produces:
  - `walk_inline(paragraph, children, base_cn='宋体', base_en='Times New Roman', base_size=10.5, bold=False, italic=False, underline=False)` — 向下传格式状态
  - `_apply_omml_style(omml, bold=False, italic=False) -> omml` — 为 OMML 内所有 `m:r` 加 `<m:sty m:val="b|i|bi"/>`，原地修改返回

- [ ] **Step 1: 写失败验证脚本**

创建 `/tmp/verify_walk_inline.py`：

```python
import mistune
from mistune.plugins import table as mistune_table, math as mistune_math
from docx import Document
import sys
sys.path.insert(0, '/home/kkk/project/mddocx/skills/mddoc/scripts')
import md2docx

md = mistune.create_markdown(renderer='ast', plugins=[mistune_table.table, mistune_math.math])
inline = md('x **bold** *italic* $x^2$ **$y^2$**')[0]['children']
doc = Document()
p = doc.add_paragraph()
md2docx.walk_inline(p, inline)
xml = p._element.xml
assert '<w:b/>' in xml, '加粗 run 缺失'
assert '<w:i/>' in xml, '斜体 run 缺失'
assert 'oMath' in xml, '行内公式缺失'
assert 'm:val="b"' in xml, '嵌套加粗公式样式缺失 (**$y^2$**)'
print('OK')
```

- [ ] **Step 2: 运行确认失败**

Run: `cd /home/kkk/project/mddocx && source .venv/bin/activate && python /tmp/verify_walk_inline.py`
Expected: `AssertionError: 嵌套加粗公式样式缺失 (**$y^2$**)`（当前 strong 分支只拷 runs，丢 OMML）

- [ ] **Step 3: 实现**

删除 `_make_temp_para`（`md2docx.py:1375-1378`），整体替换 `walk_inline`（`md2docx.py:1381-1458`）：

```python
def _apply_omml_style(omml, bold=False, italic=False):
    """为 OMML 内所有 m:r 应用加粗/斜体样式（m:sty）

    参数：
        omml: m:oMath 元素
        bold: 是否加粗
        italic: 是否斜体（默认数学即斜体，仅强加时设置）

    返回:
        应用样式后的 omml（原地修改）
    """
    if not (bold or italic):
        return omml
    val = ('b' if bold else '') + ('i' if italic else '')
    for r in omml.iter(f'{{{NSM}}}r'):
        rPr = r.find(f'{{{NSM}}}rPr')
        if rPr is None:
            rPr = _m_elem('rPr')
            r.insert(0, rPr)
        if rPr.find(f'{{{NSM}}}sty') is None:
            sty = _m_elem('sty')
            sty.set(_m_qn('val'), val)
            rPr.append(sty)
    return omml


def walk_inline(paragraph, children, base_cn='宋体', base_en='Times New Roman',
                base_size=10.5, bold=False, italic=False, underline=False):
    """递归遍历 inline AST，将格式化文本和 OMML 公式附加到段落

    参数：
        paragraph: python-docx Paragraph 对象
        children: mistune inline AST 子节点列表
        base_cn: 中文字体名（默认 '宋体'）
        base_en: 英文字体名（默认 'Times New Roman'）
        base_size: 基础字号，单位 pt（默认 10.5）
        bold: 加粗状态（嵌套 strong 递归传 True）
        italic: 斜体状态（嵌套 emphasis 递归传 True）
        underline: 下划线状态（嵌套 link 递归传 True）
    """
    for child in children:
        ct = child['type']

        if ct == 'text':
            raw = child.get('raw', '')
            if raw:
                run = paragraph.add_run(raw)
                set_run_font(run, base_cn, en_font=base_en, size_pt=base_size, bold=bold)
                if italic:
                    run.font.italic = True
                if underline:
                    run.font.underline = True

        elif ct == 'strong':
            walk_inline(paragraph, child.get('children', []), base_cn, base_en, base_size,
                        bold=True, italic=italic, underline=underline)

        elif ct == 'emphasis':
            walk_inline(paragraph, child.get('children', []), base_cn, base_en, base_size,
                        bold=bold, italic=True, underline=underline)

        elif ct == 'inline_math':
            try:
                omml = latex_to_omml(child['raw'], display=False)
                omml = _apply_omml_style(omml, bold, italic)
                paragraph._element.append(omml)
            except Exception:
                run = paragraph.add_run(child.get('raw', ''))
                set_run_font(run, base_cn, en_font=base_en, size_pt=base_size, bold=bold)
                run.font.italic = True

        elif ct == 'codespan':
            run = paragraph.add_run(child.get('raw', ''))
            set_run_font(run, base_cn, en_font=base_en, size_pt=base_size)

        elif ct == 'link':
            walk_inline(paragraph, child.get('children', []), base_cn, base_en, base_size,
                        bold=bold, italic=italic, underline=True)

        elif ct == 'image':
            url = child.get('attrs', {}).get('url', '')
            alt = child.get('alt', '')
            img_path = download_image(url)
            if img_path is None:
                img_path = add_placeholder_image(None, alt or url)
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

- [ ] **Step 4: 运行确认通过**

Run: `cd /home/kkk/project/mddocx && source .venv/bin/activate && python /tmp/verify_walk_inline.py`
Expected: `OK`

- [ ] **Step 5: 回归现有正文/列表**

Run: `cd /home/kkk/project/mddocx && source .venv/bin/activate && python skills/mddoc/scripts/md2docx.py skills/mddoc/evals/test-sample.md -o /tmp/test.docx && echo CONVERT_OK`
Expected: `CONVERT_OK`（正文/列表正常文本输出不变，仅修复嵌套公式丢失）

- [ ] **Step 6: Commit**

```bash
git add skills/mddoc/scripts/md2docx.py
git commit -m "refactor: walk_inline 改为格式状态递归，支持嵌套加粗/斜体公式"
```

---

### Task 2: 新增 `_set_para_mark_size` + 改造 `_render_table`

**Files:**
- Modify: `skills/mddoc/scripts/md2docx.py`（import 区、`set_run_font` 后、`_render_table` 1613-1696）

**Interfaces:**
- Consumes: `walk_inline`（Task 1）、`_extract_ast_text`（`md2docx.py:1357`）、`WD_ALIGN_PARAGRAPH`
- Produces:
  - `_set_para_mark_size(paragraph, size_pt)` — 设置段落标记 `w:rPr/w:sz`（半磅），供 OMML 公式继承字号
  - `_render_table` 内部：`head_cells`/`body_rows` 元素类型从 `str` 改为 **children AST 列表**

- [ ] **Step 1: 写失败验证脚本**

创建 `/tmp/tbl_test.md`：

```markdown
表 参数表

| 名称 | 公式 |
|------|------|
| **批量** | *32* |
| $x_i$ | **$y^2$** |

| 甲 | 乙 |
|----|----|
| 1 | 2 |
```

创建 `/tmp/verify_table.py`：

```python
from docx import Document
doc = Document('/tmp/tbl_test.docx')
table = doc.tables[0]
t0 = table.rows[1].cells[0]._tc.xml   # **批量**
t1 = table.rows[1].cells[1]._tc.xml   # *32*
m0 = table.rows[2].cells[0]._tc.xml   # $x_i$
m1 = table.rows[2].cells[1]._tc.xml   # **$y^2$**
assert '<w:b/>' in t0, '单元格加粗缺失'
assert '<w:i/>' in t1, '单元格斜体缺失'
assert 'oMath' in m0, '单元格公式缺失'
assert 'oMath' in m1 and 'm:val="b"' in m1, '单元格加粗公式缺失'
hdr = table.rows[0].cells[0]._tc.xml
assert 'vAlign' in hdr and 'center' in hdr, '表头垂直居中缺失'
pm = table.rows[1].cells[0].paragraphs[0]._p.xml
assert 'w:val="18"' in pm, '单元格段落标记字号未设 9pt'
# 表题回退：第二个表无"表 xxx"段，回退到第一列表头
cap = [p.text for p in doc.paragraphs if p.text.startswith('表')]
assert cap and '甲' in cap[-1], '表题回退失败'
print('OK')
```

- [ ] **Step 2: 运行确认失败**

Run: `cd /home/kkk/project/mddocx && source .venv/bin/activate && python skills/mddoc/scripts/md2docx.py /tmp/tbl_test.md -o /tmp/tbl_test.docx && python /tmp/verify_table.py`
Expected: `AssertionError: 单元格公式缺失`（当前 `_extract_ast_text` 丢弃公式）；表题回退因 `head_cells[0]` 类型问题可能再报错，同样属预期失败

- [ ] **Step 3: 实现**

3a. 头部 import 追加（`md2docx.py:27` 后）：

```python
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
```

3b. `set_run_font` 后新增 `_set_para_mark_size`（`md2docx.py:1085` 后）：

```python
def _set_para_mark_size(paragraph, size_pt):
    """设置段落标记字号（半磅），供行内公式 OMML 继承

    参数：
        paragraph: python-docx Paragraph 对象
        size_pt: 字号，单位 pt（如 9 → 9pt）
    """
    pPr = paragraph._element.get_or_add_pPr()
    rPr = pPr.get_or_add_rPr()
    sz = rPr.find(qn('w:sz'))
    if sz is None:
        sz = OxmlElement('w:sz')
        rPr.append(sz)
    sz.set(qn('w:val'), str(int(size_pt * 2)))
```

3c. 重写 `_render_table`（`md2docx.py:1613-1696`）头部数据收集与渲染：

```python
def _render_table(doc, node, tab_counter, chapter_path, has_chapter, caption=''):
    """渲染三线表，适配 mistune table AST。

    参数：
        doc: Document 对象
        node: mistune table AST 节点
        tab_counter: 表格计数 dict
        chapter_path: 章节路径 list
        has_chapter: 是否有章标题
        caption: 表题文本（来自"表 xxx"段落缓冲，或为空回退到 header_cells[0]）
    """
    children = node.get('children', [])
    head_cells = []
    body_rows = []

    for c in children:
        if c['type'] == 'table_head':
            # table_head 的子节点直接是 table_cell（无 row 包装）
            for cell in c.get('children', []):
                if cell['type'] == 'table_cell':
                    head_cells.append(cell.get('children', []))
        elif c['type'] == 'table_body':
            # table_body > table_row > table_cell
            for row in c.get('children', []):
                if row['type'] != 'table_row':
                    continue
                row_data = []
                for cell in row.get('children', []):
                    if cell['type'] == 'table_cell':
                        row_data.append(cell.get('children', []))
                body_rows.append(row_data)

    if not head_cells:
        return

    ncols = len(head_cells)

    # 表题文本：优先使用传入的 caption，否则回退到第一列表头
    if not caption:
        caption = _extract_ast_text(head_cells[0]) if head_cells else ''
    tab_key = tuple(chapter_path[:1]) if has_chapter else None
    if tab_key:
        tab_counter[tab_key] = tab_counter.get(tab_key, 0) + 1
        tab_num = tab_counter[tab_key]
    else:
        tab_counter[None] = tab_counter.get(None, 0) + 1
        tab_num = tab_counter[None]

    tab_label = f"表{chapter_path[0]}-{tab_num}" if has_chapter else f"表{tab_num}"

    add_empty_para(doc)
    p_tab_cap = doc.add_paragraph()
    p_tab_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_tc = p_tab_cap.add_run(f'{tab_label} {caption}')
    set_run_font(run_tc, '宋体', size_pt=10.5, bold=True)

    table = doc.add_table(rows=len(body_rows) + 1, cols=ncols)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_three_line_table(table)

    # 表头行设为重复标题行（跨页自动出现）
    tblHeader_el = OxmlElement('w:tblHeader')
    trPr = table.rows[0]._tr.get_or_add_trPr()
    trPr.append(tblHeader_el)

    for j, cell_children in enumerate(head_cells):
        cell = table.rows[0].cells[j]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = cell.paragraphs[0]
        p.clear()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        _set_para_mark_size(p, 9)
        walk_inline(p, cell_children, base_size=9)

    for i, row_data in enumerate(body_rows):
        for j, cell_children in enumerate(row_data):
            if j < ncols:
                cell = table.rows[i + 1].cells[j]
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                p = cell.paragraphs[0]
                p.clear()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.first_line_indent = Pt(0)
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                _set_para_mark_size(p, 9)
                walk_inline(p, cell_children, base_size=9)

    add_empty_para(doc)
```

- [ ] **Step 4: 运行确认通过**

Run: `cd /home/kkk/project/mddocx && source .venv/bin/activate && python skills/mddoc/scripts/md2docx.py /tmp/tbl_test.md -o /tmp/tbl_test.docx && python /tmp/verify_table.py`
Expected: `OK`

- [ ] **Step 5: 回归冒烟**

Run: `cd /home/kkk/project/mddocx && source .venv/bin/activate && python skills/mddoc/scripts/md2docx.py skills/mddoc/evals/test-sample.md -o /tmp/test.docx && echo CONVERT_OK`
Expected: `CONVERT_OK`，纯文本表格/正文输出不变

- [ ] **Step 6: Commit**

```bash
git add skills/mddoc/scripts/md2docx.py
git commit -m "feat: 表格单元格支持加粗/斜体/公式，垂直居中，公式字号9pt"
```

---

### Task 3: `test-sample.md` 追加格式化示例行（回归 fixture）

**Files:**
- Modify: `skills/mddoc/evals/test-sample.md:27`（表格最后一个数据行后追加）

**Interfaces:**
- Consumes: 无
- Produces: 含加粗/斜体/公式单元格的持久化回归样例

- [ ] **Step 1: 追加行**

在 `test-sample.md` 表格（`md2docx.py` 转换的 evals 样本）末尾数据行后追加：

```markdown
| **ViT-B/32** | *85.0%* | $87.7$M | 2021 |
```

- [ ] **Step 2: 转换验证**

Run: `cd /home/kkk/project/mddocx && source .venv/bin/activate && python skills/mddoc/scripts/md2docx.py skills/mddoc/evals/test-sample.md -o /tmp/test.docx && python - <<'PY'
from docx import Document
doc = Document('/tmp/test.docx')
last = doc.tables[-1].rows[-1]
xml = ''.join(c._tc.xml for c in last.cells)
assert '<w:b/>' in xml and '<w:i/>' in xml and 'oMath' in xml, '格式化行渲染失败'
print('OK')
PY`
Expected: `OK`

- [ ] **Step 3: Commit**

```bash
git add skills/mddoc/evals/test-sample.md
git commit -m "test: test-sample 表格加格式化单元格回归样例"
```

---

### Task 4: SKILL.md 同步（walk_inline 示例 + 表格代码段）

**Files:**
- Modify: `skills/mddoc/SKILL.md:347-379`（walk_inline 示例）、`skills/mddoc/SKILL.md:286-301`（表格单元格代码段）

**Interfaces:**
- Consumes: Task 1、Task 2 的行为定义
- Produces: AI Agent 读取的规范代码，与内置脚本一致

- [ ] **Step 1: 替换 walk_inline 示例（347-379 行）**

```python
def walk_inline(paragraph, children, base_cn='宋体', base_en='Times New Roman',
                base_size=10.5, bold=False, italic=False, underline=False):
    for child in children:
        ct = child['type']
        if ct == 'text':
            if child.get('raw', ''):
                run = paragraph.add_run(child['raw'])
                set_cn_font(run, base_cn, en_font=base_en, size_pt=base_size, bold=bold)
                if italic: run.font.italic = True
                if underline: run.font.underline = True
        elif ct == 'strong':  # **加粗** — 递归传 bold=True
            walk_inline(paragraph, child['children'], base_cn, base_en, base_size,
                        bold=True, italic=italic, underline=underline)
        elif ct == 'emphasis':  # *斜体* — 递归传 italic=True
            walk_inline(paragraph, child['children'], base_cn, base_en, base_size,
                        bold=bold, italic=True, underline=underline)
        elif ct == 'link':  # 链接 — 递归传 underline=True
            walk_inline(paragraph, child['children'], base_cn, base_en, base_size,
                        bold=bold, italic=italic, underline=True)
        elif ct == 'inline_math':  # $...$
            omml = latex_to_omml(child['raw'], display=False)
            paragraph._element.append(apply_omml_style(omml, bold, italic))
        # codespan/image/linebreak 逻辑同前
```

并在片段后补一句：`apply_omml_style()` 为公式内所有 `m:r` 添加 `<m:sty m:val="b/i/bi"/>`，使 `**$x$**`、`*$x$*` 嵌套公式生效。

- [ ] **Step 2: 替换表格单元格代码段（286-301 行）**

`header`/`data` 含义改为 mistune inline children AST 列表，渲染用 walk_inline：

```python
# 表头行 — 居中、9pt、垂直居中、无缩进
for j, children in enumerate(header):
    c = tb.rows[0].cells[j]
    c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p = c.paragraphs[0]; p.clear()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    _set_para_mark_size(p, 9)  # 公式字号跟随 9pt
    walk_inline(p, children, '宋体', size_pt=9)

# 数据行 — 左对齐、9pt、垂直居中、无缩进
for i, row in enumerate(data):
    for j, children in enumerate(row):
        c = tb.rows[i+1].cells[j]
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = c.paragraphs[0]; p.clear()
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        _set_para_mark_size(p, 9)
        walk_inline(p, children, '宋体', size_pt=9)
```

并在表格段前补一句：`header`/`data` 为 mistune `table_cell['children']` AST 列表，单元格内 **加粗**、*斜体*、$公式$ 由 `walk_inline` 渲染。

- [ ] **Step 3: 一致性检查**

Run: `cd /home/kkk/project/mddocx && grep -n "_make_temp_para\|split_inline_math" skills/mddoc/SKILL.md`
Expected: 无 `_make_temp_para` 引用（除可能遗留的 `split_inline_math` 独立章节，该章节属历史文档，本次不改）

- [ ] **Step 4: 冒烟回归**

Run: `cd /home/kkk/project/mddocx && source .venv/bin/activate && python skills/mddoc/scripts/md2docx.py examples/demo.md -o /tmp/demo.docx && echo OK`
Expected: `OK`

- [ ] **Step 5: Commit**

```bash
git add skills/mddoc/SKILL.md
git commit -m "docs: SKILL walk_inline 与表格示例同步格式状态递归"
```

---

## Self-Review

**Spec coverage:**
- 方案 A 重构 walk_inline → Task 1 ✅
- `_apply_omml_style` → Task 1 ✅
- `_render_table` 存 AST + 垂直居中 + para mark 字号 → Task 2 ✅
- 表题回退 `_extract_ast_text(head_cells[0])` → Task 2（3c）✅
- `_set_para_mark_size` → Task 2 ✅
- SKILL.md 同步（walk_inline + 表格）→ Task 4 ✅
- test-sample 回归行 → Task 3 ✅
- 版本号不改、正文布局不改 → Global Constraints ✅

**Placeholder scan:** 无 TBD/TODO；每步含完整代码或命令。

**Type consistency:** `walk_inline` 新签名在 Task 1 定义、Task 2/4 调用一致（`base_size=9`/`size_pt=9`）；`head_cells`/`body_rows` 从 `str` 改 children AST 在 Task 2 定义、Task 4 文档同步；`_apply_omml_style`/`_set_para_mark_size` 签名一致。
