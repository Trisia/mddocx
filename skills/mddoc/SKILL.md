---
name: mddoc
description: 将 Markdown 内容转换为特定学术格式的 Word 文档 (.docx)。当用户想要将 Markdown 文件、粘贴的 Markdown 文本转换为格式化的 docx 文档时使用，特别是学术论文、技术报告、毕业论文等需要严格格式要求的场景。触发词包括：/mddoc、markdown转docx、md转word、生成格式化文档、学术格式转换。即使用户只说"把这个转成word"而内容是Markdown，也应使用此技能。
---

# mddoc — Markdown 转学术格式 DOCX

## 快速开始

```bash
# 安装依赖（仅首次）
pip install python-docx Pillow requests mistune

# 转换 Markdown 文件 → 输出到同目录
python <skill-path>/scripts/md2docx.py paper.md

# 指定输出路径
python <skill-path>/scripts/md2docx.py paper.md -o /path/to/output.docx

# 直接转换粘贴的文本
python <skill-path>/scripts/md2docx.py --text "# 标题\n\n正文内容" -o out.docx
```

其中 `<skill-path>` = `/home/kkk/.claude/skills/mddoc`

## 工作流程

1. **读取输入** — 若用户粘贴 Markdown 文本则直接读取；若用户提供文件路径（含 `@` 引用）则读取该文件
2. **安装依赖**（如未安装）— `pip install python-docx Pillow requests mistune`
3. **执行转换** — 优先使用内置脚本 `scripts/md2docx.py`；若 Markdown 结构特殊则参照下方格式规范编写自定义脚本
4. **确定输出** — 文件路径输入→同目录；粘贴内容→当前目录；文件名=「题目.docx」（题目从第一个 `# 标题` 提取；若无 `#` 标题则使用输入文件名，`--text` 模式为「未命名文档.docx」）
5. **验证** — 检查 outline level、图片嵌入、页眉、分页符

---

## 格式参考

> 每个元素配可直接使用的 python-docx 代码。公共导入和工具函数：

```python
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cn_font(run, cn_name, en_name='Times New Roman', size_pt=10.5, bold=False):
    """设置 run 的中英文字体、字号、加粗、颜色(黑)"""
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.name = en_name
    run.font.color.rgb = None
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts'); rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), cn_name)
    rFonts.set(qn('w:ascii'), en_name)
    rFonts.set(qn('w:hAnsi'), en_name)

def add_empty(doc):
    """五号空行"""
    p = doc.add_paragraph()
    run = p.add_run('')
    set_cn_font(run, '宋体', size_pt=10.5)
    return p

def set_outline(para, level):
    """设置 outline level（XML方式，兼容所有python-docx版本）"""
    pPr = para._element.get_or_add_pPr()
    ol = OxmlElement('w:outlineLvl')
    ol.set(qn('w:val'), str(level))
    pPr.append(ol)
```

### 基础设置

五号=10.5pt，1.3倍行距，段前段后0磅，全黑，A4页边距左3cm右2cm上2cm下2cm，页脚距底1cm。

```python
doc = Document()
section = doc.sections[0]
# 页边距：左3cm 右2cm 上2cm 下2cm
section.left_margin = Cm(3)
section.right_margin = Cm(2)
section.top_margin = Cm(2)
section.bottom_margin = Cm(2)

sty = doc.styles['Normal']
sty.font.size = Pt(10.5)
sty.font.name = 'Times New Roman'
sty.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
sty.paragraph_format.line_spacing = 1.3
sty.paragraph_format.space_before = Pt(0)
sty.paragraph_format.space_after = Pt(0)
```

### 空行规则

- **段落之间**：不空行
- **题目**：上下各空一行
- **一级标题**：上下各空一行
- **二级/三级标题**：上下不空行
- **图片**：上方空一行，图题下方空一行
- **表格**：表题上方空一行，表格下方空一行

### 题目（第一个 `#`）

三号黑体(16pt)、居中、上下各空一行、**不设 outline level**。

```python
add_empty(doc)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
# 题目不设 outline_level
run = p.add_run(title_text)
set_cn_font(run, '黑体', size_pt=16)
add_empty(doc)
```

### 无标题文档

当输入 Markdown 没有 `#` 标题时：
- 跳过标题页，文档从第一个内容节点开始排版
- 页眉右侧显示「未命名文档」代替标题
- 页眉左侧「xxxxx」保持不变
- 输出文件名：文件输入使用输入文件名（如 `paper.md` → `paper.docx`），`--text` 模式使用「未命名文档.docx」

### 一级标题（后续 `#`）

三号黑体(16pt)、居中、上下各空一行、outline_level=1、**前加分页符**。


```python
add_empty(doc)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_outline(p, 1)
# 分页符
pPr = p._element.get_or_add_pPr()
pPr.append(OxmlElement('w:pageBreakBefore'))
run = p.add_run(text)
set_cn_font(run, '黑体', size_pt=16)
add_empty(doc)
```

### 二级标题（`##`）

四号黑体(14pt)、**不加粗**、顶格(无缩进)、outline_level=2。上下不空行。

```python
p = doc.add_paragraph()
set_outline(p, 2)
p.paragraph_format.first_line_indent = Pt(0)
run = p.add_run(text)
set_cn_font(run, '黑体', size_pt=14, bold=False)
```

### 三级标题（`###`）

小四宋体(12pt)、**不加粗**、首行缩进Pt(21)(=2个中文字)、outline_level=3。

```python
p = doc.add_paragraph()
set_outline(p, 3)
p.paragraph_format.first_line_indent = Pt(21)
run = p.add_run(text)
set_cn_font(run, '宋体', size_pt=12, bold=False)
```

### 正文段落

五号(10.5pt)宋体/TNR、首行缩进Pt(21)、1.3倍行距、段前段后0磅。

```python
p = doc.add_paragraph()
p.paragraph_format.first_line_indent = Pt(21)
p.paragraph_format.line_spacing = 1.3
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
run = p.add_run(text)
set_cn_font(run, '宋体', size_pt=10.5)
```

### 图片

下载嵌入(不压缩)、等比缩放(8-12cm规则)、居中、上方空一行。

```python
import requests
from PIL import Image

def dl_image(url):
    h = {'User-Agent': 'Mozilla/5.0 (compatible; mddoc/1.0)'}
    r = requests.get(url, headers=h, timeout=30); r.raise_for_status()
    return r.content  # bytes, 可保存为临时文件

def img_width(path):
    img = Image.open(path)
    dpi = img.info.get('dpi', (96, 96))
    dx = dpi[0] if dpi and dpi[0] > 0 else 96
    w_cm = img.width / dx * 2.54
    if   w_cm > 12: tw = 12.0
    elif w_cm < 8:  tw = 8.0
    else:           tw = w_cm
    return Cm(tw)

add_empty(doc)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run().add_picture(img_path, width=img_width(img_path))
```

### 图题

图片下方(不空行)、小五(9pt)宋体加粗居中、逐章编号。有章标题时格式 `图<章>-<序号> <alt>`，无章标题时格式 `图<序号> <alt>`。若 alt 为空，自动使用图片文件名（不含扩展名）作为图题。

```python
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run(f'图{ch}-{idx} {alt_text}')
set_cn_font(run, '宋体', size_pt=9, bold=True)
add_empty(doc)  # 图题下方空一行
```

### 表格 + 表题

**三线表**(顶线1.5pt/表头底线0.75pt/底线1.5pt粗，无竖线)。表题在上方(不空行)、五号(10.5pt)宋体加粗居中。有章标题时格式`表<章>-<序号> <描述>`，无章标题时格式`表<序号> <描述>`。

```python
# 表题
add_empty(doc)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run(f'表{ch}-{idx} {caption}')
set_cn_font(run, '宋体', size_pt=10.5, bold=True)

# 表格
tb = doc.add_table(rows=nrows, cols=ncols)
tb.alignment = WD_ALIGN_PARAGRAPH.CENTER

# 三线表边框
Pr = tb._tbl.tblPr or OxmlElement('w:tblPr')
B = OxmlElement('w:tblBorders')
for tag, sz in [('top','12'),('bottom','12')]:
    e = OxmlElement(f'w:{tag}'); e.set(qn('w:val'),'single')
    e.set(qn('w:sz'),sz); e.set(qn('w:space'),'0'); e.set(qn('w:color'),'000000')
    B.append(e)
for tag in ('left','right','insideH','insideV'):
    e = OxmlElement(f'w:{tag}'); e.set(qn('w:val'),'none')
    e.set(qn('w:sz'),'0'); e.set(qn('w:space'),'0'); e.set(qn('w:color'),'auto')
    B.append(e)
Pr.append(B)
# 表头行每格底部加细线 0.75pt（cell级别，不影响数据行）
for cell in tb.rows[0].cells:
    tcPr = cell._tc.get_or_add_tcPr()
    tcB = OxmlElement('w:tcBorders')
    btm = OxmlElement('w:bottom')
    btm.set(qn('w:val'),'single'); btm.set(qn('w:sz'),'6')
    btm.set(qn('w:space'),'0'); btm.set(qn('w:color'),'000000')
    tcB.append(btm); tcPr.append(tcB)

# 表头行 — 居中、9pt、无缩进
for j, txt in enumerate(header):
    c = tb.rows[0].cells[j]; c.paragraphs[0].clear()
    c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = c.paragraphs[0].add_run(txt)
    set_cn_font(run, '宋体', size_pt=9)
    c.paragraphs[0].paragraph_format.first_line_indent = Pt(0)

# 数据行 — 左对齐、9pt、无缩进
for i, row in enumerate(data):
    for j, txt in enumerate(row):
        c = tb.rows[i+1].cells[j]; c.paragraphs[0].clear()
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = c.paragraphs[0].add_run(txt)
        set_cn_font(run, '宋体', size_pt=9)
        c.paragraphs[0].paragraph_format.first_line_indent = Pt(0)

add_empty(doc)  # 表格下方空一行
```

### 页眉

左顶格"xxxxx"、右顶格文档题目、小五(9pt)黑体。

```python
sec = doc.sections[0]; hdr = sec.header
hdr.is_linked_to_previous = False
r = hdr.paragraphs[0].add_run('xxxxx')
set_cn_font(r, '黑体', size_pt=9)
p2 = hdr.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r2 = p2.add_run(title_text)
set_cn_font(r2, '黑体', size_pt=9)
```

### 列表

有序列表用（1）、（2）、（3）序号；无序列表用 • 符号。按正文格式（首行缩进Pt(21)）。

```python
# 有序列表 — (1)、(2)、(3) 格式
for idx, item in enumerate(items, 1):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(21)
    run = p.add_run(f'（{idx}）{item}')
    set_cn_font(run, '宋体', size_pt=10.5)

# 无序列表
for item in items:
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(21)
    run = p.add_run(f'• {item}')
    set_cn_font(run, '宋体', size_pt=10.5)
```

### 代码块

Courier New 等宽字体、五号、左缩进。

```python
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(1)
run = p.add_run(code)
set_cn_font(run, '宋体', en_name='Courier New', size_pt=10.5)
```

### 行内公式（`$...$`）

嵌于段落中，转换为 OMML（Office Math Markup Language）格式，WPS/Word 可直接渲染。
公式在段落中作为 `<m:oMath>` 元素插入，与文字 run 同级。

```python
# 行内公式通过 split_inline_math() 拆分为 segments
# text segment → 普通 run，math segment → latex_to_omml() → m:oMath 元素
p = doc.add_paragraph()
p.paragraph_format.first_line_indent = Pt(21)
p.paragraph_format.line_spacing = 1.3
for seg in segments:
    if seg['type'] == 'text':
        if seg['content']:
            run = p.add_run(seg['content'])
            set_cn_font(run, '宋体', size_pt=10.5)
    else:  # math — 转换为 OMML 并插入段落 XML
        omml = latex_to_omml(seg['content'], display=False)
        p._element.append(omml)
```

### 行间公式（`$$...$$`）

独立成行，上下各空一行。公式通过 OMML 居中渲染，编号右对齐。编号格式同图/表：`(章-序号)`，逐章编号。编号括号用五号宋体，数字用五号 Times New Roman。

布局方式：段落设置居中和右对齐 tab stop，OMML `<m:oMath>` 元素位于两个 tab 之间，实现公式居中、编号右对齐。

```python
add_empty(doc)  # 上方空一行
p = doc.add_paragraph()

# 设置 tab stops：居中和右对齐
# 左边距3cm右边距2cm → 可用宽度16cm → 中心8cm(4536twips), 右边16cm(9072twips)
pPr = p._element.get_or_add_pPr()
tabs = OxmlElement('w:tabs')
ct = OxmlElement('w:tab'); ct.set(qn('w:val'), 'center'); ct.set(qn('w:pos'), '4536')
rt = OxmlElement('w:tab'); rt.set(qn('w:val'), 'right'); rt.set(qn('w:pos'), '9072')
tabs.append(ct); tabs.append(rt)
pPr.append(tabs)

# tab → 公式 OMML 居中
run_t1 = p.add_run()
run_t1._r.append(OxmlElement('w:tab'))
omml = latex_to_omml(formula_text, display=False)
p._element.append(omml)

# tab → 编号右对齐：(章-序号)
run_t2 = p.add_run()
run_t2._r.append(OxmlElement('w:tab'))
run_lp = p.add_run('(')
set_cn_font(run_lp, '宋体', en_name='宋体', size_pt=10.5)
run_num = p.add_run(f'{ch}-{eq_num}')
set_cn_font(run_num, '宋体', en_name='Times New Roman', size_pt=10.5)
run_rp = p.add_run(')')
set_cn_font(run_rp, '宋体', en_name='宋体', size_pt=10.5)

add_empty(doc)  # 下方空一行
```

### LaTeX → OMML 转换

内置 `latex_to_omml()` 函数将 LaTeX 数学公式转换为 OMML XML，支持以下构造：

- 上下标：`a^b`、`a_b`、`a_b^c`
- 分数：`\frac{a}{b}`
- 根式：`\sqrt{x}`、`\sqrt[n]{x}`
- 希腊字母：`\alpha`、`\beta`、`\Gamma` 等
- 大运算符：`\sum`、`\prod`、`\int`（含上下限）
- 函数名：`\sin`、`\cos`、`\log`、`\lim` 等
- 重音：`\hat`、`\bar`、`\vec`、`\tilde` 等
- 符号：`\infty`、`\cdot`、`\times`、`\partial`、`\nabla` 等
- 定界符：`\left(...\right)`、`\left[...\right]`
- 文本：`\text{...}`

### 页码

从第一章（第一个一级标题所在页）开始编页码，格式 `第×页  共×页`，页脚边距 1cm。

```python
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

section = doc.sections[0]
section.footer_distance = Cm(1.1)  # 页脚边距

footer = section.footer
footer.is_linked_to_previous = False
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER

# "第" 字
run1 = fp.add_run('第')
set_cn_font(run1, '宋体', size_pt=10.5)
# 页码字段
run_page = fp.add_run()
fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
run_page._r.append(fldChar1)
instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve')
instrText.text = ' PAGE '
run_page._r.append(instrText)
fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
run_page._r.append(fldChar2)
# "页  共"
run2 = fp.add_run('页  共')
set_cn_font(run2, '宋体', size_pt=10.5)
# 总页数字段
run_total = fp.add_run()
fldChar3 = OxmlElement('w:fldChar'); fldChar3.set(qn('w:fldCharType'), 'begin')
run_total._r.append(fldChar3)
instrText2 = OxmlElement('w:instrText'); instrText2.set(qn('xml:space'), 'preserve')
instrText2.text = ' NUMPAGES '
run_total._r.append(instrText2)
fldChar4 = OxmlElement('w:fldChar'); fldChar4.set(qn('w:fldCharType'), 'end')
run_total._r.append(fldChar4)
# "页"
run3 = fp.add_run('页')
set_cn_font(run3, '宋体', size_pt=10.5)
```

> 注意：页码从第一章开始计数需在第一个一级标题处插入分节符并重置页码。简单场景可统一从第1页开始。

### 跨页表格（续表）

表头行设为重复标题行，跨页自动重复。续表在右上方标"续表<编号>"。

```python
# 表头行设为重复标题行（跨页自动出现）
tblHeader = OxmlElement('w:tblHeader')
trPr = table.rows[0]._tr.get_or_add_trPr()
trPr.append(tblHeader)

# 对于确实跨页的大表格，在分页处手动加"续表"标记：
# （python-docx无法检测分页位置，建议生成后人工检查）
# 续表右对齐段落示例：
p_cont = doc.add_paragraph()
p_cont.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run_cont = p_cont.add_run(f'续表{tab_label}')
set_cn_font(run_cont, '宋体', size_pt=10.5)
```

---

## 生成后自检清单

- [ ] 题目：16pt黑体、居中、无outline、上下各空一行
- [ ] 无标题文档：跳过标题页、页眉显示「未命名文档」
- [ ] 一级标题：16pt黑体、outline_level=1、前有分页符、上下各空一行
- [ ] 二级标题：14pt黑体、顶格、outline_level=2、不加粗、**上下不空行**
- [ ] 三级标题：12pt宋体、首行缩进Pt(21)、outline_level=3、不加粗、**上下不空行**
- [ ] 正文：10.5pt、首行缩进Pt(21)、1.3倍行距、**段落间不空行**
- [ ] 图片：嵌入、等比缩放8-12cm、居中、上下各空一行
- [ ] 图题：9pt宋体加粗居中、图片下方
- [ ] 表格：三线表(顶/底粗、表头格底线细、数据行无线)、居中
- [ ] 表头行：居中、9pt、无缩进、tblHeader重复
- [ ] 表题：10.5pt宋体加粗居中、表格上方
- [ ] 列表：有序用（1）（2）（3）、无序用•、首行缩进
- [ ] 页眉：左"xxxxx"右题目、9pt黑体
- [ ] 行内公式：$...$ 嵌于段落、OMML格式、WPS/Word可渲染
- [ ] 行间公式：$$...$$ OMML居中、上下各空一行、编号(章-序号)右对齐、括号宋体数字TNR
- [ ] 页码："第×页 共×页"、页脚边距1cm
- [ ] 续表：跨页表头重复、"续表xx"右上标注
- [ ] 页边距：左3cm 右2cm 上2cm 下2cm
- [ ] 全部黑色、无额外参数

## 注意事项

- **第一个 `#` 是题目**（不设 outline），后续 `#` 是一级标题（outline_level=1）；若无 `#` 标题仍可正常生成文档
- **图片尺寸用 `Cm()`，不手算 EMU** — `add_picture(width=Cm(x))` 自动转换
- **图片下载必设 User-Agent** — 否则 CDN/Wikipedia 返回 400
- **图题编号自动生成** — alt text 作为描述文字，图/表编号独立逐章编序
- **表题识别** — 表格前含"表"字的段落作为表题
- **outline_level 用 `set_outline()` 设置** — `paragraph_format.outline_level` 在部分python-docx版本不生效，统一用XML方式写入；读取时也从XML读取
- **公式编号自动生成** — 图/表/公式独立逐章编序，编号格式统一为 `章-序号`
- **不添加**：目录页、背景色
- **单位**：字体用 pt、尺寸用 cm
