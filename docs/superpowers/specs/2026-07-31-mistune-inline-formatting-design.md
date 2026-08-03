# Markdown 内联格式化与列表支持 — 设计文档

日期: 2026-07-31
状态: 待审批

## 概述

用 mistune 3.x 全面替换手写 Markdown 解析器，同时新增 5 项格式支持：

1. **分隔线 `---`** → 分页符（3+ 连续 `-`，多个连续视为一个）
2. **加粗 `**text**`** → `run.font.bold = True`
3. **斜体 `*text*`** → `run.font.italic = True`
4. **加粗斜体 `***text***`** → bold + italic
5. **列表** → Word 原生有序/无序列表，带富文本 inline

## 架构变化

```
当前流程:
  Markdown → parse_markdown() [手写150行, 行级正则可读性差]
          → split_inline_math()
          → 自定义 nodes [dict, type/level/text/children/header/rows]
          → generate_docx() [逐个节点生成, inline仅支持文本+$公式]

改为:
  Markdown → mistune(ast, math插件) [1行]
          → AST [dict, 标准格式]
          → generate_docx() [AST walker + walk_inline()递归]
```

代码量: 删除 ~190 行手写解析, 新增 ~160 行 walk 函数, 净变化 -30 行。

## 解析层

### mistune 配置

```python
import mistune
from mistune.plugins import math

md = mistune.create_markdown(
    renderer='ast',
    plugins=[math.math]
)
ast = md(markdown_text)  # → list[dict]
```

### 关键 AST 节点映射

| Markdown | mistune AST type |
|----------|-----------------|
| `# Title` | `heading` (attrs.level=1) |
| `## H2` | `heading` (attrs.level=2) |
| 普通段落 | `paragraph` (children: inline nodes) |
| `**bold**` | `strong` |
| `*italic*` | `emphasis` |
| `***both***` | `emphasis > strong` (嵌套) |
| `` `code` `` | `codespan` |
| `$x^2$` | `inline_math` (raw="x^2") |
| `$$...$$` | `block_math` (raw="...") |
| `---` | `thematic_break` |
| `- item` | `list` (ordered=False) |
| `1. item` | `list` (ordered=True) |
| `|...|` | `table` |
| ` ``` ` | `block_code` |
| `![alt](url)` | `image` (inline, in paragraph) |
| `[text](url)` | `link` |

## 生成层: generate_docx 重构

### 主分发循环

```python
def generate_docx(ast, output_path, title_text=None):
    doc = Document()
    # ... 样式/页边距/页眉/页脚设置（不变） ...

    prev_node = None  # 表题识别: 前一个 paragraph 节点
    for node in ast:
        t = node['type']

        if t == 'heading':
            render_heading(doc, node)
        elif t == 'paragraph':
            p = doc.add_paragraph()
            # 正文样式
            p.paragraph_format.first_line_indent = Pt(21)
            p.paragraph_format.line_spacing = 1.3
            walk_inline(p, node.get('children', []))
            prev_node = node
        elif t == 'thematic_break':
            # 插入空段落 + pageBreakBefore
            p = doc.add_paragraph()
            run = p.add_run('')
            set_run_font(run, '宋体', size_pt=10.5)
            add_page_break_before(p)
        elif t == 'list':
            render_list(doc, node)
        elif t == 'block_math':
            render_display_math(doc, node['raw'])
        elif t == 'block_code':
            render_code_block(doc, node['raw'])
        elif t == 'table':
            # 表题: 检查 prev_node 是否为 paragraph 且含"表"字
            caption = extract_table_caption(prev_node)
            render_table(doc, node, caption)
        elif t == 'blank_line':
            pass  # 忽略

    doc.save(output_path)
```

### walk_inline(paragraph, children)

递归遍历 inline AST，为每个文本片段创建带格式的 run：

```python
def walk_inline(paragraph, children):
    """递归遍历 inline AST, 将格式化文本/OMML 公式附加到段落。"""
    for child in children:
        ct = child['type']

        if ct == 'text':
            if child.get('raw', ''):
                run = paragraph.add_run(child['raw'])
                set_run_font(run, '宋体', size_pt=10.5)

        elif ct == 'strong':
            # 加粗: 新建 run, 设置 bold=True, 递归处理 children
            sub_p = _create_temp_paragraph()
            walk_inline(sub_p, child.get('children', []))
            for r in sub_p.runs:
                new_run = paragraph.add_run(r.text)
                set_run_font(new_run, '宋体', size_pt=10.5, bold=True)
                # 如果内部有 emphasis: bold+italic
                if r.font.italic:
                    new_run.font.italic = True

        elif ct == 'emphasis':
            sub_p = _create_temp_paragraph()
            walk_inline(sub_p, child.get('children', []))
            for r in sub_p.runs:
                is_bold = r.font.bold  # 检测嵌套 strong
                new_run = paragraph.add_run(r.text)
                set_run_font(new_run, '宋体', size_pt=10.5, bold=is_bold)
                new_run.font.italic = True

        elif ct == 'inline_math':
            omml = latex_to_omml(child['raw'], display=False)
            paragraph._element.append(omml)

        elif ct == 'codespan':
            run = paragraph.add_run(child.get('raw', ''))
            set_run_font(run, '宋体', en_font='Times New Roman', size_pt=10.5)

        elif ct == 'link':
            # 链接: 蓝色带下划线
            sub_p = _create_temp_paragraph()
            walk_inline(sub_p, child.get('children', []))
            for r in sub_p.runs:
                new_run = paragraph.add_run(r.text)
                set_run_font(new_run, '宋体', size_pt=10.5)
                new_run.font.color.rgb = RGBColor(0, 0, 255)
                new_run.font.underline = True

        elif ct == 'image':
            # 行内图片
            url = child.get('attrs', {}).get('url', '')
            alt = child.get('alt', '')
            img_path = download_image(url)
            if img_path:
                w, h = calc_image_size(img_path)
                run = paragraph.add_run()
                run.add_picture(img_path, width=w)
            # alt 作为 run text 追加（图题另处理）

        elif ct == 'linebreak':
            run = paragraph.add_run()
            run.add_break()
```

**嵌套处理**: `***both***` → AST: `emphasis > strong > text("both")`。
walk_inline 先遇 emphasis → 创建临时段落 → 遇 strong → 设置 bold → 遇 text。
最终一个 run: bold=True + italic=True。自然处理，无需特殊分支。

### render_list(doc, node)

list AST 结构：
```json
{
  "type": "list",
  "ordered": false,
  "bullet": "-",
  "children": [
    {
      "type": "list_item",
      "children": [
        {"type": "block_text", "children": [{"type": "text", "raw": "item 1"}]}
      ]
    }
  ]
}
```

每条 `list_item` 对应一个段落的 paragraph，包含 `block_text` children（可含富文本 inline）。

```python
def render_list(doc, node):
    """渲染 Word 原生列表（有序/无序），支持富文本 inline。"""
    ordered = node.get('ordered', False)
    
    for item in node.get('children', []):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.3
        
        if ordered:
            # Word 原生编号: 通过 XML numPr 设置
            _set_num_style(p, ilvl=0)  # 1, 2, 3...
        else:
            # Word 原生项目符号: 黑圆点 •
            _set_bullet_style(p, ilvl=0)
        
        # 悬挂缩进: 左缩进 0.63cm, 首行缩进 -0.63cm (符号悬出)
        p.paragraph_format.left_indent = Cm(0.63)
        p.paragraph_format.first_line_indent = Cm(-0.63)
        
        # 渲染列表项内富文本
        for child in item.get('children', []):
            if child['type'] == 'block_text':
                walk_inline(p, child.get('children', []))
```

列表项内自动支持 `**加粗**`、`*斜体*`、`$公式$` 等 inline 格式（walk_inline 处理）。

**跨页编号连续**：同一 `numId` 的所有段落由 Word 自动维护编号序列。列表内容跨页时，第二页自动从上一页末项编号继续递增，无需额外处理。

### 表题识别

唯一需要相邻节点判断的场景。Markdown 用 `表 xxx` 段落 + 表格表示，无专用语法。

```python
def extract_table_caption(prev_node):
    """从表格前一个段落节点提取表题文本。"""
    if prev_node and prev_node['type'] == 'paragraph':
        # 收集所有 text run 的 raw
        text = ''.join(
            c['raw'] for c in flatten_inline(prev_node.get('children', []'))
            if c['type'] == 'text'
        )
        if '表' in text:
            return text.strip()
    return ''
```

## Word 标题样式清理

文档生成前，删除/覆盖 Word 内置 Heading 1-6 样式，替换为学术格式定义。防止 Word 默认样式（蓝色、Cambria 字体、加粗等）覆盖我们的自定义格式。

```python
def _override_builtin_heading_styles(doc):
    """将 Word 内置 Heading 1-6 样式改为自定义学术格式

    参数：
        doc: python-docx Document 对象
    """
    heading_specs = {
        'Heading 1': {'font': '黑体', 'size': 16, 'bold': True, 'align': WD_ALIGN_PARAGRAPH.CENTER},
        'Heading 2': {'font': '黑体', 'size': 14, 'bold': False, 'align': WD_ALIGN_PARAGRAPH.LEFT},
        'Heading 3': {'font': '宋体', 'size': 12, 'bold': False, 'align': WD_ALIGN_PARAGRAPH.LEFT},
        'Heading 4': {'font': '宋体', 'size': 10.5, 'bold': False, 'align': WD_ALIGN_PARAGRAPH.LEFT},
        'Heading 5': {'font': '宋体', 'size': 10.5, 'bold': False, 'align': WD_ALIGN_PARAGRAPH.LEFT},
        'Heading 6': {'font': '宋体', 'size': 10.5, 'bold': False, 'align': WD_ALIGN_PARAGRAPH.LEFT},
    }

    for style_name, spec in heading_specs.items():
        style = doc.styles[style_name]
        style.font.size = Pt(spec['size'])
        style.font.name = 'Times New Roman'
        style.font.bold = spec['bold']
        style.font.color.rgb = None  # 黑色（None = 自动 = 黑色）
        style.font.italic = False
        style.font.underline = False
        # 中文字体
        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:eastAsia'), spec['font'])
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        # 段落格式
        pPr = style.element.get_or_add_pPr()
        # 清除默认段前段后间距
        spacing = pPr.find(qn('w:spacing'))
        if spacing is None:
            spacing = OxmlElement('w:spacing')
            pPr.append(spacing)
        spacing.set(qn('w:before'), '0')
        spacing.set(qn('w:after'), '0')
        spacing.set(qn('w:line'), '312')  # 1.3x line spacing in 240th of a line
        spacing.set(qn('w:lineRule'), 'auto')
```

在 `generate_docx()` 开头调用：
```python
def generate_docx(ast, output_path, title_text=None):
    doc = Document()
    _override_builtin_heading_styles(doc)  # ← 新增
    # ... 后续样式/页眉页脚设置 ...
```

效果：Word 样式面板中 Heading 1-6 显示为黑色宋体/黑体，打开文档时不会出现蓝色标题等意外格式。

## 不变部分

以下完全不变：

- `_LatexParser` 类（~690 行）— LaTeX → OMML 解析器
- `latex_to_omml()` — 公式转换入口
- `set_run_font()` — 字体设置
- `set_three_line_table()` — 三线表边框
- `set_outline_level()` — outline 级别
- `calc_image_size()` / `download_image()` — 图片处理
- `add_page_break_before()` — 分页符
- 页眉/页脚/页码/页边距设置
- 表题/图题编号计数器

## 文件变更清单

| 文件 | 变更 |
|------|------|
| `skills/mddoc/scripts/md2docx.py` | 删除 `parse_markdown()`, `split_inline_math()`<br>新增 `walk_inline()`, `render_list()`, `_override_builtin_heading_styles()`, `_create_temp_paragraph()`, `flatten_inline()`<br>重写主循环为 AST walker<br>删除 `import re` 中不再需要的部分 |
| `skills/mddoc/SKILL.md` | 更新 inline 格式和列表的代码示例<br>版本号不变 |

## mistune 版本

当前 `.venv` 中 mistune 版本: 3.3.3。依赖 `mistune.plugins.math` 插件（内置，无需额外安装）。

## 测试验证

转换 `skills/mddoc/evals/test-sample.md`，验证：

- [ ] `**加粗**` 在 Word 中显示为 bold
- [ ] `*斜体*` 在 Word 中显示为 italic
- [ ] `***加粗斜体***` 显示为 bold+italic
- [ ] 列表项内富文本（如 `- **加粗** *斜体*`）正确渲染
- [ ] 有序列表（`1.` `2.`）使用 Word 原生编号
- [ ] 无序列表（`-`）使用 Word 原生黑圆点
- [ ] `---` 后内容在新页开始，连续多个 `---` 视为一个
- [ ] 行内公式 `$x^2$` 仍为 OMML 格式，WPS/Word 可渲染
- [ ] 行间公式 `$$...$$` 仍为 OMML 居中 + 编号右对齐
- [ ] 题目/标题/图片/表格/代码块 格式不变
- [ ] 页眉/页脚/页码 不变
- [ ] Heading 1-6 样式在 Word 样式面板中显示为黑色黑体/宋体（非 Word 默认蓝色）

## 风险

- **表格 AST 结构差异**: mistune 的 table 节点结构与当前手写解析不同（attrs 中 header=...body=...），需适配。风险低。
- **嵌套列表**: 当前不支持嵌套，mistune 支持。若输入含嵌套列表，需确定 `ilvl` 处理。当前设计仅处理 ilvl=0。后续可扩展。
