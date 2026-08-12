# 表格单元格内联格式渲染设计

日期：2026-08-12
状态：已批准

## 背景与问题

当前 `_render_table`（`skills/mddoc/scripts/md2docx.py:1613`）用 `_extract_ast_text` 把每个单元格的 mistune inline AST 拍平成纯文本，再 `add_run` 单段文字。导致：

- **加粗 / 斜体丢失**：`**批量**` 只输出 `批量`，无 `w:b`
- **行内公式整个被丢弃**：`inline_math` 节点既无 `raw` 处理也无 `children`，`$x$` 在单元格中消失

正文段落和列表已用 `walk_inline`（`md2docx.py:1381`）渲染富文本，但 `walk_inline` 自身有缺陷：`strong` / `emphasis` / `link` 分支通过临时段落收集 runs 再拷贝，**嵌套在加粗/斜体/链接里的 OMML 公式会丢失**（如 `**$x$**`）。该缺陷同样存在于正文。

## 目标

1. 表格单元格支持 **斜体、加粗、行内 LaTeX 公式**，含嵌套组合（`**$x$**`、`*$x$*`）
2. 顺带修复 `walk_inline` 正文中嵌套公式丢失的同一缺陷
3. 布局微调：单元格垂直居中、单倍行距、公式字号跟随单元格 9pt
4. `SKILL.md` 同步，AI Agent 生成代码与内置脚本一致

## 方案（已确认：方案 A）

### 1. 重构 `walk_inline`（`md2docx.py:1381`）

签名增加格式状态，递归下传，替代临时段落拷贝：

```python
def walk_inline(paragraph, children, base_cn='宋体', base_en='Times New Roman',
                base_size=10.5, bold=False, italic=False, underline=False):
```

分支逻辑：

| AST 类型 | 处理 |
|----------|------|
| `text` | `add_run(raw)`，`set_run_font(bold=bold)`，`italic/underline` 按状态设置 |
| `strong` | 递归 `bold=True` |
| `emphasis` | 递归 `italic=True` |
| `link` | 递归 `underline=True` |
| `inline_math` | `latex_to_omml(display=False)` 后 `_apply_omml_style(omml, bold, italic)` |
| `codespan` | 原逻辑：`add_run(raw)` 基础字体 |
| `image` | 原逻辑：下载/占位 + `add_picture` |
| `linebreak` | 原逻辑：`add_break` |

**删除 `_make_temp_para`**（不再使用）。`_extract_ast_text` 保留（表题、标题等纯文本场景仍用）。

### 2. 新增 `_apply_omml_style(omml, bold, italic)`

给 OMML 内所有 `m:r` 增加 `<m:rPr><m:sty m:val="..."/>`：

- `m:val` 组合：仅 bold → `"b"`，仅 italic → `"i"`，都加 → `"bi"`
- 若 `m:r` 已有 `m:rPr` 且已含 `m:sty`，不重复添加
- 与已有 `m:nor`（正体 run，来自 `\text{}`）共存，`nor` 管正斜体、`sty` 管粗细，互不冲突
- bold 与 italic 都为 False 时直接返回原 omml（默认数学本就是斜体，无样式需求）

### 3. 改造 `_render_table`（`md2docx.py:1613`）

- `head_cells` / `body_rows` 从存文本改为存 **children AST 列表**（`cell.get('children', [])`）
- 表题回退逻辑（`md2docx.py:1652`）同步改为 `caption = _extract_ast_text(head_cells[0])`，否则 caption 会变成 AST dict
- 表头单元格：`vertical_alignment = CENTER`，`walk_inline(p, children, base_size=9)`，居中
- 数据单元格：同上，左对齐
- 单元格段落统一：`line_spacing = 1.0`、`space_before/after = Pt(0)`、`first_line_indent = Pt(0)`
- 每个单元格段落调用 `_set_para_mark_size(p, 9)`，让行内公式字号继承 9pt

### 4. 新增 `_set_para_mark_size(paragraph, size_pt)`

设置段落标记字号，供 OMML 继承：

```python
def _set_para_mark_size(paragraph, size_pt):
    pPr = paragraph._element.get_or_add_pPr()
    rPr = pPr.get_or_add_rPr()
    sz = rPr.find(qn('w:sz'))
    if sz is None:
        sz = OxmlElement('w:sz')
        rPr.append(sz)
    sz.set(qn('w:val'), str(int(size_pt * 2)))  # 半磅
```

需要新 import：`from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT`。

## 数据流

```
Markdown 表格
  ↓ mistune AST（table_cell.children 为 inline AST，已验证含 emphasis/strong/inline_math）
  ↓ _render_table 保留 children AST
  ↓ walk_inline(状态递归)  →  文本 run(带 w:b/w:i/w:u) + m:oMath(带 m:sty)
```

## 错误处理

- 公式解析失败：沿用 `walk_inline` 现有行为，回退为斜体纯文本（`run.font.italic = True`）
- 单元格 children 为空：walk_inline 无输出，段落为空，正常

## 测试

1. 给 `skills/mddoc/evals/test-sample.md` 表格追加一行，含 `**加粗**`、`*斜体*`、`$x_i$`、`**$y^2$**` 四种单元格
2. 转换后检查生成的 docx：
   - 加粗 run 带 `w:b`
   - 斜体 run 带 `w:i`
   - 单元格含 `m:oMath`
   - `**$y^2$**` 单元格的 `m:oMath` 内 `m:r` 带 `m:sty m:val="b"`
   - `tcPr` 含垂直居中（`vAlign center`）
   - 段落标记 `w:sz` = 18（9pt）
3. 回归：现有冒烟测试（`test-sample.md` → `/tmp/test.docx`）纯文本单元格输出不变，正文段落输出不变

## 同步

- `skills/mddoc/SKILL.md`
  - 347-379 行 `walk_inline` 示例改为状态递归版（含 `_apply_omml_style` 说明）
  - 254-304 行表格代码段：单元格改 `walk_inline` + 垂直居中 + 段落标记字号

## 非目标

- 不改版本号（本次不发布）
- 不改正文段落布局（仅修复嵌套公式丢失，正常文本输出与现状一致）
- 不做列宽自适应、行高调整等其他表格布局调整
