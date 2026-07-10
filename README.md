# mddocx — Markdown 转学术格式 DOCX

[![Claude Code](https://img.shields.io/badge/Agent-Claude%20Code-orange?logo=claude)](https://claude.com/code)
[![Codex](https://img.shields.io/badge/Agent-Codex-blue?logo=openai)](https://github.com/openai/codex)
[![Cursor](https://img.shields.io/badge/Agent-Cursor-6c47ff?logo=cursor)](https://cursor.com)
[![OpenCode](https://img.shields.io/badge/Agent-OpenCode-teal)](https://opencode.ai)
[![npm](https://img.shields.io/npm/v/@cliven/mddocx?color=red)](https://www.npmjs.com/package/@cliven/mddocx)
[![Version](https://img.shields.io/github/v/release/Trisia/mddocx)](https://github.com/Trisia/mddocx/releases)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue?logo=python)](https://python.org)
[![CI](https://github.com/Trisia/mddocx/actions/workflows/release.yml/badge.svg)](https://github.com/Trisia/mddocx/actions/workflows/release.yml)
[![Clawhub](https://img.shields.io/badge/Clawhub-mddocx-8b5cf6)](https://clawhub.ai/Trisia/mddoc)

将 Markdown 转换为符合学术规范的 Word 文档的Agent Skill，支持 LaTeX 公式（OMML）、三线表、图题/表题自动编号、页码、页眉等学术论文排版规范。

![MD与DOCX生成效果](./examples/demo.png)

## 安装

### 🤖 一键安装

将以下内容复制粘贴给任意智能体（Claude Code / Codex / OpenCode / Cursor）即可安装。

> 各平台专属安装说明详见 **[INSTALL.md](INSTALL.md)**。

```copy
请根据 https://github.com/Trisia/mddocx/blob/main/INSTALL.md 中对应平台的说明，帮我安装 mddocx
```

### 依赖

```bash
pip install python-docx Pillow requests mistune
```

## 使用

### npx（无需安装）

```bash
npx @cliven/mddocx paper.md                    # 转换文件
npx @cliven/mddocx paper.md -o output.docx     # 指定输出
npx @cliven/mddocx --text "# 标题" -o out.docx # 转换文本
```

### Claude Code 中

```
/mddoc paper.md                    # 转换 Markdown 文件
/mddoc @paper.md                   # @引用文件
/mddoc 把这段内容转成Word         # 粘贴 Markdown 文本
```

### 命令行直接使用

```bash
# 转换文件（输出到同目录）
python skills/mddoc/scripts/md2docx.py paper.md

# 指定输出路径
python skills/mddoc/scripts/md2docx.py paper.md -o output.docx

# 直接转换文本
python skills/mddoc/scripts/md2docx.py --text "# 标题\n\n正文" -o out.docx
```

## 格式规范

生成的文档自动应用以下学术排版规范：

| 元素 | 格式 |
|------|------|
| 题目 | 三号黑体(16pt)、居中、上下空一行 |
| 一级标题 | 三号黑体、居中、前加分页符、outline_level=1 |
| 二级标题 | 四号黑体(14pt)、顶格、不加粗、outline_level=2 |
| 三级标题 | 小四宋体(12pt)、首行缩进、不加粗、outline_level=3 |
| 正文 | 五号(10.5pt)、首行缩进2字符、1.3倍行距 |
| 表格 | 三线表(顶线粗/表头底线细/底线粗)、表头重复 |
| 图题 | 小五(9pt)宋体加粗居中、"图1-1 xxx"格式 |
| 表题 | 五号(10.5pt)宋体加粗居中、"表1-1 xxx"格式 |
| 页码 | "第×页 共×页"、页脚边距1cm |
| 列表 | 有序列表用（1）（2）（3）序号 |
| 行内公式 | $...$ 转 OMML、嵌于段落、WPS/Word 可渲染 |
| 行间公式 | $$...$$ 转 OMML 居中、编号(章-序号)右对齐 |
| 页边距 | 左3cm 右2cm 上2cm 下2cm |

