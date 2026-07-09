# mddocx — Markdown 转学术格式 DOCX

将 Markdown 转换为符合学术规范的 Word 文档的Agent Skill，支持三线表、图题/表题自动编号、页码、页眉等学术论文排版规范。

[![Claude Code](https://img.shields.io/badge/Agent-Claude%20Code-orange?logo=claude)](https://claude.com/code)
[![Codex](https://img.shields.io/badge/Agent-Codex-blue?logo=openai)](https://github.com/openai/codex)
[![OpenCode](https://img.shields.io/badge/Agent-OpenCode-teal)](https://opencode.ai)
[![Version](https://img.shields.io/github/v/release/Trisia/mddocx)](https://github.com/Trisia/mddocx/releases)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue?logo=python)](https://python.org)

![Markdown → DOCX 对比](assets/comparison.png)

## 安装

### 通过 Git 安装（推荐）

```bash
# 克隆到本地
git clone https://github.com/Trisia/mddocx ~/.claude/plugins/mddocx

# 在 Claude Code 中安装
/plugin install ~/.claude/plugins/mddocx
```

### Codex (OpenAI)

在 Codex 中安装：

```bash
/plugin install git:https://github.com/Trisia/mddocx
```

### OpenCode

在 `opencode.json` 中添加插件：

```json
{
  "plugins": ["git:https://github.com/Trisia/mddocx"]
}
```

详见 [.opencode/INSTALL.md](.opencode/INSTALL.md)

### 直接复制

```bash
cp -r ~/project/mddocx/skills/mddoc ~/.claude/skills/mddoc
```

### 依赖安装

```bash
pip install python-docx Pillow requests mistune
```

## 使用

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
| 页码 | "第×页 共×页"、页脚边距1.1cm |
| 列表 | 有序列表用（1）（2）（3）序号 |

