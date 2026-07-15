# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

mddocx 是一个多平台 AI Agent 插件，将 Markdown 转换为符合中国学术论文排版规范的 Word 文档。支持 Claude Code、Codex、Cursor、OpenCode 四个平台，同时发布为 npm 包。

核心功能：Markdown 解析 → LaTeX 公式转 OMML → 学术格式 DOCX（三线表、图题/表题自动编号、页码、页眉）。

## 架构

```
输入 (Markdown 文件/文本)
  ↓
bin/mddocx.js              ← npm CLI 入口，调用 Python 脚本
  ↓
skills/mddoc/scripts/md2docx.py  ← 核心转换引擎（~1600行）
  ├── mistune 解析 Markdown AST
  ├── _LatexParser 类：LaTeX → OMML XML 递归下降解析器
  ├── 图片下载/嵌入（PIL + requests）
  └── python-docx 生成 DOCX（三线表、OMML 公式、页码字段）
  ↓
skills/mddoc/SKILL.md      ← 技能定义（格式规范 + python-docx 代码示例）
  ↓
各平台插件层：
  .claude-plugin/plugin.json   → /plugin install 注册 + hooks/SessionStart
  hooks/session-start          → 会话启动时将 SKILL.md 注入 Claude Code 上下文
  .opencode/plugins/mddocx.js  → OpenCode 插件，chat.messages.transform 注入
  .codex-plugin/plugin.json    → Codex 插件注册
  .cursor-plugin/plugin.json   → Cursor 插件注册
```

**关键设计**：
- 四个平台共享同一份 `skills/mddoc/`，各平台 `plugin.json` 指向 `./skills/`
- SessionStart hook 和 OpenCode transform hook 都在会话启动时将 SKILL.md 全文注入 Agent 上下文，让 Agent 直接按规范生成代码
- npm 包 `@cliven/mddocx` 包含所有平台配置，`npx @cliven/mddocx` 可直接使用
- 虚拟环境位于 `.venv/`，不应提交

## 常用命令

```bash
# 安装依赖（首次）
python -m venv .venv && source .venv/bin/activate
pip install python-docx Pillow requests mistune

# 转换测试
python skills/mddoc/scripts/md2docx.py skills/mddoc/evals/test-sample.md -o /tmp/test.docx

# 转换 demo（发布前冒烟测试）
python skills/mddoc/scripts/md2docx.py examples/demo.md -o /tmp/test.docx

# 验证生成的 docx 格式
python -c "
from docx import Document
doc = Document('/tmp/test.docx')
for p in doc.paragraphs:
    ol = p._element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}outlineLvl')
    lvl = ol.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if ol is not None else '-'
    print(f'[{lvl}] {p.text[:80]}')
"
```

## 版本号位置

发布时需同步更新以下 **6 个文件**的版本号：

| 文件 | 字段 |
|------|------|
| `.claude-plugin/plugin.json` | `"version"` |
| `.codex-plugin/plugin.json` | `"version"` |
| `.cursor-plugin/plugin.json` | `"version"` |
| `package.json` | `"version"` |
| `package-lock.json` | `"version"`（两处） |
| `skills/mddoc/SKILL.md` | frontmatter `version:` |

## 发布流程

1. 更新上述 6 个文件的版本号
2. `git add` 并提交（commit message: `chore: bump version to x.y.z`）
3. `git tag vX.Y.Z`
4. `git push --tags`（需用户确认）

推送 tag 后 CI（`.github/workflows/release.yml`）自动执行：
- 冒烟测试（转换 `examples/demo.md`）
- 打包 `mddoc.skill.zip` 和源码包
- 发布到 npm
- 同步到 ClawHub
- 创建 GitHub Release

## 代码修改同步

修改格式规范时，**两处必须同步**：
1. `skills/mddoc/SKILL.md` — 格式规范和 python-docx 代码示例（AI Agent 读取此文件生成代码）
2. `skills/mddoc/scripts/md2docx.py` — 内置转换脚本（命令行直接调用）

二者逻辑须保持一致，否则命令行结果和 Agent 生成结果会不同。

## 测试

```bash
# 冒烟测试
python skills/mddoc/scripts/md2docx.py skills/mddoc/evals/test-sample.md -o /tmp/test.docx

# evals 定义在 skills/mddoc/evals/evals.json（3 个用例）
```
