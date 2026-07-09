# mddocx 贡献指南

## 项目概述

mddocx 是一个 Claude Code 插件，将 Markdown 内容转换为符合中国学术论文排版规范的 Word 文档。

## 技能开发

技能位于 `skills/mddoc/`：
- `SKILL.md` — 技能定义，包含完整的格式规范和 python-docx 代码示例
- `scripts/md2docx.py` — 内置 Python 转换脚本
- `evals/` — 测试用例

## 修改技能

1. 修改 `skills/mddoc/SKILL.md` 中的规范或代码
2. 同步更新 `skills/mddoc/scripts/md2docx.py` 中的实现
3. 运行测试验证格式输出正确

## 发布

1. 更新 `.claude-plugin/plugin.json` 和 `package.json` 中的版本号
2. `git tag vX.Y.Z`
3. `git push --tags`

## 测试

```bash
# 命令行测试
python skills/mddoc/scripts/md2docx.py skills/mddoc/evals/test-sample.md -o /tmp/test.docx

# 验证格式
python -c "
from docx import Document
doc = Document('/tmp/test.docx')
# 检查标题、outline level、表格边框等
"
```
