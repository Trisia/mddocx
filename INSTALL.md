# 安装 mddocx

复制下面代码块，粘贴给任意智能体即可自动完成安装。

---

## npm（推荐命令行用户）

```bash
npm install -g mddocx        # 全局安装
npx mddocx paper.md          # 或直接 npx 运行（无需安装）
```

发布：

```bash
npm login
npm publish --access public
```

---

## Claude Code

```copy
请帮我安装 mddocx 插件：

1. 克隆并安装技能：
git clone https://github.com/Trisia/mddocx /tmp/mddocx
mkdir -p ~/.claude/skills/mddoc
cp -r /tmp/mddocx/skills/mddoc/* ~/.claude/skills/mddoc/
rm -rf /tmp/mddocx

2. 安装 Python 依赖：
python3 -m venv ~/.claude/venvs/mddocx
~/.claude/venvs/mddocx/bin/pip install python-docx Pillow requests mistune

3. 验证安装：
python ~/.claude/skills/mddoc/scripts/md2docx.py --help
```

手动安装：

```bash
# 插件方式（含 SessionStart hook）
git clone https://github.com/Trisia/mddocx ~/.claude/plugins/mddocx
# 在 Claude Code 中执行: /plugin install ~/.claude/plugins/mddocx

# 或仅安装技能
git clone https://github.com/Trisia/mddocx /tmp/mddocx
mkdir -p ~/.claude/skills/mddoc
cp -r /tmp/mddocx/skills/mddoc/* ~/.claude/skills/mddoc/
```

---

## Codex (OpenAI)

```copy
请帮我安装 mddocx 插件：

git clone https://github.com/Trisia/mddocx /tmp/mddocx
mkdir -p ~/.codex/skills/mddoc
cp -r /tmp/mddocx/skills/mddoc/* ~/.codex/skills/mddoc/
rm -rf /tmp/mddocx
pip install python-docx Pillow requests mistune
```

手动安装：

```bash
/plugin install git:https://github.com/Trisia/mddocx
```

---

## OpenCode

```copy
请帮我在 OpenCode 项目中安装 mddocx：

1. 在 opencode.json 的 plugins 数组中添加：
"git:https://github.com/Trisia/mddocx"

2. 安装 Python 依赖：
python3 -m venv .venv
source .venv/bin/activate
pip install python-docx Pillow requests mistune
```

手动安装：按 [.opencode/INSTALL.md](.opencode/INSTALL.md) 操作。

---

## Cursor

```copy
请帮我安装 mddocx：

git clone https://github.com/Trisia/mddocx /tmp/mddocx
mkdir -p ~/.cursor/skills/mddoc
cp -r /tmp/mddocx/skills/mddoc/* ~/.cursor/skills/mddoc/
rm -rf /tmp/mddocx
pip install python-docx Pillow requests mistune
```

---

## 通用（仅技能，无插件hook）

```copy
请帮我安装 mddocx 技能：

git clone https://github.com/Trisia/mddocx /tmp/mddocx
mkdir -p ~/.claude/skills/mddoc
cp -r /tmp/mddocx/skills/mddoc/* ~/.claude/skills/mddoc/
rm -rf /tmp/mddocx
pip install python-docx Pillow requests mistune
```

---

## 依赖

所有平台均需 Python 依赖：

```bash
# 推荐虚拟环境
python3 -m venv .venv
source .venv/bin/activate  # Linux/macOS

pip install python-docx Pillow requests mistune
```

> 若提示 `externally-managed-environment`，使用虚拟环境或 `--break-system-packages`。
