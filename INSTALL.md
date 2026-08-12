# 安装 mddocx

复制下面代码块，粘贴给任意智能体即可自动完成安装。

---


## npm（推荐，全平台通用）

```copy
请帮我用 npm 安装 mddocx：

npm install -g @cliven/mddocx
```

Python 依赖由 `setup_env.py` 自动安装，见下方「依赖」一节。

安装后即可使用 `mddocx` 命令，或直接用 `npx @cliven/mddocx` 无需安装。

> npm 包内含 Claude Code / Codex / OpenCode / Cursor 的插件配置，各平台可通过插件机制自动注册技能。

---

## Claude PluginHub（Claude Code 专属）

在 Claude Code 中执行（添加 marketplace 仅需一次）：

```copy
/plugin marketplace add https://www.claudepluginhub.com/api/plugins/trisia-mddocx/marketplace.json
/plugin install trisia-mddocx@cpd-trisia-mddocx
```

安装后按下方「依赖」一节执行环境自检。

---

## Claude Code

```copy
请帮我安装 mddocx 插件：

1. 克隆并安装技能：
git clone --depth 1 https://github.com/Trisia/mddocx /tmp/mddocx
mkdir -p ~/.claude/skills/mddoc
cp -r /tmp/mddocx/skills/mddoc/* ~/.claude/skills/mddoc/
rm -rf /tmp/mddocx

2. 准备 Python 依赖（自动）：
python3 ~/.claude/skills/mddoc/scripts/setup_env.py

3. 验证安装：
python ~/.claude/skills/mddoc/scripts/md2docx.py --help
```

手动安装：

```bash
# 插件方式（含 SessionStart hook）
git clone --depth 1 https://github.com/Trisia/mddocx ~/.claude/plugins/mddocx
# 在 Claude Code 中执行: /plugin install ~/.claude/plugins/mddocx

# 或仅安装技能
git clone --depth 1 https://github.com/Trisia/mddocx /tmp/mddocx
mkdir -p ~/.claude/skills/mddoc
cp -r /tmp/mddocx/skills/mddoc/* ~/.claude/skills/mddoc/
```

---

## Codex (OpenAI)

```copy
请帮我安装 mddocx 插件：

git clone --depth 1 https://github.com/Trisia/mddocx /tmp/mddocx
mkdir -p ~/.codex/skills/mddoc
cp -r /tmp/mddocx/skills/mddoc/* ~/.codex/skills/mddoc/
rm -rf /tmp/mddocx
python3 ~/.codex/skills/mddoc/scripts/setup_env.py
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

2. 准备 Python 依赖（自动）：
python3 skills/mddoc/scripts/setup_env.py
```

手动安装：按 [.opencode/INSTALL.md](.opencode/INSTALL.md) 操作。

---

## Cursor

```copy
请帮我安装 mddocx：

git clone --depth 1 https://github.com/Trisia/mddocx /tmp/mddocx
mkdir -p ~/.cursor/skills/mddoc
cp -r /tmp/mddocx/skills/mddoc/* ~/.cursor/skills/mddoc/
rm -rf /tmp/mddocx
python3 ~/.cursor/skills/mddoc/scripts/setup_env.py
```

手动安装：

```bash
# 克隆仓库
git clone --depth 1 https://github.com/Trisia/mddocx ~/.cursor/plugins/mddocx

# 准备 Python 依赖（自动）
python3 ~/.cursor/plugins/mddocx/skills/mddoc/scripts/setup_env.py
```

---

## 通用（仅技能，无插件hook）

```copy
请帮我安装 mddocx 技能：

git clone --depth 1 https://github.com/Trisia/mddocx /tmp/mddocx
mkdir -p ~/.claude/skills/mddoc
cp -r /tmp/mddocx/skills/mddoc/* ~/.claude/skills/mddoc/
rm -rf /tmp/mddocx
python3 ~/.claude/skills/mddoc/scripts/setup_env.py
```

---

## ClawHub（openclaw CLI）

```copy
请帮我用 ClawHub 安装 mddocx：

openclaw skills install @trisia/mddoc
```

安装后按下方「依赖」一节执行环境自检。

---

## 升级

### npm（全平台通用）

```bash
npm update -g @cliven/mddocx       # 全局安装升级
npx @cliven/mddocx@latest          # npx 始终使用最新版
```

### Claude Code

```bash
# 插件方式 → git pull
cd ~/.claude/plugins/mddocx && git pull

# 仅技能 → 重新复制
git clone --depth 1 https://github.com/Trisia/mddocx /tmp/mddocx
cp -rf /tmp/mddocx/skills/mddoc ~/.claude/skills/mddoc
rm -rf /tmp/mddocx
```

### Codex

```bash
cd ~/.codex/plugins/mddocx && git pull
# 或
/plugin update mddocx
```

### OpenCode

重启 OpenCode 即可自动拉取插件最新版本。或手动：

```bash
# 删除缓存后重启
rm -rf ~/.opencode/plugins/mddocx
```

### Cursor

```bash
cd ~/.cursor/plugins/mddocx && git pull
# 或重新克隆
git clone --depth 1 https://github.com/Trisia/mddocx /tmp/mddocx
cp -rf /tmp/mddocx/skills/mddoc ~/.cursor/skills/mddoc
rm -rf /tmp/mddocx
```

### Claude PluginHub

在 Claude Code 中执行：

```copy
/plugin install trisia-mddocx@cpd-trisia-mddocx   # 重新安装即更新
```

### ClawHub

```bash
openclaw skills install @trisia/mddoc   # 重新安装拉取最新版
```

### Python 依赖

```bash
python3 skills/mddoc/scripts/setup_env.py   # 幂等,自动补齐缺失依赖
```

---

## 依赖

所有平台均需 Python 依赖，装在专用虚拟环境 `~/.cache/mddocx/venv`（Windows: `%LOCALAPPDATA%/mddocx/venv`），不污染用户项目目录。**无需手动 pip 安装**，执行环境自检：环境就绪则零操作，**仅当环境缺失或依赖缺失时才创建/安装**（幂等，可重复执行）：

依赖清单：`python-docx`、`Pillow`、`requests`、`mistune`。pip 安装使用清华镜像源 `https://pypi.tuna.tsinghua.edu.cn/simple`。

```bash
# 环境自检与准备（一次性）
python3 skills/mddoc/scripts/setup_env.py
```

> 若提示 `externally-managed-environment`，请使用 `setup_env.py`（自动使用 venv）而非系统 pip。
