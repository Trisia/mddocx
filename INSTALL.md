# 安装 mddocx

## 快速安装（推荐）

### npx skills — 全平台技能安装

将 mddoc 技能安装到 Claude Code / Codex / Cursor / OpenCode：

```bash
npx skills add Trisia/mddocx -g -a claude-code codex cursor opencode -y
```

装到各平台 `~/.claude/skills/mddoc` 等用户目录；升级见下方「升级」。

### npm — 全平台通用

```bash
npm install -g @cliven/mddocx
```

安装后可用 `mddocx` 命令，或免安装直接用 `npx @cliven/mddocx`。

> npm 包内含各平台插件配置，可通过插件机制自动注册技能。

安装后按下方「依赖」一节执行环境自检。

---

## 各平台专属方式

### Claude Code

```bash
# PluginHub（推荐）
/plugin marketplace add https://www.claudepluginhub.com/api/plugins/trisia-mddocx/marketplace.json
/plugin install trisia-mddocx@cpd-trisia-mddocx

# 手动（插件，含 SessionStart hook）
git clone --depth 1 https://github.com/Trisia/mddocx ~/.claude/plugins/mddocx
/plugin install ~/.claude/plugins/mddocx
```

### Codex

```bash
/plugin install git:https://github.com/Trisia/mddocx
```

### OpenCode

在 `opencode.json` 的 `plugins` 数组添加 `"git:https://github.com/Trisia/mddocx"`，详见 [.opencode/INSTALL.md](.opencode/INSTALL.md)。

### Cursor

```bash
git clone --depth 1 https://github.com/Trisia/mddocx ~/.cursor/plugins/mddocx
```

### ClawHub（openclaw CLI）

```bash
openclaw skills install @trisia/mddoc
```

---

## 升级

| 方式 | 命令 |
|------|------|
| npx skills | `npx skills update` |
| npm | `npm update -g @cliven/mddocx` |
| Claude PluginHub | `/plugin install trisia-mddocx@cpd-trisia-mddocx`（重新安装即更新） |
| ClawHub | `openclaw skills install @trisia/mddoc`（重新安装拉取最新版） |
| 插件（git 安装的） | 进入插件目录 `git pull` 后重启 |

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
