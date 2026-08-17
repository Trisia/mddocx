# mddocx — DeepSeek Harness (DSH) 原生插件

这是 mddocx 的 DSH **插件**：npm 包 `@cliven/mddocx` 内的一个 Cordis 插件文件（`dsh/mddocx.mjs`），向任意 agent 预设注册原生 `mddocx` 模型工具（Markdown → 学术格式 DOCX，一步完成，无需手动运行 Python 脚本）。

它**不是一个完整的 agent 预设**，只是一行组合：装到哪个预设，那个预设的会话就能用 `mddocx`。

## 安装

1. 安装 npm 包（含转换引擎 `skills/mddoc/scripts/` 与插件文件）：

   ```bash
   npm install -g @cliven/mddocx
   ```

2. 查看插件文件的绝对路径：

   ```bash
   echo "$(npm root -g)/@cliven/mddocx/dsh/mddocx.mjs"
   ```

3. 在你预设的 `~/.dsh/.agent-presets/<你的预设>/agent.cordis.yml` 末尾追加一行（`name` 填入上一步输出的绝对路径）：

   ```yaml
   - id: mddocx
     name: '/usr/lib/node_modules/@cliven/mddocx/dsh/mddocx.mjs'
   ```

重新开一个使用该预设的会话即可生效。

> DSH loader 支持组合行 `name` 为绝对路径（挂载时转 `file:` URL 导入）。引擎脚本随 npm 包分发，插件从自身位置解析（`import.meta.url`），全局安装即自包含；也支持 `MDDOCX_SKILL_PATH` 环境变量覆盖与 `~/.claude/skills/mddoc` 技能安装回退。

## 使用

模型会自动调用 `mddocx`，入参：

| 参数 | 说明 |
|------|------|
| `text` | 直接粘贴的 Markdown 文本（与 `path` 二选一） |
| `path` | 工作区内要转换的 `.md` 文件路径（与 `text` 二选一） |
| `output` | 可选，输出 `.docx` 路径（相对工作区）；默认输出到工作区根目录，文件名取自第一个 `#` 标题 |

返回生成的 `.docx` 绝对路径。

## 说明

- **引擎单一来源**：插件驱动包内 `skills/mddoc/scripts/md2docx.py`；缺失时工具会报错并提示安装命令。
- **Python 环境**：优先使用专用虚拟环境 `~/.cache/mddocx/venv`，缺失时自动运行包内 `setup_env.py` 自检（幂等）。
- **技能位置覆盖**：环境变量 `MDDOCX_SKILL_PATH` 指向技能目录可覆盖包内引擎。

## 动态插件形态（会话级，Cordis 面板显示用）

除组合插件外，同一能力还有一个**动态 Cordis 插件**形态：在某个会话内临时定义并激活，显示在 GUI 的 Cordis 插件面板（侧边栏底部「Cordis Plugin」按钮），模型可直接调用 `mddocx` 工具。动态插件**只在进程内存在**，重启即消失，因此其 host/client 源码以参考文件保存在 **`dsh/mddocx.dynamic.js`**（导出 `host` / `client` 两个函数体字符串）。

任一会话重建方法：

1. 读取 `dsh/mddocx.dynamic.js`，取 `host` / `client` 字符串；
2. `cordis_define(kind:"new", idPrefix:"mddoc", name:"mddocx", code:{ host, client })`；
3. `cordis_run` 激活（带 client 半区，首次需在界面批准一次）。

> 动态版与组合版是两种实现：动态版沙箱兼容（无 Node 模块，走 `ctx.shell`/`ctx.fs`，工作目录取 `exec.agent.session.header.cwd`）；组合版（`dsh/mddocx.mjs`）是真实 ESM，用于 npm 包 / 预设行分发。改动时注意两处同步。
