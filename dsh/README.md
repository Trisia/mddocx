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
