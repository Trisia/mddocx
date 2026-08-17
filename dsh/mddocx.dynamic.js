// mddocx —— 动态 Cordis 插件源码（会话级运行时形态）
//
// 动态 Cordis 插件只在运行时经 cordis_define / cordis_run 定义与激活，
// 源码不落盘（进程重启即消失）。本文件把 host / client 两个函数体作为
// 版本化参考保存：任一会话中把 host / client 字符串交给 cordis_define，
// 即可重建与本仓库一致的 mddocx 动态插件。
//
// 与 dsh/mddocx.mjs（组合插件，npm/预设分发形态）是同一能力的两种实现：
//   - 本文件（动态版）：沙箱兼容，无 Node 模块；经 ctx.shell / ctx.fs 执行，
//     工作目录取 exec.agent.session.header.cwd。
//   - dsh/mddocx.mjs（组合版）：真实 ESM，Node 内置模块直接可用。
//
// 重建方法（任意会话）：
//   1) 读取本文件得到 host / client 字符串
//   2) cordis_define(kind:"new", idPrefix:"mddoc", name:"mddocx", code:{ host, client })
//   3) cordis_run 激活；带 client 半区，首次需在界面批准一次
// 之后模型即可调用 mddocx 工具（text / path / output）。

export const host = `return {
  name: 'mddocx',
  apply(ctx) {
    const shell = ctx.get('shell')
    const fs = ctx.get('fs')
    const sandboxPolicy = ctx.get('sandboxPolicy')
    const fallbackRoot = (sandboxPolicy && typeof sandboxPolicy.workspaceRoot === 'string' && sandboxPolicy.workspaceRoot)
      ? sandboxPolicy.workspaceRoot
      : '/home/kkk/project/mddocx'

    function shq(s) {
      return "'" + String(s).replace(/'/g, "'\\\\''") + "'"
    }

    function sanitizeName(s) {
      const t = String(s).replace(/[\\\\/*?:"<>|]/g, '').trim()
      return t || '未命名文档'
    }

    const tool = harness.defineTool({
      name: 'mddocx',
      description: '将 Markdown 内容转换为符合中国学术论文排版规范的 Word 文档(.docx)：三线表、图题/表题自动编号、LaTeX 公式转 OMML、页码与页眉。当用户要求把 Markdown 文本或 .md 文件转成学术格式 Word 文档（md转word、markdown转docx、生成格式化文档）时使用。通过 text（直接粘贴的 Markdown 文本）或 path（工作区内的 .md 文件）提供输入，可选 output 指定输出路径；返回生成的 .docx 绝对路径。',
      parameters: {
        text: { type: 'string', description: '要转换的 Markdown 文本内容（直接粘贴）。与 path 二选一。' },
        path: { type: 'string', description: '工作区内要转换的 Markdown 文件路径（相对或绝对）。与 text 二选一。' },
        output: { type: 'string', description: '可选，输出 .docx 文件的路径（相对工作区）。默认输出到工作区根目录，文件名取自第一个 # 标题或输入文件名。' },
      },
      output: {
        schema: {
          type: 'object',
          properties: {
            outputPath: { type: 'string' },
            exitCode: { type: 'integer' },
            stdout: { type: 'string' },
            stderr: { type: 'string' },
          },
          additionalProperties: false,
        },
        render(args, value) {
          if (value.exitCode !== 0) {
            return [{ type: 'text', text: 'mddocx 转换失败 (exit code ' + value.exitCode + ')\\n[stderr]\\n' + (value.stderr || '') }]
          }
          return [{ type: 'text', text: '已生成 DOCX: ' + value.outputPath }]
        },
      },
      async execute(args, exec) {
        if (!shell || !fs) {
          throw new Error('mddocx 依赖的 shell/fs 服务不可用')
        }

        // 会话工作目录（与组合版一致；sandboxPolicy 仅作回退）
        const cwd = (exec.agent && exec.agent.session && exec.agent.session.header && exec.agent.session.header.cwd) || fallbackRoot
        const abs = (p) => (typeof p === 'string' && p.charAt(0) === '/') ? p : cwd + '/' + p

        // 输入
        let inputPath = null
        let tempTarget = null
        let baseName = '未命名文档'
        if (typeof args.text === 'string' && args.text.trim() !== '') {
          const tmpAbs = abs('.mddocx-input-' + Date.now() + '.md')
          tempTarget = await fs.resolve(tmpAbs)
          await fs.writeText(tempTarget, args.text)
          inputPath = fs.processPath(tempTarget)
          const m = args.text.match(/^#\\s+(.+)$/m)
          if (m && m[1]) baseName = sanitizeName(m[1])
        } else if (typeof args.path === 'string' && args.path.trim() !== '') {
          const inAbs = abs(args.path)
          const t = await fs.resolve(inAbs)
          const info = await fs.stat(t)
          if (!info) throw new Error('输入文件不存在: ' + inAbs)
          inputPath = fs.processPath(t)
          const lastSlash = inputPath.lastIndexOf('/')
          let base = lastSlash >= 0 ? inputPath.slice(lastSlash + 1) : inputPath
          base = base.replace(/\\.(md|markdown)$/i, '')
          if (base) baseName = sanitizeName(base)
        } else {
          throw new Error('mddocx 需要提供 text（粘贴的 Markdown 文本）或 path（工作区内 .md 文件路径）之一')
        }

        // 输出
        let outputPath = null
        if (typeof args.output === 'string' && args.output.trim() !== '') {
          const ot = await fs.resolve(abs(args.output))
          outputPath = fs.processPath(ot)
        } else {
          outputPath = abs(baseName + '.docx')
        }

        // 引擎 + venv（本机安装于用户级 npm prefix；bash 内解析 $HOME）
        const engine = '"$HOME/.npm-global/lib/node_modules/@cliven/mddocx/skills/mddoc/scripts/md2docx.py"'
        const cmd =
          'VENV="$HOME/.cache/mddocx/venv/bin/python"; ' +
          'if [ -x "$VENV" ]; then PY="$VENV"; else PY=python3; fi; ' +
          '"$PY" ' + engine + ' ' + shq(inputPath) + ' -o ' + shq(outputPath) +
          (tempTarget ? '; rm -f ' + shq(inputPath) : '')

        const policy = sandboxPolicy
          ? sandboxPolicy.resolve(exec.agent ? { session: exec.agent.session } : {})
          : undefined
        const result = await shell.run(shell.resolve({
          command: cmd,
          workdir: cwd,
          signal: exec.signal,
          ...(policy ? { sandboxPolicy: policy } : {}),
        }))

        if (result.aborted) {
          const err = new Error('mddocx 已取消')
          err.name = 'AbortError'
          throw err
        }

        const exitCode = typeof result.exitCode === 'number' ? result.exitCode : 1
        const stdout = result.stdout && typeof result.stdout.text === 'string' ? result.stdout.text : ''
        let stderr = result.stderr && typeof result.stderr.text === 'string' ? result.stderr.text : ''
        if (exitCode !== 0 && stderr.trim() === '') stderr = stdout || '(无输出)'

        return { outputPath: outputPath, exitCode: exitCode, stdout: stdout, stderr: stderr }
      },
    })

    return harness.registerTool(ctx, tool)
  },
}`

export const client = `return {
  apply(ctx) {
    // 客户端半区：让包在浏览器侧完整加载，Cordis 面板显示为 running。
  },
}`
