// Native mddoc conversion tool for DeepSeek Harness (DSH).
//
// Registers the `mddocx` model tool: Markdown → academic DOCX via the
// mddocx package's conversion engine. The engine (`skills/mddoc/scripts/`)
// ships inside the same npm package (`@cliven/mddocx`), so this plugin resolves
// it RELATIVE TO ITSELF (`import.meta.url`) — a global install is fully
// self-contained. Resolution order: `MDDOCX_SKILL_PATH` env override → the
// package's own `skills/mddoc` → the mddoc skill install at
// `~/.claude/skills/mddoc`.
//
// Registered as a composition row with an ABSOLUTE path to this file (the DSH
// loader supports absolute paths), so this module is a real ESM plugin, not a
// sandboxed dynamic package: Node builtins and `process` are available.

import { execFile } from 'node:child_process'
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { homedir, tmpdir } from 'node:os'
import { basename, dirname, isAbsolute, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const name = 'mddocx'
const inject = ['tools', 'systemPrompt']

/** The skill directory that ships inside this package (`<pkg>/skills/mddoc`). */
const pkgSkillDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'skills', 'mddoc')

/** Resolve the conversion-engine directory: env override → package → skill install. */
function resolveScriptDir() {
  if (process.env.MDDOCX_SKILL_PATH) return process.env.MDDOCX_SKILL_PATH
  if (existsSync(join(pkgSkillDir, 'scripts', 'md2docx.py'))) return pkgSkillDir
  return join(homedir(), '.claude', 'skills', 'mddoc')
}

/** Remove characters illegal in a Windows filename; empty result → 未命名文档. */
function sanitizeName(s) {
  const t = String(s).replace(/[\\/*?:"<>|]/g, '').trim()
  return t || '未命名文档'
}

/** Run one command, always resolving with { exitCode, stdout, stderr }. */
function run(cmd, args, opts) {
  return new Promise((resolve) => {
    execFile(cmd, args, opts, (err, stdout, stderr) => {
      resolve({
        exitCode: err ? (typeof err.code === 'number' ? err.code : 1) : 0,
        stdout: String(stdout || ''),
        stderr: String(stderr || ''),
      })
    })
  })
}

function apply(ctx) {
  ctx.systemPrompt.section({
    name: 'tool:mddocx',
    order: 130,
    text: '当用户要求把 Markdown 转成 Word / 学术格式文档（md转word、markdown转docx、生成格式化文档）时，直接调用 mddocx 工具，不要再手动运行 Python 脚本。',
  })

  ctx.tools.register({
    name: 'mddocx',
    description: '将 Markdown 内容转换为符合中国学术论文排版规范的 Word 文档(.docx)：三线表、图题/表题自动编号、LaTeX 公式转 OMML、页码与页眉。当用户要求把 Markdown 文本或 .md 文件转成学术格式 Word 文档（md转word、markdown转docx、生成格式化文档）时使用。通过 text（直接粘贴的 Markdown 文本）或 path（工作区内的 .md 文件）提供输入，可选 output 指定输出路径；返回生成的 .docx 绝对路径。',
    parameters: {
      type: 'object',
      properties: {
        text: { type: 'string', description: '要转换的 Markdown 文本内容（直接粘贴）。与 path 二选一。' },
        path: { type: 'string', description: '工作区内要转换的 Markdown 文件路径（相对或绝对）。与 text 二选一。' },
        output: { type: 'string', description: '可选，输出 .docx 文件的路径（相对工作区）。默认输出到工作区根目录，文件名取自第一个 # 标题或输入文件名。' },
      },
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
          return [{ type: 'text', text: 'mddoc 转换失败 (exit code ' + value.exitCode + ')\n[stderr]\n' + (value.stderr || '') }]
        }
        return [{ type: 'text', text: '已生成 DOCX: ' + value.outputPath }]
      },
    },
    async execute(args, exec) {
      // 1) Locate the conversion engine (package-relative → skill install).
      const skillDir = resolveScriptDir()
      const scriptPath = join(skillDir, 'scripts', 'md2docx.py')
      if (!existsSync(scriptPath)) {
        throw new Error('找不到 mddoc 转换脚本: ' + scriptPath + '。请检查 npm 包安装（npm install -g @cliven/mddocx）或设置 MDDOCX_SKILL_PATH 指向技能目录')
      }

      // 2) Resolve the session working directory.
      const cwd = (exec.agent && exec.agent.session && exec.agent.session.header && exec.agent.session.header.cwd) || process.cwd()

      // 3) Input: pasted text → temp file; file path → use directly.
      let inputPath = null
      let tmpDir = null
      let baseName = '未命名文档'

      if (typeof args.text === 'string' && args.text.trim() !== '') {
        tmpDir = mkdtempSync(join(tmpdir(), 'mddocx-'))
        inputPath = join(tmpDir, 'input.md')
        writeFileSync(inputPath, args.text, 'utf-8')
        const m = args.text.match(/^#\s+(.+)$/m)
        if (m && m[1]) baseName = sanitizeName(m[1])
      } else if (typeof args.path === 'string' && args.path.trim() !== '') {
        inputPath = isAbsolute(args.path) ? args.path : join(cwd, args.path)
        const base = basename(inputPath).replace(/\.(md|markdown)$/i, '')
        if (base) baseName = sanitizeName(base)
      } else {
        throw new Error('mddocx 需要提供 text（粘贴的 Markdown 文本）或 path（工作区内 .md 文件路径）之一')
      }
      if (!existsSync(inputPath)) {
        throw new Error('输入文件不存在: ' + inputPath)
      }

      // 4) Output path: explicit or derived from title / input name.
      let outputPath
      if (typeof args.output === 'string' && args.output.trim() !== '') {
        outputPath = isAbsolute(args.output) ? args.output : join(cwd, args.output)
      } else {
        outputPath = join(cwd, baseName + '.docx')
      }
      mkdirSync(dirname(outputPath), { recursive: true })

      // 5) Python interpreter: dedicated venv, auto-setup when missing.
      const cacheDir = process.platform === 'win32'
        ? (process.env.LOCALAPPDATA || homedir())
        : (process.env.XDG_CACHE_HOME || join(homedir(), '.cache'))
      const venvPython = join(cacheDir, 'mddocx', 'venv',
        process.platform === 'win32' ? 'Scripts' : 'bin',
        process.platform === 'win32' ? 'python.exe' : 'python')

      let python = venvPython
      if (!existsSync(venvPython)) {
        const setup = join(skillDir, 'scripts', 'setup_env.py')
        if (existsSync(setup)) {
          await run('python3', [setup], { cwd, signal: exec.signal })
        }
        if (!existsSync(venvPython)) python = 'python3'
      }

      // 6) Convert.
      const result = await run(python, [scriptPath, inputPath, '-o', outputPath], { cwd, signal: exec.signal })
      if (tmpDir) rmSync(tmpDir, { recursive: true, force: true })

      let stderr = result.stderr
      if (result.exitCode !== 0 && stderr.trim() === '') stderr = result.stdout || '(无输出)'
      return { outputPath, exitCode: result.exitCode, stdout: result.stdout, stderr }
    },
  })
}

export { apply, inject, name }
