#!/usr/bin/env node
// mddocx CLI — 调用 Python 转换脚本

import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { existsSync } from 'node:fs';
import { homedir } from 'node:os';

const __dirname = dirname(fileURLToPath(import.meta.url));
const script = join(__dirname, '..', 'skills', 'mddoc', 'scripts', 'md2docx.py');
const setupEnv = join(__dirname, '..', 'skills', 'mddoc', 'scripts', 'setup_env.py');

if (!existsSync(script)) {
  console.error('错误: 找不到转换脚本', script);
  process.exit(1);
}

// 专用虚拟环境解释器（Linux/macOS: ~/.cache/mddocx/venv；Windows: %LOCALAPPDATA%/mddocx/venv）
const cacheDir = process.platform === 'win32'
  ? (process.env.LOCALAPPDATA || homedir())
  : (process.env.XDG_CACHE_HOME || join(homedir(), '.cache'));
const venvPython = join(cacheDir, 'mddocx', 'venv',
  process.platform === 'win32' ? 'Scripts' : 'bin',
  process.platform === 'win32' ? 'python.exe' : 'python');

// 优先用专用 venv 解释器，不存在则回退系统 python3
const pythonCmd = existsSync(venvPython) ? venvPython : 'python3';

const args = [script, ...process.argv.slice(2)];

const child = spawn(pythonCmd, args, { stdio: 'inherit' });

child.on('close', (code) => {
  if (code !== 0) {
    console.error(`\n提示: 依赖缺失或环境未就绪，请先执行环境自检: python3 ${setupEnv}`);
  }
  process.exit(code);
});
