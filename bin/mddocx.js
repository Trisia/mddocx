#!/usr/bin/env node
// mddocx CLI — 调用 Python 转换脚本

import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { existsSync } from 'node:fs';

const __dirname = dirname(fileURLToPath(import.meta.url));
const script = join(__dirname, '..', 'skills', 'mddoc', 'scripts', 'md2docx.py');

if (!existsSync(script)) {
  console.error('错误: 找不到转换脚本', script);
  process.exit(1);
}

// 优先找 venv python，否则用系统 python3
const pythonCmd = process.env.MDDOCX_PYTHON || 'python3';

const args = [script, ...process.argv.slice(2)];

const child = spawn(pythonCmd, args, { stdio: 'inherit' });

child.on('close', (code) => {
  if (code !== 0) {
    console.error(`\n提示: 确认已安装依赖: pip install python-docx Pillow requests mistune`);
  }
  process.exit(code);
});
