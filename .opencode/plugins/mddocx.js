// mddocx — Markdown 转学术格式 DOCX
// OpenCode 插件：注册 mddoc 技能并在会话启动时注入引导

import { readFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const PLUGIN_ROOT = join(__dirname, '..', '..');
const SKILL_MD_PATH = join(PLUGIN_ROOT, 'skills', 'mddoc', 'SKILL.md');

/** @param {import('opencode').Plugin} plugin */
export default async function mddocx(plugin) {
  // 注册技能目录
  plugin.config({
    skills: [join(PLUGIN_ROOT, 'skills')],
  });

  // 会话启动时注入 mddoc 技能引导
  plugin.hook('experimental.chat.messages.transform', async (messages) => {
    try {
      const skillContent = readFileSync(SKILL_MD_PATH, 'utf-8');
      const bootstrap = `<EXTREMELY_IMPORTANT>
本会话已安装 mddocx 插件。以下是 mddoc 技能：

${skillContent}
</EXTREMELY_IMPORTANT>`;

      // 注入到第一条用户消息之前
      const systemMsg = messages.find(m => m.role === 'system');
      if (systemMsg) {
        systemMsg.content = bootstrap + '\n\n' + (systemMsg.content || '');
      } else {
        messages.unshift({ role: 'system', content: bootstrap });
      }
    } catch (e) {
      console.warn('[mddocx] 无法注入技能引导:', e.message);
    }
    return messages;
  });
}
