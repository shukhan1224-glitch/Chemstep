/**
 * 把 apps-script/ 底下的 .gs 合并成单一档案 ALL-IN-ONE.gs,
 * 让使用者只需要复制贴上一次。
 *
 * 用法: node tools/build-bundle.js
 */
const fs = require('fs');
const path = require('path');

const SRC_DIR = path.join(__dirname, '..', 'apps-script');
const OUT = path.join(SRC_DIR, 'ALL-IN-ONE.gs');
const BUNDLE_NAME = 'ALL-IN-ONE.gs';

const files = fs.readdirSync(SRC_DIR)
  .filter(f => f.endsWith('.gs') && f !== BUNDLE_NAME)
  .sort();

if (!files.length) {
  console.error('找不到任何 .gs 档案');
  process.exit(1);
}

const header = [
  '/**',
  ' * ══════════════════════════════════════════════════════════════',
  ' *  线上补习点名 + 收费提醒系统  —— 合并版(单一档案)',
  ' * ══════════════════════════════════════════════════════════════',
  ' *',
  ' *  这个档案由 tools/build-bundle.js 自动产生,请勿直接编辑。',
  ' *  要改程式码请改 apps-script/ 底下的个别档案,再重新执行:',
  ' *      node tools/build-bundle.js',
  ' *',
  ' *  安装:把整个档案内容贴进 Apps Script 编辑器,储存,',
  ' *        重新整理试算表,然后点 📚 补习管理 → 🔧 初始化 / 修复表格。',
  ' *',
  ' *  合并自:' + files.join('、'),
  ' */',
  ''
].join('\n');

const body = files.map(f => {
  const bar = '─'.repeat(Math.max(60 - f.length, 3));
  return '/* ── ' + f + ' ' + bar + ' */\n\n' + fs.readFileSync(path.join(SRC_DIR, f), 'utf8').trim() + '\n';
}).join('\n\n');

fs.writeFileSync(OUT, header + '\n' + body);

// 基本防呆:合并后不该有重复的函式或常数名称
const names = {};
const dupes = [];
const re = /^(?:function\s+([A-Za-z0-9_$]+)|const\s+([A-Z][A-Z0-9_]*)\s*=)/gm;
let m;
const merged = fs.readFileSync(OUT, 'utf8');
while ((m = re.exec(merged)) !== null) {
  const name = m[1] || m[2];
  if (names[name]) dupes.push(name); else names[name] = true;
}
if (dupes.length) {
  console.error('❌ 合并后有重复定义:' + [...new Set(dupes)].join(', '));
  process.exit(1);
}

console.log('✅ 已产生 ' + path.relative(process.cwd(), OUT) +
            '(' + files.length + ' 个档案,' + merged.split('\n').length + ' 行,' +
            Object.keys(names).length + ' 个顶层定义)');
