// Word 与 Excel 共用 chat.js 的繁忙态协议：chat.send 要在整轮完成后才返回。
// Word 如果提前返回 started，面板会立即清掉“正在处理”气泡，用户看不到任何进展。

import fs from 'node:fs';

const source = fs.readFileSync('src/ChatWord.AddIn/Bridge/WordBridge.cs', 'utf8');

let passed = 0;
let failed = 0;

function check(label, condition, detail = '') {
  if (condition) {
    passed += 1;
    console.log(`  通过  ${label}`);
    return;
  }

  failed += 1;
  console.log(`  失败  ${label}${detail ? `：${detail}` : ''}`);
}

console.log('检查 Word chat.send 繁忙态协议：');

check('Word chat.send 是异步完成协议',
  /private\s+async\s+Task<object>\s+SendAsync\s*\(/.test(source));
check('Word chat.send 等待文档任务结束',
  /await\s+Task\.Run\([\s\S]*?_agent\.RunAsync/.test(source));
check('Word 正常结束返回 completed',
  /return\s+new\s*\{\s*completed\s*=\s*true\s*\}/.test(source));
check('Word 不再返回即时 started',
  !/\{\s*started\s*=\s*true\s*\}/.test(source));
check('Word 取消仍推送 stopped 状态',
  /Kind\s*=\s*"stopped"/.test(source) && /stopped\s*=\s*true/.test(source));

console.log(`=== Word chat.send 繁忙态协议：通过 ${passed}，失败 ${failed} ===`);
if (failed > 0) {
  process.exitCode = 1;
}
