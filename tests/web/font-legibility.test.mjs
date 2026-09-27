import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = fileURLToPath(new URL('.', import.meta.url));
const css = readFileSync(join(here, '..', '..', 'src', 'web', 'styles', 'app.css'), 'utf8')
  .replace(/\r\n/g, '\n');

let passed = 0;
let failed = 0;

function check(label, condition, detail = '') {
  if (condition) {
    passed += 1;
    console.log(`  通过  ${label}`);
  } else {
    failed += 1;
    console.log(`  失败  ${label}${detail ? `：${detail}` : ''}`);
  }
}

function block(selector) {
  const start = css.indexOf(selector);
  if (start < 0) { return ''; }
  const open = css.indexOf('{', start);
  const close = css.indexOf('}', open);
  return open >= 0 && close >= 0 ? css.slice(open + 1, close) : '';
}

function has(blockText, pattern) {
  return pattern.test(blockText.replace(/\/\*[\s\S]*?\*\//g, ''));
}

const body = block('body {\n  /*');
const version = block('.app-version {');
const summary = block('.advanced-summary {');
const hint = block('.field-hint {');
const helpBody = block('.field-help-body {');
const workBuddy = block('.workbuddy-status,');
const workBuddyCheckin = block('.workbuddy-checkin {');
const notice = block('.notice {');
const pickerTrigger = block('.picker-trigger {');

console.log('检查字体清晰度基线：');
console.log('');

check(
  '拉丁字母与数字优先使用 Segoe UI',
  /font-family:\s*"Segoe UI"[\s\S]*"Microsoft YaHei UI"/.test(body),
);
check(
  '正文使用 13px / 19px，避免 100% DPI 下过小',
  has(body, /font-size:\s*13px/) && has(body, /line-height:\s*19px/),
);
check(
  '版本号使用清晰的 UI 字体栈且字号可读',
  /font-family:\s*"Segoe UI"[\s\S]*"Microsoft YaHei UI"/.test(version) && has(version, /font-size:\s*12px/),
);
check(
  '高级参数标题提升到 12px / 17px',
  has(summary, /font-size:\s*12px/) && has(summary, /line-height:\s*17px/),
);
check(
  '设置页提示提升到 12px / 17px',
  has(hint, /font-size:\s*12px/) && has(hint, /line-height:\s*17px/) &&
    has(helpBody, /font-size:\s*12px/) && has(helpBody, /line-height:\s*17px/),
);
check(
  'WorkBuddy 状态芯片提升到 12px / 17px',
  has(workBuddy, /font-size:\s*12px/) && has(workBuddy, /line-height:\s*17px/),
);
check(
  '顶部签到状态提升到 12px / 17px',
  has(workBuddyCheckin, /font-size:\s*12px/) && has(workBuddyCheckin, /line-height:\s*17px/),
);
check(
  '提示胶囊提升到 12px / 17px',
  has(notice, /font-size:\s*12px/) && has(notice, /line-height:\s*17px/),
);
check(
  '模型选择器触发器提升到 12px / 17px',
  has(pickerTrigger, /font-size:\s*12px/) && has(pickerTrigger, /line-height:\s*17px/),
);
check(
  '不使用会让 Windows WebView2 变糊的 antialiased hack',
  !/\-webkit-font-smoothing\s*:\s*antialiased/.test(css),
);

console.log('');
console.log(`=== 字体清晰度：通过 ${passed}，失败 ${failed} ===`);
process.exit(failed === 0 ? 0 : 1);
