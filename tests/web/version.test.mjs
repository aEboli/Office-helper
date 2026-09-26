import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = fileURLToPath(new URL('.', import.meta.url));
const webDir = join(here, '..', '..', 'src', 'web');
const { formatVersion, updateVersionDisplay } = await import('../../src/web/scripts/version.js');

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

console.log('检查标题栏版本显示：');

check(
  '三段版本显示为阿拉伯数字',
  formatVersion('0.10.1') === '0.10.1',
  formatVersion('0.10.1'),
);
check(
  '四段版本保留可重复部署构建号',
  formatVersion('0.10.1.1') === '0.10.1.1',
  formatVersion('0.10.1.1'),
);
check(
  '无效版本号不渲染',
  formatVersion('preview') === '',
  formatVersion('preview'),
);

const node = { dataset: { version: '0.10.3.16' }, textContent: '' };
const documentRoot = {
  getElementById: (id) => id === 'app-version' ? node : null,
};
check(
  '宿主版本返回后更新标题栏',
  updateVersionDisplay('0.10.3.16', documentRoot) && node.textContent === '0.10.3.16',
  node.textContent,
);
check(
  '没有宿主版本时使用静态回退',
  updateVersionDisplay('', documentRoot) && node.textContent === '0.10.3.16',
  node.textContent,
);

const html = readFileSync(join(webDir, 'index.html'), 'utf8');
const css = readFileSync(join(webDir, 'styles', 'app.css'), 'utf8');
check(
  '标题栏包含版本节点和当前回退版本',
  /id="app-version"[^>]*data-version="0\.10\.3\.16">0\.10\.3\.16<\/span>/.test(html),
);
check(
  '版本节点使用清晰的 UI 字体',
  /\.app-version\s*\{[^}]*font-family:\s*"Segoe UI"[^}]*"Microsoft YaHei UI"/.test(css),
);

console.log('');
console.log(`=== 标题栏版本显示：通过 ${passed}，失败 ${failed} ===`);
process.exit(failed === 0 ? 0 : 1);
