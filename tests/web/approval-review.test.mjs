// 审批卡要让人看见自己在批什么。
//
// 本期锁住四件事，任何一件退回去，「允许」就又变成一次跳跃：
//   一、写值/写公式给出「现在 → 将改为」的逐格对照，而不只是 values: 20 行 × 3 列；
//   二、截断要用文字说出来，且报的是剩余**格数**——截断同时发生在行与列两个方向；
//   三、空单元格与「读不到当前值」必须分得开：前者是正常写入，后者是探测失败；
//   四、卡片上的范围是跳进 Excel 的入口，点它发 sheet.goto。
//
// 另外锁住授权语义：「本轮同类允许」不得顺带放行结构，
// 结构要另一个写明含结构的按钮，且芯片必须用文字说清放行了哪张表的哪一类。
//
// 运行：node tests/web/approval-review.test.mjs

const posted = [];
let hostHandler = null;

globalThis.window = {
  chrome: {
    webview: {
      addEventListener: (kind, handler) => {
        if (kind === 'message') { hostHandler = handler; }
      },
      postMessage: (message) => posted.push(message),
    },
  },
  innerWidth: 420,
  location: { hash: '' },
};

/*
  假 DOM。querySelector 必须是真的：审批卡的渲染全靠按类名找子节点，
  返回 null 会让被测代码抛异常或对着空节点通过断言——那种全绿什么也没验。
*/
function makeNode(tag = 'div') {
  const node = {
    tag,
    textContent: '',
    innerHTML: '',
    title: '',
    value: '',
    type: '',
    disabled: false,
    hidden: false,
    open: false,
    scrollTop: 0,
    scrollHeight: 40,
    clientHeight: 17,
    style: {},
    dataset: {},
    attributes: {},
    children: [],
    parent: null,
    listeners: new Map(),
    classes: new Set(),
    append: (...kids) => {
      for (const kid of kids) {
        if (kid && typeof kid === 'object') { kid.parent = node; }
        node.children.push(kid);
      }
    },
    remove: () => {
      const parent = node.parent;
      if (!parent) { return; }
      parent.children = parent.children.filter((n) => n !== node);
      node.parent = null;
    },
    replaceWith: (other) => {
      const parent = node.parent;
      if (!parent) { return; }
      const index = parent.children.indexOf(node);
      if (index >= 0) { parent.children[index] = other; }
      if (other && typeof other === 'object') { other.parent = parent; }
      node.parent = null;
    },
    replaceChildren: (...kids) => {
      for (const kid of node.children) {
        if (kid && typeof kid === 'object') { kid.parent = null; }
      }
      node.children = [];
      node.append(...kids);
    },
    setAttribute: (name, value) => { node.attributes[name] = value; },
    getAttribute: (name) => node.attributes[name],
    focus: () => {},
    querySelector: (selector) => descendants(node).find((n) => matches(n, selector)) ?? null,
    querySelectorAll: (selector) => descendants(node).filter((n) => matches(n, selector)),
    addEventListener: (kind, handler) => node.listeners.set(kind, handler),
    classList: {
      // add / remove 收多个类名，与真实 DOM 一致。
      // 只收一个的话 classList.add('a', 'b') 会静默丢掉 'b'，
      // 断言表现成「产品没加那个类」，指向的方向完全不对。
      add: (...names) => names.forEach((n) => node.classes.add(n)),
      remove: (...names) => names.forEach((n) => node.classes.delete(n)),
      contains: (name) => node.classes.has(name),
      toggle: (name, on) => (on ? node.classes.add(name) : node.classes.delete(name)),
    },
  };

  Object.defineProperty(node, 'className', {
    get: () => [...node.classes].join(' '),
    set: (value) => {
      node.classes.clear();
      for (const name of String(value).split(/\s+/).filter(Boolean)) {
        node.classes.add(name);
      }
    },
  });

  return node;
}

function descendants(node) {
  const out = [];
  for (const kid of node.children) {
    if (!kid || typeof kid !== 'object') { continue; }
    out.push(kid);
    if (Array.isArray(kid.children)) { out.push(...descendants(kid)); }
  }
  return out;
}

/*
  认类名、标签，以及「类名 + data 属性」这一种组合。

  属性形态不能省：模型发起的操作卡片是靠
  `.tool-card[data-tool-id="…"]` 找回来的，只认类名会让这次查找返回 null，
  于是结果永远填不进卡片——而断言看起来只是「没有撤销按钮」，
  指向的方向完全不对。
*/
function matches(node, selector) {
  for (const part of String(selector).split(',')) {
    const one = part.trim();
    const attr = /^\.([\w-]+)\[data-([\w-]+)="([^"]*)"\]$/.exec(one);
    if (attr) {
      const [, cls, rawName, value] = attr;
      const key = rawName.replace(/-([a-z])/g, (_, c) => c.toUpperCase());
      if (node.classes?.has(cls) && node.dataset?.[key] === value) { return true; }
      continue;
    }

    if (one.startsWith('.')) {
      if (node.classes?.has(one.slice(1))) { return true; }
    } else if (one && node.tag === one) {
      return true;
    }
  }
  return false;
}

const nodes = new Map();
function nodeFor(id) {
  if (!nodes.has(id)) { nodes.set(id, makeNode()); }
  return nodes.get(id);
}

const fitPop = nodeFor('fit-pop');
for (const align of ['left', 'center', 'right']) {
  const item = makeNode('button');
  item.className = 'fit-item';
  item.dataset.align = align;
  fitPop.append(item);
}

globalThis.document = {
  getElementById: (id) => nodeFor(id),
  querySelector: () => makeNode(),
  querySelectorAll: () => [],
  addEventListener: () => {},
  createElement: (tag) => makeNode(tag),
};

globalThis.getComputedStyle = () => ({ lineHeight: '17' });

const { initChat } = await import('../../src/web/scripts/chat.js');

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

initChat();

const transcript = nodeFor('transcript');
const tick = () => new Promise((resolve) => setImmediate(resolve));

/** 投一条宿主推送，驱动真实代码路径——不在测试里复刻卡片渲染。 */
function push(payload) {
  hostHandler({ data: payload });
}

const approvals = () => transcript.children.filter((n) => n.classes?.has('approval'));
const toolCards = () => transcript.children.filter((n) => n.classes?.has('tool-card'));
const lastApproval = () => approvals()[approvals().length - 1];
const findIn = (card, cls) => descendants(card).find((n) => n.classes?.has(cls));
const allIn = (card, cls) => descendants(card).filter((n) => n.classes?.has(cls));
const textIn = (card, cls) => findIn(card, cls)?.textContent ?? '';
const buttonIn = (card, label) =>
  descendants(card).find((n) => n.tag === 'button' && n.textContent === label);
const click = (node) => node?.listeners.get('click')?.({
  preventDefault: () => {},
  stopPropagation: () => {},
});

console.log('检查缺少 root 的工具调用补全：');

push({
  kind: 'agent',
  stage: 'tool-start',
  payload: { id: 'call-root', name: 'excel_object', args: { path: [], action: 'get', member: 'Name' } },
});
push({ kind: 'tool-root-request', id: 'call-root' });
await tick();

const rootCard = toolCards().find((item) => item.dataset.toolId === 'call-root');
const rootChoices = allIn(rootCard, 'tool-root-options')[0]?.children ?? [];
check('缺 root 时展开原工具卡等待选择', Boolean(rootCard?.open)
  && textIn(rootCard, 'tool-state').includes('等待选择'));
check('显示三个对象入口和取消按钮', rootChoices.length === 4
  && rootChoices.map((item) => item.textContent).join('|')
    === '当前工作簿|当前工作表|当前窗口|取消操作');

click(rootChoices[1]);
await tick();
const rootChoice = [...posted].reverse().find((item) => item.channel === 'tool.root.respond');
check('选择工作表时回传原调用 ID 和白名单值', rootChoice?.payload?.id === 'call-root'
  && rootChoice.payload.root === 'worksheet' && rootChoice.payload.cancelled === false,
  JSON.stringify(rootChoice?.payload));
hostHandler({ data: { kind: 'response', id: rootChoice.id, ok: true, data: { accepted: true } } });
await tick();
push({
  kind: 'agent',
  stage: 'tool-args-updated',
  payload: { id: 'call-root', args: { root: 'worksheet', path: [], action: 'get', member: 'Name' } },
});
await tick();
check('同一卡片显示补全后的 root', textIn(rootCard, 'tool-args').includes('root: worksheet'));
check('选择后选项不可重复提交', rootChoices.every((item) => item.disabled));

push({
  kind: 'agent',
  stage: 'tool-start',
  payload: { id: 'call-root-cancel', name: 'excel_object', args: { path: [], action: 'get', member: 'Name' } },
});
push({ kind: 'tool-root-request', id: 'call-root-cancel' });
await tick();
const cancelCard = toolCards().find((item) => item.dataset.toolId === 'call-root-cancel');
click(buttonIn(cancelCard, '取消操作'));
await tick();
const rootCancel = [...posted].reverse().find((item) => item.channel === 'tool.root.respond');
check('取消会回传取消标记而不附加 root', rootCancel?.payload?.id === 'call-root-cancel'
  && rootCancel.payload.cancelled === true && rootCancel.payload.root === null,
  JSON.stringify(rootCancel?.payload));
hostHandler({ data: { kind: 'response', id: rootCancel.id, ok: true, data: { accepted: true } } });
await tick();
check('取消后原卡片明确说明没有执行', textIn(cancelCard, 'tool-root-status').includes('该操作未执行'));

push({
  kind: 'agent',
  stage: 'tool-start',
  payload: { id: 'call-root-stop', name: 'excel_object', args: { path: [], action: 'get', member: 'Name' } },
});
push({ kind: 'tool-root-request', id: 'call-root-stop' });
await tick();
const stoppedRootCard = toolCards().find((item) => item.dataset.toolId === 'call-root-stop');
push({ kind: 'agent', stage: 'stopped', text: '已停止生成。' });
await tick();
check('停止任务会关闭入口选项并说明未执行',
  allIn(stoppedRootCard, 'tool-root-options')[0].children.every((item) => item.disabled)
    && textIn(stoppedRootCard, 'tool-root-status').includes('该操作未执行'));

console.log('检查审批卡的可见性与授权语义：');

// ---- 一、写值给出逐格对照 ----

push({
  kind: 'approval-request',
  id: 'ap1',
  tool: 'write_values',
  risk: 'Write',
  impact: '',
  impactRange: { sheet: 'Sheet1', address: '$B$2:$C$3', cells: 4 },
  preview: {
    currentUnreadable: false,
    formattingMixed: false,
    omittedCells: 0,
    cells: [
      { row: 1, column: 1, before: '甲', after: '丙', beforeEmpty: false, afterEmpty: false },
      { row: 1, column: 2, before: '', after: '0', beforeEmpty: true, afterEmpty: false },
      { row: 2, column: 1, before: '乙', after: '丁', beforeEmpty: false, afterEmpty: false },
      { row: 2, column: 2, before: '2', after: '', beforeEmpty: false, afterEmpty: true },
    ],
  },
  args: { range: '$B$2:$C$3', values: [['丙', 0], ['丁', null]] },
});
await tick();

const card = lastApproval();
check('审批请求落成一张审批卡', approvals().length === 1, `卡片 ${approvals().length} 张`);
check('卡片上出现前后对照', Boolean(findIn(card, 'approval-preview')));

const rows = allIn(card, 'approval-preview-table').flatMap((t) => t.children.filter((n) => n.tag === 'tr'));
// 一行表头 + 四行数据。少了表头就分不清哪列是「现在」哪列是「将改为」。
check('对照表含表头与四行数据', rows.length === 5, `共 ${rows.length} 行`);

const dataCells = rows.slice(1).map((row) => row.children.map((c) => c.textContent));
check('第一格给出原值与新值',
  dataCells[0]?.[1] === '甲' && dataCells[0]?.[2] === '丙',
  JSON.stringify(dataCells[0]));
check('位置用范围内的相对行列',
  (dataCells[0]?.[0] ?? '').includes('第 1 行') && (dataCells[0]?.[0] ?? '').includes('第 1 列'),
  dataCells[0]?.[0]);
// 空与 0 必须分得开：这正是「合成一个样子」会让用户看错的地方。
check('空的原值显示为（空）', dataCells[1]?.[1] === '（空）', JSON.stringify(dataCells[1]));
check('新值 0 照常显示为 0', dataCells[1]?.[2] === '0', JSON.stringify(dataCells[1]));
check('要写成空的那格也标（空）', dataCells[3]?.[2] === '（空）', JSON.stringify(dataCells[3]));
check('全部列出时不出现省略说明', allIn(card, 'approval-preview-more').length === 0);

// 对照在参数之前：参数区只报形状，用户要先看见内容才有决定可做。
const order = card.children.map((n) => n.className);
const previewAt = order.findIndex((c) => c.includes('approval-preview'));
const argsAt = order.findIndex((c) => c.includes('approval-args'));
check('对照排在参数区之前', previewAt >= 0 && argsAt >= 0 && previewAt < argsAt,
  JSON.stringify(order));

// ---- 二、范围是跳进 Excel 的入口 ----

const jump = findIn(card, 'range-jump');
check('影响范围是可点的控件', Boolean(jump) && jump.tag === 'button', jump?.tag);
check('范围按钮写出行列位置', (jump?.textContent ?? '').includes('行'), jump?.textContent);
check('悬停说明会改变当前选区', (jump?.title ?? '').includes('选区'), jump?.title);

click(jump);
await tick();

const goto = [...posted].reverse().find((m) => m.channel === 'sheet.goto');
check('点范围会发出 sheet.goto', Boolean(goto), JSON.stringify(posted.slice(-1)));
check('跳转带上工作表与地址',
  goto?.payload?.sheet === 'Sheet1' && goto?.payload?.address === '$B$2:$C$3',
  JSON.stringify(goto?.payload));

// ---- 三、「本轮同类允许」不含结构，含结构另有按钮 ----

const approveAll = buttonIn(card, '本轮同类允许');
const approveStructure = buttonIn(card, '含结构允许');
check('存在「本轮同类允许」', Boolean(approveAll));
check('存在独立的「含结构允许」', Boolean(approveStructure));
check('同类允许的说明写明不含结构',
  (approveAll?.title ?? '').includes('不会'), approveAll?.title);
check('含结构允许的说明点明建表建图',
  (approveStructure?.title ?? '').includes('建图'), approveStructure?.title);

click(approveAll);
await tick();

const respond = [...posted].reverse().find((m) => m.channel === 'approval.respond');
check('同类允许回传 approveRest', respond?.payload?.approveRest === true,
  JSON.stringify(respond?.payload));
check('同类允许不回传含结构',
  respond?.payload?.approveStructureRest === false,
  JSON.stringify(respond?.payload));

// ---- 四、截断要说出来，且报剩余格数 ----

push({
  kind: 'approval-request',
  id: 'ap2',
  tool: 'write_values',
  risk: 'Write',
  impactRange: { sheet: 'Sheet1', address: '$A$1:$C$20', cells: 60 },
  preview: {
    currentUnreadable: false,
    formattingMixed: false,
    omittedCells: 36,
    cells: Array.from({ length: 24 }, (_, i) => ({
      row: Math.floor(i / 3) + 1,
      column: (i % 3) + 1,
      before: `旧${i}`,
      after: `新${i}`,
      beforeEmpty: false,
      afterEmpty: false,
    })),
  },
  args: { range: '$A$1:$C$20', values: [] },
});
await tick();

const truncated = lastApproval();
const more = findIn(truncated, 'approval-preview-more');
check('截断时给出文字说明', Boolean(more), '缺少 approval-preview-more');
check('省略说明报的是格数',
  (more?.textContent ?? '').includes('36') && (more?.textContent ?? '').includes('单元格'),
  more?.textContent);

// ---- 五、读不到当前值时不画空表 ----

push({
  kind: 'approval-request',
  id: 'ap3',
  tool: 'write_values',
  risk: 'Write',
  impactRange: { sheet: 'Sheet1', address: '$A$1:$Z$5000', cells: 130000 },
  preview: { currentUnreadable: true, formattingMixed: false, omittedCells: 0, cells: [] },
  args: { range: '$A$1:$Z$5000', values: [] },
});
await tick();

const unreadable = lastApproval();
check('读不到当前值时不渲染对照表',
  allIn(unreadable, 'approval-preview-table').length === 0,
  '不应出现表格');
check('读不到当前值时说明原因',
  (textIn(unreadable, 'approval-preview-note') ?? '').includes('读不到'),
  textIn(unreadable, 'approval-preview-note'));

// ---- 六、格式不统一时如实说一句 ----

push({
  kind: 'approval-request',
  id: 'ap4',
  tool: 'format_range',
  risk: 'Write',
  impactRange: { sheet: 'Sheet1', address: '$A$1:$D$9', cells: 36 },
  preview: { currentUnreadable: false, formattingMixed: true, omittedCells: 0, cells: [] },
  args: { range: '$A$1:$D$9', bold: true },
});
await tick();

const mixed = lastApproval();
check('格式不统一时给出说明',
  (textIn(mixed, 'approval-preview-note') ?? '').includes('无法完整还原'),
  textIn(mixed, 'approval-preview-note'));
check('格式卡片不渲染对照表',
  allIn(mixed, 'approval-preview-table').length === 0,
  '不应出现表格');

// ---- 七、授权芯片用文字说清放行了什么 ----

const chip = nodeFor('approval-grants');
push({ kind: 'agent', stage: 'approval-grants', payload: { grants: [{ sheet: 'Sheet1', approvalClass: 'Format' }] } });
await tick();

check('授权后芯片可见', chip.hidden === false, `hidden=${chip.hidden}`);
check('芯片写出工作表与类别',
  chip.textContent.includes('Sheet1') && chip.textContent.includes('格式'),
  chip.textContent);
check('芯片悬停说明新一轮会重新确认',
  chip.title.includes('新一轮'), chip.title);

// 芯片本身是收回入口：只显示不给收回，用户中途改主意时唯一出路会变成掐掉整轮。
// 「它是不是真的 button」在 icon-buttons.test.mjs 上验——那里读真实 HTML，
// 而这里的假 DOM 一律建 div，断言标签只是自问自答。
check('悬停说明提到收回', (chip.title ?? '').includes('收回'), chip.title);

click(chip);
await tick();

const revoke = [...posted].reverse().find((m) => m.channel === 'approval.revoke');
check('点芯片发出收回请求', Boolean(revoke), JSON.stringify(posted.slice(-1)));
if (revoke) {
  hostHandler({ data: { kind: 'response', id: revoke.id, ok: true, data: { ok: true, revoked: 1 } } });
}
await tick();

check('收回后芯片隐藏', chip.hidden === true, `hidden=${chip.hidden}`);
check('收回后留一条说明',
  transcript.children.some((n) => n.classes?.has('notice')
    && (n.textContent ?? '').includes('重新逐个询问')),
  '缺少收回说明');

push({ kind: 'agent', stage: 'approval-grants', payload: { grants: [] } });
await tick();
check('授权清空后芯片隐藏', chip.hidden === true, `hidden=${chip.hidden}`);

// ---- 七之二、操作卡上的范围也可跳转 ----

push({
  kind: 'agent',
  stage: 'tool-start',
  payload: { id: 'call-jump', name: 'write_values', risk: 'Write', args: { range: 'B2:C3' } },
});
push({
  kind: 'agent',
  stage: 'tool-result',
  payload: {
    id: 'call-jump',
    name: 'write_values',
    ok: true,
    data: { sheet: 'Sheet1', address: '$B$2:$C$3', cells_written: 4 },
    canUndo: false,
  },
});
await tick();

const opCard = transcript.children.filter((n) => n.classes?.has('tool-card')).pop();
const opJump = descendants(opCard).find((n) => n.classes?.has('range-jump'));
check('操作卡成功摘要里的范围可点', Boolean(opJump), '没有 range-jump');
click(opJump);
await tick();
const opGoto = [...posted].reverse().find((m) => m.channel === 'sheet.goto');
check('操作卡跳转带上工作表与地址',
  opGoto?.payload?.sheet === 'Sheet1' && opGoto?.payload?.address === '$B$2:$C$3',
  JSON.stringify(opGoto?.payload));

// ---- 八、切换处理方式的回执是居中胶囊，且只留一条 ----

// 设置返回后才有三档的标签与说明。
hostHandler({
  data: {
    kind: 'push',
    channel: 'session',
    payload: {},
  },
});

const approvalIcon = nodeFor('approval-icon');
const noticesOf = () => transcript.children.filter((n) => n.classes?.has('notice'));
const approvalNotices = () => transcript.children.filter((n) => n.classes?.has('notice-approval'));

const beforeCount = noticesOf().length;
click(approvalIcon);
await tick();

// 切换要落盘，回应它，否则 Promise 悬着、回执不出现。
const saved = [...posted].reverse().find((m) => m.channel === 'session.update');
check('切换会保存处理方式', Boolean(saved), JSON.stringify(posted.slice(-1)));
if (saved) {
  hostHandler({ data: { kind: 'response', id: saved.id, ok: true, data: {} } });
}
await tick();

check('回执落在对话流里而不是状态行',
  approvalNotices().length === 1,
  `胶囊 ${approvalNotices().length} 条`);

const modeNotice = approvalNotices()[0];
check('回执是居中胶囊（沿用 notice）',
  modeNotice.classes.has('notice'), modeNotice.className);
check('回执写出档位名', (modeNotice.textContent ?? '').includes('处理方式'),
  modeNotice.textContent);
check('回执带三档配色之一',
  ['is-strict', 'is-medium', 'is-auto'].some((c) => modeNotice.classes.has(c)),
  modeNotice.className);
check('回执不复用出错样式',
  !modeNotice.classes.has('notice-error'), modeNotice.className);
check('回执有悬停说明', Boolean(modeNotice.title), modeNotice.title);

// 连点：轮换三档只应留最后那一条，否则三条自相矛盾的记录同时在场。
for (let i = 0; i < 2; i++) {
  click(approvalIcon);
  await tick();
  const again = [...posted].reverse().find((m) => m.channel === 'session.update');
  hostHandler({ data: { kind: 'response', id: again.id, ok: true, data: {} } });
  await tick();
}

check('连点三次只留一条回执',
  approvalNotices().length === 1,
  `胶囊 ${approvalNotices().length} 条`);
check('没有额外堆积其他提示',
  noticesOf().length === beforeCount + 1,
  `提示总数 ${noticesOf().length}，切换前 ${beforeCount}`);

// ---- 八之二、不给撤销按钮时必须说原因 ----
//
// 缺按钮本身是可见的，缺原因会被当成故障——而它其实是「保不住足以完整还原的
// 依据就不承诺撤销」这一有意为之的取舍。面板「适配」早就这样做，
// 模型发起的这条路上曾经两种情形都是静默的。

const withheld = [
  {
    id: 'call-chart-noname',
    name: 'create_chart',
    note: '这张图表不能撤销：宿主没有回报图表的名称，撤销时无法定位它。需要时请让我删掉重建。',
    label: '建图没名字',
    keyword: '图表的名称',
  },
  {
    id: 'call-format-mixed',
    name: 'format_range',
    note: '这次格式改动不能撤销：这片范围原本的外观逐项都不一样，宿主读不出统一值，保不住足以完整还原的快照。改动本身已经生效。',
    label: '格式全项不一致',
    keyword: '逐项都不一样',
  },
];

for (const item of withheld) {
  push({
    kind: 'agent',
    stage: 'tool-start',
    payload: { id: item.id, name: item.name, risk: 'Write', args: { range: 'A1' } },
  });
  push({
    kind: 'agent',
    stage: 'tool-result',
    payload: {
      id: item.id,
      name: item.name,
      ok: true,
      data: { sheet: 'Sheet1', address: '$A$1' },
      canUndo: false,
      undoNote: item.note,
    },
  });
  await tick();

  const card = transcript.children.filter((n) => n.classes?.has('tool-card')).pop();
  check(`${item.label}：不给撤销按钮`,
    !descendants(card).some((n) => n.classes?.has('tool-undo')),
    '仍有撤销按钮');
  check(`${item.label}：卡片说明为什么不能撤销`,
    descendants(card).some((n) => n.classes?.has('tool-note')
      && (n.textContent ?? '').includes(item.keyword)),
    descendants(card).filter((n) => n.classes?.has('tool-note')).map((n) => n.textContent).join(' | '));
}

// ---- 九、记录被挤掉之后不留一个永远失败的按钮 ----

push({
  kind: 'agent',
  stage: 'tool-start',
  payload: { id: 'call-evicted', name: 'write_values', risk: 'Write', args: { range: 'A1' } },
});
push({
  kind: 'agent',
  stage: 'tool-result',
  payload: {
    id: 'call-evicted',
    name: 'write_values',
    ok: true,
    data: { sheet: 'Sheet1', address: '$A$1', cells_written: 1 },
    canUndo: true,
    undoSummary: '写入值 Sheet1!$A$1',
    canRedoAfterUndo: true,
  },
});
await tick();

const evictedCard = transcript.children.filter((n) => n.classes?.has('tool-card')).pop();
const evictedUndo = descendants(evictedCard).find((n) => n.classes?.has('tool-undo'));
check('操作卡上先有撤销按钮', Boolean(evictedUndo));

click(evictedUndo);
await tick();

const undoCall = [...posted].reverse().find((m) => m.channel === 'undo.apply');
hostHandler({
  data: {
    kind: 'response',
    id: undoCall.id,
    ok: true,
    data: { ok: false, errorCode: 'NOT_FOUND', message: '找不到该操作记录，可能已超出保留范围。' },
  },
});
await tick();

check('记录被挤掉后撤掉按钮',
  !descendants(evictedCard).some((n) => n.classes?.has('tool-undo')),
  '按钮仍在');
check('记录被挤掉时说明原因',
  descendants(evictedCard).some((n) => n.classes?.has('tool-note')
    && (n.textContent ?? '').includes('保留条数')),
  descendants(evictedCard).filter((n) => n.classes?.has('tool-note')).map((n) => n.textContent).join(' | '));

console.log('');
console.log(`=== 审批卡可见性：通过 ${passed}，失败 ${failed} ===`);
process.exit(failed === 0 ? 0 : 1);
