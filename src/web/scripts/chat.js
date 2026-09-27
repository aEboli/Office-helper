import { request, on, isHosted, logToHost } from './bridge.js';
import { renderMarkdown } from './markdown.js';
import { initPicker, setPickerTurnInFlight, syncPicker } from './picker.js';
import { describeRange, rangeLabel } from './range-label.js';
import { prefersReducedMotion } from './motion.js';
import { updateChannelHeader } from './channel-header.js';
import {
  initAttachments,
  getImages,
  getFiles,
  hasAttachments,
  clearAttachments,
  createFileGlyph,
} from './attachments.js';
import { isWordOfficeHost, welcomeCopy } from './host-ui.js';

const TOOL_LABELS = {
  get_document_info: '读取文档信息',
  read_document_structure: '读取文档结构',
  read_paragraphs: '读取段落',
  read_table: '读取表格',
  find_text: '查找文字',
  replace_text: '替换文字',
  insert_text: '插入文字',
  insert_document_draft: '插入文档草稿',
  delete_range: '删除文档范围',
  format_text: '设置字符格式',
  format_paragraph: '设置段落格式',
  apply_style: '应用文档样式',
  edit_table: '编辑表格',
  set_page_setup: '设置页面',
  insert_page_break: '插入分页符',
  excel_object: '操作表格对象',
  get_workbook_info: '读取工作簿结构',
  get_selection: '读取当前选区',
  read_range: '读取范围',
  write_values: '写入值',
  write_formulas: '写入公式',
  format_range: '设置格式',
  set_number_format: '设置数字格式',
  autofit_range: '自动调整',
  fit_range: '适配',
  merge_cells: '合并单元格',
  unmerge_cells: '取消合并',
  clear_range: '清除',
  add_worksheet: '新增工作表',
  rename_worksheet: '重命名工作表',
  sort_range: '排序',
  create_table: '创建表格',
  create_chart: '创建图表',
};

const RISK_LABELS = {
  Read: '读取',
  Write: '修改内容',
  Structure: '改变结构',
};

/**
 * 操作统计里对风险等级的简称。
 *
 * 与审批卡片上的 RISK_LABELS 分开：那里要把「将要发生什么」说清楚，
 * 这里是一行摘要里的计数单位，宽度只够一个字（「3 改 1 读」）。
 * 改与结构都算「改」——摘要要回答的是「那一轮动过表没有」，
 * 再分成两类反而要多读一遍才知道加起来是几。
 */
const RISK_TALLY = {
  Read: '读',
  Write: '改',
  Structure: '改',
};

/**
 * 值为 A1 地址的字段。展示时在原值后补一句「几行 × 哪几列」。
 *
 * 按字段名而非按值形态判断：像 format_code 的 0.00 或 key_column 的 B
 * 本身也能被当作地址解析，误译成位置说明反而更难懂。
 */
const ADDRESS_KEYS = new Set(['range', 'address', 'source_range']);

let transcript;
let composer;
let sendButton;
let statusLine;
let usageLine;
let queueStrip;

let busy = false;
let draftOperationPending = false;
let currentAssistant = null;
let currentThinking = null;

/**
 * 待处理的输入队列。
 *
 * 为什么要队列：加载项同一时刻只跑一轮（后端见 chat.send 的 BUSY 守卫），
 * 早先的做法是处理中直接禁用输入框，于是想到的下一步只能干等或记在别处。
 * 改为入队后，输入随时可写，上一轮结束即自动接着跑。
 *
 * 队列留在面板侧而非加载项：排队中的条目要能看见、能取消，
 * 这两件事都是界面的事；后端多一个队列反而要再开一套查询与撤单通道。
 * 代价是刷新面板会丢掉未发出的排队项，但那与丢掉输入框里的草稿同级。
 *
 * 排队条目显示在输入区上方的排队条里（见 renderQueueStrip），只有真正开跑时
 * 才进对话流。早先的做法是入队即上屏成气泡，实测的问题是：还没发生的事混在
 * 已发生的消息之间会被当成已经处理过，而且对话一长就被顶出可视区——
 * 用户想确认「刚排的那条还在不在」反而要往上翻。
 *
 * 同理，取消掉的条目直接消失（见 cancelQueued）：它从未发出，对话流里不该有它。
 */
const queue = [];
let queueSequence = 0;

/** 正在轮转队列。防止多个入口并发触发同一条队列。 */
let pumping = false;

function toolLabel(name) {
  return TOOL_LABELS[name] ?? name;
}

function isWordHostUi() {
  const root = typeof document === 'undefined' ? null : document.documentElement;
  return isWordOfficeHost(root?.dataset?.officeHost);
}

function describeWordTarget(target) {
  if (!target || typeof target !== 'object') { return ''; }
  const story = target.story ? `Story=${target.story}` : '';
  const span = Number.isInteger(target.start) || Number.isInteger(target.end)
    ? `Start=${target.start ?? '?'} End=${target.end ?? '?'}` : '';
  const paragraph = Number.isInteger(target.paragraph_index) ? `段落=${target.paragraph_index}` : '';
  const table = Number.isInteger(target.table_index)
    ? `表格=${target.table_index} 行=${target.row ?? '?'} 列=${target.column ?? '?'}` : '';
  const named = target.bookmark ? `书签=${target.bookmark}` :
    (target.content_control ? `内容控件=${target.content_control}` : '');
  return [story, span, paragraph, table, named].filter(Boolean).join('，');
}

function describeWordResult(data) {
  if (!data || typeof data !== 'object') { return '完成文档操作'; }
  const target = describeWordTarget(data);
  if (typeof data.text === 'string') {
    return `${target || '明确文档目标'}，读回文本 ${data.text.length > 80 ? `${data.text.slice(0, 80)}…` : data.text}`;
  }
  return target || '完成文档操作';
}

function scrollToBottom() {
  // 用 scrollTop 而非 scrollIntoView：后者在窄栏里会带偏水平位置。
  transcript.scrollTop = transcript.scrollHeight;
}

/**
 * 上屏顺序的计数器。每个进对话流的节点记一个递增序号。
 *
 * 为什么要记：操作卡片被收进轮次组后，「还原」要把它们放回原来的位置。
 * 记「插在谁后面」是不够的——那个锚点自己可能也被收进了别的组，
 * 于是还原时找不到落点。序号是每个节点自己的属性，不依赖邻居是否还在原处，
 * 还原时把对话流按序号重排即可，不管中间又形成过几个组。
 */
let mountSequence = 0;

/**
 * 把节点挂到对话流末尾并记下上屏顺序。
 *
 * 所有进对话流的东西都走这里，没有例外：漏一个的话它就没有序号，
 * 还原时会被当成最早的节点排到最前面。
 *
 * 重复挂载同一个节点（showPending 会这么做，把指示器移到末尾）会刷新它的序号，
 * 这是对的：它此刻确实在末尾。
 */
function mountToTranscript(node) {
  if (node.dataset.seq) {
    // 重挂。此刻必须把进场类摘掉，否则接下来的 append 会把动画从头重播一遍。
    //
    // 为什么：append 一个已是子节点的元素等于「先摘再插」，而移出文档会取消
    // 元素上的动画，插回去又重新起播。因此「首挂才加类」这一条挡不住重播——
    // 类是首挂时加的，可它还在，重挂就又放一次。真实 WebView2 里量到的是
    // 动画进度从 170ms 退回 0ms（PaneHarness --motion 的第二条断言）。
    //
    // 会撞上的路径都是毫秒级的：showPending 在下一个事件里把指示器移到末尾、
    // sealOpsBatch 把「已完成」胶囊移到组后面。表现是那个气泡可见地闪两下，
    // 而代码里看不出任何异常。
    node.classList.remove('is-entering');
  } else if (!prefersReducedMotion()) {
    // 首次挂载放进场动画（淡入上浮，见 app.css 的 is-entering）。
    // 「挂没挂过」直接看 seq 在不在——它恰好只在这里写。
    //
    // 减少动效时连类都不加：全局 CSS 把 animation 关掉后动画不起播，
    // animationend/animationcancel 都不会来，类会永久留在节点上。
    // 那本身不可见（类只带动画），但用户中途在系统设置里关掉「减少动效」时
    // 媒体查询会实时翻转，整条对话流所有残留节点同时起播——毫无操作却整屏闪
    // 一下。不加类就没有这条路。
    node.classList.add('is-entering');

    // animationcancel 与 animationend 都要听：动画被取消时只有前者会来
    // （sealOpsBatch 把仍在动的卡片搬进未渲染的组容器就是这种情形），
    // 只听 animationend 的话类会残留，日后该节点被重插时再淡入一次。
    //
    // 核对 target：气泡里还有自己会动的东西（那圈点），
    // 冒泡上来的结束事件不该提前把外层的进场动画掐掉。
    const done = (event) => {
      if (event.target === node) {
        node.classList.remove('is-entering');
      }
    };
    node.addEventListener('animationend', done);
    node.addEventListener('animationcancel', done);
  }

  mountSequence += 1;
  node.dataset.seq = String(mountSequence);
  transcript.append(node);
  return node;
}

/** 正在处理的助手气泡。同一时刻只会有一个。 */
let pendingBubble = null;

/**
 * 在对话流末尾显示一个「正在处理」的助手气泡。
 *
 * 放在回复气泡里而不是底部状态行：答案将出现的位置就是用户目光所在，
 * 进展写在那里不必来回扫视；底部那行留给错误与配置提示。
 * 气泡本身会被后续流式文本直接接管（见 ensureAssistant），
 * 因此从「正在处理」到正文是同一个气泡在原地变化，不会闪一下再重排。
 */
/**
 * 指示器上那圈点的数量。
 *
 * 必须与 app.css 里 .pending-dots i:nth-child(n) 的规则条数一致：那些规则给每个点
 * 定角度与相位，多出来的点没有 transform，会全都堆在圆圈顶端同一个位置。
 * 常量放在这里而不是写死 8：两处数字要一起改，至少让其中一处有名字。
 */
const PENDING_DOT_COUNT = 8;

/**
 * 显示处理中的指示器。
 *
 * label 省略时用兜底文案。调用方知道在做什么时应当传具体的那句（正在思考…、
 * 正在读取…），兜底只覆盖「说不出更具体的阶段」这一种情形。
 */
function showPending(label = '正在忙着办…') {
  if (pendingBubble) {
    pendingBubble.label.textContent = label;
    // 重新追加到末尾：工具卡片等内容可能已插到它后面。
    mountToTranscript(pendingBubble.wrapper);
    scrollToBottom();
    return;
  }

  const wrapper = document.createElement('div');
  wrapper.className = 'msg msg-assistant msg-pending';

  const body = document.createElement('div');
  body.className = 'msg-body';

  const indicator = document.createElement('div');
  indicator.className = 'pending';

  const dots = document.createElement('span');
  dots.className = 'pending-dots';
  dots.setAttribute('aria-hidden', 'true');
  for (let i = 0; i < PENDING_DOT_COUNT; i++) {
    dots.append(document.createElement('i'));
  }

  const text = document.createElement('span');
  text.className = 'pending-label';
  text.textContent = label;

  // 进展变化要让读屏软件播报，但不能打断用户当前操作。
  indicator.setAttribute('role', 'status');
  indicator.setAttribute('aria-live', 'polite');

  indicator.append(dots, text);
  body.append(indicator);
  wrapper.append(body);
  mountToTranscript(wrapper);
  scrollToBottom();

  pendingBubble = { wrapper, body, indicator, label: text };
}

function clearPending() {
  if (!pendingBubble) {
    return;
  }

  pendingBubble.wrapper.remove();
  pendingBubble = null;
}

/**
 * 造一个消息气泡，但不插入对话流。
 *
 * 与 addBubble 分开，是因为排队中的用户消息需要拿到气泡外框本身——
 * 队列标记与取消按钮挂在外框上，而 addBubble 只交出文本容器。
 */
function buildBubble(role, text, images = [], files = []) {
  const wrapper = document.createElement('div');
  wrapper.className = `msg msg-${role}`;

  const body = document.createElement('div');
  body.className = 'msg-body';

  // 图片放在文字之上，与发送给模型的顺序一致。
  if (images.length > 0) {
    const gallery = document.createElement('div');
    gallery.className = 'msg-images';
    for (const image of images) {
      const thumb = document.createElement('img');
      thumb.className = 'msg-image';
      thumb.src = image.dataUrl;
      thumb.alt = image.name ?? '图片';
      thumb.title = image.name ?? '图片';
      gallery.append(thumb);
    }

    body.append(gallery);
  }

  // 文件只列名字，不铺内容：一个几百行的 CSV 铺在气泡里会把对话顶走，
  // 而用户已经知道自己附了什么，需要的只是「确实附上了」这个确认。
  //
  // 带图标，且与输入区附件条用的是同一个（createFileGlyph）。此前这里只有
  // 纯文字，同一个文件发送前带图标、发送后成了灰色文字胶囊，看着像两种东西。
  if (files.length > 0) {
    const list = document.createElement('div');
    list.className = 'msg-files';
    for (const file of files) {
      const chip = document.createElement('span');
      chip.className = 'msg-file';

      const label = document.createElement('span');
      label.className = 'msg-file-name';
      label.textContent = file.name ?? '文件';

      chip.append(createFileGlyph('file-glyph'), label);
      chip.title = `${file.name ?? '文件'}（内容已随消息发送）`;
      list.append(chip);
    }

    body.append(list);
  }

  // 助手气泡即使初始为空也要建文本容器：流式增量要往里写。
  if (text || role === 'assistant') {
    const textNode = document.createElement('div');
    textNode.className = 'msg-text';
    if (role === 'user') {
      textNode.textContent = text ?? '';
    } else {
      textNode.innerHTML = renderMarkdown(text ?? '');
    }

    body.append(textNode);
  }

  wrapper.append(body);
  // 助手气泡的流式更新需要拿到文本容器本身。
  return { wrapper, body, text: body.querySelector('.msg-text') ?? body };
}

function addBubble(role, text, images = [], files = []) {
  const bubble = buildBubble(role, text, images, files);
  mountToTranscript(bubble.wrapper);
  scrollToBottom();
  return bubble.text;
}

/**
 * 取当前助手气泡，没有就建一个。
 *
 * 优先接管「正在处理」气泡：用户已经在看那个位置，就地换成正文
 * 比另起一个气泡更连贯，也少一次重排。
 */
function ensureAssistant() {
  if (currentAssistant) {
    return currentAssistant;
  }

  if (pendingBubble) {
    const { wrapper, body } = pendingBubble;
    pendingBubble = null;
    wrapper.classList.remove('msg-pending');
    body.replaceChildren();

    const textNode = document.createElement('div');
    textNode.className = 'msg-text';
    body.append(textNode);

    currentAssistant = { element: textNode, raw: '' };
    return currentAssistant;
  }

  currentAssistant = { element: addBubble('assistant', ''), raw: '' };
  return currentAssistant;
}

function appendAssistantText(delta) {
  const target = ensureAssistant();
  target.raw += delta;
  // 每次增量都整体重渲染：Markdown 结构会随新内容变化，
  // 增量拼 HTML 会在列表和表格上出错。文本量在单轮内可控。
  target.element.innerHTML = renderMarkdown(target.raw);
  scrollToBottom();
}

function appendThinking(delta) {
  if (!currentThinking) {
    const wrapper = document.createElement('details');
    wrapper.className = 'thinking';
    const summary = document.createElement('summary');
    summary.textContent = '思考过程';
    const body = document.createElement('div');
    body.className = 'thinking-body';
    wrapper.append(summary, body);
    mountToTranscript(wrapper);
    currentThinking = { element: body, raw: '' };
  }

  currentThinking.raw += delta;
  currentThinking.element.textContent = currentThinking.raw;
  scrollToBottom();
}

/**
 * 添加工具操作卡片，默认折叠。
 *
 * 用 details/summary 而非自制折叠：原生元素自带键盘操作与无障碍语义，
 * 且折叠状态由浏览器维护，不需要额外状态管理。
 *
 * 默认折叠的理由：一轮任务常有多次工具调用，全部展开会把对话正文
 * 挤到看不见；摘要行已包含工具名与结果，需要细节时再展开。
 */
function addToolCard(payload) {
  const card = document.createElement('details');
  card.className = payload.manual ? 'tool-card is-manual' : 'tool-card';
  card.dataset.toolId = payload.id ?? '';

  const head = document.createElement('summary');
  head.className = 'tool-head';

  const name = document.createElement('span');
  name.className = 'tool-name';
  name.textContent = toolLabel(payload.name);

  head.append(name);

  // 手动操作加一枚标记。
  //
  // 只靠边条颜色不够：颜色说不出区别在哪，色觉障碍下也可能根本看不出来。
  // 标记与卡片同宽同高，折叠时也在，是这两类操作唯一始终可读的差别。
  if (payload.source) {
    card.dataset.source = payload.source;
  }

  if (payload.manual) {
    const origin = document.createElement('span');
    origin.className = 'tool-origin';
    origin.textContent = '手动';
    origin.title = payload.source === 'ribbon'
      ? '你在功能区点按钮直接执行的，不是模型发起的'
      : '你在面板上点按钮直接执行的，不是模型发起的';
    head.append(origin);
  }

  const state = document.createElement('span');
  state.className = 'tool-state';
  state.textContent = '执行中…';

  // 撤销按钮占位，执行成功且可撤销时才填充。
  // 放在摘要行右端：与操作本身同处一行，不必展开就能撤销。
  const actions = document.createElement('span');
  actions.className = 'tool-actions';

  head.append(state, actions);

  const body = document.createElement('div');
  body.className = 'tool-body';

  const argsTitle = document.createElement('div');
  argsTitle.className = 'tool-section-title';
  argsTitle.textContent = '参数';

  const args = document.createElement('pre');
  args.className = 'tool-args';
  args.textContent = summarizeArgs(payload.args) || '（无参数）';

  body.append(argsTitle, args);
  card.append(head, body);
  mountToTranscript(card);
  // 收进当前这一批，下一轮开始时一起成组。
  joinOpsBatch(card, payload);
  scrollToBottom();
  return card;
}

/**
 * 描述二维数据的尺寸。
 *
 * 一律带「共」字：尺寸与位置说明的措辞必须能区分，
 * 否则「3 行 × 2 列」既可能指第 1-3 行，也可能指三行数据。
 */
function describeMatrixSize(value) {
  if (!Array.isArray(value[0])) {
    // 一维数组不是矩阵，按行列描述会凭空造出「1 列」。
    return `共 ${value.length} 项`;
  }

  return `共 ${value.length} 行 × ${value[0].length} 列`;
}

/** 地址字段的展示文本：保留原值，并在括号里给出位置说明。 */
function describeAddressValue(text) {
  const label = rangeLabel(text);
  return label === '' ? text : `${text}（${label}）`;
}

/** 参数摘要：完整 JSON 在窄栏里会挤占太多空间，只展示关键字段。 */
function summarizeArgs(args) {
  if (!args || typeof args !== 'object') {
    return '';
  }

  const parts = [];
  for (const [key, value] of Object.entries(args)) {
    if (value === null || value === undefined || value === '') {
      continue;
    }

    if (Array.isArray(value)) {
      parts.push(`${key}: ${describeMatrixSize(value)}`);
      continue;
    }

    if (isWordHostUi() && key === 'target' && value && typeof value === 'object') {
      parts.push(`目标：${describeWordTarget(value) || '未指定'}`);
      continue;
    }

    const text = String(value);
    if (ADDRESS_KEYS.has(key)) {
      parts.push(`${key}: ${describeAddressValue(text)}`);
      continue;
    }

    parts.push(`${key}: ${text.length > 60 ? `${text.slice(0, 60)}…` : text}`);
  }

  return parts.join('\n');
}

function finishToolCard(payload) {
  const card = transcript.querySelector(`.tool-card[data-tool-id="${payload.id}"]`);
  if (!card) {
    return;
  }

  fillToolCard(card, payload);
}

/**
 * 把结果填进一张已经在屏上的卡片。
 *
 * 与 finishToolCard 分开是为了面板直接发起的操作：它们的撤销标识要等宿主
 * 执行完才知道，没法在开跑时就按标识把卡片找回来，但卡片引用一直在手里。
 */
function fillToolCard(card, payload) {
  const state = card.querySelector('.tool-state');
  if (payload.ok) {
    fillSuccessState(state, payload.data);
    state.className = 'tool-state is-ok';
  } else {
    state.textContent = payload.error ?? '失败';
    state.className = 'tool-state is-error';
    card.classList.add('is-error');
    // 失败时自动展开：出错的细节是用户当下最需要看到的。
    card.open = true;
  }

  appendToolDetail(card, payload);

  if (payload.canUndo) {
    attachUndoButton(card, payload);
  }

  // 撤销能还原到什么程度。写在卡片上而不是只回给模型：
  // 「内容能撤、格式撤不回」这句话是用户决定要不要接着改的依据。
  if (payload.ok && payload.undoNote) {
    appendToolNote(card, payload.undoNote);
  }
}

/**
 * 给操作卡片挂上撤销按钮，撤销后原地变为「恢复」。
 *
 * 同一个按钮承担两个方向，而不是并排放两个：同一时刻只有一个动作有效，
 * 放两个会让另一个始终处于禁用状态，反而不易看懂当前处于哪个状态。
 */
function attachUndoButton(card, payload) {
  const actions = card.querySelector('.tool-actions');
  if (!actions) {
    return;
  }

  actions.replaceChildren();

  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'tool-undo';
  button.dataset.undone = 'false';
  button.textContent = '撤销';

  // 撤销说明里带的是原始地址，同样补上行列位置：
  // 悬停就能确认要还原的是哪几行哪几列，不必先展开卡片。
  const where = rangeLabel(payload.data?.address ?? payload.data?.source_range ?? '');
  if (payload.undoSummary) {
    button.title = where === ''
      ? `撤销：${payload.undoSummary}`
      : `撤销：${payload.undoSummary}（${where}）`;
  } else {
    const target = isWordHostUi() ? describeWordResult(payload.data) : '';
    button.title = target ? `撤销文档操作：${target}` : '撤销此操作';
  }

  button.addEventListener('click', async (event) => {
    // 阻止冒泡：按钮在 summary 内，点击会连带折叠或展开卡片。
    event.preventDefault();
    event.stopPropagation();

    const redo = button.dataset.undone === 'true';
    button.disabled = true;
    const original = button.textContent;
    button.textContent = redo ? '恢复中…' : '撤销中…';

    try {
      // 用户已经看过重叠警告并再点一次，这次带上 force。
      // 只对撤销有意义：恢复走「后快照」，不存在盖掉更晚改动的问题。
      const force = !redo && button.dataset.overlapWarned === 'true';
      const result = await request('undo.apply', { id: payload.id, redo, force });

      if (!result?.ok) {
        button.textContent = original;

        // 范围相交：第一次不执行，把话说清并把按钮改成明确的「仍然撤销」。
        // 乱序撤销本身是允许的，要拦的只是静默盖掉之后那次写入。
        if (result?.errorCode === 'OVERLAP_WARNING') {
          button.dataset.overlapWarned = 'true';
          button.textContent = '仍然撤销';
          button.title = result.message ?? '这次撤销会覆盖之后的一次改动。';
          appendToolNote(card, result.message ?? '这次撤销会覆盖之后的一次改动。');
          return;
        }

        // 撤销栈有条数上限，早期记录会被挤掉。此后这个按钮永远失败，
        // 留着它等于每次点都得到同一条错误。撤掉并说明原因——
        // 缺按钮本身是可见的，缺原因会被当成故障。
        if (result?.errorCode === 'NOT_FOUND') {
          button.remove();
          appendToolNote(card, '这一步已经不能撤销了：撤销记录超出保留条数，已被更晚的操作挤掉。');
          return;
        }

        addNotice(result?.message ?? '操作失败。', 'error');
        return;
      }

      const undone = result.undone === true;
      button.dataset.undone = undone ? 'true' : 'false';
      delete button.dataset.overlapWarned;
      card.classList.toggle('is-undone', undone);

      const badge = card.querySelector('.tool-state');
      if (badge) {
        if (undone) {
          badge.dataset.original = badge.dataset.original ?? badge.textContent;
          badge.textContent = '已撤销';
          badge.className = 'tool-state is-undone';
        } else if (badge.dataset.original) {
          badge.textContent = badge.dataset.original;
          badge.className = 'tool-state is-ok';
        }
      }

      // 撤销之后不一定能恢复：图表删掉就重建不回来。
      // 这时撤掉按钮并说明原因，而不是摆一个点下去必然失败的「恢复」——
      // 那只是把同一个谎言从一个按钮挪到另一个。
      if (undone && payload.canRedoAfterUndo === false) {
        button.remove();
        appendToolNote(card, '这一步撤销后无法自动恢复，需要时请让我重新创建。');
      } else {
        button.textContent = undone ? '恢复' : '撤销';
      }

      // 卡片可能已经收进某个轮次组，组的摘要里带着「已撤销」的计数，
      // 不刷新就停在成组那一刻的状态——而收起来时那正是唯一可见的说法。
      refreshOpsGroupFor(card);
    } catch (error) {
      button.textContent = original;
      addNotice(`操作失败：${error.message}`, 'error');
    } finally {
      button.disabled = false;
    }
  });

  actions.append(button);
}

/** 把执行结果追加到卡片折叠区，展开后可看到完整细节。 */
function appendToolDetail(card, payload) {
  const body = card.querySelector('.tool-body');
  if (!body) {
    return;
  }

  const title = document.createElement('div');
  title.className = 'tool-section-title';
  title.textContent = payload.ok ? '结果' : '失败原因';

  const detail = document.createElement('pre');
  detail.className = 'tool-args';

  if (payload.ok) {
    detail.textContent = describeResultDetail(payload.data);
  } else {
    detail.textContent = payload.errorCode
      ? `${payload.errorCode}\n${payload.error ?? ''}`
      : (payload.error ?? '未提供原因');
  }

  body.append(title, detail);
}

/**
 * 展开后展示的结果细节。
 * 原始 JSON 里的读回校验数据可能很长，这里按字段整理并截断，
 * 既能看清改了什么，又不会让卡片撑满整屏。
 */
function describeResultDetail(data) {
  if (!data || typeof data !== 'object') {
    return '（无返回数据）';
  }

  const lines = [];
  for (const [key, value] of Object.entries(data)) {
    if (value === null || value === undefined) {
      continue;
    }

    if (Array.isArray(value)) {
      const preview = JSON.stringify(value.slice(0, 3));
      lines.push(`${key}: ${describeMatrixSize(value)}`);
      lines.push(`  前几行 ${preview.length > 200 ? `${preview.slice(0, 200)}…` : preview}`);
      continue;
    }

    const text = String(value);
    if (ADDRESS_KEYS.has(key)) {
      lines.push(`${key}: ${describeAddressValue(text)}`);
      continue;
    }

    lines.push(`${key}: ${text.length > 200 ? `${text.slice(0, 200)}…` : text}`);
  }

  return lines.length > 0 ? lines.join('\n') : '（无返回数据）';
}

/**
 * 成功时展示影响面而非原始数据，用户关心的是改了哪里、改了多少。
 *
 * 位置放在前面：单元格数量能说明规模，但说不清落在哪几行哪几列，
 * 而后者才是用户核对结果时第一眼要找的。
 */
/**
 * 成功摘要里的范围可点进 Excel。
 *
 * 审批卡已经这样做了。操作卡上的地址原先是死文字，核对要自己去表里翻。
 * 规范要求凡是已经译成行列位置的，都是这个动作的入口。
 */
function fillSuccessState(state, data) {
  const summary = describeSuccess(data);
  const sheet = data?.sheet;
  const address = data?.address ?? data?.source_range ?? '';
  const where = rangeLabel(address);
  const jump = where !== '' ? addRangeJumpControl(sheet, address, where) : null;
  if (!jump) {
    state.textContent = summary;
    return;
  }

  state.textContent = '';
  // 把位置那一段换成可点控件，其余原文保留。
  const idx = summary.indexOf(where);
  if (idx <= 0) {
    state.append(jump);
    const rest = summary.slice(where.length);
    if (rest) { state.append(rest); }
    return;
  }

  state.append(summary.slice(0, idx), jump, summary.slice(idx + where.length));
}

function describeSuccess(data) {
  if (!data || typeof data !== 'object') {
    return isWordHostUi() ? '完成文档操作' : '完成';
  }

  if (isWordHostUi()) {
    return describeWordResult(data);
  }

  const where = rangeLabel(data.address ?? data.source_range ?? '');

  if (typeof data.dimensions_adjusted === 'number') {
    const unit = data.target === 'rows' ? '行' : '列';
    return where === ''
      ? `已调整 ${data.dimensions_adjusted} ${unit}的整${unit}`
      : `已调整 ${where} 的整${unit}（${data.dimensions_adjusted}）`;
  }
  if (typeof data.rows_adjusted === 'number' && typeof data.columns_adjusted === 'number') {
    return where === ''
      ? `已适配 ${data.rows_adjusted} 行 × ${data.columns_adjusted} 列（整行整列）`
      : `已适配 ${where}（整行整列）`;
  }
  // 面板「适配」回传 rows/columns，与模型工具的 rows_adjusted 不是同一份字段。
  // 有 horizontalAlignment 就是适配：动的是整行整列，不能报成 N 个单元格。
  if (data.horizontalAlignment && typeof data.rows === 'number' && typeof data.columns === 'number') {
    return where === ''
      ? `已适配 ${data.rows} 行 × ${data.columns} 列（整行整列）`
      : `已适配 ${where}（整行整列）`;
  }
  if (typeof data.cells_written === 'number') {
    return where === ''
      ? `已写入 ${data.cells_written} 个单元格`
      : `已写入 ${where}，共 ${data.cells_written} 个单元格`;
  }
  if (typeof data.cells_affected === 'number') {
    return where === ''
      ? `影响 ${data.cells_affected} 个单元格`
      : `影响 ${where}，共 ${data.cells_affected} 个单元格`;
  }
  if (typeof data.rows === 'number' && typeof data.columns === 'number') {
    return where === ''
      ? `共 ${data.rows} 行 × ${data.columns} 列`
      : `${where}，共 ${data.rows * data.columns} 个单元格`;
  }
  if (data.created_sheet) {
    return `已创建 ${data.created_sheet}`;
  }
  if (data.new_name) {
    return `已改名为 ${data.new_name}`;
  }
  if (data.table_name) {
    return `已创建表格 ${data.table_name}`;
  }
  // 自动调整、创建图表这类结果没有计数字段，至少说明作用位置。
  return where === '' ? '完成' : where;
}

/**
 * 审批卡片的影响范围说明。
 *
 * 加载项探到具体范围时回传 impactRange（表名、地址、单元格数），
 * 由这里组装文案：位置说明与工具卡片同源，两处措辞才不会一处说位置、
 * 一处说尺寸。探测失败或操作本就没有范围（如新增工作表）时，
 * 退回加载项给的 impact 文本。
 */
function describeImpact(message) {
  if (isWordHostUi()) {
    const target = describeWordTarget(message?.args?.target);
    return target ? `文档目标：${target}` : (message.impact ?? '需要明确文档目标');
  }
  const target = message.impactRange;
  if (!target || !target.address) {
    return message.impact ?? '';
  }

  const where = describeRange(target.address);
  const prefix = target.sheet ? `${target.sheet} 的 ` : '';
  const size = typeof target.cells === 'number' ? `，共 ${target.cells} 个单元格` : '';
  return `${prefix}${where}${size}`;
}

/**
 * 「将改成什么」的对照表。
 *
 * 审批卡此前只报形状（values: 20 行 × 3 列），用户批准的是一个尺寸而不是内容。
 * 加载项为估算影响已经把当前值读出来过一次，这里用的正是那一次的结果。
 *
 * 三件事必须分得开：
 *   一、空单元格显示「（空）」——原值是空、新值是 0 是一次正常写入，
 *       与「读不到当前值」不是一回事，合成一个样子会让人把后者看成前者；
 *   二、截断要用文字说出来。只靠少画几行，看起来就是一张完整的表；
 *   三、读不到当前值时不画空表，直接说读不到。
 */
function buildPreviewTable(preview) {
  if (!preview) {
    return null;
  }

  const wrap = document.createElement('div');
  wrap.className = 'approval-preview';

  if (preview.formattingMixed) {
    const note = document.createElement('div');
    note.className = 'approval-preview-note';
    note.textContent = '当前这片范围的格式逐项都不一样，改过之后无法完整还原。';
    wrap.append(note);
    return wrap;
  }

  if (preview.currentUnreadable) {
    const note = document.createElement('div');
    note.className = 'approval-preview-note';
    note.textContent = '读不到这片范围当前的内容（范围过大或地址无法解析），因此无法给出前后对照。';
    wrap.append(note);
    return wrap;
  }

  const cells = Array.isArray(preview.cells) ? preview.cells : [];
  const discarding = preview.kind === 'merge' || preview.kind === 'clear';

  // 抹除类即使一格值都没有也要出一句话：「这片范围里没有会丢的值」
  // 与「没给对照」是两件事，后者会让人以为功能没生效。
  if (cells.length === 0) {
    if (!discarding) {
      return null;
    }

    const none = document.createElement('div');
    none.className = 'approval-preview-note';
    none.textContent = preview.kind === 'merge'
      ? '这片范围里没有会被丢弃的值。'
      : '这片范围里没有要清掉的内容。';
    wrap.append(none);
    return wrap;
  }

  const table = document.createElement('table');
  table.className = 'approval-preview-table';

  const head = document.createElement('tr');
  // 抹除类的第三列不是「将改为」——那一侧是确定的（清完是空，
  // 合并只留左上角）。写成「将改为」会让人以为还有别的结果可选。
  const headers = discarding
    ? ['位置', '会丢掉', '之后']
    : ['位置', '现在', '将改为'];
  for (const label of headers) {
    const th = document.createElement('th');
    th.textContent = label;
    head.append(th);
  }
  table.append(head);

  for (const cell of cells) {
    const row = document.createElement('tr');

    const where = document.createElement('td');
    where.className = 'approval-preview-where';
    // 位置用范围内的相对行列，不是工作表行号：卡片顶上已经写了绝对范围，
    // 这里再写一次绝对地址反而要用户自己去减。
    where.textContent = `第 ${cell.row} 行 第 ${cell.column} 列`;

    const before = document.createElement('td');
    before.className = cell.beforeEmpty ? 'approval-preview-empty' : '';
    before.textContent = cell.beforeEmpty ? '（空）' : cell.before;

    const after = document.createElement('td');
    after.className = cell.afterEmpty ? 'approval-preview-empty' : '';
    after.textContent = cell.afterEmpty ? '（空）' : cell.after;

    row.append(where, before, after);
    table.append(row);
  }

  wrap.append(table);

  // 丢几个值必须是范围内的总数，且要显著。
  //
  // 合并是唯一静默丢值的写操作，事后没有痕迹可查；这个数字在用户点
  // 「允许」之前只有这里能看到。只报卡上列出的那几格是不够的——
  // 列出 8 行而实际丢 300 格时，按卡面判断就错了一个量级。
  if (discarding && typeof preview.discardedValues === 'number' && preview.discardedValues > 0) {
    const total = document.createElement('div');
    total.className = 'approval-preview-total';
    total.textContent = preview.kind === 'merge'
      ? `合并会丢弃 ${preview.discardedValues} 个有内容的单元格，只保留左上角那一个。`
      : `会清掉 ${preview.discardedValues} 个有内容的单元格。`;
    wrap.append(total);
  }

  if (typeof preview.omittedCells === 'number' && preview.omittedCells > 0) {
    const more = document.createElement('div');
    more.className = 'approval-preview-more';
    // 报格数而不是行数：截断同时发生在行和列两个方向上。
    more.textContent = discarding
      ? `其中 ${preview.omittedCells} 个未列出。`
      : `另有 ${preview.omittedCells} 个单元格未列出。`;
    wrap.append(more);
  }

  return wrap;
}

function buildWordPreview(preview) {
  if (!preview || preview.supported !== true) { return null; }
  const wrap = document.createElement('div');
  wrap.className = 'approval-preview';
  const heading = document.createElement('div');
  heading.className = 'approval-preview-note';
  heading.textContent = `文档目标：${preview.target || `Story=${preview.story}，Start=${preview.start}，End=${preview.end}`}`;
  const table = document.createElement('table');
  table.className = 'approval-preview-table';
  const header = document.createElement('tr');
  for (const label of ['目标', '修改前', '修改后']) {
    const cell = document.createElement('th'); cell.textContent = label; header.append(cell);
  }
  const row = document.createElement('tr');
  const target = document.createElement('td'); target.textContent = preview.target || '明确文档目标';
  const before = document.createElement('td'); before.textContent = preview.before ?? '（空）';
  const after = document.createElement('td'); after.textContent = preview.after ?? '（空）';
  row.append(target, before, after); table.append(header, row); wrap.append(heading, table); return wrap;
}

function addRangeJumpControl(sheet, address, label) {
  if (!address) {
    return null;
  }

  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'range-jump';
  button.textContent = label;
  button.title = '跳到 Excel 中的这个范围（会改变当前选区）';
  button.addEventListener('click', async (event) => {
    event.preventDefault();
    event.stopPropagation();
    button.disabled = true;
    try {
      const result = await request('sheet.goto', { sheet, address });
      if (!result?.ok) {
        addNotice(result?.message ?? '无法跳转到该范围。', 'error');
      }
    } catch (error) {
      addNotice(`无法跳转到该范围：${error.message}`, 'error');
    } finally {
      button.disabled = false;
    }
  });
  return button;
}

function addApprovalCard(message) {
  const card = document.createElement('div');
  card.className = 'approval';

  const title = document.createElement('div');
  title.className = 'approval-title';
  title.textContent = `需要确认：${toolLabel(message.tool)}`;

  const risk = document.createElement('div');
  risk.className = 'approval-risk';
  risk.textContent = RISK_LABELS[message.risk] ?? message.risk;

  const impact = document.createElement('div');
  impact.className = 'approval-impact';
  const impactText = describeImpact(message);
  impact.textContent = impactText ? '影响范围：' : '';
  const rangeJump = message.impactRange?.address
    ? addRangeJumpControl(message.impactRange.sheet, message.impactRange.address, impactText)
    : null;
  if (rangeJump) {
    impact.append(rangeJump);
  } else if (impactText) {
    impact.append(impactText);
  }

  const note = document.createElement('div');
  note.className = 'approval-impact-note';
  if (message.impactNote) {
    note.textContent = message.impactNote;
  }

  const preview = isWordHostUi()
    ? buildWordPreview(message.hostPreview)
    : buildPreviewTable(message.preview);

  const args = document.createElement('pre');
  args.className = 'approval-args';
  args.textContent = summarizeArgs(message.args);

  const actions = document.createElement('div');
  actions.className = 'approval-actions';

  const approve = document.createElement('button');
  approve.type = 'button';
  approve.className = 'btn btn-primary';
  approve.textContent = '允许';

  const approveAll = document.createElement('button');
  approveAll.type = 'button';
  approveAll.className = 'btn';
  approveAll.textContent = '本轮同类允许';
  approveAll.title = isWordHostUi()
    ? '只允许本轮当前文档中同一类目标操作；不会扩大到未明确指定的 Story 或全文。'
    : '只允许本轮中同一工作表、同一类操作；不会允许新建工作表、建表或建图。';

  const approveStructure = document.createElement('button');
  approveStructure.type = 'button';
  approveStructure.className = 'btn';
  approveStructure.textContent = '含结构允许';
  approveStructure.title = isWordHostUi()
    ? '允许本轮当前文档中同类的段落、表格、节和分页结构操作；仍要求工具明确 Story 和目标。'
    : '允许本轮在当前工作表继续进行同类操作，并允许新建或重命名工作表、建表、建图。';

  const reject = document.createElement('button');
  reject.type = 'button';
  reject.className = 'btn btn-danger';
  reject.textContent = '拒绝';

  const settle = async (approved, approveRest, approveStructureRest = false) => {
    approve.disabled = true;
    approveAll.disabled = true;
    approveStructure.disabled = true;
    reject.disabled = true;

    try {
      await request('approval.respond', {
        id: message.id,
        approved,
        approveRest,
        approveStructureRest,
      });
      // 用户已决定，处理随即继续，重新显示进展。
      showPending();
      const outcome = document.createElement('div');
      outcome.className = approved ? 'approval-outcome is-ok' : 'approval-outcome is-error';
      outcome.textContent = approved
        ? (approveStructureRest
          ? (isWordHostUi() ? '已允许，本轮当前文档含结构的后续操作不再询问' : '已允许，本轮当前表含结构的后续操作不再询问')
          : (approveRest ? (isWordHostUi() ? '已允许，本轮当前文档同类操作不再询问' : '已允许，本轮当前表同类操作不再询问') : '已允许'))
        : '已拒绝';
      actions.replaceWith(outcome);
    } catch (error) {
      addNotice(`回复审批失败：${error.message}`, 'error');
    }
  };

  approve.addEventListener('click', () => void settle(true, false));
  approveAll.addEventListener('click', () => void settle(true, true));
  approveStructure.addEventListener('click', () => void settle(true, true, true));
  reject.addEventListener('click', () => void settle(false, false));

  actions.append(approve, approveAll, approveStructure, reject);

  // 等用户决定期间撤掉进展指示器：此刻真正在等的是用户，
  // 旁边还跳着「正在处理」既不准确，也会分散对审批卡片的注意力。
  clearPending();

  card.append(title, risk);
  if (impactText) { card.append(impact); }
  if (message.impactNote) { card.append(note); }
  // 对照放在参数之前：参数区只报形状（values: 20 行 × 3 列），
  // 而用户要决定的是内容，先看见内容才有决定可做。
  if (preview) { card.append(preview); }
  if (args.textContent) { card.append(args); }
  card.append(actions);

  mountToTranscript(card);
  scrollToBottom();
}

/**
 * 在对话流中间插入一条居中的胶囊提示。
 *
 * 不做成气泡：错误与系统提示不是对话中任何一方说的话，套上气泡会让人
 * 误当成模型的回复——错误文本里常常就带着模型口吻。居中胶囊是系统级
 * 消息的通行表达，与左右两侧的对话一眼可分。
 */
function addNotice(text, variant = 'info') {
  const notice = document.createElement('div');
  notice.className = `notice notice-${variant}`;
  notice.textContent = text;
  mountToTranscript(notice);

  // 纯圆角只适合单行：文字换行后首末行会被挤进弧内。
  // 必须等插入文档后才能量到真实行数。
  markMultiline(notice);

  scrollToBottom();
  return notice;
}

/**
 * 一轮正常收尾时插一条「已完成」。
 *
 * 位置与错误、停止、步数上限同处一处——都是对话流中间的居中胶囊。
 * 一轮怎么结束的只有这一类消息在说，正常结束却原先什么都不留：
 * 于是「模型说完了」与「中途断了但最后一段话看起来像结论」在屏幕上
 * 长得一模一样，只能靠日志区分（见 chatsheet-turn-ended-early-diagnosis）。
 * 补上这条之后，没有它就是没正常收完。
 *
 * 只在 turn-complete 时插。加载项的四条终止路径互斥：stalled、step-limit、
 * stopped、error 都各自 return 而不再发 turn-complete，因此不会出现
 * 「已停止」紧跟着「已完成」这种自相矛盾的收尾。
 */
function markTurnComplete() {
  const notice = addNotice('已完成', 'ok');
  notice.classList.add('notice-complete');
  notice.title = '这一轮已正常结束。没有这一条就说明中途断了（被停止、达上限或出错）';
  return notice;
}

/** 水平对齐的三个选项。顺序与浮层里的 DOM 一致，默认取 center。 */
const FIT_ALIGNMENTS = {
  left: '靠左',
  center: '居中',
  right: '靠右',
};

/**
 * 当前选中的水平对齐。只存在会话内：适配是即时动作，
 * 记住上次选择方便连续排版，但不值得为它写一条持久设置。
 */
let fitAlignment = 'center';

/**
 * 适配按钮与对齐浮层。
 *
 * 点按钮就按当前对齐直接适配。原先点按钮只是展开浮层，必须再点一次
 * 「居中」之类的选项才动手——那让默认对齐形同不存在：想按默认排一次表要点两下，
 * 连续排几张表就是连续的两下。既然按钮上写着「适配」，点它就该适配。
 *
 * 浮层退到「换成哪一种」这一个职责上：悬停展开顺手，键盘按上下方向键展开，
 * 选中哪一项就记住它并立刻按那一项适配（换对齐本身也是一次适配意图，
 * 选完还要再点一次按钮等于把一件事拆成两步）。
 */
function initFit() {
  const wrap = document.getElementById('fit-wrap');
  const button = document.getElementById('fit');
  const pop = document.getElementById('fit-pop');
  if (!wrap || !button || !pop) {
    return;
  }

  const items = [...pop.querySelectorAll('.fit-item')];

  const isOpen = () => !pop.hidden;

  const setOpen = (open) => {
    pop.hidden = !open;
    button.setAttribute('aria-expanded', open ? 'true' : 'false');
    if (open) {
      markActive();
    }
  };

  // 标出当前选项。三选一里「现在是哪个」是用户最先要看的信息。
  // 按钮上的说明同步改写：点下去会按哪一种对齐排版，光看图标是看不出来的。
  const markActive = () => {
    for (const item of items) {
      const active = item.dataset.align === fitAlignment;
      item.classList.toggle('is-active', active);
      item.setAttribute('aria-checked', active ? 'true' : 'false');
    }

    const label = FIT_ALIGNMENTS[fitAlignment] ?? '居中';
    button.title =
      `适配当前表（${label}）：水平${label}、垂直居中，列宽上限 100 并自动换行后调整行高。上下方向键可换对齐方式`;
    button.setAttribute('aria-label', `适配当前表，当前对齐：${label}`);
  };

  markActive();

  // 悬停展开。离开容器才收起：浮层紧贴按钮且有透明补丁接缝，
  // 鼠标从按钮移向选项的途中不会误判为移开。
  wrap.addEventListener('mouseenter', () => setOpen(true));
  wrap.addEventListener('mouseleave', () => setOpen(false));

  // 点按钮直接适配，不再只是开关浮层。浮层此时若因悬停开着就收起——
  // 动作已经发出，留着一张待选菜单只会让人以为还要再选一次。
  button.addEventListener('click', () => {
    setOpen(false);
    const source = window.__chatsheetFitSource === 'ribbon' ? 'ribbon' : 'panel';
    window.__chatsheetFitSource = undefined;
    void runFit(fitAlignment, source);
  });

  // 键盘用户换对齐的入口。点击被适配占用之后，方向键就是唯一能展开浮层的键；
  // 浮层向上弹出，上下两个方向都收下，不必先猜它在按钮的哪一侧。
  button.addEventListener('keydown', (event) => {
    if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') {
      return;
    }

    event.preventDefault();
    setOpen(true);

    // 焦点落在当前项上，接着按 Tab 或方向键就能挑另一种，
    // 直接回车则是「就按现在这种」，与点按钮一致。
    const current = items.find((item) => item.dataset.align === fitAlignment);
    (current ?? items[0])?.focus();
  });

  for (const item of items) {
    item.addEventListener('click', () => {
      const align = item.dataset.align;
      if (!FIT_ALIGNMENTS[align]) {
        return;
      }

      fitAlignment = align;
      markActive();
      setOpen(false);
      void runFit(align, 'panel');
    });
  }

  // 点击浮层外部关闭。捕获阶段以免被内部 stopPropagation 阻断。
  document.addEventListener(
    'click',
    (event) => {
      if (isOpen() && !wrap.contains(event.target)) {
        setOpen(false);
      }
    },
    true,
  );

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && isOpen()) {
      setOpen(false);
      button.focus();
    }
  });
}

/** 手动操作卡片的临时标识计数。撤销标识要等宿主执行完才知道。 */
let fitSequence = 0;

/**
 * 执行一次适配。
 *
 * 不进对话历史：这是确定性的排版动作，点按钮已经表达了意图，模型不必知道。
 * 但在对话流里按工具卡片呈现——它和模型发起的写入是同一类事（改了哪个范围、
 * 影响多少格、能不能撤销），拆成两种样式只会让人对着两处找同一种信息。
 * 区别靠「手动」标记和边条颜色说明，而不是靠换一套结构。
 *
 * 撤销入口是必须的：加载项经 COM 的写入会清空 Excel 自己的撤销栈，
 * 用户按 Ctrl+Z 是拿不回来的。
 */
async function runFit(alignment, source = 'panel') {
  const button = document.getElementById('fit');
  const label = FIT_ALIGNMENTS[alignment] ?? '居中';

  if (button) { button.disabled = true; }
  setStatus(`正在适配当前表（${label}）…`);

  // 卡片先上屏再发请求。整表适配可能跑上一两分钟，这段时间里
  // 摘要行的「执行中…」就是进度反馈，与模型发起的工具一致。
  const card = addToolCard({
    id: `fit-pending-${++fitSequence}`,
    name: 'fit_range',
    args: { horizontal_alignment: alignment },
    manual: true,
    source,
    // 手动操作没有加载项给的风险等级，由这里声明。适配改的是排版，
    // 算「改」——统计里把它当读取会让「那一轮动过表没有」答错。
    risk: 'Write',
  });

  try {
    // 默认 30 秒会误报超时，而宿主那边其实还在正常执行——这种失败最难排查。
    const result = await request(
      'sheet.fit',
      { horizontalAlignment: alignment },
      { timeout: 300000 },
    );

    if (!result?.ok) {
      fillToolCard(card, { ok: false, error: result?.message ?? '适配失败。' });
      return;
    }

    // 撤销要按宿主登记的记录标识来点，卡片的标识随之改写。
    if (result.undoId) {
      card.dataset.toolId = result.undoId;
    }

    const where = rangeLabel(result.address);

    fillToolCard(card, {
      id: result.undoId,
      ok: true,
      data: result,
      canUndo: Boolean(result.undoId),
      undoSummary: `适配 ${where === '' ? (result.address ?? '') : where}`.trim(),
    });

    // 没有撤销入口时把原因写进卡片：只是少个按钮的话看起来像功能坏了，
    // 而它其实是保不住完整快照时的有意取舍。
    if (!result.undoId && result.undoUnavailableReason) {
      appendToolNote(card, result.undoUnavailableReason);
    }
  } catch (error) {
    fillToolCard(card, { ok: false, error: error.message });
  } finally {
    // 结果已经写在卡片上，状态行不再复述一遍。
    setStatus('');
    if (button) { button.disabled = false; }
  }
}

/**
 * 功能区快捷入口用的三个函数。挂在 window 上是因为宿主经 ExecuteScript 调它们，
 * 模块作用域里的名字宿主看不见。
 *
 * 适配仍然点真实按钮：当前对齐、操作卡片、错误与撤销都走已经验过的同一条路。
 * 用一次性标记告诉 click 处理器这次来自功能区，卡片据此打上 data-source=ribbon，
 * 功能区的撤销按钮就只认这些卡——面板里点出来的、模型发起的，都不碰。
 */
function ribbonFitCards() {
  // 不靠 CSS 属性选择器：面板单测的假 DOM 只认类名，
  // 真页面里从 transcript 过滤也更窄，不会误伤别的区域。
  const root = document.getElementById('transcript');
  if (!root || typeof root.querySelectorAll !== 'function') {
    return [];
  }

  return [...root.querySelectorAll('.tool-card')].filter(
    (card) => card.dataset.source === 'ribbon',
  );
}

function newestRibbonUndoTarget() {
  const cards = ribbonFitCards();
  for (let i = cards.length - 1; i >= 0; i--) {
    const button = cards[i].querySelector('.tool-undo');
    if (button && button.dataset.undone !== 'true') {
      return { card: cards[i], button };
    }
  }

  return null;
}

function ribbonUndoSummary(target) {
  if (!target) {
    return '';
  }

  const title = target.button.title ?? '';
  const trimmed = title.replace(/^撤销：/, '').trim();
  if (trimmed !== '') {
    return trimmed;
  }

  const name = target.card.querySelector('.tool-name')?.textContent ?? '';
  return name;
}

function ribbonUndoStateText() {
  const cards = ribbonFitCards();
  let count = 0;
  let newest = null;
  for (let i = cards.length - 1; i >= 0; i--) {
    const button = cards[i].querySelector('.tool-undo');
    if (button && button.dataset.undone !== 'true') {
      count += 1;
      if (!newest) {
        newest = { card: cards[i], button };
      }
    }
  }

  const warned = newest?.button.dataset.overlapWarned === 'true';
  const summary = ribbonUndoSummary(newest);
  return `count=${count}|summary=${summary}|warned=${warned ? 'true' : 'false'}`;
}

function exposeRibbonShortcuts() {
  window.__chatsheetRibbonFit = () => {
    const button = document.getElementById('fit');
    if (!button) {
      return 'no-button';
    }

    if (button.disabled) {
      return 'busy';
    }

    window.__chatsheetFitSource = 'ribbon';
    button.click();
    return 'clicked';
  };

  window.__chatsheetRibbonUndo = () => {
    const target = newestRibbonUndoTarget();
    if (!target) {
      return 'nothing';
    }

    if (target.button.disabled) {
      return 'busy';
    }

    const warnedBefore = target.button.dataset.overlapWarned === 'true';
    target.button.click();
    return warnedBefore ? 'forced' : 'clicked';
  };

  window.__chatsheetRibbonUndoState = () => ribbonUndoStateText();
}

/**
 * 往卡片折叠区追加一段说明文字。
 * 展开才看得到，因此只放「为什么没有某个按钮」这类不影响判断结果的补充。
 */
function appendToolNote(card, text) {
  const body = card.querySelector('.tool-body');
  if (!body) {
    return;
  }

  const note = document.createElement('div');
  note.className = 'tool-note';
  note.textContent = text;
  body.append(note);
}

/* ---- 操作按轮次成组 ---- */

/**
 * 当前这一批操作。下一轮开始时收成一组。
 *
 * 为什么按「批」而不是严格按轮：面板上点「适配」产生的卡片可能落在两轮之间，
 * 它不属于任何一轮，但确实发生在这段时间里。按批收集就不必为它另立一类——
 * 从上一轮开始到下一轮开始之间发生的操作是同一批，谁发起的由卡片上的
 * 「手动」标记区分（见 addToolCard）。
 */
let opsBatch = [];

/** 当前批次开始时的轮次号。0 表示这批还没跑过任何一轮，摘要写「手动操作」。 */
let opsBatchTurn = 0;

/** 已经形成的组。撤销后要回来刷新摘要，所以得记住卡片与组的对应关系。 */
const opsGroups = [];

/** 轮次计数。只用于摘要文案，不参与任何逻辑判断。 */
let turnNumber = 0;

/** 把一张卡片收进当前批次。 */
function joinOpsBatch(card, payload) {
  opsBatch.push({
    card,
    name: toolLabel(payload.name),
    // 手动操作没有加载项给的风险等级，由发起方自己声明（见 runFit）。
    risk: payload.risk ?? 'Read',
  });
}

/**
 * 一张卡片当前的状态。摘要要据此计数，所以只认卡片上的类名与状态文字，
 * 不另存一份——撤销是异步发生的，另存的那份一定会过期。
 */
function cardOutcome(entry) {
  if (entry.card.classList.contains('is-undone')) { return 'undone'; }
  if (entry.card.classList.contains('is-error')) { return 'error'; }
  return 'ok';
}

/**
 * 组的摘要文字。
 *
 * 给统计而不只给条数：用户合上一轮之后要判断的是「那一轮动过表没有」，
 * 「3 改」直接回答了它，「4 个操作」回答不了。
 */
function opsSummaryText(record) {
  const total = record.entries.length;

  // 按「改」「读」计数。顺序固定为改在前：动过表是用户更关心的那一件。
  const tally = new Map();
  let failed = 0;
  let undone = 0;
  for (const entry of record.entries) {
    const unit = RISK_TALLY[entry.risk] ?? '读';
    tally.set(unit, (tally.get(unit) ?? 0) + 1);

    const outcome = cardOutcome(entry);
    if (outcome === 'error') { failed += 1; }
    if (outcome === 'undone') { undone += 1; }
  }

  const parts = [];
  for (const unit of ['改', '读']) {
    if (tally.has(unit)) { parts.push(`${tally.get(unit)} ${unit}`); }
  }
  if (failed > 0) { parts.push(`${failed} 失败`); }
  if (undone > 0) { parts.push(`${undone} 已撤销`); }

  const lead = record.turn > 0 ? `第 ${record.turn} 轮` : '手动操作';
  return `${lead} ${total} 个操作（${parts.join('，')}）`;
}

/** 悬停说明：逐条列出组里有什么，不展开也能确认。 */
function opsSummaryTitle(record) {
  const marks = { ok: '', error: '（失败）', undone: '（已撤销）' };
  const lines = record.entries.map((entry) => `· ${entry.name}${marks[cardOutcome(entry)]}`);
  return ['点击展开这一组操作', ...lines].join('\n');
}

/** 重画一个组的摘要。撤销、恢复之后要调用，否则计数停在成组那一刻。 */
function renderOpsSummary(record) {
  record.label.textContent = opsSummaryText(record);
  record.head.title = opsSummaryTitle(record);
  // 组里有失败时标红：失败的卡片自己会展开，但它在收起的组里看不见——
  // 组上不留记号的话，用户合上一轮之后就再也不知道那轮里出过错。
  record.group.classList.toggle(
    'is-error',
    record.entries.some((entry) => cardOutcome(entry) === 'error'),
  );
}

/** 卡片状态变了，刷新它所在的组。不在任何组里（当前批次）时什么也不做。 */
function refreshOpsGroupFor(card) {
  const record = opsGroups.find((r) => r.entries.some((entry) => entry.card === card));
  if (record) { renderOpsSummary(record); }
}

/**
 * 把当前批次收成一个折叠组，落在这一批内容之后。
 *
 * 时机是「下一轮开始时」而不是「当前轮结束时」：一轮刚跑完，用户往往正要看
 * 结果、点撤销。收在那个时候等于刚做完就把东西收走。
 *
 * 组追加在对话流末尾，也就是这一批全部内容（含模型的收尾回复）之后。
 * 不搬到对话流最底部：操作要跟着它所属的那轮对话走，否则「查看某轮对应的
 * 操作」就成了去别处找。
 */
function sealOpsBatch() {
  const entries = opsBatch;
  opsBatch = [];

  // 这一批没有操作（纯对话的一轮）就不留空组。
  if (entries.length === 0) {
    return;
  }

  const group = document.createElement('details');
  group.className = 'ops-group';

  const head = document.createElement('summary');
  head.className = 'ops-head';

  const label = document.createElement('span');
  label.className = 'ops-label';

  const actions = document.createElement('span');
  actions.className = 'ops-actions';

  const restore = document.createElement('button');
  restore.type = 'button';
  restore.className = 'ops-restore';
  restore.textContent = '还原';
  restore.title = '把这几张卡片放回对话流原来的位置，按发生顺序穿插在回复之间';

  actions.append(restore);
  head.append(label, actions);

  const body = document.createElement('div');
  body.className = 'ops-body';
  // 卡片从对话流搬进组里。append 会把它们从原父节点摘走，
  // 因此不必先逐个 remove——原位不会留下空节点。
  body.append(...entries.map((entry) => entry.card));

  group.append(head, body);

  const record = { group, head, label, entries, turn: opsBatchTurn };
  mountToTranscript(group);

  // 「已完成」是这一轮的收尾，组不该排在它下面：那样这一轮读下来是
  //   …回复 → 已完成 → 一组操作
  // 收尾之后又冒出内容，而用户扫「这轮完没完」正是找那条胶囊。
  // 重新挂一次把它移到末尾——mountToTranscript 会同时刷新它的挂载序号，
  // 于是还原时按序号重排的结果与此刻的顺序一致，两处不会打架。
  const tail = transcript.children[transcript.children.length - 2];
  if (tail?.classList?.contains('notice-complete')) {
    mountToTranscript(tail);
  }

  opsGroups.push(record);
  renderOpsSummary(record);

  restore.addEventListener('click', (event) => {
    // 按钮在 summary 里，不拦下来点一次会连带展开或收起这个组。
    event.preventDefault();
    event.stopPropagation();
    restoreOpsGroup(record);
  });
}

/**
 * 解散一个组，卡片回到对话流里原来的位置。
 *
 * 按挂载序号重排整条对话流，而不是把卡片插回某个锚点后面：锚点自己可能
 * 已经被收进了别的组，那时插到哪里都不对。序号是每个节点自己的属性，
 * 无论此前形成过几个组、还原过几次，重排的结果都是当初的上屏顺序。
 *
 * 还原后这些卡片不再进入任何批次：还原是明确的「我要看原位」，
 * 下一轮开始时再把它们收回去等于撤销了用户刚做的事。
 */
function restoreOpsGroup(record) {
  const index = opsGroups.indexOf(record);
  if (index >= 0) { opsGroups.splice(index, 1); }

  const seq = (node) => Number(node.dataset?.seq ?? 0);

  // 组本身要从对话流里去掉，卡片则加回来，然后整体按序号排。
  const rest = [...transcript.children].filter((node) => node !== record.group);
  const merged = [...rest, ...record.entries.map((entry) => entry.card)]
    .sort((a, b) => seq(a) - seq(b));

  transcript.replaceChildren(...merged);

  // 不滚到底：用户点还原正是为了看回上面那几处，把视口拉走等于白点一次。
  void logToHost(
    `还原操作组：第 ${record.turn} 轮 ${record.entries.length} 张卡片回到原位`,
  );
}

/** 超过一行时换成圆角块。用行高判断，避免把圆角开得过大导致文字压边。 */
function markMultiline(node) {
  const lineHeight = parseFloat(getComputedStyle(node).lineHeight);
  if (Number.isFinite(lineHeight) && node.clientHeight > lineHeight * 1.6) {
    node.classList.add('is-multiline');
  }
}

/**
 * 底部那行短暂状态。只承担「刚做了什么」这类过渡说明。
 *
 * 错误不走这里：这一行与本轮用量并排，宽度只有半栏，稍长的报错就被挤成
 * 一截看不全的文字；它还会被下一次状态更新直接覆盖，用户回头找不到。
 * 错误一律用 addNotice 插进对话流，与出错的那一步留在同一位置。
 */
function setStatus(text) {
  statusLine.textContent = text ?? '';
  statusLine.className = 'status';
  // 与排队条争位时这一行会被压到只剩两三个字（实测 320px 栏宽下如此），
  // 压缩优先让给状态是有意的——排队项上有取消按钮，裁掉就没法点了。
  // 代价是状态可能读不全，所以原文放进悬停说明。
  statusLine.title = text ?? '';
}

/**
 * 切换忙闲。
 *
 * 发送按钮在忙时不禁用而是改变含义（见 updateSendAffordance）。禁用会让
 * 唯一的中断入口在最需要它的时候没法点——上一版另有一个停止按钮，合并后
 * 若照旧禁用，运行中就完全无从中断了。
 *
 * 输入框也不再禁用：处理中写下一步是常态，写好的内容进队列，
 * 上一轮结束自动接着跑。禁用会把用户逼成「干等」或「记在别处」。
 */
function setBusy(value) {
  busy = value;
  sendButton.classList.toggle('is-busy', value);
  syncDraftControls();
  updateSendAffordance();
  // 模型选择器要知道这件事：加载项在对话进行中会拒绝批量测试，而那条拒绝
  // 只落在宿主日志里，面板上不出任何东西。告诉它，好把「测试」先禁掉并说明原因。
  setPickerTurnInFlight(value);
  if (!value) {
    currentAssistant = null;
    currentThinking = null;
    // 兜底清理：异常路径可能不会走到 turn-complete。
    clearPending();
  }
}

/**
 * 按当前状态给发送按钮定含义。四种：
 *   空闲                        → 发送
 *   处理中 + 有输入              → 加入队列
 *   处理中 + 输入为空 + 队列非空  → 清空队列（不动正在跑的那一轮）
 *   处理中 + 输入为空 + 队列为空  → 停止
 *
 * 有输入时让「加入队列」压过其余含义：输入框里有字说明用户正打算安排下一步，
 * 此时点按钮几乎不可能是想中断。
 *
 * 「清空队列」这一态是后加的，为的是堵住一个真实的误操作：
 * 打字 → 回车入队（输入框随即清空）→ 再点一下按钮。最后这一下在合并按钮的
 * 旧规则下就是停止，于是正在跑的任务被掐掉、队列也一并清空，而用户以为自己
 * 只是在发消息。分出这一态后，刚入队时按钮只可能清队列；要停当前任务得先清
 * 队列（有胶囊回执），再点第二下。破坏性动作因此需要两次明确的点击。
 *
 * 图形跟着含义换（见 app.css 的 is-queueing 与 is-clearing）：按钮上画什么，
 * 必须和点下去会发生的事一致，否则用户是照着图标点的，含义写在 title 里也
 * 来不及看。
 */
function updateSendAffordance() {
  const willQueue = busy && hasComposerContent();
  const willClearQueue = busy && !willQueue && queue.length > 0;

  sendButton.classList.toggle('is-queueing', willQueue);
  sendButton.classList.toggle('is-clearing', willClearQueue);

  if (!busy) {
    sendButton.title = '发送（Enter）';
    sendButton.setAttribute('aria-label', '发送');
    return;
  }

  if (willQueue) {
    const ahead = queue.length;
    sendButton.title = ahead === 0
      ? '正在处理，点击排到下一个（Enter 同样入队）'
      : `正在处理，点击排到第 ${ahead + 1} 位（Enter 同样入队）`;
    sendButton.setAttribute('aria-label', '加入队列');
    return;
  }

  if (willClearQueue) {
    sendButton.title = `正在处理，点击取消排队中的 ${queue.length} 条` +
      '（当前任务继续；清空后再点一下才是停止）';
    sendButton.setAttribute('aria-label', '清空队列');
    return;
  }

  sendButton.title = '正在处理，点击停止';
  sendButton.setAttribute('aria-label', '停止');
}

function updateUsage(payload) {
  if (!payload) { return; }
  const { promptTokens, completionTokens } = payload;
  usageLine.textContent = `本轮 ${promptTokens ?? 0} 入 / ${completionTokens ?? 0} 出`;
}

// 圆环周长，用于按占比设置 stroke-dashoffset。半径 9 → 2πr。
const RING_CIRCUMFERENCE = 2 * Math.PI * 9;

let contextState = null;
let compactPrompted = false;

/**
 * 更新上下文进度圆环。
 *
 * 到达阈值（后端判定，当前为 90%）时变色并提示是否压缩。
 * 阈值判断放在后端，避免界面与压缩逻辑各算一套导致不一致。
 */
function updateContextRing(payload, source = '未标注') {
  if (!payload) { return; }
  contextState = payload;

  const ring = document.getElementById('ring-value');
  const text = document.getElementById('context-percent');
  const button = document.getElementById('context-ring');
  if (!ring || !text || !button) { return; }

  const percent = payload.percent ?? 0;
  const ratio = Math.max(0, Math.min(1, payload.ratio ?? 0));

  ring.style.strokeDasharray = `${RING_CIRCUMFERENCE}`;
  ring.style.strokeDashoffset = `${RING_CIRCUMFERENCE * (1 - ratio)}`;
  text.textContent = String(percent);

  button.classList.toggle('is-near-limit', payload.nearLimit === true);
  button.title =
    `上下文 ${payload.used ?? 0} / ${payload.budget ?? 0} tokens（${percent}%）\n` +
    `达到 ${payload.threshold ?? 90}% 会自动压缩，也可点击立即压缩`;

  // 记录圆环的实际渲染值与来源：
  // 「圆环不动」需要对比推送值与渲染值，
  // 「同一状态重复渲染」则需要知道是哪条路径触发的。
  void logToHost(
    `上下文圆环[${source}]：${payload.used ?? 0}/${payload.budget ?? 0} tokens ` +
      `= ${percent}%（阈值 ${payload.threshold ?? 90}%，` +
      `dashoffset=${Math.round(RING_CIRCUMFERENCE * (1 - ratio))}/${Math.round(RING_CIRCUMFERENCE)}）` +
      `${payload.nearLimit ? ' 已达阈值' : ''}`,
  );

  // 达到阈值时只示警一次，不弹询问。
  //
  // 不问「是否压缩」的原因：系统在阈值处已自动压缩，此时再问已无实际选择，
  // 徒增一次打断。圆环变色说明正在临界，压缩完成后另有说明消息；
  // 想提前压缩可以直接点圆环。
  if (payload.nearLimit && !compactPrompted) {
    compactPrompted = true;
    addNotice(
      `上下文已达 ${payload.percent}%（${payload.used} / ${payload.budget} tokens），` +
        `超过 ${payload.threshold}% 阈值，正在自动压缩较早的记录。` +
        `若想保留完整上下文，可开新会话。`,
      'warn',
    );
  }

  // 回落到阈值以下后重置，下次再次接近时仍会示警。
  if (!payload.nearLimit) {
    compactPrompted = false;
  }
}

async function refreshContext(source = '主动查询') {
  try {
    updateContextRing(await request('context.state'), source);
  } catch (error) {
    // 取不到不影响对话，圆环保持上一次的值。
    // 但必须记录：静默吞掉异常会让「圆环一直是 0」变成无从排查的问题。
    void logToHost(`刷新上下文状态失败：${error.message}`, 'warn');
  }
}

const DRAFT_STYLE_LABELS = {
  follow_document: '沿用文档',
  minimal: '简洁',
  business: '商务',
  report: '报告',
};

function appendDraftRuns(parent, runs) {
  (Array.isArray(runs) ? runs : []).forEach((run) => {
    let node = document.createTextNode(typeof run?.text === 'string' ? run.text : '');
    if (run?.bold) {
      const strong = document.createElement('strong');
      strong.append(node);
      node = strong;
    }
    if (run?.italic) {
      const emphasis = document.createElement('em');
      emphasis.append(node);
      node = emphasis;
    }
    parent.append(node);
  });
}

function createDraftBlock(block) {
  const type = block?.type;
  if (type === 'table') {
    const table = document.createElement('table');
    const head = document.createElement('thead');
    const headerRow = document.createElement('tr');
    (Array.isArray(block.headers) ? block.headers : []).forEach((value) => {
      const cell = document.createElement('th');
      cell.textContent = String(value ?? '');
      headerRow.append(cell);
    });
    head.append(headerRow);
    table.append(head);
    const body = document.createElement('tbody');
    (Array.isArray(block.rows) ? block.rows : []).forEach((row) => {
      const rowElement = document.createElement('tr');
      (Array.isArray(row) ? row : []).forEach((value) => {
        const cell = document.createElement('td');
        cell.textContent = String(value ?? '');
        rowElement.append(cell);
      });
      body.append(rowElement);
    });
    table.append(body);
    return table;
  }

  if (type === 'bullets' || type === 'numbered') {
    const list = document.createElement(type === 'bullets' ? 'ul' : 'ol');
    (Array.isArray(block.items) ? block.items : []).forEach((value) => {
      const item = document.createElement('li');
      item.textContent = String(value ?? '');
      list.append(item);
    });
    return list;
  }

  const tag = type === 'title' ? 'h3'
    : type === 'heading' ? `h${Math.min(5, Math.max(3, (Number(block.level) || 1) + 2))}`
      : type === 'quote' ? 'blockquote' : 'p';
  const element = document.createElement(tag);
  if (type === 'title') { element.className = 'word-draft-title'; }
  appendDraftRuns(element, block?.runs);
  return element;
}

function syncDraftControls() {
  transcript?.querySelectorAll('.word-draft-card:not(.is-closed)').forEach((card) => {
    const disabled = busy || card.dataset.pending === 'true';
    card.querySelectorAll('button, select, textarea').forEach((control) => {
      control.disabled = disabled;
    });
  });
}

function addDocumentDraft(payload) {
  if (!payload?.id || !Array.isArray(payload.blocks)) { return; }
  let card = [...transcript.querySelectorAll('.word-draft-card')]
    .find((item) => item.dataset.draftId === payload.id);
  const isNew = !card;
  if (isNew) {
    card = document.createElement('section');
    card.className = 'word-draft-card';
    card.dataset.draftId = payload.id;

    const heading = document.createElement('div');
    heading.className = 'word-draft-heading';
    const title = document.createElement('strong');
    title.className = 'word-draft-heading-title';
    title.textContent = 'Word 草稿';
    const documentName = document.createElement('span');
    documentName.className = 'word-draft-document';
    heading.append(title, documentName);

    const target = document.createElement('p');
    target.className = 'word-draft-target';
    const selectedText = document.createElement('details');
    selectedText.className = 'word-draft-selection';
    const selectedSummary = document.createElement('summary');
    selectedSummary.textContent = '查看待替换内容';
    const selectedValue = document.createElement('div');
    selectedValue.className = 'word-draft-selection-text';
    selectedText.append(selectedSummary, selectedValue);

    const controls = document.createElement('div');
    controls.className = 'word-draft-controls';
    const styleLabel = document.createElement('label');
    styleLabel.className = 'word-draft-style-label';
    styleLabel.append(document.createTextNode('版式'));
    const styleSelect = document.createElement('select');
    styleSelect.className = 'word-draft-style';
    styleSelect.setAttribute('aria-label', '草稿版式');
    Object.entries(DRAFT_STYLE_LABELS).forEach(([value, label]) => {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = label;
      styleSelect.append(option);
    });
    styleLabel.append(styleSelect);

    const preview = document.createElement('div');
    preview.className = 'word-draft-preview';
    const actions = document.createElement('div');
    actions.className = 'word-draft-actions';
    const adjust = document.createElement('button');
    adjust.type = 'button';
    adjust.className = 'word-draft-button';
    adjust.textContent = '继续调整';
    const insert = document.createElement('button');
    insert.type = 'button';
    insert.className = 'word-draft-button is-primary';
    const discard = document.createElement('button');
    discard.type = 'button';
    discard.className = 'word-draft-button';
    discard.textContent = '丢弃';
    actions.append(adjust, insert, discard);

    const adjustForm = document.createElement('form');
    adjustForm.className = 'word-draft-adjust-form';
    adjustForm.hidden = true;
    const adjustment = document.createElement('textarea');
    adjustment.rows = 2;
    adjustment.maxLength = 2000;
    adjustment.placeholder = '描述要调整的内容';
    adjustment.setAttribute('aria-label', '草稿调整要求');
    const formActions = document.createElement('div');
    formActions.className = 'word-draft-form-actions';
    const applyAdjustment = document.createElement('button');
    applyAdjustment.type = 'submit';
    applyAdjustment.className = 'word-draft-button is-primary';
    applyAdjustment.textContent = '发送调整';
    const cancelAdjustment = document.createElement('button');
    cancelAdjustment.type = 'button';
    cancelAdjustment.className = 'word-draft-button';
    cancelAdjustment.textContent = '取消';
    formActions.append(applyAdjustment, cancelAdjustment);
    adjustForm.append(adjustment, formActions);
    controls.append(styleLabel);
    card.append(heading, target, selectedText, controls, preview, actions, adjustForm);

    styleSelect.addEventListener('change', async () => {
      const previous = card.dataset.style ?? 'follow_document';
      const style = styleSelect.value;
      card.dataset.pending = 'true';
      card.dataset.style = style;
      syncDraftControls();
      try {
        const result = await request('draft.style', { id: payload.id, style });
        if (!result?.ok) { throw new Error(result?.message ?? '无法更新草稿版式。'); }
      } catch (error) {
        card.dataset.style = previous;
        styleSelect.value = previous;
        addNotice(`调整草稿版式失败：${error.message}`, 'error');
      } finally {
        card.dataset.pending = 'false';
        syncDraftControls();
      }
    });

    adjust.addEventListener('click', () => {
      adjustForm.hidden = false;
      adjustment.focus();
    });
    cancelAdjustment.addEventListener('click', () => {
      adjustForm.hidden = true;
      adjustment.value = '';
    });
    adjustForm.addEventListener('submit', (event) => {
      event.preventDefault();
      const instruction = adjustment.value.trim();
      if (!instruction) { adjustment.focus(); return; }
      if (submit(instruction, payload.id)) {
        adjustment.value = '';
        adjustForm.hidden = true;
      }
    });
    insert.addEventListener('click', async () => {
      draftOperationPending = true;
      card.dataset.pending = 'true';
      syncDraftControls();
      try {
        const result = await request('draft.insert', { id: payload.id }, { timeout: 0 });
        if (!result?.ok) { throw new Error(result?.message ?? '无法插入草稿。'); }
      } catch (error) {
        addNotice(`插入草稿失败：${error.message}`, 'error');
      } finally {
        draftOperationPending = false;
        card.dataset.pending = 'false';
        syncDraftControls();
        void pumpQueue();
      }
    });
    discard.addEventListener('click', async () => {
      card.dataset.pending = 'true';
      syncDraftControls();
      try {
        const result = await request('draft.discard', { id: payload.id });
        if (!result?.ok) { throw new Error(result?.message ?? '无法丢弃草稿。'); }
      } catch (error) {
        addNotice(`丢弃草稿失败：${error.message}`, 'error');
      } finally {
        card.dataset.pending = 'false';
        syncDraftControls();
      }
    });
  }

  const storyLabels = {
    main_text: '正文', primary_header: '页眉', primary_footer: '页脚',
    footnotes: '脚注', endnotes: '尾注', comments: '批注',
  };
  const target = payload.target ?? {};
  const story = storyLabels[target.story] ?? target.story ?? '正文';
  card.dataset.style = payload.style;
  card.querySelector('.word-draft-document').textContent = payload.documentName || '当前文档';
  card.querySelector('.word-draft-style').value = payload.style;
  card.querySelector('.word-draft-target').textContent = payload.operation === 'replace'
    ? `确认后将替换 ${story} ${target.start ?? ''}-${target.end ?? ''} 的选中内容`
    : `确认后将在 ${story} ${target.start ?? ''} 处插入`;
  const selection = card.querySelector('.word-draft-selection');
  selection.hidden = payload.operation !== 'replace';
  selection.querySelector('.word-draft-selection-text').textContent = target.text || '（选区为空）';
  card.querySelector('.word-draft-actions .is-primary').textContent = payload.operation === 'replace'
    ? '确认替换选区' : '插入到文档';
  card.querySelector('.word-draft-preview').replaceChildren(...payload.blocks.map(createDraftBlock));
  if (isNew) { mountToTranscript(card); }
  syncDraftControls();
  scrollToBottom();
}

function closeDocumentDraft(payload) {
  const card = [...transcript.querySelectorAll('.word-draft-card')]
    .find((item) => item.dataset.draftId === payload?.id);
  if (!card) { return; }
  card.classList.add('is-closed');
  card.dataset.pending = 'false';
  card.querySelector('.word-draft-heading-title').textContent = payload.status === 'inserted' ? '已插入文档' : '草稿已丢弃';
  card.querySelector('.word-draft-controls').remove();
  card.querySelector('.word-draft-actions').remove();
  card.querySelector('.word-draft-adjust-form')?.remove();
}

function handleAgent(message) {
  switch (message.stage) {
    case 'text':
      appendAssistantText(message.text ?? '');
      break;
    case 'thinking':
      appendThinking(message.text ?? '');
      // 思考中还没有正文，气泡里改成对应说明，让用户知道不是卡住了。
      if (!currentAssistant) { showPending('正在思考…'); }
      break;
    case 'tool-start': {
      // 工具开始时结束当前助手气泡，让后续文本另起一段。
      currentAssistant = null;
      const payload = message.payload ?? {};
      addToolCard(payload);
      // 指示器重新落到末尾，并说明正在做什么。
      showPending(`正在${toolLabel(payload.name)}…`);
      break;
    }
    case 'tool-result':
      finishToolCard(message.payload ?? {});
      showPending();
      break;
    case 'document-draft':
      addDocumentDraft(message.payload ?? {});
      break;
    case 'draft-closed':
      closeDocumentDraft(message.payload ?? {});
      break;
    case 'approval-grants':
      renderApprovalGrants(message.payload?.grants ?? []);
      break;
    case 'retry':
      // 重试提示写进同一个气泡：这是「还在处理」的一种，不必单独占一条消息。
      showPending(message.text ?? '正在重试…');
      break;
    case 'usage':
      updateUsage(message.payload);
      break;
    case 'context':
      updateContextRing(message.payload, 'agent推送');
      break;
    case 'context-trimmed': {
      const p = message.payload ?? {};
      addNotice(
        `上下文超出预算，已压缩 ${p.compressed ?? 0} 条工具结果、移除 ${p.dropped ?? 0} 条早期记录（${p.before} → ${p.after} tokens）。`,
        'warn',
      );
      break;
    }
    case 'auto-continue':
      // 不落一条通知：这是加载项自己接上去的，用户什么都不用做，
      // 每次截断都插一条提示只会把对话记录塞满。写进处理指示器即可。
      showPending(message.text ?? '正在自动继续…');
      break;
    case 'tool-fallback':
      // 落一条通知而不是只写进指示器：工具形态变了会改变这个模型能做什么，
      // 用户需要在事后回看时还能看到这件事发生过。
      addNotice(message.text ?? '已改用其他方式调用工具。', 'warn');
      break;
    case 'vision-fallback':
      addNotice(message.text ?? '当前模型无法读取图片。', 'warn');
      break;
    case 'stalled':
      clearPending();
      addNotice(message.text ?? '模型输出反复被截断，已停止。', 'warn');
      break;
    case 'step-limit':
      clearPending();
      addNotice(message.text ?? '已达步数上限。', 'warn');
      break;
    case 'stopped':
      clearPending();
      addNotice(message.text ?? '已停止。', 'warn');
      break;
    case 'error':
      clearPending();
      addNotice(message.text ?? '发生错误。', 'error');
      break;
    case 'turn-complete':
      setStatus('');
      clearPending();
      markTurnComplete();
      // 本轮结束后上报一次布局：此时工具卡片已生成，
      // 才能核对它们是否默认折叠、有没有把正文挤出可视区。
      void logToHost(describeChatLayout());
      break;
    default:
      break;
  }
}

/** 输入区是否有可提交的内容。附件单独也算：贴张截图问「这个怎么填」很常见。 */
function hasComposerContent() {
  return composer.value.trim() !== '' || hasAttachments();
}

/**
 * 提交输入框里的内容。
 *
 * 空闲时立刻开跑，处理中则排到队尾，上一轮结束后自动接着跑。
 * 两条路合成一个入口：用户按 Enter 时不必先判断「现在能不能发」，
 * 界面也不必再靠禁用输入框来表达「等一下」。
 */
function submit(textOverride = null, draftId = null) {
  const fromComposer = textOverride === null;
  const text = fromComposer ? composer.value.trim() : String(textOverride ?? '').trim();
  if (!text && !(fromComposer && hasAttachments())) {
    return;
  }

  if (!isHosted()) {
    addNotice('未运行在 Office-helper 宿主内，无法发送。', 'error');
    return;
  }

  const entry = {
    id: `q${++queueSequence}`,
    text,
    draftId,
    // 附件在入队时就取出快照：输入框随即清空，之后加的附件属于下一条。
    images: fromComposer ? getImages() : [],
    files: fromComposer ? getFiles() : [],
  };

  if (fromComposer) {
    composer.value = '';
    clearAttachments();
    autoGrow();
  }

  queue.push(entry);
  renderQueueStrip();
  updateSendAffordance();

  void pumpQueue();
  return true;
}

/** 排队项在排队条上显示的一行字。没有正文时用附件充当标题。 */
function entrySummary(entry) {
  if (entry.text) {
    return entry.text;
  }

  const parts = [];
  if (entry.images.length > 0) { parts.push(`${entry.images.length} 张图片`); }
  if (entry.files.length > 0) { parts.push(`${entry.files.length} 个文件`); }
  return parts.join('、') || '（空）';
}

/**
 * 重画排队条。
 *
 * 整条重画而不是增量改：队列一动（入队、开跑、取消）位次就要全部重排，
 * 逐个节点去改反而要多维护一份「哪个节点对应哪条」的对应关系。
 * 队列最多也就几条，重画的代价可以忽略。
 */
function renderQueueStrip() {
  if (!queueStrip) {
    return;
  }

  queueStrip.replaceChildren();
  queueStrip.hidden = queue.length === 0;

  queue.forEach((entry, index) => {
    const chip = document.createElement('div');
    chip.className = 'queue-chip';
    chip.setAttribute('role', 'listitem');

    const position = document.createElement('span');
    position.className = 'queue-chip-pos';
    position.setAttribute('aria-hidden', 'true');
    position.textContent = String(index + 1);

    const text = document.createElement('span');
    text.className = 'queue-chip-text';
    text.textContent = entrySummary(entry);

    chip.append(position, text);

    const attachments = entry.images.length + entry.files.length;
    if (attachments > 0) {
      const mark = document.createElement('span');
      mark.className = 'queue-chip-files';
      const count = document.createElement('span');
      count.textContent = String(attachments);
      mark.append(createFileGlyph('file-glyph'), count);
      chip.append(mark);
    }

    const cancel = document.createElement('button');
    cancel.type = 'button';
    cancel.className = 'queue-chip-cancel';
    cancel.textContent = '×';
    cancel.title = '把这条从队列里去掉，不发送';
    cancel.setAttribute('aria-label', `取消排队中的第 ${index + 1} 条`);
    cancel.addEventListener('click', () => cancelQueued(entry.id));
    chip.append(cancel);

    // 位次与完整正文都在悬停说明里：条上只有数字与截断的一行，
    // 图标没有文字标签，title 是它唯一的自解释途径。
    const ordinal = index === 0 ? '下一个就发这条' : `排在第 ${index + 1} 位`;
    chip.title = `${ordinal}\n${entrySummary(entry)}` +
      (attachments > 0 ? `\n带 ${attachments} 个附件` : '');

    queueStrip.append(chip);
  });

  // 排队条最多显示三条（高度上限在 app.css），其余靠滑动看。每次重画后把视口
  // 归到队首那一端：replaceChildren 不会清掉滚动位置，用户若滑上去看早排的那几条，
  // 之后每次队列变化都会仍停在那里，而他要确认的永远是「下一个发哪条」。
  //
  // 归位给 0：column-reverse 的滚动范围是负的，0 是队首（第 1 位）那一端，
  // 滑到早排的那端是负值（实测 Chromium/WebView2 如此）。
  queueStrip.scrollTop = 0;
}

/**
 * 把条目的气泡插进对话流。真正开跑时才调用。
 *
 * 排队期间不进对话流：对话流记录的是已经发生的事，
 * 待办由排队条负责（见 renderQueueStrip）。
 */
function mountEntryBubble(entry) {
  const bubble = buildBubble('user', entry.text, entry.images, entry.files);
  mountToTranscript(bubble.wrapper);
  scrollToBottom();
}

/**
 * 取消一条排队中的输入：出队、重画，不留痕。
 *
 * 不往对话流里留划掉的气泡：对话流记的是已经发生的事，而这条从未发出。
 * 取消几次就积几条无法重发、也不进上下文的噪声，读起来只是干扰。
 * 代价是那段文字就此没了——按下取消时的意思本就是「这条不要了」。
 */
function cancelQueued(id) {
  const index = queue.findIndex((entry) => entry.id === id);
  // 找不到说明它已经开跑了，此刻要停只能用停止按钮。
  if (index < 0) {
    return;
  }

  queue.splice(index, 1);
  renderQueueStrip();
  updateSendAffordance();
}

/** 清空队列，返回被取消的条数。与单条取消一样不往对话流留痕。 */
function clearQueue() {
  const dropped = queue.splice(0, queue.length);
  renderQueueStrip();
  updateSendAffordance();
  return dropped.length;
}

/**
 * 依次跑完队列。
 *
 * 单实例轮转：pumping 作为闸门，多个入口（提交、上一轮结束）同时触发时
 * 也只有一条链在跑。否则两条链会各自 shift 出条目并发调用 chat.send，
 * 而加载项只接受一轮，第二条会撞上 BUSY 而白丢一次输入。
 */
async function pumpQueue() {
  if (pumping || draftOperationPending) {
    return;
  }

  pumping = true;
  try {
    while (queue.length > 0) {
      const entry = queue.shift();

      // 新一轮开始，上一批操作到此为止，收成一组。
      //
      // 收在这里而不是上一轮结束时：一轮刚跑完，用户往往正要看结果、点撤销。
      // 也必须在下面那句 mountEntryBubble 之前——组要落在上一轮内容的后面，
      // 而新的用户气泡一上屏，末尾就不再是上一轮了。
      sealOpsBatch();
      turnNumber += 1;
      opsBatchTurn = turnNumber;

      // 轮到它才上屏：从排队条挪进对话流，两处不会同时出现同一条。
      mountEntryBubble(entry);
      renderQueueStrip();
      await runTurn(entry);
    }
  } finally {
    pumping = false;
    updateSendAffordance();
  }
}

/** 跑一轮。气泡已由 pumpQueue 上屏，这里只负责请求与收尾。 */
async function runTurn(entry) {
  setBusy(true);
  // 新一轮绝不继承授权。授权本来只活在 Agent 的 RunAsync 里；
  // 面板也在这里先清芯片，避免桥的异步推送到达前显示上一轮的陈旧状态。
  renderApprovalGrants([]);
  // 进展显示在回复气泡里；状态行清空，免得上一轮的短暂提示看着像本轮的。
  setStatus('');
  showPending();

  try {
    const result = await request(
      'chat.send',
      { text: entry.text, images: entry.images, files: entry.files, draftId: entry.draftId },
      { timeout: 0 },
    );

    // 不再把 result.error 显示出来：加载项在返回这个字段之前，
    // 已经把同一条消息作为 error 推给了面板并渲染成胶囊，
    // 这里再显示一次就是同样的话说两遍。只留一条日志便于事后对账。
    if (result?.error) {
      void logToHost(`本轮以失败结束：${result.error}`, 'warn');
    }
  } catch (error) {
    // 走到这里是通道本身失败（超时、桥断开、加载项抛异常），
    // 加载项没有机会推送，必须由面板自己插一条。
    addNotice(error.message, 'error');
  } finally {
    setBusy(false);
  }
}

function autoGrow() {
  composer.style.height = 'auto';
  composer.style.height = `${Math.min(composer.scrollHeight, 160)}px`;
}

/**
 * 取消队列里所有待发内容，但不动正在跑的那一轮。
 *
 * 这是「停止」之前的一级：队列非空时按钮先干这件事。用户刚把内容排进队列
 * 又点了一下按钮，本意几乎不可能是掐掉正在跑的任务。
 *
 * 与单条取消一样不往对话流留痕（它们从未发出），只用一行胶囊回执说明取消了
 * 几条——同时也在告诉用户「当前任务还在跑」，下一下点击才是停止。
 */
function clearQueueByButton() {
  const dropped = clearQueue();
  if (dropped > 0) {
    addNotice(`已取消 ${dropped} 条排队中的输入，当前任务继续。再点一下可停止当前任务。`, 'warn');
  }
}

/**
 * 请求中断当前一轮。
 *
 * 停止不是瞬时的：正在进行的请求要收束。指示器改文案，
 * 让用户知道已经收到而不是没反应。
 *
 * 不再连带清空队列：队列非空时按钮的含义是「清空队列」，走不到这里
 * （见 updateSendAffordance）。因此走到这里队列本就是空的，
 * 「停完又自动开跑下一条」这件事不会发生。
 */
async function stopRun() {
  try {
    await request('chat.stop');
    showPending('正在停止…');
  } catch (error) {
    addNotice(`停止失败：${error.message}`, 'error');
  }
}

export function initChat() {
  transcript = document.getElementById('transcript');
  composer = document.getElementById('composer');
  sendButton = document.getElementById('send');
  statusLine = document.getElementById('status');
  usageLine = document.getElementById('usage');
  queueStrip = document.getElementById('queue-strip');

  // 显式落一次空闲态。index.html 里的 title 只是脚本就位前的兜底，
  // 真正的文案由 setBusy 统一给——两处各写一份的话，改了一处就会不一致。
  setBusy(false);
  // 设置读取可能需要等待 WorkBuddy 探测；先绘制默认的逐项审批图标，
  // 避免工具栏出现一个空白按钮或看起来发生错位。
  renderApprovalIcon();

  // 同一个按钮四种含义，见 updateSendAffordance。
  // 这里按点击时的实时状态分派，不看按钮上的类名——附件可能刚被粘进来。
  sendButton.addEventListener('click', () => {
    if (busy && !hasComposerContent()) {
      // 队列非空时先清队列，不动正在跑的那一轮：刚入队又点一下按钮，
      // 本意不可能是掐掉当前任务。停止要等队列空了那一下。
      if (queue.length > 0) {
        clearQueueByButton();
        return;
      }

      void stopRun();
      return;
    }

    submit();
  });

  composer.addEventListener('keydown', (event) => {
    // Enter 发送，Shift+Enter 换行——与主流对话界面一致。
    // 处理中按 Enter 同样受理，只是排到队尾而不是立刻发。
    if (event.key === 'Enter' && !event.shiftKey && !event.isComposing) {
      event.preventDefault();
      submit();
    }
  });

  composer.addEventListener('input', () => {
    autoGrow();
    // 输入框由空变非空会把按钮从「停止」换成「加入队列」，必须实时跟随。
    updateSendAffordance();
  });

  document.getElementById('approval-icon')?.addEventListener('click', () => void cycleApproval());
  document.getElementById('approval-grants')?.addEventListener('click', () => void revokeApprovalGrants());

  // 图片与文本文件附件：粘贴、拖入两种入口。
  // 附件变化也要刷新按钮含义：处理中只贴了张图、一个字没打，
  // 按钮同样该是「加入队列」，而粘贴不触发输入框的 input 事件。
  initAttachments(
    (message, variant) => addNotice(message, variant),
    () => updateSendAffordance(),
  );

  // 模型与思考等级由 picker 模块处理，切换后同步一次上下文占用，
  // 因为换模型可能改变协议，进而影响上下文预算的解释。
  initPicker(() => void refreshContext('切换模型'));

  document.getElementById('context-ring')?.addEventListener('click', async () => {
    if (!contextState) {
      await refreshContext('圆环点击');
      return;
    }

    try {
      const result = await request('context.compact');
      addNotice(
        result.trimmed
          ? `已压缩上下文：${result.before} → ${result.after} tokens`
          : '当前无可压缩内容。',
        result.trimmed ? 'info' : 'warn',
      );
      await refreshContext('圆环压缩后');
    } catch (error) {
      addNotice(`压缩失败：${error.message}`, 'error');
    }
  });

  initFit();
  exposeRibbonShortcuts();

  document.getElementById('reset').addEventListener('click', async () => {
    // 先清队列：留着它们就会在新会话里悄悄开跑，而用户已看不到任何痕迹。
    clearQueue();

    try {
      await request('chat.reset');
      transcript.replaceChildren();
      // 清空 DOM 不会清掉这个引用，漏掉会让下一轮往已移除的节点里写。
      pendingBubble = null;
      // 分组状态同样要清：留着的话新会话第一轮开始时会为上个会话那批已经
      // 不在 DOM 里的卡片建一个空组，而组里的卡片一张也点不到。
      opsBatch = [];
      opsGroups.length = 0;
      opsBatchTurn = 0;
      turnNumber = 0;
      mountSequence = 0;
      usageLine.textContent = '';
      compactPrompted = false;
      // 重新走一次就绪检查：既刷新上下文圆环，也让欢迎语重新出现。
      await checkReady('新会话');
      setStatus('已开始新会话。');
    } catch (error) {
      addNotice(`开始新会话失败：${error.message}`, 'error');
    }
  });

  on('agent', handleAgent);
  on('approval-request', addApprovalCard);

  // 就绪检查交由路由切换统一触发，此处不再调用，避免启动时重复请求。
}

/**
 * 重新检查配置就绪状态。
 *
 * 必须可重复调用：用户在设置页改完再回到对话页时，
 * 旧的判断结果已经过期。只在初始化时查一次会一直显示陈旧的错误。
 */
export async function refreshReady(source = '路由切换') {
  await checkReady(source);
}

/**
 * 处理方式的图标定义。
 *
 * 用盾牌的填充程度表达防护强度：
 *   逐项审批 = 实心盾（最严），每轮确认 = 半填充，全自动 = 空心带感叹号。
 * 颜色同步变化，绿→琥珀→红，与风险递增一致。
 */
const APPROVAL_ICONS = {
  PerWrite: {
    // 实心盾
    path: 'M8 1.5 3 3.3v4.2c0 3 2.1 5.6 5 6.5 2.9-.9 5-3.5 5-6.5V3.3L8 1.5z',
    variant: 'strict',
    label: '逐项审批',
  },
  PerTurn: {
    // 盾牌轮廓 + 下半填充，表示只在开头确认一次
    path: 'M8 1.5 3 3.3v4.2c0 3 2.1 5.6 5 6.5 2.9-.9 5-3.5 5-6.5V3.3L8 1.5zm0 1.6 3.6 1.3v3.1c0 2.2-1.5 4.2-3.6 5V3.1z',
    variant: 'medium',
    label: '每轮确认',
  },
  Automatic: {
    // 空心盾 + 感叹号，提示无人工把关
    path: 'M8 1.5 3 3.3v4.2c0 3 2.1 5.6 5 6.5 2.9-.9 5-3.5 5-6.5V3.3L8 1.5zm0 1.6 3.6 1.3v3.1c0 2.2-1.5 4.2-3.6 5-2.1-.8-3.6-2.8-3.6-5V4.4L8 3.1zM7.3 5h1.4v3.6H7.3V5zm0 4.4h1.4v1.4H7.3V9.4z',
    variant: 'auto',
    label: '全自动',
  },
};

const APPROVAL_ORDER = ['PerWrite', 'PerTurn', 'Automatic'];

let approvalOptions = [];
let currentApproval = 'PerWrite';

const APPROVAL_CLASS_LABELS = {
  Format: '格式',
  Write: '写入',
  // 抹除类单独成档：清除、合并、排序会抹掉或搬动既有内容，
  // 而合并还会静默丢值。芯片上必须与「写入」区分得开，
  // 否则用户以为自己只放行了填数据。
  Destructive: '抹除',
  Structure: '结构',
};

/**
 * 当前轮活着的授权。盾牌的填充程度只说明所选策略，
 * 不足以说清「哪个表的哪类操作已不再询问」，因此另用文字芯片。
 *
 * 芯片同时是撤回入口：点一下就收回本轮全部授权，后续操作重新逐个询问。
 * 只显示不给撤回是不够的——用户中途改主意时，唯一出路会变成掐掉整轮，
 * 而他想停的只是「别再自动放行」这一件事。
 */
function renderApprovalGrants(grants) {
  const chip = document.getElementById('approval-grants');
  if (!chip) {
    return;
  }

  if (!Array.isArray(grants) || grants.length === 0) {
    chip.hidden = true;
    chip.textContent = '';
    chip.title = '';
    return;
  }

  const labels = grants.map((grant) => {
    const kind = APPROVAL_CLASS_LABELS[grant.approvalClass] ?? grant.approvalClass;
    // 新建/重命名工作表没有范围参数，宿主拿不到表名，改按工作簿记。
    // 那个键是个内部标记，不能原样显示给用户。
    const where = grant.workbookWide ? '整个工作簿' : grant.sheet;
    return `${where} · ${kind}`;
  });
  chip.textContent = labels.join('、');
  chip.title = `本轮已允许：${labels.join('、')}。点这里收回，后续操作会重新逐个询问。新一轮也会重新确认。`;
  chip.hidden = false;
}

/** 收回本轮全部授权。之后的写操作重新逐个弹卡。 */
async function revokeApprovalGrants() {
  const chip = document.getElementById('approval-grants');
  if (!chip || chip.hidden) {
    return;
  }

  try {
    const result = await request('approval.revoke', {});
    if (result?.ok === false) {
      addNotice(result.message ?? '收回授权失败。', 'error');
      return;
    }

    renderApprovalGrants([]);
    addNotice('已收回本轮授权，后续操作会重新逐个询问。', 'info');
  } catch (error) {
    addNotice(`收回授权失败：${error.message}`, 'error');
  }
}

/**
 * 填充输入区控件：模型与思考等级由 picker 模块统一负责，
 * 这里只处理处理方式图标。
 */
function fillQuickControls(settings) {
  approvalOptions = settings.approvalOptions ?? approvalOptions;
  currentApproval = settings.approval ?? currentApproval;
  renderApprovalIcon();
  syncPicker(settings);
}

/** 按当前策略绘制图标与颜色。 */
function renderApprovalIcon() {
  const button = document.getElementById('approval-icon');
  const glyph = document.getElementById('approval-glyph');
  if (!button || !glyph) {
    return;
  }

  const icon = APPROVAL_ICONS[currentApproval] ?? APPROVAL_ICONS.PerWrite;
  glyph.setAttribute('d', icon.path);

  for (const name of ['strict', 'medium', 'auto']) {
    button.classList.toggle(`is-${name}`, icon.variant === name);
  }

  const option = approvalOptions.find((o) => o.id === currentApproval);
  const hint = option?.hint ?? '';
  const label = option?.label ?? icon.label;
  button.title = `处理方式：${label}${hint ? `\n${hint}` : ''}\n点击切换`;
  button.setAttribute('aria-label', `处理方式：${label}，点击切换`);
}

/**
 * 切换处理方式的回执。
 *
 * 原先写在输入区那条状态行上，而它与排队条争位——320px 栏宽下会被压到只剩
 * 两三个字，正文读不全只能靠悬停。而「现在改成哪一档」恰恰是必须一眼看清的：
 * 三档的差别是模型能不能不问就动手。
 *
 * 改成与「已完成」同一种居中胶囊：那类系统级消息本来就不属于对话任何一方，
 * 位置在对话流中间、宽度随文字伸缩，不与任何控件抢空间。
 *
 * 只保留一条。切换是「点着轮换三档」的交互，每次追加会在连点三次后
 * 留下三条自相矛盾的记录，而只有最后那条是真的。
 */
let approvalNotice = null;

function announceApproval(option, id) {
  const label = option?.label ?? id;
  const hint = option?.hint ?? '';

  // 前一条就地移除，不做「改写文字」：胶囊要重新量行数决定圆角，
  // 而且移除再插能让新的一条重放进场动画，切换才有可见的回执。
  if (approvalNotice?.parent) {
    approvalNotice.remove();
  }

  const notice = addNotice(hint ? `处理方式：${label} · ${hint}` : `处理方式：${label}`, 'info');

  // 三档用与盾牌一致的颜色：绿 → 琥珀 → 红，与风险递增同向。
  // 刻意不复用 notice-error——那是「出错了」的语义，而全自动是用户主动选的。
  notice.classList.add('notice-approval', `is-${APPROVAL_ICONS[id]?.variant ?? 'strict'}`);
  notice.title = hint ? `${label}：${hint}` : label;
  approvalNotice = notice;
  return notice;
}

/** 点击循环切换处理方式并立即保存。 */
async function cycleApproval() {
  const index = APPROVAL_ORDER.indexOf(currentApproval);
  currentApproval = APPROVAL_ORDER[(index + 1) % APPROVAL_ORDER.length];
  renderApprovalIcon();

  try {
    await request('session.update', { approval: currentApproval });
    announceApproval(approvalOptions.find((o) => o.id === currentApproval), currentApproval);
  } catch (error) {
    addNotice(`调整处理方式失败：${error.message}`, 'error');
  }
}

/**
 * 对话区布局摘要，在一轮结束后上报。
 *
 * 与 app.js 的整页布局自检互补：这里关注对话内容本身
 * （工具卡片是否折叠、消息是否按 4/5 宽度分列），
 * 那些元素只有在对话产生内容后才存在。
 */
function describeChatLayout() {
  const cards = transcript.querySelectorAll('.tool-card');
  const opened = transcript.querySelectorAll('.tool-card[open]');

  const widthOf = (selector) => {
    const node = transcript.querySelector(`${selector} .msg-body`);
    if (!node) { return '无'; }
    const own = node.getBoundingClientRect().width;
    const parent = transcript.getBoundingClientRect().width;
    return parent > 0 ? `${Math.round((own / parent) * 100)}%` : '?';
  };

  return `对话布局：工具卡片 ${cards.length} 个（展开 ${opened.length}）` +
    // 轮次组：本轮的卡片还平铺着，往前几轮的应当都已成组。
    // 组数与「批中待收」对不上就说明有卡片漏了收集或收了两次。
    ` 轮次组 ${transcript.querySelectorAll('.ops-group').length} 个` +
    `（展开 ${transcript.querySelectorAll('.ops-group[open]').length}，` +
    `批中待收 ${opsBatch.length} 个）` +
    ` 助手消息宽 ${widthOf('.msg-assistant')} 用户消息宽 ${widthOf('.msg-user')}` +
    ` 欢迎语 ${transcript.querySelectorAll('.welcome').length} 个` +
    // 本轮已结束，指示器必须已清除；留下就说明有路径漏了清理。
    ` 处理指示器 ${transcript.querySelectorAll('.msg-pending').length} 个` +
    // 完成标记：正常收完的轮数。这行日志是在 turn-complete 里发的，
    // 因此至少应有 1 个；数目应与正常收完的轮数一致。
    ` 完成标记 ${transcript.querySelectorAll('.notice-complete').length} 个` +
    // 队列状态一并记录：排队条上的条目与内部队列必须数目一致，
    // 对不上就说明有条目丢了显示或显示丢了条目。
    ` 队列 ${queue.length} 条（排队条 ${queueStrip?.querySelectorAll('.queue-chip').length ?? 0} 个）`;
}

/**
 * 显示欢迎语，说明身份、能力与边界。
 *
 * 只在对话流为空时显示，避免每次切回对话页都插一遍。
 * 配置有问题时追加一段说明，让用户知道该去哪里处理，
 * 而不是发了消息才发现不能用。
 */
function showWelcome(settings) {
  // 已有内容（含上一次的欢迎语）就不再插入。
  if (transcript.querySelector('.welcome') || transcript.children.length > 0) {
    return;
  }

  const card = document.createElement('div');
  card.className = 'welcome';

  const title = document.createElement('div');
  title.className = 'welcome-title';
  const copy = welcomeCopy(isWordHostUi());
  title.textContent = copy.title;

  const body = document.createElement('div');
  body.className = 'welcome-body';
  body.innerHTML = renderMarkdown(copy.body);

  card.append(title, body);

  if (!settings.ready && !isWorkBuddyAuthorizationDetail(settings.readyDetail)) {
    const warn = document.createElement('div');
    warn.className = 'welcome-warn';
    const detail = (settings.readyDetail ?? '').replace(/[。.]+$/, '');
    warn.textContent = `暂时还不能开始：${detail || '配置未完成'}。请到上方「设置」页完成配置。`;
    card.append(warn);
  }

  mountToTranscript(card);
}

/**
 * 宿主信息异步返回时，欢迎语可能已经先画出来；此处补一次宿主专属文案。
 * 只在 Word/WPS 下调用，Excel 欢迎语保持原样。
 */
export function refreshWelcomeHostUi() {
  if (!isWordHostUi() || !transcript) { return; }

  const card = transcript.querySelector('.welcome');
  if (!card) { return; }

  const copy = welcomeCopy(true);
  const title = card.querySelector('.welcome-title');
  const body = card.querySelector('.welcome-body');
  if (title) { title.textContent = copy.title; }
  if (body) { body.innerHTML = renderMarkdown(copy.body); }
}

function isWorkBuddyAuthorizationDetail(detail) {
  return /workbuddy|授权状态|授权登录/i.test(String(detail ?? ''));
}

/**
 * 进入对话页时检查配置是否可用。
 * 否则用户要先发一条消息才知道没配好，反馈太晚。
 */
async function checkReady(source = '未标注') {
  if (!isHosted()) {
    return;
  }

  try {
    // 就绪判断由后端给出：只有它知道 CLI 配置里有没有模型、密钥能否解开。
    const settings = await request('settings.get');
    updateChannelHeader(settings);
    fillQuickControls(settings);
    await refreshContext(source);

    // 移除上一次的未就绪提示，否则配置修好后旧提示仍留在对话流里，
    // 看起来像是「配置完成了但还在报错」。
    for (const stale of transcript.querySelectorAll('.notice-config')) {
      stale.remove();
    }

    showWelcome(settings);

    if (settings.ready) {
      // 就绪时不显示任何文字：模型与接入信息已在输入区的选择器上可见，
      // 再占一行重复展示 url 只会挤压对话空间。
      setStatus('');
      return;
    }

    // 去掉尾部句号再拼接：后端的错误消息自带句号，直接拼会出现「。。」。
    const detail = isWorkBuddyAuthorizationDetail(settings.readyDetail)
      ? '当前连接需要在「设置」页完成授权'
      : (settings.readyDetail ?? '').replace(/[。.]+$/, '') || '配置未完成';
    const notice = addNotice(`还不能开始对话：${detail}。请到「设置」页完成配置。`, 'warn');
    // 打标记以便下次检查时清理。
    notice.classList.add('notice-config');
    // 状态行不再重复一遍「配置未完成」：上面那条胶囊已经把原因和去处说清了。
    setStatus('');
  } catch (error) {
    addNotice(`读取配置失败：${error.message}`, 'error');
  }
}
