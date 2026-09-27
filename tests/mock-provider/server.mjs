// OpenAI 兼容的最小 mock 服务，用于端到端验证而不消耗真实额度。
//
// 覆盖三条关键路径：
//   1. 模型列表   GET  /v1/models
//   2. 纯文本流式 POST /v1/chat/completions
//   3. 工具调用流式（分帧下发参数，验证增量拼接与工具执行）
//
// 用法：node server.mjs [端口] [场景]
//   场景 text   只回文本
//   场景 tool   先调用一个工具，收到结果后再回文本（默认）
//   场景 bulk   连续请求多次读取，快速堆高上下文，
//               用于验证 90% 阈值触发压缩的路径
//   场景 flaky  前两次以 503 拒绝再放行，验证失败重试确实会重来并最终成功
//   场景 reject 一律以 401 拒绝，验证配置类错误不被重试
//   场景 cut    第一轮模拟输出被长度上限截断（无正文、无工具调用，且不回
//               finish_reason，与实测网关一致），第二轮才正常干活。
//               用于验证加载项会自动续跑，而不是把半途当成结束
//   场景 cutloop 每一轮都被截断，永不产出。验证续跑有上限、不会无限空转（只应收到一次请求）
//   场景 slow   慢慢回一段话，并把收到的最后一句用户输入原样念回来。
//              慢是为了让一轮长时间停在处理中，好在这期间验证排队；
//              念回输入是为了能断言队列到底按什么顺序发出去的。
//   场景 notool 带 tools 的请求一律以 400 拒绝，不带则正常干活，并用文本指令块
//              发起调用。用于验证「不支持原生工具调用」会自动改用文本协议，
//              且解析出的调用照样走审批与执行。
//   场景 toolrefusal 接受原生 tools 但回复没有工具调用能力，随后用文本指令块干活。
//   场景 deepseek-flash 首个请求直接走文本指令块，验证型号兼容提示不会发送 tools。
//   场景 novision 带图片的请求以 400 拒绝，不带则正常回文本。
//              用于验证视觉回退：去图重发或经中转模型转写后继续。

import { createServer } from 'node:http';

const port = Number(process.argv[2]) || 58940;
const scenario = process.argv[3] || 'tool';

/** flaky 场景先失败几次。取 2 是为了既覆盖多次重试，又不必等满退避总时长。 */
const FLAKY_FAILURES = 2;

/** grant 场景走到第几步。每步下发一个工具调用，四步后收尾。 */
let grantStep = 0;

/**
 * novision 场景里「看得了图」的那个模型名。
 * 验证脚本用 -VisionRelayModel 传同一个值，两处必须一致。
 */
const VISION_RELAY_MODEL = 'mock-vision';

/**
 * slow 场景每帧的间隔毫秒数。
 *
 * 一轮总时长要够长，让脚本有从容的时间在处理中再投两条输入并读回队列状态；
 * 又不能长到把整个验证拖成分钟级。8 帧 × 700ms 约 5.6 秒，两者兼顾。
 */
const SLOW_FRAME_MS = 700;

/** 对话请求计数。重试验证要靠它判断到底来了几次。 */
let chatAttempts = 0;

function sse(res, payload) {
  res.write(`data: ${JSON.stringify(payload)}\n\n`);
}

function textFrame(content) {
  return { choices: [{ delta: { content }, index: 0 }] };
}

function finish(reason) {
  return { choices: [{ delta: {}, index: 0, finish_reason: reason }] };
}

function usage(prompt, completion) {
  return { choices: [], usage: { prompt_tokens: prompt, completion_tokens: completion } };
}

const server = createServer((req, res) => {
  const url = new URL(req.url, `http://127.0.0.1:${port}`);
  // 日志一律英文：本进程输出被 PowerShell 重定向读取，
  // Node 写 UTF-8 而 PowerShell 按 ANSI 读，中文会成乱码。
  console.log(`[mock] ${req.method} ${url.pathname}`);

  if (url.pathname === '/v1/models') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    // 混合名单：前两个正常，其余各自对应对话处理器里按模型名分流的一种情形。
    // 名单本身不会让探测产出差异，差异全在 /v1/chat/completions 那边。
    res.end(JSON.stringify({
      data: [
        { id: 'mock-model' },
        { id: 'mock-model-mini' },
        { id: 'mock-absent' },
        { id: 'mock-forbidden' },
        { id: 'mock-keybad' },
        { id: 'mock-ratelimit' },
        { id: 'mock-aliasbroken' },
        { id: 'mock-silent' },
      ],
    }));
    return;
  }

  if (url.pathname !== '/v1/chat/completions') {
    res.writeHead(404).end('not found');
    return;
  }

  let body = '';
  req.on('data', (chunk) => { body += chunk; });
  req.on('end', () => {
    let parsed = {};
    try { parsed = JSON.parse(body); } catch {}

    const messages = parsed.messages ?? [];
    // 已经收到过工具结果，说明这是第二轮，该收尾了。
    const hasToolResult = messages.some((m) => m.role === 'tool');

    // 统计收到的图片：验证多模态载荷是否真的送达。
    let imageCount = 0;
    let imageTypes = [];
    for (const message of messages) {
      if (!Array.isArray(message.content)) { continue; }
      for (const block of message.content) {
        if (block?.type === 'image_url' && block.image_url?.url?.startsWith('data:')) {
          imageCount += 1;
          const match = /^data:([^;]+);base64,/.exec(block.image_url.url);
          if (match) { imageTypes.push(match[1]); }
        }
      }
    }

    console.log(
      `[mock] messages=${messages.length} hasToolResult=${hasToolResult} ` +
        `tools=${(parsed.tools ?? []).length} images=${imageCount}` +
        (imageTypes.length > 0 ? ` types=${imageTypes.join(',')}` : ''),
    );

    chatAttempts += 1;

    // 按模型名分流，供「按需确认」的端到端验证使用。
    //
    // 这几条与 scenario 无关：可用性判定是按模型给出的，因此 mock 也必须能
    // 对不同模型给不同答案，否则「确认出可用/不可用/未知」三种结论无法在真实
    // 宿主里区分——扩 /v1/models 的名单不会在这条链上产生任何差异。
    const model = String(parsed.model ?? '');

    // 不存在的模型：点名它。加载项应判为不可用。
    if (model.includes('absent')) {
      console.log(`[mock] ${model} -> 404 model_not_found`);
      res.writeHead(404, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        error: { code: 'model_not_found', message: `The model '${model}' does not exist` },
      }));
      return;
    }

    // 无权访问某一个模型：点名模型的 403，同样判不可用。
    if (model.includes('forbidden')) {
      console.log(`[mock] ${model} -> 403 (names the model)`);
      res.writeHead(403, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        error: { message: `Your key has no access to model ${model}` },
      }));
      return;
    }

    // 只说密钥的 403：说的是账号，判未知。
    if (model.includes('keybad')) {
      console.log(`[mock] ${model} -> 403 (account only)`);
      res.writeHead(403, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: { message: 'invalid api key' } }));
      return;
    }

    // 限流：判未知，绝不能判不可用。
    if (model.includes('ratelimit')) {
      console.log(`[mock] ${model} -> 429`);
      res.writeHead(429, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: { message: 'rate limit exceeded' } }));
      return;
    }

    // 别名：200 开流，但体内就是一条错误。判不可用。
    // Chat Completions 的解析器此前没有 error 分支，这条正是为它补的。
    if (model.includes('aliasbroken')) {
      console.log(`[mock] ${model} -> 200 with in-body error`);
      res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        Connection: 'keep-alive',
      });
      res.write(`data: ${JSON.stringify({
        error: { code: 'model_not_found', message: `The model '${model}' does not exist` },
      })}\n\n`);
      res.end();
      return;
    }

    // 200 但一个事件都不给：判未知，不判可用。
    if (model.includes('silent')) {
      console.log(`[mock] ${model} -> 200 with no events`);
      res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        Connection: 'keep-alive',
      });
      res.end();
      return;
    }

    // 只接受 max_completion_tokens 的推理模型。
    if (model.includes('reasoning') && parsed.max_tokens !== undefined) {
      console.log(`[mock] ${model} -> 400 (needs max_completion_tokens)`);
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        error: {
          message: "Unsupported parameter: 'max_tokens' is not supported with this model. " +
            "Use 'max_completion_tokens' instead.",
        },
      }));
      return;
    }

    // 只认 low/medium/high 的网关：拒绝 "none"，但不带思考参数就放行。
    // 探测若把「关闭」当成「不传」，这条会让它恒判未知。
    if (model.includes('strictthinking') && parsed.reasoning_effort !== undefined) {
      console.log(`[mock] ${model} -> 400 (reasoning_effort=${parsed.reasoning_effort} rejected)`);
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        error: { message: `Unsupported value for reasoning_effort: ${parsed.reasoning_effort}` },
      }));
      return;
    }

    // 一律拒绝：401 属于配置错误，重试多少次都一样，
    // 因此加载项应当只请求一次就把错误报出来。
    if (scenario === 'reject') {
      console.log(`[mock] attempt ${chatAttempts} -> 401 (must not be retried)`);
      res.writeHead(401, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: { message: 'invalid api key (mock)' } }));
      return;
    }

    // 只要请求里带 tools 就拒绝，措辞照抄真实网关。
    // 加载项应当据此改用文本指令协议，而不是把整轮报失败。
    if (scenario === 'notool' && (parsed.tools ?? []).length > 0) {
      console.log(`[mock] attempt ${chatAttempts} -> 400 (tools not supported)`);
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        error: { message: 'This model does not support tools / function calling.' },
      }));
      return;
    }

    // 视觉中转模型是这条规则的例外：它就是那个「看得了图」的模型。
    // 认模型名而不是别的，正因为加载项转述时只换模型名、连接照旧。
    if (scenario === 'novision' && imageCount > 0 && parsed.model === VISION_RELAY_MODEL) {
      console.log(`[mock] relay ${VISION_RELAY_MODEL} describing ${imageCount} image(s)`);
      res.writeHead(200, {
        'Content-Type': 'text/event-stream; charset=utf-8',
        'Cache-Control': 'no-cache',
        Connection: 'keep-alive',
      });

      // 转写内容刻意写成表格截图该有的样子：加载项会把它当文本注入上下文，
      // 主模型据此作答，因此这里要能与「没收到说明」明确区分开。
      const described = '这是一张表格截图。可见范围 A1:B3，表头为「名称」「数量」，' +
        '第 2 行为「铅笔 10」，第 3 行为「橡皮 5」。无合并单元格，无报错提示。';
      for (const piece of described.match(/.{1,10}/gs) ?? []) {
        sse(res, textFrame(piece));
      }

      sse(res, finish('stop'));
      sse(res, usage(220, 60));
      res.write('data: [DONE]\n\n');
      res.end();
      return;
    }

    // 带图片就拒绝。加载项应当去图重发，或先经视觉中转模型转写。
    if (scenario === 'novision' && imageCount > 0) {
      console.log(`[mock] attempt ${chatAttempts} -> 400 (no vision, images=${imageCount})`);
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        error: { message: 'Invalid content type image_url: this model has no vision capability.' },
      }));
      return;
    }

    // 先失败几次再放行：验证重试真的会重来，且带 Retry-After 时按它等待。
    if (scenario === 'flaky' && chatAttempts <= FLAKY_FAILURES) {
      console.log(`[mock] attempt ${chatAttempts} -> 503 (expect retry)`);
      res.writeHead(503, { 'Content-Type': 'application/json', 'Retry-After': '1' });
      res.end(JSON.stringify({ error: { message: 'upstream temporarily unavailable (mock)' } }));
      return;
    }

    if (scenario === 'flaky') {
      console.log(`[mock] attempt ${chatAttempts} -> 200 (recovered)`);
    }

    res.writeHead(200, {
      'Content-Type': 'text/event-stream; charset=utf-8',
      'Cache-Control': 'no-cache',
      Connection: 'keep-alive',
    });

    // bulk 场景：每轮都请求读取一大片范围，让工具结果不断累积。
    // 读取上限是 5000 个单元格，单次结果已相当可观，数轮即可逼近上下文预算。
    const toolResultCount = messages.filter((m) => m.role === 'tool').length;

    // 取最后一条用户输入的纯文本。slow 场景要把它念回来，
    // 好让验证脚本据此判断队列的实际发送顺序。
    const lastUserText = (() => {
      for (let i = messages.length - 1; i >= 0; i--) {
        const message = messages[i];
        if (message.role !== 'user') { continue; }
        if (typeof message.content === 'string') { return message.content; }
        if (Array.isArray(message.content)) {
          const text = message.content.find((b) => b?.type === 'text');
          if (text) { return text.text ?? ''; }
        }
      }

      return '';
    })();

    const send = async () => {
      // slow 场景：慢慢回一句话，并原样念回收到的输入。
      // 不调用任何工具：这里要验证的是面板的排队与顺序，
      // 掺进工具执行只会让一轮的时长变得不好预期。
      if (scenario === 'slow') {
        const reply = `收到：${lastUserText}`;
        // 先按帧数切分再逐帧下发，确保总时长与帧数无关于文本长度。
        const size = Math.max(1, Math.ceil(reply.length / 8));
        for (const piece of reply.match(new RegExp(`.{1,${size}}`, 'gs')) ?? []) {
          await new Promise((r) => setTimeout(r, SLOW_FRAME_MS));
          sse(res, textFrame(piece));
        }

        sse(res, finish('stop'));
        sse(res, usage(100, 20));
        res.write('data: [DONE]\n\n');
        res.end();
        return;
      }

      if (scenario === 'toolrefusal' && (parsed.tools ?? []).length > 0) {
        console.log('[mock] toolrefusal: model says it cannot call tools');
        const refusal = '我没有工具调用的能力。';
        for (const piece of refusal.match(/.{1,6}/gs) ?? []) {
          sse(res, textFrame(piece));
        }

        sse(res, finish('stop'));
        sse(res, usage(240, 30));
        res.write('data: [DONE]\n\n');
        res.end();
        return;
      }

      // notool、toolrefusal 与 deepseek-flash 场景：用文本指令块发起调用。
      //
      // 文本协议下工具结果是以 user 消息回灌的（协议里没有 tool 角色可用），
      // 因此不能靠 hasToolResult 判断轮次，要认那条消息里的标记。
      if (scenario === 'notool' || scenario === 'toolrefusal' || scenario === 'deepseek-flash') {
        if (scenario === 'toolrefusal') {
          const retainedRefusal = messages.some((m) =>
            m.role === 'assistant' &&
            typeof m.content === 'string' &&
            m.content.includes('我没有工具调用的能力'));
          console.log(retainedRefusal
            ? '[mock] toolrefusal: retry retained the refusal'
            : '[mock] toolrefusal: retry omitted the refusal');
          console.log((parsed.tools ?? []).length === 0
            ? '[mock] toolrefusal: retry omitted native tools'
            : '[mock] toolrefusal: retry still has native tools');
        }

        if (scenario === 'deepseek-flash') {
          console.log((parsed.tools ?? []).length === 0
            ? '[mock] deepseek-flash: text protocol request (tools=0)'
            : '[mock] deepseek-flash: unexpected native tools');
        }

        // 认工具名而不是那句中文抬头：抬头的措辞随时可能改，
        // 而工具名是协议里的固定标识。用户的原始提问里不会出现它。
        const fed = messages.some((m) =>
          m.role === 'user' &&
          typeof m.content === 'string' &&
          m.content.includes('write_values'));

        if (!fed) {
          const block =
            '我先写入表头。\n\n```chatsheet:tool\n' +
            '{"tool": "write_values", "args": {"range": "A1:B1", "values": [["名称", "数量"]]}}\n' +
            '```\n';

          for (const piece of block.match(/.{1,8}/gs) ?? []) {
            await new Promise((r) => setTimeout(r, 30));
            sse(res, textFrame(piece));
          }

          sse(res, finish('stop'));
          sse(res, usage(300, 50));
          res.write('data: [DONE]\n\n');
          res.end();
          return;
        }

        for (const piece of '已写入 A1:B1。'.match(/.{1,4}/gs) ?? []) {
          await new Promise((r) => setTimeout(r, 30));
          sse(res, textFrame(piece));
        }

        sse(res, finish('stop'));
        sse(res, usage(400, 30));
        res.write('data: [DONE]\n\n');
        res.end();
        return;
      }

      // novision 场景：到这里说明请求已经不带图片了（带图的在上面被拒）。
      // 把是否收到过「有图但你看不到」的说明念回来，便于断言回退路径确实
      // 在上下文里留了痕迹——静默丢图是这条链路最危险的失败方式。
      if (scenario === 'novision') {
        // 认「系统提示」这个固定前缀：加载项注入的两种说法（转写、未送达）
        // 都以它开头，而用户自己的提问不会。
        const told = messages.some((m) =>
          typeof m.content === 'string' && m.content.includes('系统提示'));
        const reply = told
          ? '我收到了关于图片的说明，但看不到图片本身。'
          : '我没有收到任何图片，也没有相关说明。';

        for (const piece of reply.match(/.{1,6}/gs) ?? []) {
          await new Promise((r) => setTimeout(r, 30));
          sse(res, textFrame(piece));
        }

        sse(res, finish('stop'));
        sse(res, usage(180, 30));
        res.write('data: [DONE]\n\n');
        res.end();
        return;
      }

      // image 场景：只回文本并报出收到的图片数，用于验证多模态链路。
      if (scenario === 'image') {
        const reply = imageCount > 0
          ? `我收到了 ${imageCount} 张图片（${imageTypes.join('、')}）。`
          : '我没有收到任何图片。';

        for (const piece of reply.match(/.{1,6}/gs) ?? []) {
          await new Promise((r) => setTimeout(r, 30));
          sse(res, textFrame(piece));
        }

        sse(res, finish('stop'));
        sse(res, usage(150, 30));
        res.write('data: [DONE]\n\n');
        res.end();
        return;
      }

      if (scenario === 'bulk') {
        // 固定轮数后收尾，避免无限循环。
        if (toolResultCount < 12) {
          sse(res, textFrame(`第 ${toolResultCount + 1} 轮读取。`));
          await new Promise((r) => setTimeout(r, 60));

          // 每轮读不同范围，避免被上层去重。
          const startRow = toolResultCount * 40 + 1;
          const args = JSON.stringify({ range: `A${startRow}:T${startRow + 39}` });

          sse(res, {
            choices: [{
              index: 0,
              delta: {
                tool_calls: [{
                  index: 0,
                  id: `call_bulk_${toolResultCount}`,
                  type: 'function',
                  function: { name: 'read_range', arguments: args },
                }],
              },
            }],
          });

          sse(res, finish('tool_calls'));
          sse(res, usage(500 * (toolResultCount + 1), 60));
          res.write('data: [DONE]\n\n');
          res.end();
          return;
        }

        sse(res, textFrame('读取完毕。'));
        sse(res, finish('stop'));
        sse(res, usage(9000, 80));
        res.write('data: [DONE]\n\n');
        res.end();
        return;
      }

      // cutloop 场景：每一轮都贴顶截断，永不产出。
      // 加载项必须在有限次续跑后停下并说明原因，否则会一直空转到步数上限，
      // 白烧额度还看不出发生了什么。
      if (scenario === 'cutloop') {
        const cap = parsed.max_tokens ?? 8192;
        console.log(`[mock] cutloop: stalled turn ${chatAttempts}, usage ${cap}/${cap}`);
        sse(res, usage(500, cap));
        res.write('data: [DONE]\n\n');
        res.end();
        return;
      }

      // cut 场景：第一轮什么都不产出，只把用量顶到上限。
      //
      // 刻意不回 finish_reason：实测的中转网关就是这样（日志里结束原因为
      // 「未提供」），因此加载项只能靠「输出用量贴住上限」判断被截断。
      // 若这里回了 length，就绕过了那条更难的判据。
      if (scenario === 'cut') {
        const cap = parsed.max_tokens ?? 8192;
        // 按会话形态判断轮次，不按文本匹配：系统提示里本来就写着「截断」
        // 这类字样，拿关键词认轮次会在第一轮就误判。
        // 助手消息的有无才是可靠信号——第一轮一条都没有。
        const hasAssistantTurn = messages.some((m) => m.role === 'assistant');

        // 第一轮：无正文、无工具调用、用量贴顶。加载项应当自动续跑。
        if (!hasToolResult && !hasAssistantTurn) {
          console.log(`[mock] cut: stalled turn, usage ${cap}/${cap}, no finish_reason`);
          sse(res, usage(500, cap));
          res.write('data: [DONE]\n\n');
          res.end();
          return;
        }

        // 续跑轮：收到催促后正常调用工具。
        if (!hasToolResult) {
          console.log('[mock] cut: continued, issuing tool call');
          sse(res, {
            choices: [{
              index: 0,
              delta: {
                tool_calls: [{
                  index: 0,
                  id: 'call_cut_1',
                  type: 'function',
                  function: {
                    name: 'write_values',
                    arguments: '{"range":"A1:B1","values":[["名称","数量"]]}',
                  },
                }],
              },
            }],
          });

          sse(res, finish('tool_calls'));
          sse(res, usage(600, 40));
          res.write('data: [DONE]\n\n');
          res.end();
          return;
        }

        // 收尾轮。
        for (const piece of '已写入 A1:B1。'.match(/.{1,4}/gs) ?? []) {
          await new Promise((r) => setTimeout(r, 40));
          sse(res, textFrame(piece));
        }

        sse(res, finish('stop'));
        sse(res, usage(700, 30));
        res.write('data: [DONE]\n\n');
        res.end();
        return;
      }

      // 授权分档：同表连发三次格式类，再来一次结构类。
      //
      // 验的是「同表同类只问一次，结构仍单独问」。三次格式必须是**不同工具**
      // （format_range / set_number_format / fit_range）：同名工具连发只能证明
      // 去重，证不了整个格式类共用一笔授权。最后那次 add_worksheet 是关键——
      // 格式授权不能把它捎带放行。
      if (scenario === 'grant') {
        const step = grantStep++;
        const calls = [
          { name: 'format_range', args: '{"range":"A1:B2","sheet":"Sheet1","bold":true}' },
          { name: 'set_number_format', args: '{"range":"A1:B2","sheet":"Sheet1","format_code":"0.00"}' },
          { name: 'fit_range', args: '{"range":"A1:B2","sheet":"Sheet1"}' },
          { name: 'add_worksheet', args: '{"name":"新表"}' },
        ];

        if (step < calls.length) {
          const call = calls[step];
          sse(res, textFrame(`第 ${step + 1} 步：${call.name}。`));
          await new Promise((r) => setTimeout(r, 60));
          sse(res, {
            choices: [{
              index: 0,
              delta: {
                tool_calls: [{
                  index: 0,
                  id: `call_grant_${step}`,
                  type: 'function',
                  function: { name: call.name, arguments: call.args },
                }],
              },
            }],
          });
          sse(res, finish('tool_calls'));
          sse(res, usage(80, 20));
          res.write('data: [DONE]\n\n');
          res.end();
          return;
        }

        sse(res, textFrame('四步都做完了。'));
        sse(res, finish('stop'));
        sse(res, usage(120, 30));
        res.write('data: [DONE]\n\n');
        res.end();
        return;
      }

      if (scenario === 'tool' && !hasToolResult) {
        // 分帧下发工具调用参数，验证增量拼接。
        sse(res, textFrame('我先看一下工作簿结构。'));
        await new Promise((r) => setTimeout(r, 120));

        sse(res, {
          choices: [{
            index: 0,
            delta: {
              tool_calls: [{
                index: 0,
                id: 'call_mock_1',
                type: 'function',
                function: { name: 'write_values', arguments: '' },
              }],
            },
          }],
        });

        const argChunks = [
          '{"range"',
          ':"A1:B2"',
          ',"values":[["名称","数量"]',
          ',["铅笔",10]]}',
        ];
        for (const chunk of argChunks) {
          await new Promise((r) => setTimeout(r, 80));
          sse(res, {
            choices: [{
              index: 0,
              delta: { tool_calls: [{ index: 0, function: { arguments: chunk } }] },
            }],
          });
        }

        await new Promise((r) => setTimeout(r, 80));
        sse(res, finish('tool_calls'));
        sse(res, usage(120, 45));
        res.write('data: [DONE]\n\n');
        res.end();
        return;
      }

      // 收尾轮：逐字下发一段含 Markdown 的文本。
      const reply = '已写入 **A1:B2**。\n\n| 名称 | 数量 |\n| --- | --- |\n| 铅笔 | 10 |\n\n还需要我做什么？';
      for (const piece of reply.match(/.{1,6}/gs) ?? []) {
        await new Promise((r) => setTimeout(r, 40));
        sse(res, textFrame(piece));
      }

      sse(res, finish('stop'));
      sse(res, usage(200, 60));
      res.write('data: [DONE]\n\n');
      res.end();
    };

    void send();
  });
});

server.listen(port, '127.0.0.1', () => {
  console.log(`mock listening http://127.0.0.1:${port}/v1  scenario=${scenario}`);
});
