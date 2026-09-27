using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ChatSheet.AddIn.Hosts;
using ChatSheet.AddIn.Providers;
using ChatSheet.AddIn.Storage;
using ChatSheet.AddIn.Tools;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ChatSheet.AddIn.Agent
{
    /// <summary>
    /// 一笔授权覆盖的操作类别。
    ///
    /// 分类的唯一标准是「用户看过这一次之后，同一笔授权还会放过什么」。
    /// 因此按破坏性分层，而不是按「都改单元格」这种实现上的相似性：
    ///
    /// - <see cref="Format"/> 只改外观与尺寸，值不动。
    /// - <see cref="Write"/> 覆盖值或公式，改动可由快照完整还原。
    /// - <see cref="Destructive"/> 抹除或重排：清除会抹掉内容与格式，
    ///   合并会静默丢弃非锚点单元格的值，排序会整行搬动。批准一次写入
    ///   不该顺带放行清空整表——那是量级不同的两件事。
    /// - <see cref="Structure"/> 改工作簿骨架，永远单独问。
    /// </summary>
    internal enum ApprovalClass
    {
        Format = 0,
        Write = 1,
        Destructive = 2,
        Structure = 3,
    }

    /// <summary>本轮里的一个授权：宿主解析后的工作表名 + 操作类别。</summary>
    internal sealed class ApprovalGrant
    {
        internal string SheetName { get; set; }

        internal ApprovalClass Class { get; set; }
    }

    /// <summary>Agent 向面板推送的进度事件。</summary>
    internal sealed class AgentUpdate
    {
        internal string Kind { get; set; }

        internal string Text { get; set; }

        internal object Payload { get; set; }
    }

    /// <summary>
    /// 审批卡片要展示的影响范围。
    ///
    /// 探到具体范围时给出结构化字段，由面板负责组装文案——地址的中文说明
    /// （「2-10 行 × B-D 列」）在面板侧已有一份实现，这里再拼一遍字符串
    /// 就会出现两套措辞。拿不到范围时才退回 Text。
    /// </summary>
    internal sealed class ImpactEstimate
    {
        /// <summary>没有范围可报时的兜底说明，可为空。</summary>
        internal string Text { get; set; } = string.Empty;

        internal string SheetName { get; set; }

        /// <summary>宿主解析后的 A1 地址；为空表示没探到范围。</summary>
        internal string Address { get; set; }

        internal long? CellCount { get; set; }

        /// <summary>
        /// 「将改成什么」的截断对照。只有写入类调用会有，且只送给面板。
        ///
        /// 审批卡此前只报形状（values: 20 行 × 3 列），用户批准的是一个尺寸
        /// 而不是内容。而估算影响时已经把整片当前值读出来过一次，
        /// 原先读完就丢——预览用的正是那一次的结果，不额外读第二遍。
        /// </summary>
        internal ChangePreview Preview { get; set; }

        /// <summary>
        /// 补充一句影响说明，与单元格数分开。
        ///
        /// 适配与自动调整动的是整列整行，不是所选那 N 个单元格；
        /// 新建工作表会成为活动表。这些不能塞进 CellCount。
        /// </summary>
        internal string Note { get; set; }

        /// <summary>宿主专属审批预览；Word 使用 Story/Start/End 与前后文本，不复用单元格对照。</summary>
        internal object HostPreview { get; set; }
    }

    /// <summary>审批请求的结果。</summary>
    internal sealed class ApprovalDecision
    {
        internal bool Approved { get; set; }

        internal string Reason { get; set; }

        /// <summary>本轮后续同类操作是否一并批准。</summary>
        internal bool ApproveRest { get; set; }

        /// <summary>
        /// 用户明确选择「含结构允许」。
        ///
        /// 与 ApproveRest 分开是硬边界：今天那个布尔一旦为真，格式清理之后的
        /// add_worksheet / create_chart 也不再问。结构必须是一次另说清楚的选择。
        /// </summary>
        internal bool ApproveStructureRest { get; set; }
    }

    /// <summary>
    /// Agent 主循环：请求模型 → 收集工具调用 → 审批 → 执行 → 回灌结果 → 继续，
    /// 直到模型不再请求工具或达到步数上限。
    /// </summary>
    internal sealed class AgentRunner
    {
        private readonly Func<object> _applicationAccessor;
        private readonly ToolExecutor _tools;
        private readonly WorkbookContext _context;
        private readonly Conversation _conversation = new Conversation();

        /// <summary>
        /// 把委托切到 UI 线程执行。
        ///
        /// 宿主的 COM 对象是 STA 绑定的，而 Agent 循环在 await 之后运行于线程池线程。
        /// 跨单元调用 Excel 会不稳定，可能抛 RPC_E_SERVERCALL_RETRYLATER，
        /// 也可能在宿主繁忙时死锁。因此所有触碰工作簿的操作都必须回到 UI 线程。
        ///
        /// 注意工具层的单元测试跑在 [STAThread] 的 Main 上，天然是 STA，
        /// 因此掩盖了这个问题——真实加载项里必须显式切换。
        /// </summary>
        private readonly Func<Func<object>, Task<object>> _uiInvoker;

        /// <summary>本轮已批准的范围：工作表 + 风险类。只活到本轮结束，绝不落盘。</summary>
        private readonly List<ApprovalGrant> _approvalGrants = new List<ApprovalGrant>();

        /// <summary>
        /// 本轮传入的设置对象。
        ///
        /// 审批策略必须读这份活对象，而不是开轮时拍一份布尔：对话页切换策略
        /// 改的是同一份实例，拍成布尔的话轮内切换永远不生效——规范明确要求
        /// 下一刀写操作就要按新策略走。
        /// </summary>
        private Settings _liveSettings;
        private string _lastFailedCall;
        private int _sameFailureCount;

        /// <summary>
        /// 本轮实际使用的工具形态。可能在轮内降级（原生 → 文本 → 顾问），
        /// 因此不能每步都从档案里重取——降级发生在步中间。
        /// </summary>
        private ToolProtocolMode _toolMode = ToolProtocolMode.Native;

        /// <summary>本轮的能力档案。降级要写回它，下一轮才不必重蹈。</summary>
        private ModelCapability _capability;

        /// <summary>本轮判定要记到哪个「连接 + 模型」上。</summary>
        private string _connectionKey = string.Empty;

        private string _model = string.Empty;

        /// <summary>
        /// 图片转述的缓存，键是图片在本轮里的序号。
        /// 一轮可能跑几十步，每步都重新转述同一张图既慢又多花钱。
        /// </summary>
        private readonly Dictionary<int, string> _relayedImages = new Dictionary<int, string>();

        /// <summary>
        /// 文本协议下连续多少步没能给出可用指令块。
        ///
        /// 必须按轮归零，且一有进展就归零——与 consecutiveStalls 同一个道理。
        /// 累计计数会把「一轮寒暄 + 下一轮寒暄」误判成「这个模型不会用指令块」，
        /// 于是永久降级成顾问模式：此后它再也不能改表格，而它其实一直都能。
        /// </summary>
        private int _textProtocolMisses;
        private WorkbookTarget _turnWorkbook;
        private WorkbookTarget _operationWorkbook;
        private CancellationToken _turnCancellation;

        internal AgentRunner(Func<object> applicationAccessor, Func<Func<object>, Task<object>> uiInvoker = null)
        {
            _applicationAccessor = applicationAccessor;
            _tools = new ToolExecutor(applicationAccessor);
            _context = new WorkbookContext(applicationAccessor);
            // 未提供切换器时原地执行，便于在 STA 测试宿主中直接使用。
            _uiInvoker = uiInvoker ?? (work => Task.FromResult(work()));
        }

        /// <summary>在 UI 线程上执行工具，避免跨 COM 单元调用宿主。</summary>
        private async Task<ToolResult> ExecuteOnUiAsync(string name, JObject args, string undoId = null)
        {
            var result = await InvokeForWorkbookAsync(() => {
                _tools.BeforeWriteChunk = () => _turnCancellation.ThrowIfCancellationRequested();
                try { return _tools.Execute(name, args, undoId); }
                finally { _tools.BeforeWriteChunk = null; }
            }).ConfigureAwait(false);
            return (ToolResult)result;
        }

        private Task<object> InvokeForWorkbookAsync(Func<object> work)
        {
            return _uiInvoker(() =>
            {
                _turnCancellation.ThrowIfCancellationRequested();
                if ((_turnWorkbook != null && !_turnWorkbook.IsCurrent(_applicationAccessor())) ||
                    (_operationWorkbook != null && !_operationWorkbook.IsCurrent(_applicationAccessor())))
                {
                    throw new ProviderException("WORKBOOK_CHANGED", "工作簿已切换，本轮已停止。请回到原工作簿后重新发送请求。");
                }
                return work();
            });
        }

        internal ToolExecutor Tools => _tools;

        internal Conversation Conversation => _conversation;

        internal void Reset()
        {
            _conversation.Clear();
        }

        /// <summary>
        /// 收回本轮全部授权，后续写操作重新逐个询问。
        ///
        /// 轮次进行中也能调用：授权只是「下一次要不要弹卡」的判据，
        /// 清掉它不影响正在执行的那一次调用，也不停止本轮。
        /// 用户想停的往往只是自动放行，不是整个任务。
        /// </summary>
        internal int RevokeApprovalGrants()
        {
            lock (_approvalGrants)
            {
                var count = _approvalGrants.Count;
                _approvalGrants.Clear();
                return count;
            }
        }

        /// <summary>
        /// 执行一轮用户请求。
        /// </summary>
        /// <param name="userInput">用户输入。</param>
        /// <param name="settings">当前设置。</param>
        /// <param name="onUpdate">进度回调，用于把流式内容推给面板。</param>
        /// <param name="requestApproval">审批回调，返回用户决定。</param>
        internal async Task RunAsync(
            string userInput,
            Settings settings,
            Func<AgentUpdate, Task> onUpdate,
            Func<ToolDefinition, JObject, ImpactEstimate, Task<ApprovalDecision>> requestApproval,
            CancellationToken cancellationToken,
            IReadOnlyList<ImageAttachment> images = null)
        {
            _turnCancellation = cancellationToken;
            cancellationToken.ThrowIfCancellationRequested();
            _turnWorkbook = (WorkbookTarget)await _uiInvoker(() => WorkbookTarget.Capture(_applicationAccessor())).ConfigureAwait(false);
            try
            {
                await RunCoreAsync(userInput, settings, onUpdate, requestApproval, cancellationToken, images).ConfigureAwait(false);
            }
            finally
            {
                var target = _turnWorkbook;
                _turnWorkbook = null;
                target?.Dispose();
            }
        }

        private async Task RunCoreAsync(
            string userInput,
            Settings settings,
            Func<AgentUpdate, Task> onUpdate,
            Func<ToolDefinition, JObject, ImpactEstimate, Task<ApprovalDecision>> requestApproval,
            CancellationToken cancellationToken,
            IReadOnlyList<ImageAttachment> images = null)
        {
            // 只带图片不写文字也应允许：贴一张截图问「这个怎么填」是常见用法。
            if (string.IsNullOrWhiteSpace(userInput) && (images == null || images.Count == 0))
            {
                throw new ProviderException("EMPTY_INPUT", "请输入内容或附加图片。");
            }

            var connection = settings.ResolveConnection();
            if (!connection.IsWorkBuddy && string.IsNullOrWhiteSpace(connection.Model))
            {
                throw new ProviderException("MODEL_REQUIRED", "尚未选择模型，请到设置页选择。");
            }

            // 授权只活本轮。关面板不会销毁 runner，若不在这里清空，
            // 下一轮甚至下一个工作簿会继承上一次的「允许」，等于悄悄变成全自动。
            _liveSettings = settings;
            _lastFailedCall = null;
            _sameFailureCount = 0;
            RevokeApprovalGrants();

            // 本轮的能力档案与起始工具形态。
            //
            // 手动指定时直接采用用户的选择，不再探测：服务端静默忽略工具声明的
            // 情形探测无从触发，而用户已经知道结果。
            _capability = ModelCapabilities.For(settings.ConnectionKey(), connection.Model);
            // ACP 当前不声明 ChatSheet 的原生工具 schema，因此固定使用现有文本指令协议，
            // 这样审批、执行和工具结果回灌仍走同一条 Agent 链路。
            _toolMode = connection.IsWorkBuddy
                ? ToolProtocolMode.Text
                : ModelCapabilities.ResolveMode(settings.ToolProtocol, _capability, connection.Model);
            _relayedImages.Clear();

            // 本轮的可用性判定要记到哪。与能力档案同键，但两者互不改写。
            _connectionKey = settings.ConnectionKey();
            _model = connection.Model;

            // 过程计数按轮归零。留到下一轮就变成「凭历史降级」：
            // 上一轮末尾那步本来就不该有工具调用（模型在作答），不是它不会用。
            _textProtocolMisses = 0;

            if (_toolMode != ToolProtocolMode.Native)
            {
                Log.Info($"本轮工具形态：{_toolMode}" +
                    (settings.ToolProtocol == ToolProtocolPreference.Auto ? "（按已探测到的能力）" : "（用户指定）"));
            }

            // 每轮刷新系统提示：工作簿可能已被用户手动改动。
            await RefreshSystemPromptAsync(settings).ConfigureAwait(false);
            _conversation.Add(ChatMessage.FromUser(userInput ?? string.Empty, images));

            if (images != null && images.Count > 0)
            {
                Log.Info($"本轮附带 {ImageSupport.Describe(images)}");

                // 已知这个模型看不了图时不必先撞一次 400，直接走回退。
                if (_capability.VisionUnsupported)
                {
                    await ApplyVisionFallbackAsync(settings, connection, onUpdate, cancellationToken)
                        .ConfigureAwait(false);
                }
            }

            // 连续被截断的次数。只要有一步正常产出就归零：
            // 用累计次数会让一轮里偶发几次截断也把续跑额度耗尽。
            var consecutiveStalls = 0;

            using (var client = CreateChatClient(settings, connection))
            {
                for (var step = 0; settings.MaxSteps == 0 || step < settings.MaxSteps; step++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // 先推送压缩前的占用，再压缩。
                    //
                    // 顺序很关键：若先压缩再推送，界面永远只看到压缩后的低占比，
                    // 用户无从知道曾经触达阈值，圆环也不会有任何提示——
                    // 「到达 90% 时提示是否压缩」这件事就完全不可见了。
                    await PushContextAsync(onUpdate, settings).ConfigureAwait(false);

                    var trim = _conversation.TrimToBudget(settings.ContextBudgetTokens);
                    if (trim.Trimmed)
                    {
                        Log.Info($"上下文压缩：{trim.TokensBefore} → {trim.TokensAfter} tokens" +
                            $"（阈值 {trim.TriggerTokens}，预算 {trim.BudgetTokens}），" +
                            $"压缩 {trim.CompressedToolResults} 条工具结果、移除 {trim.DroppedMessages} 条早期记录");

                        await onUpdate(new AgentUpdate
                        {
                            Kind = "context-trimmed",
                            Payload = new
                            {
                                budget = trim.BudgetTokens,
                                trigger = trim.TriggerTokens,
                                before = trim.TokensBefore,
                                after = trim.TokensAfter,
                                compressed = trim.CompressedToolResults,
                                dropped = trim.DroppedMessages,
                            },
                        }).ConfigureAwait(false);

                        // 压缩后再推一次，让圆环回落到实际值。
                        await PushContextAsync(onUpdate, settings).ConfigureAwait(false);
                    }

                    // 本步的请求与交付。能力回退（换工具形态、处理图片）会在这里
                    // 原地重跑本步，因此重跑不消耗步数——上一次尝试没有任何进展。
                    var outcome = await RunStepAsync(
                        client, connection, settings, onUpdate, cancellationToken).ConfigureAwait(false);

                    var assistantText = outcome.Text;
                    var pendingCalls = outcome.Calls;
                    var finishReason = outcome.FinishReason;
                    var stepCompletionTokens = outcome.CompletionTokens;

                    // 用户中途点了停止：必须在这里显式抛出，不能往下走。
                    //
                    // 取消是通过关闭 HTTP 流实现的，而流被关掉在读取侧表现为
                    // 正常结束（EOF），不抛异常。于是这一步会被当成「模型说完了」：
                    // 日志记下「对话结束」、面板收到 turn-complete，用户点了停止却
                    // 既看不到停止回执，那句被截断的半截回复还会作为完整回答留在
                    // 上下文里。下一轮模型便以为自己已经答完了。
                    //
                    // 也必须在判定停顿之前：停止时若正文恰好为空，会被当成「无进展」
                    // 而触发自动续跑——用户要的是停下来，不是接着跑。
                    cancellationToken.ThrowIfCancellationRequested();

                    // 本步是被打断还是真的说完了。
                    //
                    // 必须先判这一条，两条能力判据才有意义：被长度上限截断的一步
                    // 既没有工具调用也没有指令块，形态与「不会用工具」一模一样。
                    // 顺序反了，输出上限设小一点就会被判成模型不支持工具。
                    var stall = TurnOutcome.Classify(
                        finishReason,
                        pendingCalls.Count > 0,
                        assistantText.Length,
                        stepCompletionTokens,
                        settings.MaxOutputTokens);

                    // 模型收下了工具声明却一个都没用，还回了一句「我碰不到你的表格」。
                    //
                    // 这是不具备工具能力的模型最常见的表现——服务端不报任何错，
                    // 所以只能从这句推辞认出来。认出后换成文本协议原地重跑本步，
                    // 并且不把这句推辞留进上下文：留着的话模型下一步会把
                    // 「我没有权限」当成已经确立的事实继续沿用。
                    if (stall == StepStall.None &&
                        ShouldProbeToolRefusal(settings, pendingCalls.Count, assistantText))
                    {
                        _capability.ToolRefusalProbed = true;
                        await SwitchToolModeAsync(
                            ToolProtocolMode.Text,
                            settings,
                            onUpdate,
                            "该模型没有发起工具调用，并表示工具能力不可用。已改用文本指令方式重试。")
                            .ConfigureAwait(false);

                        // 原地重跑本步：上一次尝试没有任何进展，不该占掉一个步数。
                        step--;
                        continue;
                    }

                    // 文本协议下连续几步既没有指令块也没有正常收尾，说明它连指令块
                    // 也写不对。此时退为顾问模式，避免最坏情况：模型既动不了手，
                    // 又被系统提示反复告知「你有读写权限」，于是编造出「已经填好了」。
                    if (stall == StepStall.None &&
                        ShouldDegradeToAdvisor(settings, pendingCalls.Count, outcome.SawToolBlock))
                    {
                        await SwitchToolModeAsync(
                            ToolProtocolMode.None,
                            settings,
                            onUpdate,
                            "该模型既不支持原生工具调用，也未能按格式发出指令块，无法直接操作表格。" +
                                "已切换为顾问模式：它会给出公式与操作步骤，由你在表格里执行。")
                            .ConfigureAwait(false);
                    }

                    // 记录助手消息，含工具调用，供下一轮作为历史发送。
                    // 被打断且一个字都没留下时给一句占位：Anthropic 要求 user 与
                    // assistant 交替出现，空内容的助手消息会被服务端拒绝，
                    // 而下面紧接着要追加一条催促用的 user 消息。
                    //
                    // 文本协议下用未经闸门过滤的原文：指令块要留在历史里，
                    // 模型才看得见自己发过什么调用。用户那侧看到的是操作卡片。
                    var assistantMessage = ChatMessage.FromAssistant(
                        stall != StepStall.None && assistantText.Length == 0
                            ? "（本次输出被长度上限截断，未能给出内容。）"
                            : outcome.RawText);
                    assistantMessage.ToolCalls.AddRange(pendingCalls);
                    _conversation.Add(assistantMessage);

                    if (stall != StepStall.None)
                    {
                        consecutiveStalls++;

                        if (consecutiveStalls > TurnOutcome.MaxAutoContinues)
                        {
                            Log.Warn($"连续 {consecutiveStalls - 1} 次输出被截断后仍无进展，停止续跑" +
                                $"（输出上限 {settings.MaxOutputTokens}）");

                            await onUpdate(new AgentUpdate
                            {
                                Kind = "stalled",
                                Text = $"模型的输出连续 {TurnOutcome.MaxAutoContinues} 次被长度上限截断" +
                                    $"（当前上限 {settings.MaxOutputTokens} tokens），已停止。" +
                                    "请到设置页调高「最大输出 tokens」，或调低思考档位后重试。",
                            }).ConfigureAwait(false);
                            return;
                        }

                        // 自动续跑，不再要求用户手动催一次。
                        Log.Info($"第 {step + 1} 步无进展（{stall}，结束原因={finishReason ?? "<未提供>"}，" +
                            $"输出={stepCompletionTokens}/{settings.MaxOutputTokens}），" +
                            $"自动继续（第 {consecutiveStalls}/{TurnOutcome.MaxAutoContinues} 次）");

                        _conversation.Add(ChatMessage.FromUser(ContinueNudge(stall)));

                        await onUpdate(new AgentUpdate
                        {
                            Kind = "auto-continue",
                            Text = stall == StepStall.Truncated
                                ? "输出被长度上限截断，正在自动继续…"
                                : "上一步没有产出，正在自动继续…",
                            Payload = new
                            {
                                reason = stall.ToString(),
                                attempt = consecutiveStalls,
                                maxAttempts = TurnOutcome.MaxAutoContinues,
                                completionTokens = stepCompletionTokens,
                                maxOutputTokens = settings.MaxOutputTokens,
                            },
                        }).ConfigureAwait(false);

                        continue;
                    }

                    consecutiveStalls = 0;

                    if (pendingCalls.Count == 0)
                    {
                        Log.Info($"对话结束：步数={step + 1} 结束原因={finishReason ?? "<未提供>"} " +
                            $"回复长度={assistantText.Length} 用量={_conversation.LastPromptTokens}入/{_conversation.LastCompletionTokens}出");
                        await onUpdate(new AgentUpdate { Kind = "turn-complete", Payload = new { finishReason } })
                            .ConfigureAwait(false);
                        return;
                    }

                    Log.Info($"第 {step + 1} 步：模型请求 {pendingCalls.Count} 个工具调用（" +
                        string.Join("、", pendingCalls.ConvertAll(c => c.Name)) + "）");

                    foreach (var call in pendingCalls)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        await ExecuteOneAsync(call, settings, onUpdate, requestApproval).ConfigureAwait(false);
                    }
                }

                // 达到步数上限：明确告知而不是静默停止。
                await onUpdate(new AgentUpdate
                {
                    Kind = "step-limit",
                    Text = $"已达到单轮步数上限（{settings.MaxSteps} 步）。任务可能尚未完成，可继续输入以让我接着处理，或到设置页提高上限。",
                }).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// 自动续跑时给模型的提示。
        ///
        /// 必须写明「不要重述」：模型看到自己上一条是残句时，默认会从头再讲一遍，
        /// 于是又在同一个位置被截断。也必须写明「先动手再解释」——被截断的正是
        /// 长篇推理与说明，把工具调用排到前面才有机会在上限内发出去。
        /// </summary>
        private static string ContinueNudge(StepStall stall)
        {
            if (stall == StepStall.Truncated)
            {
                return "（系统提示：你上一次的输出因达到长度上限被截断，用户没有看到完整内容，" +
                    "也没有需要用户确认的事情。请直接继续未完成的任务：先发出下一个工具调用，" +
                    "再用一两句话说明结果。不要重述已经说过的部分，不要复述计划，" +
                    "单次写入的数据量要小一些以免再被截断。）";
            }

            return "（系统提示：你上一次没有任何输出，任务尚未完成。请直接继续：" +
                "该调用工具就调用工具，该作答就作答，不要等待用户回应。）";
        }

        /// <summary>一步的交付结果。</summary>
        private sealed class StepOutcome
        {
            /// <summary>给用户看的正文（文本协议下已剔除指令块）。</summary>
            internal string Text { get; set; } = string.Empty;

            /// <summary>进上下文的原文，含指令块。</summary>
            internal string RawText { get; set; } = string.Empty;

            internal List<ToolCall> Calls { get; } = new List<ToolCall>();

            internal string FinishReason { get; set; }

            internal int CompletionTokens { get; set; }

            /// <summary>文本协议下本步是否出现过指令块。</summary>
            internal bool SawToolBlock { get; set; }
        }

        private static IChatStreamClient CreateChatClient(Settings settings, ResolvedConnection connection)
        {
            return settings.Mode.IsWorkBuddy()
                ? WorkBuddyProvider.CreateChatClient(settings.Mode)
                : (IChatStreamClient)new ChatClient(connection.ProxyCandidates != null && connection.ProxyCandidates.Count > 0
                    ? connection.ProxyCandidates : new[] { connection.Proxy });
        }

        /// <summary>
        /// 跑一步，并在能力不匹配时就地换个形态重来。
        ///
        /// 重来放在这一层而不是主循环，是因为「换形态重试」与主循环的步数、
        /// 截断续跑、停顿判定都无关——上一次尝试连一个字都没交付出去，
        /// 对上层来说等于没发生过。
        /// </summary>
        private async Task<StepOutcome> RunStepAsync(
            IChatStreamClient client,
            ResolvedConnection connection,
            Settings settings,
            Func<AgentUpdate, Task> onUpdate,
            CancellationToken cancellationToken)
        {
            // 两次足够：工具形态降一级、图片处理一次。再多说明判断有误，
            // 继续重试只会反复撞同一个 400。
            const int maxFallbacks = 2;

            for (var attempt = 0; ; attempt++)
            {
                try
                {
                    return await StreamStepAsync(client, connection, settings, onUpdate, cancellationToken)
                        .ConfigureAwait(false);
                }
                // 判定记在 when 子句里，借的是「过滤器先于 catch 体求值、且无论
                // 哪个 catch 最终命中都会跑过」这一点：同一条「模型不存在」既该被
                // 记成不可用，也可能因为模型名含 image 而落进底下的视觉分支，
                // 记在任一 catch 体里都会漏掉另一条路。RecordFailure 恒返回 false，
                // 只记账、不拦截。
                catch (ProviderException ex) when (RecordFailure(ex, connection.Model))
                {
                    throw new InvalidOperationException("不可达：RecordFailure 恒为假");
                }
                catch (ProviderException ex) when (
                    attempt < maxFallbacks &&
                    _toolMode == ToolProtocolMode.Native &&
                    ModelCapabilities.DetectionEnabled(settings.ToolProtocol) &&
                    CapabilitySignals.LooksLikeToolUnsupported(ex))
                {
                    await SwitchToolModeAsync(
                        ToolProtocolMode.Text,
                        settings,
                        onUpdate,
                        $"该模型不支持原生工具调用（接口回复：{ex.Message}）。已改用文本指令方式，功能不变。")
                        .ConfigureAwait(false);
                }
                // 输出上限的字段名用错了：换一个重来。
                //
                // 排在工具与视觉之后是因为它最窄——判据要求错误同时点名字段与「不支持」，
                // 不会与那两条抢。而它必须在这里而不是只在探测里：真实对话发的是同一个
                // 字段名，推理模型在真实对话里也会被这条 400 挡住。
                catch (ProviderException ex) when (
                    attempt < maxFallbacks &&
                    connection.Protocol == ProtocolKind.OpenAiChatCompletions &&
                    _capability != null &&
                    _capability.OutputLimit != OutputLimitField.MaxCompletionTokens &&
                    CapabilitySignals.LooksLikeOutputLimitFieldWrong(ex))
                {
                    _capability.OutputLimit = OutputLimitField.MaxCompletionTokens;
                    Log.Warn($"模型 {connection.Model} 不接受 max_tokens，改用 max_completion_tokens 重试（{ex.Message}）");
                }
                catch (ProviderException ex) when (
                    attempt < maxFallbacks &&
                    ConversationHasImages() &&
                    CapabilitySignals.LooksLikeVisionUnsupported(ex))
                {
                    _capability.VisionUnsupported = true;
                    Log.Warn($"模型 {connection.Model} 不支持图片输入（{ex.Message}），转入视觉回退");

                    await ApplyVisionFallbackAsync(settings, connection, onUpdate, cancellationToken)
                        .ConfigureAwait(false);
                }
            }
        }

        /// <summary>
        /// 把一轮失败折算成可用性判定。恒返回 false——它只记账，不参与异常筛选。
        ///
        /// 放在 when 子句里而不是 catch 体里：本方法的调用点有三条出路（换工具形态、
        /// 转视觉回退、异常上抛），记在任何一条里都会漏掉另外两条。
        /// </summary>
        private bool RecordFailure(ProviderException ex, string model)
        {
            try
            {
                var verdict = ModelAvailability.Classify(ex, model);
                if (verdict == AvailabilityVerdict.Unavailable)
                {
                    Log.Info($"模型 {model} 判为不可用（服务端原文点名了它：{ex.Detail}）");
                }

                ModelAvailability.Record(_connectionKey, model, verdict);
            }
            catch (Exception recordError)
            {
                // 记账失败绝不能改变这一轮的走向。
                Log.Warn("记录可用性判定失败：" + recordError.Message);
            }

            return false;
        }

        /// <summary>发一次请求并把流式事件交付出去。</summary>
        private async Task<StepOutcome> StreamStepAsync(
            IChatStreamClient client,
            ResolvedConnection connection,
            Settings settings,
            Func<AgentUpdate, Task> onUpdate,
            CancellationToken cancellationToken)
        {
            var request = BuildRequest(connection, settings);
            var outcome = new StepOutcome();
            var raw = new System.Text.StringBuilder();
            var visible = new System.Text.StringBuilder();

            // 文本协议下正文要过一道闸门，把指令块拦在气泡之外。
            // 原生模式不建闸门：那条路上正文里不会有指令块，
            // 多一层缓冲只会让逐字输出显得迟滞。
            var gate = _toolMode == ToolProtocolMode.Text ? new TextToolGate() : null;

            await client.StreamAsync(
                request,
                async chatEvent =>
                {
                    // 收到任何事件就说明请求到达了模型——这就是可用性的全部主张，
                    // 之后无论出什么错都不改这个结论。
                    //
                    // 刻意不用 ChatClient 里的 delivered：那是重试循环内的局部变量，
                    // 只说明「这一次尝试」交付过，退避重来时会重新置假。
                    ModelAvailability.Record(_connectionKey, _model, AvailabilityVerdict.Available);

                    switch (chatEvent.Kind)
                    {
                        case ChatEventKind.TextDelta:
                            raw.Append(chatEvent.Text);

                            if (gate == null)
                            {
                                visible.Append(chatEvent.Text);
                                await onUpdate(new AgentUpdate { Kind = "text", Text = chatEvent.Text }).ConfigureAwait(false);
                                break;
                            }

                            var passed = gate.Push(chatEvent.Text);
                            if (passed.Length > 0)
                            {
                                visible.Append(passed);
                                await onUpdate(new AgentUpdate { Kind = "text", Text = passed }).ConfigureAwait(false);
                            }

                            break;

                        case ChatEventKind.ThinkingDelta:
                            await onUpdate(new AgentUpdate { Kind = "thinking", Text = chatEvent.Text }).ConfigureAwait(false);
                            break;

                        case ChatEventKind.ToolCall:
                            outcome.Calls.Add(chatEvent.Call);
                            break;

                        case ChatEventKind.Usage:
                            if (chatEvent.CompletionTokens > 0)
                            {
                                outcome.CompletionTokens = chatEvent.CompletionTokens;
                            }

                            _conversation.RecordUsage(chatEvent.PromptTokens, chatEvent.CompletionTokens);
                            await onUpdate(new AgentUpdate
                            {
                                Kind = "usage",
                                Payload = new
                                {
                                    promptTokens = chatEvent.PromptTokens,
                                    completionTokens = chatEvent.CompletionTokens,
                                    estimatedContext = _conversation.EstimateTotalTokens(),
                                    budget = settings.ContextBudgetTokens,
                                },
                            }).ConfigureAwait(false);
                            break;

                        case ChatEventKind.Completed:
                            // 只认非空值：OpenAI 兼容协议会先在带 finish_reason 的帧
                            // 上报一次，随后的 [DONE] 又报一次且不带原因。
                            // 直接覆盖会把刚拿到的 length 抹成空，截断也就无从判断。
                            if (!string.IsNullOrEmpty(chatEvent.FinishReason))
                            {
                                outcome.FinishReason = chatEvent.FinishReason;
                            }

                            break;
                    }
                },
                cancellationToken,
                // 重试期间界面必须有说明：否则退避等待期表现为「发出去了但毫无反应」，
                // 用户往往会以为卡死而反复点发送。
                (attempt, delay, reason) => onUpdate(new AgentUpdate
                {
                    Kind = "retry",
                    Text = RetryPolicy.Describe(attempt, delay, reason),
                    Payload = new
                    {
                        attempt,
                        maxRetries = RetryPolicy.MaxRetries,
                        delaySeconds = (int)Math.Round(delay.TotalSeconds),
                        reason,
                    },
                })).ConfigureAwait(false);

            if (gate != null)
            {
                // 收束闸门：流可能断在围栏中间，攥着的内容必须在这里定性，
                // 否则这段文字既不显示也不执行，凭空消失。
                var tail = gate.Flush();
                if (tail.Length > 0)
                {
                    visible.Append(tail);
                    await onUpdate(new AgentUpdate { Kind = "text", Text = tail }).ConfigureAwait(false);
                }

                outcome.SawToolBlock = gate.SawToolBlock;

                // 指令块转成与原生调用同形的 ToolCall，后续链路完全共用。
                var index = 0;
                foreach (var parsed in gate.Calls)
                {
                    outcome.Calls.Add(new ToolCall
                    {
                        Id = "txt" + (++index) + "-" + Guid.NewGuid().ToString("N").Substring(0, 6),
                        Name = parsed.Name,
                        ArgumentsJson = parsed.ArgumentsJson,
                    });
                }
            }

            outcome.RawText = raw.ToString();
            outcome.Text = visible.ToString();
            return outcome;
        }

        /// <summary>
        /// 换一种工具形态并写回档案。
        ///
        /// 必须同时刷新系统提示：文本协议要把工具清单写进去，顾问模式要把
        /// 「你已经连上工作簿」收回来。漏了这一步，模型收到的还是上一种形态的说明。
        /// </summary>
        private async Task SwitchToolModeAsync(
            ToolProtocolMode mode,
            Settings settings,
            Func<AgentUpdate, Task> onUpdate,
            string notice)
        {
            _toolMode = mode;
            _capability.ToolMode = mode;

            // 换了形态，上一种形态攒下的未命中数就不再说明任何事情。
            _textProtocolMisses = 0;

            Log.Warn($"工具形态降级为 {mode}：{notice}");
            await RefreshSystemPromptAsync(settings).ConfigureAwait(false);

            await onUpdate(new AgentUpdate
            {
                Kind = "tool-fallback",
                Text = notice,
                Payload = new { mode = mode.ToString() },
            }).ConfigureAwait(false);
        }

        /// <summary>
        /// 是否该用「推辞」这条启发式判定模型没有工具能力。
        ///
        /// 三个条件都必要：必须是自动探测（手动指定时用户已经表态）、
        /// 必须还没试过（推辞也可能是合理拒绝，不能每轮都为同一句话重跑）、
        /// 必须是原生模式下零调用且正文确实在说自己碰不到表格。
        /// </summary>
        private bool ShouldProbeToolRefusal(Settings settings, int callCount, string assistantText)
        {
            return _toolMode == ToolProtocolMode.Native &&
                callCount == 0 &&
                ModelCapabilities.DetectionEnabled(settings.ToolProtocol) &&
                !_capability.ToolRefusalProbed &&
                CapabilitySignals.LooksLikeToolRefusal(assistantText);
        }

        /// <summary>
        /// 文本协议是否已被证明也走不通。
        ///
        /// 只在自动探测下降级，且要求连续两步都没有指令块：一步没有可能是它
        /// 正在正常作答（用户问的就是概念问题），连续两步则说明它根本不会用。
        /// </summary>
        private bool ShouldDegradeToAdvisor(Settings settings, int callCount, bool sawToolBlock)
        {
            if (_toolMode != ToolProtocolMode.Text ||
                !ModelCapabilities.DetectionEnabled(settings.ToolProtocol))
            {
                return false;
            }

            var tally = ModelCapabilities.TallyTextProtocolStep(
                _textProtocolMisses,
                madeProgress: callCount > 0 || sawToolBlock);

            _textProtocolMisses = tally.Misses;
            return tally.ShouldDegrade;
        }

        /// <summary>上下文里是否还有图片待发送。</summary>
        private bool ConversationHasImages()
        {
            foreach (var message in _conversation.Messages)
            {
                if (message.Images.Count > 0)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 处理「模型看不了图」：有中转模型就转述，没有就去图并告知。
        ///
        /// 两条路都必须在上下文里留下痕迹。静默丢掉图片是最坏的做法——
        /// 用户以为模型看过自己的截图，于是会相信一个从未基于它的回答。
        /// </summary>
        private async Task ApplyVisionFallbackAsync(
            Settings settings,
            ResolvedConnection connection,
            Func<AgentUpdate, Task> onUpdate,
            CancellationToken cancellationToken)
        {
            var images = new List<ImageAttachment>();
            foreach (var message in _conversation.Messages)
            {
                images.AddRange(message.Images);
            }

            if (images.Count == 0)
            {
                return;
            }

            var relay = settings.ResolveVisionRelay(connection);
            string notice;
            string injected;

            if (relay != null)
            {
                var descriptions = new List<string>();
                string failure = null;

                for (var i = 0; i < images.Count; i++)
                {
                    // 一轮可能跑几十步，同一张图只转述一次。
                    if (_relayedImages.TryGetValue(i, out var cached))
                    {
                        descriptions.Add(cached);
                        continue;
                    }

                    try
                    {
                        await onUpdate(new AgentUpdate
                        {
                            Kind = "retry",
                            Text = $"正在用 {relay.Model} 转写第 {i + 1}/{images.Count} 张图片…",
                        }).ConfigureAwait(false);

                        var described = await VisionRelay
                            .DescribeAsync(relay, images[i], cancellationToken)
                            .ConfigureAwait(false);

                        _relayedImages[i] = described;
                        descriptions.Add(described);
                    }
                    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        Log.Warn($"视觉中转失败：{ex.Message}");
                        failure = ex.Message;
                        break;
                    }
                }

                if (failure == null)
                {
                    injected = VisionRelay.ComposeDescriptions(descriptions, images.Count);
                    notice = $"当前模型没有视觉能力，已用 {relay.Model} 把 {images.Count} 张图片转写成文字后交给它。";
                    Log.Info($"视觉中转完成：{images.Count} 张图片经 {relay.Model} 转写");
                }
                else
                {
                    injected = VisionRelay.ComposeUnavailableNotice(
                        images.Count, $"视觉中转模型 {relay.Model} 也失败了（{failure}）。");
                    notice = $"当前模型没有视觉能力，视觉中转模型 {relay.Model} 也未能转写（{failure}）。" +
                        "已去掉图片继续这一轮，请改用带视觉的模型，或把图中内容用文字说明。";
                }
            }
            else
            {
                injected = VisionRelay.ComposeUnavailableNotice(images.Count, string.Empty);
                notice = $"当前模型没有视觉能力，{images.Count} 张图片未能送达，已去掉图片继续这一轮。" +
                    "可在设置页把模型换成带视觉的型号，或填写「视觉中转模型」让另一个模型先把图转成文字。";
            }

            // 去掉图片：留着只会在下一步再撞一次同样的 400。
            foreach (var message in _conversation.Messages)
            {
                message.Images.Clear();
            }

            _conversation.Add(ChatMessage.FromUser(injected));

            await onUpdate(new AgentUpdate
            {
                Kind = "vision-fallback",
                Text = notice,
                Payload = new
                {
                    images = images.Count,
                    relayModel = relay?.Model,
                    relayed = relay != null && _relayedImages.Count > 0,
                },
            }).ConfigureAwait(false);
        }

        /// <summary>
        /// 推送上下文占用情况，供界面绘制进度圆环。
        /// nearLimit 由后端判定，界面不重复计算阈值，避免两处规则不一致。
        /// </summary>
        private Task PushContextAsync(Func<AgentUpdate, Task> onUpdate, Settings settings)
        {
            var used = _conversation.EstimateTotalTokens();
            var budget = Math.Max(1, settings.ContextBudgetTokens);
            var ratio = Math.Min(1.0, (double)used / budget);

            return onUpdate(new AgentUpdate
            {
                Kind = "context",
                Payload = new
                {
                    used,
                    budget,
                    ratio,
                    percent = (int)Math.Round(ratio * 100),
                    threshold = (int)(Conversation.CompressionThreshold * 100),
                    nearLimit = ratio >= Conversation.CompressionThreshold,
                    lastPromptTokens = _conversation.LastPromptTokens,
                    lastCompletionTokens = _conversation.LastCompletionTokens,
                },
            });
        }

        private async Task ExecuteOneAsync(
            ToolCall call,
            Settings settings,
            Func<AgentUpdate, Task> onUpdate,
            Func<ToolDefinition, JObject, ImpactEstimate, Task<ApprovalDecision>> requestApproval)
        {
            _operationWorkbook = (WorkbookTarget)await InvokeForWorkbookAsync(() => WorkbookTarget.Capture(_applicationAccessor())).ConfigureAwait(false);
            try
            {
                await ExecuteOneCoreAsync(call, settings, onUpdate, requestApproval).ConfigureAwait(false);
            }
            finally
            {
                var target = _operationWorkbook;
                _operationWorkbook = null;
                target?.Dispose();
            }
        }

        private async Task ExecuteOneCoreAsync(
            ToolCall call,
            Settings settings,
            Func<AgentUpdate, Task> onUpdate,
            Func<ToolDefinition, JObject, ImpactEstimate, Task<ApprovalDecision>> requestApproval)
        {
            // 文本协议下的指令块可能连工具名都没写出来（例如整块 JSON 被截断）。
            // 这时说「未知工具：」会让模型以为自己拼错了名字，转而换个名字重发；
            // 说清是块本身写坏了，它才会重新写一个完整的块。
            if (string.IsNullOrWhiteSpace(call.Name))
            {
                var truncated = TurnOutcome.LooksTruncatedJson(call.ArgumentsJson);
                await FeedToolResultAsync(
                    call,
                    ToolResult.Failure(
                        truncated ? "ARGS_TRUNCATED" : "BLOCK_INVALID",
                        truncated
                            ? "指令块在传输中途被长度上限截断，未能读出工具名。请缩小单次数据量后重新发出完整的指令块。"
                            : "指令块里没有可用的工具名。块内必须是一个 JSON 对象，含 tool 与 args 两个字段。"),
                    onUpdate).ConfigureAwait(false);
                return;
            }

            var definition = ToolCatalog.Find(call.Name);
            if (definition == null)
            {
                await FeedToolResultAsync(call, ToolResult.Failure("UNKNOWN_TOOL", $"未知工具：{call.Name}。"), onUpdate)
                    .ConfigureAwait(false);
                return;
            }

            JObject args;
            try
            {
                args = string.IsNullOrWhiteSpace(call.ArgumentsJson)
                    ? new JObject()
                    : JObject.Parse(call.ArgumentsJson);
            }
            catch (Exception ex)
            {
                // 区分「被截断」与「写错格式」：只说不是合法 JSON 的话，
                // 模型会认为自己拼错了标点，用同样的参数重发，于是断在同一处。
                // 说清是长度截断，它才会改为分批写入。
                if (TurnOutcome.LooksTruncatedJson(call.ArgumentsJson))
                {
                    Log.Warn($"工具 {call.Name} 的参数被截断，长度 {call.ArgumentsJson.Length}");
                    await FeedToolResultAsync(
                        call,
                        ToolResult.Failure(
                            "ARGS_TRUNCATED",
                            TurnOutcome.DescribeTruncatedArguments(call.Name, call.ArgumentsJson.Length)),
                        onUpdate).ConfigureAwait(false);
                    return;
                }

                await FeedToolResultAsync(
                    call,
                    ToolResult.Failure("ARGS_INVALID", $"参数不是合法 JSON：{ex.Message}"),
                    onUpdate).ConfigureAwait(false);
                return;
            }

            await onUpdate(new AgentUpdate
            {
                Kind = "tool-start",
                Payload = new
                {
                    id = call.Id,
                    name = call.Name,
                    risk = definition.Risk.ToString(),
                    args,
                },
            }).ConfigureAwait(false);

            if (definition.RequiresApproval && !IsAutomaticApproval())
            {
                // 审批前先算出影响范围，也用它拿到宿主解析后的工作表名。
                // 授权不信模型 args 里的 sheet：省略时它只是「当前活动表」，
                // 活动表可能在用户批准期间已经变了。
                var impact = await DescribeImpactAsync(definition, args).ConfigureAwait(false);
                var approvalClass = ApprovalClassFor(definition.Name);

                // 把解析结果钉进真正要执行的那份参数。
                //
                // 估算影响与执行是两次独立的 UI 线程往返，中间夹着等用户点按钮
                // 那段时间。省略 sheet 时两次各自去取「当前活动表」，用户在批准
                // 期间切了表，卡上说的是 Sheet1，写入落到另一张表上。
                // fit_range 连范围也要钉：它省略 range 时两侧会各算一次已用范围。
                PinResolvedTarget(definition, args, impact);

                var resolvedSheet = GrantKeyFor(approvalClass, impact, args);

                var granted = IsGranted(resolvedSheet, approvalClass);

                // 分流留痕：卡片为什么出现、为什么没出现，事后只有这一行能回答。
                // 工作表名可能是业务信息，隐私段已据此更新。
                Log.Info($"审批分流：{definition.Name} 类={approvalClass} 表=「{resolvedSheet}」 " +
                    $"策略={(_liveSettings?.Approval.ToString() ?? "?")} 已授权={granted}");

                if (!granted)
                {
                    Log.Info($"审批请求：{definition.Name}");
                    var decision = await requestApproval(definition, args, impact).ConfigureAwait(false);

                    if (!decision.Approved)
                    {
                        var reason = string.IsNullOrWhiteSpace(decision.Reason)
                            ? "用户拒绝了此操作。"
                            : $"用户拒绝了此操作：{decision.Reason}";

                        await FeedToolResultAsync(call, ToolResult.Failure("USER_REJECTED", reason), onUpdate)
                            .ConfigureAwait(false);
                        return;
                    }

                    // PerTurn 的第一次允许自动建立同表同类授权；PerWrite 只有用户点
                    // 「本轮同类允许」时才建立。结构必须显式点「含结构允许」。
                    if (settings.Approval == ApprovalPolicy.PerTurn
                        || decision.ApproveRest
                        || decision.ApproveStructureRest)
                    {
                        AddGrant(resolvedSheet, approvalClass);
                    }

                    // 结构是独立的一类。只有用户明确要求「含结构允许」时才会加，
                    // 格式/内容的 ApproveRest 永远不会让下一张建表卡片消失。
                    if (decision.ApproveStructureRest)
                    {
                        AddGrant(resolvedSheet, ApprovalClass.Structure);
                    }

                    await PublishApprovalGrantsAsync(onUpdate).ConfigureAwait(false);
                }
            }

            // 用工具调用标识作为撤销标识：面板拿到的 tool-result 里带同一个 id，
            // 因此能直接把撤销按钮挂到对应的操作卡片上。
            var result = await ExecuteOnUiAsync(call.Name, args, call.Id).ConfigureAwait(false);
            await FeedToolResultAsync(call, result, onUpdate).ConfigureAwait(false);
        }

        /// <summary>供测试核对工具分类。运行时不用它，只是把私有判据开一个只读入口。</summary>
        internal static ApprovalClass ClassOfToolForTest(string toolName)
        {
            return ApprovalClassFor(toolName);
        }

        /// <summary>供测试核对「不给撤销按钮时有没有说原因」。</summary>
        internal static string UndoNoteForTest(string toolName, ToolResult result, UndoRecord record)
        {
            return UndoNoteFor(toolName, result, record);
        }

        /// <summary>供测试核对「解析出的目标钉进了执行参数」。</summary>
        internal static void PinResolvedTargetForTest(
            ToolDefinition definition, JObject args, ImpactEstimate impact)
        {
            PinResolvedTarget(definition, args, impact);
        }

        /// <summary>供测试核对授权键的选取，尤其是没有工作表名的结构调用。</summary>
        internal static string GrantKeyForTest(
            ApprovalClass approvalClass, ImpactEstimate impact, JObject args)
        {
            return GrantKeyFor(approvalClass, impact, args);
        }

        private bool IsAutomaticApproval()
        {
            var live = _liveSettings ?? Settings.Load();
            return live.Approval == ApprovalPolicy.Automatic;
        }

        /// <summary>
        /// 把工具名归入授权的风险类。
        ///
        /// 清除、合并、排序刻意与写入分开：它们抹除或重排既有内容，而写入只是覆盖。
        /// 放在一类会让「往这几格写个值」的一次批准，顺带放行「清空整张表」——
        /// 用户看过的那一次和被放过去的那一次，破坏性根本不在一个量级。
        /// 合并尤其如此：它是唯一会静默丢值的写操作。
        ///
        /// 结构永远单列，不被任何其他授权捎带放行。
        /// </summary>
        private static ApprovalClass ApprovalClassFor(string toolName)
        {
            switch (toolName)
            {
                case "format_range":
                case "set_number_format":
                case "autofit_range":
                case "fit_range":
                    return ApprovalClass.Format;

                case "clear_range":
                case "merge_cells":
                case "unmerge_cells":
                case "sort_range":
                    return ApprovalClass.Destructive;

                case "add_worksheet":
                case "rename_worksheet":
                case "create_table":
                case "create_chart":
                case "excel_object":
                    return ApprovalClass.Structure;

                default:
                    return ApprovalClass.Write;
            }
        }

        /// <summary>
        /// 工作簿级授权的键。
        ///
        /// Excel 的工作表名不能含控制字符，所以这个键与任何真实表名都不会撞。
        /// </summary>
        private const string WorkbookWideGrantKey = "\u0000workbook";

        /// <summary>
        /// 把宿主解析出的目标钉进真正要执行的那份参数。
        ///
        /// 估算影响与执行是两次独立的 UI 线程往返，中间夹着等用户决定那段时间。
        /// 省略 sheet 时两次各自去取「当前活动表」——用户在批准期间切了表，
        /// 卡上说的是一张表，写入落到另一张。原先只有授权的键用了解析结果，
        /// 执行仍拿模型给的原始参数，这道防线只挡住了一半。
        /// </summary>
        private static void PinResolvedTarget(ToolDefinition definition, JObject args, ImpactEstimate impact)
        {
            if (impact == null)
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(impact.SheetName)
                && string.IsNullOrWhiteSpace(args.Value<string>("sheet")))
            {
                args["sheet"] = impact.SheetName;
            }

            // fit_range 省略 range 时，执行侧会再算一次已用范围。期间用户可能
            // 又填了几行，于是适配的范围与卡上说的不是同一片。
            if (!string.IsNullOrWhiteSpace(impact.Address)
                && string.IsNullOrWhiteSpace(args.Value<string>("range"))
                && string.Equals(definition.Name, "fit_range", StringComparison.Ordinal))
            {
                args["range"] = impact.Address;
            }
        }

        /// <summary>
        /// 这次调用该按什么键记授权。
        ///
        /// 结构类里 add_worksheet 与 rename_worksheet 压根没有 range 参数，
        /// 影响估算因此拿不到工作表名。原先的后果是：授权写不进去也查不出来，
        /// 用户点了「含结构允许」，芯片是空的，下一张建表卡照样弹——
        /// 而按钮的悬停说明承诺了相反的事。
        ///
        /// 这类调用改用工作簿级的键：结构操作本来就作用于整个工作簿，
        /// 按表记本身没有意义。
        /// </summary>
        private static string GrantKeyFor(ApprovalClass approvalClass, ImpactEstimate impact, JObject args)
        {
            if (!string.IsNullOrWhiteSpace(impact?.SheetName))
            {
                return impact.SheetName;
            }

            // 只对结构类兜底。范围类拿不到表名说明估算失败了，
            // 那时宁可每次都问，不要凭一个兜底键把整个工作簿都授权出去。
            return approvalClass == ApprovalClass.Structure ? WorkbookWideGrantKey : null;
        }

        private bool IsGranted(string sheetName, ApprovalClass approvalClass)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return false;
            }

            lock (_approvalGrants)
            foreach (var grant in _approvalGrants)
            {
                if (grant.Class == approvalClass
                    && string.Equals(grant.SheetName, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private void AddGrant(string sheetName, ApprovalClass approvalClass)
        {
            lock (_approvalGrants)
            {
            if (string.IsNullOrWhiteSpace(sheetName) || IsGranted(sheetName, approvalClass))
            {
                return;
            }

            _approvalGrants.Add(new ApprovalGrant { SheetName = sheetName, Class = approvalClass });
            }
        }

        /// <summary>把当前授权交给面板。芯片必须是文字，不让盾牌图标独自承担状态。</summary>
        private async Task PublishApprovalGrantsAsync(Func<AgentUpdate, Task> onUpdate)
        {
            var grants = new List<object>();
            lock (_approvalGrants)
            foreach (var grant in _approvalGrants)
            {
                var workbookWide = string.Equals(
                    grant.SheetName, WorkbookWideGrantKey, StringComparison.Ordinal);

                grants.Add(new
                {
                    // 工作簿级授权的键含控制字符，是内部标记，不能让面板原样显示。
                    sheet = workbookWide ? string.Empty : grant.SheetName,
                    workbookWide,
                    approvalClass = grant.Class.ToString(),
                });
            }

            await onUpdate(new AgentUpdate
            {
                Kind = "approval-grants",
                Payload = new { grants },
            }).ConfigureAwait(false);
        }

        /// <summary>
        /// 估算操作影响，用于审批卡片展示。
        /// 失败不阻塞审批：拿不到精确范围时给出可用的近似描述。
        /// </summary>
        private async Task<ImpactEstimate> DescribeImpactAsync(ToolDefinition definition, JObject args)
        {
            try
            {
                var range = args.Value<string>("range");

                // fit_range 允许省略 range。审批卡不能因此留一行空白的影响范围：
                // 用户看不出要动多少格，而它恰恰是整表操作。这里先解析成已用范围，
                // 与撤销登记前那段解析同一个来源。
                if (string.IsNullOrWhiteSpace(range)
                    && string.Equals(definition.Name, "fit_range", StringComparison.Ordinal))
                {
                    var resolved = await ExecuteOnUiAsync("get_workbook_info", new JObject()).ConfigureAwait(false);
                    range = ResolveUsedRangeForFit(args, resolved);
                }

                if (string.IsNullOrWhiteSpace(range))
                {
                    return new ImpactEstimate
                    {
                        Text = StructureImpactText(definition.Name, args),
                    };
                }

                var sheet = args.Value<string>("sheet");

                // 写公式要对照公式文本。默认读取只给 Value2，拿它作对照会把
                // =A1+1 显示成算出来的数字，批准的人以为在改值。
                var formulas = string.Equals(definition.Name, "write_formulas", StringComparison.Ordinal);
                var probe = await ExecuteOnUiAsync("read_range", new JObject
                {
                    ["range"] = range,
                    ["sheet"] = sheet,
                    ["include_formulas"] = formulas,
                }).ConfigureAwait(false);

                // 大范围会触发读取上限，此时退回用模型给的地址描述。
                // 它未经宿主规范化、也数不出单元格数，但仍能被面板译成行列说明。
                if (!probe.Ok)
                {
                    return new ImpactEstimate
                    {
                        SheetName = sheet,
                        Address = range,

                        // 读不到当前值也要出卡，并且说清是读不到——
                        // 拿一张空对照表假装读过，比没有预览更糟。
                        Preview = BuildPreview(definition, args, null, 0, unreadable: true),
                    };
                }

                var payload = JObject.FromObject(probe.Data);
                var rows = payload.Value<int?>("total_rows") ?? payload.Value<int?>("rows") ?? 0;
                var columns = payload.Value<int?>("total_columns") ?? payload.Value<int?>("columns") ?? 0;
                var currentMatrix = formulas
                    ? (payload["formulas"] ?? payload["values"])
                    : payload["values"];

                var resolvedSheet = payload.Value<string>("sheet") ?? sheet;
                var resolvedAddress = payload.Value<string>("address") ?? range;

                // 写公式对照公式文本——那是用户批准的对象。
                // 其余操作对照格子里显示出来的那一串：日期、时间、百分比、
                // 千分位用 Value2 会变成序列号或裸小数，摆在「将改为 2025-08-31」
                // 旁边看起来像换了一种东西，而不是改了一个值。
                if (!formulas)
                {
                    var display = await InvokeForWorkbookAsync(
                        () => _tools.ReadDisplayMatrix(
                            resolvedAddress,
                            resolvedSheet,
                            ChangePreviewBuilder.MaxRows,
                            ChangePreviewBuilder.MaxColumns)).ConfigureAwait(false)
                        as List<List<object>>;
                    if (display != null)
                    {
                        currentMatrix = JArray.FromObject(display);
                    }
                }
                var preview = BuildPreview(definition, args, currentMatrix,
                    (int)Math.Min(int.MaxValue, (long)rows * columns), unreadable: false);

                // 格式类操作的参数已经说清要改成什么，缺的是「现在是什么」。
                // 逐项都不一致时如实说一句，而不是把一份格式矩阵倒进卡片。
                if (preview == null && IsFormatTool(definition.Name))
                {
                    var mixed = (bool?)await InvokeForWorkbookAsync(
                        () => _tools.IsFormattingMixed(resolvedAddress, resolvedSheet)).ConfigureAwait(false);

                    if (mixed == true)
                    {
                        preview = new ChangePreview { FormattingMixed = true };
                    }
                }

                var note = RangeImpactNote(definition.Name, args, resolvedAddress, rows, columns);
                var countMeaningful = note == null
                    || (!string.Equals(definition.Name, "autofit_range", StringComparison.Ordinal)
                        && !string.Equals(definition.Name, "fit_range", StringComparison.Ordinal));

                return new ImpactEstimate
                {
                    SheetName = resolvedSheet,
                    Address = resolvedAddress,
                    CellCount = countMeaningful ? (long)rows * columns : (long?)null,
                    Preview = preview,
                    Note = note,
                };
            }
            catch (Exception ex)
            {
                Log.Warn("估算影响范围失败：" + ex.Message);
                return new ImpactEstimate();
            }
        }

        /// <summary>
        /// 没有范围的结构操作，审批卡上要说清的那一句。
        /// 新建工作表会成为活动表：之后省略 sheet 的调用都会落到它上面。
        /// </summary>
        private static string StructureImpactText(string toolName, JObject args)
        {
            switch (toolName)
            {
                case "add_worksheet":
                    return "将新增一张工作表，并成为当前活动表。之后未指定工作表的操作都会落到这张新表上。";
                case "rename_worksheet":
                    return "将重命名工作表。";
                case "create_chart":
                    return "将改变工作簿结构。";
                default:
                    return "将改变工作簿结构。";
            }
        }

        /// <summary>
        /// 有范围的操作里，单元格数说不清的那部分。
        /// 自动调整与适配动的是整列整行，不是所选那 N 个单元格。
        /// </summary>
        private static string RangeImpactNote(string toolName, JObject args, string address, int rows, int columns)
        {
            switch (toolName)
            {
                case "autofit_range":
                {
                    var target = args?.Value<string>("target");
                    if (string.Equals(target, "columns", StringComparison.OrdinalIgnoreCase))
                    {
                        return $"将调整所选范围内每一列的整列列宽，效果到达该列全部行，不限于这 {rows} 行。";
                    }

                    if (string.Equals(target, "rows", StringComparison.OrdinalIgnoreCase))
                    {
                        return $"将调整所选范围内每一行的整行行高，效果到达该行全部列，不限于这 {columns} 列。";
                    }

                    return "将按整列或整行自动调整尺寸，不限于所选单元格。";
                }

                case "fit_range":
                    return "将把这片范围水平与垂直居中，按整列自动调整列宽并把超过 100 的列限制为 100，启用自动换行后按整行调整行高。";

                case "create_table":
                {
                    var hasHeader = args?.Value<bool?>("has_header") ?? true;
                    return hasHeader
                        ? "将把首行当作标题（不再当数据）。撤销会去掉表格，但表格样式会留在格子上。"
                        : "将把范围转成表格。撤销会去掉表格，但表格样式会留在格子上。";
                }

                default:
                    return null;
            }
        }

        /// <summary>
        /// 只改外观、不改内容的那几个工具。
        ///
        /// 与审批授权的「格式类」是同一组工具，但刻意各自成表：这里回答的是
        /// 「该不该提一句当前格式统不统一」，那里回答的是「一次允许覆盖到哪」。
        /// 合成一处会让改动其中一个语义时静默改掉另一个。
        /// </summary>
        private static bool IsFormatTool(string name)
        {
            switch (name)
            {
                case "format_range":
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// 为写入类调用生成「原值 → 新值」的对照。
        ///
        /// 只覆盖写值与写公式：那两个的参数里带着完整的新内容，对照才有两边。
        /// 格式类操作的参数已经是可读的属性列表（加粗、填充色），审批卡照旧展示参数，
        /// 不为它们倒一份格式矩阵——那会把卡片撑成一张表，而参数本身已经说清要做什么。
        /// </summary>
        private static ChangePreview BuildPreview(
            ToolDefinition definition,
            JObject args,
            JToken currentMatrix,
            int totalCells,
            bool unreadable)
        {
            var isValues = string.Equals(definition.Name, "write_values", StringComparison.Ordinal);
            var isFormulas = string.Equals(definition.Name, "write_formulas", StringComparison.Ordinal);

            if (isValues || isFormulas)
            {
                var incoming = args?[isFormulas ? "formulas" : "values"];
                if (incoming == null)
                {
                    return null;
                }

                var preview = ChangePreviewBuilder.Build(
                    unreadable ? null : currentMatrix,
                    incoming,
                    totalCells,
                    formulas: isFormulas);

                if (preview != null && unreadable)
                {
                    preview.CurrentUnreadable = true;
                }

                return preview;
            }

            // 抹除类：要害不是「改成什么」而是「丢掉什么」。
            //
            // 合并是唯一会静默丢值的写操作——宿主只留左上角那一格，其余一概丢弃，
            // 事后没有痕迹可查。工具会把 discarded_values 回给模型，而在这之前
            // 用户要先点「允许」，所以那个数字必须先出现在卡上。
            // 清除同理：它抹掉的内容能撤，但撤销之前用户得知道抹的是什么。
            switch (definition.Name)
            {
                case "merge_cells":
                    return ChangePreviewBuilder.BuildDiscard(
                        unreadable ? null : currentMatrix, totalCells, keepAnchor: true, kind: "merge");

                case "clear_range":
                    // 只清格式时不丢值，列「会丢什么」是错的。
                    var scope = args?.Value<string>("scope");
                    if (string.Equals(scope, "formats", StringComparison.OrdinalIgnoreCase))
                    {
                        return null;
                    }

                    return ChangePreviewBuilder.BuildDiscard(
                        unreadable ? null : currentMatrix, totalCells, keepAnchor: false, kind: "clear");

                default:
                    return null;
            }
        }

        /// <summary>
        /// 取 fit_range 省略 range 时它实际会作用的已用范围地址。
        ///
        /// 审批卡不能因为参数里没有 range 就留一行空白的影响范围：那恰恰是整表操作，
        /// 用户最需要知道要动多少格。取不到时返回 null，卡片退回结构类兜底文案。
        /// </summary>
        private static string ResolveUsedRangeForFit(JObject args, ToolResult workbookInfo)
        {
            if (workbookInfo == null || !workbookInfo.Ok)
            {
                return null;
            }

            try
            {
                var payload = JObject.FromObject(workbookInfo.Data);
                var target = args.Value<string>("sheet");
                if (string.IsNullOrWhiteSpace(target))
                {
                    target = payload.Value<string>("active_sheet");
                }

                var sheets = payload["sheets"] as JArray;
                if (sheets == null || string.IsNullOrWhiteSpace(target))
                {
                    return null;
                }

                foreach (var entry in sheets)
                {
                    if (string.Equals(entry.Value<string>("name"), target, StringComparison.OrdinalIgnoreCase))
                    {
                        var used = entry.Value<string>("used_range");
                        return string.IsNullOrWhiteSpace(used) ? null : used;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Warn("解析适配的已用范围失败：" + ex.Message);
            }

            return null;
        }

        private async Task FeedToolResultAsync(ToolCall call, ToolResult result, Func<AgentUpdate, Task> onUpdate)
        {
            var json = JsonConvert.SerializeObject(result.ToPayload(), Formatting.None);
            var failedCall = result.Ok ? null : call.Name + "\n" + call.ArgumentsJson + "\n" + json;
            _sameFailureCount = failedCall != null && failedCall == _lastFailedCall ? _sameFailureCount + 1 : 1;
            _lastFailedCall = failedCall;

            // 文本协议下没有 tool_call_id 可用，结果只能作为 user 消息回灌。
            // 仍标记出工具结果身份，好让上下文压缩继续优先压它们。
            if (_toolMode == ToolProtocolMode.Text)
            {
                _conversation.Add(ChatMessage.FromTextProtocolToolResult(
                    call.Name,
                    $"（工具执行结果 · {call.Name}）\n{json}"));
            }
            else
            {
                _conversation.Add(ChatMessage.FromToolResult(call.Id, call.Name, json));
            }

            Log.Info(result.Ok
                ? $"工具 {call.Name} 执行成功"
                : $"工具 {call.Name} 执行失败：{result.ErrorCode} {result.Error}");

            // 告知面板这条操作能否撤销，以便决定是否显示撤销按钮。
            var undoRecord = _tools.Undo.Find(call.Id);

            await onUpdate(new AgentUpdate
            {
                Kind = "tool-result",
                Payload = new
                {
                    id = call.Id,
                    name = call.Name,
                    ok = result.Ok,
                    error = result.Error,
                    errorCode = result.ErrorCode,
                    data = result.Data,
                    canUndo = undoRecord?.CanUndo ?? false,
                    undoSummary = undoRecord?.Summary,

                    // 撤销之后还能不能恢复。图表删掉就重建不回来，
                    // 不带这一项面板会把按钮改成「恢复」，再点必然拿到 UNSUPPORTED。
                    canRedoAfterUndo = undoRecord?.Structure == null || undoRecord.Structure.CanRestore,

                    // 撤销能还原到什么程度。清除时格式逐项都不一致，
                    // 值找得回来、格式还原不了，这话要说在卡片上而不是藏在日志里。
                    undoNote = UndoNoteFor(call.Name, result, undoRecord),
                },
            }).ConfigureAwait(false);
            if (!result.Ok && _sameFailureCount >= 8)
            {
                throw new ProviderException("NO_PROGRESS", "同一工具和参数连续八次返回相同错误，已停止无进展重试；请按具体错误调整请求。");
            }
        }

        /// <summary>
        /// 撤销范围的补充说明。没有要说的就返回 null。
        ///
        /// 两种情形：撤销登记不成（写操作成功但没有记录，用户看到的是「少一个按钮」），
        /// 以及记录成立但某个维度还原不了。缺按钮不说原因会被当成故障，
        /// 而它其实是「保不住足以完整还原的快照就不承诺撤销」这一有意为之的取舍。
        /// </summary>
        private static string UndoNoteFor(string toolName, ToolResult result, UndoRecord record)
        {
            if (!result.Ok)
            {
                return null;
            }

            // 只读工具本来就不该有撤销入口，缺它不是需要解释的事。
            var definition = ToolCatalog.Find(toolName);
            if (definition == null || !definition.RequiresApproval)
            {
                return null;
            }

            // 记录成立：只剩「能还原到什么程度」要交代。
            if (record != null && record.CanUndo)
            {
                var before = record.Before;
                if (before == null)
                {
                    return null;
                }

                // 边框不在采集范围内（逐边读的 COM 成本过高），而清除格式会把它
                // 一并抹掉。撤销能还原九项，边框永久消失——这句话必须说在卡上。
                if (before.ClearsUncapturedFormats)
                {
                    // 纯文本：这条说明经 textContent 落到卡上，markdown 不会被渲染。
                    return "撤销能找回内容与字体、填充、对齐这些，但边框不会回来："
                        + "边框没有进快照，而清除格式已经把它抹掉了。需要的话请让我重新加边框。";
                }

                return before.FormatIncomplete
                    ? "撤销能找回内容，但格式只能部分还原：这片范围原本有几项外观逐格不同，"
                        + "宿主读不出统一值，那几项撤销后不会恢复原样。"
                    : null;
            }

            // 走到这里就是「操作成功了，但不提供撤销」。缺按钮本身是可见的，
            // 缺原因会被当成故障——而它其实是「保不住足以完整还原的依据，
            // 就不承诺可以撤销」这一有意为之的取舍。面板「适配」早就这样做了，
            // 模型发起的这条路上一直没做，两种新的不给按钮的情形都是静默的。
            switch (toolName)
            {
                case "excel_object":
                    return "通用表格对象操作不提供自动撤销，请通过读取目标状态确认结果。";
                case "create_chart":
                    return "这张图表不能撤销：宿主没有回报图表的名称，撤销时无法定位它。"
                        + "需要时请让我删掉重建。";

                case "create_table":
                case "add_worksheet":
                case "rename_worksheet":
                    return "这一步不能撤销：没能记下足以回退的信息。";

                case "format_range":
                case "set_number_format":
                    return "这次格式改动不能撤销：这片范围原本的外观逐项都不一样，"
                        + "宿主读不出统一值，保不住足以完整还原的快照。改动本身已经生效。";

                default:
                    return "这一步不能撤销：范围太大或原有状态过于参差，"
                        + "保不住足以完整还原的快照。改动本身已经生效。";
            }
        }

        private Task<object> RefreshSystemPromptAsync(Settings settings)
        {
            return InvokeForWorkbookAsync(() => { RefreshSystemPrompt(settings); return null; });
        }

        private void RefreshSystemPrompt(Settings settings)
        {
            WorkbookSummary summary = null;
            SelectionInfo selection = null;

            try
            {
                summary = _context.GetSummary();
                if (settings.AutoIncludeSelection)
                {
                    selection = _context.GetSelection();
                }
            }
            catch (Exception ex)
            {
                // 取不到工作簿信息时仍要能对话，系统提示里会写明这一点。
                Log.Warn("刷新工作簿上下文失败：" + ex.Message);
            }

            _conversation.SetSystemPrompt(
                SystemPrompt.Build(summary, selection, settings.AutoIncludeSelection, _toolMode));
        }

        private ChatRequest BuildRequest(ResolvedConnection connection, Settings settings)
        {
            var request = new ChatRequest
            {
                Protocol = connection.Protocol,
                BaseUrl = connection.BaseUrl,
                Token = connection.Token,
                Model = connection.Model,
                Thinking = settings.Thinking,
                Temperature = settings.Temperature,
                MaxOutputTokens = settings.MaxOutputTokens,
                // 只有原生形态才附带函数声明。文本协议把清单写进系统提示，
                // 顾问模式压根不给工具。
                IncludeTools = _toolMode == ToolProtocolMode.Native,
                // 这个模型被拒过 max_tokens 就改用 max_completion_tokens。
                OutputLimitOverride = _capability?.OutputLimit,
            };

            request.Messages.AddRange(_conversation.Messages);
            return request;
        }
    }
}
