using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ChatSheet.AddIn.Agent;
using ChatSheet.AddIn.Providers;
using ChatSheet.AddIn.Storage;
using ChatSheet.AddIn.Tools;
using Newtonsoft.Json.Linq;

namespace ChatSheet.AddIn.Bridge
{
    /// <summary>
    /// 面板与 Agent 之间的通道实现。
    ///
    /// 审批是这里最需要小心的部分：加载项发起审批请求后必须挂起等待
    /// 面板回传用户决定，用 TaskCompletionSource 按请求标识配对。
    /// 若面板在等待期间被关闭，必须让等待方以「拒绝」收束，
    /// 否则 Agent 会永久卡住。
    /// </summary>
    internal sealed class AgentChannels : IDisposable
    {
        private readonly AgentRunner _agent;
        private readonly Func<AgentUpdate, Task> _push;
        private readonly Func<object, Task> _pushRaw;

        /// <summary>
        /// 切到 UI 线程执行。撤销要访问宿主 COM 对象，
        /// 而通道回调运行在消息处理链上，未必是 UI 线程。
        /// </summary>
        private readonly Func<Func<object>, Task<object>> _uiInvoker;
        private readonly Func<Settings> _loadSettings;
        private readonly bool _wordHost;

        private readonly ConcurrentDictionary<string, TaskCompletionSource<ApprovalDecision>> _pendingApprovals =
            new ConcurrentDictionary<string, TaskCompletionSource<ApprovalDecision>>(StringComparer.Ordinal);

        private CancellationTokenSource _currentRun;

        /// <summary>
        /// 批量确认的取消源。
        ///
        /// 与 _currentRun 分开是硬要求：Stop() Cancel 的是 _currentRun 这一个对话槽位，
        /// 复用它会让「停止」在批量与对话之间产生歧义——那正是「发消息不再误停当前
        /// 任务」已经付过代价的故障。停批量不许碰对话，反之同理。
        /// </summary>
        private CancellationTokenSource _currentBulkProbe;

        private int _approvalSequence;
        private Settings _settings;
        private sealed class WorkBuddyProbeState
        {
            internal CancellationTokenSource Cancellation;
            internal Task<WorkBuddyModelsResult> Task;
        }

        private readonly object _workBuddyProbeLock = new object();
        private readonly Dictionary<ConnectionMode, WorkBuddyProbeState> _workBuddyProbes =
            new Dictionary<ConnectionMode, WorkBuddyProbeState>();
        private readonly Dictionary<ConnectionMode, WorkBuddyModelsResult> _workBuddyAuthorizations =
            new Dictionary<ConnectionMode, WorkBuddyModelsResult>();
        private readonly Dictionary<ConnectionMode, DateTime> _workBuddyAuthorizationAtUtc =
            new Dictionary<ConnectionMode, DateTime>();
        private readonly CancellationTokenSource _workBuddyLifetime = new CancellationTokenSource();
        private sealed class WorkBuddyActionState
        {
            internal CancellationTokenSource Cancellation;
            internal string Id;
            internal ConnectionMode Mode;
            internal string AuthUrl;
        }
        private WorkBuddyActionState _workBuddyAction;
        private int _disposed;

        internal AgentChannels(
            Func<object> applicationAccessor,
            Func<AgentUpdate, Task> push,
            Func<object, Task> pushRaw,
            Func<Func<object>, Task<object>> uiInvoker,
            Func<Settings> loadSettings = null,
            bool createAgent = true,
            bool wordHost = false)
        {
            if (applicationAccessor == null)
            {
                throw new ArgumentNullException(nameof(applicationAccessor));
            }

            // Word 只复用这里的 WorkBuddy/模型通道；它有自己的文档 Agent，
            // 不需要再构造一份带 Excel 工具的 AgentRunner。
            _agent = createAgent ? new AgentRunner(applicationAccessor, uiInvoker) : null;
            _push = push;
            _pushRaw = pushRaw;
            _uiInvoker = uiInvoker ?? (work => Task.FromResult(work()));
            _loadSettings = loadSettings ?? Settings.Load;
            _wordHost = wordHost;
            _settings = _loadSettings();
        }

        internal void Register(IDictionary<string, Func<JObject, Task<object>>> handlers)
        {
            handlers["settings.get"] = GetSettingsAsync;
            handlers["workbuddy.runtime"] = payload =>
            {
                var mode = WorkBuddyMode(payload);
                return Task.FromResult<object>(mode.IsWorkBuddy() ? WorkBuddyRuntime.Status(mode) : null);
            };
            handlers["workbuddy.account"] = async payload =>
            {
                var mode = WorkBuddyMode(payload);
                return mode.IsDomesticWorkBuddy()
                    ? await WorkBuddyAccount.RefreshAsync(mode, payload.Value<bool?>("force") == true, _workBuddyLifetime.Token).ConfigureAwait(false)
                    : null;
            };
            handlers["workbuddy.open-auth-url"] = payload =>
            {
                var mode = WorkBuddyMode(payload);
                var action = _workBuddyAction;
                var opened = action != null && action.Mode == mode && !action.Cancellation.IsCancellationRequested &&
                    action.Id == payload.Value<string>("operationId") &&
                    WorkBuddyAccountConnection.TryOpenAuthUrl(action.AuthUrl, mode);
                return Task.FromResult<object>(new
                {
                    ok = opened,
                    detail = opened ? "登录页面已重新打开。" : "登录页面无法打开，请检查默认浏览器设置后重试。",
                });
            };
            handlers["workbuddy.login"] = payload => RunWorkBuddyActionAsync(false, WorkBuddyMode(payload),
                payload.Value<string>("operationId"), payload.Value<bool?>("switchAccount") == true);
            handlers["workbuddy.install"] = payload => RunWorkBuddyActionAsync(true, WorkBuddyMode(payload), payload.Value<string>("operationId"));
            handlers["workbuddy.cancel"] = payload =>
            {
                var action = _workBuddyAction;
                try { if (action?.Id == payload.Value<string>("operationId")) { action?.Cancellation.Cancel(); } }
                catch (ObjectDisposedException) { }
                return Task.FromResult<object>(new { ok = true });
            };

            // 卡片上的范围不是死文字：用户要在允许之前亲眼看看那几格。
            // 必须经 UI 线程访问宿主 COM；解析失败返回现有 RangeResolver 错误，
            // 不静默跳到当前表的什么位置。
            handlers["sheet.goto"] = async payload =>
            {
                var address = payload.Value<string>("address");
                var sheet = payload.Value<string>("sheet");
                if (string.IsNullOrWhiteSpace(address))
                {
                    return new { ok = false, message = "缺少范围地址" };
                }

                try
                {
                    var result = (ToolResult)await _uiInvoker(
                        () => _agent.Tools.GotoRange(address, sheet)).ConfigureAwait(false);

                    return result.Ok
                        ? new { ok = true, data = result.Data }
                        : new { ok = false, message = result.Error, errorCode = result.ErrorCode };
                }
                catch (ToolException ex)
                {
                    return new { ok = false, message = ex.Message, errorCode = ex.Code };
                }
                catch (Exception ex)
                {
                    Log.Error("跳转范围失败", ex);
                    return new { ok = false, message = "无法跳转到该范围：" + ex.Message };
                }
            };

            // 图片能力约束下发给面板，避免前端与后端各写一套上限。
            handlers["image.limits"] = _ => Task.FromResult<object>(new
            {
                maxCount = ImageSupport.MaxImagesPerTurn,
                maxBytes = ImageSupport.MaxBytesPerImage,
                mediaTypes = ImageSupport.SupportedMediaTypes,
            });

            // 文件附件的约束同理。面板据此在拖入时就给出拒绝原因，
            // 不必先发一轮再被这边退回。
            handlers["file.limits"] = _ => Task.FromResult<object>(new
            {
                maxCount = FileSupport.MaxFilesPerTurn,
                maxBytes = FileSupport.MaxBytesPerFile,
                maxTotalBytes = FileSupport.MaxTotalBytes,
                extensions = FileSupport.SupportedExtensions,
            });
            handlers["settings.save"] = SaveSettingsAsync;
            handlers["cli.probe"] = _ => Task.FromResult(ProbeCliPayload());
            handlers["models.list"] = ListModelsAsync;
            handlers["proxy.test"] = TestProxyAsync;
            handlers["chat.send"] = SendAsync;
            handlers["chat.stop"] = _ => Task.FromResult(Stop());
            handlers["chat.reset"] = _ =>
            {
                _agent.Reset();
                return Task.FromResult<object>(new { reset = true });
            };
            handlers["approval.respond"] = RespondApprovalAsync;

            // 收回本轮授权。不停止本轮——用户想停的是自动放行，不是整个任务。
            handlers["approval.revoke"] = _ =>
            {
                var count = _agent.RevokeApprovalGrants();
                Log.Info($"用户收回本轮授权，清掉 {count} 条");
                return Task.FromResult<object>(new { ok = true, revoked = count });
            };

            // 对话页的快捷切换：模型、思考档位、审批策略。
            // 这三项每次任务都可能调整，走完整的设置保存太重。
            handlers["session.update"] = payload =>
            {
                // 改正在跑的那一份设置对象，不要 Load 一份新的再替换引用。
                // AgentRunner 开轮时拿到的就是 _settings，换引用的话轮内切策略
                // 永远进不了正在跑的那一轮——规范要求下一刀写操作按新策略走。
                var settings = _settings ?? Settings.Load();
                var changed = new List<string>();

                var model = payload.Value<string>("model");
                if (model != null && !string.Equals(settings.Model, model.Trim(), StringComparison.Ordinal))
                {
                    settings.Model = model.Trim();
                    // 对话页的模型列表就是当前连接拉来的，选中即属于当前连接。
                    settings.StampModelConnection();
                    changed.Add("模型=" + settings.Model);
                }

                if (Thinking.TryParse(payload.Value<string>("thinking"), out var level) &&
                    settings.Thinking != level)
                {
                    settings.Thinking = level;
                    changed.Add("思考档位=" + level);
                }

                if (Enum.TryParse(payload.Value<string>("approval"), out ApprovalPolicy policy) &&
                    settings.Approval != policy)
                {
                    settings.Approval = policy;
                    changed.Add("审批策略=" + policy);
                }

                // 「只看名单」开关走这条通道而不是 settings.save：选择器在对话页，
                // 而 settings.save 发的是设置页那份 current 全量快照，且 initSettings
                // 每个面板生命周期只跑一次。让开关走那条路，用户在选择器里拨完再去
                // 设置页点保存，就会把面板启动时的旧值写回来。
                var onlyFavorites = payload.Value<bool?>("onlyFavoriteModels");
                if (onlyFavorites.HasValue && settings.OnlyFavoriteModels != onlyFavorites.Value)
                {
                    settings.OnlyFavoriteModels = onlyFavorites.Value;
                    changed.Add("只看名单=" + onlyFavorites.Value);
                }

                if (changed.Count > 0)
                {
                    settings.Save();
                    _settings = settings;
                    Log.Info("对话页快捷调整：" + string.Join("，", changed));
                }

                return Task.FromResult<object>(new
                {
                    model = settings.Model,
                    thinking = settings.Thinking.ToString(),
                    approval = settings.Approval.ToString(),
                    onlyFavoriteModels = settings.OnlyFavoriteModels,
                    thinkingSupported = Thinking.SupportedLevels(EffectiveProtocol()),
                });
            };

            handlers["models.probe"] = ProbeModelAsync;
            handlers["models.probe.bulk"] = ProbeFavoritesAsync;
            // 批量测试整份目录。与 models.probe.bulk 分开注册而不是加个参数：
            // 二者的作用范围与代价差一个数量级（名单几个 vs 目录几十个），
            // 而「几十次计费请求」这件事必须在调用点就看得见，不该藏在一个布尔里。
            //
            // 目标数（stopAfterAvailable）反过来走同一条 channel：它不改作用范围也不改
            // 代价模型——仍然是「对整份目录的那一次运行」，只是给了个停止规则，
            // 而条数仍明写在 payload 与按钮上。真开第三条 channel 要把守卫、去重、
            // 四处推送、收尾回复整段复制，而 StopBulkProbe 只有一个 _currentBulkProbe，
            // 两条路都往里塞会让「现在跑的是哪一批」重新变成猜。
            handlers["models.test.all"] = TestAllModelsAsync;
            handlers["models.probe.stop"] = _ => Task.FromResult(StopBulkProbe());

            // 常用名单的读写。名单是用户意图，落盘；判定是外部事实，只在内存里。
            handlers["models.favorites"] = payload =>
            {
                var settings = Settings.Load();
                var key = settings.FavoritesKey();
                var action = payload.Value<string>("action") ?? "get";
                var model = payload.Value<string>("model");

                switch (action)
                {
                    case "toggle":
                        FavoriteModels.Toggle(key, model);
                        break;
                    case "add":
                        // 手填的 ID 自动进名单：肯花力气打出来的 ID 就是要用的。
                        FavoriteModels.Add(key, model);
                        break;
                }

                return Task.FromResult<object>(new
                {
                    favorites = FavoriteModels.Load(key),
                    availability = AvailabilityPayload(settings),
                });
            };

            // 面板初次渲染时需要主动取一次上下文占用，
            // 否则进度圆环要等到第一轮对话产生推送后才有数值。
            handlers["context.state"] = _ =>
            {
                var used = _agent.Conversation.EstimateTotalTokens();
                var budget = Math.Max(1, _settings.ContextBudgetTokens);
                var ratio = Math.Min(1.0, (double)used / budget);

                return Task.FromResult<object>(new
                {
                    used,
                    budget,
                    ratio,
                    percent = (int)Math.Round(ratio * 100),
                    threshold = (int)(Conversation.CompressionThreshold * 100),
                    nearLimit = ratio >= Conversation.CompressionThreshold,
                });
            };

            // 「适配」按钮：把活动表的已用范围整片排好，不经过模型。
            //
            // 不走对话是刻意的：这是个确定性的排版动作，用户点按钮就是已经表达了
            // 意图，再让模型转述一遍只会增加延迟、token 开销和被误解的可能。
            // 但仍登记撤销记录并回传标识，面板据此给出撤销入口——
            // 加载项通过 COM 的写入会清空 Excel 自身的撤销栈，Ctrl+Z 救不回来。
            //
            // 不传 range：由 fit_range 自己取已用范围，省一次跨线程往返，
            // 也让「适配到哪」这个判断只存在一处。
            handlers["sheet.fit"] = async payload =>
            {
                var undoId = "fit-" + Guid.NewGuid().ToString("N").Substring(0, 8);
                var args = new JObject();

                // 水平对齐由面板给。缺省交给 fit_range 兜底成 center，
                // 这样「默认居中」只在一处定义。
                var alignment = payload.Value<string>("horizontalAlignment");
                if (!string.IsNullOrWhiteSpace(alignment))
                {
                    args["horizontal_alignment"] = alignment.Trim();
                }

                var result = (ToolResult)await _uiInvoker(
                    () => _agent.Tools.Execute("fit_range", args, undoId)).ConfigureAwait(false);

                if (!result.Ok)
                {
                    Log.Warn($"适配失败：{result.ErrorCode} {result.Error}");
                    return new { ok = false, message = result.Error };
                }

                var data = JObject.FromObject(result.Data);

                // 只有确实登记成了记录才回传标识。
                //
                // 这里曾经无条件回传：快照采集失败时没有记录，面板却照样显示
                // 撤销按钮，点下去只能得到「找不到该操作记录」。宁可不给按钮，
                // 也不能给一个注定失败的按钮——后者会让用户以为改动可以回退。
                var undoRecord = _agent.Tools.Undo.Find(undoId);
                var canUndo = undoRecord?.CanUndo == true;

                Log.Info($"适配 {data.Value<string>("address")}：{data.Value<int>("cells_affected")} 个单元格" +
                    (canUndo ? string.Empty : "（未登记撤销记录）"));

                return new
                {
                    ok = true,
                    undoId = canUndo ? undoId : null,
                    // 没有撤销入口时说明原因。缺按钮本身是可见的，缺原因则会
                    // 被当成故障——而它其实是「保不住足以完整还原的快照，
                    // 那就不承诺可以撤销」这一有意为之的取舍。
                    //
                    // 只说事实加最常见的成因，不逐一枚举：采集失败有几种情形
                    // （范围行列数过大、原排版逐格各异且单元格过多、宿主读取失败），
                    // 把三种都摆给用户并不能帮他做任何决定，具体原因记在日志里。
                    undoUnavailableReason = canUndo
                        ? null
                        : "这次适配不能撤销：范围太大，保不住足以完整还原的排版快照" +
                            "（原本的对齐逐格不同时尤其容易触发）。适配本身已经生效。",
                    address = data.Value<string>("address"),
                    sheet = data.Value<string>("sheet"),
                    rows = data.Value<int>("rows_adjusted"),
                    columns = data.Value<int>("columns_adjusted"),
                    horizontalAlignment = data.Value<string>("horizontal_alignment"),
                };
            };

            // 撤销与恢复。必须切到 UI 线程：还原要访问宿主 COM 对象。
            handlers["undo.apply"] = async payload =>
            {
                var id = payload.Value<string>("id");
                var redo = payload.Value<bool?>("redo") ?? false;

                // 用户已经看过重叠警告并选择继续。只对撤销有意义：
                // 恢复走的是「后快照」，不存在盖掉更晚改动的问题。
                var force = payload.Value<bool?>("force") ?? false;

                if (string.IsNullOrEmpty(id))
                {
                    return new { ok = false, message = "缺少操作标识" };
                }

                var outcome = (UndoOutcome)await _uiInvoker(
                    () => redo ? _agent.Tools.Undo.Redo(id) : _agent.Tools.Undo.Undo(id, force)).ConfigureAwait(false);

                Log.Info($"{(redo ? "恢复" : "撤销")}操作 {id}：{(outcome.Ok ? "成功" : "失败 " + outcome.ErrorCode)} {outcome.Message}");

                return new
                {
                    ok = outcome.Ok,
                    message = outcome.Message,
                    errorCode = outcome.ErrorCode,
                    undone = outcome.Undone,
                };
            };

            // 手动压缩：圆环到达阈值后由用户决定是否立即压缩。
            handlers["context.compact"] = _ =>
            {
                var settings = Settings.Load();
                var trim = _agent.Conversation.TrimToBudget(settings.ContextBudgetTokens);
                Log.Info($"手动压缩上下文：{trim.TokensBefore} → {trim.TokensAfter} tokens");

                return Task.FromResult<object>(new
                {
                    trimmed = trim.Trimmed,
                    before = trim.TokensBefore,
                    after = trim.TokensAfter,
                    compressed = trim.CompressedToolResults,
                    dropped = trim.DroppedMessages,
                    budget = trim.BudgetTokens,
                });
            };
        }

        /// <summary>
        /// 当前连接下已有判定的模型，键是模型 ID，值是三态之一。
        ///
        /// 只下发已有判定的：没判定就是「未确认」，让面板自己把缺席渲染成那个状态，
        /// 省掉为整份目录逐个下发 Unknown。
        /// </summary>
        private static object AvailabilityPayload(Settings settings)
        {
            var payload = new JObject();
            foreach (var pair in ModelAvailability.SnapshotFor(settings.ConnectionKey()))
            {
                payload[pair.Key] = pair.Value.ToString();
            }

            return payload;
        }

        /// <summary>
        /// 确认一个模型。
        ///
        /// 只走已保存设置，不接受面板传来的候选配置。ListModelsAsync 刻意接受未保存值
        /// 让设置页能试连，但那条路对探测不成立：判定的键取自已保存的连接，
        /// 拿候选网关探出来的结论会盖到用户当前正在用的连接上。
        /// </summary>
        private async Task<object> ProbeModelAsync(JObject payload)
        {
            var model = (payload.Value<string>("model") ?? string.Empty).Trim();
            if (model.Length == 0)
            {
                throw new ProviderException("MODEL_REQUIRED", "没有指定要确认的模型。");
            }

            // 对话在飞时不许探测。SendAsync 的 BUSY 守卫只在它自己内部，
            // 这条通道继承不到——而几条请求压在同一个账号上会招限流，
            // 限流判未知等于花了钱没答案，还可能把用户那一轮带着上下文的请求限掉。
            if (_currentRun != null)
            {
                throw new ProviderException(
                    "BUSY", "正在对话中，等这一轮结束再确认模型。");
            }

            var settings = Settings.Load();
            var connection = settings.ResolveConnection();
            var connectionKey = settings.ConnectionKey();
            var capability = ModelCapabilities.For(connectionKey, model);

            var verdict = await ModelProbe.ProbeAsync(
                connection, model, capability.OutputLimit, _workBuddyLifetime.Token)
                .ConfigureAwait(false);

            ModelAvailability.Record(connectionKey, model, verdict);
            Log.Info($"确认模型 {model}：{verdict}");

            return new
            {
                model,
                verdict = verdict.ToString(),
                availability = AvailabilityPayload(settings),
            };
        }

        /// <summary>
        /// 把名单里的模型逐个确认完。串行、带进度、可中断，已得结果保留。
        ///
        /// 只对名单提供批量，不对完整目录：几十个 ID 就是几十次付费请求，
        /// 而那正是用户想避开的开销。
        /// </summary>
        private async Task<object> ProbeFavoritesAsync(JObject payload)
        {
            if (_currentRun != null)
            {
                throw new ProviderException(
                    "BUSY", "正在对话中，等这一轮结束再确认模型。");
            }

            if (_currentBulkProbe != null)
            {
                throw new ProviderException("BUSY", "已经在确认名单了。");
            }

            var settings = Settings.Load();
            var connection = settings.ResolveConnection();
            var connectionKey = settings.ConnectionKey();
            var models = FavoriteModels.Load(settings.FavoritesKey());

            if (models.Count == 0)
            {
                return new { confirmed = 0, total = 0, stopped = false, availability = AvailabilityPayload(settings) };
            }

            var cts = CancellationTokenSource.CreateLinkedTokenSource(_workBuddyLifetime.Token);
            _currentBulkProbe = cts;

            var confirmed = 0;
            var stopped = false;

            try
            {
                for (var i = 0; i < models.Count; i++)
                {
                    if (cts.IsCancellationRequested)
                    {
                        stopped = true;
                        break;
                    }

                    var model = models[i];

                    await _pushRaw(new
                    {
                        kind = "probe-progress",
                        model,
                        index = i + 1,
                        total = models.Count,
                        starting = true,
                    }).ConfigureAwait(false);

                    AvailabilityVerdict? settled = null;
                    try
                    {
                        var verdict = await ModelProbe.ProbeAsync(
                            connection, model, ModelCapabilities.For(connectionKey, model).OutputLimit, cts.Token)
                            .ConfigureAwait(false);

                        ModelAvailability.Record(connectionKey, model, verdict);
                        settled = verdict;
                        confirmed++;
                    }
                    catch (OperationCanceledException)
                    {
                        // 用户停了。已得结果保留，剩下的不发。
                        stopped = true;
                        break;
                    }
                    catch (ProviderException ex)
                    {
                        // 单个失败不该中断整批：它本身就是一条判定。
                        var verdict = ModelAvailability.Classify(ex, model);
                        ModelAvailability.Record(connectionKey, model, verdict);
                        settled = verdict;
                        confirmed++;
                    }

                    // 探完这一个就把判定推出去，让这一行当场变绿或变红。
                    //
                    // 不推的话，颜色只能等整批结束时由回复里的 availability 一次性补上，
                    // 而名单有十几个模型就要跑十几次往返——中途一列全是「未确认」，
                    // 用户看到的就是「批量探测点了没反应，失败的没变红、成功的没标绿」。
                    // 批量测试那条路（TestAllModelsAsync）一直是这么推的，这里是漏的。
                    //
                    // settled 标住「这一个已经有结论」：面板据此把扫光从这一行摘掉，
                    // 否则已经上色的行会继续挂着「正在测」的高光。
                    if (settled.HasValue)
                    {
                        await _pushRaw(new
                        {
                            kind = "probe-progress",
                            model,
                            index = i + 1,
                            total = models.Count,
                            verdict = settled.Value.ToString(),
                            settled = true,
                        }).ConfigureAwait(false);
                    }
                }
            }
            finally
            {
                _currentBulkProbe = null;
                cts.Dispose();
            }

            Log.Info($"批量确认结束：已确认 {confirmed}/{models.Count}" + (stopped ? "（用户中止）" : string.Empty));

            await _pushRaw(new
            {
                kind = "probe-progress",
                model = string.Empty,
                index = confirmed,
                total = models.Count,
                done = true,
            }).ConfigureAwait(false);

            return new
            {
                confirmed,
                total = models.Count,
                stopped,
                availability = AvailabilityPayload(settings),
            };
        }

        /// <summary>
        /// 测试整份目录：每个模型一条最小请求，按给定并发跑。
        ///
        /// 与 ProbeFavoritesAsync 的区别不只是范围：
        ///   · 范围是面板传来的整份目录，不是落盘的名单。目录来自 GET /models，
        ///     后端并不留存，所以由面板把它带过来——它决定探哪些模型，
        ///     而这是用户点按钮时就已经表达的意图。
        ///   · 并发跑（默认 5），不是串行。并发确实会招限流，而限流判「未确认」；
        ///     这个代价由面板写在按钮上，让用户点之前就知道会发多少条。
        ///
        /// 停止走同一个 models.probe.stop：对用户而言「停下正在跑的那一批」只有一个意思，
        /// 两个停止通道反而要先想清楚停的是哪一批。
        /// </summary>
        private async Task<object> TestAllModelsAsync(JObject payload)
        {
            if (_currentRun != null)
            {
                throw new ProviderException(
                    "BUSY", "正在对话中，等这一轮结束再测试模型。");
            }

            if (_currentBulkProbe != null)
            {
                throw new ProviderException("BUSY", "已经在测试了。");
            }

            var settings = Settings.Load();
            var connection = settings.ResolveConnection();
            var connectionKey = settings.ConnectionKey();

            var models = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var token in payload?["models"] as JArray ?? new JArray())
            {
                var id = (token?.ToString() ?? string.Empty).Trim();
                if (id.Length > 0 && seen.Add(id))
                {
                    models.Add(id);
                }
            }

            if (models.Count == 0)
            {
                // 形状与正常路径保持一致：面板拿到的字段集合不该因为「目录是空的」
                // 而少几个，否则它得为这一种情形单独写一套读法。
                return new
                {
                    confirmed = 0,
                    total = 0,
                    stopped = false,
                    target = 0,
                    availableFound = 0,
                    attempted = 0,
                    targetMet = false,
                    outcome = "completed",
                    availability = AvailabilityPayload(settings),
                };
            }

            var concurrency = (int?)payload?["concurrency"] ?? 5;

            // 目标数：探出这么多个可用的就不再派发。0 表示测完整份目录。
            //
            // 不用裸 (int?) 强转：那会让一个非数字的值走 Convert 抛 FormatException，
            // 于是整次运行一条请求都不发，而面板只看到一句「批量测试失败」，
            // 看不出是参数问题。上面 concurrency 那行就是这个写法，不是照抄的理由。
            var stopAfterAvailable = 0;
            var targetToken = payload?["stopAfterAvailable"];
            if (targetToken != null &&
                (targetToken.Type == JTokenType.Integer || targetToken.Type == JTokenType.Float))
            {
                stopAfterAvailable = (int)targetToken;
            }

            // 目标数不小于候选数时等价于全部测试。收拢成 0 才能让下面的结局不谎报
            // 「达标」、日志不写出「剩余 0 个未发请求」。
            if (stopAfterAvailable < 0 || stopAfterAvailable >= models.Count)
            {
                stopAfterAvailable = 0;
            }

            var cts = CancellationTokenSource.CreateLinkedTokenSource(_workBuddyLifetime.Token);
            _currentBulkProbe = cts;

            var stopped = false;
            var confirmed = 0;

            // 实发条数与可用个数都记在这里，不经 ProbeManyAsync 的返回值带回来：
            // 用户取消那条路是抛出而不是返回的，结构体在那条路上永远拿不到值，
            // 而那正是唯一需要「发了几条却没拿到判定」这个差额的地方。
            // attempted 记在 onStart 里——拿到槽位之后、发请求之前，
            // 那是「这一条真的要计费了」最近的证据。
            var attempted = 0;
            var available = 0;
            var outcome = default(ProbeSweepOutcome);

            try
            {
                // target 必须出现在这一轮的每一条推送上。面板只对 index 与 total
                // 做了「沿用上一条」的兜底，缺字段的推送会让列头文字在 starting 与
                // settled 交替时忽明忽暗（一会儿按目标算、一会儿按目录条数算）。
                await _pushRaw(new
                {
                    kind = "probe-progress",
                    model = string.Empty,
                    index = 0,
                    total = models.Count,
                    target = stopAfterAvailable,
                    availableFound = 0,
                }).ConfigureAwait(false);

                outcome = await ModelProbe.ProbeManyAsync(
                    connection,
                    models,
                    concurrency,
                    model => ModelCapabilities.For(connectionKey, model).OutputLimit,
                    async (model, verdict, completed) =>
                    {
                        ModelAvailability.Record(connectionKey, model, verdict);
                        Interlocked.Increment(ref confirmed);

                        // 可用个数必须单独推给面板，不能让它从 confirmed 猜：
                        // Record 遇 Unknown 直接 return，所以一批全被限流时
                        // confirmed 会是十几，而列表一行都没上色。
                        if (verdict == AvailabilityVerdict.Available)
                        {
                            Interlocked.Increment(ref available);
                        }

                        // 每探完一个就推一次：几十个模型跑下来要一阵，
                        // 只在结束时推一次的话中途看起来像卡住。
                        //
                        // settled 标住「这一个有结论了」，面板据此把它从「正在测」里摘掉。
                        // 少了这个字段，扫光会留在一行刚刚变绿或变红的行上——已经有结论
                        // 却还挂着「正在测」的高光，而真正在飞的那几个一个都没标。
                        await _pushRaw(new
                        {
                            kind = "probe-progress",
                            model,
                            index = completed,
                            total = models.Count,
                            verdict = verdict.ToString(),
                            settled = true,
                            target = stopAfterAvailable,
                            availableFound = Volatile.Read(ref available),
                        }).ConfigureAwait(false);
                    },
                    cts.Token,
                    // 某个模型真的开始探时推一条不带判定的：并发 5 就是同时五行在测，
                    // 而「在飞的是哪几个」只有这里说得出来。不带 index——已完成数由
                    // onResult 那条推进，这条只负责把这一行标成正在测。
                    async model =>
                    {
                        // 这一条是「真的要发这条请求了」最近的证据，实发条数记在这里。
                        Interlocked.Increment(ref attempted);

                        await _pushRaw(new
                        {
                            kind = "probe-progress",
                            model,
                            total = models.Count,
                            starting = true,
                            target = stopAfterAvailable,
                            availableFound = Volatile.Read(ref available),
                        }).ConfigureAwait(false);
                    },
                    stopAfterAvailable).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                // 用户停了。已得判定保留——那些请求已经付过钱了。
                //
                // 达标停**不**走这条路：它是「停止派发 + 照常等在飞的跑完」，正常返回。
                // 于是这个布尔仍然只表示用户中止一件事，两种结局在类型层面就分开了。
                // 别为了「统一」把达标也改成 Cancel——那会丢掉在飞那几条已付费的判定。
                stopped = true;
            }
            finally
            {
                _currentBulkProbe = null;
                cts.Dispose();
            }

            // 三种结局，用户中止优先：达标之后要排空在飞的那几条（每条最长 15 秒截止
            // 时间），这段窗口里用户完全能点停止，那时两者同时成立。必须报用户中止,
            // 因为只有它意味着有已付费的判定被丢了。
            //
            // 「达标」这一档额外要求确实还有没发的：最后一个候选刚好凑够目标时
            // 一条都没省下，报「剩余 0 个未发请求」读起来像有 bug。
            var targetMet = !stopped && outcome.TargetMet && attempted < models.Count;
            var reason = stopped ? "stopped" : (targetMet ? "target" : "completed");

            Log.Info(
                $"批量测试结束：已测 {confirmed}/{models.Count}，实发 {attempted} 条，" +
                $"可用 {available} 个，并发 {concurrency}" +
                (stopped
                    ? "（用户中止）"
                    : (targetMet
                        ? $"（达到目标 {stopAfterAvailable} 个可用，剩余 {models.Count - attempted} 个未发请求）"
                        : string.Empty)));

            await _pushRaw(new
            {
                kind = "probe-progress",
                model = string.Empty,
                index = confirmed,
                total = models.Count,
                done = true,
                target = stopAfterAvailable,
                availableFound = available,
                attempted,
                targetMet,
                outcome = reason,
            }).ConfigureAwait(false);

            return new
            {
                confirmed,
                total = models.Count,
                stopped,
                target = stopAfterAvailable,
                availableFound = available,
                attempted,
                targetMet,
                outcome = reason,
                availability = AvailabilityPayload(settings),
            };
        }

        /// <summary>
        /// 停批量。刻意不碰 _currentRun——停批量与停对话是两件事。
        /// </summary>
        private object StopBulkProbe()
        {
            var running = _currentBulkProbe;
            if (running == null)
            {
                return new { stopped = false, reason = "没有正在进行的批量确认" };
            }

            try
            {
                running.Cancel();
            }
            catch (ObjectDisposedException)
            {
                // 刚好自己结束了，等价于已停。
            }

            Log.Info("用户中止批量确认");
            return new { stopped = true };
        }

        private ConnectionMode WorkBuddyMode(JObject payload)
        {
            if (Enum.TryParse(payload?.Value<string>("mode"), out ConnectionMode requested) && requested.IsWorkBuddy())
            {
                return requested;
            }

            return _settings.Mode;
        }

        private Task<object> GetSettingsAsync(JObject payload)
        {
            // ACP 查询会跨线程、跨多个 await。必须把本次请求的设置冻结下来，
            // 否则用户在等待期间切换模式后，旧模式的授权目录会被新模式包装返回。
            var settings = CloneSettings(_settings);
            WorkBuddyModelsResult authorization = null;
            if (settings.Mode.IsWorkBuddy() &&
                _workBuddyAuthorizations.TryGetValue(settings.Mode, out var cachedAuthorization))
            {
                // settings.get 只返回已缓存的授权快照。真正的 ACP 探测由设置页
                // 的连接刷新显式触发，避免聊天页启动时被 WorkBuddy 拖住。
                authorization = cachedAuthorization;
            }
            return Task.FromResult(GetSettingsPayload(settings, authorization));
        }

        private Task<WorkBuddyModelsResult> GetWorkBuddyModelsAsync(ConnectionMode mode, bool force)
        {
            if (!mode.IsWorkBuddy())
            {
                throw new ArgumentException(nameof(mode));
            }

            lock (_workBuddyProbeLock)
            {
                if (_workBuddyProbes.TryGetValue(mode, out var pending) &&
                    pending.Task != null && !pending.Task.IsCompleted)
                {
                    if (!force) { return pending.Task; }

                    // 登录、安装或用户点击“自动获取”必须真正刷新。旧探测可能
                    // 仍在等待未授权 ACP 响应，不能把它原样复用成新的结果。
                    try { pending.Cancellation.Cancel(); } catch (ObjectDisposedException) { }
                }

                if (!force && _workBuddyAuthorizations.TryGetValue(mode, out var cached) && cached.IsAuthorized &&
                    _workBuddyAuthorizationAtUtc.TryGetValue(mode, out var cachedAt) &&
                    DateTime.UtcNow - cachedAt < TimeSpan.FromSeconds(15))
                {
                    return Task.FromResult(cached);
                }

                var state = new WorkBuddyProbeState
                {
                    Cancellation = CancellationTokenSource.CreateLinkedTokenSource(_workBuddyLifetime.Token),
                };
                // 先登记状态再启动异步方法：没有候选路径时 provider 可能同步完成，
                // 也必须让完成回调看到自己是当前探测，而不能把结果丢掉。
                _workBuddyProbes[mode] = state;
                state.Task = ProbeWorkBuddyAsync(mode, state);
                return state.Task;
            }
        }

        private async Task<WorkBuddyModelsResult> ProbeWorkBuddyAsync(
            ConnectionMode mode,
            WorkBuddyProbeState state)
        {
            WorkBuddyModelsResult result;
            try
            {
                result = await WorkBuddyProvider.GetModelsAsync(mode, state.Cancellation.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (
                state.Cancellation.IsCancellationRequested && !_workBuddyLifetime.IsCancellationRequested)
            {
                // force 探测替换了旧请求时，让旧调用者跟随新任务拿结果；
                // 这样 settings.get 不会因为一次刷新短暂变成“组件不可用”。
                WorkBuddyProbeState replacement = null;
                lock (_workBuddyProbeLock)
                {
                    if (_workBuddyProbes.TryGetValue(mode, out var active) &&
                        !ReferenceEquals(active, state))
                    {
                        replacement = active;
                    }
                }

                if (replacement?.Task != null)
                {
                    try { return await replacement.Task.ConfigureAwait(false); }
                    finally { state.Cancellation.Dispose(); }
                }

                result = new WorkBuddyModelsResult
                {
                    State = WorkBuddyAuthorizationState.Unavailable,
                    Code = "WORKBUDDY_PROBE_SUPERSEDED",
                    Detail = "授权状态正在刷新，请稍后重试。",
                    Models = new List<WorkBuddyModelInfo>(),
                };
            }
            catch
            {
                // provider 已过滤外部进程细节；这里再兜底，避免缓存任务永久卡住。
                result = new WorkBuddyModelsResult
                {
                    State = WorkBuddyAuthorizationState.Unavailable,
                    Code = "WORKBUDDY_UNAVAILABLE",
                    Detail = "无法读取 WorkBuddy 授权状态，请确认 WorkBuddy 已安装并可运行。",
                    Models = new List<WorkBuddyModelInfo>(),
                };
            }

            lock (_workBuddyProbeLock)
            {
                if (_workBuddyProbes.TryGetValue(mode, out var active) &&
                    ReferenceEquals(active, state))
                {
                    _workBuddyAuthorizations[mode] = result;
                    _workBuddyAuthorizationAtUtc[mode] = DateTime.UtcNow;
                    _workBuddyProbes.Remove(mode);
                }
            }

            state.Cancellation.Dispose();
            return result;
        }

        private object GetSettingsPayload(Settings settings, WorkBuddyModelsResult authorization = null)
        {
            if (settings == null)
            {
                settings = CloneSettings(_settings);
            }

            // ACP 目录是当前账号的权威边界。设置文件里的模型可能来自另一
            // 个 WorkBuddy 版本（尤其是国内/国际切换后的旧快照），返回面板
            // 前先校正，避免旧 ID 被继续显示或直接发给新账号。
            ReconcileWorkBuddyModel(settings, authorization);

            var hasCustomToken = SecretStore.Exists(Settings.CustomApiSecretKey);
            var maskedToken = string.Empty;
            if (hasCustomToken)
            {
                // 只回传掩码，密钥本身绝不出加载项进程。
                maskedToken = SecretStore.Mask(SecretStore.Load(Settings.CustomApiSecretKey));
            }

            var activeProxySecretKey = Settings.ProxySecretKeyFor(settings.ActiveProxyProfileId);
            var hasProxyPassword = SecretStore.Exists(activeProxySecretKey);
            var maskedProxyPassword = hasProxyPassword
                ? SecretStore.Mask(SecretStore.Load(activeProxySecretKey))
                : string.Empty;

            var proxyProfiles = new List<object>();
            foreach (var profile in settings.ProxyProfiles ?? new List<ProxyProfile>())
            {
                var secretKey = Settings.ProxySecretKeyFor(profile.Id);
                var hasPassword = SecretStore.Exists(secretKey);
                proxyProfiles.Add(new
                {
                    id = profile.Id,
                    name = profile.Name,
                    proxyType = profile.Kind.ToString(),
                    proxyHost = profile.Host,
                    proxyPort = profile.Port,
                    proxyUsername = profile.Username,
                    hasPassword,
                    maskedPassword = hasPassword ? SecretStore.Mask(SecretStore.Load(secretKey)) : string.Empty,
                });
            }

            // 由后端给出权威的就绪判断：它才知道 CLI 配置里有没有模型、
            // 密钥是否真的能解开。前端自行推断会与实际不一致。
            var ready = false;
            var readyDetail = string.Empty;
            var effectiveModel = settings.Model;

            if (settings.Mode.IsWorkBuddy())
            {
                if (authorization == null)
                {
                    readyDetail = "正在读取 WorkBuddy 授权状态。";
                }
                else if (!authorization.IsAuthorized)
                {
                    readyDetail = authorization.Detail;
                }
                else
                {
                    effectiveModel = string.IsNullOrWhiteSpace(settings.Model)
                        ? authorization.CurrentModelId
                        : settings.Model;
                    ready = !string.IsNullOrWhiteSpace(effectiveModel);
                    var sourceLabel = settings.Mode == ConnectionMode.AuthorizedInternational
                        ? "WorkBuddy 国际版授权"
                        : "WorkBuddy 授权";
                    readyDetail = ready
                        ? $"{sourceLabel} · {effectiveModel}"
                        : authorization.Detail + " 请在模型列表中选择模型。";
                }
            }
            else
            {
                try
                {
                    var connection = settings.ResolveConnection();
                    effectiveModel = connection.Model;

                    if (string.IsNullOrWhiteSpace(connection.Model))
                    {
                        readyDetail = $"{connection.SourceLabel} 的配置未指定模型，请选择或填写模型名";
                    }
                    else
                    {
                        ready = true;
                        readyDetail = $"{connection.SourceLabel} · {connection.BaseUrl} · {connection.Model}";
                    }
                }
                catch (ProviderException ex)
                {
                    readyDetail = ex.Message;
                }
                catch (Exception ex)
                {
                    readyDetail = "配置解析失败：" + ex.Message;
                }
            }

            return new
            {
                mode = settings.Mode.ToString(),
                channelLabel = ChannelLabel(settings, authorization),
                cliSource = settings.CliSource.ToString(),
                customProtocol = Protocols.Get(settings.CustomProtocol).Id,
                customBaseUrl = settings.CustomBaseUrl,
                proxyType = settings.ProxyType.ToString(),
                proxyHost = settings.ProxyHost,
                proxyPort = settings.ProxyPort,
                proxyUsername = settings.ProxyUsername,
                hasProxyPassword,
                maskedProxyPassword,
                proxyProfiles,
                activeProxyProfileId = settings.ActiveProxyProfileId,
                autoSwitchProxy = settings.AutoSwitchProxy,
                model = settings.Model,
                // CLI 配置自带模型时，这里会是那个值，而 model 字段仍为空。
                effectiveModel,
                thinking = settings.Thinking.ToString(),
                approval = settings.Approval.ToString(),
                temperature = settings.Temperature,
                maxOutputTokens = settings.MaxOutputTokens,
                contextBudgetTokens = settings.ContextBudgetTokens,
                maxSteps = settings.MaxSteps,
                autoIncludeSelection = settings.AutoIncludeSelection,
                toolProtocol = settings.ToolProtocol.ToString(),
                visionRelayModel = settings.VisionRelayModel,
                onlyFavoriteModels = settings.OnlyFavoriteModels,
                favorites = FavoriteModels.Load(settings.FavoritesKey()),
                // 三态由后端给权威判断，面板只做投影。
                availability = AvailabilityPayload(settings),
                hasCustomToken,
                maskedToken,
                ready,
                readyDetail,
                authorization = settings.Mode.IsWorkBuddy()
                    ? BuildWorkBuddyAuthorizationPayload(authorization)
                    : null,
                workbuddyRuntime = settings.Mode.IsWorkBuddy()
                    ? WorkBuddyRuntime.Status(settings.Mode)
                    : null,
                protocols = ProtocolOptions(),
                thinkingOptions = ThinkingOptions(),
                // 当前协议实际支持的档位，界面据此标注哪些会被降级。
                thinkingSupported = Thinking.SupportedLevels(EffectiveProtocol(settings)),
                approvalOptions = ApprovalOptions(_wordHost),
                toolProtocolOptions = ToolProtocolOptions(),
            };
        }

        /// <summary>宿主专属桥保存设置后刷新共享通道的快照。</summary>
        internal void ReloadSettings()
        {
            _settings = _loadSettings();
        }

        internal static string ChannelLabel(Settings settings, WorkBuddyModelsResult authorization = null)
        {
            settings = settings ?? new Settings();
            if (settings.Mode == ConnectionMode.CustomApi) { return "DIY"; }
            if (settings.Mode.IsWorkBuddy())
            {
                var identity = authorization?.UserName?.Trim();
                return string.IsNullOrWhiteSpace(identity)
                    ? "work buddy"
                    : "work buddy · " + identity;
            }

            try
            {
                var cli = LocalCliConfig.Resolve(settings.CliSource);
                return CliChannelLabel(cli);
            }
            catch
            {
                if (settings.CliSource == CliKind.Codex) { return "codex cli"; }
                if (settings.CliSource == CliKind.Claude) { return "claude cli"; }
                return "本机 CLI";
            }
        }

        internal static string CliChannelLabel(CliCredentials cli)
        {
            if (cli?.Source != CliKind.Codex) { return "claude cli"; }
            return string.IsNullOrWhiteSpace(cli.ProviderName)
                ? "codex cli"
                : "codex cli · " + cli.ProviderName.Trim();
        }

        private static bool ReconcileWorkBuddyModel(
            Settings settings,
            WorkBuddyModelsResult authorization)
        {
            if (settings == null || !settings.Mode.IsWorkBuddy() || authorization == null)
            {
                return false;
            }

            if (authorization.State == WorkBuddyAuthorizationState.Unavailable) { return false; }
            if (!authorization.IsAuthorized)
            {
                // 明确未授权才清除；暂时不可用保留同一连接的选择，发送仍会校验授权。
                var cleared = !string.IsNullOrWhiteSpace(settings.Model);
                settings.Model = string.Empty;
                settings.ModelConnection = string.Empty;
                return cleared;
            }

            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var model in authorization.Models ?? new List<WorkBuddyModelInfo>())
            {
                var id = (model?.ModelId ?? string.Empty).Trim();
                if (id.Length > 0) { ids.Add(id); }
            }

            var current = (settings.Model ?? string.Empty).Trim();
            var selected = ids.Contains(current) ? current : null;
            var preferred = (authorization.CurrentModelId ?? string.Empty).Trim();
            if (selected == null && preferred.Length > 0 && ids.Contains(preferred))
            {
                selected = preferred;
            }

            if (selected == null)
            {
                selected = string.Empty;
                foreach (var id in ids)
                {
                    selected = id;
                    break;
                }
            }

            var changed = !string.Equals(settings.Model ?? string.Empty, selected, StringComparison.Ordinal);
            settings.Model = selected;
            settings.StampModelConnection();
            return changed;
        }

        /// <summary>
        /// 复制设置中的值类型和字符串，供跨 await 的响应使用。
        /// 不保存引用，避免 settings.save/session.update 改写共享对象后污染旧响应。
        /// </summary>
        private static Settings CloneSettings(Settings source)
        {
            return (source ?? new Settings()).Clone();
        }

        internal static object BuildWorkBuddyAuthorizationPayload(WorkBuddyModelsResult result)
        {
            var models = new List<object>();
            foreach (var model in result?.Models ?? new List<WorkBuddyModelInfo>())
            {
                models.Add(new
                {
                    modelId = model.ModelId,
                    name = model.Name,
                    supportsImages = model.SupportsImages,
                    supportsReasoning = model.SupportsReasoning,
                    supportsToolCall = model.SupportsToolCall,
                    maxInputTokens = model.MaxInputTokens,
                    multiplier = model.CreditMultiplier,
                });
            }

            return result == null
                ? null
                : new
                {
                    status = WorkBuddyProvider.StatusId(result.State),
                    code = result.Code,
                    detail = result.Detail,
                    currentModelId = result.CurrentModelId,
                    models,
                };
        }

        private async Task<object> RunWorkBuddyActionAsync(bool install, ConnectionMode mode, string operationId, bool switchAccount = false)
        {
            if (string.IsNullOrWhiteSpace(operationId)) { return new { ok = false, detail = "登录操作已失效，请刷新面板后重试。" }; }
            if (install && mode != ConnectionMode.AuthorizedInternational)
            {
                return new
                {
                    ok = false,
                    detail = "独立授权组件仅用于国际版；国内版请启动或安装 WorkBuddy 桌面端。",
                };
            }

            var cancellation = CancellationTokenSource.CreateLinkedTokenSource(_workBuddyLifetime.Token);
            var action = new WorkBuddyActionState { Cancellation = cancellation, Id = operationId, Mode = mode };
            if (Interlocked.CompareExchange(ref _workBuddyAction, action, null) != null)
            {
                cancellation.Dispose();
                return new { ok = false, detail = "已有组件操作正在进行，请等待完成或取消。" };
            }
            var operation = install ? "install" : "login";
            var stage = "discover";
            try
            {
                if (switchAccount && !install)
                {
                    return new { ok = true, detail = WorkBuddyAccount.OpenAccountClient(mode) };
                }
                WorkBuddyModelsResult authorization = null;
                var found = WorkBuddyProvider.TryFindPaths(mode, out var paths);
                Log.Info($"WorkBuddy 操作开始：操作={operation} 模式={mode} 位数={(Environment.Is64BitProcess ? "x64" : "x86")} " +
                    $"完全信任={AppDomain.CurrentDomain.IsFullyTrusted} 基目录={AppDomain.CurrentDomain.BaseDirectory} " +
                    $"已发现={found} CLI={paths?.CliPath ?? "<none>"} Node={paths?.NodePath ?? "<none>"} " +
                    $"配置目录={paths?.ConfigDirectory ?? "<none>"} 独立组件={WorkBuddyRuntime.FindNativePath() ?? "<none>"}");
                if (!found)
                {
                    Log.Warn("WorkBuddy 路径诊断：" + WorkBuddyRuntime.NativePathDiagnostics());
                }
                if (install)
                {
                    stage = "install";
                    await WorkBuddyRuntime.InstallAsync(detail =>
                    {
                        _ = _pushRaw(new { kind = "workbuddy.progress", operationId, detail });
                    }, cancellation.Token).ConfigureAwait(false);
                }
                else
                {
                    stage = "authenticate";
                    authorization = await WorkBuddyAccount.LoginAsync(mode, cancellation.Token, authUrl =>
                    {
                        if (!ReferenceEquals(_workBuddyAction, action) || cancellation.IsCancellationRequested) { return; }
                        action.AuthUrl = authUrl;
                        _ = _pushRaw(new
                        {
                            kind = "workbuddy.auth-url",
                            operationId,
                            mode = mode.ToString(),
                            url = authUrl,
                        });
                    }).ConfigureAwait(false);
                }

                // force 探测会取消并替换登录前遗留的 pending 请求；这里不能先
                // 等它自然结束，否则旧的未授权结果会把登录流程卡住几十秒。
                stage = "refresh-models";
                if (authorization == null) { authorization = await GetWorkBuddyModelsAsync(mode, force: true).ConfigureAwait(false); }
                else
                {
                    // 采用登录来源已确认的结果，废弃登录前的探测，避免旧未授权覆盖新账号。
                    lock (_workBuddyProbeLock)
                    {
                        if (_workBuddyProbes.TryGetValue(mode, out var pending))
                        {
                            _workBuddyProbes.Remove(mode);
                            try { pending.Cancellation.Cancel(); } catch (ObjectDisposedException) { }
                        }
                        _workBuddyAuthorizations[mode] = authorization;
                        _workBuddyAuthorizationAtUtc[mode] = DateTime.UtcNow;
                    }
                }
                stage = "refresh-checkin";
                var checkin = authorization.IsAuthorized && mode.IsDomesticWorkBuddy()
                    ? await WorkBuddyAccount.RefreshAsync(mode, true, cancellation.Token).ConfigureAwait(false)
                    : null;
                Log.Info($"WorkBuddy 操作完成：操作={operation} 模式={mode} 授权={authorization.State} 模型数={authorization.Models?.Count ?? 0}");
                return new
                {
                    ok = install || authorization.IsAuthorized,
                    detail = install ? "独立组件已安装，请点击浏览器登录。" : authorization.IsAuthorized ? "登录成功，账号与模型已刷新。" : authorization.Detail,
                    runtime = WorkBuddyRuntime.Status(mode),
                    authorization = BuildWorkBuddyAuthorizationPayload(authorization),
                    checkin,
                };
            }
            catch (OperationCanceledException)
            {
                var detail = cancellation.IsCancellationRequested ? "操作已取消。" : "操作超时，请重新尝试。";
                Log.Warn($"WorkBuddy 操作取消：操作={operation} 模式={mode} 阶段={stage} 详情={detail}");
                return new { ok = false, code = "WORKBUDDY_OPERATION_CANCELLED", detail };
            }
            catch (ProviderException ex)
            {
                Log.Warn($"WorkBuddy 操作失败：操作={operation} 模式={mode} 阶段={stage} 错误码={ex.Code} 详情={ex.Message}");
                return new { ok = false, code = ex.Code, detail = ex.Message };
            }
            catch (Exception ex)
            {
                Log.Error($"WorkBuddy 操作异常：操作={operation} 模式={mode} 阶段={stage}", ex);
                return new
                {
                    ok = false,
                    code = "WORKBUDDY_OPERATION_FAILED",
                    detail = $"组件操作未完成（阶段：{stage}），请重试并打开诊断查看错误码。",
                };
            }
            finally
            {
                action.AuthUrl = null;
                Interlocked.CompareExchange(ref _workBuddyAction, null, action);
                cancellation.Dispose();
            }
        }

        /// <summary>取当前生效的协议，用于判断思考档位支持范围。</summary>
        private ProtocolKind EffectiveProtocol()
        {
            return EffectiveProtocol(_settings);
        }

        private static ProtocolKind EffectiveProtocol(Settings settings)
        {
            settings = settings ?? new Settings();
            try
            {
                return settings.ResolveConnection().Protocol;
            }
            catch
            {
                return settings.Mode == ConnectionMode.CustomApi
                    ? settings.CustomProtocol
                    : Protocols.Default;
            }
        }

        private static object ThinkingOptions()
        {
            var list = new List<object>();
            foreach (var option in Thinking.Options)
            {
                list.Add(new { id = option.Id, label = option.Label, hint = option.Hint });
            }

            return list;
        }

        private static object ApprovalOptions(bool wordHost)
        {
            return new[]
            {
                new { id = "PerWrite", label = "逐项审批", hint = "写操作逐项确认，读操作自动执行" },
                new { id = "PerTurn", label = "每轮确认", hint = wordHost ? "本轮当前文档每类操作首次确认，后续同类不再问；结构单独确认" : "本轮第一次写操作问一次，之后同一工作表同一类不再问；结构单独问" },
                new { id = "Automatic", label = "全自动", hint = wordHost ? "写操作直接执行；仅工具提供真实快照时可撤销" : "不询问，依赖 Excel 撤销兜底" },
            };
        }

        private static object ToolProtocolOptions()
        {
            return new[]
            {
                new { id = "Auto", label = "自动探测", hint = "先按原生方式发，被拒或模型推辞后自动改用文本指令" },
                new { id = "Native", label = "原生函数调用", hint = "多数模型支持，效果最好" },
                new { id = "Text", label = "文本指令", hint = "把工具清单写进提示词，适合不支持函数调用的模型" },
                new { id = "None", label = "不用工具", hint = "只给方案与公式，不读写表格" },
            };
        }

        private static object ProtocolOptions()
        {
            var list = new List<object>();
            foreach (var protocol in Protocols.All)
            {
                list.Add(new { id = protocol.Id, label = protocol.Label });
            }

            return list;
        }

        private static void ApplyProxyProfilesPayload(Settings settings, JObject payload)
        {
            var array = payload?["proxyProfiles"] as JArray;
            if (array == null) { return; }

            var profiles = new List<ProxyProfile>();
            foreach (var token in array)
            {
                var item = token as JObject;
                if (item == null) { continue; }
                var profile = new ProxyProfile
                {
                    Id = item.Value<string>("id") ?? string.Empty,
                    Name = item.Value<string>("name") ?? string.Empty,
                    Host = item.Value<string>("proxyHost") ?? string.Empty,
                    Port = item.Value<int?>("proxyPort") ?? 0,
                    Username = item.Value<string>("proxyUsername") ?? string.Empty,
                };
                if (Enum.TryParse(item.Value<string>("proxyType"), out ProxyKind kind)) { profile.Kind = kind; }
                profiles.Add(profile);
            }

            if (profiles.Count > 0)
            {
                settings.ProxyProfiles = profiles;
                settings.ActiveProxyProfileId = payload.Value<string>("activeProxyProfileId") ?? string.Empty;
                settings.AutoSwitchProxy = payload.Value<bool?>("autoSwitchProxy") ?? true;
                settings.SyncLegacyFieldsFromActive();
            }
        }

        private async Task<object> SaveSettingsAsync(JObject payload)
        {
            var settings = Settings.Load();
            var previousConnectionKey = settings.ConnectionKey();
            // 面板确认「这个模型是在当前这套接入配置下选的」。缺少确认时，
            // 一旦连接发生变化就只能当作上一套配置的残留处理。
            var modelChosenForConnection = payload.Value<bool?>("modelChosenForConnection") ?? false;

            if (Enum.TryParse(payload.Value<string>("mode"), out ConnectionMode mode)) { settings.Mode = mode; }
            if (Enum.TryParse(payload.Value<string>("cliSource"), out CliKind cli)) { settings.CliSource = cli; }
            if (Protocols.TryParse(payload.Value<string>("customProtocol"), out var protocol)) { settings.CustomProtocol = protocol; }
            if (Enum.TryParse(payload.Value<string>("proxyType"), out ProxyKind proxyType)) { settings.ProxyType = proxyType; }
            if (Thinking.TryParse(payload.Value<string>("thinking"), out var thinking)) { settings.Thinking = thinking; }
            if (Enum.TryParse(payload.Value<string>("approval"), out ApprovalPolicy approval)) { settings.Approval = approval; }

            if (payload["customBaseUrl"] != null) { settings.CustomBaseUrl = payload.Value<string>("customBaseUrl") ?? string.Empty; }
            if (payload["proxyHost"] != null) { settings.ProxyHost = payload.Value<string>("proxyHost") ?? string.Empty; }
            if (payload["proxyPort"] != null) { settings.ProxyPort = payload.Value<int?>("proxyPort") ?? 0; }
            if (payload["proxyUsername"] != null) { settings.ProxyUsername = payload.Value<string>("proxyUsername") ?? string.Empty; }
            if (payload["model"] != null) { settings.Model = payload.Value<string>("model") ?? string.Empty; }
            if (payload["proxyProfiles"] is JArray)
            {
                ApplyProxyProfilesPayload(settings, payload);
            }
            else
            {
                // 兼容旧版面板只提交单代理字段的保存请求。
                settings.SyncActiveProxyProfileFromLegacy();
            }
            // 模型归属必须在协议、地址、CLI 来源都写完后再判定，
            // 否则算出的连接键还是旧的。
            settings.KeepModelOnlyIfChosenForConnection(previousConnectionKey, modelChosenForConnection);
            if (payload["temperature"] != null)
            {
                settings.Temperature = payload["temperature"].Type == JTokenType.Null
                    ? (double?)null
                    : payload.Value<double>("temperature");
            }

            if (payload["maxOutputTokens"] != null) { settings.MaxOutputTokens = payload.Value<int>("maxOutputTokens"); }
            if (payload["contextBudgetTokens"] != null) { settings.ContextBudgetTokens = payload.Value<int>("contextBudgetTokens"); }
            if (payload["maxSteps"] != null) { settings.MaxSteps = payload.Value<int>("maxSteps"); }
            if (payload["autoIncludeSelection"] != null) { settings.AutoIncludeSelection = payload.Value<bool>("autoIncludeSelection"); }
            if (Enum.TryParse(payload.Value<string>("toolProtocol"), out ToolProtocolPreference toolProtocol))
            {
                // 用户改了工具形态就把探测结果作废：从「文本指令」改回「自动探测」时，
                // 留着上次探出的降级档等于这个选项没生效。
                if (settings.ToolProtocol != toolProtocol)
                {
                    ModelCapabilities.Reset();
                }

                settings.ToolProtocol = toolProtocol;
            }

            if (payload["visionRelayModel"] != null)
            {
                settings.VisionRelayModel = payload.Value<string>("visionRelayModel") ?? string.Empty;
            }

            // 密钥单独走加密存储；面板传空字符串表示清除。
            // Authorized 不使用 ChatSheet 的自定义密钥槽，WorkBuddy 令牌始终
            // 留在 WorkBuddy 自己的授权存储中。
            var token = payload.Value<string>("customToken");
            if (settings.Mode == ConnectionMode.CustomApi && token != null)
            {
                if (string.IsNullOrWhiteSpace(token))
                {
                    SecretStore.Delete(Settings.CustomApiSecretKey);
                }
                else
                {
                    SecretStore.Save(Settings.CustomApiSecretKey, token.Trim());
                }

                // 写了密钥就作废该连接的可用性判定：一个账号能碰到哪些模型跟着密钥走。
                // 按「写了密钥」触发而不去比对新旧——比对要把已存的密钥读回来，
                // 无谓地多碰一次密钥，而多作废一次只是让下一轮重新记一遍。
                //
                // 只有自定义接口这条路会经过这里。本机 CLI 的密钥在 CLI 自己的配置里、
                // 不经 SecretStore，那条路上没有这个触发点——不是漏了。
                ModelAvailability.ResetConnection(settings.ConnectionKey());
            }

            var proxyPasswordChanges = payload["proxyPasswordChanges"] as JArray;
            if (proxyPasswordChanges != null)
            {
                foreach (var changeToken in proxyPasswordChanges)
                {
                    var change = changeToken as JObject;
                    if (change == null) { continue; }
                    var profileId = change.Value<string>("profileId") ?? string.Empty;
                    var password = change.Value<string>("password");
                    if (string.IsNullOrWhiteSpace(profileId) || password == null) { continue; }
                    var secretKey = Settings.ProxySecretKeyFor(profileId);
                    if (string.IsNullOrWhiteSpace(password))
                    {
                        SecretStore.Delete(secretKey);
                    }
                    else
                    {
                        SecretStore.Save(secretKey, password.Trim());
                    }
                }
            }
            else
            {
                // 兼容旧版面板的一次性单代理密码字段。
                var proxyPassword = payload.Value<string>("proxyPassword");
                if (proxyPassword != null)
                {
                    var profileId = payload.Value<string>("proxyPasswordProfileId") ?? settings.ActiveProxyProfileId;
                    var secretKey = Settings.ProxySecretKeyFor(profileId);
                    if (string.IsNullOrWhiteSpace(proxyPassword))
                    {
                        SecretStore.Delete(secretKey);
                    }
                    else
                    {
                        SecretStore.Save(secretKey, proxyPassword.Trim());
                    }
                }
            }

            // 换了连接同样作废：判定的键含连接，旧连接那份留着也不会被查到，
            // 但新连接可能与某个旧连接同键（改回来），那时留着的就是过期结论。
            if (previousConnectionKey != settings.ConnectionKey())
            {
                ModelAvailability.ResetConnection(settings.ConnectionKey());
            }

            settings.Save();
            _settings = settings;

            // 与探测一起返回同一份保存快照，不能在 await 后回读可能已经被
            // 另一条 settings.save 或 chat.send 替换的 _settings。
            var responseSettings = CloneSettings(settings);
            var authorization = responseSettings.Mode.IsWorkBuddy()
                ? await GetWorkBuddyModelsAsync(responseSettings.Mode, force: true).ConfigureAwait(false)
                : null;
            if (ReconcileWorkBuddyModel(responseSettings, authorization) && responseSettings.Mode.IsWorkBuddy())
            {
                // 面板可能带着旧目录提交了保存。把后端校正后的模型一并落盘，
                // 避免下一次打开或直接发送又回到失效的国内/国际模型。
                settings.Model = responseSettings.Model;
                settings.StampModelConnection();
                settings.Save();
                _settings = settings;
            }
            return GetSettingsPayload(responseSettings, authorization);
        }

        private object ProbeCliPayload()
        {
            var list = new List<object>();
            foreach (var probe in LocalCliConfig.Probe())
            {
                list.Add(new
                {
                    kind = probe.Kind.ToString(),
                    displayName = probe.DisplayName,
                    configPath = probe.ConfigPath,
                    exists = probe.Exists,
                    usable = probe.Usable,
                    protocol = Protocols.Get(probe.Protocol).Id,
                    baseUrl = probe.BaseUrl,
                    model = probe.Model,
                    detail = probe.Detail,
                });
            }

            return new { candidates = list };
        }

        private async Task<object> TestProxyAsync(JObject payload)
        {
            var settings = Settings.Load();
            if (Enum.TryParse(payload.Value<string>("proxyType"), out ProxyKind proxyType))
            {
                settings.ProxyType = proxyType;
            }
            if (payload["proxyHost"] != null) { settings.ProxyHost = payload.Value<string>("proxyHost") ?? string.Empty; }
            if (payload["proxyPort"] != null) { settings.ProxyPort = payload.Value<int?>("proxyPort") ?? 0; }
            if (payload["proxyUsername"] != null) { settings.ProxyUsername = payload.Value<string>("proxyUsername") ?? string.Empty; }
            if (payload["proxyProfiles"] is JArray)
            {
                ApplyProxyProfilesPayload(settings, payload);
            }
            else
            {
                settings.SyncActiveProxyProfileFromLegacy();
            }

            try
            {
                var password = payload["proxyPassword"] != null
                    ? payload.Value<string>("proxyPassword")?.Trim() ?? string.Empty
                    : null;
                var proxy = settings.ResolveProxy(password);
                using (var cts = CancellationTokenSource.CreateLinkedTokenSource(_workBuddyLifetime.Token))
                {
                    cts.CancelAfter(TimeSpan.FromSeconds(15));
                    return await ProxyTransport.TestAsync(proxy, cts.Token).ConfigureAwait(false);
                }
            }
            catch (ProviderException ex)
            {
                return new { ok = false, detail = ex.Message, elapsedMs = 0L };
            }
        }

        private async Task<object> ListModelsAsync(JObject payload)
        {
            var settings = Settings.Load();

            // 面板可能已切换模式但尚未保存，必须按它当前选择的模式解析，
            // 否则会拿旧设置去连（例如界面已切到本机 CLI，却仍按空的自定义地址解析而报错）。
            if (Enum.TryParse(payload.Value<string>("mode"), out ConnectionMode pendingMode))
            {
                settings.Mode = pendingMode;
            }

            if (Enum.TryParse(payload.Value<string>("cliSource"), out CliKind pendingCli))
            {
                settings.CliSource = pendingCli;
            }

            if (Enum.TryParse(payload.Value<string>("proxyType"), out ProxyKind pendingProxyType))
            {
                settings.ProxyType = pendingProxyType;
            }
            if (payload["proxyHost"] != null) { settings.ProxyHost = payload.Value<string>("proxyHost") ?? string.Empty; }
            if (payload["proxyPort"] != null) { settings.ProxyPort = payload.Value<int?>("proxyPort") ?? 0; }
            if (payload["proxyUsername"] != null) { settings.ProxyUsername = payload.Value<string>("proxyUsername") ?? string.Empty; }
            if (payload["proxyProfiles"] is JArray)
            {
                ApplyProxyProfilesPayload(settings, payload);
            }
            else
            {
                settings.SyncActiveProxyProfileFromLegacy();
            }

            if (settings.Mode.IsWorkBuddy())
            {
                var force = payload.Value<bool?>("force") ?? false;
                var authorization = await GetWorkBuddyModelsAsync(settings.Mode, force).ConfigureAwait(false);
                var models = new List<object>();
                foreach (var model in authorization.Models ?? new List<WorkBuddyModelInfo>())
                {
                    models.Add(new
                    {
                        modelId = model.ModelId,
                        name = model.Name,
                        supportsImages = model.SupportsImages,
                        supportsReasoning = model.SupportsReasoning,
                        supportsToolCall = model.SupportsToolCall,
                        maxInputTokens = model.MaxInputTokens,
                        multiplier = model.CreditMultiplier,
                    });
                }

                return new
                {
                    protocol = WorkBuddyProvider.ProtocolId,
                    baseUrl = string.Empty,
                    models,
                    authorization = BuildWorkBuddyAuthorizationPayload(authorization),
                    runtime = WorkBuddyRuntime.Status(settings.Mode),
                    manualEntryRequired = models.Count == 0,
                };
            }

            // 允许面板传入尚未保存的地址与密钥，以便保存前先试连。
            var protocolId = payload.Value<string>("protocol");
            var baseUrlInput = payload.Value<string>("baseUrl");
            var tokenInput = payload.Value<string>("token");

            ProtocolKind protocol;
            string baseUrl;
            string token;
            IReadOnlyList<ProxyOptions> proxies;

            if (settings.Mode == ConnectionMode.CustomApi && !string.IsNullOrWhiteSpace(baseUrlInput))
            {
                protocol = Protocols.TryParse(protocolId, out var parsed) ? parsed : settings.CustomProtocol;
                baseUrl = Protocols.NormalizeBaseUrl(baseUrlInput, protocol);
                token = string.IsNullOrWhiteSpace(tokenInput)
                    ? SecretStore.Load(Settings.CustomApiSecretKey)
                    : tokenInput.Trim();
                var pendingProxyPassword = payload["proxyPassword"] != null
                    ? payload.Value<string>("proxyPassword")?.Trim() ?? string.Empty
                    : null;
                proxies = settings.ResolveProxyCandidates(pendingProxyPassword);
            }
            else
            {
                var connection = settings.ResolveConnection();
                protocol = connection.Protocol;
                baseUrl = connection.BaseUrl;
                token = connection.Token;
                proxies = connection.ProxyCandidates != null && connection.ProxyCandidates.Count > 0
                    ? connection.ProxyCandidates : new[] { connection.Proxy };
            }

            // 预算 = 单次请求的 30 秒 + 全部重试的退避时长。
            // 只给 30 秒会让重试还没走完就被超时掐断。
            var budget = TimeSpan.FromSeconds(30) + RetryPolicy.TotalBackoff;

            using (var client = new ChatClient(proxies))
            using (var cts = CancellationTokenSource.CreateLinkedTokenSource(_workBuddyLifetime.Token))
            {
                cts.CancelAfter(budget);
                var models = await client.ListModelsAsync(
                    protocol,
                    baseUrl,
                    token,
                    cts.Token,
                    // 重试期间界面仍显示「获取中…」，不说明就像卡住了。
                    (attempt, delay, reason) => _pushRaw(new
                    {
                        kind = "models-retry",
                        requestId = payload.Value<string>("requestId"),
                        text = RetryPolicy.Describe(attempt, delay, reason),
                        attempt,
                        maxRetries = RetryPolicy.MaxRetries,
                    })).ConfigureAwait(false);

                return new
                {
                    protocol = Protocols.Get(protocol).Id,
                    baseUrl,
                    models,
                    // 空列表不是错误：部分网关不提供模型列表，需允许手填。
                    manualEntryRequired = models.Count == 0,
                };
            }
        }

        private object Stop()
        {
            var cts = _currentRun;
            if (cts == null)
            {
                return new { stopped = false, reason = "当前没有进行中的任务" };
            }

            try
            {
                cts.Cancel();
                return new { stopped = true };
            }
            catch (Exception ex)
            {
                return new { stopped = false, reason = ex.Message };
            }
        }

        private async Task<object> SendAsync(JObject payload)
        {
            if (Volatile.Read(ref _disposed) != 0) { throw new OperationCanceledException("面板已关闭"); }
            var cts = CancellationTokenSource.CreateLinkedTokenSource(_workBuddyLifetime.Token);
            if (Interlocked.CompareExchange(ref _currentRun, cts, null) != null)
            {
                cts.Dispose();
                throw new ProviderException("BUSY", "上一轮任务尚未结束，请稍后重发这条内容。");
            }
            try
            {
                return await SendCoreAsync(payload, cts).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                await _push(new AgentUpdate { Kind = "stopped", Text = "已停止生成。" }).ConfigureAwait(false);
                return new { completed = false, stopped = true };
            }
            finally
            {
                FailPendingApprovals("任务已结束");
                Interlocked.CompareExchange(ref _currentRun, null, cts);
                cts.Dispose();
            }
        }

        private async Task<object> SendCoreAsync(JObject payload, CancellationTokenSource cts)
        {
            cts.Token.ThrowIfCancellationRequested();
            var input = payload.Value<string>("text");
            var images = ParseImages(payload);
            var files = ParseFiles(payload);

            // 文件内容拼进用户输入。图片走协议的多模态字段，文件走文本——
            // 四种协议都没有「文本文件」这类内容块，带围栏的代码块才是通用形式。
            var composed = FileSupport.Compose(input, files);

            // 同一时刻只跑一轮。面板侧会把处理中的新输入排进队列并在上一轮
            // 结束后自动接着发，因此正常使用不会撞上这里；真撞上说明有第二个
            // 入口绕过了队列（例如面板刷新后旧的请求仍在途），此时如实回报，
            // 而不是让两轮交替写同一个工作簿。
            _settings = _loadSettings();
            var settings = _settings;

            // 记录本轮的接入配置（不含密钥），这是排查「发了没反应」的第一现场。
            try
            {
                if (settings.Mode.IsWorkBuddy())
                {
                    // 发送前再以 ACP 目录校正一次模型。设置页可能还没打开，
                    // 或者账号刚在另一套 WorkBuddy 中切换；不能让磁盘里的旧
                    // 模型 ID 直接进入新的 session/set_model。
                    var probe = GetWorkBuddyModelsAsync(settings.Mode, force: false);
                    var stopped = new TaskCompletionSource<WorkBuddyModelsResult>(TaskCreationOptions.RunContinuationsAsynchronously);
                    WorkBuddyModelsResult authorization;
                    using (cts.Token.Register(() => stopped.TrySetCanceled()))
                    {
                        authorization = await (await Task.WhenAny(probe, stopped.Task).ConfigureAwait(false)).ConfigureAwait(false);
                    }
                    cts.Token.ThrowIfCancellationRequested();
                    if (!authorization.IsAuthorized)
                    {
                        throw new ProviderException(
                            authorization.Code ?? "WORKBUDDY_UNAVAILABLE",
                            authorization.Detail ?? "WorkBuddy 当前不可用，请先完成授权。");
                    }

                    if (ReconcileWorkBuddyModel(settings, authorization) && ReferenceEquals(settings, _settings))
                    {
                        settings.Save();
                    }
                }

                var connection = settings.ResolveConnection();
                Log.Info($"开始对话：模式={settings.Mode} 来源={connection.SourceLabel} " +
                    $"协议={Protocols.Get(connection.Protocol).Id} 地址={connection.BaseUrl} " +
                    $"模型={connection.Model} 思考={settings.Thinking} 审批={settings.Approval} " +
                    $"输入长度={input?.Length ?? 0}" +
                    // 拼接后的长度单独记：只看输入长度会以为用户只发了一句话，
                    // 而实际进上下文的可能是几万字符的附件。
                    (files.Count > 0 ? $" 拼接后长度={composed.Length}" : string.Empty));

                if (files.Count > 0)
                {
                    Log.Info($"本轮附带 {FileSupport.Describe(files)}");
                }
            }
            catch (ProviderException ex)
            {
                // 配置不完整时直接回报，避免用户面对无反馈的界面。
                Log.Warn($"接入配置不可用：{ex.Code} {ex.Message}");
                await _push(new AgentUpdate { Kind = "error", Text = ex.Message, Payload = new { code = ex.Code } })
                    .ConfigureAwait(false);
                return new { completed = false, error = ex.Message, code = ex.Code };
            }

            try
            {
                cts.Token.ThrowIfCancellationRequested();
                await _agent.RunAsync(
                    composed,
                    settings,
                    _push,
                    RequestApprovalAsync,
                    cts.Token,
                    images).ConfigureAwait(false);

                return new { completed = true };
            }
            catch (OperationCanceledException)
            {
                await _push(new AgentUpdate { Kind = "stopped", Text = "已停止生成。" }).ConfigureAwait(false);
                return new { completed = false, stopped = true };
            }
            catch (ProviderException ex)
            {
                // 必须记日志：失败的一轮此前只把消息推给面板，加载项日志里
                // 「开始对话」之后再无下文，事后无从判断是没发出去还是被拒绝。
                Log.Warn($"对话失败：{ex.Code} {ex.Message}");
                await _push(new AgentUpdate { Kind = "error", Text = ex.Message, Payload = new { code = ex.Code } })
                    .ConfigureAwait(false);
                return new { completed = false, error = ex.Message, code = ex.Code };
            }
            catch (Exception ex)
            {
                Log.Error("Agent 运行失败", ex);
                await _push(new AgentUpdate { Kind = "error", Text = "运行失败：" + ex.Message }).ConfigureAwait(false);
                return new { completed = false, error = ex.Message };
            }
        }

        /// <summary>
        /// 解析面板传来的图片。
        ///
        /// 单张不合规就整轮拒绝，而不是静默丢弃：用户以为图片发出去了、
        /// 模型却看不到，那比明确报错更难排查。
        /// </summary>
        private static List<ImageAttachment> ParseImages(JObject payload)
        {
            var result = new List<ImageAttachment>();
            if (!(payload["images"] is JArray array) || array.Count == 0)
            {
                return result;
            }

            if (array.Count > ImageSupport.MaxImagesPerTurn)
            {
                throw new ProviderException(
                    "TOO_MANY_IMAGES",
                    $"一次最多附带 {ImageSupport.MaxImagesPerTurn} 张图片，当前有 {array.Count} 张。");
            }

            foreach (var item in array)
            {
                if (!(item is JObject image))
                {
                    continue;
                }

                var name = image.Value<string>("name") ?? "图片";
                var dataUrl = image.Value<string>("dataUrl");
                result.Add(ImageSupport.ParseDataUrl(dataUrl, name));
            }

            return result;
        }

        /// <summary>
        /// 解析面板传来的文本文件。
        ///
        /// 与图片同样的取舍：一个不合规就整轮拒绝。静默丢弃会让用户以为
        /// 文件发出去了，而模型的回答其实完全没看过它。
        /// </summary>
        private static List<TextAttachment> ParseFiles(JObject payload)
        {
            var result = new List<TextAttachment>();
            if (!(payload["files"] is JArray array) || array.Count == 0)
            {
                return result;
            }

            if (array.Count > FileSupport.MaxFilesPerTurn)
            {
                throw new ProviderException(
                    "TOO_MANY_FILES",
                    $"一次最多附带 {FileSupport.MaxFilesPerTurn} 个文件，当前有 {array.Count} 个。");
            }

            var total = 0;
            foreach (var item in array)
            {
                if (!(item is JObject file))
                {
                    continue;
                }

                var attachment = FileSupport.Create(
                    file.Value<string>("name"),
                    file.Value<string>("text"));

                total += attachment.ByteLength;
                if (total > FileSupport.MaxTotalBytes)
                {
                    throw new ProviderException(
                        "FILES_TOO_LARGE",
                        $"文件合计 {total / 1024.0:F0} KB，超过 {FileSupport.MaxTotalBytes / 1024} KB 上限。" +
                            "文件内容会整段进入上下文，因此总量也有限制。");
                }

                result.Add(attachment);
            }

            return result;
        }

        private async Task<ApprovalDecision> RequestApprovalAsync(
            ToolDefinition definition,
            JObject args,
            ImpactEstimate impact)
        {
            var id = "ap" + Interlocked.Increment(ref _approvalSequence);
            var completion = new TaskCompletionSource<ApprovalDecision>(TaskCreationOptions.RunContinuationsAsynchronously);
            _pendingApprovals[id] = completion;

            var token = _currentRun?.Token ?? _workBuddyLifetime.Token;
            using (token.Register(() => completion.TrySetCanceled()))
            try
            {
            token.ThrowIfCancellationRequested();

            await _pushRaw(new
            {
                kind = "approval-request",
                id,
                tool = definition.Name,
                description = definition.Description,
                risk = definition.Risk.ToString(),
                impact = impact?.Text ?? string.Empty,
                impactNote = impact?.Note,
                // 探到范围时另给结构化字段，面板据此把地址译成行列说明。
                impactRange = string.IsNullOrEmpty(impact?.Address)
                    ? null
                    : new
                    {
                        sheet = impact.SheetName,
                        address = impact.Address,
                        cells = impact.CellCount,
                    },

                // 「将改成什么」的截断对照。只在这条推送里出现：
                // 写进工具结果或对话历史，等于每步批准都再付一次读取的税，
                // 而最近几条消息本来就不参与上下文压缩。
                preview = impact?.Preview == null
                    ? null
                    : new
                    {
                        currentUnreadable = impact.Preview.CurrentUnreadable,
                        formattingMixed = impact.Preview.FormattingMixed,
                        omittedCells = impact.Preview.OmittedCells,
                        discardedValues = impact.Preview.DiscardedValues,
                        kind = impact.Preview.Kind,
                        cells = impact.Preview.Cells.ConvertAll(cell => new
                        {
                            row = cell.Row,
                            column = cell.Column,
                            before = cell.Before,
                            after = cell.After,
                            beforeEmpty = cell.BeforeEmpty,
                            afterEmpty = cell.AfterEmpty,
                        }),
                    },
                args,
            }).ConfigureAwait(false);

            return await completion.Task.ConfigureAwait(false);
            }
            finally { _pendingApprovals.TryRemove(id, out _); }
        }

        private Task<object> RespondApprovalAsync(JObject payload)
        {
            var id = payload.Value<string>("id");
            if (string.IsNullOrEmpty(id) || !_pendingApprovals.TryRemove(id, out var completion))
            {
                return Task.FromResult<object>(new { accepted = false, reason = "该审批请求已失效" });
            }

            completion.TrySetResult(new ApprovalDecision
            {
                Approved = payload.Value<bool?>("approved") ?? false,
                Reason = payload.Value<string>("reason"),
                ApproveRest = payload.Value<bool?>("approveRest") ?? false,
                ApproveStructureRest = payload.Value<bool?>("approveStructureRest") ?? false,
            });

            return Task.FromResult<object>(new { accepted = true });
        }

        /// <summary>
        /// 让所有挂起的审批以拒绝收束。
        /// 面板关闭或任务结束时必须调用，否则 Agent 侧会永久等待。
        /// </summary>
        private void FailPendingApprovals(string reason)
        {
            foreach (var key in new List<string>(_pendingApprovals.Keys))
            {
                if (_pendingApprovals.TryRemove(key, out var completion))
                {
                    completion.TrySetResult(new ApprovalDecision { Approved = false, Reason = reason });
                }
            }
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0) { return; }
            _workBuddyLifetime.Cancel();
            StopBulkProbe();
            try
            {
                _currentRun?.Cancel();
            }
            catch
            {
            }

            FailPendingApprovals("面板已关闭");
            _agent?.Tools.Undo.Clear();
        }
    }
}
