using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ChatSheet.AddIn;
using ChatSheet.AddIn.Agent;
using ChatSheet.AddIn.Bridge;
using ChatSheet.AddIn.Providers;
using ChatSheet.AddIn.Storage;
using ChatSheet.AddIn.Tools;
using ChatWord.AddIn.Agent;
using ChatWord.AddIn.Hosts;
using ChatWord.AddIn.Tools;
using Microsoft.Web.WebView2.Core;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ChatWord.AddIn.Bridge
{
    internal sealed class WordBridge : IPanelBridge
    {
        private readonly CoreWebView2 _core;
        private readonly Func<object> _applicationAccessor;
        private readonly SynchronizationContext _uiContext;
        private readonly CancellationTokenSource _lifetime = new CancellationTokenSource();
        private readonly Dictionary<string, Func<JObject, Task<object>>> _handlers = new Dictionary<string, Func<JObject, Task<object>>>(StringComparer.Ordinal);
        private readonly ConcurrentDictionary<string, TaskCompletionSource<ApprovalDecision>> _approvals = new ConcurrentDictionary<string, TaskCompletionSource<ApprovalDecision>>(StringComparer.Ordinal);
        private readonly AgentChannels _sharedChannels;
        private readonly WordAgentRunner _agent;
        private readonly WordDocumentContext _context;
        private Settings _settings;
        private CancellationTokenSource _run;
        private bool _disposed;
        private int _approvalSequence;

        internal WordBridge(CoreWebView2 core, Func<object> applicationAccessor)
        {
            _core = core ?? throw new ArgumentNullException(nameof(core)); _applicationAccessor = applicationAccessor ?? throw new ArgumentNullException(nameof(applicationAccessor)); _uiContext = SynchronizationContext.Current;
            _settings = Settings.Load();
            // WorkBuddy ACP、模型目录和授权操作必须与 Excel 入口走同一套通道。
            // Word 只替换聊天/文档执行器，不能把这些能力退化为空响应。
            _sharedChannels = new AgentChannels(applicationAccessor, PushAgentUpdateAsync, PushRawAsync, InvokeOnUiAsync, Settings.Load, createAgent: false, wordHost: true);
            _agent = new WordAgentRunner(applicationAccessor, "OfficeHelper.Word.AddIn", InvokeOnUiAsync);
            _context = _agent.Tools.Context;
            _sharedChannels.Register(_handlers);
            RegisterHandlers();
        }

        public Func<int, int, double, int> WidthAdjuster { get; set; }
        public Func<int> WidthPersister { get; set; }
        public Func<string, bool> ThemeApplier { get; set; }

        public void Start() { _core.WebMessageReceived += OnWebMessageReceived; }

        public void PostNavigate(string route) { Post(new { kind = "navigate", route }); }
        public void PostRaw(object message) { Post(message); }

        private void RegisterHandlers()
        {
            _handlers["ping"] = _ => Task.FromResult<object>(new { pong = true });
            _handlers["host.info"] = _ => Task.FromResult(BuildHostInfo());
            _handlers["document.summary"] = _ => Task.FromResult<object>(DocumentSummaryPayload());
            _handlers["document.selection"] = _ => Task.FromResult<object>(DocumentSelectionPayload());
            _handlers["workbook.summary"] = _ => Task.FromResult<object>(DocumentSummaryPayload());
            _handlers["workbook.selection"] = _ => Task.FromResult<object>(DocumentSelectionPayload());
            _handlers["chat.send"] = SendAsync;
            _handlers["chat.stop"] = _ => Task.FromResult<object>(Stop());
            _handlers["chat.reset"] = _ => { _agent.Reset(); return Task.FromResult<object>(new { reset = true }); };
            _handlers["approval.respond"] = RespondApprovalAsync;
            _handlers["approval.revoke"] = _ => Task.FromResult<object>(new { ok = true, revoked = _agent.RevokeApprovalGrants() });
            _handlers["draft.insert"] = InsertDraftAsync;
            _handlers["draft.style"] = UpdateDraftStyleAsync;
            _handlers["draft.discard"] = DiscardDraftAsync;
            _handlers["context.state"] = _ => Task.FromResult<object>(new { used = _agent.Conversation.EstimateTotalTokens(), budget = Math.Max(1, _settings.ContextBudgetTokens), ratio = 0.0, percent = 0, threshold = 85, nearLimit = false });
            _handlers["context.compact"] = _ => Task.FromResult<object>(new { trimmed = false, before = _agent.Conversation.EstimateTotalTokens(), after = _agent.Conversation.EstimateTotalTokens(), compressed = 0, dropped = 0, budget = _settings.ContextBudgetTokens });
            _handlers["session.update"] = UpdateSessionAsync;
            _handlers["pane.ensureWidth"] = EnsureWidth;
            _handlers["pane.saveWidth"] = _ => Task.FromResult<object>(new { saved = (WidthPersister?.Invoke() ?? -1) > 0 });
            _handlers["pane.saveTheme"] = SaveTheme;
            _handlers["undo.apply"] = UndoAsync;
            _handlers["sheet.goto"] = _ => Task.FromResult<object>(new { ok = false, message = "Word 文档没有工作表地址；请使用 Story、Start/End、段落、表格、书签或内容控件目标。", errorCode = "UNSUPPORTED_TARGET" });
            _handlers["sheet.fit"] = _ => Task.FromResult<object>(new { ok = false, message = "Word 文档不支持表格排版快捷操作。", errorCode = "UNSUPPORTED_TARGET" });
            _handlers["client.log"] = payload => { ChatSheet.AddIn.Log.Info("[Word 面板] " + (payload.Value<string>("message") ?? string.Empty)); return Task.FromResult<object>(new { logged = true }); };
        }

        private Task<object> EnsureWidth(JObject payload)
        {
            var current = payload.Value<int?>("currentCss") ?? 0; var target = payload.Value<int?>("targetCss") ?? 0; var dpr = payload.Value<double?>("devicePixelRatio") ?? 1;
            var applied = current > 0 && target > current ? WidthAdjuster?.Invoke(current, target, dpr) ?? -1 : -1;
            return Task.FromResult<object>(new { adjusted = applied > 0, hostWidth = applied });
        }

        private Task<object> SaveTheme(JObject payload)
        {
            var theme = payload.Value<string>("theme") ?? string.Empty; return Task.FromResult<object>(new { applied = ThemeApplier?.Invoke(theme) == true, theme });
        }

        private async Task<object> SendAsync(JObject payload)
        {
            var text = payload.Value<string>("text") ?? string.Empty; var settings = Settings.Load(); _settings = settings; var cts = BeginRun();
            try
            {
                // chat.js 会一直保持忙碌态，直到本次请求完成。Excel 桥遵守同一协议；
                // 如果这里立即返回 started，Word 还没收到任何进度事件就会移除处理中气泡。
                await Task.Run(
                    () => _agent.RunAsync(text, settings, PushAgentUpdateAsync, RequestApprovalAsync, cts.Token, payload.Value<string>("draftId")),
                    cts.Token).ConfigureAwait(false);
                return new { completed = true };
            }
            catch (OperationCanceledException)
            {
                await PushAgentUpdateAsync(new AgentUpdate { Kind = "stopped", Text = "文档任务已停止。" }).ConfigureAwait(false);
                return new { completed = false, stopped = true };
            }
            catch (Exception ex)
            {
                var errorCode = ex is ProviderException provider ? provider.Code : "WORD_AGENT_ERROR";
                await PushAgentUpdateAsync(new AgentUpdate
                {
                    Kind = "error",
                    Text = ex.Message,
                    Payload = new { errorCode },
                }).ConfigureAwait(false);
                return new { completed = false, error = ex.Message, code = errorCode };
            }
            finally
            {
                EndRun(cts);
            }
        }

        private object Stop() { try { Volatile.Read(ref _run)?.Cancel(); } catch { } return new { stopped = true }; }

        private async Task<object> InsertDraftAsync(JObject payload)
        {
            var settings = Settings.Load();
            var cts = BeginRun();
            try
            {
                var result = await _agent.InsertDraftAsync(payload.Value<string>("id"), settings, PushAgentUpdateAsync,
                    RequestApprovalAsync, cts.Token).ConfigureAwait(false);
                return DraftResult(result);
            }
            finally
            {
                EndRun(cts);
            }
        }

        private async Task<object> UpdateDraftStyleAsync(JObject payload)
        {
            if (Volatile.Read(ref _run) != null) { throw new ProviderException("BUSY", "文档任务完成后才能调整草稿风格。"); }
            var result = await _agent.UpdateDraftStyleAsync(payload.Value<string>("id"), payload.Value<string>("style"),
                PushAgentUpdateAsync).ConfigureAwait(false);
            return DraftResult(result);
        }

        private async Task<object> DiscardDraftAsync(JObject payload)
        {
            if (Volatile.Read(ref _run) != null) { throw new ProviderException("BUSY", "文档任务完成后才能丢弃草稿。"); }
            var result = await _agent.DiscardDraftAsync(payload.Value<string>("id"), PushAgentUpdateAsync).ConfigureAwait(false);
            return DraftResult(result);
        }

        private static object DraftResult(ToolResult result)
        {
            return new { ok = result.Ok, message = result.Error, errorCode = result.ErrorCode, data = result.Data };
        }

        private CancellationTokenSource BeginRun()
        {
            var cts = new CancellationTokenSource();
            if (Interlocked.CompareExchange(ref _run, cts, null) == null) { return cts; }
            cts.Dispose();
            throw new ProviderException("BUSY", "当前已有文档任务正在运行。");
        }

        private void EndRun(CancellationTokenSource cts)
        {
            if (ReferenceEquals(Interlocked.CompareExchange(ref _run, null, cts), cts)) { cts.Dispose(); }
        }

        private Task<ApprovalDecision> RequestApprovalAsync(ToolDefinition definition, JObject args, ImpactEstimate impact)
        {
            var id = "word-approval-" + Interlocked.Increment(ref _approvalSequence); var source = new TaskCompletionSource<ApprovalDecision>(TaskCreationOptions.RunContinuationsAsynchronously); _approvals[id] = source;
            _ = PushRawAsync(new { kind = "approval-request", id, tool = definition.Name, risk = definition.Risk.ToString(), args, impact = impact?.Text, impactNote = impact?.Note, hostPreview = impact?.HostPreview, host = "Word/WPS Writer" });
            return AwaitApproval(id, source.Task);
        }

        private async Task<ApprovalDecision> AwaitApproval(string id, Task<ApprovalDecision> task)
        {
            try { return await task.ConfigureAwait(false); } finally { _approvals.TryRemove(id, out _); }
        }

        private Task<object> RespondApprovalAsync(JObject payload)
        {
            var id = payload.Value<string>("id"); if (string.IsNullOrEmpty(id) || !_approvals.TryRemove(id, out var source)) { return Task.FromResult<object>(new { ok = false, message = "审批请求已过期。" }); }
            source.TrySetResult(new ApprovalDecision { Approved = payload.Value<bool?>("approved") == true, Reason = payload.Value<string>("reason"), ApproveRest = payload.Value<bool?>("approveRest") == true, ApproveStructureRest = payload.Value<bool?>("approveStructureRest") == true }); return Task.FromResult<object>(new { ok = true });
        }

        private async Task<object> UndoAsync(JObject payload)
        {
            var id = payload.Value<string>("id");
            if (string.IsNullOrWhiteSpace(id)) { return new { ok = false, message = "缺少文档快照标识。", errorCode = "UNDO_UNAVAILABLE" }; }
            var result = (object)await InvokeOnUiAsync(() =>
            {
                var document = _context.ActiveDocument();
                try { var ok = _agent.Tools.Undo.TryUndo(id, document, out var message); return new { ok, message, errorCode = ok ? null : "UNDO_FAILED" }; }
                finally { WordCom.Release(document); }
            }).ConfigureAwait(false);
            return result;
        }

        private Task<object> UpdateSessionAsync(JObject payload)
        {
            // 设置页通过共享 AgentChannels 保存连接后，不能继续改写构造时的旧快照；
            // 否则切到 WorkBuddy 后对话页的快捷模型选择会把模式恢复成旧值。
            _settings = Settings.Load();
            if (payload.Value<string>("model") != null) { _settings.Model = payload.Value<string>("model").Trim(); _settings.StampModelConnection(); }
            if (Thinking.TryParse(payload.Value<string>("thinking"), out var thinking)) { _settings.Thinking = thinking; }
            if (Enum.TryParse(payload.Value<string>("approval"), out ApprovalPolicy approval)) { _settings.Approval = approval; _agent.UpdateApprovalPolicy(approval); }
            _settings.Save();
            _sharedChannels.ReloadSettings();
            return Task.FromResult<object>(new { model = _settings.Model, thinking = _settings.Thinking.ToString(), approval = _settings.Approval.ToString(), thinkingSupported = true });
        }

        private object DocumentSummaryPayload() { var summary = _context.GetSummary(); return new { hasDocument = summary.HasDocument, name = summary.Name, path = summary.Path, saved = summary.Saved, readOnly = summary.ReadOnly, host = summary.Host, pages = summary.Pages, sections = summary.Sections, paragraphs = summary.Paragraphs, tables = summary.Tables, promptText = summary.ToPromptText() }; }
        private object DocumentSelectionPayload() { var selection = _context.GetSelection(); return new { hasSelection = selection.HasSelection, story = selection.Story, storyType = selection.StoryType, start = selection.Start, end = selection.End, text = selection.Text, promptText = selection.ToPromptText() }; }
        private object BuildHostInfo() { var application = _applicationAccessor(); var kind = WordHostProbe.Detect(application, "OfficeHelper.Word.AddIn"); return new { host = WordHostProbe.DisplayName(kind), hostKind = kind.ToString(), hostMode = "word", process = WordHostProbe.CurrentProcessName() + ".exe", progId = "OfficeHelper.Word.AddIn", bitness = Environment.Is64BitProcess ? "x64" : "x86", hostName = WordCom.String(application, "Name"), hostVersion = WordCom.String(application, "Version"), hostBuild = WordCom.String(application, "Build"), webview2 = SafeBrowserVersion(), clr = Environment.Version.ToString(), logPath = ChatSheet.AddIn.Log.CurrentPath, addInVersion = typeof(WordBridge).Assembly.GetName().Version.ToString() }; }
        private string SafeBrowserVersion() { try { return _core?.Environment?.BrowserVersionString ?? string.Empty; } catch { return string.Empty; } }

        private Task PushAgentUpdateAsync(AgentUpdate update) { Post(new { kind = "agent", stage = update.Kind, text = update.Text, payload = update.Payload }); return Task.CompletedTask; }
        private Task PushRawAsync(object message) { Post(message); return Task.CompletedTask; }
        private Task<object> InvokeOnUiAsync(Func<object> work)
        {
            if (_disposed) { return Task.FromException<object>(new OperationCanceledException("文档面板已关闭。")); }
            if (_uiContext == null || SynchronizationContext.Current == _uiContext) { try { return Task.FromResult(work()); } catch (Exception ex) { return Task.FromException<object>(ex); } }
            var tcs = new TaskCompletionSource<object>(TaskCreationOptions.RunContinuationsAsynchronously); _uiContext.Post(_ => { try { tcs.TrySetResult(work()); } catch (Exception ex) { tcs.TrySetException(ex); } }, null); return tcs.Task;
        }
        private async void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            if (_disposed || !ChatSheet.AddIn.TaskPaneControl.IsTrustedPanelUri(e.Source)) { return; }
            string id = null; try { var envelope = JObject.Parse(e.WebMessageAsJson); id = envelope.Value<string>("id"); var channel = envelope.Value<string>("channel"); if (!_handlers.TryGetValue(channel ?? string.Empty, out var handler)) { Reply(id, false, null, "未知通道：" + channel); return; } var result = await handler(envelope["payload"] as JObject ?? new JObject()).ConfigureAwait(true); Reply(id, true, result, null); } catch (Exception ex) { Reply(id, false, null, ex.Message); }
        }
        private void Reply(string id, bool ok, object data, string error) { if (!string.IsNullOrEmpty(id)) { Post(new { kind = "response", id, ok, data, error }); } }
        private void Post(object message) { if (_disposed) { return; } try { var json = JsonConvert.SerializeObject(message); if (_uiContext != null && SynchronizationContext.Current != _uiContext) { _uiContext.Post(_ => PostCore(json), null); } else { PostCore(json); } } catch { } }
        private void PostCore(string json) { try { if (!_disposed) { _core.PostWebMessageAsJson(json); } } catch { } }
        public void Dispose() { if (_disposed) { return; } _disposed = true; try { _run?.Cancel(); } catch { } foreach (var pair in _approvals) { pair.Value.TrySetResult(new ApprovalDecision { Approved = false, Reason = "文档面板已关闭。" }); } _approvals.Clear(); _sharedChannels.Dispose(); _lifetime.Cancel(); _core.WebMessageReceived -= OnWebMessageReceived; _lifetime.Dispose(); }
    }
}
