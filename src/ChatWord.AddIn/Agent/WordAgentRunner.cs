using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using ChatSheet.AddIn.Agent;
using ChatSheet.AddIn.Providers;
using ChatSheet.AddIn.Storage;
using ChatSheet.AddIn.Tools;
using ChatWord.AddIn.Hosts;
using ChatWord.AddIn.Tools;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ChatWord.AddIn.Agent
{
    internal sealed class WordAgentRunner
    {
        private readonly Func<object> _applicationAccessor;
        private readonly Func<Func<object>, Task<object>> _uiInvoker;
        private readonly WordToolExecutor _tools;
        private readonly WordDocumentContext _context;
        private readonly Conversation _conversation = new Conversation();
        private readonly Dictionary<string, ToolDefinition> _definitions = new Dictionary<string, ToolDefinition>(StringComparer.Ordinal);
        private readonly object _approvalLock = new object();
        private readonly object _draftLock = new object();
        private readonly List<WordApprovalGrant> _approvalGrants = new List<WordApprovalGrant>();
        private readonly Dictionary<string, WordDocumentDraft> _drafts = new Dictionary<string, WordDocumentDraft>(StringComparer.Ordinal);
        private Settings _activeSettings;
        private bool _draftOnlyForCurrentTurn;
        private bool _draftInsertAllowedForCurrentTurn;

        internal WordAgentRunner(Func<object> applicationAccessor, string progId, Func<Func<object>, Task<object>> uiInvoker = null)
        {
            _applicationAccessor = applicationAccessor ?? throw new ArgumentNullException(nameof(applicationAccessor));
            _uiInvoker = uiInvoker ?? (work => Task.FromResult(work()));
            _tools = new WordToolExecutor(applicationAccessor, progId);
            _context = _tools.Context;
            foreach (var definition in WordToolCatalog.All) { _definitions[definition.Name] = definition; }
        }

        internal Conversation Conversation => _conversation;
        internal WordToolExecutor Tools => _tools;
        internal void Reset()
        {
            _conversation.Clear();
            lock (_draftLock) { _drafts.Clear(); }
        }
        internal int RevokeApprovalGrants()
        {
            lock (_approvalLock)
            {
                var count = _approvalGrants.Count;
                _approvalGrants.Clear();
                return count;
            }
        }

        internal void UpdateApprovalPolicy(ApprovalPolicy approval)
        {
            lock (_approvalLock)
            {
                if (_activeSettings != null) { _activeSettings.Approval = approval; }
            }
        }

        internal async Task RunAsync(
            string userInput,
            Settings settings,
            Func<AgentUpdate, Task> onUpdate,
            Func<ToolDefinition, JObject, ImpactEstimate, Task<ApprovalDecision>> requestApproval,
            CancellationToken cancellationToken,
            string draftId = null)
        {
            if (string.IsNullOrWhiteSpace(userInput)) { throw new ProviderException("EMPTY_INPUT", "请输入内容。"); }
            lock (_approvalLock)
            {
                _activeSettings = settings;
                _draftInsertAllowedForCurrentTurn = string.IsNullOrWhiteSpace(draftId) && IsExplicitDraftInsertion(userInput);
                _draftOnlyForCurrentTurn = RequiresDraftPreview(userInput) || !string.IsNullOrWhiteSpace(draftId) || _draftInsertAllowedForCurrentTurn;
            }
            RevokeApprovalGrants();
            await PublishApprovalGrantsAsync(onUpdate).ConfigureAwait(false);
            try
            {
                var connection = settings.ResolveConnection();
                if (!connection.IsWorkBuddy && string.IsNullOrWhiteSpace(connection.Model)) { throw new ProviderException("MODEL_REQUIRED", "尚未选择模型，请到设置页选择。"); }
                var documentSnapshot = (Tuple<WordDocumentSummary, WordSelectionInfo, WordSelectionInfo>)await _uiInvoker(() =>
                {
                    var currentSummary = _context.GetSummary();
                    var currentSelection = _context.GetSelection(settings.AutoIncludeSelection);
                    var draftTarget = settings.AutoIncludeSelection ? currentSelection : _context.GetSelection();
                    return Tuple.Create(currentSummary, currentSelection, draftTarget);
                }).ConfigureAwait(false);
                var summary = documentSnapshot.Item1;
                var documentKey = WordDocumentContext.DocumentKey(summary);
                var selection = documentSnapshot.Item2;
                var draftTarget = documentSnapshot.Item3;
                var systemPrompt = WordSystemPrompt.Build(summary, selection, approval: CurrentApproval(settings)) + "\n" + WordToolCatalog.PromptSection(connection.IsWorkBuddy);
                if (!string.IsNullOrWhiteSpace(draftId))
                {
                    systemPrompt += "\n用户正在继续调整侧栏中的现有 Word 草稿。必须调用 draft_document 并原样传入 draft_id=" + draftId +
                        "，保留草稿的 operation 和目标；只按用户要求更新内容或风格，不得新建另一份草稿或插入文档。";
                }
                _conversation.SetSystemPrompt(systemPrompt);
                _conversation.Add(ChatMessage.FromUser(userInput));

                using (var client = CreateClient(settings, connection))
                {
                    for (var step = 0; settings.MaxSteps == 0 || step < Math.Max(1, settings.MaxSteps); step++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var outcome = await StreamAsync(client, connection, settings, onUpdate, cancellationToken).ConfigureAwait(false);
                        var assistant = ChatMessage.FromAssistant(outcome.RawText ?? outcome.Text ?? string.Empty);
                        assistant.ToolCalls.AddRange(outcome.Calls);
                        _conversation.Add(assistant);
                        if (outcome.Calls.Count == 0)
                        {
                            await onUpdate(new AgentUpdate { Kind = "turn-complete", Payload = new { finishReason = outcome.FinishReason } }).ConfigureAwait(false);
                            return;
                        }

                        foreach (var call in outcome.Calls)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            await ExecuteCallAsync(call, settings, documentKey, draftTarget, onUpdate, requestApproval, cancellationToken).ConfigureAwait(false);
                        }
                    }
                }
                await onUpdate(new AgentUpdate { Kind = "step-limit", Text = "已达到文档任务步数上限，请继续输入以接着处理。" }).ConfigureAwait(false);
            }
            finally
            {
                lock (_approvalLock)
                {
                    if (ReferenceEquals(_activeSettings, settings))
                    {
                        _activeSettings = null;
                        _draftOnlyForCurrentTurn = false;
                        _draftInsertAllowedForCurrentTurn = false;
                    }
                }
            }
        }

        private async Task<ToolResult> ExecuteCallAsync(
            ToolCall call,
            Settings settings,
            string documentKey,
            WordSelectionInfo draftTarget,
            Func<AgentUpdate, Task> onUpdate,
            Func<ToolDefinition, JObject, ImpactEstimate, Task<ApprovalDecision>> requestApproval,
            CancellationToken cancellationToken,
            ToolDefinition definitionOverride = null,
            bool recordInConversation = true)
        {
            if (call.Name == "draft_document")
            {
                return await CreateDraftAsync(call, documentKey, draftTarget, onUpdate, recordInConversation).ConfigureAwait(false);
            }

            JObject args;
            try { args = string.IsNullOrWhiteSpace(call.ArgumentsJson) ? new JObject() : JObject.Parse(call.ArgumentsJson); }
            catch { args = new JObject(); }

            if (definitionOverride == null && !_definitions.TryGetValue(call.Name, out definitionOverride))
            {
                var unknown = ToolResult.Failure("UNKNOWN_TOOL", "Word 工具不存在：" + call.Name);
                await PublishResultAsync(call, unknown, onUpdate, recordInConversation).ConfigureAwait(false);
                return unknown;
            }
            var definition = definitionOverride;

            if (IsDraftOnlyTurn() && definition.Risk != ToolRisk.Read &&
                !(call.Name == "insert_document_draft" && IsDraftInsertAllowed()))
            {
                var blocked = ToolResult.Failure("DRAFT_CONFIRMATION_REQUIRED", "生成和草稿调整阶段只能创建或更新预览；请在草稿卡确认插入，或在之后的新对话轮明确要求插入已有草稿。");
                await PublishResultAsync(call, blocked, onUpdate, recordInConversation).ConfigureAwait(false);
                return blocked;
            }

            string draftIdToClose = null;
            if (call.Name == "insert_document_draft")
            {
                var requestedId = args.Value<string>("draft_id");
                WordDocumentDraft draft = null;
                if (!string.IsNullOrWhiteSpace(requestedId))
                {
                    lock (_draftLock) { _drafts.TryGetValue(requestedId, out draft); }
                }
                if (draft == null)
                {
                    var missing = ToolResult.Failure("DRAFT_NOT_FOUND", "待插入的草稿已不存在。");
                    await PublishResultAsync(call, missing, onUpdate, recordInConversation).ConfigureAwait(false);
                    return missing;
                }
                args = draft.ToExecutionArgs();
                draftIdToClose = draft.Id;
            }

            await onUpdate(new AgentUpdate { Kind = "tool-start", Payload = new { id = call.Id, name = call.Name, risk = definition.Risk.ToString(), args } }).ConfigureAwait(false);
            var approvalClass = ApprovalClassFor(call.Name);
            var approval = CurrentApproval(settings);
            if (definition.RequiresApproval && approval != ApprovalPolicy.Automatic && !IsGranted(documentKey, approvalClass))
            {
                var approvalPreview = (ApprovalPreview)await _uiInvoker(() =>
                {
                    var currentSummary = _context.GetSummary();
                    return new ApprovalPreview
                    {
                        DocumentKey = WordDocumentContext.DocumentKey(currentSummary),
                        DocumentName = currentSummary.Name,
                        Preview = _tools.BuildPreview(call.Name, args),
                    };
                }).ConfigureAwait(false);
                if (!string.Equals(documentKey, approvalPreview.DocumentKey, StringComparison.Ordinal))
                {
                    var changed = ToolResult.Failure("DOCUMENT_CHANGED", "任务开始后当前文档已切换，已拒绝在另一份文档中继续操作。");
                    await PublishResultAsync(call, changed, onUpdate, recordInConversation).ConfigureAwait(false);
                    return changed;
                }

                var preview = approvalPreview.Preview == null ? null : JObject.FromObject(approvalPreview.Preview);
                if (call.Name == "insert_document_draft" && preview?.Value<bool?>("supported") != true)
                {
                    var unsupported = ToolResult.Failure("DRAFT_PREVIEW_FAILED", preview?.Value<string>("error") ?? "无法安全预览此草稿目标。");
                    await PublishResultAsync(call, unsupported, onUpdate, recordInConversation).ConfigureAwait(false);
                    return unsupported;
                }

                var decision = await requestApproval(definition, args, new ImpactEstimate
                {
                    Text = BuildImpact(definition, args),
                    Note = "Word/WPS Writer 写操作会在当前审批后执行；审批前仅读取目标生成预览。",
                    HostPreview = approvalPreview.Preview,
                }).ConfigureAwait(false);
                if (decision == null || !decision.Approved)
                {
                    var denied = ToolResult.Failure("APPROVAL_DENIED", decision?.Reason ?? "用户拒绝了文档操作。");
                    await PublishResultAsync(call, denied, onUpdate, recordInConversation).ConfigureAwait(false);
                    return denied;
                }

                approval = CurrentApproval(settings);
                if (approval == ApprovalPolicy.PerTurn || decision.ApproveRest || decision.ApproveStructureRest)
                {
                    AddGrant(documentKey, approvalPreview.DocumentName, approvalClass);
                }
                if (decision.ApproveStructureRest)
                {
                    AddGrant(documentKey, approvalPreview.DocumentName, WordApprovalClass.Structure);
                }
                await PublishApprovalGrantsAsync(onUpdate).ConfigureAwait(false);
            }

            var undoId = definition.RequiresApproval && call.Name != "insert_document_draft"
                ? "word-" + Guid.NewGuid().ToString("N").Substring(0, 10) : null;
            ToolResult result;
            try
            {
                result = (ToolResult)await _uiInvoker(() =>
                {
                    var currentKey = WordDocumentContext.DocumentKey(_context.GetSummary());
                    return string.Equals(documentKey, currentKey, StringComparison.Ordinal)
                        ? _tools.Execute(call.Name, args, undoId)
                        : ToolResult.Failure("DOCUMENT_CHANGED", "任务开始后当前文档已切换，已拒绝在另一份文档中继续操作。");
                }).ConfigureAwait(false);
            }
            catch (Exception ex) { result = ToolResult.Failure("HOST_ERROR", ex.Message); }
            await PublishResultAsync(call, result, onUpdate, recordInConversation).ConfigureAwait(false);
            if (result.Ok && draftIdToClose != null)
            {
                lock (_draftLock) { _drafts.Remove(draftIdToClose); }
                await onUpdate(new AgentUpdate { Kind = "draft-closed", Payload = new { id = draftIdToClose, status = "inserted" } }).ConfigureAwait(false);
            }
            return result;
        }

        private async Task<ToolResult> CreateDraftAsync(ToolCall call, string documentKey, WordSelectionInfo draftTarget, Func<AgentUpdate, Task> onUpdate, bool recordInConversation)
        {
            JObject args;
            try { args = string.IsNullOrWhiteSpace(call.ArgumentsJson) ? new JObject() : JObject.Parse(call.ArgumentsJson); }
            catch { args = new JObject(); }

            var currentSummary = (WordDocumentSummary)await _uiInvoker(() => _context.GetSummary()).ConfigureAwait(false);
            if (!string.Equals(documentKey, WordDocumentContext.DocumentKey(currentSummary), StringComparison.Ordinal))
            {
                return await StoreDraftToolResult(call, ToolResult.Failure("DOCUMENT_CHANGED", "生成期间活动文档已切换；请在目标文档中重新生成草稿。"), onUpdate, recordInConversation).ConfigureAwait(false);
            }

            var requestedId = args.Value<string>("draft_id");
            WordDocumentDraft existing = null;
            if (!string.IsNullOrWhiteSpace(requestedId))
            {
                lock (_draftLock) { _drafts.TryGetValue(requestedId, out existing); }
                if (existing == null)
                {
                    return await StoreDraftToolResult(call, ToolResult.Failure("DRAFT_NOT_FOUND", "待更新的草稿已不存在；请重新生成草稿。"), onUpdate, recordInConversation).ConfigureAwait(false);
                }
            }

            WordDocumentDraft draft;
            try { draft = WordDocumentDraft.Create(documentKey, currentSummary.Name, draftTarget, args, existing); }
            catch (WordToolException ex)
            { return await StoreDraftToolResult(call, ToolResult.Failure(ex.Code, ex.Message), onUpdate, recordInConversation).ConfigureAwait(false); }

            lock (_draftLock) { _drafts[draft.Id] = draft; }
            if (recordInConversation) { AddResult(call, ToolResult.Success(draft.ToModelResult())); }
            await onUpdate(new AgentUpdate { Kind = "document-draft", Payload = draft.ToPreviewPayload() }).ConfigureAwait(false);
            return ToolResult.Success(draft.ToModelResult());
        }

        private async Task<ToolResult> StoreDraftToolResult(ToolCall call, ToolResult result, Func<AgentUpdate, Task> onUpdate, bool recordInConversation)
        {
            if (recordInConversation) { AddResult(call, result); }
            await onUpdate(new AgentUpdate { Kind = "tool-result", Payload = ResultPayload(call, result) }).ConfigureAwait(false);
            return result;
        }

        internal async Task<ToolResult> InsertDraftAsync(
            string id,
            Settings settings,
            Func<AgentUpdate, Task> onUpdate,
            Func<ToolDefinition, JObject, ImpactEstimate, Task<ApprovalDecision>> requestApproval,
            CancellationToken cancellationToken)
        {
            WordDocumentDraft draft;
            lock (_draftLock) { _drafts.TryGetValue(id ?? string.Empty, out draft); }
            if (draft == null) { return ToolResult.Failure("DRAFT_NOT_FOUND", "待插入的草稿已不存在。"); }

            lock (_approvalLock)
            {
                _activeSettings = settings;
                _draftOnlyForCurrentTurn = false;
                _draftInsertAllowedForCurrentTurn = true;
            }
            try
            {
                var call = new ToolCall
                {
                    Id = "word-draft-" + Guid.NewGuid().ToString("N").Substring(0, 8),
                    Name = "insert_document_draft",
                    ArgumentsJson = draft.ToExecutionArgs().ToString(Formatting.None),
                };
                var result = await ExecuteCallAsync(call, settings, draft.DocumentKey, null, onUpdate, requestApproval,
                    cancellationToken, WordToolCatalog.InsertDraft, recordInConversation: false).ConfigureAwait(false);
                return result;
            }
            finally
            {
                lock (_approvalLock)
                {
                    if (ReferenceEquals(_activeSettings, settings))
                    {
                        _activeSettings = null;
                        _draftOnlyForCurrentTurn = false;
                        _draftInsertAllowedForCurrentTurn = false;
                    }
                }
            }
        }

        internal async Task<ToolResult> UpdateDraftStyleAsync(string id, string style, Func<AgentUpdate, Task> onUpdate)
        {
            WordDocumentDraft draft;
            try
            {
                lock (_draftLock)
                {
                    if (!_drafts.TryGetValue(id ?? string.Empty, out draft))
                    { return ToolResult.Failure("DRAFT_NOT_FOUND", "待调整的草稿已不存在。"); }
                    draft.WithStyle(style);
                }
            }
            catch (WordToolException ex) { return ToolResult.Failure(ex.Code, ex.Message); }
            await onUpdate(new AgentUpdate { Kind = "document-draft", Payload = draft.ToPreviewPayload() }).ConfigureAwait(false);
            return ToolResult.Success(new { id = draft.Id, style = draft.Style });
        }

        internal async Task<ToolResult> DiscardDraftAsync(string id, Func<AgentUpdate, Task> onUpdate)
        {
            WordDocumentDraft draft;
            lock (_draftLock)
            {
                if (!_drafts.TryGetValue(id ?? string.Empty, out draft))
                { return ToolResult.Failure("DRAFT_NOT_FOUND", "待丢弃的草稿已不存在。"); }
                _drafts.Remove(draft.Id);
            }
            await onUpdate(new AgentUpdate { Kind = "draft-closed", Payload = new { id = draft.Id, status = "discarded" } }).ConfigureAwait(false);
            return ToolResult.Success(new { id = draft.Id, discarded = true });
        }

        private async Task PublishResultAsync(ToolCall call, ToolResult result, Func<AgentUpdate, Task> onUpdate, bool recordInConversation = true)
        {
            if (recordInConversation) { AddResult(call, result); }
            await onUpdate(new AgentUpdate { Kind = "tool-result", Payload = ResultPayload(call, result) }).ConfigureAwait(false);
        }

        private ApprovalPolicy CurrentApproval(Settings settings)
        {
            lock (_approvalLock)
            {
                return ReferenceEquals(_activeSettings, settings) ? _activeSettings.Approval : settings.Approval;
            }
        }

        private bool IsDraftOnlyTurn()
        {
            lock (_approvalLock) { return _draftOnlyForCurrentTurn; }
        }

        private bool IsDraftInsertAllowed()
        {
            lock (_approvalLock) { return _draftInsertAllowedForCurrentTurn; }
        }

        private static bool RequiresDraftPreview(string input)
        {
            return Regex.IsMatch(input ?? string.Empty,
                "(写|起草|续写|补充|扩写|改写|重写|生成|设计|美化|排版|版式|风格|样式|调整|优化|draft|write|compose|rewrite|generate|continue|adjust|improve)",
                RegexOptions.IgnoreCase);
        }

        private static bool IsExplicitDraftInsertion(string input)
        {
            var hasInsertAction = Regex.IsMatch(input ?? string.Empty, "(插入|替换|写入|应用|insert|replace|write|apply)", RegexOptions.IgnoreCase);
            var refersToExistingDraft = Regex.IsMatch(input ?? string.Empty,
                "(刚才|刚刚|上一轮|上轮|上面|之前|现有|这个|这份|该|已有|已生成|previous|prior|earlier|existing|this|that|already generated).{0,16}(草稿|draft)|(草稿|draft).{0,16}(刚才|刚刚|上一轮|上轮|上面|之前|现有|这个|这份|该|已有|previous|prior|earlier|existing|this|that)",
                RegexOptions.IgnoreCase);
            return hasInsertAction && refersToExistingDraft;
        }

        private bool IsGranted(string documentKey, WordApprovalClass approvalClass)
        {
            if (string.IsNullOrEmpty(documentKey)) { return false; }
            lock (_approvalLock)
            {
                return _approvalGrants.Any(grant => grant.ApprovalClass == approvalClass &&
                    string.Equals(grant.DocumentKey, documentKey, StringComparison.Ordinal));
            }
        }

        private void AddGrant(string documentKey, string documentName, WordApprovalClass approvalClass)
        {
            if (string.IsNullOrEmpty(documentKey)) { return; }
            lock (_approvalLock)
            {
                if (_approvalGrants.Any(grant => grant.ApprovalClass == approvalClass &&
                    string.Equals(grant.DocumentKey, documentKey, StringComparison.Ordinal))) { return; }
                _approvalGrants.Add(new WordApprovalGrant
                {
                    DocumentKey = documentKey,
                    DocumentName = documentName,
                    ApprovalClass = approvalClass,
                });
            }
        }

        private async Task PublishApprovalGrantsAsync(Func<AgentUpdate, Task> onUpdate)
        {
            object[] grants;
            lock (_approvalLock)
            {
                grants = _approvalGrants.Select(grant => (object)new
                {
                    sheet = grant.DocumentName ?? "当前文档",
                    workbookWide = false,
                    approvalClass = grant.ApprovalClass.ToString(),
                }).ToArray();
            }
            await onUpdate(new AgentUpdate { Kind = "approval-grants", Payload = new { grants } }).ConfigureAwait(false);
        }

        private static WordApprovalClass ApprovalClassFor(string toolName)
        {
            switch (toolName)
            {
                case "format_text":
                case "format_paragraph":
                case "apply_style": return WordApprovalClass.Format;
                case "delete_range": return WordApprovalClass.Destructive;
                case "set_page_setup":
                case "insert_page_break": return WordApprovalClass.Structure;
                default: return WordApprovalClass.Write;
            }
        }

        private void AddResult(ToolCall call, ToolResult result)
        {
            _conversation.Add(ChatMessage.FromToolResult(call.Id, call.Name, JsonConvert.SerializeObject(result.ToPayload())));
        }

        private async Task<StepOutcome> StreamAsync(
            IChatStreamClient client,
            ResolvedConnection connection,
            Settings settings,
            Func<AgentUpdate, Task> onUpdate,
            CancellationToken cancellationToken)
        {
            var outcome = new StepOutcome();
            var text = new System.Text.StringBuilder();
            var raw = new System.Text.StringBuilder();
            var gate = connection.IsWorkBuddy ? new TextToolGate() : null;
            var request = new ChatRequest
            {
                Protocol = connection.Protocol, BaseUrl = connection.BaseUrl, Token = connection.Token, Model = connection.Model,
                Thinking = settings.Thinking, Temperature = settings.Temperature, MaxOutputTokens = settings.MaxOutputTokens,
                IncludeTools = !connection.IsWorkBuddy, ToolDefinitions = WordToolCatalog.All,
            };
            request.Messages.AddRange(_conversation.Messages);
            await client.StreamAsync(request, async e =>
            {
                switch (e.Kind)
                {
                    case ChatEventKind.TextDelta:
                        raw.Append(e.Text);
                        var visible = gate == null ? e.Text : gate.Push(e.Text);
                        if (!string.IsNullOrEmpty(visible)) { text.Append(visible); await onUpdate(new AgentUpdate { Kind = "text", Text = visible }).ConfigureAwait(false); }
                        break;
                    case ChatEventKind.ThinkingDelta: await onUpdate(new AgentUpdate { Kind = "thinking", Text = e.Text }).ConfigureAwait(false); break;
                    case ChatEventKind.ToolCall: outcome.Calls.Add(e.Call); break;
                    case ChatEventKind.Completed: if (!string.IsNullOrEmpty(e.FinishReason)) { outcome.FinishReason = e.FinishReason; } break;
                }
            }, cancellationToken).ConfigureAwait(false);
            if (gate != null)
            {
                var tail = gate.Flush();
                if (!string.IsNullOrEmpty(tail)) { text.Append(tail); await onUpdate(new AgentUpdate { Kind = "text", Text = tail }).ConfigureAwait(false); }
                var index = 0;
                foreach (var parsed in gate.Calls)
                {
                    outcome.Calls.Add(new ToolCall { Id = "word-txt" + (++index) + "-" + Guid.NewGuid().ToString("N").Substring(0, 6), Name = parsed.Name, ArgumentsJson = parsed.ArgumentsJson });
                }
            }
            outcome.Text = text.ToString();
            outcome.RawText = raw.ToString();
            return outcome;
        }

        private static IChatStreamClient CreateClient(Settings settings, ResolvedConnection connection)
        {
            return settings.Mode.IsWorkBuddy() ? WorkBuddyProvider.CreateChatClient(settings.Mode) : (IChatStreamClient)new ChatClient(connection.Proxy);
        }

        private static object ResultPayload(ToolCall call, ToolResult result)
        {
            var data = result.Data as JObject ?? (result.Data == null ? null : JObject.FromObject(result.Data));
            return new { id = call.Id, name = call.Name, ok = result.Ok, data, error = result.Error, errorCode = result.ErrorCode, canUndo = data?.Value<bool?>("canUndo") == true, undoId = data?.Value<string>("undoId"), undoNote = data?.Value<string>("undoNote") };
        }

        private static string BuildImpact(ToolDefinition definition, JObject args)
        {
            var target = args?["target"]?.ToString(Formatting.None);
            return definition.Name + "：" + (string.IsNullOrEmpty(target) ? "需要明确文档目标" : target);
        }

        private sealed class StepOutcome
        {
            internal string Text;
            internal string RawText;
            internal string FinishReason;
            internal List<ToolCall> Calls { get; } = new List<ToolCall>();
        }

        private enum WordApprovalClass
        {
            Format,
            Write,
            Destructive,
            Structure,
        }

        private sealed class WordApprovalGrant
        {
            internal string DocumentKey;
            internal string DocumentName;
            internal WordApprovalClass ApprovalClass;
        }

        private sealed class ApprovalPreview
        {
            internal string DocumentKey;
            internal string DocumentName;
            internal object Preview;
        }
    }
}
