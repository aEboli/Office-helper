using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
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
        private readonly List<WordApprovalGrant> _approvalGrants = new List<WordApprovalGrant>();
        private Settings _activeSettings;

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
        internal void Reset() { _conversation.Clear(); }
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
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(userInput)) { throw new ProviderException("EMPTY_INPUT", "请输入内容。"); }
            lock (_approvalLock) { _activeSettings = settings; }
            RevokeApprovalGrants();
            await PublishApprovalGrantsAsync(onUpdate).ConfigureAwait(false);
            try
            {
                var connection = settings.ResolveConnection();
                if (!connection.IsWorkBuddy && string.IsNullOrWhiteSpace(connection.Model)) { throw new ProviderException("MODEL_REQUIRED", "尚未选择模型，请到设置页选择。"); }
                var summary = (WordDocumentSummary)await _uiInvoker(() => _context.GetSummary()).ConfigureAwait(false);
                var documentKey = WordDocumentContext.DocumentKey(summary);
                var selection = (WordSelectionInfo)await _uiInvoker(() => _context.GetSelection(settings.AutoIncludeSelection)).ConfigureAwait(false);
                _conversation.SetSystemPrompt(WordSystemPrompt.Build(summary, selection, approval: CurrentApproval(settings)) + "\n" + WordToolCatalog.PromptSection(connection.IsWorkBuddy));
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
                            await ExecuteCallAsync(call, settings, documentKey, onUpdate, requestApproval, cancellationToken).ConfigureAwait(false);
                        }
                    }
                }
                await onUpdate(new AgentUpdate { Kind = "step-limit", Text = "已达到文档任务步数上限，请继续输入以接着处理。" }).ConfigureAwait(false);
            }
            finally
            {
                lock (_approvalLock)
                {
                    if (ReferenceEquals(_activeSettings, settings)) { _activeSettings = null; }
                }
            }
        }

        private async Task ExecuteCallAsync(
            ToolCall call,
            Settings settings,
            string documentKey,
            Func<AgentUpdate, Task> onUpdate,
            Func<ToolDefinition, JObject, ImpactEstimate, Task<ApprovalDecision>> requestApproval,
            CancellationToken cancellationToken)
        {
            if (!_definitions.TryGetValue(call.Name, out var definition))
            {
                var unknown = ToolResult.Failure("UNKNOWN_TOOL", "Word 工具不存在：" + call.Name);
                AddResult(call, unknown);
                await onUpdate(new AgentUpdate { Kind = "tool-result", Payload = ResultPayload(call, unknown) }).ConfigureAwait(false);
                return;
            }
            JObject args;
            try { args = string.IsNullOrWhiteSpace(call.ArgumentsJson) ? new JObject() : JObject.Parse(call.ArgumentsJson); }
            catch { args = new JObject(); }

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
                    await PublishResultAsync(call, ToolResult.Failure("DOCUMENT_CHANGED", "任务开始后当前文档已切换，已拒绝在另一份文档中继续操作。"), onUpdate).ConfigureAwait(false);
                    return;
                }

                var decision = await requestApproval(definition, args, new ImpactEstimate
                {
                    Text = BuildImpact(definition, args),
                    Note = "Word/WPS Writer 写操作会在当前审批后执行；审批前仅读取目标生成预览。",
                    HostPreview = approvalPreview.Preview,
                }).ConfigureAwait(false);
                if (decision == null || !decision.Approved)
                {
                    await PublishResultAsync(call, ToolResult.Failure("APPROVAL_DENIED", decision?.Reason ?? "用户拒绝了文档操作。"), onUpdate).ConfigureAwait(false);
                    return;
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

            var undoId = definition.RequiresApproval ? "word-" + Guid.NewGuid().ToString("N").Substring(0, 10) : null;
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
            await PublishResultAsync(call, result, onUpdate).ConfigureAwait(false);
        }

        private async Task PublishResultAsync(ToolCall call, ToolResult result, Func<AgentUpdate, Task> onUpdate)
        {
            AddResult(call, result);
            await onUpdate(new AgentUpdate { Kind = "tool-result", Payload = ResultPayload(call, result) }).ConfigureAwait(false);
        }

        private ApprovalPolicy CurrentApproval(Settings settings)
        {
            lock (_approvalLock)
            {
                return ReferenceEquals(_activeSettings, settings) ? _activeSettings.Approval : settings.Approval;
            }
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
